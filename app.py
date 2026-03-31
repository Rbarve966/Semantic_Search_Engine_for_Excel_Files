import streamlit as st
import pickle
import numpy as np
import os
import io
import tempfile
import shutil

import pythoncom
import win32com.client
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

# ── CONFIG ────────────────────────────────────────────────────────────────────

INDEX_PATH = "cache/png_index.pkl"

# ── Model & data ──────────────────────────────────────────────────────────────

@st.cache_resource(show_spinner=False)
def load_model():
    return SentenceTransformer("all-MiniLM-L6-v2")


@st.cache_data(show_spinner=False)
def load_data():
    with open("documents.pkl", "rb") as f:
        docs = pickle.load(f)
    embs = np.load("embeddings.npy")
    return docs, embs


@st.cache_data(show_spinner=False)
def load_png_index() -> dict:
    if not os.path.exists(INDEX_PATH):
        return {}
    with open(INDEX_PATH, "rb") as f:
        raw_index = pickle.load(f)
    basename_index = {}
    for (file_path, sheet_name), png_path in raw_index.items():
        base_key = (os.path.basename(file_path), sheet_name)
        basename_index[base_key] = png_path
    return basename_index


def get_png_path(png_index: dict, file_path: str, sheet_name: str) -> str | None:
    key      = (os.path.basename(file_path), sheet_name)
    png_path = png_index.get(key)
    if png_path and os.path.exists(png_path):
        return png_path
    return None


def extract_sheet_as_xlsx(file_path: str, sheet_name: str) -> bytes | None:
    """
    Use Excel COM to copy just the relevant sheet into a new workbook
    and return it as .xlsx bytes. Works for both .xls and .xlsx.
    """
    pythoncom.CoInitialize()  # required when COM is called from a Streamlit thread
    work_dir = tempfile.mkdtemp()
    excel    = None
    src_wb   = None
    dst_wb   = None

    try:
        excel                    = win32com.client.Dispatch("Excel.Application")
        excel.Visible            = False
        excel.DisplayAlerts      = False
        excel.AutomationSecurity = 3

        src_wb = excel.Workbooks.Open(
            os.path.abspath(file_path),
            UpdateLinks=False,
            ReadOnly=True,
        )

        # Find the sheet
        target_sheet = None
        for sheet in src_wb.Sheets:
            if sheet.Name == sheet_name or sheet.Name.strip() == sheet_name.strip():
                target_sheet = sheet
                break

        if target_sheet is None:
            return None

        # Copy sheet to a new workbook
        target_sheet.Copy()
        dst_wb = excel.ActiveWorkbook

        # Save as .xlsx
        out_path = os.path.join(work_dir, "sheet.xlsx")
        dst_wb.SaveAs(
            out_path,
            FileFormat=51,  # 51 = xlOpenXMLWorkbook (.xlsx)
        )
        dst_wb.Close(SaveChanges=False)
        dst_wb = None

        with open(out_path, "rb") as f:
            return f.read()

    except Exception as e:
        st.caption(f"Sheet extraction error: {e}")
        return None

    finally:
        try:
            if dst_wb is not None:
                dst_wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if src_wb is not None:
                src_wb.Close(SaveChanges=False)
        except Exception:
            pass
        try:
            if excel is not None:
                excel.Quit()
        except Exception:
            pass
        shutil.rmtree(work_dir, ignore_errors=True)
        pythoncom.CoUninitialize()


# ── UI ────────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Excel Semantic Search", layout="wide")
st.title("📊 Excel Semantic Search")

model                 = load_model()
documents, embeddings = load_data()
png_index             = load_png_index()

if not png_index:
    st.warning(
        "⚠️ PNG cache not found. "
        "Run `python generate_png_cache.py` first, then restart the app."
    )

top_k = 3

query = st.text_input(
    "🔍 Enter your query",
    placeholder="e.g. quarterly revenue by region",
)

if query:

    query_embedding = model.encode([query])
    scores          = cosine_similarity(query_embedding, embeddings)[0]
    top_indices     = scores.argsort()[-top_k:][::-1]

    st.subheader("Relevant Sheets")

    for rank, idx in enumerate(top_indices, 1):

        doc        = documents[idx]
        meta       = doc["metadata"]
        file_name  = meta["file_name"]
        sheet_name = meta["sheet_name"]
        score      = round(float(scores[idx]), 3)

        with st.expander(
            f"#{rank} · {os.path.basename(file_name)} — sheet: {sheet_name} (score: {score})",
            expanded=(rank == 1),
        ):

            col1, col2 = st.columns([3, 1])

            with col1:
                st.caption(
                    f"File: {file_name} | Sheet: {sheet_name} | Score: {score}"
                )

            with col2:
                sheet_bytes = extract_sheet_as_xlsx(file_name, sheet_name)

                if sheet_bytes:
                    download_name = f"{os.path.splitext(os.path.basename(file_name))[0]} — {sheet_name}.xlsx"
                    st.download_button(
                        label="⬇ Download Sheet",
                        data=sheet_bytes,
                        file_name=download_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{idx}",
                    )
                else:
                    # Fallback — download entire file
                    try:
                        with open(file_name, "rb") as fh:
                            data = fh.read()
                        ext  = os.path.splitext(file_name)[1].lower()
                        mime = (
                            "application/vnd.ms-excel"
                            if ext == ".xls"
                            else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.download_button(
                            label="⬇ Download Excel",
                            data=data,
                            file_name=os.path.basename(file_name),
                            mime=mime,
                            key=f"dl_{idx}",
                        )
                    except FileNotFoundError:
                        st.warning("File not found for download.")

            st.markdown("**Text preview:**")
            st.code(doc["text"][:300], language=None)

            st.markdown("**Sheet preview:**")

            png_path = get_png_path(png_index, file_name, sheet_name)

            if png_path:
                st.image(png_path, caption=sheet_name, use_container_width=True)
            else:
                st.info(f"Preview not available for: {sheet_name}")