import streamlit as st
import pickle
import numpy as np
import os
import io
import openpyxl
from sentence_transformers import SentenceTransformer
from sklearn.metrics.pairwise import cosine_similarity

# ── CONFIG ────────────────────────────────────────────────────────────────────
CHUNKED_DATA = "chunked_data.pkl"
EMBEDDINGS   = "embeddings.npy"
PNG_INDEX    = "png_index.pkl"
SAVED_MODEL  = "saved_model"       # local folder in repo
TOP_K        = 3

# ── loaders ───────────────────────────────────────────────────────────────────

@st.cache_resource(show_spinner=False)
def load_model():
    return SentenceTransformer(SAVED_MODEL)

@st.cache_data(show_spinner=False)
def load_data():
    with open(CHUNKED_DATA, "rb") as f:
        chunks = pickle.load(f)
    embs = np.load(EMBEDDINGS)
    return chunks, embs

@st.cache_data(show_spinner=False)
def load_png_index() -> dict:
    if not os.path.exists(PNG_INDEX):
        return {}
    with open(PNG_INDEX, "rb") as f:
        raw = pickle.load(f)
    # Normalise: strip directory prefix from key, fix backslashes in value
    index = {}
    for (fp, sheet), img_path in raw.items():
        index[(os.path.basename(fp), sheet)] = img_path.replace("\\", "/")
    return index

# ── helpers ───────────────────────────────────────────────────────────────────

def get_image(png_index, file_name, sheet_name):
    path = png_index.get((os.path.basename(file_name), sheet_name))
    return path if path and os.path.exists(path) else None

def score_chunks(query, chunks, embeddings, model):
    """Semantic score + exact/keyword boost."""
    q_emb   = model.encode([query])
    scores  = cosine_similarity(q_emb, embeddings)[0].copy()
    q_lower = query.lower().strip()

    for i, chunk in enumerate(chunks):
        text = chunk["text"].lower()
        if f"clause no.: {q_lower}" in text:
            scores[i] += 1.0
        elif q_lower in text:
            scores[i] += 0.3
        else:
            kws = q_lower.split()
            scores[i] += 0.05 * sum(1 for kw in kws if kw in text)

    return scores

# ── app ───────────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Excel Semantic Search", layout="wide")
st.title("📊 Excel Semantic Search")

model            = load_model()
chunked_data, embeddings = load_data()
png_index        = load_png_index()

with st.sidebar:
    st.metric("Chunks indexed", len(chunked_data))
    st.metric("Images mapped",  len(png_index))
    if not png_index:
        st.warning("png_index.pkl not found — previews unavailable.")

query = st.text_input(
    "🔍 Enter your query",
    placeholder="e.g. customer complaints, clause 7.1.6, process audit",
)

if query:
    scores      = score_chunks(query, chunked_data, embeddings, model)
    top_indices = scores.argsort()[-TOP_K:][::-1]

    st.subheader("Top Matching Sheets")

    for rank, idx in enumerate(top_indices, 1):
        chunk      = chunked_data[idx]
        meta       = chunk["metadata"]
        file_name  = meta["file_name"]
        sheet_name = meta["sheet_name"]
        score      = round(float(scores[idx]), 3)
        base_name  = os.path.basename(file_name)

        with st.expander(
            f"#{rank} · {base_name}  —  {sheet_name}  (score: {score})",
            expanded=(rank == 1),
        ):
            col_info, col_dl = st.columns([3, 1])

            with col_info:
                st.caption(f"**File:** {base_name} &nbsp;|&nbsp; **Sheet:** {sheet_name} &nbsp;|&nbsp; **Score:** {score}")

            with col_dl:
                # On cloud the Excel files are not present, so download is
                # not available — we show a clean message instead of an error.
                if os.path.exists(file_name):
                    try:
                        src_wb = openpyxl.load_workbook(file_name, read_only=True, data_only=True)
                        matched = sheet_name
                        if sheet_name not in src_wb.sheetnames:
                            matched = next(
                                (s for s in src_wb.sheetnames if s.strip() == sheet_name.strip()),
                                None
                            )
                        if matched:
                            dst_wb = openpyxl.Workbook()
                            dst_ws = dst_wb.active
                            dst_ws.title = matched
                            for row in src_wb[matched].iter_rows(values_only=True):
                                dst_ws.append(list(row))
                            buf = io.BytesIO()
                            dst_wb.save(buf)
                            buf.seek(0)
                            dl_name = f"{os.path.splitext(base_name)[0]} — {sheet_name}.xlsx"
                            st.download_button(
                                "⬇ Download Sheet",
                                data=buf.read(),
                                file_name=dl_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key=f"dl_{idx}",
                            )
                    except Exception as e:
                        st.caption(f"Download error: {e}")
                else:
                    st.caption("_(download unavailable on cloud)_")

            st.markdown("**Text preview:**")
            st.code(chunk["text"][:300], language=None)

            st.markdown("**Sheet preview:**")
            img = get_image(png_index, file_name, sheet_name)
            if img:
                st.image(img, caption=f"{base_name} › {sheet_name}", use_container_width=True)
            else:
                st.info("No preview available for this sheet.")