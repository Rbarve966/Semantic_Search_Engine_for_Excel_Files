"""
Run this script LOCALLY (on the same Windows PC where images were generated).
It rebuilds png_index.pkl using the exact same hash formula as app.py.

Usage:
    python rebuild_png_index.py

Then commit the resulting png_index.pkl to your repo.
"""

import os
import pickle
import hashlib

# ── config: adjust these paths to match your local setup ─────────────────────
IMAGE_FOLDER   = r"images_final"          # folder where your PNGs live (in repo)
DOCUMENTS_FILE = r"chunked_data.pkl"      # your documents pickle
OUTPUT_FILE    = r"png_index.pkl"         # output mapping file
# ─────────────────────────────────────────────────────────────────────────────

def _cache_key(file_path: str, sheet_name: str) -> str:
    """Must exactly match the _cache_key() function in app.py."""
    raw = f"{os.path.abspath(file_path)}::{sheet_name}"
    return hashlib.md5(raw.encode()).hexdigest()

# Load documents
with open(DOCUMENTS_FILE, "rb") as f:
    documents = pickle.load(f)

# Get all PNG hashes that exist in the image folder
actual_hashes = {
    fname[:-4]  # strip .png
    for fname in os.listdir(IMAGE_FOLDER)
    if fname.endswith(".png")
}
print(f"Found {len(actual_hashes)} PNG files in '{IMAGE_FOLDER}'")

# Build mapping: (basename, sheet_name) -> relative image path
png_index = {}
unmatched = []

for doc in documents:
    file_name  = doc["metadata"]["file_name"]
    sheet_name = doc["metadata"]["sheet_name"]
    key        = (os.path.basename(file_name), sheet_name)

    h = _cache_key(file_name, sheet_name)
    if h in actual_hashes:
        png_index[key] = f"{IMAGE_FOLDER}/{h}.png"
    else:
        unmatched.append((file_name, sheet_name, h))

print(f"Mapped:   {len(png_index)}")
print(f"Unmatched: {len(unmatched)}")

if unmatched:
    print("\nUnmatched entries (first 10):")
    for fn, sn, h in unmatched[:10]:
        print(f"  file={fn!r}  sheet={sn!r}  expected_hash={h}")

with open(OUTPUT_FILE, "wb") as f:
    pickle.dump(png_index, f)

print(f"\nSaved {OUTPUT_FILE}  ({len(png_index)} entries)")
