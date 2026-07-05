# -*- coding: utf-8 -*-
import csv
import io
import os
import posixpath
import shutil
import sys
import zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "_整理")
DATA_DIR = os.path.join(OUT, "数据")
IMG_DIR = os.path.join(OUT, "图")
MANIFEST = os.path.join(OUT, "整理清单.csv")

IMG_EXTS = {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff",
            ".svg", ".webp", ".pdf"}
ZIP_EXTS = {".zip"}

DRY_RUN = "--apply" not in sys.argv

def classify(name):
    ext = os.path.splitext(name)[1].lower()
    return "img" if ext in IMG_EXTS else "data"

def safe_name(dest_dir, filename):
    """Return a non-colliding path inside dest_dir."""
    base, ext = os.path.splitext(filename)
    candidate = filename
    i = 1
    while os.path.exists(os.path.join(dest_dir, candidate)):
        candidate = f"{base}_{i}{ext}"
        i += 1
    return os.path.join(dest_dir, candidate)

def decode_zipname(info):
    """ZIP entries are often cp437 or GBK encoded for Chinese filenames."""
    raw = info.filename
    if info.flag_bits & 0x800:  # UTF-8 flag set
        return raw
    try:
        data = raw.encode("cp437")
    except UnicodeEncodeError:
        return raw
    for encoding in ("utf-8", "gbk"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            pass
    return raw

def is_junk(name):
    normalized = name.replace("\\", "/")
    base_name = posixpath.basename(normalized)
    if not base_name:
        return True
    lower = base_name.lower()
    return (
        base_name.startswith("._")
        or lower in {".ds_store", ".rhistory"}
        or "__macosx/" in normalized.lower()
    )

def iter_zip_files(zf, source_label):
    for info in zf.infolist():
        if info.is_dir():
            continue
        name = decode_zipname(info)
        if is_junk(name):
            continue

        base_name = posixpath.basename(name.replace("\\", "/"))
        ext = os.path.splitext(base_name)[1].lower()
        nested_source = f"{source_label}/{name}"

        with zf.open(info) as src:
            if ext in ZIP_EXTS:
                blob = src.read()
                try:
                    with zipfile.ZipFile(io.BytesIO(blob)) as nested:
                        yield from iter_zip_files(nested, nested_source)
                except zipfile.BadZipFile:
                    yield "data", base_name, nested_source, io.BytesIO(blob)
            else:
                blob = src.read()
                yield classify(base_name), base_name, nested_source, io.BytesIO(blob)

def main():
    zips = [f for f in os.listdir(BASE) if f.lower().endswith(".zip")]
    zips.sort()
    print(f"Found {len(zips)} zip files. DRY_RUN={DRY_RUN}\n")

    if not DRY_RUN:
        os.makedirs(DATA_DIR, exist_ok=True)
        os.makedirs(IMG_DIR, exist_ok=True)

    counts = {"data": 0, "img": 0}
    samples = {"data": [], "img": []}
    manifest_rows = []

    for z in zips:
        zpath = os.path.join(BASE, z)
        print(f"== {z} ==")
        try:
            with zipfile.ZipFile(zpath) as zf:
                for kind, base_name, source_path, stream in iter_zip_files(zf, z):
                    counts[kind] += 1
                    dest_dir = IMG_DIR if kind == "img" else DATA_DIR
                    if len(samples[kind]) < 8:
                        samples[kind].append(base_name)
                    dest_name = base_name
                    if not DRY_RUN:
                        target = safe_name(dest_dir, base_name)
                        dest_name = os.path.basename(target)
                        with open(target, "wb") as out:
                            shutil.copyfileobj(stream, out)
                    manifest_rows.append({
                        "category": "图" if kind == "img" else "数据",
                        "output_name": dest_name,
                        "source_path": source_path,
                    })
        except zipfile.BadZipFile:
            print(f"   !! Could not read (bad/corrupt zip): {z}")
        print()

    print("==== SUMMARY ====")
    print(f"Data files: {counts['data']}")
    print(f"Image/PDF files: {counts['img']}")
    print("\nSample data files:", samples["data"])
    print("Sample image files:", samples["img"])
    if DRY_RUN:
        print("\n(DRY RUN — nothing written. Re-run with --apply to extract.)")
    else:
        with open(MANIFEST, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=["category", "output_name", "source_path"])
            writer.writeheader()
            writer.writerows(manifest_rows)
        print(f"\nDone. Output in: {OUT}")

if __name__ == "__main__":
    main()
