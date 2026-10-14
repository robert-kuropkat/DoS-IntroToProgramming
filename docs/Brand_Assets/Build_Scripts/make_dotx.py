import sys, os, re, zipfile, shutil

def make_dotx(docx_path, dotx_path):
    workdir = docx_path + "_unpack_dotx"
    if os.path.exists(workdir):
        shutil.rmtree(workdir)
    os.makedirs(workdir)
    with zipfile.ZipFile(docx_path) as z:
        z.extractall(workdir)

    # 1. [Content_Types].xml: main document part becomes a template part
    ct_path = os.path.join(workdir, "[Content_Types].xml")
    with open(ct_path, "r", encoding="utf-8") as f:
        ct = f.read()
    ct2 = ct.replace(
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
    )
    assert ct2 != ct, "content type override not found/replaced"
    with open(ct_path, "w", encoding="utf-8") as f:
        f.write(ct2)

    # 2. word/document.xml itself is unchanged - templates use the same
    #    document.xml schema, just a different content type + extension.

    if os.path.exists(dotx_path):
        os.remove(dotx_path)
    with zipfile.ZipFile(dotx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(workdir):
            for file in files:
                full = os.path.join(root, file)
                arcname = os.path.relpath(full, workdir)
                zf.write(full, arcname)

    shutil.rmtree(workdir)
    print("wrote", dotx_path)

if __name__ == "__main__":
    make_dotx(sys.argv[1], sys.argv[2])
