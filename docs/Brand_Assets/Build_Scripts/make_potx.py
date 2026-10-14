import sys, os, zipfile, shutil

def make_potx(pptx_path, potx_path):
    workdir = pptx_path + "_unpack_potx"
    if os.path.exists(workdir):
        shutil.rmtree(workdir)
    os.makedirs(workdir)
    with zipfile.ZipFile(pptx_path) as z:
        z.extractall(workdir)

    # [Content_Types].xml: main presentation part becomes a template part
    ct_path = os.path.join(workdir, "[Content_Types].xml")
    with open(ct_path, "r", encoding="utf-8") as f:
        ct = f.read()
    ct2 = ct.replace(
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml",
        "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml",
    )
    assert ct2 != ct, "content type override not found/replaced"
    with open(ct_path, "w", encoding="utf-8") as f:
        f.write(ct2)

    if os.path.exists(potx_path):
        os.remove(potx_path)
    with zipfile.ZipFile(potx_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(workdir):
            for file in files:
                full = os.path.join(root, file)
                arcname = os.path.relpath(full, workdir)
                zf.write(full, arcname)

    shutil.rmtree(workdir)
    print("wrote", potx_path)

if __name__ == "__main__":
    make_potx(sys.argv[1], sys.argv[2])
