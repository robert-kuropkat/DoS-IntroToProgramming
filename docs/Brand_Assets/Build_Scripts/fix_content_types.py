import sys, os, re, zipfile, shutil

def fix(pptx_path):
    workdir = pptx_path + "_unpacked"
    if os.path.exists(workdir):
        shutil.rmtree(workdir)
    os.makedirs(workdir)
    with zipfile.ZipFile(pptx_path) as z:
        names = z.namelist()
        z.extractall(workdir)

    ct_path = os.path.join(workdir, "[Content_Types].xml")
    with open(ct_path, "r", encoding="utf-8") as f:
        ct = f.read()

    # Which slideMaster / slideLayout / slide files actually exist in the package?
    actual_masters = set(n for n in names if re.match(r"ppt/slideMasters/slideMaster\d+\.xml$", n))
    actual_layouts = set(n for n in names if re.match(r"ppt/slideLayouts/slideLayout\d+\.xml$", n))

    removed = []

    def strip_orphans(xml, actual_set, label):
        nonlocal removed
        def repl(m):
            part = m.group(1)  # e.g. /ppt/slideMasters/slideMaster7.xml
            rel = part.lstrip("/")
            if rel in actual_set:
                return m.group(0)
            removed.append(part)
            return ""
        pattern = re.compile(r'<Override PartName="(/ppt/(?:slideMasters/slideMaster|slideLayouts/slideLayout)\d+\.xml)"[^>]*?/>')
        return pattern.sub(repl, xml)

    new_ct = strip_orphans(ct, actual_masters | actual_layouts, "master/layout")

    with open(ct_path, "w", encoding="utf-8") as f:
        f.write(new_ct)

    # --- Fix: PptxGenJS 4.0.1 always emits a notesMaster with 6 malformed
    # placeholder shapes (Header/Date/SlideImage/Body/Footer/SlideNumber)
    # that PowerPoint's strict parser rejects on EVERY file it produces,
    # regardless of content -- see gitbrent/PptxGenJS issue #1443.
    # We don't use speaker notes, so replace the spTree with the minimal
    # empty structure PowerPoint itself produces after repairing the file.
    nm_path = os.path.join(workdir, "ppt", "notesMasters", "notesMaster1.xml")
    if os.path.exists(nm_path):
        with open(nm_path, "r", encoding="utf-8") as f:
            nm = f.read()
        empty_sptree = (
            '<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
            '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
            '</p:spTree>'
        )
        new_nm, n = re.subn(r'<p:spTree>.*?</p:spTree>', empty_sptree, nm, flags=re.S)
        if n:
            with open(nm_path, "w", encoding="utf-8") as f:
                f.write(new_nm)
            print(f"Patched notesMaster1.xml: removed {n} malformed spTree block(s) with 6 placeholder shapes")
        else:
            print("WARNING: could not find spTree in notesMaster1.xml to patch")

    # --- Cleanup: drop empty scaffold dirs pptxgenjs always creates
    # (ppt/charts, ppt/embeddings) even when unused. Low severity but
    # unnecessary and flagged as a contributing factor in combination
    # with the other defects (gitbrent/PptxGenJS #1449).
    for junk in ("ppt/charts", "ppt/embeddings"):
        jp = os.path.join(workdir, *junk.split("/"))
        if os.path.exists(jp) and not any(
            os.path.isfile(os.path.join(jp, f)) and not f.startswith("_rels")
            for f in os.listdir(jp)
        ):
            shutil.rmtree(jp)
            print("Removed empty scaffold dir:", junk)

    out_path = pptx_path
    if os.path.exists(out_path):
        os.remove(out_path)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(workdir):
            for file in files:
                full = os.path.join(root, file)
                arcname = os.path.relpath(full, workdir)
                zf.write(full, arcname)

    shutil.rmtree(workdir)
    print(f"Removed {len(removed)} orphaned Content_Types overrides:")
    for r in removed:
        print("  ", r)
    print("Rewrote", out_path)

if __name__ == "__main__":
    fix(sys.argv[1])
