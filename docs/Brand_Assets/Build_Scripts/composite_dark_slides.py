from PIL import Image

DPI = 96  # 1280px / 13.333in

def paste(base_path, overlay_path, x_in, y_in, w_in, h_in, out_path):
    base = Image.open(base_path).convert("RGBA")
    overlay = Image.open(overlay_path).convert("RGBA")
    w_px = round(w_in * DPI)
    h_px = round(h_in * DPI)
    x_px = round(x_in * DPI)
    y_px = round(y_in * DPI)
    overlay_r = overlay.resize((w_px, h_px), Image.LANCZOS)
    base.alpha_composite(overlay_r, (x_px, y_px))
    base.convert("RGB").save(out_path)
    print("wrote", out_path)

SW = 13.333

# Title slide: vgc_crest centered
paste("bg_title.png", "vgc_crest.png", (SW - 2.05) / 2, 1.0, 2.05, 2.54, "bg_title_with_vgc.png")

# Closing slide: slhe_crest centered (smaller, higher up)
paste("bg_title.png", "slhe_crest.png", (SW - 1.45) / 2, 0.7, 1.45, 1.8, "bg_title_with_slhe.png")

# Section dividers: icon badge at fixed position
for icon in ["book", "shield", "design", "controller", "wrench"]:
    paste("bg_title.png", f"icons/{icon}.png", 1.0, 2.75, 1.7, 1.7, f"bg_title_{icon}.png")

print("ALL DONE")
