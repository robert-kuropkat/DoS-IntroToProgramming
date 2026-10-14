import cairosvg, os

os.makedirs("icons", exist_ok=True)

NAVY = "#16305c"
GOLD = "#c99a2e"
GOLD_L = "#e8c25a"

def badge(glyph, name, ring=GOLD, fill=NAVY, glyph_color="#f4f0e6"):
    svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200" width="200" height="200">
      <circle cx="100" cy="100" r="94" fill="{fill}" stroke="{ring}" stroke-width="7"/>
      <g fill="{glyph_color}" stroke="{glyph_color}">
        {glyph}
      </g>
    </svg>'''
    cairosvg.svg2png(bytestring=svg.encode(), write_to=f"icons/{name}.png", output_width=200, output_height=200)
    print("made", name)

# Shield (internet safety)
badge('''<path d="M100,45 L145,60 L145,100 Q145,140 100,158 Q55,140 55,100 L55,60 Z"
          fill="none" stroke-width="8" stroke-linejoin="round"/>
       <path d="M82,100 L95,113 L120,85" fill="none" stroke-width="8" stroke-linecap="round" stroke-linejoin="round"/>''',
      "shield", fill=NAVY, glyph_color=GOLD_L)

# Design (lightbulb / pencil)
badge('''<path d="M100,50 Q130,50 130,82 Q130,102 115,112 L115,128 L85,128 L85,112 Q70,102 70,82 Q70,50 100,50 Z"
          fill="none" stroke-width="7"/>
       <line x1="88" y1="140" x2="112" y2="140" stroke-width="7" stroke-linecap="round"/>
       <line x1="90" y1="150" x2="110" y2="150" stroke-width="7" stroke-linecap="round"/>''',
      "design", fill=NAVY, glyph_color=GOLD_L)

# Programming (code brackets)
badge('''<path d="M78,70 L48,100 L78,130" fill="none" stroke-width="9" stroke-linecap="round" stroke-linejoin="round"/>
       <path d="M122,70 L152,100 L122,130" fill="none" stroke-width="9" stroke-linecap="round" stroke-linejoin="round"/>
       <line x1="112" y1="60" x2="88" y2="140" stroke-width="7" stroke-linecap="round"/>''',
      "programming", fill=NAVY, glyph_color=GOLD_L)

# Visual Art (palette)
badge('''<path d="M100,55 Q145,55 145,95 Q145,115 125,115 Q118,115 118,108 Q118,100 126,100 Q133,100 133,93
          Q133,68 100,68 Q65,68 65,100 Q65,132 100,132 Q112,132 120,127"
          fill="none" stroke-width="7" stroke-linecap="round"/>
       <circle cx="85" cy="88" r="7" fill="{0}" stroke="none"/>
       <circle cx="82" cy="108" r="7" fill="{0}" stroke="none"/>
       <circle cx="100" cy="115" r="7" fill="{0}" stroke="none"/>
       <circle cx="115" cy="80" r="7" fill="{0}" stroke="none"/>'''.replace("{0}", GOLD_L),
      "art", fill=NAVY, glyph_color=GOLD_L)

# Audio (music note)
badge('''<circle cx="78" cy="135" r="14" fill="{0}" stroke="none"/>
       <circle cx="128" cy="122" r="14" fill="{0}" stroke="none"/>
       <line x1="92" y1="135" x2="92" y2="65" stroke-width="7"/>
       <line x1="142" y1="122" x2="142" y2="55" stroke-width="7"/>
       <line x1="92" y1="65" x2="142" y2="55" stroke-width="7"/>'''.replace("{0}", GOLD_L),
      "audio", fill=NAVY, glyph_color=GOLD_L)

# Production (clipboard/checklist)
badge('''<rect x="65" y="55" width="70" height="95" rx="8" fill="none" stroke-width="7"/>
       <rect x="85" y="48" width="30" height="16" rx="4" fill="{0}" stroke="none"/>
       <line x1="78" y1="85" x2="122" y2="85" stroke-width="6" stroke-linecap="round"/>
       <line x1="78" y1="105" x2="122" y2="105" stroke-width="6" stroke-linecap="round"/>
       <line x1="78" y1="125" x2="105" y2="125" stroke-width="6" stroke-linecap="round"/>'''.replace("{0}", GOLD_L),
      "production", fill=NAVY, glyph_color=GOLD_L)

# Wrench (lab/setup)
badge('''<path d="M120,55 Q140,55 140,75 Q140,85 132,90 L92,130 Q85,137 78,130 Q71,123 78,116 L118,76
          Q112,68 120,55 Z" fill="none" stroke-width="7" stroke-linejoin="round"/>
       <circle cx="70" cy="138" r="6" fill="{0}" stroke="none"/>'''.replace("{0}", GOLD_L),
      "wrench", fill=NAVY, glyph_color=GOLD_L)

# Book (homework)
badge('''<path d="M100,65 Q80,52 55,58 L55,128 Q80,122 100,135 Q120,122 145,128 L145,58 Q120,52 100,65 Z"
          fill="none" stroke-width="7" stroke-linejoin="round"/>
       <line x1="100" y1="65" x2="100" y2="135" stroke-width="6"/>''',
      "book", fill=NAVY, glyph_color=GOLD_L)

# Controller (small, reuse VGC style) - gold badge navy glyph variant
badge('''<path d="M62,88 Q62,68 82,68 L118,68 Q138,68 138,88 L138,108
          Q138,126 124,126 Q115,126 111,116 L109,111 Q106,105 98,105 L82,105
          Q74,105 71,111 L69,116 Q65,126 56,126 Q42,126 42,108 Z"
          fill="{0}" stroke="none"/>
       <rect x="76" y="82" width="8" height="20" rx="2" fill="{1}" stroke="none"/>
       <rect x="68" y="90" width="24" height="8" rx="2" fill="{1}" stroke="none"/>
       <circle cx="118" cy="82" r="5" fill="{1}" stroke="none"/>
       <circle cx="128" cy="90" r="5" fill="{1}" stroke="none"/>
       <circle cx="118" cy="98" r="5" fill="{1}" stroke="none"/>
       <circle cx="108" cy="90" r="5" fill="{1}" stroke="none"/>'''.format(GOLD_L, NAVY),
      "controller", fill=NAVY, ring=GOLD)

# Chat bubble (discussion points)
badge('''<path d="M55,60 L145,60 Q152,60 152,67 L152,110 Q152,117 145,117 L95,117 L75,138 L75,117 L55,117
          Q48,117 48,110 L48,67 Q48,60 55,60 Z" fill="none" stroke-width="7" stroke-linejoin="round"/>
       <circle cx="78" cy="88" r="5" fill="{0}" stroke="none"/>
       <circle cx="100" cy="88" r="5" fill="{0}" stroke="none"/>
       <circle cx="122" cy="88" r="5" fill="{0}" stroke="none"/>'''.replace("{0}", GOLD_L),
      "chat", fill=NAVY, glyph_color=GOLD_L)

# Checkmark (exit check / goal)
badge('''<path d="M62,102 L88,128 L140,72" fill="none" stroke-width="12" stroke-linecap="round" stroke-linejoin="round"/>''',
      "check", fill=NAVY, glyph_color=GOLD_L)

print("ALL DONE")
