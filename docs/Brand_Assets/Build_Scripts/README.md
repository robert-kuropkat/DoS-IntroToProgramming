# Brand_Assets / Build_Scripts

This folder holds the source scripts that generate every branded Word
document and PowerPoint deck in the course, plus the small utilities that
patch known bugs in the libraries they depend on. Regenerating a document
means editing the relevant script and re-running it - not hand-editing the
.docx/.pptx in Word/PowerPoint (those are the *output*, not the source).

Only the final, correct version of each script is kept here. Earlier
draft versions (v2, v3, "unbranded", etc.) lived in Claude's temporary
working scratchpad and are gone - this folder is the only copy from now
on, so treat it as the master source.

## One-time setup to regenerate anything

You need Node.js and Python 3. From a terminal, in a working folder
containing the script(s) you want to run:

```
npm install docx pptxgenjs
pip install python-docx python-pptx cairosvg openpyxl --break-system-packages
```

(LibreOffice, run headless as `soffice --headless --convert-to pdf file.docx`,
is what was used throughout to render a PDF for visual proofing before
shipping anything - not required to generate the .docx/.pptx itself, only
to preview it.)

Versions last verified against: docx 9.7.1, pptxgenjs 3.x, python-docx 1.2.0,
python-pptx 1.0.2, cairosvg 2.9.0, openpyxl 3.1.5.

## The working-directory convention

Every script below uses **relative** paths for the images it embeds
(`"vgc_crest.png"`, `"icons/book.png"`, `"bg_title.png"`, etc.) - this
matches how they were originally run, all together in one flat folder.
To regenerate:

1. Copy the script (and `brand.js`, always required) into an empty working
   folder.
2. Copy in whatever image assets that specific script needs (see the list
   below) from the sibling `Logos_and_Crests/`, `Icons/`, `Slide_Backgrounds/`,
   and `Diagrams/` folders. Scripts that reference `icons/xxx.png` need an
   `icons` subfolder in that same working folder, not the images loose.
3. Run `node the_script.js`. It writes a .docx or .pptx into that same
   folder.
4. For slide decks only: run `python3 fix_content_types.py YourDeck.pptx`
   afterward - pptxgenjs's own output has a few known defects (a phantom
   Content_Types entry, a malformed notes-master, junk empty folders) that
   PowerPoint's strict parser can choke on. This script strips them.
   Skipping this step means the .pptx may still open fine in LibreOffice
   but risk a "repair" prompt in real PowerPoint.
5. For the two template files only: after generating the base .docx/.pptx,
   run `make_dotx.py in.docx out.dotx` or `make_potx.py in.pptx out.potx`
   to convert it into a true template file.

## Script -> output -> required assets

| Script | Produces | Needs (beyond brand.js) |
|---|---|---|
| `build_course_syllabus.js` | Course_Syllabus.docx | crest only (via brand.js) |
| `build_engine_and_platform_guide.js` | Engine_and_Platform_Guide.docx | crest only |
| `build_engine_installation_guide.js` | Engine_Installation_and_Hello_World_Guide.docx | crest only, plus the engine-specific annotated screenshots already embedded in the shipped doc (those source screenshots were one-time captures and were not preserved separately - see note below) |
| `build_engine_showcase.js` | Engine_Showcase.docx | crest, plus `showcase_<engine>_{setup,code,result}.png` per engine (24 images, not preserved separately - same note as above) |
| `build_game_design_document.js` | Game_Design_Document_Template.docx | `vgc_crest.png`, `player_arrow_white.png`, `player_walk_spritesheet.png` |
| `build_gpu_troubleshooting.js` | Older_GPU_DirectX_Troubleshooting.docx | crest only |
| `build_lesson01_gdd.js` | Lesson_01_GDD.docx | `vgc_crest.png` only - a condensed, non-fill-in "sneak peek" version of the GDD template, handed out at the end of Lesson 1 |
| `build_lesson01_plan.js` | Lesson_01_Plan.docx | crest only |
| `build_lesson01_slides_deck.js` | Lesson_01_Slides.pptx | `bg_title.png`, `bg_content.png`, `bg_title_with_vgc.png`, `bg_title_with_slhe.png`, `icons/` (check, production, chat, book, + per-engine icons) |
| `build_lesson02_gdd.js` | Lesson_02_GDD.docx | `vgc_crest.png` - full GDD template plus a "Lesson 2 Goals" tri-state checkbox table (Plan to Implement / Implemented / Not Planned) pre-filled with that lesson's 7 Level 1/2/3 goals |
| `build_lesson02_handouts_keys_checklist.js` | Lesson_02_Prep_Checklist.docx + 8 Handouts + 8 Keys | crest only |
| `build_lesson02_plan_teacher_script.js` | Lesson_02_Plan.docx + Lesson_02_Teacher_Script.docx | crest, `diagram_game_loop.png`, `diagram_sequences_events.png`, `diagram_activity_keypress.png`, `diagram_state_player.png` |
| `build_lesson02_slides_deck.js` | Lesson_02_Slides.pptx | same backgrounds as Lesson 1 deck, `icons/` (check, production, wrench, chat, book, programming, controller), all 4 diagram PNGs |
| `build_lesson02_superstretch_ideas.js` | Lesson_02_SuperStretch_Ideas.docx | crest only |
| `build_lesson_production_standards.js` | Lesson_Production_Standards.docx | crest only |
| `build_powerpoint_template.js` | SLHE_VGC_PowerPoint_Template.pptx (then run `make_potx.py`) | `bg_title.png`, `bg_content.png`, `bg_title_with_vgc.png`, `bg_title_with_slhe.png` |
| `build_semester01_capstone_gdd.js` | Semester_01_Capstone_GDD.docx | `vgc_crest.png` only - the instructor-facing master design for the Semester 1 compulsory mini-game (top-down shooter): High Concept, full Level 1/2/3 feature breakdown, and a Feature-to-Lesson Map. Lives at Lessons/ top level, not inside one lesson folder, since it's referenced across Lessons 2-15. Not distributed to students. |
| `build_student_resource_list.js` | Student_Resource_List.docx | crest only |
| `build_template_howto_guide.js` | SLHE_VGC_Template_HowTo_Guide.docx | crest only |
| `build_word_template.js` | SLHE_VGC_Word_Template.docx (then run `make_dotx.py`) | crest only |
| `brand.js` | *(shared module, not run directly)* | `vgc_crest.png` must be present alongside it - every other script requires this file |

Note on screenshots: the Engine Installation Guide and Engine Showcase
documents embed real annotated screenshots of each engine's editor. The
`annotate_*.py` scripts that produced those (one per engine, referenced
in the course's build history) are a good reference for the *technique*
- crop coordinates, callout style - but their raw input screenshots were
one-time captures of each engine's UI and were not archived as reusable
assets. Regenerating those two documents from scratch would mean
re-capturing fresh screenshots, not just re-running a script.

## What's NOT here

- The 30-Lesson Master Schedule (xlsx) was built/edited directly with
  openpyxl rather than through a reusable branded template - there's no
  equivalent "just re-run this" script for it.
- Anything Lesson 3 onward - this library only covers what's been built
  through Lesson 2 plus the shared templates. Extend it as new lessons
  are produced, following the same one-script-per-document pattern.

## Where the visual rules live

This README is about the *code*. For the actual design rules (exact
colors, when to use which gold, logo clear-space, etc.) see
`Brand_Style_Guide.docx` in `Lessons/Lesson 00 - Standards/`, alongside
the Word/PowerPoint templates.

## PDF rendering note: the Consolas font

Code boxes use the Consolas font, which Word has but a plain Linux/LibreOffice
render does not. Without it, LibreOffice silently substitutes a serif font, and
underscores in names like `_ready()` and `on_enemy_escaped` get clipped and look
like spaces in the PDF. Before converting any .docx to PDF with soffice, install
a monospace font under the family name "Consolas" (for example, copy Liberation
Mono and rename its family to Consolas with fontTools, save it in ~/.fonts, run
fc-cache -f), then check the PDF with `pdffonts` and a rendered page. The .docx
files themselves were never affected - only PDFs rendered without the font.
