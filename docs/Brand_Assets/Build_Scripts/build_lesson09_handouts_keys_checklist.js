const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign, ImageRun
} = require("docx");
const brand = require("./brand.js");

const CONTENT_WIDTH = 9360;
const border = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function link(text, url) { return new ExternalHyperlink({ children: [new TextRun({ text, style: "Hyperlink" })], link: url }); }
function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function numbered(text, level = 0) { return new Paragraph({ numbering: { reference: "numbers", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function checkbox(text) { return new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 80 }, children: [new TextRun(text)] }); }
function labelBullet(label, text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)] });
}
function code(lines, legend) {
  // Lines starting with "@@ " are NEW this week and get a yellow highlight (marker stripped).
  const children = [];
  if (lines.some(l => l.startsWith("@@ "))) {
    children.push(new TextRun({ text: legend || "Highlighted lines are new this week; everything else is code you already have.", italics: true, size: 18, color: "555555" }));
    children.push(new TextRun({ text: "", break: 1 }));
  }
  lines.forEach((line, i) => {
    const isNew = line.startsWith("@@ ");
    const opts = { text: isNew ? line.slice(3) : line, font: "Consolas", size: 20 };
    if (isNew) opts.shading = { type: ShadingType.CLEAR, fill: "FFE98A" };
    children.push(new TextRun(opts));
    if (i < lines.length - 1) children.push(new TextRun({ text: "", break: 1 }));
  });
  return new Paragraph({
    spacing: { after: 160, before: 80 }, indent: { left: 360 },
    shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
    children
  });
}
function diagram(file, widthIn, heightIn, caption) {
  const data = fs.readFileSync(file);
  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: caption ? 40 : 160 },
      children: [new ImageRun({ data, transformation: { width: widthIn * 96, height: heightIn * 96 }, type: "png" })],
    }),
  ];
  if (caption) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 },
      children: [new TextRun({ text: caption, italics: true, size: 16, color: "888888" })],
    }));
  }
  return children;
}
function headerCell(text, width) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: brand.NAVY2, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 18 })] })] });
}
function bodyCell(children, width, opts = {}) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.TOP,
    children: Array.isArray(children) ? children : [new Paragraph({ children: [new TextRun({ text: children, size: 18 })] })] });
}

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "checks", levels: [{ level: 0, format: LevelFormat.BULLET, text: "☐", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]
};

const stylesBlock = {
  default: { document: { run: { font: "Arial", size: 22 } } },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 30, bold: true, font: "Arial", color: brand.NAVY },
      paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 24, bold: true, font: "Arial", color: brand.NAVY2 },
      paragraph: { spacing: { before: 220, after: 120 }, outlineLevel: 1 } },
    { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 22, bold: true, font: "Arial", color: "444444" },
      paragraph: { spacing: { before: 140, after: 80 }, outlineLevel: 2 } },
  ]
};

function pageHeader(text) {
  return new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text, size: 16, color: "888888" })] })] });
}
function pageFooter() {
  return new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 16 }), new TextRun({ children: [PageNumber.CURRENT], size: 16 })] })] });
}

const jobs = [];

// =========================================================================
// PREP CHECKLIST
// =========================================================================
{
  const doc = new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 9 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 9 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("This Week: The Payoff of Deferring UI"),
        p("Since Lesson 6, every engine's win/lose messages have deliberately used the simplest possible one-off text tool, with a promise that \"the real UI system comes later, once there's a score and lives to actually display.\" This week is that payoff. Some engines (Godot, GameMaker, GDevelop, TIC-80, Scratch) already had the right tool the whole time - their Lesson 6/8 text mechanism just needs to run continuously instead of once. Two engines (Unity, Unreal) are genuinely introducing a new system this week: Unity's UI Text + Canvas, and Unreal's UMG Widget Blueprint. 001 Game Creator also introduces something new: an Interface with bound Fields, replacing one-off Message Boxes for this purpose."),
        checkbox("Read through each engine's Handout once before class - Unity, Unreal, and 001 Game Creator involve genuinely new UI concepts this week, more setup than a typical lesson."),

        h2("001 Game Creator: Verified Mechanism"),
        p("001 Game Creator's HUD approach (an Interface containing Text/Graphic Fields, bound to a Variable via \"Retrieve Text/Value\" with \"Update Real-time\" ticked) was verified directly against the official wiki (001gamecreator.com/wiki/field/) before this lesson was written. The most important checkbox to not miss: Update Real-time. Without it, the Interface only refreshes when reopened, which looks like a bug but is really just a missed setting."),
        checkbox("If you have Windows access before class, walk through creating one Interface with one bound Field yourself - this is new territory even for an instructor comfortable with 001 Game Creator's trigger-based patterns."),

        h2("Two Update Patterns Worth Naming"),
        p("This lesson quietly introduces two different, both-correct ways to keep a display in sync with data: PUSH (call an update function every time the data changes - used by Godot, GDevelop, 001 Game Creator's real-time binding, Unreal) and POLL (just re-check the data every single frame, no matter what - used naturally by GameMaker's Draw GUI event, TIC-80's print() calls, and this week's Unity HUD script). Neither is \"more correct\" - it's worth drawing this contrast out loud if a strong student asks why the approaches differ across engines."),

        h2("Self-Contained Lesson Assets"),
        p("No new assets this week - the HUD is built from each engine's own text/UI tools."),

        h2("Continuity From Lesson 8"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 8, not create a new one."),
        checkbox("If a student's Lesson 8 project is lost or broken, have a pre-built \"Lesson 8 complete\" starter available per engine (from the Lesson 8 Teacher's Keys) so they aren't blocked."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop's Lesson 6 and 8 exercises still work for all 8 engines."),
        checkbox("Unity: confirm GameObject > UI (Canvas) > Legacy > Text is available and creates a Canvas automatically - some Unity versions nudge toward TextMeshPro first; Legacy Text is intentionally used here to avoid an extra package-import step."),
        checkbox("Unreal: confirm Widget Blueprint creation and \"Create Widget\" / \"Add to Viewport\" nodes work smoothly on the demo laptop - this is the most setup-heavy engine this week."),
        checkbox("001 Game Creator (Windows demo laptop): walk through Interface creation once, and specifically confirm the Update Real-time checkbox's effect is visible (toggle it off briefly to show students what \"not real-time\" looks like, if time allows)."),
        checkbox("Godot / GameMaker / GDevelop / TIC-80 / Scratch: confirm score increments by exactly 10 per kill and the on-screen display updates immediately, every time, with no stale numbers."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies."),

        h2("Known Gotchas This Week"),
        bullet("001 Game Creator: forgetting Update Real-time is the single most likely mistake - the Interface will look broken (frozen numbers) even though the underlying Variable is updating correctly."),
        bullet("Unity / Unreal: students may accidentally create multiple Canvas or Widget instances by repeating setup steps - if a HUD element appears duplicated or doubled, check for this first."),
        bullet("Every engine: a score readout stuck at 0 usually means the increment was added to the wrong code path (or forgotten entirely), not a display bug - have students confirm the underlying variable is changing (via a temporary debug print/log if needed) before troubleshooting the display itself."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_09_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 9: On the Scoreboard", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: track a real score, and show it - along with your lives - on screen, all the time, not just at the end.", { bold: true }),
    p("Open your own [YourName]Game project from Lesson 8 - don't start a new one.", { bold: true }),
    p("Tip: no new assets this week - this is a logic and UI lesson.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("Destroying an enemy adds 10 points to your score. Score and lives are both visible on screen, all the time, not just at the start or end of the game.", { italics: true }),
    ...diagram("diagram_activity_hud_sync.png", 6.3, 4.375, "Score joins the hit path, then Push or Poll takes over - the same shape in every engine, whatever the exact code or blocks look like."),
    ...data.level1.map(s => numbered(s)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  if (data.level1EngineDiagram) children.push(...diagram(data.level1EngineDiagram.file, data.level1EngineDiagram.w, data.level1EngineDiagram.h, data.level1EngineDiagram.caption));
  if (data.altPath) {
    children.push(h2("Alternate Path: " + data.altPath.title));
    children.push(p(data.altPath.intro, { italics: true, color: "555555" }));
    children.push(...data.altPath.steps.map(s => numbered(s)));
    if (data.altPath.code) children.push(code(data.altPath.code));
    if (data.altPath.level1Diagram) children.push(...diagram(data.altPath.level1Diagram.file, data.altPath.level1Diagram.w, data.altPath.level1Diagram.h, data.altPath.level1Diagram.caption));
    if (data.altPath.note) children.push(p(data.altPath.note, { italics: true, color: "555555" }));
  }
  children.push(
    h2("Level 2 - Complete (Alpha/Beta)"),
    p("The score and lives readouts never show a stale or wrong number - they always match reality, immediately.", { italics: true }),
    ...data.level2.map(s => bullet(s)),
    h2("Level 3 - Refined (Ready to Publish) - pick any that interest you"),
    ...data.level3.map(s => bullet(s)),
    h2("If Something's Wrong"),
    ...data.trouble.map(s => bullet(s)),
    h2("Want More?"),
    p(data.resource),
  );

  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 9 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 8.",
      "Create a new variable score (for all sprites, not \"for this sprite only\").",
    ],
    level1: [
      "On GameManager's existing \"when I receive [enemy hit]\" script - the same script that already decrements enemies_remaining - add \"change score by 10\". The Enemy sprite itself doesn't change: it still only ever broadcasts enemy hit and never touches score directly, the same report-only role it's had since Lesson 6.",
      "In the Variables palette, check the checkbox next to score AND the checkbox next to lives so both display automatically on stage - Scratch's simplest built-in way to show a live number, no coding required.",
      "Click the stage, shoot a few enemies, and let one escape - watch score go up and lives go down, both immediately, right on stage.",
    ],
    level1Code: [
      "// GameManager (hidden sprite) - existing script:",
      "when I receive [enemy hit]",
      "change enemies_remaining by -1",
      "@@ change score by 10",
      "if enemies_remaining = 0 then",
      "    broadcast [you win]",
      "",
      "// Enemy sprite: no changes - it still only broadcasts [enemy hit] and never touches score.",
      "",
      "// Variables palette: tick the checkbox beside score AND beside lives so both show on stage.",
    ],
    level2: [
      "Shoot 3 enemies and confirm score reads exactly 30 - not 20, not 40.",
      "Let an enemy escape and confirm score does NOT change - only lives should move.",
      "The on-screen numbers can never be \"stale\" here, since Scratch's variable watchers always show the live value automatically - there's nothing extra to keep in sync.",
    ],
    level3: [
      "Visual: right-click the score or lives watcher on stage and choose \"large readout\" for a bigger, more game-like display; try dragging both watchers to a corner that doesn't overlap your sprites.",
      "Visual: change points-per-kill to something more interesting, like a random value between 5 and 15 (using \"pick random\"), for a bit of scoring variety.",
      "Structure: keep the score change inside the same \"when I receive [enemy hit]\" script as the enemies_remaining change, on GameManager - both are direct consequences of the exact same confirmed-hit event, so they belong together, on the one object that already owns all of this game's shared state.",
    ],
    trouble: [
      "Score doesn't appear on stage at all: confirm the checkbox next to the score variable (in the Variables palette) is actually checked.",
      "Score changes when an enemy escapes: the \"change score by 10\" block was added to GameManager's wrong script - move it into \"when I receive [enemy hit]\", not \"when I receive [enemy escaped]\".",
      "Score jumps by more than 10 per kill: check for a duplicate \"change score by 10\" block on GameManager, possibly left over from testing.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 8.",
      "Add two Label nodes to your main scene: ScoreText and LivesText, positioned in a top corner where they won't overlap gameplay. Leave both visible (unlike WinText/LoseText, these stay on screen the whole game).",
    ],
    level1: [
      "In GameManager's script - the same node that already owns enemies_remaining, lives, on_enemy_destroyed(), and on_enemy_escaped() - add var score := 0.",
      "Write a new function update_hud() on GameManager's script that sets both Labels' text from score and lives, reaching them via get_node(\"../ScoreText\") and get_node(\"../LivesText\") - the same relative-path pattern GameManager already uses to reach WinText and LoseText (code below).",
      "Add a _ready() function to GameManager's script that calls update_hud() once, so the HUD shows correct starting values the instant the scene starts.",
      "In on_enemy_destroyed(), add score += 10 and call update_hud(). In on_enemy_escaped(), call update_hud() too, right after decrementing lives.",
      "Run the scene, shoot a few enemies, and let one escape - both labels should update immediately every time.",
    ],
    level1Code: [
      "# GameManager's script (complete):",
      "var enemies_remaining := 5",
      "var lives := 3",
      "@@ var score := 0",
      "",
      "@@ func _ready():",
      "@@     update_hud()",
      "",
      "@@ func update_hud():",
      "@@     get_node(\"../ScoreText\").text = \"Score: \" + str(score)",
      "@@     get_node(\"../LivesText\").text = \"Lives: \" + str(lives)",
      "",
      "func on_enemy_destroyed():",
      "    enemies_remaining -= 1",
      "@@     score += 10",
      "@@     update_hud()",
      "    if enemies_remaining <= 0:",
      "        get_node(\"../WinText\").show()",
      "",
      "func on_enemy_escaped():",
      "    lives -= 1",
      "@@     update_hud()",
      "    if lives <= 0:",
      "        get_node(\"../LoseText\").show()",
    ],
    level2: [
      "Shoot 3 enemies and confirm ScoreText reads exactly \"Score: 30\".",
      "Let an enemy escape and confirm ScoreText does NOT change - only LivesText should move.",
      "Confirm both labels show correct values the instant the scene starts, not just after the first kill or escape (this is why update_hud() is called once in GameManager's own _ready()).",
    ],
    level3: [
      "Visual: use the Inspector's Theme Overrides on each Label to increase font size or change color for a more game-like HUD look.",
      "Visual: change points-per-kill to a randomized value (e.g. randi_range(5, 15)) for scoring variety.",
      "Structure: notice update_hud() is now the ONLY code that touches ScoreText or LivesText directly - every other function just changes the underlying numbers and calls update_hud() to reflect them. This separation (game logic vs. display) is a new idea this week, distinct from the hit-vs-escape separation from Lessons 6-8 - and it still lives entirely on GameManager, the same dedicated manager node that's owned all of this game's shared state since Lesson 6.",
    ],
    trouble: [
      "Labels show \"Score: 0\" forever: confirm update_hud() is actually being called inside on_enemy_destroyed(), not just defined and forgotten.",
      "Error referencing ScoreText or LivesText: confirm the Label node names in the Scene tree exactly match (case-sensitive) what's referenced in GameManager's get_node() calls, and that the Labels are children of the main scene (siblings of GameManager), not children of GameManager itself.",
      "HUD doesn't show correct values at the very start: confirm update_hud() is called once inside GameManager's own _ready() function, not only inside the two handler functions.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Open your [YourName]Game project from Lesson 8.",
      "GameObject > UI (Canvas) > Legacy > Text, twice - rename them ScoreText and LivesText, and position them in a top corner (this automatically creates a Canvas if you don't have one yet).",
      "Create an empty GameObject named HUD. Create a new C# script named HUD, attach it to this GameObject.",
    ],
    level1: [
      "In HUD.cs, add public Text scoreText and public Text livesText fields, then an Update() that continuously sets their text from GameManager.score and GameManager.lives (code below).",
      "In the Inspector, drag your ScoreText and LivesText objects onto the HUD script's matching fields on the HUD GameObject.",
      "In GameManager.cs, add public static int score = 0, and add score += 10 inside RegisterKill().",
      "Press Play, shoot a few enemies, and let one escape - both texts should update immediately every time.",
    ],
    level1Code: [
      "// HUD.cs (new script, attached to the HUD GameObject):",
      "@@ using UnityEngine;",
      "@@ using UnityEngine.UI;",
      "@@ ",
      "@@ public class HUD : MonoBehaviour {",
      "@@     public Text scoreText;",
      "@@     public Text livesText;",
      "@@ ",
      "@@     void Update() {",
      "@@         scoreText.text = \"Score: \" + GameManager.score;",
      "@@         livesText.text = \"Lives: \" + GameManager.lives;",
      "@@     }",
      "@@ }",
      "",
      "// GameManager.cs - the new field and the revised RegisterKill():",
      "@@ public static int score = 0;",
      "",
      "public static void RegisterKill() {",
      "    enemiesRemaining--;",
      "@@     score += 10;",
      "    if (enemiesRemaining <= 0) hasWon = true;",
      "}",
    ],
    level2: [
      "Shoot 3 enemies and confirm ScoreText reads exactly \"Score: 30\".",
      "Let an enemy escape and confirm ScoreText does NOT change - only LivesText should move.",
      "Because HUD.cs's Update() re-reads the values every single frame, the display can never go stale - unlike a system that only updates on specific events, there's nothing to forget to refresh.",
    ],
    level3: [
      "Visual: adjust font size/color on the Text (Legacy) components in the Inspector for a more game-like HUD look.",
      "Visual: change points-per-kill to Random.Range(5, 16) for scoring variety.",
      "Structure: HUD.cs's Update() polls (checks every frame, no matter what), while GameManager's RegisterKill()/RegisterEscape() are event-driven (only run when something happens). Both are valid ways to keep a display in sync - discuss with your instructor when each style fits better.",
    ],
    trouble: [
      "NullReferenceException on Play: confirm you actually dragged ScoreText and LivesText onto the HUD script's fields in the Inspector - an unassigned public field is the most common cause.",
      "Score doesn't appear correctly: confirm public static int score is spelled and cased exactly the same everywhere it's referenced.",
      "Text is invisible: confirm the Text (Legacy) objects' Rect Transform positions are actually within the Canvas/screen bounds, and that their color isn't the same as the background.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same poll-every-frame HUD pattern works in a Script Graph instead of C#.",
      steps: [
        "On the HUD GameObject's graph, add an Update event node feeding two \"Format String\" nodes (String category) - one building \"Score: {0}\" with the Score graph variable as its argument, one building \"Lives: {0}\" with Lives - each feeding its own \"Set Text\" node (one per Text object), the same poll-every-frame shape as the code path's Update().",
        "On the Bullet graph's existing kill-handling nodes - the same place the EnemiesRemaining decrement already happens - add an \"Add\" node (Math/Scalar category) taking the Score graph variable and 10, feeding a \"Set Variable\" node targeting Score.",
      ],
      level1Diagram: { file: "diagram_unity_visual_scripting_l9_level1.png", w: 6.3, h: 4.6, caption: "The HUD GameObject's graph and the Bullet graph's kill-handling nodes - the node-for-node equivalent of the Level 1 code." },
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: ["Open your [YourName]Game project from Lesson 8."],
    level1: [
      "On obj_gamemanager's Create event, add __dnd_score = 0; alongside your existing enemies_remaining/has_won/__dnd_lives/has_lost. (__dnd_score is the instance variable GameMaker's own drag-and-drop Score blocks use - the \"dnd\" stands for drag-and-drop. Use it instead of the old built-in score, which is deprecated, just as Lesson 8 did for lives.)",
      "On obj_bullet's Collision event with obj_enemy, add obj_gamemanager.__dnd_score += 10; alongside the existing enemies_remaining decrement.",
      "On obj_gamemanager's existing Draw GUI event, add two unconditional draw_text() calls near the top, before the has_won/has_lost checks - one for score, one for lives - so they display every single frame, not just at a win/lose moment (code below).",
      "Press F5 (Run), shoot a few enemies, and let one escape - both numbers should update immediately every time.",
    ],
    level1Code: [
      "// obj_gamemanager's Create event:",
      "enemies_remaining = 5;",
      "has_won = false;",
      "__dnd_lives = 3;",
      "has_lost = false;",
      "@@ __dnd_score = 0;",
      "",
      "// obj_bullet's Collision event with obj_enemy:",
      "audio_play_sound(sfx_hit, 1, false);",
      "obj_gamemanager.enemies_remaining -= 1;",
      "@@ obj_gamemanager.__dnd_score += 10;",
      "if (obj_gamemanager.enemies_remaining <= 0) {",
      "    obj_gamemanager.has_won = true;",
      "}",
      "instance_destroy(other);",
      "instance_destroy(self);",
      "",
      "// obj_gamemanager's Draw GUI event:",
      "@@ draw_text(20, 20, \"Score: \" + string(__dnd_score));",
      "@@ draw_text(20, 50, \"Lives: \" + string(__dnd_lives));",
      "if (has_won) {",
      "    draw_text(400, 50, \"You Win!\");",
      "}",
      "if (has_lost) {",
      "    draw_text(400, 90, \"Game Over\");",
      "}",
    ],
    level2: [
      "Shoot 3 enemies and confirm the display reads exactly \"Score: 30\".",
      "Let an enemy escape and confirm the score display does NOT change - only the lives display should move.",
      "Because Draw GUI runs every single frame and reads __dnd_score/__dnd_lives directly (no separate copy stored anywhere), the display can never be out of sync with the real values.",
    ],
    level3: [
      "Visual: use draw_set_font() and draw_set_color() before the draw_text() calls to style the HUD text.",
      "Visual: change points-per-kill to irandom_range(5, 15) for scoring variety.",
      "Structure: the score increment lives in obj_bullet's Collision event - the same place enemies_remaining already changes, since both are direct consequences of the exact same confirmed hit.",
    ],
    trouble: [
      "Score doesn't appear on screen: confirm the new draw_text() calls were added to Draw GUI, not the regular Draw event (Draw GUI is unaffected by camera/room position, which is why it's used for HUD elements).",
      "Score changes when an enemy escapes: confirm the score increment is only in obj_bullet's Collision event, not accidentally added to obj_enemy's Step event too.",
      "Score and lives text overlap other text: adjust the Y coordinates in the draw_text() calls so they don't collide with the has_won/has_lost messages.",
      "Visual path: Set Score needs its Target (drop-down arrow, upper right of the block) set to obj_gamemanager - left on the default it adds points to the bullet, which is about to be destroyed, and the score stays at 0.",
    ],
    resource: "Student Resource List > GameMaker section, or gamemaker.io/en/tutorials",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same continuous-HUD pattern works with drag-and-drop blocks instead of typed code. GameMaker's Score and Lives blocks are purpose-built: they keep the instance variables __dnd_score and __dnd_lives (not globals), so each block that changes or checks them needs its Target set to obj_gamemanager. The code path uses those same two variables, so the paths are interchangeable.",
      steps: [
        "On obj_gamemanager's Create event, add a \"Set Score\" block (Instance Variables category) with Value 0 and Relative unchecked, alongside the \"Set Lives\" block from Lesson 8.",
        "On obj_bullet's Collision event with obj_enemy, add a \"Set Score\" block: click the Target drop-down arrow (upper right of the block), choose obj_gamemanager, check Relative, and set Value to 10. Keep it alongside the existing enemies_remaining block, which stays as a Set Variable block since enemies_remaining is a custom variable with no dedicated block.",
        "On obj_gamemanager's existing Draw GUI event, near the top before the has_won/has_lost checks, add a \"Draw Instance Score\" block (Drawing category) with Caption \"Score: \", X 20, Y 20, and a \"Draw Instance Lives\" block with Caption \"Lives: \", Sprite spr_player, Stack Order Row, X 20, Y 50 - it draws one small ship per life remaining instead of a number. (Prefer a number? A \"Draw Value\" block with Caption \"Lives: \" and Value __dnd_lives does that.)",
      ],
      level1Diagram: { file: "diagram_gml_visual_l9_level1.png", w: 6.3, h: 4.6, caption: "obj_bullet's Collision event and obj_gamemanager's Draw GUI event - the block-for-block equivalent of the Level 1 code." },
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 8.",
      "Add two Text objects, name them ScoreText and LivesText, and place them in a corner of the scene. Leave both visible (unlike WinText/LoseText, these stay on screen the whole game).",
      "Add a new scene variable Score, default value 0.",
    ],
    level1: [
      "In the Combat event group's collision event, add a new action: Score = Score + 10, alongside the existing EnemiesRemaining action.",
      "Add a new \"always\" event (no condition) with two actions: change the text of ScoreText and LivesText to show the current values (code below).",
      "Click the green Preview button, shoot a few enemies, and let one escape - both text objects should update immediately every time.",
    ],
    level1Code: [
      "(Combat event group)",
      "Condition: Bullet is in collision with Enemy",
      "Action: Play a sound (sfx_hit.wav)",
      "Action: EnemiesRemaining = EnemiesRemaining - 1",
      "@@ Action: Score = Score + 10",
      "Action: Delete Bullet",
      "Action: Delete Enemy",
      "",
      "@@ (new event, always - no condition)",
      "@@ Action: Change text of ScoreText: \"Score: \" + ToString(Score)",
      "@@ Action: Change text of LivesText: \"Lives: \" + ToString(Lives)",
    ],
    level2: [
      "Shoot 3 enemies and confirm ScoreText reads exactly \"Score: 30\".",
      "Let an enemy escape and confirm ScoreText does NOT change - only LivesText should move.",
      "Because the HUD-update event has no condition, it runs every single frame - the display can never be stale.",
    ],
    level3: [
      "Visual: style ScoreText and LivesText's font/color in the object editor for a more game-like HUD look.",
      "Visual: change points-per-kill to a randomized value using GDevelop's Random() expression, e.g. Score = Score + (5 + Random(10)).",
      "Structure: group the new HUD-update event into its own labeled group (e.g. \"HUD\"), separate from Combat, Spawn, and Movement - continuing the organizational habit from Lesson 7.",
    ],
    trouble: [
      "Score text never updates: confirm the always-event's action actually references the Score variable (ToString(Score)), not a typo'd or different variable name.",
      "Score changes when an enemy escapes: confirm the Score = Score + 10 action is only inside the Combat group, not accidentally added to the escape-cleanup event too.",
      "ScoreText and LivesText overlap on screen: reposition one of them in the scene editor.",
    ],
    resource: "Student Resource List > GDevelop section, or gdevelop.io/tutorials",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: ["Launch TIC-80 and load your cart from Lesson 8.", "Press Esc to reach the code editor."],
    level1: [
      "Add score=0 near the top of your code, alongside lives and enemies_remaining.",
      "In the bullets loop's hit branch (where enemies_remaining decrements), add score=score+10.",
      "At the bottom of TIC(), add two unconditional print() calls - one for score, one for lives - so they display every single frame (code below).",
      "Press Esc, type run, press Enter, then shoot a few enemies and let one escape - both numbers should update immediately every time.",
    ],
    level1Code: [
      "-- top of your code, alongside lives and enemies_remaining:",
      "@@ score=0",
      "",
      "-- bullets loop's hit branch:",
      "-- ...existing overlap check...",
      "e.alive=false",
      "enemies_remaining=enemies_remaining-1",
      "@@ score=score+10",
      "",
      "-- bottom of TIC(), unconditional every frame:",
      "@@ print(\"Score: \"..score,5,5,12)",
      "@@ print(\"Lives: \"..lives,5,15,12)",
    ],
    level2: [
      "Shoot 3 enemies and confirm the display reads exactly \"Score: 30\".",
      "Let an enemy escape and confirm the score display does NOT change - only the lives display should move.",
      "Because these print() calls run every single frame with no caching, the display can never be out of sync with the real values.",
    ],
    level3: [
      "Visual: change the color argument (the last number) in the print() calls, or reposition them, for a more game-like HUD look.",
      "Visual: change points-per-kill to math.random(5,15) for scoring variety.",
      "Structure: keep the score increment inside the bullets loop's hit branch - the same place enemies_remaining already changes, since both are direct consequences of the exact same confirmed hit.",
    ],
    trouble: [
      "Score doesn't appear: confirm the print() calls are actually inside TIC() (which runs every frame), not accidentally placed above it in the one-time setup code.",
      "Score changes when an enemy escapes: confirm score=score+10 is only in the bullets loop's hit branch, not the enemy-drawing loop's escape branch.",
      "Score and lives text overlap other text: adjust the Y coordinates in the print() calls so they don't collide with the You Win!/Game Over messages.",
    ],
    resource: "Student Resource List > TIC-80 section, or tic80.com",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 8.",
      "Create a new Variable (Resources > Variables/Switches/Collections/Tables) named score, default value 0.",
      "Create a new Interface (this is a persistent on-screen layout, different from a one-off Message Box). Add a \"Text / Graphic\" Field to it for score, and another for lives.",
      "On each Field, under Use Values, set Retrieve Text/Value to the matching Variable (score or lives) - and IMPORTANT: tick \"Update Real-time\" on both Fields, so they refresh automatically as the Variables change.",
    ],
    level1: [
      "On the enemy Actor Template's existing \"Collided with actor\" trigger, add a Variable event that increases score by 10, alongside the existing enemies_remaining decrement (code below).",
      "Save, then Play/Test and confirm your new Interface is showing (Interfaces may need to be set to display automatically at game start - check your Interface's settings if it doesn't appear).",
      "Shoot a few enemies and let one escape - both Fields should update immediately every time, with no need to reopen anything.",
    ],
    level1Code: [
      "enemy Actor Trigger: Collided with actor ->",
      "@@   Variable – Modify Value(score, Add, 10)",
      "  Variable – Modify Value(enemies_remaining, Subtract, 1)",
      "  If(enemies_remaining <= 0)",
      "    Message Box(\"You Win!\")",
      "  Actor – Destroy/Revive Actor(relate, Destroy)",
      "  Actor – Destroy/Revive Actor(this, Destroy)",
      "",
      "@@ Interface Field (Score): Use Values -> Retrieve Text/Value -> Variable: score, Update Real-time: ON",
      "@@ Interface Field (Lives): Use Values -> Retrieve Text/Value -> Variable: lives, Update Real-time: ON",
    ],
    level2: [
      "Shoot 3 enemies and confirm the Score Field reads exactly 30.",
      "Let an enemy escape and confirm the Score Field does NOT change - only the Lives Field should move.",
      "If a Field's number only changes after you close and reopen the Interface, its Update Real-time checkbox isn't actually ticked - go back and confirm it on both Fields.",
    ],
    level3: [
      "Visual: try the \\V[score] markup inside a styled Field label instead of a raw bound value, for more control over formatting.",
      "Visual: use a \"Units\" Field type (e.g. heart icons) for the lives display instead of a plain number.",
      "Structure: keep the score increment inside the same \"Collided with actor\" trigger as the enemies_remaining decrement - both are direct consequences of the exact same confirmed hit.",
    ],
    trouble: [
      "Score Field shows the right number, but only after reopening the Interface: Update Real-time isn't ticked on that Field - this is the single most common mistake this week.",
      "Interface doesn't appear at all when testing: check the Interface's own display settings - it may need to be explicitly set to show automatically at game/level start.",
      "Score changes when an enemy escapes: confirm the score increment is only in \"Collided with actor,\" not accidentally added to \"Left/Entered Camera\" too.",
    ],
    resource: "001gamecreator.com/wiki/field/, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 8.",
      "Content Browser > Add > User Interface > Widget Blueprint, name it WBP_HUD. Open it and, in the Designer view, add two Text Blocks: ScoreText and LivesText, positioned in a top corner.",
      "In BP_GameManager - the placed Actor from Lesson 6 that already holds EnemiesRemaining and Lives - add a new Integer variable Score, default value 0.",
    ],
    level1: [
      "In WBP_HUD's own Event Graph, add a Custom Event named UpdateHUD with two Integer inputs (Score, Lives), that sets both Text Blocks' text from those inputs (code below).",
      "In BP_GameManager, on Event BeginPlay, add \"Create Widget\" (Class = WBP_HUD) -> \"Add to Viewport\", and store the result in a new HUDWidget variable on BP_GameManager. Right after, call HUDWidget's UpdateHUD, passing Score and Lives, so the HUD shows correct starting values immediately.",
      "In BP_GameManager's existing HandleEnemyDestroyed, add Score = Score + 10 and call HUDWidget's UpdateHUD again. In BP_GameManager's existing HandleEnemyEscaped, call UpdateHUD again too, right after decreasing Lives.",
      "Click Compile, then Play, shoot a few enemies, and let one escape - both Text Blocks should update immediately every time.",
    ],
    level1Code: [
      "@@ WBP_HUD's Event Graph (new):",
      "@@ Custom Event UpdateHUD(Score: Integer, Lives: Integer) ->",
      "@@   Set Text(ScoreText, \"Score: \" + ToString(Score))",
      "@@   Set Text(LivesText, \"Lives: \" + ToString(Lives))",
      "",
      "BP_GameManager:",
      "@@ Event BeginPlay (new) ->",
      "@@   Create Widget(Class = WBP_HUD) -> Add to Viewport -> Set HUDWidget",
      "@@   Call UpdateHUD(HUDWidget, Score, Lives)",
      "",
      "Custom Event HandleEnemyDestroyed ->",
      "  EnemiesRemaining = EnemiesRemaining - 1",
      "@@   Score = Score + 10",
      "@@   Call UpdateHUD(HUDWidget, Score, Lives)",
      "  Branch(EnemiesRemaining <= 0) -> True: Print String(\"You Win!\", Print to Screen = true)",
      "",
      "Custom Event HandleEnemyEscaped ->",
      "  Lives = Lives - 1",
      "@@   Call UpdateHUD(HUDWidget, Score, Lives)",
      "  Branch(Lives <= 0) -> True: Print String(\"Game Over\", Print to Screen = true)",
    ],
    level2: [
      "Shoot 3 enemies and confirm ScoreText reads exactly \"Score: 30\".",
      "Let an enemy escape and confirm ScoreText does NOT change - only LivesText should move.",
      "Confirm both Text Blocks show correct values the instant Play starts, not just after the first kill or escape (this is why UpdateHUD is called once right after Add to Viewport).",
    ],
    level3: [
      "Visual: style the Text Blocks' font size/color in WBP_HUD's Designer view for a more game-like HUD look.",
      "Visual: change points-per-kill to a randomized value using a Random Integer in Range node.",
      "Structure: notice UpdateHUD is now the ONLY place that touches ScoreText or LivesText directly - BP_GameManager just changes numbers and calls UpdateHUD to reflect them, the same push-based pattern used in Godot this week, and the same one dedicated manager Actor that's owned all of this game's shared state since Lesson 6.",
    ],
    trouble: [
      "HUD never appears: confirm Create Widget and Add to Viewport both actually ran on BP_GameManager's BeginPlay (check that it's properly wired, not left disconnected).",
      "HUD shows \"Score: 0\" forever: confirm Call UpdateHUD is actually wired inside BP_GameManager's HandleEnemyDestroyed, not just defined in WBP_HUD and never invoked.",
      "Compile error on UpdateHUD: confirm the Custom Event's Score and Lives inputs are typed as Integer, matching BP_GameManager's variable types exactly.",
      "Score or HUDWidget not found: confirm Score, HUDWidget, EnemiesRemaining, and Lives are all variables on BP_GameManager - not accidentally created on the Level Blueprint instead.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    level1EngineDiagram: { file: "diagram_unreal_blueprint_l9_level1.png", w: 6.3, h: 4.65, caption: "WBP_HUD's Event Graph and BP_GameManager's BeginPlay and HandleEnemyDestroyed - the node-for-node equivalent of the Level 1 code." },
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same push-based HUD pattern works in C++ instead of Blueprint nodes, calling into a UMG widget from game code.",
      steps: [
        "Create a UUserWidget subclass (e.g. UCppHUDWidget) with UPROPERTY(meta=(BindWidget)) UTextBlock* ScoreText and LivesText, and a function UpdateHUD(int32 Score, int32 Lives) that sets both texts.",
        "In your GameManager-equivalent C++ Actor class - the same class already holding EnemiesRemaining and Lives - create and add the widget once at BeginPlay, then call UpdateHUD() from the same places that already change Score and Lives.",
      ],
      note: "This mirrors the Blueprint path's Custom Event exactly - same push-based shape, just from code.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_09_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 9 Teacher's Key: On the Scoreboard", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine + " - Instructor Reference Only", size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    h2("What This Lesson Teaches"),
    p(data.teaches),
    h2("Level 1 - Full Worked Solution"),
    ...data.level1.map(s => bullet(s)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  if (data.level1EngineDiagram) children.push(...diagram(data.level1EngineDiagram.file, data.level1EngineDiagram.w, data.level1EngineDiagram.h, data.level1EngineDiagram.caption));
  children.push(
    h2("Why This Structure"),
    ...data.why.map(s => bullet(s)),
    h2("Level 2 / Level 3 Talking Points"),
    ...data.talkingPoints.map(s => bullet(s)),
    ...(data.level2Code ? [p("Level 2 - complete code:", { bold: true }), code(data.level2Code)] : []),
    ...(data.level3Code ? [p("Level 3 - complete code (all goals combined):", { bold: true }), code(data.level3Code, "Highlighted lines are the Level 3 additions on top of the Level 1 code.")] : []),
    h2("Common Student Mistakes"),
    ...data.mistakes.map(s => bullet(s)),
  );
  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 9 Key - " + data.engine + " (Instructor Only)") },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    teaches: "That a genuinely \"real\" UI can be entirely free - Scratch's built-in variable watchers are not a placeholder or a workaround, they're the actual, correct tool for this job in this engine, the same way draw_gui is GameMaker's correct tool.",
    level1: [
      "score is a second \"for all sprites\" variable, exactly parallel to lives and enemies_remaining.",
      "The score change lives inside GameManager's existing \"when I receive [enemy hit]\" script, alongside the enemies_remaining change - one broadcast, two consequences, both handled on the one object that already owns this game's shared state. The Enemy sprite still only ever broadcasts enemy hit; it never touches score itself.",
      "Checking the watcher checkboxes is the entire \"UI implementation\" - there is no additional code, no additional sprite, nothing else to build.",
    ],
    why: [
      "This is a valuable contrast to draw next to Unity/Unreal's much more involved setup this week: sometimes the simplest tool really is the right tool, and recognizing that (rather than assuming more setup = more correct) is itself a production skill.",
      "Keeping score and enemies_remaining changes in the same GameManager script mirrors the reuse principle established since Lesson 6 - anything that's a direct, simultaneous consequence of the same broadcast belongs in the same place, on the one dedicated manager sprite.",
    ],
    talkingPoints: [
      "Ask students who also work in another engine this week to compare setup effort - Scratch's total UI work is two checkboxes, versus several new nodes/objects elsewhere. Worth naming explicitly as a tradeoff of using a block-based, watcher-driven engine.",
      "Level 3 visual: \"large readout\" mode is worth demonstrating live - it's a genuinely different-looking display from the default small watcher, useful for students who want their HUD to feel more like a real game.",
    ],
    mistakes: [
      "Forgetting to check the watcher checkbox for one of the two variables (usually lives, since it was added in a prior lesson and may already be hidden by default).",
      "Adding \"change score by 10\" to GameManager's wrong script (the \"when I receive [enemy escaped]\" script instead of \"when I receive [enemy hit]\").",
      "Confusing the show/hide checkbox in the Variables palette with the sprite's own show/hide - these are unrelated; the variable watcher's visibility is independent of any sprite's.",
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    teaches: "The push-based update-function pattern: one small function (update_hud()) is the single source of truth for what the display shows, called every time - and only when - the underlying data actually changes.",
    level1: [
      "update_hud() lives on GameManager, alongside score, lives, enemies_remaining, on_enemy_destroyed(), and on_enemy_escaped() - the same dedicated manager node that's owned all of this game's shared state since Lesson 6. It reads score and lives and writes both Labels' text via get_node(\"../ScoreText\")/get_node(\"../LivesText\") - the same relative-path pattern already used to reach WinText and LoseText - and never changes score or lives itself, only reflects them.",
      "Every place that changes score or lives (currently on_enemy_destroyed() and on_enemy_escaped(), both on GameManager) calls update_hud() immediately afterward - plus one extra call in GameManager's own _ready() so the HUD is correct before anything has happened yet.",
      "This is a genuinely new structural idea this week: previous lessons separated different KINDS of game logic (hit vs. escape); this lesson separates game logic from DISPLAY - all still on the one manager node.",
    ],
    why: [
      "Centralizing all display-writing into one function means there's exactly one place to look if the HUD ever shows something wrong, and exactly one place to change if the HUD's format ever needs to change (e.g. adding a comma to large scores later).",
      "This pattern scales well: as more stats get added in later lessons (wave number, high score, etc.), they extend update_hud() rather than requiring new scattered Label-writing code - and they belong on GameManager too, the same one-shared-state-home-per-project principle Rule 13 established.",
    ],
    talkingPoints: [
      "Ask students to predict what would go wrong if a function changed score directly without calling update_hud() - the number would be correct internally but wrong on screen, a good illustration of why the two need to travel together.",
      "Compare with Unity's approach this same lesson (a polling Update() instead of a push function) - ask which one this Godot pattern most resembles conceptually to Unity's own signal-based design from Lessons 6-8 (it's the same push philosophy, just applied to display instead of destruction).",
    ],
    mistakes: [
      "Forgetting to call update_hud() inside GameManager's own _ready() - the HUD shows blank or default text until the first kill or escape.",
      "Changing score or lives directly inside update_hud() by mistake (it should only read them, never write them).",
      "Defining update_hud() or score on the main scene's script instead of GameManager's - breaks the one-dedicated-manager pattern and makes get_node(\"../ScoreText\") resolve from the wrong place.",
      "Node name typos in the get_node() calls (e.g. \"../ScoreLable\" vs. the actual node name in the Scene tree) - case-sensitive and exact-match required.",
    ],
  },
  {
    engine: "Unity", file: "Unity",
    teaches: "The polling pattern (checking data every single frame in Update(), regardless of whether it changed) as a legitimate, simple alternative to Godot/Unreal's push-based approach - and when each is the more natural fit.",
    level1: [
      "HUD.cs is a small, separate script whose only job is displaying data - it never changes score or lives, only reads them from GameManager every frame.",
      "Because Update() runs every frame unconditionally, there's no need to remember to call a refresh function anywhere else in the codebase - the tradeoff is doing slightly more work than strictly necessary (checking even when nothing changed), which is harmless for two small ints.",
      "GameManager.cs continues to be the single source of truth for score/lives - HUD.cs only reads it, never writes it.",
    ],
    why: [
      "Polling is simpler to reason about for beginners specifically because there's no need to identify every single place data changes and remember to call a refresh - a real advantage for a script that only needs to display two cheap values.",
      "This contrasts usefully with GameManager's own RegisterKill()/RegisterEscape() methods, which ARE event-driven (called only when something happens) - the lesson isn't \"polling is better,\" it's \"different jobs suit different approaches.\"",
    ],
    talkingPoints: [
      "Ask students: would polling still make sense if the HUD needed to update 500 different values every frame? (Probably not - this is a good entry point for why push-based approaches exist and matter at larger scale, previewing ideas they'll meet again in more advanced programming.)",
      "For students who did Godot too: have them explicitly compare update_hud() (called only when needed) against HUD.cs's Update() (called unconditionally, every frame) - both are correct, valid solutions to the same problem.",
    ],
    mistakes: [
      "Forgetting to assign ScoreText/LivesText in the Inspector after writing HUD.cs - causes a NullReferenceException the instant Play starts.",
      "Declaring score as a non-static field on GameManager by mistake, breaking the static-class pattern established since Lesson 6.",
      "Placing the new UI Text objects outside the Canvas (rare, but possible if a Canvas wasn't auto-created correctly) - text won't render if this happens.",
    ],
    level1EngineDiagram: { file: "diagram_unity_visual_scripting_l9_level1.png", w: 6.3, h: 4.6, caption: "The HUD GameObject's graph and the Bullet graph's kill-handling nodes - the Visual Scripting equivalent of this solution, for students comparing both paths." },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    teaches: "That GameMaker's Draw GUI event has been the 'real' HUD tool the whole time, not a placeholder - this lesson simply asks it to draw two more things, every frame, the same way it's already drawing conditional win/lose messages.",
    level1: [
      "__dnd_score is a new instance variable on obj_gamemanager, parallel to enemies_remaining and __dnd_lives - the same variable GameMaker's drag-and-drop Score blocks use, so Code and Visual students are working with identical data.",
      "The score increment lives in obj_bullet's Collision event - the exact same place enemies_remaining already changes, since both are consequences of the exact same confirmed hit.",
      "The two new draw_text() calls are unconditional (no if-check around them, unlike has_won/has_lost) - they simply run every single frame, which is exactly what a persistent HUD needs.",
    ],
    why: [
      "Draw GUI was already the correct, idiomatic GameMaker tool for HUD elements as far back as Lesson 6's win message - this lesson doesn't introduce a new mechanism, it just uses the existing one unconditionally instead of conditionally.",
      "Because Draw GUI re-reads __dnd_score and __dnd_lives directly from the variables every frame (no separate cached copy), there's no possible way for the display to fall out of sync - a nice, simple guarantee worth naming explicitly.",
      "Visual path: purpose-built blocks throughout - Set Score (Relative, Value 10, Target obj_gamemanager) in obj_bullet's Collision event, then Draw Instance Score and Draw Instance Lives (spr_player) in Draw GUI. Because the code path uses the same __dnd_score/__dnd_lives variables, a student can mix typed code and blocks in one project. Draw Instance Lives only draws frame 0 of its sprite and ignores image_xscale/image_blend - fine for spr_player.",
    ],
    talkingPoints: [
      "Ask students to contrast this week's unconditional draw_text() calls with the existing conditional ones (has_won/has_lost) - same event, same function, different condition logic. A good moment to reinforce that \"always show\" and \"show only when X\" are both just conditions (or the absence of one).",
      "For students curious about visual polish: draw_set_font() and draw_set_color() are worth a quick live demo, since GameMaker's default draw_text() styling is quite plain.",
    ],
    mistakes: [
      "Adding the new draw_text() calls to the regular Draw event instead of Draw GUI (causes the HUD to scroll/zoom with the room camera instead of staying fixed on screen).",
      "Adding the score increment to obj_enemy's Step event by mistake instead of obj_bullet's Collision event (makes escaping award points, which is backwards).",
      "Y-coordinate collisions between the new score/lives text and the existing has_won/has_lost messages, especially once both are visible at once near end-of-game.",
    ],
    level1EngineDiagram: { file: "diagram_gml_visual_l9_level1.png", w: 6.3, h: 4.6, caption: "obj_bullet's Collision event and obj_gamemanager's Draw GUI event - the Visual (drag-and-drop) equivalent of this solution, for students comparing both paths." },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    teaches: "That a persistent HUD in GDevelop needs one small addition beyond what score/lives already needed: an explicit \"always update the display\" event, since Text objects don't automatically reflect variable changes the way Scratch's watchers or 001GC's real-time Fields do.",
    level1: [
      "Score is a new scene variable, parallel to EnemiesRemaining and Lives.",
      "The Score = Score + 10 action lives in the Combat group's collision event - the exact same event that already changes EnemiesRemaining, since both are consequences of the exact same confirmed hit.",
      "The new always-event (no condition) is what makes this a genuinely continuous HUD - GDevelop's Text objects hold whatever text they were last explicitly set to, so this event needs to run every frame to stay accurate.",
    ],
    why: [
      "This is a useful contrast for students moving between engines this week: GDevelop's Text objects are 'dumb' displays that need to be told what to say, unlike Scratch's watchers (always live) or 001GC's real-time Fields (bound automatically) - worth naming this difference explicitly.",
      "Keeping the HUD-update event unconditional (no condition at all) guarantees the display never falls behind, at the small, harmless cost of re-setting text that sometimes hasn't actually changed.",
    ],
    talkingPoints: [
      "Ask students to predict what would happen if the HUD-update event had a condition attached (e.g., only running on a collision) - the display would only refresh at those specific moments, which happens to work for score/lives here but would break for anything that changes for other reasons (like a countdown timer).",
      "Level 3 structure: this is a good moment to introduce or reinforce GDevelop's event groups as a scaling technique - a project with Combat, Spawn, Movement, and now HUD groups is starting to look like a real, organized project.",
    ],
    mistakes: [
      "Forgetting the always-event entirely and expecting the Text objects to update automatically just because the Score variable changed - they won't, GDevelop requires an explicit Change Text action.",
      "Adding Score = Score + 10 to the wrong event (escape-cleanup instead of Combat).",
      "ScoreText and LivesText positioned at the same coordinates, causing visual overlap.",
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    teaches: "That TIC-80's print() has been the 'real' HUD tool the whole time - this lesson just moves it from conditional (You Win!/Game Over) to unconditional (always-visible score/lives), the same underlying mechanism used differently.",
    level1: [
      "score=0 is a new global, parallel to lives and enemies_remaining.",
      "The score increment lives inside the bullets loop's hit branch - the exact same branch that already changes enemies_remaining, since both are consequences of the exact same confirmed hit.",
      "The two new print() calls at the bottom of TIC() run every single frame with no surrounding if-check, unlike the existing conditional You Win!/Game Over prints.",
    ],
    why: [
      "print() was already the correct, idiomatic TIC-80 tool for on-screen text as far back as Lesson 6 - nothing new is being introduced, only a new (unconditional) way of calling something already familiar.",
      "Because these calls re-read score and lives directly every frame, the display can never fall out of sync - the same guarantee GameMaker's Draw GUI approach offers, for the same underlying reason (direct read, every frame, no caching).",
    ],
    talkingPoints: [
      "Ask students to contrast the always-run score/lives prints with the conditional You Win!/Game Over prints in the same TIC() function - both use print(), the only difference is whether an if surrounds the call.",
      "For students working with larger carts: this is a good moment to ask whether TIC-80's single long TIC() function is starting to feel unwieldy, and whether breaking pieces into named functions (draw_hud(), update_enemies(), etc.) might help - a preview of concerns that grow through Lesson 11.",
    ],
    mistakes: [
      "Placing the new print() calls above TIC() (in the one-time setup code) instead of inside it - they'd only ever draw once, not every frame.",
      "Adding score=score+10 to the enemy-drawing loop's escape branch instead of the bullets loop's hit branch (makes escaping award points, which is backwards).",
      "Y-coordinate collisions between the new score/lives prints and the existing You Win!/Game Over prints.",
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    teaches: "001 Game Creator's actual, wiki-documented persistent-UI mechanism: an Interface containing Fields, distinct from the one-off Message Box used in Lessons 6 and 8. This was specifically researched and verified against the official wiki before writing this lesson, given this engine's history of needing extra care.",
    level1: [
      "An Interface is a scene-level, persistent UI layout - fundamentally different from Message Box, which is a one-off popup. Interfaces are the correct, documented tool for HUDs in this engine (confirmed via the official 001 Game Creator wiki's Field page, and the wiki's own \"RPG – Menus/HUD\" tutorial, which uses this exact mechanism for a health display).",
      "Each Field's \"Use Values\" section can bind directly to a Variable via \"Retrieve Text/Value\" - and critically, ticking \"Update Real-time\" makes that binding live, refreshing automatically the instant the Variable changes, with no manual re-set action needed anywhere in the Trigger scripts.",
      "The score increment itself lives in the enemy's existing \"Collided with actor\" Trigger, alongside the enemies_remaining decrement - no new Trigger needed, since the display updates itself via the real-time binding.",
    ],
    why: [
      "This is meaningfully different from every other engine's HUD approach this week: 001 Game Creator's binding is neither push (an explicit function call after each change) nor exactly poll (no per-frame check written by the student) - the engine itself handles keeping the display in sync, once Update Real-time is configured correctly.",
      "This is exactly why the setup step matters so much this week: get the Field's binding and Update Real-time setting right once, and the Trigger-script side of this lesson becomes almost trivially simple (one new Variable action in an existing Trigger).",
    ],
    talkingPoints: [
      "This is worth explicitly contrasting with GDevelop's Text objects this same lesson: GDevelop needs an explicit \"always\" event to push text updates every frame, while 001 Game Creator's real-time Field binding handles that automatically once configured - two different engines solving the same problem with genuinely different underlying mechanisms.",
      "If time allows, demonstrate what happens with Update Real-time UNCHECKED (the Field freezes until the Interface is closed and reopened) - seeing the broken behavior makes the correct setting's purpose click for students immediately.",
    ],
    mistakes: [
      "Forgetting to tick Update Real-time on one or both Fields - the single most likely mistake this week, and one that looks like a mysterious bug rather than an obviously missing setting.",
      "Forgetting to set the Interface to display automatically at game/level start (a separate setting from the Field bindings themselves).",
      "Binding a Field to the wrong Variable (e.g., both Fields accidentally bound to score) - always test both independently, by shooting an enemy alone and separately letting one escape alone.",
    ],
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    teaches: "The UMG (Unreal Motion Graphics) Widget Blueprint system - Unreal's actual, standard on-screen UI framework, introduced properly this week after Lessons 6-8 deliberately deferred to simple Print String messages.",
    level1: [
      "WBP_HUD is a Widget Blueprint - a separate, reusable UI asset containing Text Blocks, fundamentally different from the one-off Print String calls used in Lessons 6 and 8.",
      "UpdateHUD is a push-based Custom Event on the widget itself, called from BP_GameManager - the same placed Actor that already owns EnemiesRemaining and Lives - every time Score or Lives actually changes. Score, HUDWidget, and the Create Widget/Add to Viewport setup all live on BP_GameManager too, not the Level Blueprint, the same one-dedicated-manager pattern used since Lesson 6.",
      "Create Widget + Add to Viewport (done once, in BP_GameManager's BeginPlay) is the standard way to actually put a Widget Blueprint on screen - without this step, WBP_HUD would exist as an asset but never actually display.",
    ],
    why: [
      "This mirrors Godot's push-based update_hud() function almost exactly, just expressed through Unreal's widget/dispatcher idiom instead of GDScript - worth naming this cross-engine parallel explicitly for students working in both. Godot's version lives on GameManager; Unreal's lives on BP_GameManager - the same architectural choice, two engines.",
      "Keeping ScoreText/LivesText updates entirely inside UpdateHUD (never set directly from BP_GameManager) mirrors the same discipline established in Godot: one function is the sole owner of what the display shows.",
    ],
    talkingPoints: [
      "For students who also did Unity this week: contrast Unreal's push-based UpdateHUD (called only when something changes) against Unity's poll-based HUD.cs Update() (checked every frame) - both are legitimate, and this is a good moment to ask which one they'd choose for a hypothetical HUD with many rapidly-changing values.",
      "This is the most setup-heavy engine this week (Widget Blueprint creation, Designer view layout, Create Widget/Add to Viewport wiring) - worth acknowledging directly with students working in Unreal that more scaffolding this week is expected, not a sign they're doing something wrong.",
    ],
    mistakes: [
      "Forgetting Add to Viewport after Create Widget - the widget exists in memory but is never actually shown on screen.",
      "Calling UpdateHUD before HUDWidget is actually set (wiring order matters - Create Widget/Add to Viewport/Set HUDWidget must all happen before the first UpdateHUD call).",
      "Text Block name mismatches between the Designer view and the Set Text nodes in the Event Graph - these must reference the exact same Text Block.",
    ],
    level1EngineDiagram: { file: "diagram_unreal_blueprint_l9_level1.png", w: 6.3, h: 4.65, caption: "WBP_HUD's Event Graph and BP_GameManager's BeginPlay and HandleEnemyDestroyed - the node-for-node equivalent of this solution." },
  },
];


// ---- Rule 14: Teacher's Keys carry complete code for all three levels ----
const keyLevel2Default = (c) => [
  c + " Level 2 is verification only - the Level 1 code is already correct, so no lines change.",
  c + " Checks: 3 kills read exactly \"Score: 30\"; an escape moves lives but never score;",
  c + " both numbers are correct the instant the game starts, not only after the first kill or escape.",
];
const keyComment = { "Scratch": "//", "Godot": "#", "TIC-80": "--", "GDevelop": "//", "001GameCreator": "//", "Unreal": "//", "Unity": "//", "GameMaker": "//" };
const keyLevel3 = {
  "Scratch": [
    "// GameManager - when I receive [enemy hit] (Level 3 variety):",
    "change enemies_remaining by -1",
    "@@ change score by (pick random 5 to 15)",
    "if enemies_remaining = 0 then",
    "    broadcast [you win]",
    "",
    "// Level 3 visual: right-click the score or lives watcher on stage and choose \"large readout\"; drag the watchers to a free corner.",
  ],
  "Godot": [
    "# GameManager's script - Level 3 (Theme Overrides for font size/color are set on each Label in the Inspector):",
    "func on_enemy_destroyed():",
    "    enemies_remaining -= 1",
    "@@     score += randi_range(5, 15)",
    "    update_hud()",
    "    if enemies_remaining <= 0:",
    "        get_node(\"../WinText\").show()",
  ],
  "Unity": [
    "// GameManager.cs - Level 3 RegisterKill():",
    "public static void RegisterKill() {",
    "    enemiesRemaining--;",
    "@@     score += Random.Range(5, 16);",
    "    if (enemiesRemaining <= 0) hasWon = true;",
    "}",
    "",
    "// HUD.cs - Level 3 visual (or set these in the Inspector instead):",
    "void Start() {",
    "@@     scoreText.fontSize = 28;",
    "@@     scoreText.color = Color.yellow;",
    "}",
  ],
  "GameMaker": [
    "// obj_bullet's Collision event with obj_enemy - Level 3:",
    "audio_play_sound(sfx_hit, 1, false);",
    "obj_gamemanager.enemies_remaining -= 1;",
    "@@ obj_gamemanager.__dnd_score += irandom_range(5, 15);",
    "if (obj_gamemanager.enemies_remaining <= 0) {",
    "    obj_gamemanager.has_won = true;",
    "}",
    "instance_destroy(other);",
    "instance_destroy(self);",
    "",
    "// obj_gamemanager's Draw GUI event - Level 3 styling:",
    "@@ draw_set_color(c_yellow);",
    "draw_text(20, 20, \"Score: \" + string(__dnd_score));",
    "draw_text(20, 50, \"Lives: \" + string(__dnd_lives));",
    "@@ draw_set_color(c_white);",
    "if (has_won) {",
    "    draw_text(400, 50, \"You Win!\");",
    "}",
    "if (has_lost) {",
    "    draw_text(400, 90, \"Game Over\");",
    "}",
    "",
    "// Visual path: Set Score with Value 10 becomes Value irandom_range(5, 15) (Expression); a Set Color Draw block before the Draw Instance Score block gives the styling.",
  ],
  "GDevelop": [
    "(Combat event group - Level 3)",
    "Condition: Bullet is in collision with Enemy",
    "Action: Play a sound (sfx_hit.wav)",
    "Action: EnemiesRemaining = EnemiesRemaining - 1",
    "@@ Action: Score = Score + (5 + Random(10))",
    "Action: Delete Bullet",
    "Action: Delete Enemy",
    "",
    "(HUD group - Level 3: move the always event into its own event group named HUD)",
    "Action: Change text of ScoreText: \"Score: \" + ToString(Score)",
    "Action: Change text of LivesText: \"Lives: \" + ToString(Lives)",
    "@@ (font and color for ScoreText/LivesText are set in each object's editor)",
  ],
  "TIC-80": [
    "-- bullets loop's hit branch - Level 3:",
    "e.alive=false",
    "enemies_remaining=enemies_remaining-1",
    "@@ score=score+math.random(5,15)",
    "",
    "-- bottom of TIC() - Level 3 (color is the last argument):",
    "@@ print(\"Score: \"..score,5,5,14)",
    "@@ print(\"Lives: \"..lives,5,15,6)",
  ],
  "001GameCreator": [
    "enemy Actor Trigger: Collided with actor -> (unchanged, still the only place score changes)",
    "  Variable – Modify Value(score, Add, 10)",
    "  Variable – Modify Value(enemies_remaining, Subtract, 1)",
    "",
    "@@ Interface Field (Score) - Level 3: label text \"Score: \\\\V[score]\" with a styled font (formatting markup instead of a raw bound value)",
    "@@ Interface Field (Lives) - Level 3: a Units Field type with a heart icon, bound to lives",
  ],
  "Unreal": [
    "BP_GameManager - Level 3:",
    "Custom Event HandleEnemyDestroyed ->",
    "  EnemiesRemaining = EnemiesRemaining - 1",
    "@@   Score = Score + Random Integer in Range(5, 15)",
    "  Call UpdateHUD(HUDWidget, Score, Lives)",
    "  Branch(EnemiesRemaining <= 0) -> True: Print String(\"You Win!\", Print to Screen = true)",
    "",
    "@@ WBP_HUD Designer view: set font size and color on ScoreText and LivesText for the styled HUD.",
  ],
};
keyData.forEach(k => {
  const h = handoutData.find(x => x.file === k.file);
  if (!k.level1Code && h && h.level1Code) k.level1Code = h.level1Code;
  if (!k.level2Code) k.level2Code = keyLevel2Default(keyComment[k.file] || "//");
  if (!k.level3Code && keyLevel3[k.file]) k.level3Code = keyLevel3[k.file];
});

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_09_Key_${d.file}.docx`]));

// =========================================================================
// WRITE ALL FILES
// =========================================================================
(async () => {
  for (const [doc, filename] of jobs) {
    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(filename, buf);
    console.log("wrote", filename);
  }
})();
