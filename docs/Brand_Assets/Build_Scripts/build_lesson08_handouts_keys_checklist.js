const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign,
  ImageRun
} = require("docx");
const brand = require("./brand.js");

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
  const hasNew = lines.some(l => l.startsWith("@@ "));
  if (hasNew) {
    children.push(new TextRun({ text: legend || "Highlighted lines are new this week; everything else is code you already have.", italics: true, size: 18, color: "555555" }));
    children.push(new TextRun({ text: "", break: 1 }));
  }
  lines.forEach((line, i) => {
    const isNew = line.startsWith("@@ ");
    const txt = isNew ? line.slice(3) : line;
    const opts = { text: txt, font: "Consolas", size: 20 };
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
      headers: { default: brand.pageHeader("Lesson 8 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 8 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("This Week: A Payoff, Not a New Concept"),
        p("Good news for prep this week - there's no engine-specific bugfix, unlike Lesson 7. Every engine already has a clean, separate escape-cleanup code path (built in Lesson 7, and for GameMaker/001 Game Creator, specifically fixed there). This week just gives that existing path a real consequence: losing a life instead of a silent, free escape. Students who understood Lesson 7 should find this week's addition small and satisfying - reusing code they already trust."),

        h2("Important: We Are NOT Building Restart Flow Yet"),
        p("An enemy reaching the bottom now costs a life, and reaching zero lives shows a \"Game Over\" message - but the game does not stop, freeze, or offer a restart button this week. That's Lesson 10's job specifically (\"Game-over and restart flow\"). If a student asks why the game keeps running after \"Game Over\" appears, that's expected - remind them we're building this in deliberate layers, the same way the win message in Lesson 6 was just a message before anything more elaborate was added."),
        checkbox("Read through each engine's Handout once before class - the shape is consistent across all 8, but the exact variable/signal names differ."),

        h2("Self-Contained Lesson Assets"),
        p("Core exercise needs no new assets - lives tracking and the lose message use only code/logic changes, reusing each engine's existing basic text tool from Lesson 6. Three placeholder sound effects are provided in this lesson's Assets/ folder for the Level 3 sound idea, so no student has to find or make their own: sfx_lifelost.wav (a short descending bonk, play it each time a life is lost), sfx_gameover.wav (a longer descending game-over stinger, play it when lives reaches zero), and sfx_start.wav (the same game-start stinger from Lesson 6, repeated here so the folder is self-contained)."),

        h2("Continuity From Lesson 7"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 7, not create a new one."),
        checkbox("If a student's Lesson 7 project is lost or broken, have a pre-built \"Lesson 7 complete\" starter available per engine (from the Lesson 7 Teacher's Keys) so they aren't blocked."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop's Lesson 6 and 7 exercises still work for all 8 engines."),
        checkbox("Godot / Unreal: confirm the new escaped/OnEnemyEscaped signal or dispatcher connects cleanly alongside the existing destroyed/OnEnemyDestroyed one - it's easy to only wire one of the two when copying the spawn-loop connection code."),
        checkbox("GameMaker / GDevelop / TIC-80 / Unity / 001 Game Creator: confirm the lives decrement was added to the existing escape-cleanup code, not accidentally duplicated into the hit-detection code too."),
        checkbox("Test on the demo laptop: let all 3 lives run out by letting enemies escape unshot, and confirm \"Game Over\" appears exactly once, at exactly zero lives."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies."),

        h2("Known Gotchas This Week"),
        bullet("The most likely mistake: a student adds the lives decrement to the wrong code path (hit-detection instead of escape-cleanup, or vice versa), which either makes shooting an enemy cost a life or makes escaping count as a kill. Test both directions - shoot an enemy (should NOT lose a life) and let one escape (SHOULD lose a life) - to confirm they're not mixed up."),
        bullet("A student who lets all enemies through without shooting any will see lives hit zero quickly - that's a good built-in test case, not a bug, and worth pointing out during lab."),
        bullet("No continuous \"Lives: 3\" readout is required this week (only the end-state \"Game Over\" message) - a live counter is offered as an optional Level 3 idea, with the full on-screen UI treatment properly arriving in Lesson 9. Don't let students get stuck trying to build a polished HUD this week."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_08_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 8: Under Pressure", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: give the enemies you dodge last week actual stakes. An enemy reaching the bottom now costs a life - lose all 3, and \"Game Over\" appears.", { bold: true }),
    p("Open your own [YourName]Game project from Lesson 7 - don't start a new one.", { bold: true }),
    p("Tip: the core exercise needs no new assets - this reuses last week's escape-cleanup code and just gives it a real consequence. For Level 3 sound, use the provided sfx_lifelost.wav and sfx_gameover.wav from the Lesson 8 Assets folder.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("Every enemy that reaches the bottom unshot now costs the player one life, tracked starting at 3. The moment lives reach zero, \"Game Over\" appears on screen.", { italics: true }),
    ...diagram("diagram_activity_lose_condition.png", 6.3, 4.375, "The lose-condition logic behind this week's exercise - the same shape in every engine, whatever the exact code or blocks look like."),
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
    p("Shooting an enemy never costs a life, and letting one escape never counts as a kill - the two counters stay completely independent.", { italics: true }),
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
      headers: { default: brand.pageHeader("Lesson 8 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 7.",
      "Create a new variable lives (for all sprites, not \"for this sprite only\").",
    ],
    level1: [
      "On GameManager's existing \"when green flag clicked\" script, right where you set enemies_remaining, also add \"set lives to 3\".",
      "In the Enemy's forever loop, find the escape branch you built last week (checking y position < -170). Right before \"delete this clone\", add \"broadcast [enemy escaped]\" - the Enemy only ever reports what happened, it never counts or decides anything itself.",
      "On GameManager, add a new script: \"when I receive [enemy escaped]\" -> \"change lives by -1\" -> \"if lives = 0 then broadcast [game over]\" - a direct parallel to its existing \"when I receive [enemy hit]\" script.",
      "On the Player sprite, add a new script: \"when I receive [game over]\" -> \"say [Game Over] for 3 seconds\".",
      "Click the stage, then let all 3 lives run out by letting enemies escape unshot - \"Game Over\" should appear the instant the third one escapes.",
    ],
    level1Code: [
      "// GameManager (hidden sprite) - existing green flag script:",
      "when green flag clicked",
      "hide",
      "set enemies_remaining to 5",
      "@@ set lives to 3",
      "",
      "// Enemy sprite - the escape branch inside its forever loop:",
      "if y position < -170 then",
      "@@     broadcast [enemy escaped]",
      "    delete this clone",
      "",
      "// GameManager - new script:",
      "@@ when I receive [enemy escaped]",
      "@@ change lives by -1",
      "@@ if lives = 0 then",
      "@@     broadcast [game over]",
      "",
      "// Player sprite - new script:",
      "@@ when I receive [game over]",
      "@@ say [Game Over] for 3 seconds",
    ],
    level2: [
      "Shoot an enemy and confirm lives does NOT change - only GameManager's \"when I receive [enemy hit]\" script, and enemies_remaining, should be involved.",
      "Let an enemy escape and confirm enemies_remaining does NOT change - only GameManager's \"when I receive [enemy escaped]\" script, and lives, should be involved.",
      "Confirm \"Game Over\" appears exactly once, exactly when lives reaches zero.",
    ],
    level3: [
      "Visual: check the box next to the lives variable in the Variables palette to show a live readout on stage - Scratch's simplest built-in way to display a running number, the same idea as \"You Win!\" using a basic tool instead of a fancy one.",
      "Sound: import the provided sfx_gameover.wav and add \"start sound\" right before the \"say\" block on the Player's \"when I receive [game over]\" script. For a life-lost sound too, import sfx_lifelost.wav and add \"start sound\" to GameManager's \"when I receive [enemy escaped]\" script, right after \"change lives by -1\".",
      "Structure: keep lives, the zero-check, and the game-over broadcast entirely inside GameManager's own \"when I receive [enemy escaped]\" script - the Enemy sprite only ever reports that it escaped, it never counts or decides anything itself, exactly mirroring how enemy hit already works.",
    ],
    trouble: [
      "Lives decreases when you shoot an enemy: confirm the Enemy's touching-Bullet branch only ever broadcasts enemy hit, never enemy escaped, and that GameManager's \"when I receive [enemy escaped]\" script is the only place lives changes.",
      "\"Game Over\" never appears: confirm GameManager's green-flag script actually sets lives to 3 (easy to miss since GameManager is hidden), and that the escape branch actually broadcasts enemy escaped.",
      "\"Game Over\" appears multiple times: confirm GameManager has only ONE \"when I receive [enemy escaped]\" script - a second copy would double-count every escape - and that the check uses \"lives = 0\" (exact equality), matching the same pattern the win condition used in Lesson 6.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 7.",
      "Add a Label node to your main scene, rename it LoseText, set its Text to \"Game Over\", and hide it (uncheck Visible in the Inspector).",
      "In enemy.tscn's own script, add a new signal at the top, alongside the existing destroyed signal: signal escaped.",
    ],
    level1: [
      "In enemy.tscn's _process(), find the escape branch you built last week. Right before queue_free(), add escaped.emit() (code below).",
      "In GameManager's script - the same node that already owns enemies_remaining and on_enemy_destroyed() - add var lives := 3 and a new function on_enemy_escaped() that decrements it and shows LoseText at zero (code below).",
      "In the main scene's spawn loop, right where you connect each enemy's destroyed signal to $GameManager.on_enemy_destroyed, also connect its escaped signal to $GameManager.on_enemy_escaped.",
      "Run the scene and let all 3 lives run out by letting enemies escape unshot - \"Game Over\" should appear the instant the third one escapes.",
    ],
    level1Code: [
      "# enemy.tscn's script:",
      "@@ signal escaped",
      "",
      "func _process(delta):",
      "    position.y += enemy_speed * delta",
      "    if position.y > 700:",
      "@@         escaped.emit()",
      "        queue_free()",
      "",
      "# GameManager's script - add alongside enemies_remaining and on_enemy_destroyed():",
      "@@ var lives := 3",
      "",
      "@@ func on_enemy_escaped():",
      "@@     lives -= 1",
      "@@     if lives <= 0:",
      "@@         get_node(\"../LoseText\").show()",
      "",
      "# Main scene's script - in the spawn loop, alongside connecting destroyed:",
      "@@ new_enemy.escaped.connect($GameManager.on_enemy_escaped)",
    ],
    level2: [
      "Shoot an enemy and confirm lives does NOT change - only the destroyed signal, and enemies_remaining, should be involved.",
      "Let an enemy escape and confirm enemies_remaining does NOT change - only the escaped signal, and lives, should be involved.",
      "Confirm LoseText appears exactly once, exactly when lives reaches zero.",
    ],
    level3: [
      "Visual: add a third Label showing the current lives count, updated inside GameManager's on_enemy_escaped() - a nice touch, though the full scoreboard-style UI is coming properly in Lesson 9.",
      "Sound: add two AudioStreamPlayer nodes (not 2D, since they're UI, not positional) as children of GameManager, assign the provided sfx_lifelost.wav to one and sfx_gameover.wav to the other, then call .play() on the first inside on_enemy_escaped() right after lives decreases - and on the second instead, in the branch where lives reaches zero.",
      "Structure: notice that escaped and destroyed are two completely separate signals, each connected once at spawn time, each handled by its own function on GameManager - this is the same pattern that made Godot's Lesson 7 bugfix unnecessary, now extended to a second signal and kept on the same dedicated manager node introduced in Lesson 6.",
    ],
    trouble: [
      "Lives decreases when you shoot an enemy: confirm _on_hitbox_area_entered() only emits destroyed, and _process()'s escape branch only emits escaped - the two should never call into each other.",
      "\"Game Over\" never appears: confirm the spawn loop actually connects new_enemy.escaped.connect($GameManager.on_enemy_escaped) for every spawned enemy, not just destroyed.",
      "Nothing happens when enemies escape: confirm on_enemy_escaped() is actually defined on GameManager's script (not accidentally left on the main scene's own script), and that LoseText's node path ($../LoseText, referenced from GameManager) matches exactly.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: ["Open your [YourName]Game project from Lesson 7."],
    level1: [
      "In GameManager.cs, add a public static int lives = 3, a public static bool hasLost = false, and a static method RegisterEscape() that decrements lives and sets hasLost at zero (code below).",
      "In Enemy.cs's Update(), find the escape branch you built last week. Right before Destroy(gameObject), call GameManager.RegisterEscape().",
      "In GameManager.cs's OnGUI(), add a check that draws \"Game Over\" when hasLost is true, alongside the existing hasWon check.",
      "Press Play and let all 3 lives run out by letting enemies escape unshot - \"Game Over\" should appear the instant the third one escapes.",
    ],
    level1Code: [
      "// GameManager.cs - add fields and a new method:",
      "@@ public static int lives = 3;",
      "@@ public static bool hasLost = false;",
      "",
      "@@ public static void RegisterEscape() {",
      "@@     lives--;",
      "@@     if (lives <= 0) hasLost = true;",
      "}",
      "",
      "// GameManager.cs's OnGUI() - add alongside the hasWon check:",
      "@@ if (hasLost) {",
      "@@     GUI.Label(new Rect(Screen.width/2 - 50, 90, 200, 30), \"Game Over\");",
      "}",
      "",
      "// Enemy.cs's Update() - escape branch (revised):",
      "if (transform.position.y < -6f) {",
      "@@     GameManager.RegisterEscape();",
      "    Destroy(gameObject);",
      "}",
    ],
    level2: [
      "Shoot an enemy and confirm lives does NOT change - only RegisterKill(), called from Bullet.cs, should run.",
      "Let an enemy escape and confirm enemiesRemaining does NOT change - only RegisterEscape(), called from Enemy.cs's own Update(), should run.",
      "Confirm \"Game Over\" appears exactly once, exactly when lives reaches zero.",
    ],
    level3: [
      "Visual: add a GUI.Label in OnGUI() showing the current lives value, updated every frame - a nice touch, though the full scoreboard-style UI is coming properly in Lesson 9.",
      "Sound: add two AudioClip fields to GameManager, assign the provided sfx_lifelost.wav and sfx_gameover.wav to them in the Inspector, and call AudioSource.PlayClipAtPoint() inside RegisterEscape() - the life-lost clip when a life is lost, the game-over clip instead the moment hasLost first becomes true.",
      "Structure: notice RegisterKill() and RegisterEscape() are two separate static methods, each called from exactly one script - this mirrors the exact separation that made GameManager safe from the Lesson 7 bug in the first place.",
    ],
    trouble: [
      "Lives decreases when you shoot an enemy: confirm RegisterEscape() is called only from Enemy.cs's Update(), never from Bullet.cs.",
      "\"Game Over\" never appears: confirm lives actually starts at 3 and RegisterEscape() is actually being called - add a temporary Debug.Log(lives) to check.",
      "Compile errors: confirm hasLost is declared as public static bool, matching hasWon's declaration style exactly.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same lives-tracking pattern works in a Script Graph instead of C#.",
      steps: [
        "On the Enemy prefab's graph, find the escape Branch from Lesson 7, right before its Destroy node. Add a \"Subtract\" node (Math/Scalar category) taking the Lives graph variable and 1, feeding into a \"Set Variable\" node targeting Lives - the node-graph equivalent of lives--.",
        "Feed Lives into a \"Less Than or Equal\" node (Comparison category) against 0, feeding a Branch node.",
        "On that Branch's True output, add an \"On GUI\" event node with a GUI Label node showing \"Game Over\", the same pattern already used for the win message.",
        "Confirm this new logic only touches the Lives graph variable, never the separate variable used for enemies-remaining/win-tracking.",
      ],
      level1Diagram: { file: "diagram_unity_visual_scripting_l8_level1.png", w: 6.3, h: 4.9, caption: "The Enemy prefab's Script Graph - the node-for-node equivalent of the Level 1 lives-tracking logic." },
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: ["Open your [YourName]Game project from Lesson 7."],
    level1: [
      "On obj_gamemanager's Create event, add __dnd_lives = 3; and has_lost = false; alongside your existing enemies_remaining and has_won. (__dnd_lives is the instance variable GameMaker's own drag-and-drop Lives blocks use - the \"dnd\" stands for drag-and-drop. Use it instead of the old built-in lives, which is deprecated, so code and blocks stay interchangeable.)",
      "On obj_enemy's Step event, find the escape-cleanup code you built last week. Right before instance_destroy(), add the lives decrement and zero-check, targeting obj_gamemanager (code below).",
      "On obj_gamemanager's existing Draw GUI event, add a second if block that draws \"Game Over\" when has_lost is true, right after the existing has_won block (full code below, new lines highlighted).",
      "Press F5 (Run) and let all 3 lives run out by letting enemies escape unshot - \"Game Over\" should appear the instant the third one escapes.",
    ],
    level1Code: [
      "// obj_gamemanager's Create event:",
      "enemies_remaining = 5;",
      "has_won = false;",
      "@@ __dnd_lives = 3;",
      "@@ has_lost = false;",
      "",
      "// obj_enemy's Step event:",
      "y += enemy_speed;",
      "if (y > room_height) {",
      "@@     obj_gamemanager.__dnd_lives -= 1;",
      "@@     if (obj_gamemanager.__dnd_lives <= 0) {",
      "@@         obj_gamemanager.has_lost = true;",
      "@@     }",
      "    instance_destroy();",
      "}",
      "",
      "// obj_gamemanager's Draw GUI event:",
      "if (has_won) {",
      "    draw_text(400, 50, \"You Win!\");",
      "}",
      "@@ if (has_lost) {",
      "@@     draw_text(400, 90, \"Game Over\");",
      "@@ }",
    ],
    level2: [
      "Shoot an enemy and confirm lives does NOT change - only obj_bullet's Collision event, and enemies_remaining, should be involved.",
      "Let an enemy escape and confirm enemies_remaining does NOT change - only obj_enemy's Step event, and lives, should be involved.",
      "Confirm \"Game Over\" appears exactly once, exactly when lives reaches zero.",
    ],
    level3: [
      "Visual: add a second draw_text() call in obj_gamemanager's Draw GUI event showing the current __dnd_lives value every frame (e.g. draw_text(20, 20, \"Lives: \" + string(__dnd_lives))) - a nice touch, though the full scoreboard-style UI is coming properly in Lesson 9. Visual path: a \"Draw Instance Lives\" block (Drawing category) on obj_gamemanager's Draw GUI event, choosing spr_player as the sprite, draws one small ship per life - or a \"Draw Value\" block with Caption \"Lives: \" and Value __dnd_lives.",
      "Sound: import the provided sfx_lifelost.wav and sfx_gameover.wav as new sound resources (keep the default names, sfx_lifelost and sfx_gameover), then call audio_play_sound() right where __dnd_lives decreases in obj_enemy's Step event - sfx_lifelost normally, sfx_gameover instead when lives reaches zero and has_lost becomes true.",
      "Structure: obj_enemy's Step event now handles movement, escape cleanup, AND reaching out to obj_gamemanager for lives - all genuinely the same concern (\"this enemy left the screen unshot\"), while obj_bullet's Collision event stays the only place enemies_remaining changes, and obj_gamemanager stays the only place any shared game state lives at all. Two clearly separated jobs, one dedicated manager, same as last week.",
    ],
    trouble: [
      "Lives decreases when you shoot an enemy: confirm the lives decrement is only in obj_enemy's Step event, not accidentally copied into obj_bullet's Collision event too.",
      "\"Game Over\" never appears: confirm __dnd_lives actually starts at 3 in obj_gamemanager's Create event (a Set Lives block with Value 3, Relative unchecked, in the Visual path), and that the Step event's threshold (y > room_height) is actually being reached.",
      "Visual path: Set Lives and If Lives both need their Target (drop-down arrow, upper right of the block) set to obj_gamemanager - left on the default, they change or check lives on the enemy itself, and Game Over never appears.",
      "Error referencing obj_gamemanager: confirm there's actually exactly one obj_gamemanager instance in the room - the same requirement enemies_remaining already relies on.",
      "\"Game Over\" and \"You Win!\" overlap on screen: adjust the Y coordinate in the has_lost draw_text() call (e.g. 90 instead of 50) so the two messages don't draw on top of each other.",
    ],
    resource: "Student Resource List > GameMaker section, or gamemaker.io/en/tutorials",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same lives-tracking pattern works with drag-and-drop blocks instead of typed code, targeting obj_gamemanager the same way the code path does. GameMaker's Lives blocks are purpose-built: they keep an instance variable named __dnd_lives (not a global), so each block needs its Target set to obj_gamemanager. The code path uses that same variable, so the two paths are interchangeable.",
      steps: [
        "On obj_gamemanager's Create event, add a \"Set Lives\" block (Instance Variables category) with Value 3 and Relative unchecked, then a Set Variable block setting has_lost to false (has_lost is a custom variable with no dedicated block).",
        "On obj_enemy's Step event, inside the existing If Variable block checking y against room_height, add a \"Set Lives\" block: click the Target drop-down arrow (upper right of the block) and choose obj_gamemanager, set Value to -1 and check Relative. Then add an \"If Lives\" block, also with Target set to obj_gamemanager, Expression set to less than or equal to, and Value 0. Drop a \"Set Variable\" block to the side of it (inside the If) setting has_lost to true on obj_gamemanager.",
        "On obj_gamemanager's existing Draw GUI event, add a second \"If Variable\" block checking has_lost, with a \"Draw Text\" block inside it - set its X to 400, Y to 90, and Text to \"Game Over\", matching the code path's exact values.",
      ],
      level1Diagram: { file: "diagram_gml_visual_l8_level1.png", w: 6.3, h: 5.75, caption: "obj_enemy's Step event and obj_gamemanager's Draw GUI event, in GML Visual - the block-for-block equivalent of the Level 1 code." },
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 7.",
      "Add a new Text object, name it LoseText, set its content to \"Game Over\", place it on screen, and hide it (uncheck its visibility).",
      "Add a new scene variable Lives, default value 3.",
    ],
    level1: [
      "Find the escape-cleanup event you built last week (Condition: Enemy.Y() > 750 -> Action: Delete Enemy). Add a new action to it: Lives = Lives - 1 (code below).",
      "Add a new event: Condition \"Lives <= 0\" -> Action \"Show\" LoseText.",
      "Click the green Preview button and let all 3 lives run out by letting enemies escape unshot - \"Game Over\" should appear the instant the third one escapes.",
    ],
    level1Code: [
      "(existing escape event, revised)",
      "Condition: Enemy.Y() > 750",
      "@@ Action: Lives = Lives - 1",
      "Action: Delete Enemy",
      "",
      "@@ (new event)",
      "@@ Condition: Lives <= 0",
      "@@ Action: Show LoseText",
    ],
    level2: [
      "Shoot an enemy and confirm Lives does NOT change - only the Combat group's collision event, and EnemiesRemaining, should be involved.",
      "Let an enemy escape and confirm EnemiesRemaining does NOT change - only the escape-cleanup event, and Lives, should be involved.",
      "Confirm LoseText appears exactly once, exactly when Lives reaches zero.",
    ],
    level3: [
      "Visual: add a second Text object showing the current Lives value, updated by the escape-cleanup event - a nice touch, though the full scoreboard-style UI is coming properly in Lesson 9.",
      "Sound: import the provided sfx_lifelost.wav and sfx_gameover.wav, then add a \"Play a sound\" action to the escape-cleanup event (sfx_lifelost.wav) and another to the Lives <= 0 check (sfx_gameover.wav).",
      "Structure: the escape-cleanup event now changes Lives, and the Combat group changes EnemiesRemaining - two clearly separate events, neither one referencing the other's variable. Keep it that way as the event sheet grows.",
    ],
    trouble: [
      "Lives decreases when you shoot an enemy: confirm the Lives = Lives - 1 action is only in the escape-cleanup event, not accidentally added to the Combat group too.",
      "\"Game Over\" never appears: confirm Lives actually starts at 3 (check the scene variable's default value) and that the Enemy.Y() > 750 condition is actually true for escaping enemies.",
      "LoseText and WinText overlap on screen: reposition one of the two Text objects in the scene editor so they don't draw on top of each other.",
    ],
    resource: "Student Resource List > GDevelop section, or gdevelop.io/tutorials",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: ["Launch TIC-80 and load your cart from Lesson 7.", "Press Esc to reach the code editor."],
    level1: [
      "Add lives=3 near the top of your code, alongside enemies_remaining.",
      "In the enemy-drawing loop's escape branch (where you set e.alive=false for reaching the bottom), add the lives decrement (code below).",
      "At the bottom of TIC(), add an if check that calls print(\"Game Over\", ...) when lives reaches 0, alongside your existing You Win! check.",
      "Press Esc, type run, press Enter, then let all 3 lives run out by letting enemies escape unshot - \"Game Over\" should appear the instant the third one escapes.",
    ],
    level1Code: [
      "@@ lives=3",
      "",
      "-- enemy movement/escape (drawing loop, revised):",
      "for enemy_index=1,#enemies do",
      " local enemy=enemies[enemy_index]",
      " if enemy.alive then",
      "  enemy.ey=enemy.ey+0.3",
      "  if enemy.ey>136 then",
      "   enemy.alive=false",
      "@@    lives=lives-1",
      "  else",
      "   spr(3,enemy.ex,enemy.ey,0)",
      "  end",
      " end",
      "end",
      "",
      "-- bottom of TIC(), alongside the You Win! check:",
      "@@ if lives<=0 then",
      "@@  print(\"Game Over\",40,60,8)",
      "end",
    ],
    level2: [
      "Shoot an enemy and confirm lives does NOT change - only the bullets loop's collision branch, and enemies_remaining, should be involved.",
      "Let an enemy escape and confirm enemies_remaining does NOT change - only the enemy-drawing loop's escape branch, and lives, should be involved.",
      "Confirm \"Game Over\" appears exactly once, exactly when lives reaches zero.",
    ],
    level3: [
      "Visual: add a second print() call showing the current lives value every frame - a nice touch, though the full scoreboard-style UI is coming properly in Lesson 9.",
      "Sound: import the provided sfx_lifelost.wav and sfx_gameover.wav into two new SFX slots (TIC-80's Surf/import tools accept an external .wav) rather than designing new effects from scratch, then call sfx(slot_number) in the escape branch - the life-lost slot normally, the game-over slot instead when lives reaches 0, guarded so it doesn't replay every frame.",
      "Structure: the two loops from Lesson 7 stay separate - the enemy-drawing loop only ever changes lives, and the bullets loop only ever changes enemies_remaining. Resist the urge to merge them even though both touch the enemies table.",
    ],
    trouble: [
      "Lives decreases when you shoot an enemy: confirm the lives decrement is only in the enemy-drawing loop's escape branch, not accidentally added to the bullets loop too.",
      "\"Game Over\" never appears: confirm lives actually starts at 3 and that the e.ey>136 threshold is actually being reached for escaping enemies.",
      "\"Game Over\" and \"You Win!\" overlap on screen: adjust the Y coordinate in one of the print() calls (e.g. 60 instead of 50) so the two messages don't draw on top of each other.",
    ],
    resource: "Student Resource List > TIC-80 section, or tic80.com",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 7.",
      "Create a new Variable (Resources > Variables/Switches/Collections/Tables) named lives, default value 3.",
    ],
    level1: [
      "On the enemy Actor Template's existing \"Left/Entered Camera\" trigger (the Left case, built last week for escape cleanup), add the lives decrement and zero-check BEFORE the existing Destroy event (code below).",
      "Save, then Play/Test and let all 3 lives run out by letting enemies escape unshot - \"Game Over\" should appear the instant the third one escapes.",
    ],
    level1Code: [
      "enemy Actor Trigger: Left/Entered Camera (Left case) -> (revised)",
      "@@   Variable – Modify Value(lives, Subtract, 1)",
      "@@   If(lives <= 0)",
      "@@     Message Box(\"Game Over\")",
      "  Actor – Destroy/Revive Actor(this, Destroy)",
    ],
    level2: [
      "Shoot an enemy and confirm lives does NOT change - only the \"Collided with actor\" trigger, and enemies_remaining, should be involved.",
      "Let an enemy escape and confirm enemies_remaining does NOT change - only the \"Left/Entered Camera\" trigger, and lives, should be involved.",
      "Confirm \"Game Over\" appears exactly once, exactly when lives reaches zero.",
    ],
    level3: [
      "Visual: add a Text/HUD display element showing the current lives value, updated inside the \"Left/Entered Camera\" trigger - a nice touch, though the full scoreboard-style UI is coming properly in Lesson 9.",
      "Sound: import the provided sfx_lifelost.wav and sfx_gameover.wav, then add a \"Play Sound\" event to the \"Left/Entered Camera\" trigger's script - sfx_lifelost.wav on each escape, sfx_gameover.wav instead right before the Message Box when lives reaches zero.",
      "Structure: the enemy Template now has exactly two Triggers, each doing one clear job - \"Collided with actor\" (a kill, changes enemies_remaining) and \"Left/Entered Camera\" (an escape, changes lives). Neither one references the other's variable. (Unlike Scratch/Godot/GameMaker/Unreal, 001 Game Creator's Actor Triggers are bound to the actor they're attached to, so lives stays on the enemy Template itself rather than moving to a separate \"GameManager\" Actor - the same documented exception enemies_remaining already relies on since Lesson 6.)",
    ],
    trouble: [
      "Lives decreases when you shoot an enemy: confirm the lives decrement is only in \"Left/Entered Camera,\" not accidentally added to \"Collided with actor\" too.",
      "\"Game Over\" never appears: confirm lives actually starts at 3 (check the Variable's default value) and that \"Left/Entered Camera\" is actually firing for the Left case, not Entered.",
      "\"Game Over\" and \"You Win!\" Message Boxes both pop up oddly: this can happen if an enemy escapes at the same moment as the last enemy is shot - not a bug, just an edge case worth a quick laugh with students if it comes up.",
    ],
    resource: "001gamecreator.com/wiki, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 7.",
      "Open BP_Enemy and add a second Event Dispatcher (Details panel under My Blueprint > Event Dispatchers, click +) named OnEnemyEscaped.",
      "In BP_GameManager - the placed Actor introduced in Lesson 6 that already holds EnemiesRemaining - add a second Integer variable Lives, default value 3.",
    ],
    level1: [
      "In BP_Enemy's Event Graph, find the escape Branch you built last week. Right before its Destroy Actor node, add a \"Call OnEnemyEscaped\" node (code below).",
      "In the Level Blueprint's spawn For Loop, right where you bind OnEnemyDestroyed to your BP_GameManager reference, also add a \"Bind Event to OnEnemyEscaped\" node targeting the same spawned Actor, wired to the same BP_GameManager reference's new HandleEnemyEscaped custom event.",
      "In BP_GameManager, build HandleEnemyEscaped: decrease Lives by 1, Branch on Lives <= 0, and on True, add a Print String node with text \"Game Over\" and \"Print to Screen\" checked.",
      "Click Compile, then Play and let all 3 lives run out by letting enemies escape unshot - \"Game Over\" should appear the instant the third one escapes.",
    ],
    level1Code: [
      "BP_Enemy's Event Graph (escape Branch, revised):",
      "Event Tick ->",
      "  Branch(GetActorLocation().Y < -400) -> True:",
      "@@     Call OnEnemyEscaped (Broadcast)",
      "    Destroy Actor(Self)",
      "",
      "Level Blueprint spawn loop (alongside binding OnEnemyDestroyed to the same BP_GameManager reference):",
      "@@   Bind Event to OnEnemyEscaped(Target = spawned Actor, Event = BP_GameManager ref's HandleEnemyEscaped)",
      "",
      "BP_GameManager:",
      "@@ Custom Event HandleEnemyEscaped ->",
      "@@   Lives = Lives - 1",
      "@@   Branch(Lives <= 0) -> True: Print String(\"Game Over\", Print to Screen = true)",
    ],
    level2: [
      "Shoot an enemy and confirm Lives does NOT change - only OnEnemyDestroyed, and EnemiesRemaining, should be involved.",
      "Let an enemy escape and confirm EnemiesRemaining does NOT change - only OnEnemyEscaped, and Lives, should be involved.",
      "Confirm \"Game Over\" prints exactly once, exactly when Lives reaches zero.",
    ],
    level3: [
      "Visual: add a Print String in HandleEnemyEscaped showing the current Lives value every time it changes - a nice touch, though the full UMG-based UI is coming properly in Lesson 9.",
      "Sound: import the provided sfx_lifelost.wav and sfx_gameover.wav as Sound Wave assets, then add a Play Sound 2D node in HandleEnemyEscaped - sfx_lifelost.wav right after Lives decreases, sfx_gameover.wav on the Branch's True output before the Print String node.",
      "Structure: OnEnemyDestroyed and OnEnemyEscaped are two separate dispatchers, each broadcast from exactly one place in BP_Enemy, each bound to its own handler on BP_GameManager - the same separation pattern used since Lesson 7, just doubled, and both handlers still living on the one dedicated manager Actor introduced in Lesson 6.",
    ],
    trouble: [
      "Lives decreases when you shoot an enemy: confirm OnEnemyEscaped is only broadcast from the escape Branch, never from the successful-Cast (hit) branch.",
      "\"Game Over\" never prints: confirm the Level Blueprint's spawn loop actually binds OnEnemyEscaped for every spawned Actor, targeting BP_GameManager's HandleEnemyEscaped - not just OnEnemyDestroyed, and not a Custom Event on the Level Blueprint itself.",
      "Both messages print at the same screen position: adjust HandleEnemyEscaped's Print String Text Position/Duration slightly so it doesn't overlap the win message.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    level1EngineDiagram: { file: "diagram_unreal_blueprint_l8_level1.png", w: 6.3, h: 4.45, caption: "BP_Enemy's escape Branch, a new dispatcher, and BP_GameManager's new custom event, in Unreal Blueprint - the node-for-node equivalent of the Level 1 code." },
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same escape-vs-hit-dispatcher pattern works in C++ instead of Blueprint nodes.",
      steps: [
        "In ACppEnemy.h, declare a second dispatcher-equivalent delegate, e.g. DECLARE_MULTICAST_DELEGATE(FOnEnemyEscaped), alongside the existing OnEnemyDestroyed delegate.",
        "In ACppEnemy::Tick()'s escape branch, call OnEnemyEscaped.Broadcast() right before Destroy(), mirroring how OnEnemyDestroyed.Broadcast() is called from OnOverlap().",
      ],
      code: [
        "// ACppEnemy.cpp - Tick()'s escape branch (revised):",
        "if (GetActorLocation().Y < -400.f) {",
        "    OnEnemyEscaped.Broadcast();",
        "    Destroy();",
        "}",
      ],
      note: "This mirrors the Blueprint path's escape Branch - same underlying broadcast-then-destroy shape, just from code.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_08_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 8 Teacher's Key: Under Pressure", bold: true, size: 30, color: brand.NAVY })] }),
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
    ...(data.level3Code ? [p("Level 3 - complete code (all three goals combined):", { bold: true }), code(data.level3Code, "Highlighted lines are the Level 3 additions on top of the Level 1 code.")] : []),
    h2("Common Student Mistakes"),
    ...data.mistakes.map(s => bullet(s)),
  );
  if (data.altNotes) {
    children.push(h2("Alternate Path Notes: " + data.altNotes.title));
    children.push(...data.altNotes.notes.map(s => p(s)));
    if (data.altNotes.code) children.push(code(data.altNotes.code));
    children.push(...data.altNotes.why.map(s => bullet(s)));
  }
  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 8 Key - " + data.engine + " (Instructor Only)") },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    teaches: "Extending GameManager - the dedicated sprite introduced in Lesson 6 for win-tracking - with a second, independent counter for a lose condition, reusing the same broadcast/receive pattern and the escape branch already built in Lesson 7.",
    level1: [
      "lives is a second variable living on GameManager, exactly parallel to enemies_remaining - same sprite, same broadcast-based announcement pattern, different broadcast name and different trigger condition.",
      "The Enemy's escape branch (y position < -170) now broadcasts enemy escaped instead of touching lives directly - the Enemy only ever reports what happened, it never counts or decides anything, exactly like the touching-Bullet branch already does for enemy hit.",
      "GameManager's new \"when I receive [enemy escaped]\" script owns the decrement, the zero-check, and the game-over broadcast - a direct parallel to its existing \"when I receive [enemy hit]\" script.",
      "\"lives = 0\" (exact equality) mirrors \"enemies_remaining = 0\" from Lesson 6 exactly - the same reasoning applies: a variable that only ever decreases by 1 can only cross zero once, so exact equality is safe and simpler than <=.",
    ],
    why: [
      "This is a deliberate payoff for good Lesson 6/7 structure: because GameManager already owns all counting and decision-making, and the escape branch was already cleanly separated from the touching-Bullet branch, adding a second consequence required only one new broadcast and one new receiving script - zero risk of accidentally touching enemies_remaining.",
      "Using broadcast/receive again, with GameManager as the single listener, keeps this lesson consistent with Lesson 6's win-announcement pattern and Lesson 7's Rule 13 retrofit - students are extending a tool and a structure they already trust, not learning either one fresh.",
    ],
    talkingPoints: [
      "Level 3 sound, worked (Scratch): import the provided sfx_gameover.wav and sfx_lifelost.wav. On the Player, \"when I receive [game over]\" becomes: start sound [sfx_gameover] -> say [Game Over] for 3 seconds. On GameManager, \"when I receive [enemy escaped]\" becomes: change lives by -1 -> start sound [sfx_lifelost] -> if lives = 0 then broadcast [game over]. Expect both sounds to overlap on the final life - acceptable, or students can wrap the lifelost sound in an if lives > 0.",
      "Ask students to point out exactly which broadcast the Enemy sends on escape, and which script on GameManager receives it - checking their own reasoning out loud catches the most common mistake (touching lives from the wrong place) before it happens.",
      "Level 3 visual: the built-in variable-watcher checkbox is worth calling out explicitly as \"the Scratch equivalent of draw_text() or print()\" - genuinely the simplest possible way to show a running number, no blocks needed at all.",
    ],
    mistakes: [
      "Adding a lives-changing block directly to the Enemy sprite instead of broadcasting enemy escaped and letting GameManager handle it - even if the number comes out right, this bypasses the structure the rest of the course relies on.",
      "Broadcasting enemy escaped from the touching-Bullet branch instead of the escape branch (makes shooting cost a life, which is backwards).",
      "Using \"lives < 1\" or forgetting the check entirely and just broadcasting every frame lives is low.",
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    teaches: "A second, independent signal (escaped) added alongside an existing one (destroyed) - the cleanest possible illustration in the course of two distinct event types living side by side without interfering, extending the exact pattern that made Godot immune to the Lesson 7 bug.",
    level1: [
      "signal escaped is declared in enemy.tscn's script alongside the pre-existing signal destroyed - same mechanism, different name, different meaning.",
      "escaped.emit() is added to the EXISTING escape branch in _process(), right before the pre-existing queue_free() call - one new line in an already-correct location.",
      "lives and on_enemy_escaped() are added to GameManager's script, alongside its existing enemies_remaining and on_enemy_destroyed() - same node, same pattern, different signal.",
      "The main scene connects BOTH signals per spawned enemy, both targeting GameManager (destroyed to $GameManager.on_enemy_destroyed from Lesson 6, escaped to $GameManager.on_enemy_escaped new this week) - two independent connections, two independent handler functions, one dedicated manager node.",
    ],
    why: [
      "Godot's signal system makes this lesson's core idea almost visually obvious: two separate signals, two separate handlers, two separate counters, all owned by the same GameManager node - there's no shared hook to misuse the way GameMaker's Destroy event or 001GC's Defeated trigger could be misused.",
      "This is worth explicitly naming as the payoff for Godot's Lesson 6/7 structure: because GameManager already owned all counting, and escape-cleanup and hit-detection were already fully separate code paths, adding a second signal and a second handler required zero changes to either existing one.",
    ],
    talkingPoints: [
      "Level 3 sound, worked (Godot): GameManager gets two AudioStreamPlayer children, LifeLostSound (sfx_lifelost.wav) and GameOverSound (sfx_gameover.wav). In on_enemy_escaped(): lives -= 1, then if lives <= 0: get_node(\"../LoseText\").show() and $GameOverSound.play(), else: $LifeLostSound.play(). Plain AudioStreamPlayer, not 2D, since this is UI-level audio.",
      "Ask students what would happen if _on_hitbox_area_entered() emitted BOTH destroyed and escaped by mistake (a copy-paste error) - a good diagnostic exercise, since the symptom (shooting costs a life) is confusing without understanding the two-signal structure.",
      "Level 3 structure: point out that a project could keep adding new signals this way indefinitely (poisoned, healed, upgraded, etc.), all still handled on GameManager, without ever needing to touch existing signal-handling code - this is a preview of how larger, real projects stay maintainable.",
    ],
    mistakes: [
      "Emitting escaped from _on_hitbox_area_entered() by mistake, or destroyed from _process()'s escape branch by mistake - either one causes cross-contamination between lives and enemies_remaining.",
      "Writing on_enemy_escaped() on the main scene's own script instead of GameManager's - even if the logic is otherwise correct, this splits counting logic across two different owners instead of keeping it on the one dedicated node.",
      "Forgetting to connect the new escaped signal in the spawn loop (only destroyed connected, copied from Lesson 6 without adding the new line) - lives never changes even though the signal fires correctly.",
      "LoseText and WinText positioned at the same screen coordinates, causing them to overlap if both happen to show near the same time.",
    ],
  },
  {
    engine: "Unity", file: "Unity",
    teaches: "A second static method (RegisterEscape()) added to GameManager alongside the existing RegisterKill() - reinforcing that a single manager class can safely host multiple, clearly-named, single-purpose methods without them interfering.",
    level1: [
      "lives and hasLost are new static fields, parallel to enemiesRemaining and hasWon - same class, same pattern, different names.",
      "RegisterEscape() is called from the EXISTING escape branch in Enemy.cs's Update(), right before the pre-existing Destroy(gameObject) call - one new line in an already-correct location.",
      "OnGUI() gets a second if-block for hasLost, alongside the existing hasWon block - two independent checks, two independent messages.",
    ],
    why: [
      "This is the direct payoff of Lesson 7's fix: RegisterKill() lives exclusively in Bullet.cs, and now RegisterEscape() lives exclusively in Enemy.cs's own Update() - two methods, two call sites, each unambiguous about what it means.",
      "Keeping both methods on the same GameManager class (rather than splitting into multiple manager scripts) is fine here because the class itself isn't doing the deciding - each calling script decides which method applies to its own situation.",
    ],
    talkingPoints: [
      "Level 3 sound, worked (Unity): GameManager.cs gets public AudioClip lifeLostSound and gameOverSound, assigned in the Inspector from the provided sfx_lifelost.wav / sfx_gameover.wav. Because RegisterEscape() is static and the clips are instance fields, GameManager also needs a static instance reference set in Awake() (see the Level 3 code below). In RegisterEscape(): lives--, then if (lives <= 0 && !hasLost) { hasLost = true; AudioSource.PlayClipAtPoint(gameOverSound, Camera.main.transform.position); } else AudioSource.PlayClipAtPoint(lifeLostSound, Camera.main.transform.position). Visual Scripting path: the same two Play Clip At Point-style calls on the Branch's True/False outputs.",
      "Ask students: what's the actual difference between RegisterKill() and RegisterEscape(), code-wise? (One decrements enemiesRemaining and sets hasWon; the other decrements lives and sets hasLost - genuinely parallel, unrelated pieces of state.)",
      "Level 3 structure talking point: this GameManager class is starting to look like a real \"game state\" hub - ask students to predict what else might eventually live here (score, wave number, difficulty) as a preview of Lesson 9 and beyond.",
    ],
    mistakes: [
      "Calling RegisterEscape() from Bullet.cs, or RegisterKill() from Enemy.cs - either one causes cross-contamination between lives and enemiesRemaining.",
      "Declaring hasLost as a non-static field by mistake (compiles fine but never updates the way hasWon does), copied incorrectly from the static hasWon declaration.",
      "OnGUI() messages overlapping on screen if hasWon and hasLost happen to become true around the same time - remind students this is a display/layout detail, not a logic bug.",
    ],
    level1EngineDiagram: { file: "diagram_unity_visual_scripting_l8_level1.png", w: 6.3, h: 4.9, caption: "The Enemy prefab's Script Graph - the node-for-node equivalent of the Level 1 lives-tracking logic." },
    altNotes: {
      title: "Visual Scripting",
      notes: [
        "The Enemy prefab's graph gets a second branch off its existing escape Branch node: a \"Subtract\" node (Math/Scalar category) computing Lives minus 1, feeding a \"Set Variable\" node that writes it back to the Lives graph variable - the node-graph equivalent of lives--.",
        "Lives feeds a \"Less Than or Equal\" node (Comparison category) against 0, feeding a second Branch node. Its True output feeds an \"On GUI\" event node with a GUI Label showing \"Game Over\" - the exact same pattern Lesson 7's win message already used, just gated on a different variable.",
        "Lives is its own graph variable, entirely separate from the enemies-remaining/win-tracking variable - the same separation RegisterKill()/RegisterEscape() enforce in the C# path, just expressed as two variables instead of two methods.",
      ],
      why: [
        "This mirrors the C# path's structure exactly: one new decrement-and-check sequence, attached to the same escape Branch that already existed, touching only the Lives variable.",
        "Point out to students moving between C# and Visual Scripting that \"static method call\" and \"node sequence wired to an event\" are two different surfaces for the identical idea: a single, named place where a specific consequence happens.",
      ],
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    teaches: "Extending an already-correct code path (obj_enemy's Step event, fixed in Lesson 7 to only handle movement/escape) with a second consequence routed through obj_gamemanager - reinforcing exactly why last week's fix mattered, and why Rule 13's dedicated manager object earns its keep the moment a second piece of shared state shows up.",
    level1: [
      "__dnd_lives and has_lost are added to obj_gamemanager's Create event alongside enemies_remaining and has_won - same object, same pattern, different variables, all still living on the one dedicated manager instance introduced in Lesson 6.",
      "The lives decrement and zero-check are added directly inside the EXISTING escape branch of obj_enemy's Step event (the if (y > room_height) block), right before instance_destroy() - not a new event, just new lines in an existing, already-correct one, now reaching out to obj_gamemanager instead of obj_spawner.",
      "obj_gamemanager's existing Draw GUI event gets a second if-block for has_lost, alongside the existing has_won block - the same object that already draws \"You Win!\" now also draws \"Game Over\".",
    ],
    why: [
      "This is the most direct possible payoff of Lesson 7's bugfix AND this week's Rule 13 retrofit: because obj_enemy's Step event was already the single, unambiguous home for \"this enemy left the screen unshot,\" and obj_gamemanager was already the single home for shared game state, adding a new consequence required touching nothing else.",
      "Contrast this with what would have happened before the Rule 13 retrofit, when lives briefly lived on obj_spawner: a student following last week's habits would have had two different objects to check depending on which piece of state they wanted, instead of one.",
    ],
    talkingPoints: [
      "Level 3 sound, worked (GameMaker): import sfx_lifelost.wav and sfx_gameover.wav keeping their default names, sfx_lifelost and sfx_gameover. Where obj_enemy's Step event does obj_gamemanager.__dnd_lives -= 1, follow with: if (obj_gamemanager.__dnd_lives <= 0) { obj_gamemanager.has_lost = true; audio_play_sound(sfx_gameover, 1, false); } else { audio_play_sound(sfx_lifelost, 1, false); }. Visual path: a Play Sound block inside the If Lives block (game over) and another after it (life lost) - use If Lives' Not option or a second If Lives (Expression greater than, Value 0) so only one sound plays.",
      "This is a great lesson to explicitly connect back to last week: ask students to recall why the GameMaker fix mattered, and now show them it's the reason this week was easy.",
      "Level 3 structure: obj_bullet's Collision event and obj_enemy's Step event remain the entire win/lose logic surface of the game so far, and obj_gamemanager remains the only object holding any of that state - worth pointing out how compact and consistent that still is, even as features accumulate.",
    ],
    mistakes: [
      "Adding the lives decrement to obj_bullet's Collision event instead of obj_enemy's Step event (makes shooting cost a life, which is backwards).",
      "Adding __dnd_lives/has_lost to obj_spawner instead of obj_gamemanager, out of habit from before the Lesson 7 GameManager retrofit - watch for this specifically, since it's the most likely regression this week.",
      "Forgetting the has_lost flag and just checking __dnd_lives <= 0 directly in Draw GUI without ever setting a flag (works, but inconsistent with the has_won pattern already established - point this out as a style/consistency issue, not a bug).",
      "\"Game Over\" and \"You Win!\" text overlapping on screen because both draw_text() calls use the same Y coordinate.",
    ],
    level1Code: [
      "// obj_gamemanager's Create event:",
      "enemies_remaining = 5;",
      "has_won = false;",
      "@@ __dnd_lives = 3;",
      "@@ has_lost = false;",
      "",
      "// obj_enemy's Step event:",
      "y += enemy_speed;",
      "if (y > room_height) {",
      "@@     obj_gamemanager.__dnd_lives -= 1;",
      "@@     if (obj_gamemanager.__dnd_lives <= 0) {",
      "@@         obj_gamemanager.has_lost = true;",
      "@@     }",
      "    instance_destroy();",
      "}",
      "",
      "// obj_gamemanager's Draw GUI event:",
      "if (has_won) {",
      "    draw_text(400, 50, \"You Win!\");",
      "}",
      "@@ if (has_lost) {",
      "@@     draw_text(400, 90, \"Game Over\");",
      "@@ }",
    ],
    level2Code: [
      "// Level 2 is verification only - the Level 1 code is already correct, so no lines change.",
      "// Checks: shooting an enemy never touches obj_gamemanager.__dnd_lives (only obj_bullet's Collision event",
      "// changes enemies_remaining); an escape never touches enemies_remaining; \"Game Over\" draws once, when",
      "// __dnd_lives first reaches 0.",
    ],
    level3Code: [
      "// obj_gamemanager's Create event:",
      "enemies_remaining = 5;",
      "has_won = false;",
      "__dnd_lives = 3;",
      "has_lost = false;",
      "",
      "// obj_enemy's Step event (Structure: still the only place lives change; Sound: both clips):",
      "y += enemy_speed;",
      "if (y > room_height) {",
      "    obj_gamemanager.__dnd_lives -= 1;",
      "    if (obj_gamemanager.__dnd_lives <= 0) {",
      "        obj_gamemanager.has_lost = true;",
      "@@         audio_play_sound(sfx_gameover, 1, false);",
      "@@     } else {",
      "@@         audio_play_sound(sfx_lifelost, 1, false);",
      "    }",
      "    instance_destroy();",
      "}",
      "",
      "// obj_gamemanager's Draw GUI event (Visual: live lives readout, drawn every frame):",
      "@@ draw_text(20, 20, \"Lives: \" + string(__dnd_lives));",
      "if (has_won) {",
      "    draw_text(400, 50, \"You Win!\");",
      "}",
      "if (has_lost) {",
      "    draw_text(400, 90, \"Game Over\");",
      "}",
      "",
      "// Visual path alternative for the readout: replace the first draw_text() with a Draw Instance Lives block",
      "// (spr_player, Caption \"Lives: \", X 20, Y 20) or a Draw Value block (Caption \"Lives: \", Value __dnd_lives).",
    ],
    level1EngineDiagram: { file: "diagram_gml_visual_l8_level1.png", w: 6.3, h: 5.75, caption: "obj_enemy's Step event and obj_gamemanager's Draw GUI event, in GML Visual - the block-for-block equivalent of the Level 1 code." },
    altNotes: {
      title: "GML Visual (no typing)",
      notes: [
        "obj_enemy's Step event works identically whether it holds blocks or code - the event itself already fires exactly once per instance either way, and reaches out to obj_gamemanager the same way in both paths.",
        "The block path uses GameMaker's purpose-built Lives blocks - \"Set Lives\" (Relative checked, Value -1) and \"If Lives\" (Expression less than or equal to, Value 0) - each with its Target drop-down set to obj_gamemanager. Behind the scenes these keep an instance variable named __dnd_lives (the \"dnd\" is drag-and-drop; it is NOT a global). The code path deliberately uses that exact variable, so a student can mix code and blocks in one project - which is why the Target matters: left at the default, the blocks would act on the enemy that is about to be destroyed. has_lost, a custom variable with no dedicated block, still uses the generic \"Set Variable\" block, exactly as the code path's plain assignment does.",
        "Create-event setup mirrors the code path: a \"Set Lives\" block with Value 3 and Relative unchecked on obj_gamemanager's Create event.",
        "Level 3 visual: \"Draw Instance Lives\" (Drawing category, Draw GUI event, spr_player as the sprite) is the purpose-built way to show lives as icons; \"Draw Value\" with Value __dnd_lives is the text alternative. Draw Instance Lives only draws frame 0 of the sprite and ignores image_xscale/image_blend - fine for spr_player.",
        "obj_gamemanager's Draw GUI event gets the same second \"If Variable\" + \"Draw Text\" block pair whether Lesson 7's win check was built in Code or Visual - same object, same event, same targeting.",
      ],
      why: [
        "Keeping the same Structure principle (obj_gamemanager owns all shared state, obj_enemy only ever reaches out to modify it) whether a student is in Code or Visual reinforces that Rule 13 is a design idea, not a syntax rule tied to one coding style.",
        "A student switching between Code and Visual mid-course (or comparing notes with a classmate using the other track) should recognize obj_gamemanager playing the exact same role either way.",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    teaches: "Adding a new action to an existing event (the Lesson 7 escape-cleanup event) rather than creating a new event - reinforcing that GDevelop's event-based model lets a single trigger drive multiple, independent consequences cleanly.",
    level1: [
      "Lives is a new scene variable, parallel to EnemiesRemaining - same scope, same pattern, different name.",
      "The Lives = Lives - 1 action is added directly to the EXISTING escape-cleanup event from Lesson 7 (Enemy.Y() > 750 -> Delete Enemy), alongside the Delete action - one new action line in an already-correct event.",
      "A new, separate event (Lives <= 0 -> Show LoseText) mirrors the existing EnemiesRemaining <= 0 -> Show WinText event exactly.",
    ],
    why: [
      "GDevelop's per-event action lists make this lesson's core idea very visible: one event (the escape condition) now has two actions (delete the enemy, AND decrement Lives), while the Combat group's collision event remains untouched and still only ever changes EnemiesRemaining.",
      "This is the direct payoff of good Lesson 7 structure: the escape-cleanup event already existed as a clean, single-purpose event, so this week's addition was one new action line, not a redesign.",
    ],
    talkingPoints: [
      "Level 3 sound, worked (GDevelop): import both files. In the escape-cleanup event, after Lives = Lives - 1, add Action: Play a sound (sfx_lifelost.wav). In the event with Condition Lives <= 0, add Action: Play a sound (sfx_gameover.wav). Make the lifelost event only fire when Lives > 0 if students don't want the overlap.",
      "Ask students to open their event sheet and point to the exact event they added the new action to - confirming they can distinguish it from the Combat group is the real test of understanding, not just getting the game to work.",
      "Level 3 structure: suggest renaming the escape-cleanup event group to something like \"Enemy Movement & Escape\" if it isn't already labeled, continuing the organizational habit introduced in Lesson 7.",
    ],
    mistakes: [
      "Adding Lives = Lives - 1 to the Combat group's collision event instead of the escape-cleanup event (makes shooting cost a life, which is backwards).",
      "Forgetting Lives starts at 3 (scene variables default to 0 unless explicitly set) - check the variable's default value in the Scene Variables editor, not just referenced in an event.",
      "LoseText and WinText positioned at the same coordinates on the scene, causing them to overlap if both are shown near the same time.",
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    teaches: "Adding a second piece of state (lives) that changes inside an existing loop's existing branch, alongside a first (enemies_remaining) that changes in a completely different loop - reinforcing the two-separate-loops structure introduced in Lesson 7.",
    level1: [
      "lives=3 is declared near the top, parallel to enemies_remaining - same pattern, different variable, different loop responsible for changing it.",
      "The lives decrement is added directly inside the EXISTING escape branch of the enemy-drawing loop (where e.alive=false already happens for reaching the bottom) - one new line in an already-correct branch.",
      "The Game Over print() check at the bottom of TIC() mirrors the existing You Win! check exactly, just testing lives instead of enemies_remaining.",
    ],
    why: [
      "This is the direct payoff of Lesson 7's two-loop structure: the enemy-drawing loop was already the single, unambiguous home for \"this enemy reached the bottom,\" so lives could be added there with zero risk of touching the separate bullets-vs-enemies collision loop.",
      "Worth explicitly re-stating the tradeoff from Lesson 7 here: keeping the two loops separate (rather than merging them) is exactly what made this week's addition a one-line change instead of a careful untangling.",
    ],
    talkingPoints: [
      "Level 3 sound, worked (TIC-80): import both .wav files into two SFX slots (e.g. 1 and 2). In the enemy-drawing loop's escape branch, after lives=lives-1: if lives<=0 then (if not has_lost then has_lost=true; sfx(2) end) else sfx(1) end. The has_lost guard stops the game-over sound replaying every frame.",
      "Ask advanced students to trace, out loud, which loop owns which variable now - the enemy-drawing loop owns lives, the bullets loop owns enemies_remaining - as a way of testing whether the Lesson 7 structure actually stuck.",
      "Level 3 structure: this is a good moment to ask whether TIC-80's cart is starting to feel like it needs more organization (functions instead of one long TIC()) - a preview of concerns that become more pressing as the capstone grows through Lesson 11.",
    ],
    mistakes: [
      "Adding the lives decrement inside the bullets loop instead of the enemy-drawing loop (makes shooting cost a life, which is backwards).",
      "Forgetting lives=3 is declared before TIC() runs (a global, not something reset each frame) - a common Lua scoping confusion worth checking directly if lives behaves oddly.",
      "\"Game Over\" and \"You Win!\" print() calls overlapping on screen because both use the same Y coordinate.",
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    teaches: "Adding a second consequence to an existing Trigger (\"Left/Entered Camera,\" built in Lesson 7 for escape cleanup) - reinforcing exactly why last week's fix (splitting into two specific Triggers) mattered: this week's addition was trivial to make correctly BECAUSE the escape Trigger was already isolated from the hit Trigger.",
    level1: [
      "lives is a new Variable, parallel to enemies_remaining - same pattern, different variable, different Trigger responsible for changing it.",
      "The lives decrement and Message Box are added directly inside the EXISTING 'Left/Entered Camera' Trigger from Lesson 7, BEFORE the existing Destroy event - new Events added to an already-correct Trigger, not a new Trigger.",
      "'Collided with actor' remains completely untouched this week - it's still the only place enemies_remaining changes.",
    ],
    why: [
      "This is the most direct possible payoff of Lesson 7's fix: because 'Left/Entered Camera' was already the single, unambiguous home for 'this enemy left the screen unshot,' adding a new consequence to it required touching nothing else.",
      "Contrast this with what would have happened if Lesson 7's fix hadn't been made (i.e. if 'Defeated / Destroyed' were still doing double duty): this week's lives logic would have had to be untangled from the old broken win-counter at the same time as being added.",
    ],
    talkingPoints: [
      "Level 3 sound, worked (001 Game Creator): import both files. In the enemy Template's \"Left/Entered Camera\" Trigger (Left case): after the lives change, add Play Sound (sfx_lifelost.wav); inside the lives = 0 check, add Play Sound (sfx_gameover.wav) before the Message Box.",
      "This is a great lesson to explicitly connect back to last week: ask students to recall why the 001 Game Creator fix mattered, and now show them it's the reason this week was easy.",
      "Level 3 structure: the enemy Actor Template now has exactly two Triggers doing the entire win/lose logic of the game so far - worth pointing out how compact that still is, even as features accumulate.",
    ],
    mistakes: [
      "Adding the lives decrement to 'Collided with actor' instead of 'Left/Entered Camera' (makes shooting cost a life, which is backwards).",
      "Forgetting lives starts at 3 (check the Variable's default value in the Variables list, not just referenced in an Event).",
      "\"Game Over\" and \"You Win!\" Message Boxes both appearing if an enemy escapes at the same moment the last enemy is shot - an edge case, not a bug, worth a quick acknowledgment if a student hits it.",
    ],
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    teaches: "A second Event Dispatcher (OnEnemyEscaped) added to BP_Enemy alongside the existing OnEnemyDestroyed, with its handler living on BP_GameManager - the Blueprint-native version of the same two-signal / two-dispatcher pattern seen in Godot and Unity this lesson, reinforcing that both the pattern AND the dedicated-manager structure are cross-engine principles, not engine-specific tricks.",
    level1: [
      "Lives is a new variable on BP_GameManager, parallel to EnemiesRemaining - same placed Actor, same pattern, different variable, different dispatcher responsible for changing it.",
      "Call OnEnemyEscaped is added to the EXISTING escape Branch in BP_Enemy's Event Graph, right before the pre-existing Destroy Actor node - one new node in an already-correct location.",
      "The Level Blueprint's spawn loop binds BOTH dispatchers per spawned enemy to the SAME BP_GameManager reference (OnEnemyDestroyed to HandleEnemyDestroyed from Lesson 6, OnEnemyEscaped to HandleEnemyEscaped new this week) - two independent bindings, one dedicated manager Actor holding both handlers.",
    ],
    why: [
      "This is the same principle as Godot's two-signal design and Unity's two-static-method design, expressed in Unreal's dispatcher/delegate idiom: two separate broadcasts, two separate handlers, two separate pieces of state, all owned by BP_GameManager, with no shared hook to misuse.",
      "This is the direct payoff of Lesson 7's structure AND this week's Rule 13 retrofit: because the escape Branch was already the single, unambiguous home for 'this enemy left the play area unshot,' and BP_GameManager was already the single home for shared game state, adding a new dispatcher call and a new handler required zero changes to the overlap event's hit-detection logic or to EnemiesRemaining.",
    ],
    talkingPoints: [
      "Level 3 sound, worked (Unreal): import both files as Sound Wave assets. In BP_GameManager's HandleEnemyEscaped: Lives = Lives - 1 -> Play Sound 2D(sfx_lifelost.wav) -> Branch(Lives <= 0) -> True: Play Sound 2D(sfx_gameover.wav) -> Print String(\"Game Over\"). C++ path: UGameplayStatics::PlaySound2D with the same two assets.",
      "For students who did both Godot and Unreal (or who are comparing engines): ask them to name the equivalent piece in each engine (signal/emit vs. dispatcher/broadcast, GameManager node vs. BP_GameManager Actor) - a good vocabulary-building exercise that reinforces the underlying concepts are engine-independent.",
      "Level 3 structure: point out that BP_Enemy now has exactly two dispatchers, cleanly named for what they mean (Destroyed, Escaped), and BP_GameManager now holds exactly two pieces of shared state (EnemiesRemaining, Lives) - contrast this with a hypothetical single generic 'OnEnemyGone' dispatcher or a Lives variable left on the Level Blueprint, and ask why either would be worse.",
    ],
    mistakes: [
      "Calling OnEnemyEscaped from the successful-Cast (hit) branch, or OnEnemyDestroyed from the escape Branch - either one causes cross-contamination between Lives and EnemiesRemaining.",
      "Adding Lives and HandleEnemyEscaped to the Level Blueprint instead of BP_GameManager, out of habit from before the Lesson 7 GameManager retrofit - watch for this specifically, since it's the most likely regression this week.",
      "Forgetting to bind OnEnemyEscaped in the Level Blueprint's spawn loop (only OnEnemyDestroyed bound, copied from Lesson 6 without adding the new binding node) - Lives never changes even though the dispatcher fires correctly.",
      "Both Print String messages appearing at the same screen position - adjust one's position/duration so \"Game Over\" and \"You Win!\" don't overlap.",
    ],
    level1EngineDiagram: { file: "diagram_unreal_blueprint_l8_level1.png", w: 6.3, h: 4.45, caption: "BP_Enemy's escape Branch, a new dispatcher, and BP_GameManager's new custom event, in Unreal Blueprint - the node-for-node equivalent of the Level 1 code." },
    altNotes: {
      title: "C++ (code, optional advanced track)",
      notes: [
        "ACppEnemy.h declares FOnEnemyEscaped as a second multicast delegate, alongside the existing FOnEnemyDestroyed - same mechanism, different name, different meaning.",
        "ACppEnemy::Tick()'s escape branch calls OnEnemyEscaped.Broadcast() right before Destroy() - one new line in an already-correct location, mirroring how OnEnemyDestroyed.Broadcast() is called from OnOverlap().",
        "The GameManager actor class's BeginPlay() binds a second UFUNCTION to each spawned enemy's FOnEnemyEscaped delegate using AddDynamic(), tracking a new Lives member variable alongside the existing EnemiesRemaining - the same class, the same binding pattern, a second piece of state.",
      ],
      code: [
        "// ACppEnemy.h - add alongside the existing delegate:",
        "DECLARE_MULTICAST_DELEGATE(FOnEnemyEscaped);",
        "FOnEnemyEscaped OnEnemyEscaped;",
        "",
        "// ACppEnemy.cpp - Tick()'s escape branch (revised):",
        "if (GetActorLocation().Y < -400.f) {",
        "    OnEnemyEscaped.Broadcast();",
        "    Destroy();",
        "}",
        "",
        "// GameManager's handler:",
        "int32 Lives = 3;",
        "",
        "void AGameManager::HandleEnemyEscaped() {",
        "    Lives--;",
        "    if (Lives <= 0) {",
        "        GEngine->AddOnScreenDebugMessage(-1, 5.f, FColor::Red, TEXT(\"Game Over\"));",
        "    }",
        "}",
      ],
      why: [
        "This mirrors the Blueprint path's structure exactly: a second delegate, broadcast from the same escape branch, handled by a second method on the same GameManager class that already owns EnemiesRemaining - same principle as BP_GameManager, expressed in C++.",
      ],
    },
  },
];


// ---- Rule 14: Teacher's Keys carry complete code for all three levels ----
const keyLevel2Default = (c) => [
  c + " Level 2 is verification only - the Level 1 code is already correct, so no lines change.",
  c + " Checks: shooting an enemy never changes lives (only the kill path changes enemies_remaining);",
  c + " an escape never changes enemies_remaining (only the escape path changes lives);",
  c + " \"Game Over\" appears once, when lives first reaches 0.",
];
const keyComment = { "Scratch": "//", "Godot": "#", "TIC-80": "--", "GDevelop": "//", "001GameCreator": "//", "Unreal": "//", "Unity": "//", "GameMaker": "//" };
const keyLevel3 = {
  "Scratch": [
    "// Player sprite - Level 3 sound:",
    "when I receive [game over]",
    "@@ start sound [sfx_gameover]",
    "say [Game Over] for 3 seconds",
    "",
    "// GameManager - Level 3 sound:",
    "when I receive [enemy escaped]",
    "change lives by -1",
    "@@ start sound [sfx_lifelost]",
    "if lives = 0 then",
    "    broadcast [game over]",
    "",
    "// Level 3 visual: tick the checkbox beside the lives variable in the Variables palette.",
    "// On the last life both sounds overlap - acceptable, or wrap start sound [sfx_lifelost] in: if lives > 0.",
  ],
  "Godot": [
    "# GameManager's script - Level 3 (children: LifeLostSound and GameOverSound, both plain AudioStreamPlayer;",
    "# a Label named LivesText next to LoseText in the main scene):",
    "var lives := 3",
    "",
    "func on_enemy_escaped():",
    "    lives -= 1",
    "@@     get_node(\"../LivesText\").text = \"Lives: \" + str(lives)",
    "    if lives <= 0:",
    "        get_node(\"../LoseText\").show()",
    "@@         $GameOverSound.play()",
    "@@     else:",
    "@@         $LifeLostSound.play()",
  ],
  "Unity": [
    "// GameManager.cs - Level 3 (assign lifeLostSound and gameOverSound in the Inspector):",
    "public static int lives = 3;",
    "public static bool hasLost = false;",
    "@@ public static GameManager instance;",
    "@@ public AudioClip lifeLostSound;",
    "@@ public AudioClip gameOverSound;",
    "",
    "@@ void Awake() {",
    "@@     instance = this;",
    "@@ }",
    "",
    "public static void RegisterEscape() {",
    "    lives--;",
    "    if (lives <= 0 && !hasLost) {",
    "        hasLost = true;",
    "@@         AudioSource.PlayClipAtPoint(instance.gameOverSound, Camera.main.transform.position);",
    "@@     } else if (lives > 0) {",
    "@@         AudioSource.PlayClipAtPoint(instance.lifeLostSound, Camera.main.transform.position);",
    "    }",
    "}",
    "",
    "// GameManager.cs's OnGUI():",
    "@@ GUI.Label(new Rect(20, 20, 200, 30), \"Lives: \" + lives);",
    "if (hasLost) {",
    "    GUI.Label(new Rect(Screen.width/2 - 50, 90, 200, 30), \"Game Over\");",
    "}",
  ],
  "GDevelop": [
    "(existing escape event, Level 3)",
    "Condition: Enemy.Y() > 750",
    "Action: Lives = Lives - 1",
    "@@ Action: Play a sound (sfx_lifelost.wav) - only if Lives > 0, add Condition: Lives > 0 on a separate event if the overlap bothers students",
    "Action: Delete Enemy",
    "",
    "(game over event, Level 3)",
    "Condition: Lives <= 0",
    "Action: Show LoseText",
    "@@ Action: Play a sound (sfx_gameover.wav)",
    "",
    "(Level 3 visual)",
    "@@ Action on the escape event: set the content of LivesText (a second Text object) to \"Lives: \" + ToString(Lives)",
  ],
  "TIC-80": [
    "lives=3",
    "@@ has_lost=false",
    "",
    "-- enemy-drawing loop (Level 3):",
    "for enemy_index=1,#enemies do",
    " local enemy=enemies[enemy_index]",
    " if enemy.alive then",
    "  enemy.ey=enemy.ey+0.3",
    "  if enemy.ey>136 then",
    "   enemy.alive=false",
    "   lives=lives-1",
    "@@    if lives<=0 then",
    "@@     if not has_lost then",
    "@@      has_lost=true",
    "@@      sfx(2)",
    "@@     end",
    "@@    else",
    "@@     sfx(1)",
    "@@    end",
    "  else",
    "   spr(3,enemy.ex,enemy.ey,0)",
    "  end",
    " end",
    "end",
    "",
    "-- bottom of TIC() (Level 3 visual readout):",
    "@@ print(\"Lives: \"..lives,4,4,12)",
    "if lives<=0 then",
    " print(\"Game Over\",40,60,8)",
    "end",
    "-- sfx slots: 1 = sfx_lifelost.wav, 2 = sfx_gameover.wav",
  ],
  "001GameCreator": [
    "enemy Actor Trigger: Left/Entered Camera (Left case) -> (Level 3)",
    "  Variable – Modify Value(lives, Subtract, 1)",
    "@@   Play Sound(sfx_lifelost.wav)",
    "  If(lives <= 0)",
    "@@     Play Sound(sfx_gameover.wav)",
    "    Message Box(\"Game Over\")",
    "  Actor – Destroy/Revive Actor(this, Destroy)",
    "",
    "@@ Level 3 visual: add a Text/HUD element bound to the lives variable and refresh it in this same trigger.",
  ],
  "Unreal": [
    "BP_GameManager - Level 3:",
    "Custom Event HandleEnemyEscaped ->",
    "  Lives = Lives - 1",
    "@@   Play Sound 2D(sfx_lifelost.wav)",
    "@@   Print String(\"Lives: \" + Lives, Print to Screen = true)",
    "  Branch(Lives <= 0) -> True:",
    "@@     Play Sound 2D(sfx_gameover.wav)",
    "    Print String(\"Game Over\", Print to Screen = true)",
  ],
};
keyData.forEach(k => {
  const h = handoutData.find(x => x.file === k.file);
  if (!k.level1Code && h && h.level1Code) k.level1Code = h.level1Code;
  if (!k.level2Code) k.level2Code = keyLevel2Default(keyComment[k.file] || "//");
  if (!k.level3Code && keyLevel3[k.file]) k.level3Code = keyLevel3[k.file];
});

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_08_Key_${d.file}.docx`]));

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
