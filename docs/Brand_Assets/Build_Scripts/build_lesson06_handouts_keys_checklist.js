const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign
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
function code(lines) {
  const children = [];
  lines.forEach((line, i) => {
    children.push(new TextRun({ text: line, font: "Consolas", size: 20 }));
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
      headers: { default: brand.pageHeader("Lesson 6 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 6 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("This Is a Milestone Lesson"),
        p("Today's exercise completes the Level 1 MVP milestone from the Semester 1 Capstone GDD - by the end of class, every student's game will have a complete, playable core loop: move, shoot, face multiple enemies, and win. Worth calling out explicitly in class; it's a real accomplishment six weeks in, not just another incremental step."),

        h2("Self-Contained Lesson Assets"),
        p("Everything needed for the core exercise is already in students' projects from Lessons 2-5, and the \"win text\" is drawn/displayed directly by each engine's own basic text tools, not an image asset. Two new placeholder sound effects were added to this lesson's Assets/ folder for the Level 3 sound idea: sfx_win.wav (victory stinger, play it the moment the win check passes) and sfx_start.wav (a game-start/announcement stinger, provided for general use)."),

        h2("Continuity From Lesson 5"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 5, not create a new one - today adds a win check on top of the enemy row they already built."),
        checkbox("If a student's Lesson 5 project is lost or broken, have a pre-built \"Lesson 5 complete\" starter available per engine (from the Lesson 5 Teacher's Keys) so they aren't blocked."),

        h2("Engine Verification"),
        p("This week introduces each engine's simplest built-in way to show text on screen, deliberately avoiding a full UI/Canvas system - that's coming properly in Lesson 9."),
        checkbox("Confirm the demo laptop's Lesson 4 and 5 exercises still work for all 8 engines."),
        checkbox("Godot: confirm a simple Label node, hidden by default and shown via script, displays correctly."),
        checkbox("Unity: confirm OnGUI() with GUI.Label() actually renders in your installed Editor version (it's legacy IMGUI, still fully supported, but worth a quick check)."),
        checkbox("GameMaker: confirm obj_enemy's Destroy event fires reliably and can reference a variable on another instance (obj_spawner.enemies_remaining) without error."),
        checkbox("GDevelop: confirm a Text object can be placed hidden and shown via an action in your installed version."),
        checkbox("TIC-80: confirm print() draws text at the coordinates you expect in your installed version - low risk, but worth a 30-second check."),
        checkbox("001 Game Creator (Windows demo laptop): confirm the \"Defeated / Destroyed\" Actor Trigger and the \"Message Box\" event both work as expected - this is a different trigger than \"Collided with actor\" from Lesson 4, so re-verify."),
        checkbox("If any advanced-track students are on Unreal this week, confirm Event Dispatchers (Add Event Dispatcher, Bind Event to) work as expected in your installed version, and that Print String's \"Print to Screen\" option is enabled by default."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies."),

        h2("Known Gotchas This Week"),
        bullet("Every engine: the most common bug is an off-by-one error in the enemy count - double check the starting count actually matches the number of enemies the Lesson 5 spawn loop creates (5, unless a student changed it)."),
        bullet("Godot / Unity / GameMaker / GDevelop / Unreal: the win check needs to run from the code that actually destroys an enemy, not from a separate \"check every frame\" loop - a per-frame check is more error-prone and unnecessary here."),
        bullet("Unity: OnGUI() only draws while the GameObject it's attached to is active and Play Mode is running - if nothing appears, confirm the GameManager GameObject itself wasn't accidentally deleted or disabled."),
        bullet("001 Game Creator: \"Defeated / Destroyed\" is a different trigger from \"Collided with actor\" - both belong on the enemy Actor Template, doing two different jobs (one for the entire fight, one specifically for counting down)."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_06_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 6: Victory Condition", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: a real win condition. Your job is to track how many enemies are left and show \"You Win!\" the moment the last one is destroyed - completing your game's Level 1 MVP milestone.", { bold: true }),
    p("Open your own [YourName]Game project from Lesson 5 - don't start a new one. Everything this week builds on the row of enemies you already have.", { bold: true }),
    p("Tip: no new assets this week - \"You Win!\" is shown with each engine's own basic text tools, not an image.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("The moment the last enemy is destroyed, \"You Win!\" appears on screen. This completes your game's Level 1 MVP milestone - a complete, playable core loop.", { italics: true }),
    ...data.level1.map(s => numbered(s)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  if (data.altPath) {
    children.push(h2("Alternate Path: " + data.altPath.title));
    children.push(p(data.altPath.intro, { italics: true, color: "555555" }));
    if (data.altPath.warning) children.push(p(data.altPath.warning, { bold: true, color: "B00020" }));
    children.push(...data.altPath.steps.map(s => numbered(s)));
    if (data.altPath.code) children.push(code(data.altPath.code));
    if (data.altPath.note) children.push(p(data.altPath.note, { italics: true, color: "555555" }));
  }
  children.push(
    h2("Level 2 - Complete (Alpha/Beta)"),
    p("The win text appears exactly once, exactly when the last enemy is destroyed - not early, not late, and not repeatedly.", { italics: true }),
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
      headers: { default: brand.pageHeader("Lesson 6 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 5.",
      "Create a new sprite (any simple placeholder costume), name it GameManager, and hide it at the green flag - it never needs to be seen, it just holds win-tracking logic.",
      "Create a new variable enemies_remaining (for all sprites, not \"for this sprite only\").",
    ],
    level1: [
      "On GameManager, add \"when green flag clicked\" -> \"hide\" -> \"set enemies_remaining to 5\".",
      "In the Enemy's \"when I start as a clone\" script, inside the existing \"if touching Bullet? then\" branch, replace the old decrement with \"broadcast [enemy hit]\" right before \"delete this clone\" - the Enemy now only reports that it was hit, it no longer counts anything itself.",
      "On GameManager, add a new script: \"when I receive [enemy hit]\" -> \"change enemies_remaining by -1\" -> \"if enemies_remaining = 0 then broadcast [you win]\".",
      "On the Player sprite, add a new script: \"when I receive [you win]\" -> \"say [You Win!] for 3 seconds\".",
      "Click the stage, then shoot every enemy to test - the message should appear the instant the last one is destroyed.",
    ],
    level2: [
      "Destroy all 5 enemies and confirm \"You Win!\" appears exactly once, exactly when the last one is destroyed - not one enemy early, and not showing up multiple times.",
      "This already works correctly because all the counting logic lives in exactly one place - GameManager's \"when I receive [enemy hit]\" script - which runs once per broadcast, and each Enemy clone only ever broadcasts once, right when it's actually hit.",
    ],
    level3: [
      "Sound: add \"start sound\" using the provided sfx_win.wav right before the \"say\" block on the Player's \"when I receive [you win]\" script.",
      "Visual: make the win message more of a celebration - try a few \"change size by\" pulses in a row, or a costume/color-effect change, instead of a plain \"say\" block.",
      "Structure: keep all win-tracking (enemies_remaining, the zero-check, the win broadcast) inside GameManager's own script - the Enemy sprite only ever reports that it was hit, it never counts or decides anything itself.",
    ],
    trouble: [
      "\"You Win!\" never appears: confirm GameManager's \"when green flag clicked\" script actually runs and sets enemies_remaining to 5 - an easy thing to miss since GameManager is hidden and easy to forget about.",
      "\"You Win!\" appears too early or too late: confirm enemies_remaining is being set to the correct starting count (matching your actual spawn loop's repeat count) on GameManager's green-flag script.",
      "Message appears multiple times: confirm GameManager has only ONE \"when I receive [enemy hit]\" script - a second copy would double-count every hit.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 5.",
      "Add a Label node to your main scene, rename it WinText, set its Text to \"You Win!\", and hide it (uncheck Visible in the Inspector, or call hide() in code).",
      "Add a new Node as a child of your main scene, rename it GameManager, and attach a new script to it - this is the one place win-tracking will live.",
      "In enemy.tscn's own script, add a new signal at the top: signal destroyed.",
    ],
    level1: [
      "In enemy.tscn's _on_hitbox_area_entered(), call destroyed.emit() right before queue_free() (code below).",
      "In GameManager's script, add a var enemies_remaining and a function on_enemy_destroyed() that decrements it and shows WinText at zero.",
      "In the main scene's script (which still does the spawning, same as Lesson 5), connect each new enemy's destroyed signal to $GameManager.on_enemy_destroyed instead of a method on the main scene itself.",
      "Run the scene and destroy every enemy to test - \"You Win!\" should appear the instant the last one is destroyed.",
    ],
    level1Code: [
      "# enemy.tscn's script - add at the top and inside the hit handler:",
      "signal destroyed",
      "",
      "func _on_hitbox_area_entered(area):",
      "    area.get_parent().queue_free()",
      "    destroyed.emit()",
      "    queue_free()",
      "",
      "# GameManager's script:",
      "var enemies_remaining := 5",
      "",
      "func on_enemy_destroyed():",
      "    enemies_remaining -= 1",
      "    if enemies_remaining <= 0:",
      "        get_node(\"../WinText\").show()",
      "",
      "# Main scene's script:",
      "func _ready():",
      "    for enemy_count in range(5):",
      "        var new_enemy := enemy_scene.instantiate()",
      "        new_enemy.position = Vector2(40 + enemy_count * 60, 40)",
      "        new_enemy.destroyed.connect($GameManager.on_enemy_destroyed)",
      "        add_child(new_enemy)",
    ],
    level2: [
      "Destroy all 5 enemies and confirm \"You Win!\" appears exactly once, exactly when the last one is destroyed - not one enemy early, and not repeatedly.",
      "This already works correctly because destroyed.emit() fires exactly once per enemy (right where it's actually destroyed), and each enemy's signal is connected exactly once, always to GameManager's on_enemy_destroyed() - there's no way for the count to drift.",
    ],
    level3: [
      "Sound: add an AudioStreamPlayer node (not 2D, since it's UI, not positional) as a child of GameManager, assign the provided sfx_win.wav as its stream, and call it from GameManager's own script right where WinText is shown.",
      "Visual: instead of just show()-ing WinText, animate it in - a quick scale tween using create_tween() reads more like a celebration than a plain appearance.",
      "Structure: keep the win-tracking (enemies_remaining, the check for zero) entirely inside GameManager's own script - the main scene only ever spawns enemies and wires up their signals, and enemy.tscn only ever announces that it was destroyed. Neither one tracks or decides anything itself.",
    ],
    trouble: [
      "\"You Win!\" never appears: confirm destroyed.emit() is actually being called, and that new_enemy.destroyed.connect($GameManager.on_enemy_destroyed) runs for every spawned enemy, not just some of them.",
      "\"You Win!\" appears too early: confirm enemies_remaining's starting value on GameManager actually matches the number of enemies your loop spawns.",
      "Error about a missing signal: confirm signal destroyed is declared at the very top of enemy.tscn's script, outside any function.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Open your [YourName]Game project from Lesson 5.",
      "Create a new empty GameObject named GameManager. Create a new C# script named GameManager, attach it to this GameObject.",
    ],
    level1: [
      "In GameManager.cs, add a public static int enemiesRemaining, a public static bool hasWon, and a static method RegisterKill() that decrements the count and sets hasWon at zero. Add an OnGUI() that draws \"You Win!\" when hasWon is true (code below).",
      "In Bullet.cs's OnTriggerEnter2D, call GameManager.RegisterKill() right where you already destroy the enemy.",
      "In GameManager.cs's Start() (or a field default), set enemiesRemaining to match your EnemySpawner's enemyCount.",
      "Press Play and destroy every enemy to test - \"You Win!\" should appear the instant the last one is destroyed.",
    ],
    level1Code: [
      "// GameManager.cs:",
      "public static int enemiesRemaining = 5;",
      "public static bool hasWon = false;",
      "",
      "public static void RegisterKill() {",
      "    enemiesRemaining--;",
      "    if (enemiesRemaining <= 0) hasWon = true;",
      "}",
      "",
      "void OnGUI() {",
      "    if (hasWon) GUI.Label(new Rect(10, 10, 200, 30), \"You Win!\");",
      "}",
      "",
      "// Bullet.cs - add inside OnTriggerEnter2D, next to Destroy(other.gameObject):",
      "GameManager.RegisterKill();",
    ],
    level2: [
      "Destroy all 5 enemies and confirm \"You Win!\" appears exactly once, exactly when the last one is destroyed - not one enemy early, and not flickering or repeating.",
      "This already works correctly because RegisterKill() runs exactly once per actual hit (right next to the existing Destroy(other.gameObject) call), and hasWon only ever flips from false to true, never back.",
    ],
    level3: [
      "Sound: add an AudioClip field to GameManager, assign the provided sfx_win.wav to it in the Inspector, and call AudioSource.PlayClipAtPoint() the moment hasWon first becomes true inside RegisterKill().",
      "Visual: style OnGUI()'s label with a custom GUIStyle (bigger font size, centered) instead of the plain default look.",
      "Structure: keep all win-tracking (enemiesRemaining, hasWon, RegisterKill()) inside GameManager.cs - Bullet.cs only ever calls RegisterKill(), it never touches the counting logic directly.",
    ],
    trouble: [
      "\"You Win!\" never appears: confirm RegisterKill() is actually being called from Bullet.cs, and that the GameManager GameObject is present and active in the scene.",
      "\"You Win!\" appears too early: confirm enemiesRemaining's starting value actually matches your EnemySpawner's enemyCount.",
      "Nothing draws at all: OnGUI() only runs on a MonoBehaviour that's active in the scene - confirm GameManager.cs is attached to a GameObject that hasn't been disabled or deleted.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same kill-tracking pattern works in a Script Graph instead of C#.",
      steps: [
        "On the GameManager GameObject, add a Script Machine with a Scene Variable enemiesRemaining set to 5.",
        "On the Bullet prefab's graph, right after the existing Destroy nodes in its trigger event, add a \"Decrement\" node on the scene's enemiesRemaining variable, followed by a Branch checking if it's <= 0.",
        "On True, use an \"On GUI\" event node on GameManager's graph with a GUI Label node showing \"You Win!\", gated by a Scene Variable hasWon that the Branch sets to true.",
      ],
      note: "Level 2/3 already work the same way as the C# path, since the counting logic lives in one place (the scene variable) either way.",
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: [
      "Open your [YourName]Game project from Lesson 5.",
      "Create a new object obj_gamemanager (no sprite needed) and drag one instance of it into your room, anywhere off to the side - it just needs to exist once, and it's the one place win-tracking will live.",
    ],
    level1: [
      "On obj_gamemanager's Create event, add enemies_remaining = 5; and has_won = false;.",
      "Add a Destroy event to obj_enemy - GameMaker automatically runs this once whenever an instance of obj_enemy is destroyed, for any reason. Inside it, decrement obj_gamemanager's enemies_remaining and check for zero (code below).",
      "Add a Draw GUI event to obj_gamemanager that draws \"You Win!\" text when has_won is true.",
      "Press F5 (Run) and destroy every enemy to test - \"You Win!\" should appear the instant the last one is destroyed.",
    ],
    level1Code: [
      "// obj_gamemanager's Create event:",
      "enemies_remaining = 5;",
      "has_won = false;",
      "",
      "// obj_enemy's Destroy event:",
      "obj_gamemanager.enemies_remaining -= 1;",
      "if (obj_gamemanager.enemies_remaining <= 0) {",
      "    obj_gamemanager.has_won = true;",
      "}",
      "",
      "// obj_gamemanager's Draw GUI event:",
      "if (has_won) {",
      "    draw_text(400, 50, \"You Win!\");",
      "}",
    ],
    level2: [
      "Destroy all 5 enemies and confirm \"You Win!\" appears exactly once, exactly when the last one is destroyed - not one enemy early, and not repeatedly.",
      "This already works correctly because GameMaker's Destroy event fires exactly once per instance, no matter what caused the destruction, and has_won only ever flips from false to true, never back.",
    ],
    level3: [
      "Sound: import the provided sfx_win.wav as a new sound resource (keep its default name, sfx_win), then call audio_play_sound() inside obj_enemy's Destroy event, right where enemies_remaining reaches zero.",
      "Visual: use draw_set_font() and draw_set_halign() before draw_text() in obj_gamemanager's Draw GUI event to make the win text bigger and centered instead of the small default look.",
      "Structure: keep all win-tracking (enemies_remaining, has_won) as variables on obj_gamemanager - obj_spawner only ever spawns enemies, and obj_enemy's Destroy event only ever reaches out to modify obj_gamemanager's variables. Neither one draws anything itself.",
    ],
    trouble: [
      "\"You Win!\" never appears: confirm obj_gamemanager actually has one instance placed in the room, that obj_enemy's Destroy event references it correctly, and that it has its own Draw GUI event - a missing Draw GUI event means has_won can be true with nothing ever drawing it.",
      "\"You Win!\" appears too early: confirm enemies_remaining's starting value in obj_gamemanager's Create event actually matches your spawn loop's count.",
      "Error referencing obj_gamemanager: confirm there's actually exactly one obj_gamemanager instance in the room - referencing an object with zero instances by name causes an error.",
    ],
    resource: "Student Resource List > GameMaker section, or gamemaker.io/en/tutorials",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same kill-tracking pattern works with drag-and-drop blocks instead of typed code.",
      steps: [
        "On obj_enemy's Destroy event, switch to GML Visual. Drag a \"Set Variable (Relative)\" block targeting obj_gamemanager's enemies_remaining, then an \"If Variable\" block checking it against 0, with a \"Set Variable\" block inside setting has_won to true.",
        "On obj_gamemanager's Draw GUI event, drag an \"If Variable\" block checking has_won, with a \"Draw Text\" block inside it.",
        "Press F5 (Run) and test the same way as the code path.",
      ],
      note: "Level 2/3's per-instance logic already lives the same way in both paths, since GameMaker's Destroy event behaves identically whether it contains blocks or code.",
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 5.",
      "Add a new Text object, name it WinText, set its content to \"You Win!\", place it on screen, and hide it (uncheck its visibility, or add a \"Hide\" action at the beginning of the scene).",
    ],
    level1: [
      "In your Combat event group's collision event (Bullet is in collision with Enemy), add an action to decrease EnemiesRemaining by 1, right alongside the existing Delete actions.",
      "Add a new event: Condition \"EnemiesRemaining <= 0\" -> Action \"Show\" WinText.",
      "Click the green Preview button and destroy every enemy to test - \"You Win!\" should appear the instant the last one is destroyed.",
    ],
    level1Code: [
      "Combat event group:",
      "Condition: Bullet is in collision with Enemy",
      "Action: Play a sound (sfx_hit.wav)",
      "Action: EnemiesRemaining = EnemiesRemaining - 1",
      "Action: Delete Bullet",
      "Action: Delete Enemy",
      "",
      "(new event, outside the group)",
      "Condition: EnemiesRemaining <= 0",
      "Action: Show WinText",
    ],
    level2: [
      "Destroy all 5 enemies and confirm \"You Win!\" appears once the count reaches zero and stays visible - not flickering on and off.",
      "This already works correctly because the win-check event's condition (EnemiesRemaining <= 0) stays true forever once it's reached, and \"Show\" on an already-visible object simply does nothing extra.",
    ],
    level3: [
      "Sound: add a \"Play a sound\" action to the win-check event, using the provided sfx_win.wav.",
      "Visual: add a \"Change the scale of an object\" or \"Tween\" action to the win-check event so WinText grows in instead of just appearing.",
      "Structure: keep the win-check as its own separate event, clearly outside the \"Combat\" event group - it reads EnemiesRemaining but isn't really part of the hit-detection logic itself. (GDevelop's scene variables are already global rather than tied to any one object, so there's no separate \"GameManager\" object to introduce the way Scratch/Godot/GameMaker/Unreal now use - the scene variable already plays that role natively.)",
    ],
    trouble: [
      "\"You Win!\" never appears: confirm EnemiesRemaining is actually being decreased inside the collision event, and that its starting value (set when spawning) matches the number of enemies actually created.",
      "\"You Win!\" appears immediately at the start: confirm EnemiesRemaining's starting value is being set correctly (5, not 0) before any enemies are destroyed.",
      "WinText never shows even though the condition seems true: confirm WinText was actually given a \"Hide\" action at the start rather than being deleted or never placed on the scene.",
    ],
    resource: "Student Resource List > GDevelop section, or gdevelop.io/tutorials",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: [
      "Launch TIC-80 and load your cart from Lesson 5.",
      "Press Esc to reach the code editor.",
    ],
    level1: [
      "Add enemies_remaining=5 near the top of your code, alongside your enemies table.",
      "Inside the bullets loop's hit branch (where an enemy's alive is set to false), decrement enemies_remaining.",
      "At the bottom of TIC(), add an if check that calls print(\"You Win!\", ...) when enemies_remaining reaches 0 (code below).",
      "Press Esc, type run, press Enter, then shoot every enemy to test - \"You Win!\" should appear the instant the last one is destroyed.",
    ],
    level1Code: [
      "enemies={}",
      "for enemy_count=0,4 do",
      " table.insert(enemies,{ex=20+enemy_count*20,ey=10,alive=true})",
      "end",
      "enemies_remaining=5",
      "bullets={}",
      "",
      "function TIC()",
      " cls(0)",
      " if btn(3) then x=x+1 end",
      " if btnp(4) then",
      "  table.insert(bullets,{bx=x,by=y})",
      " end",
      " for i=#bullets,1,-1 do",
      "  local b=bullets[i]",
      "  b.by=b.by-2",
      "  local hit=false",
      "  for j=1,#enemies do",
      "   local e=enemies[j]",
      "   if e.alive and b.bx<e.ex+8 and b.bx+8>e.ex and b.by<e.ey+8 and b.by+8>e.ey then",
      "    e.alive=false",
      "    enemies_remaining=enemies_remaining-1",
      "    hit=true",
      "   end",
      "  end",
      "  if hit then",
      "   table.remove(bullets,i)",
      "  elseif b.by<0 then",
      "   table.remove(bullets,i)",
      "  else",
      "   spr(2,b.bx,b.by,0)",
      "  end",
      " end",
      " for j=1,#enemies do",
      "  local e=enemies[j]",
      "  if e.alive then spr(3,e.ex,e.ey,0) end",
      " end",
      " spr(1,x,y,0)",
      " if enemies_remaining<=0 then",
      "  print(\"You Win!\",84,60,12)",
      " end",
      "end",
    ],
    level2: [
      "Destroy all 5 enemies and confirm \"You Win!\" appears exactly when enemies_remaining reaches 0, staying on screen from then on.",
      "This already works correctly because enemies_remaining is only ever decremented inside the same guarded branch (e.alive and ...) that already prevents a single enemy from being counted as destroyed more than once.",
    ],
    level3: [
      "Sound: import the provided sfx_win.wav into a new SFX slot (TIC-80's Surf/import tools accept an external .wav) rather than designing a new effect from scratch, then call sfx(slot_number) the moment enemies_remaining first reaches 0 - guard it with a has_won flag so it doesn't replay every single frame while the win text is showing.",
      "Visual: use print()'s extra parameters (scale, fixed width) or draw a rectangle behind the text with rect() first, so the message stands out more.",
      "Structure: keep enemies_remaining declared once at the top with your other game state, and only ever change it from inside the collision check - the same principle you've already been applying to enemies_alive and bullets. (TIC-80 has no object system at all, so there's no separate \"GameManager\" object to introduce here - a top-level global variable already is this engine's native equivalent.)",
    ],
    trouble: [
      "\"You Win!\" never appears: confirm enemies_remaining is actually being decremented inside the hit branch, not just enemies[j].alive.",
      "\"You Win!\" appears immediately: confirm enemies_remaining starts at 5 (or your actual enemy count), not 0.",
      "Text flashes then disappears: confirm the print() call is guarded by enemies_remaining<=0, which stays true forever once reached - if it disappears, something else may be resetting the variable.",
    ],
    resource: "Student Resource List > TIC-80 section, or tic80.com",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 5.",
      "Create a new Variable (Resources > Variables/Switches/Collections/Tables) named enemies_remaining, default value 5.",
    ],
    level1: [
      "On the enemy Actor Template, add a second Actor Trigger: \"Defeated / Destroyed\" - this fires automatically whenever a copy of this Actor is destroyed, for any reason, separately from the \"Collided with actor\" trigger you built in Lesson 4.",
      "In that trigger's script, add a Variable event that decreases enemies_remaining by 1.",
      "Add an \"If\" event checking whether enemies_remaining is 0 or less - inside it, add a \"Message Box\" event with the text \"You Win!\".",
      "Save, then Play/Test and destroy every enemy to test - \"You Win!\" should appear the instant the last one is destroyed.",
    ],
    level1Code: [
      "enemy Actor Trigger: Defeated / Destroyed ->",
      "  Variable – Modify Value(enemies_remaining, Subtract, 1)",
      "  If(enemies_remaining <= 0)",
      "    Message Box(\"You Win!\")",
    ],
    level2: [
      "Destroy all 5 enemies and confirm \"You Win!\" appears exactly once, exactly when the last one is destroyed - not one enemy early, and not repeatedly.",
      "This already works correctly because \"Defeated / Destroyed\" fires exactly once per Actor copy, no matter what destroyed it, and the If check only ever passes once enemies_remaining actually reaches zero.",
    ],
    level3: [
      "Sound: import the provided sfx_win.wav, then add a \"Play Sound\" event to the \"Defeated / Destroyed\" trigger's script, right before the Message Box.",
      "Visual: look at the Message Box's own settings (or Visual Settings for message boxes generally) for a background/border color that stands out as a \"win\" moment rather than a plain dialogue box.",
      "Structure: keep this counting logic entirely on its own \"Defeated / Destroyed\" trigger, separate from \"Collided with actor\" - one trigger handles the fight itself, the other tracks overall progress. (Unlike Scratch/Godot/GameMaker/Unreal, 001 Game Creator's Actor Triggers are bound to the actor they're attached to, so this logic stays on the enemy Template itself rather than moving to a separate \"GameManager\" Actor - see the Book Notes entry on this decision.)",
    ],
    trouble: [
      "\"You Win!\" never appears: confirm the \"Defeated / Destroyed\" trigger is actually attached to the enemy Actor Template (not main or the bullet), and that the Variable event is targeting the correct Variable name.",
      "\"You Win!\" appears too early: confirm enemies_remaining's default value actually matches how many enemy copies you placed in Lesson 5.",
      "Message Box doesn't appear even though the trigger seems to fire: confirm the If event's condition is actually enemies_remaining <= 0 and that the Message Box event is nested inside it, not after it unconditionally.",
    ],
    resource: "001gamecreator.com/wiki, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 5.",
      "Open BP_Enemy and add an Event Dispatcher (Details panel under My Blueprint > Event Dispatchers, click +) named OnEnemyDestroyed.",
      "Create a new Actor Blueprint named BP_GameManager (Blueprint Class > Actor), and drag one instance of it into the level - it just needs to exist once, and it's the one place win-tracking will live.",
    ],
    level1: [
      "In BP_Enemy's Event Graph, inside the successful-Cast branch (right before the two Destroy Actor nodes), add a \"Call OnEnemyDestroyed\" node.",
      "In BP_GameManager, add an Integer variable EnemiesRemaining, default value 5.",
      "In the Level Blueprint, drag your placed BP_GameManager instance from the level (or the World Outliner) into the Event Graph to create a reference to it. In the spawn For Loop, right after each Spawn Actor from Class, add a \"Bind Event to OnEnemyDestroyed\" node targeting that spawned Actor, wired to that BP_GameManager reference's HandleEnemyDestroyed custom event.",
      "In BP_GameManager, build HandleEnemyDestroyed: decrease EnemiesRemaining by 1, Branch on EnemiesRemaining <= 0, and on True, add a Print String node with text \"You Win!\" and \"Print to Screen\" checked.",
      "Click Compile, then Play and destroy every enemy to test - \"You Win!\" should appear on screen the instant the last one is destroyed.",
    ],
    level1Code: [
      "BP_Enemy's Event Graph (inside the successful Cast branch):",
      "Call OnEnemyDestroyed (Broadcast) -> Destroy Actor(Cast result) -> Destroy Actor(Self)",
      "",
      "Level Blueprint:",
      "Event BeginPlay -> (BP_GameManager reference, dragged in from the level) ->",
      "  For Loop -> Spawn Actor from Class(BP_Enemy, ...) ->",
      "    Bind Event to OnEnemyDestroyed(Target = spawned Actor, Event = BP_GameManager ref's HandleEnemyDestroyed)",
      "",
      "BP_GameManager:",
      "Custom Event HandleEnemyDestroyed ->",
      "  EnemiesRemaining = EnemiesRemaining - 1",
      "  Branch(EnemiesRemaining <= 0) -> True: Print String(\"You Win!\", Print to Screen = true)",
    ],
    level2: [
      "Destroy all 5 enemies and confirm \"You Win!\" appears exactly once, exactly when the last one is destroyed - not one enemy early, and not repeatedly.",
      "This already works correctly because each spawned enemy's dispatcher is bound exactly once, at spawn time, always to the same BP_GameManager instance, and Call OnEnemyDestroyed only ever broadcasts once per enemy, right at the moment it's actually destroyed.",
    ],
    level3: [
      "Sound: import the provided sfx_win.wav as a Sound Wave asset, then add a Play Sound 2D node right before the Print String node in BP_GameManager's HandleEnemyDestroyed.",
      "Visual: use Print String's Text Color parameter to make the win message gold or bright green instead of the default white, and increase Duration so it stays on screen longer.",
      "Structure: keep all counting logic (EnemiesRemaining, HandleEnemyDestroyed) inside BP_GameManager - the Level Blueprint only ever spawns enemies and binds their dispatchers, and BP_Enemy only ever broadcasts that it was destroyed. Neither one knows or cares how many enemies remain overall.",
    ],
    trouble: [
      "\"You Win!\" never appears: confirm Bind Event to OnEnemyDestroyed is actually wired for every spawned enemy inside the For Loop, and that it's targeting BP_GameManager's HandleEnemyDestroyed, not a Custom Event on the Level Blueprint itself.",
      "\"You Win!\" appears too early: confirm EnemiesRemaining's default value on BP_GameManager actually matches your For Loop's enemy count.",
      "Compile error about the event signature: confirm HandleEnemyDestroyed's parameters (or lack of them) match OnEnemyDestroyed's Event Dispatcher signature exactly.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same kill-tracking pattern works in C++ instead of Blueprint nodes.",
      steps: [
        "Add a DECLARE_DYNAMIC_MULTICAST_DELEGATE to ACppEnemy.h and a UPROPERTY(BlueprintAssignable) instance of it, then Broadcast() it from OnOverlap() right before calling Destroy().",
        "In a GameManager actor class's BeginPlay(), bind a UFUNCTION to each spawned enemy's delegate using AddDynamic(), and track EnemiesRemaining as a member variable - the C++ equivalent of BP_GameManager.",
        "In the bound handler, decrement EnemiesRemaining and call GEngine->AddOnScreenDebugMessage() to show \"You Win!\" when it reaches zero - the C++ equivalent of Print String's Print to Screen option.",
      ],
      code: [
        "// ACppEnemy.h",
        "DECLARE_DYNAMIC_MULTICAST_DELEGATE(FOnEnemyDestroyed);",
        "UPROPERTY(BlueprintAssignable)",
        "FOnEnemyDestroyed OnEnemyDestroyed;",
        "",
        "// Inside OnOverlap(), before Destroy():",
        "OnEnemyDestroyed.Broadcast();",
        "",
        "// In your GameManager's handler:",
        "EnemiesRemaining--;",
        "if (EnemiesRemaining <= 0) {",
        "    GEngine->AddOnScreenDebugMessage(-1, 5.f, FColor::Yellow, TEXT(\"You Win!\"));",
        "}",
      ],
      note: "This mirrors the Blueprint path's Event Dispatcher and Print String node - same underlying engine calls, just from code. Level 2/3's logic already lives the same way in both paths.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_06_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS (one per engine) - instructor only
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.0, { after: 120 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 6 Teacher's Key", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "INSTRUCTOR ONLY - do not distribute to students", bold: true, size: 18, color: "B00020" })] }),
    h2("Level 1 - Complete Worked Solution"),
    ...data.l1Notes.map(s => p(s)),
    code(data.l1Code),
    h2("Level 2 - Complete Worked Solution"),
    ...data.l2Notes.map(s => p(s)),
  ];
  if (data.l2Code) children.push(code(data.l2Code));
  if (data.l2Extra) children.push(...data.l2Extra.map(s => p(s)));
  children.push(
    h2("Level 3 - All Worked Solutions"),
    p("Every stretch-goal idea from the handout, fully worked - not just one. Useful for validating each idea actually works before class, and as a ready reference no matter which one a student ends up picking."),
  );
  data.l3Ideas.forEach(idea => {
    children.push(h3(idea.title));
    children.push(...idea.notes.map(s => p(s)));
    if (idea.code) children.push(code(idea.code));
  });
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
      headers: { default: brand.pageHeader("Lesson 6 Teacher's Key - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    l1Notes: ["A dedicated GameManager sprite owns all win-tracking, listening for a broadcast from the Enemy's existing clone script:"],
    l1Code: [
      "-- Enemy sprite (spawn script, unchanged from Lesson 5):",
      "when green flag clicked",
      "  set [enemy_number] to 0",
      "  repeat 5",
      "    change [enemy_number] by 1",
      "    set [enemy_x] to ((enemy_number * 60) - 150)",
      "    create clone of [myself]",
      "",
      "-- Enemy sprite (clone script):",
      "when I start as a clone",
      "  show",
      "  go to x: (enemy_x) y: 150",
      "  forever",
      "    if <touching [Bullet]?> then",
      "      broadcast [enemy hit]",
      "      delete this clone",
      "",
      "-- GameManager sprite (new, hidden):",
      "when green flag clicked",
      "  hide",
      "  set [enemies_remaining] to 5",
      "when I receive [enemy hit]",
      "  change [enemies_remaining] by -1",
      "  if <enemies_remaining = 0> then",
      "    broadcast [you win]",
      "",
      "-- Player sprite (new script):",
      "when I receive [you win]",
      "  say [You Win!] for 3 seconds",
    ],
    l2Notes: ["The decrement and the win check both live inside GameManager's single \"when I receive [enemy hit]\" script, which runs once per broadcast, and each Enemy clone broadcasts exactly once, right on its own hit - there's no separate code path that could double-count or skip an enemy, so \"You Win!\" fires exactly once, exactly on the last hit."],
    l3Ideas: [
      {
        title: "Sound: celebration sound",
        notes: ["Add the provided sfx_win.wav right before the say block on the Player's win script:"],
        code: ["when I receive [you win]", "  start sound [sfx_win]", "  say [You Win!] for 3 seconds"],
      },
      {
        title: "Visual: bigger celebration",
        notes: ["A couple of size pulses read more like a celebration than a plain message:"],
        code: ["when I receive [you win]", "  set size to 120%", "  wait 0.1 seconds", "  set size to 100%", "  say [You Win!] for 3 seconds"],
      },
      {
        title: "Structure: one place tracks the win",
        notes: ["The decrement, the zero-check, and the broadcast all stay inside GameManager's own script - the Enemy sprite only ever broadcasts that it was hit, and no second script anywhere else also reads or writes enemies_remaining."],
      },
      {
        title: "Complete Level 3 solution (all 2 goals combined)",
        notes: ["The win-tracking logic unchanged, plus the Player's enhanced win script:"],
        code: [
          "-- Enemy sprite (clone script, unchanged):",
          "when I start as a clone",
          "  show",
          "  go to x: (enemy_x) y: 150",
          "  forever",
          "    if <touching [Bullet]?> then",
          "      broadcast [enemy hit]",
          "      delete this clone",
          "",
          "-- GameManager sprite (unchanged):",
          "when green flag clicked",
          "  hide",
          "  set [enemies_remaining] to 5",
          "when I receive [enemy hit]",
          "  change [enemies_remaining] by -1",
          "  if <enemies_remaining = 0> then",
          "    broadcast [you win]",
          "",
          "-- Player sprite:",
          "when I receive [you win]",
          "  start sound [sfx_win]",
          "  set size to 120%",
          "  wait 0.1 seconds",
          "  set size to 100%",
          "  say [You Win!] for 3 seconds",
        ],
      },
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    l1Notes: ["A custom signal on enemy.tscn, connected once per instance at spawn time to a dedicated GameManager node's script, which owns a hidden Label shown at zero:"],
    l1Code: [
      "# enemy.tscn's script:",
      "signal destroyed",
      "",
      "func _on_hitbox_area_entered(area):",
      "    area.get_parent().queue_free()",
      "    destroyed.emit()",
      "    queue_free()",
      "",
      "# GameManager's script (GameManager is a child Node of the main scene, sibling of WinText):",
      "var enemies_remaining := 5",
      "",
      "func on_enemy_destroyed():",
      "    enemies_remaining -= 1",
      "    if enemies_remaining <= 0:",
      "        get_node(\"../WinText\").show()",
      "",
      "# Main scene's script:",
      "var enemy_scene := preload(\"res://enemy.tscn\")",
      "",
      "func _ready():",
      "    for enemy_count in range(5):",
      "        var new_enemy := enemy_scene.instantiate()",
      "        new_enemy.position = Vector2(40 + enemy_count * 60, 40)",
      "        new_enemy.destroyed.connect($GameManager.on_enemy_destroyed)",
      "        add_child(new_enemy)",
    ],
    l2Notes: ["destroyed.emit() fires exactly once per enemy, right where it's actually destroyed, and each enemy's signal is connected exactly once at spawn time, always to the same GameManager method - there's no way for enemies_remaining to drift or for WinText to show more than once."],
    l3Ideas: [
      {
        title: "Sound: play on win",
        notes: ["Add a plain AudioStreamPlayer (non-2D, since it's UI) as a child of GameManager, with the provided sfx_win.wav assigned as its stream:"],
        code: ["# GameManager's script:", "func on_enemy_destroyed():", "    enemies_remaining -= 1", "    if enemies_remaining <= 0:", "        get_node(\"../WinText\").show()", "        $WinSound.play()"],
      },
      {
        title: "Visual: animate the win text in",
        notes: ["A quick scale tween reads more like a celebration than a plain show():"],
        code: ["# GameManager's script:", "func on_enemy_destroyed():", "    enemies_remaining -= 1", "    if enemies_remaining <= 0:", "        var win_text := get_node(\"../WinText\")", "        win_text.show()", "        win_text.scale = Vector2(0.5, 0.5)", "        var tween := create_tween()", "        tween.tween_property(win_text, \"scale\", Vector2(1, 1), 0.3)"],
      },
      {
        title: "Structure: GameManager owns the counting",
        notes: ["enemy.tscn only ever announces that it was destroyed (destroyed.emit()), and the main scene only ever spawns enemies and wires up their signals - all counting and the win-state decision live entirely inside GameManager's own script."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, in GameManager's script:"],
        code: [
          "func on_enemy_destroyed():",
          "    enemies_remaining -= 1",
          "    if enemies_remaining <= 0:",
          "        var win_text := get_node(\"../WinText\")",
          "        win_text.show()",
          "        win_text.scale = Vector2(0.5, 0.5)",
          "        var tween := create_tween()",
          "        tween.tween_property(win_text, \"scale\", Vector2(1, 1), 0.3)",
          "        $WinSound.play()",
        ],
      },
    ],
  },
  {
    engine: "Unity", file: "Unity",
    l1Notes: ["A small static GameManager holding win state, called from Bullet.cs's existing trigger handler:"],
    l1Code: [
      "// GameManager.cs:",
      "public static int enemiesRemaining = 5;",
      "public static bool hasWon = false;",
      "",
      "public static void RegisterKill() {",
      "    enemiesRemaining--;",
      "    if (enemiesRemaining <= 0) hasWon = true;",
      "}",
      "",
      "void OnGUI() {",
      "    if (hasWon) GUI.Label(new Rect(10, 10, 200, 30), \"You Win!\");",
      "}",
      "",
      "// Bullet.cs - inside OnTriggerEnter2D:",
      "void OnTriggerEnter2D(Collider2D other) {",
      "    if (other.CompareTag(\"Enemy\")) {",
      "        AudioSource.PlayClipAtPoint(hitSound, transform.position);",
      "        GameManager.RegisterKill();",
      "        Destroy(other.gameObject);",
      "        Destroy(gameObject);",
      "    }",
      "}",
    ],
    l2Notes: ["RegisterKill() runs exactly once per actual hit, right next to the existing Destroy(other.gameObject) call, and hasWon only ever flips from false to true - there's no path where the win text could show early, late, or more than once."],
    l3Ideas: [
      {
        title: "Sound: play on win",
        notes: ["Assign the provided sfx_win.wav to winSound in the Inspector, and play it the moment hasWon first becomes true, guarding against re-triggering every frame:"],
        code: [
          "public static AudioClip winSound;",
          "",
          "public static void RegisterKill() {",
          "    enemiesRemaining--;",
          "    if (enemiesRemaining <= 0 && !hasWon) {",
          "        hasWon = true;",
          "        AudioSource.PlayClipAtPoint(winSound, Camera.main.transform.position);",
          "    }",
          "}",
        ],
      },
      {
        title: "Visual: styled win text",
        notes: ["A custom GUIStyle makes the default OnGUI text look far more intentional:"],
        code: [
          "void OnGUI() {",
          "    if (hasWon) {",
          "        GUIStyle style = new GUIStyle();",
          "        style.fontSize = 32;",
          "        style.normal.textColor = Color.yellow;",
          "        GUI.Label(new Rect(Screen.width/2 - 80, 20, 300, 50), \"You Win!\", style);",
          "    }",
          "}",
        ],
      },
      {
        title: "Structure: GameManager owns the counting",
        notes: ["All win-tracking (enemiesRemaining, hasWon, RegisterKill(), OnGUI()) lives inside GameManager.cs - Bullet.cs only ever calls RegisterKill(), it never touches the counting logic or draws anything itself."],
      },
      {
        title: "Complete Level 3 solution (all 2 goals combined)",
        notes: ["GameManager.cs with sound and styled text combined:"],
        code: [
          "public static int enemiesRemaining = 5;",
          "public static bool hasWon = false;",
          "public static AudioClip winSound;",
          "",
          "public static void RegisterKill() {",
          "    enemiesRemaining--;",
          "    if (enemiesRemaining <= 0 && !hasWon) {",
          "        hasWon = true;",
          "        AudioSource.PlayClipAtPoint(winSound, Camera.main.transform.position);",
          "    }",
          "}",
          "",
          "void OnGUI() {",
          "    if (hasWon) {",
          "        GUIStyle style = new GUIStyle();",
          "        style.fontSize = 32;",
          "        style.normal.textColor = Color.yellow;",
          "        GUI.Label(new Rect(Screen.width/2 - 80, 20, 300, 50), \"You Win!\", style);",
          "    }",
          "}",
        ],
      },
    ],
    altNotes: {
      title: "Visual Scripting",
      notes: [
        "A Scene Variable enemiesRemaining plus a \"Decrement\" node called from the Bullet's trigger graph mirrors RegisterKill().",
        "The Level 3 goals map onto nodes the same way: an Audio node gated by a \"has already played\" check for sound, and an On GUI event node with styled GUI Label nodes for the visual idea.",
      ],
      why: [
        "Keeping the same Structure principle (one place owns win-tracking) whether a student is in C# or graphs reinforces that it's a design idea, not a syntax rule.",
      ],
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    l1Notes: ["obj_enemy's built-in Destroy event does the counting, reaching out to a dedicated obj_gamemanager instance rather than the spawner:"],
    l1Code: [
      "// obj_spawner's Create event (spawning only, unchanged from Lesson 5):",
      "for (var i = 0; i < 5; i++) {",
      "    instance_create_layer(80 + i * 96, 80, \"Instances\", obj_enemy);",
      "}",
      "",
      "// obj_gamemanager's Create event (one instance, placed once in the room):",
      "enemies_remaining = 5;",
      "has_won = false;",
      "",
      "// obj_enemy's Destroy event:",
      "obj_gamemanager.enemies_remaining -= 1;",
      "if (obj_gamemanager.enemies_remaining <= 0) {",
      "    obj_gamemanager.has_won = true;",
      "}",
      "",
      "// obj_gamemanager's Draw GUI event:",
      "if (has_won) {",
      "    draw_text(400, 50, \"You Win!\");",
      "}",
    ],
    l2Notes: ["GameMaker's Destroy event fires exactly once per instance, no matter what caused the destruction, and obj_gamemanager's has_won only ever flips from false to true - there's no path where the win text could show early, late, or more than once."],
    l3Ideas: [
      {
        title: "Sound: play on win",
        notes: ["Import the provided sfx_win.wav as a new sound resource (keep its default name, sfx_win), and play it exactly once, right when has_won first flips true:"],
        code: [
          "// obj_enemy's Destroy event:",
          "obj_gamemanager.enemies_remaining -= 1;",
          "if (obj_gamemanager.enemies_remaining <= 0 && !obj_gamemanager.has_won) {",
          "    obj_gamemanager.has_won = true;",
          "    audio_play_sound(sfx_win, 1, false);",
          "}",
        ],
      },
      {
        title: "Visual: styled win text",
        notes: ["Center and enlarge the text using draw_set_halign() and draw_set_font():"],
        code: [
          "// obj_gamemanager's Draw GUI event:",
          "if (has_won) {",
          "    draw_set_halign(fa_center);",
          "    draw_text_transformed(400, 50, \"You Win!\", 2, 2, 0);",
          "}",
        ],
      },
      {
        title: "Structure: obj_gamemanager owns the counting",
        notes: ["All win-tracking (enemies_remaining, has_won) lives as variables on obj_gamemanager - obj_spawner only ever spawns enemies, and obj_enemy's Destroy event only ever reaches out to modify obj_gamemanager's variables. None of the three draws anything except obj_gamemanager itself."],
      },
      {
        title: "Complete Level 3 solution (all 2 goals combined)",
        notes: ["The Destroy event and Draw GUI event with both ideas added:"],
        code: [
          "// obj_enemy's Destroy event:",
          "obj_gamemanager.enemies_remaining -= 1;",
          "if (obj_gamemanager.enemies_remaining <= 0 && !obj_gamemanager.has_won) {",
          "    obj_gamemanager.has_won = true;",
          "    audio_play_sound(sfx_win, 1, false);",
          "}",
          "",
          "// obj_gamemanager's Draw GUI event:",
          "if (has_won) {",
          "    draw_set_halign(fa_center);",
          "    draw_text_transformed(400, 50, \"You Win!\", 2, 2, 0);",
          "}",
        ],
      },
    ],
    altNotes: {
      title: "GML Visual",
      notes: [
        "obj_enemy's Destroy event works identically whether it holds blocks or code - the event itself already fires exactly once per instance either way, and reaches out to obj_gamemanager the same way in both paths.",
        "The Level 3 goals map onto blocks the same way: an Audio category block gated by an \"If Variable\" checking has_won, and \"Draw Text\" blocks with transform options for the visual idea, both still living on obj_gamemanager.",
      ],
      why: [
        "Keeping the same Structure principle (obj_gamemanager owns counting, obj_spawner only spawns, obj_enemy only reports) whether a student is in Code or Visual reinforces that it's a design idea, not a syntax rule.",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    l1Notes: ["A scene variable updated inside the existing Combat event, plus a new event to show a pre-placed, hidden Text object:"],
    l1Code: [
      "At the beginning of the scene:",
      "Action: EnemiesRemaining = 5",
      "Action: Hide WinText",
      "",
      "Event group: \"Combat\"",
      "Condition: Bullet is in collision with Enemy",
      "Action: Play a sound (sfx_hit.wav)",
      "Action: EnemiesRemaining = EnemiesRemaining - 1",
      "Action: Delete Bullet",
      "Action: Delete Enemy",
      "",
      "(new event, outside the group)",
      "Condition: EnemiesRemaining <= 0",
      "Action: Show WinText",
    ],
    l2Notes: ["EnemiesRemaining is only ever decreased inside the same collision event that already runs once per actual hit, and the win-check condition stays true forever once reached - \"Show\" on an already-visible object does nothing extra, so the message never flickers or double-fires."],
    l3Ideas: [
      {
        title: "Sound: victory sound",
        notes: ["Add a sound action to the win-check event itself, using the provided sfx_win.wav:"],
        code: ["Condition: EnemiesRemaining <= 0", "Action: Show WinText", "Action: Play a sound (sfx_win.wav)"],
      },
      {
        title: "Visual: grow the win text in",
        notes: ["A \"Change the scale of an object\" action gives it a bit of punch:"],
        code: ["Condition: EnemiesRemaining <= 0", "Action: Show WinText", "Action: Change the scale of WinText to 0.2", "Action: Tween the scale of WinText to 1 over 0.4 seconds"],
      },
      {
        title: "Structure: win-check stays separate from Combat",
        notes: ["The win-check event lives on its own, outside the \"Combat\" event group - it reads EnemiesRemaining but isn't part of hit-detection itself, keeping the two concerns easy to find independently."],
      },
      {
        title: "Complete Level 3 solution (all 2 goals combined)",
        notes: ["The win-check event with both ideas added:"],
        code: [
          "Condition: EnemiesRemaining <= 0",
          "Action: Show WinText",
          "Action: Change the scale of WinText to 0.2",
          "Action: Tween the scale of WinText to 1 over 0.4 seconds",
          "Action: Play a sound (sfx_win.wav)",
        ],
      },
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    l1Notes: ["One new counter variable, decremented inside the existing collision check, and a print() call gated on it reaching zero:"],
    l1Code: [
      "enemies={}",
      "for enemy_count=0,4 do",
      " table.insert(enemies,{ex=20+enemy_count*20,ey=10,alive=true})",
      "end",
      "enemies_remaining=5",
      "bullets={}",
      "",
      "function TIC()",
      " cls(0)",
      " if btn(3) then x=x+1 end",
      " if btnp(4) then",
      "  table.insert(bullets,{bx=x,by=y})",
      " end",
      " for i=#bullets,1,-1 do",
      "  local b=bullets[i]",
      "  b.by=b.by-2",
      "  local hit=false",
      "  for j=1,#enemies do",
      "   local e=enemies[j]",
      "   if e.alive and b.bx<e.ex+8 and b.bx+8>e.ex and b.by<e.ey+8 and b.by+8>e.ey then",
      "    e.alive=false",
      "    enemies_remaining=enemies_remaining-1",
      "    hit=true",
      "   end",
      "  end",
      "  if hit then",
      "   table.remove(bullets,i)",
      "  elseif b.by<0 then",
      "   table.remove(bullets,i)",
      "  else",
      "   spr(2,b.bx,b.by,0)",
      "  end",
      " end",
      " for j=1,#enemies do",
      "  local e=enemies[j]",
      "  if e.alive then spr(3,e.ex,e.ey,0) end",
      " end",
      " spr(1,x,y,0)",
      " if enemies_remaining<=0 then",
      "  print(\"You Win!\",84,60,12)",
      " end",
      "end",
    ],
    l2Notes: ["enemies_remaining is only ever decremented inside the same e.alive and ... guarded branch that already prevents a single enemy from being counted twice - so the win text appears exactly once it's actually earned, and stays showing from then on since the check is <= 0, not == 0."],
    l3Ideas: [
      {
        title: "Sound: victory sound, once",
        notes: ["Import the provided sfx_win.wav into a new SFX slot rather than designing a new effect, and add a has_won flag so it plays exactly once, not every frame the text is on screen:"],
        code: [
          "has_won=false",
          "...",
          " if enemies_remaining<=0 then",
          "  if not has_won then",
          "   has_won=true",
          "   sfx(3)",
          "  end",
          "  print(\"You Win!\",84,60,12)",
          " end",
        ],
      },
      {
        title: "Visual: make it stand out",
        notes: ["Draw a filled rectangle behind the text first so it's readable against any background:"],
        code: [" if enemies_remaining<=0 then", "  rect(74,52,80,16,12)", "  print(\"You Win!\",84,60,15)", " end"],
      },
      {
        title: "Structure: one counter, one place to change it",
        notes: ["enemies_remaining is declared once at the top with the rest of the game state, and only ever changed from inside the collision check - the same principle already used for enemies_alive."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined into the one TIC() function:"],
        code: [
          "enemies={}",
          "for enemy_count=0,4 do",
          " table.insert(enemies,{ex=20+enemy_count*20,ey=10,alive=true})",
          "end",
          "enemies_remaining=5",
          "has_won=false",
          "bullets={}",
          "",
          "function TIC()",
          " cls(0)",
          " if btn(3) then x=x+1 end",
          " if btnp(4) then",
          "  table.insert(bullets,{bx=x,by=y})",
          " end",
          " for i=#bullets,1,-1 do",
          "  local b=bullets[i]",
          "  b.by=b.by-2",
          "  local hit=false",
          "  for j=1,#enemies do",
          "   local e=enemies[j]",
          "   if e.alive and b.bx<e.ex+8 and b.bx+8>e.ex and b.by<e.ey+8 and b.by+8>e.ey then",
          "    e.alive=false",
          "    enemies_remaining=enemies_remaining-1",
          "    hit=true",
          "   end",
          "  end",
          "  if hit then",
          "   table.remove(bullets,i)",
          "  elseif b.by<0 then",
          "   table.remove(bullets,i)",
          "  else",
          "   spr(2,b.bx,b.by,0)",
          "  end",
          " end",
          " for j=1,#enemies do",
          "  local e=enemies[j]",
          "  if e.alive then spr(3,e.ex,e.ey,0) end",
          " end",
          " spr(1,x,y,0)",
          " if enemies_remaining<=0 then",
          "  if not has_won then",
          "   has_won=true",
          "   sfx(3)",
          "  end",
          "  rect(74,52,80,16,12)",
          "  print(\"You Win!\",84,60,15)",
          " end",
          "end",
        ],
      },
    ],
  },
  {
    engine: "001 Game Creator", file: "001GameCreator",
    l1Notes: ["A new Variable and a second Actor Trigger on the enemy Template - completely separate from Lesson 4's \"Collided with actor\" trigger:"],
    l1Code: [
      "New Variable: enemies_remaining, default value 5",
      "",
      "enemy Actor Trigger: Defeated / Destroyed ->",
      "  Variable – Modify Value(enemies_remaining, Subtract, 1)",
      "  If(enemies_remaining <= 0)",
      "    Message Box(\"You Win!\")",
    ],
    l2Notes: ["\"Defeated / Destroyed\" fires exactly once per Actor copy, no matter what destroyed it, and the If check only ever passes once enemies_remaining actually reaches zero - there's no path where the message could appear early, late, or more than once."],
    l3Ideas: [
      {
        title: "Sound: victory sound",
        notes: ["Add a Play Sound event before the Message Box, using the provided sfx_win.wav:"],
        code: [
          "enemy Actor Trigger: Defeated / Destroyed ->",
          "  Variable – Modify Value(enemies_remaining, Subtract, 1)",
          "  If(enemies_remaining <= 0)",
          "    Play Sound(sfx_win.wav)",
          "    Message Box(\"You Win!\")",
        ],
      },
      {
        title: "Visual: distinct message box style",
        notes: ["Check Visual Settings for message boxes for a background/border color option that reads as a celebration rather than ordinary dialogue - applied globally, or via any per-box style override your installed version supports."],
      },
      {
        title: "Structure: two triggers, two jobs",
        notes: ["\"Collided with actor\" handles the fight itself (destroying the bullet and the enemy) - \"Defeated / Destroyed\" only ever tracks overall progress. Neither trigger needs to know about the other."],
      },
      {
        title: "Complete Level 3 solution (all 2 goals combined)",
        notes: ["The Defeated/Destroyed trigger with sound added:"],
        code: [
          "enemy Actor Trigger: Defeated / Destroyed ->",
          "  Variable – Modify Value(enemies_remaining, Subtract, 1)",
          "  If(enemies_remaining <= 0)",
          "    Play Sound(sfx_win.wav)",
          "    Message Box(\"You Win!\")",
        ],
      },
    ],
  },
  {
    engine: "Unreal", file: "Unreal",
    l1Notes: ["An Event Dispatcher on BP_Enemy, bound once per spawned instance from the Level Blueprint's existing spawn loop, targeting a dedicated BP_GameManager instance placed in the level:"],
    l1Code: [
      "BP_Enemy's Event Graph (inside the successful Cast branch):",
      "Call OnEnemyDestroyed (Broadcast) -> Destroy Actor(Cast result) -> Destroy Actor(Self)",
      "",
      "Level Blueprint:",
      "Event BeginPlay -> (BP_GameManager reference, dragged in from the level) ->",
      "  For Loop -> Spawn Actor from Class(BP_Enemy, ...) ->",
      "    Bind Event to OnEnemyDestroyed(Target = spawned Actor, Event = BP_GameManager ref's HandleEnemyDestroyed)",
      "",
      "BP_GameManager:",
      "Custom Event HandleEnemyDestroyed ->",
      "  EnemiesRemaining = EnemiesRemaining - 1",
      "  Branch(EnemiesRemaining <= 0) -> True: Print String(\"You Win!\", Print to Screen = true)",
    ],
    l2Notes: ["Each spawned enemy's dispatcher is bound exactly once, at spawn time, always to the same BP_GameManager instance, and Call OnEnemyDestroyed only ever broadcasts once per enemy, right at the moment it's actually destroyed - there's no path where EnemiesRemaining could drift or the win message could double-fire."],
    l3Ideas: [
      {
        title: "Sound: victory sound",
        notes: ["Add a Play Sound 2D node using the provided sfx_win.wav right before the Print String node, inside BP_GameManager:"],
        code: ["Branch(EnemiesRemaining <= 0) -> True:", "  Play Sound 2D(sfx_win.wav)", "  Print String(\"You Win!\", Print to Screen = true)"],
      },
      {
        title: "Visual: styled win message",
        notes: ["Set Print String's Text Color to gold and increase Duration so it's clearly a celebratory moment, not a debug message:"],
        code: ["Print String(\"You Win!\", Print to Screen = true, Text Color = (Gold), Duration = 5.0)"],
      },
      {
        title: "Structure: BP_GameManager owns the counting",
        notes: ["EnemiesRemaining and HandleEnemyDestroyed live entirely inside BP_GameManager - the Level Blueprint only ever spawns enemies and binds their dispatchers, and BP_Enemy only ever broadcasts that it was destroyed. Neither one knows or cares how many enemies remain overall."],
      },
      {
        title: "Complete Level 3 solution (all 2 goals combined)",
        notes: ["HandleEnemyDestroyed with both ideas added:"],
        code: [
          "Custom Event HandleEnemyDestroyed ->",
          "  EnemiesRemaining = EnemiesRemaining - 1",
          "  Branch(EnemiesRemaining <= 0) -> True:",
          "    Play Sound 2D(sfx_win.wav)",
          "    Print String(\"You Win!\", Print to Screen = true, Text Color = (Gold), Duration = 5.0)",
        ],
      },
    ],
    altNotes: {
      title: "C++",
      notes: [
        "A DECLARE_DYNAMIC_MULTICAST_DELEGATE plus a BlueprintAssignable UPROPERTY on ACppEnemy is the code equivalent of a Blueprint Event Dispatcher - Broadcast() called from OnOverlap() mirrors Call OnEnemyDestroyed.",
        "The Level 3 goals map onto C++ calls the same way: UGameplayStatics::PlaySound2D() for sound, and GEngine->AddOnScreenDebugMessage() (with a custom FColor and duration) for the visual idea - both called from the GameManager actor's bound handler, mirroring BP_GameManager's role in the Blueprint path.",
      ],
      why: [
        "Showing both C++ and Blueprint side by side for the same behavior keeps demystifying Blueprints for students who assume \"visual\" means \"not real code\" - it's the same engine calls either way.",
        "This is squarely the optional advanced track - most students should stay on Blueprints; C++ is for students already comfortable with typed languages from Godot, GameMaker, or Unity's C# path.",
      ],
    },
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_06_Key_${d.file}.docx`]));

// =========================================================================
// WRITE ALL FILES
// =========================================================================
(async () => {
  for (const [doc, filename] of jobs) {
    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(filename, buf);
    console.log("wrote", filename);
  }
  console.log("done,", jobs.length, "files");
})();
