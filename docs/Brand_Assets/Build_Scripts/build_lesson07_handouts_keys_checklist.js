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
      headers: { default: brand.pageHeader("Lesson 7 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 7 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("New Unit: Core Mechanics"),
        p("Lessons 7-11 form the \"Core Mechanics\" unit (collision, scoring, win/lose states, simple UI), per the Course Syllabus's milestone table. Level 1 MVP is done as of Lesson 6 - everything from here builds toward the Level 2 Complete milestone at Lesson 11."),

        h2("Important: A Bugfix Lives In This Lesson"),
        p("Two engines - GameMaker and 001 Game Creator - need an actual correction to Lesson 6's win-counting code this week, not just an addition. Both used a generic \"this instance was destroyed, for any reason\" hook to count kills (GameMaker's Destroy event, 001 Game Creator's \"Defeated / Destroyed\" trigger). That was fine when the only way an enemy could be destroyed was by being shot - but starting this week, enemies can also be removed by escaping off the bottom of the screen, and that must NOT count as a kill. Both engines' Handouts and Keys walk students through moving the win-counter to the actual hit-detection code instead. Read through the GameMaker and 001 Game Creator Keys before class so this doesn't catch you by surprise."),
        checkbox("Read the GameMaker and 001 Game Creator Keys' Level 1 sections in full before class - both involve editing existing Lesson 6 logic, not just adding new code."),
        p("The other six engines (Scratch, Godot, Unity, GDevelop, TIC-80, Unreal) already tied their win-counter directly to the hit-detection code in Lesson 6, so they need no correction - only the new movement/escape-cleanup logic.", { italics: true, color: "555555" }),

        h2("Self-Contained Lesson Assets"),
        p("No new assets this week - enemy movement and cleanup use only code/logic changes."),

        h2("Continuity From Lesson 6"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 6, not create a new one."),
        checkbox("If a student's Lesson 6 project is lost or broken, have a pre-built \"Lesson 6 complete\" starter available per engine (from the Lesson 6 Teacher's Keys) so they aren't blocked."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop's Lesson 5 and 6 exercises still work for all 8 engines."),
        checkbox("GameMaker: walk through the Collision-event fix yourself once before class - it's easy to leave the old Destroy-event decrement in place by accident, double-counting kills."),
        checkbox("001 Game Creator (Windows demo laptop): confirm the \"Left/Entered Camera\" trigger reliably fires for an Actor leaving via the bottom of the screen, not just the top (as used for the bullet in Lesson 3) - same trigger, different edge, worth a quick test."),
        checkbox("Godot / Unity / GameMaker / GDevelop / Unreal: confirm a slow, continuous downward movement plus an off-screen cleanup check both work smoothly for several enemies at once."),
        checkbox("TIC-80: confirm the enemies table's per-enemy movement and bottom-of-screen cleanup don't interfere with the existing bullet-collision nested loop."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies."),

        h2("Known Gotchas This Week"),
        bullet("GameMaker / 001 Game Creator: the single most important thing to watch for this week - confirm the win-counter decrement actually moved to the hit-detection code, not left on the generic destroy hook. Test by letting an enemy escape off the bottom unshot and confirming \"You Win!\" does NOT appear."),
        bullet("Every engine: enemies moving continuously means testing takes a few seconds longer each time - remind students to be patient watching a test run through instead of assuming something's broken after 1 second."),
        bullet("A very slow enemy speed is fine and expected this week - there's no lose condition wired up yet (that's Lesson 8), so a fast enemy just means less time to observe and test."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_07_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 7: On the Move", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: enemies stop standing still. Your job is to make every enemy advance slowly toward the player, and clean up any that reach the bottom of the screen without being shot.", { bold: true }),
    p("Open your own [YourName]Game project from Lesson 6 - don't start a new one.", { bold: true }),
    p("Tip: no new assets this week - this is a logic-only lesson.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("Every enemy moves slowly downward, toward the player, instead of sitting still. An enemy that reaches the bottom of the screen unshot is removed - and does NOT count toward your win condition, since it escaped rather than being defeated.", { italics: true }),
    ...diagram("diagram_activity_enemy_removal.png", 6.3, 4.914, "The general escape-vs-hit distinction behind this week's exercise - the same shape in every engine, whatever the exact code or blocks look like."),
    ...data.level1.map(s => numbered(s)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  if (data.level1EngineDiagram) children.push(...diagram(data.level1EngineDiagram.file, data.level1EngineDiagram.w, data.level1EngineDiagram.h, data.level1EngineDiagram.caption));
  if (data.altPath) {
    children.push(h2("Alternate Path: " + data.altPath.title));
    children.push(p(data.altPath.intro, { italics: true, color: "555555" }));
    if (data.altPath.warning) children.push(p(data.altPath.warning, { bold: true, color: "B00020" }));
    children.push(...data.altPath.steps.map(s => numbered(s)));
    if (data.altPath.code) children.push(code(data.altPath.code));
    if (data.altPath.level1Diagram) children.push(...diagram(data.altPath.level1Diagram.file, data.altPath.level1Diagram.w, data.altPath.level1Diagram.h, data.altPath.level1Diagram.caption));
    if (data.altPath.note) children.push(p(data.altPath.note, { italics: true, color: "555555" }));
  }
  children.push(
    h2("Level 2 - Complete (Alpha/Beta)"),
    p("A moving enemy can still be shot and destroyed correctly, and an enemy that escapes off the bottom never triggers \"You Win!\" by accident.", { italics: true }),
    ...diagram("diagram_state_enemy.png", 6.3, 3.906, "The enemy's two possible endings this week - only one of them should count as a kill."),
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
      headers: { default: brand.pageHeader("Lesson 7 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: ["Open your [YourName]Game project from Lesson 6."],
    level1: [
      "On the Enemy sprite's \"when I start as a clone\" script, inside the forever loop, add \"change y by -1\" (Scratch's y-axis increases upward, so decreasing y moves an enemy down) - add this as the very first thing inside the loop, before the existing touching check.",
      "Still inside the forever loop, after the existing touching-check branch, add a second check: \"if <y position < -170> then delete this clone\" - this removes an enemy that escapes off the bottom, with no broadcast and no effect on enemies_remaining.",
      "Click the stage and let the enemies drift downward to test - shooting one should still work exactly as before, and letting one reach the bottom unshot should just make it vanish, with no \"You Win!\" message.",
    ],
    level2: [
      "Shoot a moving enemy to confirm collision still works correctly while it's in motion.",
      "Let at least one enemy escape off the bottom unshot, and confirm \"You Win!\" does NOT appear - only enemies_remaining, only ever changed inside GameManager's \"when I receive [enemy hit]\" script, controls that message. The Enemy's touching-Bullet branch broadcasts enemy hit but never touches enemies_remaining itself.",
    ],
    level3: [
      "Visual: give each enemy a slightly different downward speed based on enemy_number (its position in the row) at spawn time, so they don't all move in a perfectly straight line together.",
      "Structure: keep the two checks inside the forever loop clearly separate - one only ever runs on \"touched by a bullet\" (and only ever broadcasts enemy hit), the other only ever runs on \"reached the bottom\" (and never broadcasts anything). Never let one branch do both jobs.",
    ],
    trouble: [
      "Enemies don't move at all: confirm \"change y by -1\" is actually inside the forever loop, not outside it or in a script that only runs once.",
      "\"You Win!\" appears even when enemies escaped unshot: confirm the escape-cleanup branch never broadcasts enemy hit - only the touching-Bullet branch may do that, and only GameManager's script ever changes enemies_remaining.",
      "Enemies move but never disappear at the bottom: confirm the y position < -170 check is inside the same forever loop, using the correct comparison direction (less than, not greater than).",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: ["Open your [YourName]Game project from Lesson 6."],
    level1: [
      "In enemy.tscn's script, add var enemy_speed := 60.0 near the top.",
      "Add a _process(delta) function that moves the enemy downward and frees it (without emitting destroyed) if it goes past the bottom of the screen (code below).",
      "Run the scene and let the enemies drift downward - shooting one should still work exactly as before, and letting one reach the bottom unshot should just remove it, with no \"You Win!\" message.",
    ],
    level1Code: [
      "# enemy.tscn's script - add near the top and as a new function:",
      "var enemy_speed := 60.0",
      "",
      "func _process(delta):",
      "    position.y += enemy_speed * delta",
      "    if position.y > 700:",
      "        queue_free()",
    ],
    level2: [
      "Shoot a moving enemy to confirm the Hitbox's area_entered collision still works correctly while it's in motion.",
      "Let at least one enemy escape off the bottom unshot, and confirm \"You Win!\" does NOT appear - only the destroyed signal (emitted exclusively from _on_hitbox_area_entered) can ever decrement enemies_remaining, since the main scene connects that signal directly to $GameManager.on_enemy_destroyed().",
    ],
    level3: [
      "Visual: give each enemy a slightly randomized enemy_speed at spawn time (e.g. randf_range(50.0, 80.0) set on the instance right after instantiate()) so they don't all move in perfect lockstep.",
      "Structure: keep _process()'s escape-cleanup call as a plain queue_free() only - never let it call destroyed.emit(), which is reserved exclusively for _on_hitbox_area_entered().",
    ],
    trouble: [
      "Enemies don't move at all: confirm _process(delta) is spelled exactly right (case-sensitive) and is a new function, not accidentally merged into _ready().",
      "\"You Win!\" appears even when enemies escaped unshot: confirm _process()'s escape cleanup calls queue_free() only, with no destroyed.emit() anywhere near it.",
      "Enemies move but never disappear at the bottom: confirm the position.y > 700 threshold actually matches your project's visible screen height - adjust the number if your window size differs.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: ["Open your [YourName]Game project from Lesson 6."],
    level1: [
      "In Enemy.cs, add a public float enemySpeed and an Update() method that moves the enemy downward and destroys it (without calling GameManager.RegisterKill()) if it goes past the bottom of the screen (code below).",
      "Press Play and let the enemies drift downward - shooting one should still work exactly as before, and letting one reach the bottom unshot should just destroy it, with no \"You Win!\" message.",
    ],
    level1Code: [
      "// Enemy.cs - add a field and a new method:",
      "public float enemySpeed = 1.5f;",
      "",
      "void Update() {",
      "    transform.Translate(Vector3.down * enemySpeed * Time.deltaTime);",
      "    if (transform.position.y < -6f) {",
      "        Destroy(gameObject);",
      "    }",
      "}",
    ],
    level2: [
      "Shoot a moving enemy to confirm Bullet.cs's OnTriggerEnter2D still fires correctly while it's in motion.",
      "Let at least one enemy escape off the bottom unshot, and confirm \"You Win!\" does NOT appear - only Bullet.cs's OnTriggerEnter2D calls GameManager.RegisterKill(), and Enemy.cs's own Update() never calls it.",
    ],
    level3: [
      "Visual: give each enemy a slightly randomized enemySpeed in Start() (e.g. enemySpeed = Random.Range(1f, 2f)) so they don't all move in perfect lockstep.",
      "Structure: keep Enemy.cs's own Update() cleanup as a plain Destroy(gameObject) only - never let it call GameManager.RegisterKill(), which is reserved exclusively for an actual hit in Bullet.cs.",
    ],
    trouble: [
      "Enemies don't move at all: confirm Update() is spelled exactly right (case-sensitive) and enemySpeed is actually greater than 0.",
      "\"You Win!\" appears even when enemies escaped unshot: confirm Enemy.cs's Update() never calls GameManager.RegisterKill() - that call belongs exclusively in Bullet.cs.",
      "Enemies move but never disappear at the bottom: confirm the transform.position.y < -6f threshold actually matches your camera's visible area - adjust the number if your camera setup differs.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same movement-and-cleanup pattern works in a Script Graph instead of C#.",
      steps: [
        "Open the Enemy prefab's Script Graph (the same graph carrying its existing On Trigger Enter logic from earlier lessons) and add an Update event node - it fires once per frame, the same as Update() in the C# path.",
        "Feed the Update event into a \"Translate\" node (Space: World), with its Translation input set to Vector3.down * enemySpeed * deltaTime - the same expression as transform.Translate(Vector3.down * enemySpeed * Time.deltaTime) in the code path.",
        "From the same Update event, add a Branch node whose Condition compares the Enemy's transform.position.y against your off-screen threshold (transform.position.y < -6f), with its True output feeding a Destroy node targeting itself.",
        "Confirm the Destroy node here is a plain destroy, with no call into the GameManager's kill-registering logic - that stays exclusive to the Bullet's own graph, exactly as in the C# path.",
      ],
      level1Diagram: { file: "diagram_unity_visual_scripting_l7_level1.png", w: 6.3, h: 4.41, caption: "The Enemy prefab's Update graph, in Unity Visual Scripting - the node-for-node equivalent of the Level 1 code." },
      note: "Level 2/3 already work the same way as the C# path, since the escape-vs-hit distinction lives in which graph calls Destroy, not in any engine-specific detail.",
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: [
      "Open your [YourName]Game project from Lesson 6.",
      "IMPORTANT - fixing a hidden bug from Lesson 6: obj_enemy's Destroy event currently decreases enemies_remaining whenever ANY copy of obj_enemy is destroyed, for any reason. That was fine when shooting was the only way an enemy could be destroyed - but this week, enemies can also be removed by escaping off the bottom of the screen, and that must NOT count as a kill. Delete the code currently inside obj_enemy's Destroy event entirely - it's being replaced below.",
    ],
    level1: [
      "Add a Create event to obj_enemy that sets enemy_speed = 1; and a Step event that uses it to move the instance downward, destroying it if it passes the bottom of the room (code below) - this is a plain destroy, unrelated to win-counting.",
      "On obj_bullet's existing Collision event with obj_enemy, add the win-counter decrement there instead, right alongside the existing instance_destroy() calls (code below) - this is the ONLY place enemies_remaining should ever change now, and it lives on obj_gamemanager, the dedicated manager object introduced in Lesson 6.",
      "Press F5 (Run) and let the enemies drift downward - shooting one should still work exactly as before, and letting one reach the bottom unshot should just destroy it, with no \"You Win!\" message.",
    ],
    level1Code: [
      "// obj_enemy's Create event (new):",
      "enemy_speed = 1;",
      "",
      "// obj_enemy's Step event (new):",
      "y += enemy_speed;",
      "if (y > room_height) {",
      "    instance_destroy();",
      "}",
      "",
      "// obj_bullet's Collision event with obj_enemy (revised):",
      "audio_play_sound(sfx_hit, 1, false);",
      "obj_gamemanager.enemies_remaining -= 1;",
      "if (obj_gamemanager.enemies_remaining <= 0) {",
      "    obj_gamemanager.has_won = true;",
      "}",
      "instance_destroy(other);",
      "instance_destroy(self);",
      "",
      "// obj_enemy's Destroy event: now empty - delete it, or leave it blank.",
    ],
    level2: [
      "Shoot a moving enemy to confirm the Collision event still fires correctly while it's in motion.",
      "Let at least one enemy escape off the bottom unshot, and confirm \"You Win!\" does NOT appear - only obj_bullet's Collision event can ever change obj_gamemanager's enemies_remaining now, and the Step event's escape cleanup never touches it.",
    ],
    level3: [
      "Visual: give each obj_enemy a slightly randomized enemy_speed in its Create event (e.g. enemy_speed = random_range(0.7, 1.3);) instead of the fixed value, so they don't all move in perfect lockstep.",
      "Structure: keep obj_enemy's Step event limited to movement and escape cleanup only - it should never reference or change enemies_remaining. That variable only ever changes from obj_bullet's Collision event, on obj_gamemanager, now.",
    ],
    trouble: [
      "\"You Win!\" appears even when enemies escaped unshot: this means the old Destroy-event decrement wasn't fully removed - double-check obj_enemy's Destroy event is actually empty.",
      "Enemies aren't destroyed when shot anymore: confirm the win-counter code was added to obj_bullet's Collision event, not accidentally left out when you moved it.",
      "Enemies don't move: confirm the Step event is attached to obj_enemy (not obj_spawner or obj_bullet), and that y += enemy_speed; actually runs every step with enemy_speed set in the Create event.",
    ],
    resource: "Student Resource List > GameMaker section, or gamemaker.io/en/tutorials",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same movement-and-cleanup pattern, and the same win-counter fix, work with drag-and-drop blocks instead of typed code.",
      steps: [
        "On obj_enemy's Create event, add a \"Set Variable\" block setting enemy_speed to 1. On obj_enemy's Step event, switch to GML Visual: a \"Jump To Point\" block (Movement category) with Relative checked for Y, X set to 0 and Y set to enemy_speed - prefer this purpose-built block over the generic Assign Variable/Set Variable for moving a built-in coordinate. Then an \"If Variable\" block comparing y against room_height, with an \"Instance - Destroy Instance\" block inside it.",
        "On obj_bullet's Collision event with obj_enemy, add the enemies_remaining blocks there (Set Variable targeting obj_gamemanager's enemies_remaining, If Variable checking <= 0), same as the code path.",
        "Delete any blocks remaining in obj_enemy's Destroy event.",
      ],
      level1Diagram: { file: "diagram_gml_visual_l7_level1.png", w: 6.3, h: 5.355, caption: "obj_enemy's Step event and obj_bullet's Collision event, in GML Visual - the block-for-block equivalent of the Level 1 code and the win-counter fix." },
      note: "Level 2/3's escape-vs-hit distinction lives in which event contains the win-counter logic, the same in both paths.",
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: ["Open your [YourName]Game project from Lesson 6."],
    level1: [
      "Add a new \"always\" event (no condition): Action \"Enemy.Y = Enemy.Y() + 30 * TimeDelta()\" - this moves every Enemy instance downward, every frame.",
      "Add a second new event: Condition \"Enemy.Y() > 750\" -> Action \"Delete Enemy\" - this is a plain deletion, with no effect on EnemiesRemaining.",
      "Click the green Preview button and let the enemies drift downward - shooting one should still work exactly as before, and letting one reach the bottom unshot should just delete it, with no \"You Win!\" message.",
    ],
    level1Code: [
      "(always, no condition)",
      "Action: Enemy.Y = Enemy.Y() + 30 * TimeDelta()",
      "",
      "Condition: Enemy.Y() > 750",
      "Action: Delete Enemy",
    ],
    level2: [
      "Shoot a moving enemy to confirm the Combat event group's collision condition still fires correctly while it's in motion.",
      "Let at least one enemy escape off the bottom unshot, and confirm \"You Win!\" does NOT appear - only the Combat group's collision event ever changes EnemiesRemaining, and the new escape-cleanup event never touches it.",
    ],
    level3: [
      "Visual: use a random per-instance speed instead of a fixed 30 - GDevelop's \"Enemy.Y = Enemy.Y() + (30 + Random(20)) * TimeDelta()\" gives each instance a slightly different feel without extra events.",
      "Structure: group the movement and escape-cleanup events into their own labeled group (e.g. \"Enemy Movement\"), separate from \"Combat\" and \"Spawn Enemies\" - three clearly separate concerns as the event sheet grows.",
    ],
    trouble: [
      "\"You Win!\" appears even when enemies escaped unshot: confirm the escape-cleanup event only has a \"Delete Enemy\" action, with no EnemiesRemaining action anywhere near it.",
      "Enemies don't move: confirm the movement event truly has no condition attached (an empty condition list, not a false one) so it runs every single frame.",
      "Enemies move but never disappear at the bottom: confirm the Enemy.Y() > 750 threshold actually matches your project's scene height.",
    ],
    resource: "Student Resource List > GDevelop section, or gdevelop.io/tutorials",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: ["Launch TIC-80 and load your cart from Lesson 6.", "Press Esc to reach the code editor."],
    level1: [
      "In the enemy-drawing loop at the bottom of TIC(), add movement (enemy.ey=enemy.ey+0.3) before the alive check, and a new check that sets enemy.alive=false if an enemy passes the bottom of the screen - without touching enemies_remaining (code below).",
      "Press Esc, type run, press Enter, then let the enemies drift downward to test - shooting one should still work exactly as before, and letting one reach the bottom unshot should just make it vanish, with no \"You Win!\" message.",
    ],
    level1Code: [
      "for enemy_index=1,#enemies do",
      " local enemy=enemies[enemy_index]",
      " if enemy.alive then",
      "  enemy.ey=enemy.ey+0.3",
      "  if enemy.ey>136 then",
      "   enemy.alive=false",
      "  else",
      "   spr(3,enemy.ex,enemy.ey,0)",
      "  end",
      " end",
      "end",
    ],
    level2: [
      "Shoot a moving enemy to confirm the bullets-vs-enemies collision check still works correctly while it's in motion.",
      "Let at least one enemy escape off the bottom unshot, and confirm \"You Win!\" does NOT appear - only the collision branch inside the bullets loop (guarded by enemy.alive and the overlap check) ever decrements enemies_remaining, and the movement loop's escape check never touches it.",
    ],
    level3: [
      "Visual: give each enemy a slightly different descent speed based on its table index enemy_index (e.g. enemy.ey=enemy.ey+0.2+enemy_index*0.05) so they don't all move in perfect lockstep.",
      "Structure: keep the movement/escape-cleanup loop and the bullets-vs-enemies collision loop as two clearly separate blocks (as shown above) - resist merging them into one loop even though both touch the enemies table.",
    ],
    trouble: [
      "\"You Win!\" appears even when enemies escaped unshot: confirm enemies_remaining is decremented only inside the bullets loop's collision branch, never inside the enemy-drawing loop.",
      "Enemies don't move: confirm enemy.ey=enemy.ey+0.3 is inside the if enemy.alive then branch of the drawing loop, not skipped or placed outside it.",
      "Enemies move but never disappear at the bottom: confirm the enemy.ey>136 threshold actually matches TIC-80's 136-pixel-tall default screen - adjust if you changed resolution.",
    ],
    resource: "Student Resource List > TIC-80 section, or tic80.com",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 6.",
      "IMPORTANT - fixing a hidden bug from Lesson 6: the enemy Actor Template's \"Defeated / Destroyed\" trigger currently decreases enemies_remaining whenever ANY copy of the enemy Actor is destroyed, for any reason. That was fine when being shot was the only way an enemy could be destroyed - but this week, enemies can also leave the screen by moving off the bottom, and that must NOT count as a kill. Delete the \"Defeated / Destroyed\" trigger from the enemy Actor Template entirely - it's being replaced below.",
    ],
    level1: [
      "On the enemy Actor Template's Automatic Movement settings, tick \"Move\" and set its direction to Down, at a slow, comfortable speed - the exact same feature the bullet already used for its own upward movement in Lesson 3.",
      "Add a new Actor Trigger to the enemy Template: \"Left/Entered Camera\" (the Left case) -> a script with an \"Actor – Destroy/Revive Actor\" event (Destroy) targeting itself - the exact same trigger the bullet already used for its own off-screen cleanup in Lesson 3.",
      "On the enemy Template's existing \"Collided with actor\" trigger, add the win-counter decrement there instead, right alongside the existing Destroy events (code below) - this is the ONLY place enemies_remaining should ever change now.",
      "Save, then Play/Test and let the enemies drift downward - shooting one should still work exactly as before, and letting one reach the bottom unshot should just remove it, with no \"You Win!\" message.",
    ],
    level1Code: [
      "enemy Actor Template settings:",
      "  Automatic Movement -> Move = true, Direction = Down",
      "",
      "enemy Actor Trigger: Left/Entered Camera (Left case) ->",
      "  Actor – Destroy/Revive Actor(this, Destroy)",
      "",
      "enemy Actor Trigger: Collided with actor (revised) ->",
      "  Variable – Modify Value(enemies_remaining, Subtract, 1)",
      "  If(enemies_remaining <= 0)",
      "    Message Box(\"You Win!\")",
      "  Actor – Destroy/Revive Actor(relate, Destroy)",
      "  Actor – Destroy/Revive Actor(this, Destroy)",
      "",
      "enemy Actor Trigger: Defeated / Destroyed -> deleted entirely",
    ],
    level2: [
      "Shoot a moving enemy to confirm \"Collided with actor\" still fires correctly while it's in motion.",
      "Let at least one enemy escape off the bottom of the camera unshot, and confirm \"You Win!\" does NOT appear - only \"Collided with actor\" can ever change enemies_remaining now, and \"Left/Entered Camera\" never touches it.",
    ],
    level3: [
      "Visual: give a couple of enemy copies a slightly different Automatic Movement speed (per-instance override) so they don't all move in perfect lockstep.",
      "Structure: keep the enemy Template down to exactly two Triggers now - \"Collided with actor\" (the fight, including win-counting) and \"Left/Entered Camera\" (plain escape cleanup). Neither one references the other's job.",
    ],
    trouble: [
      "\"You Win!\" appears even when enemies escaped unshot: this means \"Defeated / Destroyed\" wasn't fully removed - double check it's actually deleted from the enemy Template, not just disabled.",
      "Enemies don't move: confirm \"Move\" is ticked in the enemy Template's Automatic Movement window, with Direction actually set to Down and a nonzero speed.",
      "Enemies vanish immediately at spawn: confirm \"Left/Entered Camera\" is checking the \"Left\" case specifically (exited view), not \"Entered\" - the wrong case fires the moment they first appear.",
    ],
    resource: "001gamecreator.com/wiki, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: ["Open your [YourName]Game project from Lesson 6."],
    level1: [
      "In BP_Enemy, add a float variable EnemySpeed (default around 50).",
      "In BP_Enemy's Event Graph, add an Event Tick node feeding an \"Add Actor World Offset\" node. Right-click its Delta Location pin and choose \"Split Struct Pin\" first (it starts collapsed as one Vector pin, and won't accept a value into just Y until it's split), then wire -EnemySpeed * DeltaSeconds into Y (negative, since this course's Y+ convention is \"up\"). Add a Branch checking its own Y location against an off-screen-at-the-bottom threshold, feeding a plain \"Destroy Actor\" node on True - not the OnEnemyDestroyed dispatcher.",
      "Click Compile, then Play and let the enemies drift downward - shooting one should still work exactly as before, and letting one reach the bottom unshot should just destroy it, with no \"You Win!\" message.",
    ],
    level1Code: [
      "BP_Enemy's Event Graph (new):",
      "Event Tick ->",
      "  Add Actor World Offset(Self, DeltaLocation = (0, -EnemySpeed*DeltaSeconds, 0))",
      "Event Tick ->",
      "  Branch(GetActorLocation().Y < -400) -> True: Destroy Actor(Self)",
      "  (plain Destroy Actor here - do NOT call OnEnemyDestroyed)",
    ],
    level2: [
      "Shoot a moving enemy to confirm Event ActorBeginOverlap still fires correctly while it's in motion.",
      "Let at least one enemy escape off the bottom unshot, and confirm \"You Win!\" does NOT appear - only the successful-Cast branch (a real hit) ever calls OnEnemyDestroyed, and the escape-cleanup Branch never does.",
    ],
    level3: [
      "Visual: randomize EnemySpeed slightly in BeginPlay (e.g. EnemySpeed = FMath::RandRange(40.f, 70.f) via a Random Float in Range node) so enemies don't all move in perfect lockstep.",
      "Structure: keep BP_Enemy's escape-cleanup Branch calling only Destroy Actor, with the OnEnemyDestroyed broadcast staying exclusively inside the overlap event's successful-Cast branch.",
    ],
    trouble: [
      "Enemies don't move at all: confirm Event Tick is actually wired to the Add Actor World Offset node, and EnemySpeed is greater than 0.",
      "Can't wire -EnemySpeed*DeltaSeconds into just the Y of Delta Location - it only shows one combined Vector pin: right-click that pin and choose \"Split Struct Pin\" to break it into separate X/Y/Z pins first, the same step Lessons 2 and 3 needed.",
      "\"You Win!\" appears even when enemies escaped unshot: confirm the escape-cleanup Branch's True case only has a plain Destroy Actor node, with no Call OnEnemyDestroyed nearby.",
      "Enemies move but never disappear at the bottom: confirm the GetActorLocation().Y < -400 threshold actually matches your level's visible play area.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    level1EngineDiagram: { file: "diagram_unreal_blueprint_l7_level1.png", w: 6.3, h: 4.536, caption: "BP_Enemy's Event Tick graph, in Unreal Blueprint - the node-for-node equivalent of the Level 1 code." },
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same movement-and-cleanup pattern, and the same escape-vs-hit distinction, work in C++ instead of Blueprint nodes.",
      steps: [
        "In ACppEnemy's Tick(), call AddActorWorldOffset() along -Y, then check GetActorLocation().Y against your off-screen threshold and call a plain Destroy() when past it - not OnEnemyDestroyed.Broadcast().",
        "Confirm OnEnemyDestroyed.Broadcast() still only appears inside OnOverlap(), never inside Tick().",
      ],
      code: [
        "// ACppEnemy.cpp",
        "void ACppEnemy::Tick(float DeltaSeconds) {",
        "    Super::Tick(DeltaSeconds);",
        "    AddActorWorldOffset(FVector(0.f, -EnemySpeed * DeltaSeconds, 0.f));",
        "    if (GetActorLocation().Y < -400.f) {",
        "        Destroy();",
        "    }",
        "}",
      ],
      note: "This mirrors the Blueprint path's Event Tick and escape Branch - same underlying engine calls, just from code. The escape-vs-hit distinction lives in which function calls Destroy(), the same in both paths.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_07_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 7 Teacher's Key: On the Move", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine + " - Instructor Reference Only", size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    h2("What This Lesson Teaches"),
    p(data.teaches),
  ];
  if (data.bugfixNote) {
    children.push(h2("The Bug Being Fixed"));
    children.push(p(data.bugfixNote));
  }
  children.push(
    h2("Level 1 - Full Worked Solution"),
    ...diagram("diagram_activity_enemy_removal.png", 6.3, 4.914, "The general escape-vs-hit distinction behind this week's exercise - the same shape in every engine, whatever the exact code or blocks look like."),
    ...data.level1.map(s => bullet(s)),
  );
  if (data.level1Code) children.push(code(data.level1Code));
  if (data.level1EngineDiagram) children.push(...diagram(data.level1EngineDiagram.file, data.level1EngineDiagram.w, data.level1EngineDiagram.h, data.level1EngineDiagram.caption));
  children.push(
    h2("Why This Structure"),
    ...data.why.map(s => bullet(s)),
    ...diagram("diagram_state_enemy.png", 6.3, 3.906, "The enemy's two possible endings this week - only one of them should count as a kill."),
    h2("Level 2 / Level 3 Talking Points"),
    ...data.talkingPoints.map(s => bullet(s)),
    h2("Common Student Mistakes"),
    ...data.mistakes.map(s => bullet(s)),
  );
  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 7 Key - " + data.engine + " (Instructor Only)") },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    teaches: "Continuous per-frame movement inside a forever loop, and a second independent condition (position threshold) inside that same loop that removes a clone without ever broadcasting enemy hit. This is the first lesson where two separate \"remove this instance\" pathways coexist in the same script.",
    level1: [
      "\"change y by -1\" goes first in the forever loop, before the touching check, so movement happens every frame regardless of collision state.",
      "The bottom-of-screen check (\"y position < -170\") is a second, independent if-block in the same forever loop, using delete this clone with no broadcast and no reference to enemies_remaining.",
      "Scratch's stage is 360 units tall (y from -180 to 180); -170 leaves a small margin so the enemy is fully off-stage before deletion, avoiding a visible pop.",
    ],
    why: [
      "Scratch's Lesson 6 retrofit already moved win-counting off the Enemy sprite entirely, onto a dedicated GameManager sprite that listens for an enemy hit broadcast - so this lesson's escape-cleanup check only has to avoid ever broadcasting that message, not avoid touching a shared variable directly. There's no 'on clone deleted' event for GameManager to accidentally overload the way GameMaker and 001GC's generic destroy hooks can be.",
      "Keeping both checks in the same forever loop (rather than splitting into two scripts) is fine here because each check's action (delete this clone, with only one of the two also broadcasting) is self-contained and neither one depends on the other's outcome.",
    ],
    talkingPoints: [
      "Ask students: what would happen if the y<-170 check ran before the touching check instead of after? (Nothing bad in this case, since both branches end in delete this clone and only one runs per frame - a good moment to discuss when ordering matters and when it doesn't.)",
      "Level 3 visual: varying speed by enemy_number produces a fanned-out descent that reads as more three-dimensional/organic than uniform movement.",
    ],
    mistakes: [
      "Placing \"change y by -1\" outside the forever loop (runs once at clone creation, enemy never actually moves).",
      "Using > instead of < in the bottom-check comparison (enemy deletes itself immediately since starting y is greater than -170).",
      "Forgetting Scratch's y-axis convention (up is positive) and moving enemies up off the top instead of down.",
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    teaches: "A second script pathway for removing a node (queue_free() called directly from the enemy's own _process(), vs. the destroyed signal emitted only from the hitbox collision callback) - the clearest illustration in the course so far of 'not every destruction is a kill.'",
    bugfixNote: "Godot's Lesson 6 design already emitted destroyed exclusively from _on_hitbox_area_entered(), so there is no bug to fix here - only new code to add. This engine is a good one to point to as the 'already did it right' example when explaining the GameMaker/001GC fix to the class.",
    level1: [
      "var enemy_speed := 60.0 is a plain per-instance property; delta-scaled movement (enemy_speed * delta) keeps motion frame-rate independent, consistent with movement patterns taught since Lesson 2.",
      "_process(delta) runs every frame automatically once defined - no manual wiring needed, unlike Unity's Update() which behaves the same way but is worth contrasting since students are comparing engines.",
      "The queue_free() escape path is a completely separate code path from destroyed.emit() - _process() never touches the signal at all.",
    ],
    why: [
      "Signals are still the only thing GameManager listens to for win-counting (the main scene connects each new enemy's destroyed signal to $GameManager.on_enemy_destroyed() at spawn time, per the Lesson 6 retrofit), and _process()'s cleanup simply doesn't use that mechanism - there's no shared hook to misuse the way GameMaker's Destroy event or 001GC's Defeated trigger can be.",
      "This is a good moment to name the general principle explicitly for students: 'ask what specifically happened, not just that something happened, before deciding whether it counts.'",
    ],
    talkingPoints: [
      "Level 3 structure talking point: this is the cleanest possible illustration of the escape-vs-hit distinction, since the two removal paths use genuinely different underlying mechanisms (direct call vs. signal) rather than the same mechanism used two different ways.",
      "If a student asks 'why not just check enemies_remaining inside _process() too, for safety' - explain that's redundant and actually risks re-introducing exactly the bug GameMaker/001GC has, by giving _process() any awareness of win-counting at all.",
    ],
    mistakes: [
      "Calling destroyed.emit() from _process()'s cleanup 'to be thorough' - reintroduces the exact bug this lesson is designed to avoid.",
      "Forgetting delta in the movement line (speed becomes frame-rate dependent, moves differently on different machines).",
      "Using a fixed pixel threshold that doesn't match the project's actual viewport height - remind students to check Project Settings > Display > Window if 700 doesn't look right.",
    ],
  },
  {
    engine: "Unity", file: "Unity",
    teaches: "The same escape-vs-hit distinction as Godot, expressed via Destroy(gameObject) called directly from the enemy's own Update() vs. GameManager.RegisterKill() called only from the bullet's OnTriggerEnter2D.",
    bugfixNote: "Unity's Lesson 6 design already called RegisterKill() exclusively from Bullet.cs's OnTriggerEnter2D, so there is no bug to fix here - only new code to add, same as Godot.",
    level1: [
      "transform.Translate(Vector3.down * enemySpeed * Time.deltaTime) is frame-rate-independent movement, consistent with patterns taught since Lesson 2.",
      "The Y threshold (-6f) assumes the default Lesson 1-3 orthographic camera setup; note in class if a student customized their camera size.",
      "Destroy(gameObject) in Enemy.cs's own Update() never calls GameManager.RegisterKill() - that call stays exclusively in Bullet.cs.",
    ],
    why: [
      "Same principle as Godot: RegisterKill() is only ever invoked from the one code path that actually represents a hit (Bullet.cs's trigger callback), and Enemy.cs's own escape-cleanup code has no reason to know that method exists.",
      "Worth pointing out to students that this is a general engineering habit worth keeping for the rest of the course: when adding a new way for something to end, check whether existing 'on end' code was written assuming there was only one way.",
    ],
    level1EngineDiagram: { file: "diagram_unity_visual_scripting_l7_level1.png", w: 6.3, h: 4.41, caption: "The Enemy prefab's Update graph, in Unity Visual Scripting - the node-for-node equivalent of the Level 1 code." },
    talkingPoints: [
      "Level 3 structure: same talking point as Godot - ask students to explain in their own words why Enemy.cs shouldn't call RegisterKill() even though it 'also' destroys the enemy.",
      "For strong students: ask what would happen if Bullet.cs's OnTriggerEnter2D destroyed the enemy but Update() also destroyed it the same frame (a double-destroy race) - a good lead-in to why Level 3 error-hardening exists later in the course.",
    ],
    mistakes: [
      "Calling GameManager.RegisterKill() from Enemy.cs's Update() 'just in case' - same bug pattern as Godot, explicitly avoid.",
      "Using Time.deltaTime incorrectly (multiplying twice, or forgetting it) resulting in wildly wrong speeds on different frame rates.",
      "Not adjusting the -6f threshold for a customized camera - enemies either never despawn or despawn while still visible.",
    ],
  },
  {
    engine: "GameMaker", file: "GameMaker",
    teaches: "The core lesson-7 bugfix: GameMaker's Destroy event fires for ANY instance_destroy() call regardless of cause, so it cannot be trusted for win-counting once a second destruction pathway (off-screen escape) exists. The fix moves the counter to the one event that actually represents a hit.",
    bugfixNote: "This is the primary bugfix engine for this lesson, alongside 001 Game Creator. Lesson 6 wired enemies_remaining -= 1 into obj_enemy's Destroy event. That event fires on literally any instance_destroy() call against an obj_enemy instance - including the new escape-cleanup instance_destroy() this lesson adds inside obj_enemy's Step event. Left unfixed, an enemy escaping unshot would still decrement enemies_remaining and could trigger a false \"You Win!\". The fix: delete the Destroy event's logic entirely, and add the decrement (plus the win check) into obj_bullet's Collision event with obj_enemy instead - the one place that actually represents a hit, targeting obj_gamemanager, the dedicated manager object Lesson 6 introduced for exactly this kind of shared state.",
    level1: [
      "obj_enemy's Step event handles movement (y += enemy_speed) and escape cleanup (instance_destroy() when y > room_height) - a plain destroy with zero awareness of win-counting.",
      "obj_bullet's Collision event with obj_enemy is now the single source of truth for obj_gamemanager's enemies_remaining, right alongside the sound effect and the two instance_destroy() calls already there from Lesson 4/6.",
      "obj_enemy's Destroy event is now empty (or deleted outright) - it has no job left to do.",
    ],
    level1Code: [
      "// obj_enemy Create event:",
      "enemy_speed = 1;",
      "",
      "// obj_enemy Step event:",
      "y += enemy_speed;",
      "if (y > room_height) {",
      "    instance_destroy();",
      "}",
      "",
      "// obj_bullet Collision event with obj_enemy:",
      "audio_play_sound(sfx_hit, 1, false);",
      "obj_gamemanager.enemies_remaining -= 1;",
      "if (obj_gamemanager.enemies_remaining <= 0) {",
      "    obj_gamemanager.has_won = true;",
      "}",
      "instance_destroy(other);",
      "instance_destroy(self);",
    ],
    why: [
      "GameMaker's Destroy event is a lifecycle hook, not a gameplay event - it fires on ANY instance_destroy(), which is exactly why it was safe to use in Lesson 6 (only one destruction path existed) and exactly why it breaks in Lesson 7 (a second path was added).",
      "Collision events are inherently specific - obj_bullet's Collision event with obj_enemy can only ever fire when those two specific instance types actually overlap, which is a much stronger and more accurate signal for 'a kill happened' than 'something destroyed an enemy.'",
      "This mirrors the exact same principle taught with Scratch's per-clone touching check in Lesson 5 and Godot/Unity's signal-vs-direct-call distinction above: attach state-changing logic to the specific event that means what you think it means, not to a generic lifecycle hook that fires for unrelated reasons too.",
    ],
    level1EngineDiagram: { file: "diagram_gml_visual_l7_level1.png", w: 6.3, h: 5.355, caption: "obj_enemy's Step event and obj_bullet's Collision event, in GML Visual - the block-for-block equivalent of the Level 1 code and the win-counter fix." },
    talkingPoints: [
      "This is a good lesson to spend extra teacher-led time on, since it's a genuine bug students would have shipped in Lesson 6 without knowing it (the bug was latent - correct as of Lesson 6, broken only once Lesson 7's second destruction path exists).",
      "Ask students to predict, before testing: 'if we only fixed the Collision event but left the old Destroy event decrement in place, what would happen?' (Answer: double-counting - an enemy that's shot would decrement enemies_remaining twice, once from each event, causing the win condition to trigger after only half the enemies are actually destroyed.)",
      "Level 3 structure talking point: obj_enemy's Step event should never need to know enemies_remaining exists at all - if a student's Step event references it, that's a sign the fix wasn't fully applied.",
    ],
    mistakes: [
      "Leaving the old decrement in obj_enemy's Destroy event AND adding the new one to obj_bullet's Collision event - causes double-counting (see talking point above). This is the single most likely mistake this lesson; test for it explicitly.",
      "Moving the decrement to the Collision event but forgetting to delete/empty the Destroy event.",
      "Placing the escape-cleanup instance_destroy() in the Collision event instead of the Step event (unrelated concerns merged back together).",
    ],
  },
  {
    engine: "GDevelop", file: "GDevelop",
    teaches: "GDevelop's Lesson 6 design already tied EnemiesRemaining to a specific collision condition in the Combat event group, so this lesson only adds new events - no fix needed - but it's a good moment to reinforce why that structure paid off.",
    bugfixNote: "No bugfix needed for GDevelop - the Combat group's collision event was always the only place EnemiesRemaining changes, and the new escape-cleanup event (added to a new or existing event group) has no reason to touch it.",
    level1: [
      "The always-true movement event applies Enemy.Y() + 30 * TimeDelta() to every Enemy instance each frame - GDevelop applies per-instance events to every existing instance of that object automatically, no explicit loop needed.",
      "The escape-cleanup event's condition (Enemy.Y() > 750) and action (Delete Enemy) are new, independent of the Combat group.",
      "TimeDelta() keeps movement frame-rate independent, consistent with patterns taught since Lesson 2.",
    ],
    why: [
      "GDevelop's per-object event model (an event's actions apply to every matching instance automatically) means the movement and escape-cleanup events don't need to reference EnemiesRemaining at all - they operate purely on position.",
      "The Combat group's collision condition remains the only event in the entire sheet that changes EnemiesRemaining - worth pointing this out explicitly as the reason nothing broke.",
    ],
    talkingPoints: [
      "Ask students to locate every place in their event sheet that changes EnemiesRemaining before and after this lesson - it should still be exactly one place.",
      "Level 3 structure: grouping movement/escape-cleanup into a labeled 'Enemy Movement' group (separate from 'Combat' and 'Spawn Enemies') previews the kind of organization a much larger event sheet will need by Lesson 11.",
    ],
    mistakes: [
      "Adding a condition to the movement event 'so it doesn't run all the time' - it should have no condition; that's what makes it run every frame.",
      "Using Scene.TimeDelta() instead of TimeDelta() by itself, or omitting it entirely (speed becomes frame-rate dependent).",
      "Setting the Enemy.Y() > 750 threshold using the wrong axis direction and having enemies vanish immediately.",
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    teaches: "The same escape-vs-hit distinction expressed as two clearly separate loops over the same enemies table - one that only ever sets enemy.alive=false for movement/escape reasons, and a separate one (already existing since Lesson 4) that both sets enemy.alive=false AND decrements enemies_remaining for an actual hit. This lesson also retrofits the drawing loop's variable names (j/e) to enemy_index/enemy, matching the meaningful-name standard the rest of the course already follows.",
    bugfixNote: "No bugfix needed for TIC-80 - Lesson 6's win-counter decrement already lived only inside the bullets-vs-enemies collision loop, and this lesson's new movement/escape logic is added to the separate enemy-drawing loop, which never touches enemies_remaining. This lesson does rename the drawing loop's j/e variables to enemy_index/enemy while it's already being edited - a small Rule 11 cleanup carried over from earlier lessons, not a functional change.",
    level1: [
      "Movement (enemy.ey=enemy.ey+0.3) and the escape check (enemy.ey>136) both live in the existing enemy-drawing loop at the bottom of TIC(), guarded by if enemy.alive then, same structural location as before.",
      "enemy.alive=false in the escape branch removes the enemy from being drawn or collision-checked on future frames, with zero reference to enemies_remaining.",
      "The bullets-vs-enemies collision loop (from Lesson 4/6) is untouched by this lesson - it remains the only place enemies_remaining changes.",
    ],
    level1Code: [
      "-- enemy movement/escape (drawing loop):",
      "for enemy_index=1,#enemies do",
      " local enemy=enemies[enemy_index]",
      " if enemy.alive then",
      "  enemy.ey=enemy.ey+0.3",
      "  if enemy.ey>136 then",
      "   enemy.alive=false",
      "  else",
      "   spr(3,enemy.ex,enemy.ey,0)",
      "  end",
      " end",
      "end",
    ],
    why: [
      "TIC-80's table-driven design already forces students to be explicit about which loop does what, since there's no automatic event system doing it for them - this lesson leans on that explicitness rather than fighting it.",
      "Keeping the two loops textually separate (rather than merging movement/escape into the collision loop 'to save a pass over the table') keeps each loop's single job legible, even though it costs a small amount of performance running two loops instead of one - worth naming that tradeoff explicitly for older/more advanced students.",
      "The rename from j/e to enemy_index/enemy is purely cosmetic (Lua doesn't care what a local is called), but it matches the meaningful-name standard already applied to every other engine's loop variables since Lesson 5 - worth a one-sentence mention in class if a sharp-eyed student compares this week's code to a screenshot of an earlier lesson.",
    ],
    talkingPoints: [
      "For advanced students: ask whether the two loops could be merged into one pass over the enemies table, and what would be gained (a few percent of performance) versus lost (harder-to-follow code, higher risk of accidentally mixing the two concerns) - this is a genuine, defensible design tradeoff, not a right/wrong question.",
      "Point out that enemy.alive is the single shared piece of state both loops respect - it's what lets two independently-written loops cooperate safely without either needing to know about the other's internals.",
    ],
    mistakes: [
      "Adding the enemies_remaining decrement into the movement/escape branch by mistake, copy-pasting from the collision loop.",
      "Forgetting the if enemy.alive then guard on the movement loop, causing already-dead (but not yet cleared) table entries to keep moving/drawing.",
      "Using the wrong screen-height constant (TIC-80's default is 136 pixels tall, not 128 or 240 - easy to misremember).",
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    teaches: "The 001 Game Creator equivalent of the GameMaker bugfix: the 'Defeated / Destroyed' Actor Trigger fires on ANY destruction of an Actor instance, so it cannot be trusted for win-counting once a second destruction pathway (the new 'Left/Entered Camera' escape-cleanup trigger) exists.",
    bugfixNote: "This is the second primary bugfix engine for this lesson, alongside GameMaker. Lesson 6 wired the enemies_remaining decrement and win Message Box into the enemy Actor Template's 'Defeated / Destroyed' trigger. That trigger fires on any destruction of the Actor, regardless of cause - including the new 'Left/Entered Camera' escape-cleanup trigger this lesson adds. Left unfixed, an enemy escaping off the bottom unshot would still decrement enemies_remaining and could trigger a false 'You Win!' message. The fix: delete the 'Defeated / Destroyed' trigger entirely, and move the decrement plus win check into the existing 'Collided with actor' trigger instead - the one trigger that actually represents a hit.",
    level1: [
      "The enemy Actor Template's Automatic Movement setting (Move = true, Direction = Down) reuses the exact mechanism already taught for the bullet's upward movement in Lesson 3 - no new concept, just a new direction and a slower speed.",
      "The new 'Left/Entered Camera' (Left case) trigger reuses the exact mechanism already taught for the bullet's off-screen cleanup in Lesson 3 - a plain Destroy, no awareness of enemies_remaining.",
      "The existing 'Collided with actor' trigger is now the single source of truth for enemies_remaining, with the decrement and win-check Events added alongside the existing Destroy Events already there from Lesson 4/6.",
      "The old 'Defeated / Destroyed' trigger is deleted entirely from the enemy Actor Template - it has no job left to do.",
    ],
    level1Code: [
      "enemy Actor Template settings:",
      "  Automatic Movement -> Move = true, Direction = Down",
      "",
      "enemy Actor Trigger: Left/Entered Camera (Left case) ->",
      "  Actor - Destroy/Revive Actor(this, Destroy)",
      "",
      "enemy Actor Trigger: Collided with actor (revised) ->",
      "  Variable - Modify Value(enemies_remaining, Subtract, 1)",
      "  If(enemies_remaining <= 0)",
      "    Message Box(\"You Win!\")",
      "  Actor - Destroy/Revive Actor(relate, Destroy)",
      "  Actor - Destroy/Revive Actor(this, Destroy)",
      "",
      "enemy Actor Trigger: Defeated / Destroyed -> deleted entirely",
    ],
    why: [
      "'Defeated / Destroyed' is 001 Game Creator's equivalent of a generic lifecycle hook - it fires whenever the Actor instance stops existing, for any reason, exactly like GameMaker's Destroy event. Both engines share this exact pitfall because both offer a convenient 'on destroyed' hook that's tempting to use for win-counting, and both need the identical fix: move the logic to the specific trigger that represents the actual game event.",
      "'Collided with actor' fires on overlap with any Solid Actor, not just the bullet specifically - but the bullet is the only other Solid Actor moving toward the enemy at this point in the semester, so in practice it's still a much stronger signal for 'a kill happened' than 'this Actor stopped existing.'",
      "Reusing the bullet's exact Lesson 3 patterns (Automatic Movement, Left/Entered Camera cleanup) for the enemy this lesson is a deliberate continuity choice - students who understood those mechanisms for the bullet don't need to learn anything new to apply them to the enemy, just recognize the same tool solving a similar problem.",
    ],
    talkingPoints: [
      "Same as GameMaker: spend extra time here, since this is a genuine latent bug from Lesson 6, not something students did wrong at the time.",
      "Ask students to predict what happens if they add the new logic to 'Collided with actor' but forget to delete 'Defeated / Destroyed' - answer: double-counting, since a shot enemy triggers both triggers (the collision AND its own consequent destruction), decrementing enemies_remaining twice per actual kill.",
      "If a student asks why the enemy needs a NEW trigger for escape cleanup instead of just reusing 'Defeated / Destroyed' for that too - explain that's exactly backwards: the problem is that a single generic trigger tries to mean two different things at once. Two specific triggers (Collided with actor for kills, Left/Entered Camera for escapes) is the fix, not a workaround.",
    ],
    mistakes: [
      "Leaving 'Defeated / Destroyed' in place AND adding the new logic to 'Collided with actor' - causes double-counting exactly like the GameMaker equivalent. Test for this explicitly: shoot exactly half the enemies and confirm the win message does NOT appear early.",
      "Setting 'Left/Entered Camera' to the 'Entered' case instead of 'Left' - fires immediately when the enemy spawns into view instead of when it exits at the bottom.",
      "Forgetting to set Direction to Down on Automatic Movement (leaves it at whatever default/previous direction was set, if the Template was copied from another Actor).",
    ],
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    teaches: "The same escape-vs-hit distinction as Godot/Unity, expressed as a plain Destroy Actor node (called from the enemy Blueprint's own Tick-driven escape Branch) vs. the OnEnemyDestroyed dispatcher (broadcast only from the successful-Cast branch of the overlap event).",
    bugfixNote: "No bugfix needed for Unreal - Lesson 6's design already broadcast OnEnemyDestroyed exclusively from the overlap event's successful-Cast branch, and this lesson's new Tick-driven escape logic calls a plain Destroy Actor with no dispatcher broadcast, same pattern as Godot/Unity.",
    level1: [
      "Event Tick drives both the movement (Add Actor World Offset along -Y, Delta Location's pin Split Struct Pinned first so -EnemySpeed*DeltaSeconds can be wired into just Y - same step Lessons 2 and 3 needed) and the escape check (a Branch comparing GetActorLocation().Y against a threshold) every frame - Blueprint's Tick behaves the same as Godot's _process or Unity's Update.",
      "The escape Branch's True case calls a plain Destroy Actor node - no connection to OnEnemyDestroyed at all.",
      "DeltaSeconds scaling on the movement offset keeps it frame-rate independent, consistent with patterns taught since Lesson 2.",
    ],
    why: [
      "Same principle as Godot/Unity: the dispatcher broadcast is reserved exclusively for the overlap event's successful-Cast branch (an actual hit), and the escape-cleanup Branch has no reason to know the dispatcher exists.",
      "This is the third engine in a row (after Godot and Unity) where the Lesson 6 structure already prevents this bug - worth using as a three-way comparison to reinforce that 'attach state changes to the specific event, not the generic one' is a cross-engine principle, not an engine-specific trick.",
    ],
    level1EngineDiagram: { file: "diagram_unreal_blueprint_l7_level1.png", w: 6.3, h: 4.536, caption: "BP_Enemy's Event Tick graph, in Unreal Blueprint - the node-for-node equivalent of the Level 1 code." },
    talkingPoints: [
      "For students on the C++ track: point out that the same distinction (Tick calling a plain Destroy() vs. OnOverlap calling Broadcast()) holds identically in code - the principle isn't Blueprint-specific.",
      "Ask advanced students: what's the Unreal-specific risk of doing heavy logic in Tick for many enemies at once (performance, since Tick runs every frame for every actor) - a good preview of optimization topics for students continuing past this course.",
    ],
    mistakes: [
      "Wiring the escape Branch's True case to Call OnEnemyDestroyed 'to keep things consistent' - reintroduces the exact bug this lesson is designed to avoid.",
      "Forgetting DeltaSeconds in the Add Actor World Offset node (frame-rate-dependent movement).",
      "Using the wrong sign on the Y offset (positive instead of negative, given this course's 'Y+ is up' convention) and having enemies move upward off the top instead of downward.",
    ],
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_07_Key_${d.file}.docx`]));

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
