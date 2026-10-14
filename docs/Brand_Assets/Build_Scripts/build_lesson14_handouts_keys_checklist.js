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
      headers: { default: brand.pageHeader("Lesson 14 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 14 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("This Week: The Foreshadowed Feature"),
        p("Since Lesson 11, several Teacher's Keys have deliberately flagged that hard-coding 5 enemies per wave and a fixed speed were 'deliberate simplifications' with 'a preview of Lesson 14's difficulty scaling.' This is that payoff lesson: enemy count and enemy speed both increase with the wave number, using the exact wave variable every engine has had since Lesson 11."),
        checkbox("Read the Handout's Level 1/2 split before class - Level 1 is enemy-count scaling only (a formula change, no new variables anywhere); Level 2 adds enemy-speed scaling (a new per-instance variable in engines that don't already have one)."),

        h2("Also This Week: Student's Choice"),
        p("Level 3 this week is different from every prior lesson: instead of 2-3 small stretch ideas, students choose ONE feature from a menu of five capstone-level options (destructible shields, a power-up, a bonus enemy, formation-marching movement, or a title/game-over screen). These are intentionally lighter-touch than a full worked Level 1 - point students to the general shape and let them problem-solve using tools they already know."),
        checkbox("Skim the Level 3 menu before class so you can point students toward the option that best matches their engine and comfort level."),

        h2("Self-Contained Lesson Assets"),
        p("No new assets this week - difficulty scaling reuses existing enemy sprites and spawn logic. Any Level 3 stretch feature a student chooses may call for its own placeholder art; that's fine to improvise from existing lesson assets or simple shapes."),

        h2("Continuity From Lesson 13"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 13, not create a new one."),
        checkbox("If a student's Lesson 13 project is lost or broken, have a pre-built \"Lesson 13 complete\" starter available per engine (from the Lesson 13 Teacher's Keys) so they aren't blocked."),
        checkbox("001 Game Creator (Windows demo laptop): confirm the two additional pre-placed enemy copies (enemy_6, enemy_7) needed for this lesson's count-scaling are added and named correctly before class - see this engine's Handout Setup section."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop's Lesson 11-13 exercises (multiple waves, sound, hit reactions) still work for all 8 engines."),
        checkbox("Test a full 3-wave playthrough on every engine before class: wave 1 should feel like Lesson 13's game, and wave 3 should be noticeably faster and busier - but still fair and beatable."),
        checkbox("Confirm restarting still correctly resets to wave 1's original count/speed, not whatever wave 3 had scaled to."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies."),

        h2("Known Gotchas This Week"),
        bullet("The most likely mistake: forgetting to reset enemies_remaining using the SAME per-wave count formula used in spawn_wave() - if these two places disagree, the win-check can fire early or never."),
        bullet("A close second: forgetting that restart must return everything to wave 1's values, not leave speed/count stuck wherever they last scaled to."),
        bullet("Difficulty scaling can make a game feel unfair if increments are too aggressive - if a cluster's wave 3 feels impossible rather than exciting, help them dial back the per-wave increment rather than pushing through."),
        bullet("Since Level 3 this week is open-ended student choice, expect a wider range of finishing points than usual - that's fine; the goal is a good-faith attempt at ONE chosen feature, not a complete implementation."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_14_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 14: Raising the Stakes", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: later waves get genuinely harder - more enemies, moving faster - using the exact wave variable you've had since Lesson 11.", { bold: true }),
    p("Open your own [YourName]Game project from Lesson 13 - don't start a new one.", { bold: true }),
    p("Tip: no new assets this week - difficulty scaling reuses your exact existing enemy sprites and spawn logic.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("Each wave spawns one more enemy than the last (5, 6, 7). Speed stays the same for now - that's Level 2.", { italics: true }),
    ...data.level1.map(s => numbered(s)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  if (data.altPath) {
    children.push(h2("Alternate Path: " + data.altPath.title));
    children.push(p(data.altPath.intro, { italics: true, color: "555555" }));
    children.push(...data.altPath.steps.map(s => numbered(s)));
    if (data.altPath.code) children.push(code(data.altPath.code));
    if (data.altPath.note) children.push(p(data.altPath.note, { italics: true, color: "555555" }));
  }
  children.push(
    h2("Level 2 - Complete (Alpha/Beta)"),
    p("Enemy speed ALSO increases each wave, on top of the growing count. Restarting correctly returns everything to wave 1's original values.", { italics: true }),
    ...data.level2.map(s => bullet(s)),
    h2("Level 3 - Refined - Choose ONE Stretch Feature"),
    p("Pick exactly one of the five ideas below. These are intentionally open-ended - a good-faith attempt at the general idea is a complete Level 3 finish, not a fully polished feature.", { italics: true, color: "555555" }),
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
      headers: { default: brand.pageHeader("Lesson 14 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

// Shared Level 3 stretch-feature menu (same five capstone options, phrased once and reused per engine with light tailoring)
function stretchMenu(engineHint) {
  return [
    "Destructible shields/barriers: a few simple blocks placed between the player and the enemies that disappear (or shrink) after absorbing a hit or two - reuses your existing collision-detection and hit-reaction patterns from Lessons 4 and 13, just on a new kind of object.",
    "A power-up: an item that occasionally drops or spawns, which the player collects to temporarily change how shooting works (faster fire rate, multiple bullets at once, or a brief shield) - reuses your existing timer/cooldown pattern from Lesson 3.",
    "A bonus enemy: a rare, faster-moving, or oddly-patterned enemy worth extra points, spawned occasionally alongside a normal wave - reuses this week's own speed/count scaling ideas, just for one special case instead of the whole wave.",
    "Formation-marching movement: instead of each enemy independently drifting down (this course's approach since Lesson 4), have the whole wave move together as a group and occasionally step sideways - a genuinely different movement pattern, inspired by classic arcade shooters.",
    "A title screen and/or game-over screen: a proper opening screen (press a key to start) and/or a dedicated game-over screen with a final score, instead of just text overlaid on the gameplay - reuses your existing game_active flag and UI tools from Lessons 9-10." + (engineHint ? " " + engineHint : ""),
  ];
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 13.",
      "No new variables needed for Level 1 - Level 2 will add one (enemy_speed).",
    ],
    level1: [
      "On the Enemy sprite's \"when I receive [spawn wave]\" script, change the fixed \"repeat 5\" loop into \"repeat (4 + wave)\" (code below) - wave 1 spawns 5, wave 2 spawns 6, wave 3 spawns 7.",
      "Wherever enemies_left gets reset to 5 for a new wave (both in the spawn script and in the wave-advance branch), change that to \"set enemies_left to (4 + wave)\" too - these two places must always agree.",
      "Click the stage and play through a full round - wave 2 should visibly spawn one more enemy than wave 1, and wave 3 one more than that.",
    ],
    level1Code: [
      "Enemy's \"when I receive [spawn wave]\" script (revised):",
      "when I receive [spawn wave]",
      "set enemies_left to (4 + wave)",
      "repeat (4 + wave)",
      "  create clone of myself",
    ],
    level2: [
      "Add a new variable enemy_speed (for all sprites). In the spawn-wave script, add \"set enemy_speed to (1 + (wave - 1) * 0.5)\".",
      "In the Enemy clone's forever loop, change \"change y by -1\" to \"change y by (enemy_speed * -1)\".",
      "Press R to restart mid-game and confirm both count and speed return correctly to wave 1's values, not wherever wave 3 left them.",
      "Play a full round and confirm wave 3 is noticeably busier and faster than wave 1, but still fair.",
    ],
    level3: stretchMenu("Scratch's \"switch backdrop\" blocks are a natural fit for a simple title/game-over screen."),
    trouble: [
      "Every wave still spawns exactly 5: confirm the loop actually uses \"(4 + wave)\" and not a leftover hardcoded 5.",
      "\"You Win!\" fires early or never: confirm enemies_left is reset using the SAME formula in every place it's set, not 5 in one spot and (4 + wave) in another.",
      "Speed doesn't change between waves: confirm enemy_speed is recalculated inside the spawn-wave script (not just once at game start) and that the Enemy's forever loop is actually reading enemy_speed, not a leftover -1.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 13.",
      "No new variables needed for Level 1 - Level 2 will use enemy.tscn's existing speed variable (from Lesson 7).",
    ],
    level1: [
      "In spawn_wave(), change the fixed \"for i in range(5)\" loop into \"for i in range(4 + wave)\" (code below) - wave 1 spawns 5, wave 2 spawns 6, wave 3 spawns 7.",
      "Wherever enemies_remaining gets reset to 5 for a new wave, change that to \"enemies_remaining = 4 + wave\" too - these two places must always agree.",
      "Run the scene and play through a full round - wave 2 should visibly spawn one more enemy than wave 1, and wave 3 one more than that.",
    ],
    level1Code: [
      "func spawn_wave():",
      "    for i in range(4 + wave):",
      "        var enemy = enemy_scene.instantiate()",
      "        add_child(enemy)",
      "        enemy.add_to_group(\"enemies\")",
      "        enemy.destroyed.connect(_on_enemy_destroyed)",
      "        enemy.escaped.connect(_on_enemy_escaped)",
      "        # ...existing per-enemy positioning...",
      "    enemies_remaining = 4 + wave",
    ],
    level2: [
      "In spawn_wave(), right after instantiating each enemy, set its speed based on the wave: enemy.speed = 60.0 + (wave - 1) * 20.0.",
      "Confirm enemy.tscn's own speed variable (from Lesson 7) is still what its _process(delta) movement actually reads - no changes needed there.",
      "Restart mid-game and confirm both count and speed return correctly to wave 1's values.",
      "Play a full round and confirm wave 3 is noticeably busier and faster than wave 1, but still fair.",
    ],
    level3: stretchMenu("Godot's Tween node is a natural fit for a smooth title/game-over screen transition."),
    trouble: [
      "Every wave still spawns exactly 5: confirm the loop actually uses \"range(4 + wave)\" and not a leftover range(5).",
      "WinText fires early or never: confirm enemies_remaining is reset using the SAME formula everywhere it's set, not 5 in one spot and 4 + wave in another.",
      "Speed doesn't change between waves: confirm enemy.speed is being set INSIDE spawn_wave() on each new instance, not left at enemy.tscn's default forever.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Open your [YourName]Game project from Lesson 13.",
      "No new fields needed for Level 1 - Level 2 will use Enemy.cs's existing speed field (from Lesson 7).",
    ],
    level1: [
      "In EnemySpawner.cs's SpawnEnemies(), change the fixed loop count of 5 into (4 + GameManager.wave) (code below) - wave 1 spawns 5, wave 2 spawns 6, wave 3 spawns 7.",
      "Wherever enemiesRemaining gets reset to 5 for a new wave, change that to 4 + wave too - these two places must always agree.",
      "Press Play and play through a full round - wave 2 should visibly spawn one more enemy than wave 1, and wave 3 one more than that.",
    ],
    level1Code: [
      "// EnemySpawner.cs:",
      "public void SpawnEnemies() {",
      "    int count = 4 + GameManager.wave;",
      "    for (int i = 0; i < count; i++) {",
      "        GameObject e = Instantiate(enemyPrefab);",
      "        // ...existing per-enemy positioning...",
      "    }",
      "    GameManager.enemiesRemaining = count;",
      "}",
    ],
    level2: [
      "In SpawnEnemies(), right after instantiating each enemy, set its speed based on the wave: e.GetComponent<Enemy>().speed = 1.5f + (GameManager.wave - 1) * 0.5f.",
      "Confirm Enemy.cs's own speed field (from Lesson 7) is still what its Update() movement actually reads - no changes needed there.",
      "Restart mid-game and confirm both count and speed return correctly to wave 1's values.",
      "Play a full round and confirm wave 3 is noticeably busier and faster than wave 1, but still fair.",
    ],
    level3: stretchMenu("Unity's UI Canvas system (already used for your HUD since Lesson 9) is a natural fit for a title/game-over screen."),
    trouble: [
      "Every wave still spawns exactly 5: confirm the loop actually uses (4 + GameManager.wave) and not a leftover literal 5.",
      "The win message fires early or never: confirm enemiesRemaining is reset using the SAME formula everywhere it's set.",
      "Speed doesn't change between waves: confirm each new enemy's speed field is being set INSIDE SpawnEnemies(), not left at the prefab's default forever.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same wave-dependent formulas work as Script Graph math nodes instead of C# expressions.",
      steps: [
        "On your spawn-loop graph, replace the fixed loop count of 5 with an Add node combining 4 and the current Wave value.",
        "On each spawned enemy, add a \"Set Member\" node targeting the enemy's own EnemyMover script and its speed field, fed by an Add node combining 1.5 and (Wave - 1) times 0.5 - Set Member targets a real, type-checked field on the component (pick it from a list), so prefer it over the generic Variables > Set Variable node, which stores under a loose string key instead.",
      ],
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: [
      "Open your [YourName]Game project from Lesson 13.",
      "Add a new instance variable to obj_enemy's Create event: enemy_speed = 1;.",
    ],
    level1: [
      "In scr_start_game's spawn_wave() function, change the fixed \"repeat (5)\" loop into \"repeat (4 + obj_spawner.wave)\" (code below) - wave 1 spawns 5, wave 2 spawns 6, wave 3 spawns 7.",
      "Wherever enemies_remaining gets reset to 5 for a new wave, change that to 4 + wave too - these two places must always agree.",
      "Press F5 (Run) and play through a full round - wave 2 should visibly spawn one more enemy than wave 1, and wave 3 one more than that.",
    ],
    level1Code: [
      "// spawn_wave() (revised):",
      "function spawn_wave() {",
      "    var count = 4 + wave;",
      "    repeat (count) {",
      "        var inst = instance_create_layer(x, y, \"Instances\", obj_enemy);",
      "        // ...existing per-enemy positioning...",
      "    }",
      "    enemies_remaining = count;",
      "}",
    ],
    level2: [
      "In spawn_wave(), right after creating each enemy instance, set its speed based on the wave: inst.enemy_speed = 1 + (wave - 1) * 0.5;.",
      "In obj_enemy's Step event, change \"y += 1;\" to \"y += enemy_speed;\" - the enemy now reads its own speed instead of a fixed literal.",
      "Restart mid-game and confirm both count and speed return correctly to wave 1's values.",
      "Play a full round and confirm wave 3 is noticeably busier and faster than wave 1, but still fair.",
    ],
    level3: stretchMenu("GameMaker's Room system and Draw GUI event (already used for your HUD) are a natural fit for a title/game-over screen."),
    trouble: [
      "Every wave still spawns exactly 5: confirm the repeat loop actually uses (4 + wave) and not a leftover literal 5.",
      "\"You Win!\" fires early or never: confirm enemies_remaining is reset using the SAME formula everywhere it's set.",
      "Speed doesn't change between waves: confirm each new instance's enemy_speed is being set INSIDE spawn_wave(), and that obj_enemy's Step event actually reads enemy_speed instead of a leftover literal 1.",
    ],
    resource: "Student Resource List > GameMaker section, or manual.gamemaker.io",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same wave-dependent formulas work as drag-and-drop math blocks instead of typed code.",
      steps: [
        "On spawn_wave()'s Repeat block, replace the fixed value 5 with an Addition block combining 4 and the wave Variable.",
        "On each created instance, add a \"Set Instance Variable\" block for enemy_speed, fed by an Addition block combining 1 and (wave - 1) times 0.5.",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 13.",
      "No new variables needed for Level 1 - Level 2 will use a new Speed variable on the Enemy object.",
    ],
    level1: [
      "On the Setup event group's \"Repeat 5 times -> Create Enemy\" action (and the matching one in the advance-wave event from Lesson 11), change the repeat count from a fixed 5 to the expression 4 + Wave (code below) - wave 1 spawns 5, wave 2 spawns 6, wave 3 spawns 7.",
      "Change EnemiesRemaining's reset value from a fixed 5 to the same 4 + Wave expression, in both events - these two places must always agree.",
      "Click Preview and play through a full round - wave 2 should visibly spawn one more enemy than wave 1, and wave 3 one more than that.",
    ],
    level1Code: [
      "Setup event group (revised):",
      "Action: Repeat (4 + Wave) times -> Create object Enemy",
      "Action: EnemiesRemaining = 4 + Wave",
      "",
      "Advance-wave event (revised, from Lesson 11):",
      "Action: Wave = Wave + 1",
      "Action: EnemiesRemaining = 4 + Wave",
      "Action: Repeat (4 + Wave) times -> Create object Enemy",
    ],
    level2: [
      "Add a new object variable, Speed, on the Enemy object (default value 30).",
      "On the same Create-Enemy actions, add: Change Speed of Enemy to (30 + (Wave - 1) * 10).",
      "Change the Enemy's existing downward-movement action to use Enemy.Variable(Speed) instead of a fixed number.",
      "Restart mid-game and confirm both count and speed return correctly to wave 1's values.",
      "Play a full round and confirm wave 3 is noticeably busier and faster than wave 1, but still fair.",
    ],
    level3: stretchMenu("GDevelop's Panel Sprite and layer-switching actions are a natural fit for a title/game-over screen."),
    trouble: [
      "Every wave still spawns exactly 5: confirm both the Setup group and the advance-wave event actually use the (4 + Wave) expression, not a leftover fixed 5.",
      "WinText fires early or never: confirm EnemiesRemaining is reset using the SAME expression in every event that sets it.",
      "Speed doesn't change between waves: confirm the Enemy's movement action is actually reading its Speed variable, not a leftover fixed number.",
    ],
    resource: "Student Resource List > GDevelop section, or wiki.gdevelop.io",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: [
      "Open your [YourName]Game project from Lesson 13.",
      "No new globals needed for Level 1 - Level 2 will add a speed field to each enemy's table entry.",
    ],
    level1: [
      "In spawn_wave(), change the fixed loop bound of 5 into 4 + wave (code below) - wave 1 spawns 5, wave 2 spawns 6, wave 3 spawns 7.",
      "Wherever enemies_remaining gets set to 5 for a new wave, change that to 4 + wave too - these two places must always agree.",
      "Press Esc, type run, press Enter, then play through a full round - wave 2 should visibly spawn one more enemy than wave 1, and wave 3 one more than that.",
    ],
    level1Code: [
      "function spawn_wave()",
      " enemies = {}",
      " local count = 4 + wave",
      " for i=1,count do",
      "  table.insert(enemies, {x=..., y=..., flash=nil})",
      " end",
      " enemies_remaining = count",
      "end",
    ],
    level2: [
      "In spawn_wave(), add a speed field to each new enemy entry: speed = 1 + (wave - 1) * 0.5.",
      "In your enemy-movement code, change the fixed downward increment to use e.speed instead of a literal number.",
      "Restart mid-game and confirm both count and speed return correctly to wave 1's values.",
      "Play a full round and confirm wave 3 is noticeably busier and faster than wave 1, but still fair.",
    ],
    level3: stretchMenu("TIC-80's map() function (for drawing a background) is a natural fit for a simple title screen."),
    trouble: [
      "Every wave still spawns exactly 5: confirm the loop bound actually uses (4 + wave) and not a leftover literal 5.",
      "\"You Win!\" fires early or never: confirm enemies_remaining is set using the SAME formula everywhere it's set.",
      "Speed doesn't change between waves: confirm each enemy's speed field is being set INSIDE spawn_wave(), and that your movement code actually reads e.speed instead of a leftover literal.",
    ],
    resource: "Student Resource List > TIC-80 section, or github.com/nesbox/TIC-80/wiki",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 13.",
      "In the Map editor, copy-paste 2 more enemy copies from your existing row, naming them enemy_6 and enemy_7 and keeping at least a full enemy-width of gap from their neighbors, same as your original row from Lesson 5 - you now have 7 total pre-placed copies to work with (this engine's enemies are always pre-placed, never runtime-spawned, so growing the count means adding more copies now, at design time).",
    ],
    level1: [
      "In your ResetGame Custom Event, add Actor – Destroy/Revive(Destroy) for enemy_6 and enemy_7 - by default, only the wave 1 count (5) should be active.",
      "In SpawnWave, add an If(wave >= 2) branch that also revives/repositions enemy_6, and an If(wave >= 3) branch that also revives/repositions enemy_7 (code below) - each wave activates one more pre-placed copy than the last.",
      "Update enemies_remaining's reset value in both ResetGame and the win-check's Else branch to match: 4 + wave, instead of a fixed 5.",
      "Save, then Play/Test and play through a full round - wave 2 should visibly show enemy_6 join in, and wave 3 should show enemy_7 join in too.",
    ],
    level1Code: [
      "Custom Event: SpawnWave (revised) ->",
      "  Actor – Destroy/Revive(Revive) + Teleport(enemy_1..enemy_5, starting positions)",
      "  If(wave >= 2)",
      "    Actor – Destroy/Revive(Revive) + Teleport(enemy_6, starting position)",
      "  If(wave >= 3)",
      "    Actor – Destroy/Revive(Revive) + Teleport(enemy_7, starting position)",
      "",
      "Custom Event: ResetGame (revised) ->",
      "  Variable – Modify Value(wave, Set, 1)",
      "  Variable – Modify Value(enemies_remaining, Set, 5)",
      "  Actor – Destroy/Revive(Destroy) for enemy_6, enemy_7",
      "  Call Custom Event(SpawnWave)",
    ],
    level2: [
      "This engine's confirmed actions don't include a verified way to change Automatic Movement speed at runtime - if your installed version exposes one (check the enemy Template's action list for a Movement-related action), scaling speed per wave is a great Level 3 stretch instead of a required Level 2 item here.",
      "For this engine, Level 2 instead means: confirm all 7 pre-placed copies revive/reposition correctly across a full 3-wave playthrough, with no copy left stuck mid-Destroy or appearing in the wrong wave.",
      "Restart mid-game and confirm enemy_6 and enemy_7 correctly return to Destroyed until their wave is reached again.",
    ],
    level3: stretchMenu("001 Game Creator's Interface/Field system (already used for your HUD since Lesson 9) is a natural fit for a title/game-over screen."),
    trouble: [
      "enemy_6 or enemy_7 never appears: confirm their SpawnWave If-branch conditions use >= (not just =), so they stay active once revived, not just on the exact wave they're introduced.",
      "\"You Win!\" fires early or never: confirm enemies_remaining is reset using the SAME (4 + wave) formula everywhere it's set.",
      "Restarting leaves enemy_6/7 visible from a previous game: confirm ResetGame explicitly Destroys them before SpawnWave runs.",
    ],
    resource: "001gamecreator.com/wiki/, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 13.",
      "No new variables needed for Level 1 - Level 2 will use BP_Enemy's existing EnemySpeed variable (from Lesson 7).",
    ],
    level1: [
      "In your SpawnWave Custom Event, change the fixed For Loop count of 5 into (4 + Wave) (code below) - wave 1 spawns 5, wave 2 spawns 6, wave 3 spawns 7.",
      "Wherever EnemiesRemaining gets reset to 5 for a new wave, change that to 4 + Wave too - these two places must always agree.",
      "Play in Editor and play through a full round - wave 2 should visibly spawn one more enemy than wave 1, and wave 3 one more than that.",
    ],
    level1Code: [
      "Custom Event SpawnWave (revised) ->",
      "  For Loop (0 to (4 + Wave) - 1) -> Spawn Actor from Class(BP_Enemy) -> Bind Event to OnEnemyDestroyed / OnEnemyEscaped",
      "  EnemiesRemaining = 4 + Wave",
    ],
    level2: [
      "On each spawned BP_Enemy, right after Spawn Actor from Class, add Set EnemySpeed = 50.0 + (Wave - 1) * 15.0.",
      "Confirm BP_Enemy's own Event Tick movement (from Lesson 7) is still what actually reads EnemySpeed - no changes needed there.",
      "Restart mid-game and confirm both count and speed return correctly to wave 1's values.",
      "Play a full round and confirm wave 3 is noticeably busier and faster than wave 1, but still fair.",
    ],
    level3: stretchMenu("Unreal's Widget Blueprint system (already used for your HUD since Lesson 9) is a natural fit for a title/game-over screen."),
    trouble: [
      "Every wave still spawns exactly 5: confirm the For Loop actually uses (4 + Wave) - 1 as its end value, not a leftover literal 4.",
      "The win message fires early or never: confirm EnemiesRemaining is reset using the SAME formula everywhere it's set.",
      "Speed doesn't change between waves: confirm Set EnemySpeed is wired right after each Spawn Actor from Class node inside SpawnWave, not left at BP_Enemy's default forever.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same wave-dependent formulas work as plain C++ expressions instead of Blueprint math nodes.",
      steps: [
        "In SpawnWave(), replace the fixed loop bound of 5 with (4 + Wave), and set each newly spawned enemy's EnemySpeed to 50.f + (Wave - 1) * 15.f right after spawning it.",
      ],
      note: "This mirrors the Blueprint path's For Loop and Set EnemySpeed nodes exactly - same formulas, just in code.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_14_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 14 Teacher's Key: Raising the Stakes", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine + " - Instructor Reference Only", size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    h2("What This Lesson Teaches"),
    p(data.teaches),
    h2("Level 1 - Full Worked Solution"),
    ...data.level1.map(s => bullet(s)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  children.push(
    h2("Why This Structure"),
    ...data.why.map(s => bullet(s)),
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
      headers: { default: brand.pageHeader("Lesson 14 Key - " + data.engine + " (Instructor Only)") },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    teaches: "Replacing two hardcoded literals (a fixed spawn count, a fixed movement speed) with formulas keyed off the wave variable - the payoff lesson for simplifications flagged as deliberate back in Lesson 11.",
    level1: [
      "\"repeat (4 + wave)\" and \"set enemies_left to (4 + wave)\" must use the identical expression - any mismatch between them is exactly the kind of bug this Key's troubleshooting section calls out.",
      "No new variables are needed for Level 1 - this is purely a formula change to numbers that were already there.",
      "Wave 1 still spawns exactly 5 (4 + 1 = 5), so nothing about the game's opening feel changes - only later waves escalate.",
    ],
    why: [
      "Deriving both the spawn count and the enemies_left reset from the SAME formula (rather than one being \"4 + wave\" and the other a separately-updated literal) is the exact same discipline as Lesson 11's spawn_wave() reuse - one source of truth, referenced everywhere it's needed.",
      "Starting the formula at \"4 + wave\" (rather than, say, \"5 * wave\") keeps wave 1 unchanged and growth gentle - worth discussing as a game-balance decision, not just a math one.",
    ],
    talkingPoints: [
      "Ask students to compute the wave 1/2/3 enemy counts by hand from the formula before running it, then confirm the game matches.",
      "For strong students: discuss what would happen with a much steeper formula (e.g. \"5 * wave\") - wave 3 would have 15 enemies, likely overwhelming for a 3-wave game. A good real discussion about tuning difficulty, not just implementing it.",
    ],
    mistakes: [
      "Updating the spawn loop's count but forgetting to update enemies_left's reset value (or vice versa), causing the win-check to fire early or never.",
      "Using a different formula in two different places by accident (e.g. \"4 + wave\" in one, \"5 + wave\" in another).",
      "For Level 2: forgetting to recalculate enemy_speed inside the spawn-wave script itself, so it only ever gets wave 1's value.",
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    teaches: "Same idea as every engine this week - replace hardcoded literals with wave-dependent formulas - expressed as GDScript expressions inside the spawn_wave() function this course has been building since Lesson 10.",
    level1: [
      "range(4 + wave) and enemies_remaining = 4 + wave must use the identical expression - any mismatch is exactly the kind of bug this Key's troubleshooting section calls out.",
      "No new variables are needed for Level 1 - this is purely a formula change inside a function that already exists.",
      "Wave 1 still spawns exactly 5 (4 + 1 = 5), so nothing about the game's opening feel changes - only later waves escalate.",
    ],
    why: [
      "Keeping both the loop bound and enemies_remaining's value inside the SAME function (spawn_wave()) rather than split across files makes the 'one source of truth' discipline easy to see and easy to verify at a glance.",
      "For Level 2, setting enemy.speed on each freshly-instantiated enemy (rather than changing enemy.tscn's own default) keeps the scaling logic entirely inside spawn_wave() - the enemy scene itself stays exactly as simple as Lesson 7 left it.",
    ],
    talkingPoints: [
      "Ask students to trace through spawn_wave() for wave 3 by hand: range(4+3)=range(7), enemies_remaining=7 - a good comprehension check for reading a formula inside real code.",
      "For students comparing engines: point out that Unreal's identical (4 + Wave) formula, expressed as Blueprint math nodes instead of GDScript, is conceptually the exact same idea - a good cross-engine parallel.",
    ],
    mistakes: [
      "Updating the loop's range() but forgetting to update enemies_remaining (or vice versa).",
      "Setting enemy.speed BEFORE add_child() instead of after - harmless in Godot 4 either way, but worth confirming the assignment target is actually the enemy just created, not a stale reference.",
      "Leaving enemy.tscn's default speed := 60.0 as the ONLY place speed is set, forgetting the new per-instance override in spawn_wave() entirely.",
    ],
  },
  {
    engine: "Unity", file: "Unity",
    teaches: "Same idea as every engine this week - replace hardcoded literals with wave-dependent formulas - expressed as C# expressions inside EnemySpawner.cs's SpawnEnemies() method.",
    level1: [
      "int count = 4 + GameManager.wave computes the count once, then reuses it for both the loop and enemiesRemaining - a single source of truth by construction, not by discipline alone.",
      "No new fields are needed for Level 1 - this is purely a formula change inside a method that already exists.",
      "Wave 1 still spawns exactly 5 (4 + 1 = 5), so nothing about the game's opening feel changes - only later waves escalate.",
    ],
    why: [
      "Storing count in a local variable (rather than repeating the (4 + GameManager.wave) expression twice) is good practice worth naming explicitly - it guarantees the loop and the reset value can never disagree, since they read the same computed value.",
      "For Level 2, setting each enemy's speed field right after Instantiate() (rather than changing the prefab's own default) keeps the scaling logic entirely inside SpawnEnemies() - Enemy.cs itself stays exactly as simple as Lesson 7 left it.",
    ],
    talkingPoints: [
      "Ask students why storing count in a local variable is safer than writing the formula twice - a good, concrete illustration of avoiding duplicated logic.",
      "For students comparing engines: point out that GameMaker's identical (4 + wave) formula, expressed in GML instead of C#, is conceptually the exact same idea - a good cross-engine parallel.",
    ],
    mistakes: [
      "Writing the formula twice (once for the loop, once for enemiesRemaining) instead of storing it in a local variable - works if both are correct, but an easy place for a typo to cause disagreement.",
      "Forgetting GetComponent<Enemy>() when trying to set the new enemy's speed field, attempting to set it on the wrong object type.",
      "Leaving Enemy.cs's default public float speed = 1.5f as the ONLY place speed is set, forgetting the new per-instance override in SpawnEnemies() entirely.",
    ],
  },
  {
    engine: "GameMaker", file: "GameMaker",
    teaches: "Same idea as every engine this week - replace hardcoded literals with wave-dependent formulas - expressed as GML inside scr_start_game's spawn_wave() function, introduced in Lesson 11.",
    level1: [
      "var count = 4 + wave; computes the count once, then reuses it for both the repeat loop and enemies_remaining - a single source of truth by construction.",
      "No new instance variables are needed for count scaling - only Level 2's speed scaling introduces one (enemy_speed on obj_enemy).",
      "Wave 1 still spawns exactly 5 (4 + 1 = 5), so nothing about the game's opening feel changes - only later waves escalate.",
    ],
    why: [
      "Storing count in a local var (rather than repeating the formula) guarantees the repeat loop and the reset value can never disagree - the same discipline as every other engine this week, in GML's own idiom.",
      "Setting enemy_speed on each freshly-created instance (inst.enemy_speed = ...) rather than changing obj_enemy's Create-event default keeps the scaling logic entirely inside spawn_wave() - obj_enemy's Step event stays exactly as simple as Lesson 7 left it, just reading a variable instead of a literal.",
    ],
    talkingPoints: [
      "Ask students to trace spawn_wave() for wave 3 by hand: count = 4 + 3 = 7 - a good comprehension check.",
      "For students comparing engines: point out that TIC-80's identical (4 + wave) formula, expressed as plain Lua instead of GML, is conceptually the exact same idea - a good cross-engine parallel, and both engines this week needed to introduce a NEW per-instance/per-entry speed value where none existed before.",
    ],
    mistakes: [
      "Updating the repeat loop's count but forgetting to update enemies_remaining (or vice versa) - avoided entirely if count is stored in a local var and reused, which is why that's the recommended pattern.",
      "Setting enemy_speed on the wrong instance (e.g. self instead of inst, or vice versa) inside spawn_wave().",
      "Forgetting to change obj_enemy's Step event from \"y += 1;\" to \"y += enemy_speed;\" - the variable exists but is never actually read.",
    ],
  },
  {
    engine: "GDevelop", file: "GDevelop",
    teaches: "Same idea as every engine this week - replace hardcoded literals with wave-dependent formulas - expressed as GDevelop expressions inside the Setup and advance-wave events built since Lesson 10/11.",
    level1: [
      "(4 + Wave) as a repeat count and as EnemiesRemaining's reset value must be identical in BOTH the Setup event group and the advance-wave event - two places, one formula, each independently editable, so both need checking.",
      "No new variables are needed for Level 1 - this is purely an expression change inside events that already exist.",
      "Wave 1 still spawns exactly 5 (4 + 1 = 5), so nothing about the game's opening feel changes - only later waves escalate.",
    ],
    why: [
      "GDevelop's expression fields accept the SAME formula ((4 + Wave)) typed directly wherever a number is expected - no separate 'store it in a variable first' step is required the way GameMaker's or Unity's code needed, though it's still good practice to double check both events agree.",
      "For Level 2, adding a Speed object variable to Enemy (rather than a scene-wide variable) keeps each enemy's speed as its own property, consistent with how Score/Lives/Wave are scene-wide but per-enemy state (like position) has always lived on the object itself.",
    ],
    talkingPoints: [
      "Ask students to find both places (4 + Wave) needs to appear (Setup group, advance-wave event) and confirm they're both actually updated - a good comprehension check for events versus a single function.",
      "For students comparing engines: point out that Scratch's identical (4 + wave) formula, expressed as blocks instead of GDevelop's expression fields, is conceptually the exact same idea.",
    ],
    mistakes: [
      "Updating the Setup group's formula but forgetting the advance-wave event's matching formula (or vice versa) - these are two separate events since Lesson 11, easy to update only one.",
      "Adding Speed as a scene variable instead of an object variable on Enemy, which would give every enemy the SAME speed rather than each one carrying its own.",
      "Forgetting to change the Enemy's movement action to read Enemy.Variable(Speed), leaving it at a hardcoded number.",
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    teaches: "Same idea as every engine this week - replace hardcoded literals with wave-dependent formulas - expressed as plain Lua inside spawn_wave(), introduced in Lesson 11.",
    level1: [
      "local count = 4 + wave computes the count once, then reuses it for both the spawn loop and enemies_remaining - a single source of truth by construction.",
      "No new fields are needed for Level 1 count scaling - only Level 2's speed scaling adds a new field (speed) to each enemy's table entry.",
      "Wave 1 still spawns exactly 5 (4 + 1 = 5), so nothing about the game's opening feel changes - only later waves escalate.",
    ],
    why: [
      "Storing count in a local variable (rather than repeating the formula) guarantees the spawn loop and enemies_remaining can never disagree - the same discipline as every other engine this week, in plain Lua.",
      "Adding a speed field to each enemy's own table entry (rather than a single global speed) keeps each enemy's full state - position, alive-ness, flash status from Lesson 13, and now speed - together in one place, the same discipline established since Lesson 4.",
    ],
    talkingPoints: [
      "Ask students to trace spawn_wave() for wave 3 by hand: count = 4 + 3 = 7 - a good comprehension check.",
      "For students comparing engines: point out that GameMaker's identical (4 + wave) formula, expressed as GML instead of Lua, is conceptually the exact same idea - both engines needed a new per-instance/per-entry speed value where none existed before.",
    ],
    mistakes: [
      "Updating the spawn loop's count but forgetting to update enemies_remaining (or vice versa) - avoided entirely if count is stored in a local variable and reused.",
      "Setting speed on the table entry AFTER table.insert() instead of as part of the inserted table itself, which can silently fail depending on how the insert is written.",
      "Forgetting to change the movement code to read e.speed instead of a leftover literal - the field exists but is never actually used.",
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    teaches: "This engine's version of difficulty scaling looks different from the other seven: since enemies are always pre-placed at design time (never runtime-spawned, per Lesson 5's precedent), 'more enemies per wave' means activating MORE pre-placed copies, not creating new ones - an honest, confirmed-mechanism-only approach given this engine's unverified runtime speed API.",
    level1: [
      "enemy_6 and enemy_7 are added at design time, in the Map editor, as two more copies alongside the original enemy_1..enemy_5, spaced with the same full-enemy-width gap as the rest of the row - a direct continuation of Lesson 5's design-time-placement precedent.",
      "SpawnWave's nested If(wave >= 2) / If(wave >= 3) branches are what actually implements the scaling - each wave activates one more already-existing copy than the last, using only the confirmed Revive/Teleport mechanism from Lesson 11.",
      "ResetGame Destroying enemy_6/7 by default (before SpawnWave runs) is what keeps wave 1 looking exactly like Lesson 13's game - only 5 active copies until later waves need more.",
    ],
    why: [
      "Rather than fabricate an unconfirmed 'set movement speed at runtime' action, this lesson leans entirely on mechanisms already verified in this course (Spawn/Destroy/Revive, design-time Actor placement) - the same honest-uncertainty approach used for this engine's restart flow in Lesson 10 and hit-flash timing in Lesson 13.",
      "This actually turns this engine's fixed-actor-count constraint into a genuine, different-but-valid solution to the same problem every other engine solves at runtime - worth presenting to students as a real design trade-off, not a lesser version of the 'real' approach.",
    ],
    talkingPoints: [
      "Ask students on this engine to explain, in their own words, why 001 Game Creator's approach to more-enemies-per-wave is structurally different from every other engine this week, and why that's a reasonable consequence of a design decision made all the way back in Lesson 5.",
      "For advanced students: if their installed version DOES expose a runtime movement-speed action (check together), implementing per-wave speed scaling as a Level 3 stretch is a great extension, using the exact same wave-gated If-branch pattern as the count scaling.",
    ],
    mistakes: [
      "Forgetting to Destroy enemy_6/7 in ResetGame, leaving them visible on a fresh wave 1 game.",
      "Using If(wave = 2) instead of If(wave >= 2) for enemy_6, which would deactivate it again once wave 3 begins instead of keeping it active.",
      "Naming the new copies inconsistently (not exactly enemy_6/enemy_7), breaking SpawnWave's references to them.",
    ],
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    teaches: "Same idea as every engine this week - replace hardcoded literals with wave-dependent formulas - expressed as Blueprint math nodes inside the SpawnWave Custom Event introduced in Lesson 11.",
    level1: [
      "(4 + Wave) as the For Loop's end value and as EnemiesRemaining's reset value must be identical - both live inside the same SpawnWave graph, making it easy to verify they agree.",
      "No new variables are needed for Level 1 - this is purely a formula change inside a Custom Event that already exists.",
      "Wave 1 still spawns exactly 5 (4 + 1 = 5), so nothing about the game's opening feel changes - only later waves escalate.",
    ],
    why: [
      "Computing (4 + Wave) once via a shared math node chain (rather than two separate calculations) keeps the For Loop and EnemiesRemaining tied to the exact same value - the same discipline as every other engine this week, expressed as Blueprint wiring instead of a stored variable.",
      "Setting EnemySpeed on each spawned BP_Enemy right after Spawn Actor from Class (rather than changing BP_Enemy's own default) keeps the scaling logic entirely inside SpawnWave - BP_Enemy's Event Tick movement stays exactly as simple as Lesson 7 left it.",
    ],
    talkingPoints: [
      "Ask students to trace SpawnWave for wave 3 by hand: (4 + 3) = 7 - a good comprehension check for reading Blueprint math wiring.",
      "For students on the C++ track: point out that the identical (4 + Wave) formula, expressed as a plain arithmetic expression instead of Blueprint nodes, is conceptually the exact same idea - worth naming that parallel explicitly.",
    ],
    mistakes: [
      "Using (4 + Wave) as the For Loop's end value without the \"- 1\" adjustment For Loop nodes typically need (since For Loop's end value is inclusive) - resulting in one extra enemy per wave. Double-check against your specific For Loop node's exact semantics.",
      "Updating the For Loop's count but forgetting to update EnemiesRemaining (or vice versa).",
      "Setting EnemySpeed on BP_Enemy's class defaults instead of the just-spawned instance, which would affect ALL enemies at once rather than scaling per wave.",
    ],
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_14_Key_${d.file}.docx`]));

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
