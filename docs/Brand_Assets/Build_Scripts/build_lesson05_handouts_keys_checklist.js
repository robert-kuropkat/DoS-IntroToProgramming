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
      headers: { default: brand.pageHeader("Lesson 5 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 5 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("Self-Contained Lesson Assets"),
        p("No new assets this week - the first lesson where that's true. Everything needed is already sitting in students' own projects from Lessons 3 and 4: bullet_yellow.png, sfx_shoot.wav, enemy_invader_red.png, and sfx_hit.wav."),
        checkbox("No new Assets folder this week - confirm students still have their Lesson 3 and 4 Assets folders available if they need to re-import anything."),

        h2("Continuity From Lesson 4"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 4, not create a new one - today turns last week's single enemy into a full row of them."),
        checkbox("If a student's Lesson 4 project is lost or broken, have a pre-built \"Lesson 4 complete\" starter available per engine (from the Lesson 4 Teacher's Keys) so they aren't blocked."),

        h2("Engine Verification"),
        p("This week's core new concept - using a loop to create several instances instead of placing or spawning just one - has a genuinely different flavor per engine. Worth a quick personal test run before class."),
        checkbox("Confirm the demo laptop's Lesson 3 and 4 exercises still work for all 8 engines."),
        checkbox("Pre-build the full Scratch multi-clone exercise on top of a finished Lesson 4 project - this is the trickiest engine this week, since hit-detection has to move from the Bullet's script to the Enemy's own clone script (see the Key for why)."),
        checkbox("Godot / Unity / GameMaker / Unreal: confirm a simple for-loop spawning several instances at different positions works cleanly in your installed version - standard, low-risk functionality, but worth a 60-second sanity check."),
        checkbox("GDevelop: confirm the \"Repeat\" event type (right-click an event sheet > Add Repeat Event) is available in your installed version."),
        checkbox("TIC-80: no new API here, just a for loop populating a Lua table at startup - low risk, but worth a quick test since the bullets-vs-enemies collision check now needs a nested loop instead of one comparison."),
        checkbox("001 Game Creator (Windows demo laptop): confirm dragging several copies of the enemy Actor Template onto the Map each keep their own working \"Collided with actor\" trigger with no extra scripting - standard Template behavior, but worth confirming once in your installed version."),
        checkbox("If any advanced-track students are on Unreal this week, confirm the Blueprint For Loop node is available and behaves as expected (Blueprints > Utilities > Flow Control)."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies."),

        h2("Known Gotchas This Week"),
        bullet("Scratch: this is the one engine where the hit-detection logic actually has to move this week - Scratch clones have no unique ID, so a Bullet clone can't tell which specific Enemy clone it's touching. The fix (checking from each Enemy clone's own script instead) is the correct, standard Scratch technique - flag this clearly, since it's a real change from last week's code, not just an extension of it."),
        bullet("Scratch (continued): because a hit Bullet clone is no longer explicitly deleted on contact, it keeps flying through a hidden/destroyed Enemy and cleans itself up normally off-screen - this is expected, not a bug, and worth explaining if a student asks why the bullet doesn't visibly vanish exactly on contact."),
        bullet("Godot / Unity / GameMaker / GDevelop / Unreal: the most common slip is leaving the old, manually-placed single enemy in the scene/room/level alongside the newly spawned row - remind students to delete it once the spawn loop is working."),
        bullet("001 Game Creator: no scripting changes at all this week - just more copies of the same Actor Template placed on the Map. If a student expects to write a loop and finds nothing to write, that's correct, not a sign they're missing something."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_05_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 5: Enemy Ranks", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: one enemy becomes many. Your job is to turn last week's single enemy into a whole row of them, using a loop instead of repeating yourself by hand."),
    p("Open your own [YourName]Game project from Lesson 4 - don't start a new one. Everything this week builds on the enemy and collision detection you already have.", { bold: true }),
    p("Tip: no new assets this week - you're reusing bullet_yellow.png, sfx_shoot.wav, enemy_invader_red.png, and sfx_hit.wav from your Lesson 3 and 4 Assets folders.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("Several enemies sit on screen in a row, created by a loop instead of placed or spawned one at a time. Shooting any one of them destroys just that one. This is a complete finish for today - reaching this is enough to move on to next week.", { italics: true }),
    ...diagram("diagram_spawn_loop.png", 6.2, 4.712, "The spawn loop, in general - the same counter, repeat check, creation step, and increment behind Level 1 in every engine, whatever the exact blocks or syntax look like."),
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
    if (data.altPath.level3) {
      children.push(h3("Level 3 Visual-Track Equivalent"));
      children.push(...data.altPath.level3.map(s => bullet(s)));
      if (data.altPath.level3Code) children.push(code(data.altPath.level3Code));
      if (data.altPath.level3Diagram) children.push(...diagram(data.altPath.level3Diagram.file, data.altPath.level3Diagram.w, data.altPath.level3Diagram.h, data.altPath.level3Diagram.caption));
    }
    if (data.altPath.note) children.push(p(data.altPath.note, { italics: true, color: "555555" }));
  }
  children.push(
    h2("Level 2 - Complete (Alpha/Beta)"),
    p("Every enemy in the row can be individually shot and destroyed without affecting the others, and a bullet that misses every enemy still behaves exactly like it did last week.", { italics: true }),
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
      headers: { default: brand.pageHeader("Lesson 5 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 4.",
      "Switch to the Enemy sprite. Add a \"for this sprite only\" variable named enemy_x (right-click Variables > Make a Variable, choose \"For this sprite only\").",
      "Add a \"when green flag clicked\" script to the Enemy sprite with just \"hide\" - like Bullet, the original Enemy sprite should never be seen itself, only its clones.",
    ],
    level1: [
      "Still on the Enemy sprite, build a spawn script: \"when green flag clicked\" (a second one, alongside the hide script) -> repeat 5 -> inside the loop, set enemy_x to a different value each time (e.g. using a counter variable, or (loop count * 100) - 300) -> create clone of myself.",
      "Add a \"when I start as a clone\" script to the Enemy: show, go to x: (enemy_x) y: 150.",
      "In that same clone script, add a forever loop containing \"if <touching [Bullet]?> then delete this clone\" - this replaces last week's Bullet-side check (see the note below for why).",
      "IMPORTANT: remove the touching-Enemy check from the Bullet sprite's own forever loop if you still have it from Lesson 4 - Bullet's script goes back to exactly what it was at the end of Lesson 3 (move up, clean up off-screen), nothing about enemies.",
      "Click the stage, then press spacebar aimed at different enemies to test - each one should disappear independently when hit.",
    ],
    level2: [
      "Shoot several different enemies in a row, including some misses - each hit should remove exactly the one enemy clone that was actually touched, and every other enemy should be unaffected.",
      "Note: Scratch clones don't have unique ID numbers the way objects do in other engines, so a Bullet clone can't identify exactly which Enemy clone it's touching from the Bullet's side. That's why the touching-check lives on the Enemy's own clone script instead - each Enemy clone only ever asks about itself, which works correctly no matter how many enemies or bullets are on screen.",
      "A hit bullet isn't explicitly deleted the instant it connects - it keeps moving upward and cleans itself up off-screen exactly like a miss would. This is a deliberate simplification, not a bug.",
    ],
    level3: [
      "Sound: this already works with no changes - if you added \"start sound [sfx_hit]\" to the Enemy's clone script back in Lesson 4, every enemy clone plays it independently when hit.",
      "Visual: give enemies a striped look by alternating a costume or color effect based on their position in the loop (e.g., every other one gets \"set color effect to 50\"). Scratch's \"color effect\" isn't a literal red/green/blue color - it's a hue dial from 0 to 200 that cycles all the way around the color wheel. A value of 50 rotates about a quarter of the way around it, so on a red-ish enemy sprite the shifted enemies land somewhere around yellow-green, not on a specific named color - the goal is \"visibly different,\" not a precise shade.",
      "Structure: keep the spawn loop's only job being creating clones and giving each one its starting position - the hit-detection and hide-on-hit logic stays entirely inside the \"when I start as a clone\" script, not the spawn script.",
    ],
    trouble: [
      "All enemies stack on top of each other: confirm enemy_x is actually being set to a new value before each \"create clone of myself\" - if it's set once outside the loop, every clone reads the same value.",
      "Shooting one enemy removes all of them, or none: confirm the touching check moved to the Enemy's own \"when I start as a clone\" script, not left on the Bullet - a leftover Bullet-side check will misbehave with multiple enemies.",
      "Project freezes: this almost always means a clone is creating another clone inside its own \"when I start as a clone\" script - the spawn loop belongs in its own separate green-flag script.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 4.",
      "Delete the single Enemy instance you manually placed in your main scene last week - it's about to be replaced by several spawned copies.",
      "If your main scene's root node doesn't already have a script attached, attach one now (extends Node2D, or whatever your root's type is).",
    ],
    level1: [
      "In the main scene's script, preload enemy.tscn near the top, then in _ready(), use a for loop to instantiate several copies, each positioned at a different x (code below).",
      "Run the scene - you should see a row of enemies, and shooting any one of them should destroy just that one.",
    ],
    level1Code: [
      "# Main scene's script:",
      "var enemy_scene := preload(\"res://enemy.tscn\")",
      "",
      "func _ready():",
      "    for enemy_count in range(5):",
      "        var new_enemy := enemy_scene.instantiate()",
      "        new_enemy.position = Vector2(150 + enemy_count * 200, 50)",
      "        add_child(new_enemy)",
    ],
    level2: [
      "Shoot several different enemies in a row, including some misses - each hit destroys exactly the one enemy that was actually touched, and every other enemy is unaffected.",
      "This already works with no extra code: enemy.tscn's own script (its area_entered connection and hit-response) is self-contained, so every instance the loop creates behaves independently and correctly, exactly the way Lesson 4 built it.",
    ],
    level3: [
      "Sound: this already works with no changes - each spawned enemy carries its own Hit AudioStreamPlayer2D (if you built the Lesson 4 sound idea), so every instance plays it independently.",
      "Visual: vary each enemy's modulate color based on its loop index (enemy_count) at spawn time, for a striped row effect - alternate Color(0.8, 0.1, 0.2), a deep crimson red, with Color(0.95, 0.4, 0.5), a lighter dusty rose/pink (Godot's Color values run 0-1 instead of 0-255, so 0.8 is roughly 200 out of 255). The result is a row where every other enemy is a visibly different shade of red, not a smooth gradient blend.",
      "Structure: keep the spawn loop's only job being creating and positioning enemies - resist the urge to also handle collision-response logic in the main scene's script; that still belongs entirely inside enemy.tscn's own script.",
    ],
    trouble: [
      "Only one enemy appears, or they're all stacked in the same spot: confirm the loop variable enemy_count is actually used in the position calculation, not a fixed number.",
      "Nothing appears at all: confirm add_child(new_enemy) is actually being called for each instance - instantiate() alone creates the node in memory but doesn't add it to the scene tree.",
      "Old and new enemies both appear: confirm the manually-placed Lesson 4 Enemy instance was actually deleted from the scene, not just hidden.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Open your [YourName]Game project from Lesson 4.",
      "Delete the single Enemy prefab instance you manually placed in your scene last week.",
      "Create a new empty GameObject named EnemySpawner. Create a new C# script named EnemySpawner, attach it to this new GameObject.",
    ],
    level1: [
      "In EnemySpawner.cs, add a public GameObject enemyPrefab field and an int enemyCount, then use a for loop in Start() to Instantiate() several copies at different positions (code below).",
      "In the Inspector, drag your Enemy prefab onto EnemySpawner's Enemy Prefab field.",
      "Press Play - you should see a row of enemies, and shooting any one of them should destroy just that one.",
    ],
    level1Code: [
      "// EnemySpawner.cs:",
      "public GameObject enemyPrefab;",
      "public int enemyCount = 5;",
      "",
      "void Start() {",
      "    for (int enemyIndex = 0; enemyIndex < enemyCount; enemyIndex++) {",
      "        Vector3 pos = new Vector3(-4f + enemyIndex * 2f, 4f, 0f);",
      "        Instantiate(enemyPrefab, pos, Quaternion.identity);",
      "    }",
      "}",
    ],
    level2: [
      "Shoot several different enemies in a row, including some misses - each hit destroys exactly the one enemy that was actually touched, and every other enemy is unaffected.",
      "This already works with no extra code: Bullet.cs's OnTriggerEnter2D and Enemy.cs's own Start() are self-contained per instance, so every prefab copy the loop creates behaves independently and correctly, exactly the way Lesson 4 built it.",
    ],
    level3: [
      "Sound: this already works with no changes - AudioSource.PlayClipAtPoint() in Bullet.cs runs independently for every enemy that gets hit.",
      "Visual: vary each enemy's SpriteRenderer.color based on the loop index enemyIndex at spawn time, for a striped row effect - this needs a small public setter method or a spawner-side GetComponent<SpriteRenderer>() call right after Instantiate(). Alternate new Color(0.8f, 0.1f, 0.2f), a deep crimson red, with new Color(0.95f, 0.4f, 0.5f), a lighter dusty rose/pink (Unity's Color values run 0-1 instead of 0-255, so 0.8f is roughly 200 out of 255) - the result is every other enemy showing up as a visibly different shade of red, not a smoothly blended color.",
      "Structure: keep EnemySpawner.cs's only job being creating and positioning enemies - resist the urge to also handle collision-response logic there; that still belongs entirely inside Enemy.cs and Bullet.cs.",
    ],
    trouble: [
      "Only one enemy appears, or they're all stacked in the same spot: confirm the loop variable enemyIndex is actually used in the position calculation, not a fixed Vector3.",
      "Console error about a missing prefab reference: confirm the Enemy Prefab field on EnemySpawner is actually assigned in the Inspector.",
      "Old and new enemies both appear: confirm the manually-placed Lesson 4 Enemy instance was actually deleted from the Hierarchy.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same for-loop spawn pattern works in a Script Graph instead of C#.",
      steps: [
        "On the EnemySpawner GameObject, add a Script Machine component and open a new Script Graph.",
        "Add an On Start event node feeding a \"For Loop\" node (0 to enemyCount), with an \"Instantiate\" node inside it (Enemy prefab as the asset input) feeding a \"Set Position\" node using the loop's current index.",
        "Set up the Enemy Prefab field the same way as the C# path.",
      ],
      level1Diagram: { file: "diagram_unity_visual_scripting_level1.png", w: 6.2, h: 4.96, caption: "The Script Graph, node by node - Instantiate's own output pin plugs straight into Set Position's Target, no extra step needed." },
      note: "Level 2/3 already work the same way as the C# path, since the per-instance logic lives on the Enemy and Bullet prefabs' own graphs, not the spawner's.",
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: [
      "Open your [YourName]Game project from Lesson 4.",
      "Remove the single obj_enemy instance you manually placed in your room last week: open the Room Editor, click that instance in the room, and press Delete (or right-click it > Delete). Do NOT delete the obj_enemy object itself from the Asset Browser - obj_spawner still needs that object to exist so it can create new instances of it.",
      "Create a new object obj_spawner (no sprite needed), and drag one instance of it into your room.",
    ],
    level1: [
      "Add a Create event to obj_spawner with a for loop that calls instance_create_layer() several times at different positions (code below).",
      "Press F5 (Run) - you should see a row of enemies, and shooting any one of them should destroy just that one.",
    ],
    level1Code: [
      "// obj_spawner's Create event:",
      "for (var enemy_count = 0; enemy_count < 5; enemy_count++) {",
      "    instance_create_layer(80 + enemy_count * 96, 80, \"Instances\", obj_enemy);",
      "}",
    ],
    level2: [
      "Shoot several different enemies in a row, including some misses - each hit destroys exactly the one enemy that was actually touched, and every other enemy is unaffected.",
      "This already works with no extra code: obj_bullet's Collision event with obj_enemy already reacts per-pair of overlapping instances, so every instance the loop creates behaves independently and correctly, exactly the way Lesson 4 built it.",
    ],
    level3: [
      "Sound: this already works with no changes - audio_play_sound() in obj_bullet's Collision event runs independently for every enemy that gets hit.",
      "Visual: vary each obj_enemy's image_blend based on enemy_count in obj_spawner's loop, using with (instance_create_layer(...)) { image_blend = ... } to set a property on the instance right after creating it, for a striped row effect. make_color_rgb(200, 25, 55) is a deep crimson red and make_color_rgb(230, 100, 130) is a lighter dusty rose/pink - image_blend multiplies your sprite's pixels by this color, so alternating the two tints every other enemy a visibly different shade of red rather than blending into one new color.",
      "Structure: keep obj_spawner's Create event limited to just creating and positioning enemies - resist the urge to also handle collision-response logic there; that still belongs entirely inside obj_enemy and obj_bullet's own events.",
    ],
    trouble: [
      "Only one enemy appears, or they're all stacked in the same spot: confirm the loop variable enemy_count is actually used in the position calculation, not a fixed number.",
      "\"instance_create_layer\" errors about a missing layer: confirm the room's instance layer is still actually named \"Instances\", matching Lesson 3's setup.",
      "Old and new enemies both appear: confirm the manually-placed Lesson 4 obj_enemy instance was actually deleted from the room.",
      "\"obj_enemy doesn't exist\" or obj_spawner's Create event won't compile: this means the obj_enemy object asset itself got deleted from the Asset Browser, not just its placed instance - Object (the reusable template in the Asset Browser) and Instance (one placed copy of it in the room) are different things in GameMaker; you can freely delete an instance from a room without touching the object it was made from, but instance_create_layer() needs the object to still exist. Recreate obj_enemy from Lesson 4's Key if this happens.",
    ],
    resource: "Student Resource List > GameMaker section, or gamemaker.io/en/tutorials",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same for-loop spawn pattern works with drag-and-drop blocks instead of typed code.",
      steps: [
        "On obj_spawner's Create event, switch to the GML Visual tab (the toggle at the top-left of the event, next to the code-view icon).",
        "From the Variables category, drag a \"Set Variable\" block into the event. Set its Variable field to enemy_count and its Value field to 0 - this is the loop counter, the block equivalent of var enemy_count = 0 in the code path.",
        "From the Control category, drag a \"Repeat\" block below it. Set its Times field to 5.",
        "Inside the Repeat block's body, drag an \"Instance - Create Instance\" block (from the Instances category). Set its Object field to obj_enemy, its Layer field to \"Instances\" (matching the layer name from Lesson 3's setup), its X field to the expression 80 + (enemy_count * 96), and its Y field to 80.",
        "Still inside the Repeat block, right after the Create Instance block, drag a second \"Set Variable\" block: Variable field = enemy_count, Value field = enemy_count + 1 - this increments the counter each pass, the block equivalent of enemy_count++ in the code for loop.",
        "Press F5 (Run) and test the same way as the code path: you should see 5 enemies in a row, and shooting any one destroys just that one.",
      ],
      level1Diagram: { file: "diagram_gml_visual_level1.png", w: 6.2, h: 4.712, caption: "The five blocks above, wired together - the block-for-block equivalent of the Level 1 code." },
      level3: [
        "Sound and Structure need no changes from the code path either way - both already live on obj_enemy and obj_bullet's own events, which behave identically no matter which path spawned the instance.",
        "Visual (color-switching), on the other hand, DOES need its own block sequence - it happens inside the spawn loop itself, so it's built differently in Visual than in code. Give the \"Instance - Create Instance\" block a Target field (in its own block, alongside Object/Layer/X/Y) and set it to a new local variable, e.g. new_enemy - this captures a reference to the specific instance that call just created, the block equivalent of var new_enemy = instance_create_layer(...) in the code path (the code path uses this exact same name for it now, too).",
        "Then drag an \"Instance - Apply To\" block (Instances category) right after it. Open its Action Scope control (the small dropdown arrow in the block's own top-right corner, the same control every action block has) - instead of picking one of the listed Self/Other/object options, type new_enemy directly into the Expression field at the top of that pop-up. This re-points \"Self\" to that one captured instance for everything nested inside Apply To, without needing a wire or socket to carry the reference.",
        "Attach an \"If/Else\" block (Control category) to the RIGHT of Apply To, not below it - that's what nests the If/Else inside Apply To's scope rather than just running it after. Set its condition to enemy_count mod 2 = 0.",
        "Inside the If branch, drag a \"Set Instance Colour\" block (Instances category) - Colour field: make_color_rgb(200, 25, 55), a deep crimson red; Use Alpha: off. Inside the Else branch, the same block again - Colour: make_color_rgb(230, 100, 130), a lighter dusty rose/pink; Use Alpha: off. Because both live inside Apply To's scope, \"Self\" already means the one enemy just spawned - no targeting dropdown to fight with.",
        "This is the purpose-built block after all, not a workaround: Set Instance Colour's own Action Scope dropdown looks like it only offers Self, Other, or a specific object asset (which would recolor every enemy that exists so far, not just the new one) - but nesting it inside an Apply To block whose Expression field names the captured instance variable is what lets \"Self\" mean that one instance. Apply To is the piece that reaches a dynamically-created instance; once it's in place, every ordinary purpose-built block nested inside it just works normally.",
      ],
      level3Diagram: { file: "diagram_gml_visual_wireframe.png", w: 6.2, h: 4.34, caption: "The exact block wiring behind that Target field - Create Instance's Target captures the new instance, and Apply To's Action Scope Expression field is what carries that capture into every block nested to its right." },
      note: "See the Level 3 Visual-Track Equivalent above for exactly how the color-switching idea maps into GML Visual - it's the one Level 3 idea that isn't identical between Code and Visual, since it lives inside the spawn loop itself.",
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 4.",
      "Delete the single Enemy instance you manually placed on your scene last week.",
      "Add a new scene variable named EnemyIndex (Scene Properties > Scene Variables), default value 0.",
    ],
    level1: [
      "At the beginning of the scene, add a \"Repeat 5 times\" event. Inside it, add an action to create an Enemy instance at a position based on EnemyIndex, then increment EnemyIndex (code below).",
      "Click the green Preview button - you should see a row of enemies, and shooting any one of them should destroy just that one.",
    ],
    level1Code: [
      "At the beginning of the scene:",
      "Repeat 5 times:",
      "  Action: Create object Enemy at (40 + EnemyIndex * 80, 40)",
      "  Action: EnemyIndex = EnemyIndex + 1",
    ],
    level2: [
      "Shoot several different enemies in a row, including some misses - each hit deletes exactly the one Enemy instance that was actually touched, and every other enemy is unaffected.",
      "This already works with no extra code: the \"Bullet is in collision with Enemy\" condition and its Delete actions already operate per pair of overlapping instances, so every instance the Repeat event creates behaves independently and correctly, exactly the way Lesson 4 built it.",
    ],
    level3: [
      "Sound: this already works with no changes - the \"Play a sound\" action in the Combat event group runs independently for every enemy that gets hit.",
      "Visual: give each Enemy a different \"Tint color\" as it's created, based on whether EnemyIndex is even or odd, for a striped row effect. GDevelop's Tint color action takes an \"R;G;B\" string - use \"200;25;55\" (a deep crimson red) for even-indexed enemies and \"230;100;130\" (a lighter dusty rose/pink) for odd-indexed ones, so every other enemy is a visibly different shade of red rather than a blended new color.",
      "Structure: keep the spawn Repeat event's only job being creating and positioning enemies - group it under its own labeled event group (e.g. \"Spawn Enemies\"), separate from the \"Combat\" group from Lesson 4.",
    ],
    trouble: [
      "All enemies stack on top of each other: confirm EnemyIndex is actually being incremented inside the Repeat event, and used in the position action.",
      "Only one enemy appears: confirm the Repeat event's count is actually set to 5 (or your chosen number), not 1.",
      "Old and new enemies both appear: confirm the manually-placed Lesson 4 Enemy instance was actually deleted from the scene.",
    ],
    resource: "Student Resource List > GDevelop section, or gdevelop.io/tutorials",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: [
      "Launch TIC-80 and load your cart from Lesson 4.",
      "Press Esc to reach the code editor.",
    ],
    level1: [
      "Replace your single enemy_alive/ex/ey variables with an enemies table, populated by a for loop at the top of your code, above the TIC() function.",
      "Update the bullets loop so each bullet checks against every enemy in the enemies table (a nested loop), instead of just one fixed position.",
      "Update the drawing code at the bottom of TIC() to loop over the enemies table too, only drawing the ones still alive (code below).",
      "Press Esc, type run, press Enter, then shoot different enemies to test - each one should disappear independently when hit.",
    ],
    level1Code: [
      "enemies={}",
      "for enemy_count=0,4 do",
      " table.insert(enemies,{ex=20+enemy_count*20,ey=10,alive=true})",
      "end",
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
      "end",
    ],
    level2: [
      "Shoot several different enemies in a row, including some misses - each hit sets exactly the one enemies[j].alive that was actually touched to false, and every other enemy is unaffected.",
      "The inner enemies loop checks every enemy for every bullet, every frame, so this naturally scales to any number of enemies without extra code - a hit stops mattering for that enemy immediately, since e.alive and ... guards the check the same way it did for the single-enemy version.",
    ],
    level3: [
      "Sound: this already works with no changes - sfx(2) still fires from inside the same hit branch, once per actual hit, no matter which enemy or how many exist.",
      "Visual: draw a slightly different sprite index or use spr()'s flip argument based on each enemy's position in the table (j), for a bit of visual variety across the row.",
      "Structure: keep enemy state entirely inside the enemies table (never as loose separate variables again) - every new place that needs an enemy's position or status should read it from that table, the same way the bullets loop does.",
    ],
    trouble: [
      "Shooting one enemy removes all of them: confirm the alive flag is being set on enemies[j] specifically (e.alive=false), not on some variable shared by all enemies.",
      "Enemies don't line up in a row: confirm the spawn loop's ex=20+enemy_count*20 (or your chosen spacing) actually uses the loop variable enemy_count, not a fixed number.",
      "\"attempt to index a nil value\" error: confirm enemies={} and its populating for loop run before TIC(), at the very top of the code, same as bullets={}.",
    ],
    resource: "Student Resource List > TIC-80 section, or tic80.com",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 4.",
      "Select your existing enemy Actor on the Map, copy it (Ctrl+C), then paste (Ctrl+V) 4 more times, spacing each copy out in a row with at least a full enemy-width of empty gap between each one so they don't touch - unlike the other engines this week, 001 Game Creator doesn't need a loop at all for this part.",
    ],
    level1: [
      "Save, then Play/Test and shoot different enemies in the row - each one should disappear independently when hit, with zero new scripting.",
    ],
    level1Code: [
      "No new script needed - every pasted copy of the enemy Actor automatically keeps its own \"Collided with actor\" Trigger script, since Trigger scripts belong to the Actor Template, not to any one placed copy.",
    ],
    level2: [
      "Shoot several different enemies in a row, including some misses - each hit destroys exactly the one enemy Actor that was actually touched, and every other enemy is unaffected.",
      "This already works with no extra scripting: every placed copy of an Actor Template runs its own independent copy of that Template's Triggers, the same way every copy of an object/prefab/scene behaves independently in every other engine this course uses.",
    ],
    level3: [
      "Sound: this already works with no changes - the \"Play Sound\" event in the enemy's \"Collided with actor\" script runs independently for every copy that gets hit.",
      "Visual: give a couple of the pasted copies a slightly different Tint (Actor > Graphics Options > Tint, set per-instance rather than on the Template - a hex color picker) for a bit of row-to-row variety. Try #C81937 (a deep crimson red) on some copies and #E66482 (a lighter dusty rose/pink) on others, so the row shows visibly different shades of red rather than a blended new color.",
      "Structure: no change needed - the same \"everything about being hit lives on the enemy Actor's own Trigger\" structure from Lesson 4 already scales to any number of copies for free.",
    ],
    trouble: [
      "Only one enemy reacts to being shot: confirm each pasted copy is actually a separate placed Actor (Ctrl+V creates independent copies) and not, for example, the same Actor accidentally dragged back and forth.",
      "Enemies overlap each other: spread the pasted copies out further apart on the Map before saving.",
      "A copy doesn't react at all: confirm \"Initially Existent\" is still ticked on that copy - copy/paste sometimes needs this double-checked per copy depending on version.",
    ],
    resource: "001gamecreator.com/wiki, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 4.",
      "Delete the single BP_Enemy instance you manually placed in your level last week.",
    ],
    level1: [
      "In the Level Blueprint, add an Event BeginPlay node feeding a For Loop node (First Index = 0, Last Index = 4).",
      "Inside the loop, wire a \"Spawn Actor from Class\" node (Class = BP_Enemy), with its Transform's Location built from a base position offset by the loop's index (e.g. X = base + Index * 150).",
      "Click Compile, then Play - you should see a row of enemies, and shooting any one of them should destroy just that one.",
    ],
    level1Code: [
      "Level Blueprint:",
      "Event BeginPlay ->",
      "  For Loop(First Index = 0, Last Index = 4) ->",
      "    Spawn Actor from Class(BP_Enemy, Transform.Location = (BaseX + Index * 150, BaseY, 0))",
    ],
    level1EngineDiagram: { file: "diagram_unreal_blueprint_level1.png", w: 6.2, h: 4.712, caption: "The Level Blueprint's Event Graph, node by node - the node-for-node equivalent of the Level 1 nodes above." },
    level2: [
      "Shoot several different enemies in a row, including some misses - each hit destroys exactly the one BP_Enemy that was actually touched, and every other enemy is unaffected.",
      "This already works with no extra nodes: BP_Enemy's own Event ActorBeginOverlap and Cast To BP_Bullet logic is self-contained per instance, so every Actor the loop spawns behaves independently and correctly, exactly the way Lesson 4 built it.",
    ],
    level3: [
      "Sound: this already works with no changes - the Play Sound 2D node inside BP_Enemy's overlap event runs independently for every enemy that gets hit.",
      "Visual: vary each spawned enemy's material color based on the loop's Index, right after Spawn Actor from Class, calling a custom Event on the newly spawned Actor to set it - for a striped row effect. Build the two colors with a Make Linear Color node (not Make Color - Set Vector Parameter Value expects a Linear Color, whose R/G/B fields are 0-1 floats, not Make Color's 0-255 bytes): Dark Red = R 0.8, G 0.1, B 0.2 (a deep crimson red), Light Red = R 0.95, G 0.4, B 0.5 (a lighter dusty rose/pink) - the same two shades used in this course's Unity and Godot tracks, just expressed here as a 0-1 float instead of 0-255. A Select node driven by Index % 2 == 0 picks which one to pass in, so every other enemy is a visibly different shade of red, not a blended new color.",
      "Structure: keep the Level Blueprint's For Loop limited to just spawning and positioning enemies - resist the urge to also handle collision-response logic there; that still belongs entirely inside BP_Enemy's own Event Graph.",
    ],
    trouble: [
      "Only one enemy appears, or they're all stacked in the same spot: confirm the loop's Index output is actually wired into the Transform's Location, not a fixed value.",
      "Nothing spawns at all: confirm the For Loop's Last Index is actually greater than or equal to its First Index, and that the loop's body pin is connected to Spawn Actor from Class.",
      "Old and new enemies both appear: confirm the manually-placed Lesson 4 BP_Enemy instance was actually deleted from the level.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same for-loop spawn pattern works in C++ instead of Blueprint nodes.",
      steps: [
        "In your Game Mode's (or a dedicated spawner Actor's) BeginPlay(), add a for loop that calls GetWorld()->SpawnActor<ACppEnemy>() several times, each with a different location.",
        "Compile, drag a test setup into the level in place of your existing Blueprint spawn logic, Play, and test the same way as the Blueprint path.",
      ],
      code: [
        "// In BeginPlay() of your Game Mode or a spawner Actor:",
        "for (int32 EnemyCount = 0; EnemyCount < 5; EnemyCount++) {",
        "    FVector SpawnLoc(BaseX + EnemyCount * 150.f, BaseY, 0.f);",
        "    GetWorld()->SpawnActor<ACppEnemy>(SpawnLoc, FRotator::ZeroRotator);",
        "}",
      ],
      note: "This mirrors the Blueprint path's For Loop and Spawn Actor from Class node - same underlying engine calls, just from code. Level 2/3's per-instance logic already lives on ACppEnemy's own BeginPlay()/OnOverlap(), so nothing about them changes based on which path spawned the instances.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_05_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS (one per engine) - instructor only
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.0, { after: 120 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 5 Teacher's Key", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "INSTRUCTOR ONLY - do not distribute to students", bold: true, size: 18, color: "B00020" })] }),
    h2("Level 1 - Complete Worked Solution"),
    ...diagram("diagram_spawn_loop.png", 6.2, 4.712, "The spawn loop, in general - the same counter, repeat check, creation step, and increment behind Level 1 in every engine, whatever the exact blocks or syntax look like."),
    ...data.l1Notes.map(s => p(s)),
    code(data.l1Code),
  ];
  if (data.level1EngineDiagram) children.push(...diagram(data.level1EngineDiagram.file, data.level1EngineDiagram.w, data.level1EngineDiagram.h, data.level1EngineDiagram.caption));
  children.push(
    h2("Level 2 - Complete Worked Solution"),
    ...data.l2Notes.map(s => p(s)),
  );
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
  if (data.level3EngineDiagram) {
    children.push(h3(data.level3EngineDiagram.heading || "Complete Visual Sequence (all blocks combined)"));
    if (data.level3EngineDiagram.intro) children.push(p(data.level3EngineDiagram.intro));
    children.push(...diagram(data.level3EngineDiagram.file, data.level3EngineDiagram.w, data.level3EngineDiagram.h, data.level3EngineDiagram.caption));
  }
  if (data.altNotes) {
    children.push(h2("Alternate Path Notes: " + data.altNotes.title));
    if (data.altNotes.level1Diagram) children.push(...diagram(data.altNotes.level1Diagram.file, data.altNotes.level1Diagram.w, data.altNotes.level1Diagram.h, data.altNotes.level1Diagram.caption));
    children.push(...data.altNotes.notes.map(s => p(s)));
    if (data.altNotes.code) children.push(code(data.altNotes.code));
    if (data.altNotes.diagram) children.push(...diagram(data.altNotes.diagram.file, data.altNotes.diagram.w, data.altNotes.diagram.h, data.altNotes.diagram.caption));
    if (data.altNotes.fullDiagram) {
      children.push(h3(data.altNotes.fullDiagram.heading || "Complete Visual Sequence (all blocks combined)"));
      if (data.altNotes.fullDiagram.intro) children.push(p(data.altNotes.fullDiagram.intro));
      children.push(...diagram(data.altNotes.fullDiagram.file, data.altNotes.fullDiagram.w, data.altNotes.fullDiagram.h, data.altNotes.fullDiagram.caption));
    }
    children.push(...data.altNotes.why.map(s => bullet(s)));
  }

  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 5 Teacher's Key - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    l1Notes: ["A spawn script and a clone script on the Enemy sprite, plus a reverted (simpler) Bullet script - the hit-check moves onto the Enemy this week:"],
    l1Code: [
      "-- Enemy sprite:",
      "when green flag clicked",
      "  hide",
      "",
      "when green flag clicked",
      "  set [enemy_number] to 0",
      "  repeat 5",
      "    change [enemy_number] by 1",
      "    set [enemy_x] to ((enemy_number * 100) - 300)",
      "    create clone of [myself]",
      "",
      "when I start as a clone",
      "  show",
      "  go to x: (enemy_x) y: 150",
      "  forever",
      "    if <touching [Bullet]?> then",
      "      delete this clone",
      "",
      "-- Bullet sprite (back to Lesson 3's version - no Enemy check):",
      "when I start as a clone",
      "  show",
      "  go to x: (Player's x) y: (Player's y)",
      "  forever",
      "    change y by 10",
      "    if <y position > 170> then",
      "      delete this clone",
    ],
    l2Notes: ["Scratch clones have no unique ID, so a Bullet clone can't identify exactly which Enemy clone it's touching from the Bullet's side - that's why the check moved to each Enemy clone's own script, which only ever asks about itself. This already scales correctly to any number of enemies with no extra code, and a hit bullet simply isn't deleted early - it keeps moving and cleans itself up off-screen exactly like a miss."],
    l3Ideas: [
      {
        title: "Sound: already works",
        notes: ["If sfx_hit.wav was added to the Enemy's clone script in Lesson 4, it already plays correctly per-clone with zero changes needed this week:"],
        code: ["when I start as a clone", "  show", "  go to x: (enemy_x) y: 150", "  forever", "    if <touching [Bullet]?> then", "      start sound [sfx_hit]", "      delete this clone"],
      },
      {
        title: "Visual: striped row",
        notes: ["Alternate a color effect based on whether enemy_number is even or odd, right before creating each clone. Reminder for class: Scratch's color effect is a 0-200 hue dial, not an RGB color - a value like 50 just rotates the sprite's hue about a quarter of the way around the color wheel (roughly red toward yellow-green on a red sprite), so this produces \"every other enemy looks noticeably different,\" not two specific named colors:"],
        code: ["when green flag clicked", "  set [enemy_number] to 0", "  repeat 5", "    change [enemy_number] by 1", "    set [enemy_x] to ((enemy_number * 100) - 300)", "    create clone of [myself]"],
      },
      {
        title: "Structure: spawn vs. behavior stay separate",
        notes: ["The green-flag spawn script's only job is creating clones and setting their starting enemy_x - all hit-detection and hide-on-hit logic stays inside the \"when I start as a clone\" script. This is what let the Sound idea above keep working with zero changes."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, across the Enemy and Bullet sprites:"],
        code: [
          "-- Enemy sprite:",
          "when green flag clicked",
          "  hide",
          "",
          "when green flag clicked",
          "  set [enemy_number] to 0",
          "  repeat 5",
          "    change [enemy_number] by 1",
          "    set [enemy_x] to ((enemy_number * 100) - 300)",
          "    create clone of [myself]",
          "",
          "when I start as a clone",
          "  show",
          "  go to x: (enemy_x) y: 150",
          "  forever",
          "    if <touching [Bullet]?> then",
          "      start sound [sfx_hit]",
          "      delete this clone",
          "",
          "-- Bullet sprite:",
          "when I start as a clone",
          "  show",
          "  go to x: (Player's x) y: (Player's y)",
          "  forever",
          "    change y by 10",
          "    if <y position > 170> then",
          "      delete this clone",
        ],
      },
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    l1Notes: ["A for loop in the main scene's own script - enemy.tscn itself needs zero changes from Lesson 4:"],
    l1Code: [
      "# Main scene's script:",
      "var enemy_scene := preload(\"res://enemy.tscn\")",
      "",
      "func _ready():",
      "    for enemy_count in range(5):",
      "        var new_enemy := enemy_scene.instantiate()",
      "        new_enemy.position = Vector2(150 + enemy_count * 200, 50)",
      "        add_child(new_enemy)",
    ],
    l2Notes: ["enemy.tscn's own script (its Hitbox's area_entered connection and hit-response) is fully self-contained, so every instance the loop creates behaves independently and correctly with no extra code needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: already works",
        notes: ["Every spawned enemy carries its own Hit AudioStreamPlayer2D from Lesson 4 - no changes needed."],
      },
      {
        title: "Visual: striped row",
        notes: ["Set modulate based on the loop index right after instancing, from the spawner's own script. Color(0.8, 0.1, 0.2) is a deep crimson red (mostly red, small amounts of green/blue) and Color(0.95, 0.4, 0.5) is a lighter dusty rose/pink (same red-dominant mix, just brighter with more green/blue mixed in) - alternating them produces a row where every other enemy is a visibly different shade of red, not a blended new color:"],
        code: ["for enemy_count in range(5):", "    var new_enemy := enemy_scene.instantiate()", "    new_enemy.position = Vector2(150 + enemy_count * 200, 50)", "    if enemy_count % 2 == 0:", "        new_enemy.modulate = Color(0.8, 0.1, 0.2)", "    else:", "        new_enemy.modulate = Color(0.95, 0.4, 0.5)", "    add_child(new_enemy)"],
      },
      {
        title: "Structure: spawner vs. behavior stay separate",
        notes: ["The main scene's _ready() only ever creates and positions enemies - it never reaches into an enemy's hit-response logic, which stays entirely inside enemy.tscn's own script."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["The spawner loop with the visual idea added - sound needs no changes since it already lives on each enemy instance:"],
        code: [
          "var enemy_scene := preload(\"res://enemy.tscn\")",
          "",
          "func _ready():",
          "    for enemy_count in range(5):",
          "        var new_enemy := enemy_scene.instantiate()",
          "        new_enemy.position = Vector2(150 + enemy_count * 200, 50)",
          "        if enemy_count % 2 == 0:",
          "            new_enemy.modulate = Color(0.8, 0.1, 0.2)",
          "        else:",
          "            new_enemy.modulate = Color(0.95, 0.4, 0.5)",
          "        add_child(new_enemy)",
        ],
      },
    ],
  },
  {
    engine: "Unity", file: "Unity",
    l1Notes: ["A new EnemySpawner.cs on a new empty GameObject - Enemy.cs and Bullet.cs need zero changes from Lesson 4:"],
    l1Code: [
      "// EnemySpawner.cs:",
      "public GameObject enemyPrefab;",
      "public int enemyCount = 5;",
      "",
      "void Start() {",
      "    for (int enemyIndex = 0; enemyIndex < enemyCount; enemyIndex++) {",
      "        Vector3 pos = new Vector3(-4f + enemyIndex * 2f, 4f, 0f);",
      "        Instantiate(enemyPrefab, pos, Quaternion.identity);",
      "    }",
      "}",
    ],
    l2Notes: ["Bullet.cs's OnTriggerEnter2D and Enemy.cs's own Start() are fully self-contained per instance, so every prefab copy the loop creates behaves independently and correctly with no extra code needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: already works",
        notes: ["AudioSource.PlayClipAtPoint() in Bullet.cs runs independently for every enemy that gets hit - no changes needed."],
      },
      {
        title: "Visual: striped row",
        notes: ["Set the SpriteRenderer color right after Instantiate() in the spawner, using the loop index. new Color(0.8f, 0.1f, 0.2f) is a deep crimson red and new Color(0.95f, 0.4f, 0.5f) is a lighter dusty rose/pink - alternating them by enemyIndex % 2 makes every other enemy a visibly different shade of red, not a blended new color. Written as a plain if/else rather than a one-line ternary - readable at a glance for a student who hasn't met the ternary operator yet (see this lesson's Book Notes entry: that's deliberately saved for a later, more advanced project):"],
        code: [
          "for (int enemyIndex = 0; enemyIndex < enemyCount; enemyIndex++) {",
          "    Vector3 pos = new Vector3(-4f + enemyIndex * 2f, 4f, 0f);",
          "    GameObject newEnemy = Instantiate(enemyPrefab, pos, Quaternion.identity);",
          "    if (enemyIndex % 2 == 0) {",
          "        newEnemy.GetComponent<SpriteRenderer>().color = new Color(0.8f, 0.1f, 0.2f);",
          "    } else {",
          "        newEnemy.GetComponent<SpriteRenderer>().color = new Color(0.95f, 0.4f, 0.5f);",
          "    }",
          "}",
        ],
      },
      {
        title: "Structure: spawner vs. behavior stay separate",
        notes: ["EnemySpawner.cs only ever creates and positions enemies - collision-response and hit-sound logic stay entirely inside Enemy.cs and Bullet.cs."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["The spawner loop with the visual idea added - sound needs no changes since it already lives on Bullet.cs:"],
        code: [
          "public GameObject enemyPrefab;",
          "public int enemyCount = 5;",
          "",
          "void Start() {",
          "    for (int enemyIndex = 0; enemyIndex < enemyCount; enemyIndex++) {",
          "        Vector3 pos = new Vector3(-4f + enemyIndex * 2f, 4f, 0f);",
          "        GameObject newEnemy = Instantiate(enemyPrefab, pos, Quaternion.identity);",
          "        if (enemyIndex % 2 == 0) {",
          "            newEnemy.GetComponent<SpriteRenderer>().color = new Color(0.8f, 0.1f, 0.2f);",
          "        } else {",
          "            newEnemy.GetComponent<SpriteRenderer>().color = new Color(0.95f, 0.4f, 0.5f);",
          "        }",
          "    }",
          "}",
        ],
      },
    ],
    altNotes: {
      title: "Visual Scripting",
      level1Diagram: { file: "diagram_unity_visual_scripting_level1.png", w: 6.2, h: 4.96, caption: "The Script Graph, node by node - Instantiate's own output pin plugs straight into Set Position's Target, no extra step needed." },
      notes: [
        "A For Loop node (0 to enemyCount) feeding Instantiate + Set Position nodes is the graph equivalent of the C# for loop.",
        "The Level 3 goals map onto nodes the same way: a \"Set Member\" node on the newly spawned instance's Sprite Renderer color for the visual idea - Sound needs no changes either way, since it lives on the Bullet's own graph, not the spawner's.",
      ],
      why: [
        "Keeping the same Structure principle (a spawner only spawns; behavior lives on the spawned object) whether a student is in C# or graphs reinforces that it's a design idea, not a syntax rule.",
      ],
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    l1Notes: ["A new obj_spawner with a Create event loop - obj_enemy and obj_bullet need zero changes from Lesson 4:"],
    l1Code: [
      "// obj_spawner's Create event:",
      "for (var enemy_count = 0; enemy_count < 5; enemy_count++) {",
      "    instance_create_layer(80 + enemy_count * 96, 80, \"Instances\", obj_enemy);",
      "}",
    ],
    l2Notes: ["obj_bullet's Collision event with obj_enemy already reacts per pair of overlapping instances, so every instance the loop creates behaves independently and correctly with no extra code needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: already works",
        notes: ["audio_play_sound() in obj_bullet's Collision event runs independently for every enemy that gets hit - no changes needed."],
      },
      {
        title: "Visual: striped row",
        notes: ["Set image_blend on each new instance right after creating it, using with(). make_color_rgb(200, 25, 55) is a deep crimson red and make_color_rgb(230, 100, 130) is a lighter dusty rose/pink - image_blend multiplies the sprite's own pixels by this tint, so alternating the two makes every other enemy a visibly different shade of red, not a blended new color. Written as a plain if/else rather than a one-line ternary - readable at a glance for a student who hasn't met the ternary operator yet (see this lesson's Book Notes entry: that's deliberately saved for a later, more advanced project):"],
        code: [
          "for (var enemy_count = 0; enemy_count < 5; enemy_count++) {",
          "    var new_enemy = instance_create_layer(80 + enemy_count * 96, 80, \"Instances\", obj_enemy);",
          "    with (new_enemy) {",
          "        if (enemy_count % 2 == 0) {",
          "            image_blend = make_color_rgb(200, 25, 55);",
          "        } else {",
          "            image_blend = make_color_rgb(230, 100, 130);",
          "        }",
          "    }",
          "}",
        ],
      },
      {
        title: "Structure: spawner vs. behavior stay separate",
        notes: ["obj_spawner's Create event only ever creates and positions enemies - collision-response and hit-sound logic stay entirely inside obj_enemy and obj_bullet's own events."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["The spawner loop with the visual idea added - sound needs no changes since it already lives on obj_bullet:"],
        code: [
          "// obj_spawner's Create event:",
          "for (var enemy_count = 0; enemy_count < 5; enemy_count++) {",
          "    var new_enemy = instance_create_layer(80 + enemy_count * 96, 80, \"Instances\", obj_enemy);",
          "    with (new_enemy) {",
          "        if (enemy_count % 2 == 0) {",
          "            image_blend = make_color_rgb(200, 25, 55);",
          "        } else {",
          "            image_blend = make_color_rgb(230, 100, 130);",
          "        }",
          "    }",
          "}",
        ],
      },
    ],
    altNotes: {
      title: "GML Visual",
      level1Diagram: { file: "diagram_gml_visual_level1.png", w: 6.2, h: 4.712, caption: "The five blocks above, wired together - the block-for-block equivalent of the Level 1 code." },
      notes: [
        "A \"Set Variable\" block (Variable = enemy_count, Value = 0) initializes the loop counter, a \"Repeat\" block (Times = 5) is the block equivalent of the for loop's condition, an \"Instance - Create Instance\" block inside it (Object = obj_enemy, Layer = \"Instances\", X = 80 + (enemy_count * 96), Y = 80) does the spawning, and a second \"Set Variable\" block (Variable = enemy_count, Value = enemy_count + 1) right after it increments the counter - together these five blocks are the exact equivalent of the code path's for (var enemy_count = 0; enemy_count < 5; enemy_count++) { instance_create_layer(...) }.",
        "The Level 3 Visual idea is the one piece that isn't identical between Code and Visual, since it lives inside the spawn loop itself rather than on obj_enemy: give Create Instance a Target field set to a new local variable, new_enemy, to capture a reference to the instance it just made - the code path now captures it under this exact same name. Then drag an \"Instance - Apply To\" block right after it, and in Apply To's own Action Scope pop-up, type new_enemy into the Expression field at the top instead of picking Self/Other/an object - this re-points \"Self\" to that one captured instance for anything nested inside Apply To. Attach an If/Else block to the RIGHT of Apply To (not below it, which is what actually nests the If/Else inside its scope) on enemy_count mod 2 = 0, and inside each branch drop a \"Set Instance Colour\" block - Colour = make_color_rgb(200, 25, 55) in the If branch, make_color_rgb(230, 100, 130) in the Else branch, matching the code path's if/else exactly, block for statement.",
        "This resolves a case that briefly looked like a genuine exception to preferring purpose-built blocks: Set Instance Colour's own Action Scope dropdown only lists Self, Other, or a specific object asset (which would recolor every existing enemy, not just the new one), so it looks like it can't reach an instance held in a variable - the same limit Set Instance Variable has. Apply To is what closes that gap: its Expression field accepts the exact same captured variable Create Instance's Target field produced, and nesting the purpose-built color block inside that scope makes \"Self\" mean the one instance that variable refers to. No generic block or dot-notation workaround needed after all - see this lesson's Book Notes entry for how this was found, correcting an earlier note.",
        "Sound needs no changes either way, since it lives on obj_bullet's own Collision event, not the spawner's.",
      ],
      diagram: { file: "diagram_gml_visual_wireframe.png", w: 6.2, h: 4.34, caption: "The exact block wiring behind that Target field - Create Instance's Target captures the new instance, and Apply To's Action Scope Expression field is what carries that capture into every block nested to its right, including the purpose-built Set Instance Colour block." },
      fullDiagram: {
        heading: "Complete Visual Sequence (all blocks combined)",
        intro: "Same idea as the Complete Level 3 solution shown in code above - here's every block in that solution, wired in the order it actually runs, not just the two blocks doing the tricky part.",
        file: "diagram_gml_visual_full_sequence.png", w: 6.2, h: 4.77,
        caption: "obj_spawner's entire Create event in GML Visual - Set Variable, Repeat, Create Instance, Apply To with its Expression-scoped If/Else color branch, and the counter increment, exactly mirroring the code path's for loop.",
      },
      why: [
        "Keeping the same Structure principle (a spawner only spawns; behavior lives on the spawned object) whether a student is in Code or Visual reinforces that it's a design idea, not a syntax rule.",
        "Naming the exact block, category, and field for every step (rather than describing the shape of the logic) matters more in GML Visual than in most engines' visual tracks - and this lesson's own color-switch idea is a good example of why: the obvious-looking purpose-built block (Set Instance Colour) looks like it can't target a specific, dynamically-created instance at first glance, the same limit Set Instance Variable has. What actually unlocks it is a separate scope-setting block, Apply To, whose Action Scope Expression field takes the same captured variable Create Instance's Target produced - confirmed by hands-on testing, and worth remembering the next time a purpose-built block's dropdown looks like a dead end: check whether an Apply To block nested around it would solve it before reaching for a generic, typed-expression workaround.",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    l1Notes: ["A \"Repeat 5 times\" event at the beginning of the scene, using a new EnemyIndex scene variable - the Combat event group from Lesson 4 needs zero changes:"],
    l1Code: [
      "At the beginning of the scene:",
      "Repeat 5 times:",
      "  Action: Create object Enemy at (40 + EnemyIndex * 80, 40)",
      "  Action: EnemyIndex = EnemyIndex + 1",
    ],
    l2Notes: ["The \"Bullet is in collision with Enemy\" condition and its Delete actions already operate per pair of overlapping instances, so every instance the Repeat event creates behaves independently and correctly with no extra events needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: already works",
        notes: ["The \"Play a sound\" action inside the Combat event group runs independently for every enemy that gets hit - no changes needed."],
      },
      {
        title: "Visual: striped row",
        notes: ["Tint each Enemy differently as it's created, based on whether EnemyIndex is even or odd. \"200;25;55\" is a deep crimson red and \"230;100;130\" is a lighter dusty rose/pink (GDevelop's Tint color action takes an \"R;G;B\" string) - alternating them makes every other enemy a visibly different shade of red, not a blended new color:"],
        code: [
          "Repeat 5 times:",
          "  Action: Create object Enemy at (40 + EnemyIndex * 80, 40)",
          "  Condition: EnemyIndex mod 2 = 0",
          "  Action: Tint color of Enemy = \"200;25;55\"",
          "  Condition: EnemyIndex mod 2 = 1",
          "  Action: Tint color of Enemy = \"230;100;130\"",
          "  Action: EnemyIndex = EnemyIndex + 1",
        ],
      },
      {
        title: "Structure: group the spawn loop separately",
        notes: ["Put the Repeat spawn event in its own labeled group (e.g. \"Spawn Enemies\"), separate from the \"Combat\" group from Lesson 4 - the two mechanics stay easy to find as the event sheet grows."],
      },
      {
        title: "Complete Level 3 solution (all 2 goals combined)",
        notes: ["The spawn Repeat event with the visual idea added, as its own group - sound needs no changes since it already lives in the Combat group:"],
        code: [
          "Event group: \"Spawn Enemies\"",
          "At the beginning of the scene:",
          "Repeat 5 times:",
          "  Action: Create object Enemy at (40 + EnemyIndex * 80, 40)",
          "  Condition: EnemyIndex mod 2 = 0",
          "  Action: Tint color of Enemy = \"200;25;55\"",
          "  Condition: EnemyIndex mod 2 = 1",
          "  Action: Tint color of Enemy = \"230;100;130\"",
          "  Action: EnemyIndex = EnemyIndex + 1",
        ],
      },
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    l1Notes: ["An enemies table populated by a for loop at startup, plus a nested loop inside the existing bullets loop:"],
    l1Code: [
      "enemies={}",
      "for enemy_count=0,4 do",
      " table.insert(enemies,{ex=20+enemy_count*20,ey=10,alive=true})",
      "end",
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
      "end",
    ],
    l2Notes: ["The inner enemies loop checks every enemy for every bullet, every frame, so this naturally scales to any number of enemies with no extra code - a hit sets e.alive=false for exactly the one enemy whose bounds actually overlapped that bullet."],
    l3Ideas: [
      {
        title: "Sound: already works",
        notes: ["sfx(2) can stay exactly where it was added in Lesson 4's hit branch - add it inside the hit=true assignment if not already there:"],
        code: [
          "   if e.alive and b.bx<e.ex+8 and b.bx+8>e.ex and b.by<e.ey+8 and b.by+8>e.ey then",
          "    e.alive=false",
          "    hit=true",
          "    sfx(2)",
          "   end",
        ],
      },
      {
        title: "Visual: row variety",
        notes: ["Use each enemy's index j to flip alternating enemies for a bit of visual variety:"],
        code: [" for j=1,#enemies do", "  local e=enemies[j]", "  if e.alive then spr(3,e.ex,e.ey,0,1,j%2) end", " end"],
      },
      {
        title: "Structure: enemy state lives in the table",
        notes: ["Every place that needs an enemy's position or status reads it from the enemies table - resist ever declaring a loose ex/ey/alive variable outside it again."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined into the one TIC() function:"],
        code: [
          "enemies={}",
          "for enemy_count=0,4 do",
          " table.insert(enemies,{ex=20+enemy_count*20,ey=10,alive=true})",
          "end",
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
          "    hit=true",
          "    sfx(2)",
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
          "  if e.alive then spr(3,e.ex,e.ey,0,1,j%2) end",
          " end",
          " spr(1,x,y,0)",
          "end",
        ],
      },
    ],
  },
  {
    engine: "001 Game Creator", file: "001GameCreator",
    l1Notes: ["No loop and no new scripting at all - just more copies of the same Actor Template placed on the Map. This is the one engine this week where \"multiple enemies\" is a placement task, not a code task:"],
    l1Code: [
      "Copy/paste the enemy Actor 4 more times on the Map, spaced out in a row with at least a full enemy-width of gap between each one.",
      "Each copy automatically keeps its own working \"Collided with actor\" Trigger script from Lesson 4:",
      "enemy Actor Trigger: Collided with actor ->",
      "  Actor – Destroy/Revive Actor(relate, Destroy)",
      "  Actor – Destroy/Revive Actor(this, Destroy)",
    ],
    l2Notes: ["Every placed copy of an Actor Template runs its own independent copy of that Template's Triggers - this is the same underlying idea as \"every instance behaves independently\" in every other engine this course uses, just expressed through placement instead of a spawn loop. No extra scripting is needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: already works",
        notes: ["The \"Play Sound\" event in the enemy's \"Collided with actor\" script runs independently for every copy that gets hit - no changes needed."],
      },
      {
        title: "Visual: row variety",
        notes: ["Select an individual pasted copy and adjust its own Tint (Actor > Graphics Options > Tint, per-instance override, not the Template) for a couple of enemies to look slightly different from the rest. #C81937 is a deep crimson red and #E66482 is a lighter dusty rose/pink - the same two shades used across this course's other engines this lesson, just as hex codes here instead of RGB numbers."],
      },
      {
        title: "Structure: no change needed",
        notes: ["The same \"everything about being hit lives on the enemy Actor's own Trigger\" structure from Lesson 4 already scales to any number of copies for free - there's nothing new to organize this week."],
      },
      {
        title: "Complete Level 3 solution (all 2 goals combined)",
        notes: ["The enemy Actor Template's Trigger script, unchanged from Lesson 4, plus per-instance Tint variation on a couple of the pasted copies:"],
        code: [
          "enemy Actor Template Trigger: Collided with actor ->",
          "  Play Sound(sfx_hit.wav)",
          "  Actor – Destroy/Revive Actor(relate, Destroy)",
          "  Actor – Destroy/Revive Actor(this, Destroy)",
          "(plus: 2 of the 5 pasted copies given a per-instance Tint of #C81937 - deep crimson red - and the other 3 left default or set to #E66482 - lighter dusty rose/pink)",
        ],
      },
    ],
  },
  {
    engine: "Unreal", file: "Unreal",
    l1Notes: ["A For Loop node in the Level Blueprint - BP_Enemy needs zero changes from Lesson 4:"],
    l1Code: [
      "Level Blueprint:",
      "Event BeginPlay ->",
      "  For Loop(First Index = 0, Last Index = 4) ->",
      "    Spawn Actor from Class(BP_Enemy, Transform.Location = (BaseX + Index * 150, BaseY, 0))",
    ],
    level1EngineDiagram: { file: "diagram_unreal_blueprint_level1.png", w: 6.2, h: 4.712, caption: "The Level Blueprint's Event Graph, node by node - the node-for-node equivalent of the Level 1 nodes above." },
    l2Notes: ["BP_Enemy's own Event ActorBeginOverlap and Cast To BP_Bullet logic is fully self-contained per instance, so every Actor the loop spawns behaves independently and correctly with no extra nodes needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: already works",
        notes: ["The Play Sound 2D node inside BP_Enemy's overlap event runs independently for every enemy that gets hit - no changes needed."],
      },
      {
        title: "Visual: striped row",
        notes: ["Add a custom Event (e.g. SetStripeColor) to BP_Enemy that creates its dynamic material instance with a passed-in color, then call it from the Level Blueprint's loop right after each Spawn Actor from Class, alternating colors based on Index. Build Dark Red and Light Red with a Make Linear Color node, not Make Color - Set Vector Parameter Value expects a Linear Color, whose R/G/B fields are 0-1 floats, not Make Color's 0-255 bytes: Dark Red = R 0.8, G 0.1, B 0.2 (a deep crimson red), Light Red = R 0.95, G 0.4, B 0.5 (a lighter dusty rose/pink) - the same two shades used in this course's Unity/Godot tracks, just as a 0-1 float here instead of 0-255:"],
        code: [
          "BP_Enemy: Custom Event SetStripeColor(Color) ->",
          "  Create Dynamic Material Instance(Self's Material) -> Set Vector Parameter Value(\"Color\", Color)",
          "",
          "Level Blueprint:",
          "Dark Red = Make Linear Color(R=0.8, G=0.1, B=0.2)",
          "Light Red = Make Linear Color(R=0.95, G=0.4, B=0.5)",
          "",
          "Inside the For Loop:",
          "Spawn Actor from Class(BP_Enemy, ...) -> (Return Value) -> SetStripeColor(Index % 2 == 0 ? Dark Red : Light Red)",
        ],
      },
      {
        title: "Structure: spawner vs. behavior stay separate",
        notes: ["The Level Blueprint's For Loop only ever spawns and positions enemies - collision-response and hit-sound logic stay entirely inside BP_Enemy's own Event Graph."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["The spawn loop with the visual idea added - sound needs no changes since it already lives on BP_Enemy:"],
        code: [
          "BP_Enemy: Custom Event SetStripeColor(Color) ->",
          "  Create Dynamic Material Instance(Self's Material) -> Set Vector Parameter Value(\"Color\", Color)",
          "",
          "Level Blueprint:",
          "Dark Red = Make Linear Color(R=0.8, G=0.1, B=0.2)",
          "Light Red = Make Linear Color(R=0.95, G=0.4, B=0.5)",
          "Event BeginPlay ->",
          "  For Loop(First Index = 0, Last Index = 4) ->",
          "    Spawn Actor from Class(BP_Enemy, Transform.Location = (BaseX + Index * 150, BaseY, 0))",
          "    -> (Return Value) -> SetStripeColor(Index % 2 == 0 ? Dark Red : Light Red)",
        ],
      },
    ],
    level3EngineDiagram: {
      heading: "Complete Visual Sequence (all nodes combined)",
      intro: "Every node in the Complete Level 3 solution above, wired in the order it actually runs.",
      file: "diagram_unreal_blueprint_full_sequence.png", w: 6.2, h: 6.448,
      caption: "Spawn Actor from Class's own Return Value pin plugs straight into Set Vector Parameter Value's Target - the same direct-wire pattern as Unity's Instantiate, no promote-to-variable step needed.",
    },
    altNotes: {
      title: "C++",
      notes: [
        "A plain C++ for loop calling GetWorld()->SpawnActor<ACppEnemy>() several times is the code equivalent of the Blueprint For Loop node.",
        "The Level 3 visual idea maps onto a public method on ACppEnemy called right after each SpawnActor call - Sound needs no changes either way, since it lives on ACppEnemy's own OnOverlap(), not the spawner.",
      ],
      why: [
        "Showing both C++ and Blueprint side by side for the same behavior keeps demystifying Blueprints for students who assume \"visual\" means \"not real code\" - it's the same engine calls either way.",
        "This is squarely the optional advanced track - most students should stay on Blueprints; C++ is for students already comfortable with typed languages from Godot, GameMaker, or Unity's C# path.",
      ],
    },
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_05_Key_${d.file}.docx`]));

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
