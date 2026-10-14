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
      headers: { default: brand.pageHeader("Lesson 11 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 11 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("This Is a Milestone Lesson"),
        p("Once multiple waves work today, the capstone reaches the Level 2 Complete milestone from the Feature-to-Lesson Map - the second of the year's big checkpoints, after Lesson 6's Level 1 MVP. This deserves the same extra celebratory framing Lesson 6 got: a dedicated Milestone Moment discussion, and a real share-out at the end of lab, not just a quick check."),
        checkbox("Read the Teacher's Script's Milestone Moment section before class, and plan to give it genuine time - this is a good one to slow down for."),

        h2("This Week: Splitting a Decision Into Two Branches"),
        p("Every engine's Lesson 6 win-check has always asked one question: 'are all enemies destroyed?' This week, that question grows a second part: 'AND is this the last wave?' If yes to both, it's a real win. If enemies are cleared but MORE waves remain, the round continues instead - a fresh wave spawns, and score/lives/game_active are untouched."),
        p("The wave-spawning logic itself is written ONCE and reused three ways: at the very first game start (wave 1), every time a wave is cleared with more remaining (wave 2, wave 3), and after every restart (back to wave 1) - the same reuse principle from Lesson 10, applied one level deeper."),
        checkbox("Read through each engine's Handout once before class - every engine restructures its existing win-check into an if/else this week, not just adds new code."),

        h2("Self-Contained Lesson Assets"),
        p("No new assets this week - multiple waves reuse the exact same enemy spawn logic and placeholder sprites already in every project."),

        h2("Continuity From Lesson 10"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 10, not create a new one."),
        checkbox("If a student's Lesson 10 project is lost or broken, have a pre-built \"Lesson 10 complete\" starter available per engine (from the Lesson 10 Teacher's Keys) so they aren't blocked."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop's Lesson 9-10 exercises (scoreboard, freeze/restart) still work for all 8 engines."),
        checkbox("Test a full 3-wave playthrough on every engine before class: clear wave 1 (round continues), clear wave 2 (round continues), clear wave 3 (real win, with restart still working correctly back to wave 1)."),
        checkbox("Confirm losing a life mid-wave (at any wave number) still correctly triggers Game Over, not just a wave transition - this is an easy thing to accidentally break while restructuring the win-check."),
        checkbox("001 Game Creator (Windows demo laptop): confirm the new SpawnWave Custom Event correctly revives and repositions all 5 enemy copies when called mid-game (not just from ResetGame)."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies."),

        h2("Known Gotchas This Week"),
        bullet("The most likely mistake: forgetting the wave-count check entirely, so clearing wave 1 immediately triggers the full win instead of advancing to wave 2. Test this explicitly - the win message should NOT appear until wave 3 (or your total_waves value) is cleared."),
        bullet("A close second: the wave-advance branch accidentally also resetting score or lives (it shouldn't - only enemies_remaining and the wave counter should change between waves)."),
        bullet("Restarting mid-way through wave 2 or 3 should always bring the player back to wave 1, not wherever they left off - test this directly, not just from a fresh start."),
        bullet("This is a genuinely exciting lesson to finish - encourage a full 3-wave playthrough as part of the end-of-lesson share-out, and treat it with the same energy as Lesson 6's milestone celebration."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_11_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 11: Wave After Wave", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: clearing a wave of enemies no longer means the game is over - it means the NEXT wave begins. Clear all 3, and that's when you really win.", { bold: true }),
    p("Open your own [YourName]Game project from Lesson 10 - don't start a new one.", { bold: true }),
    p("Tip: no new assets this week - multiple waves reuse your exact existing enemy spawn logic. This completes the Level 2 Complete milestone!", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("Clearing a wave's enemies spawns a fresh wave, up to 3 total. Clearing wave 3 triggers the real win. The current wave number is shown on screen.", { italics: true }),
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
    p("Wave transitions are seamless and correct every time - no double-spawns, no skipped waves, exactly 3 waves before the real win, and restarting always returns to wave 1.", { italics: true }),
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
      headers: { default: brand.pageHeader("Lesson 11 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 10.",
      "Create two new variables: wave (for all sprites) and total_waves (for all sprites).",
    ],
    level1: [
      "On the Enemy sprite, find your \"when I receive [start game]\" script. Add \"set wave to 1\" and \"set total_waves to 3\" alongside the existing resets, then replace the spawn loop at the end with \"broadcast [spawn wave]\".",
      "Add a new script on the Enemy sprite: \"when I receive [spawn wave]\" containing exactly your old spawn loop (repeat 5: create clone of myself; ...) - this is now the single reusable piece that creates a wave of enemies.",
      "In the Enemy's touching-Bullet branch, right after changing enemies_left, revise the win-check into an if/else (code below): if this was the last wave, win as before; otherwise, advance the wave and spawn a new one.",
      "Add wave and total_waves to your visible variables on stage (check their checkboxes) so the current wave is visible during play.",
      "Click the stage and play through - clearing wave 1 or 2 should spawn a fresh wave instead of ending the game; clearing wave 3 should show \"You Win!\".",
    ],
    level1Code: [
      "Enemy's touching-Bullet branch (revised):",
      "if <touching Bullet?> then",
      "  change score by 10",
      "  change enemies_left by -1",
      "  if <enemies_left = 0> then",
      "    if <wave = total_waves> then",
      "      set game_active to 0",
      "      broadcast [you win]",
      "    else",
      "      change wave by 1",
      "      set enemies_left to 5",
      "      broadcast [spawn wave]",
      "  delete this clone",
    ],
    level2: [
      "Play a full 3-wave game: confirm waves 1 and 2 each spawn a fresh set of enemies with no win message, and only wave 3 triggers \"You Win!\".",
      "Let a life run out partway through wave 2 and confirm \"Game Over\" still works correctly, not a wave transition.",
      "Press R after a win or loss and confirm wave resets to 1, along with score and lives.",
    ],
    level3: [
      "Visual: briefly announce each new wave, e.g. \"say [Wave 2!] for 1 seconds\" on the Player sprite right when \"spawn wave\" is broadcast for a wave 2 or later.",
      "Structure: notice \"when I receive [spawn wave]\" is now called from THREE different moments - the very first game start, every wave clear, and (indirectly, via \"start game\") every restart. The actual enemy-creation code only exists once.",
    ],
    trouble: [
      "Wave 1 immediately shows \"You Win!\": confirm the if/else checks \"wave = total_waves\" (not just enemies_left = 0 alone) - without the wave check, any cleared wave looks like a full win.",
      "New waves never spawn: confirm \"broadcast [spawn wave]\" is inside the else branch, and that the Enemy sprite actually has a \"when I receive [spawn wave]\" script.",
      "Score or lives reset unexpectedly between waves: confirm the else branch only changes wave and enemies_left - nothing else.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 10.",
      "In your main scene's script, add var wave := 1 and var total_waves := 3.",
      "Add a WaveText node to your main scene, visible, positioned near ScoreText/LivesText.",
    ],
    level1: [
      "Split your existing enemy-creation code out of start_game() into its own function spawn_wave() (code below) - just the loop that creates 5 enemies, nothing else.",
      "In start_game(), add wave = 1 and call spawn_wave() at the end instead of spawning directly.",
      "In update_hud(), add a line updating WaveText's text from wave and total_waves.",
      "In _on_enemy_destroyed(), revise the win-check into an if/else (code below): if this was the last wave, win as before; otherwise, advance the wave and call spawn_wave() again.",
      "Run the scene and play through - clearing wave 1 or 2 should spawn a fresh wave instead of ending the game; clearing wave 3 should show WinText.",
    ],
    level1Code: [
      "func spawn_wave():",
      "    for i in range(5):",
      "        var enemy = enemy_scene.instantiate()",
      "        add_child(enemy)",
      "        enemy.add_to_group(\"enemies\")",
      "        enemy.destroyed.connect(_on_enemy_destroyed)",
      "        enemy.escaped.connect(_on_enemy_escaped)",
      "        # ...existing per-enemy positioning...",
      "",
      "func start_game():",
      "    for enemy in get_tree().get_nodes_in_group(\"enemies\"):",
      "        enemy.queue_free()",
      "    score = 0",
      "    lives = 3",
      "    wave = 1",
      "    enemies_remaining = 5",
      "    game_active = true",
      "    $WinText.hide()",
      "    $LoseText.hide()",
      "    update_hud()",
      "    spawn_wave()",
      "",
      "func _on_enemy_destroyed():",
      "    enemies_remaining -= 1",
      "    score += 10",
      "    update_hud()",
      "    if enemies_remaining <= 0:",
      "        if wave >= total_waves:",
      "            game_active = false",
      "            $WinText.show()",
      "        else:",
      "            wave += 1",
      "            enemies_remaining = 5",
      "            update_hud()",
      "            spawn_wave()",
    ],
    level2: [
      "Play a full 3-wave game: confirm waves 1 and 2 each spawn a fresh set of enemies with no win message, and only wave 3 triggers WinText.",
      "Let a life run out partway through wave 2 and confirm LoseText still works correctly, not a wave transition.",
      "Press R after a win or loss and confirm wave resets to 1, along with score and lives.",
    ],
    level3: [
      "Visual: briefly show a \"Wave 2!\" style message using a Timer-based fade or a temporary Label, right when a new wave spawns (not the win/lose labels, which stay reserved for the round's actual end).",
      "Structure: notice spawn_wave() is now called from three different moments - start_game() (indirectly, for wave 1), and _on_enemy_destroyed()'s else branch (every subsequent wave). The actual enemy-creation code only exists once.",
    ],
    trouble: [
      "Wave 1 immediately shows WinText: confirm the if/else checks \"wave >= total_waves\" (not just enemies_remaining <= 0 alone).",
      "New waves never spawn: confirm spawn_wave() is actually called inside the else branch, and that it's spelled identically to its function definition.",
      "Enemies from the new wave don't get cleaned up correctly: confirm every enemy created in spawn_wave() is still added to the \"enemies\" group, not just the ones from start_game().",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Open your [YourName]Game project from Lesson 10.",
      "GameObject > UI (Canvas) > Legacy > Text, rename it WaveText, position it near ScoreText/LivesText.",
    ],
    level1: [
      "In GameManager.cs, add public static int wave = 1 and public static int totalWaves = 3.",
      "In EnemySpawner.cs, rename your SpawnEnemies() usage so it's clearly reusable as-is (it already just spawns, no changes needed there) - confirm it doesn't reset any GameManager variables itself.",
      "In HUD.cs, add a public Text waveText field and update Update() to also set it from GameManager.wave and totalWaves.",
      "In GameManager.cs's RegisterKill(), revise the win-check into an if/else (code below): if this was the last wave, win as before; otherwise, advance the wave and call EnemySpawner.instance.SpawnEnemies() again.",
      "Press Play and play through - clearing wave 1 or 2 should spawn a fresh wave instead of ending the game; clearing wave 3 should show the win message.",
    ],
    level1Code: [
      "// GameManager.cs:",
      "public static int wave = 1;",
      "public static int totalWaves = 3;",
      "",
      "public static void RegisterKill() {",
      "    enemiesRemaining--;",
      "    score += 10;",
      "    if (enemiesRemaining <= 0) {",
      "        if (wave >= totalWaves) {",
      "            hasWon = true;",
      "        } else {",
      "            wave++;",
      "            enemiesRemaining = 5;",
      "            EnemySpawner.instance.SpawnEnemies();",
      "        }",
      "    }",
      "}",
      "",
      "// ResetGame() - add:",
      "wave = 1;",
    ],
    level2: [
      "Play a full 3-wave game: confirm waves 1 and 2 each spawn a fresh set of enemies with no win message, and only wave 3 triggers the win message.",
      "Let a life run out partway through wave 2 and confirm the lose message still works correctly, not a wave transition.",
      "Press R after a win or loss and confirm wave resets to 1, along with score and lives.",
    ],
    level3: [
      "Visual: briefly flash a \"Wave 2!\" style message using a separate GUI.Label with a short timer, right when a new wave spawns.",
      "Structure: notice EnemySpawner.instance.SpawnEnemies() is now called from two different moments in gameplay (ResetGame(), and RegisterKill()'s else branch) plus the very first Start() - the actual spawning code only exists once.",
    ],
    trouble: [
      "Wave 1 immediately shows the win message: confirm the if/else checks \"wave >= totalWaves\" (not just enemiesRemaining <= 0 alone).",
      "New waves never spawn: confirm EnemySpawner.instance.SpawnEnemies() is actually called inside the else branch.",
      "Restarting leaves wave at whatever it was: confirm ResetGame() explicitly sets wave = 1 alongside its other resets.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same wave-advance pattern works in a Script Graph instead of C#.",
      steps: [
        "On the kill-handling graph, replace the single \"you win\" branch with a Branch node comparing Wave against TotalWaves - True calls your existing win nodes, False increments Wave and re-triggers your spawn graph.",
      ],
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: ["Open your [YourName]Game project from Lesson 10."],
    level1: [
      "On obj_spawner's Create event (or scr_start_game), add wave = 1; and total_waves = 3;.",
      "In scr_start_game, split the enemy-creation loop out into its own function spawn_wave() (code below) - just the loop, nothing else.",
      "In start_game(), call spawn_wave() at the end instead of spawning directly.",
      "On obj_spawner's Draw GUI event, add a draw_text() call showing the current wave alongside score and lives.",
      "On obj_bullet's Collision event, revise the win-check into an if/else (code below): if this was the last wave, win as before; otherwise, advance the wave and call spawn_wave() again.",
      "Press F5 (Run) and play through - clearing wave 1 or 2 should spawn a fresh wave instead of ending the game; clearing wave 3 should show \"You Win!\".",
    ],
    level1Code: [
      "// scr_start_game (Script resource, add a second function):",
      "function spawn_wave() {",
      "    // ...your existing enemy-creation loop, creating 5 obj_enemy instances...",
      "}",
      "",
      "function start_game() {",
      "    with (obj_enemy) { instance_destroy(); }",
      "    score = 0;",
      "    lives = 3;",
      "    wave = 1;",
      "    enemies_remaining = 5;",
      "    has_won = false;",
      "    has_lost = false;",
      "    game_active = true;",
      "    spawn_wave();",
      "}",
      "",
      "// obj_bullet's Collision event with obj_enemy (revised):",
      "audio_play_sound(sfx_hit, 1, false);",
      "obj_spawner.enemies_remaining -= 1;",
      "obj_spawner.score += 10;",
      "if (obj_spawner.enemies_remaining <= 0) {",
      "    if (obj_spawner.wave >= obj_spawner.total_waves) {",
      "        obj_spawner.has_won = true;",
      "    } else {",
      "        obj_spawner.wave += 1;",
      "        obj_spawner.enemies_remaining = 5;",
      "        obj_spawner.spawn_wave();",
      "    }",
      "}",
      "instance_destroy(other);",
      "instance_destroy(self);",
    ],
    level2: [
      "Play a full 3-wave game: confirm waves 1 and 2 each spawn a fresh set of enemies with no win message, and only wave 3 triggers \"You Win!\".",
      "Let a life run out partway through wave 2 and confirm \"Game Over\" still works correctly, not a wave transition.",
      "Press R after a win or loss and confirm wave resets to 1, along with score and lives.",
    ],
    level3: [
      "Visual: briefly flash a \"Wave 2!\" message in Draw GUI using a short-lived timer variable, separate from the has_won/has_lost text.",
      "Structure: notice spawn_wave() is now called from two moments - start_game() and obj_bullet's Collision event's else branch. The actual enemy-creation loop only exists once.",
    ],
    trouble: [
      "Wave 1 immediately shows \"You Win!\": confirm the if/else checks \"wave >= total_waves\" (not just enemies_remaining <= 0 alone).",
      "New waves never spawn: confirm obj_spawner.spawn_wave() is actually called inside the else branch (with the obj_spawner. prefix, since the code runs from obj_bullet).",
      "Restarting leaves wave at whatever it was: confirm start_game() explicitly sets wave = 1.",
    ],
    resource: "Student Resource List > GameMaker section, or gamemaker.io/en/tutorials",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same wave-advance pattern works with drag-and-drop blocks instead of typed code.",
      steps: [
        "On obj_bullet's Collision event, replace the single win-check block with an \"If Variable\" block comparing wave against total_waves - True runs your existing win blocks, False increments wave, resets enemies_remaining, and calls your spawn_wave script block.",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 10.",
      "Add two new scene variables: Wave (default 1) and TotalWaves (default 3).",
      "Add a new Text object, WaveText, visible, positioned near ScoreText/LivesText.",
    ],
    level1: [
      "In the Setup event group (from Lesson 10), add Wave = 1 alongside the other resets.",
      "In the HUD always-event, add an action changing WaveText's content to show Wave and TotalWaves.",
      "Split your Combat group's win-check into two separate events (code below): one for \"this was the last wave\" (the real win), one for \"more waves remain\" (advance and respawn).",
      "Click Preview and play through - clearing wave 1 or 2 should spawn a fresh wave instead of ending the game; clearing wave 3 should show WinText.",
    ],
    level1Code: [
      "(Combat event group, revised)",
      "Condition: Bullet is in collision with Enemy",
      "Action: Play a sound (sfx_hit.wav)",
      "Action: EnemiesRemaining = EnemiesRemaining - 1",
      "Action: Score = Score + 10",
      "Action: Delete Bullet",
      "Action: Delete Enemy",
      "",
      "(new event - full win)",
      "Condition: EnemiesRemaining <= 0",
      "Condition: Wave >= TotalWaves",
      "Action: GameActive = 0",
      "Action: Show WinText",
      "",
      "(new event - advance wave)",
      "Condition: EnemiesRemaining <= 0",
      "Condition: Wave < TotalWaves",
      "Action: Wave = Wave + 1",
      "Action: EnemiesRemaining = 5",
      "Action: Repeat 5 times -> Create Enemy...",
    ],
    level2: [
      "Play a full 3-wave game: confirm waves 1 and 2 each spawn a fresh set of enemies with no win message, and only wave 3 triggers WinText.",
      "Let Lives run out partway through wave 2 and confirm LoseText still works correctly, not a wave transition.",
      "Press R after a win or loss and confirm Wave resets to 1, along with Score and Lives.",
    ],
    level3: [
      "Visual: briefly show a \"Wave 2!\" style message using a separate Text object with a short Timer-based hide action, distinct from WinText/LoseText.",
      "Structure: notice the \"Repeat 5 times -> Create Enemy\" action now appears in two places (Setup, and the advance-wave event) - for extra practice, try extracting it into a Function (Project > Functions) callable from both, if you're comfortable exploring that feature.",
    ],
    trouble: [
      "Wave 1 immediately shows WinText: confirm the full-win event has BOTH conditions (EnemiesRemaining <= 0 AND Wave >= TotalWaves), not just the first.",
      "New waves never spawn: confirm the advance-wave event's Wave < TotalWaves condition is using the correct comparison direction.",
      "Both the full-win and advance-wave events fire on the same frame: confirm their Wave conditions are opposites (>= vs. <) so only one can ever be true at a time.",
    ],
    resource: "Student Resource List > GDevelop section, or gdevelop.io/tutorials",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: ["Launch TIC-80 and load your cart from Lesson 10.", "Press Esc to reach the code editor."],
    level1: [
      "Add wave=1 and total_waves=3 near the top of your code.",
      "Split your enemy-table setup out of start_game() into its own function spawn_wave() (code below) - just the table-populating loop, nothing else.",
      "In start_game(), set wave=1 and call spawn_wave() at the end instead of populating the table directly.",
      "At the bottom of TIC(), add a wave readout alongside your score/lives prints.",
      "In the bullets loop's hit branch, revise the win-check into an if/else (code below): if this was the last wave, win as before; otherwise, advance the wave and call spawn_wave() again.",
      "Press Esc, type run, press Enter, then play through - clearing wave 1 or 2 should spawn a fresh wave instead of ending the game; clearing wave 3 should show \"You Win!\".",
    ],
    level1Code: [
      "function spawn_wave()",
      " enemies={}",
      " for i=0,4 do",
      "  table.insert(enemies,{ex=20+i*20,ey=10,alive=true})",
      " end",
      "end",
      "",
      "function start_game()",
      " score=0",
      " lives=3",
      " wave=1",
      " enemies_remaining=5",
      " game_active=true",
      " bullets={}",
      " spawn_wave()",
      "end",
      "",
      "start_game()",
      "",
      "-- bullets loop's hit branch (revised):",
      "e.alive=false",
      "enemies_remaining=enemies_remaining-1",
      "score=score+10",
      "if enemies_remaining<=0 then",
      " if wave>=total_waves then",
      "  game_active=false",
      " else",
      "  wave=wave+1",
      "  enemies_remaining=5",
      "  spawn_wave()",
      " end",
      "end",
    ],
    level2: [
      "Play a full 3-wave game: confirm waves 1 and 2 each spawn a fresh set of enemies with no win message, and only wave 3 triggers \"You Win!\".",
      "Let lives run out partway through wave 2 and confirm \"Game Over\" still works correctly, not a wave transition.",
      "Press R after a win or loss and confirm wave resets to 1, along with score and lives.",
    ],
    level3: [
      "Visual: briefly print a \"Wave 2!\" style message using a short countdown variable, distinct from the You Win!/Game Over prints.",
      "Structure: notice spawn_wave() is now called from two moments - start_game() and the bullets loop's else branch. The actual table-populating code only exists once.",
    ],
    trouble: [
      "Wave 1 immediately ends the game: confirm the if/else checks \"wave>=total_waves\" (not just enemies_remaining<=0 alone).",
      "New waves never spawn: confirm spawn_wave() is actually called inside the else branch, and that it replaces the enemies table with a fresh one (enemies={}) each time.",
      "Restarting leaves wave at whatever it was: confirm start_game() explicitly sets wave=1.",
    ],
    resource: "Student Resource List > TIC-80 section, or tic80.com",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 10.",
      "Create two new Variables: wave, default value 1, and total_waves, default value 3.",
      "Create a new Custom Event named SpawnWave.",
    ],
    level1: [
      "Move the enemy-reviving/repositioning actions out of your ResetGame Custom Event and into the new SpawnWave Custom Event instead - just the revive/reposition steps, nothing else.",
      "In ResetGame, add wave=1 to your resets, and replace the old revive/reposition actions with a single Call Custom Event(SpawnWave).",
      "On the enemy Actor Template's \"Collided with actor\" trigger, revise the win-check into an if/else (code below): if this was the last wave, win as before; otherwise, advance the wave, reset enemies_remaining, and call SpawnWave again.",
      "Add a wave readout to your Interface (a third Text/Graphic Field, bound to the wave Variable with Update Real-time ticked) alongside score and lives.",
      "Save, then Play/Test and play through - clearing wave 1 or 2 should revive a fresh set of enemies instead of ending the game; clearing wave 3 should show \"You Win!\".",
    ],
    level1Code: [
      "Custom Event: SpawnWave ->",
      "  Actor – Destroy/Revive Actor(enemy_1, Revive) -> Actor – Teleport(enemy_1, [starting position])",
      "  (repeat for enemy_2, enemy_3, enemy_4, enemy_5)",
      "",
      "Custom Event: ResetGame (revised) ->",
      "  Variable – Modify Value(score, Set, 0)",
      "  Variable – Modify Value(lives, Set, 3)",
      "  Variable – Modify Value(wave, Set, 1)",
      "  Variable – Modify Value(enemies_remaining, Set, 5)",
      "  Variable – Modify Value(game_active, Set, 1)",
      "  Call Custom Event(SpawnWave)",
      "",
      "enemy Actor Trigger: Collided with actor (revised) ->",
      "  If(game_active = 1)",
      "    Variable – Modify Value(score, Add, 10)",
      "    Variable – Modify Value(enemies_remaining, Subtract, 1)",
      "    If(enemies_remaining <= 0)",
      "      If(wave >= total_waves)",
      "        Variable – Modify Value(game_active, Set, 0)",
      "        Message Box(\"You Win! - Press R to Restart\")",
      "      Else",
      "        Variable – Modify Value(wave, Add, 1)",
      "        Variable – Modify Value(enemies_remaining, Set, 5)",
      "        Call Custom Event(SpawnWave)",
      "    Actor – Destroy/Revive Actor(relate, Destroy)",
      "    Actor – Destroy/Revive Actor(this, Destroy)",
    ],
    level2: [
      "Play a full 3-wave game: confirm waves 1 and 2 each revive a fresh set of enemies with no win message, and only wave 3 triggers \"You Win!\".",
      "Let lives run out partway through wave 2 and confirm \"Game Over\" still works correctly, not a wave transition.",
      "Press R after a win or loss and confirm wave resets to 1, along with score and lives.",
    ],
    level3: [
      "Visual: briefly show a \"Wave 2!\" Message Box (or a dedicated Interface Field, shown for a moment) distinct from the win/lose Message Box.",
      "Structure: notice SpawnWave is now called from two moments - ResetGame, and the win-check's Else branch. The actual revive/reposition actions only exist once.",
    ],
    trouble: [
      "Wave 1 immediately shows \"You Win!\": confirm the nested If checks wave against total_waves, not just enemies_remaining alone.",
      "New waves never revive: confirm Call Custom Event(SpawnWave) is actually inside the Else branch, and that SpawnWave's actor names match your actual enemy copy names exactly.",
      "Restarting leaves wave at whatever it was: confirm ResetGame explicitly sets wave to 1.",
    ],
    resource: "001gamecreator.com/wiki/customevents/, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 10.",
      "In the Level Blueprint, add two new Integer variables: Wave, default value 1, and TotalWaves, default value 3.",
      "In WBP_HUD's Designer view, add a fourth Text Block: WaveText, positioned near ScoreText/LivesText.",
    ],
    level1: [
      "Extract the spawn For Loop out of your ResetGame Custom Event into its own new Custom Event SpawnWave (code below) - just the loop, nothing else.",
      "In ResetGame, add Wave = 1 to your resets, and call SpawnWave at the end instead of spawning directly.",
      "In WBP_HUD's UpdateHUD, add a Wave parameter and a Set Text call for WaveText.",
      "In HandleEnemyDestroyed, revise the win-check into a nested Branch (code below): if this was the last wave, win as before; otherwise, advance the wave and call SpawnWave again.",
      "Click Compile, then Play, and play through - clearing wave 1 or 2 should spawn a fresh wave instead of ending the game; clearing wave 3 should show the win message.",
    ],
    level1Code: [
      "Custom Event SpawnWave ->",
      "  For Loop (0 to 4) -> Spawn Actor from Class(BP_Enemy) -> Bind Event to OnEnemyDestroyed / OnEnemyEscaped",
      "",
      "Custom Event ResetGame (revised) ->",
      "  For Each (Get All Actors of Class(BP_Enemy)) -> Destroy Actor",
      "  Score = 0",
      "  Lives = 3",
      "  Wave = 1",
      "  EnemiesRemaining = 5",
      "  GameActive = true",
      "  Call SetStatus(HUDWidget, \"\")",
      "  Call UpdateHUD(HUDWidget, Score, Lives, Wave)",
      "  Call SpawnWave",
      "",
      "Custom Event HandleEnemyDestroyed (revised) ->",
      "  EnemiesRemaining = EnemiesRemaining - 1",
      "  Score = Score + 10",
      "  Call UpdateHUD(HUDWidget, Score, Lives, Wave)",
      "  Branch(EnemiesRemaining <= 0) -> True:",
      "    Branch(Wave >= TotalWaves) -> True:",
      "      GameActive = false",
      "      Call SetStatus(HUDWidget, \"You Win! Press R to Restart\")",
      "    False:",
      "      Wave = Wave + 1",
      "      EnemiesRemaining = 5",
      "      Call UpdateHUD(HUDWidget, Score, Lives, Wave)",
      "      Call SpawnWave",
    ],
    level2: [
      "Play a full 3-wave game: confirm waves 1 and 2 each spawn a fresh set of enemies with no win message, and only wave 3 triggers the win message.",
      "Let Lives run out partway through wave 2 and confirm the lose message still works correctly, not a wave transition.",
      "Press R after a win or loss and confirm Wave resets to 1, along with Score and Lives.",
    ],
    level3: [
      "Visual: briefly show a \"Wave 2!\" message via SetStatus with a Delay node clearing it shortly after, distinct from the final win/lose message.",
      "Structure: notice Call SpawnWave now appears in two places (ResetGame, and HandleEnemyDestroyed's False branch) - the actual spawn loop only exists once, in the SpawnWave Custom Event itself.",
    ],
    trouble: [
      "Wave 1 immediately shows the win message: confirm the nested Branch checks Wave >= TotalWaves, not just EnemiesRemaining <= 0 alone.",
      "New waves never spawn: confirm Call SpawnWave is actually wired inside the False case of the nested Branch.",
      "Restarting leaves Wave at whatever it was: confirm ResetGame explicitly sets Wave = 1.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same nested win-check pattern works in C++ instead of Blueprint nodes.",
      steps: [
        "Extract your enemy-spawning loop into a SpawnWave() function, called from both ResetGame() and the else-branch of your kill-handling function's nested if.",
        "In the kill-handling function: if (EnemiesRemaining <= 0) { if (Wave >= TotalWaves) { bGameActive = false; } else { Wave++; EnemiesRemaining = 5; SpawnWave(); } }",
      ],
      note: "This mirrors the Blueprint path's nested Branch exactly - same shape, just from code.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_11_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 11 Teacher's Key: Wave After Wave", bold: true, size: 30, color: brand.NAVY })] }),
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
      headers: { default: brand.pageHeader("Lesson 11 Key - " + data.engine + " (Instructor Only)") },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    teaches: "Splitting a single-condition check (Lesson 6's 'are all enemies gone?') into a two-condition decision ('gone, AND is this the last wave?') - the first time this course asks students to nest a decision inside another decision.",
    level1: [
      "wave and total_waves are two new 'for all sprites' variables, exactly parallel to lives and score in how they're declared and scoped.",
      "The spawn loop itself moves, unchanged, into its own 'when I receive [spawn wave]' script - a direct continuation of Lesson 10's 'when I receive [start game]' refactor, one layer more granular.",
      "The if/else nested inside the existing 'if enemies_left = 0' check is the actual new idea this week: same outer trigger as Lesson 6, a new inner decision.",
    ],
    why: [
      "Reusing the broadcast mechanism (now three broadcasts: start game, spawn wave, you win) keeps this lesson consistent with everything Scratch has done since Lesson 6 - no new Scratch concept, just a new broadcast and a nested if.",
      "Keeping spawn_wave/'spawn wave' as its own broadcast (rather than copy-pasting the loop into both 'start game' and the win-check) is exactly the reuse principle from Lesson 10, applied one level deeper this week.",
    ],
    talkingPoints: [
      "Ask students to trace what happens on wave 2's clear, step by step: enemies_left hits 0, wave (currently 1) is compared to total_waves (3), they're not equal, so wave becomes 2 and 'spawn wave' broadcasts again - a good comprehension check for the nested if.",
      "For strong students: ask what would need to change to make each wave harder (more enemies, faster movement) - they should recognize this is explicitly NOT this lesson's job (that's Lesson 14), but the wave counter built today is exactly what a difficulty-scaling system would key off of.",
    ],
    mistakes: [
      "Checking only 'enemies_left = 0' without the wave comparison, causing wave 1's clear to immediately trigger the win.",
      "Broadcasting 'spawn wave' from outside the else branch (e.g., after the if/else instead of inside it), causing an extra, unwanted wave spawn even on the winning wave.",
      "Forgetting to add wave=1 to the 'start game' script, leaving wave stuck wherever it was after a restart.",
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    teaches: "Extracting spawn_wave() as a sibling of start_game() - two functions with a clear relationship (start_game() calls spawn_wave(), but spawn_wave() is also called independently) - a natural next step from Lesson 10's single-function refactor.",
    level1: [
      "spawn_wave() is purely the enemy-creation loop, with no knowledge of score, lives, or win/lose state - it does exactly one job.",
      "start_game() is now a thin wrapper: clear everything, reset all state including wave = 1, then delegate the actual spawning to spawn_wave().",
      "_on_enemy_destroyed()'s nested if is where the real branching logic lives - this function now makes a genuinely two-part decision instead of last week's one-part decision.",
    ],
    why: [
      "This mirrors real software structure: small, single-purpose functions (spawn_wave()) composed by a slightly higher-level one (start_game()), which is itself called from an even higher-level trigger (_ready(), or the restart key).",
      "Keeping the win/lose Labels reserved for the ROUND's actual end (not wave transitions) maintains the same 'don't reach for the big tool before you need it' discipline from Lesson 6 - a mid-run wave transition doesn't need the same weight as a true win or loss.",
    ],
    talkingPoints: [
      "Ask students to draw (on paper or verbally describe) the call graph: _ready() -> start_game() -> spawn_wave(), and separately _on_enemy_destroyed() -> spawn_wave() directly. Two different paths reach the same function.",
      "For students who want a stretch goal: ask how they'd store per-wave enemy counts (e.g., an array like [5, 7, 9]) instead of a flat 5 every wave - not required this week, but a natural preview of Lesson 14's difficulty scaling.",
    ],
    mistakes: [
      "Checking only 'enemies_remaining <= 0' without the wave comparison.",
      "Calling spawn_wave() from inside the True branch (the win case) by mistake, spawning an unwanted extra wave right as the game is won.",
      "Forgetting to add newly-spawned enemies to the \"enemies\" group inside spawn_wave() specifically (easy to miss if spawn_wave() was written as a fresh function rather than copy-adjusted from start_game()'s old code).",
    ],
  },
  {
    engine: "Unity", file: "Unity",
    teaches: "Reusing EnemySpawner.SpawnEnemies() from TWO different call sites within GameManager's own logic (ResetGame() and RegisterKill()'s else branch) - the same reuse principle from Lesson 10, now triggered by gameplay itself rather than only by player input.",
    level1: [
      "wave and totalWaves are new static fields on GameManager, parallel to score and lives.",
      "SpawnEnemies() itself needs zero changes this week - it was already a clean, single-purpose method as of Lesson 10, which is exactly why it was trivial to call again from a second place.",
      "RegisterKill()'s nested if is where the real branching logic lives - GameManager now makes a genuinely two-part decision instead of last week's one-part decision.",
    ],
    why: [
      "This is the clearest payoff yet of Lesson 10's EnemySpawner.instance singleton pattern: because SpawnEnemies() was already reachable from anywhere via a static instance reference, calling it again from RegisterKill() required no new plumbing.",
      "Keeping the win/lose OnGUI() messages reserved for the round's true end (not wave transitions) continues the same restraint from Lesson 6 and Lesson 9 - not every state change deserves the same visual weight.",
    ],
    talkingPoints: [
      "Ask students to name every place SpawnEnemies() is now called from (Start(), ResetGame(), RegisterKill()'s else branch) - three call sites, one implementation.",
      "For students comparing engines: Unity's EnemySpawner.instance.SpawnEnemies() and Godot's spawn_wave() are conceptually identical this week - both are 'the one place enemies get created,' called from multiple locations.",
    ],
    mistakes: [
      "Checking only 'enemiesRemaining <= 0' without the wave comparison.",
      "Calling SpawnEnemies() inside the True (win) branch by mistake.",
      "Forgetting wave = 1 inside ResetGame(), leaving it at whatever value it had before a restart.",
    ],
  },
  {
    engine: "GameMaker", file: "GameMaker",
    teaches: "Adding a second function (spawn_wave()) to the same Script resource introduced in Lesson 10 - GameMaker's Scripts can hold multiple related functions, and this is the first time this course uses that.",
    level1: [
      "spawn_wave() and start_game() now live in the same scr_start_game Script resource, a natural place for two closely related functions to live together.",
      "start_game() delegates its actual spawning work to spawn_wave() - the same 'thin wrapper calling a focused helper' shape as Godot's version, just in GML.",
      "obj_bullet's Collision event grows a nested if - the same two-part decision every other engine builds this week, expressed in GML's if/else syntax.",
    ],
    why: [
      "GameMaker Scripts aren't limited to one function each - grouping spawn_wave() alongside start_game() in the same resource keeps closely-related reset/setup logic organized together, rather than scattered across many small Script resources.",
      "The obj_spawner. prefix throughout (from obj_bullet's Collision event) is worth re-emphasizing this week specifically, since students are now writing a more complex nested condition and it's easy to drop a prefix somewhere in the middle.",
    ],
    talkingPoints: [
      "Ask students to find every place spawn_wave() is called from within scr_start_game and obj_bullet - two call sites, one function.",
      "For advanced students: ask what would happen if total_waves were increased to, say, 5, without touching any other code - the game should just work, needing more clears to win. This is a good test of whether the structure is genuinely decoupled from the specific number 3.",
    ],
    mistakes: [
      "Checking only 'enemies_remaining <= 0' without the wave comparison.",
      "Calling spawn_wave() inside the win branch by mistake.",
      "Dropping the obj_spawner. prefix somewhere inside the new nested if, causing a variable-scope error or silently referencing the wrong instance's variable.",
    ],
  },
  {
    engine: "GDevelop", file: "GDevelop",
    teaches: "Splitting one event (Lesson 6's single win-check) into two mutually exclusive events (full-win vs. advance-wave), each guarded by an opposite comparison on the same Wave variable - GDevelop's event-based way of expressing an if/else.",
    level1: [
      "Wave and TotalWaves are new scene variables, parallel to Score and Lives.",
      "The full-win event and the advance-wave event share the same first condition (EnemiesRemaining <= 0) but have opposite second conditions (Wave >= TotalWaves vs. Wave < TotalWaves) - together, they cover every possibility exactly once, with no overlap.",
      "The advance-wave event's spawn action is a direct copy of the Setup group's spawn action - GDevelop doesn't require extracting a callable function for this to work, though the optional Level 3 Functions idea shows students how they could.",
    ],
    why: [
      "GDevelop's event-condition model expresses 'if/else' as two separate, oppositely-guarded events rather than a single event with a branching action list - a genuinely different (and equally valid) way of achieving the same logical structure other engines build with code.",
      "Keeping WinText/LoseText reserved for the round's true end (not wave transitions) continues the established discipline from Lesson 6 and Lesson 9.",
    ],
    talkingPoints: [
      "Ask students to explain, in their own words, why the two new events' Wave conditions need to be exact opposites (>= vs. <) - if they overlapped or left a gap, some wave-clear moments could trigger both, neither, or the wrong one.",
      "For students ready for a stretch: point out the Level 3 idea about extracting the spawn action into a genuine GDevelop Function (Project > Functions) - not required, but a natural bridge toward the 'named reusable routine' pattern every other engine uses this week.",
    ],
    mistakes: [
      "Only one of the two new events' Wave conditions actually added (e.g., forgetting Wave < TotalWaves on the advance event) - can cause both events to fire on the same frame, or neither.",
      "The advance-wave event's spawn action missing or referencing the wrong count.",
      "Forgetting Wave = 1 in the Setup group, leaving it unreset after a restart.",
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    teaches: "The plainest possible version of this week's idea: one more function (spawn_wave()), one more nested if - built from the exact same three Lua primitives (functions, variables, conditionals) TIC-80 has used all along.",
    level1: [
      "spawn_wave() replaces the table-populating portion of start_game() - a clean split between 'reset everything' (start_game) and 'create a wave's worth of enemies' (spawn_wave).",
      "wave and total_waves are new globals, parallel to lives and score.",
      "The bullets loop's hit branch grows a nested if - the same two-part decision every other engine builds this week.",
    ],
    why: [
      "TIC-80 has no special 'wave system' feature to reach for - this lesson, like Lesson 10's restart flow, is built entirely from ordinary Lua tools, which is a good, honest illustration of how far a few simple primitives can go.",
      "Replacing the enemies table with a fresh one (enemies={}) inside spawn_wave(), rather than trying to modify the old (now-empty) one in place, avoids any risk of leftover stale entries between waves.",
    ],
    talkingPoints: [
      "Ask students to trace what 'enemies={}' actually does each time spawn_wave() runs - it doesn't just clear old entries, it replaces the whole table with a brand new empty one, which is a subtly different (and safer) operation worth understanding.",
      "For advanced students: point out that hard-coding 5 enemies per wave is a deliberate simplification - ask them to imagine a table like waves={5,7,9} indexed by the current wave, as a preview of Lesson 14's difficulty scaling.",
    ],
    mistakes: [
      "Checking only 'enemies_remaining<=0' without the wave comparison.",
      "Calling spawn_wave() inside the win branch (game_active=false case) by mistake.",
      "Forgetting wave=1 inside start_game(), leaving it at whatever value it had before a restart.",
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    teaches: "Splitting the existing ResetGame Custom Event into two purpose-specific Custom Events (ResetGame for full resets, SpawnWave for just reviving enemies) - directly extending last week's introduction of Custom Events as this engine's reusable-subroutine feature.",
    level1: [
      "SpawnWave contains exactly the revive/reposition actions that used to live inside ResetGame - moved, not duplicated.",
      "ResetGame is now a thin wrapper: reset all Variables including the new wave, then Call Custom Event(SpawnWave) to handle the actual enemy revival.",
      "The 'Collided with actor' trigger's nested If is where the real branching logic lives - a genuinely two-part decision this week, verified against the same wiki-confirmed Custom Event mechanism used last week.",
    ],
    why: [
      "This is the clearest payoff yet of introducing Custom Events last week: because the revive/reposition logic was already isolated into its own callable block, splitting it out further and calling it from a second location (the win-check's Else branch) required no new mechanism, just a second Call Custom Event action.",
      "Keeping the win Message Box reserved for the true, final win (not wave transitions) continues the same discipline from Lesson 6, applied consistently across every lesson since.",
    ],
    talkingPoints: [
      "Ask students to name every place Call Custom Event(SpawnWave) now appears (ResetGame, and the win-check's Else branch) - two call sites, one Custom Event.",
      "This is a good lesson to revisit the Lesson 5 design-time-placement decision one more time: because there are only ever exactly 5 named enemy copies (never created or destroyed at runtime, only revived and repositioned), 'a new wave' in this engine literally means the same 5 actors coming back to life - conceptually identical to every other engine's fresh spawn, just implemented differently.",
    ],
    mistakes: [
      "Checking only 'enemies_remaining <= 0' without the nested wave check.",
      "Calling SpawnWave inside the win branch by mistake, causing enemies to revive right as the game is won.",
      "Forgetting wave=1 inside ResetGame, leaving it at whatever value it had before a restart.",
    ],
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    teaches: "Extracting SpawnWave as its own Custom Event, called from both ResetGame and HandleEnemyDestroyed's nested Branch - continuing this course's Blueprint-native reuse pattern (UpdateHUD, SetStatus, ResetGame) into a fourth, closely related Custom Event.",
    level1: [
      "SpawnWave is purely the enemy-spawning For Loop, extracted unchanged from what used to live inline inside ResetGame.",
      "ResetGame is now a thin wrapper: clear existing enemies, reset all variables including the new Wave, update the HUD, then Call SpawnWave.",
      "HandleEnemyDestroyed's nested Branch is where the real decision-making lives - Unreal's Blueprint equivalent of every other engine's if/else this week.",
    ],
    why: [
      "This mirrors Godot's spawn_wave() and Unity's SpawnEnemies() almost exactly, just expressed as a fourth Blueprint Custom Event alongside UpdateHUD, SetStatus, and ResetGame - worth naming this cross-engine parallel explicitly.",
      "Keeping the win/lose SetStatus messages reserved for the round's true end (not wave transitions) continues the same discipline established since Lesson 6.",
    ],
    talkingPoints: [
      "For students who've built several Custom Events by now: ask them to sketch (verbally or on paper) the full call graph - BeginPlay -> ResetGame -> SpawnWave, and separately HandleEnemyDestroyed -> SpawnWave. A good comprehension check for how many Custom Events this project now has and how they relate.",
      "This is a good moment to ask advanced students what would need to change to make SpawnWave create a DIFFERENT number of enemies depending on Wave - not required this week (that's Lesson 14), but the nested Branch built today is exactly the hook a difficulty system would extend.",
    ],
    mistakes: [
      "Checking only EnemiesRemaining <= 0 without the nested Wave Branch.",
      "Calling SpawnWave from inside the True (win) case by mistake.",
      "Forgetting Wave = 1 inside ResetGame, leaving it at whatever value it had before a restart.",
    ],
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_11_Key_${d.file}.docx`]));

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
