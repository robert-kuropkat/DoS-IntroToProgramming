const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign
, ImageRun } = require("docx");
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
function numbered(text, level = 0, instance = 0) { return new Paragraph({ numbering: { reference: "numbers", level, instance }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
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
      headers: { default: brand.pageHeader("Lesson 10 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 10 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("This Week: Closing the Loop"),
        p("Since Lesson 8, winning or losing has just meant a message appearing while the game kept quietly running underneath. This week finally closes that loop: the game actually stops when it's over, and pressing R brings it back to a genuinely fresh start - not a reload, not a re-launch, an in-place reset built entirely from code students already trust."),
        p("The central idea this week is reuse, one more time: instead of writing separate 'set up a new game' code and 'reset the game' code, every engine's design this week factors the ORIGINAL one-time setup (that's been running since Lesson 5) into a single reusable routine, called once at first load AND again every time R is pressed. This is the clearest, most direct example yet of the reuse principle threaded through this whole course."),
        checkbox("Read through each engine's Handout once before class - every engine involves a genuine refactor this week (moving existing setup code into a new reusable form), not just new additions."),

        h2("A game_active Flag, Introduced This Week"),
        p("Every engine gets a new boolean-ish flag (game_active, GameActive, or similar) that gates whether gameplay logic still applies. Once a round ends (win or lose), this flag flips off - existing win/lose-check code stops changing score or lives further, and (where applicable) enemy movement or player input freezes. Pressing R while the flag is off calls the reset routine, which flips it back on as its last step."),

        h2("001 Game Creator: Verified Mechanisms"),
        p("This week's 001 Game Creator design was researched and verified against the official wiki before writing: the restart key uses an Input Trigger (the same underlying mechanism already used since Lesson 2 for movement and shooting), and the reusable reset routine uses a Custom Event (a genuine, documented reusable-subroutine feature - 001gamecreator.com/wiki/customevents/). A built-in \"Restart Game\" action does exist in the engine, but this lesson deliberately teaches the manual, verified reset approach instead, since it's better pedagogy and its exact reset behavior isn't documented clearly enough to teach with confidence."),
        checkbox("If you have Windows access before class, walk through creating one Custom Event and one Input Trigger yourself - both are new concepts for this course, even for an instructor comfortable with 001 Game Creator's other patterns."),

        h2("TIC-80: Verified Key Code"),
        p("TIC-80's restart key detection uses keyp(18) - confirmed directly against the engine's own GitHub wiki key code table (keycode 18 = R) and the keyp() function's documented edge-trigger behavior (fires true only on the frame a key is first pressed, not while held). This was verified this week given this engine's history of needing extra care with exact API details."),

        h2("Self-Contained Lesson Assets"),
        p("No new assets this week - restart flow is entirely code/logic."),

        h2("Continuity From Lesson 9"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 9, not create a new one."),
        checkbox("If a student's Lesson 9 project is lost or broken, have a pre-built \"Lesson 9 complete\" starter available per engine (from the Lesson 9 Teacher's Keys) so they aren't blocked."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop's Lesson 6-9 exercises still work for all 8 engines."),
        checkbox("Test a FULL round-trip on every engine before class: play to a win, press R, confirm a clean restart; separately play to a loss, press R, confirm a clean restart. Both directions matter."),
        checkbox("001 Game Creator (Windows demo laptop): walk through the Custom Event + Input Trigger combination once, and specifically confirm the 5 pre-placed enemy copies revive and reposition correctly on reset."),
        checkbox("Godot / Unity / GameMaker / GDevelop / Unreal: confirm any enemies still on screen at game-over are properly cleared before the fresh set spawns on restart - a common thing to miss when first building this."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies."),

        h2("Known Gotchas This Week"),
        bullet("The most likely mistake across every engine: forgetting to clear/destroy leftover enemies before respawning on restart, resulting in double the enemies (or a wrong enemies_remaining count) after the first restart."),
        bullet("A close second: pressing R while the game is still active accidentally triggers a restart, because the game_active check was left out of the restart-key logic. Test this explicitly - R should do nothing mid-game."),
        bullet("Score/lives sometimes still change after game-over because the win/lose-check code wasn't actually wrapped in a game_active check - only the display was frozen, not the underlying logic. This is a subtle miss worth testing for directly, not just eyeballing the screen."),
        bullet("This is a satisfying lesson to finish - encourage a full playthrough demo (win, restart, lose, restart) as part of the end-of-lesson share-out, since it's the first time the game truly feels replayable."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_10_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function diagram(file, widthIn, heightIn, caption) {
  const data = fs.readFileSync(file);
  const kids = [new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: caption ? 40 : 160 },
    children: [new ImageRun({ data, transformation: { width: widthIn * 96, height: heightIn * 96 }, type: "png" })] })];
  if (caption) kids.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [new TextRun({ text: caption, italics: true, size: 16, color: "888888" })] }));
  return kids;
}
const ACTIVITY = ["diagram_activity_restart_flow.png", 5.4, 5.45, "When the round ends, everything freezes - and R, pressed only while the round is over, runs the one reset routine. The same shape in every engine."];
const ENGINE_DIAGRAMS = {
  GameMaker: { file: "diagram_gml_visual_l10_level1.png", w: 5.0, h: 6.36, caption: "scr_start_game, obj_gamemanager's Step event, and obj_enemy's Step event - the block-for-block equivalent of the Level 1 code." },
  Unreal: { file: "diagram_unreal_blueprint_l10_level1.png", w: 5.6, h: 6.16, caption: "BP_GameManager's ResetGame, BeginPlay and IA_Restart event, plus the Branch at the top of BP_Enemy's Tick - the node-for-node equivalent of the Level 1 code." },
};
const STATE_AUDIT = {
 "Scratch": [
  "Find EVERY \"when green flag clicked\" script: on the Stage, on every sprite, and on the GameManager. Click each sprite in the sprite list and check. Any that set up your game (position, costume, size, show/hide, variables, backdrop) must also run on \"when I receive [start game]\", or be moved there.",
  "Look at your game-over scripts (\"when I receive [you win]\" and \"when I receive [game over]\") and anything else you added that runs when the round ends: animations, costume changes, color or ghost effects, size changes, sounds, moved or hidden sprites, backdrop switches.",
  "For each of those, add the opposite to \"when I receive [start game]\": switch back to the starting costume, \"clear graphic effects\", set size, \"go to\" the starting position, show or hide, stop sounds.",
  "Check clones and timers too: the Enemy sprite's \"delete this clone\" at the start of \"when I receive [start game]\" clears the enemy clones, but anything else you created (extra clones, \"reset timer\") needs its own reset."
 ],
 "Godot": [
  "Find EVERY _ready() in every script (Main, GameManager, Player, Enemy, and any extra node you added). _ready() runs only once, when a node first loads, so anything it sets up will NOT run again on restart. Move the setup that must repeat into start_game(), or call a reset function from start_game().",
  "Look at everything that happens at game over: Tweens, AnimationPlayer animations, modulate or color changes, hidden or shown nodes, sounds, timers you started.",
  "For each, add the opposite to start_game(): stop and reset the animation, set modulate back to white, show or hide the node, stop the sound, reset or stop the timer, and put the player back at its starting position.",
  "Remember state lives on nodes other than GameManager too (the player's position, an Enemy's speed). If you changed it during play or at game over, reset it."
 ],
 "GameMaker": [
  "Find EVERY object's Create event. A Create event runs only when that instance is first created, so objects that already exist when the game restarts (your player, obj_gamemanager) will NOT run it again. Anything it sets must also be set in start_game(), or in a reset you call from it.",
  "Look at everything that happens at game over: image_blend or image_alpha changes, sprite_index or image_speed changes, an animation or Alarm you started, draw calls, sounds, hidden (visible = false) or moved instances.",
  "For each, add the opposite to start_game() (or the Visual blocks it calls): restore the sprite, image_blend, image_alpha and image_speed, stop the Alarm, stop the sound, set visible back, and put the player at its starting x and y. In Visual, use Set Variable / Set Sprite style blocks with the Target set to the right object.",
  "Check Room Start events and Alarm events as well as Create events: anything that runs only once at launch needs a reset on restart."
 ],
 "Unity": [
  "Find EVERY Awake() and Start() in every script. Both run only once, when the object first loads, so they will NOT run again on restart. Anything they set up must also be reset by ResetGame().",
  "Remember static fields (like GameManager's score and lives) keep their value across restarts until you set them. Instance fields on other scripts (the player's speed, an enemy counter) are also state.",
  "Look at everything that happens at game over: Animator states or triggers, color or sprite changes, Time.timeScale, coroutines you started, particle effects, AudioSources, objects you turned off with SetActive(false), UI text.",
  "For each, add the opposite to ResetGame(): reset the Animator, restore the color and sprite, set Time.timeScale back to 1, stop coroutines, stop the particles and audio, SetActive(true) where needed, and put the player back at its starting position."
 ]
};
const STATE_AUDIT_GENERIC = ["Find every place your game does something once at start-up, and everything that changes when the round ends (animations, effects, sounds, hidden or moved objects). Everything that changed must be put back by your reset routine."];
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 10: Play It Again", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: when the round ends, the game actually stops - and pressing R brings it back to a genuinely fresh start.", { bold: true }),
    p("Open your own [YourName]Game project from Lesson 9 - don't start a new one.", { bold: true }),
    p("Tip: no new assets this week - this is a logic lesson, and it reuses your very first setup code from Lesson 5.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s, 0, 0)),
    h2("Level 1 - Minimal (MVP)"),
    p("When the round ends (win or lose), the game freezes - score and lives stop changing, and further enemy hits/escapes do nothing. Pressing R resets everything and starts a new round.", { italics: true }),
    ...diagram(...ACTIVITY),
    ...data.level1.map(s => numbered(s, 0, 1)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  if (data.file === "Unreal") children.push(...diagram(ENGINE_DIAGRAMS.Unreal.file, ENGINE_DIAGRAMS.Unreal.w, ENGINE_DIAGRAMS.Unreal.h, ENGINE_DIAGRAMS.Unreal.caption));
  if (data.altPath) {
    children.push(h2("Alternate Path: " + data.altPath.title));
    children.push(p(data.altPath.intro, { italics: true, color: "555555" }));
    children.push(...data.altPath.steps.map(s => numbered(s, 0, 2)));
    if (data.altPath.code) children.push(code(data.altPath.code));
    if (data.file === "GameMaker") children.push(...diagram(ENGINE_DIAGRAMS.GameMaker.file, ENGINE_DIAGRAMS.GameMaker.w, ENGINE_DIAGRAMS.GameMaker.h, ENGINE_DIAGRAMS.GameMaker.caption));
    if (data.altPath.note) children.push(p(data.altPath.note, { italics: true, color: "555555" }));
  }
  children.push(
    h2("Before You Test: Audit Your Game State"),
    p("Game state is everything your game remembers right now: score, lives, positions, costumes and colors, animations, sounds, and messages on screen. Until this lesson, your game started once and ended once, so nothing ever had to be put back. A restart means ALL of it must go back to how it was at the start - including anything extra you added that this handout doesn't mention.", { italics: true }),
    ...(STATE_AUDIT[data.file] || STATE_AUDIT_GENERIC).map(s => numbered(s, 0, 3)),
    p("Rule of thumb: if it changed when the round ended (or at any point during play), something has to change it back when the round restarts.", { bold: true }),
    h2("Level 2 - Complete (Alpha/Beta)"),
    p("A restart is a genuinely clean slate - no leftover enemies, no stale score/lives, and it works correctly no matter how many times in a row you play.", { italics: true }),
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
      headers: { default: brand.pageHeader("Lesson 10 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 9.",
      "Create a new variable game_active (for all sprites, not \"for this sprite only\"). Scratch has no true/false type, so game_active just holds the words true or false. Type them in lowercase, exactly, every time.",
    ],
    level1: [
      "On GameManager, change your existing \"when green flag clicked\" script to just: hide, then \"broadcast [start game]\" - a thin trigger instead of the full setup. GameManager already owns score, lives, and enemies_remaining, so the reset belongs here too.",
      "On GameManager, add a new script \"when I receive [start game]\" containing the setup the green-flag script used to do: set score to 0, set lives to 3, set enemies_remaining to 5, set game_active to true.",
      "On the Enemy sprite, change your existing spawn script so it starts with \"when I receive [start game]\" instead of the green flag, followed by your existing spawn loop (repeat 5: create clone of myself, and so on).",
      "Still on the Enemy sprite, add a SECOND, separate script: \"when I receive [start game]\" with a single \"delete this clone\" block under it. Clones hear broadcasts too, so on a restart every old clone deletes itself, while the original Enemy sprite (where \"delete this clone\" does nothing) keeps its spawn script and makes the fresh set. (\"delete this clone\" is an end block - nothing can connect under it - which is why it needs its own script.)",
      "In the Enemy's \"when I start as a clone\" forever loop, add an \"if game_active = false then\" block at the very top, with a \"stop this script\" block inside it (code below). When the round ends each clone stops running its movement, hit and escape checks and freezes where it is; the fresh round deletes it.",
      "On GameManager's \"when I receive [enemy hit]\" script, add \"set game_active to false\" right before \"broadcast [you win]\". On its \"when I receive [enemy escaped]\" script, add \"set game_active to false\" right before \"broadcast [game over]\". The Enemy sprite itself never touches game_active - it still only reports what happened.",
      "Update the Player's win/lose \"say\" blocks to include the restart hint: \"say [You Win! Press R to Restart] for 3 seconds\" and \"say [Game Over - Press R to Restart] for 3 seconds\".",
      "On GameManager, add a new script: \"when key r pressed -> if game_active = false then broadcast [start game]\".",
      "Click the stage, play to a win or loss, confirm the game freezes, then press R and confirm everything resets and play resumes.",
    ],
    level1Code: [
      "// GameManager (hidden sprite):",
      "when green flag clicked",
      "hide",
      "@@ broadcast [start game]",
      "",
      "@@ when I receive [start game]",
      "@@ set score to 0",
      "@@ set lives to 3",
      "@@ set enemies_remaining to 5",
      "@@ set game_active to true",
      "",
      "when I receive [enemy hit]",
      "change enemies_remaining by -1",
      "change score by 10",
      "if enemies_remaining = 0 then",
      "@@     set game_active to false",
      "    broadcast [you win]",
      "",
      "when I receive [enemy escaped]",
      "change lives by -1",
      "if lives = 0 then",
      "@@     set game_active to false",
      "    broadcast [game over]",
      "",
      "@@ when key r pressed",
      "@@ if game_active = false then",
      "@@     broadcast [start game]",
      "",
      "// Enemy sprite:",
      "@@ when I receive [start game]",
      "repeat 5",
      "    create clone of myself",
      "    // (your existing position and wait blocks stay here)",
      "",
      "@@ when I receive [start game]",
      "@@ delete this clone",
      "",
      "when I start as a clone",
      "forever",
      "@@     if game_active = false then",
      "@@         stop this script",
      "    // (your existing movement blocks stay here)",
      "    if touching Bullet? then",
      "        // (your existing hit logic stays here)",
      "    if y position < -170 then",
      "        // (your existing escape logic stays here)",
      "",
      "// Player sprite:",
      "when I receive [you win]",
      "@@ say [You Win! Press R to Restart] for 3 seconds",
      "when I receive [game over]",
      "@@ say [Game Over - Press R to Restart] for 3 seconds",
    ],
    level2: [
      "Play to a win or loss, confirm the enemies freeze where they are, then press R and confirm score/lives/enemies_remaining are exactly back to 0/3/5 with no leftover clones and a fresh set of five.",
      "Repeat for a loss - same clean reset.",
      "Press R while the game is still active (mid-round) and confirm NOTHING happens - the game_active check in your restart script is what prevents an accidental reset.",
    ],
    level3: [
      "Fix a real bug! We laid the groundwork for this bug all the way back in Lesson 6 when we defined our win scenario. Then in Lesson 8 we defined our lose scenario. By thinking in terms of win and lose separately, rather than in terms of wave over and game over, we created a bug. Find it and resolve it in the manner you, the Designer, think best. (Hint: play a round where you destroy some enemies and let the others escape, and watch what happens to the round, and to your R key.)",
      "Visual: add a brief visual cue when the game freezes, like a costume change or a \"change color effect\" pulse on the Player, separate from the win/lose message itself.",
      "Structure: notice \"when I receive [start game]\" is now the single place that sets up a fresh game - used both at the very first green-flag click and every restart. No setup code is duplicated anywhere in your project.",
    ],
    trouble: [
      "Restart does nothing when you press R: confirm the check is \"if game_active = false\" (not = true), and that game_active is actually being set to false on GameManager at the moment of the win or loss.",
      "Restart leaves old enemies on screen or doubles them: confirm the Enemy sprite has a separate \"when I receive [start game]\" script containing \"delete this clone\". It cannot go inside the spawn script, because it is an end block and nothing can follow it.",
      "Enemies keep moving after game over: confirm the \"if game_active = false then stop this script\" block is the very FIRST block inside the clone's forever loop, above all the movement, touching-Bullet and escape blocks.",
      "Pressing R mid-game causes chaos: confirm GameManager's \"when key r pressed\" script checks game_active before broadcasting - without that check, R would restart the game at any moment.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 9.",
      "In GameManager's script, add var game_active := true near the top, with the rest of the shared game state.",
      "Add the enemy scene (enemy.tscn) to a new Group named \"enemies\" (select the root node, open the Groups tab grouped with Inspector and Signals, type \"enemies\", click Add) - this lets you find and clear all enemies at once later.",
    ],
    level1: [
      "In the Main scene's script, move your spawn loop out of _ready() and into a new function spawn_enemies() (code below). Delete the old _ready() from the Main script - GameManager will now start the game.",
      "In GameManager's script, add a new function start_game() that clears any existing enemies, resets score, lives, enemies_remaining, and game_active, hides WinText and LoseText, updates the HUD, and then asks the Main scene to spawn a fresh set of enemies (code below).",
      "Change GameManager's _ready() to just call start_game.call_deferred(). Why the deferral: a child's _ready() runs BEFORE its parent has finished setting up, and a parent that is still busy refuses add_child() - so calling start_game() directly would silently spawn no enemies. call_deferred() waits one moment, until the whole scene has finished loading, and then runs it.",
      "In enemy.tscn's _process(), add \"if not get_node(\"../GameManager\").game_active: return\" as the very first line - this freezes movement once the round ends.",
      "In on_enemy_destroyed() and on_enemy_escaped(), when the win or lose condition is reached, also set game_active = false and set the WinText/LoseText text to include the restart hint.",
      "Add an _input(event) function to GameManager's script that checks for the R key and calls start_game() if the game isn't active (code below).",
      "Run the scene, play to a win or loss, confirm movement freezes, then press R and confirm everything resets and play resumes.",
    ],
    level1Code: [
      "# GameManager's script (complete):",
      "var enemies_remaining := 5",
      "var lives := 3",
      "var score := 0",
      "@@ var game_active := true",
      "",
      "func _ready():",
      "@@     start_game.call_deferred()",
      "",
      "@@ func start_game():",
      "@@     for old_enemy in get_tree().get_nodes_in_group(\"enemies\"):",
      "@@         old_enemy.queue_free()",
      "@@     enemies_remaining = 5",
      "@@     lives = 3",
      "@@     score = 0",
      "@@     game_active = true",
      "@@     get_node(\"../WinText\").hide()",
      "@@     get_node(\"../LoseText\").hide()",
      "@@     update_hud()",
      "@@     get_parent().spawn_enemies()",
      "",
      "@@ func _input(event):",
      "@@     if event is InputEventKey and event.pressed and event.keycode == KEY_R:",
      "@@         if not game_active:",
      "@@             start_game()",
      "",
      "func update_hud():",
      "    get_node(\"../ScoreText\").text = \"Score: \" + str(score)",
      "    get_node(\"../LivesText\").text = \"Lives: \" + str(lives)",
      "",
      "func on_enemy_destroyed():",
      "    enemies_remaining -= 1",
      "    score += 10",
      "    update_hud()",
      "    if enemies_remaining <= 0:",
      "@@         game_active = false",
      "@@         get_node(\"../WinText\").text = \"You Win! Press R to Restart\"",
      "        get_node(\"../WinText\").show()",
      "",
      "func on_enemy_escaped():",
      "    lives -= 1",
      "    update_hud()",
      "    if lives <= 0:",
      "@@         game_active = false",
      "@@         get_node(\"../LoseText\").text = \"Game Over - Press R to Restart\"",
      "        get_node(\"../LoseText\").show()",
      "",
      "# Main scene's script (complete):",
      "var enemy_scene := preload(\"res://enemy.tscn\")",
      "",
      "@@ func spawn_enemies():",
      "@@     for enemy_count in range(5):",
      "@@         var new_enemy := enemy_scene.instantiate()",
      "@@         new_enemy.position = Vector2(150 + enemy_count * 200, 50)",
      "@@         add_child(new_enemy)",
      "@@         new_enemy.destroyed.connect($GameManager.on_enemy_destroyed)",
      "@@         new_enemy.escaped.connect($GameManager.on_enemy_escaped)",
      "",
      "# enemy.tscn's script - add as the very first line of _process():",
      "func _process(delta):",
      "@@     if not get_node(\"../GameManager\").game_active:",
      "@@         return",
      "    position.y += enemy_speed * delta",
      "    if position.y > 700:",
      "        escaped.emit()",
      "        queue_free()",
    ],
    level2: [
      "Play to a win, confirm movement freezes, then press R and confirm score/lives/enemies_remaining are exactly back to 0/3/5 with no leftover enemy nodes.",
      "Repeat for a loss - same clean reset.",
      "Press R while the game is still active and confirm NOTHING happens - the \"if not game_active\" check in _input() is what prevents an accidental reset.",
    ],
    level3: [
      "Fix a real bug! We laid the groundwork for this bug all the way back in Lesson 6 when we defined our win scenario. Then in Lesson 8 we defined our lose scenario. By thinking in terms of win and lose separately, rather than in terms of wave over and game over, we created a bug. Find it and resolve it in the manner you, the Designer, think best. (Hint: play a round where you destroy some enemies and let the others escape, and watch what happens to the round, and to your R key.)",
      "Visual: fade WinText/LoseText in instead of an instant show(), using a Tween.",
      "Structure: notice start_game() is now the single place that sets up a fresh game - called both from GameManager's _ready() and from its own _input() restart handling. No setup code is duplicated anywhere, and every piece of shared game state still lives on GameManager.",
    ],
    trouble: [
      "Restart does nothing when you press R: confirm the keycode comparison uses KEY_R (capital, an exact Godot constant) and that game_active is actually false at that point.",
      "Restart leaves old enemies on screen: confirm every spawned enemy was actually added to the \"enemies\" group - a node not in the group won't be found and cleared by get_nodes_in_group().",
      "Enemies still move after game-over: confirm the \"if not ... game_active: return\" line is the very FIRST line inside _process(), before any movement code runs.",
      "No enemies appear at start but no error shows in the game: check the Debugger > Errors tab for \"Parent node is busy setting up children\". It means start_game() was called directly from _ready(); change it to start_game.call_deferred().",
      "Error calling spawn_enemies(): confirm the function is defined in the Main scene's script (GameManager reaches it through get_parent()), and that GameManager is a direct child of the Main scene.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Open your [YourName]Game project from Lesson 9.",
      "Select your enemy prefab and add a new Tag named \"Enemy\" (Inspector > Tag dropdown > Add Tag), then assign it to the prefab.",
    ],
    level1: [
      "In GameManager.cs, add public static bool gameActive = true, and update RegisterKill() and RegisterEscape() so that when hasWon or hasLost becomes true, gameActive is set to false too (code below).",
      "In EnemySpawner.cs, split your existing spawn code into two public methods: ClearEnemies() (destroys every GameObject tagged \"Enemy\") and SpawnEnemies() (your existing spawn loop), and call SpawnEnemies() from Start() (code below).",
      "In GameManager.cs, add a public EnemySpawner spawner field, and a public method ResetGame() (not static) that calls spawner.ClearEnemies(), resets score/lives/enemiesRemaining/hasWon/hasLost/gameActive, then calls spawner.SpawnEnemies(). Back in Unity, select the GameManager GameObject and drag the object that has EnemySpawner.cs from the Hierarchy onto its new Spawner slot in the Inspector, the same way you assigned the enemy prefab in Lesson 5.",
      "In Enemy.cs's Update(), add \"if (!GameManager.gameActive) return;\" as the very first line.",
      "In GameManager.cs, add an Update() that checks for Keyboard.current.rKey.wasPressedThisFrame and calls ResetGame() if the game isn't active (code below). If this is the first Input System check in GameManager.cs, add using UnityEngine.InputSystem; at the top of the file alongside using UnityEngine;. Update the \"You Win!\" and \"Game Over\" labels in OnGUI() to include \"Press R to Restart\".",
      "Press Play, play to a win or loss, confirm movement freezes, then press R and confirm everything resets and play resumes.",
    ],
    level1Code: [
      "// EnemySpawner.cs (complete):",
      "public GameObject enemyPrefab;",
      "public int enemyCount = 5;",
      "",
      "void Start() {",
      "@@     SpawnEnemies();",
      "}",
      "",
      "@@ public void ClearEnemies() {",
      "@@     foreach (GameObject oldEnemy in GameObject.FindGameObjectsWithTag(\"Enemy\")) {",
      "@@         Destroy(oldEnemy);",
      "@@     }",
      "@@ }",
      "",
      "@@ public void SpawnEnemies() {",
      "@@     for (int enemyIndex = 0; enemyIndex < enemyCount; enemyIndex++) {",
      "@@         Vector3 pos = new Vector3(-4f + enemyIndex * 2f, 4f, 0f);",
      "@@         Instantiate(enemyPrefab, pos, Quaternion.identity);",
      "@@     }",
      "@@ }",
      "",
      "// GameManager.cs (the new and revised members):",
      "@@ using UnityEngine.InputSystem;",
      "",
      "@@ public static bool gameActive = true;",
      "",
      "public static void RegisterKill() {",
      "    enemiesRemaining--;",
      "    score += 10;",
      "    if (enemiesRemaining <= 0) {",
      "        hasWon = true;",
      "@@         gameActive = false;",
      "    }",
      "}",
      "",
      "public static void RegisterEscape() {",
      "    lives--;",
      "    if (lives <= 0) {",
      "        hasLost = true;",
      "@@         gameActive = false;",
      "    }",
      "}",
      "",
      "@@ public EnemySpawner spawner;",
      "",
      "@@ public void ResetGame() {",
      "@@     spawner.ClearEnemies();",
      "@@     score = 0;",
      "@@     lives = 3;",
      "@@     enemiesRemaining = 5;",
      "@@     hasWon = false;",
      "@@     hasLost = false;",
      "@@     gameActive = true;",
      "@@     spawner.SpawnEnemies();",
      "@@ }",
      "",
      "@@ void Update() {",
      "@@     if (!gameActive && Keyboard.current.rKey.wasPressedThisFrame) {",
      "@@         ResetGame();",
      "@@     }",
      "@@ }",
      "",
      "// Enemy.cs - add as the very first line of Update():",
      "void Update() {",
      "@@     if (!GameManager.gameActive) return;",
      "    // ...existing movement and escape code...",
      "}",
    ],
    level2: [
      "Play to a win, confirm movement freezes, then press R and confirm score/lives/enemiesRemaining are exactly back to 0/3/5 with no leftover enemy GameObjects.",
      "Repeat for a loss - same clean reset.",
      "Press R while the game is still active and confirm NOTHING happens.",
    ],
    level3: [
      "Fix a real bug! We laid the groundwork for this bug all the way back in Lesson 6 when we defined our win scenario. Then in Lesson 8 we defined our lose scenario. By thinking in terms of win and lose separately, rather than in terms of wave over and game over, we created a bug. Find it and resolve it in the manner you, the Designer, think best. (Hint: play a round where you destroy some enemies and let the others escape, and watch what happens to the round, and to your R key.)",
      "Visual: add a fade or scale animation to the win/lose text instead of it just appearing instantly.",
      "Structure: notice ResetGame() is now the single place that sets up a fresh game - called both from Start() (via SpawnEnemies(), indirectly the first time) and directly on restart.",
    ],
    trouble: [
      "NullReferenceException inside ResetGame(): the Spawner slot on the GameManager component is empty. Select the GameManager GameObject and drag the object that has EnemySpawner.cs onto the Spawner slot.",
      "Restart leaves old enemies on screen: confirm the enemy prefab actually has the \"Enemy\" tag assigned - a forgotten tag means FindGameObjectsWithTag() finds nothing.",
      "Enemies still move after game-over: confirm \"if (!GameManager.gameActive) return;\" is the very FIRST line inside Update(), before any movement code runs.",
      "R does nothing: confirm GameManager.cs is attached to an active GameObject in the scene (its Update() only runs if it is), and that using UnityEngine.InputSystem; is at the top of the file.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same reusable-reset pattern works in a Script Graph instead of C#, using the same new Input System as the code path.",
      steps: [
        "Open PlayerControls (the Input Actions asset from Lesson 2's Player Input component) and add a new action to the Player action map named Restart, Action Type Button, bound to R. Click Save Asset.",
        "Build a Custom Event graph named \"ResetGame\" on a persistent GameObject, containing nodes that clear existing enemy objects, reset your score/lives/enemies variables, and re-trigger your spawn graph.",
        "On a graph on the Player (or PlayerInput) GameObject, right-click and add an \"On Input System Event Button\" node. Leave Target as This, set Input Action to Restart, and set Input Action Change Type to On Pressed. Wire its Trigger output into a Branch node checking GameActive is false, then into the ResetGame custom event on True.",
      ],
      note: "An On Input System Event Button node set to On Pressed only fires once per physical press (event-driven, not polled), same as wasPressedThisFrame in the code path - no extra debounce logic needed.",
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: [
      "Open your [YourName]Game project from Lesson 9.",
      "Create a new Script resource named scr_start_game - this is GameMaker's way of defining a reusable function, similar to what you'll see in several other engines this week.",
    ],
    level1: [
      "In scr_start_game, write a function start_game() that destroys all existing obj_enemy instances, resets obj_gamemanager's __dnd_score, __dnd_lives, enemies_remaining, has_won, has_lost, and game_active, then spawns a fresh set of enemies (code below).",
      "Change obj_spawner's Create event to just call start_game();.",
      "On obj_gamemanager's Create event, add game_active = true; alongside your other shared game state.",
      "On obj_enemy's Step event, add \"if (!obj_gamemanager.game_active) exit;\" as the very first line - this freezes movement once the round ends.",
      "In obj_bullet's Collision event and obj_enemy's Step event's escape branch, when setting has_won/has_lost, also set obj_gamemanager.game_active = false, and update the has_won/has_lost draw_text() strings in obj_gamemanager's Draw GUI event to include the restart hint.",
      "Add a Step event to obj_gamemanager that checks keyboard_check_pressed(ord(\"R\")) and calls start_game() if !game_active.",
      "Press F5 (Run), play to a win or loss, confirm movement freezes, then press R and confirm everything resets and play resumes.",
    ],
    level1Code: [
      "// scr_start_game (Script resource - new):",
      "@@ function start_game() {",
      "@@     with (obj_enemy) {",
      "@@         instance_destroy();",
      "@@     }",
      "@@     obj_gamemanager.__dnd_score = 0;",
      "@@     obj_gamemanager.__dnd_lives = 3;",
      "@@     obj_gamemanager.enemies_remaining = 5;",
      "@@     obj_gamemanager.has_won = false;",
      "@@     obj_gamemanager.has_lost = false;",
      "@@     obj_gamemanager.game_active = true;",
      "@@     for (var enemy_count = 0; enemy_count < 5; enemy_count++) {",
      "@@         instance_create_layer(80 + enemy_count * 96, 80, \"Instances\", obj_enemy);",
      "@@     }",
      "@@ }",
      "",
      "// obj_spawner's Create event (revised - the spawn loop moved into start_game()):",
      "@@ start_game();",
      "",
      "// obj_gamemanager's Create event:",
      "enemies_remaining = 5;",
      "has_won = false;",
      "__dnd_lives = 3;",
      "has_lost = false;",
      "__dnd_score = 0;",
      "@@ game_active = true;",
      "",
      "// obj_gamemanager's Step event (new):",
      "@@ if (!game_active && keyboard_check_pressed(ord(\"R\"))) {",
      "@@     start_game();",
      "@@ }",
      "",
      "// obj_enemy's Step event (revised):",
      "@@ if (!obj_gamemanager.game_active) exit;",
      "y += enemy_speed;",
      "if (y > room_height) {",
      "    obj_gamemanager.__dnd_lives -= 1;",
      "    if (obj_gamemanager.__dnd_lives <= 0) {",
      "        obj_gamemanager.has_lost = true;",
      "@@         obj_gamemanager.game_active = false;",
      "    }",
      "    instance_destroy();",
      "}",
      "",
      "// obj_bullet's Collision event with obj_enemy (revised):",
      "audio_play_sound(sfx_hit, 1, false);",
      "obj_gamemanager.enemies_remaining -= 1;",
      "obj_gamemanager.__dnd_score += 10;",
      "if (obj_gamemanager.enemies_remaining <= 0) {",
      "    obj_gamemanager.has_won = true;",
      "@@     obj_gamemanager.game_active = false;",
      "}",
      "instance_destroy(other);",
      "instance_destroy(self);",
      "",
      "// obj_gamemanager's Draw GUI event (revised):",
      "draw_text(20, 20, \"Score: \" + string(__dnd_score));",
      "draw_text(20, 50, \"Lives: \" + string(__dnd_lives));",
      "if (has_won) {",
      "@@     draw_text(400, 50, \"You Win! Press R to Restart\");",
      "}",
      "if (has_lost) {",
      "@@     draw_text(400, 90, \"Game Over - Press R to Restart\");",
      "}",
    ],
    level2: [
      "Play to a win, confirm movement freezes, then press R and confirm __dnd_score/__dnd_lives/enemies_remaining are exactly back to 0/3/5 with no leftover obj_enemy instances.",
      "Repeat for a loss - same clean reset.",
      "Press R while the game is still active and confirm NOTHING happens.",
    ],
    level3: [
      "Fix a real bug! We laid the groundwork for this bug all the way back in Lesson 6 when we defined our win scenario. Then in Lesson 8 we defined our lose scenario. By thinking in terms of win and lose separately, rather than in terms of wave over and game over, we created a bug. Find it and resolve it in the manner you, the Designer, think best. (Hint: play a round where you destroy some enemies and let the others escape, and watch what happens to the round, and to your R key.)",
      "Visual: add a short screen-flash or fade using draw_set_alpha() in Draw GUI when the game freezes.",
      "Structure: notice start_game() is now the single place that sets up a fresh game - called both from obj_spawner's Create event and from its own restart check in Step. No setup code is duplicated anywhere.",
    ],
    trouble: [
      "Restart does nothing when you press R: confirm keyboard_check_pressed(ord(\"R\")) is inside an event that actually runs every frame (obj_gamemanager's Step), and that !game_active is true at that point.",
      "Restart leaves old enemies on screen: confirm \"with (obj_enemy) { instance_destroy(); }\" runs BEFORE the spawn loop inside start_game(), not after.",
      "Enemies still move after game-over: confirm \"if (!obj_gamemanager.game_active) exit;\" is the very FIRST line inside obj_enemy's Step event, before y += enemy_speed.",
      "Score or lives don't reset: confirm start_game() sets obj_gamemanager.__dnd_score and obj_gamemanager.__dnd_lives (with the double leading underscore) - the old built-in score and lives are deprecated and not what the HUD reads.",
    ],
    resource: "Student Resource List > GameMaker section, or gamemaker.io/en/tutorials",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same reusable-reset pattern works with drag-and-drop blocks instead of typed code. Every Score, Lives, and Set Variable block that changes shared state needs its Target set to obj_gamemanager, exactly as in Lessons 8 and 9.",
      steps: [
        "Create a new Script resource named scr_start_game. Its top-level block must be \"Declare a New Function\" with Name = start_game, and EVERY block below goes INSIDE that function block, including the whole Repeat loop: (1) an \"Instance - Destroy Instance\" block with Applies To set to obj_enemy; (2) a \"Set Score\" block, Target obj_gamemanager, Value 0, Relative unchecked; (3) a \"Set Lives\" block, Target obj_gamemanager, Value 3, Relative unchecked; (4) four \"Set Variable\" blocks with Target obj_gamemanager: enemies_remaining = 5, has_won = false, has_lost = false, and game_active = true; (5) your existing spawn-loop blocks (Repeat 5 with Create Instance inside). Check the connectors: any block that is not nested inside the function runs once when the game starts, before any room exists, and the Create Instance block will fail with a \"layer does not exist\" error.",
        "On obj_spawner's Create event, replace the spawn-loop blocks with a single \"Function Call\" block, Function = start_game (leave Target empty and Temp unchecked). Use Function Call, not the Execute Script block: Execute Script runs the script resource itself, which only declares the function and creates no enemies.",
        "On obj_gamemanager's Create event, add a \"Set Variable\" block: game_active = true.",
        "On obj_gamemanager's Step event, add an \"If Variable\" block (the one with a variable name, an operation, and a value - not the true/false block) checking game_active equal to false. Nest an \"If Key Pressed\" block inside it: click the selector to the right of the Key field and choose Letters, then R, which fills the Key field with ord(\"R\"). Inside that, nest a \"Function Call\" block, Function = start_game.",
        "On obj_enemy's Step event, add an \"If Variable\" block at the very top: Target obj_gamemanager, game_active equal to false, with an \"Exit Event\" block nested INSIDE it (check the connector, as above). The movement blocks go after the If block, not inside it. In both places where has_won or has_lost is set to true (obj_bullet's Collision event and obj_enemy's escape branch), add a \"Set Variable\" block right after it with Target obj_gamemanager: game_active = false.",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 9.",
      "Add two new scene variables: GameActive (default value 1) and RestartRequested (default value 0).",
      "Add a new scene variable EnemyIndex (default value 0) if you don't already have one from Lesson 5.",
    ],
    level1: [
      "Find your existing spawn event from Lesson 5 (it creates the Enemy objects at the start of the scene) and your existing event that sets the starting Score, Lives, and EnemiesRemaining. Move all of them into one labeled event group called \"Setup\".",
      "Change the Setup group's condition to an \"Or\" condition: [At the beginning of the scene] OR [RestartRequested = 1] (code below) - this makes the exact same setup logic run both the first time AND every restart.",
      "Inside the Setup group, add actions to delete all existing Enemy instances, reset Score, Lives, EnemiesRemaining, EnemyIndex, and GameActive to their starting values, set RestartRequested back to 0, hide WinText and LoseText, and then run your existing spawn actions (code below).",
      "On your enemy movement event, add GameActive = 1 as an additional condition - this freezes movement once the round ends.",
      "On the Combat group's win-check and the escape event's lose-check, when EnemiesRemaining or Lives reaches zero, also set GameActive to 0, and change the WinText/LoseText content to include the restart hint.",
      "Add a new event: Condition \"Key R is pressed\" AND \"GameActive = 0\" -> Action \"RestartRequested = 1\".",
      "Click Preview, play to a win or loss, confirm movement freezes, then press R and confirm everything resets and play resumes.",
    ],
    level1Code: [
      "(Setup event group)",
      "@@ Condition: Or [At the beginning of the scene] [RestartRequested = 1]",
      "@@ Action: For each Enemy -> Delete Enemy",
      "@@ Action: Score = 0",
      "@@ Action: Lives = 3",
      "@@ Action: EnemiesRemaining = 5",
      "@@ Action: EnemyIndex = 0",
      "@@ Action: GameActive = 1",
      "@@ Action: RestartRequested = 0",
      "@@ Action: Hide WinText",
      "@@ Action: Hide LoseText",
      "Action: Repeat 5 times -> Create object Enemy at (40 + EnemyIndex * 80, 40)",
      "Action: EnemyIndex = EnemyIndex + 1   (inside the repeat)",
      "",
      "(Enemy movement event, revised)",
      "@@ Condition: GameActive = 1",
      "Action: Enemy.Y = Enemy.Y() + 30 * TimeDelta()",
      "",
      "(Combat group - win check, revised)",
      "Condition: EnemiesRemaining <= 0",
      "@@ Action: GameActive = 0",
      "@@ Action: Change text of WinText: \"You Win! Press R to Restart\"",
      "Action: Show WinText",
      "",
      "(Escape group - lose check, revised)",
      "Condition: Lives <= 0",
      "@@ Action: GameActive = 0",
      "@@ Action: Change text of LoseText: \"Game Over - Press R to Restart\"",
      "Action: Show LoseText",
      "",
      "@@ (new event)",
      "@@ Condition: Key R is pressed",
      "@@ Condition: GameActive = 0",
      "@@ Action: RestartRequested = 1",
    ],
    level2: [
      "Play to a win, confirm movement freezes, then press R and confirm Score/Lives/EnemiesRemaining are exactly back to 0/3/5 with no leftover Enemy instances.",
      "Repeat for a loss - same clean reset.",
      "Press R while the game is still active and confirm NOTHING happens - the GameActive = 0 condition on the restart event is what prevents an accidental reset.",
    ],
    level3: [
      "Visual: add a fade-in/out effect using an opacity Tween on WinText/LoseText instead of an instant show/hide.",
      "Structure: notice the Setup group is now the single place that sets up a fresh game - reached both \"At the beginning of the scene\" and on every restart via the Or condition. No setup logic is duplicated anywhere in your event sheet.",
    ],
    trouble: [
      "Restart does nothing when you press R: confirm the new event's two conditions are both present (Key R pressed AND GameActive = 0) - forgetting the GameActive check means R does nothing at all, since RestartRequested would never lead anywhere without the Setup group's Or condition also being satisfied.",
      "Restart leaves old enemies on screen: confirm \"For Each Enemy -> Delete Enemy\" runs BEFORE the Repeat-5-times spawn action inside the Setup group, not after.",
      "Enemies still move after game-over: confirm GameActive = 1 was actually added as a condition on the movement event, not just referenced elsewhere.",
    ],
    resource: "Student Resource List > GDevelop section, or gdevelop.io/tutorials",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: ["Launch TIC-80 and load your cart from Lesson 9.", "Press Esc to reach the code editor."],
    level1: [
      "Add game_active=true near the top of your code, with your other game state (score, lives, enemies_remaining).",
      "Find your existing one-time setup code (populating the enemies table, setting lives, score, enemies_remaining). Move it into a new function start_game() (code below), called once at the top level in place of the old inline setup.",
      "In the enemy-drawing loop, wrap the movement and escape code in \"if game_active then ... end\" - this freezes movement once the round ends.",
      "In the bullets loop's hit branch and the enemy loop's escape branch, when enemies_remaining or lives reaches zero, also set game_active=false, and change the print() text to include the restart hint.",
      "At the bottom of TIC(), add a check using keyp(18) (TIC-80's key code for R) that calls start_game() if game_active is false (code below).",
      "Press Esc, type run, press Enter, then play to a win or loss, confirm movement freezes, then press R and confirm everything resets and play resumes.",
    ],
    level1Code: [
      "-- player position x and y are set above this, exactly as in earlier lessons",
      "@@ game_active=true",
      "",
      "@@ function start_game()",
      "@@  enemies={}",
      "@@  for enemy_count=0,4 do",
      "@@   table.insert(enemies,{ex=20+enemy_count*20,ey=10,alive=true})",
      "@@  end",
      "@@  bullets={}",
      "@@  score=0",
      "@@  lives=3",
      "@@  enemies_remaining=5",
      "@@  game_active=true",
      "@@ end",
      "",
      "@@ start_game()",
      "",
      "function TIC()",
      " cls(0)",
      " if btn(3) then x=x+1 end",
      " if btnp(4) then",
      "  table.insert(bullets,{bx=x,by=y})",
      " end",
      " for bullet_index=#bullets,1,-1 do",
      "  local b=bullets[bullet_index]",
      "  b.by=b.by-2",
      "  local hit=false",
      "  for enemy_index=1,#enemies do",
      "   local e=enemies[enemy_index]",
      "   if e.alive and b.bx<e.ex+8 and b.bx+8>e.ex and b.by<e.ey+8 and b.by+8>e.ey then",
      "    e.alive=false",
      "    enemies_remaining=enemies_remaining-1",
      "    score=score+10",
      "@@     if enemies_remaining<=0 then game_active=false end",
      "    hit=true",
      "   end",
      "  end",
      "  if hit then",
      "   table.remove(bullets,bullet_index)",
      "  elseif b.by<0 then",
      "   table.remove(bullets,bullet_index)",
      "  else",
      "   spr(2,b.bx,b.by,0)",
      "  end",
      " end",
      " for enemy_index=1,#enemies do",
      "  local enemy=enemies[enemy_index]",
      "  if enemy.alive then",
      "@@    if game_active then",
      "    enemy.ey=enemy.ey+0.3",
      "    if enemy.ey>136 then",
      "     enemy.alive=false",
      "     lives=lives-1",
      "@@     if lives<=0 then game_active=false end",
      "    end",
      "@@    end",
      "    spr(3,enemy.ex,enemy.ey,0)",
      "  end",
      " end",
      " spr(1,x,y,0)",
      " print(\"Score: \"..score,5,5,12)",
      " print(\"Lives: \"..lives,5,15,12)",
      " if enemies_remaining<=0 then",
      "@@  print(\"You Win! Press R to Restart\",50,60,12)",
      " end",
      " if lives<=0 then",
      "@@  print(\"Game Over - Press R to Restart\",40,60,8)",
      " end",
      "@@ if keyp(18) and not game_active then",
      "@@  start_game()",
      "@@ end",
      "end",
    ],
    level2: [
      "Play to a win, confirm movement freezes, then press R and confirm score/lives/enemies_remaining are exactly back to 0/3/5 with no leftover live entries in the enemies table.",
      "Repeat for a loss - same clean reset.",
      "Press R while the game is still active and confirm NOTHING happens - the \"and not game_active\" check is what prevents an accidental reset.",
    ],
    level3: [
      "Visual: briefly flash the screen (cls() with a different color for one frame) the moment game_active flips to false.",
      "Structure: notice start_game() is now the single place that sets up a fresh game - called once when the cart first loads, and again from inside TIC() on every restart. No setup code is duplicated anywhere.",
    ],
    trouble: [
      "Restart does nothing when you press R: confirm you're using keyp(18) (edge-triggered, fires once per press) rather than key(18) (fires every frame the key is held) - using key() here would call start_game() repeatedly while R is held.",
      "Restart leaves old enemies on screen: confirm start_game() actually replaces the enemies table with a fresh one (enemies={}) rather than trying to modify the old one in place.",
      "Enemies still move after game-over: confirm the \"if game_active then ... end\" wrapper actually surrounds the movement/escape code inside the enemy-drawing loop, not just the drawing (spr()) call.",
    ],
    resource: "Student Resource List > TIC-80 section, or tic80.com",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 9.",
      "Create a new Variable (Resources > Variables/Switches/Collections/Tables) named game_active, default value 1.",
      "Create a new Custom Event (Resources > Custom Events) named ResetGame - this is 001 Game Creator's reusable-subroutine feature, callable from any trigger's script.",
      "Create a new Input Trigger for the R key (Input Set > Add Control Combination > press R > choose \"Pressed\" > Edit Script) - the same underlying mechanism your movement and shoot controls have used since Lesson 2.",
    ],
    level1: [
      "On the enemy Actor Template's \"Collided with actor\" trigger, wrap the score/enemies_remaining/win-check actions in an If(game_active = 1) check, and when enemies_remaining reaches 0, also set game_active to 0 (code below).",
      "On the enemy Actor Template's \"Left/Entered Camera\" trigger, wrap the lives-decrement action in the same If(game_active = 1) check, and when lives reaches 0, also set game_active to 0.",
      "Update both Message Box texts to include the restart hint: \"You Win! - Press R to Restart\" and \"Game Over - Press R to Restart\".",
      "On your new Input Trigger (R, Pressed), add: If(game_active = 0) -> Call Custom Event(ResetGame).",
      "Build the ResetGame Custom Event: reset score/lives/enemies_remaining/game_active to their starting values, then revive and reposition each of your enemy copies back to their original starting spots (code below).",
      "Save, then Play/Test: play to a win or loss, then press R and confirm everything resets and play resumes.",
    ],
    level1Code: [
      "enemy Actor Trigger: Collided with actor (revised) ->",
      "  If(game_active = 1)",
      "    Variable – Modify Value(score, Add, 10)",
      "    Variable – Modify Value(enemies_remaining, Subtract, 1)",
      "    If(enemies_remaining <= 0)",
      "      Variable – Modify Value(game_active, Set, 0)",
      "      Message Box(\"You Win! - Press R to Restart\")",
      "    Actor – Destroy/Revive Actor(relate, Destroy)",
      "    Actor – Destroy/Revive Actor(this, Destroy)",
      "",
      "enemy Actor Trigger: Left/Entered Camera (revised) ->",
      "  If(game_active = 1)",
      "    Variable – Modify Value(lives, Subtract, 1)",
      "    If(lives <= 0)",
      "      Variable – Modify Value(game_active, Set, 0)",
      "      Message Box(\"Game Over - Press R to Restart\")",
      "    Actor – Destroy/Revive Actor(this, Destroy)",
      "",
      "Input Trigger: R (Pressed) ->",
      "  If(game_active = 0)",
      "    Call Custom Event(ResetGame)",
      "",
      "Custom Event: ResetGame ->",
      "  Variable – Modify Value(score, Set, 0)",
      "  Variable – Modify Value(lives, Set, 3)",
      "  Variable – Modify Value(enemies_remaining, Set, 5)",
      "  Variable – Modify Value(game_active, Set, 1)",
      "  Actor – Destroy/Revive Actor(enemy_1, Revive) -> Actor – Teleport(enemy_1, [starting position])",
      "  (repeat for enemy_2, enemy_3, enemy_4, enemy_5)",
    ],
    level2: [
      "Play to a win, then press R and confirm score/lives/enemies_remaining are exactly back to 0/3/5, with all 5 enemy copies revived at their starting positions.",
      "Repeat for a loss - same clean reset.",
      "Press R while the game is still active and confirm NOTHING happens - the If(game_active = 0) check on your Input Trigger is what prevents an accidental reset.",
    ],
    level3: [
      "Visual: fade the screen briefly (a simple Camera or Interface fade effect, if you're comfortable exploring the wiki further) the moment game_active flips to 0.",
      "Structure: notice ResetGame is now the single place that sets up a fresh game, callable from your Input Trigger any time it's needed - the same setup values (score, lives, enemies_remaining) only ever need to be written once.",
      "Design note: because your enemies use Automatic Movement rather than a spawn loop, this lesson intentionally doesn't force-stop enemies already in motion at the moment of game-over - they'll simply drift off naturally within a second or two, harmlessly, since the If(game_active = 1) check means they no longer affect score or lives once the round has ended. This is a deliberate, reasonable simplification for this engine, not a cut corner.",
    ],
    trouble: [
      "Restart does nothing when you press R: confirm the Input Trigger's Control Combination is actually set to the R key and \"Pressed\" (not \"Released\"), and that game_active is actually 0 at that point.",
      "Score/lives keep changing after \"Game Over\": confirm the If(game_active = 1) check was actually added to BOTH triggers (Collided with actor AND Left/Entered Camera), not just one.",
      "Enemies don't reappear after restart: confirm each enemy copy's exact name (enemy_1, enemy_2, etc.) matches what you use in the ResetGame Custom Event - a typo'd name silently does nothing.",
    ],
    resource: "001gamecreator.com/wiki/customevents/, 001gamecreator.com/wiki/inputtriggers/, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 9.",
      "In WBP_HUD's Designer view, add a third Text Block named StatusText, positioned centrally, initially empty.",
      "Open BP_GameManager and add a new Boolean variable GameActive, default value true - it joins Score, Lives, and EnemiesRemaining as shared game state.",
      "In the Content Browser, click Add (+) > Input > Input Action, name it IA_Restart. Leave its Value Type as the default (Digital bool), same as your other Input Actions.",
      "Open IMC_Default (the Input Mapping Context from Lesson 2) and click + Add Mapping. Choose IA_Restart, then set its Key to R. Since IMC_Default is already applied to the player, this new mapping is live immediately.",
    ],
    level1: [
      "In WBP_HUD's Event Graph, add a Custom Event SetStatus(StatusText: String) that sets the StatusText Text Block's text from the input (code below).",
      "In BP_Enemy's Event Graph, add a Branch at the very start of Event Tick that reads GameActive from BP_GameManager (use Get Actor of Class(BP_GameManager), then Get GameActive). Wrap all existing movement logic inside the True case, so nothing moves once the round ends.",
      "In BP_GameManager, build a new Custom Event ResetGame: destroy all existing BP_Enemy actors, reset Score/Lives/EnemiesRemaining/GameActive, call SetStatus with an empty string, call UpdateHUD, then spawn a fresh set of enemies with the same For Loop your spawn logic already used (code below). BP_GameManager owns the shared game state, so the reset lives here and not in the Level Blueprint.",
      "Change BP_GameManager's Event BeginPlay to create the widget, add it to viewport, set HUDWidget, then call ResetGame instead of spawning enemies directly.",
      "In HandleEnemyDestroyed's win Branch and HandleEnemyEscaped's lose Branch, when the round ends, also set GameActive = false and call SetStatus with the matching restart-hint message.",
      "In BP_GameManager's Event Graph, right-click and search for IA_Restart to add an Enhanced Input Action Event node: on its Started pin, Branch on !GameActive -> True: Call ResetGame. (If BP_GameManager does not receive input in your project, enable Auto Receive Input on it in Class Defaults, or add the node to your player Blueprint and call ResetGame on BP_GameManager from there.)",
      "Click Compile, then Play, play to a win or loss, confirm movement freezes, then press R and confirm everything resets and play resumes.",
    ],
    level1Code: [
      "WBP_HUD's Event Graph (new):",
      "@@ Custom Event SetStatus(StatusText: String) ->",
      "@@   Set Text(StatusTextBlock, StatusText)",
      "",
      "BP_Enemy's Event Graph (revised):",
      "Event Tick ->",
      "@@   Branch(Get GameActive from Get Actor of Class(BP_GameManager)) -> True:",
      "    Add Actor World Offset(Delta = (0, -EnemySpeed * Delta Seconds, 0))",
      "    Branch(GetActorLocation().Y < -400) -> True: Call HandleEnemyEscaped on BP_GameManager -> Destroy Actor(Self)",
      "",
      "BP_GameManager (variables: Score, Lives, EnemiesRemaining, HUDWidget, GameActive):",
      "@@ Custom Event ResetGame ->",
      "@@   For Each (Get All Actors of Class(BP_Enemy)) -> Destroy Actor",
      "@@   Score = 0",
      "@@   Lives = 3",
      "@@   EnemiesRemaining = 5",
      "@@   GameActive = true",
      "@@   Call SetStatus(HUDWidget, \"\")",
      "@@   Call UpdateHUD(HUDWidget, Score, Lives)",
      "@@   For Loop (First Index 0, Last Index 4) -> Spawn Actor from Class(BP_Enemy) at (BaseX + Index * 150, BaseY, 0) -> Bind Event to OnEnemyDestroyed / OnEnemyEscaped",
      "",
      "Event BeginPlay ->",
      "  Create Widget(WBP_HUD) -> Add to Viewport -> Set HUDWidget",
      "@@   Call ResetGame",
      "",
      "Custom Event HandleEnemyDestroyed ->",
      "  EnemiesRemaining = EnemiesRemaining - 1",
      "  Score = Score + 10",
      "  Call UpdateHUD(HUDWidget, Score, Lives)",
      "  Branch(EnemiesRemaining <= 0) -> True:",
      "@@     GameActive = false",
      "@@     Call SetStatus(HUDWidget, \"You Win! Press R to Restart\")",
      "",
      "Custom Event HandleEnemyEscaped ->",
      "  Lives = Lives - 1",
      "  Call UpdateHUD(HUDWidget, Score, Lives)",
      "  Branch(Lives <= 0) -> True:",
      "@@     GameActive = false",
      "@@     Call SetStatus(HUDWidget, \"Game Over - Press R to Restart\")",
      "",
      "@@ Enhanced Input Action Event (IA_Restart), Started ->",
      "@@   Branch(!GameActive) -> True: Call ResetGame",
    ],
    level2: [
      "Play to a win, confirm movement freezes, then press R and confirm Score/Lives/EnemiesRemaining are exactly back to 0/3/5 with no leftover BP_Enemy actors.",
      "Repeat for a loss - same clean reset.",
      "Press R while the game is still active and confirm NOTHING happens.",
    ],
    level3: [
      "Visual: animate StatusText's opacity in WBP_HUD instead of it appearing instantly.",
      "Structure: notice ResetGame is now the single place that sets up a fresh game - called once from BeginPlay and again from the IA_Restart Enhanced Input Action Event. No spawn logic is duplicated, and all shared state still lives on BP_GameManager.",
    ],
    trouble: [
      "Restart does nothing when you press R: confirm IA_Restart actually has an R mapping inside IMC_Default, that the Enhanced Input Action Event node is wired directly in the Event Graph (not nested inside another event), that BP_GameManager is set to receive input, and that GameActive is actually false at that point.",
      "Restart leaves old enemies on screen: confirm \"Get All Actors of Class(BP_Enemy)\" -> Destroy Actor runs BEFORE the spawn For Loop inside ResetGame, not after.",
      "Enemies still move after game-over: confirm the Branch(GameActive) wraps ALL of Event Tick's movement logic, not just part of it.",
      "Get GameActive returns nothing: confirm exactly one BP_GameManager is placed in the level.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same reusable-reset pattern works in C++ instead of Blueprint nodes.",
      steps: [
        "Add a ResetGame() function to your game manager class that iterates TActorIterator<ACppEnemy> to destroy existing enemies, resets Score/Lives/EnemiesRemaining/bGameActive, and calls your existing spawn function.",
        "In ACppEnemy::Tick(), add an early return (if (!bGameActive) return;) at the very top, mirroring the Blueprint path's Branch gate.",
      ],
      note: "This mirrors the Blueprint path's ResetGame Custom Event and Tick Branch exactly - same shape, just from code.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_10_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 10 Teacher's Key: Play It Again", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine + " - Instructor Reference Only", size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    h2("What This Lesson Teaches"),
    p(data.teaches),
    h2("Level 1 - Full Worked Solution"),
    ...data.level1.map(s => bullet(s)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  children.push(...diagram(...ACTIVITY));
  if (ENGINE_DIAGRAMS[data.file]) children.push(...diagram(ENGINE_DIAGRAMS[data.file].file, ENGINE_DIAGRAMS[data.file].w, ENGINE_DIAGRAMS[data.file].h, ENGINE_DIAGRAMS[data.file].caption + (data.file === "GameMaker" ? " (the Visual path)" : "")));
  children.push(
    h2("Why This Structure"),
    ...data.why.map(s => bullet(s)),
    h2("Level 2 / Level 3 Talking Points"),
    ...data.talkingPoints.map(s => bullet(s)),
    ...(data.level2Code ? [p("Level 2 - complete code:", { bold: true }), code(data.level2Code)] : []),
    ...(data.level3Code ? [p("Level 3 - code (additions highlighted):", { bold: true }), code(data.level3Code, "Highlighted lines are the Level 3 additions on top of the Level 1 code.")] : []),
    h2("Common Student Mistakes"),
    ...data.mistakes.map(s => bullet(s)),
  );
  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 10 Key - " + data.engine + " (Instructor Only)") },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    teaches: "The clearest possible version of this lesson's core idea: a broadcast is Scratch's natural 'reusable routine' - the same broadcast can be sent from the green flag AND from a key-press script, and every receiving script runs identically either way.",
    level1: [
      "The refactor is almost entirely mechanical: what used to run directly under \"when green flag clicked\" now runs under \"when I receive [start game]\" instead, with the green flag script reduced to a single broadcast.",
      "The freeze is one new block at the top of the clone's existing forever loop: \"if game_active = false then stop this script\". It is an early exit, the same shape as Godot's early return, GameMaker's exit and Unity's early return, and it needs no re-indenting of the existing blocks. (Credit: Robert's own testing improved on the original wrap-it-in-an-if design.) A finished round freezes the enemies in place, matching the other engines.",
      "game_active is flipped to false at the exact moment score/lives would have crossed to a win/lose state, right alongside the existing broadcast - one new line in an already-correct location, twice (once per branch).",
    ],
    why: [
      "Broadcasts were already the tool Lesson 6 used to announce a win - this lesson simply reveals that the exact same tool can also be the entry point for a full game reset, with zero new Scratch concepts required.",
      "The clean-up is a SECOND \"when I receive [start game]\" script on the Enemy sprite holding only \"delete this clone\" (an end block, so it cannot sit above the spawn loop). Clones hear broadcasts too: every old clone deletes itself, while the original sprite ignores the block and its own spawn script makes the new set. Without it, surviving clones would also run the spawn loop on restart.",
    ],
    talkingPoints: [
      "Level 3 bug hunt: kills reduce enemies_remaining but escapes only reduce lives, so a round with some kills and some escapes (for example 3 and 2) ends with no enemies left, no win and no loss, and R does nothing because game_active is still true. This started in Lesson 6 (win = all killed) and Lesson 8 (lose = lives gone). A good fix, and the one shown here, is that an escape also resolves an enemy and the win check requires lives above 0. Accept any design that guarantees every round ends in exactly one outcome (for example a third 'Round Over' state, or a check that no enemies remain and none are in flight). Test it with 3 kills and 2 escapes, then 4 and 1, then 0 and 5.",
      "Ask students to trace, out loud, everything that happens when they press R while game_active is false - broadcast sent, Enemy's \"when I receive\" script runs, the clones run their delete script while the original sprite spawns the fresh five. A good comprehension check.",
      "For strong students: ask what would happen if the green-flag script were left as a full copy of the setup code (not refactored into a broadcast) alongside a separate restart script - the answer (two copies of the same logic, now capable of drifting out of sync) is exactly the problem this lesson's structure avoids.",
    ],
    mistakes: [
      "Game-state drift: the student added start-up code or a game-over effect beyond the handout (extra green-flag scripts, an animation, a color change, a hidden object) and did not reset it on restart. Symptom: a restart that looks almost right. Walk through the Handout's \"Audit Your Game State\" list with them.",
      "Forgetting to remove the setup code from the green-flag script after moving it into \"when I receive [start game]\" - results in setup effectively running twice at game start (harmless here, but worth flagging as untidy duplication).",
      "Putting the stop-this-script check anywhere but the very top of the clone's forever loop, so some movement or escape blocks still run after game over.",
      "Trying to put \"delete this clone\" inside the spawn script (it is an end block, nothing can follow it) or forgetting the separate delete script: old clones also run the spawn loop on restart, multiplying enemies.",
      "Forgetting the game_active = false check on the restart key script, allowing R to interrupt an active round.",
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    teaches: "Extracting existing _ready() code into a named, reusable function (start_game()) - the single clearest, most explicit version of the 'factor out the reset routine' idea this week, since Godot's function syntax makes the refactor completely visible.",
    level1: [
      "start_game() lives on GameManager (which owns the shared state) and resets it, then asks Main's spawn_enemies() to rebuild the wave - the spawn loop is the old _ready() body, moved into a function.",
      "_ready() shrinks to a single line: start_game.call_deferred(). This is the clearest illustration all week of 'the entry point calls the reusable routine, rather than containing the logic directly.'",
      "get_tree().get_nodes_in_group(\"enemies\") is the mechanism that makes clearing old enemies possible - without the Group assignment in Setup, there'd be no reliable way to find every existing enemy to remove.",
    ],
    why: [
      "This mirrors exactly what Lesson 9 did with update_hud() (one function, called from multiple places) - start_game() is the same idea, applied to setup instead of display.",
      "_input(event) is used here (rather than an Input Map action) specifically to keep this lesson self-contained - no Project Settings configuration required, consistent with the course's self-containment principle.",
    ],
    talkingPoints: [
      "Level 3 bug hunt: kills reduce enemies_remaining but escapes only reduce lives, so a round with some kills and some escapes (for example 3 and 2) ends with no enemies left, no win and no loss, and R does nothing because game_active is still true. This started in Lesson 6 (win = all killed) and Lesson 8 (lose = lives gone). A good fix, and the one shown here, is that an escape also resolves an enemy and the win check requires lives above 0. Accept any design that guarantees every round ends in exactly one outcome (for example a third 'Round Over' state, or a check that no enemies remain and none are in flight). Test it with 3 kills and 2 escapes, then 4 and 1, then 0 and 5.",
      "Ask students to name every place start_game() is called from (_ready(), and _input()) and confirm they understand why both call sites needed to exist before this lesson (one for the very first game, one for every subsequent one).",
      "For students comparing engines: Godot's start_game() function and GameMaker's scr_start_game Script resource are conceptually identical - both are 'a named block of reusable code,' just with different syntax.",
    ],
    mistakes: [
      "Calling start_game() directly from GameManager's _ready(): GameManager is a child of Main, its _ready() runs before Main has finished setting up, and Main refuses add_child() while busy, so no enemies appear and no error is shown in the game. The fix is start_game.call_deferred(). Worth a short explanation of ready order for students who hit it.",
      "Game-state drift: the student added start-up code or a game-over effect beyond the handout (extra green-flag scripts, an animation, a color change, a hidden object) and did not reset it on restart. Symptom: a restart that looks almost right. Walk through the Handout's \"Audit Your Game State\" list with them.",
      "Forgetting to add spawned enemies to the \"enemies\" Group, silently breaking the clear-on-restart step (they'll still exist and move, just never found by get_nodes_in_group()).",
      "Placing \"if not game_active: return\" anywhere other than the first line of _process() - enemies would still perform part of their per-frame logic before bailing out.",
      "Comparing event.keycode against a raw integer instead of the KEY_R constant - technically works if the integer happens to match, but fragile and easy to get wrong.",
    ],
  },
  {
    engine: "Unity", file: "Unity",
    teaches: "Splitting a scene object's work into named, reusable methods (ClearEnemies and SpawnEnemies) and reaching it from GameManager with an Inspector reference - the same drag-a-reference-in-the-Inspector move students already used for the enemy prefab in Lesson 5.",
    level1: [
      "GameManager gets a public EnemySpawner spawner field, filled by dragging the spawner object onto the Inspector slot - exactly how the enemy prefab slot was filled in Lesson 5. ResetGame() is an ordinary (non-static) method for that reason: Inspector slots belong to one component instance, so a static method could not use one. ResetGame() is only ever called from GameManager's own Update(), which is an instance method, so nothing else changes.",
      "Splitting the old inline spawn code into ClearEnemies() and SpawnEnemies() as two separate public methods mirrors the same 'name the reusable pieces' idea as Godot's start_game(), just split into two calls instead of one.",
      "GameManager.ResetGame() orchestrates both: clear, reset variables, respawn - the actual 'reset the game' logic lives here, one level up from EnemySpawner's more mechanical spawn/clear methods.",
    ],
    why: [
      "Static variables and singletons are deliberately NOT used for this: GameManager's shared state stays static from earlier lessons, but the one new reference is an Inspector slot, the simplest and most familiar way to connect two objects. Static members and singletons are a good refactoring topic for Semester 2.",
      "Tagging the enemy prefab \"Enemy\" and using FindGameObjectsWithTag() is the standard, idiomatic Unity way to find 'all objects of this kind' - the same category of tool as Godot's Groups, just Unity's own version.",
    ],
    talkingPoints: [
      "Level 3 bug hunt: kills reduce enemies_remaining but escapes only reduce lives, so a round with some kills and some escapes (for example 3 and 2) ends with no enemies left, no win and no loss, and R does nothing because game_active is still true. This started in Lesson 6 (win = all killed) and Lesson 8 (lose = lives gone). A good fix, and the one shown here, is that an escape also resolves an enemy and the win check requires lives above 0. Accept any design that guarantees every round ends in exactly one outcome (for example a third 'Round Over' state, or a check that no enemies remain and none are in flight). Test it with 3 kills and 2 escapes, then 4 and 1, then 0 and 5.",
      "Ask students what would go wrong if the Spawner slot were left empty - ResetGame() would throw a NullReferenceException, a good concrete illustration of why an Inspector reference has to be assigned.",
      "For students who also did Godot: compare the Inspector reference against Godot's get_tree().get_nodes_in_group() (a built-in group query) and GameMaker's direct object names - three different ways of solving 'reach those other objects,' worth naming explicitly.",
    ],
    mistakes: [
      "Game-state drift: the student added start-up code or a game-over effect beyond the handout (extra green-flag scripts, an animation, a color change, a hidden object) and did not reset it on restart. Symptom: a restart that looks almost right. Walk through the Handout's \"Audit Your Game State\" list with them.",
      "Forgetting to tag the enemy prefab \"Enemy\" - ClearEnemies() silently does nothing, and old enemies pile up after each restart.",
      "Leaving the Spawner slot on the GameManager component empty, or filling it on the wrong GameObject, so ResetGame() throws a NullReferenceException.",
      "Placing the gameActive check after movement code in Enemy.cs's Update() instead of as the very first line.",
    ],
  },
  {
    engine: "GameMaker", file: "GameMaker",
    teaches: "GameMaker's actual reusable-function mechanism (a Script resource) - introduced for the first time this course, after eight lessons of event-only code, specifically because this lesson's core idea (one setup routine, called from two places) finally needs it.",
    level1: [
      "scr_start_game is a genuinely new kind of resource for this course - previous lessons kept all logic inside object Events; this is the first time code lives in a standalone, explicitly reusable Script.",
      "with (obj_enemy) { instance_destroy(); } is the correct, idiomatic GML way to destroy every instance of a given object type in one line - a new, useful pattern worth calling out explicitly.",
      "obj_spawner's Create event and obj_gamemanager's Step event (the restart check) both call the exact same start_game() - the two call sites this lesson is built around.",
    ],
    why: [
      "Introducing a Script resource here (rather than earlier) is deliberate: this is the first lesson where the same setup logic genuinely needs to run from two different triggers, which is exactly when a reusable function starts paying for itself.",
      "keyboard_check_pressed() (edge-triggered, fires once per press) is used rather than keyboard_check() (fires every frame held) for the same reason TIC-80 uses keyp() instead of key() this week - worth drawing this cross-engine parallel for students in both.",
    ],
    talkingPoints: [
      "Level 3 bug hunt: kills reduce enemies_remaining but escapes only reduce lives, so a round with some kills and some escapes (for example 3 and 2) ends with no enemies left, no win and no loss, and R does nothing because game_active is still true. This started in Lesson 6 (win = all killed) and Lesson 8 (lose = lives gone). A good fix, and the one shown here, is that an escape also resolves an enemy and the win check requires lives above 0. Accept any design that guarantees every round ends in exactly one outcome (for example a third 'Round Over' state, or a check that no enemies remain and none are in flight). Test it with 3 kills and 2 escapes, then 4 and 1, then 0 and 5.",
      "Ask students who've been in GameMaker since Lesson 2: why didn't we need a Script resource until now? (Every previous lesson's logic only needed to run from one place - this is the first time the same setup genuinely needs two entry points.)",
      "For advanced students: point out that Scripts can also take parameters and return values, previewing more advanced GML use beyond what this lesson needs.",
    ],
    mistakes: [
      "Game-state drift: the student added start-up code or a game-over effect beyond the handout (extra green-flag scripts, an animation, a color change, a hidden object) and did not reset it on restart. Symptom: a restart that looks almost right. Walk through the Handout's \"Audit Your Game State\" list with them.",
      "Writing start_game()'s logic directly in obj_spawner's Create event instead of in the new Script resource, then trying to call it from Step anyway (Create-event code isn't callable from elsewhere the way a Script is).",
      "Placing \"with (obj_enemy) instance_destroy();\" AFTER the spawn loop instead of before it, destroying the freshly-spawned enemies by mistake.",
      "Forgetting the obj_gamemanager. prefix when obj_enemy's Step event checks game_active, causing a variable-scope error (game_active lives on obj_gamemanager, not obj_enemy).",
    ],
  },
  {
    engine: "GDevelop", file: "GDevelop",
    teaches: "GDevelop's \"Or\" condition as the mechanism for reuse this week - rather than a callable function, the SAME event group is simply reached via two different paths (scene start, or a restart flag), which is a genuinely different (and equally valid) way of achieving 'one place, two entry points.'",
    level1: [
      "The Setup group's condition changes from an implicit \"always run at scene start\" to an explicit Or: [At the beginning of the scene] [RestartRequested = 1] - now genuinely reachable two different ways.",
      "RestartRequested is a one-frame flag: the new restart-key event sets it to 1, the Setup group (seeing it's 1) does the reset AND immediately sets it back to 0, so the group doesn't fire again next frame.",
      "GameActive = 1 added as a condition on the movement event is a minimal, single-line change that freezes movement without touching the movement action itself.",
    ],
    why: [
      "This is a genuinely different mechanism from Godot's start_game() function or GameMaker's Script resource - GDevelop's event-condition model doesn't need a named callable routine, just a shared condition that's true at the right two moments. Worth naming this as a real alternative, not a lesser version of 'having a function.'",
      "The RestartRequested flag pattern (set now, act on it this same frame, then clear it) is a common, reusable technique in event-based tools generally, useful again in future GDevelop lessons whenever a one-shot trigger is needed.",
    ],
    talkingPoints: [
      "Ask students to compare this week's Or-condition approach against Godot's function call - both achieve 'one piece of logic, two ways to reach it,' just via different underlying mechanisms. A good moment to reinforce that the CONCEPT (reuse) matters more than which specific tool implements it.",
      "For strong students: ask what would happen if RestartRequested were never reset to 0 inside the Setup group - the group would keep re-running its reset logic every single frame forever, a good illustration of why clearing one-shot flags matters.",
    ],
    mistakes: [
      "Forgetting to reset RestartRequested to 0 inside the Setup group, causing a permanent reset loop the instant a restart happens.",
      "Adding the GameActive = 0 check to the restart-key event but forgetting the Key R is pressed condition (or vice versa) - both conditions are required together.",
      "Placing \"For Each Enemy -> Delete Enemy\" after the spawn action inside Setup instead of before it.",
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    teaches: "The simplest possible version of 'factor setup into a reusable function' this week - a plain Lua function, called once at cart-load time and again from inside TIC() on restart, with no additional engine machinery needed at all.",
    level1: [
      "start_game() is the old top-level setup code, unchanged in content, just wrapped in a function definition and called explicitly (once outside TIC(), and again inside it on restart).",
      "keyp(18) - not key(18) - is essential here: keyp() only returns true on the exact frame R was first pressed, so calling start_game() from it fires exactly once per press, not continuously while held.",
      "The 'if game_active then ... end' wrapper inside the enemy-drawing loop is the simplest possible freeze mechanism available in this engine: a single conditional around existing code, no special API needed.",
    ],
    why: [
      "TIC-80 has no built-in concept of 'the game restarting' - this lesson builds the entire idea from three plain Lua ingredients (a function, a boolean flag, and an edge-triggered key check), which is a good, honest illustration of how much can be built from very few primitives.",
      "Verifying keyp()'s exact behavior (vs. key()) against the official API docs mattered specifically because this engine has had real, documented API-detail mistakes earlier in this course - the same discipline applied there was applied here.",
    ],
    talkingPoints: [
      "Ask students to predict what would happen if key(18) were used instead of keyp(18) - start_game() would be called on every single frame R is held down, likely dozens of times per second, resetting the game continuously and making it un-playable while the key is held.",
      "For advanced students: point out that start_game() being called at the top level (outside TIC()) the first time, and from inside TIC() the second time, is a subtle but real detail worth understanding - Lua code outside any function runs once, immediately, when the cart loads.",
    ],
    mistakes: [
      "Using key(18) instead of keyp(18) - causes the restart to repeat/glitch while the key is held, exactly the bug this lesson's verification step was meant to prevent.",
      "Forgetting to reset the enemies table to a fresh table (enemies={}) inside start_game(), instead trying to reuse/modify the old one, leading to leftover dead entries.",
      "Wrapping only the spr() drawing call in \"if game_active then\" instead of the movement/escape logic too - enemies would stop being drawn but keep moving invisibly.",
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    teaches: "001 Game Creator's actual, wiki-verified reusable-subroutine feature (Custom Event) combined with its generic key-press mechanism (Input Trigger) - both researched and confirmed against the official wiki before writing this lesson, given this engine's history of needing extra care.",
    level1: [
      "Custom Event is a genuinely new concept for this course - previous lessons only ever used Actor Triggers and Input-driven scripts; this is the first reusable, callable block of logic, matching the same underlying need every other engine has this week.",
      "The Input Trigger for R uses the exact same mechanism (an Input Set with a Control Combination) that's been powering movement and shooting since Lesson 2 - no new input concept, just a new key bound to new logic.",
      "The If(game_active = 1) guard added to both existing Triggers is what actually freezes the GAME STATE (score/lives can no longer change) - it does not forcibly stop enemy Automatic Movement, which is a deliberate, documented design choice for this engine (see the Handout's Design Note).",
    ],
    why: [
      "A built-in \"Restart Game\" action does exist in 001 Game Creator (confirmed via community documentation), but its exact reset behavior (what state it clears, what it preserves) isn't clearly documented on the official wiki - teaching the manual, verified Custom-Event approach instead is both safer and better pedagogy, since it makes every piece of the reset explicit and visible to students.",
      "Because Lesson 5 chose design-time placement over a runtime spawn loop for this engine, 'respawning' here means REVIVING and REPOSITIONING the same named Actor copies, not creating new ones - a direct, sensible consequence of that earlier design choice, not a new complication.",
    ],
    talkingPoints: [
      "This is worth explicitly connecting back to Lesson 5's design-time-placement decision: ask students why 001 Game Creator's restart looks different (reviving named copies) from every other engine's (spawning new instances) - the answer traces directly back to that earlier, deliberate choice.",
      "If a student asks why enemies aren't forcibly stopped at game-over the way other engines freeze movement: this is a good moment to discuss that not every constraint needs a workaround - sometimes accepting a small, harmless visual imperfection (a couple of enemies finishing their slide) is the right, honest engineering call, verified-safe rather than reaching for an unconfirmed API.",
    ],
    mistakes: [
      "Forgetting the If(game_active = 1) guard on one of the two Triggers, allowing score or lives to keep changing after the round has technically ended.",
      "Typo'd Actor names inside the ResetGame Custom Event (enemy_1 vs. Enemy_1, etc.) - 001 Game Creator won't necessarily flag this loudly, so it's worth testing that all 5 copies actually reappear.",
      "Setting the Input Trigger's Control Combination to \"Released\" instead of \"Pressed,\" causing the restart to fire when R is let go rather than when it's first pressed down (functionally similar but worth getting consistent with the rest of the course's \"Pressed\" convention).",
    ],
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    teaches: "Extending WBP_HUD (built in Lesson 9) with a third, general-purpose status text element, plus a Custom Event on BP_GameManager (ResetGame) that fully replaces the old inline BeginPlay spawn logic - continuing this course's established push-based widget pattern into genuinely new territory (destroying/respawning actors) this week.",
    level1: [
      "SetStatus() on WBP_HUD is a small, single-purpose Custom Event - same shape as UpdateHUD() from Lesson 9, just for a different piece of persistent text.",
      "ResetGame is BP_GameManager's own reusable routine (it owns Score, Lives, EnemiesRemaining and GameActive) - it fully replaces what BeginPlay used to do inline (spawn enemies), and is now called from two places: BeginPlay (via a wrapping call) and the new IA_Restart Enhanced Input Action Event, whose R mapping was simply added to the existing IMC_Default from Lesson 2 - no new mapping-context wiring needed.",
      "Get All Actors of Class(BP_Enemy) -> Destroy Actor is the standard Blueprint way to find and clear every instance of an Actor class - Unreal's equivalent of Godot's Group query or Unity's tag search.",
    ],
    why: [
      "Reusing WBP_HUD (rather than falling back to a fresh Print String) keeps this lesson consistent with Lesson 9's actual UI investment - the win/lose/restart message now lives in the same persistent widget as score and lives, not a separate, older mechanism.",
      "Wrapping Event Tick's movement logic in Branch(GameActive) is the same pattern used in every scripted engine this week (Godot's early-return, Unity's early-return, GameMaker's early-exit) - Unreal's Blueprint equivalent is simply wrapping the relevant nodes in a Branch's True case instead of an early return, since Blueprint doesn't have a direct 'return' node mid-event the same way code does.",
    ],
    talkingPoints: [
      "For students who did both Unity and Unreal: compare Unity's Inspector reference to the spawner against Unreal's Get All Actors of Class - two different mechanisms for the same underlying need ('reach the enemies').",
      "This is the most Blueprint-node-heavy lesson yet for Unreal - worth acknowledging directly with students working in this engine that today's ResetGame event graph will look busier than previous weeks, and that's expected given everything it's now responsible for.",
    ],
    mistakes: [
      "Wiring the IA_Restart Enhanced Input Action Event's Started pin to call ResetGame directly, without the Branch(!GameActive) gate first - allows restarting mid-round.",
      "Forgetting to call SetStatus with an empty string inside ResetGame, leaving the old win/lose message visible after a restart even though gameplay has resumed.",
      "Leaving old enemy-spawning logic directly in BeginPlay instead of fully moving it into ResetGame - results in enemies being spawned twice at the very first game start.",
    ],
  },
];


const keyLevel2Default = (c) => [
  c + " Level 2 is verification only - the Level 1 code is already correct, so no lines change.",
  c + " Checks: after a win and after a loss, R restores 0/3/5 with no leftover enemies; R during play does nothing.",
];
const keyComment = { "Scratch": "//", "Godot": "#", "TIC-80": "--", "GDevelop": "//", "001GameCreator": "//", "Unreal": "//", "Unity": "//", "GameMaker": "//" };
const keyLevel3 = {
  "Unity": ["// Level 3 visual (GameManager.cs OnGUI): bigger, colored end message.", "void OnGUI() {", "    // ...Level 1 score and lives labels unchanged...", "    if (hasWon) {", "@@         GUIStyle bigStyle = new GUIStyle(GUI.skin.label);", "@@         bigStyle.fontSize = 36;", "@@         bigStyle.normal.textColor = Color.yellow;", "@@         GUI.Label(new Rect(250, 120, 500, 60), \"You Win! Press R to Restart\", bigStyle);", "    }", "}"],
  "GameMaker": ["// Level 3 visual: obj_gamemanager Draw GUI event.", "if (has_won) {", "@@     draw_set_color(c_yellow);", "@@     draw_set_halign(fa_center);", "    draw_text(400, 50, \"You Win! Press R to Restart\");", "@@     draw_set_halign(fa_left);", "@@     draw_set_color(c_white);", "}"],
  "Unreal": ["// Level 3 visual: WBP_HUD Designer view - StatusText Color and Opacity set to a gold color; in the Graph, a Play Animation node on a Fade In animation inside SetStatus.", "WBP_HUD's Event Graph:", "Custom Event SetStatus(StatusText: String) ->", "  Set Text(StatusTextBlock, StatusText)", "@@   Play Animation(FadeIn)"],
  "001GameCreator": ["// Level 3 visual: add to the end of each Message Box trigger branch.", "If(game_active = 0)", "@@   Camera - Fade (Fade Out, 0.5 seconds)"],
  "Scratch": ["// Level 3 visual: add a \"switch costume\" or \"set color effect\" block to the Player's win/lose receive scripts.", "when I receive [you win]", "@@ set [color v] effect to 50", "say [You Win! Press R to Restart] for 3 seconds", "@@ set [color v] effect to 0"],
  "Godot": ["# Level 3 visual: fade the message in (GameManager's script).", "func on_enemy_destroyed():", "    # ...Level 1 body unchanged...", "    if enemies_remaining <= 0:", "        game_active = false", "        get_node(\"../WinText\").text = \"You Win! Press R to Restart\"", "@@         get_node(\"../WinText\").modulate.a = 0.0", "        get_node(\"../WinText\").show()", "@@         create_tween().tween_property(get_node(\"../WinText\"), \"modulate:a\", 1.0, 0.5)"],
  "TIC-80": ["-- Level 3 visual: flash the screen when the round ends (inside TIC(), before cls(0)).", "@@ if not game_active and (time()//200)%2==0 then cls(2) else cls(0) end"],
  "GDevelop": ["// Level 3 visual: replace the instant Show with a Tween.", "(Combat group - win check)", "Condition: EnemiesRemaining <= 0", "Action: GameActive = 0", "Action: Show WinText", "@@ Action: Tween the opacity of WinText from 0 to 255 over 0.5 seconds"],
};
const bugFix = {
 "Scratch": [
  "",
  "// BUG FIX (one good answer): an escape also resolves an enemy. GameManager, \"when I receive [enemy escaped]\":",
  "change lives by -1",
  "@@ change enemies_remaining by -1",
  "if lives = 0 then",
  "    set game_active to false",
  "    broadcast [game over]",
  "@@ else",
  "@@     if enemies_remaining = 0 then",
  "@@         set game_active to false",
  "@@         broadcast [you win]"
 ],
 "Godot": [
  "",
  "# BUG FIX (one good answer): an escape also resolves an enemy. GameManager's script:",
  "func on_enemy_escaped():",
  "    lives -= 1",
  "@@     enemies_remaining -= 1",
  "    update_hud()",
  "    if lives <= 0:",
  "        game_active = false",
  "        get_node(\"../LoseText\").text = \"Game Over - Press R to Restart\"",
  "        get_node(\"../LoseText\").show()",
  "@@     elif enemies_remaining <= 0:",
  "@@         game_active = false",
  "@@         get_node(\"../WinText\").text = \"You Win! Press R to Restart\"",
  "@@         get_node(\"../WinText\").show()"
 ],
 "Unity": [
  "",
  "// BUG FIX (one good answer): an escape also resolves an enemy. GameManager.cs:",
  "public static void RegisterEscape() {",
  "    lives--;",
  "@@     enemiesRemaining--;",
  "    if (lives <= 0) {",
  "        hasLost = true;",
  "        gameActive = false;",
  "@@     } else if (enemiesRemaining <= 0) {",
  "@@         hasWon = true;",
  "@@         gameActive = false;",
  "    }",
  "}"
 ],
 "GameMaker": [
  "",
  "// BUG FIX (one good answer): an escape also resolves an enemy. obj_enemy's Step event, escape branch:",
  "if (y > room_height) {",
  "    obj_gamemanager.__dnd_lives -= 1;",
  "@@     obj_gamemanager.enemies_remaining -= 1;",
  "    if (obj_gamemanager.__dnd_lives <= 0) {",
  "        obj_gamemanager.has_lost = true;",
  "        obj_gamemanager.game_active = false;",
  "@@     } else if (obj_gamemanager.enemies_remaining <= 0) {",
  "@@         obj_gamemanager.has_won = true;",
  "@@         obj_gamemanager.game_active = false;",
  "    }",
  "    instance_destroy();",
  "}",
  "// GML Visual: add a Set Variable block (Target obj_gamemanager, enemies_remaining, Value -1, Relative checked) after the Set Lives block, and an Else branch holding an If Variable (enemies_remaining <= 0) that sets has_won and game_active."
 ]
};
Object.keys(bugFix).forEach(k => { keyLevel3[k] = (keyLevel3[k] || []).concat(bugFix[k]); });
keyData.forEach(k => {
  const h = handoutData.find(x => x.file === k.file);
  if (!k.level1Code && h && h.level1Code) k.level1Code = h.level1Code;
  if (!k.level2Code) k.level2Code = keyLevel2Default(keyComment[k.file] || "//");
  if (!k.level3Code && keyLevel3[k.file]) k.level3Code = keyLevel3[k.file];
});
keyData.forEach(d => jobs.push([buildKey(d), `Lesson_10_Key_${d.file}.docx`]));

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
