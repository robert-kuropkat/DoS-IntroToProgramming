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
      headers: { default: brand.pageHeader("Lesson 12 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 12 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("A New Phase Begins"),
        p("Lesson 11 reached Level 2 Complete - the mechanics are done. From here through Lesson 15, every lesson is about making that same loop feel good: sound this week, animation in Lesson 13, difficulty scaling in Lesson 14, and a final polish pass in Lesson 15. This is a lighter, more satisfying lesson than most - little new mechanical difficulty, mostly wiring up cues students can immediately hear."),

        h2("This Week: The Sound Payoff"),
        p("Since Lesson 2, five different lessons have quietly offered an optional 'add a sound' Level 3 idea: a movement blip (Lesson 2), a shoot sound (Lesson 3), a hit/explosion sound (Lesson 4), and a victory sound (Lesson 6). Some students built these; many didn't. This week makes four sounds mandatory for everyone - shoot, explosion, wave clear, and game over - matching the capstone's Level 3 feature list exactly. For students who already built the optional versions, today is mostly verification plus two new sounds. For everyone else, this lesson gives full instructions as if starting fresh, the same self-contained approach every lesson uses."),
        checkbox("Read the Handout's Setup and Level 1 sections before class - for students who skipped the earlier optional sound ideas, this is effectively new material, not just a quick add-on."),

        h2("Self-Contained Lesson Assets"),
        p("Two new sound effects are provided in this lesson's Assets folder: sfx_wave.wav (a short, upbeat wave-clear cue) and sfx_lose.wav (a short, low game-over cue) - deliberately distinct from each other and from the existing sfx_shoot.wav, sfx_hit.wav, and sfx_win.wav so all five are easy to tell apart by ear."),
        checkbox("Confirm sfx_wave.wav and sfx_lose.wav are present and playable before class (a quick double-click test on the demo laptop is enough)."),

        h2("Continuity From Lesson 11"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 11, not create a new one."),
        checkbox("If a student's Lesson 11 project is lost or broken, have a pre-built \"Lesson 11 complete\" starter available per engine (from the Lesson 11 Teacher's Keys) so they aren't blocked."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop's Lesson 10-11 exercises (restart flow, multiple waves) still work for all 8 engines."),
        checkbox("Test a full round on every engine before class: a shot, a hit, an early wave clear, and a loss should each play a distinct, correctly-timed sound with no missing or doubled playback."),
        checkbox("Specifically test that the wave-clear sound and the win sound never both play on the same frame (the true-win branch and the wave-advance branch are mutually exclusive, but it's an easy mistake to place a sound call outside the if/else split)."),
        checkbox("001 Game Creator and Unreal (Windows demo laptop): confirm newly imported sfx_wave.wav / sfx_lose.wav assets are recognized correctly by each engine's audio import step."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies, and headphones or speakers for the demo (this is the first lesson where audio output actually matters for the lecture demo)."),

        h2("Known Gotchas This Week"),
        bullet("Students who already built an optional sound idea in an earlier lesson may assume today needs nothing from them - point out explicitly that today still adds two brand-new sounds (wave clear, game over) they haven't built before."),
        bullet("The most likely mistake: placing a sound call outside the if/else split from Lesson 11, causing two round-ending sounds to play at once, or the wrong one to play."),
        bullet("A close second: forgetting that Scratch's \"start sound\" and most other engines' one-shot sound calls don't block gameplay - this is normal and expected, not a bug to fix."),
        bullet("Encourage a genuinely playful few minutes of testing this week - hearing a game respond to your actions is one of the most satisfying moments in the whole course, and it's worth letting students enjoy it."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_12_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 12: Hearing It Happen", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: your game finally sounds like a game. Four moments - firing, hitting, clearing a wave, and losing - each get their own distinct cue.", { bold: true }),
    p("Open your own [YourName]Game project from Lesson 11 - don't start a new one.", { bold: true }),
    p("If you already built an optional sound idea in an earlier lesson (Lessons 2, 3, 4, or 6), keep it - today reuses it directly instead of replacing it. If you didn't, today's steps are written as if starting fresh.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("Four sounds fire at the correct moment, at least once each: a shot, a hit/explosion, a wave-clear (on any non-final wave), and a game over. No silence, no obviously wrong triggers.", { italics: true }),
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
    p("All four sounds fire reliably every time, with no missing or duplicate playback, and each is clearly distinguishable from the others by ear.", { italics: true }),
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
      headers: { default: brand.pageHeader("Lesson 12 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 11.",
      "No new variables needed - today only adds sound blocks to code you already have.",
      "Confirm sfx_wave and sfx_lose are uploaded to the Enemy sprite's Sounds tab (Scratch sound assets are scoped per-sprite, so upload them to whichever sprite's script will play them).",
    ],
    level1: [
      "On the Player sprite's space-key script, confirm \"start sound [sfx_shoot]\" already fires next to \"create clone of [Bullet]\" - if it's missing (you may have skipped this as an optional idea back in Lesson 3), add it now.",
      "On the Enemy sprite's touching-Bullet branch, add \"start sound [sfx_hit]\" right after \"change enemies_left by -1\" - this plays once per hit, no matter which branch runs next.",
      "In that branch's else (wave-advance) case, add \"start sound [sfx_wave]\" right before \"broadcast [spawn wave]\".",
      "On the Enemy sprite's escape branch, add \"start sound [sfx_lose]\" right after \"change lives by -1\", before \"broadcast [game over]\".",
      "Click the stage and play through a full round - shooting, hitting, an early wave clear, and losing a life to zero should each sound distinct.",
    ],
    level1Code: [
      "Enemy's touching-Bullet branch (revised):",
      "if <touching Bullet?> then",
      "  change score by 10",
      "  change enemies_left by -1",
      "  start sound [sfx_hit]",
      "  if <enemies_left = 0> then",
      "    if <wave = total_waves> then",
      "      set game_active to 0",
      "      broadcast [you win]",
      "    else",
      "      change wave by 1",
      "      set enemies_left to 5",
      "      start sound [sfx_wave]",
      "      broadcast [spawn wave]",
      "  delete this clone",
      "",
      "Enemy's escape branch (revised):",
      "change lives by -1",
      "if <lives = 0> then",
      "  start sound [sfx_lose]",
      "  set game_active to 0",
      "  broadcast [game over]",
    ],
    level2: [
      "Play a full round: shoot, hit an enemy, clear wave 1 or 2, and separately let lives hit 0 - all four sounds should be clearly different from each other.",
      "If you built Lesson 6's optional victory sound, confirm it still plays on the true win (wave 3's clear) and doesn't also play the wave-clear sound on the same frame.",
      "Rapid shooting and a hit landing in the same instant should layer, not cut each other off - \"start sound\" doesn't block or loop, so this should already work correctly.",
    ],
    level3: [
      "Audio: if you haven't already, add Lesson 6's optional win sound now - \"start sound [sfx_win]\" right where \"broadcast [you win]\" fires. It slots into the exact same branch as today's wave-clear sound, just the opposite outcome.",
      "Audio: design a second, alternate hit sound and randomly pick between the two each time (pick random 1 to 2), so repeated hits don't sound completely identical.",
      "Structure: notice all four sounds live in the exact spot representing their event - shoot with the shot, hit with the hit, wave/win with the branch that decides them, lose with the escape that ends the round. None of them live in a generic \"forever\" script.",
    ],
    trouble: [
      "No sound at all: confirm the correct sfx_ file was uploaded via the Sounds tab for the sprite the \"start sound\" block lives on - Scratch sound assets are scoped per-sprite, not shared automatically.",
      "Wave-clear and win sounds both play on wave 3: confirm \"start sound [sfx_wave]\" is only inside the else branch, not before the if/else splits.",
      "Lose sound plays repeatedly: confirm the check uses \"if lives = 0\" as a one-time branch (matching Lesson 8's pattern), not something re-checked every frame.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 11.",
      "Add a new AudioStreamPlayer2D node to your main scene, named RoundSound - one shared player for the round-ending sounds (win, wave-clear, lose), since only one of them ever plays at a time.",
      "Import sfx_wave.wav and sfx_lose.wav (Resources > Import > Sound) alongside your existing sfx_shoot.wav, sfx_hit.wav, and (if you built it) sfx_win.wav.",
    ],
    level1: [
      "On the Player, confirm $AudioStreamPlayer2D.play() already fires in your fire function next to bullet instantiation (from Lesson 3) - add it now if it's missing.",
      "On the Enemy scene's script, confirm its own AudioStreamPlayer2D plays sfx_hit.wav when destroyed (from Lesson 4) - add it now if it's missing.",
      "At the top of your main script, preload SND_WAVE and SND_LOSE constants (and SND_WIN, if you have it).",
      "In _on_enemy_destroyed()'s nested if (from Lesson 11), assign RoundSound's stream and call .play() in each branch - SND_WIN for the true-win case (skip if you didn't build it), SND_WAVE for the wave-advance case.",
      "In _on_enemy_escaped()'s lose branch, assign RoundSound's stream to SND_LOSE and call .play().",
      "Run the scene and play through a full round - shoot, hit, an early wave clear, and a loss should each sound distinct.",
    ],
    level1Code: [
      "# Top of main script:",
      "const SND_WAVE = preload(\"res://sfx_wave.wav\")",
      "const SND_LOSE = preload(\"res://sfx_lose.wav\")",
      "const SND_WIN = preload(\"res://sfx_win.wav\")  # only if you built Lesson 6's optional idea",
      "",
      "func _on_enemy_destroyed():",
      "    enemies_remaining -= 1",
      "    score += 10",
      "    update_hud()",
      "    if enemies_remaining <= 0:",
      "        if wave >= total_waves:",
      "            game_active = false",
      "            $WinText.show()",
      "            $RoundSound.stream = SND_WIN",
      "            $RoundSound.play()",
      "        else:",
      "            wave += 1",
      "            enemies_remaining = 5",
      "            update_hud()",
      "            $RoundSound.stream = SND_WAVE",
      "            $RoundSound.play()",
      "            spawn_wave()",
      "",
      "func _on_enemy_escaped():",
      "    lives -= 1",
      "    update_hud()",
      "    if lives <= 0:",
      "        game_active = false",
      "        $LoseText.show()",
      "        $RoundSound.stream = SND_LOSE",
      "        $RoundSound.play()",
    ],
    level2: [
      "Play a full round: shoot, hit, an early wave clear, and a loss each produce a distinct sound.",
      "If you built the optional win sound, confirm it still plays on the true win, and SND_WAVE doesn't also fire on that same frame.",
      "Confirm you're setting .stream immediately before every .play() call on RoundSound, not relying on whatever it was set to last.",
    ],
    level3: [
      "Audio: if you haven't already, add the optional win sound now - it slots into the exact same $RoundSound.play() pattern as today's wave/lose sounds, just a third constant.",
      "Audio: add slight pitch variation to hits by setting the Enemy's AudioStreamPlayer2D pitch_scale to a random value (e.g. randf_range(0.9, 1.1)) right before each .play() call.",
      "Structure: notice RoundSound is shared by three different outcomes (win, wave-clear, lose) - one node, three interchangeable clips, instead of three separate always-idle nodes.",
    ],
    trouble: [
      "Wrong sound plays: confirm .stream is being reassigned immediately before each .play() call, not left over from a previous event.",
      "No sound on wave-clear: confirm the SND_WAVE play() call is actually inside the else branch, not after the if/else.",
      "Lose sound plays repeatedly: confirm the lose branch is still guarded by \"if lives <= 0\" as a one-time check (from Lesson 8), not re-triggering every frame lives stays at 0.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Open your [YourName]Game project from Lesson 11.",
      "Add an AudioSource component to the GameManager GameObject itself (Play On Awake unchecked) - one shared player for win/wave/lose sounds.",
      "Import sfx_wave.wav and sfx_lose.wav alongside your existing clips.",
    ],
    level1: [
      "In GameManager.cs, add public AudioSource roundAudioSrc and public AudioClip winClip, waveClip, loseClip fields, then drag the AudioSource component and the three clips onto them in the Inspector.",
      "If GameManager doesn't already expose itself via a public static GameManager instance (assigned in Awake()), add that now - the same singleton pattern Lesson 10 introduced for EnemySpawner, needed here so the static RegisterKill()/RegisterEscape() methods can reach the Inspector-assigned AudioSource.",
      "Confirm the Player's AudioSource already plays sfx_shoot.wav on fire (from Lesson 3) - add it now if it's missing.",
      "Confirm the Enemy prefab's AudioSource already plays sfx_hit.wav on destroy (from Lesson 4) - add it now if it's missing.",
      "In RegisterKill()'s nested if (from Lesson 11), call instance.roundAudioSrc.PlayOneShot() with winClip in the true-win branch and waveClip in the wave-advance branch.",
      "In RegisterEscape()'s lose branch, call instance.roundAudioSrc.PlayOneShot(instance.loseClip).",
      "Press Play and play through a full round - shoot, hit, an early wave clear, and a loss should each sound distinct.",
    ],
    level1Code: [
      "// GameManager.cs:",
      "public static GameManager instance;",
      "public AudioSource roundAudioSrc;",
      "public AudioClip winClip, waveClip, loseClip;",
      "",
      "void Awake() { instance = this; }",
      "",
      "public static void RegisterKill() {",
      "    enemiesRemaining--;",
      "    score += 10;",
      "    if (enemiesRemaining <= 0) {",
      "        if (wave >= totalWaves) {",
      "            hasWon = true;",
      "            instance.roundAudioSrc.PlayOneShot(instance.winClip);",
      "        } else {",
      "            wave++;",
      "            enemiesRemaining = 5;",
      "            instance.roundAudioSrc.PlayOneShot(instance.waveClip);",
      "            EnemySpawner.instance.SpawnEnemies();",
      "        }",
      "    }",
      "}",
      "",
      "public static void RegisterEscape() {",
      "    lives--;",
      "    if (lives <= 0) {",
      "        hasLost = true;",
      "        instance.roundAudioSrc.PlayOneShot(instance.loseClip);",
      "    }",
      "}",
    ],
    level2: [
      "Play a full round: shoot, hit, an early wave clear, and a loss each produce a distinct sound.",
      "If you built the optional win sound, confirm it still plays on the true win, and waveClip doesn't also fire on that same frame.",
      "PlayOneShot() lets overlapping sounds layer instead of cutting each other off - confirm rapid shooting and a hit landing together both play correctly.",
    ],
    level3: [
      "Audio: if you haven't already, add the optional win sound now - it slots into the exact same PlayOneShot() pattern as today's wave/lose sounds, just a third clip.",
      "Audio: add slight pitch variation to hits by setting the Enemy's AudioSource.pitch to a random value (e.g. Random.Range(0.9f, 1.1f)) right before Play().",
      "Structure: notice roundAudioSrc is shared by three different outcomes (win, wave-clear, lose) via PlayOneShot() - one AudioSource, three interchangeable clips.",
    ],
    trouble: [
      "NullReferenceException on instance.roundAudioSrc: confirm GameManager's Awake() actually runs before any RegisterKill()/RegisterEscape() call, and that instance is assigned there.",
      "No sound on wave-clear: confirm the waveClip PlayOneShot() call is actually inside the else branch, not after the if/else.",
      "Lose sound plays repeatedly: confirm the lose branch is still guarded by \"if (lives <= 0)\" as a one-time check (from Lesson 8), not re-triggering every frame.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same three-clip pattern works in a Script Graph instead of C#.",
      steps: [
        "On the win/wave Branch node from Lesson 11, add an Audio Play Clip One Shot node to each outcome, referencing a winClip/waveClip variable.",
        "On the lose graph, add a matching Audio Play Clip One Shot node referencing a loseClip variable.",
      ],
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: [
      "Open your [YourName]Game project from Lesson 11.",
      "Create three new Sound resources: sfx_wave and sfx_lose (and sfx_win, if you don't already have it from Lesson 6).",
    ],
    level1: [
      "Confirm sfx_shoot already plays via audio_play_sound() in your shoot input code (Lesson 3) - add it now if it's missing.",
      "Confirm sfx_hit already plays via audio_play_sound() in obj_bullet's Collision event (Lesson 4/5) - add it now if it's missing, right after the score/enemies_remaining updates.",
      "In obj_bullet's Collision event's nested if (from Lesson 11), add audio_play_sound(sfx_win, 1, false) in the true-win branch and audio_play_sound(sfx_wave, 1, false) in the else branch.",
      "In obj_enemy's Step event's escape branch (where lives decrements), add audio_play_sound(sfx_lose, 1, false) inside the lives <= 0 check.",
      "Press F5 (Run) and play through a full round - shoot, hit, an early wave clear, and a loss should each sound distinct.",
    ],
    level1Code: [
      "// obj_bullet Collision event (revised):",
      "obj_spawner.enemies_remaining -= 1;",
      "obj_spawner.score += 10;",
      "audio_play_sound(sfx_hit, 1, false);",
      "if (obj_spawner.enemies_remaining <= 0) {",
      "    if (obj_spawner.wave >= obj_spawner.total_waves) {",
      "        obj_spawner.game_active = false;",
      "        obj_spawner.has_won = true;",
      "        audio_play_sound(sfx_win, 1, false);",
      "    } else {",
      "        obj_spawner.wave += 1;",
      "        obj_spawner.enemies_remaining = 5;",
      "        audio_play_sound(sfx_wave, 1, false);",
      "        obj_spawner.spawn_wave();",
      "    }",
      "}",
      "instance_destroy();",
      "",
      "// obj_enemy Step event, escape branch (revised):",
      "obj_spawner.lives -= 1;",
      "if (obj_spawner.lives <= 0) {",
      "    obj_spawner.game_active = false;",
      "    obj_spawner.has_lost = true;",
      "    audio_play_sound(sfx_lose, 1, false);",
      "}",
    ],
    level2: [
      "Play a full round: shoot, hit, an early wave clear, and a loss each produce a distinct sound.",
      "If you built the optional win sound, confirm it still plays on the true win, and sfx_wave doesn't also fire on that same frame.",
      "audio_play_sound() calls layer naturally - confirm rapid shooting and a hit landing together both play correctly, with no sound cutting another off.",
    ],
    level3: [
      "Audio: if you haven't already, add the optional win sound now - it slots into the exact same audio_play_sound() pattern as today's wave/lose sounds, just a third resource.",
      "Audio: add slight pitch variation to hits with audio_play_sound(sfx_hit, 1, false) followed by audio_sound_pitch(inst, random_range(0.9, 1.1)), where inst is the sound's returned instance handle.",
      "Structure: group all your Sound resources under one folder in the Asset Browser for organization, now that there are five of them.",
    ],
    trouble: [
      "No sound on wave-clear: confirm the sfx_wave call is actually inside the else branch, not after the if/else.",
      "Lose sound plays repeatedly: confirm the lose check is still guarded by \"if (obj_spawner.lives <= 0)\" as a one-time check (from Lesson 8), not re-triggering every Step.",
      "Wrong instance's variables referenced: confirm the obj_spawner. prefix is present throughout the revised blocks, since this code still runs from obj_bullet and obj_enemy.",
    ],
    resource: "Student Resource List > GameMaker section, or manual.gamemaker.io",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same sound calls work as drag-and-drop blocks instead of typed code.",
      steps: [
        "On obj_bullet's Collision event (GML Visual), add an \"Audio Play Sound\" block (sfx_hit) right after your existing score/enemies_remaining Variable blocks.",
        "Add two more \"Audio Play Sound\" blocks to the win-check's If Variable block from Lesson 11 - sfx_win on the True output, sfx_wave on the False output.",
        "On obj_enemy's escape branch, add an \"Audio Play Sound\" block (sfx_lose) inside the lives-reaches-zero If Variable block.",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 11.",
      "No new resource-import step needed - GDevelop's \"Play a sound\" action references .wav files directly, the same as Lesson 3.",
    ],
    level1: [
      "Confirm the shoot event already has \"Action: Play a sound (sfx_shoot.wav)\" (Lesson 3) - add it now if it's missing.",
      "On the bullet-enemy collision event, add \"Action: Play a sound (sfx_hit.wav)\" alongside the existing Score/EnemiesRemaining actions.",
      "On the full-win event (Wave >= TotalWaves, from Lesson 11), add \"Action: Play a sound (sfx_win.wav)\".",
      "On the advance-wave event (Wave < TotalWaves, from Lesson 11), add \"Action: Play a sound (sfx_wave.wav)\".",
      "On the lose event (Lives <= 0, from Lesson 8/10), add \"Action: Play a sound (sfx_lose.wav)\".",
      "Click Preview and play through a full round - shoot, hit, an early wave clear, and a loss should each sound distinct.",
    ],
    level1Code: [
      "Bullet-Enemy collision event (revised):",
      "Condition: Bullet is in collision with Enemy",
      "Action: Score = Score + 10",
      "Action: EnemiesRemaining = EnemiesRemaining - 1",
      "Action: Play a sound (sfx_hit.wav)",
      "Action: Delete Bullet",
      "Action: Delete Enemy",
      "",
      "Full-win event (revised):",
      "Condition: EnemiesRemaining <= 0",
      "Condition: Wave >= TotalWaves",
      "Action: GameActive = 0",
      "Action: Show WinText",
      "Action: Play a sound (sfx_win.wav)",
      "",
      "Advance-wave event (revised):",
      "Condition: EnemiesRemaining <= 0",
      "Condition: Wave < TotalWaves",
      "Action: Wave = Wave + 1",
      "Action: EnemiesRemaining = 5",
      "Action: Play a sound (sfx_wave.wav)",
      "Action: Repeat 5 times -> Create Enemy",
      "",
      "Lose event (revised):",
      "Condition: Lives <= 0",
      "Action: GameActive = 0",
      "Action: Show LoseText",
      "Action: Play a sound (sfx_lose.wav)",
    ],
    level2: [
      "Play a full round: shoot, hit, an early wave clear, and a loss each produce a distinct sound.",
      "If you built the optional win sound, confirm it still plays on the true win, and sfx_wave.wav doesn't also fire on that same frame - the full-win and advance-wave events' Wave conditions should keep them mutually exclusive.",
      "\"Play a sound\" fires once per action - confirm rapid shooting and a hit landing together both play correctly, with no sound cutting another off.",
    ],
    level3: [
      "Audio: if you haven't already, add the optional win sound now - it slots into the exact same \"Play a sound\" action as today's wave/lose sounds, just a different event.",
      "Audio: add a Random object action to pick between two hit sound files each time, so repeated hits don't sound completely identical.",
      "Structure: for extra practice, try moving the repeated Play-a-sound actions into a Function (Project > Functions) that takes a filename parameter, callable from every event that needs one - not required, but a natural extension of Lesson 11's optional Functions idea.",
    ],
    trouble: [
      "No sound on wave-clear: confirm the advance-wave event's Play-a-sound action is present and its Wave < TotalWaves condition is correct.",
      "Both win and wave-clear sounds play together: confirm the two events' Wave conditions are exact opposites (>= vs. <), matching Lesson 11's original troubleshooting note.",
      "Lose sound plays repeatedly: confirm the lose event's Lives <= 0 condition is a one-time trigger tied to GameActive, matching Lesson 8/10's pattern.",
    ],
    resource: "Student Resource List > GDevelop section, or wiki.gdevelop.io",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: [
      "Open your [YourName]Game project from Lesson 11.",
      "Design two new short SFX in the SFX Editor (F4): slot 4 (wave clear) and slot 5 (game over) - distinct from your existing slot 1 (shoot), slot 2 (hit), and slot 3 (win, if built).",
    ],
    level1: [
      "Confirm sfx(1) already fires in your shoot branch (Lesson 3) - add it now if it's missing.",
      "In the bullets loop's hit branch, confirm sfx(2) already fires when a hit is detected (Lesson 4/5) - add it now if it's missing.",
      "In that same branch's nested if (from Lesson 11), add sfx(3) in the true-win case (only if you built it) and sfx(4) in the else (wave-advance) case.",
      "In the enemy escape-cleanup code where lives decrements, add sfx(5) inside the lives<=0 check.",
      "Press Esc, type run, press Enter, then play through a full round - shoot, hit, an early wave clear, and a loss should each sound distinct.",
    ],
    level1Code: [
      "-- bullets loop's hit branch (revised):",
      "score = score + 10",
      "enemies_remaining = enemies_remaining - 1",
      "sfx(2)",
      "if enemies_remaining <= 0 then",
      " if wave >= total_waves then",
      "  game_active = false",
      "  has_won = true",
      "  sfx(3)",
      " else",
      "  wave = wave + 1",
      "  enemies_remaining = 5",
      "  sfx(4)",
      "  spawn_wave()",
      " end",
      "end",
      "",
      "-- enemy escape check (revised):",
      "lives = lives - 1",
      "if lives <= 0 then",
      " game_active = false",
      " has_lost = true",
      " sfx(5)",
      "end",
    ],
    level2: [
      "Play a full round: shoot, hit, an early wave clear, and a loss each produce a distinct sound.",
      "If you built the optional win sound, confirm it still plays on the true win, and sfx(4) doesn't also fire on that same frame.",
      "None of today's five sfx() calls need an sfx(-1) stop call or a replay-guard flag - each lives inside a branch that only executes once per actual hit or escape event, not a per-frame poll.",
    ],
    level3: [
      "Audio: if you haven't already, design slot 3 and add sfx(3) now - it slots into the exact same nested if as today's wave sound, just the opposite branch.",
      "Audio: design a second hit sound in an empty slot and alternate between the two using math.random(0,1), so repeated hits don't sound completely identical.",
      "Structure: label your five SFX Editor slots with a comment at the top of your code (e.g. -- slots: 1=shoot 2=hit 3=win 4=wave 5=lose) so the mapping stays clear as the project grows.",
    ],
    trouble: [
      "Wrong sound plays: confirm each sfx() call uses the slot number matching what you designed in the SFX Editor - it's easy to mix up 3, 4, and 5.",
      "No sound on wave-clear: confirm sfx(4) is actually inside the else branch, not after the if/else.",
      "Lose sound plays repeatedly: confirm the escape check still fires from inside the collision/escape-detection code (once per actual escape), not a per-frame check in TIC() itself.",
    ],
    resource: "Student Resource List > TIC-80 section, or github.com/nesbox/TIC-80/wiki",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 11.",
      "Import sfx_wave.wav and sfx_lose.wav (Resources > Import > Sound) alongside your existing sfx_shoot.wav, sfx_hit.wav, and (if built) sfx_win.wav.",
    ],
    level1: [
      "Confirm your Spacebar Input Trigger already has Play Sound(sfx_shoot.wav) (Lesson 3) - add it now if it's missing.",
      "On the enemy Actor's \"Collided with actor\" Trigger, add Play Sound(sfx_hit.wav) as the first action, before the win-check nested If (from Lesson 11).",
      "In that same nested If, add Play Sound(sfx_win.wav) in the True branch (only if you built it) and Play Sound(sfx_wave.wav) in the Else branch.",
      "On the enemy Actor's \"Left/Entered Camera\" Trigger (gated by If(game_active=1), from Lesson 10), add Play Sound(sfx_lose.wav) inside the lives<=0 branch.",
      "Save, then Play/Test and play through a full round - shoot, hit, an early wave clear, and a loss should each sound distinct.",
    ],
    level1Code: [
      "enemy Actor Trigger: Collided with actor (revised) ->",
      "  If(game_active = 1)",
      "    Variable – Modify Value(score, Add, 10)",
      "    Variable – Modify Value(enemies_remaining, Subtract, 1)",
      "    Play Sound(sfx_hit.wav)",
      "    If(enemies_remaining <= 0)",
      "      If(wave >= total_waves)",
      "        Variable – Modify Value(game_active, Set, 0)",
      "        Play Sound(sfx_win.wav)",
      "        Message Box(\"You Win! - Press R to Restart\")",
      "      Else",
      "        Variable – Modify Value(wave, Add, 1)",
      "        Variable – Modify Value(enemies_remaining, Set, 5)",
      "        Play Sound(sfx_wave.wav)",
      "        Call Custom Event(SpawnWave)",
      "",
      "enemy Actor Trigger: Left/Entered Camera (revised) ->",
      "  If(game_active = 1)",
      "    Variable – Modify Value(lives, Subtract, 1)",
      "    If(lives <= 0)",
      "      Variable – Modify Value(game_active, Set, 0)",
      "      Play Sound(sfx_lose.wav)",
      "      Message Box(\"Game Over - Press R to Restart\")",
    ],
    level2: [
      "Play a full round: shoot, hit, an early wave clear, and a loss each produce a distinct sound.",
      "If you built the optional win sound, confirm it still plays on the true win, and sfx_wave.wav doesn't also fire on that same frame.",
      "Play Sound actions layer naturally - confirm rapid shooting and a hit landing together both play correctly, with no sound cutting another off.",
    ],
    level3: [
      "Audio: if you haven't already, add the optional win sound now - it slots into the exact same Play Sound action as today's wave/lose sounds, just a different branch.",
      "Audio: explore the Sound action's volume/pitch parameters (if available in your version) to add slight variation to the hit sound.",
      "Structure: notice Play Sound actions live directly inside the exact Trigger representing their event - no separate \"sound manager\" Actor is needed for a game this size.",
    ],
    trouble: [
      "No sound on wave-clear: confirm the Play Sound(sfx_wave.wav) action is actually inside the Else branch of the nested If, not before it.",
      "Lose sound plays repeatedly: confirm the lose branch is still guarded by If(lives <= 0) as a one-time check inside If(game_active = 1), matching Lesson 10's pattern.",
      "Imported sound doesn't play: confirm the file finished importing (check the Resources panel) before referencing it in a Play Sound action.",
    ],
    resource: "001gamecreator.com/wiki/, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 11.",
      "Import sfx_wave.wav and sfx_lose.wav as Sound Wave assets, alongside your existing sfx_shoot.wav, sfx_hit.wav, and (if built) sfx_win.wav.",
    ],
    level1: [
      "Confirm your IA_Fire Enhanced Input Action Event already has a Play Sound 2D(sfx_shoot.wav) node chained after Spawn Actor from Class (Lesson 3) - add it now if it's missing.",
      "In your bullet-enemy collision handling, confirm Play Sound 2D(sfx_hit.wav) already fires after the successful Cast (Lesson 4) - add it now if it's missing, before the win-check.",
      "In HandleEnemyDestroyed's nested Branch (from Lesson 11), add Play Sound 2D(sfx_win.wav) in the True case (only if you built it) and Play Sound 2D(sfx_wave.wav) in the False case.",
      "In HandleEnemyEscaped, add Play Sound 2D(sfx_lose.wav) inside the Lives <= 0 Branch.",
      "Play in Editor and play through a full round - shoot, hit, an early wave clear, and a loss should each sound distinct.",
    ],
    level1Code: [
      "Custom Event HandleEnemyDestroyed (revised) ->",
      "  EnemiesRemaining = EnemiesRemaining - 1",
      "  Score = Score + 10",
      "  Play Sound 2D(sfx_hit.wav)",
      "  Call UpdateHUD(HUDWidget, Score, Lives, Wave)",
      "  Branch(EnemiesRemaining <= 0) -> True:",
      "    Branch(Wave >= TotalWaves) -> True:",
      "      GameActive = false",
      "      Play Sound 2D(sfx_win.wav)",
      "      Call SetStatus(HUDWidget, \"You Win! Press R to Restart\")",
      "    False:",
      "      Wave = Wave + 1",
      "      EnemiesRemaining = 5",
      "      Play Sound 2D(sfx_wave.wav)",
      "      Call UpdateHUD(HUDWidget, Score, Lives, Wave)",
      "      Call SpawnWave",
      "",
      "Custom Event HandleEnemyEscaped (revised) ->",
      "  Lives = Lives - 1",
      "  Branch(Lives <= 0) -> True:",
      "    GameActive = false",
      "    Play Sound 2D(sfx_lose.wav)",
      "    Call SetStatus(HUDWidget, \"Game Over - Press R to Restart\")",
    ],
    level2: [
      "Play a full round: shoot, hit, an early wave clear, and a loss each produce a distinct sound.",
      "If you built the optional win sound, confirm it still plays on the true win, and sfx_wave.wav doesn't also fire on that same frame.",
      "Play Sound 2D nodes layer naturally - confirm rapid shooting and a hit landing together both play correctly.",
    ],
    level3: [
      "Audio: if you haven't already, add the optional win sound now - it slots into the exact same Play Sound 2D pattern as today's wave/lose sounds, just a different Branch case.",
      "Audio: use a Random Float in Range node to vary the Pitch Multiplier parameter on the hit sound's Play Sound 2D node for slight variety.",
      "Structure: notice every Play Sound 2D node this week lives directly inside the exact Custom Event representing its trigger - no separate audio-manager Blueprint needed for a game this size.",
    ],
    trouble: [
      "No sound on wave-clear: confirm the sfx_wave.wav Play Sound 2D node is actually wired inside the False case of the nested Branch, not before it.",
      "Lose sound plays repeatedly: confirm the lose Branch is still a one-time check tied to GameActive, matching Lesson 10's pattern.",
      "No sound at all: confirm the imported asset is actually a Sound Wave (not left as a raw .wav in the Content Browser needing re-import).",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same five sound calls work in C++ instead of Blueprint nodes.",
      steps: [
        "Add UGameplayStatics::PlaySound2D(GetWorld(), HitSound) at the same spot as the Blueprint hit sound, and matching calls for WinSound/WaveSound inside your kill-handling function's nested if/else.",
        "Add UGameplayStatics::PlaySound2D(GetWorld(), LoseSound) inside your escape-handling function's lives<=0 branch.",
      ],
      note: "This mirrors the Blueprint path's Play Sound 2D nodes exactly - same five moments, just called from code.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_12_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 12 Teacher's Key: Hearing It Happen", bold: true, size: 30, color: brand.NAVY })] }),
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
      headers: { default: brand.pageHeader("Lesson 12 Key - " + data.engine + " (Instructor Only)") },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    teaches: "Formalizing five lessons' worth of optional 'add a sound' ideas (Lessons 2, 3, 4, 6) into two mandatory new sounds (wave clear, game over), placed in the exact branches Lesson 11 already built - no new Scratch concept, just recognizing where an event already happens and adding one block there.",
    level1: [
      "sfx_hit now lives directly in the touching-Bullet branch rather than a separate broadcast - simpler than the original Lesson 4 idea, matching how this branch has been restructured since Lesson 5/7.",
      "sfx_wave and sfx_lose are the two genuinely new sounds this week, placed at the exact two spots Lesson 11 and Lesson 8 already built: the else branch of the wave-check, and the lives-reaches-zero check.",
      "If a student built Lesson 6's optional sfx_win, it needs zero changes - it already lives in the correct (true-win) branch of the same nested if.",
    ],
    why: [
      "Every sound this week lives inside a branch that only executes once per actual event (a hit, a wave clear, a loss) - Scratch's clone-and-broadcast model has made every trigger point naturally one-shot since Lesson 5, so there's no risk of a sound replaying every frame.",
      "Keeping sfx_wave and sfx_win in opposite branches of the same if/else (rather than, say, always playing sfx_wave and separately checking whether to also play sfx_win) guarantees they can never both fire on the same event - a direct benefit of Lesson 11's branching structure.",
    ],
    talkingPoints: [
      "Ask students who already had sfx_shoot/sfx_hit/sfx_win from earlier optional lessons to identify, out loud, exactly which two sounds are new to them today - a good way to surface who's starting fresh versus who's just adding two lines.",
      "For students ready for a stretch: ask what would happen if sfx_hit were accidentally placed inside the if <enemies_left = 0> block instead of right after the decrement - it would only play on the enemy that empties the wave, not every hit. A good test of whether they understand where each sound's trigger actually lives.",
    ],
    mistakes: [
      "Placing sfx_wave or sfx_win before the if/else split instead of inside their respective branches, causing both to play on every wave clear including the true win.",
      "Uploading a new sound file to the wrong sprite (Scratch sounds are per-sprite) and wondering why \"start sound\" doesn't find it.",
      "Forgetting that a student who skipped Lessons 2/3/4/6 entirely needs full instructions this week, not just \"add two sounds\" - this lesson's Handout is written to work for both cases.",
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    teaches: "Introducing one shared AudioStreamPlayer2D (RoundSound) that swaps its .stream property between three interchangeable clips (win, wave, lose) rather than adding three separate always-idle nodes - a genuinely new Godot technique, and a nice illustration of reuse applied to a resource instead of a function.",
    level1: [
      "RoundSound is new this week - a single node representing 'something happened that ends or advances the round,' regardless of which of the three outcomes it actually is.",
      "sfx_shoot and sfx_hit stay exactly where they were (Player's and Enemy's own AudioStreamPlayer2D nodes, from Lessons 3-4) - only the round-ending sounds get the new shared-node treatment, since shoot/hit can genuinely overlap each other while win/wave/lose are always mutually exclusive.",
      "Setting .stream immediately before every .play() call is the one detail that makes the shared-node approach work correctly - skipping it would replay whatever clip was assigned last.",
    ],
    why: [
      "A shared player for mutually-exclusive sounds (win/wave/lose) is a genuinely different pattern from Player's/Enemy's dedicated single-clip players (shoot/hit), and it's worth naming that difference explicitly: dedicated players suit sounds that can overlap themselves or each other; a shared swappable player suits outcomes where only one can ever be true.",
      "This is a natural, lightweight introduction to a 'sound manager' idea without building a whole new system - exactly the kind of small, justified new tool this course reaches for only when the old one stops fitting (the same discipline behind Lesson 9's real HUD and Lesson 10's reusable reset function).",
    ],
    talkingPoints: [
      "Ask students to explain, in their own words, why RoundSound doesn't need its own is_playing-style guard the way Lesson 2's beep-beep blip did - because each event that triggers it can only happen once per moment (a wave clear, a true win, a loss), not something checked continuously every frame.",
      "For students comparing engines: Unity's roundAudioSrc.PlayOneShot() with three clips is conceptually identical to Godot's RoundSound.stream-swap this week, just using a different underlying mechanism (PlayOneShot doesn't need an explicit .stream assignment at all).",
    ],
    mistakes: [
      "Forgetting to set .stream before calling .play(), so RoundSound plays whatever it played last instead of the intended clip.",
      "Placing the SND_WAVE assignment/play() call outside the else branch, causing it to also fire on the true win.",
      "Wiring the new sound calls onto the wrong AudioStreamPlayer2D (e.g. reusing Player's shoot player for a round-ending sound), which technically works but breaks the 'dedicated vs. shared' distinction this lesson is teaching.",
    ],
  },
  {
    engine: "Unity", file: "Unity",
    teaches: "Introducing a GameManager.instance singleton (mirroring Lesson 10's EnemySpawner.instance) specifically so static methods can reach an Inspector-assigned AudioSource, then using PlayOneShot() with three interchangeable clips for the round-ending sounds - genuinely new Unity technique, directly motivated by a real limitation (static methods can't see non-static Inspector fields without one).",
    level1: [
      "GameManager.instance is new this week - if it didn't already exist for some other reason, this is the first time GameManager itself needs the singleton pattern, for exactly the reason EnemySpawner needed it in Lesson 10.",
      "sfx_shoot and sfx_hit stay exactly where they were (Player's and Enemy prefab's own AudioSource components, from Lessons 3-4) - only the round-ending sounds get the new shared roundAudioSrc treatment.",
      "PlayOneShot() (mentioned as a concept back in Lesson 2's key notes but not used until now) is the right call here specifically because it doesn't require checking .isPlaying first and layers naturally - a good payoff for a concept introduced early but not needed until this lesson.",
    ],
    why: [
      "PlayOneShot() versus .Play(): .Play() (used for shoot/hit) replaces whatever the AudioSource was already playing, which is fine for sounds triggered from their own dedicated source; PlayOneShot() layers a clip on top without disturbing anything already playing, which is the right choice for a shared source that might be asked to play three different things.",
      "Routing win/wave/lose through one shared instance rather than three separate GameObjects keeps the audio setup proportional to the game's actual complexity - three mutually-exclusive outcomes don't need three permanently-idle AudioSources.",
    ],
    talkingPoints: [
      "Ask students to trace why instance.roundAudioSrc would throw a NullReferenceException if Awake() hadn't run yet - a good concrete illustration of initialization order mattering, reinforcing why EnemySpawner.instance needed the same care in Lesson 10.",
      "For students comparing engines: Godot's RoundSound.stream-swap and Unity's roundAudioSrc.PlayOneShot() solve the identical problem (one shared player, several possible outcomes) with genuinely different underlying tools - worth naming that cross-engine parallel explicitly.",
    ],
    mistakes: [
      "Missing or misplaced instance = this; in Awake(), causing a null reference the first time any RegisterKill()/RegisterEscape() call tries to reach roundAudioSrc.",
      "Placing the waveClip PlayOneShot() call outside the else branch, causing it to also fire on the true win.",
      "Forgetting to drag the AudioClip assets onto GameManager's Inspector fields after adding them in code - a common 'it compiles but plays nothing' mistake.",
    ],
  },
  {
    engine: "GameMaker", file: "GameMaker",
    teaches: "No new GameMaker mechanism this week - audio_play_sound() was already fully established by Lesson 3, so this lesson is purely about recognizing which existing branches (from Lessons 4, 5, 10, 11) are the correct trigger points for two brand-new sounds.",
    level1: [
      "sfx_hit moves to directly follow the score/enemies_remaining updates in obj_bullet's Collision event - matching the same restructured location every other engine's hit sound now lives in.",
      "sfx_wave and sfx_lose are the two genuinely new sounds, placed exactly where Lesson 11's nested if and Lesson 8's lives-check already live.",
      "audio_play_sound()'s three arguments (sound, priority, loop) stay identical to Lesson 3's original call for every one of the five sounds this week - no new parameter usage needed.",
    ],
    why: [
      "Because every trigger point (a collision, a wave transition, an escape) is already event-driven rather than polled, none of today's five sounds need any kind of replay-guard - they each fire from code that only runs once per actual occurrence.",
      "Keeping every sound call physically adjacent to the game-state change it represents (score/enemies_remaining update, then immediately the matching sound) keeps this lesson's additions easy to audit later - a direct continuation of the Structure principle from Lesson 7's original bugfix.",
    ],
    talkingPoints: [
      "Ask students to find all five audio_play_sound() calls across obj_bullet and obj_enemy and name, for each, exactly which player action or game event causes it to run.",
      "For advanced students: ask what would need to change to give each enemy type (if the project ever added more than one) its own distinct hit sound - not required this week, a natural preview of expanding the Sound resource set later in the year.",
    ],
    mistakes: [
      "Placing the sfx_wave call outside the else branch, causing it to also fire on the true win.",
      "Forgetting the obj_spawner. prefix somewhere in the revised blocks, since this code still runs from obj_bullet and obj_enemy.",
      "Creating a Sound resource but never actually assigning an imported .wav file to it, leading to a silent audio_play_sound() call with no error.",
    ],
  },
  {
    engine: "GDevelop", file: "GDevelop",
    teaches: "No new GDevelop mechanism this week - the \"Play a sound\" action was already fully established by Lesson 3, so this lesson is purely about adding it to the correct existing events (from Lessons 4, 6, 8, 10, 11).",
    level1: [
      "sfx_hit.wav's Play-a-sound action joins the existing bullet-enemy collision event, alongside the Score and EnemiesRemaining actions already there.",
      "sfx_wave.wav and sfx_lose.wav are the two genuinely new sounds, added to the exact two events Lesson 11 and Lesson 8/10 already built: the advance-wave event and the lose event.",
      "The full-win and advance-wave events' opposite Wave conditions (from Lesson 11) are exactly what keeps sfx_win.wav and sfx_wave.wav from ever firing together - no additional guard condition is needed this week.",
    ],
    why: [
      "GDevelop's action-per-event model means adding a sound is literally just one more Action row on an event that already exists - there's no separate 'wire this up to an event' step the way code-based engines need, which makes this an unusually low-friction lesson in this engine specifically.",
      "Keeping every sound action on the SAME event as the state change it corresponds to (rather than a separate audio-only event checking the same conditions again) avoids duplicating condition logic and keeps the project's event count from growing unnecessarily.",
    ],
    talkingPoints: [
      "Ask students to identify, for each of the five Play-a-sound actions, which single event it lives on and why that's the correct event (not a nearby one that looks similar).",
      "For students ready for a stretch: revisit the Lesson 11 optional Functions idea and discuss what a reusable PlaySound(file) function would look like here, even though GDevelop's built-in action already makes it nearly as simple as calling one.",
    ],
    mistakes: [
      "Adding a sound action to the wrong one of the two wave-related events (full-win vs. advance-wave), causing the wrong sound at the wrong moment.",
      "Forgetting that the full-win and advance-wave events need their opposite Wave conditions kept intact from Lesson 11 - editing one without checking the other can reopen the 'both fire at once' bug.",
      "Referencing a sound file name with a typo, which GDevelop silently accepts and simply plays nothing.",
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    teaches: "The plainest possible version of this week's idea: five sfx() calls in five existing branches, using only what Lesson 2's Sound Editor already established - and a chance to explicitly retire the hasWon-style replay-guard flag Lesson 6 needed, now that Lesson 10/11's restructuring made every trigger point naturally one-shot.",
    level1: [
      "sfx(2) (hit) moves to directly follow the score/enemies_remaining updates inside the bullets loop's hit branch - the same restructured location every other engine's hit sound now lives in.",
      "sfx(4) (wave) and sfx(5) (lose) are the two genuinely new sounds, placed exactly where Lesson 11's nested if and the enemy escape-check already live.",
      "sfx(3) (win), if a student built it back in Lesson 6, may originally have needed a hasWon guard flag to avoid replaying every frame - as of Lesson 10/11's restructuring into a one-shot hit branch, that guard is no longer necessary and can be removed if still present.",
    ],
    why: [
      "TIC-80's collision detection scans every frame, but a hit only ever causes enemies_remaining to cross to a new value once - so code placed directly inside that specific crossing (not a separately-polled check in TIC() itself) is naturally edge-triggered, with no manual guard needed.",
      "This is a good moment to point out that Lesson 6's original hasWon-flag workaround was a reasonable fix for the code structure at the time - and that Lesson 10's game_active refactor incidentally solved the same problem more thoroughly, a nice example of good structure paying dividends beyond its original purpose.",
    ],
    talkingPoints: [
      "Ask students who built Lesson 6's optional win sound to check whether they still have a hasWon guard flag anywhere, and if so, to trace through whether it's still doing anything useful post-Lesson-10 - a good structural-reasoning exercise, not just a code-reading one.",
      "For advanced students: ask them to name all five SFX Editor slots and their triggers from memory as a quick comprehension check before moving to lab time.",
    ],
    mistakes: [
      "Placing sfx(4) outside the else branch, causing it to also fire on the true win.",
      "Mixing up slot numbers between 3, 4, and 5 - easy to do since they're consecutive and were designed close together in the SFX Editor.",
      "Leaving a stale hasWon-guard flag in place from Lesson 6 that now silently prevents the win sound from ever playing a second time across multiple playthroughs without a full restart.",
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    teaches: "No new 001 Game Creator mechanism this week - the Play Sound action was already fully established by Lesson 3, so this lesson is purely about adding it to the correct existing Triggers and branches (from Lessons 4, 6, 7, 10, 11).",
    level1: [
      "Play Sound(sfx_hit.wav) joins the enemy Actor's \"Collided with actor\" Trigger, right after the score/enemies_remaining Variable updates that Lesson 7's bugfix already placed there.",
      "Play Sound(sfx_wave.wav) and Play Sound(sfx_lose.wav) are the two genuinely new sounds, placed exactly where Lesson 11's nested If and Lesson 10's gated lose branch already live.",
      "Every sound action this week sits inside the same If(game_active = 1) gate Lesson 10 introduced, so none of them can fire after the round has already ended.",
    ],
    why: [
      "Because Play Sound is a plain action with no special state, adding it anywhere in an existing Trigger's action list is safe and predictable - there's no separate registration step the way a code-based engine's AudioSource/AudioStreamPlayer needs.",
      "Keeping every sound action physically adjacent to the Variable change it represents continues the same Structure discipline this engine has followed since Lesson 7's bugfix - state changes (and now, sound) belong in the specific Trigger representing their event.",
    ],
    talkingPoints: [
      "Ask students to identify, for each of the five Play Sound actions, which single Trigger it lives on and why that's the correct one.",
      "This is a good lesson to revisit the If(game_active = 1) gate from Lesson 10 one more time: point out that it protects sound the same way it protects score/lives - once the round ends, nothing inside that gate (including sound) can fire again until a restart.",
    ],
    mistakes: [
      "Placing Play Sound(sfx_wave.wav) outside the Else branch of the nested If, causing it to also fire on the true win.",
      "Adding a sound action outside the If(game_active = 1) gate, letting it keep firing after the round has already ended.",
      "Referencing an imported sound asset before it finishes importing, leading to a silently-ignored Play Sound action.",
    ],
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    teaches: "No new Blueprint mechanism this week - Play Sound 2D was already fully established by Lesson 3, so this lesson is purely about wiring it into the correct existing Custom Events and Branches (from Lessons 4, 6, 10, 11).",
    level1: [
      "Play Sound 2D(sfx_hit.wav) joins HandleEnemyDestroyed right after the EnemiesRemaining/Score updates, before the nested Branch Lesson 11 built.",
      "Play Sound 2D(sfx_wave.wav) and Play Sound 2D(sfx_lose.wav) are the two genuinely new sounds, wired into the exact two Branch cases Lesson 11 and Lesson 10 already built.",
      "Every sound node this week is chained directly off the same Custom Event graph that already handles its corresponding state change - no separate audio Blueprint or Actor Component is needed for a game this size.",
    ],
    why: [
      "Chaining Play Sound 2D nodes directly into the existing HandleEnemyDestroyed/HandleEnemyEscaped graphs (rather than a separate audio-manager Blueprint) keeps this lesson's additions proportional to the game's actual complexity - the same restraint shown in Lessons 9 and 10 when building new UI/state systems only when genuinely needed.",
      "Placing sfx_win.wav and sfx_wave.wav in opposite cases of the same nested Branch (rather than an extra condition check) guarantees they can never both fire on the same event, for free, from Lesson 11's existing structure.",
    ],
    talkingPoints: [
      "Ask students to trace the full graph for a wave-3 clear: HandleEnemyDestroyed -> nested Branch True case -> GameActive = false, Play Sound 2D(sfx_win.wav), SetStatus - a good comprehension check tying this week's addition to last week's structure.",
      "For students on the C++ track: point out that UGameplayStatics::PlaySound2D() is the exact same function whether called from five different spots or wrapped in a small helper - a natural discussion of when extracting a helper function is worth it versus not, for a game this size.",
    ],
    mistakes: [
      "Wiring sfx_wave.wav's Play Sound 2D node outside the False case of the nested Branch, causing it to also fire on the true win.",
      "Forgetting to import a raw .wav as a proper Sound Wave asset before referencing it in a Play Sound 2D node.",
      "Placing a sound node in HandleEnemyEscaped before the Lives -= 1 update rather than after, which is harmless here but inconsistent with every other engine's ordering this week.",
    ],
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_12_Key_${d.file}.docx`]));

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
