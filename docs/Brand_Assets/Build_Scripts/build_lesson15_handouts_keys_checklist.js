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

function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function numbered(text, level = 0) { return new Paragraph({ numbering: { reference: "numbers", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function checkbox(text) { return new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 80 }, children: [new TextRun(text)] }); }

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
      headers: { default: brand.pageHeader("Lesson 15 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 15 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("This Is the Semester 1 Capstone Milestone"),
        p("Today is genuinely different from every prior lesson: there's no new mechanic to teach. Lesson 15 is a workshop-and-showcase day - students spend the first half polishing whatever's roughest in their game, and the second half presenting it to the class. This is the Level 3 Refined milestone from the Feature-to-Lesson Map, and it deserves the same celebratory weight as Lesson 6's Level 1 MVP and Lesson 11's Level 2 Complete."),
        checkbox("Read the Teacher's Script's full run-down before class - the agenda shape (polish, then present) is different from a typical lecture/lab split, and timing discipline matters more than usual since presentations have a hard end time."),

        h2("No New Assets or Mechanics This Week"),
        p("Every engine's Handout this week is a checklist and presentation guide, not new code to build. The 'Level 1/2/3' framework still appears, but it's being used as a self-assessment rubric against everything built since Lesson 2, not a new exercise."),

        h2("Continuity From Lesson 14"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 14, not create a new one - this IS the capstone; there is no separate submission."),
        checkbox("Confirm every student has SOME playable game by this point - if anyone is significantly behind, prioritize getting them to a presentable Level 1/2 state during polish time rather than chasing Level 3 today."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop can still open and run a finished example project in all 8 engines, in case a student's own project has a last-minute problem during their presentation slot."),
        checkbox("Skim each engine's Teacher's Key this week - it's reorganized as a 'known remaining quirks' reference rather than a worked solution, useful for quickly diagnosing a stuck student during polish time."),

        h2("Presentation Logistics"),
        checkbox("Decide and announce the presentation format ahead of time: whole-class demos at every student's seat/station, or a few volunteers presenting to the front of the room, depending on class size and time."),
        checkbox("Prepare a simple running order or sign-up sheet if presenting to the front of the room, so transitions between students are quick."),
        checkbox("Set a firm per-student time limit (suggested: 2 minutes each) and communicate it clearly before presentations begin."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) for quick troubleshooting during polish time."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies, and something to track presentation order/timing."),

        h2("Known Gotchas This Week"),
        bullet("Polish time can easily be swallowed by chasing new features instead of fixing existing rough edges - actively redirect students toward bug-fixing and tuning over adding anything new."),
        bullet("Presentations run long almost every time - hold the time limit firmly and warn the class at the halfway point."),
        bullet("Not every student will have finished every Level 3 item, including their Lesson 14 stretch feature - that's completely fine. A confident presentation of a Level 1/2 game is a full success today."),
        bullet("This is a genuine milestone - celebrate it. Consider a specific closing statement acknowledging that every student built and finished an actual game over 15 weeks, regardless of which engine or how far up the Level ladder they got."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_15_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS - Level 3 Refined Checklist + Presentation Guide
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 15: Ready to Show", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("Today has no new code to build. It's a workshop day: polish what's roughest, then present your finished capstone game to the class. This is the Level 3 Refined milestone - the end of Semester 1's build.", { bold: true }),
    p("Open your own [YourName]Game project from Lesson 14 - this IS your capstone. There's nothing separate to submit.", { bold: true }),

    h2("Your Full Capstone Checklist"),
    p("Work through this checklist first. Anything unchecked is your polish-time priority - fix these before adding anything new.", { italics: true, color: "555555" }),
    h3("Core Mechanics (Lessons 2-11)"),
    checkbox("Move and shoot respond correctly to input, every time."),
    checkbox("Enemies are destroyed on hit, and score increases correctly."),
    checkbox("An enemy escaping unshot costs a life, and lives are tracked correctly."),
    checkbox("Score, lives, and wave are all shown clearly on screen at all times."),
    checkbox("Winning and losing both work, and both show a clear message with a restart hint."),
    checkbox("Pressing R restarts cleanly from any state - mid-round, after a win, or after a loss - back to wave 1."),
    checkbox("Multiple waves work correctly: clearing a non-final wave spawns a fresh one; only the final wave triggers a real win."),
    h3("Sound (Lesson 12)"),
    checkbox("A shot makes a sound."),
    checkbox("A hit/explosion makes a sound."),
    checkbox("A wave clear makes a sound, distinct from the win sound."),
    checkbox("A game over makes a sound."),
    h3("Animation (Lesson 13)"),
    checkbox("A hit briefly flashes the enemy before it's removed."),
    checkbox("An independent explosion effect appears and cleans itself up on every hit."),
    h3("Difficulty (Lesson 14)"),
    checkbox("Enemy count increases with each wave."),
    checkbox("Enemy speed increases with each wave."),
    checkbox("Your chosen Level 3 stretch feature is present in at least a rough, good-faith form."),

    h2("Polish Pass"),
    p("With whatever time remains after the checklist above, work top to bottom through these polish priorities - in order:", { italics: true, color: "555555" }),
    numbered("Fix anything broken or unchecked above first - a complete, working Level 1/2 game beats a flashy but broken one."),
    numbered("Tune anything that feels unfair or confusing - difficulty too easy/hard, unclear on-screen text, a sound that's too loud or jarring."),
    numbered("Only after 1 and 2: continue building out your Lesson 14 stretch feature, or add a small extra touch if you have time and energy left."),
    ...data.engineNotes.map(s => bullet(s)),

    h2("Preparing Your Presentation"),
    p("Presentations are short - about 2 minutes each. You don't need slides or notes; just be ready to talk while you play. A simple structure that works well:", { italics: true, color: "555555" }),
    numbered("Name your game and describe its core mechanic in one sentence."),
    numbered("Play a short round - enough to show a hit, a wave clear, and either a win or a loss."),
    numbered("Point out your chosen Level 3 stretch feature specifically, even if it's not fully finished."),
    numbered("If something breaks live, that's genuinely fine - briefly say what it's supposed to do and move on. A confident presentation matters more than a perfect one."),

    h2("Want More?"),
    p(data.resource),
  ];

  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 15 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    engineNotes: [
      "Common Scratch rough edges to check: variables not set to \"for all sprites\" when they should be, a stray \"forever\" loop left in the wrong script, or a costume left in the wrong state after cloning.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    engineNotes: [
      "Common Godot rough edges to check: a signal connected twice (causing double-counted hits), a stray print() left in from debugging, or a node reference ($NodeName) that breaks after a scene tree reorganization.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    engineNotes: [
      "Common Unity rough edges to check: a prefab that wasn't updated after editing the version in the scene, an Inspector field left unassigned (null reference errors), or leftover Debug.Log() calls.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
  },
  {
    engine: "GameMaker", file: "GameMaker",
    engineNotes: [
      "Common GameMaker rough edges to check: a missing obj_spawner. prefix causing a variable-scope error, an Alarm set but never matched with an Alarm event, or a Sound resource created but never assigned an actual file.",
    ],
    resource: "Student Resource List > GameMaker section, or manual.gamemaker.io",
  },
  {
    engine: "GDevelop", file: "GDevelop",
    engineNotes: [
      "Common GDevelop rough edges to check: two events with conditions that should be opposites (like Wave >= TotalWaves and Wave < TotalWaves) that have drifted out of sync, or an object variable confused with a scene variable.",
    ],
    resource: "Student Resource List > GDevelop section, or wiki.gdevelop.io",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    engineNotes: [
      "Common TIC-80 rough edges to check: a table.remove() happening while iterating forward through a table (should count backward, i=#t,1,-1), or a leftover debug print() cluttering the screen.",
    ],
    resource: "Student Resource List > TIC-80 section, or github.com/nesbox/TIC-80/wiki",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    engineNotes: [
      "Common 001 Game Creator rough edges to check: an Actor name typo breaking a Custom Event reference, a Trigger missing the If(game_active = 1) gate, or a pre-placed Actor left in the wrong starting position after testing.",
    ],
    resource: "001gamecreator.com/wiki/, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    engineNotes: [
      "Common Unreal rough edges to check: a Custom Event dispatcher bound twice (causing double-counted hits), an unset Blueprint variable defaulting to zero/null, or a Delay node left from testing that makes the game feel sluggish.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_15_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS - Known Remaining Quirks Reference (not worked solutions)
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 15 Teacher's Key: Ready to Show", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine + " - Instructor Reference Only", size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    h2("How This Key Is Different Today"),
    p("There's no new Level 1 solution to give this week - Lesson 15 is a polish-and-present day, not a new-mechanic day. This Key instead consolidates the most common last-mile issues seen across this engine's Lessons 2-14, to help you diagnose a stuck student quickly during polish time."),
    h2("Known Recurring Quirks in This Engine"),
    ...data.quirks.map(s => bullet(s)),
    h2("Fast Diagnostic Questions"),
    ...data.diagnostics.map(s => bullet(s)),
    h2("If a Student Is Significantly Behind"),
    p(data.behindAdvice),
    h2("Presentation Day Notes"),
    ...data.presentationNotes.map(s => bullet(s)),
  ];
  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 15 Key - " + data.engine + " (Instructor Only)") },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    quirks: [
      "Variables scoped to one sprite instead of \"for all sprites,\" so a value looks right on one sprite's script but reads as 0/default elsewhere.",
      "A broadcast script (\"when I receive [x]\") accidentally left off the sprite it needs to be on, especially after copy-pasting scripts between sprites.",
      "A costume/graphic effect (from Lesson 13's flash) left set to a nonzero value after a clone is deleted, which can bleed into a freshly-cloned enemy if not reset in the spawn script.",
    ],
    diagnostics: [
      "\"Is this variable 'for all sprites' or 'for this sprite only'?\" - the single most common root cause of confusing behavior in Scratch by this point in the course.",
      "\"Which script actually broadcasts this, and which sprite is listening for it?\" - broadcast bugs are almost always a mismatch here.",
    ],
    behindAdvice: "Prioritize getting Lessons 6-11's Level 1 (win/lose, restart, multiple waves) solidly working before anything from Lessons 12-14 - a complete, simple game presents better than a broken ambitious one.",
    presentationNotes: [
      "Scratch's stage view makes for an easy, quick demo - no special setup needed beyond having the project open and the green flag ready.",
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    quirks: [
      "A signal connected more than once (often from re-running a scene multiple times during testing without restarting the editor), causing double-fired hit/escape logic.",
      "A $NodeName reference that silently breaks after a node was renamed or moved in the scene tree, producing a null/error instead of a clear symptom.",
      "await-based code (from Lesson 13) that got copy-pasted without its enclosing function's async context, causing a parse error.",
    ],
    diagnostics: [
      "\"Run the scene and check the Output/Debugger panel for errors\" - Godot surfaces most of these issues directly if a student hasn't been checking.",
      "\"Did you rename or move any nodes recently?\" - a fast way to find a broken $NodeName reference.",
    ],
    behindAdvice: "Prioritize getting Lessons 6-11's Level 1 (win/lose, restart, multiple waves) solidly working before anything from Lessons 12-14 - a complete, simple game presents better than a broken ambitious one.",
    presentationNotes: [
      "Have students run their main scene directly (not a test scene) before presenting, to avoid an awkward mid-demo scene switch.",
    ],
  },
  {
    engine: "Unity", file: "Unity",
    quirks: [
      "An Inspector field left unassigned after adding a new public variable (common this semester with explosionPrefab, roundAudioSrc, etc.), producing a NullReferenceException at the worst moment.",
      "A prefab edited in the Hierarchy but never applied back to the actual Prefab asset, so changes don't persist or show up for other instances.",
      "GameManager.instance not yet assigned if Awake() ordering is inconsistent across multiple GameManager-like objects accidentally left in the scene.",
    ],
    diagnostics: [
      "\"Check the Console for a red error\" - Unity's NullReferenceException messages usually name the exact missing field.",
      "\"Is there more than one GameManager in the scene?\" - a surprisingly common source of confusing singleton behavior by this point.",
    ],
    behindAdvice: "Prioritize getting Lessons 6-11's Level 1 (win/lose, restart, multiple waves) solidly working before anything from Lessons 12-14 - a complete, simple game presents better than a broken ambitious one.",
    presentationNotes: [
      "Have students press Play from a fresh Editor state (not mid-Play-mode-edit) before presenting, since Unity discards Play-mode-only changes on stop.",
    ],
  },
  {
    engine: "GameMaker", file: "GameMaker",
    quirks: [
      "A missing obj_spawner. prefix somewhere in code that runs from obj_bullet or obj_enemy, causing a variable-scope error or a silently-wrong instance being modified.",
      "An Alarm set (alarm[0] = N) with no matching Alarm event added, so nothing ever happens when it expires.",
      "A Sound resource created but never assigned an actual imported file, so audio_play_sound() runs with no audible result and no error.",
    ],
    diagnostics: [
      "\"Run the game (F5) and check for a red error dialog\" - GameMaker surfaces scope/reference errors directly.",
      "\"Does this object actually have the Alarm/Step/Collision event you think it does?\" - easy to add code to the wrong event by habit.",
    ],
    behindAdvice: "Prioritize getting Lessons 6-11's Level 1 (win/lose, restart, multiple waves) solidly working before anything from Lessons 12-14 - a complete, simple game presents better than a broken ambitious one.",
    presentationNotes: [
      "Have students run from the correct starting Room before presenting, in case testing left a different Room set as default.",
    ],
  },
  {
    engine: "GDevelop", file: "GDevelop",
    quirks: [
      "Two events meant to be mutually exclusive (like the full-win and advance-wave events from Lesson 11) that have drifted out of sync after later edits, causing both or neither to fire.",
      "A Speed or similar value added as a scene variable instead of an object variable (or vice versa), giving every instance the same value instead of individual ones.",
      "An extension (like Lesson 13's Flash) that got uninstalled or never installed on a specific project, silently hiding its actions from the list.",
    ],
    diagnostics: [
      "\"Click Preview and watch the Debugger tab for variable values in real time\" - GDevelop's live debugger is the fastest way to see what's actually happening.",
      "\"Are this event's conditions still exact opposites of its paired event's conditions?\" - relevant anywhere two mutually-exclusive events were built (Lesson 11's win/wave split, for example).",
    ],
    behindAdvice: "Prioritize getting Lessons 6-11's Level 1 (win/lose, restart, multiple waves) solidly working before anything from Lessons 12-14 - a complete, simple game presents better than a broken ambitious one.",
    presentationNotes: [
      "GDevelop's Preview window is quick to relaunch if something goes wrong mid-demo - encourage students not to panic over a needed restart.",
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    quirks: [
      "Iterating a table forward while removing entries from it (should count backward: for i=#t,1,-1), which skips entries or errors depending on what's being removed.",
      "A leftover debug print() statement cluttering the screen, easy to miss since TIC-80 has no separate console/output panel.",
      "SFX Editor slot numbers that drifted out of sync with what's referenced in code, especially after adding Lesson 12/14's new sounds.",
    ],
    diagnostics: [
      "\"Are you counting UP or DOWN through this table, and are you removing entries from it?\" - the single most common structural bug by this point in TIC-80.",
      "\"What's actually printed on screen right now that shouldn't be?\" - a fast way to spot leftover debug output.",
    ],
    behindAdvice: "Prioritize getting Lessons 6-11's Level 1 (win/lose, restart, multiple waves) solidly working before anything from Lessons 12-14 - a complete, simple game presents better than a broken ambitious one.",
    presentationNotes: [
      "Have students type run and confirm a clean start before presenting, since TIC-80 keeps whatever state was last active in memory.",
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    quirks: [
      "An Actor name typo (enemy_6 vs Enemy_6, for example) silently breaking a Custom Event's reference to it, with no obvious error shown.",
      "A Trigger missing the If(game_active = 1) gate from Lesson 10, allowing score/lives/wave to keep changing after the round has already ended.",
      "A pre-placed Actor left in the wrong starting position after manual testing/dragging in the editor, so it teleports incorrectly on the next SpawnWave.",
    ],
    diagnostics: [
      "\"Does this Actor's name match EXACTLY, including capitalization, everywhere it's referenced?\" - the most common root cause of a silently-broken Custom Event.",
      "\"Is this Trigger's logic wrapped in If(game_active = 1)?\" - relevant anywhere score/lives/wave could still be changing after a win or loss.",
    ],
    behindAdvice: "Prioritize getting Lessons 6-11's Level 1 (win/lose, restart, multiple waves) solidly working before anything from Lessons 12-14 - a complete, simple game presents better than a broken ambitious one.",
    presentationNotes: [
      "Have students Save and reopen the project before presenting if anything felt unstable during testing, since this engine benefits from a clean reload more than most.",
    ],
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    quirks: [
      "A Custom Event dispatcher (like OnEnemyDestroyed) bound more than once across multiple Play sessions without a full editor restart, causing double-fired logic.",
      "An unset Blueprint variable quietly defaulting to 0/null instead of erroring, especially for anything added recently (EnemySpeed, explosion references, Dynamic Material Instances).",
      "A Delay node left over from testing (or an unusually long one) making the game feel sluggish or unresponsive during a live demo.",
    ],
    diagnostics: [
      "\"Check the Output Log for a warning or error\" - Unreal surfaces most null-reference and binding issues here.",
      "\"Has this dispatcher/event been bound more than once?\" - relevant anywhere double-fired logic is suspected.",
    ],
    behindAdvice: "Prioritize getting Lessons 6-11's Level 1 (win/lose, restart, multiple waves) solidly working before anything from Lessons 12-14 - a complete, simple game presents better than a broken ambitious one. Given this track's optional/advanced status, it's also fine for a struggling student to present a strong Level 1/2 build without every Level 3 feature.",
    presentationNotes: [
      "Have students Play in a fresh Editor session (not a long-running one with many prior Play/Stop cycles) before presenting, since accumulated state occasionally causes odd one-off behavior.",
    ],
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_15_Key_${d.file}.docx`]));

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
