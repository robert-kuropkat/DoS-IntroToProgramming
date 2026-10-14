const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign,
  ImageRun
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
function pRuns(runs, opts = {}) { return new Paragraph({ spacing: { after: 160 }, ...opts, children: runs }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function numbered(text, level = 0) { return new Paragraph({ numbering: { reference: "numbers", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function checkbox(text) { return new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 80 }, children: [new TextRun(text)] }); }
function labelBullet(label, text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)] });
}
function code(lines) {
  return new Paragraph({
    spacing: { after: 160, before: 80 }, indent: { left: 360 },
    shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
    children: [new TextRun({ text: lines.join("\n"), font: "Consolas", size: 20 })]
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

const jobs = [];

// =========================================================================
// 1. ONE-PAGE LESSON PLAN
// =========================================================================
{
  const agendaRows = [
    ["0:00-0:05", "Welcome back, Lesson 10 recap (game state, the Stalemate)"],
    ["0:05-0:20", "One More Question: Is This the Last Wave?"],
    ["0:20-0:40", "The Same Spawn Code, Called a Third Way (live demo in Unity)"],
    ["0:40-0:50", "Milestone Moment: What \"Level 2 Complete\" Actually Means"],
    ["0:50-0:55", "Introduce lab exercise, hand out engine-specific handouts"],
    ["0:55-1:50", "LAB: Build multiple waves, reaching the Level 2 Complete milestone"],
    ["1:50-2:00", "Exit check / quick share-out - milestone celebration"],
  ];
  const agendaWidths = [1600, 7760];
  const agendaTable = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: agendaWidths,
    rows: agendaRows.map((row, ri) => new TableRow({
      children: [bodyCell(row[0], agendaWidths[0], { fill: ri % 2 ? "F3F6FA" : "FFFFFF" }), bodyCell(row[1], agendaWidths[1], { fill: ri % 2 ? "F3F6FA" : "FFFFFF" })]
    }))
  });

  const doc = new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 900, right: 1080, bottom: 900, left: 1080 } } },
      children: [
        new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 14, bold: true, color: brand.GOLD_TEXT, characterSpacing: 16 })] }),
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 11 - Core Mechanics V: Wave After Wave", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 11 - Level 2 Complete Milestone", italics: true, size: 18, color: brand.GOLD_TEXT, bold: true })] }),

        labelBullet("Objectives", "Nest a second decision inside an existing conditional (is this the last wave?); reuse the spawn_enemies routine from a third caller, the game itself, not just launch and player input; keep new shared state (wave) on the GameManager and reset it on restart; explain what completing the Level 2 Complete milestone means for the capstone."),
        labelBullet("IGDA Tie-In", "Programming - state machines, code reuse/refactoring, game balance and pacing (intermediate)"),
        labelBullet("Featured Engine (lecture demo)", "Unity (C#)"),
        labelBullet("Materials", "Demo laptop with the four active engines verified (see Teacher's Checklist); Student Handouts for Scratch, Godot, GameMaker (script and Visual) and Unity; students' own Lesson 10 projects (already saved) - no new assets needed this week"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): clearing a wave spawns a fresh one, up to 3 total; clearing wave 3 triggers the real win; wave number shown on screen."),
        bullet("Level 2 (Complete): wave transitions are seamless and correct every time - no double-spawns, no skipped waves, restart always returns to wave 1."),
        bullet("Level 3 (Refined): fix the Stalemate (left in on purpose from Lesson 10) so it works with waves; add a brief wave-transition announcement; and/or tighten the reusable spawn structure."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 11 Teacher's Script, Lesson 11 Prep Checklist, and the per-engine Student Handouts. This is a milestone lesson - plan extra time for the Milestone Moment and end-of-lesson celebration, same as Lesson 6.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_11_Plan.docx"]);
}

// =========================================================================
// 2. TEACHER'S SCRIPT  (data-driven: the same text is written to lesson11_script_notes.json
//    so the Slide deck's speaker notes are copied verbatim - Standards Rule 8)
// =========================================================================
const SCRIPT = [
  { heading: "[0:00-0:05] Welcome Back", slide: "welcome", paras: [
    "\"Welcome back! Last week your game became genuinely replayable - a real freeze, a real restart. Quick show of hands - who tested a full win-then-restart AND a full lose-then-restart?\" [Pause, call on 1-2 volunteers]",
    "\"Two things from last week I want you to carry into today. First, game state: when you restart, EVERYTHING that changed has to be put back, including anything extra you added. Second, the bug. Some of you found a round that simply never ends - you destroyed some enemies, let the others escape, and nothing happened: no win, no loss, and R did nothing. Let's give it a name: the Stalemate. Hold that thought. It's still in your game, on purpose, and today it's about to get more interesting.\" [Do NOT explain the fix. If a student asks, say: 'What has to be true for a round to be over?']",
    "\"Today's a big one - by the end of class, your game reaches its second major milestone of the year. Clearing one wave of enemies is about to become clearing several, in a row, before you actually win.\"",
  ]},
  { heading: "[0:05-0:20] One More Question: Is This the Last Wave?", slide: "question", paras: [
    "\"Since Lesson 6, your win-check has asked exactly one question: are all the enemies gone? Today, that question grows a second part: gone, AND is this the last wave? If both are true, you really win. If enemies are gone but more waves remain, the round just... continues.\"",
    "\"This is the first time this course asks you to nest one decision inside another. It's a small change in code, but it's a genuinely new shape of thinking - a decision whose answer depends on a second, smaller decision.\"",
    "[Discussion question] \"Can anyone think of a decision in real life that works this way - where the answer to one question depends on a second question nested inside it?\" (Examples to draw out if needed: \"can I go to the party\" depends on \"is my homework done,\" which itself might depend on \"do I understand the assignment.\")",
    "[Timing checkpoint: wrap by 0:15]",
  ]},
  { heading: "[0:15-0:20] The Whole Decision, One Picture", slide: "activity", paras: [
    "[Show the next slide: the Activity Diagram] \"Here's the whole decision in one picture. An enemy is destroyed: enemies_remaining goes down by one. Is it zero? If not, the wave simply continues. If it is, the wave is cleared - and now the second question: is this the last wave? Yes: the real win, and R restarts as before. No: the next wave. wave goes up by one, enemies_remaining goes back to five, and we call spawn_enemies again. Notice score, lives and game_active are untouched. This is the same picture in every engine, whatever the code looks like.\"",
    "[Timing checkpoint: wrap by 0:20]",
  ]},
  { heading: "[0:20-0:40] The Same Spawn Code, Called a Third Way", slide: "demo", paras: [
    "\"Here's the good news: you're not writing new enemy-spawning code today. You've had it since Lesson 5, and Lesson 10 already made it callable a second way, for restarts. Today it gets a THIRD way to run: every time a wave is cleared, if more waves remain.\"",
    "[Demo cue: switch to Unity, open last week's project and GameManager.cs]",
    "\"Watch. SpawnEnemies() on EnemySpawner already has two callers: Start() and ResetGame(). Today I add three static fields to GameManager: wave, totalWaves and nextWaveRequested. Then I turn RegisterKill()'s win check into an if/else: if wave is at least totalWaves, hasWon and gameActive, exactly as before. Otherwise wave++, enemiesRemaining back to 5, and nextWaveRequested = true.\" [Build the if/else live]",
    "\"Why a flag, and not just spawner.SpawnEnemies() right there? Try it.\" [Type spawner.SpawnEnemies() inside RegisterKill() and let the compiler complain] \"RegisterKill() is static - it belongs to the GameManager class, not to the one GameManager object in the scene, and the Spawner slot we filled in last week belongs to that object. A static method can't see it. So RegisterKill() leaves a note, and GameManager's own Update(), which does belong to the object, checks the note, clears it and calls spawner.SpawnEnemies().\" [Add the nextWaveRequested check to Update()] \"That's the real cost of static members. We'll come back to it in Semester 2.\"",
    "[Demo cue: add public Text waveText to HUD.cs, one line in Update(), drag WaveText onto the slot in the Inspector, press Play and clear wave 1 into wave 2]",
    "\"Same enemy-creation method, zero changes to it. Three callers in: EnemySpawner's Start(), every restart through ResetGame(), and now every wave clear through GameManager's Update(). That's the whole idea.\"",
    "[Timing checkpoint: wrap by 0:35]",
  ]},
  { heading: "[0:35-0:40] New State, Same Habit", slide: "state", paras: [
    "[Show the next slide: New State, Same Habit] \"Last week's habit applies the moment you add anything new. wave and totalWaves are new game state, and they live on the GameManager, just like score and lives, because more than one object cares about them. So: does the restart put wave back to 1? I'll check it right now.\" [Demo cue: in the Unity demo, reach wave 3, lose, press R, and point at the HUD: it must say Wave: 1. If you forgot ResetGame()'s wave = 1, you will see Wave: 3 - leave that on screen for a moment before fixing it.]",
    "\"And if you add a wave message today, an effect, a sound, anything: ask the same two questions. What does it change when the round ends or the wave changes? And what puts it back?\"",
    "[Timing checkpoint: wrap by 0:40]",
  ]},
  { heading: "[0:40-0:50] Milestone Moment: What \"Level 2 Complete\" Actually Means", slide: "milestone", paras: [
    "\"Take a second to think about everything your game can do now: move, shoot, face a wave of enemies, track a score, survive with limited lives, restart cleanly - and now, face MULTIPLE waves before really winning. That's not a small list. That's the Level 2 Complete milestone from your capstone's design document.\"",
    "\"Back in Lesson 6, Level 1 MVP meant 'a complete, playable loop, start to finish.' Level 2 Complete means something more: a game with actual depth - failure that matters, progress that builds, and a real sense of escalation as more waves arrive. This is the kind of thing that starts to feel like an actual arcade game, not just a working prototype.\"",
    "\"From here through Lesson 15, everything is about making this same loop feel good - sound, animation, difficulty, and polish. The mechanics are done. What's left is making them sing.\"",
    "[Timing checkpoint: wrap by 0:50]",
  ]},
  { heading: "[0:50-0:55] Today's Exercise & Handouts", slide: "exercise", paras: [
    "\"For lab today, open your own project from last week - not a new one, and no new assets needed. Today's goal: clearing a wave spawns a new one, up to 3, and only clearing wave 3 is a real win.\"",
    "[Hand out engine-specific handouts]",
    "\"Test a FULL 3-wave playthrough today, not just wave 1 - and separately, make sure losing your last life mid-wave still works correctly, whichever wave you're on.\"",
    "\"One more thing. The Stalemate is still in there, and now it can freeze the game at any of the three waves. If it starts to annoy you, that's your Level 3. You already know how to find it.\"",
    "\"Questions before we start lab? ... Let's build your second milestone.\"",
  ]},
  { heading: "Lab Time Notes", slide: "lab", paras: [
    "- The single most common bug this week: forgetting the wave-count check entirely, so clearing wave 1 immediately shows the win message. Test this explicitly - the win should NOT appear until the final wave.",
    "- Second most common: the wave-advance branch accidentally touching score or lives - only enemies_remaining and the wave counter should change between waves.",
    "- Game-state drift: wave not reset on restart. Have students test restarting from mid-wave-2 or mid-wave-3, not just from a fresh game - restart should always return to wave 1. Point them back to the 'Check Your Game State' list in the Handout.",
    "- The Stalemate: students will start reporting 'my wave never ends.' Confirm it is real, name it, do not hand over the fix, and send them back to their Lesson 10 Level 3 notes. Students who already fixed it should check that the fix works at every wave. Each Key's Level 3 section has the worked fix.",
    "- Engine notes: Godot spawns the next wave with call_deferred (a wave clears mid-collision); Unity requests the next wave with a flag that GameManager's Update() acts on; Scratch broadcasts [spawn enemies]; GameMaker calls spawn_enemies(), with a Function Call block in Visual. A doubled wave in Scratch usually means a spawn script is still on [start game].",
    "- Circulate by engine cluster as usual. At 1:50, call time and do a real celebration share-out - this is the Level 2 Complete milestone. Consider having 2-3 students demo a full 3-wave playthrough from start to finish, the same energy as Lesson 6's milestone share-out.",
  ]},
];
// Notes for the slide deck: verbatim script text, keyed by slide id
const NOTES_OUT = {};
SCRIPT.forEach(sec => { NOTES_OUT[sec.slide] = sec.paras.join("\n\n"); });
fs.writeFileSync("lesson11_script_notes.json", JSON.stringify(NOTES_OUT, null, 1));
{
  const bodyKids = [];
  SCRIPT.forEach(sec => {
    bodyKids.push(h2(sec.heading));
    sec.paras.forEach(para => {
      if (para.startsWith("[Timing checkpoint")) bodyKids.push(p(para, { italics: true, color: "888888" }));
      else if (para.startsWith("- ")) bodyKids.push(bullet(para.slice(2)));
      else bodyKids.push(p(para));
    });
  });
  const doc = new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 11 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 11 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Core Mechanics V: Wave After Wave", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
        ...bodyKids,
      ]
    }]
  });
  jobs.push([doc, "Lesson_11_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
