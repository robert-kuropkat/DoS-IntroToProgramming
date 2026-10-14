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
    ["0:00-0:05", "Welcome back, Lesson 7 recap"],
    ["0:05-0:15", "Reversing the Question: What Happens When THEY Reach YOU?"],
    ["0:15-0:35", "Reusing What You Built: The Escape Path Gets Stakes (live demo in Unity)"],
    ["0:35-0:50", "Two Counters, Never Crossed: Keeping Lives and Kills Independent"],
    ["0:50-0:55", "Why We're Not Building Game Over (Yet)"],
    ["0:55-1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["1:00-1:50", "LAB: Add a lives system and lose condition to your project"],
    ["1:50-2:00", "Exit check / quick share-out"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 8 - Core Mechanics II: Under Pressure", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 8 - Lose Condition + Lives System", italics: true, size: 18, color: brand.GOLD_TEXT, bold: true })] }),

        labelBullet("Objectives", "Reuse an existing code path (Lesson 7's escape cleanup) to drive a new, independent piece of game state; implement a lives counter and a lose condition; explain why the lose message doesn't yet stop or reset the game, and why that's a deliberate, temporary choice."),
        labelBullet("IGDA Tie-In", "Programming - state tracking, win/lose conditions, event-driven programming (intermediate); Production - deliberate feature sequencing / incremental scope"),
        labelBullet("Featured Engine (lecture demo)", "Unity (C#)"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified (see Teacher's Checklist); Student Handouts, one per engine; students' own Lesson 7 projects (already saved); three placeholder sound effects in this lesson's Assets/ folder (sfx_lifelost.wav, sfx_gameover.wav, sfx_start.wav) for the Level 3 sound idea - the core exercise itself needs no new assets"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): every enemy that reaches the bottom unshot costs a life, tracked starting at 3; \"Game Over\" appears the instant lives reach zero."),
        bullet("Level 2 (Complete): shooting an enemy never costs a life, and letting one escape never counts as a kill - the two counters stay fully independent."),
        bullet("Level 3 (Refined): add a live lives readout and/or a lose sound, and tighten the separation between escape/lives code and hit-detection/win code."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 8 Teacher's Script, Lesson 8 Prep Checklist, and the per-engine Student Handouts. No engine-specific bugfix this week - every engine's Lesson 7 structure already supports this cleanly.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_08_Plan.docx"]);
}

// =========================================================================
// 2. TEACHER'S SCRIPT
// =========================================================================
{
  const doc = new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 8 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 8 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Core Mechanics II: Under Pressure", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Last week your enemies started moving toward you. Quick show of hands - who tuned their enemy speeds, or gave them varied movement?\" [Pause, call on 1-2 volunteers]"),
        p("\"So far, an enemy reaching the bottom of the screen has been a non-event - it just vanishes, no consequence. Today, that changes.\""),

        h2("[0:05-0:15] Reversing the Question: What Happens When THEY Reach YOU?"),
        p("\"Every lesson so far has asked one kind of question: what happens when YOUR bullet reaches an enemy? Today we ask the opposite question: what happens when an ENEMY reaches you?\""),
        p("\"Think back to your capstone's design document - the lose condition has always been there on paper: an enemy reaches the player's row, or the player runs out of lives. Today we actually build it.\""),
        p("[Discussion question] \"Why do you think a game needs both a win condition AND a lose condition, instead of just one or the other?\" (Draw out: stakes, tension, replayability - a game you can't lose isn't very interesting to keep playing.)"),
        p("[Timing checkpoint: wrap by 0:15]", { italics: true, color: "888888" }),

        h2("[0:15-0:35] Reusing What You Built: The Escape Path Gets Stakes"),
        p("\"Here's the fun part: you already wrote the code that detects an enemy reaching the bottom - you built it last week for cleanup. Today, we don't write new detection code. We just give that existing detection a real consequence.\""),
        p("[Demo cue: switch to Unity, open a project with working movement/escape from Lesson 7]"),
        p("\"Watch: here's Enemy.cs's escape branch from last week, in its Update() method - reaches the bottom, calls Destroy(gameObject), and until today, that was it. Today I add one line right before that Destroy() call: a call out to a new GameManager.RegisterEscape() method.\" [Add the RegisterEscape() call, show 'Game Over' appear after 3 escapes]"),
        p("\"Notice how small that change was. That's not an accident - it's the payoff of building things in clean, separate pieces. Last week's escape-cleanup code didn't know or care about win-counting, so today it was easy to give it a second job that also has nothing to do with win-counting.\""),
        p("[Discussion question] \"Can you think of a time in your own project, or in something else you've built, where old work made new work easier because it was organized well?\" (Open-ended; validate any real example, even a small one.)"),
        p("[Timing checkpoint: wrap by 0:35]", { italics: true, color: "888888" }),

        h2("[0:35-0:50] Two Counters, Never Crossed: Keeping Lives and Kills Independent"),
        p("\"Your game now tracks two completely separate numbers: enemies_remaining (or your engine's equivalent), which only changes when you shoot something, and lives, which only changes when something gets past you. These numbers should NEVER affect each other.\""),
        p("[Demo cue: same Unity project] \"Watch what happens if I shoot an enemy...\" [shoot one] \"...lives doesn't move. And if I let one escape...\" [let one pass] \"...enemiesRemaining doesn't move either. Two separate static methods on GameManager - RegisterKill() and RegisterEscape() - each touching only its own field, and neither one knows the other exists.\""),
        p("\"This is worth testing explicitly in your own project today, in both directions - shoot one and confirm lives stays put, then let one escape and confirm enemies_remaining stays put. If either one moves when it shouldn't, that's this week's most common bug, and it's almost always the new code added to the wrong branch or event.\""),
        p("[Timing checkpoint: wrap by 0:50]", { italics: true, color: "888888" }),

        h2("[0:50-0:55] Why We're Not Building Game Over (Yet)"),
        p("\"One more thing before lab: when lives hits zero today, 'Game Over' appears on screen - and then the game just... keeps running. Enemies keep spawning, you can still move and shoot. That's on purpose.\""),
        p("\"Actually stopping the game, freezing everything, and offering a restart button is real work - enough work that it gets its own lesson, Lesson 10. For today, just like Lesson 6's win message was 'just a message' before anything fancier got added, today's lose message is the same: get the detection and the message right first, build the proper response later.\""),
        p("[Timing checkpoint: wrap by 0:55]", { italics: true, color: "888888" }),

        h2("[0:55-1:00] Today's Exercise & Handouts"),
        p("\"For lab today, open your own project from last week - not a new one, and no new assets needed. You're adding a lives counter and a lose message, both built on top of code you already trust. If you want to try the Level 3 sound idea, the sound files are already in this week's Assets folder - no hunting for your own.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"Test in both directions today - shooting shouldn't touch lives, and escaping shouldn't touch your kill counter. Your handout's Level 2 section walks through exactly how to check this.\""),
        p("\"Questions before we start lab? ... Let's put some pressure on your player.\""),

        h2("Lab Time Notes"),
        bullet("The single most important thing to verify this week: shoot an enemy (lives should NOT change) and separately let one escape (the kill counter should NOT change). Check both directions, not just that the game 'seems to work.'"),
        bullet("Most likely mistake: the lives decrement added to the wrong event/branch (hit-detection instead of escape-cleanup, or vice versa) - this is a copy-paste-adjacent mistake, easy to make quickly and easy to miss without deliberately testing both directions."),
        bullet("No engine needs a structural bugfix this week, unlike Lesson 7 - circulate normally, no cluster needs extra attention by default."),
        bullet("A student who lets every enemy through without shooting any will hit Game Over quickly - a good built-in test case, not a bug, worth pointing out if a student is confused why it happened so fast."),
        bullet("Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout (a live lives readout is a satisfying one), not given new unrelated work."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_08_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
