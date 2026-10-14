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
    ["0:00-0:05", "Welcome back, Lesson 8 recap"],
    ["0:05-0:15", "The Payoff of Deferring UI"],
    ["0:15-0:35", "Score: Reusing the Hit-Detection Path (live demo in GameMaker)"],
    ["0:35-0:50", "Two Ways to Stay in Sync: Push vs. Poll"],
    ["0:50-0:55", "Introduce lab exercise, hand out engine-specific handouts"],
    ["0:55-1:50", "LAB: Add score tracking and a real on-screen UI to your project"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 9 - Core Mechanics III: On the Scoreboard", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 9 - Score Tracking + Real On-Screen UI", italics: true, size: 18, color: brand.GOLD_TEXT, bold: true })] }),

        labelBullet("Objectives", "Add score tracking to the existing hit-detection code path; build a persistent, always-visible HUD using each engine's real UI tool; explain the difference between push-based and poll-based display updates; recognize that some engines' Lesson 6-8 text tools were already the 'real' UI all along."),
        labelBullet("IGDA Tie-In", "Programming - UI/UX implementation, state display, event-driven vs. polling architecture (intermediate)"),
        labelBullet("Featured Engine (lecture demo)", "GameMaker (GML script/text mode - callback to Lesson 5's GameMaker Visual demo)"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified (see Teacher's Checklist); Student Handouts, one per engine; students' own Lesson 8 projects (already saved) - no new assets needed this week"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): destroying an enemy adds 10 points to score; score and lives are both visible on screen at all times."),
        bullet("Level 2 (Complete): the score/lives readouts never show a stale or wrong number - they always match reality, immediately."),
        bullet("Level 3 (Refined): style the HUD, add scoring variety, and/or tighten the separation between game logic and display code."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 9 Teacher's Script, Lesson 9 Prep Checklist, and the per-engine Student Handouts. Unity, Unreal, and 001 Game Creator involve genuinely new UI setup this week - budget extra circulation time for those clusters.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_09_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 9 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 9 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Core Mechanics III: On the Scoreboard", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Last week your game got real stakes - lives, and a lose condition. Quick show of hands - who tested both directions, shooting AND escaping, to make sure the two counters stayed separate?\" [Pause, call on 1-2 volunteers]"),
        p("\"Today, your game finally gets a real scoreboard - not just a message at the end, but score and lives visible the whole time you're playing. Today's demo is also back in GameMaker - same engine as Lesson 5, but this time in its typed GML mode instead of GML Visual, so keep an eye out for how much of what you already learned there still applies.\""),

        h2("[0:05-0:15] The Payoff of Deferring UI"),
        p("\"Way back in Lesson 6, I told you something on purpose: 'a plain line of text is completely legitimate today - the real UI system comes later, once there's a score and lives to actually display.' That was a promise. Today we keep it.\""),
        p("\"Here's the interesting part: for some of you, keeping that promise is almost free. Your engine's win/lose message tool was ALREADY the real tool - it just needs to run all the time instead of once. For a couple of you, today introduces something genuinely new.\""),
        p("[Discussion question] \"Why do you think we waited until now to build a real UI, instead of doing it back in Lesson 6 alongside the win message?\" (Draw out: there was nothing worth displaying continuously yet - a UI showing nothing changing isn't useful; build the data first, display it once it means something.)"),
        p("[Timing checkpoint: wrap by 0:15]", { italics: true, color: "888888" }),

        h2("[0:15-0:35] Score: Reusing the Hit-Detection Path"),
        p("\"First, the easy part: score. Every engine's hit-detection code already runs exactly once per confirmed kill - the same place enemies_remaining already changes. Score just joins it.\""),
        p("[Demo cue: switch to GameMaker, open a project with working lives from Lesson 8]"),
        p("\"Watch: right in obj_bullet's Collision event with obj_enemy - the exact same event where enemies_remaining already goes down - I add one line: obj_gamemanager.__dnd_score += 10. That's the whole scoring system (__dnd_score is the same variable GameMaker's drag-and-drop Score blocks use, which is why it has that odd name). The hard part isn't the math, it's showing it.\" [Add the __dnd_score += 10 line, show it's not visible yet]"),
        p("\"To actually SHOW score and lives, watch where I put the display code: obj_gamemanager's Draw GUI event - the exact same event that's already been drawing 'You Win!' and 'Game Over' since Lesson 6. I add two plain draw_text() calls there, with no if-check around them at all, so they just run every single frame no matter what.\" [Add the two unconditional draw_text() calls, run the game, shoot an enemy and let one escape to show both numbers update live]"),
        p("[Timing checkpoint: wrap by 0:35]", { italics: true, color: "888888" }),

        h2("[0:35-0:50] Two Ways to Stay in Sync: Push vs. Poll"),
        p("\"What I just showed you in GameMaker is called POLL - instead of remembering to push an update every time something changes, you just check the data every single frame, no matter what. GameMaker's Draw GUI event and TIC-80's print() calls have quietly been doing this since Lesson 6 - today, Unity's HUD script does it on purpose too.\""),
        p("\"But there's a second, equally valid way, and most of the rest of you will use it today: PUSH. Every time the data actually changes, you push an update to the display, instead of checking it constantly. Godot, GDevelop, Unreal, and 001 Game Creator's real-time Field binding all work this way.\""),
        p("[Demo cue, optional: pull up Godot briefly] \"Neither approach is 'more correct.' Poll means you never forget to update, but you're technically doing a little more work than necessary, every frame, forever. Push means less wasted work, but you have to remember to call the update everywhere something changes. For two small numbers like these, the cost either way is nothing.\""),
        p("[Discussion question] \"Can you think of something in a game where checking every single frame would start to feel wasteful, if there were hundreds of things to check instead of two?\" (Open-ended; validate any reasonable guess - this previews ideas about performance they'll meet again in later, more advanced work.)"),
        p("[Timing checkpoint: wrap by 0:50]", { italics: true, color: "888888" }),

        h2("[0:50-0:55] Today's Exercise & Handouts"),
        p("\"For lab today, open your own project from last week - not a new one, and no new assets needed. You're adding score tracking and a real, always-visible HUD.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"Unity, Unreal, and 001 Game Creator folks - your setup today is more involved than usual, genuinely new UI tools. That's expected, not a sign you're behind. Everyone else - your engine's already had the right tool since Lesson 6, today's just about using it continuously.\""),
        p("\"Questions before we start lab? ... Let's get your scoreboard up.\""),

        h2("Lab Time Notes"),
        bullet("Unity, Unreal, and 001 Game Creator clusters need the most attention this week - budget extra circulation time there, and expect setup-related questions before logic questions."),
        bullet("001 Game Creator: the single most likely mistake is forgetting to tick \"Update Real-time\" on a Field - it looks like a mysterious bug (frozen numbers) rather than an obviously missing setting. Check this first if a student's HUD isn't updating."),
        bullet("Unity: NullReferenceException on Play almost always means a Text object wasn't actually dragged into the HUD script's Inspector fields - check this first."),
        bullet("A score readout stuck at 0 is almost always a missing or misplaced increment, not a display bug - have students confirm the underlying variable is actually changing before troubleshooting the display."),
        bullet("Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout (HUD styling is a satisfying, visible one), not given new unrelated work."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_09_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
