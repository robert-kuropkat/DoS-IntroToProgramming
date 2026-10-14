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
    ["0:00-0:05", "Welcome back, Lesson 11 recap, the polish phase begins"],
    ["0:05-0:15", "The Sound Payoff: Four Sounds, Four Existing Moments"],
    ["0:15-0:35", "One Shared Sound for Three Outcomes (live demo in GameMaker)"],
    ["0:35-0:45", "Why None of Today's Sounds Need a Retrigger Guard"],
    ["0:45-0:55", "Introduce lab exercise, hand out engine-specific handouts"],
    ["0:55-1:50", "LAB: Wire up all four sounds, test a full round by ear"],
    ["1:50-2:00", "Exit check / quick share-out - play a round for the class"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 12 - Hearing It Happen", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 12 - Sound Effects", italics: true, size: 18, color: brand.GOLD_TEXT, bold: true })] }),

        labelBullet("Objectives", "Wire four sound cues (shoot, explosion, wave clear, game over) into the exact code paths that already represent those events; recognize why event-driven trigger points don't need a manual replay guard; formalize sound effects that were optional 'Level 3 stretch' ideas in Lessons 2-6 into a required capstone feature."),
        labelBullet("IGDA Tie-In", "Audio - sound design fundamentals, diegetic vs. non-diegetic cues (introductory); Programming - event-driven triggers, code reuse (intermediate)"),
        labelBullet("Featured Engine (lecture demo)", "GameMaker (GML Visual - second rotation slot, callback to Lesson 5)"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified and audio output working (see Teacher's Checklist); Student Handouts, one per engine; students' own Lesson 11 projects (already saved); two new sound assets (sfx_wave.wav, sfx_lose.wav) - no other new assets needed this week"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): all four sounds play at the correct moment at least once - shoot, explosion/hit, wave clear, game over."),
        bullet("Level 2 (Complete): all four fire reliably every time with no missing or duplicate playback, and none interfere with each other on the same frame."),
        bullet("Level 3 (Refined): add the optional win sound if not already built, add pitch/volume variety, and/or tidy the sound-triggering code."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 12 Teacher's Script, Lesson 12 Prep Checklist, and the per-engine Student Handouts. This is a lighter lesson than most - mostly wiring existing branches, not new mechanical complexity. Bring headphones or speakers for the lecture demo.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_12_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 12 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 12 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Hearing It Happen", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Last week you hit Level 2 Complete - a real, replayable game with multiple waves. Today's a lighter lesson, and honestly, one of the most fun ones all semester: your game is about to start sounding like a game.\""),
        p("\"From here through Lesson 15, we're not adding new mechanics anymore. We're making the mechanics you already have feel good - sound this week, animation next week, difficulty scaling after that, and a final polish pass before the capstone.\""),

        h2("[0:05-0:15] The Sound Payoff: Four Sounds, Four Existing Moments"),
        p("\"Quick question - by a show of hands, who's built one of the optional 'add a sound' ideas along the way? Lesson 2's movement beep, Lesson 3's shoot sound, Lesson 4's hit sound, Lesson 6's victory sound?\" [Pause, note roughly how many hands go up]"),
        p("\"However many hands went up, here's the deal: today, four sounds become required for everyone - a shot, an explosion, a wave-clear cue, and a game-over cue. If you already built some of these, you're mostly just verifying and adding two brand-new ones. If you didn't build any, today's Handout gives you everything from scratch - it's written to work either way.\""),
        p("\"Here's the key idea: every one of these four sounds gets added to a spot in your code that ALREADY exists - the exact moment you fire a shot, the exact moment a bullet hits an enemy, the exact branch where a wave advances, the exact branch where you lose. No new logic. Just one more line at each of four places you already built.\""),
        p("[Timing checkpoint: wrap by 0:15]", { italics: true, color: "888888" }),

        h2("[0:15-0:35] One Shared Sound for Three Outcomes"),
        p("[Demo cue: switch to GameMaker, open a project with working multi-wave logic from Lesson 11, GML Visual tab open on obj_bullet's Collision event]"),
        p("\"Shoot and hit sounds are easy - they already live on obj_player and obj_bullet, one sound each, and if you built them back in Lessons 3 and 4, they need zero changes today. The more interesting part is the three sounds that can happen when a wave ends: a win, a wave-clear, or a loss. Only one of those three can ever be true at once - so I don't need any shared node or extra bookkeeping to keep them from overlapping. I just drop an Audio Play Sound block directly into each of the three branches that already decide which outcome happened.\""),
        p("\"Watch: right on Lesson 11's If Variable block - the one checking whether this was the last wave - I add an Audio Play Sound block (sfx_win) on the True output, and a second Audio Play Sound block (sfx_wave) on the False output. Then, on obj_enemy's escape branch, one more Audio Play Sound block (sfx_lose) inside the lives-reaches-zero check.\" [Build the three blocks live, demonstrate a wave-clear sound followed by a win sound on the final wave]"),
        p("\"This is the same reuse idea you've seen all semester, just applied to a sound block instead of a function - and it's a nice example of GameMaker's block-for-block honesty: three blocks, three branches, no shared object standing in the middle of it. Godot and Unity students do the same idea slightly differently, routing all three clips through one shared audio player node instead - different mechanism, same underlying idea of one outcome, one sound, no overlap.\""),
        p("[Timing checkpoint: wrap by 0:35]", { italics: true, color: "888888" }),

        h2("[0:35-0:45] Why None of Today's Sounds Need a Retrigger Guard"),
        p("\"Here's something worth noticing: none of today's four sounds need any kind of 'don't replay this every frame' guard. Does anyone remember why Lesson 2's beep-beep sound needed an explicit stop call, but a one-shot sound like today's shoot or hit sound doesn't?\" [Pause for answers] \"Right - because a one-shot sound fires once and finishes on its own, it doesn't need turning off. But there's a second reason today's four sounds are safe: every single one of them lives inside a branch that ONLY runs once per actual event - a real hit, a real wave transition, a real loss. Nothing here is being checked every single frame.\""),
        p("\"TIC-80 students specifically: if you built Lesson 6's optional win sound, you may have needed a guard flag back then, because the check ran every frame in a continuous loop. As of Lesson 10 and 11's restructuring, that's no longer true - the win-check now lives inside a one-time hit branch. If you've still got that old guard flag lying around, it's safe to remove it today.\""),
        p("[Timing checkpoint: wrap by 0:45]", { italics: true, color: "888888" }),

        h2("[0:45-0:55] Today's Exercise & Handouts"),
        p("\"For lab today, open your own project from last week - not a new one. Two new sound files are provided (sfx_wave.wav, sfx_lose.wav); everything else you may already have from earlier optional ideas.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"Goal for today: play a full round out loud - a shot, a hit, an early wave clear, and separately, a loss - and confirm all four sound genuinely different from each other. If you already had a win sound, make sure it still plays correctly and doesn't overlap the wave-clear sound.\""),
        p("\"Questions before we start lab? ... Let's make some noise.\""),

        h2("Lab Time Notes"),
        bullet("The most common issue this week isn't a bug so much as a missing step - remind students who skipped the optional Lessons 2/3/4/6 ideas that today's Handout gives them full instructions, not just 'add two sounds.'"),
        bullet("Watch for a sound call placed outside the correct if/else branch, causing two round-ending sounds to play together (e.g. wave-clear and win firing on the same frame)."),
        bullet("This is a good week to actually listen as you circulate, not just watch the screen - a wrong or missing sound is often faster to catch by ear than by reading code."),
        bullet("If a student's speakers/headphones aren't working, sound logic can still be verified visually for now (e.g. a temporary print/debug line) - don't let a hardware issue block progress on the actual lesson goal."),
        bullet("This is a lighter week - if a cluster finishes early, point them at the Level 3 ideas (pitch variation, the optional win sound if not yet built) rather than moving ahead to next week's material."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_12_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
