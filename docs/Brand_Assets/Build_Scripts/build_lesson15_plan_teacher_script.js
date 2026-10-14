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
    ["0:00-0:10", "Welcome back, capstone reflection, milestone framing"],
    ["0:10-0:55", "POLISH TIME: work the full capstone checklist"],
    ["0:55-1:00", "Presentation logistics, running order/sign-up"],
    ["1:00-1:50", "PRESENTATIONS: ~2 minutes per student"],
    ["1:50-2:00", "Closing celebration, look ahead to Semester 2 (incl. testing becoming ongoing)"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 15 - Ready to Show", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 15 - Level 3 Refined Milestone / Capstone Presentation", italics: true, size: 18, color: brand.GOLD_TEXT, bold: true })] }),

        labelBullet("Objectives", "Self-assess a finished capstone game against the full Level 1/2/3 checklist accumulated since Lesson 2; prioritize and complete a final polish pass; present a working game confidently to the class."),
        labelBullet("IGDA Tie-In", "Production - scoping, finishing, presenting a project (intermediate); all Five Disciplines, in review"),
        labelBullet("Format", "No lecture demo this week - a workshop-and-showcase day. No new mechanics are taught."),
        labelBullet("Materials", "Demo laptop with a finished example project per engine (backup only); Student Handouts (checklists), one per engine; students' own Lesson 14 projects - this IS the capstone, nothing new to bring"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "This lesson's \"Levels\" are a milestone checkpoint, not a new exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal) reached: core mechanics (Lessons 2-11) all work - move, shoot, score, lives, win/lose, restart, multiple waves."),
        bullet("Level 2 (Complete) reached: sound (Lesson 12) and hit reactions (Lesson 13) are both fully wired in."),
        bullet("Level 3 (Refined) reached: difficulty scaling (Lesson 14) is present, and at least a good-faith attempt at one chosen stretch feature exists - this is today's actual milestone, matching the capstone GDD's Feature-to-Lesson Map."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 15 Teacher's Script, Lesson 15 Prep Checklist, and the per-engine Student Handouts (checklists + presentation guide). This is a milestone lesson - plan real time for celebration, same as Lessons 6 and 11.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_15_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 15 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 15 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Ready to Show", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:10] Welcome Back - Capstone Reflection"),
        p("\"Welcome back! Today's different from every other day this semester. There's no new code to learn. Today you finish, and today you show what you built.\""),
        p("\"Take a second and think back to Lesson 1. You picked an engine you'd maybe never touched before. Fifteen weeks later, you have a real game - it moves, it shoots, it keeps score, it can be won and lost, it sounds and looks like something, and it gets harder the longer you survive. That's genuinely a lot.\""),
        p("\"Here's today's plan: about 45 minutes of polish time, using the checklist in your Handout to find and fix whatever's roughest. Then presentations - about 2 minutes each, just play a bit of your game and tell us about it. No slides, no stress.\""),

        h2("[0:10-0:55] Polish Time"),
        p("[Circulate continuously during this block - see Lab Time Notes below for what to prioritize]"),
        p("\"Work through your Handout's checklist top to bottom. Anything unchecked is your first priority - a complete, simple game beats a broken ambitious one, every time. Only once your checklist is solid should you spend remaining time on your Lesson 14 stretch feature or any extra polish.\""),
        p("[Timing checkpoint: give a 15-minutes-remaining warning around 0:40, and a 5-minutes-remaining warning around 0:50]", { italics: true, color: "888888" }),

        h2("[0:55-1:00] Presentation Logistics"),
        p("\"Here's how presentations will work today: [explain your chosen format - whole-class seat-to-seat, or a running order at the front]. Each of you gets about 2 minutes. Play a short round, name your game, and point out the one feature you're proudest of.\""),
        p("[Set the running order / sign-up now if presenting to the front of the room]"),

        h2("[1:00-1:50] Presentations"),
        p("\"Let's go! Remember - if something breaks live, that's genuinely fine. Just say what it's supposed to do and keep going.\""),
        p("[Hold the 2-minute limit firmly - give a 30-second warning if a student is running long, and a clear \"time\" cue at the limit]"),
        p("[Consider a brief specific, genuine comment after each presentation - naming something the student clearly worked hard on goes a long way here]"),

        h2("[1:50-2:00] Closing Celebration & What's Coming When You Get Back"),
        p("\"That's every one of you, with a real, finished game, in an engine you chose. However far up the Level ladder you got, you built something that works, start to finish, over fifteen weeks. That's the whole point of today - not a perfect game, a FINISHED one.\""),
        p("\"Semester 2 picks up in a very different place. A few of the big things waiting for you: state machines and simple AI for enemies (patrol, chase, flee - not just \\\"sit there until touched\\\"), real physics (gravity, forces, bounces), multiple levels with actual progression, menus and save data, and eventually a much bigger capstone build for the CoOp Expo. Everything you already know - variables, conditionals, loops, events, structure - carries forward completely unchanged. You're not starting over, you're building on top of a semester's worth of real skill.\""),
        p("\"One thing that's genuinely new starting next semester: testing becomes a regular part of the process, not just something you do once at the end. You've actually been tracking this all along without necessarily noticing - every week's GDD has had a 'planned / implemented / tested' checkbox, and if you're honest, that third column has probably lagged behind the other two most weeks. That's completely normal for a first semester learning to code. Starting in Semester 2, we're going to close that gap on purpose: playing your own game with a critical eye, having classmates try it and watching where they get confused or stuck, and actually fixing what you find - on an ongoing basis, most weeks, not saved up for one big session at the end. It doesn't need to be formal or complicated. It's mostly just: try to break your own game before someone else does, and take it seriously when they find something you didn't.\""),
        p("\"That's it for Semester 1 - genuinely well done. See you in the new year.\""),

        h2("Lab Time Notes (Polish Block)"),
        bullet("Redirect firmly toward fixing and tuning over adding anything new - this is the single most important thing to watch for during polish time."),
        bullet("Triage by checklist category: Core Mechanics broken > Sound/Animation broken > Difficulty/Level 3 incomplete, in that priority order, if a student is short on time."),
        bullet("For students significantly behind, help them reach a confidently presentable Level 1/2 state rather than chasing every Level 3 item - see this week's Teacher's Key for engine-specific guidance."),
        bullet("This is a good week to let students help each other across engine clusters informally, if two students on the same engine are both stuck on the same kind of issue."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_15_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
