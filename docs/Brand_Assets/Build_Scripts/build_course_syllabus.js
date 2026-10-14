const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, VerticalAlign
} = require("docx");
const brand = require("./brand.js");

const CONTENT_WIDTH = 9360;
const border = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function link(text, url) {
  return new ExternalHyperlink({ children: [new TextRun({ text, style: "Hyperlink" })], link: url });
}
function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function pRuns(runs, opts = {}) { return new Paragraph({ spacing: { after: 160 }, ...opts, children: runs }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function labelBullet(label, text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)] });
}
function headerCell(text, width) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: brand.NAVY2, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 20 })] })] });
}
function bodyCell(text, width, opts = {}) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, size: 20, bold: !!opts.bold })] })] });
}
function simpleTable(headers, widths, rows) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((t, i) => headerCell(t, widths[i])) }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((t, i) => bodyCell(t, widths[i], { fill: ri % 2 ? "F3F6FA" : "FFFFFF", bold: i === 0 }))
      }))
    ]
  });
}

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]
};

const semesterHeaders = ["Weeks", "Focus", "Milestone"];
const semesterWidths = [1700, 5660, 2000];
const semesterRows = [
  ["1", "Orientation: internet safety, game dev overview, platform selection", "Engine installed & first project opened"],
  ["2-6", "Programming foundations: variables, events, loops, conditionals, input", "First moving, controllable character"],
  ["7-11", "Core mechanics: collision, scoring, win/lose states, simple UI", "Small playable prototype"],
  ["12-15", "Game feel & polish: audio, basic animation, level structure", "Semester 1 capstone: complete Mini-Game"],
  ["16-20", "Intermediate systems: state machines, enemy/AI behavior, physics", "Prototype with opponents/obstacles"],
  ["21-25", "Expanding scope: menus, save data, multiple levels, juice/polish", "Feature-complete draft of capstone game"],
  ["26-28", "Playtesting, bug-fixing, presentation prep", "Playtested, stable build"],
  ["29-30", "Finalize, package, and export shareable copies", "Shareable build ready for CoOp Expo"],
];

const policyItems = [
  ["Internet & Digital Safety", "Lesson 1 opens with an internet-safety unit grounded in NetSmartz materials, covering safe sharing, online interaction with strangers (including in multiplayer/online game contexts), and responsible use of school/lab wifi. Digital citizenship expectations are revisited whenever a lesson touches online accounts, multiplayer features, or publishing/sharing a game."],
  ["BYOD Expectations", "Students bring a personal laptop meeting the minimums in the Engine & Platform Guide. Please charge devices before class and bring a charger - lab time is hands-on. Install large engines at home ahead of time when possible; shared classroom wifi is not fast enough for 20 students downloading Unity or Unreal simultaneously."],
  ["Engine Choice", "Students select their own engine/platform from the supported list (see Engine & Platform Guide) and may change engines later in the year without penalty - every lesson provides working examples across all supported engines. Note: 001 Game Creator requires Windows, so Mac/Chromebook students should choose a different engine."],
  ["AI-Assisted Coding", "Because this course sits at the intersection of programming education and modern development practice, thoughtful use of AI coding assistants is permitted and, in upper-level lessons, taught directly as a skill - but students must be able to explain what their code does and may not submit AI output they don't understand. Specific per-assignment guidance will be given as needed."],
  ["Academic Integrity", "Collaboration and helping classmates debug is encouraged during lab time. All capstone work should represent each student's own code and implementation, with credit given for any tutorials, assets, or AI assistance used. For the Semester 1 Mini-Game, the game design itself is the same for the whole class (see \"This Semester's Game\" in your Lesson 1 Game Design Document) - what's yours is your code, your art/sound choices, and how far into Level 3 you push. The Semester 2 Expo game is fully each student's own design as well as their own code."],
  ["Attendance & Make-Up Work", "As a homeschool co-op class, families manage their own attendance policies, but because lessons build week-to-week, missed lab time should be made up at home before the next session. Lesson materials and recordings of key demonstrations (where practical) will be made available."],
  ["Conduct", "Standard co-op classroom conduct applies during lecture and lab time. Respectful, patient help for classmates using a different engine than you is part of the culture of this class."],
];

const levelHeaders = ["Level", "Think of it as...", "What it means"];
const levelWidths = [1300, 2200, 5860];
const levelRows = [
  ["Level 1 - Minimal", "MVP (Minimum Viable Product)", "The smallest working version of that lesson's concept - enough for the student to understand it and move on to next week's lesson. This is a full, legitimate completion of the lesson, not a fallback."],
  ["Level 2 - Complete", "Internal Alpha/Beta", "All of the lesson's expected features are present and working end-to-end. Art, polish, and final balancing may still be missing, and minor bugs may remain."],
  ["Level 3 - Refined", "Ready to Publish", "Feature-complete, tested, and polished enough that the student would be comfortable sharing it outside of class."],
];

const doc = new Document({
  numbering: numberingConfig,
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: brand.NAVY },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: brand.NAVY2 },
        paragraph: { spacing: { before: 260, after: 140 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "444444" },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: brand.pageHeader("Course Syllabus") },
    footers: { default: brand.pageFooter() },
    children: [
      brand.crestParagraph(1.3, { after: 160 }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Learn Programming by Making Video Games", bold: true, size: 40, color: brand.NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Course Syllabus - 2026/27 School Year", size: 28, color: brand.NAVY2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "SLHE Homeschool Co-op  •  Instructor: Robert Kuropkat  •  robert.kuropkat@gmail.com", italics: true, size: 20, color: "666666" })] }),

      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
      new Paragraph({ children: [new PageBreak()] }),

      h1("Course Description"),
      p("This course teaches programming fundamentals through hands-on game development. Students choose a game engine or platform that fits their device and interests, then build increasingly capable games across a full school year, culminating in a finished, shareable project for the co-op's CoOp Expo. Rather than locking the class into a single tool, each week's core lecture rotates through a different supported engine - Scratch, Godot, Unity, GameMaker, GDevelop, TIC-80, 001 Game Creator, and (optionally, for capable hardware) Unreal Engine - while providing working example code in every supported engine for every lesson. Students may switch platforms at any point. Note: 001 Game Creator's editor is Windows-only, so Mac/Chromebook students will use a different engine during weeks it's featured (see the Engine & Platform Guide)."),

      h1("Meeting Schedule"),
      labelBullet("Format", "2 hours per meeting: approximately 1 hour lecture/demo, 1 hour hands-on lab"),
      labelBullet("Frequency", "Weekly, 30 total meetings across the school year"),
      labelBullet("Semester 1", "15 meetings, beginning Thursday, August 13, 2026"),
      labelBullet("Semester 2", "15 meetings, concluding with CoOp Expo project preparation"),
      labelBullet("Homework", "Light take-home practice is expected between sessions, especially finishing in-lab exercises and installing/updating software"),

      h1("Audience & Prerequisites"),
      p("Open to Middle and High School students. No prior programming experience is required. Because the class spans a wide age and maturity range, lessons are designed with flexible depth: younger or newer students can lean on block/visual coding (Scratch, GameMaker Visual, GDevelop, 001 Game Creator) and simpler project scope, while older or more experienced students are encouraged toward text-scripted engines (Godot, Unity, TIC-80, or Unreal) and more ambitious builds. Everyone works toward the same learning objectives each week at a level appropriate to them - see Assessment Approach below."),

      h1("What You'll Need"),
      p("A personal laptop (Bring Your Own Device) meeting the minimum specifications for your chosen engine, and a way to save/back up project files (USB drive or cloud storage). Full hardware guidance, install instructions, and a device-fit chart are in the companion Engine & Platform Guide, which is sent home before Lesson 1 so installation can begin ahead of time."),

      h1("Curriculum Alignment"),
      p("Course content is organized around the IGDA Game Education SIG's Curriculum Framework for the Study of Games and Game Development, adapted for a Middle/High School BYOD setting. The framework's most recent published version is v3.2 beta (2008); the IGDA Game Education SIG is actively revising it as of 2026, with a second public comment period running May 30-September 1, 2026 and publication targeted for September 30, 2026. This syllabus will be reviewed against the updated framework once published, likely partway through Semester 1."),
      p("Where a student's chosen engine is Unity, lesson content also maps to Unity's free Unity Learn Pathways - Unity Essentials, Junior Programmer, and Creative Core - so that a motivated student can work toward an official Unity Certification alongside class material."),

      h1("Semester-by-Semester Overview"),
      p("A week-by-week breakdown, including which engine is featured each week, lives in the companion 30-Lesson Master Schedule. The arc below shows the shape of the year:"),
      simpleTable(semesterHeaders, semesterWidths, semesterRows),
      new Paragraph({ spacing: { before: 300, after: 200 }, children: [] }),

      h1("CoOp Expo Capstone"),
      p("The final weeks of Semester 2 (roughly Lessons 26-30) shift from new material to finishing, playtesting, and packaging each student's game so it can be demonstrated at the co-op's Expo and shared with friends and classmates. Students will export a runnable build (or shareable link, for browser-based engines), prepare a short demo/pitch, and get structured playtesting feedback from classmates before the Expo. The exact Expo date is not yet set for 2026/27; once it is confirmed, these final lessons may shift earlier or later in the 30-Lesson Master Schedule to line up with it."),

      h1("Assessment Approach"),
      p("This is a project- and mastery-based course, not a test-driven one, and it does not use a traditional numeric grading rubric. Because students span a wide range of ages and prior experience, every lesson's hands-on work - and both semester milestones - are structured around three levels of accomplishment instead:"),
      simpleTable(levelHeaders, levelWidths, levelRows),
      new Paragraph({ spacing: { before: 200, after: 200 }, children: [] }),
      p("Reaching Level 1 is a complete, legitimate finish for that lesson and is always enough to move on to the next one - no student is ever stuck waiting to \"catch up.\" Level 2 and Level 3 are open invitations for students with more time, skill, or interest to go further, not extra assignments or busy work. Each lesson's Student Handout spells out what Level 1, 2, and 3 look like for that specific exercise, and the Semester 1 Mini-Game and Semester 2 Expo game are evaluated the same way."),

      h1("Policies"),
      ...policyItems.flatMap(([title, text]) => [h3(title), p(text)]),

      h1("Materials & Companion Documents"),
      labelBullet("Engine & Platform Guide", "engine comparison, hardware fit, and install instructions - send home before Lesson 1"),
      labelBullet("30-Lesson Master Schedule", "week-by-week topics, featured engine rotation, and IGDA framework tie-ins"),
      labelBullet("Lesson Production Standards", "defines the three-level accomplishment framework and the standard set of materials produced for every lesson"),
      labelBullet("Engine Installation & Hello World Verification Guide", "detailed setup instructions and a working test project for each supported engine, for instructor prep"),
      labelBullet("Per-Lesson Materials", "for each lesson: a 1-page lesson plan, a teacher's script, a student handout per engine, and a teacher's prep checklist"),
      labelBullet("Student Resource List", "supplemental tutorials, references, and practice sites by engine"),

      h1("Contact"),
      p("Questions about the course, engine choice, or make-up work can be directed to Robert Kuropkat at robert.kuropkat@gmail.com."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Course_Syllabus.docx", buffer);
  console.log("done");
});
