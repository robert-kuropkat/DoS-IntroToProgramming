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
function pRuns(runs) { return new Paragraph({ spacing: { after: 160 }, children: runs }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function numbered(text, level = 0) { return new Paragraph({ numbering: { reference: "numbers", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
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
    width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: widths,
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

const agendaHeaders = ["Time", "Segment"];
const agendaWidths = [2000, 7360];
const agendaRows = [
  ["0:00 - 0:10", "Welcome, roll call, course overview"],
  ["0:10 - 0:30", "Internet Safety & Digital Citizenship (NetSmartz)"],
  ["0:30 - 0:45", "What Is Game Development? (roles & disciplines overview)"],
  ["0:45 - 1:00", "Live Engine Tour: Key Press Reactor across all 8 engines"],
  ["- Lab Time -", ""],
  ["1:00 - 1:45", "Guided install/setup of chosen engine, by station"],
  ["1:45 - 1:52", "Exit check: confirm working environment, quick share-out"],
  ["1:52 - 2:00", "This Semester's Game: High Concept reveal, GDD hand-out, codename vote"],
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
    headers: { default: brand.pageHeader("Lesson 1 Plan") },
    footers: { default: brand.pageFooter() },
    children: [
      brand.crestParagraph(1.3, { after: 160 }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 1", bold: true, size: 40, color: brand.NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Welcome, Internet Safety & Choosing Your Platform", size: 28, color: brand.NAVY2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Thursday, August 13, 2026", italics: true, size: 22, color: "666666" })] }),

      h1("At a Glance"),
      labelBullet("Objectives", "Students understand class structure and expectations; can identify at least 2 key online-safety practices specific to gaming; can name the major disciplines of game development; have chosen (or shortlisted) an engine and made progress installing it"),
      labelBullet("Materials needed", "Projector/demo laptop with all 8 engines pre-installed AND the Level 1 Key Press Reactor pre-built in each (see Instructor Prep), NetSmartz video(s) queued and tested, printed or emailed Engine & Platform Guide and Engine Showcase (sent to families before this lesson), sign-in sheet, offline installer backups on USB (see Instructor Prep), printed or emailed Lesson 1 GDD (sneak peek) - one per student, 3-4 candidate codenames prepared for the end-of-class vote"),
      labelBullet("IGDA framework tie-in", "Course/discipline orientation (Production & Critical Studies content areas) and first exposure to the Programming content area"),
      h2("Agenda"),
      simpleTable(agendaHeaders, agendaWidths, agendaRows),
      new Paragraph({ spacing: { before: 300, after: 200 }, children: [] }),

      h1("Lecture (60 minutes)"),

      h2("A. Welcome & Course Overview (10 min)"),
      p("Introduce yourself and the course. Set expectations: this is a hands-on, project-based class; students choose their own engine and can switch later; every lesson provides examples across all supported engines. Walk through the Syllabus at a high level - meeting schedule, the two capstone milestones (Semester 1 mini-game, Semester 2 CoOp Expo game), and policies (BYOD, AI-assistance, conduct). Tease that everyone finds out exactly what that Semester 1 game is at the very end of today's class - a small hook to carry through the rest of the lesson."),

      h2("B. Internet Safety & Digital Citizenship (20 min)"),
      p("Grounded in NetSmartz (National Center for Missing & Exploited Children), with a gaming-specific lens since students will spend the year building and, later, sharing/playing each other's games."),
      h3("Suggested materials"),
      bullet("NSTeens “Attitude Overdrive” - appropriate conduct and etiquette while gaming online (good fit for middle schoolers)"),
      bullet("“Into the Cloud” Season 2 - covers gaming and livestreaming safety, reporting/removing inappropriate content, cyberbullying, and seeking help from a trusted adult (good fit for a mixed middle/high school room)"),
      bullet("Real-Life Stories series - for older teens, examines peer choices online and real consequences (use selectively with high schoolers if time allows)"),
      pRuns([new TextRun("Preview all clips ahead of time and check "), link("missingkids.org/netsmartz", "https://www.missingkids.org/netsmartz/home"), new TextRun(" for the current video library, since NetSmartz periodically updates its catalog.")]),
      h3("Discussion points to cover"),
      bullet("Chatting with strangers in multiplayer/voice chat - what's okay to share, what isn't"),
      bullet("Recognizing griefing, cyberbullying, and manipulation in gaming communities, and how/when to report or block"),
      bullet("In-game purchases, scams, and “too good to be true” offers"),
      bullet("Once students start sharing games with each other later this year: don't include personal info in usernames, credits, or shared builds"),
      bullet("Who to go to (parent, instructor) if something online feels wrong"),

      h2("C. What Is Game Development? (15 min)"),
      p("A fast overview of the disciplines that make a game, so students understand where coding fits into the bigger picture (this class focuses on Programming, with light Design):"),
      bullet("Design - deciding what the game is and how it plays"),
      bullet("Programming - making the rules and mechanics actually work (our main focus)"),
      bullet("Visual Art - characters, environments, UI"),
      bullet("Audio - music and sound effects"),
      bullet("Production - planning, scoping, and finishing a project (relevant all year, critical in the final weeks before the Expo)"),
      p("Optional: show 2-3 short, well-known indie games (or screenshots) and ask students to spot examples of each discipline."),

      h2("D. Live Engine Tour: Key Press Reactor (15 min)"),
      p("Walk through the comparison table from the Engine & Platform Guide, then show one consistent live demo across all 8 supported engines: Scratch, Godot, Unity, GameMaker, GDevelop, TIC-80, 001 Game Creator (note its Windows-only editor), and Unreal (framed as the optional advanced track for students with strong hardware). For each engine, open the Level 1 \"Key Press Reactor\" mini-project you pre-built ahead of time (see Instructor Prep and that engine's Lesson 2 Teacher's Key), press Play, and tap an arrow key so students watch the same small reaction - a character responding to a key press - happen 8 different ways. Budget about 60-90 seconds per engine: press Play, tap the key, name which coding style you just used (Block, Visual, or Scripting - see the Engine & Platform Guide's Hybrid Engines section), move on. Do not build these live from scratch - the whole point is a fast, consistent comparison, not a coding demo."),
      p("This is deliberately the same exercise students will build hands-on themselves next week in Lesson 2, so it doubles as a preview. If you sent the Engine Showcase document home before this lesson (see Instructor Prep), most students will already have seen still screenshots of this same exercise - today is their first time seeing it actually run.", { italics: true, color: "555555" }),
      p("Close by explaining the device-fit guidance so students with older laptops feel set up for success, not steered toward something that will frustrate them."),

      h1("Lab (~53 minutes): Install & Setup"),
      p("Goal: every student leaves with a working environment for their chosen engine - or, at minimum, a browser-based fallback (GDevelop or Scratch) open and ready, so no one leaves lab time empty-handed."),
      h3("Suggested structure"),
      numbered("Quick poll: which engine is each student planning to start with? (Show of hands or quick show on a whiteboard/slide.)"),
      numbered("Students who did home installation ahead of time (per the Engine & Platform Guide) verify their install works and open a new project."),
      numbered("Students who haven't installed yet: start with GDevelop or Scratch in-browser immediately (zero install, zero wait), then queue a larger download (Unity/Godot/GameMaker/001 Game Creator) to run in the background if bandwidth allows."),
      numbered("Instructor circulates by “station” (cluster students by chosen engine) to troubleshoot installs and answer setup questions."),
      numbered("Exit check: each student should have their engine's editor open, ideally with a blank or template project loaded."),

      h1("Closing (8 minutes): This Semester's Game"),
      p("This is the payoff of the tease from the welcome segment - keep it quick and energetic, not a full design walkthrough. The rest of the design (levels, features) is intentionally revealed a piece at a time as the class actually builds it, not all today."),
      h3("1. High Concept reveal (~3 min)"),
      p("Tell students, in your own words: this semester, the whole class is building the same game together - a top-down arcade shooter in the tradition of classics like Space Invaders and Galaga. Waves of enemies descend from above; you move, dodge, and shoot your way through before they reach you. Explain why: it needs almost no art to start, but it grows into something genuinely impressive by the end of the semester - and because nobody has to invent the game from scratch, everyone gets to focus on getting really good at their engine and their code. Personalization still happens - your own art, sound, and how far into Level 3 you push - it's just not \"what game do I make,\" the way Semester 2's fully open-ended Expo game will be."),
      h3("2. Hand out the Lesson 1 GDD (~2 min)"),
      p("Distribute Lesson_01_GDD.docx (the sneak-peek Game Design Document) - one per student. It's theirs to keep. Point out that it's a preview, not the real thing: the actual Game Design Document starts in Lesson 2, and it grows with them all year. No need to read it aloud - it repeats what you just said, plus the roadmap and the Five Disciplines, for them to revisit at home."),
      h3("3. Codename vote (~3 min)"),
      p("Real studios often work under a codename before a project's final title is locked in - so this game gets one too. Show your 3-4 prepared candidate codenames (see Instructor Prep) - or open the floor for spontaneous suggestions if you'd rather keep it looser - and take a quick show-of-hands vote. Announce the winner on the spot. It's just a codename, not final: like a real studio, the class can revisit and rename it later once the game actually starts taking shape."),

      h1("Homework / Before Next Class"),
      bullet("Finish installing your chosen engine if it wasn't complete by the end of lab"),
      bullet("Skim the Student Resource List for your engine and bookmark one beginner tutorial"),
      bullet("Optional: explore your engine's interface for 15-20 minutes just to get comfortable clicking around"),

      h1("Instructor Prep Checklist"),
      bullet("Pre-install all 8 engines on the demo/backup laptop so any engine can be demonstrated on request (001 Game Creator requires the Windows demo laptop specifically)"),
      bullet("Pre-build the Level 1 Key Press Reactor mini-project in all 8 engines ahead of time (Level 1 steps/code are in each engine's Lesson 2 Teacher's Key) so the Live Engine Tour is a quick press-Play-and-test, not live coding under time pressure"),
      bullet("Preview and queue NetSmartz video clip(s); confirm they play without needing a live internet connection if classroom wifi is unreliable"),
      bullet("Send the Engine & Platform Guide and the Engine Showcase (illustrated preview of all 8 engines) to families at least 1-2 weeks before this lesson so home installation - and engine choice - can start early"),
      bullet("Prepare offline installers on a USB drive as backup for Unity, Godot, GameMaker, and 001 Game Creator in case classroom wifi can't handle simultaneous downloads"),
      bullet("Print or digitally share a sign-in sheet and, if required by the co-op, any consent/media-release forms"),
      bullet("Have the Course Syllabus available for questions during the welcome segment"),
      bullet("Print or email a copy of Lesson_01_GDD.docx (sneak peek) for every student, to hand out during the Closing segment"),
      bullet("Come up with 3-4 candidate codenames for the Semester 1 game ahead of time for the codename vote (or plan to take spontaneous suggestions instead, if you'd rather keep it open)"),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Lesson_01_Plan.docx", buffer);
  console.log("done");
});
