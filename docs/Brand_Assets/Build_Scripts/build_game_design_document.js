const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign,
  PageBreak, TableOfContents, ImageRun, TabStopType, TabStopPosition
} = require("docx");

// Brand palette - matches the Lesson 1 slide deck (generate_deck.js) exactly.
const NAVY = "16305C";
const NAVY2 = "1C4A86";
const GOLD = "C99A2E";
const GOLD_TEXT = "8A6D1E"; // darker gold, legible as text-on-white

const CONTENT_WIDTH = 9360;
const border = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const fillBorder = { style: BorderStyle.SINGLE, size: 4, color: "AAAAAA" };

function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function checkbox(text) { return new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 80 }, children: [new TextRun(text)] }); }
function labelBullet(label, text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)] });
}
function headerCell(text, width) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: NAVY2, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 18 })] })] });
}
function bodyCell(children, width, opts = {}) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.TOP,
    children: Array.isArray(children) ? children : [new Paragraph({ children: [new TextRun({ text: children, size: 18 })] })] });
}
function table(headers, widths, rows, opts = {}) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((t, i) => headerCell(t, widths[i])) }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((t, i) => bodyCell(t, widths[i], { fill: ri % 2 ? "F3F6FA" : "FFFFFF" }))
      }))
    ]
  });
}

// A single fill-in line: bold label, then a blank ruled line for writing/typing on.
function fillLine(label) {
  return [
    new Paragraph({ spacing: { before: 140, after: 40 }, children: [new TextRun({ text: label, bold: true, size: 20, color: NAVY2 })] }),
    new Paragraph({ spacing: { after: 200 }, border: { bottom: fillBorder }, children: [new TextRun({ text: " ", size: 22 })] }),
  ];
}
// A labeled block with an optional italic hint and N blank ruled lines for longer notes.
function fillBox(label, hint, lines = 3) {
  const children = [new Paragraph({ spacing: { before: 140, after: 40 }, children: [new TextRun({ text: label, bold: true, size: 20, color: "2E5C8A" })] })];
  if (hint) children.push(new Paragraph({ spacing: { after: 80 }, children: [new TextRun({ text: hint, italics: true, size: 16, color: "888888" })] }));
  for (let i = 0; i < lines; i++) {
    children.push(new Paragraph({ spacing: { after: 220 }, border: { bottom: fillBorder }, children: [new TextRun({ text: " ", size: 22 })] }));
  }
  return children;
}
function studentNotesHeader() {
  return new Paragraph({ spacing: { before: 60, after: 40 }, children: [new TextRun({ text: "YOUR NOTES", bold: true, size: 16, color: "B00020" })] });
}

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "checks", levels: [{ level: 0, format: LevelFormat.BULLET, text: "☐", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]
};

const stylesBlock = {
  default: { document: { run: { font: "Arial", size: 22 } } },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 30, bold: true, font: "Arial", color: NAVY },
      paragraph: { spacing: { before: 340, after: 160 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 24, bold: true, font: "Arial", color: NAVY2 },
      paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 22, bold: true, font: "Arial", color: "444444" },
      paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
  ]
};

function pageHeader() {
  return new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
    new TextRun({ text: "VIDEO GAME CREATION", size: 16, color: GOLD_TEXT, bold: true, characterSpacing: 20 }),
    new TextRun({ text: "  |  Game Design Document", size: 16, color: "888888" }),
  ] })] });
}
function pageFooter() {
  return new Footer({ children: [new Paragraph({
    tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH }],
    children: [
      new TextRun({ text: "SLHE Game Programming", size: 16, color: "888888" }),
      new TextRun({ text: "\tPage ", size: 16, color: "888888" }),
      new TextRun({ children: [PageNumber.CURRENT], size: 16, color: "888888" }),
    ]
  })] });
}

// =========================================================================
// Milestone schedule (matches Course Syllabus - Meeting Schedule / Scope table)
// =========================================================================
const milestoneHeaders = ["Lessons", "Class Focus", "Class Milestone", "My Status"];
const milestoneWidths = [1100, 4260, 2500, 1500];
const milestoneRows = [
  ["1", "Orientation: internet safety, game dev overview, platform selection", "Engine installed & first project opened", " "],
  ["2-6", "Programming foundations: variables, events, loops, conditionals, input", "First moving, controllable character", " "],
  ["7-11", "Core mechanics: collision, scoring, win/lose states, simple UI", "Small playable prototype", " "],
  ["12-15", "Game feel & polish: audio, basic animation, level structure", "Semester 1 capstone: complete Mini-Game", " "],
  ["16-20", "Intermediate systems: state machines, enemy/AI behavior, physics", "Prototype with opponents/obstacles", " "],
  ["21-25", "Expanding scope: menus, save data, multiple levels, juice/polish", "Feature-complete draft of capstone game", " "],
  ["26-28", "Playtesting, bug-fixing, presentation prep", "Playtested, stable build", " "],
  ["29-30", "Finalize, package, and export shareable copies", "Shareable build ready for CoOp Expo", " "],
];

const controlsHeaders = ["Input", "Action", "Still accurate?"];
const controlsWidths = [2200, 5160, 2000];
const controlsRows = [
  ["Right Arrow", "Move right (starting point from your Lesson 2 Key Press Reactor)", " "],
  ["Left Arrow", "Move left", " "],
  ["Up Arrow", "Move up", " "],
  ["Down Arrow", "Move down", " "],
  [" ", " ", " "],
  [" ", " ", " "],
  [" ", " ", " "],
];

const levelHeaders = ["Level", "Think of it as...", "What it means"];
const levelWidths = [1500, 2360, 5500];
const levelRows = [
  ["Level 1 - Minimal", "MVP", "The smallest working version - enough to understand the idea and move on. A complete, legitimate finish."],
  ["Level 2 - Complete", "Internal Alpha/Beta", "All expected features present and working end-to-end. Art/polish/balance may still be missing."],
  ["Level 3 - Refined", "Ready to Publish", "Feature-complete, tested, and polished enough to comfortably share outside of class."],
];

const journalHeaders = ["Date", "Lesson #", "What changed and why"];
const journalWidths = [1500, 1360, 6500];
const journalRows = [[" ", " ", " "], [" ", " ", " "], [" ", " ", " "], [" ", " ", " "], [" ", " ", " "], [" ", " ", " "], [" ", " ", " "], [" ", " ", " "]];

const disciplines = [
  { name: "Design", desc: "Deciding what the game is and how it plays - the rules, the goals, the feel." },
  { name: "Programming", desc: "Making the rules and mechanics actually work in code (our main focus this year)." },
  { name: "Visual Art", desc: "Characters, environments, and UI - what the player sees." },
  { name: "Audio", desc: "Music and sound effects - what the player hears." },
  { name: "Production", desc: "Planning, scoping, and finishing a project - relevant all year, critical before the Expo." },
];

const crestData = fs.readFileSync("vgc_crest.png");
const crestWidthIn = 1.5, crestHeightIn = crestWidthIn * (1040 / 840);

const children = [
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 200 },
    children: [new ImageRun({ data: crestData, transformation: { width: crestWidthIn * 96, height: crestHeightIn * 96 }, type: "png" })],
  }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "My Game Design Document", bold: true, size: 40, color: NAVY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "SLHE Game Programming - a living document you'll update all year", size: 22, color: NAVY2 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: GOLD_TEXT, characterSpacing: 20 })] }),

  ...fillLine("Student Name:"),
  ...fillLine("Game Title:"),
  ...fillLine("Engine / Platform:"),
  ...fillLine("Started (Lesson #/Date):"),

  new Paragraph({ children: [new PageBreak()] }),

  h1("How to Use This Document"),
  p("This is not a one-time assignment - it's a notebook for your game that grows alongside it. You'll start it in Lesson 2 with almost everything blank, and you'll come back to add, cross out, and rewrite sections as your game changes all year. That's normal and expected - professional game designers rewrite their design docs constantly too."),
  bullet("There are no wrong answers this early. If you don't know yet, write your best guess - you can change it later."),
  bullet("Sections marked \"YOUR NOTES\" are blank on purpose. Everything else is either a fill-in-the-blank or something already planned for the whole class that's included here for reference."),
  bullet("Keep this file (or a printed copy) somewhere you'll find it again - you'll be asked to revisit it at several points across both semesters, especially before the Semester 1 capstone and the CoOp Expo."),

  h1("The Big Picture"),
  ...fillBox("Genre", "Examples: platformer, top-down adventure, puzzle, shooter, racing, simulation...", 1),
  ...fillBox("One-Sentence Pitch", "Example: \"A robot has to reach the top of a crumbling tower before it collapses.\"", 2),
  ...fillBox("Games it's inspired by", "What games (or movies, books, etc.) gave you this idea?", 2),

  h1("Core Gameplay"),
  p("Lesson 2's Key Press Reactor is this question in its simplest possible form: what happens when you press a key? Everything you add later builds on that same idea.", { italics: true, color: "555555" }),
  ...fillBox("Core Mechanic", "What's the ONE thing the player does over and over? Everything else in your game supports this one action.", 2),

  h3("Controls"),
  p("Starts pre-filled with your Lesson 2 Key Press Reactor mapping - update the \"Still accurate?\" column (or add rows) as your game grows.", { italics: true, color: "555555" }),
  table(controlsHeaders, controlsWidths, controlsRows),

  ...fillBox("Win Condition", "How does the player succeed?", 2),
  ...fillBox("Lose Condition", "How does the player fail? (Not every game needs one!)", 2),

  new Paragraph({ children: [new PageBreak()] }),

  h1("World & Characters"),
  ...fillBox("Player Character", "Who / what is the player controlling?", 3),
  ...fillBox("Setting / World", "Where does this take place?", 3),
  ...fillBox("Other Characters, Enemies, or NPCs", "", 3),

  h1("Levels / Structure"),
  p("Reminder - how \"Level 1 / 2 / 3\" works in this class:", { bold: true }),
  table(levelHeaders, levelWidths, levelRows),
  p(" "),
  table(["Level/Area", "What happens there", "Status"], [3200, 4560, 1600],
    [[" ", " ", " "], [" ", " ", " "], [" ", " ", " "], [" ", " ", " "]]),

  h1("Visual Art Plan"),
  p("Placeholder art is completely fine while you're prototyping - the lesson Assets folders always include ready-to-use placeholder sprites so you never have to stop and hunt for art mid-lesson. Use this section to describe what you eventually want your game to look like.", { italics: true, color: "555555" }),
  ...fillBox("Art Style", "Examples: pixel art, flat cartoon, simple geometric shapes, hand-drawn...", 2),
  ...fillBox("Color Palette / Mood", "", 2),
  h3("Placeholder assets currently in use (check any that apply)"),
  checkbox("player_circle_blue.png / player_circle_green.png / player_circle_orange.png"),
  checkbox("player_arrow_white.png"),
  checkbox("player_walk_spritesheet.png"),
  checkbox("My own art (describe below)"),
  ...fillBox("Notes on your own art assets", "", 2),

  h1("Audio Plan"),
  p("sfx_blip.wav is the placeholder sound effect provided in early lessons - swap it out here with what you actually want once you know.", { italics: true, color: "555555" }),
  ...fillBox("Music Style / Mood", "", 2),
  ...fillBox("Sound Effects Needed", "List the moments in your game that should make a sound.", 3),

  new Paragraph({ children: [new PageBreak()] }),

  h1("Scope & Milestones"),
  p("This is the class's actual 30-lesson schedule (see the Course Syllabus for full detail) with a column for your own status. Use it to check whether your game is on pace with the rest of the class, not as a strict deadline - talk to your instructor if you're falling behind or racing ahead."),
  table(milestoneHeaders, milestoneWidths, milestoneRows),

  h2("Semester 1 Capstone (Lessons 12-15)"),
  p("By the end of Lesson 15, your game should be a complete Mini-Game: it has a clear core mechanic, art and audio (even simple/placeholder), and reaches a real win or lose state. This is the same project you started in Lesson 2 - it's been growing the whole time."),
  ...fillBox("My plan for reaching the Semester 1 capstone", "", 3),

  h2("Semester 2 / CoOp Expo Capstone (Lessons 26-30)"),
  p("The final weeks of Semester 2 shift from new material to finishing, playtesting, and packaging your game so it can be demonstrated at the co-op's Expo and shared with friends and classmates. You'll export a runnable build (or shareable link, for browser-based engines) and prepare a short demo/pitch, with structured playtesting feedback from classmates before the Expo."),
  ...fillBox("My plan for the CoOp Expo", "", 3),

  new Paragraph({ children: [new PageBreak()] }),

  h1("The Five Disciplines of Game Development"),
  p("Every game (including yours) touches all five of these, even on a small student project. Programming is our main focus in class, but the other four matter too - use the space under each to jot ideas as they come to you."),
  ...disciplines.flatMap(d => [
    h3(d.name),
    p(d.desc, { italics: true, color: "555555" }),
    studentNotesHeader(),
    new Paragraph({ spacing: { after: 200 }, border: { bottom: fillBorder }, children: [new TextRun({ text: " ", size: 22 })] }),
  ]),

  h1("Design Journal / Change Log"),
  p("A quick running log of significant changes to your game and why you made them - useful for remembering your own reasoning later, and required reading for yourself before the Semester 1 and Expo capstones."),
  table(journalHeaders, journalWidths, journalRows),
];

const doc = new Document({
  numbering: numberingConfig,
  styles: stylesBlock,
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: pageHeader() },
    footers: { default: pageFooter() },
    children,
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Game_Design_Document_Template.docx", buffer);
  console.log("done");
});
