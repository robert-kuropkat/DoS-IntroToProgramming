const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign,
  PageBreak, ImageRun, TabStopType, TabStopPosition
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
function table(headers, widths, rows) {
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

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
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
    new TextRun({ text: "  |  Game Design Document - Lesson 1 Sneak Peek", size: 16, color: "888888" }),
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
// Same real content as the full GDD template - shown here as a preview,
// not a fill-in-the-blank, so students get a concrete feel for the year.
// =========================================================================
const levelHeaders = ["Level", "Think of it as...", "What it means"];
const levelWidths = [1500, 2360, 5500];
const levelRows = [
  ["Level 1 - Minimal", "MVP", "The smallest working version - enough to understand the idea and move on. A complete, legitimate finish."],
  ["Level 2 - Complete", "Internal Alpha/Beta", "All expected features present and working end-to-end. Art/polish/balance may still be missing."],
  ["Level 3 - Refined", "Ready to Publish", "Feature-complete, tested, and polished enough to comfortably share outside of class."],
];

const milestoneHeaders = ["Lessons", "Class Focus", "Class Milestone"];
const milestoneWidths = [1300, 5560, 2500];
const milestoneRows = [
  ["1", "Orientation: internet safety, game dev overview, platform selection", "Engine installed & first project opened"],
  ["2-6", "Programming foundations: variables, events, loops, conditionals, input", "First moving, controllable character"],
  ["7-11", "Core mechanics: collision, scoring, win/lose states, simple UI", "Small playable prototype"],
  ["12-15", "Game feel & polish: audio, basic animation, level structure", "Semester 1 capstone: complete Mini-Game"],
  ["16-20", "Intermediate systems: state machines, enemy/AI behavior, physics", "Prototype with opponents/obstacles"],
  ["21-25", "Expanding scope: menus, save data, multiple levels, juice/polish", "Feature-complete draft of capstone game"],
  ["26-28", "Playtesting, bug-fixing, presentation prep", "Playtested, stable build"],
  ["29-30", "Finalize, package, and export shareable copies", "Shareable build ready for CoOp Expo"],
];

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
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [new TextRun({ text: "Sneak Peek", bold: true, size: 26, color: GOLD_TEXT })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "A first look at the game we're all building this semester", size: 22, color: NAVY2 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: GOLD_TEXT, characterSpacing: 20 })] }),

  ...fillLine("Name (optional - this is just a preview to keep):"),

  h1("This Semester, We're All Making..."),
  p("A top-down arcade shooter, in the tradition of classics like Space Invaders and Galaga: waves of enemies descend from above - move, dodge, and shoot your way through before they reach you.", { bold: true }),
  p("This is exactly how a real game studio kicks off a project - a short, exciting \"high concept\" pitch, before anyone writes a full design or a line of code. We picked this idea on purpose: it needs almost no art or animation to start, but it has room to grow into something genuinely impressive by the end of the semester."),
  p("Here's the twist: everyone in class builds this same game. That's not a limit on your creativity - it's the opposite. Because nobody has to invent the game from scratch, you get to spend the semester getting really good at your engine and your code, while still making the game feel like yours through your own art, sound, and extra touches. Think of it like a gymnast's compulsory routine - same core moves for everyone, but every performance still looks different. Semester 2 is where you get a fully free-form, entirely-your-own game to design."),
  p("Full details - exactly how it works, level by level - get revealed a piece at a time as we build it together, lesson by lesson. Today, this is all you need to know: what it is, and that you're going to build it.", { italics: true, color: "555555" }),

  new Paragraph({ children: [new PageBreak()] }),

  h1("How This Class Is Leveled"),
  p("Every project in this class - starting with the very first one next week - gets judged against three levels, not just \"done\" or \"not done\":", { italics: true, color: "555555" }),
  table(levelHeaders, levelWidths, levelRows),
  p("Level 1 is always a real, complete finish - not a rough draft. Levels 2 and 3 are for anyone with time and curiosity to push further.", { italics: true, color: "555555" }),

  new Paragraph({ children: [new PageBreak()] }),

  h1("Where This Is Going"),
  p("This is the actual 30-lesson roadmap for the whole class, start to finish. It's a lot to take in on day one - the point isn't to memorize it, just to see the shape of the year and get excited about where it leads."),
  table(milestoneHeaders, milestoneWidths, milestoneRows),

  h1("The Five Disciplines of Game Development"),
  p("Every game - including the one you'll be building - touches all five of these, even on a small student project. Programming is our main focus in class, but the other four matter too, and your real Game Design Document will have space for notes on each."),
  ...disciplines.flatMap(d => [ labelBullet(d.name, d.desc) ]),

  new Paragraph({ children: [new PageBreak()] }),

  h1("Coming Up in Lesson 2"),
  p("Next class, you'll pick the engine or platform you want to build in, get your own copy of the real Game Design Document, and write your very first working program: a little character that reacts to a key press. That character is where our game starts."),
  p("Between now and then, it's worth thinking about:", { bold: true }),
  bullet("What should our ship or character look like? (Sketch it if you want - default and placeholder art are always fine too.)"),
  bullet("What would you name this game? We'll pick an official title together once we've actually started building it - real studios do this too."),
  bullet("Which engine sounds most interesting to you for building it? (Ask your instructor if you're not sure - there's no wrong choice.)"),
  p(" "),
  p("See you next class - let's build it!", { bold: true, color: NAVY2 }),
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
  fs.writeFileSync("Lesson_01_GDD.docx", buffer);
  console.log("done");
});
