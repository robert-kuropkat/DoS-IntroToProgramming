const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign,
  PageBreak, ImageRun, TabStopType, TabStopPosition
} = require("docx");

const NAVY = "16305C";
const NAVY2 = "1C4A86";
const GOLD = "C99A2E";
const GOLD_TEXT = "8A6D1E";

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
    new TextRun({ text: "  |  Semester 1 Capstone - Master GDD (Instructor Reference)", size: 16, color: "888888" }),
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

const controlsHeaders = ["Input", "Action"];
const controlsWidths = [3360, 6000];
const controlsRows = [
  ["Left / Right Arrow", "Move the player left or right along the bottom of the play area"],
  ["Spacebar (or Up Arrow, if spacebar conflicts with an engine's own shortcuts)", "Shoot a projectile upward"],
];

const mapHeaders = ["Lesson", "Feature Introduced", "Capstone Level"];
const mapWidths = [1200, 6360, 1800];
const mapRows = [
  ["2", "Player movement, left/right (this is Lesson 2's own Key Press Reactor goal - no separate capstone work needed, it's the same skill)", "Level 1"],
  ["3", "Shooting: spawn a projectile, move it upward, remove it off-screen", "Level 1"],
  ["4", "A single enemy on screen + collision detection (projectile destroys enemy)", "Level 1"],
  ["5", "Multiple enemies (a row/list of enemies), loops to spawn and manage them", "Level 1"],
  ["6", "Win condition: all enemies destroyed. Level 1 MVP milestone complete.", "Level 1"],
  ["7", "Enemies advance toward the player (lane-based descent - see Design Notes)", "Level 2"],
  ["8", "Lose condition: enemy reaches the player / player is hit, plus a lives system", "Level 2"],
  ["9", "Score tracking and a basic on-screen UI (score, lives)", "Level 2"],
  ["10", "Game-over and restart flow", "Level 2"],
  ["11", "Multiple waves. Level 2 Complete milestone reached.", "Level 2"],
  ["12", "Sound effects: shoot, explosion, game over, wave clear", "Level 3"],
  ["13", "Basic animation: explosion effect, hit flash/reaction", "Level 3"],
  ["14", "Difficulty scaling across waves; introduce optional stretch features", "Level 3"],
  ["15", "Polish pass and Semester 1 Capstone presentation. Level 3 Refined milestone reached.", "Level 3"],
];

const crestData = fs.readFileSync("vgc_crest.png");
const crestWidthIn = 1.5, crestHeightIn = crestWidthIn * (1040 / 840);

const children = [
  new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 200 },
    children: [new ImageRun({ data: crestData, transformation: { width: crestWidthIn * 96, height: crestHeightIn * 96 }, type: "png" })],
  }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Semester 1 Capstone", bold: true, size: 40, color: NAVY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [new TextRun({ text: "Master Game Design Document", bold: true, size: 26, color: GOLD_TEXT })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Instructor Reference & Target Spec - not distributed to students", size: 22, color: NAVY2 })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: GOLD_TEXT, characterSpacing: 20 })] }),

  ...fillLine("Working Title (finalize with students - naming it is itself a fun early production step):"),

  h1("How to Use This Document"),
  p("This is the instructor-facing master design for the game every student builds in Semester 1 - it's the answer key, not a student handout. It complements Game_Design_Document_Template.docx (the student's own living notebook) rather than replacing it."),
  bullet("This mirrors real pre-production: a high concept gets pitched, then locked into a full design, then sliced into a production roadmap - here, the roadmap is lesson-by-lesson instead of sprint-by-sprint."),
  bullet("Each lesson's own GDD \"Goals\" table (starting where a feature is introduced) is a slice of the Feature-to-Lesson Map below. If the design changes here, re-slice any lessons it affects."),
  bullet("Students only see the High Concept up front (Lesson 1's GDD) - the rest of this spec is intentionally revealed to them incrementally, feature by feature, the same way a real dev team learns a design through production rather than all at once."),

  h1("High Concept"),
  labelBullet("Genre", "Top-down arcade shooter, in the tradition of Space Invaders and Galaga."),
  labelBullet("One-Sentence Pitch", "Waves of enemies descend from above - move, dodge, and shoot your way through before they reach you."),
  labelBullet("Inspiration", "Classic arcade shooters (Space Invaders, Galaga) - chosen for minimal art/animation requirements and a core loop that scales cleanly from a one-lesson MVP to a full semester of features."),
  labelBullet("Tone", "Light, arcade-y, all-ages - no violence beyond abstract \"invader\" sprites, no lose-state punishment beyond \"try again\"."),

  h1("Core Gameplay"),
  labelBullet("Core Mechanic", "Move along the bottom of the screen and shoot upward to destroy enemies before they reach you."),
  h3("Controls"),
  table(controlsHeaders, controlsWidths, controlsRows),
  labelBullet("Win Condition", "Clear all enemies in the current wave (Level 1: a single wave. Level 2+: multiple waves.)"),
  labelBullet("Lose Condition", "An enemy reaches the player's row, or the player runs out of lives (Level 2+)."),

  h1("World & Characters"),
  labelBullet("Player", "A single ship or character confined to the bottom of the screen. No name/backstory needed - keep it abstract so every engine's default or placeholder assets work fine."),
  labelBullet("Enemies", "Simple, uniform \"invader\" enemies at Level 1. Level 3 may introduce a second enemy type or bonus enemy as an optional stretch."),

  new Paragraph({ children: [new PageBreak()] }),

  h1("Level 1 / 2 / 3 Feature Breakdown"),
  p("Same three levels as every other lesson in this course - Level 1 is a real, complete finish, not a rough draft. Levels 2 and 3 build on it without changing what's already there."),

  h2("Level 1 - MVP"),
  bullet("Move left/right via directional input"),
  bullet("Shoot a single projectile upward on a key press"),
  bullet("One fixed row of enemies present at start"),
  bullet("Projectile destroys an enemy on collision"),
  bullet("Win condition: all enemies destroyed"),

  h2("Level 2 - Complete"),
  bullet("Multiple rows / a full wave of enemies"),
  bullet("Enemies advance toward the player (lane-based descent - see Design Notes below)"),
  bullet("Lose condition: an enemy reaches the player, or the player is hit"),
  bullet("Player has limited lives (suggested: 3)"),
  bullet("Score tracking plus a basic UI (score, lives)"),
  bullet("Game-over / restart flow"),

  h2("Level 3 - Refined"),
  p("Core Level 3 features (expected of everyone reaching this level):", { italics: true, color: "555555" }),
  bullet("Multiple waves with increasing difficulty (more enemies, faster movement, or both)"),
  bullet("Sound effects: shoot, explosion, game over, wave clear"),
  bullet("Basic animation: explosion effect, hit flash/reaction"),
  p("Optional stretch features (student's choice - this is where personalization lives):", { italics: true, color: "555555" }),
  bullet("Destructible shields/barriers"),
  bullet("A power-up (rapid fire, spread shot, temporary shield)"),
  bullet("A bonus enemy worth extra points"),
  bullet("Classic formation-marching movement (side-to-side, dropping a row at each edge) instead of lane-based descent, for students ready to work with lists/arrays of enemies"),
  bullet("Title screen and/or game-over screen"),

  new Paragraph({ children: [new PageBreak()] }),

  h1("Feature-to-Lesson Map"),
  p("This is the slice-and-dice backbone: which capstone feature gets introduced in which lesson, matching the pacing already set by the 30-lesson milestone table in the Course Syllabus and GDD template. First draft - review and adjust before it's locked in, since every lesson built from here on inherits this pacing."),
  table(mapHeaders, mapWidths, mapRows),

  h1("Design Notes / Open Decisions"),
  labelBullet("Direction-color feedback: sprite swap vs. color tint", "Decided (Lesson 2): color change will be implemented by sprite swapping - a separate pre-made colored sprite/costume per direction - rather than tinting one sprite at runtime with each engine's quick-color property (image_blend, modulate, spriteRenderer.color, Tint color, and similar). Tint was the original plan, but it only produces a clean, true color on white/grayscale source art; on our actual colored placeholder art it multiplies into a muddy result instead, discovered while live-testing the GameMaker path. Sprite-swap was chosen over the alternative fix (adding white art variants and keeping tint) for a few reasons: it's the simpler concept for this class's age/skill range - \"show a different picture\" needs no explanation of multiply-blend color math - it needed zero new assets (the placeholder pack already has one color per direction plus a spare for idle), and it works the same way in every engine with no exceptions. This is also a deliberate, admittedly somewhat arbitrary \"faster vs. more sophisticated\" call, the kind real studios make under time and team-skill constraints: tinting is more flexible long-term (any number of colors with no new art, smooth transitions possible) and stays available to any student ready for it as a Super-Stretch idea, once they understand why it needs white source art to look clean. TIC-80 is the deliberate exception, left on its original color-tint approach (poke4() palette remapping) rather than switched to sprite-swap - unlike every other engine here, TIC-80's palette-based recolor is exact regardless of the source art's color, so it never had the muddy-blend problem sprite-swap exists to avoid; moving it to sprite-swap would trade away a technique that's already simpler and more idiomatic there for one it doesn't need."),
  labelBullet("Enemy movement: lane-based descent vs. formation marching", "Recommending lane-based descent (each enemy independently moves down its own lane) as the Level 1/2 default - it needs no shared group state and stays uniformly simple to build across all 8 engines, including block/visual tools and younger students. True Space-Invaders-style formation marching (shared movement, edge-bounce, row-drop, speed-up as enemies thin out) is offered instead as a Level 3 stretch feature, once students are comfortable with lists/arrays. Flagging for confirmation before this is locked into the Feature-to-Lesson Map above."),
  labelBullet("Working title", "Left blank on the cover - naming the game could even be a short class activity or vote once the High Concept is pitched in Lesson 1, mirroring how studios sometimes finalize a title well after a project is greenlit."),
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
  fs.writeFileSync("Semester_01_Capstone_GDD.docx", buffer);
  console.log("done");
});
