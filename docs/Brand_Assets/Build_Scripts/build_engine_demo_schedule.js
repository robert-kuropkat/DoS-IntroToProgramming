const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign,
} = require("docx");
const brand = require("./brand.js");

const CONTENT_WIDTH = 9360;
const border = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function bullet(text, level = 0) {
  return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] });
}
function headerCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: brand.NAVY2, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 19 })] })],
  });
}
function bodyCell(text, width, opts = {}) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, size: 19, bold: !!opts.bold, italics: !!opts.italics, color: opts.color })] })],
  });
}
function scheduleTable(rows) {
  const widths = [1150, 2900, 5310];
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, cantSplit: true, children: ["Lesson", "Featured Engine (Lecture Demo)", "Notes"].map((t, i) => headerCell(t, widths[i])) }),
      ...rows.map((row, ri) => new TableRow({
        cantSplit: true,
        children: [
          bodyCell(row.lesson, widths[0], { fill: ri % 2 ? "F3F6FA" : "FFFFFF", bold: true }),
          bodyCell(row.engine, widths[1], { fill: ri % 2 ? "F3F6FA" : "FFFFFF", bold: true, italics: row.noDemo, color: row.noDemo ? "666666" : undefined }),
          bodyCell(row.notes, widths[2], { fill: ri % 2 ? "F3F6FA" : "FFFFFF", italics: row.noDemo, color: row.noDemo ? "666666" : undefined }),
        ],
      })),
    ],
  });
}
function spacer(h) { return new Paragraph({ spacing: { after: h || 200 }, children: [] }); }

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ],
};

// Source of truth: each lesson's Plan document's "Featured Engine (lecture demo)"
// line (per Standards Rule 10). Verified directly against build_lessonNN_plan*.js
// for Lessons 2-14; Lessons 1 and 15 have no line by design (no demo either week).
const rows = [
  { lesson: "1", engine: "— (no demo)", notes: "Whole-class course-foundations day; no hands-on build exercise yet, so there's nothing to demo.", noDemo: true },
  { lesson: "2", engine: "Scratch", notes: "" },
  { lesson: "3", engine: "Godot Engine", notes: "Predates Rule 10 (written before the rotation was formalized), but happens to already be one of the pool's allowed choices." },
  { lesson: "4", engine: "Godot Engine", notes: "Predates Rule 10, same as Lesson 3." },
  { lesson: "5", engine: "GameMaker (GML Visual)", notes: "Rotation begins here. Primary demo is Visual, with a brief GML script contrast (see this lesson's Sidebars)." },
  { lesson: "6", engine: "Scratch", notes: "" },
  { lesson: "7", engine: "Godot Engine", notes: "Contrasted briefly with GameMaker for the bug-hunting discussion." },
  { lesson: "8", engine: "Unity (C#)", notes: "" },
  { lesson: "9", engine: "GameMaker (GML script)", notes: "Callback to Lesson 5's GameMaker Visual demo – same underlying project, script mode this time." },
  { lesson: "10", engine: "Godot Engine", notes: "" },
  { lesson: "11", engine: "Unity (C#)", notes: "" },
  { lesson: "12", engine: "GameMaker (GML Visual)", notes: "Second Visual-mode rotation slot; callback to Lesson 5." },
  { lesson: "13", engine: "Scratch", notes: "Contrasted with GameMaker's Alarm-based workaround." },
  { lesson: "14", engine: "GameMaker (GML script)", notes: "Second script-mode rotation slot; callback to Lesson 9." },
  { lesson: "15", engine: "— (no demo)", notes: "Workshop / showcase day by design, not a lecture – deliberately demo-free.", noDemo: true },
];

const children = [
  ...brand.titleBlock("Engine Demo Schedule", "Which engine is the lecture demo, week by week"),

  p("This is a quick lookup, not a policy document: it lists which engine each lesson's Teacher's Script uses for the live lecture demo, straight from that lesson's Plan document's “Featured Engine (lecture demo)” line – the single authoritative record per Standards Rule 10. If a lesson's demo engine ever changes, update its Plan first, then update this table to match; don't edit this table as the source of truth."),

  h1("Semester 1 (Lessons 1–15)"),
  scheduleTable(rows),
  spacer(200),

  h1("How the Rotation Works"),
  bullet("The current pool draws from only 5 choices – Scratch, GameMaker (GML script), GameMaker (GML Visual), Unity (C#), and Godot Engine – an intentional, temporary narrowing to the engines and coding styles actually in active use by enrolled students right now."),
  bullet("Lessons 5 through 14 use each of those five exactly twice, with no two consecutive lessons repeating the same choice. Where GameMaker appears twice, the second occurrence is an explicit callback to the first, continuing the same underlying demo project rather than starting over."),
  bullet("Lessons 1 and 15 are demo-free by design – Lesson 1 is whole-class introduction with no build exercise yet, and Lesson 15 is a workshop/showcase day, not a lecture."),
  bullet("Lessons 3 and 4 predate Rule 10 and were never assigned through the rotation process – they happen to already use Godot, which is an allowed pool choice, so no correction was needed. They're included above for completeness."),
  bullet("This 5-choice pool is not the course's permanent design. The full rotation is meant to eventually cover all 8 engines and all 3 coding styles – GDevelop, 001 Game Creator, TIC-80, and Unreal (Blueprint and C++) will join the pool once students are actively using them. Expanding the pool is a tracked Book Notes & To-Do item, not yet started."),
  bullet("Lessons 16–30 (Semester 2) haven't been authored yet. When they are, each one needs its own Featured Engine line in its Plan document per Rule 10, drawn from whatever the pool looks like at that time – don't assume a Semester 1 pattern carries forward automatically."),

  h1("Revision Note"),
  p("First version, built as a quick companion to Data_Dictionary.xlsx – same underlying motivation (a single place to look up a cross-lesson fact instead of re-deriving it from each lesson's Plan document). Update this table any time a lesson's Featured Engine changes, and add a line here when Lessons 16–30 get their own assignments."),
];

const doc = new Document({
  numbering: numberingConfig,
  styles: brand.stylesBlock,
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: brand.pageHeader("Engine Demo Schedule") },
    footers: { default: brand.pageFooter() },
    children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Engine_Demo_Schedule.docx", buffer);
  console.log("done");
});
