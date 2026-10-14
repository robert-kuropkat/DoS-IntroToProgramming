const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  HeadingLevel, LevelFormat, PageBreak,
} = require("docx");
const brand = require("./brand.js");

const CONTENT_WIDTH = brand.CONTENT_WIDTH;
const border = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function note(text) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, italics: true, color: "B00020" })] }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function numbered(text, level = 0) { return new Paragraph({ numbering: { reference: "numbers", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }

function headerCell(text, width) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: brand.NAVY2, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 18 })] })] });
}
function bodyCell(text, width, opts = {}) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.TOP,
    children: [new Paragraph({ children: [new TextRun({ text, size: 18 })] })] });
}

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ],
};

// Extend brand.stylesBlock with a "Title" style for cover pages, and a plain
// "Note" style for the light gray/italic callout paragraphs used throughout
// the course's documents.
const stylesBlock = {
  default: brand.stylesBlock.default,
  paragraphStyles: [
    ...brand.stylesBlock.paragraphStyles,
    { id: "TitleBig", name: "Title Big", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 40, bold: true, font: brand.FONT, color: brand.NAVY },
      paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 60 } } },
    { id: "Subtitle", name: "Subtitle", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 22, font: brand.FONT, color: brand.NAVY2 },
      paragraph: { alignment: AlignmentType.CENTER, spacing: { after: 300 } } },
  ],
};

const children = [
  ...brand.titleBlock("[Document Title]", "[Optional subtitle - or delete this line]"),
  p("Replace the bracketed text above with your document's real title, then delete this whole instructions page before you start writing (or just start typing over it - up to you). Everything below this point is a quick reference showing you what each style looks like; keep it as a cheat sheet in a separate file, or delete it too.", { italics: true, color: "555555" }),

  new Paragraph({ children: [new PageBreak()] }),

  h1("This is a Heading 1"),
  p("Use Heading 1 for the main sections of your document (like this one). To apply it in Word: click into a line of text, then in the Home tab's Styles gallery, click “Heading 1.” It will automatically turn navy blue and bold, in the right size."),

  h2("This is a Heading 2"),
  p("Use Heading 2 for subsections inside a Heading 1 section. Same idea - click the line, then click “Heading 2” in the Styles gallery."),

  h3("This is a Heading 3"),
  p("Use Heading 3 for a smaller subsection inside a Heading 2 section, if you need one more level."),

  h2("Regular paragraph text"),
  p("This is what normal body text looks like - Arial, black, no special formatting. Just type normally and it will look like this by default; you don't need to apply a style for regular paragraphs."),
  note("Text styled like this (italic, red) is an instruction or warning callout - the same style used for “instructor only” notices in the course's Teacher's Keys. Select this text and use Home > Font color if you want to reuse it, or just retype your own note in italics."),

  h2("Bullet and numbered lists"),
  bullet("Click the bullet-list button in the Home tab, or just start a line with a dash and a space and Word will often auto-format it."),
  bullet("Second bullet point."),
  numbered("Numbered lists work the same way - use the numbered-list button in the Home tab."),
  numbered("Second numbered item."),

  h2("A branded table"),
  p("To reuse this exact look for a new table: click anywhere inside the table below, select the whole table (the small cross-arrow icon that appears at its top-left corner), copy it, paste it wherever you need a new table, then type over the example text. This is easier than building a new table's formatting from scratch."),
  new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA },
    columnWidths: [3120, 3120, 3120],
    rows: [
      new TableRow({ children: [headerCell("Column A", 3120), headerCell("Column B", 3120), headerCell("Column C", 3120)] }),
      new TableRow({ children: [bodyCell("Row 1", 3120, { fill: "F3F6FA" }), bodyCell("Row 1", 3120, { fill: "F3F6FA" }), bodyCell("Row 1", 3120, { fill: "F3F6FA" })] }),
      new TableRow({ children: [bodyCell("Row 2", 3120), bodyCell("Row 2", 3120), bodyCell("Row 2", 3120)] }),
    ],
  }),

  h2("Header, footer, and page numbers"),
  p("The header (top of every page) and footer (bottom of every page) are already set up with the “Video Game Creation” tag and automatic page numbers. To change the document title shown in the header, double-click anywhere in the gray header text near the top of the page, then edit the text after the “|” - Word will let you type right into it. Double-click back in the main body of the page when you're done to keep working there."),

  h2("Colors, if you need them directly"),
  p("Navy (#16305C) for main headings and dark backgrounds; Navy 2 (#1C4A86) for secondary headings and accents; Gold (#C99A2E) for borders and highlights; a darker Gold (#8A6D1E) for any gold text on a white background (the brighter gold is hard to read as body text). See the Brand Style Guide document for the full reference."),
];

const doc = new Document({
  numbering: numberingConfig,
  styles: stylesBlock,
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: brand.pageHeader("[Document Title]") },
    footers: { default: brand.pageFooter() },
    children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("SLHE_VGC_Word_Template.docx", buffer);
  console.log("done");
});
