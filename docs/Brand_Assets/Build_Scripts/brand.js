// =========================================================================
// SLHE "Video Game Creation" course - shared Word-doc branding module.
// Single source of truth for the palette/fonts used by the Lesson 1 slide
// deck (generate_deck.js), the Lesson 2 slide deck (generate_deck2.js), and
// the Game Design Document template (build_gdd2.js). Require this from any
// docx-js build script to keep every course document visually consistent.
// =========================================================================
const fs = require("fs");
const path = require("path");
const {
  Header, Footer, Paragraph, TextRun, AlignmentType, HeadingLevel,
  PageNumber, ImageRun, TabStopType,
} = require("docx");

const NAVY = "16305C";       // primary heading / crest shield color
const NAVY2 = "1C4A86";      // secondary heading / accent color
const GOLD = "C99A2E";       // accent gold (table headers, rules)
const GOLD_TEXT = "8A6D1E";  // darker gold, legible as text-on-white
const TEXT_MUTE = "666666";
const CONTENT_WIDTH = 9360;  // standard US Letter content width (DXA), 1in margins

const FONT = "Arial";

// Crest asset - resolved relative to the build script's cwd (outputs/), same
// convention as every other asset reference (icons/*.png, bg_*.png, etc.)
const CREST_PATH = "vgc_crest.png";
const CREST_ASPECT = 1040 / 840; // height / width of vgc_crest.png

const stylesBlock = {
  default: { document: { run: { font: FONT, size: 22 } } },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 30, bold: true, font: FONT, color: NAVY },
      paragraph: { spacing: { before: 340, after: 160 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 24, bold: true, font: FONT, color: NAVY2 },
      paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 } },
    { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 22, bold: true, font: FONT, color: "444444" },
      paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
  ],
};

// Right-aligned running header: "VIDEO GAME CREATION | <docTitle>"
function pageHeader(docTitle) {
  return new Header({
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT,
      children: [
        new TextRun({ text: "VIDEO GAME CREATION", size: 16, color: GOLD_TEXT, bold: true, characterSpacing: 20 }),
        new TextRun({ text: "  |  " + docTitle, size: 16, color: TEXT_MUTE }),
      ],
    })],
  });
}

// Footer: "SLHE Game Programming" left, "Page N" right (tab-stop aligned)
function pageFooter() {
  return new Footer({
    children: [new Paragraph({
      tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_WIDTH }],
      children: [
        new TextRun({ text: "SLHE Game Programming", size: 16, color: TEXT_MUTE }),
        new TextRun({ text: "\tPage ", size: 16, color: TEXT_MUTE }),
        new TextRun({ children: [PageNumber.CURRENT], size: 16, color: TEXT_MUTE }),
      ],
    })],
  });
}

// Crest image paragraph, centered, for use on cover/title pages.
// widthIn defaults to 1.5in (matches the GDD template's cover treatment).
function crestParagraph(widthIn = 1.5, opts = {}) {
  const data = fs.readFileSync(path.resolve(CREST_PATH));
  const w = widthIn, h = widthIn * CREST_ASPECT;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: opts.after ?? 200 },
    children: [new ImageRun({ data, transformation: { width: w * 96, height: h * 96 }, type: "png" })],
  });
}

// Standard branded title block for a cover/first page: crest + title +
// optional subtitle + the "VIDEO GAME CREATION" tag line.
function titleBlock(title, subtitle) {
  const out = [crestParagraph(1.3, { after: 160 })];
  out.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: subtitle ? 40 : 60 },
    children: [new TextRun({ text: title, bold: true, size: 40, color: NAVY })],
  }));
  if (subtitle) {
    out.push(new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { after: 60 },
      children: [new TextRun({ text: subtitle, size: 22, color: NAVY2 })],
    }));
  }
  out.push(new Paragraph({
    alignment: AlignmentType.CENTER, spacing: { after: 300 },
    children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: GOLD_TEXT, characterSpacing: 20 })],
  }));
  return out;
}

module.exports = {
  NAVY, NAVY2, GOLD, GOLD_TEXT, TEXT_MUTE, CONTENT_WIDTH, FONT,
  stylesBlock, pageHeader, pageFooter, crestParagraph, titleBlock,
};
