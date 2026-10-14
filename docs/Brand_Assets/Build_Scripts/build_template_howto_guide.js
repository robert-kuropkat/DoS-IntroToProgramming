const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel,
} = require("docx");
const brand = require("./brand.js");

function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function note(text) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, italics: true, color: "B00020" })] }); }
function numbered(text) { return new Paragraph({ numbering: { reference: "numbers", level: 0 }, spacing: { after: 100 }, children: [new TextRun(text)] }); }
function bullet(text) { return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 }, children: [new TextRun(text)] }); }

const { LevelFormat, AlignmentType: AT } = require("docx");
const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AT.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AT.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ],
};

const children = [
  ...brand.titleBlock("How to Use the SLHE Templates", "A step-by-step guide for the Word and PowerPoint template files"),

  h1("What these files are"),
  p("There are two template files: SLHE_VGC_Word_Template.dotx (for Word documents) and SLHE_VGC_PowerPoint_Template.potx (for PowerPoint slide decks). A template is not a document you edit directly and save over - it's a starting point. Every time you open one, your computer creates a brand-new, unsaved document or presentation that already has the SLHE colors, fonts, and layouts built in. The original template file stays untouched, so you can reuse it as many times as you like."),

  h1("Step 1 - Put the templates where Word and PowerPoint expect them"),
  p("Word and PowerPoint each have a special folder for personal templates called \"Custom Office Templates.\" Once a template is there, it shows up automatically inside the app - you won't need to dig through file folders every time."),
  h2("On Windows"),
  numbered("Open File Explorer."),
  numbered("In the address bar at the top, paste this and press Enter:  %USERPROFILE%\\Documents\\Custom Office Templates"),
  numbered("If that folder doesn't exist yet, create a folder in your Documents named exactly Custom Office Templates."),
  numbered("Copy SLHE_VGC_Word_Template.dotx and SLHE_VGC_PowerPoint_Template.potx into that folder."),
  h2("On a Mac"),
  numbered("Open Word or PowerPoint, click the app's menu at the top-left, and choose Preferences."),
  numbered("Click Save, and note the folder listed next to \"Default personal templates location.\" If it's blank, set it to a folder such as ~/Documents/Custom Office Templates."),
  numbered("Copy the two template files into that folder."),
  note("You only have to do this step once. After the templates are in place, skip straight to Step 2 any time you want to start something new."),

  h1("Step 2 - Start a new document or deck from the template"),
  h2("In Word"),
  numbered("Open Word and click New (or go to File > New)."),
  numbered("Click the Personal tab near the top of the New screen."),
  numbered("Click SLHE_VGC_Word_Template. A brand-new, untitled document opens, already styled - the template file itself is not changed."),
  numbered("Save your new document with File > Save As, using whatever name makes sense (for example, \"Lesson 5 Teacher Notes.docx\")."),
  h2("In PowerPoint"),
  numbered("Open PowerPoint and click New (or go to File > New)."),
  numbered("Click the Personal tab near the top of the New screen."),
  numbered("Click SLHE_VGC_PowerPoint_Template. A brand-new, untitled presentation opens with four starter slides already showing the four available layouts."),
  numbered("Save your new deck with File > Save As, using whatever name makes sense."),
  note("If you don't see a Personal tab, the template files aren't in the right folder yet - double check Step 1, and make sure the folder name is exactly \"Custom Office Templates.\""),
  p("A quicker but slightly less tidy option: just double-click the .dotx or .potx file itself in File Explorer or Finder. That also opens a new, untitled document/presentation rather than editing the template - this works from anywhere, without doing Step 1 at all."),

  h1("Using the Word template"),
  p("The Word template opens with a cover page and, after a page break, a full \"cheat sheet\" showing every style in use with instructions. Delete both once you no longer need them, or keep the cheat sheet as a separate reference file."),
  h2("Applying heading styles"),
  numbered("Click anywhere in the line of text you want to turn into a heading (or select several lines)."),
  numbered("On the Home tab, find the Styles gallery (a row of boxes showing \"Normal,\" \"Heading 1,\" \"Heading 2,\" and so on)."),
  numbered("Click Heading 1 for a main section title, Heading 2 for a subsection, or Heading 3 for a smaller subsection."),
  h2("Editing the header and footer"),
  p("Every page already shows \"VIDEO GAME CREATION | [Document Title]\" at the top and a page number at the bottom. To change the title shown at the top, double-click directly on that gray header text near the top of the page - Word will let you edit it in place. Double-click back in the main body of the page when you're done."),
  h2("Reusing the branded table"),
  p("The cheat-sheet page includes a ready-made 3-column table in the SLHE colors. Click the small cross-shaped icon that appears at its top-left corner to select the whole table, copy it (Ctrl+C or Cmd+C), click where you want a new table, and paste. Then type over the example text."),

  h1("Using the PowerPoint template"),
  p("The template opens with one example slide for each of the four available layouts: Title Slide, Section Divider, Content, and Closing. Duplicate any of them to make more slides in that style, or delete the ones you don't need."),
  h2("Adding a new slide in a specific layout"),
  numbered("On the Home tab, click the small arrow under New Slide (not the icon itself - the arrow just below or beside it)."),
  numbered("You'll see the four SLHE layouts listed by name. Click the one you want."),
  numbered("Click into any title or text placeholder box on the new slide and start typing - the SLHE fonts and colors are already applied."),
  h2("Changing an existing slide's layout"),
  numbered("Click the slide in the left-hand thumbnail list."),
  numbered("On the Home tab, click Layout, and choose a different one of the four SLHE layouts."),
  h2("Duplicating a slide (the easiest way to add more)"),
  numbered("Right-click a slide's thumbnail on the left."),
  numbered("Choose Duplicate Slide."),
  numbered("Edit the text on the copy - the layout and styling come along automatically."),
  note("Every layout already includes the gold \"VIDEO GAME CREATION\" footer text and the SLHE background - you don't need to add these yourself."),

  h1("Quick troubleshooting"),
  bullet("\"I don't see a Personal tab in File > New\" - the template files are not in the Custom Office Templates folder yet, or are in the wrong one. Recheck Step 1, or just double-click the template file directly instead."),
  bullet("\"I accidentally edited the template file itself instead of a new document\" - close it without saving, then re-copy a clean version of the .dotx/.potx from wherever your original course materials are kept."),
  bullet("\"The colors or fonts look different than expected\" - Cambria and Calibri (used in the PowerPoint template) and Arial (used in the Word template) are standard Microsoft Office fonts and should already be installed; if a slide looks off, check that AutoCorrect or a paste operation hasn't overridden the font on that specific text box."),

  h1("Reference: the SLHE brand colors"),
  p("Navy (#16305C) - main headings, dark slide backgrounds."),
  p("Navy 2 (#1C4A86) - secondary headings and accents."),
  p("Gold (#C99A2E) - borders, highlights, gold text on dark backgrounds."),
  p("Darker Gold (#8A6D1E) - gold text on a white background (the brighter gold is hard to read as body text)."),
  p("See Brand_Style_Guide.docx for the complete visual reference, including logo usage rules."),
];

const doc = new Document({
  numbering: numberingConfig,
  styles: brand.stylesBlock,
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: brand.pageHeader("Template How-To Guide") },
    footers: { default: brand.pageFooter() },
    children,
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("SLHE_VGC_Template_HowTo_Guide.docx", buffer);
  console.log("done");
});
