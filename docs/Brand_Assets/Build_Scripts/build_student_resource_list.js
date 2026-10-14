const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Header, Footer, AlignmentType,
  LevelFormat, ExternalHyperlink, HeadingLevel, PageNumber
} = require("docx");
const brand = require("./brand.js");

function link(text, url) { return new ExternalHyperlink({ children: [new TextRun({ text, style: "Hyperlink" })], link: url }); }
function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function res(name, url, desc) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 80 },
    children: [
      new TextRun({ text: name + ": ", bold: true }),
      link(url, url),
      new TextRun(" - " + desc)
    ] });
}
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]
};

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
    ]
  },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: brand.pageHeader("Student Resource List") },
    footers: { default: brand.pageFooter() },
    children: [
      brand.crestParagraph(1.3, { after: 160 }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Student Resource List", bold: true, size: 40, color: brand.NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Supplemental tutorials, references, and free asset sites", size: 24, color: brand.NAVY2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

      p("Use these alongside weekly lessons. Engine-specific sections cover official docs and beginner tutorial paths; the general sections at the end (art, audio, design) apply no matter which engine you're using."),

      h1("Scratch"),
      res("Scratch itself", "https://scratch.mit.edu", "browser-based editor and project gallery"),
      res("Scratch Tutorials", "https://scratch.mit.edu/ideas", "official starter tutorials built into the editor"),
      res("Coding Cards (PDF)", "https://scratch.mit.edu/starter-projects", "printable step-by-step project guides"),

      h1("Godot"),
      res("Official Docs", "https://docs.godotengine.org", "the primary reference for GDScript and the editor"),
      res("Godot's own First 2D Game tutorial", "https://docs.godotengine.org/en/stable/getting_started/first_2d_game/index.html", "the standard official beginner project"),
      res("GDQuest (YouTube & free courses)", "https://www.gdquest.com", "widely-used free/paid Godot tutorials"),

      h1("Unity"),
      res("Unity Learn", "https://learn.unity.com", "official free tutorials and courses"),
      res("Unity Essentials Pathway", "https://learn.unity.com/pathway/unity-essentials", "guided intro path - start here"),
      res("Junior Programmer Pathway", "https://learn.unity.com/pathway/junior-programmer", "C# scripting-focused path toward certification"),
      res("Creative Core Pathway", "https://learn.unity.com/pathway/creative-core", "design/art-leaning path for less code-focused students"),
      res("Unity Certifications", "https://unity.com/products/unity-certifications", "official certification exams (paid, Pearson VUE)"),

      h1("GameMaker"),
      res("GameMaker Official Tutorials", "https://gamemaker.io/en/tutorials", "official beginner-to-advanced tutorial library"),
      res("GameMaker Manual", "https://manual.gamemaker.io", "full GML and GML Visual reference"),

      h1("GDevelop"),
      res("GDevelop Wiki/Docs", "https://wiki.gdevelop.io", "official documentation and tutorials"),
      res("GDevelop official tutorials page", "https://gdevelop.io/tutorials", "step-by-step starter games"),

      h1("TIC-80"),
      res("TIC-80 official site & docs", "https://tic80.com", "getting started guide and wiki"),
      res("TIC-80 Wiki (GitHub)", "https://github.com/nesbox/tic-80/wiki", "language references for Lua, JS, Python, and more"),

      h1("001 Game Creator"),
      res("Documentation Wiki", "https://001gamecreator.com/wiki/", "official reference for the graphical scripting system and editor"),
      res("STEM Learning Courses", "https://001gamecreator.com/educational-courses/", "structured classroom-oriented lessons"),
      res("Guides and Tutorials", "https://001gamecreator.com/video-guides/", "official written and video tutorials"),
      res("Educational Version Info", "https://001gamecreator.com/educational/", "classroom licensing tiers and the optional E-Book/quiz bundle"),

      h1("Unreal Engine (optional track)"),
      res("Unreal Online Learning", "https://dev.epicgames.com/community/learning", "official free courses, including Blueprints for beginners"),
      res("Unreal Blueprints Quick Start", "https://dev.epicgames.com/documentation/en-us/unreal-engine/blueprints-visual-scripting-in-unreal-engine", "official visual scripting introduction"),

      h1("Free Art, Audio & Design Assets (any engine)"),
      res("Kenney.nl", "https://kenney.nl", "large library of free, license-friendly 2D/3D game art"),
      res("OpenGameArt.org", "https://opengameart.org", "community-contributed free game art and audio"),
      res("Freesound.org", "https://freesound.org", "free sound effects (check each clip's license)"),
      res("Piskel", "https://www.piskelapp.com", "free browser-based pixel art / sprite editor"),
      res("Audacity", "https://www.audacityteam.org", "free, open-source audio editor for sound effects and music"),
      res("itch.io", "https://itch.io", "browsing inspiration, and eventually a place to publish/share your finished game"),

      h1("General Game Design & Programming Concepts"),
      bullet("IGDA Game Education SIG (framework this course is aligned to): igda.org/sigs/game-education"),
      bullet("Ask your instructor about age-appropriate books/videos on game design fundamentals if you want to go deeper between engines"),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Student_Resource_List.docx", buffer);
  console.log("done");
});
