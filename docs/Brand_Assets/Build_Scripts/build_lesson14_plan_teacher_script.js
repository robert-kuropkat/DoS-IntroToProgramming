const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign,
  ImageRun
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
function pRuns(runs, opts = {}) { return new Paragraph({ spacing: { after: 160 }, ...opts, children: runs }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function numbered(text, level = 0) { return new Paragraph({ numbering: { reference: "numbers", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function checkbox(text) { return new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 80 }, children: [new TextRun(text)] }); }
function labelBullet(label, text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)] });
}
function code(lines) {
  return new Paragraph({
    spacing: { after: 160, before: 80 }, indent: { left: 360 },
    shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
    children: [new TextRun({ text: lines.join("\n"), font: "Consolas", size: 20 })]
  });
}
function headerCell(text, width) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: brand.NAVY2, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 18 })] })] });
}
function bodyCell(children, width, opts = {}) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 60, bottom: 60, left: 100, right: 100 },
    verticalAlign: VerticalAlign.TOP,
    children: Array.isArray(children) ? children : [new Paragraph({ children: [new TextRun({ text: children, size: 18 })] })] });
}

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "checks", levels: [{ level: 0, format: LevelFormat.BULLET, text: "☐", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]
};

const stylesBlock = {
  default: { document: { run: { font: "Arial", size: 22 } } },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 30, bold: true, font: "Arial", color: brand.NAVY },
      paragraph: { spacing: { before: 300, after: 160 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 24, bold: true, font: "Arial", color: brand.NAVY2 },
      paragraph: { spacing: { before: 220, after: 120 }, outlineLevel: 1 } },
    { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 22, bold: true, font: "Arial", color: "444444" },
      paragraph: { spacing: { before: 140, after: 80 }, outlineLevel: 2 } },
  ]
};

const jobs = [];

// =========================================================================
// 1. ONE-PAGE LESSON PLAN
// =========================================================================
{
  const agendaRows = [
    ["0:00-0:05", "Welcome back, Lesson 13 recap"],
    ["0:05-0:15", "The Foreshadowed Feature: One Formula, Two Places"],
    ["0:15-0:35", "Adding Speed to the Mix (live demo in GameMaker)"],
    ["0:35-0:45", "This Week's Twist: Student's Choice"],
    ["0:45-0:55", "Introduce lab exercise, hand out engine-specific handouts"],
    ["0:55-1:50", "LAB: Scale count and speed, then pick one stretch feature"],
    ["1:50-2:00", "Exit check / quick share-out - toughest wave 3 wins bragging rights"],
  ];
  const agendaWidths = [1600, 7760];
  const agendaTable = new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: agendaWidths,
    rows: agendaRows.map((row, ri) => new TableRow({
      children: [bodyCell(row[0], agendaWidths[0], { fill: ri % 2 ? "F3F6FA" : "FFFFFF" }), bodyCell(row[1], agendaWidths[1], { fill: ri % 2 ? "F3F6FA" : "FFFFFF" })]
    }))
  });

  const doc = new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 900, right: 1080, bottom: 900, left: 1080 } } },
      children: [
        new Paragraph({ spacing: { after: 20 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 14, bold: true, color: brand.GOLD_TEXT, characterSpacing: 16 })] }),
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 14 - Raising the Stakes", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 14 - Difficulty Scaling & Stretch Features", italics: true, size: 18, color: brand.GOLD_TEXT, bold: true })] }),

        labelBullet("Objectives", "Replace hardcoded per-wave enemy count and speed with formulas keyed off the existing wave variable, keeping every referenced value in agreement; choose and make a good-faith attempt at ONE optional capstone-level stretch feature."),
        labelBullet("IGDA Tie-In", "Game Design - balance, pacing, difficulty curves (intermediate); Programming - parameterization, avoiding duplicated logic (intermediate)"),
        labelBullet("Featured Engine (lecture demo)", "GameMaker (GML text/script - second script-mode rotation slot, callback to Lesson 9)"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified; Student Handouts, one per engine; students' own Lesson 13 projects (already saved); no new assets - 001 Game Creator students add 2 pre-placed enemy copies at design time"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): enemy count increases by one each wave (5, 6, 7), using a single formula referenced everywhere it's needed."),
        bullet("Level 2 (Complete): enemy speed ALSO increases each wave; restarting correctly returns everything to wave 1's original values."),
        bullet("Level 3 (Refined): choose exactly ONE of five capstone-level stretch features (shields, power-up, bonus enemy, formation movement, title/game-over screen) and make a good-faith attempt."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 14 Teacher's Script, Lesson 14 Prep Checklist, and the per-engine Student Handouts. Level 3 this week is open-ended student choice - expect a wider range of finishing points than usual, and that's fine.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_14_Plan.docx"]);
}

// =========================================================================
// 2. TEACHER'S SCRIPT
// =========================================================================
{
  const doc = new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 14 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 14 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Raising the Stakes", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Your game moves, shoots, sounds good, and reacts to hits. Today, it starts to feel like a REAL arcade game - one that gets harder the longer you survive.\""),
        p("\"Some of you may remember a hint from Lesson 11's Teacher's Key, or a comment in the TIC-80 or Scratch materials, that hard-coding 5 enemies per wave was 'a deliberate simplification.' Today's the day that pays off.\""),

        h2("[0:05-0:15] The Foreshadowed Feature: One Formula, Two Places"),
        p("\"Here's the whole idea in one sentence: instead of always spawning 5 enemies, we spawn (4 + wave) enemies. Wave 1: still 5. Wave 2: 6. Wave 3: 7. One small formula, and your game already escalates.\""),
        p("\"The catch: that formula has to appear in TWO places that already exist in your code - wherever you spawn enemies, AND wherever enemies_remaining gets reset for a new wave. If those two disagree, your win-check breaks. This is exactly the kind of 'one source of truth' thinking we've been building since Lesson 10's reusable reset function.\""),
        p("[Timing checkpoint: wrap by 0:15]", { italics: true, color: "888888" }),

        h2("[0:15-0:35] Adding Speed to the Mix"),
        p("[Demo cue: switch to GameMaker, open a project with Lesson 13's flash and explosion working]"),
        p("\"Count alone is a good start, but speed makes it feel dangerous. Watch: in spawn_wave(), right after instance_create_layer() makes each enemy, I set inst.enemy_speed = 1 + (wave - 1) * 0.5. Wave 1 stays exactly the same. Wave 3 is noticeably faster.\" [Build it live, demonstrate wave 1 versus wave 3 side by side if time allows]"),
        p("\"That new enemy_speed variable has to actually get used, so flip over to obj_enemy's Step event: it currently just says y += 1. Change that one line to y += enemy_speed, and suddenly the number we set at spawn time is doing real work every frame. Notice what did NOT change: the rest of the Step event, from Lesson 4. It already ran every frame - today we just made one of its numbers depend on the wave.\""),
        p("\"TIC-80 folks: your engine doesn't have a per-instance speed variable yet, since Lesson 7 only offered that as an OPTIONAL idea. Today it becomes required - you're adding it for real, if you haven't already.\""),
        p("[Timing checkpoint: wrap by 0:35]", { italics: true, color: "888888" }),

        h2("[0:35-0:45] This Week's Twist: Student's Choice"),
        p("\"Level 3 works differently this week. Instead of 2-3 small stretch ideas, you're choosing ONE feature from a list of five real capstone-level ideas: destructible shields, a power-up, a bonus enemy, formation movement, or a title/game-over screen.\""),
        p("\"These are open-ended on purpose. A good-faith attempt at the general idea - even if it's rough - is a complete Level 3 finish today. Don't feel like you need to polish it. Pick whichever one sounds the most fun to YOU.\""),
        p("[Timing checkpoint: wrap by 0:45]", { italics: true, color: "888888" }),

        h2("[0:45-0:55] Today's Exercise & Handouts"),
        p("\"For lab today, open your own project from last week - not a new one. No new assets needed, except 001 Game Creator folks, who are adding two more pre-placed enemy copies in the Map editor first.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"Level 1 and 2 today, then pick your Level 3 feature. Test a full 3-wave round when you're done, and confirm restarting brings everything back to wave 1's original feel.\""),
        p("\"Questions before we start lab? ... Let's raise the stakes.\""),

        h2("Lab Time Notes"),
        bullet("The most common mistake this week: updating the spawn-count formula in one place but not the enemies_remaining reset in the other - watch for this specifically, it's an easy one to half-do."),
        bullet("001 Game Creator students need a few extra minutes up front to correctly add and name enemy_6/enemy_7 in the Map editor before any scripting begins."),
        bullet("Difficulty scaling can tip into feeling unfair if a student's increments are too aggressive - if wave 3 feels impossible rather than exciting, help them dial back the numbers rather than pushing through."),
        bullet("Level 3 this week is genuinely open-ended - circulate and help students scope their chosen feature down to something achievable in the remaining lab time, rather than letting anyone get stuck chasing a fully-polished version."),
        bullet("At 1:50, do a quick share-out - ask a few students to demo their wave 3 and their chosen Level 3 feature, whatever state it's in."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_14_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
