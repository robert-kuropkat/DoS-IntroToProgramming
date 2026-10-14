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
    ["0:00-0:05", "Welcome back, Lesson 12 recap"],
    ["0:05-0:15", "A Hit That Reacts: Two New Visual Touches"],
    ["0:15-0:40", "Pausing Mid-Script: Three Engine Families, Three Solutions (live demo in Scratch)"],
    ["0:40-0:50", "An Explosion That Minds Its Own Business"],
    ["0:50-0:55", "Introduce lab exercise, hand out engine-specific handouts"],
    ["0:55-1:50", "LAB: Wire up the flash and explosion, verify game logic is unchanged"],
    ["1:50-2:00", "Exit check / quick share-out - best explosion wins bragging rights"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 13 - A Hit That Reacts", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 13 - Basic Animation", italics: true, size: 18, color: brand.GOLD_TEXT, bold: true })] }),

        labelBullet("Objectives", "Sequence a visual reaction (flash) before finalizing an existing game-state change, without altering when score/lives/wave/win-lose logic actually resolves; spawn a fully independent, self-cleaning explosion effect; recognize the three different ways this course's engines pause mid-script (or don't) and why."),
        labelBullet("IGDA Tie-In", "Visual Art - basic animation, juice/feedback fundamentals (introductory); Programming - asynchronous/timed execution, event sequencing (intermediate)"),
        labelBullet("Featured Engine (lecture demo)", "Scratch, contrasted with GameMaker's Alarm-based workaround"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified; Student Handouts, one per engine; students' own Lesson 12 projects (already saved); one new asset (explosion_placeholder.png) - Unreal reuses its existing Cube/Sphere shapes, TIC-80 students hand-draw their own explosion sprite"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): every hit briefly flashes the enemy and spawns an independent explosion effect; score/lives/wave/win-lose logic all still resolve at exactly the same moments as before."),
        bullet("Level 2 (Complete): the flash and explosion play correctly on every hit with no stuck enemies or leftover explosions, and game-state timing is provably unaffected."),
        bullet("Level 3 (Refined): smoother color fades, a scaling/animated explosion, or a multi-frame flicker instead of a flat flash."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 13 Teacher's Script, Lesson 13 Prep Checklist, and the per-engine Student Handouts. This lesson restructures hit-handling code more than most weeks - review your demo engine's Handout closely, and skim the others before class.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_13_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 13 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 13 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "A Hit That Reacts", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Last week your game got its voice. This week, it gets a body - a hit is about to actually LOOK like a hit, not just sound like one.\""),
        p("\"Since Lesson 4, an enemy you shoot has just... vanished. Functional, but not satisfying. Today we fix that: a brief flash, then a small explosion, before it disappears.\""),

        h2("[0:05-0:15] A Hit That Reacts: Two New Visual Touches"),
        p("\"Two things are happening today, and they're different kinds of things. First: the enemy itself briefly flashes - a reaction ON the thing that got hit. Second: a completely separate explosion effect appears at that spot and plays itself out - independent of the enemy entirely.\""),
        p("\"Here's the rule for today, and it's important: your score, your lives, your wave counter, your win and lose logic - none of that changes WHEN it happens. Only how the enemy's removal LOOKS is changing. If you find yourself moving scoring code around, stop and re-check the Handout.\""),
        p("[Timing checkpoint: wrap by 0:15]", { italics: true, color: "888888" }),

        h2("[0:15-0:40] Pausing Mid-Script: Three Engine Families, Three Solutions"),
        p("[Demo cue: switch to Scratch, open a project with Lesson 12's four sounds working]"),
        p("\"Here's the actual challenge this week: to see a flash, the enemy has to stick around for a fraction of a second AFTER being hit, before it's actually removed. That means something in our code needs to pause - just for a moment - without freezing the whole game.\""),
        p("\"Scratch actually has the simplest answer of any engine in this class: a plain 'wait' block. Watch - right in the Enemy's touching-Bullet branch, after scoring and the sound, I set the color effect to flash it white, then just drop in 'wait 0.1 seconds', then move the existing win-check and delete-this-clone below that wait.\" [Build it live, demonstrate the flash-then-disappear on a hit] \"Everything below that wait block just... happens a moment later. No special keyword needed - Scratch was built from the ground up to run every sprite's scripts side by side, so pausing one of them was never a hard problem to solve.\""),
        p("\"Godot and Unity get the same result with a bit more ceremony - GDScript's await keyword and C#'s Coroutines both pause a script mid-line, the same idea as Scratch's wait, just spelled with more setup. But GameMaker and TIC-80 CAN'T do this at all - their code can't pause mid-script, full stop. So they solve it differently: GameMaker sets an Alarm and finishes the job in a second event, a fraction of a second later. TIC-80 does the Lua version of the same trick with a countdown variable. Four engines, two genuinely different families of solution to the same problem - and that's worth understanding, not just copying.\""),
        p("[Timing checkpoint: wrap by 0:40]", { italics: true, color: "888888" }),

        h2("[0:40-0:50] An Explosion That Minds Its Own Business"),
        p("\"Now the easier half: the explosion effect. It spawns at the hit location and IMMEDIATELY forgets about everything else - it doesn't know your score, doesn't know your wave, doesn't even know which enemy made it. It just plays itself out and cleans itself up, on its own.\""),
        p("\"This should feel familiar - it's the exact same discipline your bullets have followed since Lesson 3, and your enemies since Lesson 5. Every object in this game is responsible for exactly one job, and cleans up after itself. Today's explosion is just the newest member of that club.\""),
        p("[Timing checkpoint: wrap by 0:50]", { italics: true, color: "888888" }),

        h2("[0:50-0:55] Today's Exercise & Handouts"),
        p("\"For lab today, open your own project from last week - not a new one. One new asset is provided: a small explosion image. Unreal folks, you're using your existing Cube/Sphere shapes instead; TIC-80 folks, you're drawing your own tiny explosion sprite.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"Test a FULL round when you're done - shooting, a wave clear, even a loss - and confirm nothing about your SCORING changed, only the visuals. That's today's real Level 2 check.\""),
        p("\"Questions before we start lab? ... Let's make getting hit feel like something.\""),

        h2("Lab Time Notes"),
        bullet("The most common mistake this week: accidentally moving score/win-check logic along with the delay, instead of leaving it immediate and only delaying the enemy's final removal. Watch for this specifically."),
        bullet("GameMaker and TIC-80 students will need a bit more circulating time this week - splitting logic across two events/table-fields is a bigger structural change than most weeks. Budget extra minutes for these clusters."),
        bullet("001 Game Creator: remind students the enemy's own flash may be brief or inconsistent in their installed version - that's expected and documented, not something to debug endlessly. The explosion effect is the reliable part."),
        bullet("This is a fun, visual week - encourage students to actually watch a few hits happen slowly rather than just confirming it \"works,\" and consider a quick show-and-tell of favorite explosion tweaks during the share-out."),
        bullet("If a cluster finishes Level 1 quickly, point them at Level 3's animation-timing ideas (fading flash, scaling explosion, multi-frame flicker) rather than rushing ahead to Lesson 14's material."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_13_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
