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
    ["0:00-0:05", "Welcome back, Lesson 4 recap, quick share of any working combat mechanics"],
    ["0:05-0:25", "Loops: Doing Something Many Times Without Repeating Yourself (live demo in GameMaker)"],
    ["0:25-0:45", "Scaling Behavior: Why One Enemy's Code Already Works for Many"],
    ["0:45-0:50", "A Scratch-Specific Wrinkle: Clones Have No ID"],
    ["0:50-1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["1:00-1:50", "LAB: Turn your single enemy into a row of enemies in your ongoing [YourName]Game project"],
    ["1:50-2:00", "Exit check / quick share-out"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 5 - Programming Basics IV: Enemy Ranks", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 5", italics: true, size: 18, color: "666666" })] }),

        labelBullet("Objectives", "Explain what a loop is and why it avoids repeating code by hand; explain why well-structured per-instance behavior from Lesson 4 automatically scales to many instances without changes; explain Scratch's specific clone-identity limitation and its correct workaround; turn a single enemy into a row of independently-destructible enemies."),
        labelBullet("IGDA Tie-In", "Programming - loops/iteration, code reuse and object-oriented thinking (introductory)"),
        labelBullet("Featured Engine (lecture demo)", "GameMaker (GML Visual primary, with a brief GML script contrast - see this lesson's Sidebars)"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified (see Teacher's Checklist); Student Handouts, one per engine; students' own Lesson 4 projects (already saved from last week) - no new assets needed this week"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): Several enemies appear in a row, created by a loop; shooting any one destroys just that one."),
        bullet("Level 2 (Complete): Every enemy can be independently shot and destroyed with no errors; misses still behave exactly like last week."),
        bullet("Level 3 (Refined): Sound already works for free; add row-to-row visual variety and/or confirm clean spawn-vs-behavior structure."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 5 Teacher's Script, Lesson 5 Prep Checklist, and the per-engine Student Handouts.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_05_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 5 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 5 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Programming Basics IV: Enemy Ranks", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Last week you got a working enemy - shooting it destroyed it cleanly, even with a few misses mixed in. Quick show of hands - who got that working? Anyone try the Level 3 ideas?\" [Pause, call on 1-2 volunteers]"),
        p("\"Today, one enemy becomes many. By the end of today, your game will have a whole row of enemies, and you'll do it without writing the same setup code five separate times.\""),

        h2("[0:05-0:25] Loops: Doing Something Many Times Without Repeating Yourself"),
        p("\"Quick question: if I wanted five enemies on screen right now, what's the laziest way? \" [Pause for guesses - someone will likely say \"copy and paste the code five times\"] \"That would actually work! But imagine changing something about how enemies look later - now you're editing five separate copies instead of one. That's exactly the problem a loop solves.\""),
        p("\"A loop just means: run this same block of instructions a certain number of times, instead of writing it out repeatedly. You've already used a version of this without maybe noticing - your bullets loop and your movement checks run every single frame, which is itself a loop the engine drives for you. Today we write our own loop, on purpose, to create several enemies from one block of code.\""),
        p("[Demo cue: switch to GameMaker, open a project with a single enemy already built from Lesson 4, GML Visual tab open on obj_spawner's Create event]"),
        p("\"Watch: instead of placing one enemy by hand, I'm going to drag in a Repeat block set to run 5 times, and inside it, a Create Instance block that makes one more enemy at a slightly different position each pass.\" [Build the Repeat block and the Create Instance block live, run the game] \"Five blocks, and I get five independent enemies. If I want ten instead of five, I change one number on the Repeat block, not five copy-pasted objects.\""),
        p("[Demo cue: flip the same event over to its Code tab] \"Same event, same five-block idea - here's what it looks like typed instead.\" [Show the equivalent GML: a plain for loop calling instance_create_layer() five times] \"Same logic either way - blocks and typed code are just two different skins on the same idea. GameMaker's the one engine in this class that genuinely lets you flip back and forth on the same event like this, which is exactly why it's demoing loops today - we'll dig into where that flexibility actually helps versus where it doesn't in today's Sidebar.\""),
        p("[Discussion question] \"Where else might a loop save you from repeating yourself in a game? \" (Examples to draw out if needed: drawing a grid of tiles, giving every enemy in a wave the same starting health, checking every item in an inventory.)"),
        p("[Timing checkpoint: wrap by 0:25]", { italics: true, color: "888888" }),

        h2("[0:25-0:45] Scaling Behavior: Why One Enemy's Code Already Works for Many"),
        p("\"Here's the part that might surprise you. Last week, you wrote the code for what happens when ONE enemy gets hit. Today, without changing a single line of that code, it's going to correctly handle FIVE enemies. Why does that work?\" [Pause for guesses]"),
        p("\"Remember the Structure idea from Lesson 4 - that an object's own script should own its own response to being hit? Because you kept that boundary clean, every single copy the loop creates carries its own independent copy of that same behavior. Enemy #3 getting hit has absolutely nothing to do with Enemy #1's code - they're each running their own complete, separate copy.\""),
        p("[Demo cue: same GameMaker project] \"Watch what happens if I shoot enemies out of order - the third one, then the first, then the last.\" [Fire at enemies in a mixed order] \"Each one reacts completely independently. Nothing about hitting one affects the others, and I didn't have to write anything special to make that true - it's just a direct result of keeping behavior with the object it belongs to.\""),
        p("\"This is a big deal, and it's going to keep paying off all semester. Good structure now means big new features later - like a whole wave of enemies - come almost for free.\""),
        p("[Discussion question] \"Can anyone think of a moment in a game where LOTS of things need to react to a similar event independently? \" (Examples: every brick in a brick-breaker, every zombie in a horde game, every light in a level flickering on its own.)"),
        p("[Timing checkpoint: wrap by 0:43-0:45]", { italics: true, color: "888888" }),

        h2("[0:45-0:50] A Scratch-Specific Wrinkle: Clones Have No ID"),
        p("\"One quick heads-up, mainly for anyone on Scratch today. Scratch's clones are a little different from instances in every other engine we use - they don't have a unique ID number you can check. That means a Bullet clone literally cannot ask 'which specific Enemy clone am I touching?' - it can only ask 'am I touching any Enemy at all?'\""),
        p("\"So this week, Scratch students are going to move the hit-check from the Bullet's script onto the Enemy's own script instead - each Enemy clone asks 'am I touching a Bullet?' about itself, which works correctly no matter how many enemies exist. It's a real change from last week's code, not just an extension of it - your handout walks through exactly what moves where.\""),
        p("[Timing checkpoint: wrap by 0:50]", { italics: true, color: "888888" }),

        h2("[0:50-1:00] Today's Exercise & Handouts"),
        p("\"For lab today, open your own project from last week - not a new one, and no new assets needed either. Everything today builds directly on your enemy and collision detection from Lesson 4.\""),
        p("\"The goal: a row of enemies instead of one, created by a loop. Level 1 is getting that working once. Level 2 is confirming every enemy in the row can be independently destroyed with no errors. Level 3 is mostly free wins this week - your sound already works, so it's really about adding a bit of visual variety across the row.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"One reminder specific to today: most of you will need to delete the single enemy you manually placed last week before your new loop-spawned row will look right - it's easy to forget and end up with six enemies instead of five.\""),
        p("\"Questions before we start lab? ... Let's build an army - well, a small row of them.\""),

        h2("Lab Time Notes"),
        bullet("Scratch needs real hands-on support this week - the touching-check has to move from Bullet to Enemy, which is a genuine restructure, not just an addition. Expect this engine's cluster to take longer than usual."),
        bullet("The single most common bug in every other engine: forgetting to delete last week's manually-placed single enemy, ending up with one extra enemy that doesn't belong to the row."),
        bullet("Second most common: all spawned enemies stacking on top of each other - almost always means the loop's index variable isn't actually being used in the position calculation."),
        bullet("Circulate by engine cluster, starting with Scratch this week given the extra complexity there."),
        bullet("Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout, not given new unrelated work."),
        bullet("At 1:50, call time and do a 2-minute share-out: ask 2-3 students to show what they built, across different engines if possible."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_05_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
