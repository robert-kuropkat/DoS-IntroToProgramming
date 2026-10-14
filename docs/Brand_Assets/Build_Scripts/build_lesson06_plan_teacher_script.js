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
    ["0:00-0:05", "Welcome back, Lesson 5 recap, quick share of enemy rows"],
    ["0:05-0:25", "Tracking State: Counting Down to Zero (live demo in Scratch)"],
    ["0:25-0:45", "Announcing the Win: Each Engine's Simplest Text Tool"],
    ["0:45-0:50", "Milestone Moment: What \"Level 1 MVP\" Actually Means"],
    ["0:50-1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["1:00-1:50", "LAB: Add a win condition to your ongoing [YourName]Game project"],
    ["1:50-2:00", "Exit check / quick share-out - milestone celebration"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 6 - Programming Basics V: Victory Condition", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 6 - Level 1 MVP Milestone", italics: true, size: 18, color: brand.GOLD_TEXT, bold: true })] }),

        labelBullet("Objectives", "Explain what a counter/state variable is and how it's used to track \"how many are left\"; explain why each engine's simplest built-in text display is enough for now, without needing a full UI system; explain what completing the Level 1 MVP milestone means for the capstone project; implement a working win condition."),
        labelBullet("IGDA Tie-In", "Programming - state tracking, win/lose conditions, event-driven programming (introductory)"),
        labelBullet("Featured Engine (lecture demo)", "Scratch"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified (see Teacher's Checklist); Student Handouts, one per engine; students' own Lesson 5 projects (already saved from last week) - no new assets needed this week"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): \"You Win!\" appears the instant the last enemy is destroyed."),
        bullet("Level 2 (Complete): The win message appears exactly once, exactly on time - no errors, no early or repeated triggers."),
        bullet("Level 3 (Refined): Add a victory sound and/or make the win message more of a celebration."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 6 Teacher's Script, Lesson 6 Prep Checklist, and the per-engine Student Handouts.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_06_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 6 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 6 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Programming Basics V: Victory Condition", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Last week you turned one enemy into a whole row of them. Quick show of hands - who got every enemy independently destructible? Anyone try the Level 3 visual variety idea?\" [Pause, call on 1-2 volunteers]"),
        p("\"Today's a big one. By the end of class, your game will actually be winnable - shoot every enemy, and the game tells you so. This finishes the very first real milestone of your whole project.\""),

        h2("[0:05-0:25] Tracking State: Counting Down to Zero"),
        p("\"Quick question: your game has 5 enemies right now. How would you know, in code, the exact moment all 5 are gone?\" [Pause for guesses - someone may suggest checking a list's length, or counting]"),
        p("\"Both of those work! Today we'll use the simpler of the two: a single number that starts at 5 and counts down by one every time an enemy is destroyed. The moment that number hits zero, we know instantly - no need to check anything else. This kind of tracked number is usually called a counter, and it's one of the most common patterns in all of programming, not just games.\""),
        p("[Demo cue: switch to Scratch, open a project with a working enemy row from Lesson 5]"),
        p("\"Watch: I've got a variable called enemies_left, starting at 5. It's set 'for all sprites', not just this one, because both the Enemy clones and the win check need to see the same shared number. Every time a clone's touching-check destroys one, it changes enemies_left by -1. When it hits zero...\" [Destroy all enemies on screen] \"...this broadcast fires and the message appears. That's the whole idea - one shared number, counted down by one event at a time.\""),
        p("[Discussion question] \"What other things in a game might need a counter like this? \" (Examples to draw out if needed: lives remaining, ammo left, score, a timer counting down.)"),
        p("[Timing checkpoint: wrap by 0:25]", { italics: true, color: "888888" }),

        h2("[0:25-0:45] Announcing the Win: Each Engine's Simplest Text Tool"),
        p("\"Now that we know WHEN we've won, how do we tell the player? Every engine has some basic way to just draw text on screen, and that's genuinely all we need today - not a fancy menu, not a scoreboard, just words appearing.\""),
        p("[Demo cue: same Scratch project] \"In Scratch, that's just a 'say' block - a speech bubble so simple it's barely UI at all. Other engines are just as simple: Godot has a Label node, GameMaker has draw_text(), TIC-80 has print(), Unreal has a node literally called Print String. None of these need the real UI systems your games will eventually have - that's coming properly in a few weeks, once you've got a score and lives to actually display.\""),
        p("\"The lesson here: don't reach for the fanciest tool before you need it. A plain line of text is a completely legitimate way to confirm a win condition works, and it's exactly what a real developer would build first too, before ever touching final art or layout.\""),
        p("[Discussion question] \"Can you think of other times in a project where the 'quick and ugly' version is actually the right first step? \" (Examples: placeholder art, a console log instead of a UI, a text menu instead of a styled one.)"),
        p("[Timing checkpoint: wrap by 0:43-0:45]", { italics: true, color: "888888" }),

        h2("[0:45-0:50] Milestone Moment: What \"Level 1 MVP\" Actually Means"),
        p("\"Here's something worth pausing on. Once your win condition works today, you will have a complete, playable game loop: move, shoot, face enemies, and win. That's not a small thing - that's the Level 1 MVP milestone from your capstone's design document, and it's the first of several big checkpoints across the whole year.\""),
        p("\"MVP stands for Minimum Viable Product - the smallest version of something that's still a real, complete, playable experience. It's not supposed to be impressive yet. It's supposed to work, start to finish, which is genuinely the hardest part of finishing any project. Everything from here through Lesson 15 is about making this same core loop more complete, then more polished - but the loop itself, today, is done.\""),
        p("[Timing checkpoint: wrap by 0:50]", { italics: true, color: "888888" }),

        h2("[0:50-1:00] Today's Exercise & Handouts"),
        p("\"For lab today, open your own project from last week - not a new one, and no new assets needed. Everything today builds directly on your enemy row from Lesson 5.\""),
        p("\"The goal: track enemies remaining, and show 'You Win!' the instant it hits zero. Level 1 is getting that working once. Level 2 is confirming it fires exactly once, at exactly the right moment. Level 3 is turning a plain message into an actual celebration - sound, animation, whatever you want.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"One heads-up: today's win-tracking code needs to live in a slightly different place than last week's hit-detection - your handout walks through exactly where for your engine.\""),
        p("\"Questions before we start lab? ... Let's go finish your first milestone.\""),

        h2("Lab Time Notes"),
        bullet("The single most common bug this week: an off-by-one error - the starting counter not matching the actual number of enemies spawned. Have students double-check their Lesson 5 spawn count first."),
        bullet("Second most common: win-tracking code placed somewhere that runs every frame instead of only when an enemy is actually destroyed - this can cause the message to flicker or the counter to behave unpredictably."),
        bullet("Circulate by engine cluster as usual - no engine is especially harder this week, unlike Lesson 5's Scratch wrinkle."),
        bullet("Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout, not given new unrelated work."),
        bullet("At 1:50, call time and do a real celebration share-out - this is the Level 1 MVP milestone, worth more than the usual quick show-and-tell. Consider having 3-4 students demo a full win from start to finish."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_06_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
