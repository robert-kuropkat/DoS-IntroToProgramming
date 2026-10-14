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
    ["0:00-0:05", "Welcome back, Lesson 6 recap, milestone check-in"],
    ["0:05-0:20", "New Unit: Core Mechanics - what's ahead through Lesson 11"],
    ["0:20-0:40", "Movement That Doesn't Stop: Continuous Motion (live demo in Godot)"],
    ["0:40-1:00", "The Hidden Bug: When 'Destroyed' Doesn't Mean 'Defeated'"],
    ["1:00-1:05", "Introduce lab exercise, hand out engine-specific handouts"],
    ["1:05-1:55", "LAB: Advance enemies toward the player + escape cleanup"],
    ["1:55-2:00", "Exit check / quick share-out"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 7 - Core Mechanics I: On the Move", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 7 - Core Mechanics unit begins", italics: true, size: 18, color: brand.GOLD_TEXT, bold: true })] }),

        labelBullet("Objectives", "Implement continuous, frame-rate-independent movement; distinguish an object being removed for one reason (a hit) from being removed for another (escaping off-screen); explain why generic \"on destroyed\" lifecycle hooks can silently misfire once more than one destruction pathway exists; correctly relocate state-changing logic to the specific event that represents it."),
        labelBullet("IGDA Tie-In", "Programming - game loops and continuous simulation, event-driven programming, debugging and code correctness (introductory)"),
        labelBullet("Featured Engine (lecture demo)", "Godot Engine, contrasted briefly with GameMaker for the bug discussion"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified (see Teacher's Checklist); Student Handouts, one per engine; students' own Lesson 6 projects (already saved) - no new assets needed this week"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): every enemy moves slowly downward; an enemy that reaches the bottom unshot is removed without counting as a kill."),
        bullet("Level 2 (Complete): a moving enemy can still be shot correctly, and an escaped enemy never triggers \"You Win!\" by accident."),
        bullet("Level 3 (Refined): give enemies varied speeds, and/or tighten the separation between movement/escape code and hit-detection/win-counting code."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 7 Teacher's Script, Lesson 7 Prep Checklist, and the per-engine Student Handouts. GameMaker and 001 Game Creator instructors: read the Prep Checklist's bugfix note before class.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_07_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 7 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 7 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Core Mechanics I: On the Move", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Last week you hit the Level 1 MVP milestone - a full, winnable game loop. Quick show of hands - who showed their win screen to somebody at home?\" [Pause, call on 1-2 volunteers]"),
        p("\"Today we start making that loop feel like an actual game instead of a static puzzle. First change: your enemies are about to start moving.\""),

        h2("[0:05-0:20] New Unit: Core Mechanics"),
        p("\"You just finished the 'Programming Basics' stretch of the course - variables, loops, collision, state tracking. Starting today, we're in a new unit: Core Mechanics. Over the next five lessons, we'll add losing, scoring, a real UI, and multiple waves - by the end of Lesson 11, your game will be a complete, replayable arcade shooter.\""),
        p("\"Today's piece specifically: enemies stop sitting still. Every one of them will now creep toward the player. That one change touches almost everything else we build from here.\""),
        p("[Timing checkpoint: wrap by 0:20]", { italics: true, color: "888888" }),

        h2("[0:20-0:40] Movement That Doesn't Stop: Continuous Motion"),
        p("\"So far, everything in your game has moved in response to something - a key press, a spawn. Today's movement is different: it just keeps happening, every single frame, with nothing triggering it.\""),
        p("[Demo cue: switch to Godot, open a project with a working win condition from Lesson 6]"),
        p("\"Watch: I've added a few lines to the enemy's script that run automatically, every frame, forever - move down a little, check if it's gone too far. No key press, no collision, nothing triggers this. It just runs.\" [Run the scene, let enemies drift down]"),
        p("\"Notice I multiplied the movement by delta - the time since the last frame. Why do you think that matters?\" [Pause for guesses] \"Right - without it, a fast computer would move enemies faster than a slow one. Multiplying by delta keeps the speed the same no matter how fast the game is running. You've actually seen this pattern before, all the way back in Lesson 2's player movement - today's just the first time an enemy uses it too.\""),
        p("[Discussion question] \"What's the actual difference between 'runs every frame' and 'runs when something happens'? Can you think of something else in your game, or in a game you've played, that has to run continuously rather than in response to an event?\" (Examples to draw out: a timer counting down, a camera following the player, gravity.)"),
        p("[Timing checkpoint: wrap by 0:40]", { italics: true, color: "888888" }),

        h2("[0:40-1:00] The Hidden Bug: When \"Destroyed\" Doesn't Mean \"Defeated\""),
        p("\"Here's the part of today that matters most. Once enemies can move, they can now leave the screen at the bottom without ever being shot. That means there are now TWO different reasons an enemy might stop existing: getting shot, or escaping. And escaping should never count as a win toward your goal.\""),
        p("\"For most of you, this just means adding a new, separate check - no changes to what you already built. But if you're in GameMaker or 001 Game Creator, we need to talk about something that's been quietly wrong since last week, and today is when it actually breaks.\""),
        p("[Demo cue: switch to GameMaker, open a Lesson-6-complete project]"),
        p("\"Last week, your win-counter was wired to the Destroy event - a built-in event that fires whenever an enemy is destroyed, no matter why. That was totally correct last week, because the ONLY way an enemy could be destroyed was by being shot. But today we're adding a second way - escaping - and the Destroy event can't tell the difference. It'll fire, and quietly decrement your win-counter, for an enemy that should NOT have counted.\""),
        p("\"This is a really common category of bug in real software, not just games: code that was completely correct when it was written becomes wrong later, not because it changed, but because something else around it changed. The fix isn't complicated once you see it: move the win-counter to the one event that actually, specifically means 'this enemy got shot' - the collision event - and leave the generic Destroy event out of win-counting entirely.\""),
        p("[Demo cue: show the fix - Collision event gets the decrement, Destroy event goes empty] \"See the difference? Collision with a bullet is specific - it can only ever mean one thing. Destroyed is generic - it means 'stopped existing,' full stop, for any reason at all. Attach your important logic to the specific event, not the generic one.\""),
        p("\"001 Game Creator has the exact same issue, with its 'Defeated / Destroyed' trigger instead of a Destroy event - your handout walks through the identical fix there. If you're in one of the other six engines, your Lesson 6 build already tied the win-counter to the specific hit-detection code, so you have nothing to fix - just new code to add. Worth noticing why that turned out to be lucky, or well-designed, depending how you look at it.\""),
        p("[Discussion question] \"Can anyone think of a real-world example of this same idea - a rule that was fine until something else changed and it quietly started causing problems?\" (Open-ended; encourage a few guesses even if imperfect - the pattern-recognition matters more than a perfect answer.)"),
        p("[Timing checkpoint: wrap by 1:00]", { italics: true, color: "888888" }),

        h2("[1:00-1:05] Today's Exercise & Handouts"),
        p("\"One more thing before you open your projects. Looking back at Lesson 6, we noticed the win-counter ended up living in a different place in almost every engine - the main scene here, the spawner there, one project's collision branch somewhere else. Only one of you, building in Unity, had it living on its own dedicated object built just to hold that kind of information. That turned out to be the better habit, so we've updated the Lesson 6 materials to match it everywhere, and I want you to update your own project too.\""),
        p("\"It's a quick change, not a rebuild: create one object whose only job is holding game-wide information like the win-counter, and have everything else report to it instead of touching that number directly. Your updated Lesson 6 handout for your engine now shows exactly what that looks like - use it as your reference. Do this before you start today's movement work, since today's bug-fix discussion is about the exact same question: which piece of code should own which job.\""),
        p("[Give students 3-5 minutes to make the change if they haven't already, or confirm it's already in place. Circulate - GameMaker and Scratch tend to have the most questions here, since the retrofit is least like anything else they've built.]"),
        p("\"Quick question before we move on: why bother with a separate object for this at all? Couldn't the win-counter just keep living on the spawner, or the player, like it did before?\" [Pause for guesses]"),
        p("\"A few real reasons this is worth doing, not just tidiness for its own sake. First - anyone joining this project later, including future-you in three weeks, has exactly one place to look for shared information, instead of having to remember or rediscover which object happens to be holding it this time. Second - it means objects don't have to reach into each other's private variables to get work done; the Bullet doesn't need to know anything about how the Spawner is built, it just talks to the GameManager. Third, and this is the one that pays off soonest: next lesson we're adding lives, and the lesson after that, score. Both of those are exactly the same kind of shared, game-wide information as today's win-counter - and they already have an obvious home waiting for them.\""),
        p("[Demo cue, optional: pull up any one engine's GameManager object live and point out it currently holds exactly two things - point out that by Lesson 9 it'll hold four, and that's the design working as intended, not scope creep.]"),
        p("\"This pattern has a bunch of different names depending on who you ask - Game Manager, Singleton, Game State - but it's the same idea everywhere: one object, one job, holding information the rest of the game needs to share.\""),
        p("\"For lab today, open your own project from last week - not a new one, and no new assets needed. Every engine's handout has two parts: adding movement and escape cleanup, and - for GameMaker and 001 Game Creator only - fixing last week's win-counter.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"GameMaker and 001 Game Creator folks - read the setup section of your handout carefully before touching anything. You're deleting code today, not just adding it.\""),
        p("\"Questions before we start lab? ... Let's get these enemies moving.\""),

        h2("Lab Time Notes"),
        bullet("The single most important thing to verify this week, for every student regardless of engine: let an enemy escape off the bottom unshot, and confirm \"You Win!\" does NOT appear. This is the real test of Level 2, not just watching enemies move."),
        bullet("GameMaker / 001 Game Creator: watch specifically for double-counting - a student who added the new win-counter code but forgot to remove the old generic-hook code will see the win message trigger after only about half the enemies are shot. This is the most likely mistake this week; check for it directly rather than waiting for a student to notice."),
        bullet("A student whose enemies never disappear at the bottom likely has an off-screen threshold that doesn't match their project's actual screen size - have them adjust the number rather than guess at a fix."),
        bullet("Circulate by engine cluster as usual. GameMaker and 001 Game Creator clusters need the most attention this week."),
        bullet("Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout, not given new unrelated work."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_07_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
