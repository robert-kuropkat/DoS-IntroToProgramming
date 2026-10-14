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
    ["0:00-0:05", "Welcome back, Lesson 9 recap"],
    ["0:05-0:20", "The Last Missing Piece: What 'Over' Should Actually Mean"],
    ["0:20-0:40", "One Setup, Two Doors In (live demo in Godot)"],
    ["0:40-0:55", "The Freeze: A Flag That Gates Everything"],
    ["0:55-1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["1:00-1:50", "LAB: Build game-over freeze and restart flow"],
    ["1:50-2:00", "Exit check / quick share-out - full playthrough demos"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 10 - Core Mechanics IV: Play It Again", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 10 - Game-Over + Restart Flow", italics: true, size: 18, color: brand.GOLD_TEXT, bold: true })] }),

        labelBullet("Objectives", "Refactor existing one-time setup code into a reusable reset routine, callable from two different entry points; introduce a game-state flag that gates whether gameplay logic still applies; implement a full, clean restart triggered by a dedicated key press."),
        labelBullet("IGDA Tie-In", "Programming - state machines (introductory), code reuse/refactoring, input handling (intermediate)"),
        labelBullet("Featured Engine (lecture demo)", "Godot Engine"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified (see Teacher's Checklist); Student Handouts, one per engine; students' own Lesson 9 projects (already saved) - no new assets needed this week"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): when the round ends, the game freezes; pressing R resets everything and starts a new round."),
        bullet("Level 2 (Complete): a restart is a genuinely clean slate - no leftover enemies, no stale score/lives, works correctly any number of times in a row."),
        bullet("Level 3 (Refined): add a transition effect, and/or tighten the single-source-of-truth structure of the reset routine."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 10 Teacher's Script, Lesson 10 Prep Checklist, and the per-engine Student Handouts. Every engine involves a genuine refactor this week (existing setup code restructured into a reusable form), not just new additions.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_10_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 10 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 10 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Core Mechanics IV: Play It Again", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Last week your game finally got a real scoreboard. Quick show of hands - who styled their HUD, or added scoring variety?\" [Pause, call on 1-2 volunteers]"),
        p("\"Today we fix something that's probably bugged a few of you since Lesson 8: your game 'ends,' but it doesn't actually stop. Today it finally does - and you can play again without touching the Run button.\""),

        h2("[0:05-0:20] The Last Missing Piece: What 'Over' Should Actually Mean"),
        p("\"Right now, when you win or lose, a message shows up and... nothing else changes. Enemies keep coming. You can keep shooting. The numbers might even keep moving. 'Game Over' hasn't actually meant anything yet.\""),
        p("\"Today, 'over' means something real: the round stops mattering, and there's a clear way to start a fresh one. That's two pieces - freezing, and restarting - and they both lean on the exact same trick.\""),
        p("[Discussion question] \"Think of any game you've played - what usually happens the instant you lose? What are all the different things that might stop or change?\" (Draw out: enemies stop, controls lock, a menu or prompt appears, some kind of restart option shows up - list a few examples on the board if helpful.)"),
        p("[Timing checkpoint: wrap by 0:20]", { italics: true, color: "888888" }),

        h2("[0:20-0:40] One Setup, Two Doors In"),
        p("\"Here's the trick: your game already HAS code that sets everything up - score at 0, lives at 3, enemies in place. It's been running since Lesson 5, just once, right at the start. Today, we give that exact same code a second way to run: on demand, whenever R is pressed.\""),
        p("[Demo cue: switch to Godot, open a project with a working scoreboard from Lesson 9]"),
        p("\"Watch: I'm taking the setup - score at 0, lives at 3, enemies in place - and moving it into its own function on my GameManager, start_game(). The enemy-spawning loop that used to run in _ready() becomes spawn_enemies() on the Main scene, and start_game() calls it. Then GameManager's _ready() just calls start_game once, deferred - Godot runs a child's _ready() before its parent has finished loading, and the parent won't accept new children until it has, so call_deferred waits for that. That's the whole refactor.\" [Build it live] \"Now here's the payoff: I can call start_game() again, any time I want, from anywhere else in my code - like, say, when R is pressed.\" [Wire up the restart key, demonstrate a full reset]"),
        p("\"Notice nothing about the SETUP LOGIC itself changed. We didn't write a second version of it for restarting. We just gave the existing version a second doorway in.\""),
        p("[Discussion question] \"Why is it better to have ONE piece of setup code with two ways to trigger it, instead of writing the setup code twice - once for the start, once for restart?\" (Draw out: if you ever need to change how the game sets up - say, starting with 4 lives instead of 3 - you'd only have to change it in one place instead of two, and remembering to update both copies is exactly the kind of thing that's easy to forget.)"),
        p("[Timing checkpoint: wrap by 0:40]", { italics: true, color: "888888" }),

        h2("[0:40-0:55] The Freeze: A Flag That Gates Everything"),
        p("\"Restarting is only half of today. The other half: actually STOPPING the game once it's over, so nothing keeps happening behind that win/lose message.\""),
        p("[Demo cue: same Godot project] \"I've added one new variable on GameManager - game_active, true by default. The moment you win or lose, it flips to false. And now, at the very top of my enemy's movement code, one line: if the game isn't active, don't do anything else this frame.\" [Show enemies freezing on game-over]"),
        p("\"One flag, checked in the right place, freezes the whole game. That's genuinely it - no special 'pause' feature required, just a question asked at the right moment, every frame: 'is the game still going?'\""),
        p("[Show the next slide: the Activity Diagram] \"Here's the whole round in one picture. The round ends, game_active flips to false, and everything freezes. Then we ask two questions when R is pressed: was R actually pressed, and is the round actually over? Only if both answers are yes do we run start_game() - the same one reset routine that runs when the game first launches. Press R in the middle of a round and the answer to the second question is no, so nothing happens.\""),
        p("[Show the next slide: What Is Game State?] \"One more idea before lab, because it will save you an hour of confusion. Everything your game remembers right now is called its game state: the score, the lives, how many enemies are left, whether the round is over, where the player is standing, which picture the player is showing, even whether a sound or an animation is mid-play. Until today, your game only ever started once and ended once, so nothing had to be put back. A restart changes the rule: every single piece of state has to go back to how it was at the start.\""),
        p("[Show the next slide: The Restart Audit] \"So today you audit. Find every place your game does something when it starts - in Scratch, every 'when green flag clicked' script on every sprite and the stage; in Godot, every _ready(); in GameMaker, every Create event; in Unity, every Awake and Start. Then find everything you did at game over - an animation, a color change, a hidden sprite, a sound, a message. If something changed when the round ended, something has to change it back when the round restarts. If you added something extra that isn't in the handout, it is on your list too, because the handout can't know about it.\""),
        p("[Discussion question] \"What do you think would happen if we forgot to check game_active somewhere - say, in the code that handles a bullet hitting an enemy? \" (Draw out: that specific thing would keep happening even after game-over - a good preview of exactly the kind of bug students should test for today.)"),
        p("[Timing checkpoint: wrap by 0:55]", { italics: true, color: "888888" }),

        h2("[0:55-1:00] Today's Exercise & Handouts"),
        p("\"For lab today, open your own project from last week - not a new one, and no new assets needed. Every engine's handout has the same two moves: turn your existing setup code into something callable again, and add a game_active-style flag that freezes things once the round ends.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"Test in both directions today, and in both game-end states: play to a win, restart, confirm it's clean. Then play to a loss, restart, confirm that's clean too. And don't skip testing that R does NOTHING while a round is still active - that's an easy thing to accidentally break.\""),
        p("\"Questions before we start lab? ... Let's make your game actually replayable.\""),

        h2("Lab Time Notes"),
        bullet("The single most important thing to verify this week: a full round-trip, twice - win then restart, and separately lose then restart. Both should leave the game in an identical, clean starting state."),
        bullet("Most likely mistake across every engine: forgetting to clear old enemies before respawning on restart, leading to doubled enemies or a wrong count after the first restart. Check for this explicitly, don't just eyeball the screen."),
        bullet("Game state drift: any start-up code or game-over effect a student added beyond the handout (extra green-flag scripts, animations, color or costume changes, hidden or moved objects, sounds) is NOT reset by the handout steps. Walk the room with the Restart Audit checklist from the handout; the symptom is usually a restart that looks almost right (a player stuck in the wrong costume, a message still showing, an animation that never replays)."),
        bullet("Second most likely mistake: R doing something mid-game because the game_active check was left out of the restart-key logic. Test this directly by pressing R while actively playing."),
        bullet("This is a genuinely satisfying lesson to finish - encourage a full playthrough as part of the end-of-lesson share-out (win, restart, lose, restart) since it's the first time the game truly feels replayable, not just winnable once."),
        bullet("Level 3 bug hunt: there is a real bug in the win/lose logic from Lessons 6 and 8. A round with some kills and some escapes (say 3 and 2) ends with no win and no loss, and R does nothing. Ask early finishers to find it by playing a mixed round; accept any fix that guarantees every round ends in exactly one outcome. The shown fix: an escape also reduces enemies_remaining, and the win check requires lives above 0."),
        bullet("Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout, not given new unrelated work."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_10_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
