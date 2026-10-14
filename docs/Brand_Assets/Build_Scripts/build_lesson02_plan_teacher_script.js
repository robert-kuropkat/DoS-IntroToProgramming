const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign,
  ImageRun
} = require("docx");
const brand = require("./brand.js");

function diagram(file, widthIn, heightIn, caption) {
  const data = fs.readFileSync(file);
  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: caption ? 40 : 160 },
      children: [new ImageRun({ data, transformation: { width: widthIn * 96, height: heightIn * 96 }, type: "png" })],
    }),
  ];
  if (caption) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 160 },
      children: [new TextRun({ text: caption, italics: true, size: 16, color: "888888" })],
    }));
  }
  return children;
}

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

function pageHeader(text) {
  return new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text, size: 16, color: "888888" })] })] });
}
function pageFooter() {
  return new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 16 }), new TextRun({ children: [PageNumber.CURRENT], size: 16 })] })] });
}

const jobs = [];

// =========================================================================
// 1. ONE-PAGE LESSON PLAN
// =========================================================================
{
  const agendaRows = [
    ["0:00-0:05", "Welcome back, Lesson 1 recap, quick install spot-check"],
    ["0:05-0:20", "Sequences (live demo in Scratch)"],
    ["0:20-0:35", "Events (live demo in Scratch)"],
    ["0:35-0:50", "The Game Loop (concept + \"under the hood\" look)"],
    ["0:50-1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["1:00-1:50", "LAB: Build the Key Press Reactor (start of Semester 1 project)"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 2 - Programming Basics I: Sequences, Events & the Game Loop", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Thursday, August 20, 2026  •  Semester 1, Week 2", italics: true, size: 18, color: "666666" })] }),

        labelBullet("Objectives", "Explain sequences and why order matters; explain events and give a game example; describe the game loop concept; build a working Key Press Reactor - the start of the Semester 1 project."),
        labelBullet("IGDA Tie-In", "Programming - control structures and event-driven programming (introductory)"),
        labelBullet("Featured Engine (lecture demo)", "Scratch"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified (see Teacher's Checklist); Student Handouts, one per engine; printed or emailed Lesson 2 GDD (Lesson_02_GDD.docx) - one per student, the same document that was previewed in Lesson 1"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): One key press triggers one visible reaction."),
        bullet("Level 2 (Complete): All 4 arrow keys move the character in 4 directions."),
        bullet("Level 3 (Refined): Add animation, sound, and/or screen-boundary polish."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 2 Teacher's Script, Lesson 2 Prep Checklist, and the per-engine Student Handouts.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_02_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 2 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 2 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Programming Basics I: Sequences, Events & the Game Loop", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back, everyone! Last week we talked about internet safety, what game development actually involves, and looked at a bunch of different engines. Show of hands - who got their engine installed and tested since last class?\" [Pause, scan the room] \"If you didn't finish, that's okay - we've got time at the start of lab today, and remember GDevelop and Scratch work straight in a browser with zero install if you're in a pinch.\""),
        p("[Demo cue: have your own demo laptop open with Scratch ready]"),
        p("\"Today we're doing our first real programming lesson. Everything from here forward builds on itself, so today we're also starting the actual game project you'll keep working on all semester.\""),

        h2("[0:05-0:20] Sequences"),
        p("\"Let's start with the simplest idea in programming: a sequence. A sequence is just a list of instructions that happen one after another, in order. Think about your morning routine - wake up, get dressed, eat breakfast. If you get dressed before you wake up, that doesn't work! Order matters.\""),
        p("[Demo cue: switch to Scratch, open a blank project]"),
        p("\"Watch what happens when I stack these blocks: move 10 steps, then turn 15 degrees, then say 'Hello!' \" [Click blocks one at a time, then run the stack] \"See how it does them top to bottom, one after another? That's a sequence. Every single line of code in every engine we use this year - GDScript in Godot, C# in Unity, GML in GameMaker - all of it runs as sequences like this at its core.\""),
        p("[Discussion question] \"What's something else in your daily life that has to happen in a specific order to work? Turn to a neighbor and share one example.\""),
        p("[Timing checkpoint: wrap up by 0:18-0:20]", { italics: true, color: "888888" }),

        h2("[0:20-0:35] Events"),
        p("\"Now here's the problem with plain sequences: they just run once, top to bottom, and stop. But games aren't like that - games wait around and respond to what you do. That's called an event. An event is code that says 'when this happens, do that.'\""),
        p("[Demo cue: same Scratch project]"),
        p("\"In Scratch, events start with hat-shaped blocks like 'when green flag clicked' or 'when key pressed.' \" [Drag out \"when right arrow key pressed\" and attach \"move 10 steps\"] \"Watch - nothing happens until I press the right arrow key.\" [Press right arrow several times] \"Every time I press it, that whole sequence underneath runs again from the top - the character moves another 10 steps to the right. Same block as before, just triggered by a key now instead of running once at the start.\""),
        p("\"This is the shift from 'a list of steps that happens once' to 'a list of steps that can happen over and over, whenever something triggers it.' Every engine we use has some version of this - Godot has signals and _input(), Unity has Keyboard.current key checks, GameMaker has the Key Pressed event, and so on. Different names, same idea.\""),
        p("[Discussion question] \"What other events can you think of in games you've played?\" (Examples to draw out if needed: pressing a button, an enemy touching you, a timer running out, opening a chest.)"),
        p("[Display cue: put the Sequences vs. Events diagram up (slide or printed) as a wrap-up visual for both of the last two sections]", { italics: true, color: "888888" }),
        ...diagram("diagram_sequences_events.png", 6.2, 3.968, "Sequences vs. Events - the two ideas covered in the last 30 minutes, side by side."),
        p("[Timing checkpoint: wrap by 0:35]", { italics: true, color: "888888" }),

        h2("[0:35-0:50] The Game Loop"),
        p("\"Last big idea: the game loop. Here's a question - how does a game character keep moving smoothly while you hold down a key, instead of just twitching once?\" [Pause for guesses]"),
        p("\"It's because the engine isn't just waiting for one event and stopping - it's constantly asking, many times every second, usually 30 to 60 times a second: 'what's the current input? what needs to move? what needs to be redrawn?' That loop of check-update-draw, check-update-draw, happening dozens of times a second, is called the game loop. You don't usually write the loop yourself - the engine runs it for you - but almost everything you build this year runs inside it.\""),
        p("[Optional visual: flip-book analogy] \"It's like a flip book - each page is slightly different, and flipping through them fast makes it look like smooth motion. Each 'page' is one pass through the game loop.\""),
        p("[Display cue: put the Game Loop diagram up (slide or printed) - trace Input -> Update -> Render -> back to Input with your finger as you talk through it again]", { italics: true, color: "888888" }),
        ...diagram("diagram_game_loop.png", 6.2, 3.844, "The Game Loop - Input, Update, and Render, repeating dozens of times per second."),
        p("[Timing checkpoint: wrap by 0:48-0:50]", { italics: true, color: "888888" }),

        h2("[Bonus, only if time allows] Thinking in Diagrams"),
        p("This section is optional filler, not part of the timed agenda - skip it entirely if you're at time, or hold it for the start of next lesson. It's a light, no-pressure preview, not new required content.", { italics: true, color: "888888" }),
        p("\"One more thing, if we've got a few minutes: you just described the game loop out loud. Professional developers often draw it instead, using a couple of standard diagram styles. You don't need to memorize these - just start recognizing them when you see them.\""),
        p("\"An Activity Diagram is basically a flowchart - it shows the order things happen in. Here's your Key Press Reactor as one: check the key, and if it's pressed, move the player, then redraw. Same thing you just described, as a picture.\""),
        ...diagram("diagram_activity_keypress.png", 6.2, 3.844, "Activity Diagram - the Key Press Reactor's check/move/redraw flow."),
        p("\"A State Diagram shows the different 'modes' something can be in, and what switches it between them. Your player character has at least two: Idle and Moving. Pressing the key switches it one way, releasing switches it back.\""),
        ...diagram("diagram_state_player.png", 6.2, 3.844, "State Diagram - the player character's Idle and Moving states."),
        p("\"We'll come back to a third kind - Class Diagrams - later in the year once you're writing more object-oriented code. For now, just knowing these two exist, and roughly what they're for, is enough.\""),

        h2("[0:50-1:00] Today's Exercise & Handouts"),
        p("\"For lab today, you're building what I'm calling a 'Key Press Reactor' - an object that responds when you press a key. This is also the very first piece of your Semester 1 project, so save it somewhere you'll remember and give the project a name you'll recognize, like [YourName]Game.\""),
        p("\"Everyone's getting a handout specific to their engine - the steps look a little different depending on what you're using, but the goal is the same for everybody: Level 1 is one key doing one thing. Level 2 is all four arrow keys moving your character around. Level 3 is adding some polish - animation, sound, whatever you want to add.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"You're also getting your real Game Design Document today - the one your Lesson 1 sneak peek was a preview of. Find the 'Lesson 2 Goals' checklist inside it, under Levels / Structure - it's the same three levels from your handout, just in one place so you can check them off as you go. No need to fill in the rest of the document today; that can wait until you've got a working exercise to build on.\""),
        p("[Hand out Lesson_02_GDD.docx - one per student]"),
        p("\"Questions before we start lab? ... Let's go build something.\""),

        h2("Lab Time Notes"),
        bullet("Circulate by engine cluster, starting with whichever engine has the most students this week."),
        bullet("Common early snag across every engine: an event block/node that isn't actually connected to the rest of the logic won't run. If a student says \"nothing happens,\" check connections first."),
        bullet("Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout, not given new unrelated work."),
        bullet("At 1:50, call time and do a 2-minute share-out: ask 2-3 students to show what they built, across different engines if possible."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_02_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
