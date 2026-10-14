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
    ["0:00-0:05", "Welcome back, Lesson 3 recap, quick share of any working shooting mechanics"],
    ["0:05-0:25", "Collision Detection: How Engines Know Things Touched (live demo in Godot)"],
    ["0:25-0:45", "Cleanup Revisited: Destroying Two Things at Once, Safely"],
    ["0:45-0:50", "Placement: Things You Spawn vs. Things You Place"],
    ["0:50-1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["1:00-1:50", "LAB: Add an enemy and collision detection to your ongoing [YourName]Game project"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 4 - Programming Basics III: Direct Hit", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 4", italics: true, size: 18, color: "666666" })] }),

        labelBullet("Objectives", "Explain what collision/overlap detection is and why every engine needs an explicit way to ask \"did these two things touch\"; explain the difference between a trigger/area-based check and a solid physical collision; explain why responding to one collision by destroying two objects is safe even in edge cases (near-simultaneous hits); add a working single-enemy-plus-collision mechanic to the ongoing Semester 1 project."),
        labelBullet("IGDA Tie-In", "Programming - collision/physics systems, event-driven programming (introductory)"),
        labelBullet("Featured Engine (lecture demo)", "Godot Engine"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified (see Teacher's Checklist); Student Handouts, one per engine; students' own Lesson 3 projects (already saved from last week)"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): A single enemy sits on screen; a bullet touching it destroys both."),
        bullet("Level 2 (Complete): The enemy is destroyed cleanly and exactly once, with no errors even under repeated or near-simultaneous hits; misses still behave exactly like last week."),
        bullet("Level 3 (Refined): Add a hit sound, a distinct enemy look, and/or clean hit-response structure."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 4 Teacher's Script, Lesson 4 Prep Checklist, and the per-engine Student Handouts.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_04_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 4 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 4 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Programming Basics III: Direct Hit", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Last week you got your character shooting - a key press spawned a bullet that flew up the screen and cleaned itself up. Quick show of hands - who got that working? Anyone try one of the Level 3 ideas - a sound, a distinct bullet look?\" [Pause, call on 1-2 volunteers]"),
        p("\"Today your bullets finally have something to hit. By the end of today, your game will have one enemy on screen, and shooting it will destroy it. This is also going to teach us something every single game needs: how does an engine even know two things touched?\""),

        h2("[0:05-0:25] Collision Detection: How Engines Know Things Touched"),
        p("\"Here's the question: your bullet is one object, the enemy is another. As far as the computer's concerned, they're just two separate things with their own positions, updating independently. So how does the game know the exact moment they touch?\" [Pause for guesses]"),
        p("\"Every engine's answer is basically the same shape: you give each object a collision shape - usually a simple box or circle, not the exact pixel outline of the artwork - and the engine checks, every frame, whether any two of those shapes overlap. When they do, it fires off a signal or event you can react to. That's collision detection, or sometimes 'overlap detection' when nothing's meant to physically block anything, just detect contact.\""),
        p("[Demo cue: switch to Godot, open a project with bullet.tscn and enemy.tscn already built]"),
        p("\"In Godot, that shape lives on an Area2D node with a CollisionShape2D inside it, and the signal is called area_entered. GameMaker just gives you a Collision event, no shape node needed - you pick the object, and it fires when they overlap. Unity needs a Collider2D on each object, marked 'Is Trigger,' and calls OnTriggerEnter2D. Unreal calls its version Event ActorBeginOverlap. Different names, same underlying question: 'did these two things just start touching?'\""),
        p("[Live demo] \"Let's run this and shoot the enemy.\" [Fire at the enemy, point at both disappearing] \"See how both the bullet and the enemy vanished the instant they overlapped? That's the collision signal firing, and our code reacting to it.\""),
        p("[Discussion question] \"Where else in a game do you think this same idea shows up? \" (Examples to draw out if needed: picking up a coin, walking through a doorway, an attack landing on an enemy, a character falling into a hazard.)"),
        p("[Timing checkpoint: wrap by 0:25]", { italics: true, color: "888888" }),

        h2("[0:25-0:45] Cleanup Revisited: Destroying Two Things at Once, Safely"),
        p("\"Last week you learned every spawned object needs an explicit cleanup call. Today's twist: one collision needs to clean up TWO objects at once - the bullet that hit, and the enemy that got hit. That raises a real question: what if that cleanup code somehow runs twice for the same enemy? Does the game crash?\" [Pause for guesses]"),
        p("\"Good news: in every engine we use, calling 'destroy this' a second time on something that's already been destroyed is completely harmless - the engine just ignores the repeat. That matters, because in a fast-moving game, it's entirely possible for two things to overlap on the exact same frame, and you don't want that rare case to crash the game.\""),
        p("[Demo cue: same Godot project] \"Notice something else about where this code lives: it's the Enemy's own script that reacts to being hit - not the Bullet's. The bullet doesn't need to know an enemy exists at all; it just needed a collision shape added to it so the enemy could detect it. Whoever gets hit is the one responsible for deciding what happens next. That's the same 'keep behavior with the object it belongs to' principle from Lesson 3, applied to a new situation.\""),
        p("[Live demo] \"Watch what happens if I fire several bullets quickly, including a few that miss.\" [Fire a rapid burst, some hitting, some missing] \"The ones that hit destroy the enemy exactly once. The ones that miss just keep flying and disappear off-screen, exactly like last week - nothing about that changed.\""),
        p("[Discussion question] \"Can anyone think of a game moment where you've seen two things need to react to hitting each other at the same time? \" (Examples: two characters colliding head-on, a bullet and a shield meeting.)"),
        p("[Timing checkpoint: wrap by 0:43-0:45]", { italics: true, color: "888888" }),

        h2("[0:45-0:50] Placement: Things You Spawn vs. Things You Place"),
        p("\"One more quick distinction before lab. Last week's bullet is spawned while the game is running - it doesn't exist until you press fire. Today's enemy is different: it's placed once, ahead of time, sitting in the scene/room/level from the very start. Both are still instances - copies of a template - they just come into existence at different moments: one at design time, one at runtime.\""),
        p("\"You'll see this same split all semester. Some things - bullets, effects, pickups - get created while the game runs. Other things - level layout, a boss, a fixed obstacle - get placed by the designer ahead of time. Today you're doing the second kind for the first time.\""),
        p("[Timing checkpoint: wrap by 0:50]", { italics: true, color: "888888" }),

        h2("[0:50-1:00] Today's Exercise & Handouts"),
        p("\"For lab today, open your own project from last week - not a new one. Everything today gets added on top of the character and bullet you already built.\""),
        p("\"The goal: one enemy sits on screen, and shooting it destroys both the enemy and the bullet that hit it. Level 1 is getting that working once. Level 2 is making sure it stays clean even with several shots in a row, including ones that miss. Level 3 is polish - a hit sound, a distinct enemy look, or tidying up how the hit-response code is organized.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"Heads up on the most likely snag today: this is the first week that needs a real collision/physics component set up correctly - a Rigidbody2D in Unity, an Area2D with a shape in Godot, overlap events enabled in Unreal. If nothing happens when you shoot the enemy, that setup is almost always the first thing to check - it's in your handout's 'If Something's Wrong' section.\""),
        p("\"Questions before we start lab? ... Let's go make contact.\""),

        h2("Lab Time Notes"),
        bullet("The single most common bug this week: \"nothing happens when I shoot the enemy.\" This is almost always a missing or misconfigured collision component - Unity needs a Rigidbody2D on at least one of the two objects; Godot needs a real CollisionShape2D inside each Area2D; Unreal needs \"Generate Overlap Events\" ticked; 001 Game Creator needs both Actors marked \"Solid.\" Check this first before debugging the logic."),
        bullet("Second most common: the bullet destroys itself on things it shouldn't, or the enemy never gets destroyed at all - in most engines this traces back to a missing or mismatched tag/group/Scripting ID check."),
        bullet("Circulate by engine cluster, starting with whichever engine has the most students this week."),
        bullet("Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout, not given new unrelated work."),
        bullet("At 1:50, call time and do a 2-minute share-out: ask 2-3 students to show what they built, across different engines if possible."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_04_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
