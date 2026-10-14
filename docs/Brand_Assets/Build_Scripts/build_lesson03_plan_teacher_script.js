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
    ["0:00-0:05", "Welcome back, Lesson 2 recap, quick share of any finished Key Press Reactors"],
    ["0:05-0:25", "Instances: One Template, Many Copies (live demo in Godot)"],
    ["0:25-0:45", "Cleanup: Why You Have to Remove What You Create"],
    ["0:45-0:50", "One Press, One Shot (edge-triggered vs. held-down input)"],
    ["0:50-1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["1:00-1:50", "LAB: Add shooting to your ongoing [YourName]Game project"],
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
        new Paragraph({ spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 3 - Programming Basics II: Take Your Shot", bold: true, size: 26, color: brand.NAVY })] }),
        new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text: "Semester 1, Week 3", italics: true, size: 18, color: "666666" })] }),

        labelBullet("Objectives", "Explain the template/instance relationship (one object definition, many independent runtime copies); explain why spawned objects must be explicitly cleaned up; explain the difference between an edge-triggered (\"just pressed\") and level-triggered (\"currently held\") input check; add a working shoot mechanic to the ongoing Semester 1 project."),
        labelBullet("IGDA Tie-In", "Programming - object instancing, resource/memory management, input handling (introductory)"),
        labelBullet("Featured Engine (lecture demo)", "Godot Engine"),
        labelBullet("Materials", "Demo laptop with all 8 engines verified (see Teacher's Checklist); Student Handouts, one per engine; students' own Lesson 2 projects (already saved from last week)"),

        new Paragraph({ spacing: { before: 100, after: 80 }, children: [new TextRun({ text: "Agenda", bold: true, size: 20, color: brand.NAVY2 })] }),
        agendaTable,

        new Paragraph({ spacing: { before: 160, after: 60 }, children: [new TextRun({ text: "Levels for this lesson's exercise", bold: true, size: 20, color: brand.NAVY2 })] }),
        bullet("Level 1 (Minimal): One key press spawns one projectile that flies upward and disappears off-screen."),
        bullet("Level 2 (Complete): Firing repeatedly produces independent projectiles, capped by a fire-rate cooldown."),
        bullet("Level 3 (Refined): Add a fire sound, a distinct projectile look, and/or clean spawn/move/cleanup structure."),

        new Paragraph({ spacing: { before: 160 }, children: [new TextRun({ text: "See: Lesson 3 Teacher's Script, Lesson 3 Prep Checklist, and the per-engine Student Handouts.", italics: true, size: 18, color: "666666" })] }),
      ]
    }]
  });
  jobs.push([doc, "Lesson_03_Plan.docx"]);
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
      headers: { default: brand.pageHeader("Lesson 3 Teacher's Script") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 3 Teacher's Script", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Programming Basics II: Take Your Shot", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("[0:00-0:05] Welcome Back"),
        p("\"Welcome back! Last week you built a Key Press Reactor and started your Semester 1 project. Quick show of hands - who got their character moving in all 4 directions? Anyone try one of the Level 3 stretch ideas - blinking, sound, recoloring?\" [Pause, call on 1-2 volunteers to describe what they tried]"),
        p("\"Today we're adding the next real piece of your game: shooting. By the end of today, pressing a key will fire something off your character and up the screen. This is also going to teach us something new and important - what happens when your game needs to create new things while it's running, not just at the start.\""),

        h2("[0:05-0:25] Instances: One Template, Many Copies"),
        p("\"Here's a question: when you press your fire button five times fast, you want five separate bullets on screen, right? Not one bullet that teleports around. So where do those five bullets come from?\" [Pause for guesses]"),
        p("\"The answer is something almost every engine calls the same thing: an instance. You build ONE definition of what a bullet is - what it looks like, how it behaves - and then, every time you press fire, the engine stamps out a brand new, completely independent COPY of that definition. That's an instance. Change one bullet's position, and it doesn't affect any of the others - they're all separate.\""),
        p("[Demo cue: switch to Godot, open a simple project with a Bullet scene already built]"),
        p("\"Watch the vocabulary here, because it's almost identical across every engine we use. In Godot, the function is literally called instantiate(). \" [Show a line of code: var b = bullet_scene.instantiate()] \" In GameMaker it's instance_create_layer(). In Unity it's Instantiate() - same word, capital I. Unreal calls it Spawn Actor from Class. All of them mean the exact same thing: take a template, make a new independent copy, place it in the game.\""),
        p("[Live demo] \"Let's run this and press the fire key a few times.\" [Press fire repeatedly, point at multiple bullets on screen] \"See how they're all moving independently? Each one remembers its own position separately from all the others, and separately from the template they were copied from.\""),
        p("[Discussion question] \"Where else in a game might you want lots of independent copies of one thing? \" (Examples to draw out if needed: enemies, coins/pickups, particles, trees in a forest.)"),
        p("[Timing checkpoint: wrap by 0:25]", { italics: true, color: "888888" }),

        h2("[0:25-0:45] Cleanup: Why You Have to Remove What You Create"),
        p("\"Now here's the part that trips people up. If every fire button press creates a brand new bullet, and you fire a hundred times over the course of a level... what happens to all those old bullets once they've flown off the top of the screen?\" [Pause for guesses]"),
        p("\"If you don't do anything - they just stay. Forever. Even after they're invisible, off-screen, and completely useless to the player, the engine is still tracking every single one, still checking its position every frame, still drawing it if it's on-screen. A hundred forgotten bullets doesn't sound like much, but a real game spawning things constantly - bullets, particles, enemies - can pile up thousands of forgotten objects in minutes, and the game slows to a crawl or crashes.\""),
        p("[Demo cue: same Godot project] \"This is why every engine also gives you an explicit way to remove an instance once you're done with it. In Godot it's called queue_free(). GameMaker calls it instance_destroy(). Unity and Unreal both literally call it Destroy(). Different names, same job: tell the engine 'I'm done with this one, take it off the list.'\""),
        p("[Live demo] \"Watch what happens if I comment out the cleanup line and fire a bunch of times.\" [If feasible, show a debug counter or the scene tree growing; otherwise describe it] \"See the count climbing and never coming back down? Now watch with cleanup turned back on.\" [Re-enable, fire again, show the count staying flat as bullets appear and disappear]"),
        p("\"So the pattern for anything you spawn at runtime is always two-sided: something that creates it, and something that eventually destroys it. Today, that second half is triggered by the bullet leaving the screen - but later this semester, you'll destroy things for other reasons too, like an enemy getting hit, or a pickup getting collected.\""),
        p("[Discussion question] \"Can anyone think of a game where you've noticed things get removed after they're used or off-screen? \" (Examples: arrows in a shooting game, particle bursts, collected coins.)"),
        p("[Timing checkpoint: wrap by 0:43-0:45]", { italics: true, color: "888888" }),

        h2("[0:45-0:50] One Press, One Shot"),
        p("\"Quick, important one before we go build: when you check whether a key is pressed, there are actually two different questions you could be asking. 'Is this key currently being held down right now?' - that's true for as long as you hold it, checked fresh every single frame. Or: 'did this key just get pressed, this exact frame, for the first time?' - that's only true for one single frame, no matter how long you hold it after.\""),
        p("\"For movement, back in Lesson 2, you wanted the 'currently held' version - hold the key, keep moving. But for firing a bullet, if you use that same 'currently held' check, holding the fire key down for one second at 60 frames a second fires sixty bullets in that one second. That's almost never what you want.\""),
        p("[Demo cue: same Godot project] \"So firing needs the other check - 'just pressed, this one frame.' In Godot that's is_action_just_pressed(). GameMaker has keyboard_check_pressed() alongside keyboard_check(). Unity's Keyboard.current.spaceKey.wasPressedThisFrame versus .isPressed. Same split, different names, in basically every engine.\""),
        p("\"Your handout today already uses the right one for each engine, so you won't have to guess - but now you'll understand WHY it's written that way instead of just copying it.\""),
        p("[Timing checkpoint: wrap by 0:50]", { italics: true, color: "888888" }),

        h2("[0:50-1:00] Today's Exercise & Handouts"),
        p("\"For lab today, open your own project from last week - not a new one. Everything today gets added on top of the character you already built.\""),
        p("\"The goal: pressing your fire key spawns a projectile at your character, it flies up the screen on its own, and it cleans itself up once it's gone. Level 1 is getting that working once. Level 2 is making sure it works cleanly every time you press it, with a cooldown so holding the key doesn't flood the screen. Level 3 is polish - a sound, a distinct look for the bullet, or tidying up how the code is organized.\""),
        p("[Hand out engine-specific handouts]"),
        p("\"One more thing before you start: if your character from Lesson 2 still moves in all 4 directions, that's totally fine to leave as-is. Our actual game only needs Left and Right, so it's also fine to trim it down if you'd rather match the real game exactly - your call, either way works.\""),
        p("\"Questions before we start lab? ... Let's go build something that shoots.\""),

        h2("Lab Time Notes"),
        bullet("The single most common bug this week: a held-down fire key spawning dozens of bullets per second. If a student's screen fills up with bullets instantly, this is almost always the cause - check whether they're using the \"held\" input check instead of the \"just pressed\" one."),
        bullet("Second most common: bullets that spawn but never move, or move but never disappear - in every engine, this usually means the movement/cleanup code ended up in the wrong event (e.g. a Create event instead of a Step event, or vice versa)."),
        bullet("Circulate by engine cluster, starting with whichever engine has the most students this week."),
        bullet("Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout, not given new unrelated work."),
        bullet("At 1:50, call time and do a 2-minute share-out: ask 2-3 students to show what they built, across different engines if possible."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_03_Teacher_Script.docx"]);
}
Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
