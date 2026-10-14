const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, WidthType, ShadingType, VerticalAlign
} = require("docx");
const brand = require("./brand.js");

function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 80 }, children: [new TextRun(text)] }); }
function numbered(text, level = 0) { return new Paragraph({ numbering: { reference: "numbers", level }, spacing: { after: 80 }, children: [new TextRun(text)] }); }

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
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
    headers: { default: brand.pageHeader("Lesson 4 - Godot Companion") },
    footers: { default: brand.pageFooter() },
    children: [
      brand.crestParagraph(1.3, { after: 160 }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "How to Fix the Main Scene", bold: true, size: 40, color: brand.NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Godot Engine - a one-time catch-up before this week's exercise", size: 24, color: brand.NAVY2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

      h1("Why You're Seeing This"),
      p("Lesson 4's Setup says to \"open your main scene and drag an instance of enemy.tscn into it.\" If you built Lessons 2 and 3 before this document existed, you don't have a main scene - and that's not something you missed. The Godot Lesson 2 Handout should have had you create one from the very start (the same role Unity's Scene, GameMaker's Room, or Scratch's Stage already play), and it didn't. That's been fixed for future classes, but our class built Lessons 2 and 3 on the old instructions - so before touching this week's Enemy, we're doing a short, one-time catch-up to bring your project in line with the corrected materials."),
      p("Nothing about how your game plays changes. Player still moves exactly the way you built it. This just gives it a proper home to live in, the same way Bullet and (this week) Enemy already have their own .tscn files."),

      h1("Step 1: Create the Main Scene"),
      numbered("Open your existing [YourName]Game project - the same one with player.tscn and bullet.tscn already in it from Lessons 2 and 3. Don't start a new project."),
      numbered("Scene > New Scene. Choose Node2D as the root type, rename it Main, and save it as main.tscn in the project's root folder."),
      numbered("Right-click main.tscn in the FileSystem dock and choose \"Set As Main Scene\" (or Project > Project Settings > Application > Run > Main Scene, and browse to it). This is what F5 (Run Project) will launch from now on."),
      numbered("Drag player.tscn from the FileSystem dock into main.tscn's scene tree. This adds Player as a child of Main."),
      numbered("Press F5 (Run Project) and confirm Player moves exactly like it always has. If it does, this step is done - Player's own script and behavior haven't changed at all, just where it lives."),

      h1("Step 2: About Your Bullets (Lesson 3)"),
      p("Your Player's script already spawns bullets with get_parent().add_child(b), after setting b.global_position = global_position - that's exactly right, and was written that way on purpose back in Lesson 3. Here's what actually changes underneath it:"),
      bullet("Before this fix: Player was itself the scene's root when you pressed F6, so get_parent() was quietly resolving to Godot's own internal runtime root - something that works, but that you can't see or select anywhere in the Scene tree."),
      bullet("After this fix: Player is a real child of Main, so get_parent() now resolves to Main instead - a real, visible node in your Scene tree. The bullet lands as Player's sibling under Main, which is exactly what was always intended."),
      bullet("No code changes are needed for this. Press F5 and fire a few shots to confirm bullets still spawn at the Player and fly up the screen exactly like before - if they do, you're done with this part too."),

      h1("Step 3: Continue with Lesson 4"),
      p("With main.tscn created and Player inside it, Lesson 4's Setup step 5 now works exactly as written: open main.tscn (the same scene you just built) and drag in an instance of enemy.tscn, positioning it near the top of the play area. Pick up the rest of the Lesson 4 Handout from there."),
      p("One thing to know before you position it: Godot's origin (0,0) is the top-left corner, and Y increases downward - not a center or bottom-left origin. \"Near the top\" means a small positive Y (something like 50-150), not a negative one or a value guessed from a different convention. Position Enemy (and Player, if you're adjusting it) directly in the 2D editor - drag it in the viewport, or set Transform > Position in the Inspector - rather than repositioning it from code in _ready(). That's the preferred approach here, not just a shortcut: it keeps what you see in the editor exactly matching your game's actual start position."),

      h1("Troubleshooting This Retrofit"),
      bullet("F5 (Run Project) shows a blank window, or the old \"no main scene\" prompt: confirm main.tscn was actually set as the project's Main Scene (Project Settings > Application > Run > Main Scene)."),
      bullet("The game looks different after the retrofit, or Player doesn't appear: confirm Player was dragged in as a child of Main - nested underneath it in the Scene tree, not left as its own separate, unconnected scene - and that you saved main.tscn afterward."),
      bullet("Bullets spawn in the wrong place, or stop appearing, after the retrofit: confirm bullet.tscn's spawn code in Player's script still uses global_position on both sides of the assignment (b.global_position = global_position), and that Player really is a child of Main now, not still floating on its own."),
      bullet("Enemy doesn't appear anywhere on screen, even though it shows up in the Scene tree (or the Remote tab while running): this is almost always Transform > Position, not the Texture - see the origin note above. Godot's top-left, Y-down origin makes it easy to place something off the top of the visible viewport with no error at all."),
      bullet("Still stuck: this is a one-time structural change, not a rewrite - if Player, Bullet, and Enemy all worked individually before, they'll work the same way once Main actually holds them. Flag it and we'll sort it out together rather than guessing further on your own."),

      h1("A Note for Next Semester"),
      p("This is a one-time catch-up for our class specifically. Anyone starting this course fresh from a corrected Lesson 2 Handout won't ever see this document - main.tscn gets created in Lesson 2's own Setup from day one, before Player even exists."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Lesson_04_Godot_How_to_Fix_the_Main_Scene.docx", buffer);
  console.log("done");
});
