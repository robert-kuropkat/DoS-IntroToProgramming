const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, WidthType, ShadingType, PageNumber, PageBreak
} = require("docx");
const brand = require("./brand.js");

function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function labelBullet(label, text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)] });
}

// Screenshot placeholder box - shaded, clearly marked "to be filled in"
function shotBox(filename, description) {
  return new Paragraph({
    spacing: { after: 200, before: 80 },
    indent: { left: 360 },
    shading: { fill: "FFF4D6", type: ShadingType.CLEAR },
    children: [
      new TextRun({ text: "SCREENSHOT NEEDED: ", bold: true, color: "8A6100" }),
      new TextRun({ text: filename, bold: true, font: "Consolas", color: "8A6100" }),
      new TextRun({ text: "", break: 1 }),
      new TextRun({ text: description, color: "6B5200" }),
    ],
  });
}

function instructorNoteBox(children) {
  return new Paragraph({
    spacing: { after: 120, before: 80 },
    shading: { fill: "FDECEC", type: ShadingType.CLEAR },
    children,
  });
}

function pageHeader(text) {
  return new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text, size: 16, color: "888888" })] })] });
}
function pageFooter() {
  return new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Page ", size: 16 }), new TextRun({ children: [PageNumber.CURRENT], size: 16 })] })] });
}

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]
};

const stylesBlock = {
  default: { document: { run: { font: "Arial", size: 22 } } },
  paragraphStyles: [
    { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 32, bold: true, font: "Arial", color: brand.NAVY },
      paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
    { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 26, bold: true, font: "Arial", color: brand.NAVY2 },
      paragraph: { spacing: { before: 260, after: 140 }, outlineLevel: 1 } },
    { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
      run: { size: 20, bold: true, font: "Arial", color: "444444" },
      paragraph: { spacing: { before: 140, after: 60 }, outlineLevel: 2 } },
  ]
};

// ---------- Per-engine content ----------
const engines = [
  {
    name: "Scratch",
    style: "Block",
    blurb: "Press the space bar and your sprite flashes a different color for a moment - the simplest possible \"press a key, something happens\" reaction, built from two snapped-together puzzle-piece blocks.",
    shots: [
      ["showcase_scratch_setup.png", "Scratch editor open, placeholder sprite (e.g. the blue circle character) added and selected in the Sprite panel."],
      ["showcase_scratch_code.png", "The two connected blocks in the code area: \"when [space] key pressed\" on top of \"change color effect by 25.\""],
      ["showcase_scratch_result.png", "The stage right after pressing space - sprite mid color-change."],
    ],
  },
  {
    name: "Godot Engine",
    style: "Scripting",
    blurb: "Hold the Right arrow key and the character slides across the screen - a few lines of GDScript checked every frame.",
    shots: [
      ["showcase_godot_setup.png", "Godot editor with a Sprite2D node in the scene tree, placeholder image assigned as its Texture in the Inspector."],
      ["showcase_godot_code.png", "The attached script showing the _process(delta) function with the \"if Input.is_action_pressed\" check."],
      ["showcase_godot_result.png", "The running scene (F6) with the sprite partway across the screen after holding Right."],
    ],
  },
  {
    name: "Unity",
    style: "Scripting",
    blurb: "Press Play, hold an arrow key or A/D, and a placeholder GameObject slides across the Scene view - Unity's C# Update() loop in action.",
    shots: [
      ["showcase_unity_setup.png", "Unity Hierarchy panel showing the Player GameObject with a Sprite Renderer component, placeholder image assigned."],
      ["showcase_unity_code.png", "The PlayerMove script open in the editor, showing the Update() method checking Keyboard.current.rightArrowKey.isPressed."],
      ["showcase_unity_result.png", "Play mode running, Player object visibly moved right in the Game view."],
    ],
  },
  {
    name: "GameMaker",
    style: "Visual or Scripting (shown here in GML text)",
    blurb: "Hold the Right arrow key in a running room and the object steps to the right, one Step-event line of GML at a time - or the same logic built from GML Visual blocks instead.",
    shots: [
      ["showcase_gamemaker_setup.png", "GameMaker workspace with obj_player in the Room Editor, spr_player (placeholder image) assigned as its sprite."],
      ["showcase_gamemaker_code.png", "The Step event showing \"if (keyboard_check(vk_right)) x += 4;\" (or the equivalent GML Visual blocks)."],
      ["showcase_gamemaker_result.png", "The running game (F5) with the object visibly moved right after holding the key."],
    ],
  },
  {
    name: "GDevelop",
    style: "Visual",
    blurb: "Click Preview and all 4 arrow keys already move the Player object around - GDevelop's built-in \"Top-down movement\" behavior does the work with zero hand-written events.",
    shots: [
      ["showcase_gdevelop_setup.png", "GDevelop scene editor with the Player object selected and the Behaviors tab open, \"Top-down movement\" behavior listed as added."],
      ["showcase_gdevelop_code.png", "The (mostly empty) event sheet, showing that no manual events are needed for basic movement - or the manual \"Key pressed\" event if that method was used instead."],
      ["showcase_gdevelop_result.png", "Preview mode running with the Player object visibly moved after arrow key presses."],
    ],
  },
  {
    name: "TIC-80",
    style: "Scripting",
    blurb: "Hold the Right arrow key in this tiny fantasy console and a hand-drawn 8x8 sprite steps across the 240x136 screen - six lines of Lua, typed directly into the built-in code editor.",
    shots: [
      ["showcase_tic80_setup.png", "TIC-80's Sprite Editor with a simple character drawn on sprite index 1."],
      ["showcase_tic80_code.png", "The built-in code editor showing the short TIC() function with \"if btn(3) then x=x+1 end.\""],
      ["showcase_tic80_result.png", "The running cart (after typing \"run\") with the sprite moved right on screen."],
    ],
  },
  {
    name: "001 Game Creator (Windows only)",
    style: "Visual",
    blurb: "Press the Right arrow key and \"main,\" the default player Actor, steps to the right - built entirely from an Input Trigger and a Position Actor event, no code typed at all.",
    shots: [
      ["showcase_001gc_setup.png", "001 Game Creator's default Map with the \"main\" Actor selected, placeholder sprite optionally assigned in its Body settings."],
      ["showcase_001gc_code.png", "The Input Triggers window showing the Right Arrow \"Pressed\" trigger, and its attached Script Editor graph with a single Position Actor event."],
      ["showcase_001gc_result.png", "Play/Test mode with \"main\" visibly moved right after the key press."],
    ],
  },
  {
    name: "Unreal Engine (optional advanced track)",
    style: "Visual (Blueprints)",
    blurb: "Press Play, tap the Right arrow key, and a placeholder Cube slides across the level - an Enhanced Input Action Event node (IA_MoveRight, mapped to the Right arrow in an Input Mapping Context) wired to an Add Actor World Offset node in the Level Blueprint. The heaviest engine on our list, so this one's for students with capable hardware.",
    shots: [
      ["showcase_unreal_setup.png", "The level viewport with a Cube Actor placed (from the Place Actors panel), selected in the World Outliner."],
      ["showcase_unreal_code.png", "The Level Blueprint graph showing the Enhanced Input Action Event (IA_MoveRight) node wired into an Add Actor World Offset node."],
      ["showcase_unreal_result.png", "Play-in-editor with the Cube visibly moved after the key press."],
    ],
  },
];

function engineSection(e) {
  return [
    h2(e.name),
    labelBullet("Coding style", e.style),
    p(e.blurb),
    ...e.shots.flatMap(([filename, desc]) => [shotBox(filename, desc)]),
    new Paragraph({ spacing: { after: 200 }, children: [] }),
  ];
}

const doc = new Document({
  numbering: numberingConfig,
  styles: stylesBlock,
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: brand.pageHeader("Engine Showcase") },
    footers: { default: brand.pageFooter() },
    children: [
      brand.crestParagraph(1.3, { after: 160 }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Engine Showcase", bold: true, size: 40, color: brand.NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "See All 8 Engines in Action - Before You Choose", size: 26, color: brand.NAVY2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "A preview for families, ahead of Lesson 1", italics: true, size: 22, color: "666666" })] }),

      instructorNoteBox([
        new TextRun({ text: "INSTRUCTOR NOTE - delete this box before sending to families. ", bold: true, color: "B00020" }),
        new TextRun({ text: "This document has 24 screenshot placeholders (3 per engine x 8 engines), each with an exact filename and a description of what to capture. Whenever convenient - pre-building the Lesson 1 Live Engine Tour projects is a natural time, since you'll have all 8 engines open anyway - capture each shot and drop it in the Screenshots_Inbox folder using the exact filename shown in its placeholder box. No need to crop or clean anything up first. Once a batch is in the inbox, they get annotated (circles/arrows/labels as needed) and dropped into the matching placeholder below, replacing the box.", color: "6B5200" }),
      ]),
      new Paragraph({ spacing: { after: 200 }, children: [] }),

      h1("How to Use This Document"),
      p("This is a preview, not a tutorial. Every engine below is doing the exact same tiny thing: pressing a key causes a character to react. That's the very first hands-on exercise of the class (Lesson 2's \"Key Press Reactor\"), and it's a genuinely fair way to compare engines side by side, since it's the same problem solved eight different ways."),
      p("For each engine you'll see: what you're looking at, a shot of the project set up and ready, a shot of the actual code/blocks/nodes that make it work, and a shot of the result after pressing a key. Skim for whichever engines look interesting, then check the Engine & Platform Guide for the fuller picture on cost, hardware needs, and install steps before making a final choice."),
      p("Coding styles are tagged for each engine - Block (drag-and-snap puzzle pieces), Visual (connecting events/nodes without typing), or Scripting (typed text code). You are not locked into one style for the whole year: several of these engines let you use more than one style in the same project, and the class deliberately rotates through all three all year, so trying one now doesn't close any doors.", { italics: true, color: "555555" }),

      new Paragraph({ children: [new PageBreak()] }),
      h1("The Engines"),
      ...engines.flatMap(engineSection),

      h1("Which One Should I Try First?"),
      p("If you're not sure yet, the Engine & Platform Guide's \"Which Engine Fits My Laptop?\" table is the fastest filter - hardware rules out options quickly, and everything left after that is a matter of taste. A few rules of thumb:"),
      bullet("Never coded before, or want the fastest path to something playable? Start with Scratch."),
      bullet("Oldest laptop in the house, or want a zero-install option? GDevelop runs entirely in a browser."),
      bullet("Want a real scripting language without a heavy install? Godot is the class's recommended default."),
      bullet("Aiming at an industry-recognized skill and certification path? Unity."),
      bullet("On Windows with a very old or low-power laptop? 001 Game Creator is the lightest option on this list."),
      bullet("Have a newer, powerful laptop and want the most advanced 3D toolset? Unreal Engine (optional track)."),
      p("Whatever you pick, you can change your mind later - every lesson keeps working examples across all 8 engines all year."),
    ],
  }],
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Engine_Showcase.docx", buffer);
  console.log("done");
});
