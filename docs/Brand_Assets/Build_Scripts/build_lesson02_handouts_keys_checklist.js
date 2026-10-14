const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  HeadingLevel, BorderStyle, WidthType, ShadingType, PageNumber, VerticalAlign
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
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function numbered(text, level = 0) { return new Paragraph({ numbering: { reference: "numbers", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function checkbox(text) { return new Paragraph({ numbering: { reference: "checks", level: 0 }, spacing: { after: 80 }, children: [new TextRun(text)] }); }
function labelBullet(label, text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)] });
}
function code(lines) {
  const children = [];
  lines.forEach((line, i) => {
    children.push(new TextRun({ text: line, font: "Consolas", size: 20 }));
    if (i < lines.length - 1) children.push(new TextRun({ text: "", break: 1 }));
  });
  return new Paragraph({
    spacing: { after: 160, before: 80 }, indent: { left: 360 },
    shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
    children
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
// UPDATED PREP CHECKLIST (adds Self-Contained Lesson Assets section)
// =========================================================================
{
  const doc = new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 2 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 2 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before Thursday, August 20, 2026", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("Self-Contained Lesson Assets"),
        p("Per Lesson Production Standards Rule 3, nothing this lesson needs should require a student to search for or download something mid-class. Everything below is already prepared and included in this lesson's Assets folder - no download required."),
        checkbox("Placeholder sprite pack already generated (player_circle_blue.png, player_circle_green.png, player_circle_orange.png, player_arrow_white.png, player_walk_spritesheet.png) - copy into each engine's project as needed."),
        checkbox("Placeholder sound effect already generated (sfx_blip.wav) for students who want to try the Level 3 \"add a sound\" idea in Scratch, GameMaker, GDevelop, TIC-80, 001 Game Creator, or Unreal."),
        checkbox("Copy the Assets folder onto the shared classroom drive / USB backups alongside the handouts so it travels with the rest of the lesson's materials."),
        checkbox("Confirm Unreal's Blank template still ships with basic Shape actors (Cube/Sphere) in the Place Actors panel for the optional advanced-track exercise - no separate download needed there either."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop's Hello World still works for all 8 engines (ref: Engine Installation & Hello World Guide) - re-run any that haven't been touched in a few weeks."),
        checkbox("Pre-build the full Scratch Key Press Reactor exercise (all 3 levels) on the demo laptop so you have a finished example ready to show."),
        checkbox("Godot: confirm Project Settings > Input Map still has the default ui_up/ui_down/ui_left/ui_right actions unmodified, and practice the Set As Main Scene step on main.tscn once so it's a familiar click during class."),
        checkbox("Unity: confirm Edit > Project Settings > Player > Other Settings > Active Input Handling is set to \"Both\" (the default for a fresh Unity 6 project) or \"Input System Package (New)\" - the Lesson 2 code uses Keyboard.current, which needs one of those two, not \"Input Manager (Old)\" alone."),
        checkbox("GameMaker: quick-test that keyboard_check(vk_right) etc. work in a throwaway object before class."),
        checkbox("GDevelop: confirm the \"Top-down movement\" behavior is available when adding a behavior to an object in your installed version."),
        checkbox("TIC-80: confirm btn(0-3) map to Up/Down/Left/Right as expected on your keyboard."),
        checkbox("001 Game Creator (Windows demo laptop): confirm the Input Triggers window and Position Actor / Take Movement in Direction events are accessible."),
        checkbox("If any advanced-track students are on Unreal this week, confirm the editor still opens and Blueprints compile cleanly, and practice the Auto Receive Input step on BP_Player once so it's a familiar click during class."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Have Lesson 1 sign-in / install notes handy to know who still needs setup help at the start of lab."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies, Assets folder (on USB as backup), any offline installers still needed."),

        h2("Known Gotchas This Week"),
        bullet("GameMaker: movement code goes in the Step event, not Draw - students may default to Draw out of habit from the Lesson 1 Hello World exercise."),
        bullet("Godot: if a student's sprite doesn't move, the most common cause is the script being attached to the wrong node, or the scene not being saved before running. A second, separate snag this week: forgetting to drag Player into main.tscn, or forgetting to set main.tscn as the project's Main Scene - F6 on player.tscn alone still \"works\" for movement testing either way, which can hide this mistake until Lesson 3 needs Player and Bullet sharing one running scene."),
        bullet("001 Game Creator: dialog field layouts can vary slightly by version - if something looks different from the handout, the Script Editor and Actor wiki pages have the current layout."),
        bullet("Unreal: this is the heaviest engine on our list - if a student's laptop is struggling, it's fine for them to observe this week's demo and build the exercise in a lighter engine instead."),
        bullet("Unreal: the single most common \"nothing happens, no error\" cause this week is forgetting to set Auto Receive Input to Player 0 in BP_Player's Class Defaults - since BP_Player is a plain Actor rather than a possessed Pawn, it never receives keyboard input without this step, and every Enhanced Input Action Event will silently just never fire."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_02_Prep_Checklist.docx"]);
}

// =========================================================================
// UPDATED STUDENT HANDOUTS (adds asset tip)
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 2: Key Press Reactor", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: sequences, events, and the game loop. Your job is to build something that reacts when you press a key - and this is also the start of your Semester 1 project, so save it somewhere safe and give it a name you'll recognize."),
    p("Tip: this lesson's Assets folder already has a ready-to-use placeholder character image (and an optional sound effect for Level 3) - you don't need to find or make your own art today.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("One key press causes one visible reaction. This is a complete finish for today - reaching this is enough to move on to next week.", { italics: true }),
    ...data.level1.map(s => numbered(s)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  if (data.altPath) {
    children.push(h2("Alternate Path: " + data.altPath.title));
    children.push(p(data.altPath.intro, { italics: true, color: "555555" }));
    if (data.altPath.warning) children.push(p(data.altPath.warning, { bold: true, color: "B00020" }));
    children.push(...data.altPath.steps.map(s => numbered(s)));
    if (data.altPath.code) children.push(code(data.altPath.code));
    if (data.altPath.note) children.push(p(data.altPath.note, { italics: true, color: "555555" }));
  }
  children.push(
    h2("Level 2 - Complete (Alpha/Beta)"),
    p("All 4 arrow keys move your character in 4 directions.", { italics: true }),
    ...data.level2.map(s => bullet(s)),
    h2("Level 3 - Refined (Ready to Publish) - pick any that interest you"),
    ...data.level3.map(s => bullet(s)),
    h2("If Something's Wrong"),
    ...data.trouble.map(s => bullet(s)),
    h2("Want More?"),
    p(data.resource),
  );

  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 2 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open Scratch (browser or desktop). Start a new project if one isn't already open.",
      "Save your project as [YourName]Game - this is the start of your Semester 1 project.",
      "Keep the default cat sprite, or import one of the placeholder images from this lesson's Assets folder (Upload Sprite > choose player_circle_blue.png or similar).",
    ],
    level1: [
      "Click the Events category (yellow blocks).",
      "Drag \"when [right arrow] key pressed\" into the code area.",
      "Click the Motion category, drag \"change x by 10\" underneath it.",
      "Click the stage, then press the right arrow key to test.",
    ],
    level2: ["Repeat with 4 separate scripts - \"when [up/down/left/right] arrow key pressed\" - each using a Motion block (\"change y by 10\", \"change y by -10\", \"change x by -10\", \"change x by 10\") so your sprite moves in all 4 directions."],
    level3: [
      "Animate: make your sprite \"blink\" while moving and stay solid when idle. This needs its own \"when green flag clicked\" script with a forever loop, checking the Sensing category's \"key [ ] pressed?\" reporter for all 4 arrow keys - that's the one part of this exercise that genuinely needs a loop, because Scratch has no \"when key released\" block to trigger off of.",
      "Sound: play sfx_blip.wav once each time you press a movement key - silent the rest of the time, since a \"when key pressed\" script only runs at the instant of the press. Upload it from this lesson's Assets folder (Choose a Sound > Upload Sound) and add a \"start sound\" block right inside each of your 4 movement scripts - no loop or timer needed.",
      "Recolor: use a different costume for each direction, switching back to your default costume when idle - swap the picture rather than tint it. Upload the other placeholder circles from this lesson's Assets folder as extra Costumes on your sprite (Costumes tab > Upload Costume): red, blue, green, and yellow for the 4 directions, plus orange as your dedicated idle costume even if you started with a different one in Setup. Switch costume the moment you press a key (inside that key's own script), and switch back to orange in the same green-flag loop you built for the blink - since \"returning to idle\" needs to know when no key is held anymore, same as blink does.",
      "Bounce: add \"if on edge, bounce\" (Motion) right after the movement block inside each of your 4 key-press scripts, so your sprite bounces off the stage edge instead of walking off it. Heads up: this can flip your sprite left-right when it bounces, depending on its rotation style - that's a known Scratch quirk, not a bug, and it's fine to leave as-is for now. (Precisely stopping at the edge with no flip needs the position-checking technique from a later lesson.)",
      "Structure: keep this as 5 separate scripts, not one - \"when green flag clicked\" for the blink/idle-reset loop, plus one \"when key pressed\" script per arrow key that handles that direction's movement, sound, color, and bounce together. Each key-press script IS the \"if this key, do this\" logic - you don't need a separate \"if key pressed?\" check inside it. See the Key for the fully worked version.",
    ],
    trouble: [
      "Nothing happens when you press a key: make sure the blocks are actually snapped together (a white highlight shows when they connect), and click the stage first so Scratch is \"listening.\"",
      "Sprite only moves once even holding the key: that's normal - each press triggers one run. Continuous movement while held is a Level 3 stretch goal using a \"forever\" loop with a \"key [ ] pressed?\" check inside an if-block.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open Godot, click New Project, name it [YourName]Game, choose a folder, Create & Edit.",
      "Create a new scene, choose Node2D as the root type, rename it Main, and save it as main.tscn in the project's root folder - this is the scene that actually runs, the same role Unity's Scene, GameMaker's Room, or Scratch's Stage already play. Every other scene this course builds (Player this week, then Bullet, Enemy, and more in later lessons) gets instanced into this one.",
      "Right-click main.tscn in the FileSystem dock and choose \"Set As Main Scene\" (or Project > Project Settings > Application > Run > Main Scene and browse to it) - this is what F5 (Run Project) launches.",
      "Create a second new scene: a Sprite2D root node, rename it Player, and assign one of the placeholder images from this lesson's Assets folder (e.g. player_circle_blue.png) as its Texture.",
      "Save this scene as player.tscn.",
      "Back in main.tscn, drag player.tscn from the FileSystem dock into the scene tree to add Player as a child of Main.",
    ],
    level1: [
      "Select the Player node (open player.tscn directly, or select it inside main.tscn) and click the script icon in the Inspector, keep defaults, click Create.",
      "Inside the _process(delta) function, add the code below.",
      "Press F6 (Run Current Scene) with player.tscn open and hold the Right arrow key to test - F6 always runs whichever scene you currently have open, regardless of the project's Main Scene setting, so this still works for testing Player by itself.",
    ],
    level1Code: ["func _process(delta):", "    if Input.is_action_pressed(\"ui_right\"):", "        position.x += 200 * delta"],
    level2: ["Add the other 3 directions using ui_left, ui_up, and ui_down (position.x -= ..., position.y -= ..., position.y += ...)."],
    level3: [
      "Animate: make the sprite blink while moving, and stay solid (visible) when idle.",
      "Sound: play sfx_blip.wav in a beep-beep pattern while moving, silent when idle - add it via an AudioStreamPlayer2D child node.",
      "Recolor: swap in a different placeholder image for each direction (texture), rather than tinting one - assign the Sprite2D's texture property to a different preloaded image per direction, and back to your default/orange image when idle.",
      "Clamp: keep the sprite on screen using clamp() on position.",
      "Structure: keep all four of the above together in the same _process(delta) function, computed once per frame - not scattered across separate functions.",
    ],
    trouble: [
      "Nothing moves: confirm the script is attached to the Sprite2D that's actually in your running scene, and that you saved the scene before pressing F6.",
      "Red text in the Output panel points to the exact line with a typo - fix the first error listed first.",
      "F5 (Run Project) shows a blank window, or the old empty-project screen: confirm main.tscn was actually set as the project's Main Scene (Project Settings > Application > Run > Main Scene), and that Player was dragged into it as a child rather than left as its own separate, unconnected scene.",
    ],
    resource: "docs.godotengine.org, Student Resource List > Godot",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Create a new 2D project named [YourName]Game.",
      "In the Hierarchy, right-click > Create Empty, name it Player, add a Sprite Renderer component and drag one of the placeholder images from this lesson's Assets folder (e.g. player_circle_blue.png) into its Sprite field.",
    ],
    level1: [
      "With Player selected, Add Component > New Script, name it PlayerMove, Create and Add.",
      "Double-click the script to open it. Add using UnityEngine.InputSystem; as a new line right after the existing using UnityEngine; at the top - this is what makes Keyboard.current available.",
      "Replace the Update() method body with the code below.",
      "Save, return to Unity, press Play, and use the arrow keys.",
    ],
    level1Code: ["void Update() {", "    float h = 0f;", "    if (Keyboard.current.rightArrowKey.isPressed) h = 1f;", "    if (Keyboard.current.leftArrowKey.isPressed) h = -1f;", "    transform.Translate(new Vector3(h, 0, 0) * 3f * Time.deltaTime);", "}"],
    level2: ["Add vertical movement too: same if-pattern with Keyboard.current.upArrowKey and Keyboard.current.downArrowKey building a float v, then include it in the Vector3 as (h, v, 0)."],
    level3: [
      "Animate: make the sprite blink while moving, and stay solid (visible) when idle - toggle spriteRenderer.enabled based on a timer.",
      "Sound: play sfx_blip.wav in a beep-beep pattern while moving, silent when idle - add an AudioSource, check audioSrc.isPlaying before calling Play() again, and call Stop() when idle.",
      "Recolor: swap in a different placeholder image for each direction (spriteRenderer.sprite), rather than tinting one - assign a different Sprite per direction in the Inspector, and back to your default/orange sprite when idle.",
      "Clamp: keep the Player's position on screen using Mathf.Clamp().",
      "Structure: keep all four of the above together in the same Update() method - not scattered across separate functions.",
    ],
    trouble: [
      "Script errors show in red in the Console - fix the top one first; others often resolve on their own.",
      "If nothing moves, confirm the Player GameObject wasn't deleted/duplicated and the script component is enabled (checkbox visible in the Inspector).",
      "\"The type or namespace name 'Keyboard' could not be found\": confirm using UnityEngine.InputSystem; was added at the top of the script, not just using UnityEngine;.",
      "NullReferenceException on Keyboard.current, or nothing responds to key presses at all: check Edit > Project Settings > Player > Other Settings > Active Input Handling - a fresh Unity project defaults to \"Both,\" which supports this code, but if it's been changed to \"Input Manager (Old)\" only, switch it to \"Both\" or \"Input System Package (New)\" and let Unity restart the Editor when prompted.",
      "Visual Scripting: no Input System nodes show up when you search the fuzzy finder - Edit > Project Settings > Visual Scripting > Regenerate Nodes hasn't been run yet (one-time setup, not something you repeat every lesson).",
      "Visual Scripting: Move doesn't appear in the On Input System Event Vector 2 node's Input Action dropdown - the Player Input component's Actions field is still empty; assign it the PlayerControls asset.",
    ],
    resource: "learn.unity.com, Student Resource List > Unity",
    altPath: {
      title: "Visual Scripting (node graph, no typing)",
      intro: "Unity's built-in Visual Scripting package lets you build the same movement logic as a node graph instead of C# code, using the same new Input System as the code path above - not the legacy Input Manager.",
      steps: [
        "With Player selected, Add Component > Input > Player Input (Unity installs the Input System package on first use if prompted). In its Actions field, click Create Actions and save the new asset as PlayerControls.",
        "Double-click PlayerControls to open the Input Actions editor. Under the Player action map, select the existing Move action and edit its 2D Vector Composite bindings so Up/Down/Left/Right point at the Up Arrow, Down Arrow, Left Arrow, and Right Arrow keys (matching every other engine in this lesson), then click Save Asset.",
        "In Edit > Project Settings > Visual Scripting, click Regenerate Nodes (only needed once per project) so the Input System event nodes become available in the fuzzy finder.",
        "In Edit > Project Settings > Input System Package, click Create Settings Asset if you haven't already (also only needed once per project).",
        "With Player selected, Add Component > Visual Scripting > Script Machine (Unity installs the Visual Scripting package on first use if prompted), click the graph field to create a new Script Graph asset, name it PlayerMoveVisual, and click Edit Graph to open it.",
        "Right-click the canvas, search for and add an \"On Input System Event Vector 2\" node. Leave Target as This, set Input Action to Move, and set Input Action Change Type to On Hold.",
        "Add a \"Vector 2 > Get X\" node and wire the Vector 2 Value output into it.",
        "Add a \"Vector 3 > Construct Vector 3\" node and wire Get X's output into its X input.",
        "Add a \"Transform > Translate\" node; wire the On Input System Event Vector 2 node's Trigger output into it, and wire the constructed Vector 3 into its Translation input.",
        "Save the graph (Ctrl+S), press Play, and test with the arrow keys - same result as the code path.",
      ],
      note: "Level 2: add a \"Vector 2 > Get Y\" node wired to the same On Input System Event Vector 2 node's Vector 2 Value output, and wire it into the Vector 3's Y input alongside Get X into X - the Move action already carries both axes, you're just choosing to use more of its output. Level 3's five goals (blink, beep-beep sound, recolor, clamp, and grouping it all under one Branch node keyed off whether you're moving) are all reachable with more nodes from the Sprite Renderer, Audio, and Math categories - same ideas as the code path, just wired instead of typed.",
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: [
      "Create a new blank project named [YourName]Game.",
      "Create a sprite spr_player by importing one of the placeholder images from this lesson's Assets folder (e.g. player_circle_blue.png).",
      "Create an object obj_player and assign spr_player to it.",
    ],
    level1: [
      "Add a Step event to obj_player.",
      "Add the code below to the Step event.",
      "Drag obj_player into your room, then press F5 (Run) to test with the Right arrow key.",
    ],
    level1Code: ["if (keyboard_check(vk_right)) x += 4;"],
    level2: ["Add the other 3 directions using vk_left, vk_up, and vk_down."],
    level3: [
      "Animate: make the sprite blink while moving, and stay solid (visible) when idle - toggle the built-in visible variable using a counter incremented in the Step event.",
      "Sound: play sfx_blip.wav in a beep-beep pattern while moving, silent when idle - check audio_is_playing() before calling audio_play_sound() again, and audio_stop_sound() when idle.",
      "Recolor: swap in a different sprite for each direction (sprite_index), rather than tinting one - import the other 4 placeholder circles as additional sprites, and set sprite_index to the matching one per direction, back to your default/orange sprite when idle.",
      "Clamp: keep x/y within room_width and room_height using clamp().",
      "Structure: keep all four of the above together in the same Step event - not scattered across separate events.",
    ],
    trouble: [
      "Nothing moves: confirm the code is in the Step event, not Draw or Create - Step runs every frame, Draw only draws.",
      "Confirm the obj_player instance is actually placed in the room you're running (check the Room Editor's instance list).",
    ],
    resource: "gamemaker.io/en/tutorials, Student Resource List > GameMaker",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "GameMaker lets you build the same Step event using drag-and-drop blocks instead of typed code. Pick whichever path feels right for you - both are graded the same.",
      steps: [
        "Add a Step event to obj_player (same as the code path above).",
        "In the Step event's editor, click the Code/Visual toggle in the top-right corner of the panel and switch to GML Visual.",
        "From the action palette's Mouse and Keyboard category, drag an \"If Key Down\" block into the event and set its Key dropdown to Right.",
        "From the Movement category, drag \"Jump To Point\" inside the If block. Check the Relative box for both X and Y, then set X to 4 and Y to 0 - Jump To Point is the dedicated block for moving relative to your current position, so prefer it over the generic Assign Variable block whenever you're just changing x/y.",
        "Drag obj_player into your room, then press F5 (Run) to test with the Right arrow key - same result as the code path.",
      ],
      note: "Level 2: add one more \"If Key Down\" + \"Jump To Point\" (Relative) pair per direction - X -4/Y 0 for Left, X 0/Y -4 for Up, X 0/Y 4 for Down. Level 3: reach for the most specific, well-named block GameMaker offers before falling back to Assign Variable - \"Set Sprite\" (Instance category) for the recolor goal instead of Assign Variable on sprite_index, and \"If Audio Is Playing\" -> Not -> \"Play Audio\" (Audio category) as the retrigger guard for the sound goal, with \"Stop Audio\" in the idle branch. Assign Variable is still the right tool for anything without a dedicated block, like the custom blink_timer or moving variables - same ideas as the code path either way, just dragged instead of typed.",
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Create a new blank project named [YourName]Game.",
      "Add a Sprite object using one of the placeholder images from this lesson's Assets folder (e.g. player_circle_blue.png), name it Player.",
    ],
    level1: [
      "Select Player, click \"Add a behavior,\" choose \"Top-down movement,\" click Apply.",
      "Click the green Preview button and test with the arrow keys.",
      "(Alternative manual method: add an event with condition \"Key pressed [Right]\" and action \"Change the object's X position.\")",
    ],
    level2: ["With the Top-down movement behavior, all 4 arrow keys already work - confirm by testing all 4 directions in Preview. If you used the manual method for Level 1, add the other 3 direction events now."],
    level3: [
      "Animate: make Player blink while moving, and stay solid (visible) when idle. Flash visibility (blink) isn't a standalone action sitting in a general list - it belongs to a behavior, the same way Angle of movement (used in the Recolor goal below) belongs to the Top-down movement behavior. Select Player, open its Behaviors tab, click Add a behavior, search \"Flash,\" and Apply - only once that's attached will \"Flash visibility (blink)\" (start) and \"Stop flashing visibility (blink)\" (stop) show up when you search Player's actions.",
      "Sound: play sfx_blip.wav in a beep-beep pattern while moving, silent when idle - use \"Play a sound\" together with a scene timer so it retriggers instead of overlapping itself every frame.",
      "Recolor: give Player a different Animation per direction rather than tinting one. Add the other 4 placeholder circles as extra Animations on the Player object (Edit object > Animations tab > Add an animation), naming them Red/Blue/Green/Yellow, plus your default one renamed Orange. The thing you're building here is 5 separate Events, not conditions/actions on their own - each Event is one row that bundles together its own Condition(s) and Action(s), and an Event's Action(s) only run when all of that SAME Event's Condition(s) are true. There's no separate step to \"connect\" a condition to an action; they're never independent pieces to begin with, only ever parts of the one Event that owns them. Build the first one: click Add new event, then, inside that Event, click its own Add condition (small text, top-left of the row), pick Player, then look for a separate Top-down movement category in the list (not Player's own conditions) and choose Angle of movement - set it to 0 degrees with a Tolerance of about 45. Still inside that same Event, click its Add action, pick Player again, and in the action search list look for Animation (by name) under Animations and images - it's labeled that way in the picker, but displays as \"Change the animation of Player...\" once you've placed it - set it to \"Red\". Click Add new event again and repeat for the other 3 directions (90/180/270 degrees for Down/Left/Up), plus a fifth new Event using an \"Is moving\" = False condition instead, with its action set to \"Orange\" for idle.",
      "Clamp: keep Player on screen by adding the Stay On Screen (2D) behavior - zero events required.",
      "Structure: keep all four of the above together in one labeled event group - not scattered across separate, unlabeled events.",
    ],
    trouble: [
      "Preview doesn't launch: check the log panel at the bottom of the window for a red error.",
      "Object doesn't move: confirm the behavior was actually added (check the object's Behaviors tab), and that you're testing in Preview, not the static scene editor.",
      "Recolor conditions/actions ended up in the wrong Event, or all crammed into one: no need to delete and rebuild - click Add new event to create an empty one, then drag and drop the misplaced Condition(s)/Action(s) into it (or out of it) to sort them into the right Event.",
      "Can't find Flash visibility (blink), Angle of movement, or another action/condition that should exist: some actions and conditions belong to a behavior rather than to the object itself, so they only appear once that behavior has been added - check the object's Behaviors tab first (Add a behavior), then search the action/condition list again.",
    ],
    resource: "gdevelop.io/tutorials, Student Resource List > GDevelop",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: [
      "Launch TIC-80.",
      "At the console, type new lua and press Enter to start a fresh cart.",
      "Open the Sprite Editor (paintbrush icon) and draw something simple on sprite index 1 - this will be your character (feel free to use one of the images in this lesson's Assets folder as a reference to copy, since TIC-80 draws sprites in its own tiny built-in editor rather than importing files directly).",
    ],
    level1: [
      "Press Esc to reach the code editor.",
      "Type the code below, replacing any default code.",
      "Press Esc, type run, press Enter, then hold the Right arrow key to test.",
    ],
    level1Code: ["x=60 y=60", "function TIC()", " cls(0)", " if btn(3) then x=x+1 end", " spr(1,x,y,0)", "end"],
    level2: ["Add the other 3 directions using btn(0) for up, btn(1) for down, and btn(2) for left (adjust x or y accordingly)."],
    level3: [
      "Animate: make your sprite \"blink\" while moving and stay static when idle. Alternate the scale argument of spr() between 0 and 1 using a frame counter (a scale of 0 draws nothing at all) - this also means switching to spr()'s full argument list, including w/h for a bigger, 2x2-tile composite sprite if you want one.",
      "Sound: add a beep-beep while moving, silent when idle. Open the SFX Editor via the speaker icon (or press F4) - the musical note icon opens the separate Music Editor instead, which is not what you want here. Alternate between two short SFX slots for a beep-beep while a direction key is held, and call sfx(-1) the instant the key is released - a plain sfx() call has no built-in \"stop,\" so you have to turn it off yourself.",
      "Recolor: use a different color for each direction, returning to the original color when idle. Temporarily swap one of your sprite's palette colors using poke4() - open the Sprite Editor (F2, or the monster-icon button) to see the color palette strip and figure out which color index is worth changing; there's no pal()/palette() function in TIC-80 like some other fantasy consoles have. Unlike every other engine in this lesson (which swap in a whole separate colored image for this goal, since their tint functions only look clean on white art), TIC-80's poke4() genuinely recolors in place with no such catch - simpler here, not a workaround.",
      "Clamp movement so the sprite can't leave the 240x136 screen.",
      "Structure: once you're combining more than one of these, fold the new logic into the same if/elseif/else block from Level 2 rather than writing a separate block for each - see the Key for a worked example of how movement, blink, recolor, and sound can all live together in one place.",
    ],
    trouble: [
      "Nothing moves: double-check the function is named exactly TIC() (capitalization matters), and that you pressed run after editing.",
      "Sprite looks blank or wrong: make sure you drew something on sprite index 1 in the Sprite Editor first.",
    ],
    resource: "tic80.com, Student Resource List > TIC-80",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Create a new project named [YourName]Game from the Blank Template - a Warning will note this template is for experienced users with no pre-built input, menus, or actors; that's fine, click Continue. Blank starts with no Map and nothing placed on screen, but it still includes one default Player/Party roster entry named \"Hero\" - it just has no appearance or map location yet. The next four steps get Hero on screen.",
      "Create a Map: click the Maps tab in the small tabbed panel at the bottom-left, then click New. In the Map's Properties, set General > View to Top (top-down, matching our other engines), then give the Map a name (e.g. Level1).",
      "Give Hero an appearance: go to Resources > Sprites > Create New > No Template, then under the Graphics section use \"Add Animation > Image Files\" to import one of the placeholder images from this lesson's Assets folder - Resources > Import Resources does NOT accept plain image files, only pre-formatted asset sheets, so the Sprite Editor is the correct path for a simple placeholder PNG. In the Sprite Editor, check the \"Face X\" box just left of the preview image, then click the \"Top\" tab above the preview to confirm it looks right from a top-down view.",
      "Attach the sprite: go to Resources > Players and Party Members > Hero > Edit Actor > Sprite section > Body, and choose the Sprite you just created.",
      "Set the starting point: open Game Settings from the top menu, set Starting Map to the Map from step 2, then set Starting Location by dragging Hero's icon onto the map in the popup window.",
      "Press F5 to test - Hero's sprite should appear at the spot you set. Once you see it, you're ready for Level 1 below.",
    ],
    level1: [
      "Go to Resources > Input Triggers.",
      "Click \"Add Control Combination,\" press the Right Arrow key, then click \"Pressed.\"",
      "Click \"Edit Script\" for that trigger and add a Position Actor event (or Take Movement in Direction for tile-based movement) targeting main, moving it right a small amount.",
      "Save, then click Play/Test to try it.",
    ],
    level2: ["Repeat for Left, Up, and Down arrow keys as three more Control Combinations, each moving main in that direction."],
    level3: [
      "Animate: since movement here is press-and-release rather than press-and-hold, every key release is a real (if brief) idle moment, even mid-walk - so the idle-reset script you build for the Recolor goal below already makes main blink between its direction color and idle color on every step, with no separate event needed. See Recolor for the how.",
      "Sound: play sfx_blip.wav (Speaker event) each time a movement key is pressed - naturally silent whenever no key is being pressed, since Input Triggers fire on press, not continuously.",
      "Recolor: pair each direction's Position Actor event with a Change Body/Clothing Sprite event, swapping main's whole Sprite for a different pre-made colored one rather than tinting the sprite it already has - import the other placeholder circles via Resources > Sprites the same way you set up main's original appearance. To reset the sprite when idle, add a second \"Released\" Control Combination for each arrow key, and use a Multi-Control Branch checking all 4 arrow keys - its \"none pressed\" path is the real idle moment, so put the reset-to-default sprite swap there.",
      "Clamp: nothing to build - 001 Game Creator does this one automatically. Open your Map's Properties (double-click the Map, or Map menu > Properties) and look at the Boundaries section: all 4 directions default to ticked with no Map set in their Connection dropdowns, which means the engine itself blocks main from ever crossing an edge. See the Key for what would happen if you changed that.",
      "Structure: keep each direction's movement, sound, and recolor events together in that direction's own Input Trigger script - not scattered into separate triggers.",
    ],
    trouble: [
      "Nothing happens on a key press: confirm the Input Trigger is actually set to \"Pressed,\" and that the script attached to it is connected to the Start node in the Script Editor.",
      "Dialog fields may look slightly different depending on your version - the Script Editor and Actor pages at 001gamecreator.com/wiki have the current layout.",
    ],
    resource: "001gamecreator.com/wiki, Student Resource List > 001 Game Creator",
    altPath: {
      title: "Textual Scripting (typed BASIC-like code)",
      intro: "001 Game Creator has a second scripting mode that uses typed, BASIC-like commands instead of the graphical event blocks you used above.",
      warning: "IMPORTANT - READ BEFORE YOU TRY THIS: switching a project's Mode from Graphical to Textual is PERMANENT once you save. Do NOT convert your main [YourName]Game project. Create a separate throwaway test project for this exercise only.",
      steps: [
        "Create a new, separate project (e.g. [YourName]GameTextual) - do not use your main Semester 1 project for this.",
        "Open the Mode menu and switch this new project to Textual (confirm with your instructor that you understand this cannot be undone once saved).",
        "Open the Script Editor for a Right Arrow Input Trigger, same as before.",
        "Instead of dragging graphical event blocks, type the command directly - start typing \"Position\" and use the editor's autocomplete/command list to find the text equivalent of the Position Actor event you used in the graphical path, targeting main with a relative X offset.",
        "Save and Test to confirm the same result as your graphical version.",
      ],
      note: "Exact command names can vary slightly by version - the Script Editor's autocomplete and the 001gamecreator.com/wiki Textual Scripting reference show the current syntax. If you want to keep exploring this mode, do it in your throwaway test project, not your real one. Level 3's five goals (animation, sound, recolor, clamp, and grouping it all under each direction's script) all have Textual equivalents too - same ideas as the graphical path, just typed instead of dragged.",
    },
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Create a new Blueprint project named [YourName]Game (Blank template).",
      "Turn off Lumen - Unreal's dynamic lighting/reflections system, the single biggest reason this engine can feel sluggish on older or integrated-graphics laptops, and something a flat 2D-style scene has no use for anyway: Edit > Project Settings > Engine > Rendering. Under Global Illumination, set Dynamic Global Illumination Method to None. Under Reflections, set Reflection Method to None.",
      "In the World Outliner, select the level's default Directional Light and Sky Atmosphere (and Sky Light, if one is present), then delete them - the fully Unlit look from the next step needs no light sources in the level at all.",
      "Import this lesson's placeholder sprite into the project: in the Content Browser, drag player_circle_orange.png in from this lesson's Assets folder (or Import). Unreal creates a Texture asset from it automatically - accept the default settings.",
      "Create a new Material: Content Browser > Add (+) > Material, name it M_Player_Orange. Open it, select the Material output node itself, and in its Details panel change Shading Model to Unlit. Right-click in empty graph space, search \"Texture Sample,\" add one, and set its Texture to the player_circle_orange asset you just imported. Wire its RGB output into the Material node's Emissive Color input (Unlit materials show only their Emissive Color - that's what makes them render at full brightness with zero light sources). Save and close. This is the real supplied sprite, not a stand-in flat color - Unreal can texture a Material exactly the way every other engine textures a sprite, so that's what this course does here too.",
      "Create a new Blueprint Class to serve as your placeholder character: Content Browser > Add (+) > Blueprint Class, parent Static Mesh Actor, name it BP_Player. Open it and select the inherited Static Mesh component already sitting in its Components panel; in its Details panel set Static Mesh to Cube (Engine content) and Material to M_Player_Orange, then under Transform change Mobility from Static to Movable - a Static component won't actually move at runtime no matter what Blueprint logic tells it to, and setting this on the class itself (rather than a placed instance) means every future instance of BP_Player inherits it automatically.",
      "Still inside BP_Player, open its Class Defaults (the Class Defaults button in the toolbar) and, under the Input category, set Auto Receive Input to Player 0. BP_Player is a plain Actor, not a possessed Pawn, so this is what actually lets it receive keyboard input at all - skip it and every Input Action you add later will silently never fire, with no error anywhere to point at why. Compile, Save, then drag an instance of BP_Player into the level.",
      "Stop a default flying camera pawn from taking over on Play: Content Browser > Add (+) > Blueprint Class, parent Game Mode Base, name it BP_NoPawnGameMode. Open it and in Class Defaults set Default Pawn Class to None (a raw, non-Blueprint GameMode won't let you edit this field at all - it must be a Blueprint). Compile, Save. Then Window > World Settings > Game Mode section > GameMode Override > BP_NoPawnGameMode.",
      "Place Actors panel > search \"Camera\" > drag a Camera Actor into the level. Position it above BP_Player (raise its Location Z, keep X/Y matching BP_Player) and set its Rotation to Pitch = -90, Yaw = -90, Roll = 0 (looking straight down, oriented so Right reads as screen-right). In its Details panel, set Camera Settings > Projection Mode to Orthographic and Ortho Width to 2000 - a fixed, non-perspective view standing in for the fixed-size stage every other engine in this course uses, instead of Unreal's default open 3D world. Then find Auto Activate for Player and set it to Player 0, so this camera becomes the active view automatically with no Pawn required.",
      "In the Content Browser, click Add (+) > Input > Input Action four times, creating IA_MoveRight, IA_MoveLeft, IA_MoveUp, and IA_MoveDown. Leave each one's Value Type as the default (a simple on/off key state, not an axis) - Unreal calls these Input Actions, and they're the current, recommended way to handle input (the older \"Action/Axis Mappings\" system is deprecated).",
      "In the Content Browser, click Add (+) > Input > Input Mapping Context, name it IMC_Default. Open it, click the + next to Mappings four times to add all 4 Input Actions, then under each one click the + to add a key: Right arrow under IA_MoveRight, Left arrow under IA_MoveLeft, Up arrow under IA_MoveUp, Down arrow under IA_MoveDown.",
      "Open BP_Player (not the Level Blueprint - everything for this lesson lives inside BP_Player itself) and add an Event BeginPlay node, then chain: Get Player Controller (Index 0) -> Get Enhanced Input Local Player Subsystem -> Add Mapping Context (Context = IMC_Default, Priority = 0). This one-time step is what actually turns your key mappings on - without it, none of the Input Actions below will fire no matter how they're wired. Doing this inside BP_Player itself, rather than the level's own Level Blueprint, is what keeps BP_Player fully self-contained: drop it into a different level later and it still works immediately, with nothing else to remember.",
    ],
    level1: [
      "In BP_Player's own Event Graph, right-click and search for IA_MoveRight to add an Enhanced Input Action Event node for it.",
      "Add an \"Add Actor World Offset\" node. Leave its Target pin unconnected - inside BP_Player's own graph, an unwired Target automatically means Self, so there's no separate reference to drag in the way a Level Blueprint would need. Connect the Enhanced Input Action Event's Triggered pin (fires every frame for as long as the key is held down, not just once) into the Add Actor World Offset node's execution input. Right-click the Delta Location pin and choose \"Split Struct Pin\" - it starts collapsed as a single Vector pin, and you can't wire a value into just one axis until it's split into separate X/Y/Z pins. Add a \"Get World Delta Seconds\" node, multiply it by a Speed value (300 is a reasonable start), and feed that into Delta Location's X - scaling by delta time is what makes BP_Player glide smoothly instead of jumping a fixed amount, the same idea as Unity's Time.deltaTime or Godot's delta.",
      "Click Compile, then Play to test - click inside the Play viewport first so it actually has keyboard focus, then hold (not just tap) the Right arrow and confirm BP_Player glides for as long as it's held. Once you're done testing, press Shift+F1 to release the viewport's mouse capture and get your cursor back for editing - Unreal shows this same \"Shift+F1 for Mouse Cursor\" hint in the corner of the Play viewport as a reminder.",
    ],
    level2: ["Add the other 3 directions the same way: right-click and add an Enhanced Input Action Event for IA_MoveLeft, IA_MoveUp, and IA_MoveDown (already mapped to their keys in IMC_Default from Setup), each with its own Add Actor World Offset node - Target left unwired (defaults to Self, same as Right), Triggered pin (not Started) driving it, Delta Location scaled by Get World Delta Seconds x Speed the same way as Right, moving in the matching direction (Left = -X, Up = +Y, Down = -Y, matching the Camera's Yaw from Setup)."],
    level3: [
      "Animate: make BP_Player blink (toggle hidden/visible) while moving, and stay visible when idle - drive it from Event Tick with a timer and Set Actor Hidden In Game, kept separate from the per-direction Enhanced Input Action Events.",
      "Sound: chain a Play Sound 2D node (using an imported sfx_blip.wav Sound Wave asset) off the Triggered pin, gated by a short SoundCooldown timer (count it down by Get World Delta Seconds, only play and reset it back to ~0.15 when it's <= 0) - Play Sound 2D is fire-and-forget with no way to check whether it's already playing, so this cooldown is what turns it into a repeating beep-beep for as long as the key is held, instead of one blip per press, matching the other engines' continuous-while-held sound.",
      "Recolor: import the other 4 placeholder circles (player_circle_red.png, player_circle_blue.png, player_circle_green.png, player_circle_yellow.png, all in this lesson's Assets folder) the same way Setup imported the orange one, then build one Material per color the same way Setup built M_Player_Orange - a Texture Sample from that color's Texture wired straight into Emissive Color, Shading Model = Unlit. Swap which Material is applied to BP_Player with a Set Material node per direction (and back to M_Player_Orange when idle) - the same real-sprite-swap every other engine's Recolor goal uses, just with an Unreal Material standing in front of the PNG instead of the PNG being assigned directly. If Set Material doesn't show up when you search for it, see this lesson's trouble section - it's a known Blueprint search quirk, not a missing feature.",
      "Clamp: keep BP_Player inside a defined play area using Get Actor Location, Clamp (Float) on X and Y, then Set Actor Location.",
      "Structure: keep the per-direction pieces (movement, sound, recolor) chained off each direction's own Enhanced Input Action Event, and keep the continuous blink timer in its own Event Tick - don't mix the two execution contexts.",
    ],
    trouble: [
      "Nothing happens, no errors anywhere: this is almost always Auto Receive Input, not the Target pin - open BP_Player's Class Defaults and confirm Auto Receive Input is actually set to Player 0. A plain Actor Blueprint never receives keyboard input without this, and there's no error of any kind when it's missing, just an Enhanced Input Action Event that silently never fires.",
      "Nothing happens: confirm you clicked Compile after editing, and that the white execution wire actually connects your input node to the movement node.",
      "Nothing happens, even though everything looks wired correctly: click once inside the Play viewport before pressing any keys. With no Pawn being possessed (by design, from Setup), the window doesn't automatically grab keyboard focus the way it used to - a plain click gives it focus.",
      "Still nothing happens, even after checking Auto Receive Input, wiring, and viewport focus: this is almost always a missing or misconfigured Add Mapping Context call - confirm BP_Player's own BeginPlay actually calls it with IMC_Default, and that IMC_Default itself actually has all 4 keys mapped (open it and check).",
      "World Settings isn't on the main toolbar: it's under the Window menu instead - Window > World Settings.",
      "GameMode Override is set, but Default Pawn Class still shows DefaultPawn and won't change: this happens when GameMode Override is pointed at the raw \"GameModeBase\" entry instead of the BP_NoPawnGameMode Blueprint from Setup - a non-Blueprint GameMode's fields display as read-only in World Settings. Re-check GameMode Override is actually set to your BP_NoPawnGameMode asset.",
      "BP_Player moves, but Right moves it up the screen (or some other wrong direction) instead of right: adjust the Camera Actor's Yaw in 90-degree steps (try -90, 90, or 180) until Right visually reads as right - the exact value depends on how the camera ended up facing, and it's easiest to just try values and watch the result.",
      "BP_Player (or the whole scene) renders solid black: an Unlit material with nothing wired into Emissive Color renders pure black regardless of lighting - open M_Player_Orange and confirm the Texture Sample's RGB output is actually connected to Emissive Color, and that a Texture is actually assigned on the Texture Sample node itself (an empty Texture Sample defaults to black too).",
      "Movement still happens once per press instead of gliding while held: double check the Enhanced Input Action Event's Triggered pin is driving Add Actor World Offset, not the Started pin - Started only fires the instant a key goes down.",
      "Can't wire Get World Delta Seconds (or anything else) into just the X or Y of Delta Location - it only shows one combined Vector pin: right-click that pin and choose \"Split Struct Pin\" to break it into separate X/Y/Z pins first.",
      "Mouse cursor seems stuck or trapped inside the Play viewport: press Shift+F1 to release it back to the editor - the click that gives the viewport keyboard focus also captures the mouse.",
      "Searching for \"Set Material\" doesn't show the right node (or shows something unrelated like \"Set Material Info\"): this is a known Blueprint search quirk, not a missing feature - either call \"Get Static Mesh Component\" first (Target left unwired, defaults to Self same as everything else this lesson) and search for Set Material from that component's pin, or uncheck the \"Context Sensitive\" box in the top-right of the search popup to show every node regardless of what you dragged from.",
      "Nothing moves, no errors, everything else checks out: select BP_Player's instance in the level and check Mobility under its Transform section. It should already read Movable, inherited from the class default set in Setup - if it somehow shows Static, that's worth fixing directly on BP_Player's own Static Mesh component in the Blueprint editor (not just the one instance), so every instance stays correct.",
      "Unreal is still the most hardware-demanding engine on our list even with Lumen off - if your laptop struggles, it's completely fine to observe this week's demo and build your Key Press Reactor in a lighter engine instead.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "If you'd rather write code than wire Blueprint nodes, you can build the same movement in C++, reusing the same IA_MoveRight and IMC_Default assets from Setup. This needs a code editor (Visual Studio on Windows, or your platform's equivalent) hooked up to Unreal - ask if you're not sure yours is set up.",
      steps: [
        "In [YourName]Game.Build.cs, add \"EnhancedInput\" to PublicDependencyModuleNames (alongside the existing Core/CoreUObject/Engine/InputCore entries) - this is a one-time project setting that makes the Enhanced Input C++ classes available to your code.",
        "Tools > New C++ Class, choose Actor as the parent class, name it MovableActor, click Create Class (Unreal compiles and opens your code editor).",
        "In MovableActor.h, declare an override of SetupPlayerInputComponent, a MoveRight() function, and two UPROPERTY pointers - one for the Input Action, one for the Mapping Context - so you can assign IA_MoveRight and IMC_Default from the Details panel (see code below).",
        "In MovableActor.cpp's BeginPlay(), get the first player controller's Enhanced Input Local Player Subsystem and call AddMappingContext() - the C++ equivalent of the Blueprint path's Add Mapping Context node.",
        "Implement SetupPlayerInputComponent() to cast PlayerInputComponent to UEnhancedInputComponent and bind the MoveRightAction to your MoveRight() function on the Started trigger event.",
        "Implement MoveRight() to call AddActorWorldOffset() - the same underlying function the Blueprint path's node calls.",
        "Compile (the Compile button in the editor toolbar), drag a MovableActor into the level in place of BP_Player, and in its Details panel assign MoveRightAction to IA_MoveRight and MappingContext to IMC_Default.",
        "Play, and test with the Right arrow key.",
      ],
      code: [
        "// MovableActor.h",
        "virtual void SetupPlayerInputComponent(UInputComponent* PlayerInputComponent) override;",
        "void MoveRight();",
        "",
        "UPROPERTY(EditAnywhere, Category=\"Input\")",
        "class UInputAction* MoveRightAction;",
        "",
        "UPROPERTY(EditAnywhere, Category=\"Input\")",
        "class UInputMappingContext* MappingContext;",
        "",
        "// MovableActor.cpp",
        "void AMovableActor::BeginPlay() {",
        "    Super::BeginPlay();",
        "    if (APlayerController* PC = GetWorld()->GetFirstPlayerController()) {",
        "        if (UEnhancedInputLocalPlayerSubsystem* Subsystem =",
        "                ULocalPlayer::GetSubsystem<UEnhancedInputLocalPlayerSubsystem>(PC->GetLocalPlayer())) {",
        "            Subsystem->AddMappingContext(MappingContext, 0);",
        "        }",
        "    }",
        "}",
        "void AMovableActor::SetupPlayerInputComponent(UInputComponent* PlayerInputComponent) {",
        "    if (UEnhancedInputComponent* EnhancedInput = Cast<UEnhancedInputComponent>(PlayerInputComponent)) {",
        "        EnhancedInput->BindAction(MoveRightAction, ETriggerEvent::Triggered, this, &AMovableActor::MoveRight);",
        "    }",
        "}",
        "void AMovableActor::MoveRight() {",
        "    float Speed = 300.f;",
        "    AddActorWorldOffset(FVector(Speed * GetWorld()->GetDeltaSeconds(), 0.f, 0.f));",
        "}",
      ],
      note: "This mirrors the Blueprint path's Add Actor World Offset node - same underlying engine call, just from code, and the same IA_MoveRight/IMC_Default assets either path can use interchangeably. Binding to ETriggerEvent::Triggered instead of Started, and scaling by GetDeltaSeconds(), is what makes this glide smoothly while held instead of stepping once per press - the same continuous-hold behavior as the Blueprint path. Level 2/3: add MoveLeft/MoveUp/MoveDown the same way, each with its own UInputAction* property bound to the matching Input Action asset from Setup. Level 3's five goals are all reachable from C++ too: SetActorHiddenInGame() for the blink timer, UGameplayStatics::PlaySound2D() for sound, SetMaterial() on the mesh component for recolor, and FMath::Clamp() for the boundary - same ideas as the Blueprint path, just called from code.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_02_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS (one per engine) - instructor only
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.0, { after: 120 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 2 Teacher's Key", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "INSTRUCTOR ONLY - do not distribute to students", bold: true, size: 18, color: "B00020" })] }),
    h2("Level 1 - Complete Worked Solution"),
    ...data.l1Notes.map(s => p(s)),
    code(data.l1Code),
    h2("Level 2 - Complete Worked Solution"),
    ...data.l2Notes.map(s => p(s)),
    code(data.l2Code),
  ];
  if (data.l2Extra) children.push(...data.l2Extra.map(s => p(s)));
  children.push(
    h2("Level 3 - All Worked Solutions"),
    p("Every stretch-goal idea from the handout, fully worked - not just one. Useful for validating each idea actually works before class, and as a ready reference no matter which one a student ends up picking (or if a student blows past all of them, this is your own jumping-off point for exploring further right alongside them)."),
  );
  data.l3Ideas.forEach(idea => {
    children.push(h3(idea.title));
    children.push(...idea.notes.map(s => p(s)));
    if (idea.code) children.push(code(idea.code));
  });
  if (data.altNotes) {
    children.push(h2("Alternate Path Notes: " + data.altNotes.title));
    children.push(...data.altNotes.notes.map(s => p(s)));
    if (data.altNotes.code) children.push(code(data.altNotes.code));
    children.push(...data.altNotes.why.map(s => bullet(s)));
  }

  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 2 Teacher's Key - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    l1Notes: ["One hat block plus one Motion block:"],
    l1Code: ["when [right arrow] key pressed", "  change x by 10"],
    l2Notes: ["Four independent scripts, one per arrow key, each a hat block plus one Motion block:"],
    l2Code: [
      "when [up arrow] key pressed        when [down arrow] key pressed",
      "  change y by 10                     change y by -10",
      "",
      "when [left arrow] key pressed      when [right arrow] key pressed",
      "  change x by -10                    change x by 10",
    ],
    l3Ideas: [
      {
        title: "Animate: blink while moving, static when idle",
        notes: [
          "This is the one goal that genuinely needs a loop: Scratch has \"when [key] key pressed\" as a hat block, but no matching \"when [key] key released\" block, so there's no event to trigger off of the instant a key goes back up. The only way to know \"no direction key is currently held\" is to keep checking - which means a separate \"when green flag clicked\" script with a forever loop, using the Sensing category's \"key [ ] pressed?\" reporter (a boolean you can poll) rather than the \"when key pressed\" hat block (an event you can't poll).",
          "The same timer-mod trick as before drives the actual blink once the loop confirms a direction is held: Scratch's Sensing timer reporter counts up in seconds on its own, and (timer mod 0.4) compared against half its own period gives a repeating on/off flip at about 2.5 times a second.",
        ],
        code: [
          "when green flag clicked",
          "forever",
          "  if <<<key [right arrow v] pressed?> or <key [left arrow v] pressed?>> or <<key [up arrow v] pressed?> or <key [down arrow v] pressed?>>> then",
          "    if <(timer mod 0.4) < 0.2> then",
          "      show",
          "    else",
          "      hide",
          "  else",
          "    show",
          "    set color effect to 0",
        ],
      },
      {
        title: "Sound: play on movement, silent when idle",
        notes: [
          "Unlike the blink, this goal doesn't need a loop at all. Upload sfx_blip.wav from this lesson's Assets folder, then drop a single \"start sound\" block right inside each of the 4 \"when key pressed\" scripts. Because that script only runs once at the instant of the press, the sound is automatically a one-shot per press and automatically silent whenever no key is being pressed - the same \"discrete trigger, no cooldown needed\" situation as 001 Game Creator's Input Triggers and Unreal's Enhanced Input Action Events elsewhere in this lesson, not the continuous-loop-with-a-timer-gate that Godot, Unity, GameMaker, GDevelop, and TIC-80 need because their movement is polled every frame instead of event-triggered.",
        ],
        code: [
          "when [right arrow] key pressed",
          "change x by 10",
          "start sound [sfx_blip]",
        ],
      },
      {
        title: "Recolor: different costume per direction, reset when idle",
        notes: [
          "Scratch sprites can hold multiple Costumes, so this goal swaps the whole picture rather than tinting one: upload player_circle_red.png, _blue.png, _green.png, and _yellow.png as extra Costumes on your sprite (Costumes tab > Upload Costume), plus player_circle_orange.png as a dedicated idle costume - even if Setup had you start with a different color, treat orange as the one true idle appearance from here on, so idle never gets confused with a direction that happens to reuse your original color.",
          "\"switch costume to [costume]\" (Looks) takes a different costume per direction; switch it the moment a direction key is pressed, right inside that key's own script.",
          "Switching back to orange when idle has the same problem as the blink: there's no \"key released\" event to trigger it. Rather than inventing a second loop, reuse the exact loop already built for the blink above - it's already checking \"is any direction key currently held\" every frame, so the idle branch is the natural place to also switch back to the idle costume.",
        ],
        code: [
          "-- inside each \"when [direction] key pressed\" script:",
          "switch costume to [player_circle_red]   -- (blue for left, green for up, yellow for down)",
          "",
          "-- inside the green-flag loop's else branch (idle), alongside show:",
          "switch costume to [player_circle_orange]",
        ],
      },
      {
        title: "Bounce off the edges",
        notes: [
          "\"if on edge, bounce\" (Motion) is Scratch's purpose-built edge-collision block, and it's a better fit here than manually checking x position/y position against the stage's bounds - one block instead of four separate if-checks, and it's the block most Scratch tutorials reach for first.",
          "Because movement in this version happens one discrete step at a time (each key press moves 10 steps, not a smooth continuous slide), the bounce check only needs to run once, right after that step - no forever loop required. Drop it directly after the \"change x/y by\" block in each of the 4 key-press scripts.",
          "Known quirk, okay for now: \"if on edge, bounce\" also reverses the sprite's direction (heading), and if a sprite's rotation style is left-right (the default for many costumes), that direction change shows up as the sprite flipping horizontally when it bounces off a side edge. That's expected Scratch behavior, not a bug - precise, non-flipping boundary control is a later-lesson upgrade (the manual position-checking technique), not something to chase down this week.",
        ],
        code: [
          "when [right arrow] key pressed",
          "change x by 10",
          "if on edge, bounce",
        ],
      },
      {
        title: "Structure: group everything triggered by the same action in one place",
        notes: [
          "The most important structural choice in this version: stay with 5 separate scripts - one \"when green flag clicked\" and one \"when [key] key pressed\" per arrow key - instead of collapsing everything into a single forever loop. Scratch is fundamentally event-driven, and \"when key pressed\" hat blocks are the idiomatic way to react to input; a forever loop that polls \"key pressed?\" every frame is the right tool only for the one thing that genuinely can't be event-driven here (detecting \"no key is held,\" for the blink/idle-reset).",
          "Each of the 4 key-press scripts becomes its own self-contained \"if this key, do this\" - movement, sound, recolor, and bounce for that direction all live together in the one script that direction's key press triggers, so there's exactly one place to look for everything the right arrow does. That's what \"the 4 key-press scripts effectively become the if-then-else structure\" means: the branching is the four separate hat blocks themselves, not an if/elseif/else chain inside one shared script.",
        ],
        code: [
          "when [right arrow] key pressed      when [left arrow] key pressed",
          "  change x by 10                      change x by -10",
          "  if on edge, bounce                  if on edge, bounce",
          "  set color effect to 50              set color effect to 100",
          "  start sound [sfx_blip]              start sound [sfx_blip]",
          "",
          "when [up arrow] key pressed         when [down arrow] key pressed",
          "  change y by 10                      change y by -10",
          "  if on edge, bounce                  if on edge, bounce",
          "  set color effect to 150             set color effect to 25",
          "  start sound [sfx_blip]              start sound [sfx_blip]",
        ],
      },
      {
        title: "Complete Level 3 solution (all 5 goals combined)",
        notes: [
          "The full 5-script solution: one \"when green flag clicked\" script handling blink and idle color-reset (the only two things that need continuous checking), plus 4 independent \"when key pressed\" scripts, each handling that direction's movement, bounce, color, and sound in one place.",
        ],
        code: [
          "when green flag clicked",
          "forever",
          "  if <<<key [right arrow v] pressed?> or <key [left arrow v] pressed?>> or <<key [up arrow v] pressed?> or <key [down arrow v] pressed?>>> then",
          "    if <(timer mod 0.4) < 0.2> then",
          "      show",
          "    else",
          "      hide",
          "  else",
          "    show",
          "    switch costume to [player_circle_orange]",
          "",
          "when [right arrow] key pressed",
          "  change x by 10",
          "  if on edge, bounce",
          "  switch costume to [player_circle_red]",
          "  start sound [sfx_blip]",
          "",
          "when [left arrow] key pressed",
          "  change x by -10",
          "  if on edge, bounce",
          "  switch costume to [player_circle_blue]",
          "  start sound [sfx_blip]",
          "",
          "when [up arrow] key pressed",
          "  change y by 10",
          "  if on edge, bounce",
          "  switch costume to [player_circle_green]",
          "  start sound [sfx_blip]",
          "",
          "when [down arrow] key pressed",
          "  change y by -10",
          "  if on edge, bounce",
          "  switch costume to [player_circle_yellow]",
          "  start sound [sfx_blip]",
        ],
      },
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    l1Notes: ["Attach to Player's Sprite2D root - works the same whether you're editing player.tscn directly or selecting Player inside main.tscn:"],
    l1Code: ["func _process(delta):", "    if Input.is_action_pressed(\"ui_right\"):", "        position.x += 200 * delta"],
    l2Notes: ["Attach this script to the Sprite2D root. Using Input.get_vector() is the idiomatic Godot way to combine 4 directions into one movement call:"],
    l2Code: [
      "extends Sprite2D",
      "",
      "var speed := 200",
      "",
      "func _process(delta):",
      "    var dir := Input.get_vector(\"ui_left\", \"ui_right\", \"ui_up\", \"ui_down\")",
      "    position += dir * speed * delta",
    ],
    l3Ideas: [
      {
        title: "Animate: blink while moving, static when idle",
        notes: ["The simplest blink in Godot is toggling the built-in visible property (inherited from CanvasItem) on and off based on an accumulating timer - no AnimationPlayer or sprite swap needed:"],
        code: [
          "var blink_time := 0.0",
          "",
          "func _process(delta):",
          "    var dir := Input.get_vector(\"ui_left\", \"ui_right\", \"ui_up\", \"ui_down\")",
          "    position += dir * speed * delta",
          "    if dir != Vector2.ZERO:",
          "        blink_time += delta",
          "        visible = int(blink_time * 6) % 2 == 0",
          "    else:",
          "        visible = true",
          "        blink_time = 0.0",
        ],
      },
      {
        title: "Sound: beep-beep while moving, silent when idle",
        notes: ["Add an AudioStreamPlayer2D node as a child, set its Stream to the imported sfx_blip.wav. Checking .playing before calling .play() again means it naturally retriggers as soon as the short clip finishes - a beep-beep rhythm, not one continuous tone. Calling .stop() when idle is what makes it actually silent, not just \"not retriggered\":"],
        code: [
          "    if dir != Vector2.ZERO:",
          "        if not $AudioStreamPlayer2D.playing:",
          "            $AudioStreamPlayer2D.play()",
          "    else:",
          "        $AudioStreamPlayer2D.stop()",
        ],
      },
      {
        title: "Recolor: different sprite per direction, reset when idle",
        notes: ["Same Idle/Moving concept as the State Diagram from this lesson's bonus section. This swaps the texture itself rather than tinting it with modulate - preload the other 4 placeholder circles alongside your default one, then assign texture directly:"],
        code: [
          "const TEX_DEFAULT = preload(\"res://player_circle_orange.png\")",
          "const TEX_RIGHT = preload(\"res://player_circle_red.png\")",
          "const TEX_LEFT = preload(\"res://player_circle_blue.png\")",
          "const TEX_UP = preload(\"res://player_circle_green.png\")",
          "const TEX_DOWN = preload(\"res://player_circle_yellow.png\")",
          "",
          "    if dir == Vector2.ZERO:",
          "        texture = TEX_DEFAULT        # idle",
          "    elif dir.x > 0:",
          "        texture = TEX_RIGHT",
          "    elif dir.x < 0:",
          "        texture = TEX_LEFT",
          "    elif dir.y < 0:",
          "        texture = TEX_UP",
          "    else:",
          "        texture = TEX_DOWN",
        ],
      },
      {
        title: "Clamp movement to the screen",
        notes: ["clamp() on position.x/y is the standard Godot boundary pattern; get_viewport_rect().size (inherited from CanvasItem) gives the current screen size dynamically rather than hardcoding numbers:"],
        code: [
          "    var vp := get_viewport_rect().size",
          "    position.x = clamp(position.x, 0, vp.x)",
          "    position.y = clamp(position.y, 0, vp.y)",
        ],
      },
      {
        title: "Structure: group everything triggered by the same action in one place",
        notes: ["Godot's _process(delta) already runs once per frame, so the \"same place in code\" for this lesson is that one function - the trick is organizing what goes inside it, not spreading logic across multiple functions or nodes. Movement and the screen clamp always run, every frame, regardless of direction. Blink, sound, and recolor are all direction-dependent, so they all live under one shared if dir != Vector2.ZERO / else check - not four separate direction checks scattered through the function. That's the Complete Level 3 solution below."],
      },
      {
        title: "Complete Level 3 solution (all 5 goals combined)",
        notes: ["Everything above combined into one script, attached to the Sprite2D root (with an AudioStreamPlayer2D child node, Stream set to sfx_blip.wav):"],
        code: [
          "extends Sprite2D",
          "",
          "const TEX_DEFAULT = preload(\"res://player_circle_orange.png\")",
          "const TEX_RIGHT = preload(\"res://player_circle_red.png\")",
          "const TEX_LEFT = preload(\"res://player_circle_blue.png\")",
          "const TEX_UP = preload(\"res://player_circle_green.png\")",
          "const TEX_DOWN = preload(\"res://player_circle_yellow.png\")",
          "",
          "var speed := 200.0",
          "var blink_time := 0.0",
          "",
          "func _process(delta):",
          "    var dir := Input.get_vector(\"ui_left\", \"ui_right\", \"ui_up\", \"ui_down\")",
          "    position += dir * speed * delta",
          "",
          "    var vp := get_viewport_rect().size",
          "    position.x = clamp(position.x, 0, vp.x)",
          "    position.y = clamp(position.y, 0, vp.y)",
          "",
          "    if dir != Vector2.ZERO:",
          "        blink_time += delta",
          "        visible = int(blink_time * 6) % 2 == 0",
          "        if not $AudioStreamPlayer2D.playing:",
          "            $AudioStreamPlayer2D.play()",
          "        if dir.x > 0:",
          "            texture = TEX_RIGHT",
          "        elif dir.x < 0:",
          "            texture = TEX_LEFT",
          "        elif dir.y < 0:",
          "            texture = TEX_UP",
          "        else:",
          "            texture = TEX_DOWN",
          "    else:",
          "        visible = true",
          "        blink_time = 0.0",
          "        $AudioStreamPlayer2D.stop()",
          "        texture = TEX_DEFAULT",
        ],
      },
    ],
  },
  {
    engine: "Unity", file: "Unity",
    l1Notes: ["Requires using UnityEngine.InputSystem; at the top of the script (added alongside using UnityEngine;), and a fresh Unity project's default Active Input Handling of \"Both\" - see the Handout's Setup/Trouble notes. Attach to the Player GameObject:"],
    l1Code: ["void Update() {", "    float h = 0f;", "    if (Keyboard.current.rightArrowKey.isPressed) h = 1f;", "    if (Keyboard.current.leftArrowKey.isPressed) h = -1f;", "    transform.Translate(new Vector3(h, 0, 0) * 3f * Time.deltaTime);", "}"],
    l2Notes: ["Extend the Level 1 script to use both axes:"],
    l2Code: [
      "public float speed = 3f;",
      "",
      "void Update() {",
      "    float h = 0f;",
      "    if (Keyboard.current.rightArrowKey.isPressed) h = 1f;",
      "    if (Keyboard.current.leftArrowKey.isPressed) h = -1f;",
      "    float v = 0f;",
      "    if (Keyboard.current.upArrowKey.isPressed) v = 1f;",
      "    if (Keyboard.current.downArrowKey.isPressed) v = -1f;",
      "    transform.Translate(new Vector3(h, v, 0) * speed * Time.deltaTime);",
      "}",
    ],
    l3Ideas: [
      {
        title: "Animate: blink while moving, static when idle",
        notes: ["The simplest blink in Unity is toggling SpriteRenderer's built-in enabled property (inherited from Renderer) on and off based on an accumulating timer - no Animator or sprite swap needed:"],
        code: [
          "private SpriteRenderer sr;",
          "private float blinkTime = 0f;",
          "void Start() { sr = GetComponent<SpriteRenderer>(); }",
          "",
          "void Update() {",
          "    float h = 0f;",
          "    if (Keyboard.current.rightArrowKey.isPressed) h = 1f;",
          "    if (Keyboard.current.leftArrowKey.isPressed) h = -1f;",
          "    float v = 0f;",
          "    if (Keyboard.current.upArrowKey.isPressed) v = 1f;",
          "    if (Keyboard.current.downArrowKey.isPressed) v = -1f;",
          "    bool moving = h != 0 || v != 0;",
          "    if (moving) {",
          "        blinkTime += Time.deltaTime;",
          "        sr.enabled = ((int)(blinkTime * 6) % 2 == 0);",
          "    } else {",
          "        sr.enabled = true;",
          "        blinkTime = 0f;",
          "    }",
          "}",
        ],
      },
      {
        title: "Sound: beep-beep while moving, silent when idle",
        notes: ["Add an AudioSource component and assign sfx_blip.wav as its AudioClip. Checking .isPlaying before calling .Play() again means it naturally retriggers as soon as the short clip finishes - a beep-beep rhythm, not one continuous tone. Calling .Stop() when idle is what makes it actually silent - PlayOneShot() alone has no matching \"stop early\" call:"],
        code: [
          "public AudioSource audioSrc;",
          "",
          "if (moving) {",
          "    if (!audioSrc.isPlaying) audioSrc.Play();",
          "} else {",
          "    audioSrc.Stop();",
          "}",
        ],
      },
      {
        title: "Recolor: different sprite per direction, reset when idle",
        notes: ["Same Idle/Moving concept as the State Diagram from this lesson's bonus section. This swaps the Sprite itself rather than tinting it with sr.color - add 5 public Sprite fields (spriteDefault, spriteRight, spriteLeft, spriteUp, spriteDown) and drag the 5 placeholder images onto them in the Inspector:"],
        code: [
          "    if (h == 0 && v == 0) sr.sprite = spriteDefault;   // idle",
          "    else if (h > 0) sr.sprite = spriteRight;",
          "    else if (h < 0) sr.sprite = spriteLeft;",
          "    else if (v > 0) sr.sprite = spriteUp;",
          "    else sr.sprite = spriteDown;",
        ],
      },
      {
        title: "Clamp movement to the screen",
        notes: ["The -8/8 and -4.5/4.5 bounds approximate the visible camera area for Unity's default 2D orthographic camera at size 5 - adjust to match whatever camera size the class ends up using:"],
        code: [
          "    Vector3 pos = transform.position;",
          "    pos.x = Mathf.Clamp(pos.x, -8f, 8f);",
          "    pos.y = Mathf.Clamp(pos.y, -4.5f, 4.5f);",
          "    transform.position = pos;",
        ],
      },
      {
        title: "Structure: group everything triggered by the same action in one place",
        notes: ["Unity's Update() already runs once per frame, so the \"same place in code\" for this lesson is that one method - the trick is organizing what goes inside it, not spreading logic across multiple methods or components. Movement and the screen clamp always run, every frame, regardless of direction. Blink, sound, and recolor are all direction-dependent, so they all live under one shared if (moving) / else check - not four separate direction checks scattered through the method. That's the Complete Level 3 solution below."],
      },
      {
        title: "Complete Level 3 solution (all 5 goals combined)",
        notes: ["Everything above combined into one script, attached to the Player GameObject (with an AudioSource component, AudioClip set to sfx_blip.wav):"],
        code: [
          "public float speed = 3f;",
          "public AudioSource audioSrc;",
          "public Sprite spriteDefault, spriteRight, spriteLeft, spriteUp, spriteDown;",
          "private SpriteRenderer sr;",
          "private float blinkTime = 0f;",
          "",
          "void Start() {",
          "    sr = GetComponent<SpriteRenderer>();",
          "}",
          "",
          "void Update() {",
          "    float h = 0f;",
          "    if (Keyboard.current.rightArrowKey.isPressed) h = 1f;",
          "    if (Keyboard.current.leftArrowKey.isPressed) h = -1f;",
          "    float v = 0f;",
          "    if (Keyboard.current.upArrowKey.isPressed) v = 1f;",
          "    if (Keyboard.current.downArrowKey.isPressed) v = -1f;",
          "    transform.Translate(new Vector3(h, v, 0) * speed * Time.deltaTime);",
          "",
          "    Vector3 pos = transform.position;",
          "    pos.x = Mathf.Clamp(pos.x, -8f, 8f);",
          "    pos.y = Mathf.Clamp(pos.y, -4.5f, 4.5f);",
          "    transform.position = pos;",
          "",
          "    bool moving = h != 0 || v != 0;",
          "    if (moving) {",
          "        blinkTime += Time.deltaTime;",
          "        sr.enabled = ((int)(blinkTime * 6) % 2 == 0);",
          "        if (!audioSrc.isPlaying) audioSrc.Play();",
          "        if (h > 0) sr.sprite = spriteRight;",
          "        else if (h < 0) sr.sprite = spriteLeft;",
          "        else if (v > 0) sr.sprite = spriteUp;",
          "        else sr.sprite = spriteDown;",
          "    } else {",
          "        sr.enabled = true;",
          "        blinkTime = 0f;",
          "        audioSrc.Stop();",
          "        sr.sprite = spriteDefault;",
          "    }",
          "}",
        ],
      },
    ],
    altNotes: {
      title: "Visual Scripting",
      notes: [
        "Unity's Visual Scripting graphs compile to the same underlying calls as the C# path - no performance or capability loss for a beginner project.",
        "Watch for students forgetting to wire the flow (white/execution) connection from the On Input System Event Vector 2 node's Trigger output into Translate - a common visual-scripting mistake that looks identical to the graph \"doing nothing.\"",
        "If the Input System nodes don't show up in the fuzzy finder at all, the Visual Scripting Node Library needs to be regenerated once per project (Project Settings > Visual Scripting > Regenerate Nodes) - easy to forget since it's a one-time setup step, not something students do every lesson.",
        "If Move doesn't appear in the On Input System Event Vector 2 node's Input Action dropdown, the Player Input component's Actions field is probably still empty - it needs the PlayerControls asset assigned before Visual Scripting can see any of its actions.",
        "The Level 3 goals map onto nodes the same way the code maps onto members: enabled toggles to a \"Set Member\" node on Sprite Renderer, .isPlaying/.Play()/.Stop() to Audio Source nodes, and sprite to another \"Set Member\" node on Sprite Renderer (with the 5 placeholder images wired in as Sprite variables) - all wired under one Branch node keyed off the same moving check used for movement.",
      ],
      why: [
        "Visual Scripting is part of Unity's own Junior Programmer pathway content, so this stays aligned with the certification track for students who prefer it.",
        "The node-graph structure (event -> data -> action) is the same conceptual shape as the C# script - worth pointing out explicitly.",
        "Using the On Input System Event Vector 2 node (bound to a Move Input Action) instead of the old Get Axis node keeps the Visual Scripting path genuinely aligned with the C# path's new Input System, rather than secretly still depending on the legacy Input Manager underneath a different-looking node.",
      ],
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    l1Notes: ["In obj_player's Step event:"],
    l1Code: ["if (keyboard_check(vk_right)) x += 4;"],
    l2Notes: ["All 4 directions in the Step event:"],
    l2Code: [
      "if (keyboard_check(vk_right)) x += 4;",
      "if (keyboard_check(vk_left))  x -= 4;",
      "if (keyboard_check(vk_up))    y -= 4;",
      "if (keyboard_check(vk_down))  y += 4;",
    ],
    l3Ideas: [
      {
        title: "Animate: blink while moving, static when idle",
        notes: ["GameMaker objects have a built-in visible variable that controls whether the Draw event runs automatically - toggling it on and off using a counter incremented in the Step event is the simplest blink, no sprite swap needed. Add blink_timer = 0; to the Create event first:"],
        code: [
          "// Step event:",
          "var moving = keyboard_check(vk_right) || keyboard_check(vk_left) || keyboard_check(vk_up) || keyboard_check(vk_down);",
          "if (moving) {",
          "    blink_timer += 1;",
          "    visible = (blink_timer mod 20 < 10);",
          "} else {",
          "    visible = true;",
          "    blink_timer = 0;",
          "}",
        ],
      },
      {
        title: "Sound: beep-beep while moving, silent when idle",
        notes: ["Add sfx_blip.wav as a sound resource (keep its default name, sfx_blip) (right-click > Create Sound). audio_is_playing() accepts the sound asset directly and checks whether any instance of it is currently playing - checking it before calling audio_play_sound() again creates the beep-beep retrigger rhythm rather than one continuous tone. audio_stop_sound() on the asset is what actually silences it when idle:"],
        code: [
          "if (moving) {",
          "    if (!audio_is_playing(sfx_blip)) audio_play_sound(sfx_blip, 1, false);",
          "} else {",
          "    audio_stop_sound(sfx_blip);",
          "}",
        ],
      },
      {
        title: "Recolor: different sprite per direction, reset when idle",
        notes: ["Same Idle/Moving concept as the State Diagram from this lesson's bonus section. This swaps sprite_index itself rather than tinting with image_blend - import spr_player_red, spr_player_blue, spr_player_green, and spr_player_yellow from this lesson's Assets folder alongside spr_player_orange as your default:"],
        code: [
          "if (keyboard_check(vk_right))      sprite_index = spr_player_red;",
          "else if (keyboard_check(vk_left))  sprite_index = spr_player_blue;",
          "else if (keyboard_check(vk_up))    sprite_index = spr_player_green;",
          "else if (keyboard_check(vk_down))  sprite_index = spr_player_yellow;",
          "else                                sprite_index = spr_player_orange; // idle",
        ],
      },
      {
        title: "Clamp movement to the screen",
        notes: ["clamp(value, min, max) is a built-in GML function - simplest possible boundary check, no manual if/else needed. Add to the bottom of the Step event:"],
        code: ["x = clamp(x, 0, room_width);", "y = clamp(y, 0, room_height);"],
      },
      {
        title: "Structure: group everything triggered by the same action in one place",
        notes: ["The Step event already runs once per frame, so the \"same place in code\" for this lesson is that one event - the trick is organizing what goes inside it, not spreading logic across multiple events or objects. Movement and the room clamp always run, every frame, regardless of direction. Blink, sound, and recolor are all direction-dependent, so they all live under one shared if (moving) / else check - not four separate direction checks scattered through the event. That's the Complete Level 3 solution below."],
      },
      {
        title: "Complete Level 3 solution (all 5 goals combined)",
        notes: ["Everything above combined, across obj_player's two events:"],
        code: [
          "// Create event:",
          "blink_timer = 0;",
          "",
          "// Step event:",
          "var moving = keyboard_check(vk_right) || keyboard_check(vk_left) || keyboard_check(vk_up) || keyboard_check(vk_down);",
          "if (keyboard_check(vk_right)) x += 4;",
          "if (keyboard_check(vk_left))  x -= 4;",
          "if (keyboard_check(vk_up))    y -= 4;",
          "if (keyboard_check(vk_down))  y += 4;",
          "",
          "x = clamp(x, 0, room_width);",
          "y = clamp(y, 0, room_height);",
          "",
          "if (moving) {",
          "    blink_timer += 1;",
          "    visible = (blink_timer mod 20 < 10);",
          "    if (!audio_is_playing(sfx_blip)) audio_play_sound(sfx_blip, 1, false);",
          "    if (keyboard_check(vk_right))      sprite_index = spr_player_red;",
          "    else if (keyboard_check(vk_left))  sprite_index = spr_player_blue;",
          "    else if (keyboard_check(vk_up))    sprite_index = spr_player_green;",
          "    else                                sprite_index = spr_player_yellow; // down",
          "} else {",
          "    visible = true;",
          "    blink_timer = 0;",
          "    audio_stop_sound(sfx_blip);",
          "    sprite_index = spr_player_orange;",
          "}",
        ],
      },
    ],
    altNotes: {
      title: "GML Visual",
      notes: [
        "The Visual toggle is per-event, per-object - a student can mix Code and Visual across different events/objects with no penalty.",
        "Movement uses \"Jump To Point\" (Movement category) with Relative checked on both X and Y - it's the direct visual equivalent of x += 4 (and y += 4), but it's a purpose-built, self-labeled block rather than the generic Assign Variable, which reads better for a student building entirely in blocks. Leaving Relative unchecked overwrites the position outright, which is a common student mistake to watch for either way.",
        "The Level 3 goals should reach for the same kind of purpose-built block wherever GameMaker has one, instead of defaulting to Assign Variable everywhere: \"Set Sprite\" (Instance category) for the recolor goal instead of Assign Variable on sprite_index, and \"If Audio Is Playing\" -> Not -> \"Play Audio\" (Audio category) as the retrigger guard for the sound goal, with a matching \"Stop Audio\" in the idle branch (in place of audio_is_playing()/audio_play_sound()/audio_stop_sound()). Assign Variable remains the right - and only - choice for anything that isn't a recognized built-in, like the custom blink_timer or moving variables; the blink itself can still use \"Set Instance Variable\" -> Visible, since that one is a recognized built-in too.",
      ],
      why: [
        "Demonstrating both first shows students the two approaches produce identical runtime behavior, which builds intuition for what code \"really\" does under the hood.",
        "If a student picks Visual as their primary style for this course, GML Visual remains fully supported all the way through GameMaker's certification track.",
        "Reaching for the most specific, well-named block available (Jump To Point, Set Sprite, If Audio Is Playing, Play Audio, Stop Audio, Set Instance Variable) before falling back to the generic Assign Variable keeps the block graph readable as plain English - Assign Variable stays available as the fallback for anything without a dedicated block, like a custom variable.",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    l1Notes: ["Simplest path - just the behavior, no events needed. Level 1 only requires demonstrating the Right key, even though all 4 already work:"],
    l1Code: ["Object: Player", "Behavior: Top-down movement (default settings)"],
    l2Notes: ["No manual events needed - the Top-down movement behavior already reads all 4 arrow keys by default once applied to the Player object. Confirm this in Preview by testing all 4 directions."],
    l2Code: ["Object: Player", "Behavior: Top-down movement (default settings: 4-direction, Max speed ~150 px/s)"],
    l3Ideas: [
      {
        title: "Animate: blink while moving, static when idle",
        notes: ["Flash visibility (blink) is an action that belongs to a behavior, not a general-purpose action - it won't appear in Player's action search until the Flash behavior is attached. Select Player > Behaviors tab > Add a behavior > search \"Flash\" > Apply. Only then do \"Flash visibility (blink)\" (start) and \"Stop flashing visibility (blink)\" (stop) show up when you search Player's actions - the same behavior-first pattern as Angle of movement needing Top-down movement in the Recolor idea below. No counter variable or Draw-event toggle needed, unlike the code-based engines:"],
        code: [
          "Condition: Is moving  (Top-down movement)",
          "Action: Flash visibility (blink) of Player, Duration 0 (keep flashing until stopped)",
          "",
          "Condition: Is moving = False",
          "Action: Stop flashing visibility (blink) of Player",
        ],
      },
      {
        title: "Sound: beep-beep while moving, silent when idle",
        notes: ["Import sfx_blip.wav as a Sound resource. \"Play a sound\" fires once per call, so pairing it with a scene timer creates the beep-beep retrigger rhythm instead of one sound overlapping itself every frame:"],
        code: [
          "At the beginning of the scene:",
          "Action: Start (or reset) a scene timer \"Blip\"",
          "",
          "Every frame:",
          "Condition: Is moving  (Top-down movement)",
          "Condition: Timer value \"Blip\" >= 0.3 seconds",
          "Action: Play a sound (sfx_blip.wav)",
          "Action: Start (or reset) a scene timer \"Blip\"",
        ],
      },
      {
        title: "Recolor: different sprite per direction, reset when idle",
        notes: ["Same Idle/Moving concept as the State Diagram from this lesson's bonus section. This swaps which Animation is showing rather than tinting the object - add the other 4 placeholder circles as extra Animations on Player (Edit object > Animations > Add an animation), one named Red, Blue, Green, and Yellow, plus your default one renamed Orange. The thing being built here is 5 separate Events, not conditions/actions on their own - each Event is one row that bundles together its own Condition(s) and Action(s), and an Event's Action(s) only run when all of that SAME Event's Condition(s) are true. There's no separate step to \"wire\" a condition to an action; they're never independent pieces, only ever parts of the one Event that owns them - which is also why cramming multiple direction-conditions into a single Event would AND them together instead of offering a choice between them. In the Scene's Events tab, each Event row has small, subdued text at its top-left reading Add condition and Add action - no icon, easy to skim past. For each of the 5 Events: click its Add condition, pick Player, then look for a separate Top-down movement category in the list (not Player's own conditions, since this one belongs to the behavior) and choose Angle of movement - it compares to a single angle plus a tolerance, so each direction needs its own Event rather than one shared condition. Then, still inside that same Event, click its Add action, pick Player again, and search for Animation (by name) under Animations and images - it's labeled that way in the picker, but displays as \"Change the animation of Player...\" once placed:"],
        code: [
          "Event 1 - Condition: Angle of movement = 0 deg (tolerance 45)    Action: Change the animation of Player to \"Red\"     (Right)",
          "Event 2 - Condition: Angle of movement = 90 deg (tolerance 45)   Action: Change the animation of Player to \"Blue\"    (Down)",
          "Event 3 - Condition: Angle of movement = 180 deg (tolerance 45)  Action: Change the animation of Player to \"Green\"   (Left)",
          "Event 4 - Condition: Angle of movement = 270 deg (tolerance 45)  Action: Change the animation of Player to \"Yellow\"  (Up)",
          "Event 5 - Condition: Is moving = False                           Action: Change the animation of Player to \"Orange\"  (idle)",
        ],
      },
      {
        title: "Clamp movement to the screen",
        notes: ["Add the Stay On Screen (2D) behavior to Player (Add a behavior > search \"Stay On Screen\"). It constrains the object's position to the camera viewport automatically, the same idea as GameMaker's clamp() or Godot's clamp() - just as a behavior instead of a line of logic, so no events are needed at all."],
      },
      {
        title: "Structure: group everything triggered by the same action in one place",
        notes: ["GDevelop's version of \"one place in code\" is an event group - right-click the event sheet, Add a group, then drag the Animate, Sound, and Recolor events (goals 1-3) inside it and label it something like \"Direction feedback.\" The Stay On Screen behavior and the Level 1/2 movement events stay outside the group since they're not part of the same triggered action. That's the Complete Level 3 solution below."],
      },
      {
        title: "Complete Level 3 solution (all 5 goals combined)",
        notes: ["Everything above, organized as one event group named \"Direction feedback\":"],
        code: [
          "At the beginning of the scene:",
          "Action: Start (or reset) a scene timer \"Blip\"",
          "",
          "Event group: \"Direction feedback\"",
          "",
          "  Condition: Is moving  (Top-down movement)",
          "  Action: Flash visibility (blink) of Player, Duration 0",
          "",
          "  Condition: Is moving",
          "  Condition: Timer value \"Blip\" >= 0.3 seconds",
          "  Action: Play a sound (sfx_blip.wav)",
          "  Action: Start (or reset) a scene timer \"Blip\"",
          "",
          "  Condition: Angle of movement = 0 deg (tolerance 45)    ->  Action: Change the animation of Player to \"Red\"     (Right)",
          "  Condition: Angle of movement = 90 deg (tolerance 45)   ->  Action: Change the animation of Player to \"Blue\"    (Down)",
          "  Condition: Angle of movement = 180 deg (tolerance 45)  ->  Action: Change the animation of Player to \"Green\"   (Left)",
          "  Condition: Angle of movement = 270 deg (tolerance 45)  ->  Action: Change the animation of Player to \"Yellow\"  (Up)",
          "",
          "  Condition: Is moving = False",
          "  Action: Stop flashing visibility (blink) of Player",
          "  Action: Change the animation of Player to \"Orange\"",
          "",
          "(Stay On Screen (2D) behavior on Player handles Clamp automatically - no event needed)",
        ],
      },
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    l1Notes: ["Full cart:"],
    l1Code: ["x=60 y=60", "function TIC()", " cls(0)", " if btn(3) then x=x+1 end", " spr(1,x,y,0)", "end"],
    l2Notes: ["All 4 directions using the standard button indices (0=Up, 1=Down, 2=Left, 3=Right):"],
    l2Code: [
      "x=60 y=60",
      "function TIC()",
      " cls(0)",
      " if btn(0) then y=y-1 end",
      " if btn(1) then y=y+1 end",
      " if btn(2) then x=x-1 end",
      " if btn(3) then x=x+1 end",
      " spr(1,x,y,0)",
      "end",
    ],
    l3Ideas: [
      {
        title: "Animate: blink while moving, static when idle",
        notes: [
          "This solution uses every parameter of spr(): spr(id, x, y, colorkey, scale, flip, rotate, w, h) - confirmed against TIC-80's own spr() documentation. The two new ones are w and h: how many 8x8 tiles wide/tall to draw as one connected image (both default to 1). Setting w=2, h=2 draws a 2x2 block of tiles (16x16 pixels) starting at index 1 - a \"composite sprite\" - instead of the single top-left 8x8 tile that spr(1,x,y,0) was drawing before. In the Sprite Editor (F2, or the monster-icon button), draw your character across a full 2x2 block of tiles rather than a single tile, so the four cells read as one connected image instead of a quarter of one.",
          "Per TIC-80's spr() documentation: \"Fractional factors are floored to integers, and both a factor of 0 and negative factors cause the sprite to not get drawn.\" That's the trick behind the blink - alternating the scale argument between 0 and 1 makes the sprite blink on and off, no second image needed.",
          "f is a flash-toggle counter that climbs by 0.1 every frame; math.floor(f)%2 flips between 0 and 1 roughly 3 times a second (at 60fps). While a direction is held, s takes that flipping value (blink); when idle, s is pinned to 1 (always visible):",
        ],
        code: [
          "f = 0  -- flash toggle",
          "...",
          "f=f+0.1  -- advance flash speed every frame",
          "if btn(3) then",
          "  s=math.floor(f)%2  -- blinks 0/1 while moving",
          "else",
          "  s=1  -- always visible when idle",
          "end",
          "...",
          "spr( 1, x, y, 0, s, 0, 0, 2, 2 )  -- s (scale) is what makes it blink",
        ],
      },
      {
        title: "Sound: beep-beep while moving, silent when idle",
        notes: [
          "The goal here is a repeating beep-beep for as long as a direction is held, stopping the instant it's released - a different pattern than a single blip on press.",
          "b is a beep-toggle counter climbing by 0.2 every frame; math.floor(b)%2 alternates between sound slots 0 and 1 roughly every 5 frames, so you need two short SFX slots set up (e.g. two quick blips at slightly different pitches) for the alternation to read as a beep-beep rather than one repeated note.",
          "This calls sfx() every single frame a direction is held (not just on the frame it's first pressed) - that's intentional here, not a bug, because the goal is a continuous sound. Since a plain sfx() call has no built-in \"note off\" (duration defaults to play-forever), the idle/else branch explicitly calls sfx(-1) to silence channel 0 the moment no direction key is down:",
        ],
        code: [
          "b = 1  -- beep toggle",
          "...",
          "b=b+0.2  -- advance beep speed every frame",
          "if btn(3) then sfx(math.floor(b)%2)  -- alternates sound 0/1 for a beep-beep while held",
          "else sfx(-1) end  -- stop channel 0 the instant no direction key is down",
        ],
      },
      {
        title: "Recolor: different color per direction, reset when idle",
        notes: [
          "Worth calling out to the class: every other engine in this lesson does Recolor by swapping in a whole separate pre-made colored image, because their color-tint functions (image_blend, modulate, spriteRenderer.color, and so on) only give a clean result on white source art, and muddy on anything else - swapping images sidesteps that entirely. TIC-80 doesn't have that problem, so it doesn't need that workaround: poke4() rewrites an actual palette slot, not a multiplied tint, so the recolor is exact regardless of what the sprite's original color was. That makes TIC-80 the one engine in the course where true live recoloring is both possible and genuinely simpler than swapping images - a real difference in what the hardware/engine offers, not an arbitrary inconsistency.",
          "Heads up: TIC-80 has no pal()/palette()/palset() function - that's a PICO-8 thing. TIC-80's own documentation for spr() has a section called \"Color Swapping Sprites\" that uses exactly this poke4() + PALETTE_MAP technique, so this is the officially intended way to do a live recolor.",
          "In the Sprite Editor (F2, or the monster-icon button), the color palette is the strip in the lower-left. Click through it to see which index is which color, and note which index makes up most of your character's fill - in this cart that's index 10.",
          "PALETTE_MAP (0x3FF0) is a magic-looking value, but it has a plain explanation: it's a byte address in VRAM. poke4() writes in 4-bit nibbles, and TIC-80's own peek/poke documentation explains that reaching a byte address in nibble units means multiplying by 2 (each byte is 2 nibbles). So PALETTE_MAP*2 gets you to the start of that byte's nibbles, and adding the original color index (0-15) selects which of the 16 palette slots you're overriding - poke4(PALETTE_MAP*2+10, 6) means \"wherever color 10 would be drawn, draw color 6 instead.\"",
          "Important: this doesn't touch the actual palette entries - it's a live remap applied while drawing. Color 10 still exists; you're just redirecting it. That's why swapping index 10 to 6 and later poking index 10 back to 10 (itself) fully restores the original look - nothing was destroyed, just temporarily overridden:",
        ],
        code: [
          "PALETTE_MAP=0x3FF0",
          "...",
          "if btn(3) then poke4(PALETTE_MAP*2+10,6)  -- while held, redirect color 10 to green (6)",
          "else poke4(PALETTE_MAP*2+10,10) end  -- idle: redirect color 10 back to itself - fully restored",
        ],
      },
      {
        title: "Clamp movement to the screen",
        notes: ["TIC-80's screen is 240x136; clamping to 0-232/0-128 keeps the sprite fully on screen. Placed ahead of the if/elseif/else block so it applies every frame regardless of which branch runs:"],
        code: [" x=math.max(0,math.min(232,x))", " y=math.max(0,math.min(128,y))"],
      },
      {
        title: "Structure: group everything per trigger in one place",
        notes: [
          "The most important structural idea in this version: every goal above is folded into the SAME single if/elseif/else block that already handles movement from Level 2, rather than writing separate blocks that each re-check btn() on their own. Each direction's branch now does its movement, its recolor, its beep, and its blink state together in one place - and the else (idle) branch resets color, sound, and scale together too. That keeps \"what happens when I press Right\" fully readable in one spot instead of scattered across separate checks:",
        ],
        code: [
          "if btn(3) then -- right: move + recolor + beep + blink, all together",
          "  x=x+spd",
          "  poke4(PALETTE_MAP*2+10,6)",
          "  sfx(math.floor(b)%2)",
          "  s=math.floor(f)%2",
          "end",
        ],
      },
      {
        title: "Complete Level 3 solution (all 5 goals combined)",
        notes: [
          "The full working cart, combining every goal above - movement, blink, recolor, and beep-while-held - inside one if/elseif/else block, plus the clamp-to-screen check. The indentation and comment alignment are preserved exactly as authored, since the staggered spacing is what makes the branches easy to scan at a glance:",
        ],
        code: [
          "-- title:   Mini-Game",
          "-- author:  Robert Kuropkat",
          "-- desc:    Semester 01 project",
          "-- site:    https://doersofstuff.com/",
          "-- license: MIT License (change this to your license of choice)",
          "-- version: 0.1",
          "-- script:  lua",
          "",
          "PALETTE_MAP=0x3FF0",
          "",
          "x   = 60  -- starting x-position",
          "y   = 60  -- starting y-position",
          "spd = 1   -- speed",
          "b   = 1   -- beep toggle",
          "f   = 0   -- flash toggle",
          "s   = 1   -- sprite scale",
          "",
          "function TIC()",
          "",
          "    cls(0)",
          "",
          " b=b+0.2  -- set beep speed",
          " f=f+0.1  -- set flash speed",
          "",
          " x=math.max(0,math.min(232,x))",
          " y=math.max(0,math.min(128,y))",
          "",
          " if     btn(0) then -- up",
          "               poke4(PALETTE_MAP*2+10,3)",
          "               sfx(math.floor(b)%2)",
          "               y=y-spd ",
          "               s=math.floor(f)%2",
          " elseif btn(1) then -- down",
          "               poke4(PALETTE_MAP*2+10,1)",
          "               sfx(math.floor(b)%2)",
          "               y=y+spd ",
          "               s=math.floor(f)%2",
          " elseif btn(2) then -- left",
          "               poke4(PALETTE_MAP*2+10,0)",
          "               sfx(math.floor(b)%2)",
          "               x=x-spd ",
          "               s=math.floor(f)%2",
          " elseif btn(3) then -- right",
          "               poke4(PALETTE_MAP*2+10,6)",
          "               sfx(math.floor(b)%2)",
          "               x=x+spd ",
          "               s=math.floor(f)%2",
          " else          poke4(PALETTE_MAP*2+10,10) ",
          "               sfx(-1)",
          "               s=1",
          " end",
          "",
          "    spr( 1  -- index",
          "       , x  -- x-position",
          "                , y  -- y-position",
          "                , 0  -- color key",
          "                , s  -- scale",
          "                , 0        -- flip",
          "                , 0  -- rotate",
          "                , 2  -- width in sprites",
          "                , 2  -- height in sprites",
          "                )",
          "",
          "end",
        ],
      },
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    l1Notes: ["One Input Trigger:"],
    l1Code: ["Input Trigger: Right Arrow (Pressed) -> Position Actor(main, X: main.X + 8, Y: main.Y)"],
    l2Notes: [
      "Four Input Triggers, each a Control Combination bound to one arrow key set to \"Pressed,\" each running a script with a single Position Actor event targeting main:",
    ],
    l2Code: [
      "Input Trigger: Right Arrow (Pressed) -> Position Actor(main, X: main.X + 8, Y: main.Y)",
      "Input Trigger: Left  Arrow (Pressed) -> Position Actor(main, X: main.X - 8, Y: main.Y)",
      "Input Trigger: Up    Arrow (Pressed) -> Position Actor(main, X: main.X,     Y: main.Y - 8)",
      "Input Trigger: Down  Arrow (Pressed) -> Position Actor(main, X: main.X,     Y: main.Y + 8)",
    ],
    l3Ideas: [
      {
        title: "Animate: blink while moving, static when idle",
        notes: ["001 Game Creator has no press-and-hold movement without extra work (see the Super-Stretch Ideas list), so walking here means repeatedly pressing and releasing a direction key - and every release genuinely passes through a real moment with no key held, even mid-stride. That means the idle-reset script built for the Recolor goal below already produces this goal too, for free: main's sprite visibly swaps between its direction sprite and its idle sprite on every step, which reads as a blink while walking and goes static the instant a student actually stops moving. No separate event needed - see Recolor for the scripts. Worth calling out to the class directly as a deliberate consequence, not a coincidence: this is what happens when idle-detection is built from key state under a press-and-release movement scheme, and it will disappear on its own once movement switches to true press-and-hold (Change Movement X/Y, a Super-Stretch idea), since Released will then only fire on a genuine stop."],
      },
      {
        title: "Sound: play on movement, silent otherwise",
        notes: ["Import sfx_blip.wav (Resources > Import > Sound), then add a \"Play Sound\" event directly after each of the 4 Position Actor events. Because each Input Trigger fires on \"Pressed\" rather than continuously, this is naturally a one-shot per key press - no cooldown timer needed, unlike engines that poll movement every frame."],
      },
      {
        title: "Recolor: different sprite per direction, reset when idle",
        notes: ["Pair each of the 4 Position Actor events with a Change Body/Clothing Sprite event for that direction - swapping main's whole Sprite for a different pre-made colored one, rather than tinting the sprite it already has. Import the other placeholder circles as extra Sprites (Resources > Sprites > Create New > No Template, same process as main's original appearance in Setup) so you have one Sprite per color to swap between. 001 Game Creator's Input Triggers don't have a standing \"nothing is pressed\" trigger, so build idle detection instead: add a second Control Combination for each arrow key, this one set to \"Released\" rather than \"Pressed,\" and give it a Multi-Control Branch checking all 4 arrow keys - the last path Multi-Control Branch provides is \"none of these are pressed,\" which is the actual idle moment, so that's where the reset-to-default Change Body/Clothing Sprite event goes. Checking all 4 keys (not just \"this key released\") matters because a student could release Right while still holding Up - the sprite should only reset once every direction is let go. These same two scripts (per-direction Pressed sprite-swap + shared idle Released reset) are also what produces the Animate goal above, as a side effect - see that entry for why. Same Idle/Moving concept as the State Diagram from this lesson's bonus section."],
      },
      {
        title: "Clamp: automatic, no scripting needed",
        notes: ["This is a genuine freebie in 001 Game Creator, not something to build - open the Map's Properties and look at the Boundaries section (bottom-right of the window): a checkmark for each of the 4 directions plus Floor and Ceiling, all ticked by default, and a Connection dropdown per edge for linking to another Map. With every direction ticked and every Connection left at (None), the engine itself refuses to let main cross an edge - Position Actor moves that would push main past the boundary are simply stopped before they happen. That's why a \"Crossed Map's Edge\" Actor trigger doesn't behave like a live \"player tried to leave\" hook here: with the boundary actively blocking movement, main's position never actually reaches past the edge during ordinary play, so there's nothing for that trigger to catch beyond a one-time check when the Actor is first placed. It's worth understanding this rather than just accepting it, because it's easy to accidentally break: unticking one of those 4 boundary checkmarks, or setting a Connection to a second Map, lets main walk straight off that edge - the opposite of the Clamp goal - so treat those Boundaries settings as something to leave alone unless you deliberately want to add a second Map later."],
      },
      {
        title: "Structure: group everything triggered by the same action in one place",
        notes: ["Each direction's own Input Trigger script already holds that direction's movement, sound, and recolor (which also drives that direction's share of the Animate blink), with the shared idle-reset living in the Released scripts rather than duplicated per direction. Clamp needs no script anywhere - it's a Map Properties setting, not an event - so the Complete Level 3 solution below only has the 8 Input Trigger scripts (4 Pressed, 4 Released) plus a note about where Clamp actually lives."],
      },
      {
        title: "Complete Level 3 solution (all 5 goals combined)",
        notes: ["Everything above combined - one script per direction, plus one boundary script:"],
        code: [
          "Input Trigger: Right Arrow (Pressed) ->",
          "  Position Actor(main, X: main.X + 8, Y: main.Y)",
          "  Play Sound(sfx_blip.wav)",
          "  Change Body/Clothing Sprite(main, Sprite: player_circle_red)",
          "",
          "Input Trigger: Left Arrow (Pressed) ->",
          "  Position Actor(main, X: main.X - 8, Y: main.Y)",
          "  Play Sound(sfx_blip.wav)",
          "  Change Body/Clothing Sprite(main, Sprite: player_circle_blue)",
          "",
          "Input Trigger: Up Arrow (Pressed) ->",
          "  Position Actor(main, X: main.X, Y: main.Y - 8)",
          "  Play Sound(sfx_blip.wav)",
          "  Change Body/Clothing Sprite(main, Sprite: player_circle_green)",
          "",
          "Input Trigger: Down Arrow (Pressed) ->",
          "  Position Actor(main, X: main.X, Y: main.Y + 8)",
          "  Play Sound(sfx_blip.wav)",
          "  Change Body/Clothing Sprite(main, Sprite: player_circle_yellow)",
          "",
          "Input Trigger: Right/Left/Up/Down Arrow (Released - one Control Combination per key, in addition to each key's Pressed combination above) ->",
          "  Multi-Control Branch(Right Arrow, Left Arrow, Up Arrow, Down Arrow)",
          "    any path where a key is still pressed -> (no action - stay in that direction's sprite)",
          "    \"none pressed\" path -> Change Body/Clothing Sprite(main, Sprite: player_circle_orange)",
          "",
          "(This same Released/Multi-Control Branch pair is also what makes main blink while walking - see the Animate idea above for why. No separate animation event needed.)",
          "",
          "(Clamp needs no script at all - Map Properties > Boundaries is ticked on all 4 sides by default with no Connections set, so 001 Game Creator already refuses to let main cross the edge. See the Clamp idea above.)",
        ],
      },
    ],
    altNotes: {
      title: "Textual Scripting",
      notes: [
        "Reiterate to students in person: this mode switch is genuinely irreversible per-project. Confirm they created a separate throwaway project before letting them proceed.",
        "Exact text-command names should be checked against the installed version's Script Editor autocomplete or the wiki before class, since this course otherwise treats 001 Game Creator as visual-only and hasn't needed to track Textual syntax changes.",
      ],
      why: [
        "This path exists for students who've realized they prefer typing code (maybe from Godot or GameMaker) and want to see whether 001 Game Creator supports that style too - it does, with this one important caveat.",
        "Because of the irreversibility, this should stay an optional, teacher-supervised exploration rather than a path any student defaults into.",
      ],
    },
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    l1Notes: ["Requires Setup's full project configuration already done: Lumen off, default Light/Sky actors deleted, BP_Player created (parent Static Mesh Actor) with its Static Mesh set to Cube and M_Player_Orange (Unlit, a Texture Sample from the imported player_circle_orange.png wired to Emissive Color - the real sprite, not a flat color) applied, Mobility changed from Static to Movable on BP_Player's class default, Auto Receive Input set to Player 0 in BP_Player's Class Defaults, BP_NoPawnGameMode set as GameMode Override (Default Pawn Class = None), the Orthographic Camera placed and Auto-Activated for Player 0, and IA_MoveRight/Left/Up/Down + IMC_Default (all 4 keys mapped) built, with IMC_Default applied via Add Mapping Context in BP_Player's own BeginPlay. Watch for students who forget Auto Receive Input (the #1 miss now - a plain Actor Blueprint never receives keyboard input without it, and Add Actor World Offset's Target pin can be left unwired throughout, since inside BP_Player's own graph that correctly defaults to Self), who leave the Enhanced Input Action Event on its default Started pin instead of Triggered (moves once per press instead of gliding while held), who forget to right-click Delta Location and choose Split Struct Pin before trying to wire a value into just X or Y, or who forget BP_Player's Mobility (a Static component silently ignores every movement call). BP_Player's own Event Graph:"],
    l1Code: ["Enhanced Input Action Event (IA_MoveRight), Triggered -> Add Actor World Offset(Target unwired = Self, DeltaLocation = (X = 300 * Get World Delta Seconds, Y=0, Z=0))"],
    l2Notes: ["BP_Player's own Event Graph, 4 Enhanced Input Action Event nodes (one per Input Action, all on their Triggered pin) each feeding an Add Actor World Offset node with Target left unwired (defaults to Self), Delta Location scaled by Get World Delta Seconds x Speed (300) on every direction so movement stays smooth and framerate-independent:"],
    l2Code: [
      "Enhanced Input Action Event (IA_MoveRight), Triggered -> Add Actor World Offset(Self, DeltaLocation = (X = 300 * Get World Delta Seconds, Y=0, Z=0))",
      "Enhanced Input Action Event (IA_MoveLeft),  Triggered -> Add Actor World Offset(Self, DeltaLocation = (X = -300 * Get World Delta Seconds, Y=0, Z=0))",
      "Enhanced Input Action Event (IA_MoveUp),    Triggered -> Add Actor World Offset(Self, DeltaLocation = (X=0, Y = 300 * Get World Delta Seconds, Z=0))",
      "Enhanced Input Action Event (IA_MoveDown),  Triggered -> Add Actor World Offset(Self, DeltaLocation = (X=0, Y = -300 * Get World Delta Seconds, Z=0))",
    ],
    l3Ideas: [
      {
        title: "Animate: blink while moving, static when idle",
        notes: ["Toggling visibility is the Blueprint equivalent of the counter-based blink used in the code-path engines - since BP_Player's mesh is a 3D Shape rather than a 2D sprite, Set Actor Hidden In Game is the closest match to SpriteRenderer.enabled or Godot's visible. This needs an IsMoving state that's true while any direction key is held, set from the Started/Completed pins of the 4 Enhanced Input Action Events - see the Structure idea below for why it's split this way:"],
        code: [
          "Event Tick ->",
          "  Branch (IsMoving)",
          "    True:  BlinkTimer += DeltaSeconds",
          "           Branch (BlinkTimer >= 0.1) -> True: Blink = NOT Blink; BlinkTimer = 0",
          "           Set Actor Hidden In Game(Self, Blink)",
          "    False: Set Actor Hidden In Game(Self, false)",
          "           BlinkTimer = 0",
        ],
      },
      {
        title: "Sound: repeating beep-beep while held, silent when idle",
        notes: ["Play Sound 2D is fire-and-forget - it spawns a sound with no handle to check whether it's still playing, unlike Unity's AudioSource.isPlaying or Godot's AudioStreamPlayer2D.playing. To get the same repeating \"beep-beep while held\" rhythm those engines use (rather than one blip per press), gate it behind a SoundCooldown float counted down by Get World Delta Seconds on the Triggered pin - the same cooldown-timer idiom used for fire-rate limiting elsewhere in this course:"],
        code: [
          "Enhanced Input Action Event (IA_MoveRight), Triggered ->",
          "  Add Actor World Offset(Self, DeltaLocation = (X = 300 * Get World Delta Seconds, Y=0, Z=0))",
          "  SoundCooldown -= Get World Delta Seconds",
          "  Branch (SoundCooldown <= 0) -> True: Play Sound 2D(sfx_blip.wav); Set SoundCooldown = 0.15",
        ],
      },
      {
        title: "Recolor: different material per direction, reset when idle",
        notes: ["Import player_circle_red.png, player_circle_blue.png, player_circle_green.png, and player_circle_yellow.png (this lesson's Assets folder) and build one Unlit Material per color the same way Setup built M_Player_Orange - a Texture Sample from that PNG wired straight into Emissive Color. Each color is its own separate, pre-made Material (the real supplied sprite, not a flat color standing in for it), swapped onto BP_Player with a Set Material node off the Started pin - one swap per press-down is all that's needed, no reason to re-swap every frame while held. If Set Material doesn't appear in the search, call Get Static Mesh Component first (Target left unwired, defaults to Self), or uncheck \"Context Sensitive\" in the search popup - a known Blueprint search quirk. Same Idle/Moving concept as the State Diagram from this lesson's bonus section:"],
        code: [
          "Enhanced Input Action Event (IA_MoveRight), Started -> Set Material(Self, Material Index 0, M_Player_Red)",
          "Enhanced Input Action Event (IA_MoveLeft),  Started -> Set Material(Self, Material Index 0, M_Player_Blue)",
          "Enhanced Input Action Event (IA_MoveUp),    Started -> Set Material(Self, Material Index 0, M_Player_Green)",
          "Enhanced Input Action Event (IA_MoveDown),  Started -> Set Material(Self, Material Index 0, M_Player_Yellow)",
          "Event Tick, Branch(IsMoving) False -> Set Material(Self, Material Index 0, M_Player_Orange)  // idle",
        ],
      },
      {
        title: "Clamp movement to a play area",
        notes: ["After each direction's Triggered-driven Add Actor World Offset, chain: Get Actor Location -> Break Vector -> Clamp (Float) on X and Y against your chosen min/max (half the Camera's Ortho Width from Setup is a reasonable starting bound) -> Make Vector -> Set Actor Location. Clamping after the fact is simpler to explain than pre-checking bounds, and is a common real pattern in Blueprint movement code. Because movement now runs every frame while held, this check also re-runs every frame while held, keeping BP_Player inside a defined rectangle no matter how long a key stays down."],
      },
      {
        title: "Structure: group everything triggered by the same action in one place",
        notes: ["Unreal Blueprint genuinely has three different execution contexts here, and that's correct, not a mistake to avoid: each direction's movement, Clamp check, and cooldown-gated sound all chain off that direction's own Enhanced Input Action Event's Triggered pin (continuous, every frame while held), that same event's recolor and IsMoving = true chain off its Started pin (discrete, once per press), and the blink timer lives in a single Event Tick (continuous, every frame) reading a shared IsMoving boolean that the 4 Enhanced Input Action Events' Started/Completed pins set. One Enhanced Input Action Event node can drive different pins for different purposes - don't re-swap the Material every frame off Triggered (no reason to, it already matches), and don't try to drive the blink from an Enhanced Input Action Event at all. That's the Complete Level 3 solution below."],
      },
      {
        title: "Complete Level 3 solution (all 5 goals combined)",
        notes: ["Everything above combined, across BP_Player's own Enhanced Input Action Events and one Event Tick - every Target left unwired throughout, since inside BP_Player's own graph that always means Self:"],
        code: [
          "(M_Player_Red, M_Player_Blue, M_Player_Green, M_Player_Yellow, M_Player_Orange created ahead of time in the Content Browser, each Unlit with its own Texture Sample wired to Emissive Color)",
          "",
          "Enhanced Input Action Event (IA_MoveRight):",
          "  Triggered -> Add Actor World Offset(Self, DeltaLocation = (X = 300 * Get World Delta Seconds, Y=0, Z=0))",
          "               Get Actor Location -> Break Vector -> Clamp (Float) X and Y -> Make Vector -> Set Actor Location",
          "               SoundCooldown -= Get World Delta Seconds",
          "               Branch (SoundCooldown <= 0) -> True: Play Sound 2D(sfx_blip.wav); Set SoundCooldown = 0.15",
          "  Started   -> Set Material(Self, Material Index 0, M_Player_Red)",
          "               Set IsMoving = true",
          "  Completed -> Set IsMoving = false",
          "",
          "Enhanced Input Action Event (IA_MoveLeft):",
          "  Triggered -> Add Actor World Offset(Self, DeltaLocation = (X = -300 * Get World Delta Seconds, Y=0, Z=0))",
          "               Get Actor Location -> Break Vector -> Clamp (Float) X and Y -> Make Vector -> Set Actor Location",
          "               SoundCooldown -= Get World Delta Seconds",
          "               Branch (SoundCooldown <= 0) -> True: Play Sound 2D(sfx_blip.wav); Set SoundCooldown = 0.15",
          "  Started   -> Set Material(Self, Material Index 0, M_Player_Blue)",
          "               Set IsMoving = true",
          "  Completed -> Set IsMoving = false",
          "",
          "Enhanced Input Action Event (IA_MoveUp):",
          "  Triggered -> Add Actor World Offset(Self, DeltaLocation = (X=0, Y = 300 * Get World Delta Seconds, Z=0))",
          "               Get Actor Location -> Break Vector -> Clamp (Float) X and Y -> Make Vector -> Set Actor Location",
          "               SoundCooldown -= Get World Delta Seconds",
          "               Branch (SoundCooldown <= 0) -> True: Play Sound 2D(sfx_blip.wav); Set SoundCooldown = 0.15",
          "  Started   -> Set Material(Self, Material Index 0, M_Player_Green)",
          "               Set IsMoving = true",
          "  Completed -> Set IsMoving = false",
          "",
          "Enhanced Input Action Event (IA_MoveDown):",
          "  Triggered -> Add Actor World Offset(Self, DeltaLocation = (X=0, Y = -300 * Get World Delta Seconds, Z=0))",
          "               Get Actor Location -> Break Vector -> Clamp (Float) X and Y -> Make Vector -> Set Actor Location",
          "               SoundCooldown -= Get World Delta Seconds",
          "               Branch (SoundCooldown <= 0) -> True: Play Sound 2D(sfx_blip.wav); Set SoundCooldown = 0.15",
          "  Started   -> Set Material(Self, Material Index 0, M_Player_Yellow)",
          "               Set IsMoving = true",
          "  Completed -> Set IsMoving = false",
          "",
          "Event Tick ->",
          "  Branch (IsMoving)",
          "    True:  BlinkTimer += DeltaSeconds",
          "           Branch (BlinkTimer >= 0.1) -> True: Blink = NOT Blink; BlinkTimer = 0",
          "           Set Actor Hidden In Game(Self, Blink)",
          "    False: Set Actor Hidden In Game(Self, false)",
          "           BlinkTimer = 0",
          "           Set Material(Self, Material Index 0, M_Player_Orange)  // idle",
        ],
      },
    ],
    altNotes: {
      title: "C++",
      notes: [
        "Forgetting \"EnhancedInput\" in Build.cs's PublicDependencyModuleNames is the #1 thing students miss - the code compiles fine (UInputAction/UEnhancedInputComponent are still valid types), but the subsystem silently isn't registered, so nothing fires at runtime. If Enhanced Input C++ 'just doesn't work' with no compile error, check Build.cs first.",
        "The C++ path reuses the exact same IA_MoveRight/Left/Up/Down and IMC_Default assets the Blueprint path built in Setup - Input Actions and Mapping Contexts are just data assets, not tied to either track, which is why a student can freely mix a Blueprint Pawn with C++ input handling or vice versa.",
        "AddMappingContext() in C++ is the code equivalent of the Blueprint path's Add Mapping Context node - both call into the same UEnhancedInputLocalPlayerSubsystem, so they're genuinely the same operation, not two different systems that happen to look similar.",
        "The Level 3 goals map onto C++ calls the same way they map onto Blueprint nodes: SetActorHiddenInGame() in Tick() for the blink, UGameplayStatics::PlaySound2D() for sound, SetMaterial() on the Cube's mesh component (swapping between the same 5 pre-made Material assets used in the Blueprint path) for recolor, and FMath::Clamp() for the boundary check - all called from the same MoveRight()-style functions and a Tick() override.",
      ],
      why: [
        "Showing both C++ and Blueprint side by side for the same behavior is the clearest way to demystify Blueprints for students who assume \"visual\" means \"not real code\" - it's the same engine calls either way, now true of the input layer too since both tracks use Enhanced Input from Level 1 onward instead of diverging.",
        "This is squarely the optional advanced track - most students should stay on Blueprints; C++ is for students already comfortable with typed languages from Godot, GameMaker, or Unity's C# path.",
      ],
    },
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_02_Key_${d.file}.docx`]));

Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
