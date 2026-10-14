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
// PREP CHECKLIST
// =========================================================================
{
  const doc = new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 4 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 4 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("Self-Contained Lesson Assets"),
        p("Per Lesson Production Standards Rule 3, nothing this lesson needs should require a student to search for or download something mid-class. Everything below is already prepared and included in this lesson's Assets folder - no download required."),
        checkbox("New placeholder enemy sprite already generated (enemy_invader_red.png, 128x128 with transparency, a pixel-art \"invader\" silhouette in red) - distinct in color and shape from the player and bullet sprites."),
        checkbox("New sound effect already generated (sfx_hit.wav, a short impact/explosion burst - deliberately distinct from Lesson 3's sfx_shoot.wav so firing and hitting don't sound the same) for the Level 3 \"add a sound\" idea."),
        checkbox("Lesson 2 and 3's Assets folders (player sprites, bullet_yellow.png, sfx_blip.wav, sfx_shoot.wav) are still needed too, since this lesson builds on the same ongoing [YourName]Game project rather than starting fresh."),
        checkbox("Copy this lesson's Assets folder onto the shared classroom drive / USB backups alongside the handouts."),

        h2("Continuity From Lesson 3"),
        p("Same pattern as last week - students open their existing project rather than starting fresh."),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 3, not create a new one - the enemy and collision detection are added on top of the shooting mechanic they already built."),
        checkbox("If a student's Lesson 3 project is lost or broken, have a pre-built \"Lesson 3 complete\" starter available per engine (from the Lesson 3 Teacher's Keys) so they aren't blocked."),

        h2("Engine Verification"),
        p("This week's core new concept - collision/overlap detection - works differently enough across engines that it's worth a quick personal test run before class, even under the \"charge through\" pacing this course is using for Lessons 4-15."),
        checkbox("Confirm the demo laptop's Lesson 2 and 3 exercises still work for all 8 engines - re-run any that haven't been touched in a few weeks."),
        checkbox("Pre-build the full Scratch collision exercise on top of a finished Lesson 3 project, so you have a working example ready to show."),
        checkbox("Godot: confirm an Area2D's area_entered signal actually fires between two Area2D nodes (not body_entered, which is for PhysicsBody2D/TileMap nodes instead) in your installed version - this is the one new physics concept this lesson introduces."),
        checkbox("Unity: confirm a Rigidbody2D (Body Type = Kinematic) plus two Is Trigger colliders actually fires OnTriggerEnter2D in your installed Editor version - test this once before class, since a missing Rigidbody2D on either object is the most common silent failure."),
        checkbox("GameMaker: quick-test a Collision event between two throwaway objects in a test room before class."),
        checkbox("GDevelop: confirm the built-in \"is in collision with\" condition is available between two Sprite objects in your installed version (it should be, by default)."),
        checkbox("TIC-80: there's no built-in collision system here - confirm the rectangle-overlap math in the Key matches your chosen sprite size if you use anything other than the default 8x8 sprites."),
        checkbox("001 Game Creator (Windows demo laptop): confirm the \"Collided with actor\" trigger and the relate keyword are available and behave as documented in the Actor Trigger editor for your installed version - this is a different trigger than Lesson 3 used."),
        checkbox("If any advanced-track students are on Unreal this week, practice setting Collision Presets to OverlapAllDynamic on a Shape component once so it's a familiar click during class - the default BlockAllDynamic preset silently blocks Event ActorBeginOverlap from ever firing, and it's easy to mistake for a wiring problem."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies, Assets folder (on USB as backup)."),

        h2("Known Gotchas This Week"),
        bullet("Unity: trigger events need a Rigidbody2D on at least one of the two colliding objects - this is the single most common \"nothing happens\" bug this week, since it's easy to add colliders and forget the Rigidbody2D entirely."),
        bullet("Godot: Area2D only detects other Area2D nodes via area_entered (or PhysicsBody2D nodes via body_entered) - a plain Sprite2D with no Area2D/CollisionShape2D child is invisible to collision detection entirely, which is why both bullet.tscn and enemy.tscn need their own Hitbox this week. Separately, this is the first week students position something themselves in the 2D editor (dragging Enemy into main.tscn) - Godot's origin is the top-left corner with Y increasing downward, so a student used to a different convention can place Enemy well off-screen with no error at all. It'll still show up in the Remote tab while running (proving it exists), just not on screen - point them at Transform > Position in the Inspector, not the Texture, first."),
        bullet("Unreal: a fresh Shape's default Collision Preset (BlockAllDynamic) makes it physically block other dynamic objects instead of generating an overlap event - if Event ActorBeginOverlap never fires, check Collision Presets on BOTH BP_Enemy and BP_Bullet before suspecting the node wiring. A second, separate snag: BP_Bullet's Add Actor World Offset movement node from Lesson 3 needs its Sweep box checked, or a fast-moving bullet can pass straight through the enemy each frame without the move ever testing for what it passed into."),
        bullet("001 Game Creator: non-Solid Actors generate no collisions at all - if an enemy or bullet was created with \"Solid\" unchecked (or inherited that setting from its Template), nothing will ever detect it."),
        bullet("TIC-80: there's no built-in collision system - the rectangle-overlap check in the Key is deliberately simple (axis-aligned bounding box), which is normal for this engine and not a shortcut being taken."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_04_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 4: Direct Hit", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: detecting when things touch. Your job is to add one enemy to the screen and make your bullet destroy it on contact - the first piece of your capstone's actual combat loop."),
    p("Open your own [YourName]Game project from Lesson 3 - don't start a new one. Everything this week builds on the player and bullet you already have.", { bold: true }),
    p("Tip: this lesson's Assets folder has a ready-to-use placeholder enemy image (enemy_invader_red.png) and a new sound effect (sfx_hit.wav) - you don't need to find or make your own art today.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("A single enemy sits on screen. When your bullet touches it, both the bullet and the enemy disappear. This is a complete finish for today - reaching this is enough to move on to next week.", { italics: true }),
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
    p("The enemy is destroyed cleanly and exactly once, even if you fire several bullets in a row - no errors, no lingering references, and a bullet that misses still behaves exactly like it did last week.", { italics: true }),
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
      headers: { default: brand.pageHeader("Lesson 4 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 3.",
      "Create a new sprite for the enemy - import enemy_invader_red.png from this lesson's Assets folder (Upload Sprite).",
      "Rename this sprite Enemy. Add a \"when green flag clicked\" script to it with \"go to x: 0 y: 150\" (Motion) and \"show\" (Looks), so it starts in a fixed spot near the top of the stage.",
    ],
    level1: [
      "Switch to the Bullet sprite's \"when I start as a clone\" script (from Lesson 3).",
      "Inside the same forever loop that moves the bullet upward, add \"if <touching [Enemy]?> then\" (Control + Sensing), and put \"broadcast [enemy hit]\" and \"delete this clone\" inside it - check touching before the existing off-screen check, so a hit near the top of the stage is still caught.",
      "Click the New Message option in the broadcast dropdown to create the \"enemy hit\" message the first time.",
      "Switch to the Enemy sprite, click Events, and drag \"when I receive [enemy hit]\" into a new script, with \"hide\" underneath it.",
      "Click the stage, then press spacebar aimed at the Enemy to test - the enemy should vanish and the bullet that hit it should disappear too.",
    ],
    level2: [
      "Fire several bullets in a row, including some that miss - each hit should make the enemy vanish exactly once, and every bullet that misses should keep flying and disappear off-screen exactly like it did in Lesson 3.",
      "This already works with no extra code: Scratch's \"touching\" check always returns false for a hidden sprite, so once the Enemy sprite is hidden, no later bullet clone can trigger a second \"enemy hit\" broadcast, even with several bullets in flight at once.",
    ],
    level3: [
      "Sound: upload sfx_hit.wav from this lesson's Assets folder, then add \"start sound [sfx_hit]\" to the Enemy's \"when I receive [enemy hit]\" script, right before \"hide\".",
      "Visual: give the Enemy its own distinct color using the Costume editor's Fill tool, or add a \"set color effect to\" block (Looks) in its green-flag script, so it clearly reads as a threat and not just another shape on screen.",
      "Structure: the Bullet's forever loop should only ever detect a hit and report it (touching check + broadcast + delete this clone) - everything about how the Enemy responds to being hit (sound, hiding) belongs on the Enemy's own \"when I receive [enemy hit]\" script. Keeping this boundary clean is what let the Sound idea above drop into the Enemy sprite alone.",
    ],
    trouble: [
      "Bullet passes through the Enemy without anything happening: confirm \"touching [Enemy]?\" actually targets the Enemy sprite by name in the dropdown (it can default to something else), and that the check is inside the bullet's forever loop, not outside it.",
      "Enemy disappears but a leftover bullet keeps flying: confirm \"delete this clone\" runs in the same if-branch as the touching check, right alongside the broadcast.",
      "Nothing happens even up close: confirm the Enemy sprite's own \"when green flag clicked\" script actually runs \"show\" - a sprite hidden from a previous test stays hidden until the green flag is clicked again.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 3.",
      "Open bullet.tscn and add an Area2D child to the Bullet root, rename it Hitbox, then add a CollisionShape2D child to Hitbox with a new RectangleShape2D sized to roughly match bullet_yellow.png.",
      "Create a new scene: a Sprite2D root node, rename it Enemy, and assign enemy_invader_red.png (from this lesson's Assets folder) as its Texture. Add an Area2D child, rename it Hitbox, then add a CollisionShape2D child to it with a new RectangleShape2D sized to roughly match enemy_invader_red.png.",
      "Save this new scene as enemy.tscn. Attach a new script to the Enemy root, extends Sprite2D.",
      "Open your main scene and drag an instance of enemy.tscn into it, positioning it near the top of the play area - Godot's origin (0,0) is the top-left corner, and Y increases downward, so \"near the top\" means a small positive Y (something like 50-150), not a negative one or a value based on the screen's center.",
    ],
    level1: [
      "Connect Hitbox's area_entered signal to the Enemy so it can respond when something touches it - two ways to do this, pick one, not both: (a) in the editor, select Hitbox, open the Signals tab (grouped with Inspector and Groups), double-click area_entered, choose Enemy as the receiver in the dialog, and accept the suggested method name - Godot writes the function stub into Enemy's script for you; or (b) skip the editor entirely and just write $Hitbox.area_entered.connect(_on_hitbox_area_entered) yourself inside _ready() (code below). Using both at once double-connects the signal, so the handler fires twice per hit.",
      "Run the scene and shoot the Enemy to test - both the bullet and the Enemy should disappear on contact.",
    ],
    level1Code: [
      "# enemy.tscn's script (attached to the Sprite2D root):",
      "extends Sprite2D",
      "",
      "func _ready():",
      "    $Hitbox.area_entered.connect(_on_hitbox_area_entered)",
      "",
      "func _on_hitbox_area_entered(area):",
      "    area.get_parent().queue_free()",
      "    queue_free()",
    ],
    level2: [
      "Fire several bullets in a row, including some that miss - each hit destroys exactly one bullet and the Enemy, and any bullet that misses keeps flying and disappears off-screen exactly like it did in Lesson 3.",
      "queue_free() is safe to call more than once on the same node (a repeat call is simply ignored), so even the rare case of two bullets overlapping the Enemy on the very same frame doesn't cause an error - both bullets are destroyed and the Enemy is destroyed exactly once.",
    ],
    level3: [
      "Sound: add an AudioStreamPlayer2D node to the Enemy scene (named Hit), Stream set to sfx_hit.wav. Freeing the Enemy immediately would cut the sound off before it finishes, so hide the Enemy and turn off its Hitbox right away, play the sound, then await its finished signal before actually freeing the node.",
      "Visual: give the Enemy its own modulate color (distinct from the bullet's yellow) in its _ready(), so it clearly reads as a threat rather than another shot.",
      "Structure: everything about how the Enemy responds to being hit - destroying the bullet, hiding itself, playing a sound - lives inside enemy.tscn's own script. bullet.tscn's own script doesn't need to know the Enemy exists at all; it only needed a Hitbox added to it so the Enemy's Area2D could detect it.",
    ],
    trouble: [
      "Nothing happens on contact: confirm both bullet.tscn and enemy.tscn actually have a CollisionShape2D with a real shape assigned inside their Hitbox (an Area2D with no shape can't detect anything), and that neither Hitbox's Monitoring property was left off.",
      "Godot reports an error on area.get_parent(): confirm the signal connected is Hitbox's own area_entered, not the Enemy Sprite2D root's - Sprite2D nodes don't have that signal at all.",
      "Enemy disappears immediately but the bullet lingers a frame or two: this is expected - queue_free() defers actual deletion to the end of the current frame, not instantly.",
      "Enemy doesn't appear anywhere on screen, even though it shows up in the Scene tree (or the Remote tab while the game is running): before checking the Texture, check Transform > Position in the Inspector. Godot's origin is the top-left corner with Y increasing downward, so a position guessed from a different convention (screen-center origin, Y-up) can land the Enemy well off the top or side of the visible viewport with no error at all. Reposition it directly in the 2D editor - drag it in the viewport, or edit Position in the Inspector - rather than repositioning it from code in _ready(). That's the preferred fix here, not just a shortcut: it keeps what you see in the editor exactly matching your game's actual start position, the same way Player's position already works.",
      "Enemy's hit response seems to run twice (bullet and Enemy both vanish correctly, but a later Level 3 sound plays twice, or a score/lives change in a future lesson is off by double): the signal got connected two ways at once - both through the Signals tab in the editor and through a $Hitbox.area_entered.connect() call in _ready(). Select Hitbox and check its Signals tab; if area_entered already lists a connection to Enemy there, remove the code call (or vice versa) so only one connection exists.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Open your [YourName]Game project from Lesson 3.",
      "On the Bullet prefab, add a Box Collider 2D component and tick Is Trigger. Add a Rigidbody2D too, Body Type set to Kinematic - Unity's 2D physics needs at least one Rigidbody2D between two colliding objects for trigger events to fire, and Bullet is the one already moving each frame.",
      "Create a new empty GameObject, add a Sprite Renderer and assign enemy_invader_red.png (import it from this lesson's Assets folder) as its Sprite. Add a Box Collider 2D and tick Is Trigger.",
      "In the Inspector's Tag dropdown, click Add Tag, create a new tag named Enemy, then assign it to this GameObject.",
      "Drag this GameObject into the Project window to make it a Prefab named Enemy, then position an instance of it in the scene near the top of the play area.",
    ],
    level1: [
      "In Bullet.cs, add an OnTriggerEnter2D(Collider2D other) method that checks other.CompareTag(\"Enemy\") and, if true, destroys both the other object and itself (code below).",
      "Press Play and shoot the Enemy to test - both the bullet and the Enemy should disappear on contact.",
    ],
    level1Code: [
      "// Bullet.cs - add this method:",
      "void OnTriggerEnter2D(Collider2D other) {",
      "    if (other.CompareTag(\"Enemy\")) {",
      "        Destroy(other.gameObject);",
      "        Destroy(gameObject);",
      "    }",
      "}",
    ],
    level2: [
      "Fire several bullets in a row, including some that miss - each hit destroys exactly one bullet and the Enemy, and any bullet that misses keeps flying and destroys itself off-screen exactly like it did in Lesson 3.",
      "Destroy() doesn't remove a GameObject instantly - it happens at the end of the current frame - so if two bullets somehow trigger on the exact same frame, both destroy calls on the Enemy are harmless; Unity simply ignores the second one, and no error is thrown.",
    ],
    level3: [
      "Sound: import sfx_hit.wav, add a public AudioClip hitSound field to Bullet.cs, and call AudioSource.PlayClipAtPoint(hitSound, transform.position) inside OnTriggerEnter2D before either object is destroyed - this static method spawns its own temporary audio source that survives independently of the Enemy or Bullet being destroyed, so the sound always finishes playing.",
      "Visual: give the Enemy prefab its own SpriteRenderer.color (distinct from the bullet's yellow) in a small new Enemy.cs script's Start(), so it clearly reads as a threat rather than another shot.",
      "Structure: OnTriggerEnter2D in Bullet.cs should only ever detect a hit and clean up (destroy the enemy, destroy itself, play the sound) - anything about how the Enemy looks on its own belongs in its own Enemy.cs, not bolted onto Bullet.cs.",
    ],
    trouble: [
      "Nothing happens on contact: confirm both colliders have Is Trigger checked, and that the Bullet prefab actually has a Rigidbody2D - two Collider2Ds with no Rigidbody2D between them never generate trigger events in Unity's 2D physics.",
      "Console error about a missing tag: confirm the Enemy tag was actually created and assigned (Inspector's Tag dropdown, top-left) on the Enemy prefab itself, not just typed as a string somewhere.",
      "Bullet destroys itself on everything, not just the Enemy: double-check the CompareTag(\"Enemy\") check - without it, OnTriggerEnter2D fires for anything the bullet's collider touches.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same trigger-detection pattern works in a Script Graph instead of C# - Unity's Visual Scripting graphs compile to the same underlying calls.",
      steps: [
        "On the Bullet prefab, add an On Trigger Enter 2D event node, feeding a Branch node that checks the other Collider's Compare Tag against \"Enemy\".",
        "On True, wire a Destroy node targeting the other Collider's Game Object, and a second Destroy node targeting this Game Object.",
        "Set up the Bullet's Rigidbody2D and both colliders' Is Trigger exactly as in the C# path - the physics requirements are the same no matter which scripting style triggers the event.",
      ],
      note: "Level 3's ideas are all reachable with more nodes: an Audio Source \"Play Clip At Point\" node for sound, a \"Set Member\" node on the Sprite Renderer's color for the Enemy's own visual idea - same ideas as the C# path, just wired instead of typed.",
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: [
      "Open your [YourName]Game project from Lesson 3.",
      "Create a new sprite spr_enemy by importing enemy_invader_red.png from this lesson's Assets folder.",
      "Create a new object obj_enemy, assign spr_enemy to it.",
      "Drag one obj_enemy instance into your room near the top of the play area.",
    ],
    level1: [
      "On obj_bullet, add a Collision event with obj_enemy (Add Event > Collision > obj_enemy) with the code below - it destroys both the enemy it hit and itself.",
      "Press F5 (Run) and shoot the enemy to test - both the bullet and the enemy should disappear on contact.",
    ],
    level1Code: [
      "// obj_bullet's Collision event with obj_enemy:",
      "instance_destroy(other);",
      "instance_destroy(self);",
    ],
    level2: [
      "Fire several bullets in a row, including some that miss - each hit destroys exactly one bullet and the enemy, and any bullet that misses keeps flying and destroys itself off-screen exactly like it did in Lesson 3.",
      "GameMaker's Collision event already only fires once per pair of instances actually overlapping that step, and instance_destroy() is safe to call on an instance that's already been destroyed earlier in the same step (it's simply ignored) - so even the rare case of two bullets hitting the same enemy on the same step doesn't error.",
    ],
    level3: [
      "Sound: add sfx_hit.wav as a sound resource (keep its default name, sfx_hit), then call audio_play_sound(sfx_hit, 1, false) at the top of the Collision event, before either instance_destroy() call.",
      "Visual: give obj_enemy its own image_blend color in its Create event, distinct from the bullet's yellow, so it clearly reads as a threat rather than another shot.",
      "Structure: obj_bullet's Collision event should only ever detect a hit and clean up (destroy the enemy, destroy itself, play the sound) - anything about how the enemy looks on its own belongs in obj_enemy's own Create event, not bolted onto obj_bullet.",
    ],
    trouble: [
      "Nothing happens on contact: confirm the Collision event on obj_bullet is specifically targeting obj_enemy (Add Event > Collision lets you pick which object), not left as a generic collision.",
      "Errors mentioning other: other only has a meaningful value inside a Collision event - pasting this code into the Step event instead will error or silently do nothing.",
      "Bullet destroys the enemy but nothing seems to change: confirm obj_enemy was actually placed in the room (an object with zero instances can never collide with anything), and that instance_destroy(other) is really inside the Collision event.",
    ],
    resource: "Student Resource List > GameMaker section, or gamemaker.io/en/tutorials",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same collision-detection pattern works with drag-and-drop blocks instead of typed code. Pick whichever path feels right for you - both are graded the same.",
      steps: [
        "On obj_bullet, add a Collision event with obj_enemy, same as the code path.",
        "From the Instance category, drag two \"Instance - Destroy Instance\" blocks into it - set the first block's Instance dropdown to Other, and the second to Self.",
        "Press F5 (Run) and test the same way as the code path.",
      ],
      note: "Level 3's ideas map onto blocks the same way, favoring the most specific named block available: \"Play Audio\" (Audio category) for sound, and \"Set Instance Variable\" -> Image Blend for the enemy's own visual idea - Image Blend is a recognized built-in on Set Instance Variable's own dropdown, so prefer it over the generic Assign Variable/Set Variable.",
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 3.",
      "Add a new Sprite object, name it Enemy, and assign enemy_invader_red.png (import it from this lesson's Assets folder) as its image.",
      "Place one instance of Enemy on the scene near the top of the play area.",
    ],
    level1: [
      "Add a new event: Condition \"Bullet is in collision with Enemy\" -> Actions \"Delete Bullet\" and \"Delete Enemy\".",
      "Click the green Preview button and shoot the Enemy to test - both the bullet and the Enemy should disappear on contact.",
    ],
    level1Code: [
      "Condition: Bullet is in collision with Enemy",
      "Action: Delete Bullet",
      "Action: Delete Enemy",
    ],
    level2: [
      "Fire several bullets in a row, including some that miss - each hit deletes exactly one Bullet instance and the Enemy, and any bullet that misses keeps flying and is destroyed off-screen exactly like it did in Lesson 3 (the Destroy outside screen behavior).",
      "GDevelop's collision condition and Delete actions already operate per pair of overlapping instances each frame, so even if two bullets somehow overlap the Enemy on the same frame, the event simply runs twice that frame - the Enemy and both bullets are still deleted cleanly either way.",
    ],
    level3: [
      "Sound: add a \"Play a sound\" action (using the imported sfx_hit.wav) to the same event as the collision condition.",
      "Visual: give Enemy its own \"Tint color\" (distinct from the bullet's yellow), applied once at the start of the scene, so it clearly reads as a threat rather than another shot.",
      "Structure: group this lesson's event (the collision condition and its actions) into its own labeled event group (e.g. \"Combat\"), separate from the \"Shooting\" group from Lesson 3.",
    ],
    trouble: [
      "Nothing happens on contact: confirm the collision condition is comparing Bullet against Enemy (not Player against Enemy), and that both objects are Sprite objects with an actual image assigned.",
      "Enemy disappears but a bullet keeps flying: confirm \"Delete Bullet\" is in the same event as the collision condition, not a separate always-running event.",
      "Nothing on screen at all: confirm at least one Enemy instance was actually placed on the scene, not just created as an object type - an object with zero instances can never collide with anything.",
    ],
    resource: "Student Resource List > GDevelop section, or gdevelop.io/tutorials",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: [
      "Launch TIC-80 and load your cart from Lesson 3.",
      "Open the Sprite Editor (F2) and draw a small enemy shape on sprite index 3 (index 1 is your player, index 2 is your bullet).",
    ],
    level1: [
      "Press Esc to reach the code editor.",
      "Add enemy_alive=true and fixed ex, ey coordinates near the top of your code, above the TIC() function - this is the enemy's one and only position for this lesson.",
      "Inside your bullets loop (from Lesson 3), before the existing off-screen check, add a simple rectangle-overlap check against ex/ey - if it overlaps and the enemy is still alive, set enemy_alive=false and remove that bullet from the table.",
      "At the bottom of TIC(), draw the enemy only if enemy_alive is still true.",
      "Press Esc, type run, press Enter, then shoot the enemy to test - both the bullet and the enemy should disappear on contact.",
    ],
    level1Code: [
      "enemy_alive=true",
      "ex,ey=60,10",
      "bullets={}",
      "",
      "function TIC()",
      " cls(0)",
      " if btn(3) then x=x+1 end",
      " if btnp(4) then",
      "  table.insert(bullets,{bx=x,by=y})",
      " end",
      " for i=#bullets,1,-1 do",
      "  local b=bullets[i]",
      "  b.by=b.by-2",
      "  if enemy_alive and b.bx<ex+8 and b.bx+8>ex and b.by<ey+8 and b.by+8>ey then",
      "   enemy_alive=false",
      "   table.remove(bullets,i)",
      "  elseif b.by<0 then",
      "   table.remove(bullets,i)",
      "  else",
      "   spr(2,b.bx,b.by,0)",
      "  end",
      " end",
      " if enemy_alive then spr(3,ex,ey,0) end",
      " spr(1,x,y,0)",
      "end",
    ],
    level2: [
      "Fire several bullets in a row, including some that miss - the loop walks every bullet in the table each frame, so a hit only ever removes the one bullet whose position actually overlaps the enemy, and every other bullet keeps moving and gets cleaned up off-screen exactly like it did in Lesson 3.",
      "The enemy_alive check and the overlap check together mean that once enemy_alive is false, that branch can never run again for any later bullet - so the enemy can only ever be \"destroyed\" once, even if two bullets happen to overlap it on the very same frame.",
    ],
    level3: [
      "Sound: design a short impact sound in the SFX Editor (F4) in an empty slot (e.g. slot 2, since slot 1 is likely your shoot sound), then call sfx(2) right where enemy_alive is set to false.",
      "Visual: draw enemy sprite index 3 as visually distinct from the player and bullet (different shape or color) in the Sprite Editor - this alone satisfies most of \"its own look\" without any extra code.",
      "Structure: keep the enemy's own state (enemy_alive, ex, ey) declared once at the top of your code, and only ever change enemy_alive from inside the collision check shown above - resist checking or setting it from more than one place, even though everything still technically lives inside the single TIC() function.",
    ],
    trouble: [
      "Enemy vanishes but no bullet was actually near it: double check the overlap math is comparing bullet coordinates (b.bx, b.by) against the enemy's (ex, ey), not against the player's (x, y).",
      "Bullets stop appearing once the enemy is destroyed: confirm the enemy_alive check only guards the collision branch and the drawing line at the bottom, not the whole bullets loop.",
      "\"attempt to compare nil with number\" error: confirm ex,ey=60,10 (or your chosen coordinates) is declared before TIC() runs, same as bullets={}.",
    ],
    resource: "Student Resource List > TIC-80 section, or tic80.com",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 3.",
      "Create a new Actor Template for the enemy (the Character template works well) using enemy_invader_red.png (import it from this lesson's Assets folder) as its Body sprite, and name it enemy.",
      "Place one enemy Actor on the Map near the top of the play area - leave \"Initially Existent\" ticked this time, since (unlike the bullet) there's only ever one enemy and it should be there from the start.",
    ],
    level1: [
      "On the enemy Actor, add an Actor Trigger: \"Collided with actor\" -> a script with two \"Actor – Destroy/Revive Actor\" events (Destroy): one targeting relate (the actor that collided with it - your bullet) and one targeting this (itself).",
      "Save, then Play/Test and shoot the enemy - both the bullet and the enemy should disappear on contact.",
    ],
    level1Code: [
      "enemy Actor Trigger: Collided with actor ->",
      "  Actor – Destroy/Revive Actor(relate, Destroy)",
      "  Actor – Destroy/Revive Actor(this, Destroy)",
    ],
    level2: [
      "Fire several bullets in a row, including some that miss - each hit destroys exactly one bullet and the enemy, and any bullet that misses keeps flying and destroys itself off-screen exactly like it did in Lesson 3.",
      "Once the enemy Actor is destroyed, it no longer exists to collide with anything, so its own \"Collided with actor\" trigger can't run a second time - the enemy can only ever be destroyed once, even if two bullets somehow reach it on the same update.",
    ],
    level3: [
      "Sound: import sfx_hit.wav (Resources > Import > Sound), then add a \"Play Sound\" event to the enemy's \"Collided with actor\" script, before the two Destroy events.",
      "Visual: give the enemy Actor its own Tint/Opacity setting (Actor > Graphics Options), distinct from the bullet's, so it clearly reads as a threat rather than another shot.",
      "Structure: keep everything about how the enemy responds to being hit - the sound, both Destroy events - on the enemy Actor's own \"Collided with actor\" trigger. The bullet Actor doesn't need any new scripting at all this week; it only needed to exist as something the enemy's trigger could detect.",
    ],
    trouble: [
      "Nothing happens on contact: confirm the trigger is \"Collided with actor\" (not \"Collided with solid\", which is for walls/terrain), and that both the enemy and bullet Actors have \"Solid\" ticked - non-Solid Actors generate no collisions at all.",
      "Error or nothing happens when destroying relate: confirm relate is actually available in this trigger's script - it refers to whichever Actor collided with the one running the script, and only has a meaningful value inside triggers like this one.",
      "Enemy destroys itself immediately, with no bullet involved: check that the enemy and bullet Actors aren't overlapping in their initial placement on the Map - move the enemy's starting position if needed.",
    ],
    resource: "001gamecreator.com/wiki, Student Resource List > 001 Game Creator",
    altPath: {
      title: "Textual Scripting (typed BASIC-like code)",
      intro: "The same collision-detection pattern works with typed commands instead of graphical event blocks, in a separate throwaway project (see Lesson 3's warning about this switch being permanent).",
      warning: "IMPORTANT: only try this in your separate throwaway test project from Lesson 3, never your main [YourName]Game project.",
      steps: [
        "Open your throwaway Textual-mode project from Lesson 3 (or create a new one, following Lesson 3's Setup steps).",
        "Open the Script Editor for the enemy Actor's \"Collided with actor\" trigger, same as before.",
        "Instead of dragging graphical event blocks, type the command directly - start typing \"Destroy\" and use the editor's autocomplete/command list to find the text equivalent of the Actor – Destroy/Revive Actor event you used in the graphical path, once for relate and once for this.",
      ],
      note: "The Level 3 sound idea maps onto Textual commands too - it's a typed equivalent of Play Sound. Exact command names should be checked against the installed version's Script Editor autocomplete or the wiki before class.",
    },
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 3.",
      "Import enemy_invader_red.png (this lesson's Assets folder) into the Content Browser, then build M_Enemy_Red: a new Material, Shading Model = Unlit, a Texture Sample using enemy_invader_red.png. This one needs a second node too - right-click, search \"VectorParameter,\" add one, rename it Color, and leave its default value White. Multiply the Texture Sample's RGB output by the Color parameter (right-click > Multiply), then wire that result into Emissive Color. With Color left at White, multiplying does nothing and the enemy shows the plain red sprite - but it gives Lesson 5's striped-row idea and this lesson's own Visual idea a real, working tint to grab onto later, the same way a Sprite2D's modulate or a SpriteRenderer's color tints an actual sprite in the other engines.",
      "Create a new Blueprint class (Actor parent), name it BP_Enemy. Add a Static Mesh Component to it and set its shape to a small basic Shape, distinct from the player's Cube and BP_Bullet's Sphere from Lesson 3. Assign M_Enemy_Red to it in the Material slot - a default Shape's built-in material is Lit and renders solid black with this project's lights removed.",
      "Set BP_Enemy's collision component's Collision Presets to OverlapAllDynamic (Details panel, Collision section). Then open BP_Bullet from Lesson 3 and set its collision component to the same OverlapAllDynamic preset - both objects need this, not just the one whose overlap event you're about to build. The default preset a fresh Shape starts with (BlockAllDynamic) makes two dynamic objects physically block each other instead of generating an overlap event, so without this change on both sides, Event ActorBeginOverlap below will never fire no matter how correctly it's wired - and OverlapAllDynamic sets Generate Overlap Events on automatically as part of the preset, so there's no separate checkbox to hunt down.",
      "While you're in BP_Bullet, find the Add Actor World Offset node from Lesson 3's movement and check its Sweep box (click the small arrow on the node to reveal it if it's not already visible). Without Sweep, that node relocates BP_Bullet each frame without ever testing for collision along the way, so it can pass straight through BP_Enemy without the overlap ever registering - Sweep is what makes the move actually check for what it's passing into.",
      "Drag one instance of BP_Enemy into the level near the top of the play area.",
    ],
    level1: [
      "In BP_Enemy's Event Graph, add an Event ActorBeginOverlap node, feeding a \"Cast To BP_Bullet\" node on the Other Actor pin - on a successful cast, wire a \"Destroy Actor\" node targeting the cast result (the bullet), and a second \"Destroy Actor\" node targeting Self.",
      "Click Compile, then Play and shoot the enemy - both the bullet and the enemy should disappear on contact.",
    ],
    level1Code: [
      "BP_Enemy's Event Graph:",
      "Event ActorBeginOverlap(OtherActor) ->",
      "  Cast To BP_Bullet(OtherActor) ->",
      "    Success: Destroy Actor(target = Cast result)",
      "             Destroy Actor(target = Self)",
    ],
    level2: [
      "Fire several bullets in a row, including some that miss - each hit destroys exactly one BP_Bullet and BP_Enemy, and any bullet that misses keeps flying and destroys itself off-screen exactly like it did in Lesson 3.",
      "Destroy Actor is safe to call more than once on the same Actor - a repeat call on an Actor that's already pending destruction is simply ignored - so even the rare case of two bullets overlapping the enemy on the same tick doesn't cause an error.",
    ],
    level3: [
      "Sound: add a Play Sound 2D node (using an imported sfx_hit.wav Sound Wave asset) right after the successful Cast, before either Destroy Actor node.",
      "Visual: in BP_Enemy's own BeginPlay, create a dynamic material instance and set its Color parameter (from M_Enemy_Red in Setup) to a darker red, giving the enemy its own distinct tint on top of its already-distinct sprite - the same modulate/tint idea the other engines use, just reached through Unreal's Material system instead of a Sprite property.",
      "Structure: keep everything about how the enemy responds to being hit - the Cast, both Destroy Actor nodes, the sound - inside BP_Enemy's own Event Graph. BP_Bullet doesn't need any new nodes at all this week; it only needed to exist as something BP_Enemy's overlap event could detect.",
    ],
    trouble: [
      "Nothing happens on contact, no errors anywhere: this is almost always Collision Presets, not the overlap event's wiring - confirm BOTH BP_Enemy's and BP_Bullet's collision components are set to OverlapAllDynamic (Details panel, Collision section), not left on the default BlockAllDynamic. Two Actors visibly touching while either one is still set to Block instead of Overlap will never fire Event ActorBeginOverlap, and produces no error of any kind.",
      "Collision Presets are correct on both, but it still doesn't fire: check that Sweep is enabled on BP_Bullet's Add Actor World Offset node from Lesson 3 - without it, that node teleports BP_Bullet each frame without testing for collision along the way, so a fast-moving bullet can pass straight through BP_Enemy without ever registering the overlap.",
      "Everything above checks out and it still doesn't fire: select BP_Enemy's and BP_Bullet's Static Mesh components in their own Blueprint editors (not an instance in the level) and confirm each actually has a mesh assigned in its Static Mesh field - a component left on None has no real geometry to collide with, even if every collision setting on it is otherwise correct.",
      "Cast To BP_Bullet always fails: confirm the enemy is actually overlapping the bullet's collision shape, not just its visible mesh.",
      "Enemy renders solid black instead of red: open M_Enemy_Red and confirm Texture Sample -> Multiply -> Emissive Color is wired all the way through, that enemy_invader_red.png is actually assigned as the Texture Sample's Texture, and that the Color VectorParameter's default is White, not black (multiplying by black would zero out the sprite entirely).",
      "This is a heavy engine on our list - if a laptop struggles, it's fine to observe this week's demo and build the exercise in a lighter engine instead.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same overlap-detection pattern works in C++ instead of Blueprint nodes.",
      steps: [
        "In a new ACppEnemy class (Tools > New C++ Class, Actor parent, same as CppBullet in Lesson 3), bind a function to the OnActorBeginOverlap delegate, typically in BeginPlay().",
        "In that bound function, cast OtherActor to ACppBullet - on a successful cast, call Destroy() on both the cast result and this actor.",
        "Compile, drag a test setup into the level in place of your existing Blueprint enemy, Play, and test the same way as the Blueprint path.",
      ],
      code: [
        "// ACppEnemy.h",
        "virtual void BeginPlay() override;",
        "UFUNCTION()",
        "void OnOverlap(AActor* OverlappedActor, AActor* OtherActor);",
        "",
        "// ACppEnemy.cpp",
        "void ACppEnemy::BeginPlay() {",
        "    Super::BeginPlay();",
        "    OnActorBeginOverlap.AddDynamic(this, &ACppEnemy::OnOverlap);",
        "}",
        "",
        "void ACppEnemy::OnOverlap(AActor* OverlappedActor, AActor* OtherActor) {",
        "    if (Cast<ACppBullet>(OtherActor)) {",
        "        OtherActor->Destroy();",
        "        Destroy();",
        "    }",
        "}",
      ],
      note: "This mirrors the Blueprint path's Event ActorBeginOverlap and Cast To BP_Bullet nodes - same underlying engine calls, just from code. Level 2/3: Destroy() is safe to call more than once in C++ too, and UGameplayStatics::PlaySound2D() plus a dynamic material instance cover the Level 3 sound and visual ideas.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_04_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS (one per engine) - instructor only
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.0, { after: 120 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 4 Teacher's Key", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 20 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "INSTRUCTOR ONLY - do not distribute to students", bold: true, size: 18, color: "B00020" })] }),
    h2("Level 1 - Complete Worked Solution"),
    ...data.l1Notes.map(s => p(s)),
    code(data.l1Code),
    h2("Level 2 - Complete Worked Solution"),
    ...data.l2Notes.map(s => p(s)),
  ];
  if (data.l2Code) children.push(code(data.l2Code));
  if (data.l2Extra) children.push(...data.l2Extra.map(s => p(s)));
  children.push(
    h2("Level 3 - All Worked Solutions"),
    p("Every stretch-goal idea from the handout, fully worked - not just one. Useful for validating each idea actually works before class, and as a ready reference no matter which one a student ends up picking."),
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
      headers: { default: brand.pageHeader("Lesson 4 Teacher's Key - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    l1Notes: ["Two scripts on two sprites - the Bullet's existing \"when I start as a clone\" forever loop gets one new if-branch, and the Enemy gets one new listener script:"],
    l1Code: [
      "-- On the Bullet sprite (added inside the existing forever loop, before the off-screen check):",
      "when I start as a clone",
      "  show",
      "  go to x: (Player's x) y: (Player's y)",
      "  forever",
      "    if <touching [Enemy]?> then",
      "      broadcast [enemy hit]",
      "      delete this clone",
      "    change y by 10",
      "    if <y position > 170> then",
      "      delete this clone",
      "",
      "-- On the Enemy sprite:",
      "when green flag clicked",
      "  go to x: 0 y: 150",
      "  show",
      "",
      "when I receive [enemy hit]",
      "  hide",
    ],
    l2Notes: ["This already works with no extra code: Scratch's \"touching\" check always returns false for a hidden sprite, so once the Enemy sprite is hidden, no later bullet clone can trigger a second \"enemy hit\" broadcast, even with several bullets in flight at once."],
    l3Ideas: [
      {
        title: "Sound: play on hit",
        notes: ["Upload sfx_hit.wav from this lesson's Assets folder, then add \"start sound [sfx_hit]\" to the Enemy's \"when I receive [enemy hit]\" script, right before hide:"],
        code: ["when I receive [enemy hit]", "  start sound [sfx_hit]", "  hide"],
      },
      {
        title: "Visual: give the enemy its own look",
        notes: ["A Costume editor recolor, or a \"set color effect to\" block in the Enemy's green-flag script, makes it read clearly as a threat:"],
        code: ["when green flag clicked", "  go to x: 0 y: 150", "  set color effect to 50", "  show"],
      },
      {
        title: "Structure: keep hit-response on the Enemy sprite",
        notes: ["The Bullet's forever loop should only ever detect a hit and report it (touching check + broadcast + delete this clone) - everything about how the Enemy responds to being hit (sound, hiding) belongs on the Enemy's own \"when I receive [enemy hit]\" script. Keeping this boundary clean is what let the Sound idea above drop into the Enemy sprite alone."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, across the Bullet and Enemy sprites:"],
        code: [
          "-- Bullet sprite (inside the existing forever loop):",
          "when I start as a clone",
          "  show",
          "  go to x: (Player's x) y: (Player's y)",
          "  forever",
          "    if <touching [Enemy]?> then",
          "      broadcast [enemy hit]",
          "      delete this clone",
          "    change y by 10",
          "    if <y position > 170> then",
          "      delete this clone",
          "",
          "-- Enemy sprite:",
          "when green flag clicked",
          "  go to x: 0 y: 150",
          "  set color effect to 50",
          "  show",
          "",
          "when I receive [enemy hit]",
          "  start sound [sfx_hit]",
          "  hide",
        ],
      },
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    l1Notes: ["A Hitbox Area2D added to both bullet.tscn and enemy.tscn, with the collision logic living entirely in enemy.tscn's own script:"],
    l1Code: [
      "# enemy.tscn's script (attached to the Sprite2D root):",
      "extends Sprite2D",
      "",
      "func _ready():",
      "    $Hitbox.area_entered.connect(_on_hitbox_area_entered)",
      "",
      "func _on_hitbox_area_entered(area):",
      "    area.get_parent().queue_free()",
      "    queue_free()",
    ],
    l2Notes: ["queue_free() is safe to call more than once on the same node (a repeat call is simply ignored), so even the rare case of two bullets overlapping the Enemy on the very same frame doesn't cause an error - both bullets are destroyed and the Enemy is destroyed exactly once. No extra code is needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: play on hit, safely",
        notes: ["Freeing the Enemy immediately would cut its sound off before it finishes, so hide it and disable its Hitbox right away, play the sound, then await the AudioStreamPlayer2D's finished signal before actually freeing the node:"],
        code: [
          "func _on_hitbox_area_entered(area):",
          "    area.get_parent().queue_free()",
          "    hide()",
          "    $Hitbox.set_deferred(\"monitoring\", false)",
          "    $Hit.play()",
          "    await $Hit.finished",
          "    queue_free()",
        ],
      },
      {
        title: "Visual: give the enemy its own look",
        notes: ["Set modulate in the Enemy's own _ready() so it clearly reads as a threat rather than another shot:"],
        code: ["func _ready():", "    $Hitbox.area_entered.connect(_on_hitbox_area_entered)", "    modulate = Color(0.8, 0.1, 0.2)"],
      },
      {
        title: "Structure: keep hit-response in enemy.tscn's own script",
        notes: ["Everything about how the Enemy responds to being hit - destroying the bullet, hiding itself, playing a sound - lives inside enemy.tscn's own script. bullet.tscn's own script doesn't need to know the Enemy exists at all; it only needed a Hitbox added to it so the Enemy's Area2D could detect it."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, inside enemy.tscn's own script:"],
        code: [
          "extends Sprite2D",
          "",
          "func _ready():",
          "    $Hitbox.area_entered.connect(_on_hitbox_area_entered)",
          "    modulate = Color(0.8, 0.1, 0.2)",
          "",
          "func _on_hitbox_area_entered(area):",
          "    area.get_parent().queue_free()",
          "    hide()",
          "    $Hitbox.set_deferred(\"monitoring\", false)",
          "    $Hit.play()",
          "    await $Hit.finished",
          "    queue_free()",
        ],
      },
    ],
  },
  {
    engine: "Unity", file: "Unity",
    l1Notes: ["One new method on Bullet.cs, plus a Rigidbody2D (Kinematic) and Is Trigger colliders on both prefabs:"],
    l1Code: [
      "// Bullet.cs - add this method:",
      "void OnTriggerEnter2D(Collider2D other) {",
      "    if (other.CompareTag(\"Enemy\")) {",
      "        Destroy(other.gameObject);",
      "        Destroy(gameObject);",
      "    }",
      "}",
    ],
    l2Notes: ["Destroy() doesn't remove a GameObject instantly - it happens at the end of the current frame - so if two bullets somehow trigger on the exact same frame, both destroy calls on the Enemy are harmless; Unity simply ignores the second one, and no error is thrown. No extra code is needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: play on hit, safely",
        notes: ["AudioSource.PlayClipAtPoint() spawns its own temporary audio source that survives independently of the Enemy or Bullet being destroyed, so the sound always finishes playing even though both GameObjects are gone:"],
        code: [
          "public AudioClip hitSound;",
          "",
          "void OnTriggerEnter2D(Collider2D other) {",
          "    if (other.CompareTag(\"Enemy\")) {",
          "        AudioSource.PlayClipAtPoint(hitSound, transform.position);",
          "        Destroy(other.gameObject);",
          "        Destroy(gameObject);",
          "    }",
          "}",
        ],
      },
      {
        title: "Visual: give the enemy its own look",
        notes: ["A small Enemy.cs script's Start() sets its own SpriteRenderer.color, without touching Bullet.cs at all:"],
        code: ["// Enemy.cs, attached to the Enemy prefab:", "void Start() {", "    GetComponent<SpriteRenderer>().color = new Color(0.8f, 0.1f, 0.2f);", "}"],
      },
      {
        title: "Structure: keep enemy-appearance logic out of Bullet.cs",
        notes: ["OnTriggerEnter2D in Bullet.cs should only ever detect a hit and clean up (destroy the enemy, destroy itself, play the sound) - anything about how the Enemy looks on its own belongs in its own Enemy.cs. Keeping this boundary clean is what let the Visual idea above live in a totally separate script."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, across Bullet.cs and Enemy.cs:"],
        code: [
          "// Bullet.cs:",
          "public AudioClip hitSound;",
          "",
          "void OnTriggerEnter2D(Collider2D other) {",
          "    if (other.CompareTag(\"Enemy\")) {",
          "        AudioSource.PlayClipAtPoint(hitSound, transform.position);",
          "        Destroy(other.gameObject);",
          "        Destroy(gameObject);",
          "    }",
          "}",
          "",
          "// Enemy.cs:",
          "void Start() {",
          "    GetComponent<SpriteRenderer>().color = new Color(0.8f, 0.1f, 0.2f);",
          "}",
        ],
      },
    ],
    altNotes: {
      title: "Visual Scripting",
      notes: [
        "The same physics requirements apply in graphs as in code: at least one Rigidbody2D between the two colliding objects, both colliders set to Is Trigger.",
        "The Level 3 goals map onto nodes the same way they map onto code: an Audio Source \"Play Clip At Point\" node for sound, a \"Set Member\" node on the Sprite Renderer's color for the enemy's own visual idea - all inside the Enemy prefab's own graph, kept separate from the Bullet's graph the same way Enemy.cs and Bullet.cs stay separate in the code path.",
      ],
      why: [
        "Keeping the same Structure principle (a GameObject's own appearance/response lives with that GameObject) whether a student is in C# or graphs reinforces that it's a design idea, not a syntax rule.",
      ],
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    l1Notes: ["A single Collision event on obj_bullet, targeting obj_enemy:"],
    l1Code: [
      "// obj_bullet's Collision event with obj_enemy:",
      "instance_destroy(other);",
      "instance_destroy(self);",
    ],
    l2Notes: ["GameMaker's Collision event already only fires once per pair of instances actually overlapping that step, and instance_destroy() is safe to call on an instance that's already been destroyed earlier in the same step (it's simply ignored) - so even the rare case of two bullets hitting the same enemy on the same step doesn't error. No extra code is needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: play on hit",
        notes: ["Add sfx_hit.wav as a sound resource (keep its default name, sfx_hit), then call audio_play_sound() at the top of the Collision event, before either instance_destroy() call:"],
        code: ["audio_play_sound(sfx_hit, 1, false);", "instance_destroy(other);", "instance_destroy(self);"],
      },
      {
        title: "Visual: give the enemy its own look",
        notes: ["Set image_blend in obj_enemy's own Create event, without touching obj_bullet's code at all:"],
        code: ["// obj_enemy's Create event:", "image_blend = make_color_rgb(200, 25, 55);"],
      },
      {
        title: "Structure: keep enemy-appearance logic out of obj_bullet",
        notes: ["obj_bullet's Collision event should only ever detect a hit and clean up (destroy the enemy, destroy itself, play the sound) - anything about how the enemy looks on its own belongs in obj_enemy's own Create event, not bolted onto obj_bullet."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, across obj_bullet and obj_enemy:"],
        code: [
          "// obj_enemy's Create event:",
          "image_blend = make_color_rgb(200, 25, 55);",
          "",
          "// obj_bullet's Collision event with obj_enemy:",
          "audio_play_sound(sfx_hit, 1, false);",
          "instance_destroy(other);",
          "instance_destroy(self);",
        ],
      },
    ],
    altNotes: {
      title: "GML Visual",
      notes: [
        "The Collision event itself is set up identically in both paths - only what's inside it differs (blocks vs. code).",
        "The Level 3 goals map onto blocks the same way they map onto code, favoring the most specific named block available: \"Play Audio\" (Audio category) for sound, and \"Set Instance Variable\" -> Image Blend for the visual idea - Image Blend is a recognized built-in on Set Instance Variable's own dropdown, so it's the better choice over the generic Assign Variable/Set Variable. All of this stays inside obj_enemy's own events, kept separate from obj_bullet's events the same way the code path keeps the two objects separate.",
      ],
      why: [
        "Keeping the same Structure principle (an object's own appearance/response lives with that object) whether a student is in Code or Visual reinforces that it's a design idea, not a syntax rule.",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    l1Notes: ["One event: a collision condition between Bullet and Enemy, with two Delete actions:"],
    l1Code: [
      "Condition: Bullet is in collision with Enemy",
      "Action: Delete Bullet",
      "Action: Delete Enemy",
    ],
    l2Notes: ["GDevelop's collision condition and Delete actions already operate per pair of overlapping instances each frame, so even if two bullets somehow overlap the Enemy on the same frame, the event simply runs twice that frame - the Enemy and both bullets are still deleted cleanly either way. No extra code is needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: play on hit",
        notes: ["Add a \"Play a sound\" action to the same event as the collision condition:"],
        code: ["Condition: Bullet is in collision with Enemy", "Action: Play a sound (sfx_hit.wav)", "Action: Delete Bullet", "Action: Delete Enemy"],
      },
      {
        title: "Visual: give the enemy its own look",
        notes: ["\"Tint color\" applied once, at the beginning of the scene, so the Enemy clearly reads as a threat:"],
        code: ["At the beginning of the scene:", "Action: Tint color of Enemy = Dark Red"],
      },
      {
        title: "Structure: group this lesson's events together",
        notes: ["Keep the collision condition and its actions inside a labeled event group (e.g. \"Combat\"), separate from the \"Shooting\" group from Lesson 3 - GDevelop's event sheet has no hard file boundary between objects the way separate scripts/objects do in other engines, so this labeled grouping is what keeps the two mechanics from blurring together."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above, organized as one event group named \"Combat\":"],
        code: [
          "At the beginning of the scene:",
          "Action: Tint color of Enemy = Dark Red",
          "",
          "Event group: \"Combat\"",
          "Condition: Bullet is in collision with Enemy",
          "Action: Play a sound (sfx_hit.wav)",
          "Action: Delete Bullet",
          "Action: Delete Enemy",
        ],
      },
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    l1Notes: ["Fixed enemy state declared once at the top of the code, plus one new branch inside the existing bullets loop:"],
    l1Code: [
      "enemy_alive=true",
      "ex,ey=60,10",
      "bullets={}",
      "",
      "function TIC()",
      " cls(0)",
      " if btn(3) then x=x+1 end",
      " if btnp(4) then",
      "  table.insert(bullets,{bx=x,by=y})",
      " end",
      " for i=#bullets,1,-1 do",
      "  local b=bullets[i]",
      "  b.by=b.by-2",
      "  if enemy_alive and b.bx<ex+8 and b.bx+8>ex and b.by<ey+8 and b.by+8>ey then",
      "   enemy_alive=false",
      "   table.remove(bullets,i)",
      "  elseif b.by<0 then",
      "   table.remove(bullets,i)",
      "  else",
      "   spr(2,b.bx,b.by,0)",
      "  end",
      " end",
      " if enemy_alive then spr(3,ex,ey,0) end",
      " spr(1,x,y,0)",
      "end",
    ],
    l2Notes: ["The enemy_alive check and the overlap check together mean that once enemy_alive is false, that branch can never run again for any later bullet - so the enemy can only ever be \"destroyed\" once, even if two bullets happen to overlap it on the very same frame. No extra code is needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: play on hit",
        notes: ["Design a short impact sound in the SFX Editor (F4) in an empty slot (e.g. slot 2), then call sfx(2) right where enemy_alive is set to false:"],
        code: [
          " if enemy_alive and b.bx<ex+8 and b.bx+8>ex and b.by<ey+8 and b.by+8>ey then",
          "  enemy_alive=false",
          "  sfx(2)",
          "  table.remove(bullets,i)",
        ],
      },
      {
        title: "Visual: give the enemy its own look",
        notes: ["Draw enemy sprite index 3 as visually distinct from the player and bullet (different shape or color) in the Sprite Editor - this alone satisfies most of \"its own look\" without any extra code, the same way Lesson 3's bullet used its own sprite index."],
      },
      {
        title: "Structure: keep enemy state declared once",
        notes: ["Keep enemy_alive, ex, and ey declared once at the top of the code, and only ever change enemy_alive from inside the collision check - resist checking or setting it from more than one place, even though everything still technically lives inside the single TIC() function."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined into the one TIC() function:"],
        code: [
          "enemy_alive=true",
          "ex,ey=60,10",
          "bullets={}",
          "",
          "function TIC()",
          " cls(0)",
          " if btn(3) then x=x+1 end",
          " if btnp(4) then",
          "  table.insert(bullets,{bx=x,by=y})",
          " end",
          " for i=#bullets,1,-1 do",
          "  local b=bullets[i]",
          "  b.by=b.by-2",
          "  if enemy_alive and b.bx<ex+8 and b.bx+8>ex and b.by<ey+8 and b.by+8>ey then",
          "   enemy_alive=false",
          "   sfx(2)",
          "   table.remove(bullets,i)",
          "  elseif b.by<0 then",
          "   table.remove(bullets,i)",
          "  else",
          "   spr(2,b.bx,b.by,0)",
          "  end",
          " end",
          " if enemy_alive then spr(3,ex,ey,0) end",
          " spr(1,x,y,0)",
          "end",
        ],
      },
    ],
  },
  {
    engine: "001 Game Creator", file: "001GameCreator",
    l1Notes: ["One Actor Trigger script on the enemy Actor itself:"],
    l1Code: [
      "enemy Actor Trigger: Collided with actor ->",
      "  Actor – Destroy/Revive Actor(relate, Destroy)",
      "  Actor – Destroy/Revive Actor(this, Destroy)",
    ],
    l2Notes: ["Once the enemy Actor is destroyed, it no longer exists to collide with anything, so its own \"Collided with actor\" trigger can't run a second time - the enemy can only ever be destroyed once, even if two bullets somehow reach it on the same update. No extra scripting is needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: play on hit",
        notes: ["Import sfx_hit.wav (Resources > Import > Sound), then add a \"Play Sound\" event before the two Destroy events:"],
        code: [
          "enemy Actor Trigger: Collided with actor ->",
          "  Play Sound(sfx_hit.wav)",
          "  Actor – Destroy/Revive Actor(relate, Destroy)",
          "  Actor – Destroy/Revive Actor(this, Destroy)",
        ],
      },
      {
        title: "Visual: give the enemy its own look",
        notes: ["Set Tint/Opacity directly on the enemy Actor Template's own Graphics Options - a Template-level setting, so it applies with zero extra events, the same technique Lesson 3 used for the bullet's own look."],
      },
      {
        title: "Structure: keep hit-response on the enemy Actor",
        notes: ["Keep everything about how the enemy responds to being hit - the sound, both Destroy events - on the enemy Actor's own \"Collided with actor\" trigger. The bullet Actor doesn't need any new scripting at all this week; it only needed to exist as something the enemy's trigger could detect."],
      },
      {
        title: "Complete Level 3 solution (all 2 goals combined)",
        notes: ["The enemy Actor Template's own settings plus its collision trigger script:"],
        code: [
          "enemy Actor Template settings:",
          "  Graphics Options -> Tint/Opacity set",
          "enemy Actor Trigger: Collided with actor ->",
          "  Play Sound(sfx_hit.wav)",
          "  Actor – Destroy/Revive Actor(relate, Destroy)",
          "  Actor – Destroy/Revive Actor(this, Destroy)",
        ],
      },
    ],
    altNotes: {
      title: "Textual Scripting",
      notes: [
        "Reiterate to students in person: this mode switch is genuinely irreversible per-project. Confirm they're working in their separate throwaway project from Lesson 3.",
        "Exact text-command names should be checked against the installed version's Script Editor autocomplete or the wiki before class.",
        "The Level 3 sound idea maps onto Textual commands too - a typed equivalent of Play Sound. The visual idea needs no scripting either way, since it's a Template-level setting in both modes.",
      ],
      why: [
        "This path exists for students who've realized they prefer typing code and want to see whether 001 Game Creator supports that style too - it does, with the same irreversibility caveat as Lesson 3.",
      ],
    },
  },
  {
    engine: "Unreal", file: "Unreal",
    l1Notes: ["Requires Setup's collision changes actually done on both objects: Collision Presets = OverlapAllDynamic on BP_Enemy's collision component AND on BP_Bullet's (not just one of the two), plus Sweep checked on BP_Bullet's Add Actor World Offset node from Lesson 3 - watch for students who only touch BP_Enemy and leave BP_Bullet on its Lesson 3 default (BlockAllDynamic, no Sweep), which silently prevents the overlap from ever registering no matter how correct the graph below is. BP_Enemy's own Event Graph handles detection and cleanup entirely once that's done - BP_Bullet needs no new nodes at all:"],
    l1Code: [
      "BP_Enemy's Event Graph:",
      "Event ActorBeginOverlap(OtherActor) ->",
      "  Cast To BP_Bullet(OtherActor) ->",
      "    Success: Destroy Actor(target = Cast result)",
      "             Destroy Actor(target = Self)",
    ],
    l2Notes: ["Destroy Actor is safe to call more than once on the same Actor - a repeat call on an Actor that's already pending destruction is simply ignored - so even the rare case of two bullets overlapping the enemy on the same tick doesn't cause an error. No extra nodes are needed for Level 2."],
    l3Ideas: [
      {
        title: "Sound: play on hit",
        notes: ["Add a Play Sound 2D node (using an imported sfx_hit.wav Sound Wave asset) right after the successful Cast:"],
        code: [
          "Success: Play Sound 2D(sfx_hit.wav)",
          "         Destroy Actor(target = Cast result)",
          "         Destroy Actor(target = Self)",
        ],
      },
      {
        title: "Visual: give the enemy its own look",
        notes: ["Create a dynamic material instance and set its Color parameter (the one M_Enemy_Red's Setup step built specifically so this would work) in BP_Enemy's own BeginPlay - since that parameter multiplies against the enemy's real sprite instead of replacing it outright, Dark Red genuinely darkens the red invader texture instead of doing nothing:"],
        code: ["BP_Enemy's Event Graph:", "BeginPlay -> Create Dynamic Material Instance(Self's Material) -> Set Vector Parameter Value(\"Color\", Dark Red)"],
      },
      {
        title: "Structure: keep hit-response inside BP_Enemy",
        notes: ["Keep everything about how the enemy responds to being hit - the Cast, both Destroy Actor nodes, the sound - inside BP_Enemy's own Event Graph. BP_Bullet doesn't need any new nodes at all this week; it only needed to exist as something BP_Enemy's overlap event could detect."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, entirely inside BP_Enemy's own Event Graph:"],
        code: [
          "BP_Enemy's Event Graph:",
          "BeginPlay -> Create Dynamic Material Instance(Self's Material) -> Set Vector Parameter Value(\"Color\", Dark Red)",
          "Event ActorBeginOverlap(OtherActor) ->",
          "  Cast To BP_Bullet(OtherActor) ->",
          "    Success: Play Sound 2D(sfx_hit.wav)",
          "             Destroy Actor(target = Cast result)",
          "             Destroy Actor(target = Self)",
        ],
      },
    ],
    altNotes: {
      title: "C++",
      notes: [
        "OnActorBeginOverlap.AddDynamic() bound in BeginPlay() is the templated-delegate equivalent of wiring the Event ActorBeginOverlap node in Blueprint - same underlying engine event, just subscribed to from code.",
        "The Level 3 goals map onto C++ calls the same way they map onto Blueprint nodes: UGameplayStatics::PlaySound2D() for sound, and a UMaterialInstanceDynamic's SetVectorParameterValue() for the visual idea, both called from ACppEnemy's own BeginPlay()/OnOverlap() - keeping the same Structure principle as the Blueprint path.",
      ],
      why: [
        "Showing both C++ and Blueprint side by side for the same behavior is the clearest way to demystify Blueprints for students who assume \"visual\" means \"not real code\" - it's the same engine calls either way.",
        "This is squarely the optional advanced track - most students should stay on Blueprints; C++ is for students already comfortable with typed languages from Godot, GameMaker, or Unity's C# path.",
      ],
    },
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_04_Key_${d.file}.docx`]));

// =========================================================================
// WRITE ALL FILES
// =========================================================================
(async () => {
  for (const [doc, filename] of jobs) {
    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(filename, buf);
    console.log("wrote", filename);
  }
  console.log("done,", jobs.length, "files");
})();
