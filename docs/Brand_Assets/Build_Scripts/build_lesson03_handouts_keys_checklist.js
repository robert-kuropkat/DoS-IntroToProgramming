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
      headers: { default: brand.pageHeader("Lesson 3 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 3 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("Self-Contained Lesson Assets"),
        p("Per Lesson Production Standards Rule 3, nothing this lesson needs should require a student to search for or download something mid-class. Everything below is already prepared and included in this lesson's Assets folder - no download required."),
        checkbox("New placeholder projectile sprite already generated (bullet_yellow.png, 128x128 with transparency, same canvas size as Lesson 2's player sprites) - a simple yellow capsule shape, distinct from the round player sprites."),
        checkbox("New sound effect already generated (sfx_shoot.wav, a short descending \"pew\" - deliberately distinct from Lesson 2's sfx_blip.wav so shooting and moving don't sound the same) for the Level 3 \"add a sound\" idea."),
        checkbox("Lesson 2's Assets folder (player sprites, sfx_blip.wav) is still needed too, since this lesson builds on the same ongoing [YourName]Game project rather than starting fresh."),
        checkbox("Copy this lesson's Assets folder onto the shared classroom drive / USB backups alongside the handouts."),

        h2("Continuity From Lesson 2"),
        p("This is the first lesson where students open an existing project instead of starting a new one - flag this clearly at the start of lab, since it's a new habit."),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 2, not create a new one - the shooting mechanic is added on top of the player object they already built."),
        checkbox("If a student's Lesson 2 project is lost or broken, have a pre-built \"Lesson 2 complete\" starter available per engine (from the Lesson 2 Teacher's Keys) so they aren't blocked."),
        checkbox("Note for students whose Lesson 2 player still moves in all 4 directions (the Level 2 stretch): that's fine to leave as-is, or trim to Left/Right only to match the actual capstone game - either is acceptable, since the capstone itself only calls for Left/Right (see the Lesson 2 GDD)."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop's Hello World and Lesson 2 exercise still work for all 8 engines - re-run any that haven't been touched in a few weeks."),
        checkbox("Pre-build the full Scratch shooting exercise (all levels) on top of a finished Lesson 2 project, so you have a working example ready to show."),
        checkbox("Godot: confirm a second PackedScene (the projectile) can be preloaded and instantiated from the player's script without errors in your installed version. Also confirm testing this week's exercise means pressing F5 (Run Project), not F6 on player.tscn alone - the bullet-spawn code relies on Player actually running as Main's child."),
        checkbox("Unity: confirm Instantiate() and a simple prefab workflow behave as expected in your installed Editor version."),
        checkbox("GameMaker: quick-test instance_create_layer() targeting the room's existing instance layer name in a throwaway object before class."),
        checkbox("GDevelop: confirm the \"Destroy outside screen\" behavior is available when adding a behavior to an object in your installed version (it should be pre-installed by default)."),
        checkbox("TIC-80: confirm btnp() (press, not hold) is being used for firing, not btn() - btn() would fire a new bullet every single frame the key is held."),
        checkbox("001 Game Creator (Windows demo laptop): confirm the \"Map – Spawn Actor\" and \"Left/Entered Camera\" Actor Trigger options are accessible in the Script/Actor editors."),
        checkbox("If any advanced-track students are on Unreal this week, confirm Spawn Actor from Class and a simple Bullet Blueprint compile cleanly."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies, Assets folder (on USB as backup)."),

        h2("Known Gotchas This Week"),
        bullet("Every engine's Level 1 uses a \"fire once per press\" check (btnp(), \"Pressed\" pin, a boolean latch, etc.) rather than \"fire every frame the key is held\" - flag this early, since it's the single most common bug this week (a held key spawning dozens of projectiles per second)."),
        bullet("Scratch: creating and deleting a clone within the same \"when I start as a clone\" stack can freeze the project - the movement script and the off-screen cleanup check must be the same stack (a clone only runs one \"when I start as a clone\" script), but avoid also creating new clones from inside that same stack."),
        bullet("GDevelop: the Destroy Outside Screen behavior only removes an object once it's fully past the deletion margin (200px default) - a projectile won't vanish the instant it touches the very edge of the screen, and that's expected, not a bug."),
        bullet("001 Game Creator: dialog field layouts and exact event names can vary slightly by version - if something looks different from the handout, the Script Editor and Actor wiki pages have the current layout."),
        bullet("Unreal: this is the heaviest engine on our list - if a student's laptop is struggling, it's fine for them to observe this week's demo and build the exercise in a lighter engine instead."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_03_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 3: Take Your Shot", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: spawning objects at runtime. Your job is to make your player shoot a projectile that flies up the screen and cleans itself up when it leaves - the first piece of your capstone's actual shooting mechanic."),
    p("Open your own [YourName]Game project from Lesson 2 - don't start a new one. Everything this week builds on the player object you already have.", { bold: true }),
    p("Tip: this lesson's Assets folder has a ready-to-use placeholder projectile image (bullet_yellow.png) and a new sound effect (sfx_shoot.wav) - you don't need to find or make your own art today.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("Pressing the fire key spawns one projectile that flies upward and disappears when it leaves the screen. This is a complete finish for today - reaching this is enough to move on to next week.", { italics: true }),
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
    p("Firing repeatedly produces multiple independent projectiles on screen at once, each moving and disappearing correctly, and rapid-fire is capped so holding the key down doesn't flood the screen.", { italics: true }),
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
      headers: { default: brand.pageHeader("Lesson 3 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 2.",
      "Click the Sprite pane and create a new sprite for the projectile - import bullet_yellow.png from this lesson's Assets folder (Upload Sprite), or draw a small shape of your own.",
      "Rename this new sprite Bullet, then hide it in the initial state (it should only appear as a clone, never as itself) - add a \"when green flag clicked\" script to the Bullet sprite with just a \"hide\" block.",
    ],
    level1: [
      "On the Player sprite, click Events, drag \"when [space] key pressed\" into a new script.",
      "Click the Control category, drag \"create clone of [Bullet]\" underneath it - this is a separate, independent script from your Player's movement scripts.",
      "Switch to the Bullet sprite. Click Control, drag \"when I start as a clone\" into a new script area.",
      "Under it, add \"show\", then a Motion \"go to x: (Player's x) y: (Player's y)\" block (Sensing > \"[Player]'s [x position]\" reporters dragged into the x/y slots).",
      "Add a \"forever\" loop (Control) containing \"change y by 10\" (Motion).",
      "Inside that same forever loop, add \"if <y position > 170> then delete this clone\" (Control > Sensing/Operators > Motion).",
      "Click the stage, then press spacebar to test - a bullet should appear at the Player and fly upward, then vanish at the top.",
    ],
    level2: [
      "Test firing several times quickly - each press should produce its own independent bullet, all moving and disappearing on their own. Clones already behave this way automatically in Scratch, since each clone runs its own full copy of the Bullet's scripts.",
      "Add a cooldown so holding spacebar doesn't fire a new bullet every single frame: on the Player, add a \"can_shoot\" variable, set it to true at the start, and only \"create clone of [Bullet]\" when can_shoot is true - immediately set it to false, then use a short \"wait 0.2 seconds\" before setting it back to true.",
    ],
    level3: [
      "Sound: play sfx_shoot.wav (upload it from this lesson's Assets folder) the moment a bullet is fired - add \"start sound [sfx_shoot]\" right next to \"create clone of [Bullet]\" on the Player's fire script.",
      "Visual: give the Bullet its own look distinct from the Player - a different costume, size, or a brief \"change size by\" pulse when it spawns, so it reads clearly as \"a shot\" rather than a copy of the character.",
      "Structure: keep every piece of a clone's own behavior - appearing, moving, and cleaning itself up - inside that one \"when I start as a clone\" script on the Bullet sprite, and keep the Player's fire script limited to just triggering it (create clone + cooldown + sound). Don't let Bullet logic leak into the Player's scripts, or vice versa.",
    ],
    trouble: [
      "Nothing happens on spacebar: make sure \"create clone of\" is targeting Bullet (not \"myself\"), and that the Bullet sprite's \"when I start as a clone\" script actually has \"show\" - clones inherit whatever visibility the original had, which is hidden here.",
      "Project freezes or becomes unresponsive: this almost always means a clone is creating another clone inside its own \"when I start as a clone\" script - the fire logic belongs on the Player, not the Bullet.",
      "Bullets never disappear: double check the y > 170 comparison is inside the forever loop (not after it), and that the stage's visible area is roughly -180 to 180 for y.",
      "Holding a direction key, then pressing spacebar to fire, makes movement pause and the held arrow key stop responding: this is a known Scratch quirk, not a bug in your project - \"when key pressed\" scripts repeat by riding your operating system's own keyboard auto-repeat, and most operating systems only actively auto-repeat one key at a time, so a second key press can interrupt the first key's repeat stream. Release and press the movement key again to resume. See this lesson's Sidebars for why, and the Super-Stretch Ideas list for a way to rebuild movement so this can't happen.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 2.",
      "Create a new scene: a Sprite2D root node, rename it Bullet, and assign bullet_yellow.png (from this lesson's Assets folder) as its Texture.",
      "Save this new scene as bullet.tscn in the project's root folder.",
      "Attach a new script to the Bullet root, extends Sprite2D.",
    ],
    level1: [
      "In the Bullet's own script, add an exported speed and a _process(delta) that moves it upward and frees itself once it's off-screen (code below).",
      "Back in the Player's script (from Lesson 2), preload the bullet scene at the top and add an Input.is_action_just_pressed(\"ui_select\") check in _process to instantiate it, place it at global_position (not position - global_position is the safe choice for matching an object's actual on-screen location no matter where it sits in the scene tree, and you'll keep using it for Enemy and Explosion in later lessons too), and add it as a child of get_parent() rather than the Player itself, so the bullet doesn't end up riding along attached to the Player - since Player is a child of Main (from Lesson 2's Setup), get_parent() resolves to Main, a real node you can see in the Scene tree, so the bullet lands as Player's sibling, not its passenger.",
      "Press F5 (Run Project) rather than F6 this time - the bullet spawns via get_parent(), which only resolves correctly when Player is actually running as Main's child, not as its own temporary scene root. Press Space (Godot's default ui_select action) to test - a bullet should spawn at the player and fly upward, then disappear near the top.",
    ],
    level1Code: [
      "# bullet.tscn's script (attached to the Sprite2D root):",
      "extends Sprite2D",
      "var speed := 400.0",
      "",
      "func _process(delta):",
      "    position.y -= speed * delta",
      "    if position.y < -20:",
      "        queue_free()",
      "",
      "# Player's script - add near the top and inside _process:",
      "var bullet_scene := preload(\"res://bullet.tscn\")",
      "",
      "func _process(delta):",
      "    if Input.is_action_just_pressed(\"ui_select\"):",
      "        var b := bullet_scene.instantiate()",
      "        b.global_position = global_position",
      "        get_parent().add_child(b)",
    ],
    level2: [
      "Test firing several times quickly - each press should produce its own independent bullet instance, all moving and freeing themselves correctly. Since each instantiate() call creates a fully separate node, this already works with no extra code.",
      "Add a cooldown so holding Space doesn't fire every frame: is_action_just_pressed() already only fires once per physical press (not once per frame held), so Level 1 is already immune to this specific bug - add a var fire_cooldown := 0.0 timer anyway if you want a slower, tunable fire rate (count it down by delta each frame, and only allow firing when it's <= 0, resetting it to e.g. 0.2 on each shot).",
    ],
    level3: [
      "Sound: add an AudioStreamPlayer2D node to the Player, set its Stream to sfx_shoot.wav, and call $AudioStreamPlayer2D.play() right next to the instantiate() call.",
      "Visual: give the Bullet its own modulate color or a small scale pulse that fades back to normal after a fraction of a second, so it reads clearly as \"a shot\" rather than a copy of the Player - not a permanent recolor.",
      "Structure: keep everything about how a bullet behaves - its own movement, its own off-screen cleanup - inside bullet.tscn's own script, and keep the Player's script limited to just spawning it (instantiate + position + cooldown + sound). Don't reach into a bullet's internals from the Player's script, or vice versa.",
    ],
    trouble: [
      "Nothing happens on Space: confirm Space is actually bound to the ui_select action in Project Settings > Input Map (it's a Godot default, but verify).",
      "Bullet spawns but doesn't move or never disappears: confirm the script is attached to bullet.tscn's own Sprite2D root, not left unattached, and that _process is spelled exactly right (case-sensitive).",
      "Error about instantiate() on a null value: the preload() path is wrong - confirm bullet.tscn was actually saved at res://bullet.tscn (or update the path to match where you saved it).",
      "Bullet spawns way off from the player - low and to the right, or some other consistent offset: confirm both sides of the assignment use global_position (b.global_position = global_position), not plain position - it's the safe choice for matching an object's actual on-screen location no matter where it sits in the scene tree. If bullets still drift or rotate along with the Player after that, confirm Player is actually a child of Main (from Lesson 2's Setup) - get_parent().add_child(b) only lands the bullet next to Player, as its sibling, if Main is really Player's parent. Also double check the bullet was added with get_parent().add_child(b), not get_tree().current_scene.add_child(b) - since Main is current_scene here, that line would still work, but it bypasses the actual parent-child relationship the rest of this lesson relies on.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Open your [YourName]Game project from Lesson 2.",
      "Create a new empty GameObject, add a Sprite Renderer component, and assign bullet_yellow.png (import it from this lesson's Assets folder into your Project window first) as its Sprite.",
      "Drag this GameObject from the Hierarchy into the Project window to turn it into a Prefab, named Bullet - then delete the one left in the scene (the Prefab asset is all you need going forward).",
      "Create a new C# script named Bullet, attach it to the Bullet prefab.",
    ],
    level1: [
      "In the Bullet script, add a public float speed and an Update() that moves it upward and destroys it once it's off-screen (code below).",
      "On the Player script (from Lesson 2), add a public GameObject bulletPrefab field and check Keyboard.current.spaceKey.wasPressedThisFrame in Update() to Instantiate() it at the player's position.",
      "In the Inspector, drag the Bullet prefab from the Project window onto the Player's Bullet Prefab field.",
      "Press Play and hit Space to test - a bullet should spawn at the player and fly upward, then disappear.",
    ],
    level1Code: [
      "// Bullet.cs, attached to the Bullet prefab:",
      "public float speed = 6f;",
      "",
      "void Update() {",
      "    transform.Translate(Vector3.up * speed * Time.deltaTime);",
      "    if (transform.position.y > 6f) Destroy(gameObject);",
      "}",
      "",
      "// Player.cs - add a field and a check inside Update():",
      "public GameObject bulletPrefab;",
      "",
      "if (Keyboard.current.spaceKey.wasPressedThisFrame) {",
      "    Instantiate(bulletPrefab, transform.position, Quaternion.identity);",
      "}",
    ],
    level2: [
      "Test firing several times quickly - each press should produce its own independent bullet instance, all moving and destroying themselves correctly. Since each Instantiate() call creates a fully separate GameObject, this already works with no extra code.",
      "Add a cooldown so holding Space doesn't fire every frame: wasPressedThisFrame already only returns true on the single frame a key is first pressed (not every frame held), so Level 1 is already immune to this specific bug - add a private float fireCooldown = 0f timer anyway if you want a slower, tunable fire rate (count it down by Time.deltaTime each frame, and only allow firing when it's <= 0, resetting it to e.g. 0.2f on each shot).",
    ],
    level3: [
      "Sound: add an AudioSource component to the Player, assign sfx_shoot.wav as its AudioClip, and call audioSrc.Play() right next to the Instantiate() call.",
      "Visual: give the Bullet prefab its own SpriteRenderer.color or a small scale-up pulse that fades back to normal after a fraction of a second, so it reads clearly as \"a shot\" rather than a copy of the Player - not a permanent recolor.",
      "Structure: keep everything about how a bullet behaves - its own movement, its own off-screen cleanup - inside Bullet.cs, and keep Player.cs limited to just spawning it (Instantiate + cooldown + sound). Don't reach into a bullet's internals from the Player script, or vice versa.",
    ],
    trouble: [
      "Nothing happens on Space: confirm the Bullet Prefab field on the Player is actually assigned in the Inspector - a null reference here throws an error in the Console instead of silently failing.",
      "Bullet spawns but doesn't move or never disappears: confirm Bullet.cs is attached to the Bullet prefab itself (check the Project window, not just the scene), and that Update() is spelled exactly right (case-sensitive).",
      "Bullets pile up forever: double-check the y > 6f threshold roughly matches your camera's visible area - a Size-5 orthographic camera shows about -4.5 to 4.5 vertically, so 6f is safely just off-screen.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same spawn/move/cleanup pattern works in a Script Graph instead of C# - Unity's Visual Scripting graphs compile to the same underlying calls, using the same new Input System as the code path.",
      steps: [
        "Open PlayerControls (the Input Actions asset from Lesson 2's Player Input component) and add a new action to the Player action map named Fire, Action Type Button, bound to Space. Click Save Asset.",
        "On the Bullet prefab, add a Script Machine component and open a new Script Graph.",
        "Add an Update event node, a \"Translate\" node (Transform) feeding it Vector3.up * speed * deltaTime, and a Branch node checking the Y component of the position against your off-screen threshold, feeding a \"Destroy Game Object\" node on True.",
        "On the Player's own graph, right-click and add an \"On Input System Event Button\" node. Leave Target as This, set Input Action to Fire, and set Input Action Change Type to On Pressed. Wire its Trigger output into an \"Instantiate\" node (with the Bullet prefab as its asset input).",
      ],
      note: "Level 2/3: an On Input System Event Button node set to On Pressed already only fires once per physical press (event-driven, not polled), so no extra cooldown logic is required for Level 1/2. Level 3's ideas are all reachable with more nodes: an Audio Source node for sound, a \"Set Member\" node on the Sprite Renderer's color for the visual idea - same ideas as the C# path, just wired instead of typed.",
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: [
      "Open your [YourName]Game project from Lesson 2.",
      "Create a new sprite spr_bullet by importing bullet_yellow.png from this lesson's Assets folder.",
      "Create a new object obj_bullet, assign spr_bullet to it.",
    ],
    level1: [
      "Add a Step event to obj_bullet with the code below - it moves itself upward and destroys itself once it's off-screen.",
      "On obj_player (from Lesson 2), add to its existing Step event: a check for keyboard_check_pressed(vk_space) that calls instance_create_layer() to spawn an obj_bullet at the player's position.",
      "Drag obj_player into your room if it isn't already there, then press F5 (Run) and press spacebar to test - a bullet should spawn at the player and fly upward, then vanish.",
    ],
    level1Code: [
      "// obj_bullet's Step event:",
      "y -= 8;",
      "if (y < 0) instance_destroy();",
      "",
      "// Add to obj_player's existing Step event:",
      "if (keyboard_check_pressed(vk_space)) {",
      "    instance_create_layer(x, y, \"Instances\", obj_bullet);",
      "}",
    ],
    level2: [
      "Test firing several times quickly - each press should produce its own independent obj_bullet instance, all moving and destroying themselves correctly. Since instance_create_layer() creates a fully separate instance each call, this already works with no extra code.",
      "Add a cooldown so holding spacebar doesn't fire every frame: keyboard_check_pressed() already only returns true on the single frame a key is first pressed (not every frame held, unlike keyboard_check()), so Level 1 is already immune to this specific bug - add a fire_cooldown = 0; variable in obj_player's Create event anyway if you want a slower, tunable fire rate (count it down by 1 each Step, and only allow firing when it's <= 0, resetting it to e.g. 12 - about 0.2 seconds at 60fps - on each shot).",
    ],
    level3: [
      "Sound: add sfx_shoot.wav as a sound resource (keep its default name, sfx_shoot), then call audio_play_sound(sfx_shoot, 1, false) right next to the instance_create_layer() call.",
      "Visual: give obj_bullet its own image_blend color and a brief image_xscale/image_yscale pulse that reads clearly as \"a shot\" rather than a copy of the player, then fades back to normal size and color - not a permanent recolor. image_xscale/image_yscale scale around the sprite's Origin point, so set spr_bullet's Origin to Center in the Sprite Editor first (Sprite Editor > Origin) - scaling around any other origin (GameMaker's default is top-left) visibly shifts the bullet right and down as it grows, instead of scaling evenly in place.",
      "Structure: keep everything about how a bullet behaves - its own movement, its own off-screen cleanup - inside obj_bullet's own Step event, and keep obj_player's Step event limited to just spawning it (instance_create_layer + cooldown + sound). Don't reach into a bullet's variables from the player's code, or vice versa.",
    ],
    trouble: [
      "Nothing happens on spacebar: confirm the check is keyboard_check_pressed (fires once), not keyboard_check (fires every frame the key is held) - using the wrong one either does nothing visible different, or floods the room with bullets.",
      "Bullet spawns but doesn't move or never disappears: confirm the movement code is in obj_bullet's Step event, not its Create event - Create only runs once, when the instance is first spawned.",
      "\"instance_create_layer\" errors about a missing layer: confirm the room actually has an instance layer named \"Instances\" (GameMaker's default layer name) - check the Room Editor's Layers panel and adjust the layer name argument to match if it's been renamed.",
      "GML Visual only: bullet always spawns in the room's bottom-left or top-left corner instead of at the player: this means one of the Relative checkboxes on \"Jump To Point\" or \"Instance - Create Instance\" was left unchecked - every X and Y field on both blocks needs Relative checked, not just Y.",
    ],
    resource: "Student Resource List > GameMaker section, or gamemaker.io/en/tutorials",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same spawn/move/cleanup pattern works with drag-and-drop blocks instead of typed code. Pick whichever path feels right for you - both are graded the same.",
      steps: [
        "On obj_bullet's Step event, switch to GML Visual. From the Movement category, drag \"Jump To Point,\" set X to 0 and Y to -8, and check Relative for BOTH X and Y - Jump To Point is the dedicated block for moving relative to the current position, so prefer it over the generic Assign Variable/Set Variable block, but it only moves relative to the current position on the axes you actually check Relative for. Leaving X's Relative box unchecked (even with X set to 0) snaps the bullet to an absolute X of 0 every step. Then from Common, drag an \"If Variable\" block comparing y < 0, with an \"Instance - Destroy Instance\" block inside it.",
        "On obj_player's Step event, add an \"If Key Pressed\" block (Mouse and Keyboard category, Key set to Space - this is the single-press version, not the held-down \"If Key Down\" block from Lesson 2), with an \"Instance - Create Instance\" block inside it: set the Object dropdown to obj_bullet, set both its X and Y fields to 0, and check Relative for BOTH X and Y so the bullet spawns at the player's own current position instead of the room's absolute (0, 0).",
        "Drag obj_player into your room, then press F5 (Run) and test with spacebar - same result as the code path.",
      ],
      note: "Level 2/3: \"If Key Pressed\" already only fires once per press, the same as keyboard_check_pressed() in the code path, so no extra cooldown logic is required for Level 1/2. Level 3's ideas are all reachable with more blocks from the Audio and Instance categories - same ideas as the code path, just dragged instead of typed.",
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 2.",
      "Add a new Sprite object, name it Bullet, and assign bullet_yellow.png (import it from this lesson's Assets folder) as its image.",
      "Select Bullet, click \"Add a behavior,\" and add Destroy outside screen (default settings) - this is a built-in GDevelop behavior, already installed in every project.",
    ],
    level1: [
      "Add a new event: Condition \"Key pressed [Space]\" -> Action \"Create an object\" (choose Bullet, positioned at Player.X(), Player.Y()).",
      "Add a second event that always runs (no condition needed): Action \"Change the object's Y position\" for Bullet, subtracting a fixed amount times TimeDelta() each frame so it flies upward.",
      "Click the green Preview button and press Space to test - a bullet should spawn at the player and fly upward, then disappear near the top on its own (the Destroy outside screen behavior handles the cleanup automatically, no event needed)."
    ],
    level1Code: [
      "Condition: Key pressed [Space]",
      "Action: Create object Bullet at (Player.X(), Player.Y())",
      "",
      "(always, no condition)",
      "Action: Bullet.Y = Bullet.Y() - 400 * TimeDelta()",
      "",
      "(Bullet object has the Destroy outside screen behavior - handles cleanup with zero events)",
    ],
    level2: [
      "Test firing several times quickly - each press should produce its own independent Bullet instance, all moving and being destroyed correctly. Since \"Create an object\" makes a fully separate instance each time, and the movement/cleanup events apply to every Bullet instance automatically, this already works with no extra code.",
      "Add a cooldown so holding Space doesn't fire every frame: GDevelop's \"Key pressed\" condition is actually true on every frame the key is held (unlike some other engines' \"pressed\" checks), so without a cooldown this WILL flood the screen - add a scene Timer (\"Start (or reset) a scene timer\" once at the beginning of the scene), and change the fire condition to also require Timer value \"FireCooldown\" >= 0.2 seconds, resetting the timer on each shot.",
    ],
    level3: [
      "Sound: add a \"Play a sound\" action (using the imported sfx_shoot.wav) to the same event that creates a Bullet.",
      "Visual: give Bullet its own \"Tint color\" or a brief scale pulse via \"Change the scale of an object\" that fades back to normal after a fraction of a second, so it reads clearly as \"a shot\" rather than a copy of the player - not a permanent recolor.",
      "Structure: keep every Bullet-specific event (its movement, and anything from the ideas above) grouped together in one labeled event group, separate from the Player's own movement events from Lesson 2.",
    ],
    trouble: [
      "Bullets fire dozens of times per press: this is the \"Key pressed\" condition being level-triggered (true the whole time a key is held) rather than edge-triggered - see the Level 2 cooldown fix, which is required here, not optional.",
      "Bullet spawns but doesn't move: confirm the \"Change the object's Y position\" event has no condition attached (or a condition that's always true) - it needs to run every frame, not just on the frame the key was pressed.",
      "Bullet never disappears: confirm the Destroy outside screen behavior is actually attached to the Bullet object (check its Behaviors tab), not just described in a comment.",
    ],
    resource: "Student Resource List > GDevelop section, or gdevelop.io/tutorials",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: [
      "Launch TIC-80 and load your cart from Lesson 2 (or export/re-import it if you saved to disk - see the Lesson 1 Install Guide for TIC-80's save/load commands).",
      "Open the Sprite Editor (F2) and draw a small bullet shape on sprite index 2 (index 1 is still your player character from Lesson 2).",
    ],
    level1: [
      "Press Esc to reach the code editor.",
      "Add a bullets table near the top of your code (an empty Lua table that will hold one entry per active bullet).",
      "Inside your existing TIC() function, add a btnp(4) check (button 4 is Z, TIC-80's default \"A\" button) that inserts a new {x=x, y=y} entry into the bullets table when pressed.",
      "Add a loop that walks the bullets table backwards each frame, moving each bullet up and drawing it, removing any that have gone off the top of the screen.",
      "Press Esc, type run, press Enter, then press Z to test - a bullet should spawn at your character and fly upward, then vanish at the top.",
    ],
    level1Code: [
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
      "  if b.by<0 then",
      "   table.remove(bullets,i)",
      "  else",
      "   spr(2,b.bx,b.by,0)",
      "  end",
      " end",
      " spr(1,x,y,0)",
      "end",
    ],
    level2: [
      "Test firing several times quickly - each press should insert its own entry into the bullets table, all moving and being removed independently. Because the loop walks every entry in the table each frame, this already works with no extra code.",
      "Add a cooldown so holding Z doesn't fire every frame: btnp() already only returns true on the single frame a button is first pressed (not every frame held, unlike btn()), so Level 1 is already immune to this specific bug.",
    ],
    level3: [
      "Sound: play a \"pew\" sound the moment a bullet is fired. Open the SFX Editor (F4) and design a short descending-pitch effect in an empty slot, then call sfx(slot_number) right next to the table.insert() call.",
      "Visual: draw bullet sprite index 2 as visually distinct from the player (different shape or color), or use spr()'s flip/rotate argument to give the bullet a spin as it travels.",
      "Structure: keep the bullet-spawning check (btnp(4) + table.insert) and the bullet-updating loop (movement, drawing, and cleanup) as the two clearly separate pieces shown above, rather than folding bullet logic into the same lines that check player movement.",
    ],
    trouble: [
      "Bullets fire every frame instead of once per press: double-check you used btnp(4) (press) and not btn(4) (held) for the fire check.",
      "Bullets never appear: confirm you drew something on sprite index 2 in the Sprite Editor, and that spr(2, b.bx, b.by, 0) is inside the for loop, not outside it.",
      "\"attempt to index a nil value\" error: this usually means the bullets table wasn't declared before TIC() runs - confirm bullets={} is at the very top of your code, outside any function.",
    ],
    resource: "Student Resource List > TIC-80 section, or tic80.com",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 2.",
      "Create a new Actor Template for the projectile (Actor > Dynamic Object template works well) using bullet_yellow.png (import it from this lesson's Assets folder) as its Body sprite, and name it bullet.",
      "Place one bullet Actor on the Map, directly on top of main - this placed copy is only there so \"Map – Spawn Actor\" has a template to copy from; tick \"Initially Existent\" off for it, so it's invisible and inactive until spawned.",
      "On the bullet Actor's Automatic Movement settings, tick \"Move\" and set its direction to Up, at a comfortable speed - this makes every spawned copy fly upward on its own, no scripting needed for the movement itself.",
    ],
    level1: [
      "Go to Resources > Input Triggers, and add a new Control Combination for the Spacebar, set to \"Pressed.\"",
      "Click \"Edit Script\" for that trigger and add a \"Map – Spawn Actor\" event, targeting the bullet Actor.",
      "On the bullet Actor itself, add an Actor Trigger: \"Left / Entered Camera\" (the Left case) -> a script with an \"Actor – Destroy/Revive Actor\" event (Destroy) targeting itself (relate/this).",
      "Save, then click Play/Test and press spacebar - a bullet should spawn at its placed position and fly upward, then disappear once it leaves the camera view.",
    ],
    level2: [
      "Test firing several times quickly - each press should spawn its own independent bullet, all moving and being destroyed correctly. \"Map – Spawn Actor\" creates a fresh copy each time it runs, so this already works with no extra scripting.",
      "Add a cooldown so holding spacebar doesn't fire every frame: Input Triggers set to \"Pressed\" already only fire once per physical press (not every frame held), so Level 1 is already immune to this specific bug.",
    ],
    level3: [
      "Sound: import sfx_shoot.wav (Resources > Import > Sound), then add a \"Play Sound\" event directly after the \"Map – Spawn Actor\" event on the fire trigger's script.",
      "Visual: give the bullet Actor its own Tint/Opacity setting (Actor > Graphics Options) distinct from main's, so it reads clearly as \"a shot\" rather than a copy of the player.",
      "Structure: keep everything about how a bullet behaves - its automatic movement setting and its own \"Left/Entered Camera\" cleanup trigger - defined on the bullet Actor itself, and keep main's Input Trigger script limited to just spawning it (Spawn Actor + sound). Don't add bullet-specific logic to main's scripts, or vice versa.",
    ],
    trouble: [
      "Nothing happens on spacebar: confirm the Input Trigger is actually set to \"Pressed,\" and that its script is connected to the Start node in the Script Editor.",
      "Bullet spawns but doesn't move: confirm \"Move\" is ticked in the bullet Actor's Automatic Movement window, with a direction and speed actually set (not left at 0).",
      "Bullet never disappears: confirm the \"Left/Entered Camera\" trigger is on the bullet Actor itself (not main), and that it's specifically checking for the \"Left\" case, not \"Entered.\"",
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
        "Open the Script Editor for a Spacebar Input Trigger, same as before.",
        "Instead of dragging graphical event blocks, type the command directly - start typing \"Spawn\" and use the editor's autocomplete/command list to find the text equivalent of the Map – Spawn Actor event you used in the graphical path.",
        "Save and Test to confirm the same result as your graphical version.",
      ],
      note: "Exact command names can vary slightly by version - the Script Editor's autocomplete and the 001gamecreator.com/wiki Textual Scripting reference show the current syntax. If you want to keep exploring this mode, do it in your throwaway test project, not your real one.",
    },
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 2.",
      "Import bullet_yellow.png (this lesson's Assets folder) into the Content Browser the same way Lesson 2 imported the player circles, then build M_Bullet_Yellow the same way Lesson 2 built M_Player_Orange - a new Material, Shading Model = Unlit, a Texture Sample using bullet_yellow.png wired straight into Emissive Color.",
      "Create a new Blueprint class (Actor parent), name it BP_Bullet. Add a Static Mesh Component to it and set its shape to a small basic Shape (e.g. Sphere, built into the Blank template). Assign M_Bullet_Yellow to it in the Static Mesh Component's Material slot - a default Shape's built-in material is Lit, which renders solid black with this project's lights removed, so every new Shape from here on needs its own Unlit Material assigned explicitly.",
      "Open BP_Bullet and add a public float variable Speed (default something like 400).",
      "In the Content Browser, click Add (+) > Input > Input Action, name it IA_Fire. Leave its Value Type as the default (Digital bool), same as the 4 move actions from Lesson 2.",
      "Open IMC_Default (the Input Mapping Context you built in Lesson 2) and click + Add Mapping. Choose IA_Fire, then set its Key to Space Bar. Because IMC_Default was already added to the player in Lesson 2's BP_Player BeginPlay, this new mapping is live immediately - no extra Add Mapping Context step is needed just for adding one more action to a context that's already active.",
    ],
    level1: [
      "In BP_Bullet's Event Graph, add an Event Tick node feeding an \"Add Actor World Offset\" node. Right-click its Delta Location pin and choose \"Split Struct Pin\" first (it starts collapsed as one Vector pin, and won't accept a value into just Y until it's split), then wire Speed * Delta Seconds into Y. Add a Branch checking its own Y location against an off-screen threshold, feeding a \"Destroy Actor\" node on True.",
      "Back in BP_Player's own Event Graph (from Lesson 2 - not the Level Blueprint), right-click and search for IA_Fire to add an Enhanced Input Action Event node, and wire its Started pin to a \"Spawn Actor from Class\" node (Class = BP_Bullet). For Transform, add a \"Get Actor Transform\" node with its Target left unwired (defaults to Self, same as every other node this course has you leave unwired inside BP_Player), and wire its output into Spawn Actor from Class's Transform pin - that's what spawns the bullet at BP_Player's own current position and facing.",
      "Click Compile, then Play and press spacebar to test - a bullet should spawn at BP_Player and fly upward, then disappear.",
    ],
    level2: [
      "Test firing several times quickly - each press should spawn its own independent BP_Bullet instance, all moving and destroying themselves correctly. Since \"Spawn Actor from Class\" creates a fully separate Actor each call, this already works with no extra nodes.",
      "Add a cooldown so holding spacebar doesn't fire every frame: an Enhanced Input Action Event's Started pin only fires once per physical press (not every frame held, since it's event-driven rather than polled), so Level 1 is already immune to this specific bug - add a Retriggerable Delay node after the Started pin anyway if you want a slower, tunable fire rate (e.g. 0.2 seconds), gating the Spawn Actor node behind it.",
    ],
    level3: [
      "Sound: add a Play Sound 2D node (using an imported sfx_shoot.wav Sound Wave asset) chained after the Spawn Actor from Class node.",
      "Visual: give BP_Bullet a brief scale pulse on spawn that eases back to normal size after a fraction of a second - Set Actor Relative Scale3D up (e.g. 1.5, 1.5, 1.5), chain a Delay node, then set it back to (1, 1, 1). BP_Bullet's own sprite (bullet_yellow.png, from Setup) already makes it read clearly as \"a shot\" rather than a copy of the player Cube, so this idea is a pure size flourish, not a permanent recolor.",
      "Structure: keep everything about how a bullet behaves - its own movement, its own off-screen cleanup - inside BP_Bullet's own Event Graph, and keep BP_Player limited to just spawning it (Spawn Actor from Class + cooldown + sound). Don't add bullet-specific logic to BP_Player, or vice versa.",
    ],
    trouble: [
      "Nothing happens on spacebar: confirm you clicked Compile after editing, that IA_Fire actually has a Space Bar mapping inside IMC_Default (not some other, unmapped context), and that the white execution wire connects the Enhanced Input Action Event's Started pin through to Spawn Actor from Class.",
      "Bullet spawns but doesn't move or never disappears: confirm BP_Bullet's own Event Tick is wired up inside BP_Bullet itself, not accidentally built in BP_Player.",
      "Bullet renders solid black instead of yellow: open M_Bullet_Yellow and confirm its Texture Sample is actually wired into Emissive Color, and that bullet_yellow.png is actually assigned as its Texture - the same check as Lesson 2's M_Player_Orange.",
      "This is the most hardware-demanding engine on our list - if your laptop struggles, it's completely fine to observe this week's demo and build the exercise in a lighter engine instead.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same spawn/move/cleanup pattern works in C++ instead of Blueprint nodes.",
      steps: [
        "Tools > New C++ Class, choose Actor as the parent class, name it CppBullet, click Create Class.",
        "In CppBullet.h, declare a float Speed and override Tick().",
        "In CppBullet.cpp's Tick(), call AddActorWorldOffset() along Y, then check GetActorLocation().Y against your off-screen threshold and call Destroy() when past it.",
        "In your player Actor's class (the same one from Lesson 2's C++ path), add a UInputAction* FireAction UPROPERTY, bind it in SetupPlayerInputComponent() with EnhancedInputComponent->BindAction(FireAction, ETriggerEvent::Started, this, &Class::Fire) alongside the move bindings, then call GetWorld()->SpawnActor<ACppBullet>() inside that Fire() handler, passing the player's current transform.",
        "Compile, assign FireAction to the IA_Fire asset in the Details panel, Play, and test with spacebar.",
      ],
      code: [
        "// CppBullet.h",
        "virtual void Tick(float DeltaSeconds) override;",
        "float Speed = 400.f;",
        "",
        "// CppBullet.cpp",
        "void ACppBullet::Tick(float DeltaSeconds) {",
        "    Super::Tick(DeltaSeconds);",
        "    AddActorWorldOffset(FVector(0.f, Speed * DeltaSeconds, 0.f));",
        "    if (GetActorLocation().Y > 600.f) Destroy();",
        "}",
        "",
        "// Player class: UPROPERTY(EditAnywhere) class UInputAction* FireAction;",
        "// In SetupPlayerInputComponent(), alongside the Lesson 2 move bindings:",
        "EnhancedInput->BindAction(FireAction, ETriggerEvent::Started, this, &APlayerActor::Fire);",
        "",
        "void APlayerActor::Fire() {",
        "    GetWorld()->SpawnActor<ACppBullet>(BulletClass, GetActorLocation(), FRotator::ZeroRotator);",
        "}",
      ],
      note: "This mirrors the Blueprint path's Spawn Actor from Class node and BP_Bullet's own Event Tick - same underlying engine calls, just from code. FireAction reuses the exact same IMC_Default and AddMappingContext() call already set up in Lesson 2, so no new mapping-context wiring is needed here, only the new IA_Fire asset and its BindAction. Level 2/3: Enhanced Input's Started trigger only fires once per press in C++ too, so no extra cooldown logic is required for Level 1/2, and UGameplayStatics::PlaySound2D() plus SetActorRelativeScale3D() cover the Level 3 sound and visual ideas.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_03_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS (one per engine) - instructor only
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.0, { after: 120 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Lesson 3 Teacher's Key", bold: true, size: 30, color: brand.NAVY })] }),
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
      headers: { default: brand.pageHeader("Lesson 3 Teacher's Key - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    l1Notes: ["Two separate scripts on two separate sprites - this separation matters, since combining create-clone and delete-this-clone in the same \"when I start as a clone\" stack can freeze Scratch:"],
    l1Code: [
      "-- On the Player sprite:",
      "when [space] key pressed",
      "  create clone of [Bullet]",
      "",
      "-- On the Bullet sprite:",
      "when green flag clicked",
      "  hide",
      "",
      "when I start as a clone",
      "  show",
      "  go to x: (Player's x) y: (Player's y)",
      "  forever",
      "    change y by 10",
      "    if <y position > 170> then",
      "      delete this clone",
    ],
    l2Notes: ["Add a cooldown variable (can_shoot) on the Player so holding spacebar can't fire every single frame:"],
    l2Code: [
      "-- On the Player sprite:",
      "when green flag clicked",
      "  set [can_shoot] to true",
      "",
      "when [space] key pressed",
      "  if <can_shoot> then",
      "    create clone of [Bullet]",
      "    set [can_shoot] to false",
      "    wait 0.2 seconds",
      "    set [can_shoot] to true",
    ],
    l3Ideas: [
      {
        title: "Sound: play on fire",
        notes: ["Upload sfx_shoot.wav from this lesson's Assets folder, then add \"start sound\" right next to the clone-creation block - \"start sound\" doesn't block or loop, so it's fire-and-forget, the same property that made Lesson 2's beep-beep idea work:"],
        code: ["when [space] key pressed", "  if <can_shoot> then", "    create clone of [Bullet]", "    start sound [sfx_shoot]", "    set [can_shoot] to false", "    wait 0.2 seconds", "    set [can_shoot] to true"],
      },
      {
        title: "Visual: give the bullet its own look",
        notes: ["A quick \"change size by\" pulse on spawn reads clearly as a muzzle-flash-style pop, without needing a second costume:"],
        code: ["when I start as a clone", "  show", "  set size to 150%", "  go to x: (Player's x) y: (Player's y)", "  wait 0.05 seconds", "  set size to 100%"],
      },
      {
        title: "Structure: keep bullet behavior on the Bullet sprite",
        notes: ["The Player's fire script should only ever trigger a shot (create clone, cooldown, sound) - everything about what a bullet actually does once it exists (appearing, moving, cleaning itself up) belongs on the Bullet sprite's own \"when I start as a clone\" script. Keeping this boundary clean is what let the Sound and Visual ideas above drop in with one line each, instead of needing to touch the Player's code."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, across the Player and Bullet sprites:"],
        code: [
          "-- Player sprite:",
          "when green flag clicked",
          "  set [can_shoot] to true",
          "",
          "when [space] key pressed",
          "  if <can_shoot> then",
          "    create clone of [Bullet]",
          "    start sound [sfx_shoot]",
          "    set [can_shoot] to false",
          "    wait 0.2 seconds",
          "    set [can_shoot] to true",
          "",
          "-- Bullet sprite:",
          "when green flag clicked",
          "  hide",
          "",
          "when I start as a clone",
          "  show",
          "  set size to 150%",
          "  go to x: (Player's x) y: (Player's y)",
          "  wait 0.05 seconds",
          "  set size to 100%",
          "  forever",
          "    change y by 10",
          "    if <y position > 170> then",
          "      delete this clone",
        ],
      },
      {
        title: "Super-Stretch (not required): poll input instead of reacting to it, so firing can't interrupt movement",
        notes: [
          "This is the worked fix for the known Scratch quirk noted in the Handout's Trouble section and this lesson's Sidebars: \"when key pressed\" scripts repeat by riding the operating system's own keyboard auto-repeat, and most operating systems only actively auto-repeat one key at a time - so pressing Space to fire while holding a direction key can interrupt that direction key's repeat stream, pausing movement until it's pressed again.",
          "The fix is to stop depending on repeated key-press events entirely and poll real key state instead, using the Sensing category's \"key ( ) pressed?\" reporter inside a \"forever\" loop - the same technique Lesson 2's blink stretch goal already uses, now applied to movement and firing too. Because this reads whether a key is currently held down, once per frame, it's completely unaffected by the operating system's one-key-at-a-time auto-repeat limit.",
          "Two separate \"when green flag clicked\" scripts on the Player, not one - movement polls independently from firing, so the \"wait 0.2 seconds\" fire cooldown pauses only the fire script's own thread, not movement. Combining both into a single forever loop would make movement visibly freeze for the whole cooldown every time you fire, trading one bug for another:",
        ],
        code: [
          "-- Player sprite, movement script (replaces the 4 \"when [arrow] key pressed\" scripts):",
          "when green flag clicked",
          "  forever",
          "    if <key [right arrow] pressed?> then",
          "      change x by 10",
          "    if <key [left arrow] pressed?> then",
          "      change x by -10",
          "    if <key [up arrow] pressed?> then",
          "      change y by 10",
          "    if <key [down arrow] pressed?> then",
          "      change y by -10",
          "",
          "-- Player sprite, fire script (replaces \"when [space] key pressed\"):",
          "when green flag clicked",
          "  set [can_shoot] to true",
          "  forever",
          "    if <<key [space] pressed?> and <can_shoot>> then",
          "      create clone of [Bullet]",
          "      start sound [sfx_shoot]",
          "      set [can_shoot] to false",
          "      wait 0.2 seconds",
          "      set [can_shoot] to true",
        ],
      },
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    l1Notes: ["Two scripts: one attached to bullet.tscn's Sprite2D root, one added to the Player's existing script:"],
    l1Code: [
      "# bullet.tscn's script:",
      "extends Sprite2D",
      "var speed := 400.0",
      "",
      "func _process(delta):",
      "    position.y -= speed * delta",
      "    if position.y < -20:",
      "        queue_free()",
      "",
      "# Player's script - add near the top and inside _process:",
      "var bullet_scene := preload(\"res://bullet.tscn\")",
      "",
      "func _process(delta):",
      "    if Input.is_action_just_pressed(\"ui_select\"):",
      "        var b := bullet_scene.instantiate()",
      "        b.global_position = global_position",
      "        get_parent().add_child(b)",
    ],
    l2Notes: ["is_action_just_pressed() is already edge-triggered (true for exactly one frame per physical press), so Level 1's fire check needs no changes for Level 2. An optional tunable cooldown, if a slower fire rate is wanted:"],
    l2Code: [
      "var fire_cooldown := 0.0",
      "",
      "func _process(delta):",
      "    fire_cooldown -= delta",
      "    if Input.is_action_just_pressed(\"ui_select\") and fire_cooldown <= 0.0:",
      "        var b := bullet_scene.instantiate()",
      "        b.global_position = global_position",
      "        get_parent().add_child(b)",
      "        fire_cooldown = 0.2",
    ],
    l3Ideas: [
      {
        title: "Sound: play on fire",
        notes: ["Add an AudioStreamPlayer2D node to the Player, Stream set to sfx_shoot.wav:"],
        code: ["if Input.is_action_just_pressed(\"ui_select\") and fire_cooldown <= 0.0:", "    var b := bullet_scene.instantiate()", "    b.global_position = global_position", "    get_parent().add_child(b)", "    fire_cooldown = 0.2", "    $AudioStreamPlayer2D.play()"],
      },
      {
        title: "Visual: give the bullet its own look",
        notes: ["Set modulate in the bullet's own _ready() so every spawned copy looks distinct without touching the Player's script at all - then count a short flash_timer back down in _process(delta) so the flash actually fades instead of staying permanent:"],
        code: ["# bullet.tscn's script:", "var flash_timer := 0.1", "", "func _ready():", "    modulate = Color.YELLOW", "    scale = Vector2(1.3, 1.3)", "", "func _process(delta):", "    if flash_timer > 0.0:", "        flash_timer -= delta", "        if flash_timer <= 0.0:", "            modulate = Color.WHITE", "            scale = Vector2(1, 1)"],
      },
      {
        title: "Structure: keep bullet behavior in bullet.tscn's own script",
        notes: ["The Player's script should only ever trigger a shot (instantiate, position, cooldown, sound) - everything about what a bullet actually does once it exists (movement, off-screen cleanup, its own look) belongs in bullet.tscn's own script. Keeping this boundary clean is what let the Sound and Visual ideas above each drop into one file without touching the other."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, across the Player's script and bullet.tscn's own script:"],
        code: [
          "# bullet.tscn's script:",
          "extends Sprite2D",
          "var speed := 400.0",
          "var flash_timer := 0.1",
          "",
          "func _ready():",
          "    modulate = Color.YELLOW",
          "    scale = Vector2(1.3, 1.3)",
          "",
          "func _process(delta):",
          "    position.y -= speed * delta",
          "    if position.y < -20:",
          "        queue_free()",
          "    if flash_timer > 0.0:",
          "        flash_timer -= delta",
          "        if flash_timer <= 0.0:",
          "            modulate = Color.WHITE",
          "            scale = Vector2(1, 1)",
          "",
          "# Player's script:",
          "var bullet_scene := preload(\"res://bullet.tscn\")",
          "var fire_cooldown := 0.0",
          "",
          "func _process(delta):",
          "    fire_cooldown -= delta",
          "    if Input.is_action_just_pressed(\"ui_select\") and fire_cooldown <= 0.0:",
          "        var b := bullet_scene.instantiate()",
          "        b.global_position = global_position",
          "        get_parent().add_child(b)",
          "        fire_cooldown = 0.2",
          "        $AudioStreamPlayer2D.play()",
        ],
      },
    ],
  },
  {
    engine: "Unity", file: "Unity",
    l1Notes: ["Two scripts: Bullet.cs on the Bullet prefab, and an addition to the Player's existing script:"],
    l1Code: [
      "// Bullet.cs, attached to the Bullet prefab:",
      "public float speed = 6f;",
      "",
      "void Update() {",
      "    transform.Translate(Vector3.up * speed * Time.deltaTime);",
      "    if (transform.position.y > 6f) Destroy(gameObject);",
      "}",
      "",
      "// Player.cs - add a field and a check inside Update():",
      "public GameObject bulletPrefab;",
      "",
      "if (Keyboard.current.spaceKey.wasPressedThisFrame) {",
      "    Instantiate(bulletPrefab, transform.position, Quaternion.identity);",
      "}",
    ],
    l2Notes: ["wasPressedThisFrame is already edge-triggered (true for exactly one frame per physical press), so Level 1's fire check needs no changes for Level 2. An optional tunable cooldown, if a slower fire rate is wanted:"],
    l2Code: [
      "private float fireCooldown = 0f;",
      "",
      "void Update() {",
      "    fireCooldown -= Time.deltaTime;",
      "    if (Keyboard.current.spaceKey.wasPressedThisFrame && fireCooldown <= 0f) {",
      "        Instantiate(bulletPrefab, transform.position, Quaternion.identity);",
      "        fireCooldown = 0.2f;",
      "    }",
      "}",
    ],
    l3Ideas: [
      {
        title: "Sound: play on fire",
        notes: ["Add an AudioSource component to the Player, AudioClip set to sfx_shoot.wav:"],
        code: ["public AudioSource audioSrc;", "", "if (Keyboard.current.spaceKey.wasPressedThisFrame && fireCooldown <= 0f) {", "    Instantiate(bulletPrefab, transform.position, Quaternion.identity);", "    fireCooldown = 0.2f;", "    audioSrc.Play();", "}"],
      },
      {
        title: "Visual: give the bullet its own look",
        notes: ["Set color and scale in the Bullet script's own Start() so every spawned copy looks distinct without touching the Player script at all - then count a short flashTimer back down in Update() so the flash actually fades instead of staying permanent:"],
        code: ["private float flashTimer;", "", "void Start() {", "    GetComponent<SpriteRenderer>().color = Color.yellow;", "    transform.localScale *= 1.3f;", "    flashTimer = 0.1f;", "}", "", "void Update() {", "    if (flashTimer > 0f) {", "        flashTimer -= Time.deltaTime;", "        if (flashTimer <= 0f) {", "            GetComponent<SpriteRenderer>().color = Color.white;", "            transform.localScale /= 1.3f;", "        }", "    }", "}"],
      },
      {
        title: "Structure: keep bullet behavior in Bullet.cs",
        notes: ["Player.cs should only ever trigger a shot (Instantiate, cooldown, sound) - everything about what a bullet actually does once it exists (movement, off-screen cleanup, its own look) belongs in Bullet.cs. Keeping this boundary clean is what let the Sound and Visual ideas above each drop into one file without touching the other."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, across Bullet.cs and Player.cs:"],
        code: [
          "// Bullet.cs:",
          "public float speed = 6f;",
          "private float flashTimer;",
          "",
          "void Start() {",
          "    GetComponent<SpriteRenderer>().color = Color.yellow;",
          "    transform.localScale *= 1.3f;",
          "    flashTimer = 0.1f;",
          "}",
          "",
          "void Update() {",
          "    transform.Translate(Vector3.up * speed * Time.deltaTime);",
          "    if (transform.position.y > 6f) Destroy(gameObject);",
          "",
          "    if (flashTimer > 0f) {",
          "        flashTimer -= Time.deltaTime;",
          "        if (flashTimer <= 0f) {",
          "            GetComponent<SpriteRenderer>().color = Color.white;",
          "            transform.localScale /= 1.3f;",
          "        }",
          "    }",
          "}",
          "",
          "// Player.cs:",
          "public GameObject bulletPrefab;",
          "public AudioSource audioSrc;",
          "private float fireCooldown = 0f;",
          "",
          "void Update() {",
          "    fireCooldown -= Time.deltaTime;",
          "    if (Keyboard.current.spaceKey.wasPressedThisFrame && fireCooldown <= 0f) {",
          "        Instantiate(bulletPrefab, transform.position, Quaternion.identity);",
          "        fireCooldown = 0.2f;",
          "        audioSrc.Play();",
          "    }",
          "}",
        ],
      },
    ],
    altNotes: {
      title: "Visual Scripting",
      notes: [
        "The Fire action reuses the same PlayerControls Input Actions asset and Player Input component built in Lesson 2 - just one new Button action added to the existing Player action map, no new component setup required.",
        "wasPressedThisFrame's single-frame-per-press behavior has a direct Script Graph equivalent: an On Input System Event Button node set to On Pressed only triggers once per physical press, so the same no-extra-cooldown-needed-for-Level-2 logic applies in graphs too.",
        "The Level 3 goals map onto nodes the same way they map onto code: an Audio Source node for sound, \"Set Member\" nodes on the Sprite Renderer's color and the Transform's scale for the visual idea - all inside the Bullet prefab's own graph, kept separate from the Player's graph the same way Bullet.cs and Player.cs stay separate in the code path. The flash-then-fade behavior needs a Timer node (or a manually counted-down float variable checked each Update) so the same \"Set Member\" nodes fire a second time once the timer runs out, reverting color to white and dividing scale back down - without it, the color/scale change from Start stays permanent.",
      ],
      why: [
        "Keeping the same Structure principle (bullet behavior lives with the bullet, not the player) whether a student is in C# or graphs reinforces that it's a design idea, not a syntax rule.",
        "Reusing Lesson 2's PlayerControls asset instead of building a second, separate Input Actions asset keeps every Input System-driven action for the whole capstone game consolidated in one place.",
      ],
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    l1Notes: ["obj_bullet's own Step event, plus an addition to obj_player's existing Step event:"],
    l1Code: [
      "// obj_bullet's Step event:",
      "y -= 8;",
      "if (y < 0) instance_destroy();",
      "",
      "// Add to obj_player's existing Step event:",
      "if (keyboard_check_pressed(vk_space)) {",
      "    instance_create_layer(x, y, \"Instances\", obj_bullet);",
      "}",
    ],
    l2Notes: ["keyboard_check_pressed() is already edge-triggered (true for exactly one frame per physical press), so Level 1's fire check needs no changes for Level 2. An optional tunable cooldown, if a slower fire rate is wanted (add fire_cooldown = 0; to obj_player's Create event):"],
    l2Code: [
      "// obj_player's Step event:",
      "if (fire_cooldown > 0) fire_cooldown -= 1;",
      "if (keyboard_check_pressed(vk_space) && fire_cooldown <= 0) {",
      "    instance_create_layer(x, y, \"Instances\", obj_bullet);",
      "    fire_cooldown = 12;",
      "}",
    ],
    l3Ideas: [
      {
        title: "Sound: play on fire",
        notes: ["Add sfx_shoot.wav as a sound resource (keep its default name, sfx_shoot):"],
        code: ["if (keyboard_check_pressed(vk_space) && fire_cooldown <= 0) {", "    instance_create_layer(x, y, \"Instances\", obj_bullet);", "    fire_cooldown = 12;", "    audio_play_sound(sfx_shoot, 1, false);", "}"],
      },
      {
        title: "Visual: give the bullet its own look",
        notes: [
          "First, in the Sprite Editor, set spr_bullet's Origin to Center. image_xscale/image_yscale scale the sprite around its Origin point, not around its visual center - with GameMaker's default top-left Origin, scaling up makes the bullet visibly grow to the right and down (a real position shift, not just a size change) instead of scaling evenly in place. This is a one-time sprite setting, not code.",
          "The flash itself needs to be temporary, so it needs a countdown timer to reset it - the same fire_cooldown-style pattern from Level 2, just living on obj_bullet instead of obj_player. Set the flashy values plus a timer in Create, then count the timer down in Step and reset back to normal the instant it runs out:",
        ],
        code: [
          "// obj_bullet's Create event:",
          "image_blend = c_yellow;",
          "image_xscale = 1.3;",
          "image_yscale = 1.3;",
          "flash_timer = 6; // about 0.1 seconds at 60fps",
          "",
          "// obj_bullet's Step event - add alongside the existing movement code:",
          "if (flash_timer > 0) {",
          "    flash_timer -= 1;",
          "    if (flash_timer <= 0) {",
          "        image_blend = c_white;",
          "        image_xscale = 1;",
          "        image_yscale = 1;",
          "    }",
          "}",
        ],
      },
      {
        title: "Structure: keep bullet behavior in obj_bullet",
        notes: ["obj_player's Step event should only ever trigger a shot (instance_create_layer, cooldown, sound) - everything about what a bullet actually does once it exists (movement, off-screen cleanup, its own look) belongs in obj_bullet's own events. Keeping this boundary clean is what let the Sound and Visual ideas above each drop into one object without touching the other."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, across obj_bullet and obj_player:"],
        code: [
          "// spr_bullet's Origin (Sprite Editor, one-time setting): Center",
          "",
          "// obj_bullet's Create event:",
          "image_blend = c_yellow;",
          "image_xscale = 1.3;",
          "image_yscale = 1.3;",
          "flash_timer = 6; // about 0.1 seconds at 60fps",
          "",
          "// obj_bullet's Step event:",
          "y -= 8;",
          "if (y < 0) instance_destroy();",
          "if (flash_timer > 0) {",
          "    flash_timer -= 1;",
          "    if (flash_timer <= 0) {",
          "        image_blend = c_white;",
          "        image_xscale = 1;",
          "        image_yscale = 1;",
          "    }",
          "}",
          "",
          "// obj_player's Create event:",
          "fire_cooldown = 0;",
          "",
          "// obj_player's Step event:",
          "if (fire_cooldown > 0) fire_cooldown -= 1;",
          "if (keyboard_check_pressed(vk_space) && fire_cooldown <= 0) {",
          "    instance_create_layer(x, y, \"Instances\", obj_bullet);",
          "    fire_cooldown = 12;",
          "    audio_play_sound(sfx_shoot, 1, false);",
          "}",
        ],
      },
    ],
    altNotes: {
      title: "GML Visual",
      notes: [
        "\"If Key Pressed\" is already edge-triggered the same as keyboard_check_pressed(), so the same no-extra-cooldown-needed-for-Level-2 logic applies in blocks too.",
        "Movement uses \"Jump To Point\" (Movement category) with Relative checked on BOTH X and Y, the same purpose-built block from Level 1, rather than the generic Assign Variable/Set Variable - watch for students who only check Relative on the axis that's changing (Y) and leave X unchecked, which silently snaps X to an absolute 0 every step. \"Instance - Create Instance\" needs the same both-axes Relative treatment when spawning obj_bullet at the player's position.",
        "The Level 3 goals map onto blocks the same way they map onto code, favoring the most specific named block available: \"Play Audio\" (Audio category) for sound, and \"Set Instance Variable\" for image_blend/image_xscale/image_yscale/flash_timer - all recognized built-ins on Set Instance Variable's own dropdown (Image Blend, Horizontal Scale, Vertical Scale), so there's no need for the generic Assign Variable here either. The Visual idea's flash needs spr_bullet's Origin set to Center first (Sprite Editor) - image_xscale/image_yscale scale around the Origin, and GameMaker's default top-left Origin makes the bullet visibly shift right and down as it scales instead of growing in place - and needs an \"If Variable\"-gated countdown on flash_timer in Step to actually reset image_blend/image_xscale/image_yscale back to normal, the same shape as the code path's timer. All of this stays inside obj_bullet's own events, kept separate from obj_player's events the same way the code path keeps the two objects separate.",
      ],
      why: [
        "Keeping the same Structure principle (bullet behavior lives with the bullet, not the player) whether a student is in Code or Visual reinforces that it's a design idea, not a syntax rule.",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    l1Notes: ["Two events: one to spawn, one that always moves every existing Bullet instance. The Destroy outside screen behavior (added once, on the Bullet object) handles cleanup with zero events:"],
    l1Code: [
      "Condition: Key pressed [Space]",
      "Action: Create object Bullet at (Player.X(), Player.Y())",
      "",
      "(always, no condition)",
      "Action: Bullet.Y = Bullet.Y() - 400 * TimeDelta()",
      "",
      "(Bullet object has the Destroy outside screen behavior)",
    ],
    l2Notes: ["Unlike the code-based engines, GDevelop's \"Key pressed\" condition is level-triggered (true on every frame the key is held), so a cooldown is required here, not optional, to avoid flooding the screen with bullets:"],
    l2Code: [
      "At the beginning of the scene:",
      "Action: Start (or reset) a scene timer \"FireCooldown\"",
      "",
      "Every frame:",
      "Condition: Key pressed [Space]",
      "Condition: Timer value \"FireCooldown\" >= 0.2 seconds",
      "Action: Create object Bullet at (Player.X(), Player.Y())",
      "Action: Start (or reset) a scene timer \"FireCooldown\"",
    ],
    l3Ideas: [
      {
        title: "Sound: play on fire",
        notes: ["Add a \"Play a sound\" action (using sfx_shoot.wav) to the same event that creates a Bullet:"],
        code: ["Condition: Key pressed [Space]", "Condition: Timer value \"FireCooldown\" >= 0.2 seconds", "Action: Create object Bullet at (Player.X(), Player.Y())", "Action: Start (or reset) a scene timer \"FireCooldown\"", "Action: Play a sound (sfx_shoot.wav)"],
      },
      {
        title: "Visual: give the bullet its own look",
        notes: ["\"Tint color\" and \"Change the scale of an object\" applied when the Bullet is created, so every spawned copy looks distinct - plus Bullet's own per-instance object timer (not the scene's shared \"FireCooldown\" timer) counting down in a second, always-running event so the flash actually fades back to normal instead of staying permanent:"],
        code: [
          "Condition: Key pressed [Space]",
          "Condition: Timer value \"FireCooldown\" >= 0.2 seconds",
          "Action: Create object Bullet at (Player.X(), Player.Y())",
          "Action: Tint color of Bullet = Yellow",
          "Action: Change the scale of Bullet to 1.3",
          "Action: Start (or reset) the object timer \"Flash\" of Bullet",
          "",
          "Every frame (separate event, no Space condition):",
          "Condition: Bullet - Timer value \"Flash\" >= 0.1 seconds",
          "Action: Tint color of Bullet = White",
          "Action: Change the scale of Bullet to 1",
        ],
      },
      {
        title: "Structure: group Bullet-specific events together",
        notes: ["Keep every Bullet-related event - the fire trigger, the always-on movement, and anything from the ideas above - inside one labeled event group (e.g. \"Shooting\"), separate from the Player's own movement events carried over from Lesson 2. GDevelop's event sheet has no hard file boundary between objects the way separate scripts/objects do in other engines, so this labeled grouping is what keeps the two mechanics from blurring together."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above, organized as one event group named \"Shooting\":"],
        code: [
          "At the beginning of the scene:",
          "Action: Start (or reset) a scene timer \"FireCooldown\"",
          "",
          "Event group: \"Shooting\"",
          "",
          "  Condition: Key pressed [Space]",
          "  Condition: Timer value \"FireCooldown\" >= 0.2 seconds",
          "  Action: Create object Bullet at (Player.X(), Player.Y())",
          "  Action: Start (or reset) a scene timer \"FireCooldown\"",
          "  Action: Play a sound (sfx_shoot.wav)",
          "  Action: Tint color of Bullet = Yellow",
          "  Action: Change the scale of Bullet to 1.3",
          "  Action: Start (or reset) the object timer \"Flash\" of Bullet",
          "",
          "  (always, no condition)",
          "  Action: Bullet.Y = Bullet.Y() - 400 * TimeDelta()",
          "",
          "  Condition: Bullet - Timer value \"Flash\" >= 0.1 seconds",
          "  Action: Tint color of Bullet = White",
          "  Action: Change the scale of Bullet to 1",
          "",
          "(Bullet object has the Destroy outside screen behavior - handles cleanup with zero events)",
        ],
      },
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    l1Notes: ["A bullets table, a spawn check, and an update/cleanup loop, all inside the existing TIC() function:"],
    l1Code: [
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
      "  if b.by<0 then",
      "   table.remove(bullets,i)",
      "  else",
      "   spr(2,b.bx,b.by,0)",
      "  end",
      " end",
      " spr(1,x,y,0)",
      "end",
    ],
    l2Notes: ["btnp() is already edge-triggered (true for exactly one frame per physical press, confirmed against TIC-80's own documentation - it's specifically designed as the \"just pressed\" check, as opposed to btn()'s \"currently held\"), so Level 1's fire check needs no changes for Level 2."],
    l3Ideas: [
      {
        title: "Sound: play on fire",
        notes: ["Design a short descending-pitch effect in the SFX Editor (F4) in an empty slot (e.g. slot 1), then call sfx(1) right next to the table.insert() call - a single call, not looped, so it's fire-and-forget the same way Lesson 2's beep-beep idea needed an explicit sfx(-1) stop but a one-shot fire sound doesn't:"],
        code: [" if btnp(4) then", "  table.insert(bullets,{bx=x,by=y})", "  sfx(1)", " end"],
      },
      {
        title: "Visual: give the bullet its own look",
        notes: ["spr()'s flip argument (the 5th parameter) can give the bullet a rotating look as it travels, using the frame counter t() for a changing value each frame - drawn on a distinct sprite index (2) from the player (1) already satisfies most of \"its own look\" on its own:"],
        code: [" for i=#bullets,1,-1 do", "  local b=bullets[i]", "  b.by=b.by-2", "  if b.by<0 then", "   table.remove(bullets,i)", "  else", "   spr(2,b.bx,b.by,0,1,math.floor(t()/4)%4)", "  end", " end"],
      },
      {
        title: "Structure: keep the two loops separate",
        notes: ["The spawn check (btnp(4) + table.insert) and the update/cleanup loop (movement, drawing, sound, visual) are two distinct pieces, shown as two separate blocks above - resist the temptation to fold bullet-spawning logic into the same lines that check player movement, or vice versa, even though they all live inside the same TIC() function by necessity (TIC-80 only has the one entry-point function)."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined into the one TIC() function:"],
        code: [
          "bullets={}",
          "",
          "function TIC()",
          " cls(0)",
          " if btn(3) then x=x+1 end",
          " if btnp(4) then",
          "  table.insert(bullets,{bx=x,by=y})",
          "  sfx(1)",
          " end",
          " for i=#bullets,1,-1 do",
          "  local b=bullets[i]",
          "  b.by=b.by-2",
          "  if b.by<0 then",
          "   table.remove(bullets,i)",
          "  else",
          "   spr(2,b.bx,b.by,0,1,math.floor(t()/4)%4)",
          "  end",
          " end",
          " spr(1,x,y,0)",
          "end",
        ],
      },
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    l1Notes: ["One Input Trigger script, plus two Actor Trigger scripts on the bullet Actor itself:"],
    l1Code: [
      "Input Trigger: Spacebar (Pressed) ->",
      "  Map – Spawn Actor(bullet)",
      "",
      "bullet Actor: Automatic Movement -> Move = true, Direction = Up",
      "",
      "bullet Actor Trigger: Left/Entered Camera (Left) ->",
      "  Actor – Destroy/Revive Actor(this, Destroy)",
    ],
    l2Notes: ["Input Triggers set to \"Pressed\" already only fire once per physical press (not every frame held), so Level 1's fire check needs no changes for Level 2. An optional tunable cooldown, if a slower fire rate is wanted, uses a Timer/Spawn object's own interval instead of scripted variables (the more idiomatic 001 Game Creator approach):"],
    l2Notes2: [],
    l3Ideas: [
      {
        title: "Sound: play on fire",
        notes: ["Import sfx_shoot.wav (Resources > Import > Sound), then add a \"Play Sound\" event directly after Map – Spawn Actor:"],
        code: ["Input Trigger: Spacebar (Pressed) ->", "  Map – Spawn Actor(bullet)", "  Play Sound(sfx_shoot.wav)"],
      },
      {
        title: "Visual: give the bullet its own look",
        notes: ["Set Tint/Opacity and Size directly on the bullet Actor Template's own Graphics Options and General settings - since these are Template-level settings (not scripted per-spawn), every copy Map – Spawn Actor creates automatically looks this way with zero extra events."],
      },
      {
        title: "Structure: keep bullet behavior on the bullet Actor",
        notes: ["main's Input Trigger script should only ever trigger a shot (Map – Spawn Actor + sound) - everything about what a bullet actually does once it exists (its Automatic Movement, its own Left/Entered Camera cleanup trigger, its own Template-level look) belongs on the bullet Actor Template itself. Keeping this boundary clean is what let the Visual idea above need zero scripting at all."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined - the bullet Actor Template's own settings/triggers, plus main's fire script:"],
        code: [
          "bullet Actor Template settings:",
          "  Automatic Movement -> Move = true, Direction = Up",
          "  Graphics Options -> Tint/Opacity set, Size increased",
          "",
          "bullet Actor Trigger: Left/Entered Camera (Left) ->",
          "  Actor – Destroy/Revive Actor(this, Destroy)",
          "",
          "Input Trigger: Spacebar (Pressed) ->",
          "  Map – Spawn Actor(bullet)",
          "  Play Sound(sfx_shoot.wav)",
        ],
      },
    ],
    altNotes: {
      title: "Textual Scripting",
      notes: [
        "Reiterate to students in person: this mode switch is genuinely irreversible per-project. Confirm they created a separate throwaway project before letting them proceed.",
        "Exact text-command names should be checked against the installed version's Script Editor autocomplete or the wiki before class, since this course otherwise treats 001 Game Creator as visual-only and hasn't needed to track Textual syntax changes.",
        "The Level 3 goals map onto Textual commands too - the sound idea is a typed equivalent of Play Sound, and the visual idea needs no scripting either way, since it's a Template-level setting in both modes.",
      ],
      why: [
        "This path exists for students who've realized they prefer typing code (maybe from Godot or GameMaker) and want to see whether 001 Game Creator supports that style too - it does, with this one important caveat.",
        "Because of the irreversibility, this should stay an optional, teacher-supervised exploration rather than a path any student defaults into.",
      ],
    },
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    l1Notes: ["Setup adds one new Boolean Input Action (IA_Fire, mapped to Space Bar inside the existing IMC_Default from Lesson 2 - no new Add Mapping Context call needed, since that context is already applied to the player). BP_Bullet's own Event Graph, plus BP_Player's own fire trigger:"],
    l1Code: [
      "BP_Bullet's Event Graph:",
      "  Event Tick -> Add Actor World Offset(Self, DeltaLocation = (0, Speed*DeltaSeconds, 0))",
      "  Event Tick -> Branch(GetActorLocation().Y > 600) -> True: Destroy Actor(Self)",
      "",
      "BP_Player's own Event Graph:",
      "  Enhanced Input Action Event (IA_Fire), Started -> Spawn Actor from Class(BP_Bullet, Transform = Get Actor Transform(Self))",
    ],
    l2Notes: ["An Enhanced Input Action Event's Started pin already only fires once per physical press (event-driven, not polled), so Level 1's fire check needs no changes for Level 2. An optional tunable cooldown, if a slower fire rate is wanted:"],
    l2Code: [
      "Enhanced Input Action Event (IA_Fire), Started -> Retriggerable Delay (0.2s) -> Spawn Actor from Class(BP_Bullet, Transform = Get Actor Transform(Self))",
    ],
    l3Ideas: [
      {
        title: "Sound: play on fire",
        notes: ["Add a Play Sound 2D node (using an imported sfx_shoot.wav Sound Wave asset) chained after Spawn Actor from Class:"],
        code: ["Enhanced Input Action Event (IA_Fire), Started ->", "  Spawn Actor from Class(BP_Bullet, Transform = Get Actor Transform(Self))", "  Play Sound 2D(sfx_shoot.wav)"],
      },
      {
        title: "Visual: give the bullet its own look",
        notes: ["BP_Bullet's own sprite (bullet_yellow.png, wired up in M_Bullet_Yellow back in Setup) already makes it read as its own thing rather than a copy of the player Cube, so there's no \"give it a distinct color\" gap left to fill the way there is in engines where the bullet still shares the player's look. This idea is a pure size flourish instead: a brief scale pulse in BeginPlay that eases back to normal a moment later."],
        code: ["BP_Bullet's Event Graph:", "  BeginPlay ->", "    Set Actor Relative Scale3D(Self, (1.5, 1.5, 1.5)) ->", "    Delay(0.1s) ->", "    Set Actor Relative Scale3D(Self, (1.0, 1.0, 1.0))"],
      },
      {
        title: "Structure: keep bullet behavior in BP_Bullet",
        notes: ["BP_Player should only ever trigger a shot (Spawn Actor from Class + cooldown + sound) - everything about what a bullet actually does once it exists (movement, off-screen cleanup, its own look) belongs inside BP_Bullet's own Event Graph. Keeping this boundary clean is what let the Sound and Visual ideas above each drop into one Blueprint without touching the other."],
      },
      {
        title: "Complete Level 3 solution (all 3 goals combined)",
        notes: ["Everything above combined, across BP_Bullet and BP_Player:"],
        code: [
          "BP_Bullet's Event Graph:",
          "  BeginPlay ->",
          "    Set Actor Relative Scale3D(Self, (1.5, 1.5, 1.5)) ->",
          "    Delay(0.1s) ->",
          "    Set Actor Relative Scale3D(Self, (1.0, 1.0, 1.0))",
          "  Event Tick -> Add Actor World Offset(Self, DeltaLocation = (0, Speed*DeltaSeconds, 0))",
          "  Event Tick -> Branch(GetActorLocation().Y > 600) -> True: Destroy Actor(Self)",
          "",
          "BP_Player's own Event Graph:",
          "  Enhanced Input Action Event (IA_Fire), Started ->",
          "    Retriggerable Delay (0.2s) ->",
          "      Spawn Actor from Class(BP_Bullet, Transform = Get Actor Transform(Self))",
          "      Play Sound 2D(sfx_shoot.wav)",
        ],
      },
    ],
    altNotes: {
      title: "C++",
      notes: [
        "FireAction is a new UInputAction UPROPERTY bound in the player class's existing SetupPlayerInputComponent(), right alongside the 4 move bindings from Lesson 2 - EnhancedInputComponent->BindAction(FireAction, ETriggerEvent::Started, this, &Class::Fire). No new AddMappingContext() call is needed since IMC_Default is already applied from Lesson 2.",
        "SpawnActor<T>() is a templated function - the template argument (ACppBullet) both creates the instance and gives back a correctly-typed pointer, unlike the Blueprint node's generic Actor return that sometimes needs an explicit cast.",
        "This mirrors the Blueprint path's Spawn Actor from Class node and BP_Bullet's own Event Tick - same underlying engine calls, just from code.",
        "The Level 3 goals map onto C++ calls the same way they map onto Blueprint nodes: UGameplayStatics::PlaySound2D() for sound, and SetActorRelativeScale3D() plus a timer-delayed second call for the scale-pulse visual idea, both called from ACppBullet's own BeginPlay() - keeping the same Structure principle as the Blueprint path.",
      ],
      why: [
        "Showing both C++ and Blueprint side by side for the same behavior is the clearest way to demystify Blueprints for students who assume \"visual\" means \"not real code\" - it's the same engine calls either way.",
        "This is squarely the optional advanced track - most students should stay on Blueprints; C++ is for students already comfortable with typed languages from Godot, GameMaker, or Unity's C# path.",
      ],
    },
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_03_Key_${d.file}.docx`]));

Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log(`done: ${jobs.length} files`));
