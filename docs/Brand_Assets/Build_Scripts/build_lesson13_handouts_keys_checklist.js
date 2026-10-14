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
      headers: { default: brand.pageHeader("Lesson 13 Prep Checklist") },
      footers: { default: brand.pageFooter() },
      children: [
        brand.crestParagraph(1.3, { after: 160 }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 13 Teacher's Checklist / Prep Guide", bold: true, size: 32, color: brand.NAVY })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Complete before class", size: 22, color: brand.NAVY2 })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

        h2("This Week: A Hit Gets a Reaction"),
        p("Since Lesson 4, an enemy hit by a bullet has vanished instantly - functional, but not satisfying. This week adds two visual touches: a brief hit flash on the enemy right before it's removed, and an independent explosion effect that appears at that spot and plays itself out. This is the first Visual Art / animation-focused lesson of the course."),
        p("This lesson genuinely restructures each engine's hit-handling code, more than most weeks - read the Handout for your demo engine (Godot) closely before class, and skim the others so you can recognize each engine's shape of the same idea when circulating."),
        checkbox("Read the Teacher's Script's timing-mechanism comparison before class - three different engine families solve 'pause before finishing the kill' three different ways this week, and it's worth understanding all three before fielding student questions."),

        h2("Self-Contained Lesson Assets"),
        p("Two new placeholder assets are provided in this lesson's Assets folder: explosion_placeholder.png (a simple burst/star shape) and enemy_invader_white.png (the same silhouette as Lesson 4's enemy_invader_red.png, recolored solid white - this is the flash sprite for Godot, Unity, GameMaker, and Unreal, swapped in briefly on hit instead of tinted, for the same reason this course's sprite-swap Recolor from Lesson 2 exists: a multiply-based color tint can only darken/mix with a sprite's existing color, so tinting an already-colored enemy toward white produces no visible change at all - swapping to a genuinely white sprite is the only way to get an actual white flash). Unreal's own BP_Enemy Material hits this exact trap too, since Lesson 4 built its Color parameter as a multiply against the enemy's real texture (the same technique Godot's modulate and Unity's SpriteRenderer.color use) - Set Vector Parameter Value(\"Color\", White) would be just as invisible there as it is in Godot or Unity, which is why Unreal's own flash swaps in M_Enemy_White (built from this sprite) with a Set Material node instead. TIC-80 students draw their own white flash sprite directly in the Sprite Editor, and Scratch uses its own color-cycling Graphic Effect (not a multiply tint, so it isn't affected either); GDevelop's Level 1 uses the Flash extension's blink action, which just toggles visibility and works regardless of sprite color. 001 Game Creator's built-in Actor - Flash action is a genuine engine feature, not a manual tint, so it's unaffected too."),
        checkbox("Confirm explosion_placeholder.png and enemy_invader_white.png are both present and import cleanly on the demo laptop before class."),
        checkbox("Godot / Unity / GameMaker specifically: confirm the enemy visibly turns white (not just brighter or unchanged) for the flash - if it doesn't, the flash sprite likely wasn't imported/assigned, and the code silently fell back to a no-op tint."),

        h2("Continuity From Lesson 12"),
        checkbox("Remind students to open their own [YourName]Game project from Lesson 12, not create a new one."),
        checkbox("If a student's Lesson 12 project is lost or broken, have a pre-built \"Lesson 12 complete\" starter available per engine (from the Lesson 12 Teacher's Keys) so they aren't blocked."),

        h2("Engine Verification"),
        checkbox("Confirm the demo laptop's Lesson 11-12 exercises (multiple waves, all four sounds) still work for all 8 engines."),
        checkbox("Test a full round on every engine before class: a hit should flash the enemy briefly, an explosion should appear and clean itself up, and score/lives/wave/win-lose logic should all still behave exactly as before - none of today's changes should alter WHEN score or the win-check actually resolve, only how the enemy's removal looks."),
        checkbox("GameMaker and TIC-80 specifically: confirm the enemy's actual destruction now happens slightly after the hit is registered (via an Alarm or countdown), not in the same instant - this is intentional this week, not a regression."),
        checkbox("001 Game Creator (Windows demo laptop): confirm the Actor – Flash action is available in your installed version and produces a visible flash; if the enemy is destroyed before the flash completes, that's a known, acceptable limitation this week (see the Handout's troubleshooting note) - the explosion effect is this engine's primary required visual."),

        h2("Materials"),
        checkbox("Print or queue digital copies of all 8 Student Handouts (or confirm students can access their engine's version electronically)."),
        checkbox("Have the 8 Teacher's Keys on hand (instructor reference only - do not distribute) in case a student needs a worked example beyond Level 1."),
        checkbox("Bring: demo laptop, projector cable/adapter, handout copies."),

        h2("Known Gotchas This Week"),
        bullet("The most likely mistake: delaying or moving the score/win-check logic itself, rather than only delaying the enemy's final visual removal - the two must stay separate. Watch for this specifically since it's an easy line to blur."),
        bullet("A close second: forgetting to reset a flashed sprite/texture back to normal (relevant mainly for any engine where an object could theoretically be reused, e.g. an object pool) - low risk this week since every enemy that flashes is destroyed a fraction of a second later, but worth a mention if a student asks."),
        bullet("This is a genuinely fun week for students - encourage them to actually watch their explosions rather than just confirming they exist, and to compare notes with classmates on other engines during the share-out."),
        bullet("If a cluster finishes Level 1 quickly, point them at Level 3's animation-timing ideas rather than rushing ahead to Lesson 14's material."),
      ]
    }]
  });
  jobs.push([doc, "Lesson_13_Prep_Checklist.docx"]);
}

// =========================================================================
// STUDENT HANDOUTS
// =========================================================================
function buildHandout(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 13: A Hit That Reacts", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine, size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("This week: getting hit finally LOOKS like getting hit. A brief flash, then a small explosion, before the enemy disappears.", { bold: true }),
    p("Open your own [YourName]Game project from Lesson 12 - don't start a new one.", { bold: true }),
    p("Important: your score, lives, wave, and win/lose logic should all still work exactly as before. Today only changes how the enemy's removal LOOKS, not when your game-state actually changes.", { italics: true, color: "555555" }),
    h2("Setup"),
    ...data.setup.map(s => numbered(s)),
    h2("Level 1 - Minimal (MVP)"),
    p("On every hit, the enemy briefly flashes, and a small independent explosion effect appears and cleans itself up. Score, lives, wave, and win/lose logic are all unaffected.", { italics: true }),
    ...data.level1.map(s => numbered(s)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  if (data.altPath) {
    children.push(h2("Alternate Path: " + data.altPath.title));
    children.push(p(data.altPath.intro, { italics: true, color: "555555" }));
    children.push(...data.altPath.steps.map(s => numbered(s)));
    if (data.altPath.code) children.push(code(data.altPath.code));
    if (data.altPath.note) children.push(p(data.altPath.note, { italics: true, color: "555555" }));
  }
  children.push(
    h2("Level 2 - Complete (Alpha/Beta)"),
    p("The flash and explosion play correctly on every single hit, with no leftover explosions, no enemies stuck mid-flash, and no change to score/lives/wave timing.", { italics: true }),
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
      headers: { default: brand.pageHeader("Lesson 13 Handout - " + data.engine) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const handoutData = [
  {
    engine: "Scratch", file: "Scratch",
    setup: [
      "Open your [YourName]Game project from Lesson 12.",
      "Add the explosion_placeholder.png costume to a brand-new sprite named Explosion (Scratch will let you draw/import a new sprite from a file).",
      "Add \"set [color] effect to 0\" to the Enemy's existing \"when I receive [spawn wave]\" reset script (or wherever each enemy resets), so a fresh enemy never starts already flashed.",
    ],
    level1: [
      "On the Explosion sprite, hide it by default, and add a new script: \"when I receive [explosion]\" -> show -> go to x/y of Enemy (using the enemy's position at the moment of the broadcast) -> wait 0.2 seconds -> hide.",
      "In the Enemy's touching-Bullet branch, right after \"change enemies_left by -1\" and the existing sound/score lines, add: broadcast [explosion] (this spawns the visual independently, with zero effect on anything else), then \"set [color] effect to 100\", then \"wait 0.1 seconds\".",
      "Move the existing win-check if/else and \"delete this clone\" to run AFTER the wait - everything that was already there stays exactly the same, just one step later.",
      "Click the stage and play through a full round - each hit should flash the enemy briefly, and a small burst should appear at that spot for a moment.",
    ],
    level1Code: [
      "Enemy's touching-Bullet branch (revised):",
      "if <touching Bullet?> then",
      "  change score by 10",
      "  change enemies_left by -1",
      "  start sound [sfx_hit]",
      "  broadcast [explosion]",
      "  set [color] effect to 100",
      "  wait 0.1 seconds",
      "  if <enemies_left = 0> then",
      "    if <wave = total_waves> then",
      "      set game_active to 0",
      "      start sound [sfx_win]",
      "      broadcast [you win]",
      "    else",
      "      change wave by 1",
      "      set enemies_left to 5",
      "      start sound [sfx_wave]",
      "      broadcast [spawn wave]",
      "  delete this clone",
    ],
    level2: [
      "Play a full round: every single hit shows a flash and an explosion, with no enemy getting visually \"stuck\" flashed.",
      "Confirm score, wave-clear, and win/lose all still happen at exactly the moments they did last week - only the enemy's disappearance is 0.1 seconds later than before.",
      "Confirm the broadcast-based Explosion sprite never blocks or delays anything else - broadcasts don't wait for their receivers to finish.",
    ],
    level3: [
      "Audio/Visual: instead of one flat 0.1-second flash, alternate the color effect a couple of times (e.g. 100, then 0, then 100 again) with short waits between, for a stronger flicker.",
      "Visual: give the Explosion sprite 2-3 different costumes and cycle through them with \"next costume\" during its 0.2-second lifetime for a simple hand-animated burst instead of one static image.",
      "Structure: notice the Explosion sprite knows nothing about score, lives, or waves - it does exactly one job, triggered independently of the rest of the hit-handling logic.",
    ],
    trouble: [
      "Enemy vanishes instantly, no flash visible: confirm \"wait 0.1 seconds\" was actually added, and that the win-check/delete block was moved to AFTER it, not before.",
      "Explosion never appears: confirm the Explosion sprite actually has a \"when I receive [explosion]\" script, and that its costume isn't accidentally still hidden.",
      "Score changes twice per hit, or win-check fires early: confirm the win-check/scoring lines weren't accidentally duplicated when moving them below the wait.",
    ],
    resource: "Student Resource List > Scratch section, or scratch.mit.edu/ideas",
  },
  {
    engine: "Godot Engine", file: "Godot",
    setup: [
      "Open your [YourName]Game project from Lesson 12.",
      "Create a new small scene, explosion.tscn: a Sprite2D (Texture = explosion_placeholder.png) with a Timer child node (Wait Time = 0.2, One Shot on).",
      "On the explosion scene's own script, connect the Timer's timeout signal to a function that calls queue_free() - the explosion cleans itself up, with no help needed from the Enemy or main scene.",
      "Import enemy_invader_white.png (this lesson's Assets folder) into your project - it's the same enemy silhouette as Lesson 4's sprite, pre-recolored solid white, used below for the hit flash.",
    ],
    level1: [
      "On the Enemy scene's script, find the function that currently detects a bullet hit (from Lesson 4) - it likely calls queue_free() and emits a \"destroyed\" signal immediately.",
      "Right after detecting the hit, instantiate explosion.tscn at global_position and add it to the scene (code below) - this happens immediately and independently, with zero effect on anything else.",
      "Swap the Enemy's texture to enemy_invader_white.png for the flash (code below), then await get_tree().create_timer(0.1).timeout to pause this function for a tenth of a second. A color tint (modulate) won't work here: modulate multiplies the sprite's existing color, and multiplying any color by white leaves it completely unchanged - a genuinely white sprite is the only way to get a visible white flash on colored source art, the same reason Lesson 2's Recolor swaps sprites instead of tinting them.",
      "AFTER the await, keep the existing destroyed.emit() and queue_free() calls exactly as they were - just one step later in the same function.",
      "Run the scene and play through a full round - each hit should flash the enemy white briefly, and a small burst should appear at that spot for a moment.",
    ],
    level1Code: [
      "# Enemy's hit-detection function (revised):",
      "const FLASH_TEXTURE = preload(\"res://enemy_invader_white.png\")",
      "",
      "func _on_area_entered(area):",
      "    if area.is_in_group(\"bullets\"):",
      "        area.queue_free()",
      "        var explosion = explosion_scene.instantiate()",
      "        explosion.global_position = global_position",
      "        get_tree().current_scene.add_child(explosion)",
      "        texture = FLASH_TEXTURE",
      "        await get_tree().create_timer(0.1).timeout",
      "        destroyed.emit()",
      "        queue_free()",
    ],
    level2: [
      "Play a full round: every single hit shows a flash and an explosion, with no enemy left stuck mid-flash.",
      "Confirm score, wave-clear, and win/lose all still happen at exactly the moments they did last week - await only delays this ONE function's remaining lines, nothing else in the game pauses.",
      "Confirm explosion.tscn instances never pile up - each one removes itself via its own Timer, with no help from the Enemy.",
    ],
    level3: [
      "Visual: instead of one flat instant swap, alternate texture between FLASH_TEXTURE and the Enemy's normal texture 2-3 times across the 0.1 seconds (with tiny awaits between) for a flicker instead of a single snap.",
      "Visual: scale the explosion up slightly over its 0.2-second lifetime using a Tween, so it feels like it's expanding rather than just appearing and vanishing.",
      "Structure: notice explosion.tscn knows nothing about score, lives, or waves - it's a fully self-contained scene, the same discipline as this course's bullets and enemies.",
    ],
    trouble: [
      "Enemy vanishes instantly, no flash visible: confirm await get_tree().create_timer(0.1).timeout was actually added, and that destroyed.emit()/queue_free() come AFTER it.",
      "Enemy disappears but never visibly turns white: confirm enemy_invader_white.png was actually imported and FLASH_TEXTURE's preload path matches where it landed in the project - a missing/misnamed path fails silently and the sprite just keeps its normal texture.",
      "Explosion never appears, or an error about a missing scene: confirm explosion_scene is preloaded/exported and assigned to the actual explosion.tscn file.",
      "Two hits on the same enemy cause an error: confirm area.queue_free() (the bullet) happens before the await, so a second bullet can't re-trigger this same function on an enemy already mid-flash.",
    ],
    resource: "Student Resource List > Godot section, or docs.godotengine.org",
  },
  {
    engine: "Unity", file: "Unity",
    setup: [
      "Open your [YourName]Game project from Lesson 12.",
      "Create a small Explosion prefab: a GameObject with a SpriteRenderer (Sprite = explosion_placeholder.png) and a script that calls Destroy(gameObject, 0.3f) in Start() - it cleans itself up automatically, no help needed from anything else.",
      "Import enemy_invader_white.png (this lesson's Assets folder) into your project - it's the same enemy silhouette as Lesson 4's sprite, pre-recolored solid white, used below for the hit flash.",
    ],
    level1: [
      "On the Enemy prefab's script, find the function that currently detects a bullet hit (from Lesson 4) - it likely calls GameManager.RegisterKill() and Destroy(gameObject) immediately.",
      "Add public GameObject explosionPrefab, public SpriteRenderer spriteRenderer, and public Sprite flashSprite fields, and assign them in the Inspector (flashSprite = enemy_invader_white.png).",
      "Rewrite the hit-handling as a Coroutine (code below) - Unity's first Coroutine this course - so it can pause partway through without freezing the whole game.",
      "Inside the Coroutine: Instantiate the explosion immediately (independent of everything else), swap spriteRenderer.sprite to flashSprite, then yield return new WaitForSeconds(0.1f) before the existing RegisterKill()/Destroy() calls. A color tint (spriteRenderer.color) won't work here: it multiplies the sprite's existing color, and multiplying any color by white leaves it completely unchanged - a genuinely white sprite is the only way to get a visible white flash on colored source art, the same reason Lesson 2's Recolor swaps sprites instead of tinting them.",
      "Change the hit-detection code to call StartCoroutine(HandleHit()) instead of calling the old logic directly.",
      "Press Play and play through a full round - each hit should flash the enemy white briefly, and a small burst should appear at that spot for a moment.",
    ],
    level1Code: [
      "// Enemy's script (revised):",
      "public GameObject explosionPrefab;",
      "public SpriteRenderer spriteRenderer;",
      "public Sprite flashSprite;",
      "",
      "void OnTriggerEnter2D(Collider2D other) {",
      "    if (other.CompareTag(\"Bullet\")) {",
      "        Destroy(other.gameObject);",
      "        StartCoroutine(HandleHit());",
      "    }",
      "}",
      "",
      "IEnumerator HandleHit() {",
      "    Instantiate(explosionPrefab, transform.position, Quaternion.identity);",
      "    spriteRenderer.sprite = flashSprite;",
      "    yield return new WaitForSeconds(0.1f);",
      "    GameManager.RegisterKill();",
      "    Destroy(gameObject);",
      "}",
    ],
    level2: [
      "Play a full round: every single hit shows a flash and an explosion, with no enemy left stuck mid-flash.",
      "Confirm score, wave-clear, and win/lose all still happen at exactly the moments they did last week - yield only pauses this ONE Coroutine, nothing else in the game pauses.",
      "Confirm explosion prefabs never pile up in the Hierarchy - each one destroys itself via Destroy(gameObject, 0.3f), with no help from the Enemy.",
    ],
    level3: [
      "Visual: instead of one flat instant swap, alternate spriteRenderer.sprite between flashSprite and the enemy's normal sprite 2-3 times across the 0.1 seconds (with tiny waits between) for a flicker instead of a single snap.",
      "Visual: scale the explosion prefab up slightly over its lifetime with a small script using Time.deltaTime, so it feels like it's expanding.",
      "Structure: notice the explosion prefab knows nothing about score, lives, or waves - it's a fully self-contained object, the same discipline as this course's bullets and enemies.",
    ],
    trouble: [
      "Enemy vanishes instantly, no flash visible: confirm the Coroutine is actually being started (StartCoroutine(HandleHit())), not the old code still calling RegisterKill()/Destroy() directly.",
      "Enemy disappears but never visibly turns white: confirm flashSprite is actually assigned in the Inspector to enemy_invader_white.png, not left empty or still pointing at the normal sprite.",
      "Explosion never appears: confirm explosionPrefab is actually assigned in the Inspector, not left empty.",
      "Two bullets can double-trigger the same enemy: confirm other.gameObject is destroyed at the top of OnTriggerEnter2D, before StartCoroutine runs, so a second overlapping bullet can't also collide.",
    ],
    resource: "Student Resource List > Unity section, or learn.unity.com",
    altPath: {
      title: "Visual Scripting",
      intro: "The same flash-then-explode idea works in a Script Graph, using a Wait For Seconds node in place of a C# Coroutine.",
      steps: [
        "On the enemy's hit-handling graph, add an Instantiate node for the explosion prefab right at the start (unconditional, runs immediately).",
        "Set the SpriteRenderer's Sprite to flashSprite (enemy_invader_white.png) - not its Color, which would multiply against the enemy's existing color and produce no visible change - then add a Wait For Seconds (0.1) node before your existing RegisterKill/Destroy nodes.",
      ],
    },
  },
  {
    engine: "GameMaker", file: "GameMaker",
    setup: [
      "Open your [YourName]Game project from Lesson 12.",
      "Create a new object, obj_explosion, with a Sprite (import explosion_placeholder.png). In its Create event, set alarm[0] = 12. In its new Alarm 0 event, call instance_destroy() - it cleans itself up automatically.",
      "Import enemy_invader_white.png (this lesson's Assets folder) as a new sprite, spr_enemy_flash - it's the same enemy silhouette as obj_enemy's existing sprite, pre-recolored solid white, used below for the hit flash.",
    ],
    level1: [
      "GameMaker code can't pause mid-event, so this week's delay works differently: obj_bullet's Collision event starts the flash and sets an alarm on the enemy; a NEW Alarm 0 event on obj_enemy actually finishes removing it.",
      "On obj_bullet's Collision event with obj_enemy, right after the existing sound/score/win-check lines, add: instance_create_layer(other.x, other.y, \"Instances\", obj_explosion) (spawns the explosion immediately, independent of everything else), other.sprite_index = spr_enemy_flash (swaps to the flash sprite - image_blend won't work here, since it multiplies the sprite's existing color and multiplying any color by white leaves it completely unchanged, the same reason Lesson 2's Recolor swaps sprites instead of tinting them), and other.alarm[0] = 6.",
      "Remove instance_destroy(other) from obj_bullet's Collision event - the enemy is no longer destroyed here.",
      "Add a new Alarm 0 event to obj_enemy containing exactly one line: instance_destroy() - this is where the enemy actually disappears now, a fraction of a second after being hit.",
      "Press F5 (Run) and play through a full round - each hit should flash the enemy briefly, and a small burst should appear at that spot for a moment.",
    ],
    level1Code: [
      "// obj_bullet Collision event with obj_enemy (revised):",
      "audio_play_sound(sfx_hit, 1, false);",
      "instance_create_layer(other.x, other.y, \"Instances\", obj_explosion);",
      "other.sprite_index = spr_enemy_flash;",
      "other.alarm[0] = 6;",
      "obj_spawner.enemies_remaining -= 1;",
      "obj_spawner.score += 10;",
      "if (obj_spawner.enemies_remaining <= 0) {",
      "    if (obj_spawner.wave >= obj_spawner.total_waves) {",
      "        obj_spawner.game_active = false;",
      "        obj_spawner.has_won = true;",
      "        audio_play_sound(sfx_win, 1, false);",
      "    } else {",
      "        obj_spawner.wave += 1;",
      "        obj_spawner.enemies_remaining = 5;",
      "        audio_play_sound(sfx_wave, 1, false);",
      "        obj_spawner.spawn_wave();",
      "    }",
      "}",
      "instance_destroy(self);",
      "",
      "// obj_enemy's NEW Alarm 0 event:",
      "instance_destroy();",
    ],
    level2: [
      "Play a full round: every single hit shows a flash and an explosion, with no enemy left stuck mid-flash.",
      "Confirm score, wave-clear, and win/lose all still happen at exactly the moments they did last week - only obj_enemy's actual removal is 6 steps later than before.",
      "Confirm obj_explosion instances never pile up - each one destroys itself via its own Alarm 0, with no help from obj_enemy.",
    ],
    level3: [
      "Visual: instead of one flat instant swap, toggle sprite_index between spr_enemy_flash and the enemy's normal sprite every couple of steps across the 6-step alarm window, for a flicker instead of a single snap.",
      "Visual: grow obj_explosion's image_xscale/image_yscale from 0.5 to 1.5 across its Alarm 0 lifetime for an expanding-burst look.",
      "Structure: notice obj_explosion knows nothing about score, lives, or waves - it's a fully self-contained object, the same discipline as this course's bullets and enemies.",
    ],
    trouble: [
      "Enemy vanishes instantly, no flash visible: confirm instance_destroy(other) was actually removed from obj_bullet's Collision event, and that other.alarm[0] = 6 is set instead.",
      "Enemy disappears but never visibly turns white: confirm spr_enemy_flash was actually imported from enemy_invader_white.png and that other.sprite_index = spr_enemy_flash (not image_blend) is what's being set.",
      "Enemy never disappears at all: confirm obj_enemy actually has a new Alarm 0 event with instance_destroy() in it - setting an alarm does nothing without a matching Alarm event.",
      "\"You Win!\" or wave-clear timing feels off: it shouldn't - confirm score/enemies_remaining/win-check are still all inside obj_bullet's Collision event, unchanged, not accidentally moved into the new Alarm 0 event.",
    ],
    resource: "Student Resource List > GameMaker section, or manual.gamemaker.io",
    altPath: {
      title: "GML Visual (no typing)",
      intro: "The same collision-then-alarm pattern works as drag-and-drop blocks instead of typed code.",
      steps: [
        "On obj_bullet's Collision event (GML Visual), add a \"Create Instance\" block for obj_explosion, a \"Set Instance Variable\" block for other.sprite_index (targeting spr_enemy_flash, not other.image_blend - a color tint can't produce a true white flash on a colored sprite), and a \"Set Alarm\" block for other.alarm[0], right alongside your existing score/sound blocks.",
        "Remove the \"Destroy Instance\" block that targeted the enemy, and add a new Alarm 0 event on obj_enemy with a single \"Destroy Instance\" block (targeting self).",
      ],
    },
  },
  {
    engine: "GDevelop", file: "GDevelop",
    setup: [
      "Open your [YourName]Game project from Lesson 12.",
      "Add a new Sprite object named Explosion, using explosion_placeholder.png as its image.",
      "In the Extensions panel, install the official \"Flash\" extension if it isn't already installed (one click) - this makes the Flash behavior available to add to objects, but doesn't attach it to anything yet.",
      "Select Enemy, open its Behaviors tab, click Add a behavior, search \"Flash,\" and Apply. Flash visibility (blink) is an action that belongs to this behavior, not a general action - it won't show up in Enemy's action list until this step is done, even though the extension itself is already installed.",
    ],
    level1: [
      "On the bullet-enemy collision event, right after the existing score/sound actions, add: Create object Explosion at Enemy.X(), Enemy.Y() (spawns immediately, independent of everything else) and Flash visibility (blink) of Enemy, Duration 0.3 seconds - the Flash behavior's blink action (now available since the behavior was added to Enemy in Setup), which toggles the Enemy's visibility on/off for that duration. This is deliberately the blink action, not the Flash behavior's separate color-tint action: a color tint multiplies against the Enemy's existing color, and multiplying any color by white leaves it completely unchanged - it wouldn't actually flash on a colored sprite like ours. Blink works regardless of sprite color, since it's just toggling visibility.",
      "Add a Wait 0.1 seconds action right after the flash action, still inside the same event.",
      "Move Delete Enemy to come AFTER the Wait action - everything else in the event (Score, EnemiesRemaining, Play a sound, Delete Bullet) stays exactly where it was, unchanged.",
      "Add a separate event: Condition: Explosion has existed for more than 0.3 seconds -> Action: Delete Explosion, so explosions clean themselves up automatically.",
      "Click Preview and play through a full round - each hit should flash the enemy briefly, and a small burst should appear at that spot for a moment.",
    ],
    level1Code: [
      "Bullet-Enemy collision event (revised):",
      "Condition: Bullet is in collision with Enemy",
      "Action: Score = Score + 10",
      "Action: EnemiesRemaining = EnemiesRemaining - 1",
      "Action: Play a sound (sfx_hit.wav)",
      "Action: Create object Explosion at Enemy.X(), Enemy.Y()",
      "Action: Flash visibility (blink) of Enemy, Duration 0.3",
      "Action: Delete Bullet",
      "Action: Wait 0.1 seconds",
      "Action: Delete Enemy",
      "",
      "Separate event (explosion self-cleanup):",
      "Condition: Explosion has existed for more than 0.3 seconds",
      "Action: Delete Explosion",
    ],
    level2: [
      "Play a full round: every single hit shows a flash and an explosion, with no enemy left stuck mid-flash.",
      "Confirm score, wave-clear, and win/lose all still happen at exactly the moments they did last week - EnemiesRemaining changes immediately, only Delete Enemy is delayed.",
      "Confirm Explosion objects never pile up - each one deletes itself via its own timed event, with no help from the collision event.",
    ],
    level3: [
      "Visual: add the official Tween behavior to Explosion, and tween its scale from 0.5 to 1.5 over 0.3 seconds for an expanding-burst look instead of a static image.",
      "Visual: try the Flash extension's separate color-tint-flash action instead of the blink action, and see what happens on your Enemy's actual sprite color - a multiply-based tint can only darken/mix with a sprite's existing color, so a \"white\" tint on a colored sprite comes out looking like the original color, not white. It's still worth trying (and a reasonable stretch goal), but pick a color noticeably different from your Enemy's own for a result you can actually see, rather than white specifically.",
      "Structure: notice Explosion knows nothing about score, lives, or waves - it's a fully self-contained object, the same discipline as this course's bullets and enemies.",
    ],
    trouble: [
      "Enemy vanishes instantly, no flash visible: confirm Wait 0.1 seconds was actually added, and Delete Enemy comes AFTER it, not before.",
      "Explosion never appears: confirm the Flash extension is actually installed, and that Create object Explosion references the correct object name.",
      "Explosions pile up and never disappear: confirm the separate self-cleanup event's \"has existed for more than\" condition is actually targeting the Explosion object, not a typo'd name.",
      "Enemy doesn't flash at all, even though the extension is installed: installing the Flash extension only makes the behavior available to add - it doesn't attach it to any object automatically. Confirm the Flash behavior was actually added to Enemy specifically (Enemy > Behaviors tab).",
    ],
    resource: "Student Resource List > GDevelop section, or wiki.gdevelop.io",
  },
  {
    engine: "TIC-80", file: "TIC-80",
    setup: [
      "Open your [YourName]Game project from Lesson 12.",
      "Design a small explosion sprite in the Sprite Editor (a burst/star shape works well) in an empty slot, and a second \"flash\" version of your enemy sprite (the same silhouette, recolored bright white) in another empty slot - note both slot numbers.",
    ],
    level1: [
      "TIC-80's code can't pause mid-frame, so this week's delay works differently: instead of removing a hit enemy from your enemies table immediately, give it a flash countdown field and remove it once that counter reaches 0.",
      "In the bullets loop's hit branch, right after the existing sound/score/win-check lines, add e.flash = 6 to the enemy that was hit (code below) - do NOT remove it from the table yet.",
      "Add a small explosions table, and table.insert an entry (with its own short countdown) at the hit position - this is independent of the enemy's own flash countdown.",
      "Add a new update loop (runs every frame, before drawing): count down every enemy's flash field and every explosion's countdown, removing each from its table once it reaches 0.",
      "In your drawing code, draw the flash sprite instead of the normal enemy sprite whenever e.flash is set, and draw each active explosion's sprite at its position.",
      "Press Esc, type run, press Enter, then play through a full round - each hit should flash the enemy briefly, and a small burst should appear at that spot for a moment.",
    ],
    level1Code: [
      "-- bullets loop's hit branch (revised):",
      "score = score + 10",
      "enemies_remaining = enemies_remaining - 1",
      "sfx(2)",
      "e.flash = 6",
      "table.insert(explosions, {x=e.x, y=e.y, t=12})",
      "if enemies_remaining <= 0 then",
      " if wave >= total_waves then",
      "  game_active = false",
      "  has_won = true",
      "  sfx(3)",
      " else",
      "  wave = wave + 1",
      "  enemies_remaining = 5",
      "  sfx(4)",
      "  spawn_wave()",
      " end",
      "end",
      "",
      "-- new update loop (each frame, before drawing):",
      "for i=#enemies,1,-1 do",
      " local e = enemies[i]",
      " if e.flash then",
      "  e.flash = e.flash - 1",
      "  if e.flash <= 0 then table.remove(enemies, i) end",
      " end",
      "end",
      "for i=#explosions,1,-1 do",
      " explosions[i].t = explosions[i].t - 1",
      " if explosions[i].t <= 0 then table.remove(explosions, i) end",
      "end",
      "",
      "-- drawing (per enemy):",
      "if e.flash then spr(FLASH_SPRITE, e.x, e.y, 14) else spr(ENEMY_SPRITE, e.x, e.y, 14) end",
    ],
    level2: [
      "Play a full round: every single hit shows a flash and an explosion, with no enemy left stuck mid-flash.",
      "Confirm score, wave-clear, and win/lose all still happen at exactly the moments they did last week - only each enemy's actual table.remove() is delayed, not the scoring logic.",
      "Confirm the explosions table never grows without bound - each entry's countdown reaches 0 and gets removed on its own.",
    ],
    level3: [
      "Visual: cycle the explosion through 2-3 different sprite frames across its 12-frame lifetime instead of one static sprite, for a simple hand-animated burst.",
      "Visual: instead of a flat flash sprite, alternate between the flash sprite and the normal sprite every couple of frames while e.flash is active, for a flicker instead of a solid flash.",
      "Structure: notice the explosions table knows nothing about score, lives, or waves - it's a fully self-contained piece of state, the same discipline as this course's bullets table.",
    ],
    trouble: [
      "Enemy vanishes instantly, no flash visible: confirm e.flash = 6 was actually added, and that the table.remove() for a hit enemy no longer happens in the hit branch itself.",
      "Enemy freezes on screen forever, never removed: confirm the new update loop is actually being called every frame from TIC(), and that it counts down and removes correctly.",
      "Explosion sprite never appears: confirm FLASH_SPRITE/explosion sprite slot numbers in your drawing code match what you actually drew in the Sprite Editor.",
    ],
    resource: "Student Resource List > TIC-80 section, or github.com/nesbox/TIC-80/wiki",
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    setup: [
      "Open your [YourName]Game project from Lesson 12.",
      "Create a new Actor Template named Explosion, using explosion_placeholder.png as its Graphics.",
    ],
    level1: [
      "On the enemy Actor's \"Collided with actor\" Trigger, right after the existing sound/score/win-check actions, add Actor – Flash (a real, built-in action that flashes a color on the Actor for a set duration - set a short duration, e.g. 0.2 seconds).",
      "In the same action list, add Spawn Actor(Explosion) at the enemy's current position - this happens immediately and is fully independent of everything else.",
      "Leave the enemy's existing Actor – Destroy/Revive(Destroy) action exactly where it already is, unchanged - in this engine, the flash and the destroy happen close together rather than with a manual delay between them (see the note below).",
      "If your version supports a Lifetime/Duration property when spawning an Actor, set the Explosion's to a short value (e.g. 0.3 seconds) so it removes itself automatically.",
      "Save, then Play/Test and play through a full round - each hit should show a small burst at that spot; the enemy's own flash may or may not be fully visible depending on your version (see note below).",
    ],
    level1Code: [
      "enemy Actor Trigger: Collided with actor (revised) ->",
      "  If(game_active = 1)",
      "    Variable – Modify Value(score, Add, 10)",
      "    Variable – Modify Value(enemies_remaining, Subtract, 1)",
      "    Play Sound(sfx_hit.wav)",
      "    Actor – Flash(White, Duration: 0.2)",
      "    Spawn Actor(Explosion) at (this actor's position)",
      "    If(enemies_remaining <= 0)",
      "      If(wave >= total_waves)",
      "        Variable – Modify Value(game_active, Set, 0)",
      "        Play Sound(sfx_win.wav)",
      "        Message Box(\"You Win! - Press R to Restart\")",
      "      Else",
      "        Variable – Modify Value(wave, Add, 1)",
      "        Variable – Modify Value(enemies_remaining, Set, 5)",
      "        Play Sound(sfx_wave.wav)",
      "        Call Custom Event(SpawnWave)",
      "    Actor – Destroy/Revive(Destroy)",
    ],
    level2: [
      "Play a full round: an explosion appears at every hit and cleans itself up; score, lives, wave, and win/lose all still happen at exactly the moments they did last week.",
      "Confirm Explosion Actors never pile up on screen across a long playthrough.",
      "This engine's flash may complete quickly or be cut short depending on your version - Level 2 here means the explosion effect is consistently visible, even if the enemy's own flash is brief.",
    ],
    level3: [
      "Audio/Visual: if your version supports it, explore a Timer or delay mechanism to hold the enemy's Destroy action for a fraction of a second after Flash starts, so the flash is fully visible before the Actor disappears.",
      "Visual: give the Explosion Actor Template a second, brighter Graphics variant and swap to it partway through its lifetime for a simple two-frame animation.",
      "Structure: notice Explosion knows nothing about score, lives, or waves - it's a fully self-contained Actor, the same discipline as this course's bullets and enemies.",
    ],
    trouble: [
      "No flash visible at all: this is a known limitation in some versions when Destroy fires in the same action list right after Flash - the explosion effect is this engine's primary required visual this week; a fully-visible enemy flash is a Level 3 stretch, not a Level 1 requirement.",
      "Explosion never appears: confirm the Explosion Actor Template was actually created with a Graphics asset assigned, and that Spawn Actor references it by the correct name.",
      "Explosions never disappear: confirm a Lifetime/Duration was set when spawning, or add your own cleanup Trigger if your version requires one.",
    ],
    resource: "001gamecreator.com/wiki/, Student Resource List > 001 Game Creator",
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    setup: [
      "Open your [YourName]Game project from Lesson 12.",
      "Import enemy_invader_white.png (this lesson's Assets folder) into the Content Browser, then build M_Enemy_White the same way Lesson 4 built M_Enemy_Red - a new Material, Shading Model = Unlit, a Texture Sample using enemy_invader_white.png wired straight into Emissive Color. No Color parameter needed this time - unlike M_Enemy_Red, this Material exists purely to be swapped in as a whole for the flash, the same way M_Player_Red and friends were back in Lesson 2's Recolor.",
      "Import explosion_placeholder.png into the Content Browser, then build M_Explosion: a new Material, Shading Model = Unlit, and Blend Mode = Masked (Details panel, right below Shading Model) - Masked lets parts of a texture render fully see-through instead of showing as a solid rectangle, which a plain Opaque Unlit material can't do. Add a Texture Sample using explosion_placeholder.png, wire its RGB output into Emissive Color and its Alpha output into Opacity Mask.",
      "Create a new Actor Blueprint, BP_Explosion, using a Plane Static Mesh Component (Place Actors panel's Basic Shapes, same built-in set as Cube and Sphere - no download needed) with M_Explosion assigned to it. A Plane is a flat quad, the natural fit for showing a 2D sprite texture face-on to the top-down Orthographic Camera from Lesson 2, unlike a Cube or Sphere which would show the image wrapped oddly around a 3D volume.",
      "In BP_Explosion's Event Graph: BeginPlay -> a Timeline (0.3 seconds) driving the Actor's Scale from small to large -> On Timeline Finished -> Destroy Actor - it cleans itself up automatically.",
    ],
    level1: [
      "In HandleEnemyDestroyed (from Lesson 11/12), right after the existing sound/score updates, add Spawn Actor from Class(BP_Explosion) at the enemy's location - this happens immediately and independently of everything else.",
      "Swap the enemy's own Material to M_Enemy_White with a Set Material node (Material Index 0) - the same real sprite-swap flash every other engine uses this week, just reached through Unreal's Material system - then add a Delay(0.1) node before the rest of the graph continues. There's no need to set it back afterward: the enemy is destroyed a moment later anyway.",
      "Everything after the Delay (EnemiesRemaining update, the win-check nested Branch, Destroy Actor on the enemy) stays exactly as it was - just running a tenth of a second later in the same graph.",
      "Play in Editor and play through a full round - each hit should flash the enemy briefly, and a small shape should scale up and disappear at that spot.",
    ],
    level1Code: [
      "Custom Event HandleEnemyDestroyed (revised) ->",
      "  Play Sound 2D(sfx_hit.wav)",
      "  Spawn Actor from Class(BP_Explosion, Transform = enemy's current Transform)",
      "  Set Material(Self, Material Index 0, M_Enemy_White)",
      "  Delay(0.1)",
      "  EnemiesRemaining = EnemiesRemaining - 1",
      "  Score = Score + 10",
      "  Call UpdateHUD(HUDWidget, Score, Lives, Wave)",
      "  Branch(EnemiesRemaining <= 0) -> True:",
      "    Branch(Wave >= TotalWaves) -> True:",
      "      GameActive = false",
      "      Play Sound 2D(sfx_win.wav)",
      "      Call SetStatus(HUDWidget, \"You Win! Press R to Restart\")",
      "    False:",
      "      Wave = Wave + 1",
      "      EnemiesRemaining = 5",
      "      Play Sound 2D(sfx_wave.wav)",
      "      Call UpdateHUD(HUDWidget, Score, Lives, Wave)",
      "      Call SpawnWave",
      "  Destroy Actor (the enemy)",
    ],
    level2: [
      "Play a full round: every single hit shows a flash and an explosion, with no enemy left stuck mid-flash.",
      "Confirm score, wave-clear, and win/lose all still happen at exactly the moments they did last week - Delay only pauses this ONE execution wire, nothing else in the level pauses.",
      "Confirm BP_Explosion Actors never pile up in the World Outliner - each one destroys itself via its own Timeline, with no help from the enemy.",
    ],
    level3: [
      "Visual: instead of an instant Set Material snap, build M_Enemy_Flash with both enemy_invader_red.png and enemy_invader_white.png as separate Texture Samples, blended by a Lerp node driven by a scalar Blend Parameter, and animate that parameter from 1 (white) back down to 0 (red) across a Timeline instead of relying on the Delay - a genuine cross-fade instead of a hard cut.",
      "Visual: add a second Timeline track to BP_Explosion driving a new Opacity scalar Parameter on M_Explosion from 1 down to 0 as the shape scales up, so it fades out instead of popping away when destroyed - this needs M_Explosion's Blend Mode changed from Masked to Translucent first (Masked can only fully show or fully hide a pixel, not fade it), set through a Dynamic Material Instance the same way BP_Enemy's Color parameter works back in Lesson 4.",
      "Structure: notice BP_Explosion knows nothing about score, lives, or waves - it's a fully self-contained Actor, the same discipline as this course's bullets and enemies.",
    ],
    trouble: [
      "Enemy vanishes instantly, no flash visible: confirm the Delay(0.1) node was actually added, and Destroy Actor comes AFTER it, not before.",
      "Enemy renders solid black on the flash instead of white: confirm M_Enemy_White exists as its own separate Material (Texture Sample -> Emissive Color, with enemy_invader_white.png actually assigned), and that Set Material is pointing at M_Enemy_White, not accidentally back at M_Enemy_Red.",
      "Explosion never appears: confirm BP_Explosion is actually referenced correctly in Spawn Actor from Class, and that its Timeline is set to auto-play on BeginPlay.",
      "Explosion appears as a plain gray square instead of the burst shape, or doesn't appear at all: confirm M_Explosion's Texture Sample has both its RGB output wired to Emissive Color AND its Alpha output wired to Opacity Mask, that explosion_placeholder.png is actually assigned as its Texture, and that the Material's Blend Mode is actually set to Masked, not left on the default Opaque.",
      "Explosion doesn't clean itself up: confirm its Timeline's Finished pin is actually wired to Destroy Actor.",
    ],
    resource: "dev.epicgames.com/community/learning, Student Resource List > Unreal",
    altPath: {
      title: "C++ (code, optional advanced track)",
      intro: "The same flash-then-explode idea works in C++ using a Timer Handle in place of a Blueprint Delay node.",
      steps: [
        "Spawn ABPExplosion and call SetMaterial() on the enemy's mesh component to swap in M_Enemy_White at the top of your kill-handling function, then use GetWorldTimerManager().SetTimer() with a 0.1-second delay to call a second function containing the rest of the existing kill logic (EnemiesRemaining update, win-check, Destroy).",
      ],
      note: "This mirrors the Blueprint path's Delay node exactly - same pause, just implemented with a Timer Handle instead.",
    },
  },
];

handoutData.forEach(d => jobs.push([buildHandout(d), `Lesson_13_Handout_${d.file}.docx`]));

// =========================================================================
// TEACHER'S KEYS
// =========================================================================
function buildKey(data) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Lesson 13 Teacher's Key: A Hit That Reacts", bold: true, size: 30, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: data.engine + " - Instructor Reference Only", size: 24, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    h2("What This Lesson Teaches"),
    p(data.teaches),
    h2("Level 1 - Full Worked Solution"),
    ...data.level1.map(s => bullet(s)),
  ];
  if (data.level1Code) children.push(code(data.level1Code));
  children.push(
    h2("Why This Structure"),
    ...data.why.map(s => bullet(s)),
    h2("Level 2 / Level 3 Talking Points"),
    ...data.talkingPoints.map(s => bullet(s)),
    h2("Common Student Mistakes"),
    ...data.mistakes.map(s => bullet(s)),
  );
  return new Document({
    numbering: numberingConfig,
    styles: stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader("Lesson 13 Key - " + data.engine + " (Instructor Only)") },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

const keyData = [
  {
    engine: "Scratch", file: "Scratch",
    teaches: "Sequencing a visual reaction BEFORE finalizing game-state logic using Scratch's \"wait\" block, and spawning a fully independent visual effect via broadcast - the first lesson where the enemy's removal is deliberately delayed relative to when it was hit.",
    level1: [
      "The Explosion sprite is a completely separate, self-contained script - it never touches score, lives, or wave variables, and its own \"wait 0.2 seconds -> hide\" lifecycle needs no help from the Enemy.",
      "\"broadcast [explosion]\" fires and moves on immediately - Scratch broadcasts never block the sender, which is exactly why the explosion can be triggered without disturbing the flash-then-wait sequence that follows it.",
      "Moving the win-check/delete block below \"wait 0.1 seconds\" is the only structural change to existing logic - the logic itself is byte-for-byte the same as Lesson 12, just running slightly later.",
    ],
    why: [
      "Keeping the explosion as a broadcast-triggered, fully independent script (rather than something the Enemy waits on) means a slow or buggy explosion animation could never accidentally freeze the actual game logic - a good habit even at this small scale.",
      "This is the first lesson where 'wait' pauses something that MATTERS (delaying a state change) rather than just pacing a visual (Lesson 2's beep-beep) - worth naming that distinction explicitly.",
    ],
    talkingPoints: [
      "Ask students to trace, step by step, what happens between a bullet touching an enemy and the enemy actually disappearing now - flash, wait, THEN the same Lesson 6-12 logic runs. A good comprehension check for sequencing.",
      "For strong students: ask what would happen if \"broadcast [explosion]\" were accidentally placed AFTER \"wait 0.1 seconds\" instead of before it - the explosion would appear 0.1 seconds late, slightly out of sync with the flash. Not a bug exactly, but worth noticing as a timing choice.",
    ],
    mistakes: [
      "Placing the win-check/delete block BEFORE the wait instead of after, which defeats the whole point (the enemy would still vanish instantly).",
      "Forgetting to reset \"set [color] effect to 0\" somewhere in the enemy's reset script, causing a revived/respawned enemy to start already tinted.",
      "Accidentally duplicating the score/win-check lines when moving them below the wait, causing double-counting.",
    ],
  },
  {
    engine: "Godot Engine", file: "Godot",
    teaches: "Using Godot 4's await keyword to pause a single function mid-execution without freezing anything else - a genuinely new language feature this course hasn't needed until now, introduced because this is the first moment a function needs to do something, then wait, then finish.",
    level1: [
      "await get_tree().create_timer(0.1).timeout pauses ONLY this specific call of _on_area_entered() - every other node, script, and function in the game keeps running normally.",
      "The explosion is instantiated and added to the scene BEFORE the await, so it starts playing out immediately and independently, with its own Timer managing its own cleanup.",
      "destroyed.emit() and queue_free() are unchanged from Lesson 4/7 - they just now run after the await instead of immediately.",
    ],
    why: [
      "await is the right tool here specifically because the pause needs to happen INSIDE an existing function, without duplicating or restructuring the surrounding collision-detection code - a much smaller change than GameMaker or TIC-80 need for the same idea, since GDScript supports pausing execution directly.",
      "Giving the explosion its own scene with its own Timer (rather than having the Enemy manage the explosion's lifetime) keeps every object responsible for cleaning up only itself - the same discipline bullets and enemies have followed all course.",
      "The flash swaps texture rather than tinting modulate white, for the same reason Lesson 2's Recolor swaps sprites instead of tinting them: modulate multiplies against the sprite's existing color, and multiplying any color by white is a no-op - our enemy sprite is opaque red, so a white modulate tint would produce zero visible change.",
    ],
    talkingPoints: [
      "Ask students to explain, in their own words, why await here doesn't freeze the whole game the way a bug sometimes might - because it only pauses the specific function call it's written inside, not the entire engine.",
      "For students comparing engines: point out that Unity's Coroutine and Godot's await solve the identical problem (pause mid-function) with similar underlying mechanisms, while GameMaker and TIC-80 need a genuinely different structural approach (splitting into two events/checks) because their languages don't support pausing mid-script at all.",
    ],
    mistakes: [
      "Forgetting the await keyword and just calling get_tree().create_timer(0.1).timeout without it - this creates the timer but doesn't actually pause anything, so the function continues immediately.",
      "Instantiating the explosion AFTER the await instead of before, delaying the visual effect for no reason.",
      "Not guarding against a second bullet hitting the same enemy while it's mid-await, causing a double-destroyed.emit() call - freeing the bullet immediately (before the await) is what prevents this.",
    ],
  },
  {
    engine: "Unity", file: "Unity",
    teaches: "Introducing Coroutines - Unity's built-in tool for pausing a function mid-execution using yield return, without freezing the rest of the game - genuinely new this week, and directly motivated by needing exactly this capability for the first time.",
    level1: [
      "HandleHit() is a Coroutine (IEnumerator return type, started with StartCoroutine) rather than a plain function - this is required because yield return WaitForSeconds() only works inside a Coroutine.",
      "The explosion is instantiated and the flash sprite swapped in BEFORE the yield, so both happen immediately; RegisterKill() and Destroy() are unchanged from Lesson 4/7, just running after the yield instead of immediately.",
      "other.gameObject is destroyed at the very top of OnTriggerEnter2D, before StartCoroutine runs - this prevents a second overlapping bullet from double-triggering the same enemy while it's mid-Coroutine.",
    ],
    why: [
      "Coroutines are the right tool here specifically because the pause needs to happen INSIDE existing hit-handling logic without restructuring it across multiple objects - a much smaller change than GameMaker or TIC-80 need for the same idea.",
      "Giving the explosion prefab its own Destroy(gameObject, 0.3f) call in its own Start() (rather than having the Enemy manage its lifetime) keeps every object responsible for cleaning up only itself - the same discipline bullets and enemies have followed all course.",
      "The flash swaps sprite rather than tinting spriteRenderer.color white, for the same reason Lesson 2's Recolor swaps sprites instead of tinting them: color multiplies against the sprite's existing color, and multiplying any color by white is a no-op - our enemy sprite is opaque red, so a white color tint would produce zero visible change.",
    ],
    talkingPoints: [
      "Ask students to explain why HandleHit() needs to be IEnumerator instead of void - because yield return can only appear inside a function with that return type, a small but important C# rule worth naming explicitly.",
      "For students comparing engines: Godot's await and Unity's Coroutine solve the identical problem with similar underlying mechanisms, while GameMaker and TIC-80 need a genuinely different structural approach because their languages don't support pausing mid-script at all.",
    ],
    mistakes: [
      "Writing HandleHit() as a regular void function instead of IEnumerator, which won't compile once yield return is added.",
      "Calling HandleHit() directly instead of StartCoroutine(HandleHit()) - this also won't compile, and is an easy habit to fall into from months of calling regular functions directly.",
      "Forgetting to destroy the bullet before starting the Coroutine, allowing a second bullet to also trigger OnTriggerEnter2D on the same, already-dying enemy.",
    ],
  },
  {
    engine: "GameMaker", file: "GameMaker",
    teaches: "Introducing Alarms - GameMaker's built-in 'do something after N steps' mechanism - genuinely new this week, and necessary because GML can't pause mid-event the way Godot's await or Unity's Coroutine can. This requires splitting the hit-handling logic across TWO events for the first time.",
    level1: [
      "obj_bullet's Collision event keeps ALL of its existing logic (sound, score, win-check) completely unchanged and immediate - only instance_destroy(other) is replaced with a flash + an alarm.",
      "obj_enemy's new Alarm 0 event is deliberately tiny - just instance_destroy() - because all the meaningful game-state logic already ran in the Collision event, 6 steps earlier.",
      "obj_explosion is a fully separate object with its own Create-event alarm and its own Alarm 0 event - it manages its own lifetime with zero involvement from obj_enemy or obj_bullet.",
    ],
    why: [
      "Splitting into two events (Collision starts the flash + alarm, Alarm 0 finishes the removal) is the idiomatic GameMaker way to express 'do this, then after a pause, do that' - GML doesn't support blocking waits inside a single event the way Godot or Unity's languages do.",
      "Keeping the win-check/score logic in the Collision event (unchanged) rather than moving it into the Alarm 0 event keeps this week's change small and localized - GameMaker students are learning ONE new concept (Alarms), not restructuring everything they already built.",
      "The flash swaps sprite_index rather than tinting image_blend white, for the same reason Lesson 2's Recolor swaps sprites instead of tinting them: image_blend multiplies against the sprite's existing color, and multiplying any color by white is a no-op - our enemy sprite is opaque red, so a white image_blend would produce zero visible change.",
    ],
    talkingPoints: [
      "Ask students to trace exactly which lines moved and which didn't between Lesson 12 and this week - only instance_destroy(other) changed location (into a new event); everything else is untouched.",
      "For students comparing engines: point out that GameMaker's alarm[0] + Alarm 0 event is conceptually identical to TIC-80's flash-countdown-in-a-table approach this week - both engines lack a blocking wait, so both split 'trigger' and 'finish' into two separate places.",
    ],
    mistakes: [
      "Leaving instance_destroy(other) in place alongside the new alarm code, causing the enemy to still vanish instantly (the alarm gets set but never has a chance to fire).",
      "Forgetting to add the new Alarm 0 event to obj_enemy at all - setting alarm[0] does nothing without a matching Alarm event to catch it.",
      "Accidentally moving score/win-check logic INTO the new Alarm 0 event instead of leaving it in the Collision event - this would delay scoring and win detection by 6 steps, which is unnecessary and inconsistent with every other engine this week.",
    ],
  },
  {
    engine: "GDevelop", file: "GDevelop",
    teaches: "Using GDevelop's built-in \"Wait X seconds\" action to pause the remaining actions in an event, plus the official Flash extension for a one-line hit-flash - two genuinely new built-in tools this week, both requiring zero custom logic to use correctly.",
    level1: [
      "Wait 0.1 seconds pauses only the REST of this specific event's action list - EnemiesRemaining, Score, and the sound already changed by the time the wait begins, so they're unaffected.",
      "The separate full-win/advance-wave events (from Lesson 11) still key off EnemiesRemaining, which changes immediately in this event - so wave/win timing is completely unaffected by today's changes.",
      "The Explosion object's self-cleanup event (\"has existed for more than 0.3 seconds\") is a fully separate, standalone event - it needs no signal or message from the collision event at all.",
    ],
    why: [
      "GDevelop's Wait action is the right tool here specifically because it pauses in place within a single event's action list, without needing a Coroutine, Alarm, or manual timer variable - the most direct translation of \"do this, pause, then do that\" of any engine this week.",
      "Keeping the Explosion object's cleanup as its own condition-based event (rather than something the collision event manages) keeps every object responsible for cleaning up only itself - the same discipline bullets and enemies have followed all course.",
      "Level 1 deliberately uses the Flash extension's blink action rather than its color-tint action: blink just toggles visibility, so it works no matter what color the Enemy sprite is; a color tint multiplies against the sprite's existing color, and a white tint on our opaque-red Enemy would be a no-op, same underlying issue Lesson 2's Recolor was rewritten to avoid.",
    ],
    talkingPoints: [
      "Ask students to identify exactly which action in the collision event moved (Delete Enemy, to after the Wait) versus which stayed in place (everything else) - a good comprehension check for what's actually new this week.",
      "For students ready for a stretch: point out the Tween behavior confirmed in the Level 3 idea, and discuss how it's a different, more flexible tool than the Flash extension - Flash is for a quick blink or color pulse, Tween is for animating any property (scale, opacity, position) smoothly over time. If a student tries the Flash extension's color-tint action, use it as a chance to talk about why a \"white\" tint doesn't look white on a colored sprite (multiply blend), the same idea behind Lesson 2's sprite-swap Recolor.",
    ],
    mistakes: [
      "Placing Delete Enemy BEFORE the Wait action instead of after, which defeats the whole point (the enemy would still vanish instantly).",
      "Installing the Flash extension but forgetting to add the Flash behavior to Enemy specifically (Behaviors tab > Add a behavior > Flash) - the extension only makes the behavior available project-wide; the blink action still won't appear for Enemy until the behavior is actually attached to it.",
      "Building the Explosion self-cleanup condition incorrectly (e.g. checking the wrong object, or a typo'd name), leaving explosions on screen permanently.",
    ],
  },
  {
    engine: "TIC-80", file: "TIC-80",
    teaches: "Extending each enemy's table entry with a new 'flash' field, and adding a genuinely new per-frame update loop to manage it - TIC-80's version of GameMaker's Alarm, built from the same plain Lua tools (tables, loops, variables) this engine has used all course, since it has no blocking wait either.",
    level1: [
      "e.flash = 6 marks a hit enemy for delayed removal instead of removing it immediately - the enemy stays in the table, just in a new \"flashing\" state.",
      "The new update loop runs every frame from TIC(), counting down every flashing enemy's and every explosion's timer, removing each once it reaches 0 - this is the engine's substitute for a blocking wait.",
      "Score, enemies_remaining, and the win-check all still happen immediately in the hit branch, completely unchanged from Lesson 12 - only the actual table.remove() for the enemy is delayed.",
    ],
    why: [
      "Adding a flash field to each enemy's own table entry (rather than a separate parallel table) keeps each enemy's full state - position, alive-ness, now flash status - in one place, the same 'one entity, one table entry' discipline TIC-80 has used since Lesson 4.",
      "A dedicated explosions table (separate from enemies) is appropriate here because explosions and enemies are genuinely different kinds of things with different lifetimes and different drawing logic - unlike the flash field, which is naturally just an enemy's temporary state.",
    ],
    talkingPoints: [
      "Ask students to trace what happens to a single enemy's table entry across several frames after being hit: e.flash=6, drawn as the flash sprite, decremented each frame, removed at 0 - a good comprehension check for how table-based state persists across frames.",
      "For students comparing engines: point out that TIC-80's flash-countdown-in-a-table and GameMaker's alarm[0]-plus-Alarm-event solve the identical problem - both engines lack a blocking wait, so both split 'trigger' and 'finish' into two separate places in code.",
    ],
    mistakes: [
      "Removing the enemy from the table immediately in the hit branch (old habit from Lesson 4-12), which defeats the whole point - the flash field approach requires NOT removing it there anymore.",
      "Forgetting to call the new update loop from TIC() every frame, leaving flashing enemies and explosions on screen forever.",
      "Drawing logic that doesn't check e.flash at all, so hit enemies never actually show the flash sprite even though the field is being set correctly.",
    ],
  },
  {
    engine: "001 Game Creator (Windows only)", file: "001GameCreator",
    teaches: "Using the confirmed built-in Actor – Flash action and Spawn Actor for a self-managed hit reaction - this engine's version of this week's idea, with an honest, documented limitation around exact flash-then-destroy timing that the other seven engines don't share.",
    level1: [
      "Actor – Flash is a self-contained action - it manages its own color, repeat count, and duration internally, unlike every other engine this week, which needed a manually-authored delay.",
      "Spawn Actor(Explosion) fires immediately and independently, exactly like every other engine's explosion this week - this part of the lesson is fully consistent across all 8 engines.",
      "The enemy's Destroy action is left in its existing position (unchanged from Lesson 7 onward) - this is a deliberate, documented simplification for this engine specifically, not an oversight.",
    ],
    why: [
      "Unlike Godot/Unity's await/Coroutine or GameMaker/TIC-80's alarm/countdown split, 001 Game Creator's confirmed API doesn't include a verified general-purpose 'wait N seconds mid-trigger' action - rather than fabricate one, this lesson keeps Destroy unchanged and treats the explosion effect as the primary, always-reliable visual for this engine.",
      "This mirrors the same honest-uncertainty approach this course has used before (e.g. Lesson 10's choice of a verified manual restart over an unverified built-in one) - a real, if imperfect, engine limitation, clearly labeled as such rather than papered over.",
    ],
    talkingPoints: [
      "Ask students on this engine to actually test whether the flash is visible in their installed version - some versions/settings may show a brief flash even with an immediate Destroy, others may not; either result is expected and fine for Level 1.",
      "For advanced students: if their version DOES expose a Timer or delay mechanism (check the Actions list together), this is a great Level 3 stretch to hold Destroy back briefly, matching what every other engine's Level 1 already does.",
    ],
    mistakes: [
      "Assuming the flash isn't working at all when it's actually just being cut short by the immediate Destroy - the explosion effect, not the enemy's own flash, is what to check first when troubleshooting this engine.",
      "Forgetting to assign a Graphics asset to the new Explosion Actor Template, resulting in an invisible (but otherwise functioning) spawned Actor.",
      "Not setting a Lifetime/Duration on the spawned Explosion (if available), leaving them on screen indefinitely across a long playthrough.",
    ],
  },
  {
    engine: "Unreal Engine (optional advanced track)", file: "Unreal",
    teaches: "Using a Blueprint Delay node to pause a Custom Event graph mid-execution, plus a Timeline-driven BP_Explosion Actor for the visual effect - both genuinely new this week, and a clean, idiomatic Blueprint way to express 'do this, pause, then do that.'",
    level1: [
      "Delay(0.1) pauses only THIS execution wire within HandleEnemyDestroyed - every other Actor, widget, and graph in the level keeps running normally, the same non-blocking behavior as Godot's await and Unity's Coroutine.",
      "BP_Explosion is spawned and the enemy's Material swapped to M_Enemy_White BEFORE the Delay, so both happen immediately; EnemiesRemaining, the win-check Branch, and Destroy Actor are unchanged from Lesson 11/12, just running after the Delay instead of immediately.",
      "The flash is a real Set Material swap to a separate, pre-made Material (M_Enemy_White, built in this week's Setup from the supplied enemy_invader_white.png sprite) rather than a color tint - Set Vector Parameter Value(\"Color\", White) would have been a documented no-op here, since M_Enemy_Red's Color parameter already defaults to White and multiplying by White again changes nothing. This is exactly the trap every other engine's own instructions warn students away from this week (see e.g. Unity Visual Scripting's and GameMaker's own trouble notes).",
      "BP_Explosion's own Timeline (BeginPlay -> scale up -> On Finished -> Destroy Actor) makes it fully self-contained, needing no help from the enemy or level Blueprint to clean itself up.",
    ],
    why: [
      "A Delay node is the right tool here specifically because it pauses in place within an existing Custom Event graph, without restructuring anything else - a much smaller change than GameMaker or TIC-80 need for the same idea, since Blueprint execution wires support pausing directly.",
      "Swapping to a separate white-sprite Material (rather than tinting one Material's color) matches every other engine's actual technique this week - Godot/Unity/GameMaker/GDevelop/TIC-80/001 Game Creator all swap in a genuine white version of the enemy image rather than trying to tint their way there, for the same reason: a color multiply against an already-colored sprite can't produce a true, opaque white flash.",
    ],
    talkingPoints: [
      "Ask students to trace the full execution path through HandleEnemyDestroyed now: sound, spawn explosion, flash Material, Delay, THEN everything from Lesson 11/12 unchanged - a good comprehension check for how a Delay node fits into an existing graph.",
      "For students comparing engines: Godot's await, Unity's Coroutine, and Unreal's Delay node all solve the identical problem with genuinely different underlying mechanisms, while GameMaker and TIC-80 need a structurally different approach (splitting into two events/checks) because their languages don't support pausing mid-script at all. The flash itself, meanwhile, is one of the few techniques that's genuinely identical across all eight engines this week: swap in a real white sprite/Material, don't try to tint one.",
    ],
    mistakes: [
      "Wiring Destroy Actor and the rest of the kill logic BEFORE the Delay node instead of after, which defeats the whole point (the enemy would still vanish instantly).",
      "Forgetting to set BP_Explosion's Timeline to auto-play, so it spawns but never actually animates or destroys itself.",
      "Reaching for Set Vector Parameter Value(\"Color\", White) out of habit from Lesson 4/5, instead of Set Material(M_Enemy_White) - it compiles and runs with no error, but produces no visible flash at all, since White is already the Color parameter's neutral default.",
      "Editing M_Enemy_Red's own asset directly to try to change the look, instead of calling Set Material on the hit enemy - that would restyle every enemy in the level at once, not just the one that was hit.",
    ],
  },
];

keyData.forEach(d => jobs.push([buildKey(d), `Lesson_13_Key_${d.file}.docx`]));

// =========================================================================
// WRITE ALL FILES
// =========================================================================
(async () => {
  for (const [doc, filename] of jobs) {
    const buf = await Packer.toBuffer(doc);
    fs.writeFileSync(filename, buf);
    console.log("wrote", filename);
  }
})();
