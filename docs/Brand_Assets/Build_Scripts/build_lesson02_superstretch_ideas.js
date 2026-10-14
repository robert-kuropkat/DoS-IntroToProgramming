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
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function labelBullet(label, text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)] });
}

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
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
    headers: { default: brand.pageHeader("Super-Stretch Ideas") },
    footers: { default: brand.pageFooter() },
    children: [
      brand.crestParagraph(1.3, { after: 160 }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Super-Stretch Ideas", bold: true, size: 40, color: brand.NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Beyond Level 3: a curiosity list, not an assignment", size: 24, color: brand.NAVY2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

      h1("What This Is"),
      p("Every lesson's Level 3 is a short, deliberately achievable menu of stretch goals - all of them worked out and confirmed possible in every supported engine before they ever reach a Handout. This document is different. It's a longer, generic list of \"what else is out there\" ideas for a student who finishes Level 3 early, or who's just curious what game development looks like beyond this week's exercise."),
      p("A few important differences from Level 3:"),
      bullet("Nothing here is required, tracked, or graded. It's fine - expected, even - for a student to read this list and try nothing on it."),
      bullet("These ideas are written generically, not for a specific engine, and not against this lesson's exact Key Press Reactor exercise. Some will be easy in one engine and awkward or impossible in another; some describe concepts this course covers properly in a later lesson; some may need looking things up outside class."),
      bullet("Unlike the core lesson materials, self-containment doesn't apply here - a student chasing one of these ideas is welcome (encouraged, even) to search the engine's own documentation, forums, or tutorials. That kind of independent digging is the point."),
      bullet("If an idea looks interesting but the how isn't obvious, that's a good reason to ask - in class, or by flagging it for a future lesson."),

      h1("Movement & Feel"),
      labelBullet("Ease in and out", "Instead of snapping to full speed instantly, ramp up to it (acceleration) and slow back down when a key is released (deceleration/friction), instead of stopping dead."),
      labelBullet("Normalize diagonal movement", "If two direction keys are held at once, moving diagonally shouldn't be faster than moving in a single direction - scale the combined movement back down to match."),
      labelBullet("Bounce instead of stop", "At the screen edge, instead of just clamping in place, reverse direction (or velocity) so the character bounces off the edge."),
      labelBullet("Dash or sprint", "Add a second key (like Shift) that briefly boosts speed, ideally with a cooldown so it can't just be held down permanently."),
      labelBullet("Gravity and a jump", "Instead of free movement in all 4 directions, try a side-view character that's pulled downward by gravity and has a single jump button - the beginning of a platformer."),
      labelBullet("Camera follows the character", "Instead of a fixed view, have the camera/viewport scroll to keep the character in frame as it moves - useful once a level is bigger than one screen."),

      h1("Visual Polish (\"Juice\")"),
      labelBullet("True in-engine color tinting", "This lesson's Recolor goal swaps in a whole separate pre-made colored image for each direction, since most engines' quick tint options (image_blend, modulate, spriteRenderer.color, Tint color, and similar) only give a clean, true color on plain white/grayscale art - tint a colored sprite and the result comes out muddy, a multiply of the two colors rather than a clean swap. If you're curious why, try it yourself: duplicate your character sprite, fill it solid white, and tint that instead of swapping images - you should get a clean, exact color for every direction with a single sprite, no swapping needed. Worth knowing which engines can do this cleanly and which can't, and why TIC-80 is the one exception in this lesson that never had the problem to begin with (see its Recolor notes)."),
      labelBullet("Face the direction of movement", "Flip or mirror the character horizontally (or rotate it) so it visually faces the direction it's currently moving."),
      labelBullet("Squash and stretch", "Briefly scale the character slightly wider/shorter when it starts moving and slightly taller/thinner when it stops, for a springy, less rigid feel."),
      labelBullet("Trail or particle effect", "Spawn a fading trail, dust puff, or sparkle effect while the character is in motion."),
      labelBullet("Screen shake", "Briefly jitter the camera or view for an impact moment, e.g. when hitting the screen edge."),
      labelBullet("Fade in or out", "Have the character (or the whole scene) fade in at the start, or fade out when triggered, instead of just appearing/disappearing instantly."),

      h1("Audio"),
      labelBullet("Footsteps synced to animation", "Instead of one looping tone, trigger a short sound on each animation-frame change, so it reads as footsteps rather than a siren."),
      labelBullet("Background music", "Add a simple looping music track that plays for the whole scene, separate from the movement sound effect."),
      labelBullet("Vary pitch or volume", "Slightly randomize a repeated sound effect's pitch or volume each time it plays so it doesn't feel robotic on repeat."),

      h1("Game State & Feedback"),
      labelBullet("On-screen counter", "Display a running count of something - button presses, steps taken, seconds survived - updating live as text on screen."),
      labelBullet("Simple timer", "Start a countdown or count-up timer when the scene begins and display it."),
      labelBullet("Tunable variables", "Expose a value like speed as a named variable you (or the engine's own inspector/editor UI) can tweak without digging back into the logic - useful for quickly trying different \"feels.\""),

      h1("Input & Control"),
      labelBullet("A second player", "Add a second character controlled by a different set of keys (e.g. WASD alongside the arrow keys), sharing the same screen."),
      labelBullet("Gamepad support", "Get the same movement working from a connected controller, alongside (not instead of) the keyboard."),
      labelBullet("On-screen touch controls", "Add clickable/tappable on-screen buttons or a virtual joystick as an alternative to keyboard input."),
      labelBullet("True press-and-hold movement (001 Game Creator)", "This lesson's 001 Game Creator path moves the character once per key press, since Input Triggers only fire on Pressed/Released, not while a key is held. For movement that continues the whole time a key is down, use the Change Movement X / Change Movement Y events instead - set them to 1 or -1 on Pressed and back to 0 on Released, and the engine keeps applying that movement every frame on its own using the actor's Speed and Acceleration, no loop required. A nice look at how a single engine can solve \"hold to move\" a completely different way than a polling loop does."),
      labelBullet("Remappable controls (Defined Control)", "Instead of binding a script straight to one physical key, some engines let you define a named control - like \"Move Right\" - and then attach several different inputs to it (arrow key, WASD key, a gamepad button) so your scripts never have to change no matter which one the player uses. 001 Game Creator's Input Triggers have exactly this via their \"Defined Control\" option; look for an equivalent remapping system in whichever engine you're using."),
      labelBullet("Reading multiple keys reliably at once (Scratch)", "Scratch's \"when key pressed\" scripts repeat by riding your operating system's own keyboard auto-repeat feature - the same mechanism that repeats a letter when you hold it down in a text box. Most operating systems only actively auto-repeat one key at a time, so pressing a second key (like Space to fire) while a movement key is already held can interrupt the repeat stream for the first key, making movement pause until it's pressed again. The fix: poll real key state instead of reacting to key-press events, using the Sensing category's \"key ( ) pressed?\" reporter inside a \"forever\" loop - since that reads whether a key is currently held down once per frame, it isn't affected by the operating system's one-key-at-a-time auto-repeat limit at all."),

      h1("Advanced / A Peek Ahead"),
      labelBullet("A wandering or chasing companion", "Add a second sprite that moves on its own - wandering randomly, or moving toward/away from the player character - an early taste of simple AI."),
      labelBullet("Remember state between runs", "Save something (like the character's last position, or a high score) so it's still there the next time the project runs."),
      labelBullet("Randomize on start", "Have the character's color, size, or starting position be chosen randomly each time the project runs, instead of always the same."),
      labelBullet("Production-grade input systems", "Some engines offer a more advanced, remappable input system (e.g. Unreal's Enhanced Input, Unity's Input System package) as an alternative to the simplest built-in key checks - worth a look for anyone curious how bigger, shipped games handle input."),

      h1("A Note for the Curious"),
      p("None of this is a checklist to work through in order, and it's genuinely fine if nothing here appeals yet - most of it will make a lot more sense, and be a lot easier to build, after a few more lessons. It's here so that \"what else could I try?\" always has an answer, even on a week where Level 3 gets finished with time to spare."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Lesson_02_SuperStretch_Ideas.docx", buffer);
  console.log("done");
});
