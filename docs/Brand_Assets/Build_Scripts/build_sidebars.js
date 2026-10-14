const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, ImageRun,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, WidthType, ShadingType, VerticalAlign
} = require("docx");
const brand = require("./brand.js");

function diagram(file, widthIn, heightIn, caption) {
  const out = [
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { before: 160, after: caption ? 40 : 160 },
      children: [new ImageRun({ data: fs.readFileSync(file), transformation: { width: widthIn * 96, height: heightIn * 96 }, type: "png" })],
    }),
  ];
  if (caption) out.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [new TextRun({ text: caption, italics: true, size: 18, color: "555555" })] }));
  return out;
}

function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function hook(text) { return new Paragraph({ spacing: { after: 200 }, children: [new TextRun({ text, italics: true, color: brand.NAVY2, size: 24 })] }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function labelBullet(label, text) {
  return new Paragraph({ numbering: { reference: "bullets", level: 0 }, spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)] });
}
function divider() {
  return new Paragraph({ spacing: { before: 100, after: 300 }, border: { bottom: { color: brand.GOLD, space: 4, style: "single", size: 6 } }, children: [] });
}
function code(lines) {
  const runs = [];
  lines.forEach((line, i) => {
    runs.push(new TextRun({ text: line, font: "Consolas", size: 20 }));
    if (i < lines.length - 1) runs.push(new TextRun({ text: "", break: 1 }));
  });
  return new Paragraph({
    spacing: { after: 160, before: 80 }, indent: { left: 360 },
    shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
    children: runs
  });
}

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]
};

function buildSidebarDoc(lessonNum, lessonTitle, topics) {
  const children = [
    brand.crestParagraph(1.2, { after: 140 }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: `Lesson ${lessonNum} Sidebars`, bold: true, size: 32, color: brand.NAVY })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: lessonTitle, size: 22, color: brand.NAVY2 })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
    p("These aren't part of the lesson's required work. Think of them the way a textbook sidebar works: a deeper, tangential look at something this lesson's build actually ran into - why one approach was chosen over the others, what the alternatives were, or an idea worth a longer look than the Handout has room for. Pull one out for a discussion moment if it fits, hand it to a curious student, or skip it entirely - nothing here is required reading.", { italics: true, color: "555555" }),
  ];
  topics.forEach((t, i) => {
    if (i > 0) children.push(divider());
    children.push(h1(t.title));
    children.push(hook(t.hookText));
    t.body.forEach(block => {
      if (typeof block === "string") children.push(p(block));
      else if (block.bullet) children.push(labelBullet(block.bullet, block.text));
      else if (block.plainBullet) children.push(bullet(block.plainBullet));
      else if (block.subhead) children.push(h2(block.subhead));
      else if (block.code) children.push(code(block.code));
      else if (block.diagram) children.push(...diagram(block.diagram.file, block.diagram.w, block.diagram.h, block.diagram.caption));
    });
  });

  return new Document({
    numbering: numberingConfig,
    styles: brand.stylesBlock,
    sections: [{
      properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
      headers: { default: brand.pageHeader(`Lesson ${lessonNum} Sidebars`) },
      footers: { default: brand.pageFooter() },
      children,
    }]
  });
}

// =========================================================================
// LESSON 2 - three sidebars
// =========================================================================
const lesson02 = [
  {
    title: "Changing Color: Tint, Swap, or Something Fancier?",
    hookText: "Every engine in this lesson needed a way to change color per direction. Why did that one goal split into two completely different techniques - and which one would an actual studio ship?",
    body: [
      "There are two genuinely different families of technique for \"make this sprite a different color,\" and this lesson ran into both. The first is color tinting (also called blending or modulating): multiply every pixel's color by a tint color, channel by channel. Godot's modulate, Unity's spriteRenderer.color, GameMaker's image_blend, and GDevelop's Tint color are all this same operation under different names. The second is sprite swapping: instead of recoloring anything, just show a completely different, pre-made image.",
      "The catch with tinting: multiplying a pixel by a color can only ever darken or mix it - it can never brighten past what was already there. That means a tint only produces a clean, true color when the source art is white or grayscale to begin with (white times anything equals that color exactly). Tint a sprite that's already colored - like this course's placeholder art - and the result is a multiply of two colors, not a clean replacement. It comes out muddy, not wrong exactly, just not the color you asked for.",
      "This course standardizes on sprite-swapping instead, for three reasons: it's the simpler concept for this age/skill range (\"show a different picture\" needs no explanation of multiply-blend math), it needed zero new assets beyond the color-per-direction art pack already being used, and it works identically in every engine with no exceptions to explain.",
      { bullet: "The one deliberate exception: TIC-80", text: "TIC-80's poke4() genuinely recolors in place, with no muddy-tint problem at all - because it isn't multiplying pixel colors together. TIC-80 draws using a small indexed palette (a fixed list of colors, each referenced by a number), and poke4() rewrites which actual RGB color a given palette index points to. Every pixel using that index instantly shows the new color, exactly, regardless of what the old color was. It's a genuinely different operation from tinting, not a simpler version of the same one - which is why it was kept as-is rather than switched to sprite-swap." },
      { bullet: "What Scratch is actually doing", text: "Scratch's \"set color effect\" (used for this lesson's Animate/blink goal, not Recolor) is a third category on its own - it cyclically shifts hue rather than multiplying against a tint color, which is why it doesn't have the muddy-on-colored-art problem either, but also doesn't let you land on one exact target color the way a true palette swap does." },
      { bullet: "What \"the pros\" actually do", text: "Real, shipped games use both families, deliberately. Palette-swap recoloring (TIC-80's approach, at a bigger scale) is a genuinely old, genuinely common trick - it's how a lot of classic games gave enemies recolored \"palette swap\" variants cheaply, and it's still how some modern games recolor uniforms/teams efficiently. True tinting is also used constantly, but professionally it's paired with source art built specifically for it: a white or grayscale \"mask\" texture made just for tinting, kept separate from the character's normal colored art, so the multiply always lands on an exact color. That's exactly the technique this course's tint/blend Super-Stretch idea invites a curious student to try by hand." },
    ],
  },
  {
    title: "The Blink Nobody Coded the Same Way Twice",
    hookText: "All 8 engines needed a sprite to blink while moving. All 8 used a genuinely different trick to do it. None of them is \"the right way\" - each is just the natural idiom for that engine.",
    body: [
      "Worth actually lining these up side by side, because the variety is the whole point:",
      { plainBullet: "Scratch: there's no \"key released\" event to trigger off of, so blinking needs a polling forever-loop instead, using the Sensing \"key pressed?\" reporter plus a timer-mod trick ((timer mod 0.4) compared against half its own period) to get a repeating on/off flip." },
      { plainBullet: "Godot: toggling the built-in visible property (inherited from every CanvasItem) on and off from an accumulating timer." },
      { plainBullet: "Unity: the same idea, toggling SpriteRenderer's built-in enabled property from an accumulating timer." },
      { plainBullet: "GameMaker: toggling the object's built-in visible variable, driven by a counter incremented in the Step event." },
      { plainBullet: "GDevelop: no manual timer at all - a purpose-built Flash behavior ships with its own \"Flash visibility (blink)\" action." },
      { plainBullet: "TIC-80: exploiting a documented, specific quirk of spr() - a scale factor of exactly 0 causes it to draw nothing at all - so alternating the scale argument between 0 and 1 makes the sprite blink with no second image and no visibility flag involved." },
      { plainBullet: "Unreal: the same visible-toggle idea as Godot/Unity, but on an Actor rather than a sprite - Set Actor Hidden In Game, driven from Event Tick." },
      { plainBullet: "001 Game Creator: doesn't build a blink at all. It falls out for free, as an honest side effect of how this engine's movement already works - see below." },
      "The 001 Game Creator case is worth slowing down for, since it's the most surprising one. This engine has no press-and-hold movement without extra work, so every \"step\" a student takes is really a full press-release cycle. The idle-reset script built for this lesson's Recolor goal fires on every release - including the ones in the middle of a walk, not just when a student actually stops. That reads as a blink while moving, and goes static the instant movement actually stops, purely as a side effect of how idle-detection was already built. It's explicitly documented as a deliberate, explainable consequence rather than a bug, and it's expected to disappear once true press-and-hold movement is introduced (an existing Super-Stretch idea) - at that point Released will only fire on a genuine stop.",
      "The takeaway worth naming out loud: \"the right tool\" genuinely depends on what a given engine exposes. A built-in behavior beats hand-rolled code when one exists (GDevelop). A documented rendering quirk can be repurposed cleverly (TIC-80). And sometimes the real \"solution\" isn't new code at all - it's noticing what code you already wrote is already doing (001 Game Creator).",
    ],
  },
  {
    title: "Which Way Is \"Up\"? It Depends Which Engine You Ask",
    hookText: "Move a sprite \"up\" in most engines this course uses and its Y value goes down, not up. That's not a GameMaker quirk - it's the majority convention. The real exceptions are worth knowing before they cause confusion.",
    body: [
      "There are two genuinely different coordinate conventions hiding across this course's 8 engines. Screen/canvas coordinates put the origin (0,0) at the top-left corner, with Y increasing downward - so moving \"up\" means subtracting from Y, not adding to it. Cartesian coordinates work the way a math-class graph does - an origin wherever makes sense, with Y increasing upward, the way it's normally drawn on paper.",
      "GameMaker, Godot, and TIC-80 all confirmed use the screen/canvas convention (Y-down, origin top-left) in this course's own Lesson 2 movement code - GDevelop's built-in Top-down Movement behavior almost certainly works the same way under the hood, since it shares the same canvas heritage. The real exceptions are Scratch, whose origin is the CENTER of the stage with Y increasing upward (a deliberate, math-classroom-style choice), and Unity, whose world space is genuinely Cartesian and Y-up even in 2D, since Unity's \"2D\" is really a 3D scene viewed from one fixed angle.",
      "Why the screen convention exists at all: it traces back to how a raster display draws an image, one row at a time starting from the top - the same convention nearly every image file format and 2D graphics API (including a web browser's own <canvas>) still uses today. GameMaker, Godot, TIC-80, and GDevelop all inherit that heritage directly. Scratch deliberately broke from it on purpose, aiming at people - often kids - who are more likely to have already met an x/y graph in math class than a raster display's internals.",
      { bullet: "Unreal is a third case, worth its own callout", text: "Unreal is also a full 3D Cartesian engine, but by convention it's left-handed with Z as the \"up\" axis, not Y, and X as \"forward.\" That means moving between Unity and Unreal isn't just a sign flip - it's a genuinely different axis that means \"up\" in the first place. Worth flagging directly for any student working in both." },
      "The practical habit worth building: before writing a single line of movement code in a new engine, check where the origin actually is and which axis - and which direction along it - counts as \"up.\" Never assume it matches the last engine.",
    ],
  },
  {
    title: "Scope: Where a Variable Lives, and How Long It Lasts",
    hookText: "GameMaker's \"is any direction key held\" check turned into a lesson about how visual and typed code diverge for the same job. A completely separate bug, caught live while building a later lesson, turned out to be the same underlying question - variable scope - approached from the opposite direction.",
    body: [
      "Every variable in GameMaker - typed or built from GML Visual blocks - is really answering two questions the moment it's created: how long should this value stick around, and where does it actually live? GameMaker has exactly two answers, and picking the wrong one produces two very different, both genuinely common, failures.",
      "An instance variable belongs to the instance itself. Create one - typically in the Create event, with a plain assignment and no special keyword - and it keeps whatever value it was last given for as long as that instance exists, visible to every event that instance ever runs afterward. A local (or \"temporary\") variable is the opposite: it belongs only to the single event or script call that creates it, and it's gone completely, with no trace, the instant that call ends. In typed GML, the difference is one keyword - var in front of a name makes it local; leaving var off makes it an instance variable (or updates an existing one). In GML Visual, it's two different actions instead of one keyword: Assign Variable creates or updates an instance variable, and Declare Temp explicitly creates a local one, scoped to just that one event.",
      { subhead: "Case 1: moving - correctly local, recomputed every Step" },
      "moving is a textbook case for local scope, and it's exactly what the code path already does: var moving = keyboard_check(vk_right) || keyboard_check(vk_left) || keyboard_check(vk_up) || keyboard_check(vk_down); declares it fresh, every single Step, with zero memory of what it was the Step before - which is exactly right, since moving is meant to be fully recomputed from the four key checks every time, never carried forward. An earlier version of this Sidebar's Visual Solution (below) used Assign Variable for this, which works - but Assign Variable's default is actually an instance variable, just one that happens to get overwritten every Step, not a true scope-for-scope match for the code path's local var. Declare Temp is the block that actually matches it: same intent, same lifetime, not just the same visible result.",
      { subhead: "Case 2: the opposite mistake - something that needs to persist, declared local by accident" },
      "The opposite case showed up for real, debugging a later lesson: a variable meant to accumulate or persist - a counter, a timer, a running total - needs the OPPOSITE choice from moving. It has to be created once, in the Create event, as a genuine instance variable, so each Step's changes actually stick. Declare it inside Step instead - or reach for Declare Temp out of habit, since \"make a new variable\" is exactly what it sounds like it does - and it gets wiped and recreated fresh every single Step, before whatever was just added to it has any chance to matter. The symptom looks exactly like \"my increment isn't working.\" It IS working, for exactly one Step, and then getting quietly erased before the next Step even starts - which is a much harder bug to spot than an error message, because nothing actually breaks. The value just never seems to move.",
      { bullet: "The question worth asking before creating any variable", text: "Does this need to remember what it was a moment ago, or is starting fresh every time actually correct? If it needs to remember: create it once, in Create, as a plain instance variable (or Set Instance Variable, for a recognized built-in). If starting fresh is correct - or even the whole point, the way it is for moving - Declare Temp (or var) is the right choice, and the more honest one: it also tells the next reader, at a glance, \"this value never needs to outlive this one event,\" which a plain instance variable never says." },
      { subhead: "The tool-divergence story, underneath the scope question" },
      "GML Visual has no equivalent block for \"wire four booleans into one OR.\" Its closest tool, If Key Down, is a branching If action, not a value you can chain - so a literal, block-by-block translation would mean building four brand new If Key Down checks solely to compute moving, right alongside the four If Key Down checks this lesson's movement/color logic already has for those exact same keys. Same condition, tested twice, in two places that now both have to agree forever.",
      "The idiomatic Visual fix skips the duplication entirely: declare moving = false once, at the very top of the Step event, then drop one more action - Assign Variable, moving = true - directly into the branch of each of the four If Key Down blocks already handling movement and color for that direction. Nothing new to check, nothing checked twice, and every block on screen is still doing one clearly visible thing.",
      "Worth naming directly: the text solution and the Visual solution aren't shaped differently because of a different design decision - nobody decided the Visual version should be more spread out. They diverge because of what each tool actually lets you build with. A single wired boolean expression is trivial to write and awkward to represent as blocks; reusing four branches that already exist is trivial visually and would be redundant if written out as text (nobody writing GML by hand would repeat the same assignment inside four separate if-blocks when one OR does it). The shape that ends up feeling \"right\" follows from the tool, not from any deeper truth about the underlying logic.",
      "This also makes sense of something GameMaker warns about directly: converting a Visual event to typed code in the IDE is a one-way trip - once converted, that event can't cleanly convert back. That stops being surprising once the asymmetry above has actually been felt firsthand. Collapsing several blocks a person already arranged by hand down into one tidy expression is close to mechanical - just read off the structure that's already there. Going the other direction has no single obvious answer: does one compact OR expression become four new If blocks? Reuse existing ones? Something else entirely? Expanding text back into blocks would require the tool to guess at a structural decision a human hasn't actually made, which is exactly why it doesn't offer to.",
      { bullet: "The mid-path \"cheat\"", text: "GML Visual's Value fields (Assign Variable's included) accept typed expressions directly - which means keyboard_check(vk_right) || keyboard_check(vk_left) || keyboard_check(vk_up) || keyboard_check(vk_down) can genuinely be pasted straight into a visual action's Value box, and it works. That's real mix-and-match between the two styles, and a legitimate, intentional GameMaker feature, not a hack. Worth being honest about what it actually is, though: a student using it isn't building an OR out of visual pieces - they're writing a line of GML text that happens to be sitting inside one field of an otherwise drag-and-drop project." },
      "This is also not the first time this course has run into a version of the same idea. Lesson 2's own Animate (blink) goal used a genuinely different technique in all eight engines - not eight different design decisions, but eight different natural tools for \"make this flicker\" (a built-in visibility flag here, a documented rendering quirk there, a side effect of existing logic somewhere else - see \"The Blink Nobody Coded the Same Way Twice,\" above). The Recolor goal hit the same pattern from the opposite direction: sprite-swapping was chosen specifically because it's the one technique that behaves identically everywhere, sidestepping the fact that color-tinting's actual result depends on what art each engine's multiply math happens to land on (see \"Changing Color,\" above). Line all four of these up - blink, color, the OR-check shape, and now scope - and the throughline is the same: the same underlying goal does not imply the same shaped solution, and the same variable name doesn't guarantee the same scope unless you actually chose it on purpose. What a tool exposes shapes the answer as much as the goal itself does - which makes it worth noticing, every time, whether a difference sitting in front of you is really a design difference, or a tool difference wearing a design decision's costume.",
      { subhead: "So: does the tool matter, if the job is the same?" },
      "Yes. Demonstrably, in this lesson alone, four separate times now.",
      { subhead: "The Visual Solution" },
      "Declare moving as a local (Declare Temp), false, once at the very start of the Step event - before any of the four movement checks run. This matches the code path's var moving = ... scope-for-scope, not just result-for-result. Then, inside each of the four existing If Key Down blocks (Right / Left / Up / Down) that already handle this direction's movement and color change, add one more action alongside what's already there - a plain Assign Variable is correct here, since moving already exists as a local for the rest of this event and this just updates it:",
      { code: ["Step event (GML Visual):", "", "Declare Temp: moving = false        <- local, recomputed fresh every step", "", "If Key Down (Right) ->", "  [this direction's existing movement + color actions]", "  Assign Variable: moving = true", "", "If Key Down (Left) ->", "  [this direction's existing movement + color actions]", "  Assign Variable: moving = true", "", "If Key Down (Up) ->", "  [this direction's existing movement + color actions]", "  Assign Variable: moving = true", "", "If Key Down (Down) ->", "  [this direction's existing movement + color actions]", "  Assign Variable: moving = true"] },
      "No new condition checks, nothing tested twice, and moving ends up true if any direction is held and false if none are - exactly the same result as the one-line GML version, and now built with the block that actually matches its scope, not just its output.",
    ],
  },
  {
    title: "Bounce Rewrites Direction - Three Different Places to Fix It",
    hookText: "Scratch's \"if on edge, bounce\" does more than reposition a sprite - it changes its Direction, which can visibly flip or spin the costume. Three genuinely different fixes exist, and each one intervenes at a different point in the problem.",
    body: [
      "\"if on edge, bounce\" (Motion) is Scratch's purpose-built edge-collision block - one block instead of four manual boundary checks, and the block most Scratch tutorials reach for first. But its actual mechanism isn't just \"nudge the sprite back onto the stage\": it reflects the sprite's Direction as if it bounced off a wall, then repositions it along that new heading. Changing Direction isn't a side effect of the block - it's literally how the block works.",
      "The visible symptom depends on the sprite's Rotation Style (see below): a costume set to \"left-right\" - the default for many sprites - visibly flips when Direction crosses past 90 or 270 degrees; a costume set to \"all around\" visibly spins to match its new heading. Either way, the sprite's actual reported Direction value really did change, on purpose, as the block's whole point.",
      { subhead: "Fix 1: Force it back" },
      "Keep \"if on edge, bounce\" exactly as-is, and immediately follow it with \"point in direction [90]\" (or whatever direction that script represents) to overwrite whatever Direction bounce just set. Bounce still does its real job - repositioning the sprite off the wall - but the very next block cancels the part of its output that wasn't wanted. This is the cheapest fix to add (one block, dropped right after bounce in each of the four key-press scripts) and it's the one that keeps \"if on edge, bounce\" itself in the project, which is worth something on its own: it's a block most Scratch tutorials reach for first, and it reads as native Scratch idiom in a way manual boundary math doesn't. The trade-off is conceptual, not practical: this fix operates on the result, deliberately overriding a value the block just set on purpose, rather than avoiding the situation that produced it.",
      { code: ["when key [right arrow] pressed", "change x by (10)", "if on edge, bounce", "point in direction (90)          <- overwrites whatever bounce just set", "[color/sound/blink actions for this direction stay unchanged]"] },
      { subhead: "Fix 2: Clamp instead of bounce" },
      "Skip \"if on edge, bounce\" entirely. Instead, check the sprite's x position and y position directly against the stage's bounds each time it moves, and use \"set x to\" (not \"change x by\") to snap it back onto the edge if it's crossed the line - the same manual-clamp idea GameMaker's clamp(), GDevelop's Stay On Screen behavior, and Unreal's screen-bounds checks all reach for elsewhere in this course. Because this never touches Direction at all, there's nothing to override and nothing to flip - the fix works by never creating the problem in the first place. The cost is trading one clean built-in block for four separate boundary checks, one per edge, which is exactly why this course's Lesson 2 materials flag it as a later-lesson upgrade rather than the default Level 3 answer: it's the more correct fix, but also the more advanced one.",
      { subhead: "Fix 3: Turn off rotation for the sprite" },
      "A third option changes neither the block nor the script - it's a one-time, per-sprite setting instead. Click directly on the Direction field in the Sprite pane (the number next to the compass icon, above the code area) to open a small popup with three icons: all around, left-right, and don't rotate. Choosing don't rotate means the sprite's Direction value still changes exactly as before every time it bounces - nothing about the underlying data is different - but the costume itself never visually reflects that change, so there's nothing to flip or spin on screen. This is the cheapest fix in terms of ongoing work (set once, applies everywhere that sprite goes, no block added to any script), but it's also the bluntest: it silences ALL visual rotation for that sprite permanently, including any place elsewhere in a project where the sprite is actually supposed to visibly turn to face its direction of travel.",
      { bullet: "Which one is \"correct\"", text: "All three are legitimate, and they intervene at three different points in the same problem: Fix 1 corrects the result after the fact, Fix 2 avoids the cause entirely, and Fix 3 hides the visual consequence without touching the underlying value at all. None of them is wrong - which one fits best depends on whether a project needs Direction to be accurate for something else later (favoring Fix 2), whether \"if on edge, bounce\" specifically is worth keeping for its own sake (favoring Fix 1), or whether the sprite never needs to visibly rotate for any reason, ever (favoring Fix 3)." },
    ],
  },
];

// =========================================================================
// LESSON 3
// =========================================================================
const lesson03 = [
  {
    title: "One Key at a Time: Why Firing Can Stop Movement Cold",
    hookText: "Hold an arrow key, then tap Space to fire - and the arrow key seems to just stop working. What's actually happening, and why isn't it a Scratch bug?",
    body: [
      "\"when key pressed\" scripts don't have their own built-in sense of \"still held down\" the way a polled check does. Per Scratch's own documentation, holding a key makes the script run again and again by riding your operating system's own keyboard auto-repeat feature - the exact same mechanism responsible for a letter repeating when you hold it down in a text box. Scratch isn't tracking the key itself; it's just reacting to whatever repeat events the OS decides to send.",
      "That's the catch: most operating systems only actively auto-repeat one key at a time. The moment a second key goes down - Space, to fire - the OS's attention (and its repeat timer) shifts to that new key, and the repeat stream for the first key, the one you're still physically holding, stops. Scratch never gets told the arrow key was released; it just stops being told it's still pressed. Movement pauses, and pressing the same arrow key again is what restarts the repeat stream and gets it moving again.",
      "Worth pointing out why this is the first lesson where it shows up: Lesson 2 only ever asked for one key held at a time - a single movement direction. This is the first lesson where a second, independently-held-and-fired key enters the picture at the same moment as movement, which is exactly the situation the OS's one-key auto-repeat limit can't handle cleanly.",
      { subhead: "The fix, if you want it: poll instead of react" },
      "Scratch's own Wiki recommends this directly for movement: instead of a hat block that reacts to key-press events, use a \"forever\" loop (started from \"when green flag clicked\") that checks the Sensing category's \"key ( ) pressed?\" reporter every frame. That reporter reports real, current key state - it isn't generated by repeat events at all, so the operating system's one-key limit never enters into it. Two keys, five keys, it doesn't matter; every held key still reports true every single frame.",
      { code: ["when green flag clicked", "forever", "  if <key [right arrow] pressed?> then", "    change x by 10", "  [same idea for left/up/down]"] },
      { bullet: "Why this isn't just the new default", text: "Rebuilding movement this way is a bigger structural change than this course's deliberately small Level 3 goals - it replaces the lesson's whole \"one when-key-pressed script per direction\" structure with a different shape entirely. The existing design is not broken for normal, one-key-at-a-time play; it just has this one specific, well-documented rough edge once a second held-and-fired key enters the picture. That's exactly why it's offered as a Super-Stretch idea rather than a Level 3 requirement - see the Super-Stretch Ideas list, and this lesson's Teacher's Key (Scratch) for the fully worked version, including why firing needs its own separate polling loop rather than sharing one with movement." },
    ],
  },
];

// =========================================================================
// LESSON 7
// =========================================================================
const lesson07 = [
  {
    title: "A Destroy Event Doesn't Know Why It Fired",
    hookText: "This lesson quietly fixed a real bug that had been sitting in two engines since Lesson 6 - and it's a genuinely common, real-world mistake, not a beginner one.",
    body: [
      "Lesson 6 needed one thing: decrement the win-counter when an enemy is defeated. In most engines, that logic naturally landed wherever the enemy's removal was already being handled - a signal, a direct function call, a collision event specifically tied to being hit.",
      "In GameMaker and 001 Game Creator specifically, the natural-looking place for that logic was a generic lifecycle event - GameMaker's Destroy event, 001 Game Creator's \"Defeated / Destroyed\" trigger. Both fire for ANY reason an instance gets removed, not specifically for being hit. That worked perfectly in Lesson 6, because being hit was the only way an enemy could ever be destroyed. The instant Lesson 7 added a second way to remove an enemy - escaping off-screen unshot - the generic event started firing for both cases, and an escaped enemy would silently still count toward the win condition. A real bug: nothing crashes, nothing looks obviously wrong, the game just becomes winnable in a way it shouldn't be.",
      "The fix, and why it generalizes past these two engines: move the win-counting logic specifically onto the ONE event that actually, specifically represents a hit - not a generic \"this object is gone now\" lifecycle event. Stated as a general rule: a generic lifecycle event answers \"did this happen,\" not \"why did this happen.\" If the why matters - and for anything tied to scoring, counting, or game state, it almost always does - that logic belongs on the specific event that captures the reason, never the generic one.",
      "Worth naming why the other six engines never had this bug at all: Godot's signal, Unity's direct call, GDevelop's collision event, TIC-80's separate loop, and Unreal's dispatcher were all already scoped specifically to \"this was a hit\" from the moment Lesson 6 was written - there was never a generic \"any destruction\" event sitting there, tempting enough to reach for by accident. That's a real, structural difference in what a given engine makes easy by default versus easy to get subtly wrong.",
    ],
  },
  {
    title: "That Yellow Warning in the Debugger: What UNUSED_PARAMETER Actually Means",
    hookText: "A warning isn't an error - the game runs exactly the same either way. Worth knowing the difference anyway, since Godot prints all three very different kinds of complaint through the same-looking little panel.",
    body: [
      "Leave one of Godot's default stubbed functions in a script - _process(delta) is the classic case, auto-filled in by the editor the moment a script is created - and if the body never actually reads delta, Godot's own checker prints exactly this in the Debugger the next time the project runs: \"The parameter 'delta' is never used in the function '_process()'. If this is intended, prefix it with an underscore: '_delta'.\" Nothing crashes, nothing behaves differently, and the game runs identically whether that warning is fixed or ignored forever - it's a heads-up, not a problem.",
      "The recommended fix is exactly what the warning says: rename the parameter from delta to _delta, everywhere it appears in that function's signature. That's a pure rename - _delta holds the exact same value, at the exact same moment, as delta did; nothing about how the function actually runs changes at all. What does change is that GDScript's own checker specifically watches for a leading underscore on a parameter name, and treats it as a deliberate signal: \"yes, I know this parameter exists and isn't used - that's on purpose, stop flagging it.\" The underscore isn't doing anything to the code itself. It's a note left for the tool that's reading the code, not for the computer that's running it.",
      { subhead: "Three Different Kinds of Rule, Worth Telling Apart" },
      "That distinction - a note for a tool, not a rule the computer enforces - is worth pulling apart properly, because Godot's Debugger genuinely mixes three very different categories of complaint into what looks like one kind of message:",
      { bullet: "A coding convention", text: "a habit programmers agree to follow because it helps people read the code later - this course's own Standards Rule 11 (meaningful variable names over generic i/j/k) is exactly this. Nothing about the language checks it, nothing enforces it, and breaking it doesn't stop the code from running - it just makes the code harder for a person to follow." },
      { bullet: "A compiler/language-enforced rule", text: "a rule GDScript itself requires. Break one - a missing colon, a variable used before it's declared, mismatched indentation - and the script genuinely will not run at all. This is the only category where the computer refuses to proceed." },
      { bullet: "A static analyzer / linter warning", text: "the middle ground, and where UNUSED_PARAMETER actually lives. The code is completely valid GDScript and will run exactly as written - but a separate tool built into Godot's editor scans it for patterns that are usually mistakes (an unused parameter often means a forgotten piece of logic) and flags them as a courtesy, without stopping anything. The underscore prefix is GDScript's own built-in way of telling that specific tool, \"I checked, this one's intentional.\"" },
      "Worth flagging one more thing, without getting into the full explanation yet: this exact same leading-underscore character shows up elsewhere in GDScript too, meaning something related but not identical - a variable or method name starting with an underscore is a separate convention, hinting that it's meant to stay internal to that one script rather than be reached from outside it. Different meaning each time the underscore shows up, same basic idea underneath both: a leading underscore is how GDScript-flavored code signals \"this is intentional, on purpose, look here for why\" to whoever - or whatever tool - reads it next.",
      "Renaming the parameter rather than deleting the function entirely is exactly the right instinct here, too - an unused _process(_delta) costs nothing to leave in place, and it's already sitting there, ready the moment there's real per-frame logic to add.",
    ],
  },
];

// =========================================================================
// LESSON 9
// =========================================================================
const lesson08 = [
  {
    "title": "Godot Signals: Create It, Connect It, Write the Function, Emit It",
    "hookText": "Signals confuse almost everyone the first time, and it is easy to see why: one idea is spread across three different scripts, and nothing breaks when you leave a piece out - the game just quietly does nothing. Here is the whole idea in one place.",
    "body": [
      "A signal is an announcement. When something happens to an enemy - it gets destroyed, it escapes off the bottom - the enemy shouts \"this just happened!\" without knowing or caring who is listening. Somewhere else, a function has asked to be told. When the shout happens, Godot runs that function. Think of a doorbell: the person pressing the button does not walk into the house and do the work. They press the button, the bell is wired to a speaker in the right room, and the speaker makes the noise. That is the whole idea; the rest is just where each piece lives.",
      {
        "diagram": {
          "file": "diagram_godot_signal_four_pieces.png",
          "w": 6.3,
          "h": 3.28,
          "caption": "Four pieces in three scripts. Pieces 1 to 3 are set up once; piece 4 is what runs every time the event really happens."
        }
      },
      {
        "subhead": "Piece 1 - CREATE the signal (in the Enemy script)"
      },
      "The Enemy's script declares that this kind of announcement exists. This line does nothing by itself - it only gives the announcement a name, the way installing a doorbell button does not ring anything. Put it at the top of the script, with any other signals.",
      {
        "code": [
          "# enemy.tscn's script:",
          "signal escaped"
        ]
      },
      {
        "subhead": "Piece 2 - CONNECT the signal (in Main, when each enemy is created)"
      },
      "Connecting is the wiring: \"when THIS enemy shouts escaped, run THAT function.\" It happens in the Main scene's spawn code, right after each new enemy is created, because every enemy is a brand-new copy and every copy needs its own wiring. The enemy does not connect itself, and GameManager does not go looking for enemies - Main, the one place that creates them, introduces them to each other. Notice there are no parentheses after on_enemy_escaped: you are handing Godot the function itself so it can run it later, not running it right now.",
      {
        "code": [
          "# Main scene's script - inside the spawn loop:",
          "# (your existing code that creates new_enemy goes here)",
          "new_enemy.escaped.connect($GameManager.on_enemy_escaped)"
        ]
      },
      {
        "subhead": "Piece 3 - WRITE the function that will run (in the GameManager script)"
      },
      "This is the function the signal will run. It has to exist, with exactly that name, before the connect line can find it. This is where the actual game logic lives - the Enemy never touches lives; it only announces that it escaped, and this function decides what that means.",
      {
        "code": [
          "# GameManager's script:",
          "func on_enemy_escaped():",
          "    lives -= 1",
          "    if lives <= 0:",
          "        get_node(\"../LoseText\").show()"
        ]
      },
      {
        "subhead": "Piece 4 - EMIT the signal (back in the Enemy script)"
      },
      "Emitting is pressing the doorbell. When the enemy actually goes past the bottom of the screen, it calls escaped.emit(). The instant that line runs, Godot looks up every function connected to escaped - here, on_enemy_escaped on GameManager - and runs it. Everything before this was preparation; this is the only line that makes something happen while the game is running.",
      {
        "code": [
          "# enemy.tscn's script:",
          "func _process(delta):",
          "    position.y += enemy_speed * delta",
          "    if position.y > 700:",
          "        escaped.emit()   # press the doorbell",
          "        queue_free()"
        ]
      },
      {
        "subhead": "One escape, traced in order"
      },
      {
        "plainBullet": "Before the game runs: the Enemy script declares signal escaped (piece 1), and GameManager has on_enemy_escaped() written (piece 3)."
      },
      {
        "plainBullet": "When Main spawns an enemy: it connects that enemy's escaped signal to GameManager's on_enemy_escaped (piece 2)."
      },
      {
        "plainBullet": "While the game runs: the enemy drifts down. When it passes the bottom it calls escaped.emit() (piece 4)."
      },
      {
        "plainBullet": "Godot sees the connection from piece 2 and runs on_enemy_escaped() on GameManager, which lowers lives. Then the enemy removes itself with queue_free()."
      },
      {
        "subhead": "When it does nothing - and no error tells you why"
      },
      {
        "bullet": "Forgot to emit",
        "text": "the signal exists and is connected, but nobody pressed the button. Lives never change. Check that escaped.emit() is really inside the branch that runs when the enemy escapes."
      },
      {
        "bullet": "Forgot to connect",
        "text": "the enemy shouts and nobody is listening. Check the spawn loop in Main - every spawned enemy needs the connect line."
      },
      {
        "bullet": "Connected twice",
        "text": "the function runs twice for one escape, so two lives disappear. Connect once, in one place."
      },
      {
        "bullet": "Typed parentheses on the function in connect",
        "text": "writing on_enemy_escaped() instead of on_enemy_escaped runs the function immediately, once, at spawn time, instead of handing it over for later."
      },
      {
        "bullet": "Spelling mismatch",
        "text": "the function name in the connect line must match the function in GameManager exactly, including capital letters; the signal name must match between piece 1, 2 and 4."
      },
      "The same four pieces appear twice in this course so far. Lesson 6 used a signal named destroyed for hits; Lesson 8 added escaped for escapes. They are two separate announcements, each with its own connection and its own function, which is exactly why a hit can never accidentally cost a life.",
      {
        "bullet": "A useful habit",
        "text": "when a signal seems not to work, check the four pieces in order: is it created, is it connected, does the function exist, is it emitted? One of the four is almost always the missing one."
      }
    ]
  },
  {
    title: "lives or __dnd_lives? Why GameMaker Has Two Answers",
    hookText: "GameMaker has had a built-in lives variable for as long as most tutorials have existed. This lesson doesn't use it. It uses __dnd_lives - with two underscores in front. Why?",
    body: [
      "Older GameMaker tutorials, and plenty of old code, use three special global variables that GameMaker provides automatically: lives, score, and health. You could write lives = 3; anywhere without declaring it, and GameMaker kept track of it for you. Those built-ins are now deprecated, which means GameMaker still understands them for old projects but the engine's own tools no longer build on them, and they may be removed someday. The extra reason that matters most for this course: they were global, so every object in the game shared the one single value.",
      "GameMaker's drag-and-drop blocks replaced them with something different: an instance variable. Set Lives, If Lives, and Draw Instance Lives all read and write a variable named __dnd_lives (and the Score blocks use __dnd_score, the Health blocks __dnd_health). The \"dnd\" stands for drag-and-drop. The two leading underscores are a naming habit programmers use to mean \"this is internal plumbing - set up by the system, not a name you chose.\" It is a real variable with exactly that name, and the double underscore is part of the name: _dnd_lives with one underscore is a different variable that GameMaker will never create for you, and nothing will warn you if you mistype it.",
      { subhead: "Instance variable, not global" },
      "Because __dnd_lives belongs to an instance, there is one copy per object that has it. That is why this course puts it on obj_gamemanager, the single dedicated manager object introduced in Lesson 6, and why every Lives block in the Visual path has a Target drop-down that must be set to obj_gamemanager. Left at its default, a Set Lives block changes the lives of whichever object the block is sitting in - here the enemy - and the player's lives never move. That one setting is the single most likely thing to go wrong in this lesson's Visual path.",
      { subhead: "Why the code path uses the same variable" },
      "It would have been possible to leave the typed code on its own made-up variable and let the blocks use theirs. That would work, but the two paths would then be unable to share a project: a student who typed lives code and later added a Draw Instance Lives block would see the block display a number that never changes, because it is reading a different variable. Using __dnd_lives in the code means typed code, drag-and-drop blocks, and a mix of both all read and write the very same number. From another object, the code reaches it as obj_gamemanager.__dnd_lives; from inside obj_gamemanager's own events, plain __dnd_lives is enough.",
      { bullet: "The practical rule", text: "any time a lesson says lives in GameMaker, it means __dnd_lives on obj_gamemanager. The same pattern repeats with score in Lesson 9 (__dnd_score)." },
      { bullet: "Why not invent a clearer name", text: "lives_remaining would read nicer, but it would be invisible to the Lives blocks. Matching the engine's own name buys a Draw Instance Lives sprite row, If Lives, and Set Lives for free." },
      { bullet: "A good habit to take from this", text: "when a tutorial's code stops working or an engine marks a feature deprecated, the fix is rarely just a rename. Ask what replaced it and what changed about how it works - here, global became per-instance - since that difference is what actually changes your design." },
    ],
  },
];

const lesson09 = [
  {
    title: "When the Fancy Tool Is the Wrong Tool",
    hookText: "This week, five engines put a score on screen with a checkbox or a couple of code lines. Unity and Unreal needed a whole separate UI framework and several new concepts. Both are completely correct.",
    body: [
      "Scratch's built-in variable watchers - literally ticking a checkbox next to a variable - are the real, production-correct way to show a value on screen in Scratch. Not a shortcut, not a placeholder. The same is true of GameMaker's draw_gui and the equivalent lightweight text-drawing tools in Godot, GDevelop, and TIC-80: for a number or two on screen, these ARE the standard tool, not a simplified stand-in for a \"real\" one.",
      "Unity and Unreal need genuinely more setup this week - Canvas/Text objects and prefabs in Unity, a full Widget Blueprint in Unreal - because they're general-purpose engines built for games far more complex than this course's capstone. Their real UI frameworks are correspondingly heavier because they're designed to also handle deeply nested menus, controller navigation, localization, and screens of every size and aspect ratio. That weight is genuinely earned on a shipping commercial game. It's also not avoidable here, because it's the one standard, correct tool those engines offer - there isn't a lighter \"just show a number\" option built in the way there is elsewhere.",
      "001 Game Creator sits in between: it has its own dedicated, documented Interface/Field system - lighter than Unity or Unreal's frameworks, but still a purpose-built system rather than a one-line checkbox, with its own real concept to learn (an Interface's Fields can bind directly to a Variable and stay live-updating via an \"Update Real-time\" toggle).",
      "The transferable lesson, worth stating directly: more setup is not the same thing as more correct. Recognizing when a lightweight built-in tool is actually the professionally right answer - rather than reaching for a bigger system just because it exists - is itself a production skill, and it applies well outside game engines too.",
      { bullet: "A bonus contrast worth raising in discussion", text: "This lesson's Unreal HUD updates itself only the instant Score or Lives actually changes (push-based - a UpdateHUD event is explicitly called from wherever the value changes). A more Update()-driven HUD instead checks every single frame whether anything changed, whether or not it did (poll-based). Both genuinely work, and this course uses both depending on the engine. Worth asking directly: for a value that changes constantly, versus one that changes rarely, which approach would you reach for, and why?" },
    ],
  },
];

// =========================================================================
// LESSON 10
// =========================================================================
const lesson10 = [
  {
    title: "The Engine Says It Can Do That - Should You Trust It?",
    hookText: "001 Game Creator actually has a built-in \"Restart Game\" action. This lesson doesn't use it. That's on purpose.",
    body: [
      "While researching this lesson, a genuine built-in Restart Game action turned up in 001 Game Creator, confirmed to exist via community documentation. At first glance it looked like it would make this entire lesson's Custom Event/reset logic unnecessary. It wasn't used, for one specific reason: nowhere in the official documentation is it clearly specified exactly what state the action resets, and what it leaves alone. For a course built specifically around verified engine behavior - a standard held to deliberately, given how often this particular engine's real UI and behavior has turned out to differ from a reasonable-sounding guess - an undocumented black box wasn't something to build a lesson, or a student's actual game, around.",
      "The choice made instead: build the reset manually and explicitly, using a Custom Event that spells out exactly what gets reset, in full view, line by line. Genuinely more setup than clicking one built-in action - but every part of it is verifiable, debuggable, and actually understood by the student who built it, rather than trusted on faith.",
      "The general principle worth naming directly: \"the engine has a button for that\" is not the same question as \"should I press it.\" A convenient-looking built-in feature is only as trustworthy as its documentation - if you can't find or confirm exactly what it does, that's a real, professional signal to consider the manual, explicit version instead, even though it costs more time up front. This isn't just a beginner-course caution; it's a genuine, everyday engineering judgment call.",
      { bullet: "A smaller companion example, same lesson", text: "GDevelop's version of \"run this logic from two different places\" doesn't use anything like a callable function at all - it just makes the SAME event group reachable through an \"Or\" condition (scene start, OR a restart flag being set). A completely different mechanism, an equally valid solution to the same underlying problem (\"one place, two ways in\") - a good reminder that \"the same concept\" doesn't have to mean \"the same implementation,\" a theme that comes up constantly across this course's 8 engines." },
    ],
  },
];

// =========================================================================
// LESSON 13
// =========================================================================
const lesson13 = [
  {
    title: "The White Flash That Wasn't (Part Two of the Color Story)",
    hookText: "Lesson 2's color sidebar explained why tinting a colored sprite comes out muddy. This lesson's hit-flash effect ran into the exact same root cause - except this time, the result wasn't muddy. It was invisible.",
    body: [
      "A color tint is a multiply operation: result = original color × tint color, channel by channel. Multiplying by WHITE specifically leaves any color completely unchanged, since multiplying by 1 (white, in normalized color terms) is a no-op. So tinting the enemy white for a hit-flash doesn't just come out muddy on a colored enemy sprite the way Lesson 2's Recolor did - it produces zero visible change at all, since the enemy sprite wasn't white to begin with.",
      "That wasn't hypothetical here: three of eight engines' Level 1 solutions were originally written to set the enemy's tint to white on hit (Godot's modulate, Unity's spriteRenderer.color, GameMaker's image_blend) - all three would have silently done nothing. Easy to miss during a quick test, too, since the enemy still gets destroyed a moment later either way; the missing flash doesn't look obviously broken, it just isn't there.",
      "The fix is the same fix as Lesson 2's Recolor: swap in a genuinely white pre-made sprite instead of tinting toward white. TIC-80's flash was never affected, because it was already built as a sprite-swap from the very start - drawing a dedicated white \"flash\" version of the enemy directly in the Sprite Editor - the same pattern that made TIC-80 the one engine that never needed Lesson 2's Recolor fix either.",
      "The broader point worth sitting with: this is exactly why \"why does this work the way it works\" matters more than \"does this look right today.\" The same underlying limitation - a multiply-tint's total dependence on the source art's own color - surfaced twice, in two different lessons, wearing two different-looking symptoms (muddy color, then no visible change at all), because it was the same rule both times. Actually understanding the rule the first time means being able to predict where else it's going to bite - which is a big part of what separates debugging by understanding from debugging by guessing.",
    ],
  },
];

// =========================================================================
// LESSON 14
// =========================================================================
const lesson14 = [
  {
    title: "Whose Number Is It, Anyway? Per-Object State vs. Shared State",
    hookText: "This lesson gave each enemy its own speed, growing a little every wave - a small design decision hiding a genuinely important, general programming question: who actually owns a given piece of data?",
    body: [
      "Score, Lives, and Wave are shared, single values - there's exactly one of each for the whole game, so they naturally live as one shared/global/scene-level variable. An enemy's speed is different: every enemy needs its own value, since (at least potentially) different enemies could end up moving at different speeds. Putting Speed on the Enemy object itself - an object variable, or a field in that enemy's own table entry, depending on engine - means each enemy genuinely carries its own copy, exactly the same way each enemy already carries its own position; nobody would think to make X/Y a single value shared by every enemy on screen, and Speed deserves the same treatment.",
      "This lesson's Key flags the mistake this prevents directly, because it's the most likely one: accidentally making Speed a single shared/scene-level value instead. That compiles fine, runs fine, and looks correct - right up until there's ever a reason for two enemies to move at different speeds, at which point it turns out there was only ever one dial for the whole game, and every enemy has silently been sharing it.",
      "The second half of the same underlying idea, from the same lesson: the enemy spawn count and the win-counter reset both need to agree on the exact same formula (4 + wave) every single week - and every engine's solution deliberately computes that formula in exactly one place and reuses the result, rather than writing it out twice in two different spots. This is the same \"single source of truth\" discipline as Lesson 11's spawn_wave() reuse, just applied to a formula instead of a whole routine: anything that must always agree with something else should be written once, not kept in sync by hand across multiple places.",
      "Why this is worth carrying forward past this one lesson: \"does this piece of data belong to one specific thing, or to the whole game\" is worth asking explicitly every time a new variable gets added, in any engine, in any language. Get it wrong and the resulting bug is often invisible until the exact moment it starts to matter - here, the first time waves are actually meant to feel different from one another.",
    ],
  },
];

// =========================================================================
// LESSON 5
// =========================================================================
const lesson05 = [
  {
    title: "How Are Colors Defined?",
    hookText: "This lesson had every engine alternate two colors on a spawned enemy - crimson red, dusty pink, c_yellow, Make Linear Color, \"200;25;55\". Different letters and numbers for what's supposedly the same idea. Worth actually taking apart what a color IS to a computer, and why \"set the color\" and \"blend the color\" turn out to be two completely different operations.",
    body: [
      "Underneath every engine's different-looking color code, there are really only two families of answer to \"what is a color,\" and this course's 8 engines lean on both. Get comfortable with both, and every engine's color API stops looking like a pile of unrelated syntax and starts looking like the same small idea, spelled differently.",
      { subhead: "RGB: How the Screen Actually Sees It" },
      "A screen doesn't paint anything - it's a grid of tiny lights, and every one of those lights is really three lights in one: a red one, a green one, and a blue one, layered on top of each other. RGB describes a color exactly the way the screen produces it: three numbers, one per light, each saying how bright that light is - off at the bottom, full brightness at the top, nothing conceptually fancier than a dimmer switch. Because it's genuinely combining LIGHT, not pigment, RGB mixing is additive: red light plus green light plus blue light, all the way up, makes white - the opposite of what happens when paints mix, and the first thing that trips people up about it.",
      "Two scales show up across this course's engines for the exact same idea. The 8-bit scale runs each channel 0-255 (256 possible brightness steps per light) - GameMaker's make_color_rgb(), a hex code like #C81937, and Unity's Color32 all use this scale. The normalized float scale instead runs each channel 0-1, where 1.0 means the same \"full brightness\" that 255 means on the other scale - Godot's Color(r, g, b), Unity's plain Color(r, g, b), and Unreal's Linear Color all use this one. They're the same information, just counted differently: 200 out of 255 and 0.8 out of 1.0 are the same brightness, roughly 78%.",
      { subhead: "HSV/HSL: How a Person Actually Thinks About It" },
      "RGB is exactly how a screen makes color, and exactly the wrong shape for how a person thinks about color. \"Make it a little darker\" or \"make it more washed-out\" doesn't map onto changing just one of R, G, or B - usually all three numbers have to change together just to keep the same hue while adjusting how dark or vivid it looks. HSV describes the same color space along three axes that actually match how people talk about color: Hue is where it sits on the color wheel (one number from 0 to 360 degrees - 0/360 is red, 120 is green, 240 is blue, and everything between is a smooth rainbow); Saturation is how vivid or washed-out it is, from a pure, intense color down to gray; Value is how bright or dark it is, from the full color down to black. (HSL is a close cousin - same Hue, same idea, but its third axis, Lightness, runs from black up through the full color and on to white, rather than stopping at the full color the way Value does.)",
      "This course's students have actually already met a real HSV interface without necessarily knowing to call it that: Scratch's Pen extension exposes \"set pen color to ()\" plus separate saturation and brightness blocks (from the same dropdown menu block) - genuine Hue, Saturation, and Value controls, not a disguised RGB picker. Worth flagging one Scratch-specific quirk directly: its Hue block runs 0-100, not the more common 0-360 - a real, documented Scratch choice, not a mistake to correct.",
      { subhead: "Named Shortcuts, Engine by Engine" },
      "Every engine also offers some kind of shortcut past typing raw numbers for common colors - but the shortcuts differ more than they first appear to, and it's worth being precise about which is which rather than assuming they're all the same trick:",
      { plainBullet: "GameMaker: a fixed set of built-in global constants - c_red, c_yellow, c_white, c_black, c_blue, and more - each just a pre-made color value with a readable name. For anything outside that list, make_color_rgb(r, g, b) takes 0-255 per channel, and make_color_hsv(h, s, v) exists too - with a real gotcha worth flagging: GameMaker's HSV hue also runs 0-255, not the standard 0-360, so a straight mental conversion from \"180 degrees\" won't land where expected." },
      { plainBullet: "Godot: two genuinely different shortcuts exist side by side. Color.YELLOW (and friends) are named constants, same idea as GameMaker's c_yellow. Separately, Color(\"yellow\") - or Color(\"#ffff00\") - is a real, documented constructor that parses a color name or a hex string directly into a Color value. Color.from_hsv(h, s, v) builds one from HSV instead of RGB." },
      { plainBullet: "Unity: named colors are static properties, not a constructor - Color.yellow, Color.red, and friends. There is no Color(\"yellow\") string constructor in Unity's C# API the way there is in Godot - a genuinely easy mix-up between the two engines, worth calling out directly since it's exactly the kind of \"looks like it should exist\" assumption this course tries to catch before it ships. Unity's actual hex-parsing tool is ColorUtility.TryParseHtmlString(\"#RRGGBB\", out color), and Color.HSVToRGB(h, s, v) converts from HSV." },
      { plainBullet: "GDevelop: no named-color shortcuts at all - its color actions take a plain \"R;G;B\" string (this lesson's Tint color = \"200;25;55\" is exactly this), always 0-255 per channel." },
      { plainBullet: "Unreal: the Details panel's color picker shows named swatches purely as an editor convenience - under the hood, a color used for a material parameter is always a Linear Color struct (0-1 floats). Blueprint builds one from scratch with Make Linear Color - not Make Color, which builds a different, byte-based FColor struct (0-255) meant for other uses, a real distinction worth double-checking rather than assuming either node is interchangeable with the other. C++ has genuine named constants (FLinearColor::Yellow, FLinearColor::Red) for a student who progresses that far." },
      { plainBullet: "TIC-80: a genuinely different system, not a shortcut on top of RGB/HSV at all. TIC-80 draws from a small indexed palette - 16 total colors, each referenced everywhere by a plain number (0-15), not a color value directly. The palette itself (editable in the Sprite Editor, or rewritable at runtime) is what defines which actual RGB color a given index currently points to." },
      { plainBullet: "001 Game Creator: a hex color picker in Graphics Options (Actor > Graphics Options > Tint) - same underlying RGB idea as everywhere else, just entered as a hex triplet through a color-swatch UI rather than typed or wired as separate numbers." },
      { subhead: "Set vs. Blend: Two Different Operations Wearing the Same Word" },
      "Every one of this lesson's engines has some property that LOOKS like it directly sets an object's color - image_blend, modulate, SpriteRenderer.color, Tint color, Set Vector Parameter Value. None of them do that. Every one of them is actually a blend (also called a tint, or a multiply): result = original color x tint color, calculated separately per channel. That's a genuinely different operation from painting over a pixel with a fresh color, and the difference matters.",
      "The technical version, worked through with real numbers: colors here are 0-1 floats per channel. A pixel that starts out pure red is (1, 0, 0). Tint it with this lesson's dusty rose/pink, (0.95, 0.4, 0.5), and the multiply runs channel by channel: R = 1 x 0.95 = 0.95, G = 0 x 0.4 = 0, B = 0 x 0.5 = 0 - the pixel comes out (0.95, 0, 0), a very slightly dimmed red, NOT the pink that was passed in. The G and B channels of the tint never had anything to work with, because the original pixel's G and B were already zero. Two edge cases worth knowing on sight: multiplying by white, (1, 1, 1), is the identity operation - 1 times anything leaves it completely unchanged, which is exactly why Lesson 13's white hit-flash silently did nothing on three engines' colored enemy sprite. Multiplying by black, (0, 0, 0), erases everything to black no matter what was there before, since anything times zero is zero.",
      "The paint-mixing intuition gets the headline result right - a blend can only pull a color toward the tint and darken it, never brighten past what was already there, the same one-way limit real paint mixing has (you can't stir two paints together and get something lighter than the lightest one you started with). But two full pots of opaque paint stirred together isn't quite the mechanism - that's a subtractive pigment process (each pigment absorbs certain wavelengths of light, and mixing pigments compounds what gets absorbed), and it tends to produce a genuinely murkier, less predictable result than a clean per-channel multiply does. The closer real-world match for what image_blend/modulate/Tint color are actually doing is a transparent glaze or wash - a thin, see-through layer of color brushed over a painting that's already dry. The original brushstrokes still show through underneath, tinted by whatever color the glaze itself is. A clear or white glaze leaves the painting essentially untouched (the same reason a white tint is a no-op); a colored glaze can only pull what's underneath toward its own hue and darken it, never add a color, or brighten one, past what was already on the canvas.",
      "Worth being honest about this lesson's own striped-row idea in that light: enemy_invader_red.png is not white or gray source art, so tinting it toward crimson or rose is still a genuine two-color multiply, not a clean color replacement - the same underlying situation as Lesson 2's muddy Recolor problem and Lesson 13's invisible white flash. It reads fine here specifically because the source art and both chosen tints are all already in the same red family - the multiply stays inside a plausible range of reds rather than drifting somewhere unexpected. Tint toward a saturated blue or green instead, and the same muddy result Lesson 2 already documented would show up again. Same rule, third lesson, third different-looking symptom - which is exactly the point: understanding the rule once means being able to predict every place it's going to bite, instead of rediscovering it lesson by lesson.",
      { bullet: "The practical takeaway", text: "if the goal is a specific, exact target color - not just \"noticeably different\" - a blend/tint is the wrong tool no matter how carefully the numbers are chosen. That calls for an actual SET: swap in a different, pre-made sprite (this course's standard fix, covered in Lesson 2's Sidebar), or, for TIC-80 specifically, repoint the palette index itself rather than blending anything. A blend is exactly the right tool for \"shift this toward a hue, however imprecisely\" - which is genuinely all this lesson's striped-row idea ever needed." },
    ],
  },
  {
    title: "Script vs. Visual Coding: Who Gets to Hold an Instance Reference?",
    hookText: "Every engine's code path can spawn something, keep a handle to that one specific copy, and act on just it - that's just a local variable holding a reference. This lesson's two advanced-track engines answer \"how do I expose that visually?\" in genuinely different ways, and the GameMaker fix above is a direct result of which answer it picked.",
    body: [
      "Both Unreal and GameMaker's visual layers need to solve the identical underlying problem this lesson's striped-row idea runs into: spawn an instance, then act on that one instance and no other, without re-finding it later. They land in different places, and the difference isn't a bug in either one - it's a real design choice about how much of \"a reference to one specific object\" a purely visual language should expose.",
      { subhead: "Unreal Blueprint: An Instance Reference Is a First-Class Wire" },
      "In C++, this is unremarkable - just a local variable holding a pointer, then a function call through it:",
      { code: [
        "AEnemy* NewEnemy = GetWorld()->SpawnActor<AEnemy>(EnemyClass, SpawnTransform);",
        "NewEnemy->SetStripeColor(Index % 2 == 0 ? DarkRed : LightRed);",
      ] },
      "NewEnemy holds a reference to the one Actor SpawnActor just created; calling SetStripeColor through it runs that function on that Actor alone. Blueprint maps this almost directly, because it treats \"a reference to one specific instance\" as a real, first-class value - not a special case. Spawn Actor from Class has a Return Value output pin, and that pin's wire genuinely IS an Object Reference, the same kind of value as any other pin, just typed to one specific Actor. Drag it straight into the Target input of the very next node (the SetStripeColor call) and the call runs on exactly the instance the wire is carrying. Need to reuse that reference further down the graph? Right-click the pin and Promote to Variable - Blueprint creates a genuine \"Enemy Object Reference\" variable that can be dragged into any Target input anywhere in the graph, exactly like the C++ pointer.",
      { diagram: { file: "diagram_unreal_blueprint_full_sequence.png", w: 6.0, h: 6.24, caption: "Spawn Actor from Class's Return Value pin, wired straight into Set Vector Parameter Value's Target - the wire itself is the instance reference, no Promote to Variable step needed for a single, immediate use." } },
      { subhead: "GameMaker GML Visual: A Scope Block Instead of a Socket" },
      "GameMaker's visual layer draws the line in a different place. Its purpose-built property blocks - Set Instance Variable, Set Instance Colour, Set Sprite, and similar - don't have an arbitrary Target input socket at all; instead they offer a fixed Action Scope dropdown in the block's own corner, and at first glance that dropdown only lists Self, Other, or a specific Object asset (which reaches every existing instance of that object, not one). There's no wire to drag a captured instance reference into, the way Blueprint's Target pin accepts one - so it looks, briefly, like GML Visual never built a way to reach one dynamically-created instance at all. It has one; it's just a block, not a socket.",
      { subhead: "GameMaker's Answer: Apply To Re-Points What \"Self\" Means" },
      "The fix Robert found by hand, entirely within the Visual editor, is an \"Instance - Apply To\" block. Create Instance's own Target field captures the newly-made instance into a local variable (new_enemy) - the same first step as before. Apply To goes right after it, and whatever needs to act on that one instance - an If/Else block, in this lesson's case - attaches to the RIGHT of Apply To, not below it, which is what nests it inside Apply To's scope instead of just sequencing after it. Apply To's own Action Scope pop-up has an Expression field at the top, separate from its Self/Other/object list - type new_enemy into it, and \"Self\" now means that one captured instance for every block nested inside, including ordinary purpose-built ones like Set Instance Colour.",
      "That makes GML Visual's answer closer to Unreal's than it first looks: both ultimately need the same value - the reference Create Instance/Spawn Actor from Class just captured - to reach the block that acts on it. Unreal carries that value on a wire, into a Target pin any node can expose. GameMaker carries it through a scope-setting block instead: Apply To's Expression field is the equivalent handoff point, just reached through a pop-up field and a rightward attachment rather than a draggable connection on the canvas. Once Apply To is in place, every block nested inside it is fully ordinary - no generic block, no typed workaround, nothing to reach for beyond the purpose-built one already being used everywhere else in the lesson.",
      { diagram: { file: "diagram_gml_visual_wireframe.png", w: 6.0, h: 4.2, caption: "Apply To as the handoff point: Create Instance's Target field captures the new instance, and that same name goes straight into Apply To's Action Scope Expression field - a real, block-only connection, just not a wire on the canvas the way Unreal's Target pin is." } },
      { bullet: "Why this is worth noticing", text: "Unreal and GameMaker solve the identical problem - reach one specific, dynamically-created instance and act on just it - with genuinely different visual mechanisms, even though both end up letting a purpose-built block do the actual work. Unreal treats an instance reference as a full citizen of its visual language, wireable anywhere a Target is expected. GameMaker reaches the same result through a dedicated scope-setting block instead of a socket on every other block - which means a purpose-built block's own targeting dropdown looking limited isn't proof that the block can't be used; it's worth checking whether a scope block like Apply To is available before assuming a generic or typed fallback is the only option. Worth checking which mechanism a given visual engine has actually built before assuming its solution should look like another engine's." },
    ],
  },
];

const jobs = [
  [buildSidebarDoc(2, "Key Press Reactor", lesson02), "Lesson_02_Sidebars.docx"],
  [buildSidebarDoc(3, "Take Your Shot", lesson03), "Lesson_03_Sidebars.docx"],
  [buildSidebarDoc(5, "Enemy Ranks", lesson05), "Lesson_05_Sidebars.docx"],
  [buildSidebarDoc(7, "On the Move", lesson07), "Lesson_07_Sidebars.docx"],
  [buildSidebarDoc(8, "Under Pressure", lesson08), "Lesson_08_Sidebars.docx"],
  [buildSidebarDoc(9, "On the Scoreboard", lesson09), "Lesson_09_Sidebars.docx"],
  [buildSidebarDoc(10, "Play It Again", lesson10), "Lesson_10_Sidebars.docx"],
  [buildSidebarDoc(13, "A Hit That Reacts", lesson13), "Lesson_13_Sidebars.docx"],
  [buildSidebarDoc(14, "Raising the Stakes", lesson14), "Lesson_14_Sidebars.docx"],
];

Promise.all(jobs.map(([doc, filename]) =>
  Packer.toBuffer(doc).then(buffer => fs.writeFileSync(filename, buffer))
)).then(() => console.log("done: " + jobs.length + " files"));
