const pptxgen = require("pptxgenjs");

const NAVY = "16305C";
const NAVY2 = "1C4A86";
const GOLD = "C99A2E";
const GOLD_L = "E8C25A";
const BG_LIGHT = "F7F9FC";
const CREAM = "F4F0E6";
const TEXT_DARK = "222222";
const TEXT_MUTE = "5A6472";

const HEAD_FONT = "Cambria";
const BODY_FONT = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5

const SW = 13.333, SH = 7.5;
const LESSON_LABEL = "Lesson 5";

function bgSlide(dark, darkImg) {
  const slide = pres.addSlide();
  slide.addImage({ path: dark ? (darkImg || "bg_title.png") : "bg_content.png", x: 0, y: 0, w: SW, h: SH });
  return slide;
}

function footer(slide, dark, pageNum) {
  slide.addText("VIDEO GAME CREATION", {
    x: 0.5, y: 7.14, w: 5, h: 0.3, fontFace: BODY_FONT, fontSize: 10,
    color: dark ? GOLD_L : TEXT_MUTE, charSpacing: 2, bold: true, margin: 0,
  });
  slide.addText(LESSON_LABEL, {
    x: SW - 2.3, y: 7.14, w: 1.5, h: 0.3, fontFace: BODY_FONT, fontSize: 10,
    color: dark ? GOLD_L : TEXT_MUTE, align: "right", margin: 0,
  });
  slide.addText(String(pageNum), {
    x: SW - 0.75, y: 7.14, w: 0.45, h: 0.3, fontFace: BODY_FONT, fontSize: 10,
    color: dark ? GOLD_L : TEXT_MUTE, align: "right", margin: 0,
  });
}

function pill(slide, text, x, y, w = 1.3, dark = true) {
  slide.addShape("roundRect", {
    x, y, w, h: 0.38, rectRadius: 0.19,
    fill: { color: dark ? "0F2C52" : CREAM }, line: { color: GOLD, width: 1 },
  });
  slide.addText(text, {
    x, y, w, h: 0.38, align: "center", valign: "middle",
    fontFace: BODY_FONT, fontSize: 12, bold: true, color: dark ? GOLD_L : NAVY2, margin: 0,
  });
}

function sectionDivider(icon, title, timeTag, teaser) {
  const s = bgSlide(true, `bg_title_${icon}.png`);
  s.addText(title, { x: 3.1, y: 2.6, w: 9.2, h: 1.0, fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, valign: "middle", margin: 0 });
  s.addText(teaser, { x: 3.15, y: 3.65, w: 9.0, h: 0.7, fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0 });
  pill(s, timeTag, 3.15, 4.35, 1.5, true);
  footer(s, true, page++);
  return s;
}

let page = 1;

// ---------- SLIDE 1: TITLE ----------
{
  const s = bgSlide(true, "bg_title_with_vgc.png");
  s.addText("S O U T H   L A K E   H O M E   E D U C A T O R S", {
    x: 0, y: 0.55, w: SW, h: 0.35, align: "center", fontFace: BODY_FONT,
    fontSize: 13, color: GOLD_L, charSpacing: 1, margin: 0,
  });
  s.addText("Lesson 5", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Programming Basics IV: Enemy Ranks", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 5", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, color: "9FB6D6", margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 5 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Explain what a loop is and why it avoids repeating code",
    "Explain why one enemy's code already scales to many",
    "Explain Scratch's clone-identity limitation and its fix",
    "Turn a single enemy into a row of independent enemies",
  ];
  objectives.forEach((t, i) => {
    const y = 2.15 + i * 0.62;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.15, y: y - 0.03, w: 6.1, h: 0.55, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 7.6, y: 1.65, w: 5.15, h: 4.6, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("What We're Using Today", { x: 7.9, y: 1.85, w: 4.6, h: 0.4, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0 });
  const materials = [
    "Demo laptop with all 8 engines verified",
    "GameMaker, for the live lecture demo",
    "Engine-specific Student Handouts, one per student",
    "No new assets this week - everything reuses Lesson 3 and 4's art and sound",
    "Your own [YourName]Game project from Lesson 4 - not a new one",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Programming - loops/iteration, code reuse and object-oriented thinking (introductory).", {
    x: 0.6, y: 6.55, w: 12.1, h: 0.5, fontFace: BODY_FONT, fontSize: 11.5, italic: true, color: TEXT_MUTE, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 3: TODAY'S AGENDA ----------
{
  const s = bgSlide(false);
  s.addText("Today's Agenda", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/production.png", x: 11.6, y: 0.4, w: 0.9, h: 0.9 });

  const rows = [
    ["0:00 - 0:05", "Welcome back, Lesson 4 recap, quick share-out"],
    ["0:05 - 0:25", "Loops: Doing Something Many Times Without Repeating Yourself"],
    ["0:25 - 0:45", "Scaling Behavior: Why One Enemy's Code Already Works for Many"],
    ["0:45 - 0:50", "A Scratch-Specific Wrinkle: Clones Have No ID"],
    ["0:50 - 1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— Enemy Ranks —"],
    ["1:00 - 1:50", "Turn your single enemy into a row of enemies"],
    ["1:50 - 2:00", "Exit check: quick share-out"],
  ];

  let y = 1.55;
  rows.forEach(([time, seg]) => {
    const isDivider = time === "LAB TIME";
    const rowH = isDivider ? 0.5 : 0.6;
    s.addShape("roundRect", {
      x: 0.6, y, w: 12.1, h: rowH, rectRadius: 0.08,
      fill: { color: isDivider ? GOLD : "FFFFFF" },
      line: { color: "DCE3EC", width: isDivider ? 0 : 1 },
    });
    s.addText(time, {
      x: 0.9, y, w: 2.3, h: rowH, valign: "middle", fontFace: BODY_FONT,
      fontSize: isDivider ? 13 : 13.5, bold: true, color: isDivider ? NAVY : NAVY2, margin: 0,
    });
    s.addText(seg, {
      x: 3.3, y, w: 9.2, h: rowH, valign: "middle", fontFace: BODY_FONT,
      fontSize: isDivider ? 13 : 12.8, bold: isDivider, italic: isDivider,
      color: isDivider ? NAVY : TEXT_DARK, margin: 0,
    });
    y += rowH + 0.11;
  });

  footer(s, false, page++);
}

// ---------- SLIDE 4: SECTION - Welcome Back ----------
sectionDivider("book", "Welcome Back", "5 min", "Lesson 4 recap and a quick share of any working combat mechanics.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 4:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Collision detection - a bullet touching an enemy destroys both, cleanly and exactly once",
    "Everyone got one working enemy that reacts correctly to being hit",
    "Today: one enemy becomes many, using a loop instead of repeating yourself",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one - and no new assets needed either, everything builds on Lesson 4's enemy and collision detection.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  s.addNotes(`"Welcome back! Last week you got a working enemy - shooting it destroyed it cleanly, even with a few misses mixed in. Quick show of hands - who got that working? Anyone try the Level 3 ideas?" [Pause, call on 1-2 volunteers]

"Today, one enemy becomes many. By the end of today, your game will have a whole row of enemies, and you'll do it without writing the same setup code five separate times."`);

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - Loops ----------
sectionDivider("programming", "Loops", "20 min", "Doing something many times without writing it out many times.");

// ---------- SLIDE 7: Loops content ----------
{
  const s = bgSlide(false);
  s.addText("Loops", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("A loop runs the same block of instructions a set number of times, instead of writing it out repeatedly by hand.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.7, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0,
  });
  s.addText("Change something about how enemies look later? Edit one loop, not five copy-pasted blocks.", {
    x: 0.6, y: 2.1, w: 11.9, h: 0.5, fontFace: BODY_FONT, fontSize: 14, italic: true, color: TEXT_MUTE, margin: 0,
  });

  s.addText("Same underlying idea, different names, across every engine:", { x: 0.6, y: 2.85, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0 });
  const rows = [
    ["Godot / Unity / GameMaker", "for (i = 0; i < 5; i++) { ... }"],
    ["GDevelop", "\"Repeat 5 times\" event"],
    ["Unreal (Blueprint)", "For Loop node"],
    ["TIC-80", "for i=0,4 do ... end"],
  ];
  let ry = 3.3;
  rows.forEach(([eng, desc]) => {
    s.addText(eng, { x: 0.6, y: ry, w: 3.2, h: 0.42, fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY, margin: 0 });
    s.addText(desc, { x: 3.9, y: ry, w: 8.6, h: 0.42, fontFace: "Courier New", fontSize: 12.5, color: TEXT_DARK, margin: 0 });
    ry += 0.44;
  });

  s.addShape("roundRect", { x: 0.6, y: 5.15, w: 12.1, h: 1.1, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 5.35, w: 0.55, h: 0.55 });
  s.addText("Discussion: where else might a loop save you from repeating yourself in a game?", {
    x: 1.6, y: 5.3, w: 10.9, h: 0.75, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, italic: true, color: TEXT_DARK, margin: 0,
  });

  s.addNotes(`"Quick question: if I wanted five enemies on screen right now, what's the laziest way?" [Pause for guesses - someone will likely say "copy and paste the code five times"] "That would actually work! But imagine changing something about how enemies look later - now you're editing five separate copies instead of one. That's exactly the problem a loop solves."

"A loop just means: run this same block of instructions a certain number of times, instead of writing it out repeatedly. You've already used a version of this without maybe noticing - your bullets loop and your movement checks run every single frame, which is itself a loop the engine drives for you. Today we write our own loop, on purpose, to create several enemies from one block of code."

[Demo cue: switch to GameMaker, open a project with a single enemy already built from Lesson 4, GML Visual tab open on obj_spawner's Create event]

"Watch: instead of placing one enemy by hand, I'm going to drag in a Repeat block set to run 5 times, and inside it, a Create Instance block that makes one more enemy at a slightly different position each pass." [Build the Repeat block and the Create Instance block live, run the game] "Five blocks, and I get five independent enemies. If I want ten instead of five, I change one number on the Repeat block, not five copy-pasted objects."

[Demo cue: flip the same event over to its Code tab] "Same event, same five-block idea - here's what it looks like typed instead." [Show the equivalent GML: a plain for loop calling instance_create_layer() five times] "Same logic either way - blocks and typed code are just two different skins on the same idea. GameMaker's the one engine in this class that genuinely lets you flip back and forth on the same event like this, which is exactly why it's demoing loops today - we'll dig into where that flexibility actually helps versus where it doesn't in today's Sidebar."

[Discussion question] "Where else might a loop save you from repeating yourself in a game?" (Examples to draw out if needed: drawing a grid of tiles, giving every enemy in a wave the same starting health, checking every item in an inventory.)

[Timing checkpoint: wrap by 0:25]`);

  footer(s, false, page++);
}

// ---------- SLIDE 7b: The Spawn Loop diagram ----------
{
  const s = bgSlide(false);
  s.addText("The Spawn Loop, Step by Step", { x: 0.6, y: 0.45, w: 11, h: 0.6, fontFace: HEAD_FONT, fontSize: 26, bold: true, color: NAVY2, margin: 0 });
  const w = 7.2, h = w * (760 / 1000);
  s.addImage({ path: "diagram_spawn_loop.png", x: (SW - w) / 2, y: 1.15, w, h });
  s.addText("Same counter, same repeat check, same creation step, same increment - whatever the exact blocks or syntax look like in your engine.", {
    x: 0.6, y: 1.15 + h + 0.15, w: 12.1, h: 0.5, align: "center", fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: TEXT_MUTE, margin: 0,
  });
  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Scaling Behavior ----------
sectionDivider("design", "Scaling Behavior", "20 min", "Why last week's one-enemy code already handles a whole row of them.");

// ---------- SLIDE 9: Scaling Behavior content ----------
{
  const s = bgSlide(false);
  s.addText("Scaling Behavior", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("Last week's hit-response code doesn't change at all - it just runs once per enemy, automatically, because every instance carries its own independent copy of it.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.85, fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.45, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 2.68, w: 0.55, h: 0.55 });
  s.addText("This is why the Structure principle from Lesson 4 mattered: keeping an object's own response to being hit inside that object's own script is exactly what lets it scale to any number of copies for free.", {
    x: 1.6, y: 2.6, w: 10.9, h: 1.2, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.15, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 4.4, w: 0.55, h: 0.55 });
  s.addText("Live demo: shoot enemies out of order - the third one, then the first, then the last. Each reacts completely independently, with nothing extra written to make that true.", {
    x: 1.6, y: 4.3, w: 10.9, h: 1.35, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, color: CREAM, margin: 0,
  });

  s.addNotes(`"Here's the part that might surprise you. Last week, you wrote the code for what happens when ONE enemy gets hit. Today, without changing a single line of that code, it's going to correctly handle FIVE enemies. Why does that work?" [Pause for guesses]

"Remember the Structure idea from Lesson 4 - that an object's own script should own its own response to being hit? Because you kept that boundary clean, every single copy the loop creates carries its own independent copy of that same behavior. Enemy #3 getting hit has absolutely nothing to do with Enemy #1's code - they're each running their own complete, separate copy."

[Demo cue: same GameMaker project] "Watch what happens if I shoot enemies out of order - the third one, then the first, then the last." [Fire at enemies in a mixed order] "Each one reacts completely independently. Nothing about hitting one affects the others, and I didn't have to write anything special to make that true - it's just a direct result of keeping behavior with the object it belongs to."

"This is a big deal, and it's going to keep paying off all semester. Good structure now means big new features later - like a whole wave of enemies - come almost for free."

[Discussion question] "Can anyone think of a moment in a game where LOTS of things need to react to a similar event independently?" (Examples: every brick in a brick-breaker, every zombie in a horde game, every light in a level flickering on its own.)

[Timing checkpoint: wrap by 0:43-0:45]`);

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - Scratch Wrinkle ----------
sectionDivider("controller", "A Scratch-Specific Wrinkle", "5 min", "Clones have no ID - so this week the hit-check moves to the enemy.");

// ---------- SLIDE 11: Scratch Wrinkle content ----------
{
  const s = bgSlide(false);
  s.addText("Clones Have No ID", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Unlike instances in every other engine here, Scratch clones have no unique ID number to check against.", { x: 0.6, y: 1.5, w: 11.9, h: 0.5, fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0 });

  const cards = [
    ["Other engines", "A Bullet can ask \"which specific Enemy am I touching?\" and act on just that one."],
    ["Scratch", "A Bullet clone can only ask \"am I touching any Enemy?\" - not which one."],
  ];
  let cx = 0.6;
  cards.forEach(([title, desc]) => {
    s.addShape("roundRect", { x: cx, y: 2.2, w: 5.85, h: 1.9, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
    s.addText(title, { x: cx + 0.3, y: 2.4, w: 5.25, h: 0.5, fontFace: BODY_FONT, fontSize: 17, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: cx + 0.3, y: 2.95, w: 5.25, h: 1.0, fontFace: BODY_FONT, fontSize: 13, color: TEXT_DARK, margin: 0 });
    cx += 6.25;
  });

  s.addShape("roundRect", { x: 0.6, y: 4.35, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 4.58, w: 0.55, h: 0.55 });
  s.addText("Fix: move the check onto each Enemy clone's own script - each one asks \"am I touching a Bullet?\" about itself, which works correctly no matter how many enemies exist. This is a real change from last week's code, not just an addition.", {
    x: 1.6, y: 4.5, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 13, color: CREAM, margin: 0,
  });

  s.addNotes(`"One quick heads-up, mainly for anyone on Scratch today. Scratch's clones are a little different from instances in every other engine we use - they don't have a unique ID number you can check. That means a Bullet clone literally cannot ask 'which specific Enemy clone am I touching?' - it can only ask 'am I touching any Enemy at all?'"

"So this week, Scratch students are going to move the hit-check from the Bullet's script onto the Enemy's own script instead - each Enemy clone asks 'am I touching a Bullet?' about itself, which works correctly no matter how many enemies exist. It's a real change from last week's code, not just an extension of it - your handout walks through exactly what moves where."

[Timing checkpoint: wrap by 0:50]`);

  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "10 min", "Enemy Ranks - turning one enemy into a whole row of them.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("Enemy Ranks", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("Several enemies appear in a row, created by a loop. Shooting any one destroys just that one - built on top of your existing project.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "A row of enemies appears via a loop; any one can be destroyed.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "Every enemy is independently destructible with no errors.", "Everything works end to end."],
    ["Level 3 - Refined", "Sound already works for free - add row-to-row visual variety.", "Ready to share."],
  ];
  let lx = 0.6;
  levels.forEach(([title, desc, tag]) => {
    s.addShape("roundRect", { x: lx, y: 2.3, w: 3.9, h: 2.1, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(title, { x: lx + 0.2, y: 2.45, w: 3.5, h: 0.4, fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: lx + 0.2, y: 2.9, w: 3.5, h: 1.0, fontFace: BODY_FONT, fontSize: 12, color: TEXT_DARK, margin: 0 });
    s.addText(tag, { x: lx + 0.2, y: 3.9, w: 3.5, h: 0.4, fontFace: BODY_FONT, fontSize: 11.5, italic: true, color: GOLD, margin: 0 });
    lx += 4.1;
  });

  s.addShape("roundRect", { x: 0.6, y: 4.7, w: 12.1, h: 1.55, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 4.92, w: 0.55, h: 0.55 });
  s.addText("Reminder: delete the single enemy you manually placed last week before testing your new loop-spawned row, or you'll end up with one extra enemy that doesn't belong.", {
    x: 1.6, y: 4.82, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, color: CREAM, margin: 0,
  });

  s.addNotes(`"For lab today, open your own project from last week - not a new one, and no new assets needed either. Everything today builds directly on your enemy and collision detection from Lesson 4."

"The goal: a row of enemies instead of one, created by a loop. Level 1 is getting that working once. Level 2 is confirming every enemy in the row can be independently destroyed with no errors. Level 3 is mostly free wins this week - your sound already works, so it's really about adding a bit of visual variety across the row."

[Hand out engine-specific handouts]

"One reminder specific to today: most of you will need to delete the single enemy you manually placed last week before your new loop-spawned row will look right - it's easy to forget and end up with six enemies instead of five."

"Questions before we start lab? ... Let's build an army - well, a small row of them."`);

  footer(s, false, page++);
}

// ---------- SLIDE 14: Lab Time Notes ----------
{
  const s = bgSlide(false);
  s.addText("Lab Time Notes", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/wrench.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const notes = [
    "Scratch needs real hands-on support this week - the touching-check has to move from Bullet to Enemy, a genuine restructure. Expect this cluster to take longer.",
    "Most common bug elsewhere: forgetting to delete last week's manually-placed single enemy, leaving one extra enemy that doesn't belong to the row.",
    "Second most common: all spawned enemies stacking on top of each other - almost always means the loop's index variable isn't used in the position math.",
    "Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout - not given new unrelated work.",
  ];
  notes.forEach((t, i) => {
    const y = 1.75 + i * 1.15;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.9, fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addNotes(`Scratch needs real hands-on support this week - the touching-check has to move from Bullet to Enemy, which is a genuine restructure, not just an addition. Expect this engine's cluster to take longer than usual.

The single most common bug in every other engine: forgetting to delete last week's manually-placed single enemy, ending up with one extra enemy that doesn't belong to the row.

Second most common: all spawned enemies stacking on top of each other - almost always means the loop's index variable isn't actually being used in the position calculation.

Circulate by engine cluster, starting with Scratch this week given the extra complexity there.

Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout, not given new unrelated work.

At 1:50, call time and do a 2-minute share-out: ask 2-3 students to show what they built, across different engines if possible.`);

  footer(s, false, page++);
}

// ---------- SLIDE 15: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("Great Work Today!", { x: 0, y: 2.75, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 38, bold: true, color: GOLD_L, margin: 0 });
  s.addText("Next week: a win condition - clearing all enemies completes the Level 1 MVP milestone.", {
    x: 1.5, y: 3.6, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 17, color: CREAM, margin: 0,
  });
  s.addText("Keep polishing your enemy row at home if you'd like - try one of the Level 3 ideas if you haven't already.", {
    x: 1.5, y: 4.3, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_05_Slides.pptx" }).then(() => console.log("done"));
