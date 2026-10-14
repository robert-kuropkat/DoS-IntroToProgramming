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
const LESSON_LABEL = "Lesson 3";

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
  s.addText("Lesson 3", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Programming Basics II: Take Your Shot", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 3", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, color: "9FB6D6", margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 3 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Explain the template/instance relationship",
    "Explain why spawned objects must be cleaned up",
    "Explain \"just pressed\" vs. \"currently held\" input checks",
    "Add a working shoot mechanic to their Semester 1 project",
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
    "Godot, for the live lecture demo",
    "Engine-specific Student Handouts, one per student",
    "This lesson's Assets folder (new projectile sprite + sound - no download needed)",
    "Your own [YourName]Game project from Lesson 2 - not a new one",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Programming - object instancing, resource/memory management, input handling (introductory).", {
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
    ["0:00 - 0:05", "Welcome back, Lesson 2 recap, quick share-out"],
    ["0:05 - 0:25", "Instances: One Template, Many Copies (live demo in Godot)"],
    ["0:25 - 0:45", "Cleanup: Why You Have to Remove What You Create"],
    ["0:45 - 0:50", "One Press, One Shot (edge-triggered vs. held-down input)"],
    ["0:50 - 1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— Take Your Shot —"],
    ["1:00 - 1:50", "Add shooting to your ongoing Semester 1 project"],
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
sectionDivider("book", "Welcome Back", "5 min", "Lesson 2 recap and a quick share of any finished Key Press Reactors.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 2:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Sequences, events, and the game loop - and the start of your Semester 1 project",
    "Everyone built a Key Press Reactor moving a character in up to 4 directions",
    "Today: adding shooting - the first piece of your capstone's actual mechanic",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one - everything we build adds on top of the character you already have.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - Instances ----------
sectionDivider("programming", "Instances: One Template, Many Copies", "20 min", "How pressing fire five times gives you five separate, independent bullets.");

// ---------- SLIDE 7: Instances content ----------
{
  const s = bgSlide(false);
  s.addText("Instances", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("Build ONE definition of a bullet. Every press stamps out a brand-new, independent COPY of it - that's an instance.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.7, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0,
  });
  s.addText("Change one bullet's position, and none of the others are affected - they're all separate, all remembering their own data.", {
    x: 0.6, y: 2.1, w: 11.9, h: 0.5, fontFace: BODY_FONT, fontSize: 14, italic: true, color: TEXT_MUTE, margin: 0,
  });

  s.addText("Same idea, nearly the same word, across every engine:", { x: 0.6, y: 2.85, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0 });
  const rows = [
    ["Godot", "bullet_scene.instantiate()"],
    ["GameMaker", "instance_create_layer()"],
    ["Unity", "Instantiate()"],
    ["Unreal", "Spawn Actor from Class"],
  ];
  let ry = 3.3;
  rows.forEach(([eng, desc]) => {
    s.addText(eng, { x: 0.6, y: ry, w: 2.2, h: 0.42, fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY, margin: 0 });
    s.addText(desc, { x: 2.9, y: ry, w: 9.6, h: 0.42, fontFace: "Courier New", fontSize: 13, color: TEXT_DARK, margin: 0 });
    ry += 0.44;
  });

  s.addShape("roundRect", { x: 0.6, y: 5.2, w: 12.1, h: 1.05, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 5.4, w: 0.55, h: 0.55 });
  s.addText("Discussion: where else in a game might you want lots of independent copies of one thing?", {
    x: 1.6, y: 5.35, w: 10.9, h: 0.75, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, italic: true, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Cleanup ----------
sectionDivider("wrench", "Cleanup: Why You Remove What You Create", "20 min", "A hundred forgotten bullets doesn't sound like much - until it does.");

// ---------- SLIDE 9: Cleanup content ----------
{
  const s = bgSlide(false);
  s.addText("Cleanup", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("If nothing removes old bullets, the engine keeps tracking every single one - forever - even off-screen and useless.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.7, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0,
  });
  s.addText("A real game spawning constantly (bullets, particles, enemies) can pile up thousands of forgotten objects in minutes - the game slows down or crashes.", {
    x: 0.6, y: 2.1, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 14, italic: true, color: TEXT_MUTE, margin: 0,
  });

  s.addText("Every engine gives you an explicit \"I'm done with this one\" call:", { x: 0.6, y: 2.85, w: 9, h: 0.4, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0 });
  const rows = [
    ["Godot", "queue_free()"],
    ["GameMaker", "instance_destroy()"],
    ["Unity / Unreal", "Destroy()"],
  ];
  let ry = 3.3;
  rows.forEach(([eng, desc]) => {
    s.addText(eng, { x: 0.6, y: ry, w: 2.6, h: 0.42, fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY, margin: 0 });
    s.addText(desc, { x: 3.3, y: ry, w: 9.2, h: 0.42, fontFace: "Courier New", fontSize: 13, color: TEXT_DARK, margin: 0 });
    ry += 0.44;
  });

  s.addShape("roundRect", { x: 0.6, y: 4.85, w: 12.1, h: 1.4, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/book.png", x: 0.85, y: 5.08, w: 0.55, h: 0.55 });
  s.addText("The pattern for anything spawned at runtime is always two-sided: something that creates it, and something that eventually destroys it. Today, a bullet leaving the screen triggers that second half.", {
    x: 1.6, y: 5.0, w: 10.9, h: 1.1, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - One Press, One Shot ----------
sectionDivider("controller", "One Press, One Shot", "5 min", "\"Currently held\" vs. \"just pressed\" - the difference that keeps fire from flooding the screen.");

// ---------- SLIDE 11: One Press, One Shot content ----------
{
  const s = bgSlide(false);
  s.addText("One Press, One Shot", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Two different questions you can ask about a key:", { x: 0.6, y: 1.45, w: 11.9, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });

  const cards = [
    ["Currently held?", "True for as long as the key is down - checked fresh every frame. Right for movement: hold, keep moving."],
    ["Just pressed?", "True for exactly one frame, no matter how long you hold it after. Right for firing: one press, one shot."],
  ];
  let cx = 0.6;
  cards.forEach(([title, desc]) => {
    s.addShape("roundRect", { x: cx, y: 2.05, w: 5.85, h: 2.0, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
    s.addText(title, { x: cx + 0.3, y: 2.25, w: 5.25, h: 0.5, fontFace: BODY_FONT, fontSize: 17, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: cx + 0.3, y: 2.8, w: 5.25, h: 1.1, fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, margin: 0 });
    cx += 6.25;
  });

  s.addText("Same split, different names, in nearly every engine:", { x: 0.6, y: 4.35, w: 9, h: 0.4, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0 });
  const rows = [
    ["Godot", "is_action_just_pressed()   vs.   is_action_pressed()"],
    ["GameMaker", "keyboard_check_pressed()   vs.   keyboard_check()"],
    ["Unity", "wasPressedThisFrame   vs.   isPressed"],
  ];
  let ry = 4.8;
  rows.forEach(([eng, desc]) => {
    s.addText(eng, { x: 0.6, y: ry, w: 2.2, h: 0.42, fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY, margin: 0 });
    s.addText(desc, { x: 2.9, y: ry, w: 9.6, h: 0.42, fontFace: "Courier New", fontSize: 12, color: TEXT_DARK, margin: 0 });
    ry += 0.44;
  });

  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "10 min", "Take Your Shot - the first piece of your capstone's actual shooting mechanic.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("Take Your Shot", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("Pressing your fire key spawns a projectile that flies upward and disappears off-screen - added on top of your existing project.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "One press spawns one projectile that flies up and disappears.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "Repeatable, multiple projectiles, capped by a fire-rate cooldown.", "Everything works end to end."],
    ["Level 3 - Refined", "Sound, a distinct look, and/or clean code structure - pick what interests you.", "Ready to share."],
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
  s.addText("Reminder: if your Lesson 2 character still moves in all 4 directions, that's fine to keep - or trim it to Left/Right to match the actual capstone game exactly. Either way works.", {
    x: 1.6, y: 4.82, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 14: Lab Time Notes ----------
{
  const s = bgSlide(false);
  s.addText("Lab Time Notes", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/wrench.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const notes = [
    "Most common bug this week: a held-down fire key spawning dozens of bullets per second - check for a \"held\" input check where a \"just pressed\" one belongs.",
    "Second most common: bullets that spawn but never move or never disappear - usually the movement/cleanup code is in the wrong event.",
    "Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout - not given new unrelated work.",
    "At 1:50, call time and do a 2-minute share-out: ask 2-3 students to show what they built, across different engines if possible.",
  ];
  notes.forEach((t, i) => {
    const y = 1.75 + i * 1.15;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.9, fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  footer(s, false, page++);
}

// ---------- SLIDE 15: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("Great Work Today!", { x: 0, y: 2.75, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 38, bold: true, color: GOLD_L, margin: 0 });
  s.addText("Next week: a single enemy on screen, and collision detection - your projectile finally has something to hit.", {
    x: 1.5, y: 3.6, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 17, color: CREAM, margin: 0,
  });
  s.addText("Keep polishing your shooting mechanic at home if you'd like - try one of the Level 3 ideas if you haven't already.", {
    x: 1.5, y: 4.3, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_03_Slides.pptx" }).then(() => console.log("done"));
