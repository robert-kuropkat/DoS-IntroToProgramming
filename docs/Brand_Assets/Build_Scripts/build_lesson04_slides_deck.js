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
const LESSON_LABEL = "Lesson 4";

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
  s.addText("Lesson 4", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Programming Basics III: Direct Hit", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 4", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, color: "9FB6D6", margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 4 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Explain what collision/overlap detection is",
    "Explain trigger/area checks vs. solid physical collision",
    "Explain why destroying two objects on one hit is safe",
    "Add a working enemy + collision mechanic to their project",
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
    "This lesson's Assets folder (new enemy sprite + hit sound - no download needed)",
    "Your own [YourName]Game project from Lesson 3 - not a new one",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Programming - collision/physics systems, event-driven programming (introductory).", {
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
    ["0:00 - 0:05", "Welcome back, Lesson 3 recap, quick share-out"],
    ["0:05 - 0:25", "Collision Detection: How Engines Know Things Touched (live demo)"],
    ["0:25 - 0:45", "Cleanup Revisited: Destroying Two Things at Once, Safely"],
    ["0:45 - 0:50", "Placement: Things You Spawn vs. Things You Place"],
    ["0:50 - 1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— Direct Hit —"],
    ["1:00 - 1:50", "Add an enemy and collision detection to your Semester 1 project"],
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
sectionDivider("book", "Welcome Back", "5 min", "Lesson 3 recap and a quick share of any working shooting mechanics.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 3:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Instances and cleanup - spawning a bullet that flies up and disappears",
    "Everyone got a working fire key with a cooldown-capped shooting mechanic",
    "Today: adding an enemy - your bullets finally have something to hit",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one - everything we build adds on top of the character and bullet you already have.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - Collision Detection ----------
sectionDivider("shield", "Collision Detection", "20 min", "How does an engine even know two things touched?");

// ---------- SLIDE 7: Collision Detection content ----------
{
  const s = bgSlide(false);
  s.addText("Collision Detection", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("Give each object a simple collision shape (a box or circle, not exact pixel art). Every frame, the engine checks if any two shapes overlap - and fires a signal when they do.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.85, fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0,
  });

  s.addText("Same underlying question, different names, across every engine:", { x: 0.6, y: 2.45, w: 9, h: 0.4, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0 });
  const rows = [
    ["Godot", "Area2D + area_entered signal"],
    ["GameMaker", "Collision event"],
    ["Unity", "Collider2D (Is Trigger) + OnTriggerEnter2D()"],
    ["Unreal", "Event ActorBeginOverlap"],
  ];
  let ry = 2.9;
  rows.forEach(([eng, desc]) => {
    s.addText(eng, { x: 0.6, y: ry, w: 2.2, h: 0.42, fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY, margin: 0 });
    s.addText(desc, { x: 2.9, y: ry, w: 9.6, h: 0.42, fontFace: "Courier New", fontSize: 12.5, color: TEXT_DARK, margin: 0 });
    ry += 0.44;
  });

  s.addShape("roundRect", { x: 0.6, y: 4.85, w: 12.1, h: 1.4, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 5.08, w: 0.55, h: 0.55 });
  s.addText("Discussion: where else in a game does this same idea show up? (Coins, doorways, an attack landing, falling into a hazard.)", {
    x: 1.6, y: 5.0, w: 10.9, h: 1.1, valign: "middle", fontFace: BODY_FONT, fontSize: 14, italic: true, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Cleanup Revisited ----------
sectionDivider("wrench", "Cleanup Revisited", "20 min", "One collision, two things to destroy - and why that's always safe.");

// ---------- SLIDE 9: Cleanup Revisited content ----------
{
  const s = bgSlide(false);
  s.addText("Cleanup Revisited", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("One collision now needs to clean up TWO objects: the bullet that hit, and the enemy that got hit.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0,
  });
  s.addText("In every engine we use, calling \"destroy this\" a second time on something already destroyed is completely harmless - the engine just ignores the repeat.", {
    x: 0.6, y: 2.1, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 14, italic: true, color: TEXT_MUTE, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.85, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 3.08, w: 0.55, h: 0.55 });
  s.addText("Design principle: it's the Enemy's own script that reacts to being hit - not the Bullet's. Whoever gets hit is the one responsible for what happens next.", {
    x: 1.6, y: 3.0, w: 10.9, h: 1.2, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.55, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 4.8, w: 0.55, h: 0.55 });
  s.addText("Live demo: fire several bullets quickly, including some that miss. Hits destroy the enemy exactly once - misses keep flying and vanish off-screen exactly like last week.", {
    x: 1.6, y: 4.7, w: 10.9, h: 1.35, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - Placement ----------
sectionDivider("design", "Placement", "5 min", "Things you spawn while playing vs. things you place ahead of time.");

// ---------- SLIDE 11: Placement content ----------
{
  const s = bgSlide(false);
  s.addText("Placement", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Both are still instances - copies of a template. They just come into existence at different moments:", { x: 0.6, y: 1.5, w: 11.9, h: 0.5, fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0 });

  const cards = [
    ["Spawned at runtime", "Doesn't exist until the game creates it while playing.", "Last week's bullet - created on each fire-key press."],
    ["Placed at design time", "Sitting in the scene/room/level from the very start.", "Today's enemy - placed once, ahead of time."],
  ];
  let cx = 0.6;
  cards.forEach(([title, desc, ex]) => {
    s.addShape("roundRect", { x: cx, y: 2.2, w: 5.85, h: 2.4, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
    s.addText(title, { x: cx + 0.3, y: 2.4, w: 5.25, h: 0.5, fontFace: BODY_FONT, fontSize: 16.5, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: cx + 0.3, y: 2.95, w: 5.25, h: 0.85, fontFace: BODY_FONT, fontSize: 13, color: TEXT_DARK, margin: 0 });
    s.addText(ex, { x: cx + 0.3, y: 3.85, w: 5.25, h: 0.6, fontFace: BODY_FONT, fontSize: 12, italic: true, color: GOLD, margin: 0 });
    cx += 6.25;
  });

  s.addText("You'll see this split all semester: bullets, effects, and pickups get spawned; level layout, a boss, and fixed obstacles get placed.", {
    x: 0.6, y: 4.95, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 14, italic: true, color: TEXT_MUTE, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "10 min", "Direct Hit - the first piece of your capstone's actual combat loop.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("Direct Hit", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("A single enemy sits on screen. Shooting it destroys both the enemy and the bullet that hit it - added on top of your existing project.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "The enemy and the bullet that hits it both disappear on contact.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "Destroyed cleanly and exactly once - no errors, misses unaffected.", "Everything works end to end."],
    ["Level 3 - Refined", "Sound, a distinct look, and/or clean hit-response structure.", "Ready to share."],
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
  s.addText("Heads up: this is the first week needing a real collision/physics component - a Rigidbody2D in Unity, an Area2D shape in Godot, overlap events in Unreal. Check that first if nothing happens on contact.", {
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
    "Most common bug this week: \"nothing happens when I shoot the enemy\" - almost always a missing or misconfigured collision component. Check that first.",
    "Second most common: the bullet destroys things it shouldn't, or the enemy never gets destroyed - usually a missing or mismatched tag/group check.",
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
  s.addText("Next week: multiple enemies at once - loops that spawn and manage a whole row of them.", {
    x: 1.5, y: 3.6, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 17, color: CREAM, margin: 0,
  });
  s.addText("Keep polishing your combat mechanic at home if you'd like - try one of the Level 3 ideas if you haven't already.", {
    x: 1.5, y: 4.3, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_04_Slides.pptx" }).then(() => console.log("done"));
