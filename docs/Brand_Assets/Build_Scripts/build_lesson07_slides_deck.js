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
const LESSON_LABEL = "Lesson 7";

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
  s.addText("Lesson 7", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Core Mechanics I: On the Move", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 7  -  Core Mechanics Unit Begins", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, bold: true, color: GOLD_L, margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 7 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Implement continuous, frame-rate-independent movement",
    "Distinguish \"destroyed by a hit\" from \"destroyed by escaping\"",
    "Explain why generic \"on destroyed\" hooks can silently misfire",
    "Relocate state-changing logic to the specific event it belongs to",
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
    "GameMaker, for the bug-discussion contrast",
    "No new assets this week - logic-only lesson",
    "Your own [YourName]Game project from Lesson 6 - not a new one",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Programming - game loops and continuous simulation, event-driven programming, debugging and code correctness (introductory).", {
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
    ["0:00 - 0:05", "Welcome back, Lesson 6 recap, milestone check-in"],
    ["0:05 - 0:20", "New Unit: Core Mechanics - what's ahead through Lesson 11"],
    ["0:20 - 0:40", "Movement That Doesn't Stop: Continuous Motion"],
    ["0:40 - 1:00", "The Hidden Bug: When \"Destroyed\" Doesn't Mean \"Defeated\""],
    ["1:00 - 1:05", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— On the Move —"],
    ["1:05 - 1:55", "Advance enemies toward the player + escape cleanup"],
    ["1:55 - 2:00", "Exit check / quick share-out"],
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
sectionDivider("book", "Welcome Back", "5 min", "Lesson 6 recap - you hit the Level 1 MVP milestone.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 6:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "A working win condition - the final piece of the Level 1 MVP milestone",
    "A complete, playable game loop: move, shoot, face enemies, win",
    "Today: enemies stop standing still - the first step into Level 2",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one - and no new assets needed either, this is a logic-only lesson.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - New Unit ----------
sectionDivider("design", "New Unit: Core Mechanics", "15 min", "Collision, scoring, win/lose states, simple UI - Lessons 7 through 11.");

// ---------- SLIDE 7: New Unit content ----------
{
  const s = bgSlide(false);
  s.addText("Core Mechanics", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "15 min", 11.4, 0.5, 1.3, false);

  s.addText("You just finished \"Programming Basics\" - variables, loops, collision, state tracking. Starting today, a new unit builds toward a complete, replayable arcade shooter.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.85, fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0,
  });

  const rows = [
    ["Lesson 7", "Enemies advance toward the player (today)"],
    ["Lesson 8", "Lose condition + lives system"],
    ["Lesson 9", "Score tracking + basic on-screen UI"],
    ["Lesson 10", "Game-over and restart flow"],
    ["Lesson 11", "Multiple waves - Level 2 Complete milestone"],
  ];
  let ry = 2.55;
  rows.forEach(([lesson, desc], i) => {
    const highlight = i === 0;
    s.addShape("roundRect", { x: 0.6, y: ry, w: 12.1, h: 0.62, rectRadius: 0.08, fill: { color: highlight ? "0F2C52" : "EDF2F8" }, line: { type: "none" } });
    s.addText(lesson, { x: 0.9, y: ry, w: 2.0, h: 0.62, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, bold: true, color: highlight ? GOLD_L : NAVY2, margin: 0 });
    s.addText(desc, { x: 3.0, y: ry, w: 9.4, h: 0.62, valign: "middle", fontFace: BODY_FONT, fontSize: 13, color: highlight ? CREAM : TEXT_DARK, margin: 0 });
    ry += 0.72;
  });

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Movement That Doesn't Stop ----------
sectionDivider("programming", "Movement That Doesn't Stop", "20 min", "Enemies that move every frame, with nothing triggering it.");

// ---------- SLIDE 9: Movement content ----------
{
  const s = bgSlide(false);
  s.addText("Continuous Motion", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("So far, everything has moved in response to something - a key press, a spawn. Today's movement just keeps happening, every single frame, forever.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.75, fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.4, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/programming.png", x: 0.85, y: 2.65, w: 0.55, h: 0.55 });
  s.addText("Multiply movement by delta (time since the last frame) to keep speed the same no matter how fast the game runs - the same pattern player movement has used since Lesson 2.", {
    x: 1.6, y: 2.55, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addText("Every enemy also gets a bottom-of-screen check: reach the bottom unshot, and it's removed - but does NOT count as a kill.", {
    x: 0.6, y: 4.3, w: 11.9, h: 0.55, fontFace: BODY_FONT, fontSize: 14.5, color: NAVY2, bold: true, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 5.0, w: 12.1, h: 1.25, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 5.2, w: 0.55, h: 0.55 });
  s.addText("Discussion: what else in a game has to run continuously rather than in response to an event? (A countdown timer, a camera following the player, gravity.)", {
    x: 1.6, y: 5.1, w: 10.9, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - The Hidden Bug ----------
sectionDivider("wrench", "The Hidden Bug", "20 min", "When \"destroyed\" doesn't mean \"defeated.\"");

// ---------- SLIDE 11: Hidden Bug - the problem ----------
{
  const s = bgSlide(false);
  s.addText("Two Ways to Disappear", { x: 0.6, y: 0.45, w: 10.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  pill(s, "part 1", 11.4, 0.5, 1.3, false);

  s.addText("Now that enemies move, there are TWO reasons one might stop existing - and only one of them should count as a win.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.7, fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0,
  });

  const cols = [
    ["Shot", "Destroyed by a bullet - this IS a kill. Should count toward the win condition.", "icons/check.png"],
    ["Escaped", "Reached the bottom of the screen unshot - NOT a kill. Must never count.", "icons/check.png"],
  ];
  let lx = 0.6;
  cols.forEach(([title, desc]) => {
    s.addShape("roundRect", { x: lx, y: 2.4, w: 5.9, h: 1.9, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(title, { x: lx + 0.25, y: 2.6, w: 5.4, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: lx + 0.25, y: 3.05, w: 5.4, h: 1.15, fontFace: BODY_FONT, fontSize: 13, color: TEXT_DARK, margin: 0 });
    lx += 6.2;
  });

  s.addShape("roundRect", { x: 0.6, y: 4.6, w: 12.1, h: 1.65, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 4.85, w: 0.6, h: 0.6 });
  s.addText("GameMaker's Destroy event and 001 Game Creator's \"Defeated / Destroyed\" trigger both fire for EITHER reason. Last week's win-counter wiring was correct then - only one destruction path existed. This week it silently breaks.", {
    x: 1.65, y: 4.75, w: 10.8, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 13, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 11b: The Activity Diagram ----------
{
  const s = bgSlide(false);
  s.addText("Two Ways to Disappear, Side by Side", { x: 0.6, y: 0.45, w: 11, h: 0.6, fontFace: HEAD_FONT, fontSize: 26, bold: true, color: NAVY2, margin: 0 });
  const w = 6.7, h = w * (780 / 1000);
  s.addImage({ path: "diagram_activity_enemy_removal.png", x: (SW - w) / 2, y: 1.0, w, h });
  s.addText("The same shape in every engine - only the hit check is allowed to touch enemies_remaining.", {
    x: 0.6, y: 1.0 + h + 0.1, w: 12.1, h: 0.5, align: "center", fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: TEXT_MUTE, margin: 0,
  });
  footer(s, false, page++);
}

// ---------- SLIDE 12: Hidden Bug - the fix ----------
{
  const s = bgSlide(false);
  s.addText("The Fix", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "part 2", 11.4, 0.5, 1.3, false);

  s.addText("Attach win-counting to the SPECIFIC event that means \"got shot,\" not the GENERIC event that means \"stopped existing.\"", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.7, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0,
  });

  const rows = [
    ["GameMaker", "Move the decrement from the Destroy event into the Bullet's Collision-with-Enemy event."],
    ["001 Game Creator", "Move the decrement from \"Defeated / Destroyed\" into the existing \"Collided with actor\" trigger."],
    ["Other 6 engines", "Already tied the win-counter to the specific hit-detection code in Lesson 6 - nothing to fix."],
  ];
  let ry = 2.35;
  rows.forEach(([eng, desc]) => {
    s.addShape("roundRect", { x: 0.6, y: ry, w: 12.1, h: 1.0, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(eng, { x: 0.85, y: ry + 0.08, w: 3.0, h: 0.85, valign: "middle", fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: 3.9, y: ry + 0.08, w: 8.6, h: 0.85, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
    ry += 1.15;
  });

  s.addShape("roundRect", { x: 0.6, y: 5.75, w: 12.1, h: 0.95, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("General principle: attach state-changing logic to the specific event that represents it, not to a generic lifecycle hook that fires for unrelated reasons too.", {
    x: 0.9, y: 5.85, w: 11.5, h: 0.75, valign: "middle", fontFace: BODY_FONT, fontSize: 13, italic: true, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 13: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "10 min", "On the Move - continuous descent and escape cleanup.");

// ---------- SLIDE 13c: Quick Aside - Meet the GameManager (what changed) ----------
{
  const s = bgSlide(false);
  s.addText("Before We Start: Meet the GameManager", { x: 0.6, y: 0.45, w: 11, h: 0.7, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: NAVY2, margin: 0 });
  pill(s, "quick aside", 11.15, 0.5, 1.6, false);

  s.addText("One more thing before today's lab - a change already sitting in your Lesson 6 project, not something new to build today.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addText("Your win-counter (enemies_remaining) and win-flag (has_won) now live on their own small dedicated object in every engine - a GameManager - instead of scattered across whichever object happened to need them first.", {
    x: 0.6, y: 2.15, w: 11.9, h: 0.75, fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, margin: 0,
  });

  const rows = [
    ["Scratch", "A hidden GameManager sprite, reached by broadcast/receive"],
    ["Godot", "A GameManager child Node, reached by signal connection"],
    ["GameMaker", "obj_gamemanager, a placed instance"],
    ["Unreal", "BP_GameManager, a placed Actor"],
    ["Unity", "Already worked this way - nothing changed for you"],
  ];
  let ry = 3.05;
  rows.forEach(([eng, desc]) => {
    s.addShape("roundRect", { x: 0.6, y: ry, w: 12.1, h: 0.58, rectRadius: 0.08, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(eng, { x: 0.85, y: ry, w: 2.4, h: 0.58, valign: "middle", fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: 3.3, y: ry, w: 9.2, h: 0.58, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
    ry += 0.68;
  });

  s.addShape("roundRect", { x: 0.6, y: 6.5, w: 12.1, h: 0.55, rectRadius: 0.1, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("If your own project isn't retrofitted yet, do that first today - your updated Lesson 6 handout shows exactly what it should look like.", {
    x: 0.85, y: 6.5, w: 11.6, h: 0.55, valign: "middle", fontFace: BODY_FONT, fontSize: 12, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(`"One more thing before you open your projects. Looking back at Lesson 6, we noticed the win-counter ended up living in a different place in almost every engine - the main scene here, the spawner there, one project's collision branch somewhere else. Only one of you, building in Unity, had it living on its own dedicated object built just to hold that kind of information. That turned out to be the better habit, so we've updated the Lesson 6 materials to match it everywhere, and I want you to update your own project too."

"It's a quick change, not a rebuild: create one object whose only job is holding game-wide information like the win-counter, and have everything else report to it instead of touching that number directly. Your updated Lesson 6 handout for your engine now shows exactly what that looks like - use it as your reference. Do this before you start today's movement work, since today's bug-fix discussion is about the exact same question: which piece of code should own which job."

[Give students 3-5 minutes to make the change if they haven't already, or confirm it's already in place. Circulate - GameMaker and Scratch tend to have the most questions here, since the retrofit is least like anything else they've built.]

[This slide's notes are a verbatim excerpt of the "[1:00-1:05] Today's Exercise & Handouts" section of Lesson_07_Teacher_Script.docx - keep both in sync if either is edited.]`);

  footer(s, false, page++);
}

// ---------- SLIDE 13d: Quick Aside - Why a Dedicated Manager Object ----------
{
  const s = bgSlide(false);
  s.addText("Why a Dedicated Manager Object?", { x: 0.6, y: 0.45, w: 11, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  pill(s, "quick aside", 11.15, 0.5, 1.6, false);

  s.addText("This is a genuinely common real-game pattern, not just a course convention - you'll also hear it called a Game Manager, a Singleton, or Game State.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  const reasons = [
    ["One clear place to look", "Shared, game-wide information has an obvious home - no hunting across multiple objects to find where a value actually lives."],
    ["Nothing has to reach into anything else", "Any object that needs shared state just talks to the manager, instead of one object reaching directly into another's private variables."],
    ["Room to grow", "Lives, score, level number, and anything else shared game-wide all have an obvious place to land, starting now - not something to retrofit again later."],
  ];
  let ry = 2.25;
  reasons.forEach(([title, desc], i) => {
    s.addShape("roundRect", { x: 0.6, y: ry, w: 12.1, h: 1.05, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addImage({ path: "icons/check.png", x: 0.85, y: ry + 0.28, w: 0.45, h: 0.45 });
    s.addText(title, { x: 1.5, y: ry + 0.08, w: 10.9, h: 0.35, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: 1.5, y: ry + 0.42, w: 10.9, h: 0.58, fontFace: BODY_FONT, fontSize: 12, color: TEXT_DARK, margin: 0 });
    ry += 1.18;
  });

  s.addShape("roundRect", { x: 0.6, y: 5.85, w: 12.1, h: 1.25, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 6.08, w: 0.55, h: 0.55 });
  s.addText("Discussion: as we add lives (Lesson 8) and a score (Lesson 9), where should those live? Same answer both times - on the GameManager, right alongside enemies_remaining.", {
    x: 1.6, y: 5.95, w: 10.9, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 13, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(`"Quick question before we move on: why bother with a separate object for this at all? Couldn't the win-counter just keep living on the spawner, or the player, like it did before?" [Pause for guesses]

"A few real reasons this is worth doing, not just tidiness for its own sake. First - anyone joining this project later, including future-you in three weeks, has exactly one place to look for shared information, instead of having to remember or rediscover which object happens to be holding it this time. Second - it means objects don't have to reach into each other's private variables to get work done; the Bullet doesn't need to know anything about how the Spawner is built, it just talks to the GameManager. Third, and this is the one that pays off soonest: next lesson we're adding lives, and the lesson after that, score. Both of those are exactly the same kind of shared, game-wide information as today's win-counter - and they already have an obvious home waiting for them."

[Demo cue, optional: pull up any one engine's GameManager object live and point out it currently holds exactly two things - point out that by Lesson 9 it'll hold four, and that's the design working as intended, not scope creep.]

"This pattern has a bunch of different names depending on who you ask - Game Manager, Singleton, Game State - but it's the same idea everywhere: one object, one job, holding information the rest of the game needs to share."

[This slide's notes are a verbatim excerpt of the "[1:00-1:05] Today's Exercise & Handouts" section of Lesson_07_Teacher_Script.docx - keep both in sync if either is edited.]`);

  footer(s, false, page++);
}

// ---------- SLIDE 14: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("On the Move", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("Every enemy advances toward the player, and any that escape are cleaned up without counting as a kill.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "Enemies move down; an escaped enemy is removed without counting as a kill.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "A moving enemy can still be shot correctly; escape never triggers \"You Win!\"", "Everything works end to end."],
    ["Level 3 - Refined", "Varied enemy speeds, and/or a cleaner separation of movement vs. hit logic.", "Ready to share."],
  ];
  let lx = 0.6;
  levels.forEach(([title, desc, tag]) => {
    s.addShape("roundRect", { x: lx, y: 2.25, w: 3.9, h: 2.2, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(title, { x: lx + 0.2, y: 2.4, w: 3.5, h: 0.4, fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: lx + 0.2, y: 2.85, w: 3.5, h: 1.15, fontFace: BODY_FONT, fontSize: 11.8, color: TEXT_DARK, margin: 0 });
    s.addText(tag, { x: lx + 0.2, y: 3.95, w: 3.5, h: 0.4, fontFace: BODY_FONT, fontSize: 11.5, italic: true, color: GOLD, margin: 0 });
    lx += 4.1;
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 4.97, w: 0.55, h: 0.55 });
  s.addText("GameMaker and 001 Game Creator: read your handout's setup section before touching anything - you're deleting code today, not just adding it.", {
    x: 1.6, y: 4.87, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 13, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 15: Lab Time Notes ----------
{
  const s = bgSlide(false);
  s.addText("Lab Time Notes", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/wrench.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const notes = [
    "Most important check this week: let an enemy escape unshot and confirm \"You Win!\" does NOT appear - for every engine, not just the two with a fix.",
    "GameMaker / 001 Game Creator: watch for double-counting - the win message triggering after only half the enemies are shot means the old hook wasn't fully removed.",
    "Enemies never disappearing at the bottom usually means the off-screen threshold doesn't match the project's screen size - adjust the number.",
    "GameMaker and 001 Game Creator clusters need the most attention this week - circulate there first.",
  ];
  notes.forEach((t, i) => {
    const y = 1.75 + i * 1.15;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.9, fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  footer(s, false, page++);
}

// ---------- SLIDE 16: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("Enemies Are Moving!", { x: 0, y: 2.45, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, margin: 0 });
  s.addText("Continuous motion, escape cleanup, and a real fix to a latent bug - solid engineering, not just new features.", {
    x: 1.5, y: 3.35, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0,
  });
  s.addText("Next week: a lose condition and a lives system - the first threat to your player.", {
    x: 1.5, y: 4.0, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 15, color: CREAM, margin: 0,
  });
  s.addText("Try a Level 3 idea at home if you'd like - varied enemy speeds are a fun one to tune.", {
    x: 1.5, y: 4.7, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_07_Slides.pptx" }).then(() => console.log("done"));
