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
pres.layout = "LAYOUT_WIDE";

const SW = 13.333, SH = 7.5;
const LESSON_LABEL = "Lesson 14";

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
  s.addText("Lesson 14", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Raising the Stakes", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 14  -  Difficulty Scaling & Stretch Features", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, bold: true, color: GOLD_L, margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 14 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Replace hardcoded numbers with wave-dependent formulas",
    "Keep every referenced value in agreement, everywhere",
    "Scale both enemy count AND enemy speed by wave",
    "Choose and attempt one capstone-level stretch feature",
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
    "No new assets - reuses existing enemies/spawn logic",
    "001GC: 2 new pre-placed enemy copies (design time)",
    "Your own [YourName]Game project from Lesson 13",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Game Design - balance, pacing, difficulty curves (intermediate); Programming - parameterization (intermediate).", {
    x: 0.6, y: 6.55, w: 12.1, h: 0.5, fontFace: BODY_FONT, fontSize: 11.5, italic: true, color: TEXT_MUTE, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 3: TODAY'S AGENDA ----------
{
  const s = bgSlide(false);
  s.addText("Today's Agenda", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/design.png", x: 11.6, y: 0.4, w: 0.9, h: 0.9 });

  const rows = [
    ["0:00 - 0:05", "Welcome back, Lesson 13 recap"],
    ["0:05 - 0:15", "The Foreshadowed Feature: One Formula, Two Places"],
    ["0:15 - 0:35", "Adding Speed to the Mix"],
    ["0:35 - 0:45", "This Week's Twist: Student's Choice"],
    ["0:45 - 0:55", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— Raising the Stakes —"],
    ["0:55 - 1:50", "Scale count and speed, then pick one stretch feature"],
    ["1:50 - 2:00", "Exit check: toughest wave 3 wins bragging rights"],
  ];

  let y = 1.4;
  rows.forEach(([time, seg]) => {
    const isDivider = time === "LAB TIME";
    const rowH = isDivider ? 0.46 : 0.56;
    s.addShape("roundRect", {
      x: 0.6, y, w: 12.1, h: rowH, rectRadius: 0.08,
      fill: { color: isDivider ? GOLD : "FFFFFF" },
      line: { color: "DCE3EC", width: isDivider ? 0 : 1 },
    });
    s.addText(time, {
      x: 0.9, y, w: 2.3, h: rowH, valign: "middle", fontFace: BODY_FONT,
      fontSize: isDivider ? 12.5 : 13, bold: true, color: isDivider ? NAVY : NAVY2, margin: 0,
    });
    s.addText(seg, {
      x: 3.3, y, w: 9.2, h: rowH, valign: "middle", fontFace: BODY_FONT,
      fontSize: isDivider ? 12.5 : 12.3, bold: isDivider, italic: isDivider,
      color: isDivider ? NAVY : TEXT_DARK, margin: 0,
    });
    y += rowH + 0.1;
  });

  footer(s, false, page++);
}

// ---------- SLIDE 4: SECTION - Welcome Back ----------
sectionDivider("book", "Welcome Back", "5 min", "Today, your game starts fighting back.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 13:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Getting hit finally looks like getting hit - flash, then explosion",
    "Lesson 11's Key deliberately called 5-per-wave a \"simplification\"",
    "Today's the payoff: count AND speed both scale with wave",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one. No new assets needed - 001 Game Creator students add two more pre-placed enemy copies first (see your Handout).", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - The Foreshadowed Feature ----------
sectionDivider("design", "One Formula, Two Places", "10 min", "The payoff for a simplification flagged weeks ago.");

// ---------- SLIDE 7: Foreshadowed Feature content ----------
{
  const s = bgSlide(false);
  s.addText("One Formula, Two Places", { x: 0.6, y: 0.45, w: 10.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addShape("roundRect", { x: 0.6, y: 1.5, w: 12.1, h: 1.6, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 1.75, w: 0.55, h: 0.55 });
  s.addText("Instead of always spawning 5 enemies, spawn (4 + wave). Wave 1: still 5. Wave 2: 6. Wave 3: 7. One small formula, and the game already escalates.", {
    x: 1.6, y: 1.65, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 3.35, w: 12.1, h: 1.9, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("The catch: that formula has to appear in TWO places - wherever enemies spawn, AND wherever enemies_remaining resets for a new wave. If they disagree, the win-check breaks.", {
    x: 0.9, y: 3.55, w: 11.5, h: 1.5, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: GOLD_L, margin: 0,
  });

  s.addText("One source of truth, referenced everywhere - the same discipline behind Lesson 10's reusable reset function.", {
    x: 0.6, y: 5.4, w: 12.1, h: 0.5, fontFace: BODY_FONT, fontSize: 13, italic: true, color: TEXT_MUTE, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Adding Speed ----------
sectionDivider("programming", "Adding Speed to the Mix", "20 min", "Count is a start. Speed makes it dangerous.");

// ---------- SLIDE 9: Adding Speed content ----------
{
  const s = bgSlide(false);
  s.addText("Adding Speed to the Mix", { x: 0.6, y: 0.45, w: 10.8, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("Count alone is a good start, but speed makes it feel dangerous.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.2, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/programming.png", x: 0.85, y: 2.45, w: 0.55, h: 0.55 });
  s.addText("Godot demo: in spawn_wave(), right after creating each enemy, set enemy.speed = 60.0 + (wave - 1) * 20.0. Wave 1 stays the same; wave 3 is noticeably faster.", {
    x: 1.6, y: 2.35, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.1, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("The enemy's OWN movement code from Lesson 4 doesn't change at all - it already reads a speed variable. Today we just started setting it differently per wave. That's good structure paying off.", {
    x: 0.9, y: 4.3, w: 11.5, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 14, italic: true, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - Student's Choice ----------
sectionDivider("wrench", "Student's Choice", "10 min", "Pick ONE. Good-faith effort is a complete finish.");

// ---------- SLIDE 11: Student's Choice content ----------
{
  const s = bgSlide(false);
  s.addText("This Week's Twist", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("Level 3 works differently this week: pick ONE of five real capstone-level ideas.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  const items = ["Destructible shields", "A power-up", "A bonus enemy", "Formation movement", "Title / game-over screen"];
  items.forEach((t, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = 0.9 + col * 6.0;
    const y = 2.2 + row * 0.75;
    s.addImage({ path: "icons/check.png", x, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: x + 0.5, y: y - 0.05, w: 5.4, h: 0.5, valign: "middle", fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 5.0, w: 12.1, h: 1.4, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("These are open-ended on purpose. A good-faith attempt at the general idea - even if rough - is a complete Level 3 finish today. Pick whichever sounds the most fun to YOU.", {
    x: 0.9, y: 5.2, w: 11.5, h: 1.0, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "5 min", "Raising the Stakes - scale it, then choose your own.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("Raising the Stakes", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Scale enemy count and speed by wave, then pick one stretch feature - built on top of your existing project.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "Enemy count increases by one each wave (5, 6, 7).", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "Speed ALSO scales; restart returns to wave 1's values.", "Everything works end to end."],
    ["Level 3 - Refined", "Choose ONE stretch feature; a good-faith attempt.", "Ready to share."],
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
  s.addText("Test a full 3-wave round when you're done, and confirm restarting brings everything back to wave 1's original feel.", {
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
    "Watch for the spawn-count formula updated in one place but not the enemies_remaining reset in the other - an easy one to half-do.",
    "001 Game Creator students need extra time up front to correctly add and name enemy_6/enemy_7 before any scripting begins.",
    "If a cluster's wave 3 feels impossible rather than exciting, help them dial back their difficulty increments rather than pushing through.",
    "Level 3 is open-ended this week - help students scope their chosen feature down to something achievable in the remaining time.",
  ];
  notes.forEach((t, i) => {
    const y = 1.75 + i * 1.15;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.9, fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  footer(s, false, page++);
}

// ---------- SLIDE 15: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("The Stakes Are Raised!", { x: 0, y: 2.45, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, margin: 0 });
  s.addText("More enemies, faster enemies, and a feature that's all your own.", {
    x: 1.5, y: 3.35, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0,
  });
  s.addText("Next week: a final polish pass, and the Semester 1 Capstone presentation.", {
    x: 1.5, y: 4.0, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 15, color: CREAM, margin: 0,
  });
  s.addText("Keep tuning your difficulty curve at home if you'd like - and keep building on your Level 3 feature if you're excited about it.", {
    x: 1.5, y: 4.7, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_14_Slides.pptx" }).then(() => console.log("done"));
