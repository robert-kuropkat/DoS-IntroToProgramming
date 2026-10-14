const pptxgen = require("pptxgenjs");
const fs = require("fs");
// Speaker notes: the Teacher's Script text, copied verbatim (Standards Rule 8).
// lesson11_script_notes.json is written by build_lesson11_plan_teacher_script.js.
const NOTES = JSON.parse(fs.readFileSync("lesson11_script_notes.json", "utf8"));

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
const LESSON_LABEL = "Lesson 11";

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
  s.addText("Lesson 11", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Core Mechanics V: Wave After Wave", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 11  -  Level 2 Complete Milestone", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, bold: true, color: GOLD_L, margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 11 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Nest a second decision inside an existing win-check",
    "Reuse spawn_enemies from a third caller: the game itself",
    "Keep new state (wave) on the GameManager - and reset it",
    "Explain what the Level 2 Complete milestone means",
  ];
  objectives.forEach((t, i) => {
    const y = 2.15 + i * 0.62;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.15, y: y - 0.03, w: 6.1, h: 0.55, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 7.6, y: 1.65, w: 5.15, h: 4.6, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("What We're Using Today", { x: 7.9, y: 1.85, w: 4.6, h: 0.4, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0 });
  const materials = [
    "Demo laptop with the four active engines verified",
    "Unity (C#), for the live lecture demo",
    "No new assets this week - reuses the spawn routine from Lesson 10",
    "This is a milestone lesson - Level 2 Complete!",
    "Your own [YourName]Game project from Lesson 10 - not a new one",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Programming - state machines, code reuse/refactoring, game balance and pacing (intermediate).", {
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
    ["0:00 - 0:05", "Welcome back, Lesson 10 recap (game state, the Stalemate)"],
    ["0:05 - 0:20", "One More Question: Is This the Last Wave?"],
    ["0:20 - 0:40", "The Same Spawn Code, Called a Third Way"],
    ["0:40 - 0:50", "Milestone Moment: What \"Level 2 Complete\" Actually Means"],
    ["0:50 - 0:55", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— Wave After Wave —"],
    ["0:55 - 1:50", "Build multiple waves, reaching the milestone"],
    ["1:50 - 2:00", "Exit check: milestone celebration share-out"],
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
sectionDivider("book", "Welcome Back", "5 min", "Lesson 10 recap - your game got genuinely replayable.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 10:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "A game_active flag that freezes gameplay, and one reset routine run by R",
    "Game state: everything that changed must be put back on restart",
    "The Stalemate: last week's Level 3 bug - still in your game, on purpose",
    "Today: clearing one wave becomes clearing several before you truly win",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.66;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });
  s.addNotes(NOTES.welcome);

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one - and no new assets needed either, everything builds on the spawn routine you already have.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - One More Question ----------
sectionDivider("design", "One More Question", "15 min", "Is this actually the last wave?");

// ---------- SLIDE 7: One More Question content ----------
{
  const s = bgSlide(false);
  s.addText("Nesting a Decision", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "15 min", 11.4, 0.5, 1.3, false);

  s.addShape("roundRect", { x: 0.6, y: 1.5, w: 12.1, h: 1.6, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 1.75, w: 0.55, h: 0.55 });
  s.addText("Since Lesson 6: are all enemies gone? Today, a second part: gone, AND is this the last wave? Both true -> real win. Otherwise -> the round continues.", {
    x: 1.6, y: 1.65, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addText("The first time this course nests one decision inside another - a small code change, a genuinely new shape of thinking.", {
    x: 0.6, y: 3.35, w: 11.9, h: 0.75, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.3, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 4.53, w: 0.55, h: 0.55 });
  s.addText("Discussion: can you think of a real-life decision where the answer to one question depends on a second question nested inside it? (\"Can I go to the party\" depends on \"is my homework done,\" which might depend on \"do I understand the assignment.\")", {
    x: 1.6, y: 4.43, w: 10.9, h: 1.5, valign: "middle", fontFace: BODY_FONT, fontSize: 13, italic: true, color: CREAM, margin: 0,
  });
  s.addNotes(NOTES.question);

  footer(s, false, page++);
}

// ---------- SLIDE 7b: Activity Diagram ----------
{
  const s = bgSlide(false);
  s.addText("The Whole Decision, One Picture", { x: 0.6, y: 0.4, w: 12, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "diagram_activity_wave_flow.png", x: 0.8, y: 1.3, w: 6.6, h: 5.28 });
  s.addText("Two questions, in order", { x: 7.8, y: 1.5, w: 5.0, h: 0.5, fontFace: BODY_FONT, fontSize: 20, bold: true, color: NAVY2, margin: 0 });
  s.addText([
    { text: "Is the wave cleared? No: carry on.", options: { bullet: true, breakLine: true } },
    { text: "Was it the last wave? Yes: the real win.", options: { bullet: true, breakLine: true } },
    { text: "No: wave + 1, enemies_remaining = 5, spawn_enemies.", options: { bullet: true, breakLine: true } },
    { text: "Score, lives and game_active are untouched between waves.", options: { bullet: true } },
  ], { x: 7.8, y: 2.15, w: 4.9, h: 3.6, fontFace: BODY_FONT, fontSize: 16, color: TEXT_DARK, paraSpaceAfter: 10, valign: "top", margin: 0 });
  s.addNotes(NOTES.activity);
  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - The Same Spawn Code ----------
sectionDivider("programming", "A Third Door In", "20 min", "The same enemy-spawning code, one more way to call it.");

// ---------- SLIDE 9: Same Spawn Code content ----------
{
  const s = bgSlide(false);
  s.addText("Three Doors, One Room", { x: 0.6, y: 0.45, w: 10.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("No new enemy-spawning code today. You've had it since Lesson 5. Lesson 10 made it callable a second way, for restarts. Today: a THIRD way, every wave clear.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.85, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.5, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/programming.png", x: 0.85, y: 2.75, w: 0.55, h: 0.55 });
  s.addText("Unity demo: SpawnEnemies() already has two callers (Start() and ResetGame()). RegisterKill() is static, so it can't use the Spawner slot - it raises nextWaveRequested, and GameManager's Update() spawns the wave.", {
    x: 1.6, y: 2.65, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.4, w: 12.1, h: 1.4, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("Same enemy-creation code. Three doors in: the very first game, every restart, and now every wave clear. That's the whole idea.", {
    x: 0.9, y: 4.6, w: 11.5, h: 1.0, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, italic: true, color: CREAM, margin: 0,
  });
  s.addNotes(NOTES.demo);

  footer(s, false, page++);
}

// ---------- SLIDE 9b: New State, Same Habit ----------
{
  const s = bgSlide(false);
  s.addText("New State, Same Habit", { x: 0.6, y: 0.4, w: 12, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  s.addText("wave and total_waves are game state too - and they live on the GameManager.", { x: 0.6, y: 1.2, w: 12, h: 0.5, fontFace: BODY_FONT, fontSize: 18, italic: true, color: TEXT_MUTE, margin: 0 });
  const cols = [
    ["Does restart reset it?", "wave must go back to 1 in the one reset routine - test it from wave 3"],
    ["Who owns it?", "The GameManager. Objects report; the manager counts and decides"],
    ["Anything extra?", "A wave message, an effect, a sound - what changes it, and what puts it back?"],
    ["Same two questions", "What changed when the round or wave ended? What changes it back?"],
  ];
  cols.forEach(([h, d], idx) => {
    const x = 0.6 + (idx % 2) * 6.2, y = 2.0 + Math.floor(idx / 2) * 1.75;
    s.addShape("roundRect", { x, y, w: 5.9, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { color: GOLD, width: 1 } });
    s.addText(h, { x: x + 0.25, y: y + 0.15, w: 5.4, h: 0.45, fontFace: BODY_FONT, fontSize: 18, bold: true, color: NAVY2, margin: 0 });
    s.addText(d, { x: x + 0.25, y: y + 0.65, w: 5.4, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, margin: 0 });
  });
  s.addText("Test it: reach wave 3, lose, press R. The HUD must say Wave 1.", { x: 0.6, y: 5.6, w: 12.1, h: 0.9, fontFace: BODY_FONT, fontSize: 17, bold: true, color: NAVY, margin: 0 });
  s.addNotes(NOTES.state);
  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - Milestone Moment ----------
sectionDivider("shield", "Milestone Moment", "10 min", "What finishing Level 2 Complete actually means.");

// ---------- SLIDE 11: Milestone Moment content ----------
{
  const s = bgSlide(false);
  s.addText("Level 2 Complete", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addShape("roundRect", { x: 0.6, y: 1.45, w: 12.1, h: 2.0, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("Once multiple waves work today, your game can:", { x: 0.9, y: 1.65, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: NAVY2, margin: 0 });
  const items = ["Move & shoot", "Face multiple waves", "Track score", "Survive with lives", "Restart cleanly", "Truly win or lose"];
  items.forEach((t, i) => {
    s.addImage({ path: "icons/check.png", x: 0.9 + (i % 3) * 3.95, y: 2.1 + Math.floor(i / 3) * 0.6, w: 0.36, h: 0.36 });
    s.addText(t, { x: 1.32 + (i % 3) * 3.95, y: 2.08 + Math.floor(i / 3) * 0.6, w: 3.3, h: 0.4, valign: "middle", fontFace: BODY_FONT, fontSize: 13, color: TEXT_DARK, margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 3.65, w: 12.1, h: 2.55, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("Level 1 MVP meant a complete, playable loop. Level 2 Complete means real depth - failure that matters, progress that builds, escalation as more waves arrive. This is the kind of thing that starts to feel like an actual arcade game, not just a working prototype.", {
    x: 0.9, y: 3.85, w: 11.5, h: 1.2, valign: "middle", fontFace: BODY_FONT, fontSize: 14, italic: true, color: CREAM, margin: 0,
  });
  s.addText("From here through Lesson 15: sound, animation, difficulty, and polish. The mechanics are done - what's left is making them sing.", {
    x: 0.9, y: 5.1, w: 11.5, h: 1.0, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, bold: true, color: GOLD_L, margin: 0,
  });
  s.addNotes(NOTES.milestone);

  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "5 min", "Wave After Wave - building your second milestone.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("Wave After Wave", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Clearing a wave spawns a fresh one, up to 3 total. Clearing wave 3 is the real win - built on top of your existing project.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "Clearing a wave spawns a new one, up to 3; wave shown on screen.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "Seamless, correct transitions every time; restart returns to wave 1.", "Everything works end to end."],
    ["Level 3 - Refined", "Fix the Stalemate so it works with waves; a brief wave announcement; tighter reuse.", "Ready to share."],
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
  s.addText("Test a FULL 3-wave playthrough today, not just wave 1 - and confirm losing a life mid-wave still works correctly no matter which wave you're on.", {
    x: 1.6, y: 4.82, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, color: CREAM, margin: 0,
  });
  s.addNotes(NOTES.exercise);

  footer(s, false, page++);
}

// ---------- SLIDE 14: Lab Time Notes ----------
{
  const s = bgSlide(false);
  s.addText("Lab Time Notes", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/wrench.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const notes = [
    "Most common bug: forgetting the wave-count check, so clearing wave 1 immediately shows the win message. Test this explicitly.",
    "Second most common: the wave-advance branch accidentally touching score or lives - only enemies_remaining and wave should change.",
    "Game-state drift: test restarting from mid-wave-2 or mid-wave-3 - restart should always return to wave 1.",
    "The Stalemate: confirm it is real, name it, don't hand over the fix. Level 3 is where it gets solved.",
    "At 1:50, do a real celebration share-out - this is the Level 2 Complete milestone. Consider 2-3 students demoing a full 3-wave playthrough.",
  ];
  notes.forEach((t, i) => {
    const y = 1.6 + i * 0.95;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.8, fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });
  s.addNotes(NOTES.lab);

  footer(s, false, page++);
}

// ---------- SLIDE 15: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("Level 2 Complete!", { x: 0, y: 2.45, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, margin: 0 });
  s.addText("Multiple waves, real stakes, a clean restart - a genuine arcade game, not just a working prototype.", {
    x: 1.5, y: 3.35, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0,
  });
  s.addText("Next week: sound effects - shoot, explosion, game over, wave clear. The polish begins.", {
    x: 1.5, y: 4.0, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 15, color: CREAM, margin: 0,
  });
  s.addText("Keep playtesting your full 3-wave run at home if you'd like - try a Level 3 idea if you haven't already.", {
    x: 1.5, y: 4.7, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_11_Slides.pptx" }).then(() => console.log("done"));
