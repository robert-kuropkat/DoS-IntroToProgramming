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
const LESSON_LABEL = "Lesson 13";

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
  s.addText("Lesson 13", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("A Hit That Reacts", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 13  -  Basic Animation", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, bold: true, color: GOLD_L, margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 13 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Sequence a flash before finalizing an existing state change",
    "Spawn a fully independent, self-cleaning explosion effect",
    "Compare how 3 engine families pause mid-script (or don't)",
    "Confirm game logic timing is completely unchanged",
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
    "One new asset: explosion_placeholder.png",
    "Unreal reuses its existing Cube/Sphere shapes",
    "Your own [YourName]Game project from Lesson 12",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Visual Art - basic animation, juice/feedback fundamentals (introductory); Programming - timed/asynchronous execution (intermediate).", {
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
    ["0:00 - 0:05", "Welcome back, Lesson 12 recap"],
    ["0:05 - 0:15", "A Hit That Reacts: Two New Visual Touches"],
    ["0:15 - 0:40", "Pausing Mid-Script: Three Engine Families, Three Solutions"],
    ["0:40 - 0:50", "An Explosion That Minds Its Own Business"],
    ["0:50 - 0:55", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— A Hit That Reacts —"],
    ["0:55 - 1:50", "Wire up the flash and explosion, verify game logic unchanged"],
    ["1:50 - 2:00", "Exit check: best explosion wins bragging rights"],
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
sectionDivider("book", "Welcome Back", "5 min", "Your game has a voice. Now it gets a body.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 12:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Four sounds wired into their exact matching events",
    "Since Lesson 4, a hit just vanishes - functional, not satisfying",
    "Today: a brief flash, then a small explosion, before it disappears",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one. One new image is provided (explosion_placeholder.png) - Unreal students reuse existing shapes, TIC-80 students draw their own.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - A Hit That Reacts ----------
sectionDivider("design", "A Hit That Reacts", "10 min", "Two new visual touches - and one strict rule.");

// ---------- SLIDE 7: A Hit That Reacts content ----------
{
  const s = bgSlide(false);
  s.addText("Two New Visual Touches", { x: 0.6, y: 0.45, w: 10.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addShape("roundRect", { x: 0.6, y: 1.5, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 1.75, w: 0.55, h: 0.55 });
  s.addText("First: the enemy itself briefly flashes - a reaction ON the thing that got hit. Second: a completely separate explosion effect appears at that spot and plays itself out - independent of the enemy entirely.", {
    x: 1.6, y: 1.65, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 3.4, w: 12.1, h: 1.85, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("The rule for today: score, lives, wave, and win/lose logic don't change WHEN they happen. Only how the enemy's removal LOOKS is changing.", {
    x: 0.9, y: 3.6, w: 11.5, h: 1.5, valign: "middle", fontFace: BODY_FONT, fontSize: 16, bold: true, color: GOLD_L, margin: 0,
  });

  s.addText("If you find yourself moving scoring code around this week, stop and re-check the Handout.", {
    x: 0.6, y: 5.5, w: 12.1, h: 0.5, fontFace: BODY_FONT, fontSize: 13, italic: true, color: TEXT_MUTE, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Pausing Mid-Script ----------
sectionDivider("programming", "Pausing Mid-Script", "25 min", "Three engine families, three genuinely different solutions.");

// ---------- SLIDE 9: Pausing Mid-Script content ----------
{
  const s = bgSlide(false);
  s.addText("Three Solutions, One Problem", { x: 0.6, y: 0.45, w: 10.8, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  pill(s, "25 min", 11.4, 0.5, 1.3, false);

  s.addText("To see a flash, the enemy must stick around a moment after being hit - something in our code needs to pause, without freezing the whole game.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.85, fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: NAVY2, margin: 0,
  });

  const rows = [
    ["Godot & Unity", "await / Coroutine - the language itself can pause mid-function."],
    ["GameMaker & TIC-80", "No blocking wait - split into \"trigger\" and \"finish\" via an Alarm or a countdown."],
    ["Unreal & 001GC", "Delay node pauses a Blueprint wire; 001GC's Flash manages its own timing."],
  ];
  let ry = 2.45;
  rows.forEach(([title, desc]) => {
    s.addShape("roundRect", { x: 0.6, y: ry, w: 12.1, h: 1.05, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(title, { x: 0.9, y: ry + 0.12, w: 3.3, h: 0.8, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: 4.3, y: ry + 0.12, w: 8.2, h: 0.8, valign: "middle", fontFace: BODY_FONT, fontSize: 13, color: TEXT_DARK, margin: 0 });
    ry += 1.2;
  });

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - An Explosion That Minds Its Own Business ----------
sectionDivider("wrench", "Minds Its Own Business", "10 min", "One object, one job, cleans up after itself.");

// ---------- SLIDE 11: Explosion content ----------
{
  const s = bgSlide(false);
  s.addText("An Explosion That Minds Its Own Business", { x: 0.6, y: 0.45, w: 12.1, h: 0.7, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addShape("roundRect", { x: 0.6, y: 1.5, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 1.75, w: 0.55, h: 0.55 });
  s.addText("The explosion spawns at the hit location and immediately forgets everything else - no score, no wave, not even which enemy made it. It plays itself out and cleans itself up, on its own.", {
    x: 1.6, y: 1.65, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 3.4, w: 12.1, h: 1.55, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("This is the exact same discipline your bullets have followed since Lesson 3, and your enemies since Lesson 5. Every object, one job, cleans up after itself. The explosion is just the newest member of that club.", {
    x: 0.9, y: 3.6, w: 11.5, h: 1.2, valign: "middle", fontFace: BODY_FONT, fontSize: 14, italic: true, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "5 min", "A Hit That Reacts - building the feedback loop.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("A Hit That Reacts", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Add a flash and an independent explosion effect - built on top of your existing project, with zero change to game-state timing.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "Every hit flashes the enemy and spawns an explosion effect.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "Correct on every hit, no stuck enemies, timing unaffected.", "Everything works end to end."],
    ["Level 3 - Refined", "Smoother fades, a scaling explosion, or a multi-frame flicker.", "Ready to share."],
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
  s.addText("Test a FULL round when you're done - shooting, a wave clear, even a loss - and confirm nothing about your SCORING changed, only the visuals.", {
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
    "Watch for score/win-check logic accidentally moved along with the delay - it should stay immediate; only the enemy's final removal is delayed.",
    "GameMaker and TIC-80 need extra circulating time this week - splitting logic across two events/table fields is a bigger structural change than usual.",
    "001 Game Creator: the enemy's own flash may be brief or inconsistent - that's expected. The explosion effect is the reliable part to check first.",
    "This is a fun, visual week - let students actually watch a few hits happen slowly, and consider a quick show-and-tell of favorite explosion tweaks.",
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
  s.addText("It Reacts!", { x: 0, y: 2.45, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, margin: 0 });
  s.addText("A flash, a burst, then gone - getting hit finally looks like getting hit.", {
    x: 1.5, y: 3.35, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0,
  });
  s.addText("Next week: difficulty scaling across waves, and optional stretch features.", {
    x: 1.5, y: 4.0, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 15, color: CREAM, margin: 0,
  });
  s.addText("Keep playtesting a full round at home if you'd like - try a Level 3 animation idea if you haven't already.", {
    x: 1.5, y: 4.7, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_13_Slides.pptx" }).then(() => console.log("done"));
