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
const LESSON_LABEL = "Lesson 15";

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
  s.addText("Lesson 15", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Ready to Show", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 15  -  Level 3 Refined Milestone - Capstone Presentation", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, bold: true, color: GOLD_L, margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 15 objectives & format", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Self-assess against the full capstone checklist",
    "Complete a final, prioritized polish pass",
    "Present a working game confidently to the class",
    "Reach the Level 3 Refined milestone",
  ];
  objectives.forEach((t, i) => {
    const y = 2.15 + i * 0.62;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.15, y: y - 0.03, w: 6.1, h: 0.55, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 7.6, y: 1.65, w: 5.15, h: 4.6, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("Today Is Different", { x: 7.9, y: 1.85, w: 4.6, h: 0.4, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0 });
  const materials = [
    "No lecture demo, no new mechanics today",
    "A workshop-and-showcase day",
    "No new assets - this IS the capstone",
    "Bring: your own Lesson 14 project, nothing else",
    "Demo laptop on standby, backup only",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Production - scoping, finishing, and presenting a project (intermediate); all Five Disciplines, in review.", {
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
    ["0:00 - 0:10", "Welcome back, capstone reflection, milestone framing"],
    ["POLISH TIME", "— Work the Full Capstone Checklist —"],
    ["0:10 - 0:55", "Fix, tune, and finish - checklist first, extras second"],
    ["0:55 - 1:00", "Presentation logistics, running order / sign-up"],
    ["PRESENTATIONS", "— About 2 Minutes Each —"],
    ["1:00 - 1:50", "Play a round, name your game, show your feature"],
    ["1:50 - 2:00", "Closing celebration, look ahead to Semester 2"],
  ];

  let y = 1.4;
  rows.forEach(([time, seg]) => {
    const isDivider = time === "POLISH TIME" || time === "PRESENTATIONS";
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
sectionDivider("book", "Welcome Back", "10 min", "Fifteen weeks. A real, finished game.");

// ---------- SLIDE 5: Welcome content ----------
{
  const s = bgSlide(false);
  s.addText("Capstone Reflection", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("Today is different from every other day this semester:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "No new code to learn - today you finish, and today you show what you built",
    "Since Lesson 1, you've built a game that moves, shoots, scores, wins, loses, sounds, reacts, and escalates",
    "About 45 minutes of polish time, then about 2-minute presentations - no slides, no stress",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/book.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today - this IS your capstone. There's nothing separate to submit, and nothing new to bring.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - Polish Time ----------
sectionDivider("wrench", "Polish Time", "45 min", "Checklist first. Extras only after that.");

// ---------- SLIDE 7: Polish Time content ----------
{
  const s = bgSlide(false);
  s.addText("Polish Time", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "45 min", 11.4, 0.5, 1.3, false);

  s.addText("Work your Handout's checklist top to bottom - in priority order:", { x: 0.6, y: 1.45, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0 });

  const rows = [
    ["1", "Fix anything broken or unchecked", "A complete, simple game beats a broken ambitious one."],
    ["2", "Tune anything unfair or confusing", "Difficulty, text clarity, sound volume/timing."],
    ["3", "Only then: your Lesson 14 feature", "Or any extra polish, if time and energy remain."],
  ];
  let ry = 2.2;
  rows.forEach(([num, title, desc]) => {
    s.addShape("ellipse", { x: 0.6, y: ry + 0.1, w: 0.55, h: 0.55, fill: { color: GOLD }, line: { type: "none" } });
    s.addText(num, { x: 0.6, y: ry + 0.1, w: 0.55, h: 0.55, align: "center", valign: "middle", fontFace: HEAD_FONT, fontSize: 20, bold: true, color: NAVY, margin: 0 });
    s.addShape("roundRect", { x: 1.35, y: ry, w: 11.35, h: 1.1, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(title, { x: 1.6, y: ry + 0.08, w: 10.9, h: 0.45, fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: 1.6, y: ry + 0.5, w: 10.9, h: 0.5, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
    ry += 1.3;
  });

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Milestone Moment ----------
sectionDivider("shield", "Milestone Moment", "10 min", "What Level 3 Refined actually means.");

// ---------- SLIDE 9: Milestone Moment content ----------
{
  const s = bgSlide(false);
  s.addText("Level 3 Refined", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addShape("roundRect", { x: 0.6, y: 1.45, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("Level 1 MVP (Lesson 6) meant a complete, playable loop. Level 2 Complete (Lesson 11) meant real depth. Level 3 Refined means feature-complete, tested, and polished enough to comfortably share outside of class.", { x: 0.9, y: 1.65, w: 11.5, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0 });

  s.addShape("roundRect", { x: 0.6, y: 3.35, w: 12.1, h: 2.55, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("That's not a claim that everything is perfect. It's a claim that your game is DONE - a real, working, presentable arcade game, built by you, over fifteen weeks, in an engine you chose.", {
    x: 0.9, y: 3.55, w: 11.5, h: 1.2, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, italic: true, color: CREAM, margin: 0,
  });
  s.addText("However far up the Level ladder you got, a finished game is today's whole point - not a perfect one.", {
    x: 0.9, y: 4.8, w: 11.5, h: 1.0, valign: "middle", fontFace: BODY_FONT, fontSize: 14, bold: true, color: GOLD_L, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - Presentations ----------
sectionDivider("controller", "Presentations", "50 min", "Play a round. Tell us about it. That's it.");

// ---------- SLIDE 11: Presentations content ----------
{
  const s = bgSlide(false);
  s.addText("Presenting Your Game", { x: 0.6, y: 0.45, w: 10.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "50 min", 11.4, 0.5, 1.3, false);

  s.addText("About 2 minutes each. No slides, no notes required - just play while you talk.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.5, fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: NAVY2, margin: 0,
  });

  const steps = [
    "Name your game, and its core mechanic, in one sentence",
    "Play a short round - a hit, a wave clear, a win or loss",
    "Point out your chosen Level 3 stretch feature specifically",
    "If something breaks live: say what it's supposed to do, and move on",
  ];
  steps.forEach((t, i) => {
    const y = 2.15 + i * 0.85;
    s.addShape("ellipse", { x: 0.6, y, w: 0.45, h: 0.45, fill: { color: GOLD }, line: { type: "none" } });
    s.addText(String(i + 1), { x: 0.6, y, w: 0.45, h: 0.45, align: "center", valign: "middle", fontFace: HEAD_FONT, fontSize: 16, bold: true, color: NAVY, margin: 0 });
    s.addText(t, { x: 1.25, y: y - 0.05, w: 11.2, h: 0.6, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0 });
  });

  footer(s, false, page++);
}

// ---------- SLIDE 12: Polish & Presentation Notes ----------
{
  const s = bgSlide(false);
  s.addText("Polish & Presentation Notes", { x: 0.6, y: 0.45, w: 11.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/wrench.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const notes = [
    "Redirect firmly toward fixing and tuning over adding anything new during polish time.",
    "Triage: Core Mechanics broken, then Sound/Animation, then Difficulty/Level 3, in that order if time is short.",
    "A confident presentation of a Level 1/2 game is a full success today - not every Level 3 item needs to be finished.",
    "Hold the 2-minute presentation limit firmly, with a warning at the halfway point.",
  ];
  notes.forEach((t, i) => {
    const y = 1.75 + i * 1.15;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.9, fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  footer(s, false, page++);
}

// ---------- SLIDE 13: Full Checklist Recap ----------
{
  const s = bgSlide(false);
  s.addText("The Full Capstone, In Review", { x: 0.6, y: 0.45, w: 11.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });

  const cats = [
    ["Core Mechanics", "Move, shoot, score, lives, win/lose, restart, waves (L2-11)"],
    ["Sound", "Shoot, explosion, wave clear, game over (L12)"],
    ["Animation", "Hit flash, explosion effect (L13)"],
    ["Difficulty", "Enemy count and speed scale by wave (L14)"],
  ];
  let cy = 1.5;
  cats.forEach(([title, desc]) => {
    s.addShape("roundRect", { x: 0.6, y: cy, w: 12.1, h: 1.05, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(title, { x: 0.9, y: cy + 0.1, w: 3.2, h: 0.85, valign: "middle", fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: 4.3, y: cy + 0.1, w: 8.2, h: 0.85, valign: "middle", fontFace: BODY_FONT, fontSize: 13, color: TEXT_DARK, margin: 0 });
    cy += 1.2;
  });

  s.addText("Plus: one chosen Level 3 stretch feature - shields, a power-up, a bonus enemy, formation movement, or a title/game-over screen.", {
    x: 0.6, y: 6.15, w: 12.1, h: 0.6, fontFace: BODY_FONT, fontSize: 13, italic: true, color: TEXT_MUTE, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 14: Semester 2 Preview ----------
{
  const s = bgSlide(false);
  s.addText("Looking Ahead: Semester 2", { x: 0.6, y: 0.45, w: 11.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/programming.png", x: 11.6, y: 0.4, w: 0.9, h: 0.9 });

  s.addShape("roundRect", { x: 0.6, y: 1.6, w: 12.1, h: 2.2, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("Semester 2 picks up in a very different place: state machines, enemy/AI behavior, physics, expanding scope with menus and save data, and eventually a much bigger capstone for the CoOp Expo. New this term: testing your own game (and having classmates try it) becomes an ongoing habit, not a one-time thing at the end.", {
    x: 0.9, y: 1.8, w: 11.5, h: 1.8, valign: "middle", fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.0, w: 12.1, h: 1.6, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("Everything you built this semester stays exactly this valuable: it's proof you can take an idea all the way to done.", {
    x: 0.9, y: 4.2, w: 11.5, h: 1.2, valign: "middle", fontFace: BODY_FONT, fontSize: 15, italic: true, bold: true, color: GOLD_L, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 15: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("Capstone Complete!", { x: 0, y: 2.45, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, margin: 0 });
  s.addText("Fifteen weeks, one idea, a real finished game - built and shown by you.", {
    x: 1.5, y: 3.35, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0,
  });
  s.addText("Congratulations, Video Game Creation - Semester 1 is in the books.", {
    x: 1.5, y: 4.0, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 15, color: CREAM, margin: 0,
  });
  s.addText("See you next semester for state machines, AI, and an even bigger build.", {
    x: 1.5, y: 4.7, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_15_Slides.pptx" }).then(() => console.log("done"));
