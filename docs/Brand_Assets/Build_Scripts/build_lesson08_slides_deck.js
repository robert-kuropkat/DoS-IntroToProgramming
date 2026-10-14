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
const LESSON_LABEL = "Lesson 8";

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
  s.addText("Lesson 8", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Core Mechanics II: Under Pressure", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 8  -  Lose Condition + Lives System", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, bold: true, color: GOLD_L, margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 8 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Reuse an existing code path to drive new, independent state",
    "Implement a lives counter and a lose condition",
    "Keep the win-counter and the lives-counter fully independent",
    "Explain why we're deliberately not building Game Over yet",
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
    "Unity, for the live lecture demo",
    "No new assets this week - logic-only lesson",
    "No engine-specific bugfix this week, unlike Lesson 7",
    "Your own [YourName]Game project from Lesson 7 - not a new one",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Programming - win/lose conditions, event-driven programming (intermediate); Production - deliberate feature sequencing.", {
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
    ["0:00 - 0:05", "Welcome back, Lesson 7 recap"],
    ["0:05 - 0:15", "Reversing the Question: What Happens When THEY Reach YOU?"],
    ["0:15 - 0:35", "Reusing What You Built: The Escape Path Gets Stakes"],
    ["0:35 - 0:50", "Two Counters, Never Crossed"],
    ["0:50 - 0:55", "Why We're Not Building Game Over (Yet)"],
    ["0:55 - 1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— Under Pressure —"],
    ["1:00 - 1:50", "Add a lives system and lose condition to your project"],
    ["1:50 - 2:00", "Exit check / quick share-out"],
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
sectionDivider("book", "Welcome Back", "5 min", "Lesson 7 recap - your enemies started moving.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 7:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Continuous movement - enemies advance toward the player, every frame",
    "Escape cleanup - an enemy reaching the bottom unshot is simply removed, no consequence",
    "Today: that \"no consequence\" changes. An escape starts to matter.",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one - and no new assets needed for the core exercise, this is a logic-only lesson. Level 3 sound files are provided in this week's Assets folder.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  s.addNotes(`"Welcome back! Last week your enemies started moving toward you. Quick show of hands - who tuned their enemy speeds, or gave them varied movement?" [Pause, call on 1-2 volunteers]

"So far, an enemy reaching the bottom of the screen has been a non-event - it just vanishes, no consequence. Today, that changes."`);

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - Reversing the Question ----------
sectionDivider("design", "Reversing the Question", "10 min", "What happens when an enemy reaches YOU?");

// ---------- SLIDE 7: Reversing the Question content ----------
{
  const s = bgSlide(false);
  s.addText("Win AND Lose", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("Every lesson so far has asked: what happens when YOUR bullet reaches an enemy? Today's question is the opposite: what happens when an ENEMY reaches you?", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.85, fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.55, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 2.8, w: 0.55, h: 0.55 });
  s.addText("Straight from the capstone's design document: Lose Condition - an enemy reaches the player's row, or the player runs out of lives. It's been on paper since Lesson 1 - today we build it.", {
    x: 1.6, y: 2.7, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.45, w: 12.1, h: 1.4, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 4.68, w: 0.55, h: 0.55 });
  s.addText("Discussion: why does a game need both a win condition AND a lose condition? (Stakes, tension, replayability - a game you can't lose isn't very interesting to keep playing.)", {
    x: 1.6, y: 4.58, w: 10.9, h: 1.2, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(`"Every lesson so far has asked one kind of question: what happens when YOUR bullet reaches an enemy? Today we ask the opposite question: what happens when an ENEMY reaches you?"

"Think back to your capstone's design document - the lose condition has always been there on paper: an enemy reaches the player's row, or the player runs out of lives. Today we actually build it."

[Discussion question] "Why do you think a game needs both a win condition AND a lose condition, instead of just one or the other?" (Draw out: stakes, tension, replayability - a game you can't lose isn't very interesting to keep playing.)

[Timing checkpoint: wrap by 0:15]`);

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Reusing What You Built ----------
sectionDivider("programming", "Reusing What You Built", "20 min", "The escape path from Lesson 7 gets real stakes.");

// ---------- SLIDE 9: Reusing content ----------
{
  const s = bgSlide(false);
  s.addText("The Escape Path Gets Stakes", { x: 0.6, y: 0.45, w: 10.8, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("You already wrote the code that detects an enemy reaching the bottom - you built it last week for cleanup. Today we don't write new detection. We just give it a consequence.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.85, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.5, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/programming.png", x: 0.85, y: 2.75, w: 0.55, h: 0.55 });
  s.addText("Unity demo: Enemy.cs's existing escape branch gains one new line - a call to GameManager.RegisterEscape() - right before the Destroy(gameObject) call that was already there.", {
    x: 1.6, y: 2.65, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addText("That small change isn't an accident - it's the payoff of building things in clean, separate pieces.", {
    x: 0.6, y: 4.4, w: 11.9, h: 0.55, fontFace: BODY_FONT, fontSize: 14.5, color: NAVY2, bold: true, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 5.1, w: 12.1, h: 1.15, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 5.3, w: 0.55, h: 0.55 });
  s.addText("Discussion: when has good organization made new work easier for you, in this project or elsewhere?", {
    x: 1.6, y: 5.22, w: 10.9, h: 0.9, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(`"Here's the fun part: you already wrote the code that detects an enemy reaching the bottom - you built it last week for cleanup. Today, we don't write new detection code. We just give that existing detection a real consequence."

[Demo cue: switch to Unity, open a project with working movement/escape from Lesson 7]

"Watch: here's Enemy.cs's escape branch from last week, in its Update() method - reaches the bottom, calls Destroy(gameObject), and until today, that was it. Today I add one line right before that Destroy() call: a call out to a new GameManager.RegisterEscape() method." [Add the RegisterEscape() call, show 'Game Over' appear after 3 escapes]

"Notice how small that change was. That's not an accident - it's the payoff of building things in clean, separate pieces. Last week's escape-cleanup code didn't know or care about win-counting, so today it was easy to give it a second job that also has nothing to do with win-counting."

[Discussion question] "Can you think of a time in your own project, or in something else you've built, where old work made new work easier because it was organized well?" (Open-ended; validate any real example, even a small one.)

[Timing checkpoint: wrap by 0:35]`);

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - Two Counters, Never Crossed ----------
sectionDivider("shield", "Two Counters, Never Crossed", "15 min", "Kills and lives stay completely independent.");

// ---------- SLIDE 11: Two Counters content ----------
{
  const s = bgSlide(false);
  s.addText("Two Counters, Never Crossed", { x: 0.6, y: 0.45, w: 10.8, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  pill(s, "15 min", 11.4, 0.5, 1.3, false);

  const cols = [
    ["Enemies Remaining", "Changes only when you shoot something. Drives the WIN condition."],
    ["Lives", "Changes only when something gets past you. Drives the LOSE condition."],
  ];
  let lx = 0.6;
  cols.forEach(([title, desc]) => {
    s.addShape("roundRect", { x: lx, y: 1.5, w: 5.9, h: 1.7, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(title, { x: lx + 0.25, y: 1.7, w: 5.4, h: 0.4, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: lx + 0.25, y: 2.15, w: 5.4, h: 0.95, fontFace: BODY_FONT, fontSize: 13, color: TEXT_DARK, margin: 0 });
    lx += 6.2;
  });

  s.addText("These two numbers should NEVER affect each other. Today, test both directions: shoot one (lives shouldn't move) and let one escape (kill counter shouldn't move).", {
    x: 0.6, y: 3.4, w: 11.9, h: 0.75, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.35, w: 12.1, h: 2.15, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 4.6, w: 0.6, h: 0.6 });
  s.addText("Why we're not building Game Over yet: when lives hits zero, \"Game Over\" appears - and the game just keeps running. Actually freezing everything and offering a restart button is real work, enough to get its own lesson (Lesson 10). Get the message right first; build the proper response later.", {
    x: 1.65, y: 4.5, w: 10.8, h: 1.9, valign: "middle", fontFace: BODY_FONT, fontSize: 13, color: CREAM, margin: 0,
  });

  s.addNotes(`"Your game now tracks two completely separate numbers: enemies_remaining (or your engine's equivalent), which only changes when you shoot something, and lives, which only changes when something gets past you. These numbers should NEVER affect each other."

[Demo cue: same Unity project] "Watch what happens if I shoot an enemy..." [shoot one] "...lives doesn't move. And if I let one escape..." [let one pass] "...enemiesRemaining doesn't move either. Two separate static methods on GameManager - RegisterKill() and RegisterEscape() - each touching only its own field, and neither one knows the other exists."

"This is worth testing explicitly in your own project today, in both directions - shoot one and confirm lives stays put, then let one escape and confirm enemies_remaining stays put. If either one moves when it shouldn't, that's this week's most common bug, and it's almost always the new code added to the wrong branch or event."

[Timing checkpoint: wrap by 0:50]

"One more thing before lab: when lives hits zero today, 'Game Over' appears on screen - and then the game just... keeps running. Enemies keep spawning, you can still move and shoot. That's on purpose."

"Actually stopping the game, freezing everything, and offering a restart button is real work - enough work that it gets its own lesson, Lesson 10. For today, just like Lesson 6's win message was 'just a message' before anything fancier got added, today's lose message is the same: get the detection and the message right first, build the proper response later."

[Timing checkpoint: wrap by 0:55]`);

  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "5 min", "Under Pressure - a lives system and lose condition.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("Under Pressure", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("An enemy reaching the bottom now costs a life. Lose all 3, and \"Game Over\" appears - built on top of your existing project.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "An escaped enemy costs a life (start at 3); \"Game Over\" appears at zero.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "Shooting never costs a life; escaping never counts as a kill.", "Everything works end to end."],
    ["Level 3 - Refined", "A live lives readout, a lose sound, and/or tighter code separation.", "Ready to share."],
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
  s.addText("Test both directions today: shoot an enemy (lives shouldn't move), then let one escape (your kill counter shouldn't move). Your handout's Level 2 section walks through exactly how to check this.", {
    x: 1.6, y: 4.82, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, color: CREAM, margin: 0,
  });

  s.addNotes(`"For lab today, open your own project from last week - not a new one, and no new assets needed. You're adding a lives counter and a lose message, both built on top of code you already trust. If you want to try the Level 3 sound idea, the sound files are already in this week's Assets folder - no hunting for your own."

[Hand out engine-specific handouts]

"Test in both directions today - shooting shouldn't touch lives, and escaping shouldn't touch your kill counter. Your handout's Level 2 section walks through exactly how to check this."

"Questions before we start lab? ... Let's put some pressure on your player."`);

  footer(s, false, page++);
}

// ---------- SLIDE 14: Lab Time Notes ----------
{
  const s = bgSlide(false);
  s.addText("Lab Time Notes", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/wrench.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const notes = [
    "Most important check: shoot an enemy (lives should NOT change) and separately let one escape (the kill counter should NOT change). Check both directions.",
    "Most likely mistake: the lives decrement added to the wrong event/branch - easy to make quickly, easy to miss without deliberately testing both directions.",
    "No engine needs a structural bugfix this week, unlike Lesson 7 - circulate normally, no cluster needs extra attention by default.",
    "A student who lets every enemy through will hit Game Over quickly - a good built-in test case, not a bug, worth pointing out if they're confused.",
  ];
  notes.forEach((t, i) => {
    const y = 1.75 + i * 1.15;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.9, fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addNotes(`The single most important thing to verify this week: shoot an enemy (lives should NOT change) and separately let one escape (the kill counter should NOT change). Check both directions, not just that the game "seems to work."

Most likely mistake: the lives decrement added to the wrong event/branch (hit-detection instead of escape-cleanup, or vice versa) - this is a copy-paste-adjacent mistake, easy to make quickly and easy to miss without deliberately testing both directions.

No engine needs a structural bugfix this week, unlike Lesson 7 - circulate normally, no cluster needs extra attention by default.

A student who lets every enemy through without shooting any will hit Game Over quickly - a good built-in test case, not a bug, worth pointing out if a student is confused why it happened so fast.

Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout (a live lives readout is a satisfying one), not given new unrelated work.`);

  footer(s, false, page++);
}

// ---------- SLIDE 15: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("Now There Are Stakes!", { x: 0, y: 2.45, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, margin: 0 });
  s.addText("A lives system and a real lose condition, built almost entirely by reusing last week's work.", {
    x: 1.5, y: 3.35, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0,
  });
  s.addText("Next week: score tracking and a basic on-screen UI - showing your stats properly for the first time.", {
    x: 1.5, y: 4.0, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 15, color: CREAM, margin: 0,
  });
  s.addText("Try a Level 3 idea at home if you'd like - a live lives readout is a satisfying one to add.", {
    x: 1.5, y: 4.7, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_08_Slides.pptx" }).then(() => console.log("done"));
