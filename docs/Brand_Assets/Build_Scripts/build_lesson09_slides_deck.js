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
const LESSON_LABEL = "Lesson 9";

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
  s.addText("Lesson 9", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Core Mechanics III: On the Scoreboard", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 9  -  Score Tracking + Real On-Screen UI", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, bold: true, color: GOLD_L, margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 9 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Add score tracking to the existing hit-detection code",
    "Build a persistent, always-visible HUD",
    "Explain push-based vs. poll-based display updates",
    "Recognize which engines' text tools were already the real UI",
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
    "GameMaker (GML script), for the live lecture demo - callback to Lesson 5",
    "No new assets this week - logic and UI lesson",
    "Unity, Unreal, 001GC: genuinely new UI setup this week",
    "Your own [YourName]Game project from Lesson 8 - not a new one",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Programming - UI/UX implementation, state display, event-driven vs. polling architecture (intermediate).", {
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
    ["0:00 - 0:05", "Welcome back, Lesson 8 recap"],
    ["0:05 - 0:15", "The Payoff of Deferring UI"],
    ["0:15 - 0:35", "Score: Reusing the Hit-Detection Path"],
    ["0:35 - 0:50", "Two Ways to Stay in Sync: Push vs. Poll"],
    ["0:50 - 0:55", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— On the Scoreboard —"],
    ["0:55 - 1:50", "Add score tracking and a real on-screen UI"],
    ["1:50 - 2:00", "Exit check / quick share-out"],
  ];

  let y = 1.5;
  rows.forEach(([time, seg]) => {
    const isDivider = time === "LAB TIME";
    const rowH = isDivider ? 0.48 : 0.6;
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
sectionDivider("book", "Welcome Back", "5 min", "Lesson 8 recap - your game got real stakes.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 8:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "A lives system - an escaped enemy costs a life, tracked starting at 3",
    "A real lose condition - \"Game Over\" appears at zero lives",
    "Today: score and lives finally show on screen the whole time you're playing",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one - and no new assets needed either, this is a logic and UI lesson.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  s.addNotes(`"Welcome back! Last week your game got real stakes - lives, and a lose condition. Quick show of hands - who tested both directions, shooting AND escaping, to make sure the two counters stayed separate?" [Pause, call on 1-2 volunteers]

"Today, your game finally gets a real scoreboard - not just a message at the end, but score and lives visible the whole time you're playing. Today's demo is also back in GameMaker - same engine as Lesson 5, but this time in its typed GML mode instead of GML Visual, so keep an eye out for how much of what you already learned there still applies."`);

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - The Payoff of Deferring UI ----------
sectionDivider("design", "The Payoff of Deferring UI", "10 min", "A promise made in Lesson 6, kept today.");

// ---------- SLIDE 7: Payoff content ----------
{
  const s = bgSlide(false);
  s.addText("Keeping a Promise", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addShape("roundRect", { x: 0.6, y: 1.5, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 1.75, w: 0.55, h: 0.55 });
  s.addText("Lesson 6: \"a plain line of text is completely legitimate today - the real UI system comes later, once there's a score and lives to actually display.\" That's today.", {
    x: 1.6, y: 1.65, w: 10.9, h: 1.2, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addText("For some engines, keeping that promise is almost free - the tool was already right, it just needs to run continuously. For a couple, today is genuinely new.", {
    x: 0.6, y: 3.25, w: 11.9, h: 0.75, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.2, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 4.45, w: 0.55, h: 0.55 });
  s.addText("Discussion: why wait until now to build a real UI, instead of doing it back in Lesson 6? (There was nothing worth displaying continuously yet - build the data first, display it once it means something.)", {
    x: 1.6, y: 4.35, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(`"Way back in Lesson 6, I told you something on purpose: 'a plain line of text is completely legitimate today - the real UI system comes later, once there's a score and lives to actually display.' That was a promise. Today we keep it."

"Here's the interesting part: for some of you, keeping that promise is almost free. Your engine's win/lose message tool was ALREADY the real tool - it just needs to run all the time instead of once. For a couple of you, today introduces something genuinely new."

[Discussion question] "Why do you think we waited until now to build a real UI, instead of doing it back in Lesson 6 alongside the win message?" (Draw out: there was nothing worth displaying continuously yet - a UI showing nothing changing isn't useful; build the data first, display it once it means something.)

[Timing checkpoint: wrap by 0:15]`);

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Score ----------
sectionDivider("programming", "Score: Reusing the Path", "20 min", "The hit-detection code already runs once per kill.");

// ---------- SLIDE 9: Score content ----------
{
  const s = bgSlide(false);
  s.addText("Score Joins the Kill", { x: 0.6, y: 0.45, w: 10.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("Every engine's hit-detection code already runs exactly once per confirmed kill - the same place enemies_remaining already changes. Score just joins it: +10 per kill.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.85, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.5, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/programming.png", x: 0.85, y: 2.75, w: 0.55, h: 0.55 });
  s.addText("The math is the easy part. GameMaker demo: obj_bullet's Collision event gets one new line - obj_gamemanager.__dnd_score += 10 - right alongside the enemies_remaining change already there.", {
    x: 1.6, y: 2.65, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  s.addText("Showing it is the real work this week - a new kind of separation: game logic vs. display.", {
    x: 0.6, y: 4.4, w: 11.9, h: 0.75, fontFace: BODY_FONT, fontSize: 14, italic: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 5.25, w: 12.1, h: 1.0, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("Display lives in obj_gamemanager's Draw GUI event - the same event already drawing \"You Win!\" and \"Game Over\" since Lesson 6.", {
    x: 0.9, y: 5.35, w: 11.5, h: 0.8, valign: "middle", fontFace: BODY_FONT, fontSize: 13, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(`"First, the easy part: score. Every engine's hit-detection code already runs exactly once per confirmed kill - the same place enemies_remaining already changes. Score just joins it."

[Demo cue: switch to GameMaker, open a project with working lives from Lesson 8]

"Watch: right in obj_bullet's Collision event with obj_enemy - the exact same event where enemies_remaining already goes down - I add one line: obj_gamemanager.__dnd_score += 10. That's the whole scoring system (__dnd_score is the same variable GameMaker's drag-and-drop Score blocks use, which is why it has that odd name). The hard part isn't the math, it's showing it." [Add the __dnd_score += 10 line, show it's not visible yet]

"To actually SHOW score and lives, watch where I put the display code: obj_gamemanager's Draw GUI event - the exact same event that's already been drawing 'You Win!' and 'Game Over' since Lesson 6. I add two plain draw_text() calls there, with no if-check around them at all, so they just run every single frame no matter what." [Add the two unconditional draw_text() calls, run the game, shoot an enemy and let one escape to show both numbers update live]

[Timing checkpoint: wrap by 0:35]`);

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - Push vs Poll ----------
sectionDivider("wrench", "Push vs. Poll", "15 min", "Two valid ways to keep a display in sync.");

// ---------- SLIDE 11: Push vs Poll content ----------
{
  const s = bgSlide(false);
  s.addText("Two Ways to Stay in Sync", { x: 0.6, y: 0.45, w: 10.8, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  pill(s, "15 min", 11.4, 0.5, 1.3, false);

  const cols = [
    ["PUSH", "Every time the data changes, call an update function to reflect it. Godot, GDevelop, Unreal, 001GC's real-time Field binding."],
    ["POLL", "Just re-check the data every single frame, no matter what. GameMaker's Draw GUI, TIC-80's print(), and this week's Unity HUD script."],
  ];
  let lx = 0.6;
  cols.forEach(([title, desc]) => {
    s.addShape("roundRect", { x: lx, y: 1.5, w: 5.9, h: 1.9, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(title, { x: lx + 0.25, y: 1.7, w: 5.4, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: lx + 0.25, y: 2.15, w: 5.4, h: 1.15, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
    lx += 6.2;
  });

  s.addText("Neither is \"more correct.\" Push means remembering to call the update everywhere data changes. Poll means never forgetting, at the cost of a little wasted work every frame.", {
    x: 0.6, y: 3.6, w: 11.9, h: 0.75, fontFace: BODY_FONT, fontSize: 13.5, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.5, w: 12.1, h: 1.85, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 4.72, w: 0.55, h: 0.55 });
  s.addText("Discussion: can you think of something in a game where checking every single frame would start to feel wasteful, if there were hundreds of things to check instead of two?", {
    x: 1.6, y: 4.62, w: 10.9, h: 1.65, valign: "middle", fontFace: BODY_FONT, fontSize: 13, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(`"What I just showed you in GameMaker is called POLL - instead of remembering to push an update every time something changes, you just check the data every single frame, no matter what. GameMaker's Draw GUI event and TIC-80's print() calls have quietly been doing this since Lesson 6 - today, Unity's HUD script does it on purpose too."

"But there's a second, equally valid way, and most of the rest of you will use it today: PUSH. Every time the data actually changes, you push an update to the display, instead of checking it constantly. Godot, GDevelop, Unreal, and 001 Game Creator's real-time Field binding all work this way."

[Demo cue, optional: pull up Godot briefly] "Neither approach is 'more correct.' Poll means you never forget to update, but you're technically doing a little more work than necessary, every frame, forever. Push means less wasted work, but you have to remember to call the update everywhere something changes. For two small numbers like these, the cost either way is nothing."

[Discussion question] "Can you think of something in a game where checking every single frame would start to feel wasteful, if there were hundreds of things to check instead of two?" (Open-ended; validate any reasonable guess - this previews ideas about performance they'll meet again in later, more advanced work.)

[Timing checkpoint: wrap by 0:50]`);

  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "5 min", "On the Scoreboard - score tracking and a real HUD.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("On the Scoreboard", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Destroying an enemy adds 10 points. Score and lives are both visible on screen, all the time - built on top of your existing project.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "Score +10 per kill; score and lives both always visible on screen.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "Readouts never show a stale or wrong number - always match reality.", "Everything works end to end."],
    ["Level 3 - Refined", "Style the HUD, add scoring variety, tighten logic/display separation.", "Ready to share."],
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
  s.addText("Unity, Unreal, 001 Game Creator: today's setup is more involved than usual - genuinely new UI tools. Everyone else: your engine already had the right tool since Lesson 6.", {
    x: 1.6, y: 4.82, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, color: CREAM, margin: 0,
  });

  s.addNotes(`"For lab today, open your own project from last week - not a new one, and no new assets needed. You're adding score tracking and a real, always-visible HUD."

[Hand out engine-specific handouts]

"Unity, Unreal, and 001 Game Creator folks - your setup today is more involved than usual, genuinely new UI tools. That's expected, not a sign you're behind. Everyone else - your engine's already had the right tool since Lesson 6, today's just about using it continuously."

"Questions before we start lab? ... Let's get your scoreboard up."`);

  footer(s, false, page++);
}

// ---------- SLIDE 14: Lab Time Notes ----------
{
  const s = bgSlide(false);
  s.addText("Lab Time Notes", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/wrench.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const notes = [
    "Unity, Unreal, and 001 Game Creator clusters need the most attention this week - budget extra circulation time, expect setup questions before logic questions.",
    "001 Game Creator: forgetting to tick \"Update Real-time\" on a Field looks like a mystery bug, not an obviously missing setting - check this first.",
    "Unity: NullReferenceException on Play almost always means a Text object wasn't dragged into the HUD script's Inspector fields.",
    "A score readout stuck at 0 is almost always a missing/misplaced increment, not a display bug - confirm the underlying variable first.",
  ];
  notes.forEach((t, i) => {
    const y = 1.75 + i * 1.15;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.9, fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addNotes(`Unity, Unreal, and 001 Game Creator clusters need the most attention this week - budget extra circulation time there, and expect setup-related questions before logic questions.

001 Game Creator: the single most likely mistake is forgetting to tick "Update Real-time" on a Field - it looks like a mysterious bug (frozen numbers) rather than an obviously missing setting. Check this first if a student's HUD isn't updating.

Unity: NullReferenceException on Play almost always means a Text object wasn't actually dragged into the HUD script's Inspector fields - check this first.

A score readout stuck at 0 is almost always a missing or misplaced increment, not a display bug - have students confirm the underlying variable is actually changing before troubleshooting the display.

Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout (HUD styling is a satisfying, visible one), not given new unrelated work.`);

  footer(s, false, page++);
}

// ---------- SLIDE 15: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("A Real Scoreboard!", { x: 0, y: 2.45, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, margin: 0 });
  s.addText("Score and lives, visible the whole time you play - a promise kept from Lesson 6.", {
    x: 1.5, y: 3.35, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0,
  });
  s.addText("Next week: game-over and restart flow - finally closing the loop Lesson 8 opened.", {
    x: 1.5, y: 4.0, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 15, color: CREAM, margin: 0,
  });
  s.addText("Try a Level 3 idea at home if you'd like - styling your HUD is a satisfying, visible one.", {
    x: 1.5, y: 4.7, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_09_Slides.pptx" }).then(() => console.log("done"));
