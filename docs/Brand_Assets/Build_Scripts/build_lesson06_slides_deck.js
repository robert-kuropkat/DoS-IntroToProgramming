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
const LESSON_LABEL = "Lesson 6";

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
  s.addText("Lesson 6", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Programming Basics V: Victory Condition", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 6  -  Level 1 MVP Milestone", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, bold: true, color: GOLD_L, margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 6 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Explain what a counter is and how it tracks state",
    "Explain why simple text tools are enough for now",
    "Explain what the Level 1 MVP milestone means",
    "Implement a working win condition",
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
    "Scratch, for the live lecture demo",
    "Engine-specific Student Handouts, one per student",
    "No new assets this week - \"You Win!\" uses each engine's basic text tools",
    "Your own [YourName]Game project from Lesson 5 - not a new one",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Programming - state tracking, win/lose conditions, event-driven programming (introductory).", {
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
    ["0:00 - 0:05", "Welcome back, Lesson 5 recap, quick share-out"],
    ["0:05 - 0:25", "Tracking State: Counting Down to Zero"],
    ["0:25 - 0:45", "Announcing the Win: Each Engine's Simplest Text Tool"],
    ["0:45 - 0:50", "Milestone Moment: What \"Level 1 MVP\" Actually Means"],
    ["0:50 - 1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— Victory Condition —"],
    ["1:00 - 1:50", "Add a win condition to your Semester 1 project"],
    ["1:50 - 2:00", "Exit check: milestone celebration share-out"],
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
sectionDivider("book", "Welcome Back", "5 min", "Lesson 5 recap and a quick share of enemy rows.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 5:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Loops - a whole row of enemies, created from one block of code",
    "Every enemy independently destructible, thanks to good structure from Lesson 4",
    "Today: a real win condition - the final piece of a complete, playable game loop",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one - and no new assets needed either, everything builds on Lesson 5's enemy row.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  s.addNotes(`"Welcome back! Last week you turned one enemy into a whole row of them. Quick show of hands - who got every enemy independently destructible? Anyone try the Level 3 visual variety idea?" [Pause, call on 1-2 volunteers]

"Today's a big one. By the end of class, your game will actually be winnable - shoot every enemy, and the game tells you so. This finishes the very first real milestone of your whole project."`);

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - Tracking State ----------
sectionDivider("programming", "Tracking State", "20 min", "How does the game know the exact moment every enemy is gone?");

// ---------- SLIDE 7: Tracking State content ----------
{
  const s = bgSlide(false);
  s.addText("Tracking State", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("A counter is a single number that starts at 5 and counts down by one every time an enemy is destroyed. Hits zero? All enemies are gone.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.7, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0,
  });
  s.addText("One of the most common patterns in all of programming, not just games.", {
    x: 0.6, y: 2.1, w: 11.9, h: 0.5, fontFace: BODY_FONT, fontSize: 14, italic: true, color: TEXT_MUTE, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.85, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 3.1, w: 0.55, h: 0.55 });
  s.addText("enemies_left starts at 5. Every enemy destroyed subtracts one. When it reaches zero - and only then - the win condition is met, with no need to check anything else.", {
    x: 1.6, y: 3.0, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.1, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 4.95, w: 0.55, h: 0.55 });
  s.addText("Discussion: what other things in a game might need a counter like this?", {
    x: 1.6, y: 4.9, w: 10.9, h: 0.75, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(`"Quick question: your game has 5 enemies right now. How would you know, in code, the exact moment all 5 are gone?" [Pause for guesses - someone may suggest checking a list's length, or counting]

"Both of those work! Today we'll use the simpler of the two: a single number that starts at 5 and counts down by one every time an enemy is destroyed. The moment that number hits zero, we know instantly - no need to check anything else. This kind of tracked number is usually called a counter, and it's one of the most common patterns in all of programming, not just games."

[Demo cue: switch to Scratch, open a project with a working enemy row from Lesson 5]

"Watch: I've got a variable called enemies_left, starting at 5. It's set 'for all sprites', not just this one, because both the Enemy clones and the win check need to see the same shared number. Every time a clone's touching-check destroys one, it changes enemies_left by -1. When it hits zero..." [Destroy all enemies on screen] "...this broadcast fires and the message appears. That's the whole idea - one shared number, counted down by one event at a time."

[Discussion question] "What other things in a game might need a counter like this?" (Examples to draw out if needed: lives remaining, ammo left, score, a timer counting down.)

[Timing checkpoint: wrap by 0:25]`);

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Announcing the Win ----------
sectionDivider("design", "Announcing the Win", "20 min", "Each engine's simplest way to just draw text on screen.");

// ---------- SLIDE 9: Announcing the Win content ----------
{
  const s = bgSlide(false);
  s.addText("Announcing the Win", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("Today, a plain line of text is completely legitimate - the real UI system comes later, once there's a score and lives to actually display.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.7, fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0,
  });

  s.addText("Each engine's simplest built-in text tool:", { x: 0.6, y: 2.3, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0 });
  const rows = [
    ["Godot", "a Label node"],
    ["Unity", "GUI.Label() inside OnGUI()"],
    ["GameMaker", "draw_text()"],
    ["GDevelop", "a Text object, shown on demand"],
    ["TIC-80", "print()"],
    ["001 Game Creator", "Message Box"],
    ["Unreal", "Print String (Print to Screen)"],
  ];
  let ry = 2.75, colStart = 0.6;
  rows.forEach(([eng, desc], i) => {
    const col = i < 4 ? 0 : 1;
    const x = colStart + col * 6.2;
    const y = 2.75 + (i % 4) * 0.5;
    s.addText(eng, { x, y, w: 2.6, h: 0.42, fontFace: BODY_FONT, fontSize: 12.5, bold: true, color: NAVY, margin: 0 });
    s.addText(desc, { x: x + 2.6, y, w: 3.5, h: 0.42, fontFace: "Courier New", fontSize: 11.5, color: TEXT_DARK, margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 5.05, w: 12.1, h: 1.4, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 5.28, w: 0.55, h: 0.55 });
  s.addText("Discussion: what other times is the \"quick and ugly\" version actually the right first step? (Placeholder art, a console log instead of a UI, a text menu instead of a styled one.)", {
    x: 1.6, y: 5.2, w: 10.9, h: 1.2, valign: "middle", fontFace: BODY_FONT, fontSize: 13, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(`"Now that we know WHEN we've won, how do we tell the player? Every engine has some basic way to just draw text on screen, and that's genuinely all we need today - not a fancy menu, not a scoreboard, just words appearing."

[Demo cue: same Scratch project] "In Scratch, that's just a 'say' block - a speech bubble so simple it's barely UI at all. Other engines are just as simple: Godot has a Label node, GameMaker has draw_text(), TIC-80 has print(), Unreal has a node literally called Print String. None of these need the real UI systems your games will eventually have - that's coming properly in a few weeks, once you've got a score and lives to actually display."

"The lesson here: don't reach for the fanciest tool before you need it. A plain line of text is a completely legitimate way to confirm a win condition works, and it's exactly what a real developer would build first too, before ever touching final art or layout."

[Discussion question] "Can you think of other times in a project where the 'quick and ugly' version is actually the right first step?" (Examples: placeholder art, a console log instead of a UI, a text menu instead of a styled one.)

[Timing checkpoint: wrap by 0:43-0:45]`);

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - Milestone Moment ----------
sectionDivider("shield", "Milestone Moment", "5 min", "What finishing the Level 1 MVP actually means.");

// ---------- SLIDE 11: Milestone Moment content ----------
{
  const s = bgSlide(false);
  s.addText("Level 1 MVP", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("MVP = Minimum Viable Product: the smallest version of something that's still a real, complete, playable experience.", { x: 0.6, y: 1.5, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });

  s.addShape("roundRect", { x: 0.6, y: 2.35, w: 12.1, h: 2.0, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("Once your win condition works today, your game has:", { x: 0.9, y: 2.55, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: NAVY2, margin: 0 });
  const items = ["Movement", "Shooting", "Multiple enemies", "A win condition"];
  items.forEach((t, i) => {
    s.addImage({ path: "icons/check.png", x: 0.9 + (i % 2) * 5.9, y: 3.0 + Math.floor(i / 2) * 0.6, w: 0.38, h: 0.38 });
    s.addText(t, { x: 1.35 + (i % 2) * 5.9, y: 2.98 + Math.floor(i / 2) * 0.6, w: 5.0, h: 0.42, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.55, w: 12.1, h: 1.9, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("That's a complete, playable core loop - not impressive yet, but genuinely done, start to finish. Everything through Lesson 15 makes this same loop more complete, then more polished.", {
    x: 0.9, y: 4.75, w: 11.5, h: 1.5, valign: "middle", fontFace: BODY_FONT, fontSize: 15, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(`"Here's something worth pausing on. Once your win condition works today, you will have a complete, playable game loop: move, shoot, face enemies, and win. That's not a small thing - that's the Level 1 MVP milestone from your capstone's design document, and it's the first of several big checkpoints across the whole year."

"MVP stands for Minimum Viable Product - the smallest version of something that's still a real, complete, playable experience. It's not supposed to be impressive yet. It's supposed to work, start to finish, which is genuinely the hardest part of finishing any project. Everything from here through Lesson 15 is about making this same core loop more complete, then more polished - but the loop itself, today, is done."

[Timing checkpoint: wrap by 0:50]`);

  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "10 min", "Victory Condition - finishing your Level 1 MVP milestone.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("Victory Condition", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("\"You Win!\" appears the instant the last enemy is destroyed - built on top of your existing project.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "\"You Win!\" appears when the last enemy is destroyed.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "The message appears exactly once, at exactly the right time.", "Everything works end to end."],
    ["Level 3 - Refined", "Add a victory sound and/or a more celebratory message.", "Ready to share."],
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
  s.addText("Heads up: today's win-tracking code lives in a slightly different place than last week's hit-detection - your handout walks through exactly where for your engine.", {
    x: 1.6, y: 4.82, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, color: CREAM, margin: 0,
  });

  s.addNotes(`"For lab today, open your own project from last week - not a new one, and no new assets needed. Everything today builds directly on your enemy row from Lesson 5."

"The goal: track enemies remaining, and show 'You Win!' the instant it hits zero. Level 1 is getting that working once. Level 2 is confirming it fires exactly once, at exactly the right moment. Level 3 is turning a plain message into an actual celebration - sound, animation, whatever you want."

[Hand out engine-specific handouts]

"One heads-up: today's win-tracking code needs to live in a slightly different place than last week's hit-detection - your handout walks through exactly where for your engine."

"Questions before we start lab? ... Let's go finish your first milestone."`);

  footer(s, false, page++);
}

// ---------- SLIDE 14: Lab Time Notes ----------
{
  const s = bgSlide(false);
  s.addText("Lab Time Notes", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/wrench.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const notes = [
    "Most common bug this week: an off-by-one error - the starting counter not matching the actual number of enemies spawned. Check the Lesson 5 spawn count first.",
    "Second most common: win-tracking code placed somewhere that runs every frame instead of only when an enemy is destroyed - can cause flickering or unpredictable behavior.",
    "No engine is especially harder this week, unlike Lesson 5's Scratch wrinkle - circulate normally.",
    "At 1:50, do a real celebration share-out - this is the Level 1 MVP milestone. Consider 3-4 students demoing a full win from start to finish.",
  ];
  notes.forEach((t, i) => {
    const y = 1.75 + i * 1.15;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.9, fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addNotes(`The single most common bug this week: an off-by-one error - the starting counter not matching the actual number of enemies spawned. Have students double-check their Lesson 5 spawn count first.

Second most common: win-tracking code placed somewhere that runs every frame instead of only when an enemy is actually destroyed - this can cause the message to flicker or the counter to behave unpredictably.

Circulate by engine cluster as usual - no engine is especially harder this week, unlike Lesson 5's Scratch wrinkle.

Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout, not given new unrelated work.

At 1:50, call time and do a real celebration share-out - this is the Level 1 MVP milestone, worth more than the usual quick show-and-tell. Consider having 3-4 students demo a full win from start to finish.`);

  footer(s, false, page++);
}

// ---------- SLIDE 15: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("Milestone Complete!", { x: 0, y: 2.45, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, margin: 0 });
  s.addText("Level 1 MVP: move, shoot, face enemies, and win - a complete, playable game loop.", {
    x: 1.5, y: 3.35, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 17, color: CREAM, margin: 0,
  });
  s.addText("Next week: enemies that advance toward the player - the first step into Level 2.", {
    x: 1.5, y: 4.0, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 15, color: CREAM, margin: 0,
  });
  s.addText("Keep polishing your win moment at home if you'd like - try one of the Level 3 ideas if you haven't already.", {
    x: 1.5, y: 4.7, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_06_Slides.pptx" }).then(() => console.log("done"));
