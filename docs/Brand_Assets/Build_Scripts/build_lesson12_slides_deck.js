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
const LESSON_LABEL = "Lesson 12";

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
  s.addText("Lesson 12", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Hearing It Happen", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 12  -  Sound Effects", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, bold: true, color: GOLD_L, margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 12 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Wire four sound cues into their exact matching code paths",
    "Recognize why event-driven triggers need no replay guard",
    "Formalize five lessons' worth of optional sound ideas",
    "Play a full round and hear every outcome distinctly",
  ];
  objectives.forEach((t, i) => {
    const y = 2.15 + i * 0.62;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.15, y: y - 0.03, w: 6.1, h: 0.55, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 7.6, y: 1.65, w: 5.15, h: 4.6, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("What We're Using Today", { x: 7.9, y: 1.85, w: 4.6, h: 0.4, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0 });
  const materials = [
    "Demo laptop with all 8 engines, audio output working",
    "Godot, for the live lecture demo",
    "Two new sound assets: sfx_wave.wav, sfx_lose.wav",
    "Everything else reuses Lessons 2-6's optional sound ideas",
    "Your own [YourName]Game project from Lesson 11 - not a new one",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Audio - sound design fundamentals (introductory); Programming - event-driven triggers, code reuse (intermediate).", {
    x: 0.6, y: 6.55, w: 12.1, h: 0.5, fontFace: BODY_FONT, fontSize: 11.5, italic: true, color: TEXT_MUTE, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 3: TODAY'S AGENDA ----------
{
  const s = bgSlide(false);
  s.addText("Today's Agenda", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/audio.png", x: 11.6, y: 0.4, w: 0.9, h: 0.9 });

  const rows = [
    ["0:00 - 0:05", "Welcome back, Lesson 11 recap, the polish phase begins"],
    ["0:05 - 0:15", "The Sound Payoff: Four Sounds, Four Existing Moments"],
    ["0:15 - 0:35", "One Shared Sound for Three Outcomes"],
    ["0:35 - 0:45", "Why None of Today's Sounds Need a Retrigger Guard"],
    ["0:45 - 0:55", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— Hearing It Happen —"],
    ["0:55 - 1:50", "Wire up all four sounds, test a full round by ear"],
    ["1:50 - 2:00", "Exit check: play a round for the class"],
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
sectionDivider("book", "Welcome Back", "5 min", "Level 2 Complete is done. The polish phase begins.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 11:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Level 2 Complete reached - a real, replayable game with multiple waves",
    "No new mechanics from here through Lesson 15 - just making the loop feel good",
    "Today: sound. Next: animation, difficulty scaling, and a final polish pass",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/audio.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one. Two new sound files are provided - sfx_wave.wav and sfx_lose.wav - everything else may already be in your project from earlier optional ideas.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - The Sound Payoff ----------
sectionDivider("design", "The Sound Payoff", "10 min", "Four sounds, four moments you already built.");

// ---------- SLIDE 7: Sound Payoff content ----------
{
  const s = bgSlide(false);
  s.addText("Four Sounds, Four Existing Moments", { x: 0.6, y: 0.45, w: 10.8, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addShape("roundRect", { x: 0.6, y: 1.5, w: 12.1, h: 1.6, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 1.75, w: 0.55, h: 0.55 });
  s.addText("Since Lesson 2, five lessons quietly offered an optional 'add a sound' idea. Today, four sounds become required for everyone - shoot, explosion, wave clear, and game over - matching the capstone's Level 3 feature list exactly.", {
    x: 1.6, y: 1.65, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addText("No new logic today - just one more line at each of four places you already built.", {
    x: 0.6, y: 3.35, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.15, w: 12.1, h: 1.85, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  const sounds = ["Shoot - when you fire", "Explosion - when a hit lands", "Wave Clear - a non-final wave ends", "Game Over - lives reach zero"];
  sounds.forEach((t, i) => {
    s.addImage({ path: "icons/check.png", x: 0.9 + (i % 2) * 5.9, y: 4.4 + Math.floor(i / 2) * 0.7, w: 0.36, h: 0.36 });
    s.addText(t, { x: 1.32 + (i % 2) * 5.9, y: 4.38 + Math.floor(i / 2) * 0.7, w: 5.4, h: 0.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: CREAM, margin: 0 });
  });

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - One Shared Sound ----------
sectionDivider("programming", "One Shared Sound", "20 min", "Three outcomes, one player, three interchangeable clips.");

// ---------- SLIDE 9: One Shared Sound content ----------
{
  const s = bgSlide(false);
  s.addText("Three Outcomes, One Player", { x: 0.6, y: 0.45, w: 10.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("Shoot and hit sounds are easy - one sound each, already wired since Lessons 3-4. The interesting part: win, wave-clear, and lose can never happen at the same time, so they can share ONE sound player.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.9, fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.55, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/programming.png", x: 0.85, y: 2.8, w: 0.55, h: 0.55 });
  s.addText("Godot demo: one AudioStreamPlayer2D, RoundSound, added to the main scene. In the exact branches Lesson 11 already built - true-win, wave-advance, and Lesson 8's lose case - set RoundSound's stream and call play().", {
    x: 1.6, y: 2.7, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.45, w: 12.1, h: 1.4, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addText("One node, three interchangeable clips - the same reuse idea from every lesson this semester, just applied to a sound resource instead of a function. Unity: the same idea with PlayOneShot() and one shared AudioSource.", {
    x: 0.9, y: 4.65, w: 11.5, h: 1.0, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - Why No Retrigger Guard ----------
sectionDivider("wrench", "No Retrigger Guard Needed", "10 min", "Why today's four sounds are safe by construction.");

// ---------- SLIDE 11: Retrigger Guard content ----------
{
  const s = bgSlide(false);
  s.addText("Why No Retrigger Guard?", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addShape("roundRect", { x: 0.6, y: 1.5, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 1.75, w: 0.55, h: 0.55 });
  s.addText("Lesson 2's beep-beep needed an explicit stop call because it was continuous. Today's four sounds are one-shot - they finish on their own. But there's a second reason they're safe: every one lives inside a branch that only runs once per actual event.", {
    x: 1.6, y: 1.65, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 3.4, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 3.63, w: 0.55, h: 0.55 });
  s.addText("TIC-80 note: Lesson 6's optional win sound may have needed a hasWon guard flag back then, since the check ran every frame. Lesson 10/11's restructuring moved it into a one-time hit branch - that old guard flag can be removed now if it's still there.", {
    x: 1.6, y: 3.55, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 13, italic: true, color: CREAM, margin: 0,
  });

  s.addText("Good structure, built for one reason, quietly solves a second problem later.", {
    x: 0.6, y: 5.3, w: 12.1, h: 0.5, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "10 min", "Hearing It Happen - wiring up four sounds.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("Hearing It Happen", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("Wire four sounds into the exact branches that already represent their events - built on top of your existing project.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "All four sounds play at the correct moment, at least once each.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "All four fire reliably every time, none interfere with each other.", "Everything works end to end."],
    ["Level 3 - Refined", "Add the optional win sound, pitch/volume variety, or tidy the code.", "Ready to share."],
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
  s.addImage({ path: "icons/audio.png", x: 0.85, y: 4.92, w: 0.55, h: 0.55 });
  s.addText("Play a full round OUT LOUD today - a shot, a hit, an early wave clear, and separately a loss - and confirm all four sound genuinely different from each other.", {
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
    "Remind students who skipped the optional Lessons 2/3/4/6 ideas that today's Handout gives them full instructions, not just \"add two sounds.\"",
    "Watch for a sound call placed outside the correct if/else branch - causes two round-ending sounds to play together on the same frame.",
    "Listen as you circulate this week, not just watch the screen - a wrong or missing sound is often faster to catch by ear than by reading code.",
    "This is a lighter week - if a cluster finishes early, point them at Level 3 ideas rather than moving ahead to next week's material.",
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
  s.addText("It Sounds Like a Game!", { x: 0, y: 2.45, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, margin: 0 });
  s.addText("A shot, an explosion, a wave clear, a game over - four moments, four distinct cues.", {
    x: 1.5, y: 3.35, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0,
  });
  s.addText("Next week: basic animation - an explosion effect, a hit flash/reaction.", {
    x: 1.5, y: 4.0, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 15, color: CREAM, margin: 0,
  });
  s.addText("Keep playtesting a full round at home if you'd like - try a Level 3 sound idea if you haven't already.", {
    x: 1.5, y: 4.7, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_12_Slides.pptx" }).then(() => console.log("done"));
