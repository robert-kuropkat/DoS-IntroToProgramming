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
const LESSON_LABEL = "Lesson 2";

function bgSlide(dark, darkImg) {
  // See generate_deck.js (Lesson 1) for why slide.background is deliberately
  // never set - a full-bleed addImage() covers the whole canvas and avoids
  // a pptxgenjs 4.0.1 Content_Types.xml defect. Same approach here.
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
  s.addText("Lesson 2", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Programming Basics I: Sequences, Events & the Game Loop", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Thursday, August 20, 2026", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, color: "9FB6D6", margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 2 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Explain sequences and why order matters",
    "Explain events and give a game example",
    "Describe the game loop concept",
    "Build a working Key Press Reactor - the start of the Semester 1 project",
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
    "This lesson's Assets folder (placeholder sprite + sound - no download needed)",
    "Your own Game Design Document, to start filling in during lab",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Programming - control structures and event-driven programming (introductory).", {
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
    ["0:00 - 0:05", "Welcome back, Lesson 1 recap, quick install spot-check"],
    ["0:05 - 0:20", "Sequences (live demo in Scratch)"],
    ["0:20 - 0:35", "Events (live demo in Scratch)"],
    ["0:35 - 0:50", "The Game Loop (concept + \"under the hood\" look)"],
    ["0:50 - 1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— Key Press Reactor —"],
    ["1:00 - 1:50", "Build the Key Press Reactor - the start of your Semester 1 project"],
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
      fontSize: isDivider ? 13 : 13.5, bold: isDivider, italic: isDivider,
      color: isDivider ? NAVY : TEXT_DARK, margin: 0,
    });
    y += rowH + 0.11;
  });

  footer(s, false, page++);
}

// ---------- SLIDE 4: SECTION - Welcome Back ----------
sectionDivider("book", "Welcome Back", "5 min", "Lesson 1 recap and a quick install spot-check before we dive in.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 1:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Internet safety, what game development involves, and a tour of all 8 engines",
    "Everyone chose (or shortlisted) an engine and started installing it",
    "Today: our first real programming lesson - and the start of your Semester 1 project",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("If you didn't finish installing last week, that's okay - we've got a few minutes now, and Scratch and GDevelop work straight in a browser with zero install if you're in a pinch.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - Sequences ----------
sectionDivider("programming", "Sequences", "15 min", "The simplest idea in programming: instructions that happen one after another, in order.");

// ---------- SLIDE 7: Sequences content ----------
{
  const s = bgSlide(false);
  s.addText("Sequences", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "15 min", 11.4, 0.5, 1.3, false);

  s.addText("A sequence is a list of instructions that happen one after another, in order.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0,
  });
  s.addText("Think about your morning routine - wake up, get dressed, eat breakfast. If you get dressed before you wake up, that doesn't work. Order matters.", {
    x: 0.6, y: 2.05, w: 11.9, h: 0.7, fontFace: BODY_FONT, fontSize: 14, italic: true, color: TEXT_MUTE, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.95, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("Live demo (Scratch):", { x: 0.85, y: 3.15, w: 4, h: 0.4, fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY2, margin: 0 });
  s.addText("move 10 steps  →  turn 15 degrees  →  say \"Hello!\"", {
    x: 0.85, y: 3.6, w: 11.5, h: 0.5, fontFace: "Courier New", fontSize: 16, bold: false, color: NAVY, margin: 0,
  });
  s.addText("Stacked blocks run top to bottom, one after another - every language we use this year (GDScript, C#, GML...) runs sequences like this at its core.", {
    x: 0.85, y: 4.05, w: 11.5, h: 0.35, fontFace: BODY_FONT, fontSize: 12, italic: true, color: TEXT_MUTE, margin: 0,
  });

  s.addImage({ path: "icons/chat.png", x: 0.6, y: 4.85, w: 0.55, h: 0.55 });
  s.addText("Discussion: what's something else in your daily life that has to happen in a specific order to work?", {
    x: 1.3, y: 4.85, w: 11.4, h: 0.7, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, italic: true, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - Events ----------
sectionDivider("controller", "Events", "15 min", "Code that says \"when this happens, do that\" - the shift from running once to reacting.");

// ---------- SLIDE 9: Events content ----------
{
  const s = bgSlide(false);
  s.addText("Events", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "15 min", 11.4, 0.5, 1.3, false);

  s.addText("An event is code that waits, then reacts: \"when this happens, do that.\"", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.6, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.15, w: 12.1, h: 1.35, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("Live demo (Scratch):", { x: 0.85, y: 2.32, w: 4, h: 0.4, fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY2, margin: 0 });
  s.addText("when [right arrow] key pressed  →  move 10 steps", {
    x: 0.85, y: 2.72, w: 11.5, h: 0.5, fontFace: "Courier New", fontSize: 15, bold: false, color: NAVY, margin: 0,
  });
  s.addText("Nothing happens until the key is pressed - and each press runs the whole sequence underneath it again.", {
    x: 0.85, y: 3.2, w: 11.5, h: 0.3, fontFace: BODY_FONT, fontSize: 12, italic: true, color: TEXT_MUTE, margin: 0,
  });

  s.addText("Same idea, different names, across every engine:", { x: 0.6, y: 3.75, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0 });
  const rows = [
    ["Scratch", "\"when [key] pressed\" hat block"],
    ["Godot", "signals, or checking Input inside _process()"],
    ["Unity", "Keyboard.current checks inside Update()"],
    ["GameMaker", "the Key Pressed / Key Down event"],
  ];
  let ry = 4.2;
  rows.forEach(([eng, desc]) => {
    s.addText(eng, { x: 0.6, y: ry, w: 2.2, h: 0.42, fontFace: BODY_FONT, fontSize: 13, bold: true, color: NAVY, margin: 0 });
    s.addText(desc, { x: 2.9, y: ry, w: 9.6, h: 0.42, fontFace: BODY_FONT, fontSize: 13, color: TEXT_DARK, margin: 0 });
    ry += 0.44;
  });

  s.addImage({ path: "icons/chat.png", x: 9.6, y: 3.75, w: 0.5, h: 0.5 });
  s.addText("What other events can you think of in games you've played?", {
    x: 9.6, y: 4.3, w: 3.1, h: 1.5, fontFace: BODY_FONT, fontSize: 12.5, italic: true, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 10: Sequences vs. Events diagram ----------
{
  const s = bgSlide(false);
  s.addText("Sequences vs. Events - side by side", { x: 0.6, y: 0.45, w: 11, h: 0.6, fontFace: HEAD_FONT, fontSize: 26, bold: true, color: NAVY2, margin: 0 });
  const w = 8.75, h = w / (2000 / 1280);
  s.addImage({ path: "diagram_sequences_events.png", x: (SW - w) / 2, y: 1.25, w, h });
  footer(s, false, page++);
}

// ---------- SLIDE 11: SECTION - The Game Loop ----------
sectionDivider("wrench", "The Game Loop", "15 min", "How does a character keep moving smoothly while you hold a key? This is why.");

// ---------- SLIDE 12: Game Loop content ----------
{
  const s = bgSlide(false);
  s.addText("The Game Loop", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "15 min", 11.4, 0.5, 1.3, false);

  s.addText("The engine isn't waiting for one event and stopping - it's constantly asking, ~30 to 60 times every second:", {
    x: 0.6, y: 1.5, w: 11.9, h: 0.7, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0,
  });

  const asks = ["What's the current input?", "What needs to move?", "What needs to be redrawn?"];
  let ax = 0.6;
  asks.forEach(t => {
    s.addShape("roundRect", { x: ax, y: 2.3, w: 3.9, h: 0.9, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(t, { x: ax + 0.15, y: 2.3, w: 3.6, h: 0.9, valign: "middle", align: "center", fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY, margin: 0 });
    ax += 4.1;
  });

  s.addText("That loop - check, update, draw, check, update, draw - happening dozens of times a second, is the game loop. You don't usually write it yourself; the engine runs it for you. Almost everything you build this year runs inside it.", {
    x: 0.6, y: 3.55, w: 11.9, h: 1.0, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.4, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/book.png", x: 0.85, y: 4.98, w: 0.55, h: 0.55 });
  s.addText("Flip-book analogy: each page is slightly different, and flipping through them fast makes it look like smooth motion. Each \"page\" is one pass through the game loop.", {
    x: 1.6, y: 4.9, w: 10.9, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: TEXT_DARK, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 13: The Game Loop diagram ----------
{
  const s = bgSlide(false);
  s.addText("The Game Loop - Input, Update, Render", { x: 0.6, y: 0.45, w: 11, h: 0.6, fontFace: HEAD_FONT, fontSize: 26, bold: true, color: NAVY2, margin: 0 });
  const w = 9.0, h = w / (2000 / 1240);
  s.addImage({ path: "diagram_game_loop.png", x: (SW - w) / 2, y: 1.25, w, h });
  footer(s, false, page++);
}

// ---------- SLIDE 14: BONUS - Thinking in Diagrams intro ----------
{
  const s = bgSlide(false);
  s.addText("Bonus: Diagrams Programmers Use", { x: 0.6, y: 0.45, w: 10.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "BONUS", 11.4, 0.5, 1.3, false);

  s.addText("You just described the Game Loop in words. Professional game and software developers often draw it instead - using a couple of simple, standard diagram styles from UML (Unified Modeling Language).", {
    x: 0.6, y: 1.5, w: 11.9, h: 0.8, fontFace: BODY_FONT, fontSize: 16, color: TEXT_DARK, margin: 0,
  });

  const cards = [
    ["icons/programming.png", "Activity Diagram", "Shows the order things happen in - like a flowchart. Next up: your Key Press Reactor as one."],
    ["icons/controller.png", "State Diagram", "Shows the different \"modes\" something can be in, and what switches it between them - like Idle vs. Moving."],
  ];
  let cx = 0.6;
  cards.forEach(([icon, title, desc]) => {
    s.addShape("roundRect", { x: cx, y: 2.6, w: 5.85, h: 2.9, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
    s.addImage({ path: icon, x: cx + 0.35, y: 2.9, w: 0.7, h: 0.7 });
    s.addText(title, { x: cx + 1.25, y: 2.9, w: 4.2, h: 0.7, valign: "middle", fontFace: BODY_FONT, fontSize: 18, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: cx + 0.35, y: 3.75, w: 5.15, h: 1.6, fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0 });
    cx += 6.25;
  });

  s.addText("We'll come back to a third kind - Class Diagrams - later in the year, once you're writing more object-oriented code.", {
    x: 0.6, y: 5.9, w: 11.9, h: 0.5, fontFace: BODY_FONT, fontSize: 13, italic: true, color: TEXT_MUTE, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 15: Activity Diagram ----------
{
  const s = bgSlide(false);
  s.addText("Activity Diagram - Key Press Reactor", { x: 0.6, y: 0.45, w: 11, h: 0.6, fontFace: HEAD_FONT, fontSize: 26, bold: true, color: NAVY2, margin: 0 });
  pill(s, "BONUS", 11.4, 0.5, 1.3, false);
  const w1 = 9.0, h1 = w1 / (2000 / 1240);
  s.addImage({ path: "diagram_activity_keypress.png", x: (SW - w1) / 2, y: 1.2, w: w1, h: h1 });
  footer(s, false, page++);
}

// ---------- SLIDE 16: State Diagram ----------
{
  const s = bgSlide(false);
  s.addText("State Diagram - The Player Character", { x: 0.6, y: 0.45, w: 11, h: 0.6, fontFace: HEAD_FONT, fontSize: 26, bold: true, color: NAVY2, margin: 0 });
  pill(s, "BONUS", 11.4, 0.5, 1.3, false);
  const w2 = 9.0, h2 = w2 / (2000 / 1240);
  s.addImage({ path: "diagram_state_player.png", x: (SW - w2) / 2, y: 1.2, w: w2, h: h2 });
  footer(s, false, page++);
}

// ---------- SLIDE 17: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "10 min", "The Key Press Reactor - and the very first piece of your Semester 1 project.");

// ---------- SLIDE 15: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("The Key Press Reactor", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("An object that responds when you press a key - and the very first piece of your Semester 1 project. Save it somewhere you'll remember, as [YourName]Game.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "One key does one thing.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "All 4 arrow keys move your character.", "Everything works end to end."],
    ["Level 3 - Refined", "Animation, sound, and/or polish - pick what interests you.", "Ready to share."],
  ];
  let lx = 0.6;
  levels.forEach(([title, desc, tag]) => {
    s.addShape("roundRect", { x: lx, y: 2.3, w: 3.9, h: 2.1, rectRadius: 0.1, fill: { color: "EDF2F8" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(title, { x: lx + 0.2, y: 2.45, w: 3.5, h: 0.4, fontFace: BODY_FONT, fontSize: 14.5, bold: true, color: NAVY2, margin: 0 });
    s.addText(desc, { x: lx + 0.2, y: 2.9, w: 3.5, h: 0.9, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
    s.addText(tag, { x: lx + 0.2, y: 3.85, w: 3.5, h: 0.45, fontFace: BODY_FONT, fontSize: 11.5, italic: true, color: GOLD, margin: 0 });
    lx += 4.1;
  });

  s.addShape("roundRect", { x: 0.6, y: 4.7, w: 12.1, h: 1.55, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 4.92, w: 0.55, h: 0.55 });
  s.addText("New this lesson: 4 engines (GameMaker, Unity, Unreal, 001 Game Creator) get a second, alternate set of instructions on their handout - the other coding style (visual or scripted) for the exact same result. Try it if you're curious how the other side works.", {
    x: 1.6, y: 4.82, w: 10.9, h: 1.3, valign: "middle", fontFace: BODY_FONT, fontSize: 12.5, color: CREAM, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 16: Lab Time Notes ----------
{
  const s = bgSlide(false);
  s.addText("Lab Time Notes", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/wrench.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const notes = [
    "Circulate by engine cluster, starting with whichever engine has the most students this week.",
    "Common early snag, every engine: an event block/node that isn't actually connected won't run. \"Nothing happens\" → check connections first.",
    "Students who finish Level 2 early should be pointed at the Level 3 ideas on their handout - not given new unrelated work.",
    "At 1:50, call time and do a 2-minute share-out: ask 2-3 students to show what they built, across different engines if possible.",
  ];
  notes.forEach((t, i) => {
    const y = 1.75 + i * 1.15;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.9, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  footer(s, false, page++);
}

// ---------- SLIDE 17: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("Great Work Today!", { x: 0, y: 2.75, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 38, bold: true, color: GOLD_L, margin: 0 });
  s.addText("Next week: more programming foundations - variables, loops, and conditionals.", {
    x: 1.5, y: 3.6, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 17, color: CREAM, margin: 0,
  });
  s.addText("Keep building on your Key Press Reactor at home if you'd like - and start jotting notes in your Game Design Document.", {
    x: 1.5, y: 4.3, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_02_Slides.pptx" }).then(() => console.log("done"));
