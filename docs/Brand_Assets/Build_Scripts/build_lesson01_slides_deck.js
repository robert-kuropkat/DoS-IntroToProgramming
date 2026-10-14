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

function bgSlide(dark, darkImg) {
  // NOTE: deliberately not setting slide.background here. Calling
  // slide.background = {...} on every slide caused pptxgenjs 4.0.1 to emit a
  // brand-new slideMaster/slideLayout Content_Types.xml override PER SLIDE
  // without actually writing those parts to the package (16 declared,
  // 1 present) -- an invalid OPC package that PowerPoint's strict parser
  // flags as "needs repair" even though LibreOffice/python-pptx tolerate it.
  // A full-bleed addImage() already covers 100% of the canvas, so the
  // solid-color background was redundant anyway.
  const slide = pres.addSlide();
  slide.addImage({ path: dark ? (darkImg || "bg_title.png") : "bg_content.png", x: 0, y: 0, w: SW, h: SH });
  return slide;
}

function footer(slide, dark, pageNum) {
  slide.addText("VIDEO GAME CREATION", {
    x: 0.5, y: 7.14, w: 5, h: 0.3, fontFace: BODY_FONT, fontSize: 10,
    color: dark ? GOLD_L : TEXT_MUTE, charSpacing: 2, bold: true, margin: 0,
  });
  slide.addText("Lesson 1", {
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

let page = 1;

// ---------- SLIDE 1: TITLE ----------
{
  const s = bgSlide(true, "bg_title_with_vgc.png");
  s.addText("S O U T H   L A K E   H O M E   E D U C A T O R S", {
    x: 0, y: 0.55, w: SW, h: 0.35, align: "center", fontFace: BODY_FONT,
    fontSize: 13, color: GOLD_L, charSpacing: 1, margin: 0,
  });
  s.addText("Lesson 1", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Welcome, Internet Safety & Choosing Your Platform", {
    x: 1.5, y: 4.5, w: SW - 3, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Thursday, August 13, 2026", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, color: "9FB6D6", margin: 0,
  });
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 1 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  // left column - objectives
  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Understand class structure and expectations",
    "Identify at least 2 online-safety practices specific to gaming",
    "Name the major disciplines of game development",
    "Have chosen (or shortlisted) an engine and made install progress",
  ];
  objectives.forEach((t, i) => {
    const y = 2.15 + i * 0.62;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.15, y: y - 0.03, w: 6.1, h: 0.55, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  // right column - materials card
  s.addShape("roundRect", { x: 7.6, y: 1.65, w: 5.15, h: 4.6, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addText("What We're Using Today", { x: 7.9, y: 1.85, w: 4.6, h: 0.4, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0 });
  const materials = [
    "Demo laptop with all 8 engines pre-installed + Key Press Reactor loaded",
    "NetSmartz video clip(s), queued and tested",
    "Engine & Platform Guide + Engine Showcase (sent home in advance)",
    "Sign-in sheet",
    "Offline installer backups on USB (in case of wifi trouble)",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Course/discipline orientation (Production & Critical Studies) and first exposure to the Programming content area.", {
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
    ["0:00 - 0:10", "Welcome, roll call, course overview"],
    ["0:10 - 0:30", "Internet Safety & Digital Citizenship (NetSmartz)"],
    ["0:30 - 0:45", "What Is Game Development?"],
    ["0:45 - 1:00", "Live Engine Tour: Key Press Reactor across all 8 engines"],
    ["LAB TIME", "— Install & Setup —"],
    ["1:00 - 1:45", "Guided install/setup of chosen engine, by station"],
    ["1:45 - 1:52", "Exit check: confirm working environment, quick share-out"],
    ["1:52 - 2:00", "This Semester's Game: reveal, GDD hand-out, codename vote"],
  ];

  let y = 1.55;
  rows.forEach(([time, seg]) => {
    const isDivider = time === "LAB TIME";
    const rowH = isDivider ? 0.48 : 0.55;
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
    y += rowH + 0.1;
  });

  footer(s, false, page++);
}

// ---------- Section divider helper ----------
function sectionDivider(icon, title, timeTag, teaser) {
  const s = bgSlide(true, `bg_title_${icon}.png`);
  s.addText(title, { x: 3.1, y: 2.6, w: 9.2, h: 1.0, fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, valign: "middle", margin: 0 });
  s.addText(teaser, { x: 3.15, y: 3.65, w: 9.0, h: 0.7, fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0 });
  pill(s, timeTag, 3.15, 4.35, 1.5, true);
  footer(s, true, page++);
  return s;
}

// ---------- SLIDE 4: SECTION - Welcome ----------
sectionDivider("book", "Welcome & Course Overview", "10 min", "Who we are, how the class runs, and what the whole year looks like.");

// ---------- SLIDE 5: Welcome content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome & Course Overview", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "10 min", 11.4, 0.5, 1.3, false);

  s.addText("This class is:", { x: 0.6, y: 1.5, w: 6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Hands-on and project-based — you'll be building the whole year",
    "Yours to choose: pick your own engine, and you can switch later",
    "Consistent across engines: every lesson provides examples for all 8 supported engines",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 5.6, h: 0.75, fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 7.1, y: 1.5, w: 5.65, h: 5.0, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/production.png", x: 7.35, y: 1.72, w: 0.55, h: 0.55 });
  s.addText("Syllabus Highlights", { x: 8.05, y: 1.72, w: 4.5, h: 0.55, valign: "middle", fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });

  const syl = [
    ["Meeting schedule", "30 lessons across two semesters, weekly, 2 hours each (lecture + lab)"],
    ["Two capstones", "Semester 1: a finished mini-game (find out exactly what today!).  Semester 2: your own game at the Co-op Expo"],
    ["Policies", "BYOD, AI-assistance guidelines, and conduct expectations — covered in the full Syllabus"],
  ];
  let sy = 2.5;
  syl.forEach(([h, d]) => {
    s.addText(h, { x: 7.4, y: sy, w: 4.9, h: 0.35, valign: "top", fontFace: BODY_FONT, fontSize: 13.5, bold: true, color: NAVY2, margin: 0 });
    s.addText(d, { x: 7.4, y: sy + 0.35, w: 4.9, h: 0.6, valign: "top", fontFace: BODY_FONT, fontSize: 12, color: TEXT_DARK, margin: 0 });
    sy += 1.35;
  });

  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - Internet Safety ----------
sectionDivider("shield", "Internet Safety & Digital Citizenship", "20 min", "Grounded in NetSmartz, with a gaming-specific lens — you'll be building and sharing games with each other all year.");

// ---------- SLIDE 7: Internet Safety - materials ----------
{
  const s = bgSlide(false);
  s.addText("Internet Safety & Digital Citizenship", { x: 0.6, y: 0.45, w: 10, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.7, 0.5, 1.05, false);
  s.addText("Suggested NetSmartz materials (preview clips ahead of time; check missingkids.org/netsmartz for the current library):", {
    x: 0.6, y: 1.25, w: 12.1, h: 0.5, fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: TEXT_MUTE, margin: 0,
  });

  const cards = [
    ["“Attitude Overdrive”", "NSTeens", "Good fit for middle schoolers — appropriate conduct & etiquette while gaming online."],
    ["“Into the Cloud” — Season 2", "Good fit for a mixed middle/high room", "Livestreaming safety, reporting/removing inappropriate content, cyberbullying, seeking help from a trusted adult."],
    ["Real-Life Stories series", "For older teens, use selectively", "Examines peer choices online and real consequences."],
  ];
  const cw = 3.9, gap = 0.2, startX = 0.6;
  cards.forEach((c, i) => {
    const x = startX + i * (cw + gap);
    s.addShape("roundRect", { x, y: 2.05, w: cw, h: 3.15, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
    s.addText(c[0], { x: x + 0.25, y: 2.3, w: cw - 0.5, h: 0.75, valign: "top", fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
    s.addText(c[1], { x: x + 0.25, y: 3.0, w: cw - 0.5, h: 0.5, valign: "top", fontFace: BODY_FONT, fontSize: 12, italic: true, color: GOLD, bold: true, margin: 0 });
    s.addText(c[2], { x: x + 0.25, y: 3.55, w: cw - 0.5, h: 1.5, valign: "top", fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  footer(s, false, page++);
}

// ---------- SLIDE 8: Internet Safety - discussion points ----------
{
  const s = bgSlide(false);
  s.addText("Discussion Points", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/chat.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const discuss = [
    "Chatting with strangers in multiplayer/voice chat — what's okay to share, what isn't",
    "Recognizing griefing, cyberbullying, and manipulation in gaming communities — how and when to report or block",
    "In-game purchases, scams, and “too good to be true” offers",
    "No personal info in usernames, credits, or shared builds — once we start sharing games with each other this year",
    "Who to go to if something online feels wrong: a parent, or the instructor",
  ];
  let y = 1.55;
  discuss.forEach((t) => {
    s.addShape("roundRect", { x: 0.6, y, w: 12.1, h: 0.92, rectRadius: 0.1, fill: { color: "FFFFFF" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(t, { x: 1.0, y, w: 11.4, h: 0.92, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0 });
    y += 1.05;
  });

  footer(s, false, page++);
}

// ---------- SLIDE 9: SECTION - What Is Game Development? ----------
sectionDivider("design", "What Is Game Development?", "15 min", "A fast overview of the disciplines that make a game — so you can see where coding fits into the bigger picture.");

// ---------- SLIDE 10: The Five Disciplines ----------
{
  const s = bgSlide(false);
  s.addText("The Five Disciplines", { x: 0.6, y: 0.45, w: 9, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addText("This class focuses on Programming, with light Design.", { x: 0.6, y: 1.1, w: 10, h: 0.4, fontFace: BODY_FONT, fontSize: 14.5, italic: true, color: TEXT_MUTE, margin: 0 });

  const disc = [
    ["design", "Design", "Deciding what the game is and how it plays"],
    ["programming", "Programming", "Making the rules and mechanics actually work"],
    ["art", "Visual Art", "Characters, environments, and UI"],
    ["audio", "Audio", "Music and sound effects"],
    ["production", "Production", "Planning, scoping, and finishing a project"],
  ];
  const cw = 2.32, gap = 0.14, startX = 0.6, top = 1.85;
  disc.forEach((d, i) => {
    const x = startX + i * (cw + gap);
    const isMain = d[1] === "Programming";
    s.addShape("roundRect", {
      x, y: top, w: cw, h: 4.7, rectRadius: 0.12,
      fill: { color: isMain ? NAVY2 : "EDF2F8" }, line: isMain ? { color: GOLD, width: 2 } : { type: "none" },
    });
    s.addImage({ path: `icons/${d[0]}.png`, x: x + (cw - 0.9) / 2, y: top + 0.3, w: 0.9, h: 0.9 });
    s.addText(d[1], { x: x + 0.15, y: top + 1.35, w: cw - 0.3, h: 0.5, align: "center", valign: "top", fontFace: BODY_FONT, fontSize: 15, bold: true, color: isMain ? GOLD_L : NAVY2, margin: 0 });
    s.addText(d[2], { x: x + 0.2, y: top + 1.9, w: cw - 0.4, h: 1.1, align: "center", valign: "top", fontFace: BODY_FONT, fontSize: 12, color: isMain ? CREAM : TEXT_DARK, margin: 0 });
    if (isMain) {
      pill(s, "MAIN FOCUS", x + 0.16, top + 3.85, cw - 0.32, true);
    }
  });

  footer(s, false, page++);
}

// ---------- SLIDE 11: SECTION - Live Engine Tour ----------
sectionDivider("controller", "Live Engine Tour", "15 min", "One consistent demo — Key Press Reactor — run live across all 8 supported engines.");

// ---------- SLIDE 12: Engine tour table ----------
{
  const s = bgSlide(false);
  s.addText("Live Engine Tour: Key Press Reactor", { x: 0.6, y: 0.4, w: 10.5, h: 0.65, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: NAVY2, margin: 0 });
  s.addText("Press Play, tap an arrow key, watch the same reaction happen 8 different ways — about 60-90 seconds per engine.", {
    x: 0.6, y: 1.0, w: 12.1, h: 0.4, fontFace: BODY_FONT, fontSize: 13, italic: true, color: TEXT_MUTE, margin: 0,
  });

  const engines = [
    ["Scratch", "Block"],
    ["GDevelop", "Visual"],
    ["001 Game Creator", "Visual"],
    ["GameMaker", "Hybrid (Visual ↔ Script)"],
    ["Unity", "Script (+ Visual option)"],
    ["Godot", "Script"],
    ["TIC-80", "Script"],
    ["Unreal (advanced track)", "Hybrid (Blueprints ↔ C++)"],
  ];
  const cw = 2.98, ch = 1.05, gapx = 0.13, gapy = 0.13, startX = 0.6, top = 1.65;
  engines.forEach((e, i) => {
    const col = i % 4, row = Math.floor(i / 4);
    const x = startX + col * (cw + gapx);
    const y = top + row * (ch + gapy);
    s.addShape("roundRect", { x, y, w: cw, h: ch, rectRadius: 0.1, fill: { color: "FFFFFF" }, line: { color: "DCE3EC", width: 1 } });
    s.addText(e[0], { x: x + 0.18, y: y + 0.1, w: cw - 0.36, h: 0.45, fontFace: BODY_FONT, fontSize: 14, bold: true, color: NAVY2, margin: 0 });
    s.addText(e[1], { x: x + 0.18, y: y + 0.55, w: cw - 0.36, h: 0.42, fontFace: BODY_FONT, fontSize: 12, color: GOLD, bold: true, margin: 0 });
  });

  s.addText("Closing note: device-fit guidance so students with older laptops feel set up for success, not steered toward something that will frustrate them.", {
    x: 0.6, y: 4.15, w: 12.1, h: 0.5, fontFace: BODY_FONT, fontSize: 12, italic: true, color: TEXT_MUTE, margin: 0,
  });

  footer(s, false, page++);
}

// ---------- SLIDE 13: SECTION - Lab Time ----------
sectionDivider("wrench", "Lab Time: Install & Setup", "60 min", "Goal: every student leaves with a working environment for their chosen engine.");

// ---------- SLIDE 14: Lab structure ----------
{
  const s = bgSlide(false);
  s.addText("Lab: Install & Setup", { x: 0.6, y: 0.45, w: 9, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  s.addShape("roundRect", { x: 0.6, y: 1.15, w: 12.1, h: 0.6, rectRadius: 0.1, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addText("Goal: every student leaves with a working environment — or at minimum, a browser-based fallback (GDevelop or Scratch) open and ready.", {
    x: 0.85, y: 1.15, w: 11.6, h: 0.6, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, color: CREAM, italic: true, margin: 0,
  });

  const steps = [
    "Quick poll: which engine is each student planning to start with?",
    "Students who installed at home verify their install works and open a new project",
    "Students who haven't installed: start with GDevelop or Scratch in-browser immediately, then queue a larger download in the background",
    "Instructor circulates by station (clustered by chosen engine) to troubleshoot and answer questions",
    "Exit check: each student's editor is open, ideally with a blank or template project loaded",
  ];
  let y = 2.05;
  steps.forEach((t, i) => {
    s.addShape("roundRect", { x: 0.6, y, w: 0.5, h: 0.5, rectRadius: 0.25, fill: { color: GOLD }, line: { type: "none" } });
    s.addText(String(i + 1), { x: 0.6, y, w: 0.5, h: 0.5, align: "center", valign: "middle", fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY, margin: 0 });
    s.addText(t, { x: 1.3, y: y - 0.06, w: 11.4, h: 0.65, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, margin: 0 });
    y += 0.87;
  });

  footer(s, false, page++);
}

// ---------- SLIDE: SECTION - This Semester's Game ----------
sectionDivider("design", "This Semester's Game", "8 min", "The big reveal: what we're all building together this semester.");

// ---------- SLIDE: This Semester's Game content ----------
{
  const s = bgSlide(false);
  s.addText("This Semester's Game", { x: 0.6, y: 0.45, w: 10, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  pill(s, "8 min", 11.7, 0.5, 1.05, false);

  s.addShape("roundRect", { x: 0.6, y: 1.3, w: 12.1, h: 1.25, rectRadius: 0.12, fill: { color: NAVY2 }, line: { type: "none" } });
  s.addText("A top-down arcade shooter, in the tradition of classics like Space Invaders and Galaga. Move, dodge, and shoot your way through waves of descending enemies.", {
    x: 0.95, y: 1.3, w: 11.4, h: 1.25, valign: "middle", fontFace: BODY_FONT, fontSize: 16, bold: true, color: CREAM, margin: 0,
  });

  const steps = [
    ["High Concept reveal", "Everyone in class builds this same game this semester. Your code and your personal touches make it yours."],
    ["Hand out your GDD", "Your Lesson 1 Game Design Document (sneak peek). Yours to keep."],
    ["Vote on a codename", "Real studios use a codename before the final title is locked in. Let's pick ours!"],
  ];
  const cw = 3.9, gap = 0.2, startX = 0.6, top = 2.85;
  steps.forEach((h, i) => {
    const x = startX + i * (cw + gap);
    s.addShape("roundRect", { x, y: top, w: cw, h: 3.6, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
    s.addShape("ellipse", { x: x + (cw - 0.7) / 2, y: top + 0.3, w: 0.7, h: 0.7, fill: { color: GOLD }, line: { type: "none" } });
    s.addText(String(i + 1), { x: x + (cw - 0.7) / 2, y: top + 0.3, w: 0.7, h: 0.7, align: "center", valign: "middle", fontFace: HEAD_FONT, fontSize: 22, bold: true, color: NAVY, margin: 0 });
    s.addText(h[0], { x: x + 0.25, y: top + 1.25, w: cw - 0.5, h: 0.5, align: "center", valign: "top", fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0 });
    s.addText(h[1], { x: x + 0.3, y: top + 1.8, w: cw - 0.6, h: 1.6, align: "center", valign: "top", fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  footer(s, false, page++);
}

// ---------- SLIDE 15: Homework ----------
{
  const s = bgSlide(false);
  s.addText("Homework / Before Next Class", { x: 0.6, y: 0.45, w: 10, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/book.png", x: 11.5, y: 0.4, w: 1.0, h: 1.0 });

  const hw = [
    ["Finish your install", "If your chosen engine wasn't fully installed by the end of lab, finish that at home."],
    ["Bookmark a tutorial", "Skim the Student Resource List for your engine and bookmark one beginner tutorial."],
    ["Explore (optional)", "Spend 15-20 minutes just clicking around your engine's interface to get comfortable."],
  ];
  const cw = 3.9, gap = 0.2, startX = 0.6, top = 2.0;
  hw.forEach((h, i) => {
    const x = startX + i * (cw + gap);
    s.addShape("roundRect", { x, y: top, w: cw, h: 3.0, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
    // NOTE: "oval" is not a valid OOXML prstGeom value (must be "ellipse") --
    // PptxGenJS writes the string verbatim without mapping it, which is an
    // invalid-content defect PowerPoint's strict parser strips on repair.
    s.addShape("ellipse", { x: x + (cw - 0.7) / 2, y: top + 0.3, w: 0.7, h: 0.7, fill: { color: GOLD }, line: { type: "none" } });
    s.addText(String(i + 1), { x: x + (cw - 0.7) / 2, y: top + 0.3, w: 0.7, h: 0.7, align: "center", valign: "middle", fontFace: HEAD_FONT, fontSize: 22, bold: true, color: NAVY, margin: 0 });
    s.addText(h[0], { x: x + 0.25, y: top + 1.25, w: cw - 0.5, h: 0.5, align: "center", valign: "top", fontFace: BODY_FONT, fontSize: 15.5, bold: true, color: NAVY2, margin: 0 });
    s.addText(h[1], { x: x + 0.3, y: top + 1.8, w: cw - 0.6, h: 1.1, align: "center", valign: "top", fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  footer(s, false, page++);
}

// ---------- SLIDE 16: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("See You Next Week!", { x: 0, y: 2.75, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 38, bold: true, color: GOLD_L, margin: 0 });
  s.addText("Lesson 2: Key Press Reactor — build it yourself, in your chosen engine.", {
    x: 1.5, y: 3.6, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 17, color: CREAM, margin: 0,
  });
  s.addText("Questions before then? Reach out to your instructor.", {
    x: 1.5, y: 4.35, w: SW - 3, h: 0.5, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_01_Slides.pptx" }).then(() => console.log("done"));
