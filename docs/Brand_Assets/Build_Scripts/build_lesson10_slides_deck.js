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

const NOTES = {
 "welcome": "\"Welcome back! Last week your game finally got a real scoreboard. Quick show of hands - who styled their HUD, or added scoring variety?\" [Pause, call on 1-2 volunteers]\n\n\"Today we fix something that's probably bugged a few of you since Lesson 8: your game 'ends,' but it doesn't actually stop. Today it finally does - and you can play again without touching the Run button.\"",
 "missing": "\"Right now, when you win or lose, a message shows up and... nothing else changes. Enemies keep coming. You can keep shooting. The numbers might even keep moving. 'Game Over' hasn't actually meant anything yet.\"\n\n\"Today, 'over' means something real: the round stops mattering, and there's a clear way to start a fresh one. That's two pieces - freezing, and restarting - and they both lean on the exact same trick.\"\n\n[Discussion question] \"Think of any game you've played - what usually happens the instant you lose? What are all the different things that might stop or change?\" (Draw out: enemies stop, controls lock, a menu or prompt appears, some kind of restart option shows up - list a few examples on the board if helpful.)",
 "setup": "\"Here's the trick: your game already HAS code that sets everything up - score at 0, lives at 3, enemies in place. It's been running since Lesson 5, just once, right at the start. Today, we give that exact same code a second way to run: on demand, whenever R is pressed.\"\n\n[Demo cue: switch to Godot, open a project with a working scoreboard from Lesson 9]\n\n\"Watch: I'm taking the setup - score at 0, lives at 3, enemies in place - and moving it into its own function on my GameManager, start_game(). The enemy-spawning loop that used to run in _ready() becomes spawn_enemies() on the Main scene, and start_game() calls it. Then GameManager's _ready() just calls start_game once, deferred - Godot runs a child's _ready() before its parent has finished loading, and the parent won't accept new children until it has, so call_deferred waits for that. That's the whole refactor.\" [Build it live] \"Now here's the payoff: I can call start_game() again, any time I want, from anywhere else in my code - like, say, when R is pressed.\" [Wire up the restart key, demonstrate a full reset]\n\n\"Notice nothing about the SETUP LOGIC itself changed. We didn't write a second version of it for restarting. We just gave the existing version a second doorway in.\"\n\n[Discussion question] \"Why is it better to have ONE piece of setup code with two ways to trigger it, instead of writing the setup code twice - once for the start, once for restart?\" (Draw out: if you ever need to change how the game sets up - say, starting with 4 lives instead of 3 - you'd only have to change it in one place instead of two, and remembering to update both copies is exactly the kind of thing that's easy to forget.)",
 "freeze": "\"Restarting is only half of today. The other half: actually STOPPING the game once it's over, so nothing keeps happening behind that win/lose message.\"\n\n[Demo cue: same Godot project] \"I've added one new variable on GameManager - game_active, true by default. The moment you win or lose, it flips to false. And now, at the very top of my enemy's movement code, one line: if the game isn't active, don't do anything else this frame.\" [Show enemies freezing on game-over]\n\n\"One flag, checked in the right place, freezes the whole game. That's genuinely it - no special 'pause' feature required, just a question asked at the right moment, every frame: 'is the game still going?'\"\n\n[Discussion question] \"What do you think would happen if we forgot to check game_active somewhere - say, in the code that handles a bullet hitting an enemy? \" (Draw out: that specific thing would keep happening even after game-over - a good preview of exactly the kind of bug students should test for today.)",
 "activity": "[Show the next slide: the Activity Diagram] \"Here's the whole round in one picture. The round ends, game_active flips to false, and everything freezes. Then we ask two questions when R is pressed: was R actually pressed, and is the round actually over? Only if both answers are yes do we run start_game() - the same one reset routine that runs when the game first launches. Press R in the middle of a round and the answer to the second question is no, so nothing happens.\"",
 "exercise": "\"For lab today, open your own project from last week - not a new one, and no new assets needed. Every engine's handout has the same two moves: turn your existing setup code into something callable again, and add a game_active-style flag that freezes things once the round ends.\"\n\n[Hand out engine-specific handouts]\n\n\"Test in both directions today, and in both game-end states: play to a win, restart, confirm it's clean. Then play to a loss, restart, confirm that's clean too. And don't skip testing that R does NOTHING while a round is still active - that's an easy thing to accidentally break.\"\n\n\"Questions before we start lab? ... Let's make your game actually replayable.\"",
 "lab": "The single most important thing to verify this week: a full round-trip, twice - win then restart, and separately lose then restart. Both should leave the game in an identical, clean starting state.\n\nMost likely mistake across every engine: forgetting to clear old enemies before respawning on restart, leading to doubled enemies or a wrong count after the first restart. Check for this explicitly, don't just eyeball the screen.\n\nGame state drift: any start-up code or game-over effect a student added beyond the handout (extra green-flag scripts, animations, color or costume changes, hidden or moved objects, sounds) is NOT reset by the handout steps. Walk the room with the Restart Audit checklist from the handout; the symptom is usually a restart that looks almost right (a player stuck in the wrong costume, a message still showing, an animation that never replays).\n\nSecond most likely mistake: R doing something mid-game because the game_active check was left out of the restart-key logic. Test this directly by pressing R while actively playing.\n\nThis is a genuinely satisfying lesson to finish - encourage a full playthrough as part of the end-of-lesson share-out (win, restart, lose, restart) since it's the first time the game truly feels replayable, not just winnable once.\n\nLevel 3 bug hunt: there is a real bug in the win/lose logic from Lessons 6 and 8. A round with some kills and some escapes (say 3 and 2) ends with no win and no loss, and R does nothing. Ask early finishers to find it by playing a mixed round; accept any fix that guarantees every round ends in exactly one outcome. The shown fix: an escape also reduces enemies_remaining, and the win check requires lives above 0.\n\nStudents who finish Level 2 early should be pointed at the Level 3 ideas on their handout, not given new unrelated work.",
 "title": "[As students settle] Welcome everyone in. Today is Lesson 10: Play It Again - the game-over and restart flow. Objectives and the agenda come next.",
 "glance": "[Read the objectives aloud] Today's engine for the lecture demo is Godot. No new assets this week - students open their own Lesson 9 project.",
 "agenda": "[Walk the agenda top to bottom] Five minutes of welcome, then the three lecture blocks, then lab. The lecture is about 55 minutes before lab starts.",
 "closing": "[Wrap-up] \"A clean freeze and a real restart, built almost entirely by reusing code you already trust. Next week: multiple waves - the Level 2 Complete milestone. Try a Level 3 idea at home if you'd like.\"",
 "state": "[Show the next slide: What Is Game State?] \"One more idea before lab, because it will save you an hour of confusion. Everything your game remembers right now is called its game state: the score, the lives, how many enemies are left, whether the round is over, where the player is standing, which picture the player is showing, even whether a sound or an animation is mid-play. Until today, your game only ever started once and ended once, so nothing had to be put back. A restart changes the rule: every single piece of state has to go back to how it was at the start.\"",
 "audit": "[Show the next slide: The Restart Audit] \"So today you audit. Find every place your game does something when it starts - in Scratch, every 'when green flag clicked' script on every sprite and the stage; in Godot, every _ready(); in GameMaker, every Create event; in Unity, every Awake and Start. Then find everything you did at game over - an animation, a color change, a hidden sprite, a sound, a message. If something changed when the round ended, something has to change it back when the round restarts. If you added something extra that isn't in the handout, it is on your list too, because the handout can't know about it.\""
};

const SW = 13.333, SH = 7.5;
const LESSON_LABEL = "Lesson 10";

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
  s.addText("Lesson 10", {
    x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT,
    fontSize: 44, bold: true, color: GOLD_L, margin: 0,
  });
  s.addText("Core Mechanics IV: Play It Again", {
    x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT,
    fontSize: 22, color: CREAM, margin: 0,
  });
  s.addText("Semester 1, Week 10  -  Game-Over + Restart Flow", {
    x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT,
    fontSize: 14, italic: true, bold: true, color: GOLD_L, margin: 0,
  });
  s.addNotes(NOTES.title);
  page++;
}

// ---------- SLIDE 2: AT A GLANCE ----------
{
  const s = bgSlide(false);
  s.addText("At a Glance", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addText("Lesson 10 objectives & materials", { x: 0.6, y: 1.05, w: 8, h: 0.4, fontFace: BODY_FONT, fontSize: 15, italic: true, color: TEXT_MUTE, margin: 0 });

  s.addText("By the end of today, students will:", { x: 0.6, y: 1.65, w: 6.6, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const objectives = [
    "Refactor setup code into a reusable reset routine",
    "Introduce a flag that gates whether gameplay applies",
    "Implement a full, clean restart via a dedicated key",
    "Test a game that's genuinely replayable, not just winnable once",
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
    "No new assets this week - logic and refactoring lesson",
    "Every engine involves a genuine refactor, not just additions",
    "Your own [YourName]Game project from Lesson 9 - not a new one",
  ];
  materials.forEach((t, i) => {
    const y = 2.35 + i * 0.72;
    s.addText("•", { x: 7.9, y, w: 0.25, h: 0.6, fontFace: BODY_FONT, fontSize: 15, color: GOLD, margin: 0 });
    s.addText(t, { x: 8.2, y, w: 4.4, h: 0.68, fontFace: BODY_FONT, fontSize: 12.5, color: TEXT_DARK, margin: 0 });
  });

  s.addText("IGDA framework tie-in:  Programming - state machines (introductory), code reuse/refactoring, input handling (intermediate).", {
    x: 0.6, y: 6.55, w: 12.1, h: 0.5, fontFace: BODY_FONT, fontSize: 11.5, italic: true, color: TEXT_MUTE, margin: 0,
  });

  s.addNotes(NOTES.glance);
  footer(s, false, page++);
}

// ---------- SLIDE 3: TODAY'S AGENDA ----------
{
  const s = bgSlide(false);
  s.addText("Today's Agenda", { x: 0.6, y: 0.45, w: 8, h: 0.7, fontFace: HEAD_FONT, fontSize: 38, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "icons/production.png", x: 11.6, y: 0.4, w: 0.9, h: 0.9 });

  const rows = [
    ["0:00 - 0:05", "Welcome back, Lesson 9 recap"],
    ["0:05 - 0:20", "The Last Missing Piece: What \"Over\" Should Actually Mean"],
    ["0:20 - 0:40", "One Setup, Two Doors In"],
    ["0:40 - 0:55", "The Freeze: A Flag That Gates Everything"],
    ["0:55 - 1:00", "Introduce lab exercise, hand out engine-specific handouts"],
    ["LAB TIME", "— Play It Again —"],
    ["1:00 - 1:50", "Build game-over freeze and restart flow"],
    ["1:50 - 2:00", "Exit check: full playthrough demos"],
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

  s.addNotes(NOTES.agenda);
  footer(s, false, page++);
}

// ---------- SLIDE 4: SECTION - Welcome Back ----------
sectionDivider("book", "Welcome Back", "5 min", "Lesson 9 recap - your game got a real scoreboard.");

// ---------- SLIDE 5: Welcome Back content ----------
{
  const s = bgSlide(false);
  s.addText("Welcome Back", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("Quick recap of Lesson 9:", { x: 0.6, y: 1.5, w: 11.5, h: 0.4, fontFace: BODY_FONT, fontSize: 16, bold: true, color: NAVY2, margin: 0 });
  const points = [
    "Score and lives, both visible on screen the whole time you play",
    "A real, persistent HUD - the promise from Lesson 6, finally kept",
    "Today: what \"Game Over\" actually means finally gets built",
  ];
  points.forEach((t, i) => {
    const y = 2.0 + i * 0.85;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.4, h: 0.4 });
    s.addText(t, { x: 1.15, y: y - 0.08, w: 11.3, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addShape("roundRect", { x: 0.6, y: 4.75, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/wrench.png", x: 0.85, y: 5.0, w: 0.6, h: 0.6 });
  s.addText("Open your own project from last week today, not a new one - and no new assets needed either, this is a logic and refactoring lesson.", {
    x: 1.65, y: 4.95, w: 10.8, h: 1.05, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  s.addNotes(NOTES.welcome);
  footer(s, false, page++);
}

// ---------- SLIDE 6: SECTION - The Last Missing Piece ----------
sectionDivider("design", "The Last Missing Piece", "15 min", "What should \"Game Over\" actually mean?");

// ---------- SLIDE 7: Last Missing Piece content ----------
{
  const s = bgSlide(false);
  s.addText("\"Over\" Doesn't Mean Anything Yet", { x: 0.6, y: 0.45, w: 11.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: NAVY2, margin: 0 });
  pill(s, "15 min", 11.4, 0.5, 1.3, false);

  s.addShape("roundRect", { x: 0.6, y: 1.5, w: 12.1, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/design.png", x: 0.85, y: 1.75, w: 0.55, h: 0.55 });
  s.addText("Right now, a message shows up when you win or lose - and nothing else changes. Enemies keep coming. You can keep shooting. Today, \"over\" finally means something real.", {
    x: 1.6, y: 1.65, w: 10.9, h: 1.2, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addText("Two pieces, today: freezing the round, and a clear way to start a fresh one - both built from the exact same trick.", {
    x: 0.6, y: 3.25, w: 11.9, h: 0.75, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.2, w: 12.1, h: 1.6, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 4.45, w: 0.55, h: 0.55 });
  s.addText("Discussion: think of any game you've played - what usually happens the instant you lose? (Enemies stop, controls lock, a menu or prompt appears, some kind of restart option shows up.)", {
    x: 1.6, y: 4.35, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 13.5, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(NOTES.missing);
  footer(s, false, page++);
}

// ---------- SLIDE 8: SECTION - One Setup, Two Doors In ----------
sectionDivider("programming", "One Setup, Two Doors In", "20 min", "The same setup code, callable a second way.");

// ---------- SLIDE 9: One Setup content ----------
{
  const s = bgSlide(false);
  s.addText("Give Setup a Second Doorway", { x: 0.6, y: 0.45, w: 11.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 28, bold: true, color: NAVY2, margin: 0 });
  pill(s, "20 min", 11.4, 0.5, 1.3, false);

  s.addText("Your game already has code that sets everything up - running once since Lesson 5. Today, that same code gets a second way to run: on demand, whenever R is pressed.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.85, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.5, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/programming.png", x: 0.85, y: 2.75, w: 0.55, h: 0.55 });
  s.addText("Godot demo: GameManager gets start_game() (reset score, lives, enemies, then call spawn_enemies() on Main). GameManager's _ready() just calls it once, deferred (start_game.call_deferred()). Now start_game() can be called again, any time, from the restart key.", {
    x: 1.6, y: 2.65, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, margin: 0,
  });

  s.addText("Nothing about the setup logic itself changed - it just got a second doorway in.", {
    x: 0.6, y: 4.4, w: 11.9, h: 0.55, fontFace: BODY_FONT, fontSize: 14.5, italic: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 5.1, w: 12.1, h: 1.35, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 5.32, w: 0.55, h: 0.55 });
  s.addText("Discussion: why is one piece of setup code with two ways in better than writing it twice? (Change it once, not remember to update two copies.)", {
    x: 1.6, y: 5.22, w: 10.9, h: 1.15, valign: "middle", fontFace: BODY_FONT, fontSize: 13, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(NOTES.setup);
  footer(s, false, page++);
}

// ---------- SLIDE 10: SECTION - The Freeze ----------
sectionDivider("shield", "The Freeze", "15 min", "A single flag, checked in the right place.");

// ---------- SLIDE 11: The Freeze content ----------
{
  const s = bgSlide(false);
  s.addText("One Flag Gates Everything", { x: 0.6, y: 0.45, w: 10.8, h: 0.7, fontFace: HEAD_FONT, fontSize: 30, bold: true, color: NAVY2, margin: 0 });
  pill(s, "15 min", 11.4, 0.5, 1.3, false);

  s.addText("Restarting is only half of today. The other half: actually stopping the game once it's over, so nothing keeps happening behind the win/lose message.", {
    x: 0.6, y: 1.45, w: 11.9, h: 0.85, fontFace: BODY_FONT, fontSize: 15, bold: true, color: NAVY2, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 2.5, w: 12.1, h: 1.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { type: "none" } });
  s.addImage({ path: "icons/shield.png", x: 0.85, y: 2.75, w: 0.55, h: 0.55 });
  s.addText("game_active (on GameManager) starts true, flips false at win/lose. At the very top of enemy movement code: if the game isn't active, don't do anything else this frame.", {
    x: 1.6, y: 2.65, w: 10.9, h: 1.4, valign: "middle", fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  s.addShape("roundRect", { x: 0.6, y: 4.4, w: 12.1, h: 1.75, rectRadius: 0.12, fill: { color: "0F2C52" }, line: { color: GOLD, width: 1 } });
  s.addImage({ path: "icons/chat.png", x: 0.85, y: 4.63, w: 0.55, h: 0.55 });
  s.addText("Discussion: what would happen if we forgot to check game_active somewhere - say, in the bullet-hits-enemy code? (That specific thing would keep happening even after game-over - test for this today.)", {
    x: 1.6, y: 4.53, w: 10.9, h: 1.55, valign: "middle", fontFace: BODY_FONT, fontSize: 13, italic: true, color: CREAM, margin: 0,
  });

  s.addNotes(NOTES.freeze);
  footer(s, false, page++);
}

// ---------- SLIDE 11b: Activity Diagram ----------
{
  const s = bgSlide(false);
  s.addText("The Whole Round, One Picture", { x: 0.6, y: 0.4, w: 12, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  s.addImage({ path: "diagram_activity_restart_flow.png", x: 0.8, y: 1.15, w: 5.8, h: 5.86 });
  s.addText("Freeze, ask, reset", { x: 7.1, y: 1.5, w: 5.6, h: 0.5, fontFace: BODY_FONT, fontSize: 20, bold: true, color: NAVY2, margin: 0 });
  s.addText([
    { text: "The round ends: game_active becomes false and everything freezes.", options: { bullet: true, breakLine: true } },
    { text: "R pressed? Round over? Only a double yes runs start_game().", options: { bullet: true, breakLine: true } },
    { text: "R mid-game is ignored - the second check blocks it.", options: { bullet: true, breakLine: true } },
    { text: "start_game() is the same routine that runs at launch: one setup, two doorways.", options: { bullet: true } },
  ], { x: 7.1, y: 2.15, w: 5.6, h: 3.6, fontFace: BODY_FONT, fontSize: 16, color: TEXT_DARK, paraSpaceAfter: 10, valign: "top", margin: 0 });
  s.addNotes(NOTES.activity);
  footer(s, false, page++);
}

// ---------- SLIDE 11c: What Is Game State? ----------
{
  const s = bgSlide(false);
  s.addText("What Is Game State?", { x: 0.6, y: 0.4, w: 12, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  s.addText("Game state is everything your game remembers right now.", { x: 0.6, y: 1.2, w: 12, h: 0.5, fontFace: BODY_FONT, fontSize: 18, italic: true, color: TEXT_MUTE, margin: 0 });
  const cols = [
    ["Numbers and flags", "score, lives, enemies_remaining, game_active"],
    ["Where things are", "player position, enemy positions, what is visible or hidden"],
    ["How things look", "costume or sprite, color, size, animations mid-play"],
    ["What is playing", "sounds, timers, tweens, messages on screen"],
  ];
  cols.forEach(([h, d], idx) => {
    const x = 0.6 + (idx % 2) * 6.2, y = 2.0 + Math.floor(idx / 2) * 1.75;
    s.addShape("roundRect", { x, y, w: 5.9, h: 1.5, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { color: GOLD, width: 1 } });
    s.addText(h, { x: x + 0.25, y: y + 0.15, w: 5.4, h: 0.45, fontFace: BODY_FONT, fontSize: 18, bold: true, color: NAVY2, margin: 0 });
    s.addText(d, { x: x + 0.25, y: y + 0.65, w: 5.4, h: 0.75, fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, margin: 0 });
  });
  s.addText("Until today, your game started once and ended once. A restart means ALL of it has to go back to how it was at the start.", { x: 0.6, y: 5.6, w: 12.1, h: 0.9, fontFace: BODY_FONT, fontSize: 17, bold: true, color: NAVY, margin: 0 });
  s.addNotes(NOTES.state);
  footer(s, false, page++);
}

// ---------- SLIDE 11d: The Restart Audit ----------
{
  const s = bgSlide(false);
  s.addText("The Restart Audit", { x: 0.6, y: 0.4, w: 12, h: 0.7, fontFace: HEAD_FONT, fontSize: 32, bold: true, color: NAVY2, margin: 0 });
  s.addText("Two lists. Everything must be on one of them, and reset by start_game.", { x: 0.6, y: 1.2, w: 12, h: 0.5, fontFace: BODY_FONT, fontSize: 18, italic: true, color: TEXT_MUTE, margin: 0 });
  s.addShape("roundRect", { x: 0.6, y: 1.95, w: 5.9, h: 2.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { color: GOLD, width: 1 } });
  s.addText("1. Everything that happens at START", { x: 0.85, y: 2.05, w: 5.4, h: 0.45, fontFace: BODY_FONT, fontSize: 17, bold: true, color: NAVY2, margin: 0 });
  s.addText([
    { text: "Scratch: EVERY \"when green flag clicked\" - all sprites and the stage", options: { bullet: true, breakLine: true } },
    { text: "Godot: every _ready()", options: { bullet: true, breakLine: true } },
    { text: "GameMaker: every Create event", options: { bullet: true, breakLine: true } },
    { text: "Unity: every Awake() and Start()", options: { bullet: true } },
  ], { x: 0.85, y: 2.6, w: 5.4, h: 1.9, fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, paraSpaceAfter: 6, valign: "top", margin: 0 });
  s.addShape("roundRect", { x: 6.8, y: 1.95, w: 5.9, h: 2.7, rectRadius: 0.12, fill: { color: "EDF2F8" }, line: { color: GOLD, width: 1 } });
  s.addText("2. Everything that happens at the END", { x: 7.05, y: 2.05, w: 5.4, h: 0.45, fontFace: BODY_FONT, fontSize: 17, bold: true, color: NAVY2, margin: 0 });
  s.addText([
    { text: "Win and lose messages", options: { bullet: true, breakLine: true } },
    { text: "Animations, color or costume changes, sounds", options: { bullet: true, breakLine: true } },
    { text: "Anything hidden, moved, or stopped", options: { bullet: true, breakLine: true } },
    { text: "Anything extra YOU added - the handout can't know about it", options: { bullet: true } },
  ], { x: 7.05, y: 2.6, w: 5.4, h: 1.9, fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, paraSpaceAfter: 6, valign: "top", margin: 0 });
  s.addText("If it changed when the round ended, something must change it back when the round restarts.", { x: 0.6, y: 5.05, w: 12.1, h: 0.9, fontFace: BODY_FONT, fontSize: 18, bold: true, color: NAVY, margin: 0 });
  s.addNotes(NOTES.audit);
  footer(s, false, page++);
}

// ---------- SLIDE 12: SECTION - Today's Exercise & Lab ----------
sectionDivider("production", "Today's Exercise & Lab", "5 min", "Play It Again - freeze and restart flow.");

// ---------- SLIDE 13: Exercise / Levels ----------
{
  const s = bgSlide(false);
  s.addText("Play It Again", { x: 0.6, y: 0.45, w: 9.5, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 });
  pill(s, "5 min", 11.4, 0.5, 1.3, false);

  s.addText("When the round ends, the game freezes. Pressing R resets everything and starts a new round - built on top of your existing project.", {
    x: 0.6, y: 1.4, w: 11.9, h: 0.65, fontFace: BODY_FONT, fontSize: 14.5, color: TEXT_DARK, margin: 0,
  });

  const levels = [
    ["Level 1 - Minimal", "Round ends -> game freezes; R resets and starts a new round.", "MVP - a complete finish for today."],
    ["Level 2 - Complete", "A restart is a genuinely clean slate, any number of times in a row.", "Everything works end to end."],
    ["Level 3 - Refined", "A transition effect, tighter reset structure, and a real bug to find and fix.", "Ready to share."],
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
  s.addText("Test in both directions and both end-states: win then restart, lose then restart, and confirm R does NOTHING while a round is still active.", {
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
    "Most important check: a full round-trip, twice - win then restart, and separately lose then restart. Both should leave an identical, clean starting state.",
    "Most likely mistake: old enemies not cleared before respawning on restart, leading to doubled enemies after the first restart.",
    "Second most likely mistake: R doing something mid-game because the game_active check was left out of the restart-key logic - test this directly.",
    "Game state drift: anything students added beyond the handout (extra green-flag scripts, animations, color changes, hidden objects) is not reset by the handout steps - use the Restart Audit checklist.",
    "Level 3 bug hunt: a round with some kills and some escapes ends with no win and no loss, and R does nothing. Early finishers should find it and fix it their own way.",
    "This is a satisfying lesson to finish - encourage a full playthrough (win, restart, lose, restart) as part of the share-out.",
  ];
  notes.forEach((t, i) => {
    const y = 1.6 + i * 0.92;
    s.addImage({ path: "icons/check.png", x: 0.6, y, w: 0.42, h: 0.42 });
    s.addText(t, { x: 1.2, y: y - 0.15, w: 11.3, h: 0.9, fontFace: BODY_FONT, fontSize: 13.5, color: TEXT_DARK, valign: "middle", margin: 0 });
  });

  s.addNotes(NOTES.lab);
  footer(s, false, page++);
}

// ---------- SLIDE 15: Closing ----------
{
  const s = bgSlide(true, "bg_title_with_slhe.png");
  s.addText("Genuinely Replayable!", { x: 0, y: 2.45, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, margin: 0 });
  s.addText("A clean freeze and a real restart, built almost entirely by reusing code you already trust.", {
    x: 1.5, y: 3.35, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0,
  });
  s.addText("Next week: multiple waves - the Level 2 Complete milestone.", {
    x: 1.5, y: 4.0, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 15, color: CREAM, margin: 0,
  });
  s.addText("Try a Level 3 idea at home if you'd like - a transition effect is a fun, visible one.", {
    x: 1.5, y: 4.7, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 13, italic: true, color: "9FB6D6", margin: 0,
  });
  s.addNotes(NOTES.closing);
  footer(s, true, page++);
}

pres.writeFile({ fileName: "Lesson_10_Slides.pptx" }).then(() => console.log("done"));
