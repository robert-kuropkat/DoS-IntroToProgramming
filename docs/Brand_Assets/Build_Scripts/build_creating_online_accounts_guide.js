const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, LevelFormat, ExternalHyperlink,
  TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageBreak, VerticalAlign
} = require("docx");
const brand = require("./brand.js");

const CONTENT_WIDTH = 9360; // US Letter, 1" margins
const border = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function link(text, url) {
  return new ExternalHyperlink({ children: [new TextRun({ text, style: "Hyperlink" })], link: url });
}
function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function pRuns(runs) { return new Paragraph({ spacing: { after: 160 }, children: runs }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function labelBullet(label, text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)]
  });
}
function calloutBox(text) {
  return new Paragraph({
    spacing: { after: 200, before: 80 },
    indent: { left: 200, right: 200 },
    shading: { fill: "FFF4E5", type: ShadingType.CLEAR },
    children: [new TextRun({ text, size: 21 })]
  });
}

function headerCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: brand.NAVY2, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 20 })] })]
  });
}
function bodyCell(text, width, opts = {}) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, size: 20, bold: !!opts.bold })] })]
  });
}

// ---------- Summary table ----------
const sumHeaders = ["Engine", "Account Required?", "Free to Start?", "Optional Paid Upgrade"];
const sumWidths = [1900, 1900, 1660, 3900];
const sumRows = [
  ["Scratch", "No - optional", "Yes, fully", "None exists"],
  ["Godot", "No - never", "Yes, forever", "None exists"],
  ["Unity", "Yes - mandatory", "Yes (Personal tier)", "Unity Pro, $2,310/yr - not relevant to this class"],
  ["GameMaker", "Yes - mandatory", "Yes (non-commercial tier)", "$99.99 one-time, only for commercial publishing"],
  ["GDevelop", "No for desktop app; yes for cloud/online features", "Yes, fully", "Silver/Gold/Pro subscriptions, roughly $5-$33/mo"],
  ["TIC-80", "No - never", "Yes, fully", "~$5-10 Pro version, purchased via itch.io"],
  ["001 Game Creator", "Yes - Steam account, even for the trial", "7-day full trial only", "Standalone purchase, list $79.99 - regularly on sale for $15-20"],
  ["Unreal Engine", "Yes - mandatory (Epic Games account)", "Yes, fully", "None - only a 5% royalty after $1M revenue"],
];
const sumTable = new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: sumWidths,
  rows: [
    new TableRow({ tableHeader: true, children: sumHeaders.map((t, i) => headerCell(t, sumWidths[i])) }),
    ...sumRows.map((row, ri) => new TableRow({
      children: row.map((t, i) => bodyCell(t, sumWidths[i], { fill: ri % 2 ? "F3F6FA" : "FFFFFF", bold: i === 0 }))
    }))
  ]
});

// ---------- Per-engine account sections ----------
function acctSection({ name, required, whatFree, whyNeeded, upgrade, verdict, learnMore }) {
  return [
    h2(name),
    labelBullet("Account required?", required),
    labelBullet("What you get for free", whatFree),
    labelBullet("Why the account is needed", whyNeeded),
    labelBullet("Optional paid upgrade", upgrade),
    h3("Our take"),
    p(verdict, { italics: true, color: "555555" }),
    ...(learnMore ? [pRuns([new TextRun({ text: "More info: ", bold: true }), link(learnMore.text, learnMore.url)])] : []),
    new Paragraph({ spacing: { after: 260 }, children: [] }),
  ];
}

const engines = [
  {
    name: "1. Scratch",
    required: "No. The offline editor needs no account at all.",
    whatFree: "The entire editor, forever, with no feature limits.",
    whyNeeded: "An account is optional, only for saving projects to Scratch's cloud and sharing them publicly on scratch.mit.edu. Nothing in this class requires it.",
    upgrade: "None exists - Scratch has no paid tier of any kind.",
    verdict: "Skip the account unless your student wants to publish projects online for others to see. If you do create one, use a parent's email and review Scratch's community guidelines together first - it's a public, moderated site.",
    learnMore: { text: "scratch.mit.edu", url: "https://scratch.mit.edu" },
  },
  {
    name: "2. Godot Engine",
    required: "No, never, for anything - not to download, install, or use any feature.",
    whatFree: "Everything. Godot is open source (MIT license) with no gated features.",
    whyNeeded: "N/A - there is nothing behind a login in Godot.",
    upgrade: "None exists. There is no Godot Pro, no subscription, nothing to buy, ever.",
    verdict: "The simplest engine on this list from an accounts standpoint - genuinely zero to think about here.",
    learnMore: { text: "godotengine.org", url: "https://godotengine.org" },
  },
  {
    name: "3. Unity",
    required: "Yes, mandatory. A free Unity ID is required to activate Unity Hub and the free Personal tier - the editor will not run without signing in.",
    whatFree: "The full Unity Personal tier: the complete editor, no feature caps, for any individual or organization under $200,000 in annual revenue or funding (a homeschool class project is nowhere near this).",
    whyNeeded: "Unity ties every install to an account for licensing/activation purposes. The same account also unlocks free Unity Learn courses and certification-prep material tied to your profile.",
    upgrade: "Unity Pro costs $2,310/year (or $210/month) and is legally required only once a project or company crosses that $200,000 threshold. It adds things like cloud build services, priority support, and removal of the Made With Unity splash screen - none of it relevant to a class project.",
    verdict: "Create the free Unity ID - it's required, not optional, to use Unity at all. Do not buy Unity Pro; nothing this class does will ever approach the revenue threshold that makes it necessary.",
    learnMore: { text: "unity.com/pricing", url: "https://unity.com/products/pricing-updates" },
  },
  {
    name: "4. GameMaker",
    required: "Yes, mandatory, even for the free non-commercial tier. GameMaker accounts are issued by Opera (GameMaker's parent company) - sign in with an existing Google/Apple/Facebook account or any email.",
    whatFree: "The complete editor and GML language, unlimited, for any non-commercial project - everything this class needs.",
    whyNeeded: "The account activates your license (even the free one) and is how GameMaker checks which tier you're on.",
    upgrade: "A one-time $99.99 Professional license, needed only if a student wants to sell or otherwise commercially publish a finished game after the class. Educational, non-commercial use (including sharing a build at the CoOp Expo) stays fully covered by the free tier. Unlike 001 Game Creator below, GameMaker's Professional license doesn't discount steeply - its best recorded Steam sale price is $79.99 (20% off), not the deep sales some other engines see. A separate Enterprise subscription exists only for publishing to game consoles (Xbox/PlayStation/Switch) - well outside the scope of this class.",
    verdict: "The free account is required and sufficient for everything in this course. The $99.99 upgrade (rarely discounted below about $80) is a reasonable, one-time cost worth considering only if a student is serious about selling a game commercially down the road - not something to buy for class.",
    learnMore: { text: "gamemaker.io/get", url: "https://gamemaker.io/en/get" },
  },
  {
    name: "5. GDevelop",
    required: "Optional for the desktop app, which works fully signed-out. An account is needed only to use GDevelop's online services: cloud project storage, one-click publishing to gd.games/Google Play/iOS, and multiplayer lobbies.",
    whatFree: "The complete open-source editor and engine, no feature caps, always. The free account tier adds 3 cloud project slots, 1 desktop/Android publish per day, and 40 AI credits/month.",
    whyNeeded: "Only relevant if a student wants to publish a build online or store projects in the cloud instead of just saving locally - not required for weekly homework done and saved on the student's own laptop.",
    upgrade: "Paid subscriptions - Silver (roughly $5-6/mo), Gold (roughly $11/mo), and Pro (roughly $33/mo) - raise the daily publish limits, cloud storage, and AI-credit allowances, and add perks like more player feedback slots and leaderboard styling. None of these limits are likely to be hit by a single student's weekly project.",
    verdict: "Skip the account entirely unless your student wants to publish a build online to share outside class - and even then, the free tier's daily limits are generous enough that no paid tier should be necessary for this course.",
    learnMore: { text: "gdevelop.io/pricing", url: "https://gdevelop.io/pricing" },
  },
  {
    name: "6. TIC-80",
    required: "No, never, for the free version - download it directly from tic80.com or GitHub with no signup of any kind.",
    whatFree: "The complete fantasy console: all eight supported scripting languages, the built-in sprite/map/sound editors, and local export - everything this class uses.",
    whyNeeded: "N/A for the free build.",
    upgrade: "A separate Pro version (roughly $5-10) is sold only on itch.io, so buying it requires a free itch.io account for the purchase transaction. Pro adds 8 memory banks instead of 1, text-format cartridge saving (useful for version control), and exporting a standalone game without needing the editor installed.",
    verdict: "No account needed for anything in this course. The Pro upgrade is a nice-to-have for a student who really takes to TIC-80's constraints and wants to export a shareable standalone build later - not required for any lesson here.",
    learnMore: { text: "tic80.com", url: "https://tic80.com" },
  },
  {
    name: "7. 001 Game Creator",
    required: "Yes. 001 Game Creator is distributed through Steam, so a free Steam account is required even to download the free trial.",
    whatFree: "A 7-day full-feature trial (everything except exporting/saving after the trial window closes) - enough to try the engine, not enough to finish a class project.",
    whyNeeded: "The Steam account is how the trial and any purchase are delivered and licensed - there's no way to get the software without one.",
    upgrade: "A Standalone one-time purchase (through Steam), list price $79.99, which removes the trial limits entirely with no ongoing cost and no royalties on anything a student later sells. It's worth waiting for a sale: SteamDB's recorded low is $11.99 (an 80% discount), and a roughly 75%-off price - about $20 - is common.",
    verdict: "Create the free Steam account to try the trial. If a student wants their own copy to keep working with past the class, this one is worth waiting on - 001 Game Creator routinely drops to $15-20, a fraction of its $79.99 list price, so there's rarely a reason to buy it at full price.",
    learnMore: { text: "001gamecreator.com", url: "https://001gamecreator.com" },
  },
  {
    name: "8. Unreal Engine",
    required: "Yes, mandatory. A free Epic Games account is required to download the Epic Games Launcher, which is required to install Unreal Engine at all.",
    whatFree: "The complete engine, every feature, no caps - Unreal does not have a reduced free tier the way some engines do.",
    whyNeeded: "Licensing and distribution both run through the Epic Games account - there is no way to install Unreal without one.",
    upgrade: "There is no paid 'Pro' tier to upgrade to. Unreal's only cost is a 5% royalty automatically owed on a game's revenue, and only on the portion above $1,000,000 in lifetime gross revenue - a threshold no classroom project will ever approach. A separate flat annual seat fee exists only for non-game commercial uses (architecture, film, product visualization), which doesn't apply here either.",
    verdict: "The free Epic Games account is required and is the only thing to set up - there is nothing to buy, upgrade, or worry about financially with Unreal for this class.",
    learnMore: { text: "unrealengine.com", url: "https://www.unrealengine.com" },
  },
];

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]
};

const doc = new Document({
  numbering: numberingConfig,
  styles: brand.stylesBlock,
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
    },
    headers: { default: brand.pageHeader("Creating Online Accounts") },
    footers: { default: brand.pageFooter() },
    children: [
      ...brand.titleBlock("Creating Online Accounts", "A Parent's Guide to Required and Optional Sign-Ups"),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "A companion to the Engine & Platform Guide", italics: true, size: 22, color: "666666" })] }),

      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
      new Paragraph({ children: [new PageBreak()] }),

      h1("Why This Guide"),
      p("Several of our eight supported engines require a free online account before you can even open the editor - others make an account entirely optional, and a couple never ask for one at all. On top of that, most engines have some kind of paid tier sitting alongside the free one. None of this is obvious just from looking at a download page, so this guide walks through each engine and answers three questions in plain terms: is an account required, what do you actually get for free, and is a paid upgrade ever worth it for a homeschool class project."),
      p("Short version: nothing in this course requires spending money. Every engine's free tier is enough to complete every lesson and the Semester 1 capstone. The paid upgrades described below are all genuinely optional - included here so you can make an informed choice if your student wants to keep building past the class, not because any of them are expected."),

      h1("Summary Table"),
      sumTable,
      new Paragraph({ spacing: { before: 300, after: 200 }, children: [] }),
      calloutBox("A note on accounts for younger students: several of these sign-ups (Unity, GameMaker, GDevelop, Steam, Epic Games) use general-purpose adult account systems with no dedicated \"kids mode.\" For middle-school-aged students, we recommend a parent creates and holds the login (using the parent's own email), rather than the student signing up independently."),
      new Paragraph({ children: [new PageBreak()] }),

      h1("Engine by Engine"),
      ...engines.flatMap(acctSection),

      h1("Frequently Asked Questions"),
      h3("Do we need to create every account before Lesson 1?"),
      p("Only the ones your chosen engine requires just to open the editor (Unity, GameMaker, 001 Game Creator via Steam, and Unreal). If your engine doesn't require one - Scratch, Godot, GDevelop's desktop app, or TIC-80 - there's nothing to set up in advance."),
      h3("Does creating an account ever cost money by itself?"),
      p("No. Every account on this list is free to create everywhere. Money only comes into play if a family chooses one of the optional paid upgrades described above - and none of those upgrades are needed to complete this course."),
      h3("Should my student sign up with their own email, or mine?"),
      p("For most of these services, we recommend a parent's email, especially for middle-school-aged students - see the callout box above. High-school students with their own email may reasonably sign up themselves, at your discretion."),
      h3("Is any engine's account gathering more personal data than the others?"),
      p("This guide covers whether an account is required and what it costs, not a full privacy comparison - each service's own privacy policy is the authoritative source if that's a deciding factor for your family. GDevelop does note it offers an anonymous, no-signup account option specifically for classroom/education use; ask the instructor if that applies to this class."),
      h3("If my student wants to keep using an engine after the class ends, what's the one upgrade actually worth considering?"),
      p("The two worth watching for a family that wants to keep going past the class: GameMaker's one-time $99.99 Professional license (rarely discounted much below $80), and 001 Game Creator's Standalone purchase, which lists at $79.99 but regularly goes on sale for $15-20 - patience pays off there. Both are pay-once-and-own, no subscription. Unity Pro and GDevelop's paid subscriptions, by contrast, are priced and scoped for needs well beyond a single student's hobby project."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Creating_Online_Accounts_Guide.docx", buffer);
  console.log("done");
});
