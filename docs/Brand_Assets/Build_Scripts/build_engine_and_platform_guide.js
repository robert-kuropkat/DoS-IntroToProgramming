const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, VerticalAlign
} = require("docx");
const brand = require("./brand.js");

const CONTENT_WIDTH = 9360; // US Letter, 1" margins

const border = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function link(text, url) {
  return new ExternalHyperlink({ children: [new TextRun({ text, style: "Hyperlink" })], link: url });
}

function h1(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] });
}
function h2(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] });
}
function h3(text) {
  return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] });
}
function p(text, opts = {}) {
  return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] });
}
function pRuns(runs) {
  return new Paragraph({ spacing: { after: 160 }, children: runs });
}
function bullet(text, level = 0) {
  return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] });
}
function numbered(text, level = 0) {
  return new Paragraph({ numbering: { reference: "numbers", level }, spacing: { after: 60 }, children: [new TextRun(text)] });
}
function labelBullet(label, text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { after: 60 },
    children: [new TextRun({ text: label + ": ", bold: true }), new TextRun(text)]
  });
}

function headerCell(text, width) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: brand.NAVY2, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 20 })] })]
  });
}
function bodyCell(text, width, opts = {}) {
  return new TableCell({
    borders,
    width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, size: 20, bold: !!opts.bold })] })]
  });
}

// ---------- Quick comparison table ----------
const compHeaders = ["Engine", "Coding Style", "Cost", "Min. RAM", "Install", "Best Fit"];
const compWidths = [1500, 1600, 1300, 1000, 1660, 2300];
const compRows = [
  ["Scratch", "Block", "Free", "~2 GB", "None (browser) or tiny app", "Absolute beginners, oldest laptops"],
  ["Godot", "Script (GDScript/C#)", "Free", "4 GB", "~150 MB, no install", "Best all-around, low-resource main engine"],
  ["Unity", "Script (C#) + Visual Scripting", "Free (Personal)", "8-16 GB", "Multi-GB, Hub required", "Industry-standard track, Unity Learn"],
  ["GameMaker", "Hybrid (GML / GML Visual)", "Free tier", "2-4 GB", "Small installer", "2D games, light hardware, visual to code bridge"],
  ["GDevelop", "Visual (event-based) + JS", "Free", "~2 GB", "None (browser) or small app", "Web fallback, weakest laptops"],
  ["TIC-80", "Script (Lua + others)", "Free", "~1 GB", "Single tiny executable", "Retro fantasy console, ultra-light"],
  ["001 Game Creator", "Hybrid (Visual + optional text)", "Free trial; Edu from $30/seat/yr", "512 MB", "Small installer - Windows only", "Lightest option here, classroom-ready - Windows-only"],
  ["Unreal Engine", "Blueprints + C++", "Free", "32 GB rec.", "Tens of GB", "Optional/advanced track, capable machines only"],
];
const compTable = new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: compWidths,
  rows: [
    new TableRow({ tableHeader: true, children: compHeaders.map((t, i) => headerCell(t, compWidths[i])) }),
    ...compRows.map((row, ri) => new TableRow({
      children: row.map((t, i) => bodyCell(t, compWidths[i], { fill: ri % 2 ? "F3F6FA" : "FFFFFF", bold: i === 0 }))
    }))
  ]
});

// ---------- device pairing table ----------
const deviceHeaders = ["Your Laptop", "Recommended Engines"];
const deviceWidths = [3000, 6360];
const deviceRows = [
  ["Older / 4-8 GB RAM, no dedicated graphics card", "Scratch, TIC-80, GDevelop, GameMaker - all run comfortably"],
  ["Mid-range / 8-16 GB RAM", "Godot, GameMaker, GDevelop (desktop) - add Unity if closer to 16 GB"],
  ["Newer / 16 GB+ RAM with a dedicated GPU", "Unity comfortably; Unreal Engine becomes realistic (32 GB is Epic's recommended spec)"],
  ["Any Windows laptop, even a very old one - but Windows only", "001 Game Creator (512 MB RAM minimum) - the lightest option here, but there is no Mac or Linux editor"],
];
const deviceTable = new Table({
  width: { size: CONTENT_WIDTH, type: WidthType.DXA },
  columnWidths: deviceWidths,
  rows: [
    new TableRow({ tableHeader: true, children: deviceHeaders.map((t, i) => headerCell(t, deviceWidths[i])) }),
    ...deviceRows.map((row, ri) => new TableRow({
      children: row.map((t, i) => bodyCell(t, deviceWidths[i], { fill: ri % 2 ? "F3F6FA" : "FFFFFF", bold: i === 0 }))
    }))
  ]
});

function engineSection({ name, tagline, style, cost, sysreq, installComplexity, pros, cons, steps, learnMore, note }) {
  return [
    h2(name),
    p(tagline, { italics: true, color: "555555" }),
    labelBullet("Coding style", style),
    labelBullet("Cost / licensing", cost),
    labelBullet("System requirements", sysreq),
    labelBullet("Install difficulty", installComplexity),
    h3("Pros"),
    ...pros.map(t => bullet(t)),
    h3("Cons"),
    ...cons.map(t => bullet(t)),
    h3("Download & Install"),
    ...steps.map((t, i) => numbered(t)),
    pRuns([new TextRun({ text: "Learn more: ", bold: true }), link(learnMore.text, learnMore.url)]),
    ...(note ? [new Paragraph({
      spacing: { after: 160, before: 80 },
      indent: { left: 200, right: 200 },
      shading: { fill: "FFF4E5", type: ShadingType.CLEAR },
      children: [new TextRun({ text: note, size: 21 })]
    })] : []),
    new Paragraph({ spacing: { after: 300 }, children: [] }),
  ];
}

const engines = [
  {
    name: "1. Scratch",
    tagline: "Block coding - the on-ramp. Free, MIT-developed, runs on anything.",
    style: "Pure block/drag-and-drop coding. No text code option.",
    cost: "Completely free, no account required to use the offline editor.",
    sysreq: "Any laptop from the last 15 years. Runs in a browser, or as a tiny offline desktop app.",
    installComplexity: "Trivial - use it directly at scratch.mit.edu, or install the free offline editor (no internet needed in class).",
    pros: [
      "Zero hardware risk - works on the oldest device in the room.",
      "Fastest path to a finished, playable game for a first-time coder.",
      "Huge library of kid-safe example projects to remix.",
    ],
    cons: [
      "Block-only - no path to real text-based code inside Scratch itself.",
      "Can feel young for older high schoolers if used for too long.",
    ],
    steps: [
      "Go to scratch.mit.edu/download and choose your OS (Windows/macOS/Linux), OR just use it in a browser at scratch.mit.edu - no login needed to experiment.",
      "Offline app: run the installer, no additional setup.",
      "Optional: create a free Scratch account to save projects to the cloud and share with classmates.",
    ],
    learnMore: { text: "scratch.mit.edu", url: "https://scratch.mit.edu" },
  },
  {
    name: "2. Godot Engine",
    tagline: "Our recommended default - free, light, and a real scripting language.",
    style: "Traditional text scripting in GDScript (Python-like, beginner-friendly) or C#.",
    cost: "100% free forever. Open source (MIT license). No account, no revenue thresholds, nothing to buy.",
    sysreq: "Comfortable minimum: dual-core CPU, 4 GB RAM, any GPU supporting OpenGL 3.3 (Godot's Compatibility rendering mode covers older integrated graphics). Recommended: 8 GB RAM.",
    installComplexity: "Easiest real-engine install on this list - a single ~150 MB file, no installer, no account.",
    pros: [
      "Best balance of real capability and low hardware demand.",
      "GDScript is one of the gentlest true coding languages to learn.",
      "No licensing strings attached, ever - good for families wary of subscriptions.",
    ],
    cons: [
      "Smaller tutorial/certification ecosystem than Unity.",
      "Less name-recognition on a resume than Unity or Unreal (though this is changing fast).",
    ],
    steps: [
      "Go to godotengine.org/download and select your OS.",
      "Download the Standard build (skip .NET/C# build unless you plan to code in C#).",
      "Unzip and run the single executable - nothing to install.",
      "If your laptop is older, go to Editor Settings > Rendering and switch to Compatibility renderer if the default is slow.",
    ],
    learnMore: { text: "godotengine.org", url: "https://godotengine.org" },
  },
  {
    name: "3. Unity",
    tagline: "The industry-standard track - free to learn, with a full certification path.",
    style: "Traditional text scripting in C#. A built-in Visual Scripting package (node graphs) is also available - see the Hybrid Engines section below.",
    cost: "Unity Personal is free with no seat limit as of 2026 (Personal is capped by company revenue, which does not apply to student hobby projects).",
    sysreq: "8 GB RAM is Unity's stated floor; 16 GB is a realistic comfort level for the Editor plus a class project. Any modern integrated GPU works for 2D/simple 3D.",
    installComplexity: "Moderate - install Unity Hub first, then a specific Editor version (multi-GB download) plus Visual Studio Code for C#. Budget real time/bandwidth for this on install day.",
    pros: [
      "Massive learning ecosystem: Unity Learn is free, with guided Pathways (Essentials, Junior Programmer, Creative Core, Game Development).",
      "A genuine path to industry certification (Pearson VUE exams, paid, with free prep material).",
      "C# is a widely-used, resume-relevant language.",
    ],
    cons: [
      "Heaviest install/update footprint of the main engines besides Unreal.",
      "C# has a steeper first-week learning curve than GDScript or block coding - better suited to older or more experienced students.",
    ],
    steps: [
      "Go to unity.com/download and install Unity Hub.",
      "In Unity Hub, install the latest LTS Editor version (this is the big download - several GB; do this on home wifi before class if possible).",
      "Install Visual Studio Code when prompted (Unity's recommended free code editor).",
      "Create a free Unity ID account to activate Unity Personal.",
      "Create a new 2D or 3D project from Unity Hub to confirm everything works.",
    ],
    learnMore: { text: "learn.unity.com", url: "https://learn.unity.com" },
    note: "Older laptop, GPU errors on install day? Some older/business laptops only have partial DirectX 11 support and will throw GPU fence or shader errors under Unity 6.x. See the companion troubleshooting document \"Older_GPU_DirectX_Troubleshooting.docx\" (in this same folder) for the two fixes - an in-editor settings change, or downgrading to Unity 2023 LTS.",
  },
  {
    name: "4. GameMaker",
    tagline: "A light 2D engine with a genuine visual-to-code bridge.",
    style: "Hybrid: GML (a C-like scripting language) or GML Visual (drag-and-drop blocks that generate real GML code live - a great stepping stone). See the Hybrid Engines section below for how the two connect.",
    cost: "Free tier (non-commercial) has no seat limits and covers everything needed for class. A one-time $99.99 Professional license is only needed if a student wants to publish/sell a finished game outside class.",
    sysreq: "Very light: 2 GB RAM minimum (8 GB comfortable), any OpenGL 4-class integrated graphics, ~3 GB disk space.",
    installComplexity: "Simple installer, everything bundled - no separate compiler to configure.",
    pros: [
      "Because GML Visual shows the real code being generated as you drag blocks, it's a uniquely good bridge from visual to text coding.",
      "Comfortably runs on aging BYOD hardware.",
      "2D-focused, which matches most first student game projects well.",
    ],
    cons: [
      "2D only - no 3D path if a student wants to go there later.",
      "Publishing a finished game outside of class requires the paid Professional license - worth flagging to families in advance.",
    ],
    steps: [
      "Go to gamemaker.io/download.",
      "Create a free account and choose the Free (non-commercial) plan.",
      "Run the installer for your OS.",
      "Sign in on first launch to activate the free license.",
    ],
    learnMore: { text: "gamemaker.io", url: "https://gamemaker.io" },
  },
  {
    name: "5. GDevelop - the web/no-install fallback",
    tagline: "Our recommended fallback for the weakest laptops or a bad install day.",
    style: "Visual, event-based programming (conditions to actions) by default, with inline JavaScript code blocks available for advanced logic - see the Hybrid Engines section below.",
    cost: "Completely free and open source. No feature caps, no account required for the core experience.",
    sysreq: "Runs entirely in a modern browser (Chrome/Firefox) - effectively no hardware requirement. A lightweight desktop app is also available for offline use.",
    installComplexity: "Easiest possible - open a browser tab, or install a small desktop app if you want to work offline.",
    pros: [
      "Zero-install option - the true fallback for a laptop that won't cooperate that day.",
      "Not meaningfully behind Godot for the scope of 2D games this class will build in a semester.",
      "Free with no feature ceiling, unlike some competing browser tools.",
    ],
    cons: [
      "Smaller tutorial community and less industry name-recognition than Unity/Godot.",
      "Not the tool of choice for high-end 3D work.",
    ],
    steps: [
      "Go to gdevelop.io and click Create a game to use it instantly in-browser - no install needed.",
      "For offline use: download the desktop app for your OS from the same site and run the installer.",
    ],
    learnMore: { text: "gdevelop.io", url: "https://gdevelop.io" },
  },
  {
    name: "6. TIC-80 - the retro option",
    tagline: "A tiny fantasy console for learning fundamentals without a huge project scope.",
    style: "Traditional text scripting - choose from Lua, JavaScript, Python, Ruby, Wren, Fennel, Squirrel, or Moonscript, all in the same tool.",
    cost: "Free and open source.",
    sysreq: "Trivial - runs well even on a Raspberry Pi. Any laptop handles this easily.",
    installComplexity: "Single small executable, no dependencies, no account.",
    pros: [
      "Built-in constraints (small screen, 16-color palette) force small, genuinely finishable projects - great for building confidence.",
      "Language flexibility means a stronger student can pivot to Python or JavaScript within the same tool.",
      "Works identically well on the oldest laptop in the room.",
    ],
    cons: [
      "Smaller built-in tutorial/community base than its paid cousin PICO-8 (a similar concept but costs $15 and has more third-party tutorials - not required for this class).",
      "Intentionally limited scope - not a path to a modern 3D game.",
    ],
    steps: [
      "Go to tic80.com and download the free build for your OS.",
      "Unzip and run the executable - no installer needed.",
      "Type 'new lua' (or the language of your choice) at the built-in console to start a project.",
    ],
    learnMore: { text: "tic80.com", url: "https://tic80.com" },
  },
  {
    name: "7. 001 Game Creator - the classroom-built option",
    tagline: "A dedicated no-coding engine, built by SoftWeir with classrooms specifically in mind. Windows only.",
    style: "Hybrid: graphical, node-based scripting by default (with plain-language tooltips instead of programming jargon), plus an optional text-code editor for the same underlying logic - see the Hybrid Engines section below for an important caveat about switching.",
    cost: "Standalone one-time purchase, normally $79.99 (frequently on sale around $20) - no royalties or income limits after that. Separately, a purpose-built Educational tier exists: a free Basic plan (watermarked output, 10 maps/project, exporting disabled) or paid per-workstation/per-year seats - Essentials ($30/workstation/year, unlocks Windows export) and Premium ($50/workstation/year, adds mobile/web/Steam export plus a bundled E-Book curriculum with exercises, assignments, and quizzes).",
    sysreq: "By far the lightest engine on this list: minimum 1.0 GHz CPU, 512 MB RAM, 1 GB disk, OpenGL 2.1, Windows 7. Recommended: 2.0 GHz CPU, 1 GB RAM, OpenGL 3.0, Windows 11.",
    installComplexity: "Simple installer. Important: the development editor is Windows-only - there is no Mac or Linux version. (Exported games can still target Windows, Steam, HTML5, Android, and iOS.)",
    pros: [
      "Lightest system requirements of any engine here - runs on essentially any Windows laptop, no matter how old.",
      "The Educational tier is genuinely built for a classroom: content is filtered (no blood/gore, weapons, or profanity), seats are licensed per workstation, and an optional E-Book bundle ships with exercises, assignments, and quizzes already written.",
      "Graphical scripting with plain-language tooltips can be a good fit for students intimidated by traditional code syntax, while still supporting text scripting for those who want it.",
    ],
    cons: [
      "Windows only for development - students on a Mac or Chromebook cannot install the editor at all, a real limitation for a BYOD class.",
      "The free Basic tier cannot export or build a shareable game; a paid Essentials seat ($30/workstation/year) is the practical minimum for a student who wants a build to share at the CoOp Expo.",
      "Much smaller community, tutorial base, and industry footprint than Unity, Godot, or GameMaker.",
    ],
    steps: [
      "Confirm you're on Windows 7 or later - there is no Mac/Linux editor for this engine.",
      "Go to 001gamecreator.com/downloads and try the free 7-day trial (full features except exporting/saving after the trial ends), or go to 001gamecreator.com/educational to inquire about a classroom Educational license.",
      "Run the installer.",
      "To get an exportable, shareable build for the Semester 2 capstone, a paid Standalone license or an Educational Essentials/Premium seat is required - the free tiers do not support exporting a finished build.",
    ],
    learnMore: { text: "001gamecreator.com", url: "https://001gamecreator.com" },
  },
  {
    name: "8. Unreal Engine - optional advanced track",
    tagline: "The most powerful engine here, and the heaviest. Available to students with capable hardware.",
    style: "True hybrid: Blueprints (visual scripting, can build a complete game with zero text code) or C++ - see the Hybrid Engines section below for how the two work together.",
    cost: "Free to download and use. 0% royalty until $1,000,000 in lifetime revenue - irrelevant for classroom use.",
    sysreq: "Epic's own recommended spec is around 32 GB RAM and a GPU with 8 GB+ VRAM; an SSD is effectively required. This is beyond most BYOD laptops.",
    installComplexity: "Heaviest of all options - tens of GB, often paired with a full Visual Studio install for C++ work.",
    pros: [
      "Blueprints offer a real visual-scripting path to a fully playable game, not just a toy.",
      "Highest ceiling for students who want to go further into professional 3D production.",
    ],
    cons: [
      "Not realistic for most BYOD laptops, especially older ones - this is confirmed by Epic's own recommended specs.",
      "Longest install time - plan for a home download, not a shared-wifi lab download.",
    ],
    steps: [
      "Only attempt this if your laptop has 16 GB+ RAM and a dedicated graphics card.",
      "Create a free Epic Games account at unrealengine.com.",
      "Install the Epic Games Launcher, then install Unreal Engine from within it (this is a large download - do it at home on good wifi, not during class).",
      "For C++ work, also install Visual Studio Community (free) with the Game development with C++ workload.",
    ],
    learnMore: { text: "unrealengine.com", url: "https://www.unrealengine.com" },
  },
];

// ---------- Hybrid Engines section ----------
function hybridEntry({ name, styles, interact, gained, limited, bottomLine }) {
  return [
    h3(name),
    labelBullet("Two styles", styles),
    labelBullet("How they interact", interact),
    labelBullet("What you gain switching", gained),
    labelBullet("What's limited", limited),
    p(bottomLine, { italics: true, color: "555555" }),
  ];
}

const hybrids = [
  {
    name: "Unity - Visual Scripting <-> C#",
    styles: "Visual Scripting (node graphs) and C# (text code).",
    interact: "They coexist in the same project rather than converting into each other - different GameObjects can use either system, and a Visual Scripting graph can call directly into your own C# classes.",
    gained: "A student can start building real logic with nodes before learning C# syntax, then call into hand-written C# functions as their skills grow.",
    limited: "There is no official \"convert this graph to C#\" button - a graph stays a graph. Graphs reference C# code by name/reflection, so renaming a C# method can silently break a graph node that used it.",
    bottomLine: "Think of Unity's two styles as two toolkits that can talk to each other, not two views of the same logic.",
  },
  {
    name: "GameMaker - GML Visual <-> GML",
    styles: "GML Visual (drag-and-drop icon blocks) and GML (C-like text code).",
    interact: "Right-click any event's action list and choose \"Convert To GML Code\" to turn that one event's blocks into real GML text - conversion happens per event, not for the whole project at once.",
    gained: "Students see the actual code their blocks were generating the whole time - one of the best visual-to-text \"aha\" moments in the whole curriculum.",
    limited: "The conversion is one-way. GameMaker warns before converting that the generated text cannot be turned back into blocks - once converted, that event is GML from then on.",
    bottomLine: "Use \"Convert to GML Code\" deliberately, and on a simple event or a copy first - a class demo shouldn't accidentally destroy someone's visual work.",
  },
  {
    name: "001 Game Creator - Script Editor <-> Textual Scripting",
    styles: "The node-based Script Editor (default) and Textual Scripting (a BASIC-like text mode).",
    interact: "A script's mode can be switched from Graphical to Textual.",
    gained: "A curious student can see what plain, BASIC-style code looks like under the hood.",
    limited: "Per 001 Game Creator's own wiki, this conversion is one-way and destructive - switching back to Graphical loses the original events - and the wiki states plainly that Textual Scripting \"offers no advantages over visual scripting\" and is considered legacy.",
    bottomLine: "Treat 001 Game Creator as visual-only in this class. Textual Scripting is worth knowing about, but not worth building on.",
  },
  {
    name: "Unreal Engine - Blueprints <-> C++",
    styles: "Blueprints (visual node graphs) and C++ (text code).",
    interact: "This is the one true two-way bridge on our list. C++ functions and classes can be exposed directly to Blueprints (using UFUNCTION/UCLASS), Blueprint graphs can call them, and a Blueprint class can subclass a C++ parent class - both styles run in the same project and call into each other freely.",
    gained: "This mirrors how professional Unreal teams actually work: designers build and iterate in Blueprints while programmers expose gameplay systems from C++.",
    limited: "There's no automatic \"Blueprint to C++\" converter available to us - an older feature called Blueprint Nativization did this, but Epic has deprecated it. Converting a Blueprint to C++ today means a programmer manually rewriting it by hand.",
    bottomLine: "Unreal is the one engine where \"switching styles\" really means designers and programmers collaborating across styles, not one style converting into the other. (Forward-looking note: Epic has announced Blueprints and the classic Actor system are slated for eventual replacement by a new language called Verse in UE6 - but our target version, UE5, keeps Blueprints fully supported, so nothing changes for this class.)",
  },
  {
    name: "GDevelop - Event Sheets <-> JavaScript",
    styles: "Visual event sheets (conditions -> actions) and inline JavaScript code blocks.",
    interact: "The two are freely mixed - a JavaScript Code event can sit right inside a normal event sheet alongside ordinary visual conditions and actions.",
    gained: "A JavaScript Code event is a trapdoor for anything the visual event system can't express - genuinely useful once a student outgrows the built-in actions.",
    limited: "A JavaScript Code event doesn't automatically see every object in your scene - you have to explicitly choose which objects to \"pass in\" to the code block first.",
    bottomLine: "GDevelop's mixing is the most flexible of the five - closer to Unreal's two-way bridge than to GameMaker's one-way door.",
  },
];

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    { reference: "numbers", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
  ]
};

const doc = new Document({
  numbering: numberingConfig,
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: "Arial", color: brand.NAVY },
        paragraph: { spacing: { before: 360, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: brand.NAVY2 },
        paragraph: { spacing: { before: 260, after: 140 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: "444444" },
        paragraph: { spacing: { before: 160, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
    },
    headers: {
      default: brand.pageHeader("Engine & Platform Guide")
    },
    footers: {
      default: brand.pageFooter()
    },
    children: [
      brand.crestParagraph(1.3, { after: 160 }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "SLHE Game Programming Class", bold: true, size: 40, color: brand.NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [new TextRun({ text: "Engine & Platform Guide", size: 30, color: brand.NAVY2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Please read before Lesson 1 - Thursday, August 13, 2026", italics: true, size: 22, color: "666666" })] }),

      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
      new Paragraph({ children: [new PageBreak()] }),

      h1("Why This Guide"),
      p("This class lets you choose your own game engine - there is no single required platform. Every lesson's core lecture is taught live in one engine, but you'll always get working example code and instructions for every engine we support, so you can follow along and build in whichever tool fits your device, your interests, and your goals. You're also welcome to switch engines partway through the year - that's one reason we keep parallel examples for everything."),
      p("This guide exists so you and your student can choose - and install - a platform before the first class. Installing large engines (especially Unity and Unreal) over shared classroom wifi with a room full of students is slow; please install at home ahead of time wherever possible. Lesson 1's lab time is reserved for installation help and troubleshooting, but a head start makes that time much more useful."),

      h1("Quick Comparison"),
      compTable,
      new Paragraph({ spacing: { before: 300, after: 200 }, children: [] }),

      h1("Which Engine Fits My Laptop?"),
      p("As a Bring-Your-Own-Device class, hardware varies a lot. Use this as a starting filter - you can always try more than one."),
      deviceTable,
      new Paragraph({ children: [new PageBreak()] }),

      h1("Engine Profiles"),
      p("Each profile below covers coding style, cost, hardware needs, honest pros/cons, and step-by-step install instructions."),
      ...engines.flatMap(engineSection),

      h1("Hybrid Engines: Switching Between Styles"),
      p("Three distinct coding styles run through this whole class: Block (Scratch), Visual/node-based (GDevelop, GameMaker's GML Visual, Unity's Visual Scripting, 001 Game Creator, Unreal's Blueprints), and Scripting/text-based (Godot, GameMaker's GML, Unity's C#, TIC-80). Five of our engines are hybrids that offer two of these styles inside the very same project. That makes them worth a closer look on their own: a hybrid engine lets a student start in whichever style is comfortable and grow toward the other without switching tools, and it's a great case study in the difference between the styles, since you can watch the same logic represented two different ways. The five hybrid engines don't all work the same way, though - some let you convert one direction only, some don't convert at all but let both styles call into each other, and one shouldn't really be converted in class at all. The breakdown below is what to actually expect from each."),
      ...hybrids.flatMap(hybridEntry),
      p("The 30-Lesson Master Schedule calls out two \"Style Spotlight\" moments where this happens live in lecture: GameMaker reappears in Lesson 9 in GML text mode after being demoed in GML Visual mode back in Lesson 5, and Unity reappears in Lesson 12 in C# mode after being demoed in Visual Scripting mode in Lesson 7. Both are a chance to watch the same kind of problem solved in the other style, back to back, a few weeks apart.", { italics: true, color: "555555" }),
      new Paragraph({ children: [new PageBreak()] }),

      h1("A Note on the Name Game Creator"),
      p("This class supports 001 Game Creator (001gamecreator.com, by SoftWeir) - see its full profile above. If you search around and see a different tool called Game Creator 2, that's an unrelated paid visual-scripting add-on built on top of Unity, not the engine we mean. We mention it here only so families researching options aren't thrown off by the similar name."),

      h1("Frequently Asked Questions"),
      h3("Do I have to buy anything?"),
      p("No. Every engine in the core lineup (Scratch, Godot, Unity, GameMaker, GDevelop, TIC-80) is free for classroom use. Unreal Engine is also free, but requires stronger hardware. 001 Game Creator has a free trial and a free Educational Basic tier, though exporting a finished game requires a paid tier (Standalone purchase, or an Educational Essentials/Premium seat)."),
      h3("What if my laptop is really old or slow?"),
      p("Start with Scratch, GDevelop (in a browser - no install at all), TIC-80, or (if you're on Windows) 001 Game Creator. All run comfortably on hardware from a decade or more ago, and none will be a wasted investment of time."),
      h3("Can I change engines partway through the year?"),
      p("Yes. Because every lesson includes examples across all supported engines, switching is always an option - you won't be starting from zero."),
      h3("Which engine should I pick if I just want the safest, most flexible choice?"),
      p("Godot. It's free with no strings attached, runs on almost anything, and teaches a real scripting language without Unity's heavier install and steeper first weeks."),
      h3("Which engine has the strongest path to an industry certification?"),
      p("Unity, via the free Unity Learn Pathways and the paid Pearson VUE certification exams."),
      h3("My laptop is a Mac or Chromebook - can I use 001 Game Creator?"),
      p("No - 001 Game Creator's development editor only runs on Windows. Mac and Chromebook students should choose from Scratch, Godot, Unity, GameMaker, GDevelop, or TIC-80 instead; during any week 001 Game Creator is featured in lecture, you'll follow the same lesson using your own engine's parallel example."),
      h3("What are Block, Visual, and Scripting, and why do they keep getting mentioned?"),
      p("They're the three broad styles of writing game logic used across this class: Block (drag-and-snap puzzle pieces, like Scratch), Visual/node-based (connecting events, conditions, and actions without typing syntax, like GDevelop or Blueprints), and Scripting (typed text code in a real programming language, like GDScript or C#). Every lesson's featured engine rotates through all three roughly evenly across the year, so you'll get real exposure to each - see the Hybrid Engines section above and the 30-Lesson Master Schedule's Coding Style column."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Engine_and_Platform_Guide.docx", buffer);
  console.log("done");
});
