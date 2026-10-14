const pptxgen = require("pptxgenjs");

const NAVY = "16305C";
const NAVY2 = "1C4A86";
const GOLD = "C99A2E";
const GOLD_L = "E8C25A";
const CREAM = "F4F0E6";
const TEXT_DARK = "222222";
const TEXT_MUTE = "5A6472";
const HEAD_FONT = "Cambria";
const BODY_FONT = "Calibri";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
const SW = 13.333, SH = 7.5;

function footerObjs(dark) {
  return [
    { text: {
        text: "VIDEO GAME CREATION",
        options: { x: 0.5, y: 7.14, w: 5, h: 0.3, fontFace: BODY_FONT, fontSize: 10, color: dark ? GOLD_L : TEXT_MUTE, charSpacing: 2, bold: true, margin: 0 },
    } },
  ];
}

// ---------- MASTER 1: TITLE SLIDE ----------
pres.defineSlideMaster({
  title: "SLHE_TITLE",
  background: { path: "bg_title_with_vgc.png" },
  objects: [
    { text: { text: "S O U T H   L A K E   H O M E   E D U C A T O R S",
        options: { x: 0, y: 0.55, w: SW, h: 0.35, align: "center", fontFace: BODY_FONT, fontSize: 13, color: GOLD_L, charSpacing: 1, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: 0, y: 3.7, w: SW, h: 0.75, align: "center", fontFace: HEAD_FONT, fontSize: 44, bold: true, color: GOLD_L, margin: 0 },
        text: "Lesson Title Goes Here" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 1.0, y: 4.5, w: SW - 2, h: 0.7, align: "center", fontFace: BODY_FONT, fontSize: 22, color: CREAM, margin: 0 },
        text: "A short description of what this lesson/deck covers" } },
    { placeholder: { options: { name: "date", type: "body", x: 0, y: 5.35, w: SW, h: 0.4, align: "center", fontFace: BODY_FONT, fontSize: 14, italic: true, color: "9FB6D6", margin: 0 },
        text: "Date goes here" } },
    ...footerObjs(true),
  ],
});

// ---------- MASTER 2: SECTION DIVIDER ----------
pres.defineSlideMaster({
  title: "SLHE_SECTION",
  background: { path: "bg_title.png" },
  objects: [
    { text: { text: "SECTION",
        options: { x: 1.5, y: 2.15, w: 9, h: 0.4, fontFace: BODY_FONT, fontSize: 13, bold: true, color: GOLD, charSpacing: 2, margin: 0 } } },
    { placeholder: { options: { name: "title", type: "title", x: 1.5, y: 2.6, w: 10, h: 1.0, fontFace: HEAD_FONT, fontSize: 40, bold: true, color: GOLD_L, valign: "middle", margin: 0 },
        text: "Section Title" } },
    { placeholder: { options: { name: "teaser", type: "body", x: 1.55, y: 3.65, w: 9.8, h: 0.7, fontFace: BODY_FONT, fontSize: 16, color: CREAM, margin: 0 },
        text: "One sentence describing what this section covers." } },
    ...footerObjs(true),
  ],
});

// ---------- MASTER 3: CONTENT ----------
pres.defineSlideMaster({
  title: "SLHE_CONTENT",
  background: { path: "bg_content.png" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.45, w: 11, h: 0.7, fontFace: HEAD_FONT, fontSize: 34, bold: true, color: NAVY2, margin: 0 },
        text: "Slide Title" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.6, y: 1.5, w: 12.1, h: 5.3, fontFace: BODY_FONT, fontSize: 16, color: TEXT_DARK, margin: 0 },
        text: "Click to add your content. Use the Home tab's bullet-list button for bullet points, just like in Word." } },
    ...footerObjs(false),
  ],
});

// ---------- MASTER 4: CLOSING ----------
pres.defineSlideMaster({
  title: "SLHE_CLOSING",
  background: { path: "bg_title_with_slhe.png" },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0, y: 2.75, w: SW, h: 0.8, align: "center", fontFace: HEAD_FONT, fontSize: 38, bold: true, color: GOLD_L, margin: 0 },
        text: "Thank You!" } },
    { placeholder: { options: { name: "message", type: "body", x: 1.5, y: 3.6, w: SW - 3, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 17, color: CREAM, margin: 0 },
        text: "A closing message, next steps, or reminder goes here." } },
    ...footerObjs(true),
  ],
});

// =========================================================================
// Example slides - one per layout, with real starter text. Duplicate any of
// these (right-click the slide thumbnail > Duplicate Slide) to add more, or
// use Home > New Slide > pick a layout from this template.
// =========================================================================

let s1 = pres.addSlide({ masterName: "SLHE_TITLE" });
s1.addText("Lesson Title Goes Here", { placeholder: "title" });
s1.addText("A short description of what this lesson/deck covers", { placeholder: "subtitle" });
s1.addText("Date goes here", { placeholder: "date" });

let s2 = pres.addSlide({ masterName: "SLHE_SECTION" });
s2.addText("Section Title", { placeholder: "title" });
s2.addText("One sentence describing what this section covers.", { placeholder: "teaser" });

let s3 = pres.addSlide({ masterName: "SLHE_CONTENT" });
s3.addText("Slide Title", { placeholder: "title" });
s3.addText(
  [
    { text: "Click to add your content.", options: { bullet: true, breakLine: true } },
    { text: "Use the Home tab's bullet-list button for bullet points, just like in Word.", options: { bullet: true, breakLine: true } },
    { text: "Duplicate this slide (right-click its thumbnail on the left) to add more Content slides.", options: { bullet: true } },
  ],
  { placeholder: "body" }
);

let s4 = pres.addSlide({ masterName: "SLHE_CLOSING" });
s4.addText("Thank You!", { placeholder: "title" });
s4.addText("A closing message, next steps, or reminder goes here.", { placeholder: "message" });

pres.writeFile({ fileName: "SLHE_VGC_PowerPoint_Template.pptx" }).then(() => console.log("done"));
