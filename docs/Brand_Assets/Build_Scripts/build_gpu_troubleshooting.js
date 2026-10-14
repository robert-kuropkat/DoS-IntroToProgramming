const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, VerticalAlign
} = require("docx");
const brand = require("./brand.js");

const CONTENT_WIDTH = 9360;
const border = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function bullet(text, level = 0, opts = {}) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun({ text, ...opts })] }); }
function numbered(text, ref, level = 0) { return new Paragraph({ numbering: { reference: ref, level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
function code(lines) {
  const children = [];
  lines.forEach((line, i) => {
    children.push(new TextRun({ text: line, font: "Consolas", size: 20 }));
    if (i < lines.length - 1) children.push(new TextRun({ text: "", break: 1 }));
  });
  return new Paragraph({
    spacing: { after: 160, before: 80 },
    indent: { left: 360 },
    shading: { fill: "F0F0F0", type: ShadingType.CLEAR },
    children
  });
}
function calloutBox(lines, fill = "FFF4E5") {
  const children = [];
  lines.forEach((line, i) => {
    children.push(new TextRun({ text: line, size: 21 }));
    if (i < lines.length - 1) children.push(new TextRun({ text: "", break: 1 }));
  });
  return new Paragraph({
    spacing: { after: 240, before: 80 },
    indent: { left: 200, right: 200 },
    shading: { fill, type: ShadingType.CLEAR },
    children
  });
}
function headerCell(text, width) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: brand.NAVY2, type: ShadingType.CLEAR }, margins: { top: 100, bottom: 100, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 20 })] })] });
}
function bodyCell(text, width, opts = {}) {
  return new TableCell({ borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: opts.fill || "FFFFFF", type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ children: [new TextRun({ text, size: 20, bold: !!opts.bold })] })] });
}
function simpleTable(headers, widths, rows) {
  return new Table({
    width: { size: CONTENT_WIDTH, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: headers.map((t, i) => headerCell(t, widths[i])) }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((t, i) => bodyCell(t, widths[i], { fill: ri % 2 ? "F3F6FA" : "FFFFFF", bold: i === 0 }))
      }))
    ]
  });
}

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } },
      { level: 1, format: LevelFormat.BULLET, text: "◦", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 1080, hanging: 360 } } } }] },
    { reference: "numFix1", levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
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
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } },
    headers: { default: brand.pageHeader("Older GPU / DirectX Troubleshooting (Unity)") },
    footers: { default: brand.pageFooter() },
    children: [
      brand.crestParagraph(1.3, { after: 160 }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Older GPU / DirectX Troubleshooting", bold: true, size: 36, color: brand.NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "What to do when Unity throws GPU fence or shader errors on older laptops", size: 22, color: brand.NAVY2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

      h1("When You'll Hit This"),
      p("Some student and instructor laptops - especially older business/office laptops with modest integrated or low-end discrete GPUs - only properly support DirectX 10, with partial (incomplete) DirectX 11 support. Unity 6.x defaults to newer rendering paths that assume more complete DirectX 11/12 support than this hardware actually provides. The result is a machine that installs Unity fine, opens projects fine, and even runs the Hello World exercise successfully - but throws GPU-related warnings and errors in the Console along the way."),
      p("This was first documented while testing the Unity install on an older laptop with an NVIDIA NVS 3100M GPU (a low-end mobile workstation chip). The same symptoms are likely on any similarly dated hardware - Intel integrated graphics from before ~2012, or comparable low-end NVIDIA/AMD mobile chips of that era."),

      h1("How to Recognize It"),
      p("Two errors in the Unity Console are the telltale signs:"),
      bullet("GpuFence::Create(): Failed to create ID3D11Fence, error 0x80070057", 0, { font: "Consolas", size: 20 }),
      bullet("Shader 'Hidden/Universal Render Pipeline/Edge Adaptive Spatial Upsampling' is not supported (in 'Blit FSR Upscaling'). PostProcessing render passes will not execute.", 0, { font: "Consolas", size: 20 }),
      p("If you see either of these - especially alongside a game that otherwise seems to run (the Hello World \"Debug.Log\" output still appears in Console) - this is very likely the cause, not a mistake in the exercise steps."),

      h1("Fix 1: Adjust Settings in Unity (Stay on the Current Version)"),
      p("This keeps you on the installed Unity version and works around the incompatible rendering features without downgrading anything."),
      h3("Disable post-processing on the default render pipeline"),
      numbered("Edit > Project Settings > Quality > Render Pipeline Asset.", "numFix1"),
      numbered("Open the URP asset assigned there (its default name is something like \"URP-...-Renderer\" or the project's default Universal Render Pipeline / Renderer 2D Data asset).", "numFix1"),
      numbered("In the Inspector for that Renderer 2D Data asset, find Post Processing and disable it.", "numFix1"),
      h3("Reorder the Graphics API so DirectX 11 (or OpenGL) is used instead of DirectX 12"),
      numbered("Edit > Project Settings > Player > Other Settings > Graphics APIs for Windows.", "numFix1"),
      numbered("Remove DirectX 12 support from the list entirely.", "numFix1"),
      numbered("Move OpenGLCore (Deprecated) above DirectX 11 in the list, so it's tried first - or simply leave DirectX 11 as the top entry if OpenGL isn't offered as an option.", "numFix1"),
      p("Restart the Unity Editor after changing the Graphics API list - Unity will prompt you to do this automatically."),

      h1("Fix 2: Downgrade to Unity 2023 LTS"),
      p("If Fix 1 doesn't fully resolve the errors, or if a student wants to do any 3D work at all, downgrading the project to Unity 2023.2 LTS (using DirectX 11) tends to be noticeably more stable on older hardware than Unity 6.x, and avoids the GPU fence error entirely."),

      h2("When to Downgrade vs. Stay on Unity 6.x"),
      simpleTable(
        ["Situation", "Recommendation"],
        [4680, 4680],
        [
          ["Focused on 2D work (sprites, tilemaps, 2D physics)", "Stay on Unity 6.x with OpenGL (Fix 1) - current setup works fine."],
          ["Want to try any 3D, even basic 3D", "Downgrade to Unity 2023.2 LTS + DirectX 11."],
          ["Current setup runs without crashes", "No need to change anything."],
          ["Want better stability generally on this hardware", "Downgrade to Unity 2023.2 LTS + DirectX 11."],
        ]
      ),

      h2("What Older Hardware Like This Can Still Do (3D)"),
      p("On a low-end GPU like the NVS 3100M, running Unity 2023.2 LTS with DirectX 11, students can still learn real 3D fundamentals:"),
      bullet("Low-poly 3D modeling (simple shapes, basic meshes)"),
      bullet("Basic lighting (baked lightmaps, simple point/directional lights)"),
      bullet("Standard shaders (Shader Model 3.0-4.0, no compute shaders)"),
      bullet("Simple animations (rigged characters, basic mechanics)"),
      bullet("Basic physics (colliders, rigidbodies)"),
      bullet("URP rendering, without advanced post-processing or volumetrics"),

      h2("What's Off the Table (Hardware-Limited)"),
      bullet("HDRP (requires Shader Model 5.0+ and compute shaders)"),
      bullet("Real-time global illumination"),
      bullet("Advanced GPU-based particle systems"),
      bullet("High-poly models with complex (PBR-heavy) shaders"),
      bullet("VR/AR development"),

      h2("Is Hardware Like This Actually Usable for Teaching 3D?"),
      p("Yes. Many successful indie games started on far weaker hardware than this. The formula that works within these limits:"),
      bullet("Low-poly, stylized art rather than photorealistic art"),
      bullet("Baked lighting rather than real-time lighting"),
      bullet("Efficient, simple shaders rather than PBR-heavy ones"),
      bullet("Small scenes rather than open-world scale"),
      p("Unity 2023.2 LTS with DirectX 11 on hardware like the NVS 3100M handles all of the above for learning purposes. Students can complete tutorials, build small 3D prototypes, and learn the core concepts without hitting a hard wall."),

      h1("Bottom Line"),
      calloutBox([
        "For 2D learning: stay on Unity 6.x + OpenGL (Fix 1) - the current setup already works.",
        "For 3D learning: downgrade to Unity 2023.2 LTS + DirectX 11 (Fix 2) - better compatibility, no GPU fence errors.",
        "Either way: older hardware will limit what a student can ship, but it will not block them from learning the fundamentals.",
      ]),
      p("Hardware this old is \"potato\" by modern standards, but it is still fully capable of teaching game development basics. The real limit is what a project can ship, not what a student can learn."),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Older_GPU_DirectX_Troubleshooting.docx", buffer);
  console.log("done");
});
