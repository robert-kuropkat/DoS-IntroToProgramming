const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, LevelFormat, ExternalHyperlink,
  TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak, VerticalAlign, ImageRun
} = require("docx");
const brand = require("./brand.js");

const CONTENT_WIDTH = 9360;
const border = { style: BorderStyle.SINGLE, size: 2, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function link(text, url) { return new ExternalHyperlink({ children: [new TextRun({ text, style: "Hyperlink" })], link: url }); }
function h1(text) { return new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(text)] }); }
function h2(text) { return new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(text)] }); }
function h3(text) { return new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(text)] }); }
function p(text, opts = {}) { return new Paragraph({ spacing: { after: 160 }, children: [new TextRun({ text, ...opts })] }); }
function bullet(text, level = 0) { return new Paragraph({ numbering: { reference: "bullets", level }, spacing: { after: 60 }, children: [new TextRun(text)] }); }
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

// ---- screenshot embedding ----
function imgSize(path, data) {
  // PNG: width/height are big-endian uint32 at bytes 16 and 20 of the IHDR chunk.
  const w = data.readUInt32BE(16);
  const h = data.readUInt32BE(20);
  return { w, h };
}
function screenshot(filePath, caption) {
  const data = fs.readFileSync(filePath);
  const { w, h } = imgSize(filePath, data);
  const targetW = 460;
  const targetH = Math.round(targetW * (h / w));
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 120, after: 40 },
      children: [new ImageRun({ data, type: "png", transformation: { width: targetW, height: targetH } })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: caption, italics: true, size: 18, color: "555555" })],
    }),
  ];
}

function engineSection({ name, installSteps, helloSteps, expected, trouble, code: codeLines, images, imageFolder, walkthroughNote }, idx) {
  const installRef = `numbers-install-${idx}`;
  const helloRef = `numbers-hello-${idx}`;
  const out = [
    h2(name),
    h3("Detailed Install"),
    ...installSteps.map(s => numbered(s, installRef)),
    h3("Hello, World! Verification"),
    ...helloSteps.map(s => numbered(s, helloRef)),
  ];
  if (codeLines) out.push(code(codeLines));
  out.push(
    new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text: "Expected result: ", bold: true }), new TextRun(expected)] }),
    new Paragraph({ spacing: { after: 300 }, children: [new TextRun({ text: "If something's wrong: ", bold: true, italics: true }), new TextRun({ text: trouble, italics: true })] }),
  );
  if (images && images.length) {
    out.push(h3(`Visual Walkthrough${walkthroughNote ? ` (${walkthroughNote})` : ""}`));
    out.push(p("Screenshots from an actual run-through, in order. Circled/boxed elements mark exactly what to click or check at that step.", { italics: true, color: "555555" }));
    images.forEach(([file, caption]) => out.push(...screenshot(`${imageFolder}/${file}`, caption)));
  }
  return out;
}

const engines = [
  {
    name: "1. Scratch",
    installSteps: [
      "Go to scratch.mit.edu/download - this opens the \"Download the Scratch App\" page.",
      "Under Choose your OS, confirm Windows (or macOS) is selected, then click Get it from Microsoft - Windows now installs Scratch through the Microsoft Store; macOS still offers a direct .dmg download.",
      "On the Microsoft Store listing page, click Download. The browser downloads Scratch 3 Installer.exe into your Downloads - click Open file in the download flyout to run it.",
      "A small Microsoft Store window confirms the install and shows a download/progress bar. Let it finish - it's a fairly large app (~265 MB).",
      "Scratch launches automatically once installed. On first launch, a one-time dialog asks whether to share usage data with the Scratch Team - choose either option and click Close.",
      "Confirm the editor opens to a blank project with the cat sprite on stage.",
    ],
    helloSteps: [
      "With Scratch open to a blank project, click the Events category in the block palette (left side).",
      "Drag \"when green flag clicked\" into the empty code area to the right of the palette.",
      "Click the Looks category, then drag \"say Hello! for 2 seconds\" so it snaps directly underneath the when-clicked block.",
      "Click into the say block's text field (it highlights blue) and change the text to \"Hello, World!\"",
      "Click the green flag above the stage (top-right of the stage area).",
      "Optional: File > Save to your computer, name it HelloWorld, and click Save to keep a copy of the project (.sb3 file).",
    ],
    expected: "The cat sprite shows a speech bubble reading \"Hello, World!\" for 2 seconds, and the script glows yellow while it runs.",
    trouble: "If the Microsoft Store install seems stuck, check your internet connection - it's a large download for a first-run app. If Scratch won't launch, confirm you're on Windows 10+ or macOS 10.13+; the browser version at scratch.mit.edu (click Create) works as a backup with no install at all.",
    imageFolder: "scratch_annotated",
    walkthroughNote: "captured on Scratch 3.32.0",
    images: [
      ["Scratch_01_Dowload.PNG", "The Download page - Windows is selected; click Get it from Microsoft."],
      ["Scratch_02_MSStore.PNG", "The Microsoft Store listing for Scratch 3 - click Download."],
      ["Scratch_03_Downloaded.PNG", "The installer downloads as Scratch 3 Installer.exe - click Open file."],
      ["Scratch_04_Splash.PNG", "The Microsoft Store install dialog appears."],
      ["Scratch_05_Installing.PNG", "Download and install in progress (~265 MB)."],
      ["Scratch_06_UsageData.PNG", "First launch: 1) choose whether to share usage data, 2) click Close."],
      ["Scratch_07_IDE.PNG", "A blank project opens with the cat sprite on stage, ready to go."],
      ["Scratch_08_Events.PNG", "1) Click Events in the left navigation, 2) drag \"when green flag clicked\" from the palette, 3) into the code area (arrow shows the drag)."],
      ["Scratch_09_Looks.PNG", "1) Click Looks in the left navigation, 2) drag \"say Hello! for 2 seconds\", 3) snap it under the when-clicked block (arrow shows the drag)."],
      ["Scratch_10_TextChange.PNG", "1) Click into the say block's text field and change it to \"Hello, World!\", 2) then use this green flag to run it."],
      ["Scratch_11_HelloWorld.PNG", "Success - 1) the speech bubble reads \"Hello, World!\", 2) the stop button confirms the script ran."],
      ["Scratch_12_Save.PNG", "Optional: 1) click File, 2) choose Save to your computer."],
      ["Scratch_13_Save.PNG", "1) Name the file HelloWorld, 2) click Save. Arrows mark the folder path, New folder button, and Save as type field."],
    ],
  },
  {
    name: "2. Godot Engine",
    installSteps: [
      "Go to godotengine.org/download - this opens the Godot 4 download page for your OS.",
      "Click the Godot Engine button (the version badge, e.g. 4.7.1, confirms the release) - skip the \"Godot Engine - .NET\" button unless you plan to demo C# specifically.",
      "The browser downloads a .zip file (e.g. Godot_v4.7.1-stable_win64.exe.zip).",
      "In File Explorer, find the zip in your Downloads folder, right-click it, and choose Extract All...",
      "Confirm the destination folder (the default, right next to the zip, is fine) and click Extract.",
      "Extraction produces two files: a standard executable and a matching _console version (which also opens a background terminal window) - you only need the standard one.",
      "Optional: drag the standard executable (not the _console one) to your Desktop for a shortcut you can double-click anytime.",
      "Double-click the standard Godot executable to launch - there is no installer.",
      "If Windows SmartScreen warns \"Windows protected your PC,\" click More info, then Run anyway (expected for unsigned indie software).",
    ],
    helloSteps: [
      "In the Project Manager window, click Create New Project.",
      "Set Project Name to Hello World and Project Path to a folder of your choice - the Renderer and Version Control Metadata defaults are both fine to leave (or set Version Control Metadata to None if you don't plan to use Git).",
      "Click Create - since Edit Now is checked by default, the editor opens automatically afterward.",
      "The first time a project opens, Godot spends a few seconds (Re)Importing Assets - this is normal, just wait for it to finish.",
      "In the Scene panel, the empty project automatically prompts you to choose a root node type - click Other Node.",
      "In the Create New Node dialog, type Node2D into the Search bar, select it from the list, and click Create.",
      "With the Node2D root selected, scroll to the bottom of the Inspector panel (right side) to the Script field, click its dropdown arrow, and choose New Script...",
      "In the Attach Node Script dialog, keep all the defaults and click Create.",
      "The script opens with a default _ready() function containing a placeholder pass line - replace it with the code below.",
      "Notice the unsaved indicators: the editor tab, the main title bar, the script editor's own title bar, and the script's entry in the Filter Scripts list all show an asterisk (*) - press Ctrl+S to save the script.",
      "Press F6 (Run Current Scene). Since the scene has never been saved, Godot prompts you to save it first - keep the default filename (e.g. node_2d.tscn) and click Save.",
    ],
    code: ['func _ready():', '    print("Hello, World!")'],
    expected: "\"Hello, World!\" appears in the Output panel at the bottom of the editor. A small blank \"Hello World (DEBUG)\" game window also pops up - that's expected, since this exercise only prints text and doesn't draw anything to the screen. You may need to move or close that window to see the Output panel underneath it.",
    trouble: "If nothing shows in Output, confirm the scene was actually saved before running (the asterisks from the step above should be gone), and that the script is attached to the Node2D that's the scene's root.",
    imageFolder: "godot_annotated",
    walkthroughNote: "captured on Godot Engine 4.7.1",
    images: [
      ["Godot_01_DownloadPage.PNG", "The Godot 4 download page - click the Godot Engine button (not the .NET version)."],
      ["Godot_02_Downloading.PNG", "The browser downloads a .zip file."],
      ["Godot_03_FileExplorer.PNG", "1) The Downloads folder, 2) the downloaded zip file."],
      ["Godot_04_RightClickFile.PNG", "Right-click the zip file and choose Extract All..."],
      ["Godot_05_Location.PNG", "1) Confirm the destination folder, 2) click Extract."],
      ["Godot_06_Extracting.PNG", "Extraction in progress."],
      ["Godot_07_DragToDesktop.PNG", "1) The standard executable (not the _console version) - optionally drag it to the Desktop for a shortcut."],
      ["Godot_08_Defender.PNG", "Windows SmartScreen may appear on first launch - click More info."],
      ["Godot_09_RunAnyway.PNG", "Click Run anyway to continue (expected for unsigned indie software)."],
      ["Godot_10_CreateProject.PNG", "The Project Manager - click Create New Project."],
      ["Godot_11_NewProject.PNG", "The dialog's defaults: 1) Project Name, 2) Project Path, 3) Version Control Metadata."],
      ["Godot_12_NewProjectHelloWorld.PNG", "Set for this exercise: 1) Project Name = Hello World, 2) Project Path, 3) Version Control Metadata = None, 4) click Create."],
      ["Godot_13_ProjectSetup.PNG", "First load: Godot (Re)Imports Assets - just wait for it to finish."],
      ["Godot_14_ProjectWorkspace.PNG", "The empty scene prompts for a root node: 1) the + icon, 2) click Other Node."],
      ["Godot_15_NewNodeDialog.PNG", "1) Search for Node2D, 2) select it, 3) click Create."],
      ["Godot_16_Node2D.PNG", "1) The Node2D root, 2) the Inspector panel, 3) the Script field's dropdown, 4) choose New Script..."],
      ["Godot_17_NewScriptDialog.PNG", "Keep all the defaults and click Create."],
      ["Godot_18_NewScriptDefault.PNG", "The default _ready() function - replace the pass placeholder with the print line."],
      ["Godot_19_EditedScript.PNG", "1) The edited _ready() function, 2-5) all four unsaved indicators: main title bar, editor tab, script list, and the script editor's own title bar - press Ctrl+S to save."],
      ["Godot_20_SaveScene.PNG", "Running for the first time prompts a save - keep the default name and click Save."],
      ["Godot_21_FileSavedIndicators.PNG", "Saved: the asterisks are gone from all three places."],
      ["Godot_22_HelloWorld.PNG", "Success - 1) \"Hello, World!\" appears in the Output panel, 2) the Output tab. You may need to move or close the small blank game window to see it."],
    ],
  },
  {
    name: "3. Unity",
    installSteps: [
      "Go to unity.com and create a Unity account - use an existing Google or Apple account, or sign up with any email address (click Sign in instead if you already have one).",
      "Sign in to your new (or existing) Unity account.",
      "On the download page, for Windows the x64 download is correct unless you know you need Arm64.",
      "The browser downloads the Unity Hub installer.",
      "Click Open file in the downloads flyout to launch the Unity Hub Setup installer.",
      "Choose an install location (the default is fine) and click Install.",
      "Wait for Unity Hub itself to install.",
      "When complete, leave Run Unity Hub checked and click Finish.",
      "Unity Hub opens showing its Software License Terms - use the scroll bar to reach the bottom of the text; the Agree button lights up once you have.",
      "Click the blue Agree button.",
      "Sign in to Unity Hub - this is the same account you just created or signed into on the website.",
      "This reopens your browser to confirm the sign-in; click Open (checking \"Always allow\" is optional - it just skips this confirmation next time).",
      "Agree to the Unity Personal license terms.",
      "Unity Hub may begin downloading the latest Editor version automatically - click the Downloads icon in the upper right to watch its progress.",
      "If the download does not start automatically, skip ahead to the next step to start it manually from the Installs page.",
      "Once downloading, all four items (Editor application, Web Build Support, Documentation, Get Started With Unity) should finish with green checkmarks - this can take a while.",
      "To start the install manually (or add components later): click Installs in the left navigation, click the Manage dropdown next to your Editor version, choose Add Modules, then click Install Editor.",
      "In the Add Modules dialog, make sure Microsoft Visual Studio Community is checked, and scroll down to also check Documentation near the bottom of the list - a blue checkmark means an item is not yet installed and has been preselected for you.",
      "The Visual Studio Installer launches on its own to install the code editor.",
      "Let the Visual Studio Community install finish - this is a large download and can take a while the first time.",
    ],
    helloSteps: [
      "In Unity Hub, click Projects in the left navigation, then click either New Project button.",
      "Open the template filter dropdown and choose Core.",
      "Select the Universal 2D template (shown in the panel on the right), set a Project Name (e.g. HelloWorld) and Location, leave Source Code Provider set to none, then click Create Project.",
      "Accept the Unity Editor Software Terms: check the box, then click Accept and Continue.",
      "Unity's own splash screen appears while the editor loads - this can take a while on an older machine; watch the green progress bar at the bottom.",
      "If a Windows Defender Firewall prompt appears asking for network access, click Allow access - it can appear behind other windows, so watch for a glowing taskbar icon if you don't see it right away.",
      "Unity finishes importing its starter packages (shader graphs, default assets, etc.) - just wait for it to finish.",
      "In the Hierarchy panel, right-click empty space and choose Create Empty.",
      "With the new GameObject selected, in the Inspector panel click Add Component.",
      "Type \"script\" into the filter field, click New Script in the results, then click the arrow to continue.",
      "Name the script HelloWorld and click Create and Add.",
      "In the Save As dialog, confirm the file will save into your project's Assets/Scripts folder - creating a Scripts folder is optional but recommended - and click Save.",
      "Unity compiles the new script in the background; once done, confirm the Scripts folder and the HelloWorld script both appear in the Assets panel, and that the Script field in the Inspector shows it attached.",
      "Single-click the script in the Assets panel to preview its contents in the Inspector, or double-click it to open it for editing in Visual Studio.",
      "If Visual Studio prompts you to sign in, it's safe to click \"Skip and add accounts later\" for this class - a Microsoft or Xbox account works too if a student already has one.",
      "If prompted about GitHub Copilot, click \"Maybe later\" - we won't use it for this class, though students are welcome to experiment with it on their own.",
      "The very first time you open a script from Unity, Visual Studio may not load it automatically - if that happens, close Visual Studio and double-click the script in Unity's Project panel again.",
      "In Visual Studio, replace the body of the Start() function with the line below.",
      "Save the script (Ctrl+S) and switch back to Unity - it recompiles automatically.",
      "Press the Play button at the top of the Unity editor, and confirm the updated Start() method is visible in the Inspector's script preview.",
      "Click the Console tab and confirm \"Hello, World!\" appears in the output, then click the Stop button.",
      "Stop, press Ctrl+S once more to save the script, then exit.",
    ],
    code: ['void Start() {', '    Debug.Log("Hello, World!");', '}'],
    expected: "\"Hello, World!\" appears as a log entry in the Console panel (Window > General > Console if it isn't visible).",
    trouble: "If the script won't compile, check the Console for a red error first - a missing semicolon or brace is the most common cause. If Play mode does nothing, confirm the GameObject with the script is still in the Hierarchy. On some older laptops with only partial DirectX 11 support, you may also see a GpuFence::Create() error or a shader warning about \"Edge Adaptive Spatial Upsampling\" in the Console - these do not stop Hello World from working, but if you plan to do any 3D work, see the companion document Older_GPU_DirectX_Troubleshooting.docx (in this same folder) for the fix.",
    imageFolder: "unity_annotated",
    walkthroughNote: "captured on Unity 6.5 (6000.5.4f1)",
    images: [
      ["Unity_01_CreateUnityAccount.PNG", "The Create account page - if you already have a Unity account, click Sign in instead."],
      ["Unity_02_SignInToUnity.PNG", "Sign in with your chosen method (Google, Apple, or email)."],
      ["Unity_03_DownloadPage.PNG", "For Windows, the x64 download is correct unless you know you need Arm64."],
      ["Unity_04_DownloadingHub.PNG", "The Unity Hub installer downloading."],
      ["Unity_05_OpenFile.PNG", "Click Open file to launch the installer."],
      ["Unity_06_HubInstallLocation.PNG", "Choose an install location (the default is fine) and click Install."],
      ["Unity_07_HubInstalling.PNG", "Unity Hub installing."],
      ["Unity_08_HubInstallComplete.PNG", "1) Leave Run Unity Hub checked, 2) click Finish."],
      ["Unity_09_UnityLicense.PNG", "1) Scroll to the bottom of the license text, 2) until the Agree button lights up."],
      ["Unity_10_LicenseAgree.PNG", "Click the blue Agree button."],
      ["Unity_11_HubSignIn.PNG", "Click Sign in to Unity Hub with the same account."],
      ["Unity_12_HubSignedIn.PNG", "Confirming in the browser: 1) \"Always allow\" is optional, 2) click Open."],
      ["Unity_13_PersonalLicense.PNG", "Click Agree to accept the Unity Personal license."],
      ["Unity_14_UntyHub.PNG", "The Downloads icon, upper right - click to watch install progress."],
      ["Unity_15_AutoDownload.PNG", "The Editor and its components downloading automatically."],
      ["Unity_16_AutoDownloadComplete.PNG", "All items complete with green checkmarks."],
      ["Unity_17_AddModules.PNG", "To install manually instead: 1) Installs in the left nav, 2) Manage dropdown, 3) Add Modules, 4) Install Editor."],
      ["Unity_18_AddModules.PNG", "Confirm Microsoft Visual Studio Community is checked (scroll down for Documentation too) and click Continue."],
      ["Unity_19_VisualStudioSplash.PNG", "The Visual Studio Installer launches automatically."],
      ["Unity_20_VisualStudioInstaller.PNG", "Visual Studio Community installing - this can take a while the first time."],
      ["Unity_21_Projects.PNG", "1) Projects in the left nav, 2) New Project (middle), 3) New Project (top right)."],
      ["Unity_22_NewProjectDefault.PNG", "1) Open the template filter dropdown, 2) choose Core."],
      ["Unity_23_ProjectSettings.PNG", "1) Select Universal 2D, 2) confirm the right panel, 3) Project Name, 4) Location, 5) Create Project (leave Source Code Provider on none)."],
      ["Unity_24_UnitySoftwareTerms.PNG", "1) Check the box, 2) click Accept and Continue."],
      ["Unity_25_SplashScreen.PNG", "Unity's own splash screen - watch the green progress bar; this can take a while."],
      ["Unity_26_Defender.PNG", "If Windows Defender Firewall prompts, click Allow access."],
      ["Unity_27_PackageImport.PNG", "Unity imports its starter packages - note the shader warning at the bottom; this ties to the GPU troubleshooting doc, not to an error in these steps."],
      ["Unity_28_Editor.PNG", "The Editor is ready - the Hierarchy panel is boxed."],
      ["Unity_29_RightClick.PNG", "Right-click empty space in the Hierarchy and choose Create Empty."],
      ["Unity_30_EmptyObject.PNG", "1) The new GameObject in the Hierarchy, 2) click Add Component."],
      ["Unity_31_AddComponent.PNG", "1) Type \"script\" into the filter, 2) choose New Script, 3) click the arrow to continue."],
      ["Unity_32_ScriptName.PNG", "1) Name field set to HelloWorld, 2) click Create and Add."],
      ["Unity_33_SaveScript.PNG", "1) The breadcrumb path, 2) the file name, 3) click Save (a Scripts folder is optional but recommended)."],
      ["Unity_34_CompileScripts.PNG", "1) The compiling status, 2) the Scripts folder in Assets, 3) the Script field in the Inspector."],
      ["Unity_35_ScriptsFolder.PNG", "1) The HelloWorld script, 2) the breadcrumb path - single-click to preview it in the Inspector (arrow), or double-click to edit."],
      ["Unity_36_VisualStudioSignIn.PNG", "If prompted, \"Skip and add accounts later\" is fine for this class."],
      ["Unity_37_Copilot.PNG", "If prompted about Copilot, \"Maybe later\" - not needed for this class."],
      ["Unity_38_EditScript.PNG", "The HelloWorld script in the Visual Studio Project Explorer."],
      ["Unity_39_EditedScript.PNG", "The edited Start() function, boxed."],
      ["Unity_40_ReadyToPlay.PNG", "1) The Play button, 2) the updated Start() method shown in the Inspector's script preview."],
      ["Unity_41_Play.PNG", "1) The Console tab, 2) \"Hello, World!\" in the output, 3) the Stop button. The boxed line below it is the shader warning covered in the GPU troubleshooting doc."],
      ["Unity_42_RenderWarning.PNG", "A closer look at the Console: the Hello World log line together with the FSR shader warning some older GPUs show - see Older_GPU_DirectX_Troubleshooting.docx."],
      ["Unity_43_BiRPError.PNG", "On hardware with only partial DirectX 11 support, a GpuFence::Create() error may also appear - see Older_GPU_DirectX_Troubleshooting.docx for the fix."],
    ],
  },
  {
    name: "4. GameMaker",
    installSteps: [
      "Go to gamemaker.io/en/download - accept cookies if prompted.",
      "Click Download under Windows.",
      "The browser downloads the GameMaker installer (a large file, ~184 MB).",
      "Click Open file in the downloads flyout to launch it.",
      "You're taken to a sign-in page - click Register / Sign In.",
      "GameMaker accounts are handled by Opera (the company behind GameMaker) - sign in with an existing Google, Apple, or Facebook account, or any email address.",
      "Once signed in, you land on your GameMaker account page - the free tier is all this class needs, no purchase required.",
      "The installer unpacks itself.",
      "The installer's own splash screen appears (this build is the LTS - Long Term Stable - release).",
      "Choose a language and click OK.",
      "Review the License Agreement and click I Agree.",
      "On the Choose Components screen, Install Opera GX (a companion web browser) is optional and checked by default - leave it checked or uncheck it, either is fine for this class.",
      "Choose an install location (the default is fine) and click Install.",
      "Wait for GameMaker itself to install.",
      "If Opera GX was selected, its own installer opens - review the terms and click Accept & Install.",
      "Opera may show a usage-statistics screen - your choice, then continue.",
      "Opera GX installs.",
      "Opera GX opens with its own first-run welcome tour - Skip is fine, this class doesn't need it.",
      "Meanwhile, GameMaker itself continues installing (you may see a Visual C++ Redistributable step run automatically).",
      "The GameMaker installer confirms Installation Complete - click Next.",
      "Click Finish to close the installer.",
    ],
    helloSteps: [
      "Launch GameMaker. On first start, watch the Installing indicator in the upper right while it finishes setting itself up.",
      "If a Windows Defender Firewall prompt appears, click Allow access - GameMaker needs this to check for updates and license your account; you may need admin privileges to approve it.",
      "On the Welcome To GameMaker screen, a Set Up Wizard walks you through creating your first project - click the Game project type icon to begin (or click New directly if the wizard isn't showing).",
      "1) Confirm Game is selected as the project type, 2) select the Blank Game template (its description confirms what it creates), 3) set Project Name to HelloWorld, 4) confirm the Location, 5) click Let's Go!",
      "The new project opens directly into the Project Space, with a short built-in tour pointing out the toolbar, Asset Browser, Inspector, and Output Window.",
      "In the Asset Browser panel (top right), click the + icon, then Create, then choose Object from the submenu.",
      "GameMaker creates a new object and opens it for editing - 1) rename it from the default to obj_hello in the Asset Browser, 2) confirm the same name updates in the object editor's Name field, 3) click Add Event, 4) choose Draw from the event list, 5) choose the plain Draw event from the submenu (not Draw GUI or the others).",
      "The first time you add a code-based event, GameMaker asks which scripting language to use for the project - 1) choose GML Code (not GML Visual), 2) click OK.",
      "A code editor opens for the Draw event, already positioned to type - as you type, GameMaker's autocomplete suggests matching functions (for example, typing draw_t lists every draw_text-style function available).",
      "Type the line below. Notice the Draw tab now shows an asterisk (*) - the code hasn't been saved yet; press Ctrl+S to save it.",
      "Double-click Room1 in the Asset Browser to open the Room Editor.",
      "Drag obj_hello from the Asset Browser into the room canvas to place an instance of it - 1) it now appears in the room's own instance list just below the search bar, 2) notice the asterisk (*) that has appeared in the window's title bar, meaning the whole project has unsaved changes - press Ctrl+S to save everything, 3) click the Run (Play) button in the toolbar to test.",
    ],
    code: ['draw_text(100, 100, "Hello, World!");'],
    expected: "A game window titled HelloWorld opens showing the text \"Hello, World!\" near the top-left in the default font.",
    trouble: "If nothing appears, confirm the object instance was actually dropped into the room (check the Room Editor's instance list, not just the Asset Browser) and that the Draw event code has no red underlines or errors in the Feather Messages panel at the bottom. If GameMaker prompts to save before running, that's expected the first time - accept it and continue.",
    imageFolder: "gamemaker_annotated",
    walkthroughNote: "captured on GameMaker LTS2026",
    images: [
      ["GameMaker_01_Cookies.PNG", "The download page - accept cookies to continue."],
      ["GameMaker_02_Download.PNG", "Click Download under Windows."],
      ["GameMaker_03_Downloading.PNG", "The installer downloading (~184 MB)."],
      ["GameMaker_04_Downloaded.PNG", "Click Open file to launch the installer."],
      ["GameMaker_05_CreateAccount.PNG", "Click Register / Sign In to create or sign in to a GameMaker account."],
      ["GameMaker_07_GameMakerAccount.PNG", "Your GameMaker account page - the free tier is all this class needs."],
      ["GameMaker_08_UnpackingInstaller.PNG", "The installer unpacking itself."],
      ["GameMaker_09_InstallerSplashScreen.PNG", "The installer's splash screen (LTS build)."],
      ["GameMaker_10_LanguageSelection.PNG", "Choose a language and click OK."],
      ["GameMaker_11_License.PNG", "Click I Agree to accept the license."],
      ["GameMaker_12_Components.PNG", "Install Opera GX is optional - leave it checked or uncheck it, either is fine."],
      ["GameMaker_13_InstallLocation.PNG", "Confirm the install location and click Install."],
      ["GameMaker_14_Installing.PNG", "GameMaker installing."],
      ["GameMaker_16_OperaLicense.PNG", "If Opera GX was selected, its own installer screen appears - click Accept & Install to continue."],
      ["GameMaker_17_UsageStatistics.PNG", "Opera's usage-statistics choice - your preference either way."],
      ["GameMaker_18_OperaInstalling.PNG", "Opera GX installing."],
      ["GameMaker_19_OperaBrowser.PNG", "Opera GX's first-run tour - Skip is fine for this class."],
      ["GameMaker_20_GameMakerInstalling.PNG", "GameMaker continues installing its remaining components."],
      ["GameMaker_21_InstallComplete.PNG", "Installation Complete - click Next."],
      ["GameMaker_22_Finish.PNG", "Click Finish to close the installer."],
      ["GameMaker_23_FirstStart.PNG", "On first launch, watch the Installing indicator in the upper right while GameMaker finishes setting up."],
      ["GameMaker_24_FirewallVerification.PNG", "If Windows Defender Firewall prompts, click Allow access (admin privileges may be required)."],
      ["GameMaker_25_GameMakerScreen.PNG", "The Set Up Wizard - click the Game project type icon to begin."],
      ["GameMaker_26_ProjectType.PNG", "1) Game project type, 2) Blank Game template (arrow shows its description), 3) Project Name, 4) Location, 5) Let's Go!"],
      ["GameMaker_27_ProjectSpace.PNG", "The new project opens with a built-in guided tour of the workspace."],
      ["GameMaker_28_CreateObject.PNG", "1) The Asset Browser tab, 2) Create in the first popup menu, 3) Object in the submenu."],
      ["GameMaker_29_Workspace.PNG", "1) obj_hello renamed in the Asset Browser, 2) the same name in the object editor, 3) Add Event, 4) Draw in the first popup, 5) Draw in the submenu."],
      ["GameMaker_30_ScriptingPreference.PNG", "First code event: 1) choose GML Code, 2) click OK."],
      ["GameMaker_31_CodeAssist.PNG", "GameMaker's autocomplete suggests matching functions as you type."],
      ["GameMaker_32_CodeNotSaved.PNG", "The asterisk on the Draw tab means the code isn't saved yet - press Ctrl+S to save it."],
      ["GameMaker_33_RoomEditor.PNG", "Double-click Room1 in the Asset Browser to open the Room Editor."],
      ["GameMaker_34_AddObjectToRoom.PNG", "Drag obj_hello into the room (arrow) - 1) it appears in the room's instance list, 2) the unsaved asterisk in the title bar (Ctrl+S saves everything), 3) click Run to test."],
      ["GameMaker_35_HelloWorld.PNG", "Success - \"Hello, World!\" appears in the game window."],
    ],
  },
  {
    name: "5. GDevelop",
    installSteps: [
      "Go to gdevelop.io/download and click the Windows icon under Download the desktop app (the free, no-account browser version at gdevelop.io also works as a no-install backup).",
      "The browser downloads the GDevelop installer (~142 MB).",
      "Click Open file in the downloads flyout to launch it.",
      "Choose who the software installs for - Anyone who uses this computer needs admin credentials; Only for me doesn't. Either is fine for this class.",
      "Confirm the install location (the default is fine) and click Install.",
      "Wait for GDevelop to install.",
      "On the Completing GDevelop 5 Setup screen, leave Run GDevelop 5 checked and click Finish.",
    ],
    helloSteps: [
      "On first launch, GDevelop's Start Page opens - click Log in if you already have an account, or Sign up to create one.",
      "Sign in with Google, GitHub, Apple, or an email address, then click Login (or click Create an account if you're new).",
      "Once signed in, your account name appears top right - click Create in the left-hand navigation.",
      "Click + Create new game.",
      "Choose Empty project from the template grid.",
      "Pick a screen size (Desktop & Mobile landscape works fine for this class), set Project name to HelloWorld, confirm Where to store this project (GDevelop Cloud or Your computer both work), then click Create new game.",
      "The new project opens directly into the Scene Editor - in the Objects panel (right side), click Add object.",
      "Search \"text\" and choose Text (Displays a text on the screen).",
      "In the Edit dialog, set Object name to HelloWorld and Initial text to display to \"Hello, World!\", then click Apply.",
      "Drag the new HelloWorld object from the Objects panel into the middle of the scene canvas - notice the save icon in the top-left toolbar now shows an unsaved dot; press Ctrl+S to save.",
      "Click Preview in the toolbar to test.",
      "If a Windows Defender Firewall prompt appears, click Allow access - GDevelop's preview needs this to run its local server.",
    ],
    expected: "A separate \"Preview of HelloWorld\" window opens showing the text \"Hello, World!\" near the center of the screen.",
    trouble: "If nothing appears, confirm the object was actually dragged into the scene canvas (not just created in the Objects panel) and that Ctrl+S saved the change. If the Preview window doesn't seem to open at all, check for a Windows Defender Firewall prompt hiding behind other windows.",
    imageFolder: "gdevelop_annotated",
    walkthroughNote: "captured on GDevelop 5.6.274",
    images: [
      ["GDevelop_01_DownloadPage.PNG", "Click the Windows icon under Download the desktop app."],
      ["GDevelop_02_Downloading.PNG", "The installer downloading (~142 MB)."],
      ["GDevelop_03_Downloaded.PNG", "Click Open file to launch the installer."],
      ["GDevelop_04_InstallationFor.PNG", "Choose who the software installs for - \"Anyone\" needs admin credentials, \"Only for me\" doesn't. Either is fine for this class."],
      ["GDevelop_05_InstallLocation.PNG", "Confirm the install location and click Install."],
      ["GDevelop_06_Installing.PNG", "GDevelop installing."],
      ["GDevelop_07_InstallComplete.PNG", "1) Leave Run GDevelop 5 checked, 2) click Finish."],
      ["GDevelop_08_StartPage.PNG", "The Start Page - 1) Log in if you already have an account, 2) Sign up to create one."],
      ["GDevelop_09_AccountLogin.PNG", "1) Sign in with Google, GitHub, Apple, or email and click Login, 2) or click Create an account if you're new."],
      ["GDevelop_10_StartPageLoggedIn.PNG", "1) Your account name appears top right once signed in, 2) click Create in the left navigation."],
      ["GDevelop_11_CreateGamePage.PNG", "Click + Create new game."],
      ["GDevelop_12_CreateGame.PNG", "Choose Empty project."],
      ["GDevelop_13_CreateBlankGameDefault.PNG", "Pick a screen size - Desktop & Mobile landscape works fine for this class."],
      ["GDevelop_14_CreateBlankGameInfo.PNG", "1) Set Project name to HelloWorld, 2) confirm where to store the project, 3) click Create new game."],
      ["GDevelop_15_Workspace.PNG", "The Scene Editor opens directly - click Add object in the Objects panel."],
      ["GDevelop_16_CreateObject.PNG", "1) Search \"text\", 2) choose Text (Displays a text on the screen)."],
      ["GDevelop_17_TextObject.PNG", "1) Set Object name to HelloWorld, 2) set Initial text to display to \"Hello, World!\", 3) click Apply."],
      ["GDevelop_18_AddToScene.PNG", "Drag HelloWorld into the scene (arrow) - 1) it appears in Scene Objects, 2) the save icon shows an unsaved dot - press Ctrl+S, 3) click Preview to test."],
      ["GDevelop_19_DefenderFirewall.PNG", "If Windows Defender Firewall prompts, click Allow access - GDevelop's preview needs this for its local server."],
      ["GDevelop_20_HelloWorld.PNG", "Success - a separate Preview window shows \"Hello, World!\""],
    ],
  },
  {
    name: "6. TIC-80",
    installSteps: [
      "Go to tic80.com/create and scroll down to Download - click tic80-v1.1-win.zip under Windows (the free build is all this class needs).",
      "The browser downloads the zip - click Open file in the downloads flyout to reveal it in File Explorer (or just open your Downloads folder yourself).",
      "Right-click the downloaded zip file and choose Extract All...",
      "Confirm the destination folder (the default, right next to the zip, is fine) and click Extract.",
      "Drag the extracted tic80 application to your Desktop for a shortcut you can double-click anytime.",
      "Optional but recommended: right-click the tic80 icon, choose Properties, then on the Shortcut tab - 1) add a space followed by --fs \"C:\\Dev\\Tic80\" to the end of the Target field (two dashes before fs, the path in double quotes), 2) set Start in to that same folder, 3) click Apply. TIC-80's default working folder can be hard to find from the operating system - changing these properties makes your saved projects easy to locate afterward.",
      "Double-click the tic80 icon to launch - there is no installer.",
    ],
    helloSteps: [
      "TIC-80 opens directly to its console prompt with a \"hello! type help for help\" message.",
      "Type mkdir HelloWorld and press Enter to create a project folder.",
      "Type dir and press Enter to confirm the new folder is there.",
      "Type cd HelloWorld and press Enter to move into it.",
      "Type new lua and press Enter - this creates a new Lua cart. Press Esc to open the Code Editor if it doesn't open automatically; it already contains TIC-80's default starter demo code.",
      "In the Code Editor, delete the default t/x/y variable lines and the whole function TIC() body underneath them - we don't need this starter demo for Hello World.",
      "Type the code below into the now-empty function TIC() body (feel free to update the title/author/desc/site comment lines at the top with your own info too).",
      "Press Esc to return to the console, then type run and press Enter (or press Ctrl+R) to test.",
      "Once it works, press Esc again to return to the console and type save HelloWorld, then press Enter to save your cart.",
    ],
    code: ['function TIC()', '  cls(0)', '  print("Hello, World!", 40, 60)', 'end'],
    expected: "The game screen shows \"Hello, World!\" printed near the middle-left of the screen in TIC-80's default font.",
    trouble: "Make sure the TIC() function name is capitalized exactly as shown - TIC-80 calls this function every frame, and a typo means nothing will run. If commands like mkdir or cd don't seem to work, check for typos - the console is case-sensitive for folder names you create yourself.",
    imageFolder: "tic80_annotated",
    walkthroughNote: "captured on TIC-80 1.1.2837",
    images: [
      ["Tic80_01_DownloadPage.PNG", "Scroll to Download and click tic80-v1.1-win.zip under Windows."],
      ["Tic80_02_Downloaded.PNG", "Click Open file to reveal the zip in File Explorer."],
      ["Tic80_03_ZipFile.PNG", "Right-click the downloaded zip file and choose Extract All..."],
      ["Tic80_04_ExtractAll.PNG", "Confirm the destination folder and click Extract."],
      ["Tic80_05_MoveToDesktop.PNG", "Drag the extracted tic80 application to your Desktop (arrow) for an easy shortcut."],
      ["Tic80_06_ChangeWorkingFolder.PNG", "Optional: in the shortcut's Properties, 1) add the --fs \"C:\\Dev\\Tic80\" flag to Target, 2) set Start in to match, 3) click Apply - this makes your saved projects easy to find later."],
      ["Tic80_07_StartUp.PNG", "TIC-80 opens directly to its console prompt."],
      ["Tic80_08_ProjectFolder.PNG", "1) mkdir HelloWorld creates the project folder, 2) dir confirms it's there, 3) cd HelloWorld moves into it, 4) new lua creates a new cart and opens the Code Editor."],
      ["Tic80_09_CodeEditor.PNG", "The default starter demo code - 1) delete these variables, 2) and this whole function TIC() body underneath them."],
      ["Tic80_10_NewCode.PNG", "The new code in place - the print line is what draws \"Hello, World!\" to the screen."],
      ["Tic80_11_HelloWorld.PNG", "Success - \"Hello, World!\" appears on screen."],
    ],
  },
  {
    name: "7. 001 Game Creator (Windows only)",
    installSteps: [
      "In your Steam Library, select 001 Game Creator (Windows only, no Mac/Linux version exists) and click INSTALL.",
      "Confirm the install options (Create desktop shortcut and Create start menu shortcut are both fine to leave checked) and click Install.",
      "Review and click Accept on the End-User License Agreement.",
      "Wait for Steam to download and install (~2.4 GB).",
      "Once installed, click Launch from the Library page.",
    ],
    helloSteps: [
      "001 Game Creator shows its own splash screen while it loads.",
      "A New Game window may pop up automatically - click Cancel to close it for now; we want to set our save location first, and that option isn't available from this screen.",
      "Go to File > Open Game... instead.",
      "In the New / Open Game window, click into the path field at the bottom (it defaults to a Documents folder) and choose Browse... from the popup that appears, then select the folder where you'd like this class's projects saved. Once the path is set, click Create New Game.",
      "This reopens the New Game window - 1) select Blank Template, 2) set Name of Project to HelloWorld, 3) click OK.",
      "A Warning explains the Blank Template is intended for experienced users and doesn't include input triggers, menus, or actor templates - safe to ignore for this project; click Continue.",
      "The main Workspace opens with your new HelloWorld project loaded in the left-hand project list.",
      "Go to the Resources menu (not Game) and choose System Triggers.",
      "Double-click Introduction - this system trigger runs automatically at the very start of the game.",
      "In the Script Editor, double-click Message Box in the events list on the right - a direct double-click drops it onto the script and connects it to Start automatically (dragging it out is not reliable and may not work).",
      "In the Message Box properties, set Message to \"Hello, World!\", then click OK.",
      "Click the Play (Test Game) button in the main toolbar to test.",
      "A Testing Options window appears - the defaults are fine; confirm Emulation is set to None, then click OK to launch the test.",
    ],
    expected: "A debugger panel and your game's window both open, with a message box on screen displaying \"Hello, World!\"",
    trouble: "If nothing happens on Play, confirm the Message Box event is actually connected to the Start node with a visible line between them - double-clicking Message Box in the events list (rather than dragging it onto the canvas) connects it automatically, which is the reliable way to do this. If the New Game window keeps popping up before you can set your save folder, remember to close it (Cancel) and use File > Open Game instead - the save-location picker only appears on that screen, not on the New Game template screen.",
    imageFolder: "001gc_annotated",
    walkthroughNote: "captured on 001 Game Creator 2025.011.000 via Steam",
    images: [
      ["001GameCreator_01_SteamLibrary.PNG", "In your Steam Library, select 001 Game Creator and click INSTALL."],
      ["001GameCreator_02_InstallLocation.PNG", "Confirm the install options and click Install."],
      ["001GameCreator_03_License.PNG", "Review the EULA and click Accept."],
      ["001GameCreator_04_Downloading.PNG", "Steam downloads and installs (~2.4 GB)."],
      ["001GameCreator_05_InstallComplete.PNG", "Once installed, click Launch to start 001 Game Creator."],
      ["001GameCreator_06_SplashScreen.PNG", "001 Game Creator's splash screen while it loads."],
      ["001GameCreator_07_NewGame.PNG", "A New Game window may pop up automatically - click Cancel to close it for now."],
      ["001GameCreator_08_OpenGame.PNG", "Go to File > Open Game... instead."],
      ["001GameCreator_09_Browse.PNG", "1) Click the path field and choose Browse... to set your save folder, 2) once it's set, click Create New Game."],
      ["001GameCreator_10_NewGame2.PNG", "Back in the New Game window: 1) select Blank Template, 2) set Name of Project to HelloWorld, 3) click OK."],
      ["001GameCreator_11_Warning.PNG", "Safe to ignore for this project - click Continue."],
      ["001GameCreator_12_Workspace.PNG", "The main Workspace opens with your new HelloWorld project loaded on the left."],
      ["001GameCreator_13_SystemTriggers.PNG", "1) Open the Resources menu (not Game), 2) choose System Triggers."],
      ["001GameCreator_14_IntroductionTrigger.PNG", "Double-click Introduction - this runs automatically at the very start of the game."],
      ["001GameCreator_15_ScriptEditor.PNG", "In the Script Editor, double-click Message Box in the events list on the right."],
      ["001GameCreator_16_MessageBox.PNG", "Set the Message field to \"Hello, World!\", then click OK."],
      ["001GameCreator_17_GameTest.PNG", "1) Click the Play (Test Game) toolbar button, confirm Emulation is set to None, 2) click OK to launch the test."],
      ["001GameCreator_18_HelloWorld.PNG", "Success - the message box displays \"Hello, World!\""],
    ],
  },
  {
    name: "8. Unreal Engine (optional advanced track)",
    installSteps: [
      "Go to unrealengine.com and create a free Epic Games account if needed.",
      "Download and install the Epic Games Launcher.",
      "Sign in, go to the Unreal Engine tab, click Install Engine, choose the latest version.",
      "Allow the very large (tens of GB) download/install to finish - do this well ahead of time on a strong connection, ideally days before class.",
    ],
    helloSteps: [
      "Launch Unreal Engine. In current versions this opens its own standalone app (separate from the Epic Games Launcher itself), with a Home screen - click New Project.",
      "Choose the Games category > Blank template, confirm BLUEPRINT is selected (not C++), set the Project Name to HelloWorld, then click Create.",
      "Once the editor opens, click the Blueprints icon in the main toolbar and choose Open Level Blueprint from the menu.",
      "In the Event Graph, note that an Event BeginPlay node is already there by default - no need to add it.",
      "Right-click in empty space, type \"print\" to search, and add Print String. Connect the white execution line from Event BeginPlay's output to Print String's input.",
      "Click into the Print String node's In String field and set it to \"Hello, World!\"",
      "Click Compile (its tooltip reads \"Dirty; needs to be recompiled\" until you do), close the Blueprint editor, then click Play in the main toolbar.",
      "Save the level: File > Save Current (or Ctrl+S), name it HelloWorld. Confirm the saved name now appears in both the title bar and the Blueprint editor's tab.",
    ],
    expected: "\"Hello, World!\" appears as on-screen text in the top-left corner of the Play window.",
    trouble: "If nothing prints, confirm the white execution wire actually connects Event BeginPlay to Print String (not just placed nearby), and that you clicked Compile before pressing Play. If a \"Missing support for advanced rendering features\" dialog appears when the project first loads (about DirectX 12 / SM6 not being supported), this is expected on some hardware and safe to dismiss with OK - it only disables optional features like Nanite and Virtual Shadow Maps, which this class doesn't use.",
    imageFolder: "unreal_annotated",
    walkthroughNote: "captured on Unreal Editor 5.8",
    images: [
      ["Unreal_00_Splash.PNG", "The Unreal Engine app's own loading splash."],
      ["Unreal_01_StartUp.PNG", "The app's Home screen - click New Project."],
      ["Unreal_02_Project.PNG", "The New Project screen with its default values."],
      ["Unreal_03_ProjectDefaults.PNG", "1) Games category, 2) Blank template, 3) Project Location, 4) Project Name set to HelloWorld, 5) BLUEPRINT selected (not C++), 6) Create."],
      ["Unreal_04_EditorSplash.PNG", "The full Editor loading the new HelloWorld project."],
      ["Unreal_05_DX12Warning.PNG", "A DirectX 12/SM6 warning some hardware shows on first load - safe to click OK (see Troubleshooting)."],
      ["Unreal_06_GameEditor.PNG", "The Level Editor, open and ready."],
      ["Unreal_07_OpenLevelBlueprints.PNG", "Open Level Blueprint, found under the Blueprints toolbar menu."],
      ["Unreal_08_EvenGraph.PNG", "The Level Blueprint's Event Graph - Event BeginPlay is already there by default."],
      ["Unreal_09_AddPrintString.PNG", "Right-click and search \"print\" to find and add Print String."],
      ["Unreal_10_ExecutionLine.PNG", "The white execution wire connecting Event BeginPlay to Print String."],
      ["Unreal_11_SetStringValue.PNG", "The In String field, set to \"Hello, World!\""],
      ["Unreal_12_CompileButton.PNG", "The Compile button - its tooltip warns \"Dirty; needs to be recompiled\" until you click it."],
      ["Unreal_13_PlayButton.PNG", "The Play button, ready to test."],
      ["Unreal_14_HelloWorld.PNG", "Success - \"Hello, World!\" printed in the top-left of the Play window."],
      ["Unreal_15_SaveLevel.PNG", "Save Level As, with the name set to HelloWorld."],
      ["Unreal_16_LevelSaving.PNG", "The save-in-progress indicator."],
      ["Unreal_17_LevelNameSet.PNG", "Confirmed: HelloWorld now appears in the title bar, the Blueprint editor's own tab, and the Game window's tab in the background."],
    ],
  },
];

const numberingConfig = {
  config: [
    { reference: "bullets", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ...engines.flatMap((e, idx) => [
      { reference: `numbers-install-${idx}`, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: `numbers-hello-${idx}`, levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]),
  ]
};

const checklistHeaders = ["Engine", "Version Installed", "Hello World Confirmed?", "Date Tested", "Notes"];
const checklistWidths = [2200, 1800, 1800, 1500, 2060];
const checklistRows = engines.map(e => [e.name.replace(/^\d+\.\s*/, ""), "", "", "", ""]);

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
    headers: { default: brand.pageHeader("Engine Installation & Hello World Guide") },
    footers: { default: brand.pageFooter() },
    children: [
      brand.crestParagraph(1.3, { after: 160 }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [new TextRun({ text: "Engine Installation & Hello World Verification Guide", bold: true, size: 36, color: brand.NAVY })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 }, children: [new TextRun({ text: "Instructor reference - get every engine installed and prove each one actually runs", size: 22, color: brand.NAVY2 })] }),
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300 }, children: [new TextRun({ text: "VIDEO GAME CREATION", size: 16, bold: true, color: brand.GOLD_TEXT, characterSpacing: 20 })] }),

      h1("How to Use This Guide"),
      p("This is your personal setup reference so your demo laptop always has a working copy of every supported engine - the goal from the original course plan was to be able to demonstrate any of them on request, and to never be caught by a build that silently stopped working. For each engine below: follow the Detailed Install steps once, then run the Hello, World! Verification test to prove the full pipeline works end-to-end (editor opens, code/logic runs, output appears). Re-run the Hello World test after any engine update, OS update, or before the start of each semester."),
      p("These steps are also the backbone of Lesson 1's lab time and the per-lesson Teacher's Checklists - Level 1 for Lesson 1 is exactly this: a working install and a confirmed Hello World."),

      new TableOfContents("Table of Contents", { hyperlink: true, headingStyleRange: "1-2" }),
      new Paragraph({ children: [new PageBreak()] }),

      h1("Per-Engine Setup"),
      ...engines.flatMap((e, idx) => engineSection(e, idx)),

      h1("Instructor Verification Checklist"),
      p("Fill this in as you confirm each engine on your demo laptop. Re-check before Lesson 1 each year, and again mid-year if any engine has pushed a major update."),
      simpleTable(checklistHeaders, checklistWidths, checklistRows),
    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Engine_Installation_and_Hello_World_Guide.docx", buffer);
  console.log("done");
});
