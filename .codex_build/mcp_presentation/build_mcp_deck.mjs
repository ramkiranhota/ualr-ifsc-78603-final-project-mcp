import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { Presentation, PresentationFile } from "@oai/artifact-tool";

const workspaceDir = "/Users/deepaksingla/Documents/Trading/TradeAnalyser";
const SKILL_DIR = "/Users/deepaksingla/.codex/plugins/cache/openai-primary-runtime/presentations/26.909.12148/skills/presentations";
const TMP_DIR = path.join(workspaceDir, ".codex_build", "mcp_presentation");
const FINAL_PPTX = path.join(workspaceDir, ".codex_output", "trading_journal_mcp_presentation_v2.pptx");
const RUNTIME_PYTHON = "/Users/deepaksingla/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3";

const { resolvePresentationFont, finalizePresentation } = await import(
  pathToFileURL(path.join(SKILL_DIR, "container_tools/artifact_tool_utils.mjs")).href,
);

await fs.mkdir(TMP_DIR, { recursive: true });
await fs.mkdir(path.dirname(FINAL_PPTX), { recursive: true });

const font = resolvePresentationFont({ fontFamily: "Aptos" });
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });

const C = {
  bg: "071018",
  bg2: "0B1722",
  panel: "101A24",
  panel2: "162536",
  field: "0C1721",
  line: "2A3D50",
  text: "EDF6FF",
  muted: "9AAABC",
  accent: "2DD4BF",
  accent2: "38BDF8",
  green: "4ADE80",
  warn: "FB7185",
  white: "FFFFFF",
};

function addText(slide, text, x, y, w, h, opts = {}) {
  const box = slide.shapes.add({
    geometry: "textbox",
    position: { left: x, top: y, width: w, height: h },
    fill: "none",
    line: { fill: "none", width: 0 },
  });
  box.text = text;
  box.text.style = {
    typeface: font,
    fontSize: opts.size ?? 24,
    bold: opts.bold ?? false,
    color: opts.color ?? C.text,
    autoFit: "shrinkText",
    alignment: opts.align ?? "left",
  };
  return box;
}

function addTitle(slide, title, subtitle = "") {
  addText(slide, title, 58, 42, 760, 58, { size: 36, bold: true });
  if (subtitle) {
    addText(slide, subtitle, 60, 98, 820, 38, { size: 17, color: C.muted });
  }
}

function addFooter(slide, n) {
  addText(slide, `Trading Journal MCP  ${String(n).padStart(2, "0")}`, 58, 672, 260, 22, {
    size: 11,
    color: C.muted,
  });
}

function bg(slide) {
  slide.background.fill = C.bg;
  slide.shapes.add({
    geometry: "rect",
    position: { left: 0, top: 0, width: 1280, height: 720 },
    fill: { color: C.bg },
    line: { fill: "none", width: 0 },
  });
  slide.shapes.add({
    geometry: "ellipse",
    position: { left: 0, top: 0, width: 260, height: 260 },
    fill: { color: "0F766E", transparency: 58 },
    line: { fill: "none", width: 0 },
  });
  slide.shapes.add({
    geometry: "ellipse",
    position: { left: 1020, top: 0, width: 260, height: 260 },
    fill: { color: "0EA5E9", transparency: 70 },
    line: { fill: "none", width: 0 },
  });
}

function node(slide, label, title, x, y, w, h, opts = {}) {
  const s = slide.shapes.add({
    geometry: opts.shape ?? "roundRect",
    position: { left: x, top: y, width: w, height: h },
    fill: { color: opts.fill ?? C.panel },
    line: { style: "solid", fill: opts.line ?? C.line, width: opts.lineWidth ?? 1.2 },
  });
  addText(slide, label, x + 16, y + 12, w - 32, 24, {
    size: opts.labelSize ?? 11,
    bold: true,
    color: opts.labelColor ?? C.accent,
    align: opts.align ?? "center",
  });
  addText(slide, title, x + 18, y + 42, w - 36, h - 48, {
    size: opts.titleSize ?? 20,
    bold: true,
    color: opts.color ?? C.text,
    align: opts.align ?? "center",
  });
  return s;
}

function smallNote(slide, text, x, y, w, h) {
  return addText(slide, text, x, y, w, h, { size: 14, color: C.muted, align: "center" });
}

function connect(slide, a, b, fromSide = "right", toSide = "left", dashed = false) {
  return slide.shapes.connect(a, b, {
    kind: "straight",
    fromSide,
    toSide,
    line: { style: dashed ? "dashed" : "solid", fill: C.accent, width: 2.4 },
    head: { type: "arrow", width: "med", length: "med" },
  });
}

function notes(slide, text) {
  slide.speakerNotes.textFrame.setText(text);
}

// Slide 1
{
  const slide = deck.slides.add();
  bg(slide);
  addText(slide, "Model Context Protocol", 72, 92, 760, 78, { size: 46, bold: true });
  addText(slide, "General concept and Trading Journal use case", 74, 174, 620, 36, { size: 21, color: C.muted });
  const n1 = node(slide, "CLIENT", "AI app or UI", 110, 326, 230, 112, { fill: C.panel2, line: C.accent });
  const n2 = node(slide, "PROTOCOL", "MCP", 525, 304, 230, 156, { fill: "0E2A35", line: C.accent2, titleSize: 34 });
  const n3 = node(slide, "SERVER", "Tools and context", 940, 326, 230, 112, { fill: C.panel2, line: C.accent });
  connect(slide, n1, n2);
  connect(slide, n2, n3);
  addText(slide, "A common way for clients to discover and use external capabilities", 285, 506, 720, 34, {
    size: 22,
    color: C.text,
    align: "center",
  });
  addFooter(slide, 1);
  notes(slide, "Sources: Model Context Protocol architecture documentation, https://modelcontextprotocol.io/docs/concepts/architecture. Trading Journal project documentation in docs/architecture_plan.md.");
}

// Slide 2
{
  const slide = deck.slides.add();
  bg(slide);
  addTitle(slide, "MCP building blocks", "The protocol separates user experience from available capabilities");
  const host = node(slide, "HOST", "ChatGPT, IDE, or custom app", 70, 255, 250, 120, { fill: C.panel2, line: C.accent2 });
  const client = node(slide, "CLIENT", "One connection per server", 394, 255, 250, 120, { fill: C.panel, line: C.accent });
  const server = node(slide, "SERVER", "Exposes tools and context", 718, 255, 250, 120, { fill: C.panel, line: C.green });
  const data = node(slide, "DATA", "APIs, files, databases", 1038, 255, 170, 120, { fill: C.field, line: C.line, titleSize: 17 });
  connect(slide, host, client);
  connect(slide, client, server);
  connect(slide, server, data);
  node(slide, "TOOLS", "Actions", 330, 464, 170, 76, { fill: "0E2A35", line: C.accent, titleSize: 18 });
  node(slide, "RESOURCES", "Readable context", 555, 464, 210, 76, { fill: "0E2A35", line: C.accent, titleSize: 18 });
  node(slide, "PROMPTS", "Reusable workflows", 820, 464, 210, 76, { fill: "0E2A35", line: C.accent, titleSize: 18 });
  addFooter(slide, 2);
  notes(slide, "Sources: MCP architecture and concepts documentation: https://modelcontextprotocol.io/docs/concepts/architecture and specification sections for tools, resources and prompts at https://modelcontextprotocol.io/specification.");
}

// Slide 3
{
  const slide = deck.slides.add();
  bg(slide);
  addTitle(slide, "Client to server lifecycle", "Discovery happens before action");
  const a = node(slide, "1", "Connect", 95, 265, 180, 96, { fill: C.panel2, line: C.accent2, titleSize: 24 });
  const b = node(slide, "2", "Discover tools", 345, 265, 190, 96, { fill: C.panel2, line: C.accent, titleSize: 24 });
  const c = node(slide, "3", "Call a tool", 605, 265, 190, 96, { fill: C.panel2, line: C.green, titleSize: 24 });
  const d = node(slide, "4", "Return result", 865, 265, 190, 96, { fill: C.panel2, line: C.accent2, titleSize: 24 });
  connect(slide, a, b);
  connect(slide, b, c);
  connect(slide, c, d);
  addText(slide, "The client does not need to know how the server gets the data", 230, 450, 820, 38, {
    size: 26,
    bold: true,
    align: "center",
  });
  addText(slide, "It only needs the server contract: available tools, expected inputs, and structured outputs", 250, 500, 780, 52, {
    size: 18,
    color: C.muted,
    align: "center",
  });
  addFooter(slide, 3);
  notes(slide, "Sources: MCP architecture documentation on client-server interaction and capability discovery, https://modelcontextprotocol.io/docs/concepts/architecture.");
}

// Slide 4
{
  const slide = deck.slides.add();
  bg(slide);
  addTitle(slide, "Why MCP matters", "One capability layer can support many clients");
  const center = node(slide, "MCP SERVER", "Stable capability contract", 480, 270, 320, 142, {
    fill: "0E2A35",
    line: C.accent,
    titleSize: 26,
  });
  const app = node(slide, "CLIENT", "Web app", 95, 164, 200, 82, { fill: C.panel2, line: C.accent2 });
  const cli = node(slide, "CLIENT", "CLI demo", 95, 450, 200, 82, { fill: C.panel2, line: C.accent2 });
  const ai = node(slide, "CLIENT", "Future AI agent", 990, 164, 200, 82, { fill: C.panel2, line: C.accent2 });
  const ide = node(slide, "CLIENT", "Other apps", 990, 450, 200, 82, { fill: C.panel2, line: C.accent2 });
  connect(slide, app, center, "right", "left");
  connect(slide, cli, center, "right", "left");
  connect(slide, ai, center, "left", "right");
  connect(slide, ide, center, "left", "right");
  addText(slide, "MCP keeps client experience flexible while the server owns access, policy, and integration details", 240, 582, 800, 48, {
    size: 22,
    bold: true,
    align: "center",
  });
  addFooter(slide, 4);
  notes(slide, "Sources: MCP architecture documentation: https://modelcontextprotocol.io/docs/concepts/architecture. MCP 2026 roadmap for ecosystem direction: https://blog.modelcontextprotocol.io/posts/2026-mcp-roadmap/.");
}

// Slide 5
{
  const slide = deck.slides.add();
  bg(slide);
  addTitle(slide, "Trading Journal MCP topology", "The app exposes finance workflows through domain MCP servers");
  const clients = node(slide, "CLIENTS", "Web UI\nMCP Console\nCLI Client", 70, 252, 210, 170, {
    fill: C.panel2,
    line: C.accent2,
    titleSize: 23,
  });
  const host = node(slide, "APP HOST", "FastAPI app\nInternal MCP client", 350, 258, 230, 158, {
    fill: C.panel,
    line: C.accent,
    titleSize: 22,
  });
  const mcp = node(slide, "MCP SERVERS", "Portfolio\nMarket Data\nNews\nBroker\nTrading", 655, 205, 230, 264, {
    fill: "0E2A35",
    line: C.accent,
    titleSize: 22,
  });
  const services = node(slide, "SERVICES", "Portfolio, market data,\nnews, broker, orders", 970, 160, 230, 104, {
    fill: C.panel2,
    line: C.green,
    titleSize: 19,
  });
  const storage = node(slide, "STATE", "SQLite database\nAudit logs", 970, 308, 230, 104, {
    fill: C.field,
    line: C.line,
    titleSize: 20,
  });
  const providers = node(slide, "PROVIDERS", "Alpaca market data\nYahoo news", 970, 456, 230, 104, {
    fill: C.field,
    line: C.line,
    titleSize: 20,
  });
  connect(slide, clients, host);
  connect(slide, host, mcp);
  connect(slide, mcp, services);
  connect(slide, mcp, storage);
  connect(slide, mcp, providers);
  addFooter(slide, 5);
  notes(slide, "Sources: Trading Journal architecture documentation docs/architecture_plan.md and implementation plan docs/implementation_plan.md. External integrations are described in the project as Alpaca market data and Yahoo Finance demo news.");
}

// Slide 6
{
  const slide = deck.slides.add();
  bg(slide);
  addTitle(slide, "Two client paths", "The same MCP servers support the browser and the terminal");
  const leftTop = node(slide, "WEB", "User clicks Market Data", 85, 188, 230, 86, { fill: C.panel2, line: C.accent2 });
  const leftMid = node(slide, "INTERNAL", "FastAPI MCP client", 390, 188, 230, 86, { fill: C.panel, line: C.accent });
  const rightTop = node(slide, "CLI", "mcp_demo_client.py", 85, 434, 230, 86, { fill: C.panel2, line: C.accent2 });
  const rightMid = node(slide, "EXTERNAL", "Streamable HTTP client", 390, 434, 230, 86, { fill: C.panel, line: C.accent });
  const server = node(slide, "SERVER", "Market Data MCP", 735, 310, 230, 110, { fill: "0E2A35", line: C.green, titleSize: 24 });
  const api = node(slide, "PROVIDER", "Alpaca snapshots", 1030, 310, 170, 110, { fill: C.field, line: C.line, titleSize: 18 });
  connect(slide, leftTop, leftMid);
  connect(slide, rightTop, rightMid);
  connect(slide, leftMid, server);
  connect(slide, rightMid, server);
  connect(slide, server, api);
  addText(slide, "The UI and CLI use different experiences, but the capability contract stays the same", 270, 594, 740, 42, {
    size: 22,
    bold: true,
    align: "center",
  });
  addFooter(slide, 6);
  notes(slide, "Sources: app/templates/workflow.html, scripts/mcp_demo_client.py, app/services/mcp_host.py, and app/services/market_data.py in the Trading Journal project.");
}

// Slide 7
{
  const slide = deck.slides.add();
  bg(slide);
  addTitle(slide, "Trading Journal capability map", "Each financial domain has a focused server boundary");
  const items = [
    ["Portfolio", "Accounts, trades, positions"],
    ["Market Data", "Live prices through Alpaca"],
    ["News", "Stock news through Yahoo"],
    ["Broker", "IBKR sync placeholder"],
    ["Trading", "Order preview only"],
    ["Audit", "Trace user and MCP activity"],
  ];
  const startX = 78;
  const startY = 178;
  const w = 340;
  const h = 96;
  items.forEach((it, idx) => {
    const col = idx % 3;
    const row = Math.floor(idx / 3);
    const x = startX + col * 400;
    const y = startY + row * 170;
    node(slide, it[0], it[1], x, y, w, h, {
      fill: idx === 4 ? "281923" : "0E2A35",
      line: idx === 4 ? C.warn : C.accent,
      titleSize: 22,
    });
  });
  addText(slide, "Safety boundary: broker and trading workflows stay read-only or preview-only in the prototype", 210, 566, 860, 44, {
    size: 22,
    bold: true,
    color: C.warn,
    align: "center",
  });
  addFooter(slide, 7);
  notes(slide, "Sources: docs/architecture_plan.md section Active MCP Endpoints and Current Safety Boundary. The project explicitly keeps live trading disabled and Broker MCP as a placeholder.");
}

// Slide 8
{
  const slide = deck.slides.add();
  bg(slide);
  addTitle(slide, "Reference material", "Sources behind the MCP overview and application workflow");
  const refs = [
    ["MCP architecture", "Hosts, clients, servers and capability discovery"],
    ["MCP specification", "Tools, resources, prompts and transport concepts"],
    ["MCP roadmap", "Ecosystem direction and future work"],
    ["Project docs", "Trading Journal architecture and implementation plan"],
  ];
  refs.forEach((r, i) => {
    node(slide, r[0], r[1], 115 + (i % 2) * 540, 190 + Math.floor(i / 2) * 170, 430, 100, {
      fill: i < 3 ? C.panel2 : "0E2A35",
      line: i < 3 ? C.accent2 : C.accent,
      titleSize: 20,
    });
  });
  addText(slide, "Suggested talk track: define MCP first, then show how the app turns portfolio, market data and news into MCP capabilities", 140, 570, 1000, 56, {
    size: 22,
    bold: true,
    align: "center",
  });
  addFooter(slide, 8);
  notes(slide, "References: MCP architecture documentation https://modelcontextprotocol.io/docs/concepts/architecture. MCP specification https://modelcontextprotocol.io/specification. MCP roadmap https://blog.modelcontextprotocol.io/posts/2026-mcp-roadmap/. Trading Journal local files docs/architecture_plan.md, docs/implementation_plan.md and app/templates/workflow.html.");
}

const candidatePath = path.join(TMP_DIR, "candidate.pptx");
await (await PresentationFile.exportPptx(deck)).save(candidatePath);

for (let i = 0; i < deck.slides.length; i += 1) {
  const slide = deck.slides.get(i);
  const preview = await deck.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(TMP_DIR, `slide-${i + 1}.png`), new Uint8Array(await preview.arrayBuffer()));
}

const result = await finalizePresentation({
  workspaceDir,
  candidatePath,
  finalPath: FINAL_PPTX,
  pythonExecutable: RUNTIME_PYTHON,
  integrityValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_package_integrity.py"),
  layoutValidatorPath: path.join(SKILL_DIR, "container_tools/inspect_presentation_layout_geometry.py"),
  layoutArgs: [
    "--expected-slide-size-emu",
    "12192000,6858000",
    "--validate-heading-fit",
  ],
  requiredNativeTableOwnerSlides: [],
  fontPolicy: { basis: "design", families: [font] },
  verifyArtifactToolImport: true,
  receiptPath: path.join(TMP_DIR, "trading_journal_mcp_presentation_v2.validation.json"),
});

console.log(JSON.stringify({ final: FINAL_PPTX, result }, null, 2));
