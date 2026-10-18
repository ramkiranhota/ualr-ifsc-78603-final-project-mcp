import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, PresentationFile } from "@oai/artifact-tool";

const deckPath = "/Users/deepaksingla/Documents/Trading/TradeAnalyser/.codex_output/trading_journal_mcp_presentation_v2.pptx";
const outDir = "/Users/deepaksingla/Documents/Trading/TradeAnalyser/.codex_build/mcp_presentation/rendered_v2";
await fs.mkdir(outDir, { recursive: true });
const deck = await PresentationFile.importPptx(await FileBlob.load(deckPath));
for (let i = 0; i < deck.slides.count; i += 1) {
  const slide = deck.slides.getItem(i);
  const image = await deck.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(path.join(outDir, `slide-${i + 1}.png`), new Uint8Array(await image.arrayBuffer()));
}
console.log(`Rendered ${deck.slides.count} slides to ${outDir}`);
