// scripts/tiles.js
// Tile <-> Word Content Control mapping helpers

import { runBatch } from "./utils/batch.js";

/** tag schema example:
 * { "type":"text"|"image"|"table"|"checkbox"|"condition"|"comment",
 *   "name":"CustomerName", "required":true, "options":{...}, "skipOnExport":false }
 */

export async function insertTextTile(tile) {
  await runBatch(async (ctx) => {
    const range = ctx.document.getSelection();
    const cc = range.insertContentControl();
    cc.title = tile.name || "TextTile";
    cc.tag = JSON.stringify({ type: "text", ...tile });
    cc.style = tile.style || "Normal";
    cc.appearance = "BoundingBox";
  });
}

export async function insertCheckboxTile(tile) {
  await runBatch(async (ctx) => {
    const range = ctx.document.getSelection();
    const cc = range.insertContentControl();
    cc.title = tile.name || "CheckboxTile";
    cc.tag = JSON.stringify({ type: "checkbox", ...tile });
    cc.appearance = "BoundingBox";
    cc.insertText("[ ]", "Replace");
  });
}

export async function insertImageTile(tile) {
  await runBatch(async (ctx) => {
    const range = ctx.document.getSelection();
    const cc = range.insertContentControl();
    cc.title = tile.name || "ImageTile";
    cc.tag = JSON.stringify({ type: "image", ...tile }); // e.g. {fit,maxW,maxH,dpiMin}
    cc.appearance = "BoundingBox";
    // Optional: placeholder text
    cc.insertText("<<image:" + (tile.name || "Image") + ">>", "Replace");
  });
}

// Repeating Section for data tables (preview WordApi requirement notes)
export async function insertTableTile(tile) {
  await runBatch(async (ctx) => {
    const range = ctx.document.getSelection();
    const table = range.insertTable(1, (tile.columns || ["Col1"]).length, "Start");
    table.load("rows/items/cells/items/body");
    const sectionCC = table.parentContentControl || table.insertContentControl();
    sectionCC.title = tile.name || "TableTile";
    sectionCC.tag = JSON.stringify({ type: "table", ...tile });
  });
}

/** START/END condition anchors as hidden CCs */
export async function insertConditionTile(tile) {
  await runBatch(async (ctx) => {
    const sel = ctx.document.getSelection();
    const start = sel.insertContentControl();
    start.title = (tile.name || "Condition") + "_START";
    start.tag = JSON.stringify({ type: "condition", anchor: "START", ...tile });
    start.appearance = "Hidden";
    start.insertText("<!-- START_" + (tile.name || "COND") + " -->", "Replace");

    const endRange = start.getRange("End").expandTo(ctx.document.getSelection());
    const end = endRange.insertContentControl();
    end.title = (tile.name || "Condition") + "_END";
    end.tag = JSON.stringify({ type: "condition", anchor: "END", ...tile });
    end.appearance = "Hidden";
    end.insertText("<!-- END_" + (tile.name || "COND") + " -->", "Replace");
  });
}

export async function findTiles(predicateFn) {
  return Word.run(async (ctx) => {
    const ccs = ctx.document.contentControls;
    ccs.load("items/tag,title,id");
    await ctx.sync();
    return ccs.items.filter(cc => {
      try {
        const meta = JSON.parse(cc.tag || "{}");
        return predicateFn({ cc, meta });
      } catch { return false; }
    });
  });
}

export async function getTilesByName(name) {
  return findTiles(({ meta }) => meta.name === name);
}
