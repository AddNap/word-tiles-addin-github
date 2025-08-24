/* global Office, Word */
Office.onReady(() => {
  document.getElementById("addTileBtn").addEventListener("click", async () => {
    const name = document.getElementById("tileName").value.trim();
    if (!name) { alert("Podaj nazwę placeholdera."); return; }
    try {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        const cc = range.insertContentControl();
        cc.title = `TILE:${name}`;
        cc.tag = `TILE:${name}`;
        cc.appearance = "BoundingBox";
        cc.insertText(`{{${name}}}`, Word.InsertLocation.replace);
        await context.sync();
      });
    } catch (err) { console.error(err); alert("Błąd: " + err); }
  });

  document.getElementById("addBlockBtn").addEventListener("click", async () => {
    const name = prompt("Podaj nazwę bloku:"); if (!name) return;
    try {
      await Word.run(async (context) => {
        const sel = context.document.getSelection();
        const startCC = sel.insertContentControl();
        startCC.title = `BLOCK_START:${name}`;
        startCC.tag = `BLOCK_START:${name}`;
        startCC.appearance = "BoundingBox";
        startCC.insertText(`{{ START_${name} }}`, Word.InsertLocation.replace);
        const after = startCC.insertParagraph("", Word.InsertLocation.after);
        after.select();
        const endCC = after.insertContentControl();
        endCC.title = `BLOCK_END:${name}`;
        endCC.tag = `BLOCK_END:${name}`;
        endCC.appearance = "BoundingBox";
        endCC.insertText(`{{ END_${name} }}`, Word.InsertLocation.replace);
        await context.sync();
      });
    } catch (err) { console.error(err); alert("Błąd: " + err); }
  });
});
