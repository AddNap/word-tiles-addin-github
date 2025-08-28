// scripts/inspector.js
import { findTiles } from "./tiles.js";

export async function validateTiles() {
  const errors = [];
  const list = await findTiles(() => true);
  const names = new Map();
  const starts = new Set();
  const ends = new Set();

  for (const { cc, meta } of list) {
    if (!meta.type) errors.push({ id: cc.id, msg: "Brak typu w tagu JSON." });
    if (meta.required && (!meta.name || meta.name.trim() === "")) {
      errors.push({ id: cc.id, msg: "Pole wymagane bez nazwy." });
    }
    if (meta.name) {
      const count = (names.get(meta.name) || 0) + 1;
      names.set(meta.name, count);
    }
    if (meta.type === "condition") {
      if (meta.anchor === "START") starts.add(meta.name || "COND");
      if (meta.anchor === "END") ends.add(meta.name || "COND");
    }
  }
  // duplikaty
  for (const [n, c] of names.entries()) {
    if (c > 1) errors.push({ id: "-", msg: `Zduplikowana nazwa pola: ${n}` });
  }
  // pary warunków
  for (const n of starts) if (!ends.has(n)) errors.push({ id: "-", msg: `Brak END dla warunku: ${n}` });
  for (const n of ends) if (!starts.has(n)) errors.push({ id: "-", msg: `Brak START dla warunku: ${n}` });

  renderInspector(errors);
  return errors;
}

function renderInspector(errors) {
  const panel = document.getElementById("tiles-inspector");
  if (!panel) return;
  panel.innerHTML = "";
  const header = document.createElement("div");
  header.className = "ti-header";
  header.textContent = `Tiles Inspector — błędy: ${errors.length}`;
  panel.appendChild(header);

  if (errors.length === 0) {
    const ok = document.createElement("div");
    ok.className = "ti-ok";
    ok.textContent = "Brak problemów ✅";
    panel.appendChild(ok);
    return;
  }
  for (const e of errors) {
    const row = document.createElement("div");
    row.className = "ti-row";
    row.textContent = e.msg;
    panel.appendChild(row);
  }
}

window.addEventListener("load", () => {
  const btn = document.getElementById("btn-validate-tiles");
  if (btn) btn.addEventListener("click", () => validateTiles());
});
