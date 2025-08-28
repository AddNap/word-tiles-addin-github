// --- Dane przykładowych kafli (z placeholderami jinja2) ---
const TILE_DEFS = [
  {
    id: 'address-card',
    title: 'Address card',
    html: `
      <div style="border:1px solid #ddd;border-radius:8px;padding:10px;">
        <div style="font-weight:700;">{{ name }}</div>
        <div>{{ street }}</div>
        <div>{{ postal }} {{ city }}</div>
      </div>`
  },
  {
    id: 'quote',
    title: 'Quote box',
    html: `
      <blockquote style="border-left:4px solid #999;padding-left:10px;margin:0;">
        “{{ quote }}”
        <div style="font-size:12px;opacity:.7;">— {{ author }}</div>
      </blockquote>`
  },
  {
    id: 'table',
    title: 'Simple table',
    html: `
      <table style="border-collapse:collapse;width:100%;">
        <tr><th style="border:1px solid #ccc;padding:6px;">{{ col1 }}</th>
            <th style="border:1px solid #ccc;padding:6px;">{{ col2 }}</th></tr>
        <tr><td style="border:1px solid #ccc;padding:6px;">{{ v11 }}</td>
            <td style="border:1px solid #ccc;padding:6px;">{{ v12 }}</td></tr>
        <tr><td style="border:1px solid #ccc;padding:6px;">{{ v21 }}</td>
            <td style="border:1px solid #ccc;padding:6px;">{{ v22 }}</td></tr>
      </table>`
  }
]

// Opcjonalny „renderer” — na Pages tylko echo; później zamienisz na wywołanie backendu
async function renderTileHTML(tileId) {
  const t = TILE_DEFS.find(x => x.id === tileId)
  if (!t) throw new Error('Tile not found')
  return t.html.trim()
}

// --- Word helpers ---
function setHtmlToSelection(html) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (res) => res.status === Office.AsyncResultStatus.Succeeded
        ? resolve()
        : reject(res.error)
    )
  })
}

// --- UI ---
function createTileElement(tile) {
  const el = document.createElement('div')
  el.className = 'tile'
  el.draggable = true
  el.setAttribute('data-id', tile.id)

  el.addEventListener('dragstart', (e) => {
    e.dataTransfer.setData('application/x-tile-id', tile.id)
    // prosty ghost
    const ghost = document.createElement('div')
    ghost.style.padding = '6px'
    ghost.style.background = 'white'
    ghost.style.border = '1px solid #ddd'
    ghost.textContent = tile.title
    document.body.appendChild(ghost)
    e.dataTransfer.setDragImage(ghost, 0, 0)
    setTimeout(() => ghost.remove(), 0)
  })

  const title = document.createElement('div')
  title.className = 'tile-title'
  title.textContent = tile.title

  const preview = document.createElement('div')
  preview.innerHTML = tile.html // pokazujemy placeholdery

  const footer = document.createElement('div')
  footer.className = 'tile-footer'
  footer.textContent = 'Double-click, Enter lub upuść w drop-slot, aby wstawić'

  el.appendChild(title)
  el.appendChild(preview)
  el.appendChild(footer)

  // Szybkie wstawienie: dblclick / Enter
  el.addEventListener('dblclick', () => insertTile(tile.id))
  el.tabIndex = 0
  el.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') insertTile(tile.id)
  })

  return el
}

async function insertTile(tileId) {
  try {
    const html = await renderTileHTML(tileId)
    await setHtmlToSelection(html)
  } catch (err) {
    console.error(err)
    alert('Failed to insert tile. See console for details.')
  }
}

// Drop-slot: wstawia do bieżącej selekcji
function setupDropSlot() {
  const slot = document.getElementById('drop-slot')
  slot.addEventListener('dragover', (e) => {
    e.preventDefault()
    slot.classList.add('dragover')
  })
  slot.addEventListener('dragleave', () => {
    slot.classList.remove('dragover')
  })
  slot.addEventListener('drop', async (e) => {
    e.preventDefault()
    slot.classList.remove('dragover')
    const id = e.dataTransfer.getData('application/x-tile-id')
    if (id) await insertTile(id)
  })
}

async function bootstrap() {
  await Office.onReady()
  setupDropSlot()

  const list = document.getElementById('tiles')
  TILE_DEFS.forEach(t => list.appendChild(createTileElement(t)))
}

bootstrap()
