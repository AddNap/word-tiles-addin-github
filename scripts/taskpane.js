// ====== KAFLE (pozostawiamy z placeholderami „as-is”) ======
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

// ====== THEME BRIDGE: Office theme -> CSS variables ======
function applyThemeVarsFromOffice(theme) {
  // theme { bodyBackgroundColor, bodyForegroundColor, controlBackgroundColor, controlForegroundColor, hyperlinkColor, hyperlinkVisitedColor }
  const r = document.documentElement.style
  if (!theme) return

  // Użyj kolorów Office jako bazowych, ale podbij kontrast dla akcentu
  r.setProperty('--bg', theme.bodyBackgroundColor || getComputedStyle(document.documentElement).getPropertyValue('--bg'))
  r.setProperty('--fg', theme.bodyForegroundColor || getComputedStyle(document.documentElement).getPropertyValue('--fg'))
  r.setProperty('--card', theme.controlBackgroundColor || getComputedStyle(document.documentElement).getPropertyValue('--card'))
  r.setProperty('--border', 'color-mix(in srgb, var(--fg), var(--bg) 85%)')
  r.setProperty('--accent', theme.hyperlinkColor || getComputedStyle(document.documentElement).getPropertyValue('--accent'))
  r.setProperty('--accent-contrast', '#ffffff')
}

function initThemeBridge() {
  try {
    // Zastosuj na starcie
    if (Office && Office.context && Office.context.officeTheme) {
      applyThemeVarsFromOffice(Office.context.officeTheme)
    }
    // Reaguj na zmiany w locie
    if (Office && typeof Office.onOfficeThemeChanged === 'function') {
      Office.onOfficeThemeChanged((args) => applyThemeVarsFromOffice(args))
    }
  } catch (_) {
    // cicho ignorujemy jeśli środowisko nie wspiera (np. web preview)
  }
}


// ====== Word helpers ======
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

function setTextToSelection(text) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (res) => res.status === Office.AsyncResultStatus.Succeeded
        ? resolve()
        : reject(res.error)
    )
  })
}

/**
 * Wstawia Content Control jako placeholder w miejscu kursora.
 * Tag: "jinja:<name>:<type>"
 * Tekst:
 *  - text        -> "{{ name }}"
 *  - block-start -> "{{# name }}"
 *  - block-end   -> "{{/ name }}"
 */
async function addPlaceholder(name, type) {
  await Word.run(async (context) => {
    const selection = context.document.getSelection()
    const cc = selection.insertContentControl()

    cc.tag = `jinja:${name}:${type}`
    cc.title = `jinja ${name} (${type})`
    cc.appearance = Word.ContentControlAppearance.boundingBox

    let text
    if (type === 'block-start') text = `{{# ${name} }}`
    else if (type === 'block-end') text = `{{/ ${name} }}`
    else text = `{{ ${name} }}`

    cc.insertText(text, Word.InsertLocation.replace)
    cc.cannotEdit = false
    cc.cannotDelete = false

    // Opcjonalnie: lekki szary kolor, aby było widać placeholder
    cc.color = '#888888'

    await context.sync()
  })
}

/**
 * Zwraca listę unikalnych placeholderów na podstawie tagów Content Controls.
 * Wynik: [{ name, type, count }]
 */
async function listPlaceholders() {
  return Word.run(async (context) => {
    const allCcs = context.document.contentControls
    context.load(allCcs, 'items/tag')

    await context.sync()

    const map = new Map() // key = name|type, value = {name,type,count}
    for (const cc of allCcs.items) {
      if (!cc.tag || !cc.tag.startsWith('jinja:')) continue
      const parts = cc.tag.split(':') // ["jinja", name, type]
      const name = parts[1] || ''
      const type = parts[2] || 'text'
      const key = `${name}|${type}`

      const prev = map.get(key)
      if (prev) prev.count += 1
      else map.set(key, { name, type, count: 1 })
    }

    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name))
  })
}

/** Wstawia tile HTML do bieżącej selekcji */
async function insertTile(tileId) {
  const t = TILE_DEFS.find(x => x.id === tileId)
  if (!t) return alert('Nie znaleziono kafla.')
  try {
    await setHtmlToSelection(t.html.trim())
  } catch (e) {
    console.error(e)
    alert('Nie udało się wstawić kafla (sprawdź konsolę).')
  }
}

// ====== UI: Placeholdery ======
async function refreshPlaceholderList() {
  const box = document.getElementById('placeholdersList')
  box.innerHTML = 'Ładowanie…'
  try {
    const items = await listPlaceholders()
    if (items.length === 0) {
      box.innerHTML = '<div style="opacity:.7">Brak placeholderów w dokumencie.</div>'
      return
    }
    box.innerHTML = ''
    for (const it of items) {
      const row = document.createElement('div')
      row.className = 'placeholder-item'

      const left = document.createElement('div')
      left.innerHTML = `<code>{{ ${it.name} }}</code> <span class="badge">${it.type}</span>`

      const right = document.createElement('div')
      right.style.display = 'flex'
      right.style.gap = '6px'

      // (opcjonalnie) w przyszłości: przycisk "Zaznacz" / "Usuń" / "Zmień typ"
      // Tu zostawiamy tylko licznik:
      const count = document.createElement('span')
      count.className = 'badge'
      count.textContent = `×${it.count}`
      right.appendChild(count)

      row.appendChild(left)
      row.appendChild(right)
      box.appendChild(row)
    }
  } catch (e) {
    console.error(e)
    box.innerHTML = '<div style="color:#b00020">Błąd podczas czytania placeholderów.</div>'
  }
}

function wirePlaceholderForm() {
  document.getElementById('addPlaceholder').addEventListener('click', async () => {
    const name = document.getElementById('phName').value.trim()
    const type = document.getElementById('phType').value
    if (!name) {
      alert('Podaj nazwę placeholdera.')
      return
    }
    try {
      await addPlaceholder(name, type)
      await refreshPlaceholderList()
    } catch (e) {
      console.error(e)
      alert('Nie udało się dodać placeholdera.')
    }
  })

  document.getElementById('refreshPlaceholders').addEventListener('click', refreshPlaceholderList)
}

// ====== UI: Kafle + drop-slot ======
function createTileElement(tile) {
  const el = document.createElement('div')
  el.className = 'tile'
  el.draggable = true
  el.setAttribute('data-id', tile.id)

  el.addEventListener('dragstart', (e) => {
    e.dataTransfer.setData('application/x-tile-id', tile.id)
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
  preview.innerHTML = tile.html

  const buttons = document.createElement('div')
  const addBtn = document.createElement('button')
  addBtn.className = 'btn primary'
  addBtn.textContent = 'Dodaj'
  addBtn.addEventListener('click', () => insertTile(tile.id))
  buttons.appendChild(addBtn)

  el.appendChild(title)
  el.appendChild(preview)
  el.appendChild(buttons)
  return el
}

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
  wirePlaceholderForm()
  await refreshPlaceholderList()

  const list = document.getElementById('tiles')
  TILE_DEFS.forEach(t => list.appendChild(createTileElement(t)))
}

bootstrap()
