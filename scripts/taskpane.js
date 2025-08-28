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

// ====== Word helpers ======
function setH
