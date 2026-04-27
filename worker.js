const SPREADSHEET_EXTENSIONS = [".xlsx", ".xls", ".xlsm", ".csv"];

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "\"": "&quot;",
    "'": "&#39;",
  })[char]);
}

function extensionFor(key) {
  const lowerKey = key.toLowerCase();
  return SPREADSHEET_EXTENSIONS.find((extension) => lowerKey.endsWith(extension));
}

function contentTypeFor(key) {
  const lowerKey = key.toLowerCase();

  if (lowerKey.endsWith(".html")) return "text/html;charset=UTF-8";
  if (lowerKey.endsWith(".md")) return "text/markdown;charset=UTF-8";
  if (lowerKey.endsWith(".pdf")) return "application/pdf";
  if (lowerKey.endsWith(".xlsx")) return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  if (lowerKey.endsWith(".xls")) return "application/vnd.ms-excel";
  if (lowerKey.endsWith(".xlsm")) return "application/vnd.ms-excel.sheet.macroEnabled.12";
  if (lowerKey.endsWith(".csv")) return "text/csv;charset=UTF-8";

  return "application/octet-stream";
}

function spreadsheetViewer(path) {
  const fileName = escapeHtml(decodeURIComponent(path.split("/").pop() || "spreadsheet"));

  return `<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${fileName}</title>
  <style>
    :root { color-scheme: light; font-family: Inter, "Segoe UI", Arial, sans-serif; }
    body { margin: 0; color: #172033; background: #f7f8fb; }
    header { position: sticky; top: 0; z-index: 10; display: flex; align-items: center; gap: 12px; padding: 12px 16px; border-bottom: 1px solid #d8deea; background: rgba(255,255,255,.96); }
    h1 { margin: 0; min-width: 0; flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-size: 16px; font-weight: 650; }
    a.button { flex: 0 0 auto; color: #fff; background: #1769e0; border-radius: 6px; padding: 8px 11px; font-size: 13px; font-weight: 600; text-decoration: none; }
    #tabs { display: flex; gap: 6px; overflow-x: auto; padding: 10px 16px; border-bottom: 1px solid #d8deea; background: #fff; }
    #tabs button { flex: 0 0 auto; border: 1px solid #c8d1e1; border-radius: 6px; background: #fff; color: #263143; padding: 7px 10px; font-size: 13px; cursor: pointer; }
    #tabs button.active { border-color: #1769e0; background: #eaf2ff; color: #0f55bb; }
    #status { padding: 18px 16px; color: #5b6576; }
    #sheet { margin: 16px; overflow: auto; border: 1px solid #d8deea; border-radius: 8px; background: #fff; box-shadow: 0 1px 2px rgba(12, 20, 33, .05); }
    table { border-collapse: collapse; min-width: 100%; font-size: 13px; background: #fff; }
    th, td { max-width: 420px; min-width: 60px; border: 1px solid #e1e6ef; padding: 7px 9px; vertical-align: top; white-space: pre-wrap; word-break: break-word; }
    th { position: sticky; top: 0; z-index: 1; background: #f0f4fa; font-weight: 650; }
    .empty { color: transparent; }
  </style>
</head>
<body>
  <header>
    <h1>${fileName}</h1>
    <a id="download" class="button" href="?download=1">下载</a>
  </header>
  <nav id="tabs"></nav>
  <main>
    <div id="status">正在加载 Excel 文件...</div>
    <div id="sheet" hidden></div>
  </main>
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <script>
    const statusEl = document.getElementById('status');
    const sheetEl = document.getElementById('sheet');
    const tabsEl = document.getElementById('tabs');
    document.getElementById('download').href = window.location.pathname + '?download=1';

    function setStatus(message) {
      statusEl.textContent = message;
      statusEl.hidden = false;
      sheetEl.hidden = true;
    }

    function normalizeRgb(value) {
      if (!value || typeof value !== 'string') return '';
      const rgb = value.replace(/^#/, '').trim();
      if (/^[0-9a-fA-F]{8}$/.test(rgb)) return '#' + rgb.slice(2);
      if (/^[0-9a-fA-F]{6}$/.test(rgb)) return '#' + rgb;
      return '';
    }

    function colorFrom(styleColor) {
      if (!styleColor) return '';
      if (styleColor.rgb) return normalizeRgb(styleColor.rgb);
      if (styleColor.indexed === 64) return '';
      return '';
    }

    function fillColorFrom(style) {
      if (!style || !style.fill) return '';
      return colorFrom(style.fill.fgColor) || colorFrom(style.fill.bgColor);
    }

    function applyAlignment(element, alignment) {
      if (!alignment) return;

      const horizontalMap = {
        left: 'left',
        center: 'center',
        right: 'right',
        justify: 'justify',
        distributed: 'justify',
      };
      const verticalMap = {
        top: 'top',
        center: 'middle',
        bottom: 'bottom',
      };

      if (horizontalMap[alignment.horizontal]) {
        element.style.textAlign = horizontalMap[alignment.horizontal];
      }

      if (verticalMap[alignment.vertical]) {
        element.style.verticalAlign = verticalMap[alignment.vertical];
      }
    }

    function applyCellStyle(element, cell) {
      const style = cell && cell.s ? cell.s : {};
      const fillColor = fillColorFrom(style);
      const fontColor = colorFrom(style.font && style.font.color);

      if (fillColor) element.style.backgroundColor = fillColor;
      if (fontColor) element.style.color = fontColor;
      if (style.font && style.font.bold) element.style.fontWeight = '700';
      applyAlignment(element, style.alignment);
    }

    function formattedValue(cell) {
      if (!cell) return '';
      if (cell.w != null) return String(cell.w);
      if (cell.v == null) return '';
      return String(cell.v);
    }

    function buildMergeMap(sheet) {
      const starts = new Map();
      const covered = new Set();

      for (const merge of sheet['!merges'] || []) {
        const rowSpan = merge.e.r - merge.s.r + 1;
        const colSpan = merge.e.c - merge.s.c + 1;
        starts.set(merge.s.r + ':' + merge.s.c, { rowSpan, colSpan });

        for (let row = merge.s.r; row <= merge.e.r; row += 1) {
          for (let col = merge.s.c; col <= merge.e.c; col += 1) {
            if (row !== merge.s.r || col !== merge.s.c) covered.add(row + ':' + col);
          }
        }
      }

      return { starts, covered };
    }

    function renderSheet(workbook, sheetName) {
      for (const button of tabsEl.querySelectorAll('button')) {
        button.classList.toggle('active', button.dataset.sheet === sheetName);
      }

      const sheet = workbook.Sheets[sheetName];
      if (!sheet || !sheet['!ref']) {
        setStatus('这个工作表没有可显示的数据。');
        return;
      }

      const range = XLSX.utils.decode_range(sheet['!ref']);
      const mergeMap = buildMergeMap(sheet);
      const table = document.createElement('table');
      const fragment = document.createDocumentFragment();
      let hasVisibleValue = false;

      for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
        const tr = document.createElement('tr');

        for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
          const key = rowIndex + ':' + columnIndex;
          if (mergeMap.covered.has(key)) continue;

          const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
          const workbookCell = sheet[address];
          const cell = document.createElement(rowIndex === range.s.r ? 'th' : 'td');
          const text = formattedValue(workbookCell);
          const merge = mergeMap.starts.get(key);

          cell.textContent = text;
          if (!text) cell.classList.add('empty');

          if (merge) {
            if (merge.rowSpan > 1) cell.rowSpan = merge.rowSpan;
            if (merge.colSpan > 1) cell.colSpan = merge.colSpan;
          }

          applyCellStyle(cell, workbookCell);
          if (text) hasVisibleValue = true;
          tr.appendChild(cell);
        }

        fragment.appendChild(tr);
      }

      if (!hasVisibleValue) {
        setStatus('这个工作表没有可显示的数据。');
        return;
      }

      table.appendChild(fragment);
      sheetEl.replaceChildren(table);
      statusEl.hidden = true;
      sheetEl.hidden = false;
    }

    async function loadWorkbook() {
      if (!window.XLSX) {
        setStatus('Excel 查看器加载失败，请检查浏览器是否可以访问 jsDelivr CDN。');
        return;
      }

      const response = await fetch(window.location.pathname + '?raw=1', { credentials: 'same-origin' });
      if (!response.ok) throw new Error('HTTP ' + response.status);
      const data = await response.arrayBuffer();
      const workbook = XLSX.read(data, { type: 'array', cellDates: true, cellStyles: true });
      if (!workbook.SheetNames.length) {
        setStatus('这个 Excel 文件没有工作表。');
        return;
      }

      tabsEl.replaceChildren(...workbook.SheetNames.map((sheetName) => {
        const button = document.createElement('button');
        button.type = 'button';
        button.dataset.sheet = sheetName;
        button.textContent = sheetName;
        button.addEventListener('click', () => renderSheet(workbook, sheetName));
        return button;
      }));
      renderSheet(workbook, workbook.SheetNames[0]);
    }

    loadWorkbook().catch((error) => {
      setStatus('Excel 文件加载失败：' + error.message);
    });
  </script>
</body>
</html>`;
}

var worker_default = {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);
    let path = url.pathname;

    // 根路径默认访问 index.html
    if (path === "/" || path === "") {
      path = "/index.html";
    }

    // 去掉开头的 "/" 并加上 APS/ 前缀
    const key = "APS/" + (path.startsWith("/") ? path.slice(1) : path);
    const spreadsheetExtension = extensionFor(key);

    try {
      if (spreadsheetExtension && !url.searchParams.has("raw") && !url.searchParams.has("download")) {
        const objectHead = await env.MY_BUCKET.head(key);
        if (!objectHead) {
          return new Response("404 Not Found", { status: 404 });
        }

        return new Response(spreadsheetViewer(path), {
          headers: {
            "content-type": "text/html;charset=UTF-8",
          },
        });
      }

      const object = await env.MY_BUCKET.get(key);
      if (!object) {
        return new Response("404 Not Found", { status: 404 });
      }

      let extraHeaders = {};

      if (key.toLowerCase().endsWith(".pdf")) {
        extraHeaders["Content-Disposition"] = "inline";
      }

      if (spreadsheetExtension && url.searchParams.has("download")) {
        extraHeaders["Content-Disposition"] = `attachment; filename="${path.split("/").pop() || "spreadsheet"}"`;
      }

      return new Response(object.body, {
        headers: {
          "content-type": contentTypeFor(key),
          ...extraHeaders,
        },
      });
    } catch (err) {
      return new Response("Error: " + err.message, { status: 500 });
    }
  },
};

export {
  worker_default as default,
};
