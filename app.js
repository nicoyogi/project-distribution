/* Invoice Project Distribution - offline splitter
 *
 * Loads a single Excel workbook (any number of sheets, 4 expected),
 * then splits each sheet's data rows as evenly as possible across N
 * output workbooks. Each output workbook keeps the same sheet names
 * and (optionally) header row.
 */

(function () {
  'use strict';

  const fileInput = document.getElementById('fileInput');
  const dropZone = document.getElementById('dropZone');
  const dropLabel = document.getElementById('dropLabel');
  const fileInfo = document.getElementById('fileInfo');
  const splitCount = document.getElementById('splitCount');
  const hasHeader = document.getElementById('hasHeader');
  const splitBtn = document.getElementById('splitBtn');
  const statusEl = document.getElementById('status');
  const resultsEl = document.getElementById('results');

  /** @type {{ name: string, workbook: any } | null} */
  let loaded = null;

  // ---------- file selection ----------
  fileInput.addEventListener('change', (e) => {
    const f = e.target.files && e.target.files[0];
    if (f) handleFile(f);
  });

  ['dragenter', 'dragover'].forEach((ev) =>
    dropZone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropZone.classList.add('dragover');
    })
  );
  ['dragleave', 'drop'].forEach((ev) =>
    dropZone.addEventListener(ev, (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
    })
  );
  dropZone.addEventListener('drop', (e) => {
    const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });

  function handleFile(file) {
    setStatus('Reading file...', '');
    resultsEl.innerHTML = '';
    splitBtn.disabled = true;

    const reader = new FileReader();
    reader.onerror = () => setStatus('Could not read file.', 'error');
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        loaded = { name: file.name, workbook: wb };
        renderFileInfo(file, wb);
        splitBtn.disabled = false;
        setStatus('Ready to split.', 'ok');
      } catch (err) {
        console.error(err);
        loaded = null;
        setStatus('Failed to parse workbook: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function renderFileInfo(file, wb) {
    dropLabel.textContent = file.name;
    const rows = wb.SheetNames.map((name) => {
      const sheet = wb.Sheets[name];
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
      const total = aoa.length;
      const data = hasHeader.checked ? Math.max(0, total - 1) : total;
      return `<tr><td>${escapeHtml(name)}</td><td>${total}</td><td>${data}</td></tr>`;
    }).join('');

    fileInfo.classList.remove('hidden');
    fileInfo.innerHTML = `
      <div><strong>${escapeHtml(file.name)}</strong> — ${wb.SheetNames.length} sheet(s)</div>
      <table>
        <thead><tr><th>Sheet</th><th>Total rows</th><th>Data rows</th></tr></thead>
        <tbody>${rows}</tbody>
      </table>
    `;
  }

  // recompute "data rows" when header toggle changes
  hasHeader.addEventListener('change', () => {
    if (loaded) renderFileInfo({ name: loaded.name }, loaded.workbook);
  });

  // ---------- split action ----------
  splitBtn.addEventListener('click', () => {
    if (!loaded) return;
    const n = parseInt(splitCount.value, 10);
    if (!Number.isFinite(n) || n < 2) {
      setStatus('Choose at least 2 output files.', 'error');
      return;
    }
    try {
      const files = splitWorkbook(loaded.workbook, n, hasHeader.checked);
      const baseName = stripExt(loaded.name);
      renderResults(files, baseName);
      setStatus(`Created ${files.length} files.`, 'ok');
    } catch (err) {
      console.error(err);
      setStatus('Split failed: ' + err.message, 'error');
    }
  });

  /**
   * Split a workbook into n workbooks. For each sheet, the data rows are
   * partitioned as evenly as possible (extra rows go to the first chunks).
   *
   * @param {any} wb source workbook
   * @param {number} n number of output workbooks
   * @param {boolean} headerRow whether row 1 of each sheet is a header
   * @returns {{ name: string, blob: Blob }[]}
   */
  function splitWorkbook(wb, n, headerRow) {
    // Pre-extract every sheet as array-of-arrays once.
    const sheets = wb.SheetNames.map((name) => {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], {
        header: 1,
        defval: null,
      });
      let header = null;
      let data = aoa;
      if (headerRow && aoa.length > 0) {
        header = aoa[0];
        data = aoa.slice(1);
      }
      return { name, header, data };
    });

    // Compute partitions per sheet: array of [start, end) ranges, length n.
    const partitions = sheets.map((s) => evenSplit(s.data.length, n));

    const outputs = [];
    for (let i = 0; i < n; i++) {
      const outWb = XLSX.utils.book_new();
      sheets.forEach((s, sIdx) => {
        const [start, end] = partitions[sIdx][i];
        const slice = s.data.slice(start, end);
        const aoa = s.header ? [s.header, ...slice] : slice;
        // Ensure we always create the sheet, even when empty, so the
        // 4-sheet structure is preserved in every output file.
        const ws = XLSX.utils.aoa_to_sheet(aoa.length ? aoa : [[]]);
        XLSX.utils.book_append_sheet(outWb, ws, safeSheetName(s.name));
      });
      const wbout = XLSX.write(outWb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([wbout], {
        type:
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      outputs.push({ index: i + 1, blob });
    }
    return outputs;
  }

  /**
   * Split `total` items into `n` contiguous ranges of nearly equal size.
   * Earlier ranges absorb the remainder when total % n != 0.
   * @returns {[number, number][]}
   */
  function evenSplit(total, n) {
    const base = Math.floor(total / n);
    const extra = total % n;
    const ranges = [];
    let cursor = 0;
    for (let i = 0; i < n; i++) {
      const size = base + (i < extra ? 1 : 0);
      ranges.push([cursor, cursor + size]);
      cursor += size;
    }
    return ranges;
  }

  function renderResults(files, baseName) {
    resultsEl.innerHTML = '';
    files.forEach(({ index, blob }) => {
      const url = URL.createObjectURL(blob);
      const fname = `${baseName}_part_${pad(index, files.length)}.xlsx`;
      const a = document.createElement('a');
      a.href = url;
      a.download = fname;
      a.textContent = `Download ${fname}`;
      resultsEl.appendChild(a);
      // Trigger download immediately as well.
      a.click();
    });
  }

  // ---------- helpers ----------
  function setStatus(text, kind) {
    statusEl.textContent = text;
    statusEl.className = 'status' + (kind ? ' ' + kind : '');
  }

  function stripExt(name) {
    return name.replace(/\.[^.]+$/, '');
  }

  function pad(num, totalCount) {
    const width = String(totalCount).length;
    return String(num).padStart(width, '0');
  }

  // Excel sheet names cannot exceed 31 chars or contain certain symbols.
  function safeSheetName(name) {
    let s = String(name).replace(/[\\/?*[\]:]/g, '_');
    if (s.length > 31) s = s.slice(0, 31);
    return s || 'Sheet';
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }
})();
