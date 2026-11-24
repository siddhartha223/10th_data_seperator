import * as XLSX from 'xlsx';
import JSZip from 'jszip';
import { getUnrarModule } from '@unrar-browser/core';

(() => {
  const input = document.getElementById('file-input');
  const dropzone = document.getElementById('dropzone');
  const browse = document.getElementById('browse-button');
  const fileList = document.getElementById('file-list');
  const processButton = document.getElementById('process-button');
  const status = document.getElementById('status');
  const results = document.getElementById('results');
  const progressWrap = document.getElementById('progress-wrap');
  const progressBar = document.getElementById('progress-bar');
  const progressMessage = document.getElementById('progress-message');
  const progressPercent = document.getElementById('progress-percent');
  const preview = document.getElementById('preview');
  const previewTable = document.getElementById('preview-table');
  const previewCount = document.getElementById('preview-count');
  const accepted = /\.(csv|xlsx|xls|zip|rar)$/i;
  let files = [];
  let readyDownloads = [];

  browse.addEventListener('click', (event) => { event.stopPropagation(); input.click(); });
  dropzone.addEventListener('click', (event) => { if (event.target !== browse) input.click(); });
  dropzone.addEventListener('keydown', (event) => { if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); input.click(); } });
  input.addEventListener('change', () => { addFiles([...input.files]); input.value = ''; });
  ['dragenter', 'dragover'].forEach((type) => dropzone.addEventListener(type, (event) => { event.preventDefault(); dropzone.classList.add('dragging'); }));
  ['dragleave', 'drop'].forEach((type) => dropzone.addEventListener(type, (event) => { event.preventDefault(); dropzone.classList.remove('dragging'); }));
  dropzone.addEventListener('drop', (event) => addFiles([...event.dataTransfer.files]));
  document.querySelectorAll('input[name="output"]').forEach((radio) => radio.addEventListener('change', () => {
    document.querySelectorAll('.output-card').forEach((card) => card.classList.toggle('selected', card.dataset.mode === radio.value));
    results.hidden = true; preview.hidden = true; progressWrap.hidden = true; readyDownloads = [];
  }));
  processButton.addEventListener('click', processFiles);

  function addFiles(incoming) {
    const invalid = incoming.filter((file) => !accepted.test(file.name));
    const valid = incoming.filter((file) => accepted.test(file.name));
    files = [...files, ...valid.filter((file) => !files.some((existing) => existing.name === file.name && existing.size === file.size))];
    renderFiles();
    if (invalid.length) showStatus(`Skipped ${invalid.length} unsupported file${invalid.length === 1 ? '' : 's'}. Choose CSV, Excel, ZIP, or RAR files.`, true);
    else clearStatus();
    results.hidden = true;
    preview.hidden = true;
    readyDownloads = [];
  }

  function renderFiles() {
    fileList.replaceChildren();
    files.forEach((file, index) => {
      const row = document.createElement('div'); row.className = 'file-row';
      const badge = document.createElement('span'); badge.className = 'file-badge'; badge.textContent = file.name.toLowerCase().endsWith('.zip') ? 'ZIP' : file.name.split('.').pop().toUpperCase().slice(0, 4);
      const meta = document.createElement('span'); meta.className = 'file-meta';
      const name = document.createElement('strong'); name.textContent = file.name;
      const size = document.createElement('span'); size.textContent = formatSize(file.size);
      meta.append(name, size);
      const remove = document.createElement('button'); remove.type = 'button'; remove.className = 'remove-file'; remove.setAttribute('aria-label', `Remove ${file.name}`); remove.textContent = '×';
      remove.addEventListener('click', () => { files.splice(index, 1); renderFiles(); clearStatus(); results.hidden = true; preview.hidden = true; progressWrap.hidden = true; readyDownloads = []; });
      row.append(badge, meta, remove); fileList.append(row);
    });
    processButton.disabled = files.length === 0;
  }

  async function processFiles() {
    setLoading(true); results.hidden = true; preview.hidden = true; clearStatus();
    progressWrap.hidden = false; updateProgress(2, 'Preparing files…');
    try {
      const inputs = [];
      for (const [index, file] of files.entries()) {
        const base = (index / files.length) * 35;
        const span = 35 / files.length;
        updateProgress(5 + base, `Opening ${file.name}…`);
        inputs.push(...await readInputFile(file, (percent) => updateProgress(5 + base + span * percent / 100, `Opening ${file.name}…`)));
      }
      if (!inputs.length) throw new Error('No CSV or Excel files were found in the selected ZIP.');
      const extracted = [];
      let totalRows = 0;
      for (const [index, entry] of inputs.entries()) {
        updateProgress(40 + ((index / inputs.length) * 38), `Filtering ${entry.name}…`);
        await nextFrame();
        const parsed = parseWorkbook(entry.data, entry.name);
        totalRows += parsed.rows.length;
        extracted.push({ name: entry.source || entry.name, headers: parsed.headers, rows: parsed.rows.filter((row) => isClassTen(row, parsed.headers)) });
        updateProgress(40 + (((index + 1) / inputs.length) * 38), `Filtered ${index + 1} of ${inputs.length} files`);
      }
      const totalMatches = extracted.reduce((sum, item) => sum + item.rows.length, 0);
      if (!totalMatches) throw new Error(`Scanned ${totalRows.toLocaleString()} rows across ${inputs.length} file${inputs.length === 1 ? '' : 's'}; no Class X records matched. Check that the files contain classId or classDesc columns.`);
      const mode = document.querySelector('input[name="output"]:checked').value;
      const previewData = combineRows(extracted);
      updateProgress(82, 'Preparing your download…'); await nextFrame();
      if (mode === 'combined') readyDownloads = [{ name: 'class-x-students.xlsx', blob: makeWorkbook(previewData, 'Class X Students') }];
      else {
        const zip = new JSZip();
        const usedNames = new Map();
        const matched = extracted.filter((entry) => entry.rows.length);
        for (const [index, item] of matched.entries()) {
          const base = safeBaseName(item.name);
          const count = (usedNames.get(base.toLowerCase()) || 0) + 1; usedNames.set(base.toLowerCase(), count);
          const suffix = count > 1 ? ` (${count})` : '';
          const filename = `${base}${suffix} - Class X.xlsx`;
          zip.file(filename, makeWorkbook({ headers: item.headers, rows: item.rows }, 'Class X Students'));
          updateProgress(82 + ((index + 1) / matched.length) * 12, `Building workbook ${index + 1} of ${matched.length}…`);
          await nextFrame();
        }
        readyDownloads = [{ name: 'class-x-files.zip', blob: await zip.generateAsync({ type: 'blob', compression: 'DEFLATE', compressionOptions: { level: 5 } }, (meta) => updateProgress(94 + meta.percent * 0.06, 'Compressing your ZIP…')) }];
      }
      const withMatches = extracted.filter((entry) => entry.rows.length).length;
      updateProgress(100, 'Complete');
      showResult(totalMatches, extracted.length, withMatches, totalRows, previewData);
    } catch (error) {
      showStatus(error instanceof Error ? error.message : 'Something went wrong while reading those files.', true);
      progressWrap.hidden = true;
    } finally { setLoading(false); }
  }

  async function readInputFile(file, onProgress) {
    const extension = file.name.split('.').pop().toLowerCase();
    if (extension === 'rar') return extractRar(await file.arrayBuffer(), file.name, onProgress);
    if (extension === 'zip') return readZip(await file.arrayBuffer(), file.name, onProgress);
    return [{ name: file.name, data: await file.arrayBuffer() }];
  }

  async function readZip(data, archiveName, onProgress) {
    const zip = await JSZip.loadAsync(data, { onUpdate: (meta) => onProgress(meta.percent || 0) });
    const entries = Object.values(zip.files).filter((entry) => !entry.dir && accepted.test(entry.name) && !entry.name.split('/').some((part) => part.startsWith('.')));
    const output = [];
    for (const entry of entries) {
      const entryData = await entry.async('arraybuffer', (meta) => onProgress(meta.percent || 0));
      const name = entry.name.split('/').pop();
      const extension = name.split('.').pop().toLowerCase();
      if (extension === 'rar') output.push(...await extractRar(entryData, name, onProgress));
      else if (extension === 'zip') output.push(...await readZip(entryData, name, onProgress));
      else output.push({ name, data: entryData, source: `${archiveName}/${entry.name}` });
    }
    return output;
  }

  async function extractRar(data, archiveName, onProgress) {
    updateProgress(progressBar.value, `Loading RAR support for ${archiveName}…`);
    const unrar = await getUnrarModule('/');
    const archivePath = `/input-${Date.now()}-${Math.random().toString(36).slice(2)}.rar`;
    unrar.FS.writeFile(archivePath, new Uint8Array(data));
    const archive = new unrar.Archive(new unrar.CommandData());
    const extracted = [];
    try {
      if (!archive.openFile(archivePath) || !archive.isArchive(true)) throw new Error(`“${archiveName}” is not a supported RAR archive.`);
      let result;
      let fileCount = 0;
      while ((result = archive.readHeader()) > 0) {
        const isFile = archive.getHeaderType() === unrar.HeaderType.HEAD_FILE;
        if (isFile && !archive.isDirectory()) {
          const fileName = archive.getFileName().replace(/\\/g, '/').split('/').pop();
          const extension = fileName.split('.').pop().toLowerCase();
          if (accepted.test(fileName) && !fileName.startsWith('.')) {
            const fileData = archive.readFileData();
            let bytes;
            if (fileData instanceof Uint8Array) bytes = fileData;
            else if (fileData && typeof fileData.size === 'function' && typeof fileData.get === 'function') {
              bytes = new Uint8Array(fileData.size());
              for (let i = 0; i < bytes.length; i += 1) bytes[i] = fileData.get(i);
            } else throw new Error(`Could not read “${fileName}” from ${archiveName}.`);
            fileCount += 1;
            const childData = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
            if (extension === 'rar') extracted.push(...await extractRar(childData, fileName, onProgress));
            else if (extension === 'zip') extracted.push(...await readZip(childData, fileName, onProgress));
            else extracted.push({ name: fileName, data: childData, source: `${archiveName}/${fileName}` });
          } else archive.seekToNext();
        } else archive.seekToNext();
        if (fileCount > 0) { onProgress(Math.min(99, fileCount)); await nextFrame(); }
      }
      if (result < 0) throw new Error(`Could not finish extracting “${archiveName}”. It may be password-protected or damaged.`);
      return extracted;
    } finally {
      try { unrar.FS.unlink(archivePath); } catch { /* The archive may already have been released. */ }
    }
  }

  function parseWorkbook(data, name) {
    let workbook;
    try { workbook = XLSX.read(data, { type: 'array', cellDates: false, dense: false }); }
    catch { throw new Error(`Could not read “${name}”. Make sure it is a valid CSV or Excel file.`); }
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: false });
      const headerIndex = matrix.findIndex((row) => row.some((cell) => String(cell).trim() !== ''));
      if (headerIndex < 0) continue;
      const headers = matrix[headerIndex].map((cell, i) => String(cell).trim() || `Column ${i + 1}`);
      const rows = matrix.slice(headerIndex + 1).filter((row) => row.some((cell) => String(cell).trim() !== '')).map((row) => {
        const record = {}; headers.forEach((header, i) => { record[header] = row[i] ?? ''; }); return record;
      });
      if (rows.length || headers.length) return { headers, rows };
    }
    throw new Error(`“${name}” appears to be empty.`);
  }

  function normalized(value) { return String(value ?? '').trim().toUpperCase().replace(/[\s_-]+/g, ' '); }
  function isClassTen(row, headers) {
    const classIdKey = headers.find((header) => normalized(header).replace(/ /g, '') === 'CLASSID');
    const classDescKey = headers.find((header) => normalized(header).replace(/ /g, '') === 'CLASSDESC');
    const id = classIdKey ? normalized(row[classIdKey]) : '';
    const desc = classDescKey ? normalized(row[classDescKey]) : '';
    const idMatch = /^(10|X)$/.test(id);
    const descMatch = /^(X|10|10TH)$/.test(desc) || /^(CLASS|GRADE) (X|10|10TH)$/.test(desc);
    return idMatch || descMatch;
  }

  function combineRows(items) {
    const headers = []; const seen = new Set();
    items.forEach((item) => item.headers.forEach((header) => { if (!seen.has(header)) { seen.add(header); headers.push(header); } }));
    const sourceHeader = headers.some((header) => normalized(header) === 'SOURCE FILE') ? 'Input File' : 'Source File';
    const finalHeaders = [...headers, sourceHeader];
    return { headers: finalHeaders, rows: items.flatMap((item) => item.rows.map((row) => ({ ...row, [sourceHeader]: item.name }))) };
  }

  function makeWorkbook(data, sheetName) {
    const sheet = XLSX.utils.json_to_sheet(data.rows, { header: data.headers, skipHeader: false });
    sheet['!cols'] = data.headers.map((header) => ({ wch: Math.min(36, Math.max(12, header.length + 2)) }));
    if (data.rows.length) sheet['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: data.rows.length, c: data.headers.length - 1 } }) };
    const workbook = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(workbook, sheet, sheetName);
    const bytes = XLSX.write(workbook, { bookType: 'xlsx', type: 'array', compression: true });
    return new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  function safeBaseName(name) { return name.replace(/\\/g, '/').split('/').pop().replace(/\.[^.]+$/, '').replace(/[<>:"|?*]/g, '_').trim() || 'students'; }
  function formatSize(bytes) { if (bytes < 1024) return `${bytes} B`; if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`; return `${(bytes / (1024 * 1024)).toFixed(1)} MB`; }
  function setLoading(loading) { processButton.disabled = loading || files.length === 0; processButton.classList.toggle('loading', loading); processButton.querySelector('span').textContent = loading ? 'Reading and filtering files…' : 'Extract Class X data'; }
  function clearStatus() { status.textContent = ''; status.classList.remove('error'); }
  function showStatus(message, error = false) { status.textContent = message; status.classList.toggle('error', error); }
  function showResult(count, sources, matchedSources, total, data) {
    status.textContent = '';
    const resultIcon = document.createElement('span'); resultIcon.className = 'results-icon'; resultIcon.textContent = '✓';
    const copy = document.createElement('span'); copy.className = 'results-copy';
    const title = document.createElement('strong'); title.textContent = `${count.toLocaleString()} Class X student${count === 1 ? '' : 's'} ready`;
    const detail = document.createElement('span'); detail.textContent = `From ${matchedSources} of ${sources} file${sources === 1 ? '' : 's'} · ${total.toLocaleString()} rows scanned`;
    copy.append(title, detail);
    const button = document.createElement('button'); button.type = 'button'; button.className = 'download-button'; button.textContent = `Download ${readyDownloads[0].name.endsWith('.zip') ? 'ZIP' : 'Excel'}`;
    button.addEventListener('click', () => {
      try { download(readyDownloads[0]); }
      catch (error) { showStatus(`Download could not start. ${error instanceof Error ? error.message : 'Please try again.'}`, true); }
    });
    results.replaceChildren(resultIcon, copy, button); results.hidden = false;
    renderPreview(data);
  }
  function renderPreview(data) {
    previewCount.textContent = `${data.rows.length.toLocaleString()} matching rows · showing ${Math.min(data.rows.length, 10).toLocaleString()}`;
    const head = document.createElement('thead'); const headerRow = document.createElement('tr');
    data.headers.forEach((header) => { const cell = document.createElement('th'); cell.textContent = header; headerRow.append(cell); });
    head.append(headerRow);
    const body = document.createElement('tbody');
    data.rows.slice(0, 10).forEach((row) => {
      const tr = document.createElement('tr');
      data.headers.forEach((header) => { const cell = document.createElement('td'); cell.textContent = String(row[header] ?? ''); tr.append(cell); });
      body.append(tr);
    });
    previewTable.replaceChildren(head, body); preview.hidden = false;
  }
  function updateProgress(value, message) {
    const percentage = Math.max(0, Math.min(100, Math.round(value)));
    progressBar.value = percentage; progressPercent.textContent = `${percentage}%`; progressMessage.textContent = message;
  }
  function nextFrame() { return new Promise((resolve) => requestAnimationFrame(() => resolve())); }
  function download(item) {
    if (!(item.blob instanceof Blob)) throw new Error('The file could not be prepared. Run the extraction again.');
    const url = URL.createObjectURL(item.blob);
    const link = document.createElement('a'); link.href = url; link.download = item.name; link.style.display = 'none';
    document.body.append(link); link.click(); link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
  }
})();
