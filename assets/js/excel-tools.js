(function () {
  'use strict';
  const App = (window.App = window.App || {});

  function requireOffice() { if (!App.hasOffice()) throw new Error(App.t('demo')); }

  async function getWorkbookId() {
    if (!App.hasOffice()) return 'browser';
    return new Promise((resolve, reject) => {
      const settings = Office.context.document.settings;
      let v = settings.get(App.STORAGE_KEYS.workbookId);
      if (v) return resolve(v);
      v = App.id();
      settings.set(App.STORAGE_KEYS.workbookId, v);
      settings.saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? resolve(v) : reject(new Error(r.error?.message || 'Failed to save workbook ID')));
    });
  }
  async function saveOfficeSetting(key, value) {
    return new Promise((resolve, reject) => {
      Office.context.document.settings.set(key, value);
      Office.context.document.settings.saveAsync(r => r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(new Error(r.error?.message || 'Failed to save document setting')));
    });
  }
  function loadOfficeSetting(key, fallback) {
    try { return Office.context.document.settings.get(key) || fallback; } catch { return fallback; }
  }

  async function getSheetMap(context, sheets) {
    let map = {};
    try { map = JSON.parse(loadOfficeSetting(App.STORAGE_KEYS.sheetMap, '{}') || '{}'); } catch { map = {}; }
    let max = Object.values(map).reduce((a, b) => Math.max(a, Number(b) || 0), 0);
    let dirty = false;
    for (const s of sheets) {
      if (!map[s.id]) { map[s.id] = ++max; dirty = true; }
    }
    if (dirty) await saveOfficeSetting(App.STORAGE_KEYS.sheetMap, JSON.stringify(map)).catch(() => {});
    return new Map(Object.entries(map));
  }
  async function worksheetById(context, stableId) {
    const sheets = context.workbook.worksheets;
    sheets.load('items');
    await context.sync();
    for (const sheet of sheets.items) sheet.load('id,name');
    await context.sync();
    const map = await getSheetMap(context, sheets.items);
    for (const sheet of sheets.items) if (Number(map.get(sheet.id)) === Number(stableId)) return sheet;
    return null;
  }
  function colName(index) { let s = '', n = index; while (n >= 0) { s = String.fromCharCode(n % 26 + 65) + s; n = Math.floor(n / 26) - 1; } return s; }
  function a1(row, col) { return `${colName(col)}${row + 1}`; }
  function quoteSheetName(name) { return `'${String(name || '').replace(/'/g, "''")}'`; }
  function parseStart(address) {
    const part = String(address || 'A1').split('!').pop().split(':')[0].replace(/'/g, '');
    const m = part.match(/([A-Z]+)(\d+)/i);
    if (!m) return { startCol: 0, startRow: 0 };
    const col = m[1].toUpperCase().split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    return { startCol: col, startRow: Number(m[2]) - 1 };
  }
  function rangeForDimension(ref, count, dim) {
    count = Number(count || 1);
    if (dim === 'rows') {
      const start = Number(ref || 1);
      return `${start}:${start + count - 1}`;
    }
    const start = String(ref || 'A').toUpperCase();
    const n = start.split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0) - 1;
    return `${start}:${colName(n + count - 1)}`;
  }

  async function getWorkbookMetadata() {
    requireOffice();
    return Excel.run(async context => {
      const wb = context.workbook;
      wb.load('name');
      const sheets = wb.worksheets;
      sheets.load('items');
      const active = sheets.getActiveWorksheet();
      active.load('id,name');
      const selected = wb.getSelectedRange();
      selected.load('address');
      await context.sync();
      for (const sheet of sheets.items) sheet.load('id,name,position,visibility');
      await context.sync();
      const map = await getSheetMap(context, sheets.items);
      const info = [];
      for (const sheet of sheets.items) {
        const used = sheet.getUsedRangeOrNullObject();
        used.load('address,rowCount,columnCount');
        await context.sync();
        info.push({ sheetId: Number(map.get(sheet.id)), name: sheet.name, nativeId: sheet.id, position: sheet.position, visibility: sheet.visibility, maxRows: used.isNullObject ? 0 : used.rowCount, maxColumns: used.isNullObject ? 0 : used.columnCount, usedRange: used.isNullObject ? null : used.address.split('!').pop() });
      }
      return { success: true, workbookId: App.state.workbookId || 'workbook', workbookName: wb.name || '', activeSheet: { sheetId: Number(map.get(active.id)), name: active.name }, selectedRange: selected.address.includes('!') ? selected.address.split('!').pop() : selected.address, worksheets: info };
    });
  }

  async function getCellRanges(args) {
    requireOffice();
    const { sheetId, ranges, includeStyles = true, cellLimit = 2000 } = args;
    const requestedRanges = Array.isArray(ranges) ? ranges : [];
    const limit = Math.max(1, Number(cellLimit || 2000));
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      sheet.load('name');
      const used = sheet.getUsedRangeOrNullObject();
      used.load('address');
      await context.sync();
      const dimension = used.isNullObject ? 'A1' : used.address.split('!').pop();
      const cells = {}, formulas = {}, styles = {};
      let count = 0, hasMore = false;
      for (let rangeIndex = 0; rangeIndex < requestedRanges.length; rangeIndex++) {
        if (count >= limit) { hasMore = true; break; }
        const rangeText = requestedRanges[rangeIndex];
        const range = sheet.getRange(rangeText);
        range.load('values,formulas,address,rowCount,columnCount');
        await context.sync();
        const start = parseStart(range.address);
        const styleCells = [];
        let stoppedInsideRange = false;
        outer: for (let r = 0; r < range.rowCount; r++) {
          for (let c = 0; c < range.columnCount; c++) {
            if (count >= limit) { stoppedInsideRange = true; break outer; }
            const key = a1(start.startRow + r, start.startCol + c);
            const val = range.values[r][c];
            const formula = range.formulas[r][c];
            let returnedSomething = false;
            if (val !== null && val !== '' && typeof val !== 'undefined') { cells[key] = val; returnedSomething = true; }
            if (typeof formula === 'string' && formula.startsWith('=')) { formulas[key] = formula; returnedSomething = true; }
            if (returnedSomething) {
              count++;
              if (includeStyles) styleCells.push([key, range.getCell(r, c)]);
            }
          }
        }
        if (stoppedInsideRange || (count >= limit && rangeIndex < requestedRanges.length - 1)) hasMore = true;
        if (includeStyles && styleCells.length) {
          styleCells.forEach(([, cell]) => { cell.format.font.load('name,size,color,bold,italic,underline,strikethrough'); cell.format.fill.load('color'); cell.format.load('horizontalAlignment'); });
          await context.sync();
          for (const [key, cell] of styleCells) {
            const style = {};
            if (cell.format.font.name) style.fontFamily = cell.format.font.name;
            if (cell.format.font.size) style.fontSize = cell.format.font.size;
            if (cell.format.font.bold !== null) style.fontWeight = cell.format.font.bold ? 'bold' : 'normal';
            if (cell.format.font.italic !== null) style.fontStyle = cell.format.font.italic ? 'italic' : 'normal';
            if (cell.format.font.color) style.fontColor = cell.format.font.color;
            if (cell.format.fill.color) style.backgroundColor = cell.format.fill.color;
            if (cell.format.horizontalAlignment) style.horizontalAlignment = String(cell.format.horizontalAlignment).toLowerCase();
            if (Object.keys(style).length) styles[key] = style;
          }
        }
        if (hasMore) break;
      }
      return { success: true, hasMore, worksheet: { name: sheet.name, sheetId, dimension, cells, formulas, styles, borders: {} } };
    });
  }

  async function getRangeAsCsv(args) {
    requireOffice();
    const { sheetId, range, includeHeaders = true, maxRows = 500 } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      sheet.load('name');
      const r = sheet.getRange(range);
      r.load('values,rowCount,columnCount');
      await context.sync();
      const start = includeHeaders ? 0 : 1;
      const rows = [];
      for (let i = start; i < Math.min(r.rowCount, start + maxRows); i++) rows.push(r.values[i].map(csvEscape).join(','));
      return { success: true, csv: rows.join('\n'), rowCount: rows.length, columnCount: r.columnCount, hasMore: r.rowCount - start > maxRows, sheetName: sheet.name };
    });
  }
  function csvEscape(v) { if (v == null) return ''; const s = String(v); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; }

  async function searchData(args) {
    requireOffice();
    const { searchTerm, sheetId, range, offset = 0, options = {} } = args;
    const { matchCase = false, matchEntireCell = false, matchFormulas = false, useRegex = false, maxResults = 500 } = options;
    return Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      sheets.load('items'); await context.sync();
      for (const s of sheets.items) s.load('id,name'); await context.sync();
      const map = await getSheetMap(context, sheets.items);
      const targetSheets = sheetId ? [await worksheetById(context, sheetId)].filter(Boolean) : sheets.items;
      const matches = [];
      const regex = useRegex ? new RegExp(searchTerm, matchCase ? 'g' : 'ig') : null;
      for (const sheet of targetSheets) {
        const r = range ? sheet.getRange(range) : sheet.getUsedRangeOrNullObject();
        r.load('values,formulas,address,rowCount,columnCount'); await context.sync();
        if (r.isNullObject) continue;
        const start = parseStart(r.address);
        for (let row = 0; row < r.rowCount; row++) for (let col = 0; col < r.columnCount; col++) {
          const val = r.values[row][col];
          const formula = r.formulas[row][col];
          const text = String(matchFormulas && formula ? formula : (val ?? ''));
          let ok;
          if (regex) { regex.lastIndex = 0; ok = regex.test(text); }
          else { const a = matchCase ? text : text.toLowerCase(); const b = matchCase ? searchTerm : searchTerm.toLowerCase(); ok = matchEntireCell ? a === b : a.includes(b); }
          if (ok) matches.push({ sheetName: sheet.name, sheetId: Number(map.get(sheet.id)), a1: a1(start.startRow + row, start.startCol + col), value: val, formula: typeof formula === 'string' && formula.startsWith('=') ? formula : null, row: start.startRow + row + 1, column: start.startCol + col + 1 });
        }
      }
      const slice = matches.slice(offset, offset + maxResults);
      return { success: true, matches: slice, totalFound: matches.length, returned: slice.length, offset, hasMore: offset + maxResults < matches.length, searchTerm, nextOffset: offset + maxResults < matches.length ? offset + maxResults : null };
    });
  }

  async function setCellRange(args) {
    requireOffice();
    const { sheetId, range, cells, copyToRange, resizeWidth, resizeHeight, allow_overwrite = false } = args;
    const maxWriteCells = 10000;
    if (!Array.isArray(cells) || !cells.length || !cells.every(Array.isArray)) throw new Error('cells must be a non-empty two-dimensional array');
    const writeCellCount = cells.reduce((sum, row) => sum + row.length, 0);
    if (writeCellCount > maxWriteCells) throw new Error(`Refusing to write ${writeCellCount} cells in one call. Split the write into chunks of ${maxWriteCells} cells or fewer.`);
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      sheet.load('name');
      const r = sheet.getRange(range);
      r.load('values,formulas,address,rowCount,columnCount'); await context.sync();
      if (cells.length !== r.rowCount) throw new Error(`cells row count (${cells.length}) does not match target range row count (${r.rowCount})`);
      for (let i = 0; i < cells.length; i++) {
        if (cells[i].length !== r.columnCount) throw new Error(`cells[${i}] column count (${cells[i].length}) does not match target range column count (${r.columnCount})`);
      }
      const start = parseStart(r.address);
      if (!allow_overwrite) {
        const occupied = [];
        for (let i = 0; i < r.rowCount; i++) for (let j = 0; j < r.columnCount; j++) {
          if ((r.values[i][j] !== null && r.values[i][j] !== '') || (typeof r.formulas[i][j] === 'string' && r.formulas[i][j].startsWith('='))) occupied.push(a1(start.startRow + i, start.startCol + j));
        }
        if (occupied.length) throw new Error(`Would overwrite ${occupied.length} non-empty cell(s): ${occupied.slice(0, 10).join(', ')}${occupied.length > 10 ? '...' : ''}. Retry with allow_overwrite=true if confirmed.`);
      }
      const matrix = cells.map(row => row.map(cell => cell && typeof cell === 'object' ? (cell.formula || (cell.value ?? null)) : (cell ?? null)));
      r.formulas = matrix;
      const sheetPrefix = `${quoteSheetName(sheet.name)}!`;
      for (let i = 0; i < cells.length; i++) {
        for (let j = 0; j < cells[i].length; j++) {
          const cellAddress = `${sheetPrefix}${a1(start.startRow + i, start.startCol + j)}`;
          applyCellOptions(context, r.getCell(i, j), cells[i][j], cellAddress);
        }
      }
      if (copyToRange) sheet.getRange(copyToRange).copyFrom(r, Excel.RangeCopyType.all, false, false);
      applyResize(r, resizeWidth, resizeHeight);
      await context.sync();
      const dirty = [{ sheetId, range }];
      if (copyToRange) dirty.push({ sheetId, range: copyToRange });
      return { success: true, writtenRange: range, copiedTo: copyToRange || null, _dirtyRanges: dirty };
    });
  }
  function applyCellOptions(context, cell, spec, cellAddress) {
    if (!spec || typeof spec !== 'object') return;
    if (spec.note) {
      if (!context.workbook.notes || typeof context.workbook.notes.add !== 'function') throw new Error('Excel Notes API is not available in this Office host');
      context.workbook.notes.add(cellAddress, String(spec.note));
    }
    const st = spec.cellStyles || {};
    if (st.fontWeight) cell.format.font.bold = st.fontWeight === 'bold';
    if (st.fontStyle) cell.format.font.italic = st.fontStyle === 'italic';
    if (st.fontSize) cell.format.font.size = st.fontSize;
    if (st.fontFamily) cell.format.font.name = st.fontFamily;
    if (st.fontColor) cell.format.font.color = st.fontColor;
    if (st.backgroundColor) cell.format.fill.color = st.backgroundColor;
    if (st.horizontalAlignment) cell.format.horizontalAlignment = st.horizontalAlignment;
    if (st.numberFormat) cell.numberFormat = [[st.numberFormat]];
    const borderMap = { top: 'EdgeTop', bottom: 'EdgeBottom', left: 'EdgeLeft', right: 'EdgeRight' };
    for (const [side, val] of Object.entries(spec.borderStyles || {})) {
      if (!val || !borderMap[side]) continue;
      const b = cell.format.borders.getItem(borderMap[side]);
      if (val.color) b.color = val.color;
      if (val.weight) b.weight = val.weight;
      if (val.style) b.style = val.style === 'solid' ? 'Continuous' : val.style;
    }
  }
  function applyResize(r, width, height) { if (width) r.getEntireColumn().format.columnWidth = width.value; if (height) r.getEntireRow().format.rowHeight = height.value; }

  async function clearCellRange(args) {
    requireOffice();
    const { sheetId, range, clearType = 'contents' } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      const map = { contents: Excel.ClearApplyTo.contents, formats: Excel.ClearApplyTo.formats, all: Excel.ClearApplyTo.all };
      sheet.getRange(range).clear(map[clearType] || Excel.ClearApplyTo.contents);
      await context.sync();
      return { success: true, cleared: range, clearType, _dirtyRanges: [{ sheetId, range }] };
    });
  }

  async function copyTo(args) {
    requireOffice();
    const { sheetId, sourceRange, destinationRange } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      sheet.getRange(destinationRange).copyFrom(sheet.getRange(sourceRange), Excel.RangeCopyType.all, false, false);
      await context.sync();
      return { success: true, sourceRange, destinationRange, _dirtyRanges: [{ sheetId, range: destinationRange }] };
    });
  }

  async function modifySheetStructure(args) {
    requireOffice();
    const { sheetId, operation, dimension = 'rows', reference, count = 1, position = 'before' } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      if (operation === 'unfreeze') sheet.freezePanes.unfreeze();
      else if (operation === 'freeze') dimension === 'columns' ? sheet.freezePanes.freezeColumns(Number(count || reference || 1)) : sheet.freezePanes.freezeRows(Number(count || reference || 1));
      else {
        let ref = reference;
        if (operation === 'insert' && position === 'after') {
          if (dimension === 'rows') ref = String(Number(reference || 1) + 1);
          else ref = colName(String(reference || 'A').toUpperCase().split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0));
        }
        const range = sheet.getRange(rangeForDimension(ref, count, dimension));
        if (operation === 'insert') range.insert(dimension === 'rows' ? Excel.InsertShiftDirection.down : Excel.InsertShiftDirection.right);
        if (operation === 'delete') range.delete(dimension === 'rows' ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left);
        if (operation === 'hide') dimension === 'rows' ? range.rowHidden = true : range.columnHidden = true;
        if (operation === 'unhide') dimension === 'rows' ? range.rowHidden = false : range.columnHidden = false;
      }
      await context.sync();
      return { success: true, operation, dimension, sheetId, _dirtyRanges: [{ sheetId, range: '*' }] };
    });
  }

  async function modifyWorkbookStructure(args) {
    requireOffice();
    const { operation, sheetId, sheetName, newName, tabColor } = args;
    return Excel.run(async context => {
      let result = { success: true, operation };
      if (operation === 'create') {
        const s = context.workbook.worksheets.add(sheetName || 'Sheet');
        if (tabColor) s.tabColor = tabColor;
        s.load('id,name'); await context.sync();
        const map = await getSheetMap(context, [s]);
        result.sheetId = Number(map.get(s.id)); result.name = s.name;
      } else {
        const sheet = await worksheetById(context, sheetId);
        if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
        if (operation === 'delete') sheet.delete();
        if (operation === 'rename') { sheet.name = newName || sheetName || sheet.name; if (tabColor) sheet.tabColor = tabColor; }
        if (operation === 'duplicate') {
          const copy = sheet.copy(Excel.WorksheetPositionType.after, sheet);
          if (newName) copy.name = newName;
          if (tabColor) copy.tabColor = tabColor;
          copy.load('id,name'); await context.sync();
          const map = await getSheetMap(context, [copy]);
          result.sheetId = Number(map.get(copy.id)); result.name = copy.name;
        }
      }
      await context.sync();
      result._dirtyRanges = result.sheetId ? [{ sheetId: result.sheetId, range: '*' }] : (sheetId ? [{ sheetId, range: '*' }] : []);
      return result;
    });
  }

  async function resizeRange(args) {
    requireOffice();
    const { sheetId, range, width, height } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      const r = range ? sheet.getRange(range) : sheet.getUsedRangeOrNullObject();
      applyResize(r, width, height);
      await context.sync();
      return { success: true, sheetId, range: range || '*', _dirtyRanges: [{ sheetId, range: range || '*' }] };
    });
  }

  async function getAllObjects(args = {}) {
    requireOffice();
    const { sheetId, id: objectId } = args;
    return Excel.run(async context => {
      const sheets = context.workbook.worksheets;
      sheets.load('items'); await context.sync();
      for (const s of sheets.items) s.load('id,name'); await context.sync();
      const map = await getSheetMap(context, sheets.items);
      const targetSheets = sheetId ? [await worksheetById(context, sheetId)].filter(Boolean) : sheets.items;
      const objects = [];
      for (const sheet of targetSheets) {
        const charts = sheet.charts; charts.load('items');
        const pivots = sheet.pivotTables; pivots.load('items');
        await context.sync();
        for (const c of charts.items) { c.load('id,name'); await context.sync(); if (!objectId || c.id === objectId) objects.push({ id: c.id, type: 'chart', name: c.name, sheetId: Number(map.get(sheet.id)), sheetName: sheet.name }); }
        for (const p of pivots.items) { p.load('id,name'); await context.sync(); if (!objectId || p.id === objectId) objects.push({ id: p.id, type: 'pivotTable', name: p.name, sheetId: Number(map.get(sheet.id)), sheetName: sheet.name }); }
      }
      return { success: true, objects };
    });
  }

  async function modifyObject(args) {
    requireOffice();
    const { operation, sheetId, objectType, id: objectId, properties = {} } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      let result = { success: true, operation, objectType, sheetId };
      if (objectType === 'chart') {
        if (operation === 'create') {
          if (!properties.source || !properties.chartType) throw new Error('Chart creation requires source and chartType');
          const chart = sheet.charts.add(properties.chartType, sheet.getRange(properties.source), Excel.ChartSeriesBy.auto);
          if (properties.name) chart.name = properties.name;
          if (properties.title) { chart.title.text = properties.title; chart.title.visible = true; }
          if (properties.anchor) chart.setPosition(properties.anchor);
          chart.load('id,name'); await context.sync();
          result.id = chart.id; result.name = chart.name;
        } else {
          if (!objectId) throw new Error('Chart update/delete requires id');
          const chart = sheet.charts.getItem(objectId);
          if (operation === 'delete') chart.delete();
          if (operation === 'update') {
            if (properties.title) { chart.title.text = properties.title; chart.title.visible = true; }
            if (properties.name) chart.name = properties.name;
            if (properties.anchor) chart.setPosition(properties.anchor);
          }
        }
      } else if (objectType === 'pivotTable') {
        if (operation === 'create') {
          if (!properties.source || !properties.range) throw new Error('PivotTable creation requires source and range');
          const pivot = sheet.pivotTables.add(properties.name || `Pivot_${Date.now()}`, properties.source, properties.range);
          await context.sync();
          const fieldErrors = await addPivotFields(context, pivot, properties);
          pivot.load('id,name'); await context.sync();
          result.id = pivot.id; result.name = pivot.name;
          if (fieldErrors.length) { result.success = false; result.errors = fieldErrors; }
        } else {
          if (!objectId) throw new Error('PivotTable update/delete requires id');
          const pivot = sheet.pivotTables.getItem(objectId);
          if (operation === 'delete') pivot.delete();
          if (operation === 'update') {
            const fieldErrors = await addPivotFields(context, pivot, properties);
            if (fieldErrors.length) { result.success = false; result.errors = fieldErrors; }
          }
        }
      }
      await context.sync();
      result._dirtyRanges = [{ sheetId, range: properties.range || properties.anchor || '*' }];
      return result;
    });
  }
  async function addPivotFields(context, pivot, p) {
    const errors = [];
    for (const x of (p.rows || [])) {
      try {
        pivot.rowHierarchies.add(pivot.hierarchies.getItem(x.field));
        await context.sync();
      } catch (e) {
        errors.push({ area: 'rows', field: x.field, error: e.message || String(e) });
      }
    }
    for (const x of (p.columns || [])) {
      try {
        pivot.columnHierarchies.add(pivot.hierarchies.getItem(x.field));
        await context.sync();
      } catch (e) {
        errors.push({ area: 'columns', field: x.field, error: e.message || String(e) });
      }
    }
    for (const x of (p.values || [])) {
      try {
        const h = pivot.hierarchies.getItem(x.field);
        const dh = pivot.dataHierarchies.add(h);
        if (x.summarizeBy) dh.summarizeBy = pivotSummarizeBy(x.summarizeBy);
        await context.sync();
      } catch (e) {
        errors.push({ area: 'values', field: x.field, summarizeBy: x.summarizeBy || null, error: e.message || String(e) });
      }
    }
    return errors;
  }
  function pivotSummarizeBy(value) {
    const map = {
      sum: Excel.AggregationFunction && Excel.AggregationFunction.sum || 'Sum',
      count: Excel.AggregationFunction && Excel.AggregationFunction.count || 'Count',
      average: Excel.AggregationFunction && Excel.AggregationFunction.average || 'Average',
      max: Excel.AggregationFunction && Excel.AggregationFunction.max || 'Max',
      min: Excel.AggregationFunction && Excel.AggregationFunction.min || 'Min'
    };
    return map[value] || value;
  }

  async function evalOfficeJs(args) {
    requireOffice();
    const code = args.code || '';
    return Excel.run(async context => {
      const AsyncFunction = Object.getPrototypeOf(async function () {}).constructor;
      const fn = new AsyncFunction('context', 'Excel', code);
      const result = await fn(context, Excel);
      return { success: true, result: result ?? null, _dirtyRanges: [{ sheetId: -1, range: '*' }] };
    });
  }

  async function executeToolByName(name, args) {
    const fn = TOOL_EXECUTORS[name];
    if (!fn) throw new Error(`Tool ${name} not found`);
    return fn(args || {});
  }
  async function maybeFollow(result) {
    if (!App.state.settings.followMode || !result) return;
    const dirty = result._dirtyRanges;
    if (!Array.isArray(dirty) || !dirty.length) return;
    const first = dirty.find(x => x.sheetId && x.sheetId > 0);
    if (first) await selectRange(first.sheetId, first.range === '*' ? undefined : first.range).catch(console.warn);
  }
  async function selectRange(sheetId, range) {
    requireOffice();
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      sheet.activate();
      if (range) sheet.getRange(range).select();
      await context.sync();
      return { success: true };
    });
  }
  async function navigateCitation(ref) {
    const [sid, range] = ref.split('!');
    return selectRange(Number(sid), range);
  }

  const TOOL_EXECUTORS = {
    get_cell_ranges: getCellRanges,
    get_range_as_csv: getRangeAsCsv,
    search_data: searchData,
    get_all_objects: getAllObjects,
    set_cell_range: setCellRange,
    clear_cell_range: clearCellRange,
    copy_to: copyTo,
    modify_sheet_structure: modifySheetStructure,
    modify_workbook_structure: modifyWorkbookStructure,
    resize_range: resizeRange,
    modify_object: modifyObject,
    eval_officejs: evalOfficeJs
  };

  App.requireOffice = requireOffice;
  App.getWorkbookId = getWorkbookId;
  App.getWorkbookMetadata = getWorkbookMetadata;
  App.executeToolByName = executeToolByName;
  App.maybeFollow = maybeFollow;
  App.navigateCitation = navigateCitation;
  App.TOOL_EXECUTORS = TOOL_EXECUTORS;
})();
