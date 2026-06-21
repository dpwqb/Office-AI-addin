(function () {
  'use strict';
  const App = (window.App = window.App || {});

  const requireOffice = () => App.requireOffice();

  // ---- Excel 稳定 sheetId 映射 ----
  async function getSheetMap(context, sheets) {
    let map = {};
    try { map = JSON.parse(App.loadDocSetting(App.STORAGE_KEYS.sheetMap, '{}') || '{}'); } catch { map = {}; }
    let max = Object.values(map).reduce((a, b) => Math.max(a, Number(b) || 0), 0);
    let dirty = false;
    for (const s of sheets) {
      if (!map[s.id]) { map[s.id] = ++max; dirty = true; }
    }
    if (dirty) await App.saveDocSetting(App.STORAGE_KEYS.sheetMap, JSON.stringify(map)).catch(() => {});
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
    const { sheetId, operation, dimension = 'rows', reference, count, position = 'before' } = args;
    return Excel.run(async context => {
      const sheet = await worksheetById(context, sheetId);
      if (!sheet) throw new Error(`Worksheet with ID ${sheetId} not found`);
      if (operation === 'unfreeze') sheet.freezePanes.unfreeze();
      else if (operation === 'freeze') {
        let n;
        if (count != null && count !== '') {
          n = Number(count);
        } else if (reference != null && reference !== '') {
          n = dimension === 'columns'
            ? String(reference).toUpperCase().split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0)
            : Number(reference);
        } else {
          n = 1;
        }
        if (!Number.isFinite(n) || n < 1) n = 1;
        dimension === 'columns' ? sheet.freezePanes.freezeColumns(n) : sheet.freezePanes.freezeRows(n);
      }
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

  const SYSTEM_PROMPT = `You are an AI assistant integrated into Microsoft Excel with full access to read and modify spreadsheet data.

Available tools:
READ:
- get_cell_ranges: Read cell values, formulas, and formatting
- get_range_as_csv: Get data as CSV, useful for analysis
- search_data: Find text across the spreadsheet
- get_all_objects: List charts, pivot tables, and other objects

WRITE:
- set_cell_range: Write values, formulas, notes, and formatting
- clear_cell_range: Clear contents or formatting
- copy_to: Copy ranges with formula translation
- modify_sheet_structure: Insert/delete/hide/unhide rows/columns, freeze panes
- modify_workbook_structure: Create/delete/rename/duplicate sheets
- resize_range: Adjust column widths and row heights
- modify_object: Create/update/delete charts and pivot tables
- eval_officejs: Execute Office.js code when the listed tools are not enough

Citations: Use markdown links with #cite: hash to reference sheets/cells. Clicking navigates there.
- Sheet only: [Sheet Name](#cite:sheetId)
- Cell/range: [A1:B10](#cite:sheetId!A1:B10)
Example: [Exchange Ratio](#cite:3) or [see cell B5](#cite:3!B5)

When the user asks about their workbook data, read it first. Be concise. Use A1 notation for cell references. Before overwriting existing data, confirm unless the user explicitly asks to replace or overwrite.`;

  const TOOL_DEFINITIONS = [
    { type: 'function', function: { name: 'get_cell_ranges', description: 'Read cell values, formulas, and formatting from specified ranges in a worksheet. Returns cells as a sparse object with A1-notation keys.', parameters: { type: 'object', properties: { sheetId: { type: 'number', description: 'The worksheet ID (1-based index)' }, ranges: { type: 'array', items: { type: 'string' }, description: "Array of ranges in A1 notation, e.g. ['A1:C10']" }, includeStyles: { type: 'boolean', description: 'Include font/fill styling info. Default true' }, cellLimit: { type: 'number', description: 'Maximum cells to return. Default 2000' }, explanation: { type: 'string' } }, required: ['sheetId', 'ranges'] } } },
    { type: 'function', function: { name: 'get_range_as_csv', description: 'Read cell data from a range and return it as CSV format. Great for analysis.', parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, includeHeaders: { type: 'boolean' }, maxRows: { type: 'number' }, explanation: { type: 'string' } }, required: ['sheetId', 'range'] } } },
    { type: 'function', function: { name: 'search_data', description: 'Find text or values across the spreadsheet. Supports regex and case-sensitive search.', parameters: { type: 'object', properties: { searchTerm: { type: 'string' }, sheetId: { type: 'number' }, range: { type: 'string' }, offset: { type: 'number' }, options: { type: 'object', properties: { matchCase: { type: 'boolean' }, matchEntireCell: { type: 'boolean' }, matchFormulas: { type: 'boolean' }, useRegex: { type: 'boolean' }, maxResults: { type: 'number' } } }, explanation: { type: 'string' } }, required: ['searchTerm'] } } },
    { type: 'function', function: { name: 'get_all_objects', description: 'List all charts, pivot tables, and other objects in the workbook.', parameters: { type: 'object', properties: { sheetId: { type: 'number' }, id: { type: 'string' }, explanation: { type: 'string' } } } } },
    { type: 'function', function: { name: 'set_cell_range', description: "WRITE. Write values, formulas, notes, and formatting to cells. By default fails if target cells contain data. Retry with allow_overwrite=true after confirmation.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, cells: { type: 'array', items: { type: 'array', items: { type: 'object', properties: { value: {}, formula: { type: 'string' }, note: { type: 'string' }, cellStyles: { type: 'object' }, borderStyles: { type: 'object' } } } } }, copyToRange: { type: 'string' }, resizeWidth: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, resizeHeight: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, allow_overwrite: { type: 'boolean' }, explanation: { type: 'string' } }, required: ['sheetId', 'range', 'cells'] } } },
    { type: 'function', function: { name: 'clear_cell_range', description: "Clear contents, formatting, or both from a range. clearType: contents/formats/all.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, clearType: { enum: ['contents', 'formats', 'all'] }, explanation: { type: 'string' } }, required: ['sheetId', 'range'] } } },
    { type: 'function', function: { name: 'copy_to', description: 'Copy a range to another location with formula translation. If destination is larger, source pattern repeats.', parameters: { type: 'object', properties: { sheetId: { type: 'number' }, sourceRange: { type: 'string' }, destinationRange: { type: 'string' }, explanation: { type: 'string' } }, required: ['sheetId', 'sourceRange', 'destinationRange'] } } },
    { type: 'function', function: { name: 'modify_sheet_structure', description: "Insert, delete, hide, unhide, or freeze rows and columns. For insert/delete/hide/unhide, use reference like '5' (row) or 'C' (column). For freeze, prefer count = number of rows/columns to freeze; alternatively reference can specify the boundary (e.g. reference='C' freezes the first 3 columns, reference='5' freezes the first 5 rows).", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, operation: { enum: ['insert', 'delete', 'hide', 'unhide', 'freeze', 'unfreeze'] }, dimension: { enum: ['rows', 'columns'] }, reference: { type: 'string' }, count: { type: 'number' }, position: { enum: ['before', 'after'] }, explanation: { type: 'string' } }, required: ['sheetId', 'operation', 'dimension'] } } },
    { type: 'function', function: { name: 'modify_workbook_structure', description: 'Create, delete, rename, or duplicate worksheets.', parameters: { type: 'object', properties: { operation: { enum: ['create', 'delete', 'rename', 'duplicate'] }, sheetId: { type: 'number' }, sheetName: { type: 'string' }, newName: { type: 'string' }, tabColor: { type: 'string' }, explanation: { type: 'string' } }, required: ['operation'] } } },
    { type: 'function', function: { name: 'resize_range', description: "Adjust column widths or row heights. Use 'A:D' for columns, '1:5' for rows, or omit range for entire sheet.", parameters: { type: 'object', properties: { sheetId: { type: 'number' }, range: { type: 'string' }, width: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, height: { type: 'object', properties: { type: { enum: ['points', 'standard'] }, value: { type: 'number' } } }, explanation: { type: 'string' } }, required: ['sheetId'] } } },
    { type: 'function', function: { name: 'modify_object', description: 'Create, update, or delete charts and pivot tables.', parameters: { type: 'object', properties: { operation: { enum: ['create', 'update', 'delete'] }, sheetId: { type: 'number' }, objectType: { enum: ['pivotTable', 'chart'] }, id: { type: 'string' }, properties: { type: 'object', properties: { name: { type: 'string' }, source: { type: 'string' }, range: { type: 'string' }, anchor: { type: 'string' }, rows: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' } } } }, columns: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' } } } }, values: { type: 'array', items: { type: 'object', properties: { field: { type: 'string' }, summarizeBy: { enum: ['sum', 'count', 'average', 'max', 'min'] } } } }, title: { type: 'string' }, chartType: { enum: ['columnClustered', 'barClustered', 'line', 'pie', 'scatter', 'area', 'doughnut'] } } }, explanation: { type: 'string' } }, required: ['operation', 'sheetId', 'objectType'] } } },
    { type: 'function', function: { name: 'eval_officejs', description: 'Execute arbitrary Office.js code in Excel.run. Escape hatch for unsupported operations. Code receives context and Excel.', parameters: { type: 'object', properties: { code: { type: 'string' }, explanation: { type: 'string' } }, required: ['code'] } } }
  ];

  const SAMPLE_ARGS = {
    get_cell_ranges: { sheetId: 1, ranges: ['A1:D10'], includeStyles: true, cellLimit: 2000 },
    get_range_as_csv: { sheetId: 1, range: 'A1:D10', includeHeaders: true, maxRows: 500 },
    search_data: { searchTerm: 'keyword', options: { matchCase: false, useRegex: false, maxResults: 100 } },
    get_all_objects: {},
    set_cell_range: { sheetId: 1, range: 'A1:B2', cells: [[{ value: '标题1', cellStyles: { fontWeight: 'bold' } }, { value: '标题2', cellStyles: { fontWeight: 'bold' } }], [{ value: 1 }, { formula: '=A2*2' }]], allow_overwrite: false },
    clear_cell_range: { sheetId: 1, range: 'A1:B2', clearType: 'contents' },
    copy_to: { sheetId: 1, sourceRange: 'A1:B2', destinationRange: 'D1:E2' },
    modify_sheet_structure: { sheetId: 1, operation: 'insert', dimension: 'rows', reference: '5', count: 1, position: 'before' },
    modify_workbook_structure: { operation: 'create', sheetName: 'AI分析结果', tabColor: '#134cff' },
    resize_range: { sheetId: 1, range: 'A:D', width: { type: 'points', value: 90 } },
    modify_object: { operation: 'create', sheetId: 1, objectType: 'chart', properties: { source: 'A1:B10', chartType: 'columnClustered', anchor: 'E2', title: 'Chart' } },
    eval_officejs: { code: "const range = context.workbook.worksheets.getActiveWorksheet().getRange('A1');\nrange.load('values');\nawait context.sync();\nreturn range.values;" }
  };
  function defaultArgsForTool(name) { return App.pretty(SAMPLE_ARGS[name] || {}); }

  App.HOSTS.excel = {
    hostType: 'excel',
    available: true,
    metadataLabel: 'Workbook metadata',
    systemPrompt: SYSTEM_PROMPT,
    toolDefinitions: TOOL_DEFINITIONS,
    toolExecutors: TOOL_EXECUTORS,
    defaultArgsForTool,
    evalToolName: 'eval_officejs',
    getMetadata: getWorkbookMetadata,
    navigateCitation,
    follow: maybeFollow,
    i18n: {
      zh: {
        brand: 'dpoqb in Excel', brandFooter: 'dpoqb in Excel · Plain Edition',
        title: '准备好处理你的 Excel 数据', subtitle: '你可以让我分析、可视化或转换你的数据',
        input: '告诉我你想如何处理这份表格…',
        chart: '智能图表生成', chartDesc: '自动推荐图表类型并一键生成',
        fix: '公式错误诊断', fixDesc: '定位报错原因并生成修复公式',
        analyze: '跨表智能解析', analyzeDesc: '自动关联多表数据并输出结论',
        chartPrompt: '请根据当前表格数据结构，生成合适的可视化图表',
        fixPrompt: '帮我检查当前表格中的错误内容，定位问题并给出修复后的正确结果',
        analyzePrompt: '帮我全面分析表中所有内容，并输出汇总结果与关键分析结论',
        demo: '当前不在 Excel/Office 环境中，Excel 工具只能在插件侧边栏里运行。',
        confirmEval: 'AI 请求执行以下 Office.js 代码，是否允许？'
      },
      en: {
        brand: 'dpoqb in Excel', brandFooter: 'dpoqb in Excel · Plain Edition',
        title: 'Ready to work with your Excel data', subtitle: 'Ask me to analyze, visualize, or transform your data',
        input: 'Tell me what to do with this workbook…',
        chart: 'Chart Generation', chartDesc: 'One-click visualization & styling',
        fix: 'Error Fix', fixDesc: 'Auto-detect & fix formula errors',
        analyze: 'Multi-Sheet Analysis', analyzeDesc: 'Cross-sheet automation and conclusions',
        chartPrompt: 'Generate charts to visualize my data and apply professional styling',
        fixPrompt: 'Check my spreadsheet for formula errors and fix them automatically',
        analyzePrompt: 'Analyze all workbook contents and summarize key conclusions',
        demo: 'Not currently running inside Excel/Office. Excel tools only work in the add-in task pane.',
        confirmEval: 'The AI requests to run the following Office.js code. Allow it?'
      }
    }
  };
})();
