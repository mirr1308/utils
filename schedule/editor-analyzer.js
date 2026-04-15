/**
 * editor-analyzer.js  
 */

const SampleCache = {
    _domDocument:  null,
    _htmlString: null,

    _loadFromStorage() {
        if (!this._htmlString) {
            const savedHtml = AppStore.get('analysis_source_save');
            if (savedHtml) {
                this._htmlString = savedHtml;
                this._domDocument  = DomManager.parse(savedHtml);
                return true;
            }
        }
        return !!this._htmlString;
    },
    set(html) {
        const trimmedHtml = html?.trim() || '';
        if (!trimmedHtml) {
            this._htmlString = null;
            this._domDocument  = null;
            AppStore.remove('analysis_source_save');
            return;
        }
        if (trimmedHtml === this._htmlString) return;
        this._htmlString = trimmedHtml;
        this._domDocument  = DomManager.parse(trimmedHtml);
        AppStore.set('analysis_source_save', trimmedHtml);
    },
    getHtml() {
        this._loadFromStorage();
        return this._htmlString || '';
    },
    refreshUI() {
        const html  = this.getHtml();
        const inputElement = document.getElementById('analysisInput');
        if (inputElement) inputElement.value = html;
        if (window.analysisEditor) window.analysisEditor.setValue(html);
    },
    getTemplateTable() {
        const inputElement = document.getElementById('analysisInput');
        if (inputElement) {
            const inputValue = inputElement.value.trim();
            if (inputValue && inputValue !== (this._htmlString || '')) this.set(inputValue);
        }
        if (!this._domDocument) this._loadFromStorage();
        if (!this._domDocument) return null;
        const tableElement = this._domDocument.querySelector('table');
        return tableElement ? DomManager.clone(tableElement) : null;
    },
    init() {
        this._loadFromStorage();
        this.refreshUI();
    }
};
window.SampleCache = SampleCache;

const SANITIZE_CONFIG = {
    pairedTags:    ['script', 'iframe', 'object', 'svg'],   
    selfClosingTags: ['embed', 'link', 'meta', 'base'],  
};

const CSS_PROPS = {
    RADIUS: [
        'border-radius',
        'border-top-left-radius',    'border-top-right-radius',
        'border-bottom-left-radius', 'border-bottom-right-radius',
    ],
    TOP_RADIUS: [
        'border-top-left-radius', 'border-top-right-radius',
    ],
    BOTTOM_RADIUS: [
        'border-bottom-left-radius', 'border-bottom-right-radius',
    ],
};

const StyleUtils = {
    hexStyle(styleString) {
        return ColorManager.restoreColors(styleString);
    },
    toHex(color) {
		if (!color) return '';
		return ColorManager.toOriginalForm(color); 
	},

    getCleanStyle(element) {
        if (!element) return '';
        return (element.getAttribute('style') || '').replace(/cursor:[^;]+;?/g, '').trim();
    },

    replaceStyleProp(styleString, property, newValue) {
        styleString = (styleString || '').trim();
        const propRegex = new RegExp(`(^|;)\\s*${property}\\s*:[^;]+;?`, 'gi');
        if (!newValue) {
            return styleString.replace(propRegex, '$1').replace(/;+/g, ';').replace(/^;|;$/g, '');
        }
        if (propRegex.test(styleString)) {
            return styleString.replace(propRegex, `$1${property}:${newValue};`).replace(/;+/g, ';').replace(/^;|;$/g, '');
        }
        return (styleString.replace(/;?$/, '') + `;${property}:${newValue};`).replace(/^;+/, '');
    },
    _removeStyles(styleString, propsArray) {
        if (!styleString) return '';
        const obj = ColorManager.parseStyleString(styleString);
        propsArray.forEach(prop => delete obj[prop]);
        return ColorManager.serializeStyle(obj);
    },
    removeRadius(styleString)       { return this._removeStyles(styleString, CSS_PROPS.RADIUS); },
    removeTopRadius(styleString)    { return this._removeStyles(styleString, CSS_PROPS.TOP_RADIUS); },
    removeBottomRadius(styleString) { return this._removeStyles(styleString, CSS_PROPS.BOTTOM_RADIUS); },

    applyBg(styleString, backgroundColor) {
        const obj = ColorManager.parseStyleString(styleString);
        if (!backgroundColor) {
            delete obj['background'];
            delete obj['background-color'];
        } else {
            obj['background-color'] = backgroundColor;
        }
        return ColorManager.serializeStyle(obj);
    },

    extractBgColor(element) {
        if (!element) return '';
        const styleString = element.getAttribute('style') || '';
        const bgMatch = styleString.match(/background(?:-color)?\s*:\s*([^;]+)/i);
        return bgMatch ? ColorManager.toOriginalForm(bgMatch[1].trim()) : '';
    },
    parseColorInput(rawColorString) {
        return (rawColorString || '').split(',').map(p => p.trim()).filter(Boolean);
    },
    getColorForDay(colorArray, dayOfWeek) {
        const dayGroup = DayManager.getGroupType(dayOfWeek);
        if (dayGroup === DayManager.DAY_IDX.SUN) return colorArray[2];
        if (dayGroup === DayManager.DAY_IDX.SAT) return colorArray[1];
        return colorArray[0] || null;
    },
    wrapTextNodesWithColor(parentNode, hexColor) {
        const walker = document.createTreeWalker(parentNode, NodeFilter.SHOW_TEXT, {
            acceptNode: (node) => node.nodeValue.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP,
        });
        const textNodes = [];
        let node;
        while ((node = walker.nextNode())) textNodes.push(node);
        textNodes.forEach(textNode => {
            const span = document.createElement('span');
            span.setAttribute('style', `color:${hexColor}`);
            textNode.parentNode.insertBefore(span, textNode);
            span.appendChild(textNode);
        });
    },

    applyColorToDateCell(cell, hexColor) {
        if (!hexColor) return;
        const cellStyle       = cell.getAttribute('style') || '';
        const hasColorInStyle = /(?:^|;)\s*color\s*:/i.test(cellStyle);
        if (hasColorInStyle) {
            cell.setAttribute('style', cellStyle.replace(/((?:^|;)\s*color\s*:)[^;]*/i, `$1${hexColor}`).trim());
            return;
        }
        const coloredChildren = cell.querySelectorAll('p[style*="color"], span[style*="color"]');
        if (coloredChildren.length > 0) {
            coloredChildren.forEach(el => {
                el.setAttribute('style', this.replaceStyleProp(el.getAttribute('style') || '', 'color', hexColor));
            });
            return;
        }
        this.wrapTextNodesWithColor(cell, hexColor);
    },

    applyBottomBorder(styleString, borderRule) {
        if (!borderRule) return styleString;
        const borderValue = borderRule.includes(':') ? borderRule.split(':')[1].trim() : borderRule;
        return this.replaceStyleProp(styleString, 'border-bottom', borderValue);
    },

    extractBottomBorder(tableRowElements) {
        if (!tableRowElements?.length) return '';
        const joined       = tableRowElements.map(row => row.getAttribute('style') || '').join(';');
        const solidMatch   = joined.match(/border-bottom\s*:\s*[^;]*solid[^;]*/i);
        const lastRowMatch = tableRowElements.at(-1).getAttribute('style')?.match(/border-bottom\s*:[^;]+/i);
        return solidMatch?.[0] || lastRowMatch?.[0] || '';
    },

    parseBorderSides(element) {
        if (!element) return { top: '', right: '', bottom: '', left: '' };
        const s = element.style;
        const build = (w, st, c) => [w, st, c].filter(Boolean).join(' ') || s.border || '';
        return {
            top:    s.borderTop    || build(s.borderTopWidth,    s.borderTopStyle,    s.borderTopColor),
            right:  s.borderRight  || build(s.borderRightWidth,  s.borderRightStyle,  s.borderRightColor),
            bottom: s.borderBottom || build(s.borderBottomWidth, s.borderBottomStyle, s.borderBottomColor),
            left:   s.borderLeft   || build(s.borderLeftWidth,   s.borderLeftStyle,   s.borderLeftColor),
        };
    },
    parseBorderRadius(element) {
        if (!element) return { topLeft: '', topRight: '', bottomRight: '', bottomLeft: '' };
        const s = element.style;
        return {
            topLeft:     s.borderTopLeftRadius     || s.borderRadius || '',
            topRight:    s.borderTopRightRadius    || s.borderRadius || '',
            bottomRight: s.borderBottomRightRadius || s.borderRadius || '',
            bottomLeft:  s.borderBottomLeftRadius  || s.borderRadius || '',
        };
    },

    applyCellStyle(targetCell, firstRowCell, lastRowCell) {
        const firstBorder = this.parseBorderSides(firstRowCell);
        const lastBorder  = this.parseBorderSides(lastRowCell);
        targetCell.style.border = '';
        if (firstBorder.top)   targetCell.style.borderTop   = firstBorder.top;
        if (firstBorder.left)  targetCell.style.borderLeft  = firstBorder.left;
        if (firstBorder.right) targetCell.style.borderRight = firstBorder.right;
        if (lastBorder.bottom) targetCell.style.borderBottom = lastBorder.bottom;

        const firstRadius = this.parseBorderRadius(firstRowCell);
        const lastRadius  = this.parseBorderRadius(lastRowCell);
        targetCell.style.borderRadius = '';
        if (firstRadius.topLeft)     targetCell.style.borderTopLeftRadius     = firstRadius.topLeft;
        if (firstRadius.topRight)    targetCell.style.borderTopRightRadius    = firstRadius.topRight;
        if (lastRadius.bottomRight)  targetCell.style.borderBottomRightRadius = lastRadius.bottomRight;
        if (lastRadius.bottomLeft)   targetCell.style.borderBottomLeftRadius  = lastRadius.bottomLeft;

        const boxShadow = firstRowCell?.style.boxShadow || lastRowCell?.style.boxShadow || '';
        const outline   = firstRowCell?.style.outline   || '';
        if (boxShadow) targetCell.style.boxShadow = boxShadow;
        if (outline)   targetCell.style.outline   = outline;
    },

    cloneStructuralTd(sourceCell, borderBottomOverride = null) {
        const cloned = sourceCell.cloneNode(false);
        if (sourceCell.colSpan > 1) cloned.colSpan = sourceCell.colSpan;
        ['align', 'width', 'valign'].forEach(attr => {
            if (sourceCell.hasAttribute(attr)) cloned.setAttribute(attr, sourceCell.getAttribute(attr));
        });
        this.applyCellStyle(cloned, sourceCell, sourceCell);
        if (borderBottomOverride !== null) cloned.style.borderBottom = borderBottomOverride;
        return cloned;
    },

    applyRowCellStyles(targetRow, sourceRow, hexBackground, { radiusMode = 'none', skipFirstCell = false } = {}) {
        Array.from(targetRow.cells).forEach((cell, i) => {
            if (skipFirstCell && i === 0) return;
            const sourceCell = sourceRow.cells[i];
            if (!sourceCell) return;
            let styleString = this.hexStyle(sourceCell.getAttribute('style') || '');
            if      (radiusMode === 'all')    styleString = this.removeRadius(styleString);
            else if (radiusMode === 'bottom') styleString = this.removeBottomRadius(styleString);
            else if (radiusMode === 'top')    styleString = this.removeTopRadius(styleString);
            cell.setAttribute('style', this.hexStyle(this.applyBg(styleString, hexBackground)));
        });
    },

    applyRowTrStyle(targetRow, sourceRow, hexBackground) {
        const rowStyle = this.hexStyle(sourceRow?.getAttribute('style') || '');
        targetRow.setAttribute('style', this.hexStyle(this.applyBg(rowStyle, hexBackground)));
    },

    cloneRowWithEmptyCells(sourceRow, hexBackground, { skipDateCell = false } = {}) {
        if (!sourceRow) return document.createElement('tr');
        const newRow = sourceRow.cloneNode(false);
        this.applyRowTrStyle(newRow, sourceRow, hexBackground);
        Array.from(sourceRow.cells).forEach(sourceCell => {
            if (skipDateCell && (sourceCell.hasAttribute('rowspan') || sourceCell.id?.includes(CONSTANTS.USER_CONTENT_PREFIX))) return;
            newRow.appendChild(TableUtils.createEmptyCell(sourceCell, hexBackground, { removeRadius: true }));
        });
        return newRow;
    },
};
window.StyleUtils = StyleUtils;

const TableUtils = {
    getTable(element) {
        if (!element) return null;
        return element.closest('table');
    },
    getTbody(tableElement) {
        if (!tableElement) return null;
        return tableElement.querySelector(':scope > tbody') || tableElement;
    },
    getRows(tbodyOrTable) {
        if (!tbodyOrTable) return [];
        return Array.from(tbodyOrTable.querySelectorAll(':scope > tr')).filter(row => row.cells.length > 0);
    },
    createEmptyCell(sourceCell, backgroundHex = null, options = { removeRadius: false }) {
        if (!sourceCell) return document.createElement('td');
        const clonedCell = sourceCell.cloneNode(false);
        clonedCell.removeAttribute('rowspan');
        let resolvedBackground = backgroundHex;
        if (!resolvedBackground && sourceCell.hasAttribute('bgcolor')) {
            const rawBgColor = sourceCell.getAttribute('bgcolor');
            resolvedBackground = ColorManager.toOriginalForm(rawBgColor);
        }
        let styleString = sourceCell.getAttribute('style') || '';
        if (resolvedBackground) {
            styleString = StyleUtils.applyBg(styleString, resolvedBackground);
        }
        if (options.removeRadius) {
            styleString = StyleUtils.removeRadius(styleString);
        }
        clonedCell.setAttribute('style', ColorManager.restoreColors(styleString));
        clonedCell.removeAttribute('bgcolor'); 
        clonedCell.innerHTML = '&nbsp;';
        return clonedCell;
    },
    isDateCell(tableCell) {
        if (!tableCell) return false;
        const rowspan     = parseInt(tableCell.getAttribute('rowspan')) || 1;
        const hasIdPrefix  = tableCell.id?.includes(CONSTANTS.USER_CONTENT_PREFIX);
        const hasDateNumber = /\d/.test(tableCell.textContent.trim());
        return rowspan >= 2 || hasIdPrefix || hasDateNumber;
    }
};

const DateUtils = {
    padNumberToMatch(number, referenceString) {
        const reference = (referenceString || '').trim();
        const paddingLength = (reference.startsWith('0') && /\d/.test(reference)) ? reference.length : 0;
        return paddingLength > 0 ? String(number).padStart(paddingLength, '0') : String(number);
    },
    generateCellId(baseId, dateNumber, referenceString, fallbackPrefix = '') {
        const paddedNumber = this.padNumberToMatch(dateNumber, referenceString);
        if (baseId) {
            return baseId.replace(/\d+$/, '') + paddedNumber;
        } else if (fallbackPrefix) {
            return `${CONSTANTS.USER_CONTENT_PREFIX}${fallbackPrefix}${dateNumber}`;
        }
        return '';
    },
};

function getRowGroups(rows) {
    const groups = [];
    let rowIndex = 0;
    while (rowIndex < rows.length) {
        const rowspanValue = parseInt(rows[rowIndex].cells[0]?.getAttribute('rowspan')) || 1;
        groups.push(rows.slice(rowIndex, rowIndex + rowspanValue));
        rowIndex += rowspanValue;
    }
    return groups;
}
window.getRowGroups = getRowGroups;

function findDataStartIdx(rows) {
    for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
        const firstCell = rows[rowIndex].cells[0];
        if (!firstCell) continue;
        if (TableUtils.isDateCell(firstCell)) return rowIndex;
    }
    return 0;
}
window.findDataStartIdx = findDataStartIdx;

window.parseCellDate = function (cell) {
    if (!cell) return null;
    const textContent = cell.textContent.trim();
    const dateMatch = textContent.match(/\d+/);
    if (!dateMatch) return null;
    const dateNumber = parseInt(dateMatch[0]);
    const dayString  = (typeof DayManager !== 'undefined') ? DayManager.getDayStr(dateNumber) : '';
    return { date: dateNumber, day: dayString };
};

function focusCellInPreview(targetCell, markerType = 'new') {
    const preview = EditorState.get('preview');
    const editor  = EditorState.get('editor');
    if (!targetCell) return;
    const outerTable = TableUtils.getTable(targetCell) || preview.querySelector('table');
    const outerTbody = TableUtils.getTbody(outerTable);
    const allOuterCells = [];
    if (outerTbody) {
        TableUtils.getRows(outerTbody).forEach(tableRow => {
            allOuterCells.push(...Array.from(tableRow.cells));
        });
    }

    let liveCell = preview.contains(targetCell) ? targetCell : (allOuterCells.at(-1) || null);
    if (!liveCell) return;

    if (liveCell.getAttribute('contenteditable') !== 'true') makeEditableOnlyCells(preview);
    liveCell.focus();

    const selectionRange = document.createRange();
    selectionRange.selectNodeContents(liveCell);
    selectionRange.collapse(true);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(selectionRange);

    EditorState.set('savedRange', selectionRange.cloneRange());
    EditorState.currentTargetNode = liveCell;

    const allEditorCells = Array.from(preview.querySelectorAll('td, th'));
    const cellIndexInEditor = allEditorCells.indexOf(liveCell);
    if (cellIndexInEditor === -1) return;

    const editorContent = editor.getValue();
    const cellTagRegex = /<(?:td|th)\b/gi;
    let regexMatch, cellCount = 0, targetLineNumber = -1;
    while ((regexMatch = cellTagRegex.exec(editorContent)) !== null) {
        if (cellCount === cellIndexInEditor) {
            targetLineNumber = editorContent.substring(0, regexMatch.index).split('\n').length - 1;
            break;
        }
        cellCount++;
    }
    if (targetLineNumber === -1) return;

    editor.clearGutter('markers');
    const markerElement = document.createElement('div');
    markerElement.className = `working-marker working-marker--${markerType}`;
    markerElement.innerHTML = '●';
    editor.setGutterMarker(targetLineNumber, 'markers', markerElement);
    editor.scrollIntoView({ line: targetLineNumber, ch: 0 }, 200);

    if (window._headerLockRange && typeof window.applyHeaderLock === 'function') {
        requestAnimationFrame(() => window.applyHeaderLock());
    }
}

window.syncTableToEditor = function (tableElement, lockHeaderLines = false) {
    const editor  = EditorState.get('editor');
    const preview = EditorState.get('preview');
    if (!tableElement) return;
	const tempTable = DomManager.clone(tableElement);
	tempTable.querySelectorAll('td, th, [contenteditable]').forEach(el => DomManager.clean(el));
	DomManager.clean(tempTable);
	const html = StyleUtils.hexStyle(tempTable.outerHTML);
    withSyncLock(() => {
        if (typeof window.clearHeaderLock === 'function') window.clearHeaderLock();
        const beautifiedHtml = safeBeautify(html);
        editor.setValue(beautifiedHtml);
        if (editor.refresh) editor.refresh();
    });

    if (typeof window.patchPreview === 'function') window.patchPreview(html);
    else preview.innerHTML = html;

    makeEditableOnlyCells(preview);

    if (lockHeaderLines || window._headerLockRange) {
        setTimeout(() => window.applyHeaderLock(), 0);
    }
};

window.isCalendarTable = function () {
    const editor = EditorState.get('editor');
    if (!editor) return false;
    const parsedDoc   = DomManager.parse(editor.getValue());
    if (!parsedDoc) return false;
    const table = parsedDoc.querySelector('table');
    if (!table) return false;
    const firstRow = table.querySelector('tr');
    if (!firstRow) return false;
    const totalColumns = Array.from(firstRow.cells)
        .reduce((sum, cell) => sum + (parseInt(cell.getAttribute('colspan')) || 1), 0);
    return totalColumns === 7;
};

window.applyHeaderLock = function () {
    const editor = EditorState.get('editor');
    if (!editor) return;
    if (window.isCalendarTable()) { window.releaseHeaderLock(); return; }
    window.clearHeaderLock();

    const allLines      = editor.getValue().split('\n');
    const totalLineCount = allLines.length;

    const tableOpenLine  = allLines.findIndex(line => /<table[\s>]/i.test(line));
    const tbodyOpenLine  = allLines.findIndex(line => /<tbody[\s>]/i.test(line));
    let tbodyCloseLine = -1, tableCloseLine = -1;
    for (let lineIndex = totalLineCount - 1; lineIndex >= 0; lineIndex--) {
        if (tableCloseLine === -1 && /<\/table>/i.test(allLines[lineIndex])) tableCloseLine = lineIndex;
        if (tbodyCloseLine === -1 && /<\/tbody>/i.test(allLines[lineIndex]))  tbodyCloseLine = lineIndex;
        if (tableCloseLine !== -1 && tbodyCloseLine !== -1) break;
    }
    if (tableOpenLine < 0 || tbodyOpenLine <= tableOpenLine) return;

    const headerEndLine   = tbodyOpenLine;
    const footerStartLine = tbodyCloseLine >= 0 ? tbodyCloseLine : tableCloseLine;

    window._headerLockedLines = [];
    for (let lineIndex = tableOpenLine; lineIndex <= headerEndLine; lineIndex++) {
        editor.addLineClass(lineIndex, 'text', 'cm-header-locked');
        window._headerLockedLines.push(lineIndex);
    }
    if (footerStartLine > headerEndLine) {
        const footerEndLine = tableCloseLine >= 0 ? tableCloseLine : footerStartLine;
        for (let lineIndex = footerStartLine; lineIndex <= footerEndLine; lineIndex++) {
            editor.addLineClass(lineIndex, 'text', 'cm-header-locked');
            window._headerLockedLines.push(lineIndex);
        }
    }
    window._headerLockRange = {
        trStart: tbodyOpenLine + 1,
        trEnd:   footerStartLine > 0 ? footerStartLine - 1 : totalLineCount - 1,
    };
};

window.clearHeaderLock = function () {
    const editor = EditorState.get('editor');
    if (window._headerLockedLines) {
        window._headerLockedLines.forEach(lineIndex => {
            try { editor.removeLineClass(lineIndex, 'text', 'cm-header-locked'); } catch (_) {}
        });
        window._headerLockedLines = [];
    }
};

window.releaseHeaderLock = function () {
    window.clearHeaderLock();
    window._headerLockRange = null;
};

function sanitizeHtml(html) {
    let result = html;
    SANITIZE_CONFIG.pairedTags.forEach(tag => {
        result = result.replace(new RegExp(`<${tag}\\b[^<]*(?:(?!<\\/${tag}>)<[^<]*)*<\\/${tag}>`, 'gi'), '');
    });
    SANITIZE_CONFIG.selfClosingTags.forEach(tag => {
        result = result.replace(new RegExp(`<${tag}\\b[^>]*>`, 'gi'), '');
    });
    return result
        .replace(/\s+on\w+\s*=\s*["'][^"']*["']/gi, '')
        .replace(/href\s*=\s*["']\s*javascript:[^"']*/gi, 'href="#"')
        .replace(/src\s*=\s*["']\s*javascript:[^"']*/gi, 'src=""');
}

function processAnalysis() {
    const inputElement = document.getElementById('analysisInput');
    const sanitizedHtml = sanitizeHtml(inputElement?.value.trim() || '');

    if (!sanitizedHtml) {
        SampleCache.set('');
        if (inputElement) inputElement.value = '';
        window.showToast('저장된 데이터가 삭제되었습니다.', 'info');
        return;
    }
    const validationResult = DomManager.validate(sanitizedHtml, 'table');
    if (!validationResult.ok) {
        window.showToast(validationResult.reason, 'error');
        if (inputElement) inputElement.value = '';
        SampleCache.set('');
        return;
    }
    try {
        const cleanedHtml = getCleanTable(sanitizedHtml);
        if (cleanedHtml) {
            inputElement.value = cleanedHtml;
            SampleCache.set(cleanedHtml);
            window.showToast('데이터가 성공적으로 분석되었습니다.', 'success');
        } else {
            if (inputElement) inputElement.value = '';
            SampleCache.set('');
            window.showToast('유효한 테이블 구조를 찾을 수 없어 초기화되었습니다.', 'error');
        }
    } catch (_) {
        window.showToast('분석 중 오류가 발생하여 데이터를 비웁니다.', 'error');
        if (inputElement) inputElement.value = '';
        SampleCache.set('');
    }
}
window.processAnalysis = processAnalysis;

function getCleanTable(rawHtml) {
    const parsedDoc     = DomManager.parse(rawHtml);
    const sourceTable   = parsedDoc?.querySelector('table');
    if (!sourceTable) return null;

    sourceTable.querySelectorAll('td table, th table').forEach(nestedTable => nestedTable.remove());

    const cleanTable        = document.createElement('table');
    Array.from(sourceTable.attributes).forEach(attribute => cleanTable.setAttribute(attribute.name, attribute.value));
    const tableStyleAttr    = cleanTable.getAttribute('style') || '';
    if (!/border-collapse/i.test(tableStyleAttr)) {
        cleanTable.setAttribute('style', (tableStyleAttr + ';border-collapse:collapse').replace(/^;/, ''));
    }

    const captionElement = sourceTable.querySelector(':scope > caption');
    if (captionElement) cleanTable.appendChild(captionElement.cloneNode(true));
    sourceTable.querySelectorAll(':scope > colgroup').forEach(colgroup => cleanTable.appendChild(colgroup.cloneNode(true)));

    const sourceThead  = sourceTable.querySelector(':scope > thead');
    if (sourceThead) cleanTable.appendChild(sourceThead.cloneNode(true));
    const tbodyElements = Array.from(sourceTable.querySelectorAll(':scope > tbody'));

    const firstTableRow = sourceTable.querySelector('tr');
    const totalColumns  = firstTableRow
        ? Array.from(firstTableRow.cells).reduce((sum, cell) => sum + (parseInt(cell.getAttribute('colspan')) || 1), 0)
        : 3;

    function isHeaderRow(tableRow)    { return tableRow.querySelector('th') !== null; }
    function isSubHeaderRow(tableRow) {
        if (tableRow.querySelector('th')) return false;
        const cells = Array.from(tableRow.cells);
        if (!cells.length) return false;
        const totalColspan = cells.reduce((sum, cell) => sum + (parseInt(cell.getAttribute('colspan')) || 1), 0);
        return totalColspan >= totalColumns && cells.length < totalColumns;
    }
    function isDataGroupStart(tableRow) {
        return TableUtils.isDateCell(tableRow.cells[0]);
    }

    let headerRows = [], subHeaderRows = [], dataRows;

    if (sourceThead) {
        dataRows = tbodyElements.flatMap(tbody => TableUtils.getRows(tbody));
    } else {
        const allRows = tbodyElements.length > 0
            ? tbodyElements.flatMap(tbody => TableUtils.getRows(tbody))
            : TableUtils.getRows(TableUtils.getTbody(sourceTable));

        let dataStartIndex = allRows.length;
        for (let rowIndex = 0; rowIndex < allRows.length; rowIndex++) {
            if (isDataGroupStart(allRows[rowIndex])) { dataStartIndex = rowIndex; break; }
            if (isHeaderRow(allRows[rowIndex]))           headerRows.push(allRows[rowIndex]);
            else if (isSubHeaderRow(allRows[rowIndex]))   subHeaderRows.push(allRows[rowIndex]);
            else                                          headerRows.push(allRows[rowIndex]);
        }
        dataRows = allRows.slice(dataStartIndex);
    }

    const dateGroups = [];
    let remainingRowspan = 0, currentGroup = [];
    dataRows.forEach(tableRow => {
        if (remainingRowspan === 0) {
            if (currentGroup.length > 0) dateGroups.push({ rows: currentGroup });
            currentGroup = [];
            remainingRowspan = parseInt(tableRow.cells[0]?.getAttribute('rowspan')) || 1;
        }
        currentGroup.push(tableRow);
        remainingRowspan--;
    });
    if (currentGroup.length > 0) dateGroups.push({ rows: currentGroup });

    const COMPARE_STYLE_KEYS = new Set([
        'background', 'background-color', 'color',
        'border', 'border-top', 'border-right', 'border-bottom', 'border-left',
        'border-color', 'border-top-color', 'border-right-color',
        'border-bottom-color', 'border-left-color',
        'border-style', 'border-width', 'border-radius',
        'border-top-left-radius', 'border-top-right-radius',
        'border-bottom-left-radius', 'border-bottom-right-radius',
        'outline', 'outline-color',
    ]);

    function getRowStyleMap(tableRow) {
        const datumCell = tableRow.cells[0];
        const combinedStyle = [
            tableRow.getAttribute('style') || '',
            datumCell?.getAttribute('style') || '',
            datumCell?.getAttribute('bgcolor') ? `background-color:${datumCell.getAttribute('bgcolor')}` : '',
        ].join(';');
        const fullStyleMap = ColorManager.parseStyleString(combinedStyle);
        const filteredMap  = {};
        COMPARE_STYLE_KEYS.forEach(key => { if (fullStyleMap[key] !== undefined) filteredMap[key] = fullStyleMap[key]; });
        return filteredMap;
    }

    function styleMapsEqual(map1, map2) {
        const allKeys = new Set([...Object.keys(map1), ...Object.keys(map2)]);
        for (const key of allKeys) if ((map1[key] || '') !== (map2[key] || '')) return false;
        return true;
    }

    let keepGroups;
    if      (dateGroups.length === 0) keepGroups = [];
    else if (dateGroups.length === 1) keepGroups = [dateGroups[0]];
    else {
        const firstGroupStyle  = getRowStyleMap(dateGroups[0].rows[0]);
        const secondGroupStyle = getRowStyleMap(dateGroups[1].rows[0]);
        keepGroups = styleMapsEqual(firstGroupStyle, secondGroupStyle) ? [dateGroups[0]] : [dateGroups[0], dateGroups[1]];
    }

    function buildCleanRows(group) {
        const cleanRows     = [];
        const firstSourceRow = group.rows[0];
        const lastSourceRow  = group.rows[group.rows.length - 1];
        const isMultiRow     = group.rows.length > 1;
        const rowSpanValue   = isMultiRow ? 2 : 1;

        [firstSourceRow, ...(isMultiRow ? [lastSourceRow] : [])].forEach((sourceRow, orderIndex) => {
            const newRow = sourceRow.cloneNode(false);
            Array.from(sourceRow.cells).forEach((sourceCell, cellIndex) => {
                let targetCell;
                if (orderIndex === 0 && cellIndex === 0) {
                    targetCell = sourceCell.cloneNode(true);
                    targetCell.rowSpan   = rowSpanValue;
                    targetCell.innerHTML = StyleUtils.hexStyle(sourceCell.innerHTML.trim());
                    if (targetCell.hasAttribute('style')) {
                        targetCell.setAttribute('style', StyleUtils.hexStyle(targetCell.getAttribute('style')));
                    }
                } else {
                    targetCell = TableUtils.createEmptyCell(sourceCell);
                }
                newRow.appendChild(targetCell);
            });
            cleanRows.push(newRow);
        });
        return cleanRows;
    }

    if (!sourceThead && headerRows.length > 0) {
        const newThead = document.createElement('thead');
        headerRows.forEach(row => newThead.appendChild(row.cloneNode(true)));
        cleanTable.appendChild(newThead);
    }

    const finalTbody = document.createElement('tbody');
    const tbodyStyleAttr = tbodyElements[0]?.getAttribute('style');
    if (tbodyStyleAttr) finalTbody.setAttribute('style', tbodyStyleAttr);

    subHeaderRows.forEach(row => finalTbody.appendChild(row.cloneNode(true)));
    keepGroups.forEach(group => buildCleanRows(group).forEach(row => finalTbody.appendChild(row)));
    cleanTable.appendChild(finalTbody);

    const tfootElement = sourceTable.querySelector(':scope > tfoot');
    if (tfootElement) cleanTable.appendChild(tfootElement.cloneNode(true));

    const tempContainer = document.createElement('div');
    tempContainer.style.display = 'none';
    document.body.appendChild(tempContainer);
    tempContainer.appendChild(cleanTable);
    const resultHtml = tempContainer.innerHTML;
    document.body.removeChild(tempContainer);
    return resultHtml;
}

const RULE_ITEM_CONFIG = [
    { width: '25%', field: 'name', type: 'input',    placeholder: '표시 이름' },
    { width: '65%', field: 'html', type: 'textarea', placeholder: 'HTML 코드를 입력하세요' },
    { width: '10%', field: 'del',  type: 'button',   label: '×' }
];

function createRuleItemRow(item = { name: '', html: '' }) {
    const row = document.createElement('tr'); 
    const cells = RULE_ITEM_CONFIG.map(config => {
        let content = '';
        if (config.type === 'input') {
            content = `<input type="text" class="modal-input" value="${item.name || ''}" placeholder="${config.placeholder}">`;
        } else if (config.type === 'textarea') {
            content = `<textarea class="modal-input code-area" placeholder="${config.placeholder}">${item.html || ''}</textarea>`;
        } else if (config.type === 'button') {
            content = `<button class="btn-del-item" onclick="this.closest('tr').remove()">${config.label}</button>`;
        }
        return `<td style="width:${config.width}">${content}</td>`;
    });

    row.innerHTML = cells.join('');
    return row;
}

function addGroup() {
    const container = document.getElementById('ruleGroupsContainer');
    const newGroupCard  = document.createElement('div');
    newGroupCard.className = 'rule-group-card';
    newGroupCard.innerHTML = `
        <div class="group-header">
            <input type="text" class="group-name-input" placeholder="그룹 이름 (예: 카테고리 1)">
            <button class="btn-del-group" onclick="this.closest('.rule-group-card').remove()">그룹 삭제</button>
        </div>
        <table class="rule-item-table">
            <tbody class="item-list"></tbody>
        </table>
        <button class="btn-add-item-dashed" onclick="addItem(this)">+ 새 항목 추가</button>
    `;
	const itemListBody = newGroupCard.querySelector('.item-list');
    itemListBody.appendChild(createRuleItemRow());
    container.appendChild(newGroupCard);
	return { 
        card: newGroupCard, 
        itemListBody: itemListBody 
    };
}
window.addGroup = addGroup;

function addItem(button) {
    const itemListBody = button.closest('.rule-group-card').querySelector('.item-list');
    const newRow = createRuleItemRow(); 
    itemListBody.appendChild(newRow);
}
window.addItem = addItem;

function applyAndSaveRules() {
    const container  = document.getElementById('ruleGroupsContainer');
    const groupCards = container.querySelectorAll('.rule-group-card');
    const allGroups  = [];

    groupCards.forEach(card => {
        const groupName = card.querySelector('.group-name-input').value;
        const groupItems = [];
        card.querySelectorAll('.item-list tr').forEach(row => {
            const inputFields = row.querySelectorAll('input, textarea');
            if (inputFields[0].value.trim() || inputFields[1].value.trim()) {
                groupItems.push({ name: inputFields[0].value, html: inputFields[1].value });
            }
        });
        if (groupItems.length > 0 || groupName.trim()) allGroups.push({ groupName, items: groupItems });
    });

    if (allGroups.length === 0) {
        if (confirm('입력된 규칙이 없습니다. 기존 설정으로 되돌리시겠습니까?')) {
            window.renderRules?.();
            closeModal('ruleModal');
        }
        return;
    }
    AppStore.set('custom_toolbar_rules', allGroups);
    window.showToast('설정이 저장되었습니다.');
    window.updatePreview?.(true);
    closeModal('ruleModal');
}
window.applyAndSaveRules = applyAndSaveRules;

window.renderRules = function () {
    const container = document.getElementById('ruleGroupsContainer');
    const savedGroups = AppStore.get('custom_toolbar_rules');
    container.innerHTML = '';

    if (!savedGroups || savedGroups.length === 0) { 
        addGroup(); 
        return; 
    }
    savedGroups.forEach(groupData => {
        const { card, itemListBody } = addGroup();
        card.querySelector('.group-name-input').value = groupData.groupName;
        itemListBody.innerHTML = '';
        const items = groupData.items.length > 0 ? groupData.items : [{ name: '', html: '' }];      
        items.forEach(item => {
            itemListBody.appendChild(createRuleItemRow(item));
        });
    });
};

function _setCustomToolbarVisible(toolbar, visible) {
    const mainContainer = document.querySelector('.main-container');
    toolbar.style.display = visible ? 'flex' : 'none';
    if (!mainContainer) return;
    if (visible) {
        requestAnimationFrame(() => {
            const toolbarHeight = toolbar.getBoundingClientRect().height;
            mainContainer.style.marginTop = (CONSTANTS.TOOLBAR_HEIGHT + toolbarHeight) + 'px';
        });
    } else {
        mainContainer.style.marginTop = CONSTANTS.TOOLBAR_HEIGHT + 'px';
    }
}

window.updatePreview = function (forceShow = false) {
    const toolbar     = document.getElementById('customToolbar');
    const savedGroups = AppStore.get('custom_toolbar_rules');

    if (!savedGroups || savedGroups.length === 0) {
        _setCustomToolbarVisible(toolbar, false);
        toolbar.innerHTML = '';
        return;
    }
    if (!forceShow && toolbar.style.display !== 'flex') {
        _setCustomToolbarVisible(toolbar, false);
        return;
    }
    _setCustomToolbarVisible(toolbar, true);

    toolbar.innerHTML = '';
    savedGroups.forEach(group => {
        const selectElement = document.createElement('select');
        selectElement.className = 'custom-rule-select';

        const titleOption = document.createElement('option');
        titleOption.text  = group.groupName || '그룹 선택';
        titleOption.value = '';
        titleOption.disabled = titleOption.selected = true;
        selectElement.appendChild(titleOption);

        group.items.forEach(item => {
            if (item.name.trim() || item.html.trim()) {
                const option = document.createElement('option');
                option.value = item.html;
                option.text  = item.name || '내용 없음';
                selectElement.appendChild(option);
            }
        });

        let _selectDebounceTimer = null;
        selectElement.onchange = function () {
            if (!this.value) return;
            const htmlToInsert = this.value;
            this.selectedIndex = 0;
            clearTimeout(_selectDebounceTimer);
            _selectDebounceTimer = setTimeout(() => {
                const savedRange = EditorState.get('savedRange');
                EditorState.startSync();
                if (savedRange) {
                    const selection = window.getSelection();
                    selection.removeAllRanges();
                    selection.addRange(savedRange);
                    const fragment = savedRange.createContextualFragment(htmlToInsert);
                    const lastInsertedNode = fragment.lastChild;
                    savedRange.deleteContents();
                    savedRange.insertNode(fragment);
                    if (lastInsertedNode) {
                        const newRange = document.createRange();
                        newRange.setStartAfter(lastInsertedNode);
                        newRange.setEndAfter(lastInsertedNode);
                        selection.removeAllRanges();
                        selection.addRange(newRange);
                        EditorState.set('savedRange', newRange.cloneRange());
                    }
                }
                // isSyncing을 즉시 해제한 뒤 다음 tick에 동기화
                // endSync(false)의 50ms 지연이 남아있으면 syncPreviewToEditor가 guard에 걸려 실패함
                EditorState.endSync(true);
                requestAnimationFrame(() => window.syncPreviewToEditor?.());
            }, 0);
        };
        toolbar.appendChild(selectElement);
    });
};

const dayMaps = {
    ko_short: ['일','월','화','수','목','금','토'],
    ko_long:  ['일요일','월요일','화요일','수요일','목요일','금요일','토요일'],
    en_short: ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'],
    en_long:  ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'],
};

const _dayMatchOrder = ['ko_long', 'en_long', 'ko_short', 'en_short'];
const _dayMapsLower  = Object.fromEntries(
    Object.entries(dayMaps).map(([key, labels]) => [key, labels.map(label => label.toLowerCase())])
);

const DayManager = {
    DAY_IDX: { WEEKDAY: 0, SAT: 1, SUN: 2 },
    getAllPatterns() {
        return Object.values(dayMaps).flat().sort((a, b) => b.length - a.length);
    },
    getGroupType(dayOfWeek) {
        if (dayOfWeek === 6) return this.DAY_IDX.SAT;
        if (dayOfWeek === 0) return this.DAY_IDX.SUN;
        return this.DAY_IDX.WEEKDAY;
    },
    getIdxFromText(text) {
        if (!text) return -1;
        const lowerText = text.toLowerCase();
        for (const mapKey of _dayMatchOrder) {
            const foundIndex = _dayMapsLower[mapKey].findIndex(label => lowerText.includes(label));
            if (foundIndex !== -1) return foundIndex;
        }
        return -1;
    },
    getTypeFromText(text) {
        if (!text) return 'ko_short';
        const lowerText = text.toLowerCase();
        for (const mapKey of _dayMatchOrder) {
            if (_dayMapsLower[mapKey].some(label => lowerText.includes(label))) return mapKey;
        }
        return 'ko_short';
    },
    getLabel(dayIndex, type = 'ko_short') { return dayMaps[type]?.[dayIndex] || ''; },
    getDayStr(dayIndex, type = 'ko_short') { return this.getLabel(dayIndex, type); },
};

const CALENDAR_THEME = {
    COLORS: {
        SUN: 'red',
        SAT: 'blue',
        WEEKDAY: '#333333',
        DEFAULT: '#333333'
    }
};

function getCalendarColor(columnIndex, showHoliday) {
    if (!showHoliday) return CALENDAR_THEME.COLORS.DEFAULT;
    if (columnIndex === 0) return CALENDAR_THEME.COLORS.SUN;
    if (columnIndex === 6) return CALENDAR_THEME.COLORS.SAT;
    return CALENDAR_THEME.COLORS.WEEKDAY;
}

function isValidYearMonth(yearMonth) {
    const parts = yearMonth.split('/');
    if (parts.length !== 2) return false;
    const [year, month] = parts.map(Number);
    return !isNaN(year) && !isNaN(month) && year > 0 && month >= 1 && month <= 12;
}
window.isValidYearMonth = isValidYearMonth;

function generateBaseCalendar(yearMonth, options = {}) {
    const { showHoliday = true, useId = false, baseId = '', lineHeight = '2' } = options;
    const [year, month] = yearMonth.split('/').map(Number);
    const firstDayOfWeek = new Date(year, month - 1, 1).getDay();
    const lastDateOfMonth = new Date(year, month, 0).getDate();

    const weekdayLabels = dayMaps.ko_short;
    let html = `<table style="width:100%;table-layout:fixed;border-collapse:collapse;border:1px solid #ddd;text-align:center;line-height:${lineHeight};">\n`;
    html += `  <thead>\n    <tr>\n`;
    weekdayLabels.forEach((dayLabel, columnIndex) => {
        const color = getCalendarColor(columnIndex, showHoliday);
        html += `      <th style="width:14%;padding:10px;color:${color};">${dayLabel}</th>\n`;
    });
    html += `    </tr>\n  </thead>\n  <tbody>\n`;

    let currentDate = 1;
    for (let weekRow = 0; weekRow < 6; weekRow++) {
        if (currentDate > lastDateOfMonth) break;
        html += `    <tr>\n`;
        for (let columnIndex = 0; columnIndex < 7; columnIndex++) {
            if ((weekRow === 0 && columnIndex < firstDayOfWeek) || currentDate > lastDateOfMonth) {
                html += `      <td></td>\n`;
            } else {
                const color = getCalendarColor(columnIndex, showHoliday);
                const dateText = `<span style="color:${color};">${currentDate}</span>`;
                const cellId   = baseId ? `${baseId}_${currentDate}` : currentDate;
                const content  = useId
                    ? `<a href="#${CONSTANTS.USER_CONTENT_PREFIX}${cellId}" style="text-decoration:none;">${dateText}</a>`
                    : dateText;
                html += `      <td style="padding:10px;">${content}</td>\n`;
                currentDate++;
            }
        }
        html += `    </tr>\n`;
    }
    html += `  </tbody>\n</table>`;
    return html;
}
window.generateBaseCalendar = generateBaseCalendar;

function transformAdvancedCalendar(sourceHtml, fromYearMonth, toYearMonth) {
    const parsedDoc = DomManager.parse(sourceHtml);
    const table     = parsedDoc?.querySelector('table');
    if (!table) return sourceHtml;

    const [toYear, toMonth] = toYearMonth.split('/').map(Number);
    const firstDayOfWeek = new Date(toYear, toMonth - 1, 1).getDay();
    const lastDateOfMonth = new Date(toYear, toMonth, 0).getDate();

    const tbody    = TableUtils.getTbody(table);
    if (!tbody || tbody === table) return sourceHtml;

    const cellTemplates = { sun: null, sat: null, work: null };
    const sourceRows    = TableUtils.getRows(tbody);

    sourceRows.forEach(tableRow => {
        tableRow.querySelectorAll('td').forEach((cell, columnIndex) => {
            if (columnIndex > 6) return;
            const textContent   = cell.textContent.trim();
            if (!textContent.length) return;
            const isNormalDate = /^\d{1,2}$/.test(textContent) && !cell.innerHTML.includes('text-shadow');
            if      (columnIndex === 0 && !cellTemplates.sun) cellTemplates.sun = cell;
            else if (columnIndex === 6 && !cellTemplates.sat) cellTemplates.sat = cell;
            else if (columnIndex >= 1 && columnIndex <= 5) {
                if (!cellTemplates.work || (!cellTemplates.work._isNormal && isNormalDate)) {
                    cellTemplates.work = cell;
                    cellTemplates.work._isNormal = isNormalDate;
                }
            }
        });
    });

    const fallbackCell = cellTemplates.work || cellTemplates.sun || cellTemplates.sat;
    if (!fallbackCell) return sourceHtml;

    const lastSourceRowCell = sourceRows[sourceRows.length - 1]?.querySelector('td');
    const lastRowHasNoBorderBottom = lastSourceRowCell?.getAttribute('style')?.includes('border-bottom:0px');

    function buildCleanCell(sourceCell, isLastWeek) {
        let borderBottomOverride = null;
        if (isLastWeek && lastRowHasNoBorderBottom) {
            const lastRowBorders = StyleUtils.parseBorderSides(sourceRows[sourceRows.length - 1].cells[0]);
            borderBottomOverride = lastRowBorders.bottom;
        }
        return StyleUtils.cloneStructuralTd(sourceCell, borderBottomOverride);
    }

    function cloneContentWrapper(sourceWrapper, dateNumber) {
        const newWrapper = sourceWrapper.cloneNode(true);
        const updateTextRecursive = (node) => {
            if (node.nodeType === 3 && node.nodeValue.trim().length > 0) {
                node.nodeValue = String(dateNumber);
            } else {
                node.childNodes.forEach(child => updateTextRecursive(child));
            }
        };
        updateTextRecursive(newWrapper);
        const linkElement = newWrapper.tagName === 'A' ? newWrapper : newWrapper.querySelector('a');
        if (linkElement) {
            const oldHref  = linkElement.getAttribute('href') || '';
            const hrefMatch = oldHref.match(/#user_content_([^\d]*)(\d+)/);
            if (hrefMatch) {
                const paddingLength = hrefMatch[2].length;
                const paddedNumber  = String(dateNumber).padStart(paddingLength, '0');
                linkElement.setAttribute('href', `#${CONSTANTS.USER_CONTENT_PREFIX}${hrefMatch[1]}${paddedNumber}`);
            } else {
                linkElement.setAttribute('href', `#${CONSTANTS.USER_CONTENT_PREFIX}d${dateNumber}`);
            }
        }
        return newWrapper;
    }

    const newTbody     = tbody.cloneNode(false);
    let currentDate    = 1;

    for (let weekRow = 0; weekRow < 6; weekRow++) {
        if (currentDate > lastDateOfMonth) break;
        const newTableRow = document.createElement('tr');
        const rowStyleAttr = sourceRows[weekRow % sourceRows.length]?.getAttribute('style') || '';
        if (rowStyleAttr) newTableRow.setAttribute('style', rowStyleAttr);

        for (let columnIndex = 0; columnIndex < 7; columnIndex++) {
            const isLastWeek = currentDate + (7 - columnIndex) > lastDateOfMonth;
            const templateCell  = columnIndex === 0 ? (cellTemplates.sun  || fallbackCell)
                                : columnIndex === 6 ? (cellTemplates.sat  || fallbackCell)
                                :                     (cellTemplates.work || fallbackCell);
            const newCell       = buildCleanCell(templateCell, isLastWeek);
            const isCellEmpty   = (weekRow === 0 && columnIndex < firstDayOfWeek) || currentDate > lastDateOfMonth;

            if (isCellEmpty) {
                newCell.innerHTML = '&nbsp;';
            } else {
                const contentWrapper = templateCell.querySelector('a') || templateCell.querySelector('span') || templateCell.querySelector('p');
                if (contentWrapper) newCell.appendChild(cloneContentWrapper(contentWrapper, currentDate));
                else newCell.textContent = String(currentDate);
                currentDate++;
            }
            newTableRow.appendChild(newCell);
        }
        newTbody.appendChild(newTableRow);
    }

    table.replaceChild(newTbody, tbody);
    return table.outerHTML;
}

function applyCasing(sampleText, targetText) {
    if (!/[a-zA-Z]/.test(sampleText)) return targetText;
    if (sampleText === sampleText.toUpperCase()) return targetText.toUpperCase();
    if (sampleText === sampleText.toLowerCase()) return targetText.toLowerCase();
    return targetText.charAt(0).toUpperCase() + targetText.slice(1).toLowerCase();
}

window.executeExtendRow = function () {
    const editor  = EditorState.get('editor');
    const preview = EditorState.get('preview');
    if (!editor) { window.showToast('에디터가 초기화되지 않았습니다.', 'error'); return; }

    const rawTemplateTable = SampleCache.getTemplateTable();
    if (!rawTemplateTable) {
        window.showToast('샘플 코드가 없습니다. 분석창에서 먼저 설정해 주세요.');
        return;
    }
    const templateTable = rawTemplateTable.cloneNode(true);

    let fromDay         = parseInt(document.getElementById('modalExtendFrom').value) || 1;
    const toDay         = parseInt(document.getElementById('modalExtendTo').value)   || 31;
    const idSuffix      = (document.getElementById('modalDateId')?.value || '').trim();
    const colorInputRaw = (document.getElementById('modalTargetAttr').value || '').trim();
    const baseMonthValue = document.getElementById('modalBaseMonth')?.value;

    const hexColors = StyleUtils.parseColorInput(colorInputRaw)
    .map(color => ColorManager.toOriginalForm(color))
    .filter(Boolean);

    const sourceThead      = templateTable.querySelector(':scope > thead');
    const sourceTbody      = TableUtils.getTbody(templateTable);
    const allTbodyRows     = TableUtils.getRows(sourceTbody);

    const dataStartIndex   = findDataStartIdx(allTbodyRows);
    const subHeaderRows    = allTbodyRows.slice(0, dataStartIndex);
    const tbodyDataRows    = allTbodyRows.slice(dataStartIndex);
    const templateGroups   = getRowGroups(tbodyDataRows);
    const groupCount       = templateGroups.length;

    const firstSampleDateCell  = templateGroups[0]?.[0]?.cells[0];
    const sampleDateInfo       = window.parseCellDate(firstSampleDateCell);
    const sampleDateNumber     = sampleDateInfo ? sampleDateInfo.date : 1;

    const sampleCellText = firstSampleDateCell?.textContent || '';
    let sampleDayIndex   = DayManager.getIdxFromText(sampleCellText);
    let foundDayLabelList = dayMaps[DayManager.getTypeFromText(sampleCellText)] || null;

    if (sampleDayIndex === -1 && baseMonthValue) {
        const [baseYear, baseMonth] = baseMonthValue.split('-').map(Number);
        const dateObject = new Date(baseYear, baseMonth - 1, sampleDateNumber);
        if (!isNaN(dateObject.getTime())) { sampleDayIndex = dateObject.getDay(); foundDayLabelList = dayMaps.ko_short; }
    }

    let targetTable, targetTbody;
    const currentEditorContent = editor.getValue().trim();

    if (currentEditorContent.includes('<table')) {
        const currentDoc = DomManager.parse(currentEditorContent);
        targetTable  = currentDoc.querySelector('table');
        targetTbody  = TableUtils.getTbody(targetTable);

        const fromDayRawInput = document.getElementById('modalExtendFrom').value.trim();
        if (!fromDayRawInput) {
            const allExistingRows = TableUtils.getRows(targetTbody);
            let lastFoundDateNumber = 0, rowspanRemaining = 0;

            for (const tableRow of allExistingRows) {
                const firstCell = tableRow.cells[0];
                if (!firstCell) continue;
                if (rowspanRemaining > 0) { rowspanRemaining--; continue; }
                const rowspan = parseInt(firstCell.getAttribute('rowspan') || '1');
                if (TableUtils.isDateCell(firstCell)) {
                    const numberMatch = firstCell.textContent.trim().match(/(\d{1,2})/);
                    const dateNum = numberMatch ? parseInt(numberMatch[1]) : 0;
                    if (dateNum > lastFoundDateNumber) lastFoundDateNumber = dateNum;
                }
                if (rowspan > 1) rowspanRemaining = rowspan - 1;
            }
            if (lastFoundDateNumber > 0) fromDay = lastFoundDateNumber + 1;
        }
    } else {
        targetTable = templateTable;
        targetTbody = TableUtils.getTbody(targetTable);
        targetTbody.innerHTML = '';
        if (sourceThead && !targetTable.querySelector('thead')) {
            targetTable.insertBefore(sourceThead.cloneNode(true), targetTable.firstChild);
        }
        subHeaderRows.forEach(row => targetTbody.appendChild(row.cloneNode(true)));
    }

    if (fromDay > toDay) { window.showToast('시작일이 종료일보다 클 수 없습니다.'); return; }
    if (fromDay > 31)    { window.showToast('31일을 초과할 수 없습니다.'); return; }

    const allDayPatterns = DayManager.getAllPatterns();
    const rowFragment    = document.createDocumentFragment();

    const updateDateNodeRecursive = (parentNode, targetDateNumber, targetDayName) => {
        parentNode.childNodes.forEach(childNode => {
            if (childNode.nodeType === 3) {
                let textContent = childNode.nodeValue;
                if (/\d+/.test(textContent)) textContent = textContent.replace(/\d+/, targetDateNumber);
                if (targetDayName) {
                    for (const pattern of allDayPatterns) {
                        const patternRegex = new RegExp(pattern, 'gi');
                        if (patternRegex.test(textContent)) { textContent = textContent.replace(patternRegex, targetDayName); break; }
                    }
                }
                childNode.nodeValue = textContent;
            } else if (childNode.nodeType === 1) {
                updateDateNodeRecursive(childNode, targetDateNumber, targetDayName);
            }
        });
    };

    for (let dayNumber = fromDay; dayNumber <= Math.min(toDay, 31); dayNumber++) {
        const groupIndex    = (dayNumber - 1) % groupCount;
        const currentGroup  = templateGroups[groupIndex];
        const firstGroupRow = currentGroup[0];
        const lastGroupRow  = currentGroup.at(-1);
        const isMultiRow    = currentGroup.length > 1;

        const newRow = document.createElement('tr');
        Array.from(firstGroupRow.attributes).forEach(attribute => {
            newRow.setAttribute(attribute.name, attribute.name === 'style' ? StyleUtils.hexStyle(attribute.value) : attribute.value);
        });

        if (isMultiRow) {
            const bottomBorder = StyleUtils.extractBottomBorder(currentGroup);
            if (bottomBorder) {
                let currentRowStyle = newRow.getAttribute('style') || '';
                currentRowStyle = StyleUtils.applyBottomBorder(currentRowStyle, bottomBorder);
                newRow.setAttribute('style', StyleUtils.hexStyle(currentRowStyle));
            }
        }

        const currentDayIndex = sampleDayIndex !== -1
            ? (sampleDayIndex + (dayNumber - sampleDateNumber) % 7 + 7) % 7
            : null;
        const cellColor = (hexColors.length > 0 && currentDayIndex !== null)
            ? StyleUtils.getColorForDay(hexColors, currentDayIndex)
            : null;

        Array.from(firstGroupRow.cells).forEach((templateCell, cellIndex) => {
            let newCell;
            if (cellIndex === 0) {
                newCell = templateCell.cloneNode(true);
                newCell.removeAttribute('rowspan');

                if (isMultiRow) {
                    const lastRowCellBorderMatch = lastGroupRow.cells[0]?.getAttribute('style')?.match(/border-bottom\s*:[^;]+/i);
                    if (lastRowCellBorderMatch) {
                        let cellStyleString = newCell.getAttribute('style') || '';
                        cellStyleString = StyleUtils.applyBottomBorder(cellStyleString, lastRowCellBorderMatch[0]);
                        newCell.setAttribute('style', StyleUtils.hexStyle(cellStyleString));
                    }
                }

                const sampleCellTextContent = firstSampleDateCell?.textContent?.trim() || '';
                const paddedDateNumber = DateUtils.padNumberToMatch(dayNumber, sampleCellTextContent);

                let targetDayName = '';
                if (foundDayLabelList && currentDayIndex !== null) {
                    const dayLabelRaw = foundDayLabelList[currentDayIndex];
                    const sampleLabelMatches = firstSampleDateCell.textContent.match(
                        new RegExp(foundDayLabelList[sampleDayIndex], 'gi')
                    );
                    targetDayName = sampleLabelMatches ? applyCasing(sampleLabelMatches[0], dayLabelRaw) : dayLabelRaw;
                }
                updateDateNodeRecursive(newCell, paddedDateNumber, targetDayName);

                const generatedCellId = DateUtils.generateCellId(templateCell.id || '', dayNumber, sampleCellTextContent, idSuffix);
                if (generatedCellId) newCell.id = generatedCellId;
                else                 newCell.removeAttribute('id');

                if (cellColor) StyleUtils.applyColorToDateCell(newCell, cellColor);
                if (newCell.getAttribute('style')) {
                    newCell.setAttribute('style', StyleUtils.hexStyle(newCell.getAttribute('style')));
                }
            } else {
                const firstGroupCell = currentGroup[0].cells[cellIndex];
                const lastGroupCell  = isMultiRow ? (lastGroupRow.cells[cellIndex - 1] ?? firstGroupCell) : firstGroupCell;
                newCell = TableUtils.createEmptyCell(templateCell, null);
                StyleUtils.applyCellStyle(newCell, firstGroupCell, lastGroupCell);
                if (newCell.getAttribute('style')) {
                    newCell.setAttribute('style', StyleUtils.hexStyle(newCell.getAttribute('style')));
                }
                newCell.innerHTML = '&nbsp;';
            }
            newRow.appendChild(newCell);
        });
        rowFragment.appendChild(newRow);
    }
    targetTbody.appendChild(rowFragment);

    const allResultRows      = TableUtils.getRows(targetTbody);
    const resultDataStartIdx = findDataStartIdx(allResultRows);
    const firstResultDateNumber = parseInt(
        allResultRows[resultDataStartIdx]?.cells[0]?.textContent?.trim()?.match(/\d+/)?.[0] || '0'
    );
    const shouldLockHeader = (firstResultDateNumber !== 1);

    window.syncTableToEditor(targetTable, shouldLockHeader);
    if (!shouldLockHeader) window.releaseHeaderLock();

    setTimeout(() => {
        const previewTable = preview.querySelector('table');
        if (!previewTable) return;
        const previewRows = TableUtils.getRows(TableUtils.getTbody(previewTable));
        const allDateCells = previewRows
            .map(tableRow => tableRow.cells[0])
            .filter(cell => TableUtils.isDateCell(cell));
        const focusTargetCell = allDateCells[allDateCells.length - (toDay - fromDay + 1)] || allDateCells.at(-1);
        focusCellInPreview(focusTargetCell);
    }, CONSTANTS.PREVIEW_SYNC_DELAY);
};

function insertEmptyRowAfter(referenceRow, templateRow, templateCells, hexBackground) {
    const newRow = templateRow.cloneNode(false);
    StyleUtils.applyRowTrStyle(newRow, templateRow, hexBackground);
    templateCells.forEach(templateCell => {
        const newCell = templateCell.cloneNode(false);
        let styleString = StyleUtils.removeRadius(StyleUtils.hexStyle(templateCell.getAttribute('style') || ''));      
        newCell.setAttribute('style', StyleUtils.hexStyle(StyleUtils.applyBg(styleString, hexBackground)));
        newCell.innerHTML = '&nbsp;';
        newRow.appendChild(newCell);
    });
    referenceRow.parentNode.insertBefore(newRow, referenceRow.nextSibling);
    return newRow;
}

function resolveTargetCell(previewArea) {
    const selection = window.getSelection();
    if (selection?.anchorNode) {
        const cell = getResolvedNode(selection.anchorNode).closest('td');
        if (cell) return cell;
    }
    const active = document.activeElement;
    const fromActive = active?.closest('#previewArea td') 
                    || (active?.tagName === 'TD' ? active : null);
    if (fromActive) return fromActive;

    if (EditorState.get('savedRange')) {
        const cell = getResolvedNode(EditorState.get('savedRange').startContainer)
                        ?.closest('#previewArea td');
        if (cell) return cell;
    }
    const candidate = EditorState.currentTargetNode?.closest?.('td');
    if (candidate && previewArea.contains(candidate)) return candidate;
    return null;
}

window.splitCurrentRow = function () {
    const preview = EditorState.get('preview');
	const previewArea = document.getElementById('previewArea');
	const targetCell = resolveTargetCell(previewArea);
	if (!targetCell) {
        return window.showToast('분할할 칸을 선택해주세요.');
    }
    const currentTableRow = targetCell.closest('tr');
    const table           = TableUtils.getTable(currentTableRow);
    if (!table) return window.showToast('테이블을 찾을 수 없습니다.');

    const targetTbody   = TableUtils.getTbody(table);
    const tbodyRows     = TableUtils.getRows(targetTbody);
    const currentRowIndex = tbodyRows.indexOf(currentTableRow);

    const templateTable = SampleCache.getTemplateTable();
    if (!templateTable) return window.showToast('샘플코드가 없습니다.');

    const allTemplateRows = Array.from(templateTable.querySelectorAll(':scope > tbody > tr')).filter(row => row.cells.length > 0);
    const templateDataRows = allTemplateRows.slice(findDataStartIdx(allTemplateRows));
    const sampleGroups     = getRowGroups(templateDataRows);
    const referenceGroup   = sampleGroups.find(group => group.length >= 2) || sampleGroups[0];
    const firstSampleRow   = referenceGroup[0];    
    const lastSampleRow    = referenceGroup.at(-1);   
    const middleSampleRow  = referenceGroup.length > 2 ? referenceGroup[1] : firstSampleRow;  
    const sampleIsMultiRow = referenceGroup.length > 1;

    let dateCellElement = null, dateRowStartIndex = -1;
    for (let rowIndex = currentRowIndex; rowIndex >= 0; rowIndex--) {
        const cell = tbodyRows[rowIndex].cells[0];
        if (!cell) continue;
        const rowspan  = parseInt(cell.getAttribute('rowspan')) || 1;
        const hasIdPrefix = cell.id?.includes(CONSTANTS.USER_CONTENT_PREFIX);
        const hasRowspan  = rowspan >= 2;
        const hasNumber   = /\d/.test(cell.textContent.trim());

        if (rowIndex === currentRowIndex) {
            if (hasIdPrefix || hasRowspan || hasNumber) { 
                dateCellElement  = cell;
                dateRowStartIndex = rowIndex;
                break;
            }
        } else {
            if (rowspan < 2 || rowIndex + rowspan <= currentRowIndex) continue;
            if (hasIdPrefix || hasRowspan || hasNumber) {
                dateCellElement  = cell;
                dateRowStartIndex = rowIndex;
                break;
            }
        }
    }
    if (!dateCellElement) {
        const firstCellOfCurrentRow = tbodyRows[currentRowIndex]?.cells[0];
        if (firstCellOfCurrentRow) { 
            dateCellElement   = firstCellOfCurrentRow; 
            dateRowStartIndex = currentRowIndex; 
        }
    }
    if (!dateCellElement) return window.showToast('날짜 칸을 찾을 수 없습니다.');

    const currentRowspan   = parseInt(dateCellElement.getAttribute('rowspan')) || 1;
    const groupLastRowIndex = dateRowStartIndex + currentRowspan - 1;
    const isFirstRowInGroup = currentRowIndex === dateRowStartIndex;
    const isLastRowInGroup  = currentRowIndex === groupLastRowIndex;

    const hexBackground = StyleUtils.extractBgColor(tbodyRows[dateRowStartIndex]);

    if (currentRowspan === 1) {
        dateCellElement.setAttribute('rowspan', '2');
        StyleUtils.applyRowTrStyle(currentTableRow, firstSampleRow, hexBackground);
        Array.from(currentTableRow.cells).forEach((cell, cellIndex) => {
            if (cellIndex === 0) return;
            const templateCell = firstSampleRow.cells[cellIndex];
            if (!templateCell) return;
            cell.setAttribute('style', StyleUtils.hexStyle(StyleUtils.applyBg(StyleUtils.hexStyle(templateCell.getAttribute('style') || ''), hexBackground)));
        });
        const lastSampleCells = sampleIsMultiRow ? Array.from(lastSampleRow.cells) : Array.from(firstSampleRow.cells).slice(1);
		insertEmptyRowAfter(currentTableRow, lastSampleRow, lastSampleCells, hexBackground);

    } else if (isLastRowInGroup) {
        dateCellElement.setAttribute('rowspan', currentRowspan + 1);
        StyleUtils.applyRowTrStyle(currentTableRow, middleSampleRow, hexBackground);
        Array.from(currentTableRow.cells).forEach((cell, cellIndex) => {
            const isFirstRowOfGroup   = (currentTableRow === tbodyRows[dateRowStartIndex]);
            const visualColumnIndex   = isFirstRowOfGroup ? cellIndex : cellIndex + 1;
            const templateCell        = middleSampleRow.cells[visualColumnIndex];
            if (templateCell) {
                let styleString = StyleUtils.removeRadius(StyleUtils.hexStyle(templateCell.getAttribute('style') || ''));
                cell.setAttribute('style', StyleUtils.hexStyle(StyleUtils.applyBg(styleString, hexBackground)));
            }
        });
        const lastSampleCells = sampleIsMultiRow ? Array.from(lastSampleRow.cells) : Array.from(firstSampleRow.cells).slice(1);
        insertEmptyRowAfter(currentTableRow, lastSampleRow, lastSampleCells, hexBackground);

    } else {
        dateCellElement.setAttribute('rowspan', currentRowspan + 1);
        const cellsToClone = isFirstRowInGroup ? Array.from(currentTableRow.cells).slice(1) : Array.from(currentTableRow.cells);
        insertEmptyRowAfter(currentTableRow, currentTableRow, cellsToClone, hexBackground);
        if (isFirstRowInGroup) {
			StyleUtils.applyRowCellStyles(currentTableRow, currentTableRow, hexBackground, { 
				radiusMode: 'bottom', 
				skipFirstCell: true 
			});
		}
    }

    window.syncTableToEditor(table);

    setTimeout(() => {
        const previewTable = TableUtils.getTable(preview.querySelector('table'));
        if (!previewTable) return;
        const previewRows      = TableUtils.getRows(TableUtils.getTbody(previewTable));
        const newRowInPreview  = previewRows[currentRowIndex + 1];
        if (!newRowInPreview) return;
        focusCellInPreview(newRowInPreview.cells[0]);
    }, CONSTANTS.PREVIEW_SYNC_DELAY);
};
