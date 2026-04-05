/**
 * editor-analyzer.js
 */

const SampleCache = {
    _dom: null,
    _html: null,
	
	_loadFromStorage() {
        if (!this._html) {
            const saved = AppStore.get('analysis_source_save');
            if (saved) {
                this._html = saved;
                this._dom = DomManager.parse(saved);
                return true;
            }
        }
        return !!this._html;
    },
    set(html) {
        const cleanHtml = html?.trim() || "";

        if (!cleanHtml) {
            this._html = null;
            this._dom = null;
            AppStore.remove('analysis_source_save'); 
            return;
        }
        if (cleanHtml === this._html) return;

        this._html = cleanHtml;
        this._dom = DomManager.parse(cleanHtml);
        AppStore.set('analysis_source_save', cleanHtml);
    },

    getHtml() {
        this._loadFromStorage(); 
        return this._html || '';
    },

    refreshUI() {
        const html = this.getHtml();
        const input = document.getElementById('analysisInput');
        
        if (input) input.value = html;
        if (window.analysisEditor) window.analysisEditor.setValue(html);
    },

    getTemplateTable() {
        const input = document.getElementById('analysisInput');
        if (input) {
            const inputValue = input.value.trim();
            if (inputValue && inputValue !== (this._html || "")) {
                this.set(inputValue);
            }
        }

        if (!this._dom) {
            this._loadFromStorage();
        }

        if (!this._dom) return null;
        const tbl = this._dom.querySelector('table');
        return tbl ? DomManager.clone(tbl) : null;
    },

    init() {
        this._loadFromStorage(); 
        this.refreshUI();
    }
};
window.SampleCache = SampleCache;

const StyleUtils = {

    hexStyle(styleStr) {
        return ColorManager.hexifyStyle(styleStr);
    },

    normValue(v, prop) {
        v = (v || '').trim();
        if (!v) return v;
        if (prop && (prop.includes('color') || prop.includes('background'))) {
            return ColorManager.toHex(v);
        }
        return v.toLowerCase();
    },

    parseStyleObj(styleStr) {
        const el = document.createElement('div');
        el.setAttribute('style', styleStr || '');
        return el.style;   // CSSStyleDeclaration (live)
    },

    parseStyleForCompare(styleStr) {
        const map = {};
        const style = this.parseStyleObj(styleStr);
        for (let i = 0; i < style.length; i++) {
            const prop = style[i];
            const val = this.normValue(style.getPropertyValue(prop), prop);
            if (val) map[prop] = val;
        }
        return map;
    },
	getCleanStyle: function(el) {
        if (!el) return '';
        let style = el.getAttribute('style') || '';
        return style.replace(/cursor:[^;]+;?/g, '').trim(); 
    },
    parseBorderRadius(el) {
        if (!el) return { tl: '', tr: '', br: '', bl: '' };
        const s = el.style;
        return {
            tl: s.borderTopLeftRadius     || s.borderRadius || '',
            tr: s.borderTopRightRadius    || s.borderRadius || '',
            br: s.borderBottomRightRadius || s.borderRadius || '',
            bl: s.borderBottomLeftRadius  || s.borderRadius || '',
        };
    },

    parseBorderSides(el) {
        if (!el) return { top: '', right: '', bottom: '', left: '' };
        const s = el.style;
        const get = (w, st, c) => [w, st, c].filter(Boolean).join(' ') || s.border || '';
        return {
            top:    s.borderTop    || get(s.borderTopWidth,    s.borderTopStyle,    s.borderTopColor),
            right:  s.borderRight  || get(s.borderRightWidth,  s.borderRightStyle,  s.borderRightColor),
            bottom: s.borderBottom || get(s.borderBottomWidth, s.borderBottomStyle, s.borderBottomColor),
            left:   s.borderLeft   || get(s.borderLeftWidth,   s.borderLeftStyle,   s.borderLeftColor),
        };
    },

    applyCellStyle(newCell, firstRowCell, lastRowCell) {
        const fb = this.parseBorderSides(firstRowCell);
        const lb = this.parseBorderSides(lastRowCell);
        newCell.style.border = '';
        if (fb.top)    newCell.style.borderTop    = fb.top;
        if (fb.left)   newCell.style.borderLeft   = fb.left;
        if (fb.right)  newCell.style.borderRight  = fb.right;
        if (lb.bottom) newCell.style.borderBottom = lb.bottom;

        const fr = this.parseBorderRadius(firstRowCell);
        const lr = this.parseBorderRadius(lastRowCell);
        newCell.style.borderRadius = '';
        if (fr.tl) newCell.style.borderTopLeftRadius     = fr.tl;
        if (fr.tr) newCell.style.borderTopRightRadius    = fr.tr;
        if (lr.br) newCell.style.borderBottomRightRadius = lr.br;
        if (lr.bl) newCell.style.borderBottomLeftRadius  = lr.bl;

        const shadow  = firstRowCell?.style.boxShadow || lastRowCell?.style.boxShadow || '';
        const outline = firstRowCell?.style.outline   || '';
        if (shadow)  newCell.style.boxShadow = shadow;
        if (outline) newCell.style.outline   = outline;
    },

    cloneStructuralTd(srcTd, borderBottomOverride = null) {
        const td = srcTd.cloneNode(false);
        if (srcTd.colSpan > 1)          td.colSpan = srcTd.colSpan;
        ['align', 'width', 'valign'].forEach(attr => {
            if (srcTd.hasAttribute(attr)) td.setAttribute(attr, srcTd.getAttribute(attr));
        });
        this.applyCellStyle(td, srcTd, srcTd);
        if (borderBottomOverride !== null) td.style.borderBottom = borderBottomOverride;
        return td;
    },
};
window.StyleUtils = StyleUtils;

// ═══════════════════════════════════════════════════════════
//  getRowGroups — rowspan 기반 행 그룹핑 공통 함수
// ═══════════════════════════════════════════════════════════
function getRowGroups(rows) {
    const groups = [];
    let i = 0;
    while (i < rows.length) {
        const rs = parseInt(rows[i].cells[0]?.getAttribute('rowspan')) || 1;
        groups.push(rows.slice(i, i + rs));
        i += rs;
    }
    return groups;
}
window.getRowGroups = getRowGroups;

function findDataStartIdx(rows) {
    for (let i = 0; i < rows.length; i++) {
        const c0 = rows[i].cells[0];
        if (!c0) continue;
        const rs      = parseInt(c0.getAttribute('rowspan')) || 1;
        const hasDate = /\d/.test(c0.textContent.trim()) || c0.id?.includes('user_content_');
        if (rs >= 2 || hasDate) return i;
    }
    return 0;
}
window.findDataStartIdx = findDataStartIdx;

window.parseCellDate = function(cell) {
    if (!cell) return null;
    const text = cell.textContent.trim();
    const dateMatch = text.match(/\d+/);
    if (!dateMatch) return null;

    const dateNum = parseInt(dateMatch[0]);
    const dayStr = (typeof DayManager !== 'undefined') ? DayManager.getDayStr(dateNum) : '';
    return {
        date: dateNum,
        day: dayStr
    };
};

function focusCellInPreview(targetTd, markerType = 'new') {
    if (!targetTd) return;

    const allTds = Array.from(preview.querySelectorAll('td'));
    let liveTd = targetTd;

    if (!preview.contains(targetTd)) {
        const targetHtml = targetTd.outerHTML;
        liveTd = allTds.find(td => td.outerHTML === targetHtml) || allTds[allTds.length - 1];
    }
    if (!liveTd) return;

    if (liveTd.getAttribute('contenteditable') !== 'true') {
        makeEditableOnlyCells(preview);
    }

    liveTd.focus();
    const range = document.createRange();
    range.selectNodeContents(liveTd);
    range.collapse(true);
    const sel = window.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);

    savedRange = range.cloneRange();
    currentTargetNode = liveTd;

    const tdIdx = allTds.indexOf(liveTd);
    if (tdIdx === -1) return;

    const content = editor.getValue();
    const tdRegex = /<td\b/gi;
    let m, cnt = 0, tLine = -1;
    while ((m = tdRegex.exec(content)) !== null) {
        if (cnt === tdIdx) { tLine = content.substring(0, m.index).split('\n').length - 1; break; }
        cnt++;
    }
    if (tLine === -1) return;

    editor.clearGutter('markers');
    const mk = document.createElement('div');
    mk.className = `working-marker working-marker--${markerType}`;
    mk.innerHTML = '●';
    editor.setGutterMarker(tLine, 'markers', mk);
    editor.scrollIntoView({ line: tLine, ch: 0 }, 200);

    if (window._headerLockRange && typeof window.applyHeaderLock === 'function') {
        requestAnimationFrame(() => window.applyHeaderLock());
    }
}

window.syncTableToEditor = function (tableElement, lockHeaderLines = false) {
    if (!tableElement) return;
    let html = tableElement.outerHTML;

    html = html.replace(/\s*contenteditable="[^"]*"/gi, '');
    html = html.replace(/\bcursor\s*:[^;"]*(;)?/gi, '');
    html = html.replace(/\s*style=""/gi, '');
    html = StyleUtils.hexStyle(html);

    withSyncLock(() => {
        if (typeof window.clearHeaderLock === 'function') window.clearHeaderLock();
        const beautified = (typeof html_beautify !== 'undefined')
            ? html_beautify(html, BEAUTIFY_OPTIONS)
            : html;
        editor.setValue(beautified);
        if (editor.refresh) editor.refresh();
    });
    if (typeof window.patchPreview === 'function') {
        window.patchPreview(html);
    } else {
        preview.innerHTML = html;
    }
    makeEditableOnlyCells(preview);

    if (lockHeaderLines || window._headerLockRange) {
        setTimeout(() => window.applyHeaderLock(), 0);
    }
};

window.isCalendarTable = function() {
    if (!editor) return false;
    const html = editor.getValue();
    const doc = DomManager.parse(html);
    if (!doc) return false;
    const table = doc.querySelector('table');
    if (!table) return false;
    const firstRow = table.querySelector('tr');
    if (!firstRow) return false;
    const totalCols = Array.from(firstRow.cells).reduce(
        (s, td) => s + (parseInt(td.getAttribute('colspan')) || 1), 0
    );
    return totalCols === 7;
};

window.applyHeaderLock = function() {
    if (!editor) return;
    if (window.isCalendarTable()) {
        window.releaseHeaderLock();
        return;
    }
    window.clearHeaderLock();

    const lines = editor.getValue().split('\n');
    const totalLines = lines.length;

    const tableLine = lines.findIndex(l => /<table[\s>]/i.test(l));
    const tbodyLine = lines.findIndex(l => /<tbody[\s>]/i.test(l));

    let closeTbodyLine = -1, closeTableLine = -1;
    for (let i = totalLines - 1; i >= 0; i--) {
        if (closeTableLine === -1 && /<\/table>/i.test(lines[i])) closeTableLine = i;
        if (closeTbodyLine === -1 && /<\/tbody>/i.test(lines[i]))  closeTbodyLine = i;
        if (closeTableLine !== -1 && closeTbodyLine !== -1) break;
    }

    if (tableLine < 0 || tbodyLine <= tableLine) return;

    const headerEnd  = tbodyLine;          
    const footerStart = closeTbodyLine >= 0 ? closeTbodyLine : closeTableLine; 

    window._headerLockedLines = [];

    for (let i = tableLine; i <= headerEnd; i++) {
        editor.addLineClass(i, 'text', 'cm-header-locked');
        window._headerLockedLines.push(i);
    }
    if (footerStart > headerEnd) {
        const footerEnd = closeTableLine >= 0 ? closeTableLine : footerStart;
        for (let i = footerStart; i <= footerEnd; i++) {
            editor.addLineClass(i, 'text', 'cm-header-locked');
            window._headerLockedLines.push(i);
        }
    }

    window._headerLockRange = {
        trStart: tbodyLine + 1,
        trEnd:   footerStart > 0 ? footerStart - 1 : totalLines - 1
    };
};

window.clearHeaderLock = function() {
    if (window._headerLockedLines) {
        window._headerLockedLines.forEach(i => {
            try { editor.removeLineClass(i, 'text', 'cm-header-locked'); } catch(e) {}
        });
        window._headerLockedLines = [];
    }
};

window.releaseHeaderLock = function() {
    window.clearHeaderLock();
    window._headerLockRange = null;
};

// ═══════════════════════════════════════════════════════════
//  샘플 정제 
// ═══════════════════════════════════════════════════════════

function sanitizeHtml(html) {
    return html
        .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
        .replace(/<iframe\b[^<]*(?:(?!<\/iframe>)<[^<]*)*<\/iframe>/gi, '')
        .replace(/<object\b[^<]*(?:(?!<\/object>)<[^<]*)*<\/object>/gi, '')
        .replace(/<embed\b[^>]*>/gi, '')
        .replace(/<svg\b[^<]*(?:(?!<\/svg>)<[^<]*)*<\/svg>/gi, '')
        .replace(/<link\b[^>]*>/gi, '')
        .replace(/<meta\b[^>]*>/gi, '')
        .replace(/<base\b[^>]*>/gi, '')
        .replace(/\s+on\w+\s*=\s*["'][^"']*["']/gi, '')
        .replace(/\s+on\w+\s*=\s*[^\s>]+/gi, '')
        .replace(/href\s*=\s*["']\s*javascript:[^"']*/gi, 'href="#"')
        .replace(/src\s*=\s*["']\s*javascript:[^"']*/gi, 'src=""');
}

function processAnalysis() {
    const input = document.getElementById('analysisInput');
    const raw = sanitizeHtml(input?.value.trim() || '');
    if (!raw) {
        SampleCache.set("");
        if (input) input.value = "";
        window.showToast('저장된 데이터가 삭제되었습니다.', 'info');
        return;
    }
    const check = window.validateHtmlInput(raw, 'table');
    if (!check.ok) {
        window.showToast(check.reason, 'error');
        if (input) input.value = "";
        SampleCache.set("");
        return;
    }
    try {
        const cleanedHtml = getCleanTable(raw);
        if (cleanedHtml) {
            input.value = cleanedHtml;
            SampleCache.set(cleanedHtml);
            window.showToast('데이터가 성공적으로 분석되었습니다.', 'success');
        } else {
            if (input) input.value = "";
            SampleCache.set("");
            window.showToast('유효한 테이블 구조를 찾을 수 없어 초기화되었습니다.', 'error');
        }
    } catch (e) {
        window.showToast('분석 중 오류가 발생하여 데이터를 비웁니다.', 'error');
        if (input) input.value = "";
        SampleCache.set("");
    }
}
window.processAnalysis = processAnalysis;

function getCleanTable(rawHtml) {
    const doc = DomManager.parse(rawHtml);
    const sourceTable = doc.querySelector('table');
    if (!sourceTable) return null;
 
    sourceTable.querySelectorAll('td table, th table').forEach(nt => nt.remove());
    const cleanTable = document.createElement('table');
    Array.from(sourceTable.attributes).forEach(a => cleanTable.setAttribute(a.name, a.value));
    const tableStyleAttr = cleanTable.getAttribute('style') || '';
    if (!/border-collapse/i.test(tableStyleAttr)) {
        cleanTable.setAttribute('style', (tableStyleAttr + ';border-collapse:collapse').replace(/^;/, ''));
    }
 
    const caption = sourceTable.querySelector(':scope > caption');
    if (caption) cleanTable.appendChild(caption.cloneNode(true));
    sourceTable.querySelectorAll(':scope > colgroup').forEach(cg => cleanTable.appendChild(cg.cloneNode(true)));
 
    const srcThead = sourceTable.querySelector(':scope > thead');
    if (srcThead) cleanTable.appendChild(srcThead.cloneNode(true));
    const tbodies = Array.from(sourceTable.querySelectorAll(':scope > tbody'));
 
    const firstTr = sourceTable.querySelector('tr');
    const totalCols = firstTr
        ? Array.from(firstTr.cells).reduce((s, td) => s + (parseInt(td.getAttribute('colspan')) || 1), 0)
        : 3;
 
    function isHeaderRow(tr) {
        return tr.querySelector('th') !== null;
    }
    function isSubHeaderRow(tr) {
        if (tr.querySelector('th')) return false;
        const cells = Array.from(tr.cells);
        if (cells.length === 0) return false;
        const span = cells.reduce((s, td) => s + (parseInt(td.getAttribute('colspan')) || 1), 0);
        return span >= totalCols && cells.length < totalCols;
    }
    function isDataGroupStart(tr) {
        const c0 = tr.cells[0];
        if (!c0) return false;
        if ((parseInt(c0.getAttribute('rowspan')) || 1) >= 2) return true;
        if (/^\d{1,2}$/.test(c0.textContent.trim())) return true;
        if (c0.id?.includes('user_content_')) return true;
        return false;
    }
 
    let headerRows = [], subHeaderRows = [], dataRows;
 
    if (srcThead) {
        dataRows = tbodies.flatMap(tb => Array.from(tb.querySelectorAll(':scope > tr')));
    } else {
        const allRows = tbodies.length > 0
            ? tbodies.flatMap(tb => Array.from(tb.querySelectorAll(':scope > tr')))
            : Array.from(sourceTable.querySelectorAll('tr'))
                .filter(tr => tr.parentElement.closest('table') === sourceTable);
 
        let dataStart = allRows.length;
        for (let i = 0; i < allRows.length; i++) {
            if (isDataGroupStart(allRows[i])) { dataStart = i; break; }
            if (isHeaderRow(allRows[i]))           headerRows.push(allRows[i]);
            else if (isSubHeaderRow(allRows[i]))   subHeaderRows.push(allRows[i]);
            else                                   headerRows.push(allRows[i]);
        }
        dataRows = allRows.slice(dataStart);
    }
 
    const dateGroups = [];
    let spanLeft = 0, currentGroup = [];
    dataRows.forEach(tr => {
        if (spanLeft === 0) {
            if (currentGroup.length > 0) dateGroups.push({ rows: currentGroup });
            currentGroup = [];
            spanLeft = parseInt(tr.cells[0]?.getAttribute('rowspan')) || 1;
        }
        currentGroup.push(tr);
        spanLeft--;
    });
    if (currentGroup.length > 0) dateGroups.push({ rows: currentGroup });
 
    const COMPARE_KEYS = new Set([
        'background', 'background-color', 'color',
        'border', 'border-top', 'border-right', 'border-bottom', 'border-left',
        'border-color', 'border-top-color', 'border-right-color',
        'border-bottom-color', 'border-left-color',
        'border-style', 'border-width', 'border-radius',
        'border-top-left-radius', 'border-top-right-radius',
        'border-bottom-left-radius', 'border-bottom-right-radius',
        'outline', 'outline-color',
    ]);
 
    function getTrStyleMap(tr) {
        const dc = tr.cells[0];
        const combined = [
            tr.getAttribute('style') || '',
            dc?.getAttribute('style') || '',
            dc?.getAttribute('bgcolor') ? `background-color:${dc.getAttribute('bgcolor')}` : ''
        ].join(';');
        const full = StyleUtils.parseStyleForCompare(combined);
        const filtered = {};
        COMPARE_KEYS.forEach(k => { if (full[k] !== undefined) filtered[k] = full[k]; });
        return filtered;
    }
 
    function styleMapsEqual(m1, m2) {
        const keys = new Set([...Object.keys(m1), ...Object.keys(m2)]);
        for (const k of keys) if ((m1[k] || '') !== (m2[k] || '')) return false;
        return true;
    }
 
    let keepGroups;
    if      (dateGroups.length === 0) keepGroups = [];
    else if (dateGroups.length === 1) keepGroups = [dateGroups[0]];
    else {
        const s1 = getTrStyleMap(dateGroups[0].rows[0]);
        const s2 = getTrStyleMap(dateGroups[1].rows[0]);
        keepGroups = styleMapsEqual(s1, s2) ? [dateGroups[0]] : [dateGroups[0], dateGroups[1]];
    }
 
    function buildCleanRows(group) {
        const result  = [];
        const src0    = group.rows[0];
        const srcLast = group.rows[group.rows.length - 1];
        const isMulti = group.rows.length > 1;
        const rowSpanVal = isMulti ? 2 : 1;
 
        [src0, ...(isMulti ? [srcLast] : [])].forEach((src, order) => {
            const tr = src.cloneNode(false);
            Array.from(src.cells).forEach((sc, ci) => {
                const tc = sc.cloneNode(false);
                const rawStyle = sc.getAttribute('style') || '';
                if (rawStyle) tc.setAttribute('style', StyleUtils.hexStyle(rawStyle));
                if (sc.hasAttribute('bgcolor')) tc.setAttribute('bgcolor', window.rgbToHex(sc.getAttribute('bgcolor')));
                if (order === 0 && ci === 0) {
                    tc.innerHTML = StyleUtils.hexStyle(sc.innerHTML.trim());
                    tc.rowSpan = rowSpanVal;
                } else {
                    tc.innerHTML = '&nbsp;';
                    if (tc.rowSpan > 1) tc.rowSpan = 1;
                }
                tr.appendChild(tc);
            });
            result.push(tr);
        });
        return result;
    }
 
    if (!srcThead && headerRows.length > 0) {
        const newThead = document.createElement('thead');
        headerRows.forEach(r => newThead.appendChild(r.cloneNode(true)));
        cleanTable.appendChild(newThead);
    }
 
    const finalTbody = document.createElement('tbody');
    const tbodyStyle = tbodies[0]?.getAttribute('style');
    if (tbodyStyle) finalTbody.setAttribute('style', tbodyStyle);
 
    subHeaderRows.forEach(r => finalTbody.appendChild(r.cloneNode(true)));
    keepGroups.forEach(g => buildCleanRows(g).forEach(tr => finalTbody.appendChild(tr)));
    cleanTable.appendChild(finalTbody);
 
    const tfoot = sourceTable.querySelector(':scope > tfoot');
    if (tfoot) cleanTable.appendChild(tfoot.cloneNode(true));
 
    const temp = document.createElement('div');
    temp.style.display = 'none';
    document.body.appendChild(temp);
    temp.appendChild(cleanTable);
    const result = temp.innerHTML;
    document.body.removeChild(temp);
    return result;
}

// ═══════════════════════════════════════════════════════════
//  커스텀 툴바 규칙 설정
// ═══════════════════════════════════════════════════════════
function addGroup() {
    const container = document.getElementById('ruleGroupsContainer');
    const newGroup  = document.createElement('div');
    newGroup.className = 'rule-group-card';
    newGroup.innerHTML = `
        <div class="group-header">
            <input type="text" class="group-name-input" placeholder="그룹 이름 (예: 카테고리1)">
            <button class="btn-del-group" onclick="this.closest('.rule-group-card').remove()">그룹 삭제</button>
        </div>
        <table class="rule-item-table">
            <tbody class="item-list">
                <tr>
                    <td style="width:25%"><input type="text" class="modal-input" placeholder="표시 이름"></td>
                    <td style="width:65%"><textarea class="modal-input code-area" placeholder="HTML 코드를 입력하세요"></textarea></td>
                    <td style="width:10%"><button class="btn-del-item" onclick="this.closest('tr').remove()">×</button></td>
                </tr>
            </tbody>
        </table>
        <button class="btn-add-item-dashed" onclick="addItem(this)">+ 항목 추가</button>
    `;
    container.appendChild(newGroup);
}
window.addGroup = addGroup;

function addItem(btn) {
    const tbody  = btn.closest('.rule-group-card').querySelector('.item-list');
    const newRow = document.createElement('tr');
    newRow.innerHTML = `
        <td><input type="text" class="modal-input" placeholder="표시 이름"></td>
        <td><textarea class="modal-input code-area" placeholder="HTML 코드를 입력하세요"></textarea></td>
        <td><button class="btn-del-item" onclick="this.closest('tr').remove()">×</button></td>
    `;
    tbody.appendChild(newRow);
}
window.addItem = addItem;

function applyAndSaveRules() {
    const container  = document.getElementById('ruleGroupsContainer');
    const groupCards = container.querySelectorAll('.rule-group-card');
    const groups     = [];

    groupCards.forEach(card => {
        const groupName = card.querySelector('.group-name-input').value;
        const items     = [];
        card.querySelectorAll('.item-list tr').forEach(row => {
            const inputs = row.querySelectorAll('input, textarea');
            if (inputs[0].value.trim() || inputs[1].value.trim()) {
                items.push({ name: inputs[0].value, html: inputs[1].value });
            }
        });
        if (items.length > 0 || groupName.trim()) groups.push({ groupName, items });
    });

    if (groups.length === 0) {
        if (confirm('입력된 규칙이 없습니다. 기존 설정으로 되돌리시겠습니까?')) {
            window.renderRules?.();
            closeModal('ruleModal');
        }
        return;
    }
    AppStore.set('custom_toolbar_rules', groups);
    window.showToast('설정이 저장되었습니다.');
    window.updatePreview?.(true);
    closeModal('ruleModal');
}
window.applyAndSaveRules = applyAndSaveRules;

window.renderRules = function () {
    const container = document.getElementById('ruleGroupsContainer');
    const groups    = AppStore.get('custom_toolbar_rules');
    container.innerHTML = '';

    if (!groups || groups.length === 0) { addGroup(); return; }

    groups.forEach(groupData => {
        addGroup();
        const lastGroup = container.lastElementChild;
        lastGroup.querySelector('.group-name-input').value = groupData.groupName;
        const tbody = lastGroup.querySelector('.item-list');
        tbody.innerHTML = '';
        const items = groupData.items.length > 0 ? groupData.items : [{ name: '', html: '' }];
        items.forEach(item => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td style="width:25%"><input type="text" class="modal-input" value="${item.name}" placeholder="표시 이름"></td>
                <td style="width:65%"><textarea class="modal-input code-area" placeholder="HTML 코드를 입력하세요">${item.html}</textarea></td>
                <td style="width:10%"><button class="btn-del-item" onclick="this.closest('tr').remove()">×</button></td>
            `;
            tbody.appendChild(row);
        });
    });
};

window.updatePreview = function (forceShow = false) {
    const toolbar = document.getElementById('customToolbar');
    const groups  = AppStore.get('custom_toolbar_rules');

    if (!groups || groups.length === 0) {
        toolbar.style.display = 'none';
        toolbar.innerHTML = '';
        return;
    }
    if (forceShow || toolbar.style.display === 'flex') {
        toolbar.style.display = 'flex';
    } else {
        toolbar.style.display = 'none';
        return;
    }

    toolbar.innerHTML = '';
    groups.forEach(group => {
        const select = document.createElement('select');
        select.className = 'custom-rule-select';

        const titleOpt = document.createElement('option');
        titleOpt.text = group.groupName || '그룹 선택';
        titleOpt.value = '';
        titleOpt.disabled = titleOpt.selected = true;
        select.appendChild(titleOpt);

        group.items.forEach(item => {
            if (item.name.trim() || item.html.trim()) {
                const opt = document.createElement('option');
                opt.value = item.html;
                opt.text  = item.name || '내용 없음';
                select.appendChild(opt);
            }
        });

        let _customSelectTimer = null;
        select.onchange = function () {
            if (!this.value) return;
            const htmlToInsert = this.value;
            this.selectedIndex = 0;
            clearTimeout(_customSelectTimer);
            _customSelectTimer = setTimeout(() => {
                isSyncing = true;
                if (savedRange) {
                    const sel = window.getSelection();
                    sel.removeAllRanges();
                    sel.addRange(savedRange);
                    const fragment = savedRange.createContextualFragment(htmlToInsert);
                    const lastNode = fragment.lastChild;
                    savedRange.deleteContents();
                    savedRange.insertNode(fragment);
                    if (lastNode) {
                        const r = document.createRange();
                        r.setStartAfter(lastNode);
                        r.setEndAfter(lastNode);
                        sel.removeAllRanges();
                        sel.addRange(r);
                        savedRange = r.cloneRange();
                    }
                }
                setTimeout(() => {
                    isSyncing = false;
                    window.syncPreviewToEditor?.();
                }, 100);
            }, 0);
        };
        toolbar.appendChild(select);
    });
};

// ═══════════════════════════════════════════════════════════
//  캘린더 생성/변환
// ═══════════════════════════════════════════════════════════
const dayMaps = {
    ko_short: ['일','월','화','수','목','금','토'],
    ko_long:  ['일요일','월요일','화요일','수요일','목요일','금요일','토요일'],
    en_short: ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'], // 'MON' 대신 'Mon'
    en_long:  ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday']
};


const _dayMatchOrder = ['ko_long', 'en_long', 'ko_short', 'en_short'];
const _dayMapsLower = Object.fromEntries(
    Object.entries(dayMaps).map(([k, v]) => [k, v.map(d => d.toLowerCase())])
);

const DayManager = {
    DAY_IDX: {
        WEEKDAY: 0,
        SAT: 1,
        SUN: 2
    },
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
        const lower = text.toLowerCase();
        for (const key of _dayMatchOrder) {
            const idx = _dayMapsLower[key].findIndex(d => lower.includes(d));
            if (idx !== -1) return idx;
        }
        return -1;
    },
    getTypeFromText(text) {
        if (!text) return 'ko_short';
        const lower = text.toLowerCase();
        for (const key of _dayMatchOrder) {
            if (_dayMapsLower[key].some(d => lower.includes(d))) return key;
        }
        return 'ko_short';
    },
    getLabel(idx, type = 'ko_short') {
        return dayMaps[type]?.[idx] || '';
    },
    getDayStr(idx, type = 'ko_short') {
        return this.getLabel(idx, type);
    }
};

function isValidYearMonth(ym) {
    const parts = ym.split('/');
    if (parts.length !== 2) return false;
    const [year, month] = parts.map(Number);
    return !isNaN(year) && !isNaN(month) && year > 0 && month >= 1 && month <= 12;
}
window.isValidYearMonth = isValidYearMonth;

function generateBaseCalendar(yearMonth, options = {}) {
    const { showHoliday = true, useId = false, baseId = '', lineHeight = '2' } = options;
    const [year, month] = yearMonth.split('/').map(Number);
    const firstDay = new Date(year, month - 1, 1).getDay();
    const lastDate = new Date(year, month, 0).getDate();

    let html = `<table style="width:100%;table-layout:fixed;border-collapse:collapse;border:1px solid #ddd;text-align:center;line-height:${lineHeight};">\n`;
    const weekDays = dayMaps.ko_short;
    html += `  <thead>\n    <tr>\n`;
    weekDays.forEach((day, i) => {
        const color = showHoliday ? (i === 0 ? 'red' : i === 6 ? 'blue' : '#333333') : '#333333';
        html += `      <th style="width:14%;padding:10px;color:${color};">${day}</th>\n`;
    });
    html += `    </tr>\n  </thead>\n  <tbody>\n`;

    let date = 1;
    for (let i = 0; i < 6; i++) {
        if (date > lastDate) break;
        html += `    <tr>\n`;
        for (let j = 0; j < 7; j++) {
            if ((i === 0 && j < firstDay) || date > lastDate) {
                html += `      <td></td>\n`;
            } else {
                const color   = showHoliday ? (j === 0 ? 'red' : j === 6 ? 'blue' : '#333333') : '#333333';
                const dateText = `<span style="color:${color};">${date}</span>`;
                const idStr    = baseId ? `${baseId}_${date}` : date;
                const content  = useId
                    ? `<a href="#user_content_${idStr}" style="text-decoration:none;">${dateText}</a>`
                    : dateText;
                html += `      <td style="padding:10px;">${content}</td>\n`;
                date++;
            }
        }
        html += `    </tr>\n`;
    }
    html += `  </tbody>\n</table>`;
    return html;
}
window.generateBaseCalendar = generateBaseCalendar;

function transformAdvancedCalendar(sourceHtml, fromYM, toYM) {
    const doc = DomManager.parse(sourceHtml);
    const table = doc.querySelector('table');
    if (!table) return sourceHtml;

    const [toYear, toMonth] = toYM.split('/').map(Number);
    const firstDay = new Date(toYear, toMonth - 1, 1).getDay();
    const lastDate = new Date(toYear, toMonth, 0).getDate();

    const tbody = table.querySelector('tbody');
    if (!tbody) return sourceHtml;

    const tdTemplates = { sun: null, sat: null, work: null };
    const srcRows = Array.from(tbody.querySelectorAll('tr'));

    srcRows.forEach(tr => {
        tr.querySelectorAll('td').forEach((td, colIdx) => {
            if (colIdx > 6) return;
            const text = td.innerText.trim();
            if (text.length === 0) return; 

            const isNormal = /^\d{1,2}$/.test(text) && !td.innerHTML.includes('text-shadow');

            if (colIdx === 0 && !tdTemplates.sun) {
                tdTemplates.sun = td;
            } else if (colIdx === 6 && !tdTemplates.sat) {
                tdTemplates.sat = td;
            } else if (colIdx >= 1 && colIdx <= 5) {
                if (!tdTemplates.work || (!tdTemplates.work._isNormal && isNormal)) {
                    tdTemplates.work = td;
                    tdTemplates.work._isNormal = isNormal;
                }
            }
        });
    });

    const fallbackTd = tdTemplates.work || tdTemplates.sun || tdTemplates.sat;
    if (!fallbackTd) return sourceHtml;

    const lastRowSampleTd = srcRows[srcRows.length - 1]?.querySelector('td');
    const isOriginalLastRowNoBorder = lastRowSampleTd?.getAttribute('style')?.includes('border-bottom:0px');

    function cleanTd(srcTd, isLastWeek) {
        let borderBottomOverride = null;
        if (isLastWeek && isOriginalLastRowNoBorder) {
            const lastBorders = StyleUtils.parseBorderSides(srcRows[srcRows.length - 1].cells[0]);
            borderBottomOverride = lastBorders.bottom;
        }
        return StyleUtils.cloneStructuralTd(srcTd, borderBottomOverride);
    }

    function cloneWrapper(srcWrapper, dateNum) {
		const newWrapper = srcWrapper.cloneNode(true);
		const updateText = (node) => {
			if (node.nodeType === 3 && node.nodeValue.trim().length > 0) {
				node.nodeValue = String(dateNum);
			} else {
				node.childNodes.forEach(child => updateText(child));
			}
		};
		updateText(newWrapper);
		const linkEl = newWrapper.tagName === 'A' ? newWrapper : newWrapper.querySelector('a');
		if (linkEl) {
			const oldHref = linkEl.getAttribute('href') || '';
		
			const match = oldHref.match(/#user_content_([^\d]*)(\d+)/);

			if (match) {
				const userPrefix = match[1];  
				const oldNumStr = match[2];  
				
				const paddingLength = oldNumStr.length;
				const newNumStr = String(dateNum).padStart(paddingLength, '0');

				linkEl.setAttribute('href', `#user_content_${userPrefix}${newNumStr}`);
			} else {
				linkEl.setAttribute('href', `#user_content_d${String(dateNum)}`);
			}
		}

		return newWrapper;
	}

    const newTbody = tbody.cloneNode(false);
    let dateCounter = 1;

    for (let week = 0; week < 6; week++) {
        if (dateCounter > lastDate) break;
        const newTr = document.createElement('tr');
       
        const trStyle = srcRows[week % srcRows.length]?.getAttribute('style') || '';
        if (trStyle) newTr.setAttribute('style', trStyle);

        for (let col = 0; col < 7; col++) {
            const isLastWeek = dateCounter + (7 - col) > lastDate;
            const srcTd = col === 0 ? (tdTemplates.sun  || fallbackTd)
                        : col === 6 ? (tdTemplates.sat  || fallbackTd)
                        :             (tdTemplates.work || fallbackTd);

            const newTd = cleanTd(srcTd, isLastWeek);
            const isEmpty = (week === 0 && col < firstDay) || dateCounter > lastDate;

            if (isEmpty) {
                newTd.innerHTML = '&nbsp;';
            } else {
                const wrapper = srcTd.querySelector('a') || srcTd.querySelector('span') || srcTd.querySelector('p');
                if (wrapper) {
                    newTd.appendChild(cloneWrapper(wrapper, dateCounter));
                } else {
                    newTd.textContent = String(dateCounter);
                }
                dateCounter++;
            }
            newTr.appendChild(newTd);
        }
        newTbody.appendChild(newTr);
    }

    table.replaceChild(newTbody, tbody);
    return table.outerHTML;
}

// ═══════════════════════════════════════════════════════════
//  줄 확장 (executeExtendRow)
// ═══════════════════════════════════════════════════════════

function applyCasing(sample, target) {
    if (!/[a-zA-Z]/.test(sample)) return target;
    if (sample === sample.toUpperCase()) return target.toUpperCase();
    if (sample === sample.toLowerCase()) return target.toLowerCase();
    return target.charAt(0).toUpperCase() + target.slice(1).toLowerCase();
}

function removeRadius(styleStr) {
    return styleStr
        .replace(/border-radius\s*:[^;]+;?/gi, '')
        .replace(/border-top-left-radius\s*:[^;]+;?/gi, '')
        .replace(/border-top-right-radius\s*:[^;]+;?/gi, '')
        .replace(/border-bottom-left-radius\s*:[^;]+;?/gi, '')
        .replace(/border-bottom-right-radius\s*:[^;]+;?/gi, '')
        .replace(/;+/g, ';').replace(/^;|;$/g, '');
}

function applyBg(styleStr, bg) {
    if (!bg) return styleStr.replace(/background(?:-color)?\s*:[^;]+;?/gi, '').replace(/^;|;$/g, '');
    const s = styleStr.replace(/background(?:-color)?\s*:[^;]+;?/gi, '');
    return (s.replace(/;+$/, '') + `;background-color:${bg}`).replace(/^;/, '');
}

function parseColorInput(raw) {
    return (raw || '').split(',').map(s => s.trim()).filter(Boolean);
}

function getColorForDay(colors, realDayIdx) {
    const group = DayManager.getGroupType(realDayIdx);
    if (group === DayManager.DAY_IDX.SUN) return colors[2];
    if (group === DayManager.DAY_IDX.SAT) return colors[1];
    return colors[0] || null;
}

function replaceStyleProp(styleStr, prop, newVal) {
    const re = new RegExp(`(${prop}\\s*:\\s*)[^;]+`, 'i');
    if (re.test(styleStr)) return styleStr.replace(re, `$1${newVal}`);
    return styleStr.trim().replace(/;?$/, '') + `;${prop}:${newVal}`;
}

function wrapTextNodesWithColor(node, hexColor) {
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT, {
        acceptNode: (n) => n.nodeValue.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP
    });
    const nodes = [];
    let cur;
    while ((cur = walker.nextNode())) nodes.push(cur);
    nodes.forEach(textNode => {
        const span = document.createElement('span');
        span.setAttribute('style', `color:${hexColor}`);
        textNode.parentNode.insertBefore(span, textNode);
        span.appendChild(textNode);
    });
}

function applyColorToDateCell(cell, hexColor) {
    if (!hexColor) return;

    const tdStyle = cell.getAttribute('style') || '';
    const hasTdColor = /(?:^|;)\s*color\s*:/i.test(tdStyle);

    if (hasTdColor) {
        const newStyle = tdStyle.replace(/((?:^|;)\s*color\s*:)[^;]*/i, `$1${hexColor}`);
        cell.setAttribute('style', newStyle.trim());
        return;
    }

    const colorEls = cell.querySelectorAll('p[style*="color"], span[style*="color"]');
    if (colorEls.length > 0) {
        colorEls.forEach(el => {
            const elStyle = el.getAttribute('style') || '';
            el.setAttribute('style', replaceStyleProp(elStyle, 'color', hexColor));
        });
        return;
    }

    wrapTextNodesWithColor(cell, hexColor);
}

window.executeExtendRow = function() {
    const rawTemplate = SampleCache.getTemplateTable();
    if (!rawTemplate) {
        window.showToast("샘플 코드가 없습니다. 분석창에서 먼저 설정해주세요.");
        return;
    }
    const templateTable = rawTemplate.cloneNode(true);

    let fromDay        = parseInt(document.getElementById('modalExtendFrom').value) || 1;
    const toDay        = parseInt(document.getElementById('modalExtendTo').value)   || 31;
    const idSuffix     = (document.getElementById('modalDateId').value || 'd').trim();
    const colorRaw     = (document.getElementById('modalTargetAttr').value || '').trim();
    const baseMonthVal = document.getElementById('modalBaseMonth')?.value;

    const colors = parseColorInput(colorRaw).map(c => ColorManager.toHex(c)).filter(Boolean);

    const srcThead     = templateTable.querySelector(':scope > thead');
    const srcTbody     = templateTable.querySelector(':scope > tbody');
    const allTbodyRows = srcTbody
        ? Array.from(srcTbody.querySelectorAll(':scope > tr')).filter(r => r.cells.length > 0)
        : Array.from(templateTable.querySelectorAll('tr')).filter(r => r.cells.length > 0);

    let dataStartIdx = findDataStartIdx(allTbodyRows);

    const subHeaderRows = allTbodyRows.slice(0, dataStartIdx);
    const tbodyRows = allTbodyRows.slice(dataStartIdx);

    const templateGroups = getRowGroups(tbodyRows);
    const groupCount     = templateGroups.length;

    const firstSampleDateCell = templateGroups[0]?.[0]?.cells[0];
    const sampleDateInfo = window.parseCellDate(firstSampleDateCell);
    const sampleDayNum   = sampleDateInfo ? sampleDateInfo.date : 1;

    const rawText = firstSampleDateCell?.textContent || '';
	const pureDayText = rawText.replace(/[0-9]/g, '').trim(); 

	let sampleDayIdx = DayManager.getIdxFromText(pureDayText);
	let foundDayList = dayMaps[DayManager.getTypeFromText(pureDayText)];

    if (sampleDayIdx === -1 && baseMonthVal) {
        const [bYear, bMonth] = baseMonthVal.split('-').map(Number);
        const d = new Date(bYear, bMonth - 1, sampleDayNum);
        if (!isNaN(d.getTime())) { sampleDayIdx = d.getDay(); foundDayList = dayMaps.ko_short; }
    }

    let targetTable, targetTbody;
    const currentContent = editor?.getValue().trim() || '';

    if (currentContent.includes('<table')) {
        const currentDoc = DomManager.parse(currentContent);
        targetTable = currentDoc.querySelector('table');
        targetTbody = targetTable.querySelector('tbody') || targetTable;

        const modalFromRaw = document.getElementById('modalExtendFrom').value.trim();
        if (!modalFromRaw) {
            const tbodyEl = targetTable.querySelector('tbody') || targetTable;
            const allRows = Array.from(tbodyEl.querySelectorAll('tr'));
            let lastDateNum = 0;

            let rowspanRemaining = 0;
            for (const tr of allRows) {
                const c0 = tr.cells[0];
                if (!c0) continue;

                if (rowspanRemaining > 0) {
                    rowspanRemaining--;
                    continue;
                }

                const rs     = parseInt(c0.getAttribute('rowspan') || '1');
                const hasId  = c0.id && c0.id.includes('user_content_');
                const txt    = c0.textContent.trim();
                const numMatch = txt.match(/(\d{1,2})/); 

                if (hasId || numMatch) {
                    const n = numMatch ? parseInt(numMatch[1]) : 0;
                    if (n > lastDateNum) lastDateNum = n;
                }

                if (rs > 1) rowspanRemaining = rs - 1;
            }
            if (lastDateNum > 0) fromDay = lastDateNum + 1;
        }
    } else {
        targetTable = templateTable;
        targetTbody = targetTable.querySelector('tbody') || targetTable;
        targetTbody.innerHTML = '';
        if (srcThead && !targetTable.querySelector('thead')) {
            targetTable.insertBefore(srcThead.cloneNode(true), targetTable.firstChild);
        }
        subHeaderRows.forEach(r => targetTbody.appendChild(r.cloneNode(true)));
    }

    if (fromDay > toDay)  { window.showToast('시작일이 종료일보다 클 수 없습니다.'); return; }
    if (fromDay > 31)     { window.showToast('31일을 초과할 수 없습니다.'); return; }

    const allDayPatterns = DayManager.getAllPatterns();
    const fragment = document.createDocumentFragment();

    for (let d = fromDay; d <= Math.min(toDay, 31); d++) {
        const groupIdx    = (d - 1) % groupCount;
        const targetGroup = templateGroups[groupIdx];
        const firstRow    = targetGroup[0];
        const lastRow     = targetGroup.at(-1);
        const isMultiRow  = targetGroup.length > 1;

        const newRow = document.createElement('tr');
        Array.from(firstRow.attributes).forEach(attr => {
            let val = attr.value;
            if (attr.name === 'style') val = StyleUtils.hexStyle(val);
            newRow.setAttribute(attr.name, val);
        });

        if (isMultiRow) {
            const firstTrStyle = firstRow.getAttribute('style') || '';
            const lastTrStyle  = lastRow.getAttribute('style') || '';
            const allTrStyles  = targetGroup.map(r => r.getAttribute('style') || '').join(';');

            const solidMatch  = allTrStyles.match(/border-bottom\s*:\s*[^;]*solid[^;]*/i);
            const lastMatch   = lastTrStyle.match(/border-bottom\s*:[^;]+/i);
            const finalBorder = solidMatch?.[0] || lastMatch?.[0];

            if (finalBorder) {
                let curStyle = newRow.getAttribute('style') || '';
                if (/border-bottom/i.test(curStyle)) {
                    curStyle = curStyle.replace(/border-bottom\s*:[^;]+/i, finalBorder);
                } else {
                    curStyle = (curStyle + ';' + finalBorder).replace(/^;+/, '');
                }
                newRow.setAttribute('style', StyleUtils.hexStyle(curStyle));
            }
        }

        const currentDayIdx = sampleDayIdx !== -1
            ? (sampleDayIdx + (d - sampleDayNum) % 7 + 7) % 7
            : null;
        const cellColor = (colors.length > 0 && currentDayIdx !== null)
            ? getColorForDay(colors, currentDayIdx)
            : null;

        Array.from(firstRow.cells).forEach((templateCell, cellIdx) => {
            const newCell = templateCell.cloneNode(true);
            newCell.removeAttribute('rowspan'); 
            if (newCell.getAttribute('style')) {
                newCell.setAttribute('style', StyleUtils.hexStyle(newCell.getAttribute('style')));
            }

            if (cellIdx === 0) {
                const sampleId   = templateCell.id || '';
                const idInputVal = (document.getElementById('modalDateId')?.value || '').trim();
                const idNumMatch = sampleId.match(/\d+$/);
                const idPadding  = (idNumMatch?.[0]?.startsWith('0')) ? idNumMatch[0].length : 0;

                const sampleCellText = firstSampleDateCell?.textContent?.trim() || '';
                const hasLeadZero    = /^0\d/.test(sampleCellText);
                const targetDayNum   = hasLeadZero ? String(d).padStart(2, '0') : String(d);

                let targetDayName = '';
                if (foundDayList && currentDayIdx !== null) {
                    const raw     = foundDayList[currentDayIdx];
                    const matched = firstSampleDateCell.textContent.match(
                        new RegExp(foundDayList[sampleDayIdx], 'gi')
                    );
                    targetDayName = matched ? applyCasing(matched[0], raw) : raw;
                }

                const updateNode = (parentNode) => {
                    parentNode.childNodes.forEach(node => {
                        if (node.nodeType === 3) {
                            let text = node.nodeValue;
                            if (/\d+/.test(text)) text = text.replace(/\d+/, targetDayNum);
                            if (targetDayName) {
                                for (const pat of allDayPatterns) {
                                    const reg = new RegExp(pat, 'gi');
                                    if (reg.test(text)) { text = text.replace(reg, targetDayName); break; }
                                }
                            }
                            node.nodeValue = text;
                        } else if (node.nodeType === 1) {
                            updateNode(node);
                        }
                    });
                };
                updateNode(newCell);

                let finalIdValue = '';
                if (sampleId) {
                    const idBase      = sampleId.replace(/\d+$/, '');
                    const finalNumStr = idPadding > 0 ? String(d).padStart(idPadding, '0') : String(d);
                    finalIdValue = `${idBase}${finalNumStr}`;
                } else if (idInputVal) {
                    finalIdValue = `user_content_${idInputVal}${d}`;
                }
                if (finalIdValue) newCell.id = finalIdValue;
                else newCell.removeAttribute('id');

                if (cellColor) applyColorToDateCell(newCell, cellColor);

                if (isMultiRow) {
                    const lastDateCell = lastRow.cells[0];
                    if (lastDateCell) {
                        const lastCellStyle = lastDateCell.getAttribute('style') || '';
                        const lastBorderMatch = lastCellStyle.match(/border-bottom\s*:[^;]+/i);
                        if (lastBorderMatch) {
                            let cs = newCell.getAttribute('style') || '';
                            if (/border-bottom/i.test(cs)) {
                                cs = cs.replace(/border-bottom\s*:[^;]+/i, lastBorderMatch[0]);
                            } else {
                                cs = (cs + ';' + lastBorderMatch[0]).replace(/^;+/, '');
                            }
                            newCell.setAttribute('style', StyleUtils.hexStyle(cs));
                        }
                    }
                }

            } else {
                const firstRowCell = targetGroup[0].cells[cellIdx];
                const lastRowCell  = isMultiRow
                    ? (lastRow.cells[cellIdx - 1] ?? firstRowCell)
                    : firstRowCell;
                StyleUtils.applyCellStyle(newCell, firstRowCell, lastRowCell);
                if (newCell.getAttribute('style')) {
                    newCell.setAttribute('style', StyleUtils.hexStyle(newCell.getAttribute('style')));
                }
                newCell.innerHTML = '&nbsp;';
            }
            newRow.appendChild(newCell);
        });

        fragment.appendChild(newRow);
    }

    targetTbody.appendChild(fragment);

    const shouldLockHeader = (fromDay !== 1);
    window.syncTableToEditor(targetTable, shouldLockHeader);
    if (!shouldLockHeader) window.releaseHeaderLock();

    setTimeout(() => {
        const previewTable = preview.querySelector('table');
        if (!previewTable) return;
        const previewTbody = previewTable.querySelector('tbody') || previewTable;
        const previewRows  = Array.from(previewTbody.querySelectorAll('tr'));
        const allDateTds   = previewRows
            .map(tr => tr.cells[0])
            .filter(td => td && (
                parseInt(td.getAttribute('rowspan') || '1') >= 2 ||
                (td.id && td.id.includes('user_content_')) ||
                /^\d/.test(td.textContent.trim())
            ));
        const targetTd = allDateTds[allDateTds.length - (toDay - fromDay + 1)] || allDateTds.at(-1);
        focusCellInPreview(targetTd);
    }, 80);
};

// ═══════════════════════════════════════════════════════════
//  칸 분할 (splitCurrentRow)
// ═══════════════════════════════════════════════════════════
window.splitCurrentRow = function() {
    const sel = window.getSelection();
    let targetTd = sel?.anchorNode?.parentElement?.closest('td');

    if (!targetTd) {
        const ae = document.activeElement;
        targetTd = ae?.closest('#previewArea td') || (ae?.tagName === 'TD' ? ae : null);
    }

    if (!targetTd && savedRange) {
        const node = savedRange.startContainer;
        const el = node?.nodeType === 3 ? node.parentElement : node;
        targetTd = el?.closest('#previewArea td') || null;
    }
    if (!targetTd && currentTargetNode) {
        targetTd = currentTargetNode.closest?.('td') || null;
        if (targetTd && !document.getElementById('previewArea').contains(targetTd)) {
            targetTd = null;
        }
    }

    if (!targetTd || !document.getElementById('previewArea').contains(targetTd)) {
        return window.showToast('분할할 칸을 선택해주세요.');
    }

    const currentTr   = targetTd.closest('tr');
    const table       = currentTr.closest('table');
    const targetTbody = table.querySelector('tbody') || table;
    const tbodyRows   = Array.from(targetTbody.querySelectorAll('tr'));
    const currentRowIdx = tbodyRows.indexOf(currentTr);

    const templateTable = SampleCache.getTemplateTable();
    if (!templateTable) return window.showToast('샘플 코드가 없습니다.');

    const allTRows = Array.from(templateTable.querySelectorAll('tbody tr')).filter(tr => tr.cells.length > 0);
    const tRows = allTRows.slice(findDataStartIdx(allTRows));
    const sampleGroups = getRowGroups(tRows);
    const refGroup = sampleGroups.find(g => g.length >= 2) || sampleGroups[0];
    const s1 = refGroup[0];
    const sn = refGroup.at(-1);
    const sampleIsMulti = refGroup.length > 1;
    let dateTd = null, startIdx = -1;
    for (let i = currentRowIdx; i >= 0; i--) {
        const firstCell = tbodyRows[i].cells[0];
        if (!firstCell) continue;
        const span    = parseInt(firstCell.getAttribute('rowspan')) || 1;
        const dateInfo = window.parseCellDate(firstCell);
        if (i + span > currentRowIdx && (span > 1 || firstCell.id?.includes('user_content_') || dateInfo)) {
            dateTd = firstCell; startIdx = i; break;
        }
    }
    if (!dateTd) {
        const fc = tbodyRows[currentRowIdx]?.cells[0];
        if (fc) { dateTd = fc; startIdx = currentRowIdx; }
    }
    if (!dateTd) return window.showToast('날짜 칸을 찾을 수 없습니다.');

    const currentSpan = parseInt(dateTd.getAttribute('rowspan')) || 1;
    const groupLastIdx = startIdx + currentSpan - 1; 
    const isFirstRow   = currentRowIdx === startIdx;
    const isLastRow    = currentRowIdx === groupLastIdx;

    const firstRowStyleAttr = tbodyRows[startIdx]?.getAttribute('style') || '';
    const bgMatch = firstRowStyleAttr.match(/background(?:-color)?\s*:\s*([^;]+)/i);
    const rawBg   = bgMatch ? bgMatch[1].trim() : '';
    const hexBg   = rawBg ? (ColorManager.toHex(rawBg) || rawBg) : '';

    if (currentSpan === 1) {
        dateTd.setAttribute('rowspan', '2');

        const s1TrStyle = StyleUtils.hexStyle(s1.closest('tr')?.getAttribute('style') || '');
        currentTr.setAttribute('style', StyleUtils.hexStyle(applyBg(s1TrStyle, hexBg)));

        Array.from(currentTr.cells).forEach((cell, cIdx) => {
            if (cIdx === 0) return; 
            const s1Cell = s1.cells[cIdx];
            if (!s1Cell) return;
            let st = StyleUtils.hexStyle(s1Cell.getAttribute('style') || '');
            cell.setAttribute('style', StyleUtils.hexStyle(applyBg(st, hexBg)));
        });
        const newTr = sn.closest('tr')?.cloneNode(false) || document.createElement('tr');
        let snTrStyle = StyleUtils.hexStyle(sn.closest('tr')?.getAttribute('style') || '');
        newTr.setAttribute('style', StyleUtils.hexStyle(applyBg(snTrStyle, hexBg)));

        const snCells = sampleIsMulti ? Array.from(sn.cells) : Array.from(s1.cells).slice(1);
        snCells.forEach(srcCell => {
            const td = srcCell.cloneNode(false);
            let st = StyleUtils.hexStyle(srcCell.getAttribute('style') || '');
            td.setAttribute('style', StyleUtils.hexStyle(applyBg(st, hexBg)));
            td.innerHTML = '&nbsp;';
            newTr.appendChild(td);
        });
        currentTr.parentNode.insertBefore(newTr, currentTr.nextSibling);
    }

    else if (isLastRow) {
        dateTd.setAttribute('rowspan', currentSpan + 1);
		
		const middleSample = refGroup.length > 2 ? refGroup[1] : s1;
		const middleStyle = StyleUtils.hexStyle(middleSample.closest('tr')?.getAttribute('style') || '');
		currentTr.setAttribute('style', StyleUtils.hexStyle(applyBg(middleStyle, hexBg)));
		
        Array.from(currentTr.cells).forEach((cell, cIdx) => {
			const isFirstRowOfGroup = (currentTr === tbodyRows[startIdx]);
			
			const visualColumnIdx = isFirstRowOfGroup ? cIdx : cIdx + 1;

			const sampleCell = middleSample.cells[visualColumnIdx];
			
			if (sampleCell) {
				let st = StyleUtils.hexStyle(sampleCell.getAttribute('style') || '');
				st = removeRadius(st);
				cell.setAttribute('style', StyleUtils.hexStyle(applyBg(st, hexBg)));
			}
		});

        const newTr = document.createElement('tr');
        let snTrStyle = StyleUtils.hexStyle(sn.closest('tr')?.getAttribute('style') || '');
        newTr.setAttribute('style', StyleUtils.hexStyle(applyBg(snTrStyle, hexBg)));

        const snCells = sampleIsMulti ? Array.from(sn.cells) : Array.from(s1.cells).slice(1);
        snCells.forEach(srcCell => {
            const td = srcCell.cloneNode(false);
            let st = StyleUtils.hexStyle(srcCell.getAttribute('style') || '');
            td.setAttribute('style', StyleUtils.hexStyle(applyBg(st, hexBg)));
            td.innerHTML = '&nbsp;';
            newTr.appendChild(td);
        });
        currentTr.parentNode.insertBefore(newTr, currentTr.nextSibling);
    }

    else {
        dateTd.setAttribute('rowspan', currentSpan + 1);
        const newTr = currentTr.cloneNode(false);
        let trStyle = StyleUtils.hexStyle(currentTr.getAttribute('style') || '');
        newTr.setAttribute('style', StyleUtils.hexStyle(applyBg(trStyle, hexBg)));

        const cellsToClone = isFirstRow
            ? Array.from(currentTr.cells).slice(1) 
            : Array.from(currentTr.cells);   
        cellsToClone.forEach(srcCell => {
            const td = srcCell.cloneNode(false);
            let st = StyleUtils.hexStyle(srcCell.getAttribute('style') || '');
            st = removeRadius(st);
            td.setAttribute('style', StyleUtils.hexStyle(applyBg(st, hexBg)));
            td.innerHTML = '&nbsp;';
            newTr.appendChild(td);
        });
        currentTr.parentNode.insertBefore(newTr, currentTr.nextSibling);

        if (isFirstRow) {
            Array.from(currentTr.cells).slice(1).forEach(cell => {
                let st = StyleUtils.hexStyle(cell.getAttribute('style') || '');
                st = st.replace(/border-bottom-right-radius\s*:[^;]+;?/gi, '')
                       .replace(/border-bottom-left-radius\s*:[^;]+;?/gi, '')
                       .replace(/;+/g, ';').replace(/^;|;$/g, '');
                cell.setAttribute('style', StyleUtils.hexStyle(applyBg(st, hexBg)));
            });
        }
    }

    window.syncTableToEditor(table);

    setTimeout(() => {
        const previewTable = preview.querySelector('table');
        if (!previewTable) return;
        const previewTbody = previewTable.querySelector('tbody') || previewTable;
        const previewRows  = Array.from(previewTbody.querySelectorAll('tr'));
        const newRowInPreview = previewRows[currentRowIdx + 1];
        if (!newRowInPreview) return;
        focusCellInPreview(newRowInPreview.cells[0]);
    }, 80);
};
