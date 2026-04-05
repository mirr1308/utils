/**
 * editor-core.js
 */

let editor;
let preview;
let savedRange = null;
let currentTargetNode = null;
let isSyncing = false;
let syncTimer;

const BEAUTIFY_OPTIONS = {
    indent_size: 4,
    indent_char: ' ',
    indent_inner_html: true,
    wrap_line_length: 0,
    preserve_newlines: true,
    max_preserve_newlines: 1,
    unformatted: ['span', 'a', 'b', 'i', 'br']
};

const DomManager = {
    _parser: new DOMParser(),
    parse(html) {
        if (!html) return null;
        try {
            return this._parser.parseFromString(html, 'text/html');
        } catch(e) {
            return null;
        }
    },
    clone(el) { return el ? el.cloneNode(true) : null; }
};
window.DomManager = DomManager;

window.getStyleKey = function(styles) {
    if (!styles) return '';
    return Object.entries(styles)
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([k, v]) => `${k}:${v}`)
        .join(';');
};

const AppStore = {
    _cache: {},
    get(key) {
        if (this._cache[key] !== undefined) return this._cache[key];
        const raw = localStorage.getItem(key);
        if (raw === null) return null;
        try {
            const parsed = JSON.parse(raw);
            this._cache[key] = parsed;
            return parsed;
        } catch {
            console.warn(`[AppStore] "${key}" 파싱 실패 — 손상된 데이터를 삭제합니다.`);
            localStorage.removeItem(key);
            return null;
        }
    },
    set(key, value) {
        this._cache[key] = value;
        try {
            localStorage.setItem(key, JSON.stringify(value));
        } catch (e) {
            console.warn(`[AppStore] "${key}" 저장 실패:`, e);
        }
    },
    remove(key) {
        delete this._cache[key];
        localStorage.removeItem(key);
    },
    invalidate(key) {
        delete this._cache[key];
    }
};
window.AppStore = AppStore;

function withSyncLock(fn) {
    if (isSyncing) return;
    isSyncing = true;
    try {
        fn();
    } finally {
        setTimeout(() => { isSyncing = false; }, 50);
    }
}
window.withSyncLock = withSyncLock;

function applyEditableMode(html) {
    const hasTable = html.includes('<table') || html.includes('<td');
    if (hasTable) {
        preview.contentEditable = 'false';
        makeEditableOnlyCells(preview);
        preview.querySelectorAll('.preview-line-focus').forEach(el => el.classList.remove('preview-line-focus'));
    } else {
        preview.contentEditable = 'true';
    }
}
window.applyEditableMode = applyEditableMode;

window.syncToEditor = function(rawHtml, { beautify = true, refreshPreview = true } = {}) {
    if (!editor) return;
    const finalHtml = (beautify && typeof html_beautify !== 'undefined')
        ? html_beautify(rawHtml, BEAUTIFY_OPTIONS)
        : rawHtml;

    withSyncLock(() => {
        editor.setValue(finalHtml);
        if (editor.refresh) editor.refresh();
    });

    if (refreshPreview) {
        if (typeof patchPreview === 'function') {
            patchPreview(finalHtml);
        } else {
            preview.innerHTML = finalHtml;
        }
        applyEditableMode(finalHtml);
        if (finalHtml.includes('<table') || finalHtml.includes('<td')) {
            makeEditableOnlyCells(preview);
        }
    }

    if (typeof window.syncPreviewToEditor === 'function') {
        window.syncPreviewToEditor();
    }

    if (window._headerLockRange && typeof window.applyHeaderLock === 'function') {
        requestAnimationFrame(() => window.applyHeaderLock());
    }
};

const ColorManager = {
    rgbToHex(rgb) {
        if (!rgb) return rgb;
        const s = rgb.trim();
        if (/^rgba\s*\(/i.test(s)) return s;  
        const parts = s.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/i);
        if (!parts) return s;
        return '#' + [parts[1], parts[2], parts[3]]
            .map(n => parseInt(n).toString(16).padStart(2, '0'))
            .join('').toLowerCase();
    },

    expandShortHex(s) {
        if (/^#[0-9a-fA-F]{3}$/.test(s))
            return '#' + s[1]+s[1] + s[2]+s[2] + s[3]+s[3];
        return s.toLowerCase();
    },

    toHex(raw) {
        if (!raw) return '';
        const s = raw.trim();
        if (/^rgba\s*\(/i.test(s)) return s;     
        if (/^rgb\s*\(/i.test(s))  return this.rgbToHex(s);
        if (/^#[0-9a-fA-F]{3}$/.test(s)) return this.expandShortHex(s);
        if (/^#[0-9a-fA-F]{6}$/.test(s)) return s.toLowerCase();
        const tmp = document.createElement('div');
        tmp.style.color = s;
        document.body.appendChild(tmp);
        const computed = window.getComputedStyle(tmp).color;
        document.body.removeChild(tmp);
        if (!computed || computed === 'rgba(0, 0, 0, 0)') return s; 
        return computed.startsWith('rgba') ? computed : this.rgbToHex(computed);
    },

    hexifyStyle(styleStr) {
        if (!styleStr || !styleStr.includes('rgb')) return styleStr || '';
        return styleStr.replace(
            /\brgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/gi,
            (m) => this.rgbToHex(m)
        );
    },

    hexifyHtml(html) {
        if (!html || !html.includes('rgb')) return html || '';
        return html.replace(
            /\brgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/gi,
            (m) => this.rgbToHex(m)
        );
    },
};
window.ColorManager = ColorManager;

window.rgbToHex        = (v) => ColorManager.rgbToHex(v);
window.normalizeColor  = (v) => ColorManager.toHex(v);
window.hexifyHtml      = (v) => ColorManager.hexifyHtml(v);
window.normalizeColorToHex = (v) => ColorManager.toHex(v);

window.onload = function () {
    const _exceptions = new Set(['설정 불러오기', '설정 내보내기', '코드 복사', '도움말']);
    const _toolbarTools = Array.from(document.querySelectorAll('.toolbar button, .toolbar select'));
    const _toolbarLockable = _toolbarTools.filter(t => {
        const title = (t.title || '').trim();
        const text  = (t.innerText || '').trim();
        return !_exceptions.has(title) && !_exceptions.has(text);
    });

    function toggleEditTools(disabled) {
        _toolbarLockable.forEach(tool => {
            tool.disabled = disabled;
            tool.classList.toggle('toolbar-disabled', disabled);
        });
    }

    toggleEditTools(true);
    editor = CodeMirror.fromTextArea(document.getElementById('htmlInput'), {
        mode: 'htmlmixed',
        theme: 'neo',
        lineNumbers: true,
        lineWrapping: true,
        gutters: ['CodeMirror-linenumbers', 'markers']
    });

    preview = document.getElementById('previewArea');
    editor.on('focus', () => {
    });

    let _gutterTimer = null;
    editor.on('cursorActivity', () => {
        if (!editor.hasFocus()) return;
        clearTimeout(_gutterTimer);
        _gutterTimer = setTimeout(() => {
            const cursor = editor.getCursor();
            const lineContent = editor.getLine(cursor.line);
            editor.clearGutter('markers');
            if (lineContent && lineContent.includes('<td')) {
                const marker = document.createElement('div');
                marker.className = 'working-marker editor-pos';
                marker.innerHTML = '●';
                editor.setGutterMarker(cursor.line, 'markers', marker);
            }
            if (window._headerLockRange && typeof window.applyHeaderLock === 'function') {
                requestAnimationFrame(() => window.applyHeaderLock());
            }
        }, 50);
    });

    // ── DOM Diffing 헬퍼 ─────────────────────────────────────────────

    function _isSameStructureCell(oldCell, newCell) {
        if (oldCell.attributes.length !== newCell.attributes.length) return false;
        for (const a of newCell.attributes) {
            if (oldCell.getAttribute(a.name) !== a.value) return false;
        }
        const oldNodes = Array.from(oldCell.childNodes);
        const newNodes = Array.from(newCell.childNodes);
        if (oldNodes.length !== newNodes.length) return false;
        return oldNodes.every((n, i) => n.nodeType === newNodes[i].nodeType &&
            (n.nodeType !== 1 || n.tagName === newNodes[i].tagName));
    }

    function _patchTextNodes(oldNode, newNode) {
        const oldChildren = Array.from(oldNode.childNodes);
        const newChildren = Array.from(newNode.childNodes);
        oldChildren.forEach((child, i) => {
            const newChild = newChildren[i];
            if (!newChild) return;
            if (child.nodeType === 3) {
                if (child.nodeValue !== newChild.nodeValue) {
                    child.nodeValue = newChild.nodeValue;
                }
            } else if (child.nodeType === 1) {
                for (const a of newChild.attributes) {
                    if (child.getAttribute(a.name) !== a.value) child.setAttribute(a.name, a.value);
                }
                Array.from(child.attributes).forEach(a => {
                    if (!newChild.hasAttribute(a.name)) child.removeAttribute(a.name);
                });
                _patchTextNodes(child, newChild);
            }
        });
    }

    // ── DOM Diffing 동기화 ────────────────────────────────────────────

    function patchPreview(newHtml) {
        const newDoc = DomManager.parse(newHtml);
        const oldTable = preview.querySelector('table');
        const newTable = newDoc?.querySelector('table');

        if (oldTable && newTable) {
            const oldAttrNames = Array.from(oldTable.attributes).map(a => a.name);
            oldAttrNames.forEach(n => { if (!newTable.hasAttribute(n)) oldTable.removeAttribute(n); });
            Array.from(newTable.attributes).forEach(a => {
                if (oldTable.getAttribute(a.name) !== a.value) oldTable.setAttribute(a.name, a.value);
            });

            const oldThead = oldTable.querySelector('thead');
            const newThead = newTable.querySelector('thead');
            if (newThead) {
                if (!oldThead) {
                    oldTable.insertBefore(newThead.cloneNode(true), oldTable.firstChild);
                } else if (!oldThead.isEqualNode(newThead)) {
                    oldTable.replaceChild(newThead.cloneNode(true), oldThead);
                }
            } else if (oldThead) {
                oldTable.removeChild(oldThead);
            }

            const oldTbody = oldTable.querySelector('tbody');
            const newTbody = newTable.querySelector('tbody');
            if (oldTbody && newTbody) {
                const oldRows = Array.from(oldTbody.rows);
                const newRows = Array.from(newTbody.rows);
                const maxLen  = Math.max(oldRows.length, newRows.length);
                for (let i = 0; i < maxLen; i++) {
                    if (i >= newRows.length) {
                        oldTbody.removeChild(oldRows[i]);
                    } else if (i >= oldRows.length) {
                        oldTbody.appendChild(newRows[i].cloneNode(true));
                    } else if (!oldRows[i].isEqualNode(newRows[i])) {
                        const oldCells = Array.from(oldRows[i].cells);
                        const newCells = Array.from(newRows[i].cells);
                        Array.from(newRows[i].attributes).forEach(a => {
                            if (oldRows[i].getAttribute(a.name) !== a.value)
                                oldRows[i].setAttribute(a.name, a.value);
                        });
                        Array.from(oldRows[i].attributes).forEach(a => {
                            if (!newRows[i].hasAttribute(a.name)) oldRows[i].removeAttribute(a.name);
                        });
                        if (oldCells.length !== newCells.length) {
                            oldTbody.replaceChild(newRows[i].cloneNode(true), oldRows[i]);
                            continue;
                        }
                        for (let j = 0; j < newCells.length; j++) {
                            const oldCell = oldCells[j];
                            const newCell = newCells[j];
                            if (!oldCell || oldCell.isEqualNode(newCell)) continue;

                            if (_isSameStructureCell(oldCell, newCell)) {
                                _patchTextNodes(oldCell, newCell);
                            } else {
                                oldRows[i].replaceChild(newCell.cloneNode(true), oldCell);
                            }
                        }
                    }
                }
            } else if (newTbody && !oldTbody) {
                oldTable.appendChild(newTbody.cloneNode(true));
            } else if (!newTbody && oldTbody) {
                oldTable.removeChild(oldTbody);
            }
            return;
        }

        if (preview.innerHTML !== newHtml) {
            preview.innerHTML = newHtml;
        }
    }
    window.patchPreview = patchPreview;

    let _changeTimer = null;

    function _applyEditorChange() {
        if (isSyncing) return;
        const code = editor.getValue();
        patchPreview(code);
        applyEditableMode(code);
        if (window._headerLockRange && typeof window.applyHeaderLock === 'function') {
            requestAnimationFrame(() => window.applyHeaderLock());
        }
    }

    editor.on('change', (cm, changeObj) => {
        if (document.activeElement && preview.contains(document.activeElement)) return;
        if (isSyncing) return;
        clearTimeout(_changeTimer);

        if (changeObj.origin === 'paste' || changeObj.origin === 'setValue' || !changeObj.origin) {
            _applyEditorChange();
        } else {
            _changeTimer = setTimeout(_applyEditorChange, 60);
        }
    });

    editor.on('blur', () => {
        if (preview.contains(document.activeElement)) return;
        const raw = editor.getValue();
        window.syncToEditor(raw, { beautify: true, refreshPreview: true });
        if (window._headerLockRange && typeof window.applyHeaderLock === 'function') {
            requestAnimationFrame(() => window.applyHeaderLock());
        }
    });

    let previewZoom = 1;
    preview.addEventListener('wheel', (e) => {
        if (!e.ctrlKey) return;
        e.preventDefault();
        const rect = preview.getBoundingClientRect();
        const xPercent = ((e.clientX - rect.left) / rect.width) * 100;
        const yPercent = ((e.clientY - rect.top) / rect.height) * 100;

        const delta = e.deltaY > 0 ? -0.1 : 0.1;
        previewZoom = Math.min(Math.max(0.3, previewZoom + delta), 3);
        
        preview.style.transformOrigin = `${xPercent}% ${yPercent}%`;
        preview.style.transform = `scale(${previewZoom})`;
        preview.style.width = (100 / previewZoom) + '%';
    }, { passive: false });

    preview.addEventListener('input', () => {
        clearTimeout(syncTimer);
        syncTimer = setTimeout(() => {
            if (typeof window.syncPreviewToEditor === 'function') {
                window.syncPreviewToEditor();
            }
        }, 80);
    }, true);

    preview.addEventListener('click', (e) => {
        const hasTable = preview.querySelector('table');
        if (hasTable) {
            const targetTd = e.target.closest('td, th');
            if (targetTd && targetTd.getAttribute('contenteditable') !== 'true') {
                makeEditableOnlyCells(preview);
                setTimeout(() => targetTd.focus(), 0);
            }
            return;
        }
        preview.querySelectorAll('.preview-line-focus').forEach(el => el.classList.remove('preview-line-focus'));
        const block = e.target.closest('p, div, li, h1, h2, h3, h4, h5, h6, blockquote');
        if (block && preview.contains(block)) {
            block.classList.add('preview-line-focus');
        }
    });
	

    preview.addEventListener('focusin', (e) => {
        const targetTd = e.target.closest('td, th');
        if (targetTd) {
            if (targetTd.getAttribute('contenteditable') !== 'true') {
                makeEditableOnlyCells(preview);
                targetTd.focus();
            }
            const html = targetTd.innerHTML.trim();
            const text = targetTd.textContent.trim();
            if (html === '&nbsp;' || text === '\u00A0' || html === '') {
                targetTd.innerHTML = '';
                const range = document.createRange();
                range.selectNodeContents(targetTd);
                range.collapse(true);
                const sel = window.getSelection();
                sel.removeAllRanges();
                sel.addRange(range);
            }
        }
    });
    window.syncPreviewToEditor = function () {
        if (isSyncing) return;
        const tempDiv = preview.cloneNode(true);
        const hasTable = !!tempDiv.querySelector('table'); // 한 번만 체크
        tempDiv.querySelectorAll('.preview-line-focus').forEach(el => el.classList.remove('preview-line-focus'));
        tempDiv.querySelectorAll('[class=""]').forEach(el => el.removeAttribute('class'));

        if (hasTable) {
            tempDiv.querySelectorAll('td, th').forEach((el) => {
                el.removeAttribute('contenteditable');
                el.style.removeProperty('cursor');

                const htmlContent = el.innerHTML.replace(/\s/g, '');
                const textContent = el.textContent.replace(/\s/g, '');
                if (htmlContent === '&nbsp;' || textContent === '\u00A0' || htmlContent === '') {
                    el.innerHTML = '';
                }
            });

            tempDiv.querySelectorAll('[contenteditable], [style*="cursor"]').forEach(el => {
                el.removeAttribute('contenteditable');
                el.style.removeProperty('cursor');
                if (el.style.length === 0) el.removeAttribute('style');
            });
        }

        let rawHtml = ColorManager.hexifyHtml(tempDiv.innerHTML);
        if (hasTable) {
            rawHtml = rawHtml.replace(/&nbsp;|\u00A0/g, '');
        }

        const html = (typeof html_beautify !== 'undefined')
            ? html_beautify(rawHtml, BEAUTIFY_OPTIONS)
            : rawHtml;

        const currentCode = editor.getValue();
        if (currentCode === html) return;

        withSyncLock(() => {
            const oldLines = currentCode.split('\n');
            const newLines = html.split('\n');

            if (oldLines.length === newLines.length) {
                let firstDiff = -1, lastDiff = -1;
                for (let i = 0; i < oldLines.length; i++) {
                    if (oldLines[i] !== newLines[i]) {
                        if (firstDiff === -1) firstDiff = i;
                        lastDiff = i;
                    }
                }
                if (firstDiff !== -1) {
                    const lock = window._headerLockRange;
                    if (lock) {
                        const lockedLines = window._headerLockedLines || [];
                        const lockedSet = new Set(lockedLines);
                        let allLocked = true;
                        for (let i = firstDiff; i <= lastDiff; i++) {
                            if (!lockedSet.has(i)) { allLocked = false; break; }
                        }
                        if (allLocked) return;
                    }
                    const from = { line: firstDiff, ch: 0 };
                    const to   = { line: lastDiff,  ch: oldLines[lastDiff].length };
                    editor.replaceRange(newLines.slice(firstDiff, lastDiff + 1).join('\n'), from, to);
                }
            } else {
                const cursor = editor.getCursor();
                editor.setValue(html);
                editor.setCursor(cursor);
            }
        });

        if (window._headerLockRange && typeof window.applyHeaderLock === 'function') {
            requestAnimationFrame(() => window.applyHeaderLock());
        }
    };

    preview.addEventListener('focus',  () => toggleEditTools(false), true);
    preview.addEventListener('blur', (e) => {
        if (preview.contains(e.relatedTarget)) return;
        if (e.relatedTarget?.closest('.toolbar')) return;
        toggleEditTools(true);
    }, true);

    document.querySelectorAll('.toolbar .icon-btn').forEach(btn => {
        btn.addEventListener('mousedown', (e) => e.preventDefault());
    });

    // ── 툴바 버튼 캐시 ──────────────────────────────────────────────
    window._toolbarBtnCache = {
        bold:      document.querySelector(".icon-btn[onclick*=\"execStyle('bold')\"]"),
        italic:    document.querySelector(".icon-btn[onclick*=\"execStyle('italic')\"]"),
        underline: document.querySelector(".icon-btn[onclick*=\"execStyle('underline')\"]"),
        strike:    document.querySelector(".icon-btn[onclick*=\"execStyle('strike')\"]"),
        alignLeft:   document.querySelector(".icon-btn[onclick*=\"align('left')\"]"),
        alignCenter: document.querySelector(".icon-btn[onclick*=\"align('center')\"]"),
        alignRight:  document.querySelector(".icon-btn[onclick*=\"align('right')\"]"),
    };

    const initialCode = editor.getValue();
    preview.innerHTML = initialCode;
    applyEditableMode(initialCode);
};

function updateToolbarStatus() {
    const selection = window.getSelection();
    if (!selection || !selection.rangeCount) return;
    const range = selection.getRangeAt(0);

    let targets = [];
    if (!range.collapsed) {
        const frag = range.cloneContents();
        const walker = document.createTreeWalker(frag, NodeFilter.SHOW_TEXT, {
            acceptNode: n => n.nodeValue.trim() ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP
        });
        const treeWalker = document.createTreeWalker(
            range.commonAncestorContainer,
            NodeFilter.SHOW_TEXT,
            { acceptNode: n => (range.intersectsNode(n) && n.nodeValue.trim())
                ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP }
        );
        let node;
        while ((node = treeWalker.nextNode())) {
            const el = node.parentElement;
            if (el && !targets.includes(el)) targets.push(el);
        }
    }
    if (targets.length === 0) {
        let t = range.startContainer;
        if (t.nodeType === 3) t = t.parentElement;
        targets = [t];
    }

    const computedCache = new Map();
    const getComputed = (t) => {
        if (!computedCache.has(t)) computedCache.set(t, window.getComputedStyle(t));
        return computedCache.get(t);
    };
    function allHave(checkFn) {
        return targets.length > 0 && targets.every(checkFn);
    }
    const styles = {
        bold:      allHave(t => { const c = getComputed(t); return t.closest('strong,b') !== null || c.fontWeight === 'bold' || parseInt(c.fontWeight) >= 700; }),
        italic:    allHave(t => { const c = getComputed(t); return t.closest('em,i')     !== null || c.fontStyle === 'italic'; }),
        underline: allHave(t => { const c = getComputed(t); return t.closest('u')        !== null || c.textDecoration.includes('underline'); }),
        strike:    allHave(t => { const c = getComputed(t); return t.closest('strike,s,del') !== null || c.textDecoration.includes('line-through'); }),
    };

    const cache = window._toolbarBtnCache || {};
    Object.keys(styles).forEach(type => {
        const btn = cache[type] || document.querySelector(`.icon-btn[onclick*="execStyle('${type}')"]`);
        if (btn) window.setButtonActive(btn, styles[type]);
    });

    const anchorEl = targets[0] || range.startContainer;
    const block = (anchorEl.nodeType === 3 ? anchorEl.parentElement : anchorEl).closest('p, td, th, div');
    const blockStyle = block ? window.getComputedStyle(block) : null;
    let currentAlign = blockStyle ? blockStyle.textAlign : 'left';
    if (currentAlign === 'start' || !currentAlign) currentAlign = 'left';

    const alignMap = { left: cache.alignLeft, center: cache.alignCenter, right: cache.alignRight };
    ['left', 'center', 'right'].forEach(type => {
        const btn = alignMap[type] || document.querySelector(`.icon-btn[onclick*="align('${type}')"]`);
        if (!btn) return;
        window.setButtonActive(btn, currentAlign === type);
    });
}

document.getElementById('previewArea').addEventListener('click',   updateToolbarStatus);
document.getElementById('previewArea').addEventListener('keyup',   updateToolbarStatus);
document.getElementById('previewArea').addEventListener('mouseup', (e) => {
    const ae = document.activeElement;
    if (ae && ae.tagName === 'SELECT') return;
    updateToolbarStatus();
});

function makeEditableOnlyCells(container) {
    if (!container) return;
    container.querySelectorAll('th, td, p').forEach(el => {
        if (el.getAttribute('contenteditable') === 'true') return; // 이미 설정된 셀 스킵
        el.contentEditable = 'true';
        el.style.cursor = 'text';
    });
}
window.makeEditableOnlyCells = makeEditableOnlyCells;

function getSelectionData() {
    const previewArea = document.getElementById('previewArea');
    const sel = window.getSelection();
    if (sel?.anchorNode && previewArea.contains(sel.anchorNode)) {
        const range = sel.getRangeAt(0);
        return { type: 'preview', range, text: range.toString() };
    }
    
    const selectedText = editor.getSelection();
    if (editor.hasFocus() || selectedText) {
        return { type: 'editor', text: selectedText || '', isCollapsed: !selectedText };
    }
    return null;
}

let _toolbarRafId = null;
document.addEventListener('selectionchange', () => {
    const ae = document.activeElement;
    if (ae && (ae.tagName === 'SELECT' || ae.closest('select'))) return;

    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const node = sel.anchorNode;
    if (node?.parentElement?.closest('#previewArea')) {
        savedRange = sel.getRangeAt(0).cloneRange();
        const targetEl = node.nodeType === 3 ? node.parentElement : node;
        currentTargetNode = targetEl.closest('td, p');

        if (_toolbarRafId) cancelAnimationFrame(_toolbarRafId);
        _toolbarRafId = requestAnimationFrame(() => {
            _toolbarRafId = null;
            updateToolbarStatus();
        });
    }
});

function setButtonState(styleType, isActive) {
    const btn = document.querySelector(`.icon-btn[onclick*="'${styleType}'"]`);
    window.setButtonActive(btn, isActive);
}

window.showToast = function(message, type = 'info') {
    const existing = document.getElementById('editor-toast');
    if (existing) existing.remove();
    const toast = document.createElement('div');
    toast.id = 'editor-toast';
    toast.className = `editor-toast editor-toast--${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add('editor-toast--show'));
    setTimeout(() => {
        toast.classList.remove('editor-toast--show');
        setTimeout(() => toast.remove(), 400);
    }, 2500);
};

window.validateHtmlInput = function(html, requiredTag = 'table') {
    if (!html?.trim()) return { ok: false, reason: '내용이 비어있습니다.' };
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const parseError = doc.querySelector('parsererror');
    if (parseError) return { ok: false, reason: '올바르지 않은 HTML 형식입니다.' };
    if (!doc.querySelector(requiredTag)) return { ok: false, reason: `<${requiredTag}> 태그를 찾을 수 없습니다.` };
    return { ok: true };
};

window.setButtonActive = function(btn, isActive) {
    if (!btn) return;
    btn.classList.toggle('active', isActive);
};

window.cleanEl = el => {
    if(!el) return;
    el.removeAttribute('contenteditable');
    el.style.removeProperty('cursor');
    if (el.getAttribute('style') === '') el.removeAttribute('style');
	
};

document.getElementById('previewArea').addEventListener('click', (e) => {
    const targetTd = e.target.closest('td');
    if (!targetTd) return;
    if (targetTd.getAttribute('contenteditable') === 'true') {
        targetTd.focus();
    }
    const allTds = Array.from(document.querySelectorAll('#previewArea td'));
    const tdIndex = allTds.indexOf(targetTd);
    if (tdIndex === -1) return;

    const content  = editor.getValue();
    const tdRegex  = /<td\b/gi;
    let match, matchCount = 0, targetLine = -1;
    while ((match = tdRegex.exec(content)) !== null) {
        if (matchCount === tdIndex) {
            targetLine = content.substring(0, match.index).split('\n').length - 1;
            break;
        }
        matchCount++;
    }
    if (targetLine === -1) return;
	if (targetTd.getAttribute('contenteditable') === 'true') {
        targetTd.focus(); 
    }

    editor.clearGutter('markers');
    const marker = document.createElement('div');
    marker.className = 'working-marker working-marker--pos';
    marker.innerHTML = '●';
    editor.setGutterMarker(targetLine, 'markers', marker);
    editor.scrollIntoView({ line: targetLine, ch: 0 }, 200);
    const lineHandle = editor.addLineClass(targetLine, 'background', 'active-line-highlight');
    setTimeout(() => editor.removeLineClass(lineHandle, 'background', 'active-line-highlight'), 1000);
});

function copyCode() {
    const code = editor.getValue();
    if (!code) { window.showToast('복사할 코드가 없습니다.'); return; }

    let textToCopy = code;
    if (window._headerLockRange) {
        const { trStart, trEnd } = window._headerLockRange;
        const allLines = code.split('\n');
        const trLines = allLines.slice(trStart, trEnd + 1);
        textToCopy = trLines.join('\n').trim();
        if (!textToCopy) {
            window.showToast('복사할 tr 행이 없습니다.');
            return;
        }
    }

    navigator.clipboard.writeText(textToCopy)
        .then(() => window.showToast(
            window._headerLockRange ? '<tr> 행 코드가 복사되었습니다!' : 'HTML 코드가 복사되었습니다!'
        ))
        .catch(() => window.showToast('복사 중 오류가 발생했습니다.'));
}

window.downloadFile = function () {
    if (confirm('현재의 분석 샘플과 커스텀 툴바 설정을 파일(.json)로 백업하시겠습니까?')) {
        window.exportConfigToJson();
    }
};

window.exportConfigToJson = function () {
    const backupData = {
        version: '1.0',
        timestamp: new Date().toISOString(),
        analysisSample: AppStore.get('analysis_source_save') || '',
        customRules:    AppStore.get('custom_toolbar_rules') || []
    };
    const dataUri = 'data:application/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(backupData, null, 4));
    const link = document.createElement('a');
    link.setAttribute('href', dataUri);
    link.setAttribute('download', `editor_backup_${new Date().toISOString().slice(0, 10)}.json`);
    link.click();
};

window.uploadFile = function () {
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.json';
    fileInput.onchange = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            try {
                const data = JSON.parse(ev.target.result);

                if (!data || typeof data !== 'object' || Array.isArray(data)) {
                    throw new Error('올바른 설정 파일 형식이 아닙니다.');
                }
                const hasSample = 'analysisSample' in data;
                const hasRules  = 'customRules' in data;
                if (!hasSample && !hasRules) {
                    throw new Error('analysisSample 또는 customRules 키가 없습니다.');
                }
                if (hasSample && typeof data.analysisSample !== 'string') {
                    throw new Error('analysisSample은 문자열이어야 합니다.');
                }
                if (hasRules && !Array.isArray(data.customRules)) {
                    throw new Error('customRules는 배열이어야 합니다.');
                }

                if (!confirm('설정 파일을 불러오시겠습니까? 현재 샘플코드와 툴바 설정이 덮어씌워집니다.')) return;
                if (hasSample) AppStore.set('analysis_source_save', data.analysisSample);
                if (hasRules)  AppStore.set('custom_toolbar_rules', data.customRules);
                window.showToast('설정을 성공적으로 불러왔습니다. 페이지를 새로고침합니다.');
                location.reload();
            } catch (err) {
                window.showToast('파일을 읽는 중 오류가 발생했습니다: ' + err.message, 'error');
            }
        };
        reader.readAsText(file);
    };
    fileInput.click();
};

window.insertFormattedHtml = function (rawHtml) {
    if (typeof html_beautify === 'undefined') {
        editor.replaceSelection(rawHtml);
        return;
    }
    editor.replaceSelection(html_beautify(rawHtml, BEAUTIFY_OPTIONS) + '\n');
    editor.focus();
};