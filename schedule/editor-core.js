/**
 * editor-core.js
 */

window.editor  = null;
window.preview = null;

const CONSTANTS = {
    TOOLBAR_HEIGHT:       50,
    MOBILE_LOGICAL_WIDTH: 375,
    SYNC_LOCK_DELAY:      50,
    EDITOR_CHANGE_DELAY:  150,
    PREVIEW_SYNC_DELAY:   80,
    GUTTER_DELAY:         50,
    USER_CONTENT_PREFIX:  'user_content_',
    MOBILE_STYLE_ID:      'mobile-view-override',

    NON_LOCKABLE_LABELS: new Set(['설정 불러오기', '설정 내보내기', '코드 복사', '도움말']),

    SELECTORS: {
        MAIN:         '.main-container',
        RIGHT:        '.right-box',
        PREVIEW:      '#previewArea',
        WRAPPER:      '#previewWrapper',
        THEME_TOGGLE: '#themeToggle',
    },
    MOBILE_TABLE_CSS: [
        '.main-container.is-mobile-mode #previewArea { padding: 10px 1px; }',
        '.main-container.is-mobile-mode #previewArea table { width: 100% !important; min-width: 0 !important; table-layout: fixed !important; margin: 0 !important; border-collapse: collapse; border-spacing: 0; }',
        '.main-container.is-mobile-mode #previewArea th, .main-container.is-mobile-mode #previewArea td { padding: 0.5em 0.5em !important; }',
        '.main-container.is-mobile-mode #previewArea :is(th, td, span, p) { font-size: 0.6rem !important; white-space: normal !important; word-break: break-all !important; line-height: 1.5; }'
    ].join('\n'),
	STRIP_MOBILE_PROPS: [
        'width', 'white-space', 'word-break', 'font-size', 
        'padding', 'line-height', 'min-width', 'table-layout', 'margin'
    ],
};
window.CONSTANTS = CONSTANTS;

const BEAUTIFY_OPTIONS = {
    indent_size:          4,
    indent_char:          ' ',
    indent_inner_html:    true,
    wrap_line_length:     0,
    preserve_newlines:    true,
    max_preserve_newlines: 1,
    unformatted: ['span', 'a', 'strong', 'em', 'u', 's', 'b', 'i', 'br'],
};


function safeBeautify(html) {
    return typeof html_beautify !== 'undefined'
        ? html_beautify(html, BEAUTIFY_OPTIONS)
        : html;
}

function getResolvedNode(node) {
    return node?.nodeType === Node.TEXT_NODE ? node.parentElement : node;
}
window.getResolvedNode = getResolvedNode;

function withSyncLock(fn) {
    if (EditorState.get('isSyncing')) return;
    EditorState.startSync();
    try {
        fn();
    } finally {
        EditorState.endSync();
    }
}
window.withSyncLock = withSyncLock;

function _requestHeaderLockUpdate() {
    if (EditorState.get('headerLockRange') && typeof window.applyHeaderLock === 'function') {
        requestAnimationFrame(() => window.applyHeaderLock());
    }
}

function _createGutterMarker(className) {
    const el = document.createElement('div');
    el.className = className;
    el.innerHTML = '●';
    return el;
}

function _highlightEditorLineForCell(clickedTd) {
    const allTds       = Array.from(window.preview.querySelectorAll('td'));
    const cellIndex    = allTds.indexOf(clickedTd);
    if (cellIndex === -1) return;

    const code      = window.editor.getValue();
    const tdRegex   = /<td\b/gi;
    let match, count = 0, targetLine = -1;
    while ((match = tdRegex.exec(code)) !== null) {
        if (count++ === cellIndex) {
            targetLine = code.substring(0, match.index).split('\n').length - 1;
            break;
        }
    }
    if (targetLine === -1) return;

    window.editor.clearGutter('markers');
    window.editor.setGutterMarker(targetLine, 'markers', _createGutterMarker('working-marker working-marker--pos'));
    window.editor.scrollIntoView({ line: targetLine, ch: 0 }, 200);
    const lineHandle = window.editor.addLineClass(targetLine, 'background', 'active-line-highlight');
    setTimeout(() => window.editor.removeLineClass(lineHandle, 'background', 'active-line-highlight'), 1000);
}


const EditorState = {
    _data: {
        previewWrapper:       null,
        scrollBody:           null,
        savedRange:           null,
        currentTargetNode:    null,
        isSyncing:            false,
        syncTimer:            null,
        isMobileViewActive:   false,
        mobileOriginalStyles: new WeakMap(),
        headerLockRange:      null,
        headerLockedLines:    [],
        editor:               null,
        preview:              null,
    },

    get(key) {
        if (!(key in this._data)) {
            console.warn(`[EditorState] 존재하지 않는 키: ${key}`);
        }
        return this._data[key];
    },

    set(key, value) {
        this._data[key] = value;
    },

    startSync() { this._data.isSyncing = true; },

    endSync(immediate = false) {
        if (immediate) {
            this._data.isSyncing = false;
        } else {
            setTimeout(() => { this._data.isSyncing = false; }, CONSTANTS.SYNC_LOCK_DELAY);
        }
    },

    patchPreview(newHtml) {
        const previewEl = window.preview;
        if (!previewEl) return;
        if (newHtml.includes('<table')) {
            if (typeof DomManager !== 'undefined' && DomManager.validate) {
                const check = DomManager.validate(newHtml);
                if (!check.ok) {
                    console.warn('[patchPreview] 테이블 구조 오류로 패치 중단:', check.reason);
                    return;
                }
            }
        }
        DomPatchManager.patch(previewEl, newHtml);
        if (newHtml.includes('<table') && typeof MobileViewManager !== 'undefined' && MobileViewManager._isActive) {
            MobileViewManager._applyCellWidths(previewEl);
            if (typeof ZoomController !== 'undefined') ZoomController.syncAlignment();
        }
    },
};
window.EditorState = EditorState;

const AppStore = {
    _cache:     {},
    _listeners: {},

    subscribe(key, callback) {
        if (!this._listeners[key]) this._listeners[key] = [];
        this._listeners[key].push(callback);
    },

    get(key) {
        if (this._cache[key] !== undefined) return this._cache[key];
        const raw = localStorage.getItem(key);
        if (raw === null) return null;
        try {
            const parsed = JSON.parse(raw);
            this._cache[key] = parsed;
            return parsed;
        } catch {
            localStorage.removeItem(key);
            return null;
        }
    },

    set(key, value) {
        this._cache[key] = value;
        try {
            localStorage.setItem(key, JSON.stringify(value));
        } catch (error) {
            console.warn(`[AppStore] "${key}" 저장 실패:`, error);
            window.showToast?.('저장 공간이 부족합니다. 일부 설정이 저장되지 않을 수 있습니다.', 'error');
        }
        this._listeners[key]?.forEach(cb => cb(value));
    },

    remove(key) {
        delete this._cache[key];
        localStorage.removeItem(key);
    },

    invalidate(key) {
        delete this._cache[key];
    },
};
window.AppStore = AppStore;

const ColorManager = {
    rgbToHex(rgbString) {
        if (!rgbString?.includes('rgb')) return rgbString;
        const parts = rgbString.match(/\d+/g);
        if (!parts || parts.length < 3) return rgbString;
        return '#' + parts.slice(0, 3)
            .map(n => parseInt(n).toString(16).padStart(2, '0'))
            .join('')
            .toLowerCase();
    },

    toOriginalForm(value) {
        if (!value) return '';
        const t = value.trim().toLowerCase();
        if (t.startsWith('#') || t.startsWith('rgba')) return t;
        if (t.startsWith('rgb')) return this.rgbToHex(t);
        return t;
    },

    restoreColors(text) {
        if (!text?.includes('rgb')) return text || '';
        return text.replace(
            /\brgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/gi,
            (match) => this.rgbToHex(match)
        );
    },

    normValue(value, property) {
        value = (value || '').trim();
        if (!value) return value;
        return (property && (property.includes('color') || property.includes('background')))
            ? this.toOriginalForm(value)
            : value.toLowerCase();
    },

    normFont(fontFamily) {
        if (!fontFamily) return '';
        return fontFamily.split(',').map(f => `'${f.trim().replace(/['"]/g, '')}'`).join(', ');
    },

    parseStyleObj(styleString) {
        const el = document.createElement('div');
        el.setAttribute('style', styleString || '');
        return el.style;
    },

    parseStyleString(styleString) {
        const result = {};
        const parsed = this.parseStyleObj(styleString);
        for (let i = 0; i < parsed.length; i++) {
            const prop  = parsed[i];
            const value = this.normValue(parsed.getPropertyValue(prop), prop);
            if (value) result[prop] = value;
        }
        return result;
    },

    serializeStyle(styleObject) {
        const entries = Object.entries(styleObject);
        if (!entries.length) return '';
        return entries.map(([p, v]) => `${p}:${v}`).join(';') + ';';
    },
};
window.ColorManager = ColorManager;

const DomManager = {
    _parser: new DOMParser(),

    parse(html) {
        if (!html) return null;
        try { return this._parser.parseFromString(html, 'text/html'); }
        catch { return null; }
    },

    validate(html, requiredTag = 'table') {
        if (!html?.trim()) return { ok: false, reason: '내용이 비어 있습니다.' };
        const doc = this.parse(html);
        if (!doc) return { ok: false, reason: '올바르지 않은 HTML 형식입니다.' };
        if (doc.querySelector('parsererror')) return { ok: false, reason: '올바르지 않은 HTML 형식입니다.' };
        if (!doc.querySelector(requiredTag)) return { ok: false, reason: `<${requiredTag}> 태그를 찾을 수 없습니다.` };
        return { ok: true };
    },

    clean(element) {
		if (!element) return;
		element.removeAttribute('contenteditable');
		element.style.removeProperty('cursor');
		const mvm = window.MobileViewManager;
		const originalStyleStr = mvm?._originalStyles?.get(element);
		if (originalStyleStr !== undefined) {
			if (element.style.getPropertyPriority('width') === 'important') {
				const backup = mvm._originalStyles.get(element);
				if (backup) {
					const temp = document.createElement('div');
					temp.setAttribute('style', backup);
					element.style.width = temp.style.width; 
				}
			}
			const stripProps = window.CONSTANTS?.STRIP_MOBILE_PROPS || [];
			stripProps.forEach(prop => {
				if (element.style.getPropertyPriority(prop) === 'important') {
					element.style.removeProperty(prop);
				}
			});
		}
		if (element.tagName === 'TD' || element.tagName === 'TH') {
			const html = element.innerHTML.replace(/\s/g, '');
			const text = element.textContent.replace(/\s/g, '');
			if (html === '&nbsp;' || text === '\u00A0' || html === '') element.innerHTML = '';
		}
		if (element.getAttribute('style') === '' || element.style.length === 0) {
			element.removeAttribute('style');
		}
	},

    clone(element) { return element ? element.cloneNode(true) : null; },
};
window.DomManager = DomManager;


function syncAttributes(oldEl, newEl) {
    if (!oldEl || !newEl) return;
    Array.from(newEl.attributes).forEach(attr => {
        if (oldEl.getAttribute(attr.name) !== attr.value) oldEl.setAttribute(attr.name, attr.value);
    });
    Array.from(oldEl.attributes).forEach(attr => {
        if (!newEl.hasAttribute(attr.name)) oldEl.removeAttribute(attr.name);
    });
}

function _isSameStructureCell(oldCell, newCell) {
    if (oldCell.attributes.length !== newCell.attributes.length) return false;
    for (const attr of newCell.attributes) {
        if (oldCell.getAttribute(attr.name) !== attr.value) return false;
    }
    const oldChildren = Array.from(oldCell.childNodes);
    const newChildren = Array.from(newCell.childNodes);
    if (oldChildren.length !== newChildren.length) return false;
    return oldChildren.every((node, i) =>
        node.nodeType === newChildren[i].nodeType &&
        (node.nodeType !== 1 || node.tagName === newChildren[i].tagName)
    );
}

function _patchTextNodes(oldNode, newNode) {
    const oldChildren = Array.from(oldNode.childNodes);
    const newChildren = Array.from(newNode.childNodes);
    oldChildren.forEach((oldChild, i) => {
        const newChild = newChildren[i];
        if (!newChild) return;
        if (oldChild.nodeType === 3) {
            if (oldChild.nodeValue !== newChild.nodeValue) oldChild.nodeValue = newChild.nodeValue;
        } else if (oldChild.nodeType === 1) {
            for (const attr of newChild.attributes) {
                if (oldChild.getAttribute(attr.name) !== attr.value) oldChild.setAttribute(attr.name, attr.value);
            }
            Array.from(oldChild.attributes).forEach(attr => {
                if (!newChild.hasAttribute(attr.name)) oldChild.removeAttribute(attr.name);
            });
            _patchTextNodes(oldChild, newChild);
        }
    });
}

const DomPatchManager = {
    patch(oldParent, newHtml) {
        if (!oldParent) return;
        const newBody    = DomManager.parse(newHtml);
        const oldTable   = oldParent.querySelector('table');
        const newTable   = newBody?.querySelector('table');
        if (oldTable && newTable) {
            this._patchTable(oldTable, newTable);
            return;
        }
        if (oldParent.innerHTML !== newHtml) oldParent.innerHTML = newHtml;
    },

    _patchTable(oldTable, newTable) {
        syncAttributes(oldTable, newTable);
        this._patchSection(oldTable, oldTable.querySelector('thead'), newTable.querySelector('thead'));

        const oldTbody = oldTable.querySelector('tbody');
        const newTbody = newTable.querySelector('tbody');
        if (oldTbody && newTbody) {
            this._patchRows(oldTbody, newTbody);
        } else if (newTbody) {
            const cloned = newTbody.cloneNode(true);
            cloned.querySelectorAll('td, th').forEach(cell => DomManager.clean(cell));
            oldTable.appendChild(cloned);
        } else if (oldTbody) {
            oldTbody.remove();
        }
    },

    _patchSection(table, oldSec, newSec) {
        if (newSec) {
            if (!oldSec) {
                const cloned = newSec.cloneNode(true);
                cloned.querySelectorAll('td, th').forEach(cell => DomManager.clean(cell));
                table.insertBefore(cloned, table.firstChild);
            } else if (!oldSec.isEqualNode(newSec)) {
                const cloned = newSec.cloneNode(true);
                cloned.querySelectorAll('td, th').forEach(cell => DomManager.clean(cell));
                table.replaceChild(cloned, oldSec);
            }
        } else if (oldSec) {
            oldSec.remove();
        }
    },

    _patchRows(oldTbody, newTbody) {
        const oldRows = Array.from(oldTbody.rows);
        const newRows = Array.from(newTbody.rows);
        const maxRows = Math.max(oldRows.length, newRows.length);

        for (let i = 0; i < maxRows; i++) {
            if (i >= newRows.length) {
                oldTbody.removeChild(oldRows[i]);
            } else if (i >= oldRows.length) {
                const cloned = newRows[i].cloneNode(true);
                cloned.querySelectorAll('td, th').forEach(cell => DomManager.clean(cell));
                oldTbody.appendChild(cloned);
            } else if (!oldRows[i].isEqualNode(newRows[i])) {
                this._patchCells(oldRows[i], newRows[i], oldTbody);
            }
        }
    },

    _patchCells(oldRow, newRow, parentTbody) {
        syncAttributes(oldRow, newRow);
        const oldCells = Array.from(oldRow.cells);
        const newCells = Array.from(newRow.cells);

        if (oldCells.length !== newCells.length) {
            const cloned = newRow.cloneNode(true);
            cloned.querySelectorAll('td, th').forEach(cell => DomManager.clean(cell));
            parentTbody.replaceChild(cloned, oldRow);
            return;
        }

        for (let j = 0; j < newCells.length; j++) {
            const oldCell = oldCells[j];
            const newCell = newCells[j];
            if (!oldCell || oldCell.isEqualNode(newCell)) continue;

            if (_isSameStructureCell(oldCell, newCell)) {
                syncAttributes(oldCell, newCell);
                _patchTextNodes(oldCell, newCell);
            } else {
                const cloned = newCell.cloneNode(true);
                DomManager.clean(cloned);
                oldRow.replaceChild(cloned, oldCell);
            }
        }
    },
};


const ThemeManager = {
    init() {
        const toggle     = document.querySelector(CONSTANTS.SELECTORS.THEME_TOGGLE);
        const previewArea = document.querySelector(CONSTANTS.SELECTORS.PREVIEW);
        const rightBox   = document.querySelector(CONSTANTS.SELECTORS.RIGHT);
        if (!toggle || !previewArea) return;
        toggle.addEventListener('change', () => {
            const isDark = toggle.checked;
            previewArea.classList.toggle('dark-mode', isDark);
            rightBox?.classList.toggle('dark-mode', isDark);
        });
    },
};

const ZoomController = {
    _currentZoom: 1,

    init() {
        const zoomContainer = document.getElementById('zoomController');
        if (!zoomContainer) return;

        const zoomLevelEl   = document.getElementById('zoomLevel');
        const resetBtn      = document.getElementById('resetZoomBtn');
        const [decBtn, , incBtn] = zoomContainer.querySelectorAll('button');

        const ZOOM_STEP = 0.1;
        const ZOOM_MIN  = 0.3;
        const ZOOM_MAX  = 3.0;

        const applyAndUpdate = (val) => {
            const clamped = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, parseFloat(val.toFixed(2))));
            this.apply(clamped);
            if (zoomLevelEl) zoomLevelEl.textContent = Math.round(clamped * 100) + '%';
        };

        decBtn?.addEventListener('click', () => applyAndUpdate(this._currentZoom - ZOOM_STEP));
        incBtn?.addEventListener('click', () => applyAndUpdate(this._currentZoom + ZOOM_STEP));
        resetBtn?.addEventListener('click', () => applyAndUpdate(1));

        this.apply(this._currentZoom);
        window.addEventListener('editor:preview-layout-change', () => this.syncAlignment());
    },

    apply(zoomLevel) {
        this._currentZoom = zoomLevel;
        const preview = EditorState.get('preview');
        if (preview) preview.style.zoom = zoomLevel;
        this.syncAlignment();
    },

    syncAlignment() {
        const preview        = EditorState.get('preview');
        const previewWrapper = EditorState.get('previewWrapper');
        if (!preview || !previewWrapper) return;

        const isMobile       = EditorState.get('isMobileViewActive');
        const needsCenter    = isMobile || (preview.offsetWidth * this._currentZoom < previewWrapper.clientWidth);
        previewWrapper.style.justifyContent = needsCenter ? 'center' : 'flex-start';
    },

    getCurrentZoom() { return this._currentZoom; },
};

const MobileViewManager = {
    _isActive:       false,
    _originalStyles: new WeakMap(),
    _resizeObserver: null,

    init() {
        this._resizeObserver = new ResizeObserver(() => this.refreshWidth());
        document.querySelectorAll('input[name="viewMode"]').forEach(input => {
            input.addEventListener('change', (e) => {
                e.target.id === 'mobileView' ? this.enable() : this.disable();
                requestAnimationFrame(() => this._dispatchUpdate());
            });
        });
    },

    _dispatchUpdate() {
        window.dispatchEvent(new CustomEvent('editor:preview-layout-change'));
    },

    _getEls() {
        return {
            main:    document.querySelector(CONSTANTS.SELECTORS.MAIN),
            right:   document.querySelector(CONSTANTS.SELECTORS.RIGHT),
            preview: document.querySelector(CONSTANTS.SELECTORS.PREVIEW),
            wrapper: document.querySelector(CONSTANTS.SELECTORS.WRAPPER),
        };
    },

    enable() {
        this._isActive = true;
        const { main, right, preview, wrapper } = this._getEls();
        main?.classList.add('is-mobile-mode');
        if (wrapper) wrapper.style.justifyContent = 'center';
        this._injectMobileCss();
        this._applyCellWidths(preview);
        if (right) this._resizeObserver.observe(right);
        this.refreshWidth();
        EditorState.set('isMobileViewActive', true);
        this._dispatchUpdate();
    },

    disable() {
        this._isActive = false;
        const { main, right, preview } = this._getEls();
        main?.classList.remove('is-mobile-mode');
        if (right) this._resizeObserver.unobserve(right);
        this._removeMobileCss();
        this._restoreOriginalStyles(preview);
        if (preview) {
            preview.style.removeProperty('width');
            preview.style.removeProperty('min-width');
            preview.style.removeProperty('max-width');
        }
        EditorState.set('isMobileViewActive', false);
        this._dispatchUpdate();
    },

    refreshWidth() {
        if (!this._isActive) return;
        const { right, preview } = this._getEls();
        if (!right || !preview) return;

        const mobileWidth = Math.min(CONSTANTS.MOBILE_LOGICAL_WIDTH, right.clientWidth - 20) + 'px';
        ['width', 'minWidth', 'maxWidth'].forEach(prop => { preview.style[prop] = mobileWidth; });
        requestAnimationFrame(() => this._dispatchUpdate());
    },

    _injectMobileCss() {
        let styleEl = document.getElementById(CONSTANTS.MOBILE_STYLE_ID);
        if (!styleEl) {
            styleEl = document.createElement('style');
            styleEl.id = CONSTANTS.MOBILE_STYLE_ID;
            document.head.appendChild(styleEl);
        }
        styleEl.textContent = CONSTANTS.MOBILE_TABLE_CSS;
    },

    _removeMobileCss() {
        document.getElementById(CONSTANTS.MOBILE_STYLE_ID)?.remove();
    },

    _applyCellWidths(container) {
        const table = container?.querySelector('table');
        if (!table) return;
		const targets = table.querySelectorAll('th, td');
        [table, ...targets].forEach(el => {
            if (!this._originalStyles.has(el)) {
                this._originalStyles.set(el, el.getAttribute('style') || '');
            }
        });
        const tableWidth = parseFloat(table.style.width) || parseFloat(table.getAttribute('width')) || 0;
        table.querySelectorAll('th, td').forEach(cell => {
            if (!this._originalStyles.has(cell)) this._originalStyles.set(cell, cell.getAttribute('style') || '');
            const cellWidth = parseFloat(cell.style.width) || parseFloat(cell.getAttribute('width')) || 0;
            if (tableWidth > 0 && cellWidth > 0) {
				cell.style.setProperty('width', (cellWidth / tableWidth * 100).toFixed(4) + '%', 'important');
			}
            cell.style.setProperty('white-space', 'normal', 'important');
        });
    },
	stripMobileStyles(container) {
		if (!container) return;
		const table = container.querySelector('table');
		if (table) {
			const targets = table.querySelectorAll('*');
			[table, ...targets].forEach(el => window.DomManager.clean(el));
		}
		['width', 'min-width', 'max-width'].forEach(p => container.style.removeProperty(p));
	},
    _restoreOriginalStyles(container) {
        container?.querySelectorAll('th, td').forEach(cell => {
            if (!this._originalStyles.has(cell)) return;
            const original = this._originalStyles.get(cell);
            original ? cell.setAttribute('style', original) : cell.removeAttribute('style');
            this._originalStyles.delete(cell);
        });
    },
};


function syncPreviewToEditor() {
    if (EditorState.get('isSyncing')) return;
    const preview = EditorState.get('preview');
    const editor  = EditorState.get('editor');
    if (!preview || !editor) return;

    const temp     = preview.cloneNode(true);
    const hasTable = !!temp.querySelector('table');

    temp.querySelectorAll('.preview-line-focus').forEach(el => el.classList.remove('preview-line-focus'));
    temp.querySelectorAll('[class=""]').forEach(el => el.removeAttribute('class'));

    if (EditorState.get('isMobileViewActive') && hasTable) {
		MobileViewManager.stripMobileStyles(temp);
	}
    if (hasTable) {
        temp.querySelectorAll('td, th, [contenteditable], [style*="cursor"]').forEach(el => DomManager.clean(el));
    }

    const rawHtml  = ColorManager.restoreColors(temp.innerHTML);
    const html     = safeBeautify(rawHtml);
    const current  = editor.getValue();
    if (current === html) return;

    withSyncLock(() => {
        const oldLines = current.split('\n');
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
                const lock = EditorState.get('headerLockRange');
                if (lock) {
                    const lockedSet = new Set(EditorState.get('headerLockedLines') || []);
                    let allLocked = true;
                    for (let i = firstDiff; i <= lastDiff; i++) {
                        if (!lockedSet.has(i)) { allLocked = false; break; }
                    }
                    if (allLocked) return;
                }
                editor.replaceRange(
                    newLines.slice(firstDiff, lastDiff + 1).join('\n'),
                    { line: firstDiff, ch: 0 },
                    { line: lastDiff,  ch: oldLines[lastDiff].length }
                );
            }
        } else {
            const cursor = editor.getCursor();
            editor.setValue(html);
            editor.setCursor(cursor);
        }
    });

    _requestHeaderLockUpdate();
}
window.syncPreviewToEditor = syncPreviewToEditor;

function applyEditableMode(html) {
    const preview = EditorState.get('preview');
    if (!preview) return;
    const hasTable = html.includes('<table') || html.includes('<td');
    if (hasTable) {
        preview.contentEditable = 'false';
        makeEditableOnlyCells(preview);
        preview.querySelectorAll('.preview-line-focus').forEach(el => el.classList.remove('preview-line-focus'));
    } else {
        preview.contentEditable = 'true';
    }
}

function makeEditableOnlyCells(container) {
    if (!container) return;
    container.querySelectorAll('th, td, p').forEach(el => {
        if (el.getAttribute('contenteditable') === 'true') return;
        el.contentEditable = 'true';
        el.style.cursor    = 'text';
    });
}
window.makeEditableOnlyCells = makeEditableOnlyCells;

window.syncToEditor = function (rawHtml, { beautify = true, refreshPreview = true } = {}) {
    const editor = EditorState.get('editor');
    if (!editor) return;
    const finalHtml = beautify ? safeBeautify(rawHtml) : rawHtml;

    withSyncLock(() => {
        editor.setValue(finalHtml);
        editor.refresh?.();
    });

    if (refreshPreview) {
        EditorState.patchPreview(finalHtml);
        applyEditableMode(finalHtml);
    }
    syncPreviewToEditor();
    _requestHeaderLockUpdate();
};

window.insertFormattedHtml = function (rawHtml) {
    const editor = EditorState.get('editor');
    editor.replaceSelection(safeBeautify(rawHtml) + '\n');
    editor.focus();
};



window.exportConfigToJson = function () {
    const data = {
        version:        '1.0',
        timestamp:      new Date().toISOString(),
        analysisSample: AppStore.get('analysis_source_save') || '',
        customRules:    AppStore.get('custom_toolbar_rules') || [],
    };
    const uri  = 'data:application/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(data, null, 4));
    const link = document.createElement('a');
    link.setAttribute('href', uri);
    link.setAttribute('download', `editor_backup_${new Date().toISOString().slice(0, 10)}.json`);
    link.click();
};

window.downloadFile = function () {
    if (confirm('현재의 샘플코드와 커스텀 툴바 설정을 파일(.json)로 백업하시겠습니까?')) {
        window.exportConfigToJson();
    }
};

window.uploadFile = function () {
    const fileInput  = document.createElement('input');
    fileInput.type   = 'file';
    fileInput.accept = '.json';
    fileInput.onchange = (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = (ev) => {
            try {
                const data = JSON.parse(ev.target.result);
                if (!data || typeof data !== 'object' || Array.isArray(data)) throw new Error('올바른 설정 파일 형식이 아닙니다.');
                const hasSample = 'analysisSample' in data;
                const hasRules  = 'customRules' in data;
                if (!hasSample && !hasRules) throw new Error('analysisSample 또는 customRules 키가 없습니다.');
                if (hasSample && typeof data.analysisSample !== 'string') throw new Error('analysisSample은 문자열이어야 합니다.');
                if (hasRules  && !Array.isArray(data.customRules))        throw new Error('customRules는 배열이어야 합니다.');
                if (!confirm('설정 파일을 불러오시겠습니까? 현재 샘플코드와 툴바 설정을 덮어씁니다.')) return;
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

function copyCode() {
    const editor = EditorState.get('editor');
    const code   = editor.getValue();
    if (!code) { window.showToast('복사할 코드가 없습니다.'); return; }

    const headerLockRange = EditorState.get('headerLockRange');
    const useLock         = headerLockRange && !window.isCalendarTable?.();
    let textToCopy        = code;

    if (useLock) {
        const { trStart, trEnd } = headerLockRange;
        textToCopy = code.split('\n').slice(trStart, trEnd + 1).join('\n').trim();
        if (!textToCopy) { window.showToast('복사할 tr 행이 없습니다.'); return; }
    }

    navigator.clipboard.writeText(textToCopy)
        .then(() => window.showToast(useLock ? '<tr> 행 코드가 복사되었습니다!' : 'HTML 코드가 복사되었습니다!'))
        .catch(() => window.showToast('복사 중 오류가 발생했습니다.'));
}

window.showToast = function (message, type = 'info') {
    document.getElementById('editor-toast')?.remove();
    const toast     = document.createElement('div');
    toast.id        = 'editor-toast';
    toast.className = `editor-toast editor-toast--${type}`;
    toast.textContent = message;
    document.body.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add('editor-toast--show'));
    setTimeout(() => {
        toast.classList.remove('editor-toast--show');
        setTimeout(() => toast.remove(), 400);
    }, 2500);
};

window.setButtonActive = function (button, isActive) {
    button?.classList.toggle('active', isActive);
};


window.onload = function () {
    window.isCalendarTable = window.isCalendarTable ?? (() => false);
    window.applyHeaderLock = window.applyHeaderLock ?? (() => {});

    const _lockableItems = Array.from(document.querySelectorAll('.toolbar button, .toolbar select'))
        .filter(el => !CONSTANTS.NON_LOCKABLE_LABELS.has((el.title || '').trim())
                   && !CONSTANTS.NON_LOCKABLE_LABELS.has((el.innerText || '').trim()));

    function toggleEditTools(disabled) {
        _lockableItems.forEach(el => {
            el.disabled = disabled;
            el.classList.toggle('toolbar-disabled', disabled);
        });
    }
    toggleEditTools(true);
	
    const editor = CodeMirror.fromTextArea(document.getElementById('htmlInput'), {
        mode:        'htmlmixed',
        theme:       'neo',
        lineNumbers: true,
        lineWrapping: true,
        gutters:     ['CodeMirror-linenumbers', 'markers'],
    });
    window.editor = editor;
    window.preview = document.getElementById('previewArea');

    EditorState.set('editor',  editor);
    EditorState.set('preview', window.preview);

    const rightBox = window.preview.closest('.right-box');
    let scrollBody    = document.getElementById('previewScrollBody');
    let previewWrapper = document.getElementById('previewWrapper');

    if (!scrollBody) {
        const paneHeader = rightBox.querySelector('.pane-header');
        scrollBody = Object.assign(document.createElement('div'), {
            id:       'previewScrollBody',
            style:    {},
        });
        scrollBody.style.cssText = 'flex:1;overflow:auto;position:relative;';

        previewWrapper = Object.assign(document.createElement('div'), { id: 'previewWrapper' });
        previewWrapper.style.cssText = 'display:flex;justify-content:center;align-items:flex-start;min-height:100%;';

        window.preview.parentNode.removeChild(window.preview);
        previewWrapper.appendChild(window.preview);
        scrollBody.appendChild(previewWrapper);

        const insertTarget = paneHeader?.parentNode === rightBox ? paneHeader.nextSibling : null;
        rightBox.insertBefore(scrollBody, insertTarget);
    }

    EditorState.set('previewWrapper', document.getElementById('previewWrapper'));
    EditorState.set('scrollBody',     document.getElementById('previewScrollBody'));

    let _gutterTimer = null;
    editor.on('cursorActivity', () => {
        if (!editor.hasFocus()) return;
        clearTimeout(_gutterTimer);
        _gutterTimer = setTimeout(() => {
            const cursor      = editor.getCursor();
            const lineContent = editor.getLine(cursor.line);
            editor.clearGutter('markers');
            if (lineContent?.includes('<td')) {
                editor.setGutterMarker(cursor.line, 'markers', _createGutterMarker('working-marker editor-pos'));
            }
            _requestHeaderLockUpdate();
        }, CONSTANTS.GUTTER_DELAY);
    });

    let _changeTimer = null;

    function _applyEditorChange() {
        if (EditorState.get('isSyncing')) return;
        const code = editor.getValue();
        EditorState.patchPreview(code);
        applyEditableMode(code);
        _requestHeaderLockUpdate();
        requestAnimationFrame(() => ZoomController.syncAlignment());
    }

    editor.on('change', (cm, change) => {
        if (document.activeElement && window.preview.contains(document.activeElement)) return;
        if (EditorState.get('isSyncing')) return;
        clearTimeout(_changeTimer);
        const isImmediate = change.origin === 'paste'
            || change.origin === 'setValue'
            || change.origin === '+delete'
            || !change.origin;
        if (isImmediate) _applyEditorChange();
        else _changeTimer = setTimeout(_applyEditorChange, CONSTANTS.EDITOR_CHANGE_DELAY);
    });

    editor.on('blur', () => {
        if (window.preview.contains(document.activeElement)) return;
        window.syncToEditor(editor.getValue(), { beautify: true, refreshPreview: true });
        _requestHeaderLockUpdate();
    });

    const preview = window.preview;

    preview.addEventListener('input', () => {
        clearTimeout(EditorState.get('syncTimer'));
        EditorState.set('syncTimer', setTimeout(syncPreviewToEditor, CONSTANTS.PREVIEW_SYNC_DELAY));
    }, true);

    preview.addEventListener('focusin', (e) => {
        const cell = e.target.closest('td, th');
        if (!cell) return;
        if (cell.getAttribute('contenteditable') !== 'true') {
            makeEditableOnlyCells(preview);
            cell.focus();
        }
        const html = cell.innerHTML.trim();
        const text = cell.textContent.trim();
        if (html === '&nbsp;' || text === '\u00A0' || html === '') {
            cell.innerHTML = '';
            const range = document.createRange();
            range.selectNodeContents(cell);
            range.collapse(true);
            const sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(range);
        }
    });

    preview.addEventListener('click', (e) => {
        const hasTable = !!preview.querySelector('table');
        if (hasTable) {
            const clickedCell = e.target.closest('td, th');
            const clickedTd   = e.target.closest('td');

            if (clickedCell && clickedCell.getAttribute('contenteditable') !== 'true') {
                makeEditableOnlyCells(preview);
                setTimeout(() => clickedCell.focus(), 0);
            }
            if (clickedTd) _highlightEditorLineForCell(clickedTd);
        } else {
            preview.querySelectorAll('.preview-line-focus').forEach(el => el.classList.remove('preview-line-focus'));
            const block = e.target.closest('p, div, li, h1, h2, h3, h4, h5, h6, blockquote');
            if (block && preview.contains(block)) block.classList.add('preview-line-focus');
        }
        TextEditor.updateToolbarStatus();
    });

    preview.addEventListener('keyup',   TextEditor.updateToolbarStatus);
    preview.addEventListener('mouseup', (e) => {
        if (document.activeElement?.tagName === 'SELECT') return;
        TextEditor.updateToolbarStatus();
    });

    let _rafId = null;
    document.addEventListener('selectionchange', () => {
        const active = document.activeElement;
        if (active && (active.tagName === 'SELECT' || active.closest('select'))) return;

        const sel = window.getSelection();
        if (!sel?.rangeCount) return;
        const anchor = sel.anchorNode;
        if (anchor?.parentElement?.closest('#previewArea')) {
            EditorState.set('savedRange', sel.getRangeAt(0));
            EditorState.currentTargetNode = getResolvedNode(anchor).closest('td, p');
            if (_rafId) cancelAnimationFrame(_rafId);
            _rafId = requestAnimationFrame(() => {
                _rafId = null;
                TextEditor.updateToolbarStatus();
            });
        }
    });

    preview.addEventListener('focus', () => toggleEditTools(false), true);
    preview.addEventListener('blur', (e) => {
        if (preview.contains(e.relatedTarget)) return;
        if (e.relatedTarget?.closest('.toolbar')) return;
        toggleEditTools(true);
    }, true);

    document.querySelectorAll('.toolbar .icon-btn').forEach(btn => {
        btn.addEventListener('mousedown', (e) => e.preventDefault());
    });

    const initialCode = editor.getValue();
    preview.innerHTML = initialCode;
    applyEditableMode(initialCode);

    window.initToolbarCache?.();
    ThemeManager.init();
    MobileViewManager.init();
    ZoomController.init();
};
