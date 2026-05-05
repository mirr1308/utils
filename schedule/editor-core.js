/**
 * editor-core.js
 */

// 1. 정의

const CONSTANTS = {
    TOOLBAR_HEIGHT:       50,
    MOBILE_LOGICAL_WIDTH: 375,
    SYNC_LOCK_DELAY:      50,
    EDITOR_CHANGE_DELAY:  150,
    PREVIEW_SYNC_DELAY:   80,
    GUTTER_DELAY:         50,
    USER_CONTENT_PREFIX:  'user_content_',
    NON_LOCKABLE_LABELS: new Set(['설정 불러오기', '설정 내보내기', '코드 복사', '도움말']),

    SELECTORS: {
        MAIN:         '.main-container',
        RIGHT:        '.right-box',
        PREVIEW:      '#previewArea',
        WRAPPER:      '#previewWrapper',
        THEME_TOGGLE: '#themeToggle',
    },

    STRIP_MOBILE_PROPS: [
        'width', 'white-space', 'word-break', 'font-size',
        'padding', 'line-height', 'min-width', 'table-layout', 'margin'
    ],
};
window.CONSTANTS = CONSTANTS;

const BEAUTIFY_OPTIONS = {
    indent_size:           4,
    indent_char:           ' ',
    indent_inner_html:     true,
    wrap_line_length:      0,
    preserve_newlines:     true,
    max_preserve_newlines: 1,
    unformatted: ['span', 'a', 'strong', 'em', 'u', 's', 'b', 'i', 'br', 'th'],
};

// 2. 헬퍼

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
    const allTds    = Array.from(EditorState.get('preview').querySelectorAll('td'));
    const cellIndex = allTds.indexOf(clickedTd);
    if (cellIndex === -1) return;

    const code    = EditorState.get('editor').getValue();
    const tdRegex = /<td\b/gi;
    let match, count = 0, targetLine = -1;
    while ((match = tdRegex.exec(code)) !== null) {
        if (count++ === cellIndex) {
            targetLine = code.substring(0, match.index).split('\n').length - 1;
            break;
        }
    }
    if (targetLine === -1) return;

    EditorState.get('editor').clearGutter('markers');
    EditorState.get('editor').setGutterMarker(targetLine, 'markers', _createGutterMarker('working-marker working-marker--pos'));
    EditorState.get('editor').scrollIntoView({ line: targetLine, ch: 0 }, 200);
    const lineHandle = EditorState.get('editor').addLineClass(targetLine, 'background', 'active-line-highlight');
    setTimeout(() => EditorState.get('editor').removeLineClass(lineHandle, 'background', 'active-line-highlight'), 1000);
}

function _isCellEmpty(cell) {
    const html = cell.innerHTML.replace(/\s/g, '');
    const text = cell.textContent.replace(/\s/g, '');
    return html === '&nbsp;' || text === '\u00A0' || html === '';
}

function _triggerDownload(uri, filename) {
    const link = document.createElement('a');
    link.setAttribute('href', uri);
    link.setAttribute('download', filename);
    link.click();
}

// 3. 데이터 관리

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
        dirtyCell:            null,
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
        const previewEl = this.get('preview');
        if (!previewEl) return;
        if (newHtml.includes('<table')) {
            const check = DomManager.validate(newHtml);
            if (!check.ok) {
                console.warn('[patchPreview] 테이블 구조 오류로 패치 중단:', check.reason);
                return;
            }
        }
        DomPatchManager.patch(previewEl, newHtml);
        if (newHtml.includes('<table') && MobileViewManager._isActive) {
            MobileViewManager._applyCellWidths(previewEl);
            ZoomController.syncAlignment();
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

    clean(element, { stripMobile = false } = {}) {
        if (!element) return;
        element.removeAttribute('contenteditable');
        element.style.removeProperty('cursor');
        if (stripMobile) {
            const mvm = window.MobileViewManager;
            if (mvm?._originalStyles?.has(element)) {
                const backup = mvm._originalStyles.get(element);
                if (backup && element.style.getPropertyPriority('width') === 'important') {
                    const temp = document.createElement('div');
                    temp.setAttribute('style', backup);
                    element.style.width = temp.style.width;
                }
                CONSTANTS.STRIP_MOBILE_PROPS.forEach(prop => {
                    if (element.style.getPropertyPriority(prop) === 'important') {
                        element.style.removeProperty(prop);
                    }
                });
            }
        }
        if (element.tagName === 'TD' || element.tagName === 'TH') {
            element.querySelectorAll('div').forEach(div => {
                const p = document.createElement('p');
                p.innerHTML = div.innerHTML;
                const s = div.getAttribute('style');
                if (s) p.setAttribute('style', s);
                div.replaceWith(p);
            });
            element.querySelectorAll('br + br').forEach(br => br.remove());
        }

        if ((element.tagName === 'TD' || element.tagName === 'TH') && _isCellEmpty(element)) {
            element.innerHTML = '';
        }
        if (element.getAttribute('style') === '' || element.style.length === 0) {
            element.removeAttribute('style');
        }
    },
    clone(element) { return element ? element.cloneNode(true) : null; },
};
window.DomManager = DomManager;

// 4. 조작 구현

function syncAttributes(oldEl, newEl) {
    if (!oldEl || !newEl) return;
    for (const attr of newEl.attributes) {
        if (oldEl.getAttribute(attr.name) !== attr.value) oldEl.setAttribute(attr.name, attr.value);
    }
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
            syncAttributes(oldChild, newChild);
            _patchTextNodes(oldChild, newChild);
        }
    });
}

const DomPatchManager = {
    _cloneAndClean(sectionNode) {
        const cloned = sectionNode.cloneNode(true);
        cloned.querySelectorAll('td, th').forEach(cell => DomManager.clean(cell));
        return cloned;
    },

    patch(oldParent, newHtml) {
        if (!oldParent) return;
        const newBody  = DomManager.parse(newHtml);
        const oldTable = oldParent.querySelector('table');
        const newTable = newBody?.querySelector('table');
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
            oldTable.appendChild(this._cloneAndClean(newTbody));
        } else if (oldTbody) {
            oldTbody.remove();
        }
    },

    _patchSection(table, oldSec, newSec) {
        if (newSec) {
            if (!oldSec) {
                table.insertBefore(this._cloneAndClean(newSec), table.firstChild);
            } else if (!oldSec.isEqualNode(newSec)) {
                table.replaceChild(this._cloneAndClean(newSec), oldSec);
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
                oldTbody.appendChild(this._cloneAndClean(newRows[i]));
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
            parentTbody.replaceChild(this._cloneAndClean(newRow), oldRow);
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

// 5. UI/UX

const ThemeManager = {
    init() {
        const toggle      = document.querySelector(CONSTANTS.SELECTORS.THEME_TOGGLE);
        const previewArea = document.querySelector(CONSTANTS.SELECTORS.PREVIEW);
        const rightBox    = document.querySelector(CONSTANTS.SELECTORS.RIGHT);
        if (!toggle || !previewArea) return;
        toggle.addEventListener('change', () => {
            const isDark = toggle.checked;
            previewArea.classList.toggle('dark-mode', isDark);
            rightBox?.classList.toggle('dark-mode', isDark);
        });
    },
};

const ZoomController = {
    _currentZoom:  1,
    _naturalWidth: null,

    _getEls() {
        return {
            scrollBody: document.getElementById('previewScrollBody'),
            preview:    EditorState.get('preview'),
            wrapper:    EditorState.get('previewWrapper'),
        };
    },

    init() {
        const zoomContainer = document.getElementById('zoomController');
        if (!zoomContainer) return;

        const zoomLevelEl        = document.getElementById('zoomLevel');
        const resetBtn           = document.getElementById('resetZoomBtn');
        const [decBtn, , incBtn] = zoomContainer.querySelectorAll('button');

        const ZOOM_STEP = 0.1;
        const ZOOM_MIN  = 0.3;
        const ZOOM_MAX  = 5.0;

        const applyAndUpdate = (val, clientX, clientY) => {
            const clamped = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, parseFloat(val.toFixed(2))));
            this.apply(clamped, clientX, clientY);
            if (zoomLevelEl) zoomLevelEl.textContent = Math.round(clamped * 100) + '%';
        };

        decBtn?.addEventListener('click',  () => applyAndUpdate(this._currentZoom - ZOOM_STEP));
        incBtn?.addEventListener('click',  () => applyAndUpdate(this._currentZoom + ZOOM_STEP));
        resetBtn?.addEventListener('click', () => applyAndUpdate(1));

        const { scrollBody } = this._getEls();
        const wheelTarget = scrollBody || document.getElementById('previewArea') || document;
        wheelTarget.addEventListener('wheel', (e) => {
            if (!e.ctrlKey) return;
            e.preventDefault();
            const direction = e.deltaY < 0 ? 1 : -1;
            let nextZoom = Math.round((this._currentZoom + (direction * ZOOM_STEP)) * 10) / 10;
            nextZoom = Math.max(ZOOM_MIN, Math.min(nextZoom, ZOOM_MAX));
            applyAndUpdate(nextZoom, e.clientX, e.clientY);
        }, { passive: false });

        this.apply(this._currentZoom);
        window.addEventListener('editor:preview-layout-change', () => this.syncAlignment());
    },

    _clearTransformStyles(preview) {
        ['transform', 'transform-origin', 'margin-top', 'margin-bottom', 'margin-left', 'margin-right']
            .forEach(p => preview.style.removeProperty(p));
    },

    apply(zoomLevel, clientX, clientY) {
        const { scrollBody, preview, wrapper } = this._getEls();
        if (!preview) { this._currentZoom = zoomLevel; return; }

        if (zoomLevel > 1) {
            if (this._naturalWidth === null) {
                this._naturalWidth = scrollBody ? scrollBody.clientWidth : preview.offsetWidth;
                preview.style.width = this._naturalWidth + 'px';
            }

            let ratioX = 0, ratioY = 0;
            if (clientX != null && scrollBody) {
                const rect  = scrollBody.getBoundingClientRect();
                const mouseX = (clientX - rect.left) + scrollBody.scrollLeft;
                const mouseY = (clientY - rect.top)  + scrollBody.scrollTop;
                ratioX = mouseX / (this._naturalWidth * this._currentZoom);
                ratioY = mouseY / (preview.offsetHeight * this._currentZoom);
            }

            this._clearTransformStyles(preview);
            preview.style.zoom = zoomLevel;

            if (clientX != null && scrollBody) {
                const rect = scrollBody.getBoundingClientRect();
                scrollBody.scrollLeft = ratioX * (this._naturalWidth * zoomLevel) - (clientX - rect.left);
                scrollBody.scrollTop  = ratioY * (preview.offsetHeight * zoomLevel) - (clientY - rect.top);
            }
            if (wrapper) wrapper.style.justifyContent = 'flex-start';

        } else if (zoomLevel < 1) {
            if (!EditorState.get('isMobileViewActive') && scrollBody) {
                preview.style.width    = scrollBody.clientWidth + 'px';
                preview.style.minWidth = scrollBody.clientWidth + 'px';
            }
            this._clearTransformStyles(preview);
            preview.style.removeProperty('margin-bottom');
            preview.style.zoom = zoomLevel;

        } else {
            this._naturalWidth = null;
            if (!EditorState.get('isMobileViewActive')) {
                ['width', 'min-width', 'max-width'].forEach(p => preview.style.removeProperty(p));
            }
            this._clearTransformStyles(preview);
            preview.style.removeProperty('zoom');
            if (scrollBody) {
                scrollBody.scrollLeft = 0;
                scrollBody.scrollTop  = 0;
            }
            if (wrapper && !EditorState.get('isMobileViewActive')) {
                wrapper.style.justifyContent = 'center';
            }
        }

        this._currentZoom = zoomLevel;
        this._updateVisibility();
        this.syncAlignment();
    },

    _updateVisibility() {
        const zoomContainer = document.getElementById('zoomController');
        if (!zoomContainer) return;
        zoomContainer.classList.toggle('is-active', Math.abs(this._currentZoom - 1) >= 0.005);
    },

    syncAlignment() {
        const preview = EditorState.get('preview');
        const wrapper = EditorState.get('previewWrapper');
        if (!preview || !wrapper) return;
        wrapper.style.justifyContent =
            (this._currentZoom < 1 || EditorState.get('isMobileViewActive')) ? 'center' : 'flex-start';
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
        this._resizeObserver.disconnect();
        this._resizeObserver = new ResizeObserver(() => this.refreshWidth());
        this._restoreOriginalStyles(preview);
        if (preview) {
            ['width', 'min-width', 'max-width'].forEach(p => preview.style.removeProperty(p));
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
        targets.forEach(cell => {
            const cellWidth = parseFloat(cell.style.width) || parseFloat(cell.getAttribute('width')) || 0;
            if (tableWidth > 0 && cellWidth > 0) {
                cell.style.setProperty('width', (cellWidth / tableWidth * 100).toFixed(4) + '%', 'important');
            }
        });
    },

    stripMobileStyles(container) {
        if (!container) return;
        const table = container.querySelector('table');
        if (table) {
            [table, ...table.querySelectorAll('*')].forEach(el => DomManager.clean(el));
        }
        ['width', 'min-width', 'max-width'].forEach(p => container.style.removeProperty(p));
    },

    _restoreOriginalStyles(container) {
        if (!container) return;
        const table = container.querySelector('table');
        const targets = table ? [table, ...table.querySelectorAll('th, td')] : [];
        targets.forEach(el => {
            if (!this._originalStyles.has(el)) return;
            const original = this._originalStyles.get(el);
            original ? el.setAttribute('style', original) : el.removeAttribute('style');
            this._originalStyles.delete(el);
        });
    },
};

// 6. 동기화&편집

function syncPreviewToEditor({ beautify = true } = {}) {
    if (EditorState.get('isSyncing')) return;
    const preview = EditorState.get('preview');
    const editor  = EditorState.get('editor');
    if (!preview || !editor) return;

    const temp     = preview.cloneNode(true);
    const hasTable = !!temp.querySelector('table');

    temp.querySelectorAll('.preview-line-focus').forEach(el => el.classList.remove('preview-line-focus'));
    temp.querySelectorAll('[class=""]').forEach(el => el.removeAttribute('class'));

    if (hasTable) {
        if (EditorState.get('isMobileViewActive')) MobileViewManager.stripMobileStyles(temp);
        temp.querySelectorAll('td, th, [contenteditable], [style*="cursor"]').forEach(el => DomManager.clean(el));
    }

    const rawHtml = ColorManager.restoreColors(temp.innerHTML);
    const html    = beautify ? safeBeautify(rawHtml) : rawHtml;
    const current = editor.getValue();
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

function updateEditorCell(cell, isFullReplace = false) {
    if (!cell || EditorState.get('isSyncing')) return;

    const allCells  = Array.from(EditorState.get('preview').querySelectorAll('td, th'));
    const cellIndex = allCells.indexOf(cell);
    if (cellIndex === -1) return;

    const code  = EditorState.get('editor').getValue();
    const range = _findNthCellRange(code, cellIndex, cell.tagName.toLowerCase());
    if (!range) return syncPreviewToEditor({ beautify: true });

    const editor = EditorState.get('editor');
    const cloned = cell.cloneNode(true);
    DomManager.clean(cloned);
    const rawInner = ColorManager.restoreColors(cloned.innerHTML);

    const oldHtml  = code.slice(range.start, range.end);
    const charToPos = (idx) => {
        const lines = code.slice(0, idx).split('\n');
        return { line: lines.length - 1, ch: lines[lines.length - 1].length };
    };

    withSyncLock(() => {
        editor.operation(() => {
            if (isFullReplace) {
                const openTag  = oldHtml.slice(0, oldHtml.indexOf('>') + 1);
                const closeTag = oldHtml.slice(oldHtml.lastIndexOf('</'));
                const leading  = oldHtml.match(/^\s*/)[0];
                const trailing = oldHtml.match(/\s*$/)[0];
                const newHtml  = leading + safeBeautify(openTag + rawInner + closeTag).trim() + trailing;
                if (oldHtml === newHtml) return;
                editor.replaceRange(newHtml, charToPos(range.start), charToPos(range.end));
                const startLine = charToPos(range.start).line;
                const lineCount = newHtml.split('\n').length;
                for (let i = 0; i < lineCount; i++) editor.indentLine(startLine + i, 'smart');
            } else {
                const startOff = oldHtml.indexOf('>') + 1;
                const endOff   = oldHtml.lastIndexOf('</');
                editor.replaceRange(rawInner, charToPos(range.start + startOff), charToPos(range.start + endOff));
            }
        });
    });
}

function _findNthCellRange(code, cellIndex, tagName) {
    const openRe  = /<(td|th)(\s[^>]*)?>|<\/(td|th)>/gi;
    let count     = -1;
    let depth     = 0;
    let cellStart = -1;
    let match;

    while ((match = openRe.exec(code)) !== null) {
        const full    = match[0];
        const isClose = full.startsWith('</');
        const tag     = isClose ? match[3] : match[1];

        if (!isClose) {
            depth++;
            if (depth === 1) {
                count++;
                if (count === cellIndex) cellStart = match.index;
            }
        } else {
            if (depth === 1 && count === cellIndex && cellStart !== -1) {
                return { start: cellStart, end: match.index + full.length };
            }
            depth = Math.max(0, depth - 1);
        }
    }
    return null;
}

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
        syncPreviewToEditor({ beautify: false });
    }
    _requestHeaderLockUpdate();
};

window.insertFormattedHtml = function (rawHtml) {
    const editor = EditorState.get('editor');
    editor.replaceSelection(safeBeautify(rawHtml) + '\n');
    editor.focus();
};

// 7. I/O&유틸

function exportConfigToJson() {
    const data = {
        version:        '1.0',
        timestamp:      new Date().toISOString(),
        analysisSample: AppStore.get('analysis_source_save') || '',
        customRules:    AppStore.get('custom_toolbar_rules') || [],
    };
    const uri = 'data:application/json;charset=utf-8,' + encodeURIComponent(JSON.stringify(data, null, 4));
    _triggerDownload(uri, `editor_backup_${new Date().toISOString().slice(0, 10)}.json`);
};

window.downloadFile = function () {
    if (confirm('현재의 샘플코드와 커스텀 툴바 설정을 파일(.json)로 백업하시겠습니까?')) {
        exportConfigToJson();
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

window.copyCode = function () {
    const editor = EditorState.get('editor');
    const code   = editor.getValue();
    if (!code) { window.showToast('복사할 코드가 없습니다.'); return; }

    const headerLockRange = EditorState.get('headerLockRange');
    const useLock         = headerLockRange && !window.isCalendarTable?.();
    const textToCopy      = useLock
        ? code.split('\n').slice(headerLockRange.trStart, headerLockRange.trEnd + 1).join('\n').trim()
        : code;

    if (useLock && !textToCopy) { window.showToast('복사할 tr 행이 없습니다.'); return; }

    navigator.clipboard.writeText(textToCopy)
        .then(() => window.showToast(useLock ? '<tr> 행 코드가 복사되었습니다!' : 'HTML 코드가 복사되었습니다!'))
        .catch(() => window.showToast('복사 중 오류가 발생했습니다.'));
};

window.showToast = function (message, type = 'info') {
    document.getElementById('editor-toast')?.remove();
    const toast       = document.createElement('div');
    toast.id          = 'editor-toast';
    toast.className   = `editor-toast editor-toast--${type}`;
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

// 8. 메인 실행

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
        mode:         'htmlmixed',
        theme:        'neo',
        lineNumbers:  true,
        lineWrapping: true,
        gutters:      ['CodeMirror-linenumbers', 'markers'],
    });
    const previewEl = document.getElementById('previewArea');
    EditorState.set('editor',  editor);
    EditorState.set('preview', previewEl);

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
        if (document.activeElement && previewEl.contains(document.activeElement)) return;
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
        if (previewEl.contains(document.activeElement)) return;
        window.syncToEditor(editor.getValue(), { beautify: true, refreshPreview: true });
        _requestHeaderLockUpdate();
    });

    const preview = previewEl;

    let _isComposing = false;
    preview.addEventListener('compositionstart', () => { _isComposing = true;  }, true);
    preview.addEventListener('compositionend',   () => {
        _isComposing = false;
        const cell = document.activeElement?.closest('td, th');
        if (cell) {
            EditorState.set('dirtyCell', cell);
            updateEditorCell(cell, false);
        } else {
            window.syncToEditor(preview.innerHTML, { beautify: true, refreshPreview: false });
        }
    }, true);

    preview.addEventListener('keydown', (e) => {
        if (e.key !== 'Enter') return;
        const cell = e.target.closest('td, th');
        if (!cell) return;
        e.preventDefault();
        const sel = window.getSelection();
        if (!sel?.rangeCount) return;
        const range = sel.getRangeAt(0);
        range.deleteContents();
        const br = document.createElement('br');
        range.insertNode(br);
        range.setStartAfter(br);
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
        EditorState.set('dirtyCell', cell);
        updateEditorCell(cell, false);
    }, true);

    preview.addEventListener('paste', (e) => {
        const cell = e.target.closest('td, th');
        if (!cell) return;
        e.preventDefault();
        const text = e.clipboardData.getData('text/plain');
        if (!text) return;
        const sel = window.getSelection();
        if (!sel?.rangeCount) return;
        const range = sel.getRangeAt(0);
        range.deleteContents();
        const lines = text.split(/\r?\n/);
        lines.forEach((line, i) => {
            if (i > 0) range.insertNode(document.createElement('br'));
            if (line) range.insertNode(document.createTextNode(line));
        });
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
        EditorState.set('dirtyCell', cell);
        updateEditorCell(cell, false);
    }, true);

    preview.addEventListener('input', (e) => {
        if (_isComposing) return;
        clearTimeout(EditorState.get('syncTimer'));
        EditorState.set('syncTimer', setTimeout(() => {
            const cell = e.target.closest('td, th');
            if (cell) {
                EditorState.set('dirtyCell', cell);
                updateEditorCell(cell, false); // 실시간 모드
            } else {
                window.syncToEditor(preview.innerHTML, { beautify: true, refreshPreview: false });
            }
        }, CONSTANTS.PREVIEW_SYNC_DELAY));
    }, true);

    preview.addEventListener('focusin', (e) => {
        const cell = e.target.closest('td, th');
        if (!cell) return;
        if (cell.getAttribute('contenteditable') !== 'true') {
            makeEditableOnlyCells(preview);
            cell.focus();
        }
        EditorState.set('dirtyCell', cell);
        if (_isCellEmpty(cell)) {
            withSyncLock(() => { cell.innerHTML = ''; });
            const range = document.createRange();
            range.selectNodeContents(cell);
            range.collapse(true);
            const sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(range);
        }
    });

    preview.addEventListener('click', (e) => {
        const hasTable    = !!preview.querySelector('table');
        const clickedCell = e.target.closest('td, th');
        const clickedTd   = e.target.closest('td');

        if (hasTable) {
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
    preview.addEventListener('mouseup', () => {
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
        const leavingCell  = e.target.closest('td, th');
        const enteringCell = e.relatedTarget?.closest('td, th');

        if (leavingCell && preview.contains(e.relatedTarget)) {
            if (leavingCell !== enteringCell) {
                updateEditorCell(EditorState.get('dirtyCell') || leavingCell, true);
                EditorState.set('dirtyCell', null);
            }
            return;
        }

        if (e.relatedTarget?.closest('.toolbar')) return;
        toggleEditTools(true);

        const dirty = EditorState.get('dirtyCell');
        if (dirty) {
            updateEditorCell(dirty, true);
            EditorState.set('dirtyCell', null);
        }
        setTimeout(() => syncPreviewToEditor({ beautify: true }), 50);
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
