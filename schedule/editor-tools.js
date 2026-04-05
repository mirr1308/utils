/**
 * editor-tools.js
 */

function normalizeColor(color) {
    return ColorManager.toHex(color);
}

const SEMANTIC_TAG_MAP = {
    fontWeight:     { tag: 'strong', check: v => v === 'bold' || parseInt(v) >= 700 },
    fontStyle:      { tag: 'em',     check: v => v === 'italic' },
    textDecorationUnderline: { tag: 'u', check: v => v && v.includes('underline') },
    textDecorationLineThrough: { tag: 's', check: v => v && v.includes('line-through') },
};

function buildCharMap(pEl) {
    const charMap = [];
    function collectStyles(el) {
        const s = el.style;
        const r = {};
        if (s.fontWeight)      r.fontWeight      = s.fontWeight;
        if (s.fontStyle)       r.fontStyle       = s.fontStyle;
        if (s.textDecoration)  r.textDecoration  = s.textDecoration;
        if (s.color)           r.color           = normalizeColor(s.color);
        if (s.backgroundColor) r.backgroundColor = normalizeColor(s.backgroundColor);
        if (s.fontFamily)      r.fontFamily      = s.fontFamily;
        if (s.fontSize)        r.fontSize        = s.fontSize;
        const tag = el.tagName?.toUpperCase();
        if (tag === 'STRONG' && !r.fontWeight)     r.fontWeight = 'bold';
        if (tag === 'EM'     && !r.fontStyle)      r.fontStyle  = 'italic';
        if (tag === 'U'      && !r.textDecoration) r.textDecoration = 'underline';
        if (tag === 'S'      && !r.textDecoration) r.textDecoration = 'line-through';
        return r;
    }
    function walk(node, inherited) {
        if (node.nodeType === 3) {
            for (const ch of node.textContent) {
                charMap.push({ char: ch, styles: Object.assign({}, inherited) });
            }
        } else if (node.nodeType === 1) {
            if (node.tagName === 'BR') {
                charMap.push({ char: '\n', isBr: true, styles: Object.assign({}, inherited) });
                return;
            }
            let own = {};
            if (['SPAN','STRONG','EM','U','S'].includes(node.tagName)) own = collectStyles(node);
            for (const child of node.childNodes) walk(child, Object.assign({}, inherited, own));
        }
    }
    walk(pEl, collectStyles(pEl));
    return charMap;
}

function rebuildParagraph(pEl, charMap, isFullSelection) {
    const pAlign = pEl.style.textAlign;

    pEl.innerHTML = '';
    pEl.removeAttribute('class');
    pEl.removeAttribute('style');
    if (pAlign) pEl.style.textAlign = pAlign;

    let i = 0;
    while (i < charMap.length) {
        if (charMap[i].isBr) {
            pEl.appendChild(document.createElement('br'));
            i++;
            continue;
        }
        const curKey = getStyleKey(charMap[i].styles);
        let j = i + 1;
        while (j < charMap.length && !charMap[j].isBr && getStyleKey(charMap[j].styles) === curKey) j++;
        const text = charMap.slice(i, j).map(c => c.char).join('');
        const styles = charMap[i].styles;

        if (!Object.keys(styles).length) {
            pEl.appendChild(document.createTextNode(text));
        } else {
            pEl.appendChild(buildSemanticNode(text, styles));
        }
        i = j;
    }
}

function buildSemanticNode(text, styles) {
    const spanStyles = {};
    const semanticTags = [];

    const td = styles.textDecoration || '';
    if (td.includes('underline'))    semanticTags.push('u');
    if (td.includes('line-through')) semanticTags.push('s');
    if (styles.fontWeight === 'bold' || parseInt(styles.fontWeight) >= 700) semanticTags.push('strong');
    if (styles.fontStyle === 'italic') semanticTags.push('em');

    if (styles.color)           spanStyles.color           = styles.color;
    if (styles.backgroundColor) spanStyles.backgroundColor = styles.backgroundColor;
    if (styles.fontFamily)      spanStyles.fontFamily      = styles.fontFamily;
    if (styles.fontSize)        spanStyles.fontSize        = styles.fontSize;

    let innerNode;
    if (Object.keys(spanStyles).length > 0) {
        const span = document.createElement('span');
        const styleParts = [];
        if (spanStyles.color)           styleParts.push(`color:${spanStyles.color}`);
        if (spanStyles.backgroundColor) styleParts.push(`background-color:${spanStyles.backgroundColor}`);
        if (spanStyles.fontFamily) {
            const fontVal = spanStyles.fontFamily.split(',').map(s => {
                const name = s.trim().replace(/&quot;|['"]/g, '').trim();
                return name ? `'${name}'` : '';
            }).filter(Boolean).join(', ');
            styleParts.push(`font-family:${fontVal}`);
        }
        if (spanStyles.fontSize)        styleParts.push(`font-size:${spanStyles.fontSize}`);
        span.setAttribute('style', styleParts.join(';'));
        span.textContent = text;
        innerNode = span;
    } else {
        innerNode = document.createTextNode(text);
    }

    let current = innerNode;
    for (let k = semanticTags.length - 1; k >= 0; k--) {
        const wrapper = document.createElement(semanticTags[k]);
        wrapper.appendChild(current);
        current = wrapper;
    }
    return current;
}

function getSelectionCharRange(pEl, range) {
    let startIdx = -1, endIdx = -1, charCount = 0;
    function walk(node) {
        if (node.nodeType === 3) {
            let offset = 0;
            for (const ch of node.textContent) {
                if (node === range.startContainer && offset === range.startOffset) startIdx = charCount;
                if (node === range.endContainer   && offset === range.endOffset)   endIdx   = charCount;
                charCount++;
                offset += ch.length;
            }
            if (node === range.endContainer   && offset === range.endOffset)   endIdx   = charCount;
            if (node === range.startContainer && offset === range.startOffset) startIdx = charCount;
        } else if (node.nodeType === 1) {
            if (node.tagName === 'BR') { charCount++; return; }
            for (const child of node.childNodes) walk(child);
        }
    }
    walk(pEl);
    if (startIdx === -1) startIdx = 0;
    if (endIdx   === -1) endIdx   = charCount;
    return { startIdx, endIdx, total: charCount };
}

function restoreSelection(sel, pEl, startIdx, endIdx, isFullSelection) {
    try {
        const r = document.createRange();
        if (isFullSelection) {
            r.selectNodeContents(pEl);
            sel.removeAllRanges();
            sel.addRange(r);
            return;
        }
        function findNode(targetIdx) {
            let count = 0;
            function walk(node) {
                if (node.nodeType === 3) {
                    const len = node.textContent.length;
                    if (count + len > targetIdx) return { node, offset: targetIdx - count };
                    count += len;
                    return null;
                } else if (node.nodeType === 1) {
                    if (node.tagName === 'BR') { count++; return null; }
                    for (const child of node.childNodes) {
                        const res = walk(child);
                        if (res) return res;
                    }
                }
                return null;
            }
            return walk(pEl);
        }
        const s = findNode(startIdx);
        const e = findNode(endIdx);
        if (s && e) { r.setStart(s.node, s.offset); r.setEnd(e.node, e.offset); }
        else r.selectNodeContents(pEl);
        sel.removeAllRanges();
        sel.addRange(r);
    } catch (err) {
    }
}

const styleMap = {
    font: { 
        key: 'fontFamily', 
        val: (v) => v, 
        css: (v) => {
            const parts = v.split(',').map(s => {
                const name = s.trim().replace(/&quot;|['"]/g, '').trim();
                return name ? `'${name}'` : '';
            }).filter(Boolean);
            return `font-family: ${parts.join(', ')};`;
        }
    },
    size:      { key: 'fontSize',        val: (v) => v,                   css: (v) => `font-size: ${v};` },
    bold:      { key: 'fontWeight',      val: () => 'bold',               css: () => 'font-weight: bold;' },
    italic:    { key: 'fontStyle',       val: () => 'italic',             css: () => 'font-style: italic;' },
    underline: { key: 'textDecoration',  val: () => 'underline',          css: () => 'text-decoration: underline;' },
    strike:    { key: 'textDecoration',  val: () => 'line-through',       css: () => 'text-decoration: line-through;' },
    align:     { key: 'textAlign',       val: (v) => v,                   css: (v) => `text-align: ${v};` },
    text:      { key: 'color',           val: (v) => normalizeColor(v),   css: (v) => `color: ${v};` },
    bg:        { key: 'backgroundColor', val: (v) => normalizeColor(v),   css: (v) => `background-color: ${v};` },
};

function applyStyle(styleType, value) {
    let data = getSelectionData();
    if (!data && savedRange) data = { type: 'preview', range: savedRange, text: savedRange.toString() };
    if (!data) return;

    const sm = styleMap[styleType];
    if (!sm) return;

    isSyncing = true;
    try {
        if (data.type === 'preview') {
            const sel = window.getSelection();
            if (!sel || !sel.rangeCount) return;
            const range = sel.getRangeAt(0);
            if (!range.toString().trim()) return;

            const startEl = range.startContainer.nodeType === 3
                ? range.startContainer.parentElement
                : range.startContainer;
            let parentP = startEl.closest('p');

            if (!parentP) {
                const parentTd = startEl.closest('td');
                if (parentTd) {
                    if (styleType === 'align') { isSyncing = false; return; }
                    _applyToTd(parentTd, range, sm, styleType, value, sel);
                    return;
                }
            }
            if (!parentP) {
                parentP = document.createElement('p');
                range.surroundContents(parentP);
            }
            if (styleType === 'align') {
                _applyAlign(parentP, sm.val(value), sel);
                return;
            }
            _applyToP(parentP, range, sm, styleType, value, sel);

        } else if (data.type === 'editor') {
            const pure = data.text.replace(/<\/?[^>]+(>|$)/g, '');
            let wrapped;
            if (styleType === 'bold')      wrapped = `<strong>${pure}</strong>`;
            else if (styleType === 'italic')    wrapped = `<em>${pure}</em>`;
            else if (styleType === 'underline') wrapped = `<u>${pure}</u>`;
            else if (styleType === 'strike')    wrapped = `<s>${pure}</s>`;
            else wrapped = `<span style="${sm.css(value)}">${pure}</span>`;
            editor.replaceSelection(wrapped);
        }
    } catch (e) {
    } finally {
        isSyncing = false;
        if (data.type === 'editor') editor.focus();
        if (typeof window.syncPreviewToEditor === 'function') window.syncPreviewToEditor();
        if (typeof updateToolbarStatus === 'function') updateToolbarStatus();
        if (window._headerLockRange && typeof window.applyHeaderLock === 'function') {
            requestAnimationFrame(() => window.applyHeaderLock());
        }
    }
}

function _applyToP(parentP, range, sm, styleType, value, sel) {
    const charMap = buildCharMap(parentP);
    const { startIdx, endIdx, total } = getSelectionCharRange(parentP, range);
    const selected = charMap.slice(startIdx, endIdx);
    const isColor  = styleType === 'text' || styleType === 'bg';
    const isActive = !isColor && selected.length > 0 && selected.every(c => {
        const v = c.styles[sm.key];
        if (!v) return false;
        if (sm.key === 'fontWeight') return v === 'bold' || parseInt(v) >= 700;
        return v === sm.val(value);
    });
    const isFull = startIdx === 0 && endIdx === total && endIdx - startIdx === total;

    for (let i = startIdx; i < endIdx; i++) {
        if (isColor) charMap[i].styles[sm.key] = sm.val(value);
        else if (isActive) delete charMap[i].styles[sm.key];
        else charMap[i].styles[sm.key] = sm.val(value);
    }
    rebuildParagraph(parentP, charMap, isFull);
    setButtonState(styleType, isColor ? true : !isActive);
    restoreSelection(sel, parentP, startIdx, endIdx, isFull);
}

function _applyToTd(parentTd, range, sm, styleType, value, sel) {
    try {
        const charMap = buildCharMap(parentTd);
        const { startIdx, endIdx, total } = getSelectionCharRange(parentTd, range);
        const selected = charMap.slice(startIdx, endIdx).filter(c => !c.isBr);
        const isColor  = styleType === 'text' || styleType === 'bg';
        const isActive = !isColor && selected.length > 0 && selected.every(c => {
            const v = c.styles[sm.key];
            if (!v) return false;
            if (sm.key === 'fontWeight') return v === 'bold' || parseInt(v) >= 700;
            return v === sm.val(value);
        });
        const hasBr = charMap.some(c => c.isBr);
        const isFull = !hasBr && startIdx === 0 && endIdx === total;

        for (let i = startIdx; i < endIdx; i++) {
            if (charMap[i].isBr) continue; 
            if (isColor) charMap[i].styles[sm.key] = sm.val(value);
            else if (isActive) delete charMap[i].styles[sm.key];
            else charMap[i].styles[sm.key] = sm.val(value);
        }


        const tdStyle = parentTd.getAttribute('style') || '';
        rebuildParagraph(parentTd, charMap, isFull);
        if (tdStyle) parentTd.setAttribute('style', tdStyle);

        restoreSelection(sel, parentTd, startIdx, endIdx, isFull);
    } catch (e) {
    } finally {
        isSyncing = false;
        if (typeof window.syncPreviewToEditor === 'function') window.syncPreviewToEditor();
        if (typeof updateToolbarStatus === 'function') updateToolbarStatus();
        if (window._headerLockRange && typeof window.applyHeaderLock === 'function') {
            requestAnimationFrame(() => window.applyHeaderLock());
        }
    }
}

function _applyAlign(parentP, styleVal, sel) {
    const cur = parentP.style.textAlign;
    parentP.style.textAlign = (cur === styleVal || styleVal === 'left') ? '' : styleVal;
    if (parentP.getAttribute('style') === '') parentP.removeAttribute('style');
    const next = parentP.style.textAlign || 'left';
    ['left', 'center', 'right'].forEach(t => {
        const b = document.querySelector(`.icon-btn[onclick*="align('${t}')"]`);
        if (!b) return;
        const on = t !== 'left' && next === t;
        window.setButtonActive(b, on);
    });
    const fr = document.createRange();
    fr.selectNodeContents(parentP);
    sel.removeAllRanges();
    sel.addRange(fr);
}

/* 공개 헬퍼 */

let _fontDebounceTimer = null;
const changeFontFamily = (font) => {
    if (!font) return;
    clearTimeout(_fontDebounceTimer);
    _fontDebounceTimer = setTimeout(() => applyStyle('font', font), 0);
};

let _sizeDebounceTimer = null;
const changeFontSize = (size) => {
    if (!size) return;
    clearTimeout(_sizeDebounceTimer);
    _sizeDebounceTimer = setTimeout(() => applyStyle('size', size), 0);
};

const execStyle = (type) => applyStyle(type);
const align     = (type) => applyStyle('align', type);

/* 컬러피커*/
function openColor(mode, btn, event) {
    if (event) { event.stopPropagation(); event.preventDefault(); }
    if (window.activePicker) return;

    const sel = window.getSelection();
    let localRange = null;
    let startColor = mode === 'text' ? '#000000' : '#ffffff';

    if (sel && sel.rangeCount > 0) {
        localRange = sel.getRangeAt(0).cloneRange();
        const node = localRange.startContainer.nodeType === 3
            ? localRange.startContainer.parentElement
            : localRange.startContainer;
        const raw = window.getComputedStyle(node)[mode === 'text' ? 'color' : 'backgroundColor'];
        startColor = normalizeColor(raw) || startColor;
    }

    const picker = new Picker({
        parent: btn,
        popup: 'bottom',
        alpha: false,
        editor: true,
        color: startColor,
        onClose: () => {
            if (window.activePicker) { window.activePicker.destroy(); window.activePicker = null; }
        },
        onDone: (color) => {
            const finalHex = color.hex.slice(0, 7).toLowerCase();
            if (localRange) {
                const s = window.getSelection();
                s.removeAllRanges();
                s.addRange(localRange);
                applyStyle(mode, finalHex);
            }
            if (window.activePicker) { window.activePicker.destroy(); window.activePicker = null; }
        },
    });
	picker.show(); 
    window.activePicker = picker;
}

/* 하이퍼링크*/
window.prepareLinkData = function () {
    const sel = window.getSelection();
    if (!sel || !sel.rangeCount) return;
    const range = sel.getRangeAt(0);
	
    const startNode = range.startContainer;
    const parentA   = (startNode.nodeType === 3 ? startNode.parentElement : startNode).closest('a');
    const parentTd  = (startNode.nodeType === 3 ? startNode.parentElement : startNode).closest('td');

	const clone = range.cloneContents();
    const tempDiv = document.createElement('div');
    tempDiv.appendChild(clone);
	const imgTag = tempDiv.querySelector('img');
	
	const textInput = document.getElementById('modalTextDisplay');
    const previewContainer = document.getElementById('linkImgPreview');
    const previewImg = document.getElementById('modalPreviewImg');
	
	if (imgTag) {
        textInput.value = imgTag.src; 
		const previewImg = document.getElementById('modalPreviewImg');
        if (previewImg) {
            previewImg.src = imgTag.src;
            document.getElementById('linkImgPreview').style.display = 'block';
        }
        textInput.setAttribute('data-is-img', 'true');
        textInput.setAttribute('data-img-style', imgTag.style.cssText);

    } else {
        textInput.value = range.toString().trim();
        textInput.removeAttribute('data-is-img');
        if (document.getElementById('linkImgPreview')) {
            document.getElementById('linkImgPreview').style.display = 'none';
        }
    }

    document.getElementById('modalLinkHref').value    = parentA?.getAttribute('href') || '';
    document.getElementById('modalTdId').value        = parentTd?.id?.replace('user_content_', '') || '';
	
    if (parentA) {
        document.getElementById('modalTargetBlank').checked = (parentA.getAttribute('target') || '_blank') === '_blank';
        document.getElementById('modalUnderline').checked   = parentA.style.textDecoration === 'none';
    } else {
        document.getElementById('modalTargetBlank').checked = true;
        document.getElementById('modalUnderline').checked   = true;
    }
    window.savedModalRange = range.cloneRange();
};

window.applyLinkChanges = function () {
    const textInput = document.getElementById('modalTextDisplay');
    const newText   = textInput.value;
    const isImg     = textInput.getAttribute('data-is-img') === 'true';
    const href      = document.getElementById('modalLinkHref').value.trim();
    const tdIdRaw   = document.getElementById('modalTdId').value.trim();
    const tdId      = tdIdRaw ? 'user_content_' + tdIdRaw : '';
    if (!window.savedModalRange) return;

    const sel = window.getSelection();
    const range = window.savedModalRange;

    if (tdId) {
        const startNode = range.startContainer;
        const parentTd  = (startNode.nodeType === 3 ? startNode.parentElement : startNode).closest('td');
        if (parentTd) parentTd.id = tdId;
    }

    if (!href) {
        ModalManager.close('linkModal');
        if (typeof window.syncPreviewToEditor === 'function') window.syncPreviewToEditor();
        return;
    }

    const clone = range.cloneContents();
    const tempDiv = document.createElement('div');
    tempDiv.appendChild(clone);
    const originalHTML = tempDiv.innerHTML;

    sel.removeAllRanges();
    sel.addRange(range);
    
    range.extractContents(); 

    const targetParent = range.commonAncestorContainer.nodeType === 3 
                         ? range.commonAncestorContainer.parentElement 
                         : range.commonAncestorContainer;

    if (targetParent) {
        const junkSpans = targetParent.querySelectorAll('span');
        junkSpans.forEach(s => {
            if (!s.textContent.trim() && s.childNodes.length === 0) {
                s.remove();
            }
        });
    }

    const a = document.createElement('a');
    a.href = href;
    a.target = document.getElementById('modalTargetBlank').checked ? '_blank' : '_self';
    if (document.getElementById('modalUnderline').checked) a.style.textDecoration = 'none';

    if (isImg) {
        a.innerHTML = originalHTML; 
    } else {
        a.textContent = newText;
    }
    
    range.insertNode(a);

    if (a.nextSibling && a.nextSibling.tagName === 'SPAN' && !a.nextSibling.textContent.trim()) {
        a.nextSibling.remove();
    }

    range.setStartAfter(a);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);

    ModalManager.close('linkModal');
    if (typeof window.syncPreviewToEditor === 'function') window.syncPreviewToEditor();
};

/* 캘린더 생성 */
window.generateCalendar = function () {
    const ymInput = document.getElementById('calBaseYM').value.trim();
    const [year, month] = ymInput.split('/').map(Number);
    const showHoliday   = document.getElementById('calShowHoliday').checked;
    const useId         = document.getElementById('calTargetId').checked;
    const html = generateBaseCalendar(`${year}/${month}`, { showHoliday, useId, lineHeight: '1' });
    if (typeof window.insertFormattedHtml === 'function') {
        window.insertFormattedHtml(html + '<p><br></p>');
    }
    if (typeof window.closeModal === 'function') window.closeModal('calendarModal');
};
