/**
 * editor-tools.js 
 */

const SEMANTIC_TAG_MAP = {
    bold:      'strong',
    italic:    'em',
    underline: 'u',
    strike:    's',
};

const styleMap = {
    font:      { key: 'fontFamily',      val: (value) => value,                    css: (value) => `font-family:${ColorManager.normFont(value)};` },
    size:      { key: 'fontSize',        val: (value) => value,                    css: (value) => `font-size:${value};` },
    bold:      { tag: SEMANTIC_TAG_MAP.bold,      key: 'fontWeight',     val: () => STYLE_VALUES.BOLD },
    italic:    { tag: SEMANTIC_TAG_MAP.italic,    key: 'fontStyle',      val: () => STYLE_VALUES.ITALIC },
    underline: { tag: SEMANTIC_TAG_MAP.underline, key: 'textDecoration', val: () => STYLE_VALUES.UNDERLINE },
    strike:    { tag: SEMANTIC_TAG_MAP.strike,    key: 'textDecoration', val: () => STYLE_VALUES.STRIKE },
    align:     { key: 'textAlign',       val: (value) => value,                    css: (value) => `text-align:${value};` },
    text: { key: 'color', val: (value) => ColorManager.toOriginalForm(value), css: (value) => `color:${ColorManager.toOriginalForm(value)};` },
    bg: { key: 'backgroundColor', val: (value) => ColorManager.toOriginalForm(value), css: (value) => `background-color:${ColorManager.toOriginalForm(value)};` }
};

const STYLE_VALUES = {
    BOLD: 'bold',
    BOLD_NUM: 700,
    ITALIC: 'italic',
    UNDERLINE: 'underline',
    STRIKE: 'line-through'
};

const _debounceTimers = {
    font: null,
    size: null,
};

const TextFormatTools = {
    styleKey(styles) {
        if (!styles) return '';
        return Object.entries(styles)
            .sort((a, b) => a[0].localeCompare(b[0]))
            .map(([p, v]) => `${p}:${v}`)
            .join(';');
    },
    mergeDecorations: (inheritedValue, ownValue) => {
        const allParts = new Set([
            ...(inheritedValue || '').split(/\s+/), 
            ...(ownValue || '').split(/\s+/)
        ].filter(Boolean));
        return [...allParts].join(' ');
    },    
    toggleDecoration: (currentValue, targetDecoration, isActive) => {
        const parts = (currentValue || '').split(/\s+/).filter(Boolean);
        if (isActive) {
            return parts.filter(part => part !== targetDecoration).join(' ');
        } else {
            if (!parts.includes(targetDecoration)) parts.push(targetDecoration);
            return parts.join(' ');
        }
    }
};

const TextEditor = {
    changeFontFamily(fontFamily) {
        if (!fontFamily) return;
        clearTimeout(_debounceTimers.font);
        _debounceTimers.font = setTimeout(() => this.applyStyle('font', fontFamily), 0);
    },
    changeFontSize(fontSize) {
        if (!fontSize) return;
        clearTimeout(_debounceTimers.size);
        _debounceTimers.size = setTimeout(() => this.applyStyle('size', fontSize), 0);
    },
    execStyle(styleType) {
        this.applyStyle(styleType);
    },
    align(alignType) {
        this.applyStyle('align', alignType);
    },
    applyStyle(styleType, value) {
        const savedRange = EditorState.get('savedRange');
        const editor     = EditorState.get('editor');
        let selectionData = this.getSelectionData();
        if (!selectionData && savedRange) selectionData = { type: 'preview', range: savedRange, text: savedRange.toString() };
        if (!selectionData) return;

        const styleMetadata = styleMap[styleType];
        if (!styleMetadata) return;

        EditorState.startSync();
        try {
            if (selectionData.type === 'preview') {
                const selection = window.getSelection();
                if (!selection || !selection.rangeCount) return;
                const selectionRange = selection.getRangeAt(0);
                if (!selectionRange.toString().trim()) return;

                const startElement = getResolvedNode(selectionRange.startContainer);
                let paragraphElement   = startElement.closest('p');

                if (!paragraphElement) {
                    const parentCell = startElement.closest('td');
                    if (parentCell) {
                        if (styleType === 'align') {
                            const alignValue = styleMetadata.val(value);
                            const currentAlign = parentCell.style.textAlign;
                            if (currentAlign === alignValue) {
                                parentCell.style.removeProperty('text-align');
                            } else {
                                parentCell.style.textAlign = alignValue;
                            }
                            if (parentCell.getAttribute('style') === '') parentCell.removeAttribute('style');
                            const nextAlign = parentCell.style.textAlign || '';
                            if (typeof this.updateToolbarStatus === 'function') {
                                this.updateToolbarStatus();
                            }
                            EditorState.endSync(true);
                            if (typeof window.syncPreviewToEditor === 'function') window.syncPreviewToEditor();
                            return;
                        }
                        _applyToContainer(parentCell, selectionRange, styleMetadata, styleType, value, selection, true);
                        return;
                    }
                }
                if (!paragraphElement) {
                    paragraphElement = document.createElement('p');
                    selectionRange.surroundContents(paragraphElement);
                }
                if (styleType === 'align') {
                    _applyAlign(paragraphElement, styleMetadata.val(value), selection);
                    return;
                }
                _applyToContainer(paragraphElement, selectionRange, styleMetadata, styleType, value, selection, false);

            } else if (selectionData.type === 'editor') {
                const plainText = selectionData.text.replace(/<\/?[^>]+(>|$)/g, '');
                let wrappedHtml;
                if (styleMetadata.tag) {
                    wrappedHtml = `<${styleMetadata.tag}>${plainText}</${styleMetadata.tag}>`;
                } else {
                    wrappedHtml = `<span style="${styleMetadata.css(value)}">${plainText}</span>`;
                }
                editor.replaceSelection(wrappedHtml);
            }
        } catch (_) {
        } finally {
            EditorState.endSync(true);
            if (selectionData.type === 'editor') editor.focus();
            if (typeof window.syncPreviewToEditor === 'function') window.syncPreviewToEditor();
            
            if (typeof this.updateToolbarStatus === 'function') this.updateToolbarStatus();
            
            if (window._headerLockRange && typeof window.applyHeaderLock === 'function') {
                requestAnimationFrame(() => window.applyHeaderLock());
            }
        }
    },

    getSelectionData() {
        const preview   = EditorState.get('preview');
        const editor    = EditorState.get('editor');
        const previewArea = document.getElementById('previewArea');
        const selection   = window.getSelection();
        if (selection?.anchorNode && previewArea?.contains(selection.anchorNode)) {
            const range = selection.getRangeAt(0);
            return { type: 'preview', range, text: range.toString() };
        }
        const selectedText = editor?.getSelection();
        if (editor?.hasFocus() || selectedText) {
            return { type: 'editor', text: selectedText || '', isCollapsed: !selectedText };
        }
        return null;
    },

    updateToolbarStatus() {
        const selection = window.getSelection();
        if (!selection || !selection.rangeCount) return;
        const range = selection.getRangeAt(0);

        let targetElements = [];
        if (!range.collapsed) {
            const treeWalker = document.createTreeWalker(
                range.commonAncestorContainer,
                NodeFilter.SHOW_TEXT,
                { acceptNode: node => (range.intersectsNode(node) && node.nodeValue.trim())
                    ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP }
            );
            let textNode;
            while ((textNode = treeWalker.nextNode())) {
                const parentElement = textNode.parentElement;
                if (parentElement && !targetElements.includes(parentElement)) targetElements.push(parentElement);
            }
        }
        if (targetElements.length === 0) {
            let startNode = range.startContainer;
            if (startNode.nodeType === 3) startNode = startNode.parentElement;
            targetElements = [startNode];
        }

        const computedStyleCache = new Map();
        const getComputedStyleCached = (element) => {
            if (!computedStyleCache.has(element)) computedStyleCache.set(element, window.getComputedStyle(element));
            return computedStyleCache.get(element);
        };
        const allElementsHave = (checkFn) => targetElements.length > 0 && targetElements.every(checkFn);

        const activeStyles = {
            bold:      allElementsHave(element => { const computed = getComputedStyleCached(element); return element.closest('strong,b') !== null || computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 700; }),
            italic:    allElementsHave(element => { const computed = getComputedStyleCached(element); return element.closest('em,i')     !== null || computed.fontStyle === 'italic'; }),
            underline: allElementsHave(element => { const computed = getComputedStyleCached(element); return element.closest('u')        !== null || computed.textDecoration.includes('underline'); }),
            strike:    allElementsHave(element => { const computed = getComputedStyleCached(element); return element.closest('strike,s,del') !== null || computed.textDecoration.includes('line-through'); }),
        };

        const buttonCache = window._toolbarBtnCache || {};
        Object.keys(activeStyles).forEach(styleType => {
            const button = buttonCache[styleType]; 
            if (button) window.setButtonActive(button, activeStyles[styleType]);
        });

        const anchorElement = targetElements[0] || range.startContainer;
        const blockElement = getResolvedNode(anchorElement).closest('p, td, th, div');
        const blockComputedStyle = blockElement ? window.getComputedStyle(blockElement) : null;
        let currentAlign = blockComputedStyle ? blockComputedStyle.textAlign : 'left';
        if (currentAlign === 'start' || !currentAlign) currentAlign = 'left';

        ['left', 'center', 'right'].forEach(alignType => {
            const button = buttonCache['align' + alignType.charAt(0).toUpperCase() + alignType.slice(1)];
            if (button) {
                window.setButtonActive(button, currentAlign === alignType);
            }
        });
    }
};

function walkNodes(node, callback) {
    if (!node) return;
    if (callback(node) === false) return false;
    let childNode = node.firstChild;
    while (childNode) {
        if (walkNodes(childNode, callback) === false) return false;
        childNode = childNode.nextSibling;
    }
}

function buildCharMap(paragraphElement) {
    const charMap = [];

    function collectInlineStyles(element) {
        const inlineStyle = element.style;
        const styleRecord = {};
        if (inlineStyle.fontWeight)      styleRecord.fontWeight      = inlineStyle.fontWeight;
        if (inlineStyle.fontStyle)       styleRecord.fontStyle       = inlineStyle.fontStyle;
        if (inlineStyle.textDecoration)  styleRecord.textDecoration  = inlineStyle.textDecoration;
        if (inlineStyle.color)           styleRecord.color           = ColorManager.toOriginalForm(inlineStyle.color);
        if (inlineStyle.backgroundColor) styleRecord.backgroundColor = ColorManager.toOriginalForm(inlineStyle.backgroundColor);
        if (inlineStyle.fontFamily)      styleRecord.fontFamily      = inlineStyle.fontFamily;
        if (inlineStyle.fontSize)        styleRecord.fontSize        = inlineStyle.fontSize;
        const tagName = element.tagName?.toLowerCase();
        if (tagName === SEMANTIC_TAG_MAP.bold)      styleRecord.fontWeight     = STYLE_VALUES.BOLD;
        if (tagName === SEMANTIC_TAG_MAP.italic)    styleRecord.fontStyle      = STYLE_VALUES.ITALIC;
        if (tagName === SEMANTIC_TAG_MAP.underline) {
            styleRecord.textDecoration = TextFormatTools.toggleDecoration(styleRecord.textDecoration, STYLE_VALUES.UNDERLINE, false);
        }
        if (tagName === SEMANTIC_TAG_MAP.strike) {
            styleRecord.textDecoration = TextFormatTools.toggleDecoration(styleRecord.textDecoration, STYLE_VALUES.STRIKE, false);
        }
        if (tagName === 'b' && !styleRecord.fontWeight) styleRecord.fontWeight = STYLE_VALUES.BOLD;
        if (tagName === 'i' && !styleRecord.fontStyle)  styleRecord.fontStyle  = STYLE_VALUES.ITALIC;
        return styleRecord;
    }
    const semanticTagNames = Object.values(SEMANTIC_TAG_MAP).map(tag => tag.toUpperCase());

    function walkAndCollect(node, inheritedStyles) {
        if (node.nodeType === Node.TEXT_NODE) {
            for (const character of node.textContent) {
                charMap.push({ char: character, styles: Object.assign({}, inheritedStyles) });
            }
        } else if (node.nodeType === Node.ELEMENT_NODE) {
            if (node.tagName === 'BR') {
                charMap.push({ char: '\n', isBr: true, styles: Object.assign({}, inheritedStyles) });
                return;
            }
            let ownStyles = {};
            if (['SPAN', ...semanticTagNames].includes(node.tagName)) {
                ownStyles = collectInlineStyles(node);
            }
            const mergedStyles = Object.assign({}, inheritedStyles, ownStyles);
            if (inheritedStyles.textDecoration && ownStyles.textDecoration) {
                mergedStyles.textDecoration = TextFormatTools.mergeDecorations(inheritedStyles.textDecoration, ownStyles.textDecoration);
            }
            for (const childNode of node.childNodes) {
                walkAndCollect(childNode, mergedStyles);
            }
        }
    }
    walkAndCollect(paragraphElement, collectInlineStyles(paragraphElement));
    return charMap;
}

function rebuildParagraph(paragraphElement, charMap, isFullSelection) {
    const savedTextAlign = paragraphElement.style.textAlign;
    paragraphElement.innerHTML = '';
    paragraphElement.removeAttribute('class');
    paragraphElement.removeAttribute('style');
    if (savedTextAlign) paragraphElement.style.textAlign = savedTextAlign;

    let charIndex = 0;
    while (charIndex < charMap.length) {
        if (charMap[charIndex].isBr) { paragraphElement.appendChild(document.createElement('br')); charIndex++; continue; }
        const currentStyleKey = TextFormatTools.styleKey(charMap[charIndex].styles);
        let nextIndex = charIndex + 1;
        while (nextIndex < charMap.length && !charMap[nextIndex].isBr && TextFormatTools.styleKey(charMap[nextIndex].styles) === currentStyleKey) nextIndex++;
        const characters = charMap.slice(charIndex, nextIndex).map(entry => entry.char).join('');
        const styles      = charMap[charIndex].styles;
        paragraphElement.appendChild(
            Object.keys(styles).length === 0
                ? document.createTextNode(characters)
                : buildSemanticNode(characters, styles)
        );
        charIndex = nextIndex;
    }
}

function buildSemanticNode(textContent, styles) {
    const spanStyles   = {};
    const semanticTags = [];

    const textDecoration = styles.textDecoration || '';
    if (textDecoration.includes(STYLE_VALUES.UNDERLINE)) semanticTags.push('u');
    if (textDecoration.includes(STYLE_VALUES.STRIKE))    semanticTags.push('s');
    if (styles.fontWeight === STYLE_VALUES.BOLD || parseInt(styles.fontWeight) >= STYLE_VALUES.BOLD_NUM) { semanticTags.push(SEMANTIC_TAG_MAP.bold); }
    if (styles.fontStyle === STYLE_VALUES.ITALIC) { semanticTags.push(SEMANTIC_TAG_MAP.italic); }
    if (styles.color)           spanStyles.color           = styles.color;
    if (styles.backgroundColor) spanStyles.backgroundColor = styles.backgroundColor;
    if (styles.fontFamily)      spanStyles.fontFamily      = styles.fontFamily;
    if (styles.fontSize)        spanStyles.fontSize        = styles.fontSize;

    let innerNode;
    if (Object.keys(spanStyles).length > 0) {
        const spanElement = document.createElement('span');
        const cssParts = [];
        if (spanStyles.color)           cssParts.push(`color:${spanStyles.color}`);
        if (spanStyles.backgroundColor) cssParts.push(`background-color:${spanStyles.backgroundColor}`);
        if (spanStyles.fontFamily)      cssParts.push(`font-family:${ColorManager.normFont(spanStyles.fontFamily)}`);
        if (spanStyles.fontSize)        cssParts.push(`font-size:${spanStyles.fontSize}`);
        spanElement.setAttribute('style', cssParts.join(';'));
        spanElement.textContent = textContent;
        innerNode = spanElement;
    } else {
        innerNode = document.createTextNode(textContent);
    }

    let currentNode = innerNode;
    for (let tagIndex = semanticTags.length - 1; tagIndex >= 0; tagIndex--) {
        const wrapperElement = document.createElement(semanticTags[tagIndex]);
        wrapperElement.appendChild(currentNode);
        currentNode = wrapperElement;
    }
    return currentNode;
}

function getSelectionCharRange(paragraphElement, selectionRange) {
    let startCharIndex = -1, endCharIndex = -1, totalCharCount = 0;

    function walkAndCount(node) {
        if (node.nodeType === Node.TEXT_NODE) {
            let charOffset = 0;
            for (const character of node.textContent) {
                if (node === selectionRange.startContainer && charOffset === selectionRange.startOffset) startCharIndex = totalCharCount;
                if (node === selectionRange.endContainer   && charOffset === selectionRange.endOffset)   endCharIndex   = totalCharCount;
                totalCharCount++;
                charOffset += character.length;
            }
            if (node === selectionRange.endContainer   && charOffset === selectionRange.endOffset)   endCharIndex   = totalCharCount;
            if (node === selectionRange.startContainer && charOffset === selectionRange.startOffset) startCharIndex = totalCharCount;
        } else if (node.nodeType === Node.ELEMENT_NODE) {
            if (node.tagName === 'BR') { totalCharCount++; return; }
            for (const childNode of node.childNodes) walkAndCount(childNode);
        }
    }
    walkAndCount(paragraphElement);
    if (startCharIndex === -1) startCharIndex = 0;
    if (endCharIndex   === -1) endCharIndex   = totalCharCount;
    return { startIdx: startCharIndex, endIdx: endCharIndex, total: totalCharCount };
}

function restoreSelection(selection, paragraphElement, startIndex, endIndex, isFullSelection) {
    try {
        const newRange = document.createRange();
        if (isFullSelection) {
            newRange.selectNodeContents(paragraphElement);
            selection.removeAllRanges();
            selection.addRange(newRange);
            return;
        }

        function findNodeAtIndex(targetIndex) {
            let charCount = 0;
            function walkToFind(node) {
                if (node.nodeType === Node.TEXT_NODE) {
                    const textLength = node.textContent.length;
                    if (charCount + textLength > targetIndex) return { node, offset: targetIndex - charCount };
                    charCount += textLength;
                    return null;
                } else if (node.nodeType === Node.ELEMENT_NODE) {
                    if (node.tagName === 'BR') { charCount++; return null; }
                    for (const childNode of node.childNodes) {
                        const result = walkToFind(childNode);
                        if (result) return result;
                    }
                }
                return null;
            }
            return walkToFind(paragraphElement);
        }
        const startNode = findNodeAtIndex(startIndex);
        const endNode   = findNodeAtIndex(endIndex);
        if (startNode && endNode) { newRange.setStart(startNode.node, startNode.offset); newRange.setEnd(endNode.node, endNode.offset); }
        else newRange.selectNodeContents(paragraphElement);
        selection.removeAllRanges();
        selection.addRange(newRange);
    } catch (_) {
    }
}

function _applyToContainer(container, selectionRange, styleMetadata, styleType, value, selection, isTd = false) {
    const charMap = buildCharMap(container);
    const { startIdx, endIdx, total } = getSelectionCharRange(container, selectionRange);
    const selectedChars = charMap.slice(startIdx, endIdx).filter(entry => !entry.isBr);
    const isColorStyle  = styleType === 'text' || styleType === 'bg';
    const isCurrentlyActive = !isColorStyle && selectedChars.length > 0 && selectedChars.every(entry => {
        const styleValue = entry.styles[styleMetadata.key];
        if (!styleValue) return false;
        if (styleMetadata.key === 'fontWeight') return styleValue === STYLE_VALUES.BOLD || parseInt(styleValue) >= STYLE_VALUES.BOLD_NUM;
        if (styleMetadata.key === 'textDecoration') return styleValue.includes(styleMetadata.val());
        if (styleMetadata.key === 'fontStyle') return styleValue === STYLE_VALUES.ITALIC;
        return styleValue === styleMetadata.val(value);
    });
    const hasBrChar  = isTd && charMap.some(entry => entry.isBr);
    const isFullRange = (!isTd
        ? startIdx === 0 && endIdx === total && endIdx - startIdx === total
        : !hasBrChar && startIdx === 0 && endIdx === total
    );

    for (let charIndex = startIdx; charIndex < endIdx; charIndex++) {
        if (charMap[charIndex].isBr) continue;
        if (isColorStyle) {
            charMap[charIndex].styles[styleMetadata.key] = styleMetadata.val(value);
        } else if (styleMetadata.key === 'textDecoration') {
            const decorationValue = styleMetadata.val();
            const currentDecoration = charMap[charIndex].styles[styleMetadata.key] || '';    
            const nextDecoration = TextFormatTools.toggleDecoration(currentDecoration, decorationValue, isCurrentlyActive);
            if (nextDecoration) {
                charMap[charIndex].styles[styleMetadata.key] = nextDecoration;
            } else {
                delete charMap[charIndex].styles[styleMetadata.key];
            }
        } else if (isCurrentlyActive) {
            delete charMap[charIndex].styles[styleMetadata.key];
        } else {
            charMap[charIndex].styles[styleMetadata.key] = styleMetadata.val(value);
        }
    }

    const savedContainerStyle = isTd ? (container.getAttribute('style') || '') : null;
    rebuildParagraph(container, charMap, isFullRange);
    if (isTd && savedContainerStyle) container.setAttribute('style', savedContainerStyle);

    if (!isTd) {
        TextEditor.updateToolbarStatus();
    }
    restoreSelection(selection, container, startIdx, endIdx, isFullRange);
}

function _applyAlign(paragraphElement, alignValue, selection) {
    const currentAlign = paragraphElement.style.textAlign;
    paragraphElement.style.textAlign = (currentAlign === alignValue || alignValue === 'left') ? '' : alignValue;
    if (paragraphElement.getAttribute('style') === '') paragraphElement.removeAttribute('style');
    const nextAlign = paragraphElement.style.textAlign || 'left';
	if (typeof TextEditor.updateToolbarStatus === 'function') {
        TextEditor.updateToolbarStatus();
    }
    const fullRange = document.createRange();
    fullRange.selectNodeContents(paragraphElement);
    selection.removeAllRanges();
    selection.addRange(fullRange);
}

// ── 컬러 피커 ─────────────────────────────────────────────
function openColor(colorMode, triggerButton, event) {
    if (event) { event.stopPropagation(); event.preventDefault(); }
    if (window.activePicker) return;

    const selection = window.getSelection();
    let savedSelectionRange = null;
    let initialColor = colorMode === 'text' ? '#000000' : '#ffffff';

    if (selection && selection.rangeCount > 0) {
        savedSelectionRange = selection.getRangeAt(0).cloneRange();
        const anchorElement = getResolvedNode(savedSelectionRange.startContainer);
        const computedColor = window.getComputedStyle(anchorElement)[colorMode === 'text' ? 'color' : 'backgroundColor'];
        initialColor = ColorManager.toOriginalForm(computedColor) || initialColor;
    }

    const picker = new Picker({
        parent: triggerButton,
        popup:  'bottom',
        alpha:  false,
        editor: true,
        color:  initialColor,
        onClose: () => {
            if (window.activePicker) { window.activePicker.destroy(); window.activePicker = null; }
        },
        onDone: (selectedColor) => {
            const finalHex = selectedColor.hex.slice(0, 7).toLowerCase();
            if (savedSelectionRange) {
                const selection = window.getSelection();
                selection.removeAllRanges();
                selection.addRange(savedSelectionRange);
                TextEditor.applyStyle(colorMode, finalHex);
            }
            if (window.activePicker) { window.activePicker.destroy(); window.activePicker = null; }
        },
    });
    picker.show();
    window.activePicker = picker;
}

// ── 하이퍼링크 ────────────────────────────────────────────

window.prepareLinkData = function () {
    const selection = window.getSelection();
    if (!selection || !selection.rangeCount) return null;
    const selectionRange = selection.getRangeAt(0);

    const startNode     = selectionRange.startContainer;
    const parentAnchor  = getResolvedNode(startNode).closest('a');
    const parentCell    = getResolvedNode(startNode).closest('td');

    const clonedContents = selectionRange.cloneContents();
    const tempContainer  = document.createElement('div');
    tempContainer.appendChild(clonedContents);
    const imageElement = tempContainer.querySelector('img');

    const textInputElement   = document.getElementById('modalTextDisplay');
    const previewImageEl     = document.getElementById('modalPreviewImg');
    const imagePreviewBox    = document.getElementById('linkImgPreview');

    if (imageElement) {
        textInputElement.value = imageElement.src;
        if (previewImageEl) {
            previewImageEl.src = imageElement.src;
            imagePreviewBox.style.display = 'block';
        }
        textInputElement.setAttribute('data-is-img', 'true');
        textInputElement.setAttribute('data-img-style', imageElement.style.cssText);
    } else {
        textInputElement.value = selectionRange.toString().trim();
        textInputElement.removeAttribute('data-is-img');
        if (imagePreviewBox) imagePreviewBox.style.display = 'none';
    }

    document.getElementById('modalLinkHref').value = parentAnchor?.getAttribute('href') || '';
    document.getElementById('modalTdId').value     = parentCell?.id?.replace(CONSTANTS.USER_CONTENT_PREFIX, '') || '';

    if (parentAnchor) {
        document.getElementById('modalTargetBlank').checked = (parentAnchor.getAttribute('target') || '_blank') === '_blank';
        document.getElementById('modalUnderline').checked   = parentAnchor.style.textDecoration === 'none';
    } else {
        document.getElementById('modalTargetBlank').checked = true;
        document.getElementById('modalUnderline').checked   = true;
    }
    return selectionRange.cloneRange();
};

window.applyLinkChanges = function () {
    const textInputElement = document.getElementById('modalTextDisplay');
    const displayText   = textInputElement.value;
    const isImage       = textInputElement.getAttribute('data-is-img') === 'true';
    const hrefUrl       = document.getElementById('modalLinkHref').value.trim();
    const cellIdRaw     = document.getElementById('modalTdId').value.trim();
    const fullCellId    = cellIdRaw ? CONSTANTS.USER_CONTENT_PREFIX + cellIdRaw : '';
    const selection      = window.getSelection();
    const selectionRange = ModalManager._savedRange;
	if (!selectionRange) return;

    if (fullCellId) {
        const parentCell = getResolvedNode(selectionRange.startContainer).closest('td');
        if (parentCell) parentCell.id = fullCellId;
    }

    if (!hrefUrl) {
        ModalManager.close('linkModal');
        if (typeof window.syncPreviewToEditor === 'function') window.syncPreviewToEditor();
        return;
    }

    const clonedContents = selectionRange.cloneContents();
    const tempContainer  = document.createElement('div');
    tempContainer.appendChild(clonedContents);
    const originalHtml = tempContainer.innerHTML;

    selection.removeAllRanges();
    selection.addRange(selectionRange);
    selectionRange.extractContents();

    const commonAncestorElement = getResolvedNode(selectionRange.commonAncestorContainer);
    if (commonAncestorElement) {
        commonAncestorElement.querySelectorAll('span').forEach(spanElement => {
            if (!spanElement.textContent.trim() && spanElement.childNodes.length === 0) spanElement.remove();
        });
    }

    const anchorElement = document.createElement('a');
    anchorElement.href   = hrefUrl;
    anchorElement.target = document.getElementById('modalTargetBlank').checked ? '_blank' : '_self';
    if (document.getElementById('modalUnderline').checked) anchorElement.style.textDecoration = 'none';
    anchorElement[isImage ? 'innerHTML' : 'textContent'] = isImage ? originalHtml : displayText;

    selectionRange.insertNode(anchorElement);
    if (anchorElement.nextSibling?.tagName === 'SPAN' && !anchorElement.nextSibling.textContent.trim()) {
        anchorElement.nextSibling.remove();
    }

    selectionRange.setStartAfter(anchorElement);
    selectionRange.collapse(true);
    selection.removeAllRanges();
    selection.addRange(selectionRange);

    ModalManager.close('linkModal');
    if (typeof window.syncPreviewToEditor === 'function') window.syncPreviewToEditor();
};

// ── 캘린더 생성 ───────────────────────────────────────────
window.generateCalendar = function () {
    const yearMonthInput = document.getElementById('calBaseYM')?.value.trim() || '';
    if (!window.isValidYearMonth?.(yearMonthInput)) {
        window.showToast('올바른 연/월 형식(YYYY/MM)을 입력하세요.', 'error');
        return;
    }
    const [year, month] = yearMonthInput.split('/').map(Number);
    const showHoliday   = document.getElementById('calShowHoliday')?.checked ?? true;
    const useId         = document.getElementById('calTargetId')?.checked ?? false;
    const calendarHtml  = generateBaseCalendar(`${year}/${month}`, { showHoliday, useId, lineHeight: '1' });
    if (typeof window.insertFormattedHtml === 'function') {
        window.insertFormattedHtml(calendarHtml + '<p><br></p>');
    }
    if (typeof window.closeModal === 'function') window.closeModal('calendarModal');
};

window.initToolbarCache = function() {
    window._toolbarBtnCache = {
        bold:      document.querySelector(".icon-btn[onclick*=\"TextEditor.execStyle('bold')\"]"),
        italic:    document.querySelector(".icon-btn[onclick*=\"TextEditor.execStyle('italic')\"]"),
        underline: document.querySelector(".icon-btn[onclick*=\"TextEditor.execStyle('underline')\"]"),
        strike:    document.querySelector(".icon-btn[onclick*=\"TextEditor.execStyle('strike')\"]"),
        alignLeft:   document.querySelector(".icon-btn[onclick*=\"TextEditor.align('left')\"]"),
        alignCenter: document.querySelector(".icon-btn[onclick*=\"TextEditor.align('center')\"]"),
        alignRight:  document.querySelector(".icon-btn[onclick*=\"TextEditor.align('right')\"]"),
    };
};
