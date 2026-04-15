/**
 * modal-factory.js 
 */

const ModalManager = {
    currentDate: new Date(),
	_savedRange: null,
    createBase(modalId, title, icon = 'settings') {
        const overlayElement = document.createElement('div');
        overlayElement.id        = modalId;
        overlayElement.className = 'modal-overlay';
        overlayElement.innerHTML = `
            <div class="modal-content">
                <div class="modal-header">
                    <h3><span class="material-symbols-outlined">${icon}</span> ${title}</h3>
                    <button class="modal-close-x btn-cancel">
                        <span class="material-symbols-outlined">close</span>
                    </button>
                </div>
                <hr>
                <div class="modal-body"></div>
                <div class="modal-footer-between">
                    <div class="left-btns"></div>
                    <div class="main-btns">
                        <button class="btn-confirm">저장</button>
                    </div>
                </div>
            </div>`;
        overlayElement.querySelector('.btn-cancel').onclick = () => this.close(modalId);
        return overlayElement;
    },
    close(modalId) {
        const modalElement = document.getElementById(modalId);
        if (!modalElement) return;
        modalElement.querySelectorAll('[data-resetable]').forEach(inputElement => {
            if (inputElement.type === 'checkbox') {
                const defaultChecked = inputElement.dataset.defaultChecked;
                inputElement.checked = defaultChecked !== undefined ? (defaultChecked === 'true') : inputElement.defaultChecked;
            } else {
                inputElement.value = inputElement.dataset.resetValue ?? '';
            }
        });
        modalElement.style.display = 'none';
        if (typeof window.updatePreview === 'function') window.updatePreview();
        this._savedRange = null;
    },

    _mount(modalElement) {
        document.getElementById('modalContainer').appendChild(modalElement);
        modalElement.style.display = 'flex';
    },

    openAnalysisModal() {
        const modalId = 'analysisModal';
        let modalElement = document.getElementById(modalId);
        if (modalElement) {
            modalElement.style.display = 'flex';
            this._refreshData(modalElement, 'analysis_source_save', 'analysisInput');
            return;
        }

        modalElement = this.createBase(modalId, '데이터 분석/적용', 'table_chart');
        modalElement.querySelector('.modal-body').innerHTML = `
            <div class="modal-section">
                <label class="modal-label">분석할 Table 소스코드 (Sample 추출)</label>
                <textarea id="analysisInput" class="modal-input" placeholder="여기에 소스코드를 붙여넣으세요..."></textarea>
                <p class="helper-text">* 붙여넣은 데이터를 분석하여 자동으로 삽입합니다.</p>
            </div>`;

        modalElement.querySelector('.btn-confirm').onclick = () => {
            if (typeof processAnalysis === 'function') {
                processAnalysis(); 
                this.close(modalId); 
            }
        };
        this._mount(modalElement);
        this._refreshData(modalElement, 'analysis_source_save', 'analysisInput');
    },
    _refreshData(modalElement, storeKey, inputElementId) {
        const inputElement = modalElement.querySelector(`#${inputElementId}`);
        const savedValue   = AppStore.get(storeKey);
        if (inputElement && savedValue) inputElement.value = savedValue;
    },
    openRuleModal() {
        let modalElement = document.getElementById('ruleModal');
        if (modalElement) {
            modalElement.style.display = 'flex';
            if (typeof window.renderRules === 'function') window.renderRules();
            return;
        }

        modalElement = this.createBase('ruleModal', '커스텀 툴바 설정', 'playlist_add');
        modalElement.querySelector('.modal-body').innerHTML = `<div id="ruleGroupsContainer"></div>`;

        const addGroupButton = document.createElement('button');
        addGroupButton.className = 'btn-secondary';
        addGroupButton.innerText = '+ 새 그룹 추가';
        addGroupButton.onclick   = () => addGroup();
        modalElement.querySelector('.left-btns').appendChild(addGroupButton);

        modalElement.querySelector('.btn-confirm').onclick = () => {
            if (typeof applyAndSaveRules === 'function') {
                applyAndSaveRules();
                this.close('ruleModal'); 
            }
        };
        this._mount(modalElement);
        if (typeof window.renderRules === 'function') window.renderRules();
    },

    openCalendarModal() {
        let modalElement = document.getElementById('calendarModal');
        if (modalElement) { modalElement.style.display = 'flex'; return; }

        modalElement = this.createBase('calendarModal', '캘린더 생성/변환', 'calendar_month');

        const formatYearMonth = (dateObject) => `${dateObject.getFullYear()}/${String(dateObject.getMonth() + 1).padStart(2, '0')}`;
        const currentYearMonth = formatYearMonth(this.currentDate);
        const nextMonthDate    = new Date(this.currentDate.getFullYear(), this.currentDate.getMonth() + 1, 1);
        const nextYearMonth    = formatYearMonth(nextMonthDate);

        modalElement.querySelector('.modal-body').innerHTML = `
            <div class="modal-section">
                <label class="modal-label">기본형 생성</label>
                <div class="tool-group" style="gap:10px;margin-bottom:10px;">
                    <input type="text" id="calBaseYM" class="modal-input input-date-ym"
                        value="${currentYearMonth}" data-resetable data-reset-value="${currentYearMonth}">
                    <button type="button" id="btnGenBase" class="btn-confirm">생성</button>
                </div>
                <div class="checkbox-group">
                    <label class="checkbox-label"><input type="checkbox" id="calShowHoliday" checked data-resetable> 휴일 표시</label>
                    <label class="checkbox-label"><input type="checkbox" id="calTargetId" data-resetable> Id 링크 생성</label>
                    <span style="color:#1a73e8;font-size:11px;">*#${CONSTANTS.USER_CONTENT_PREFIX}(날짜)</span>
                </div>
            </div>
            <hr>
            <div class="modal-section">
                <label class="modal-label">고급형 (스타일 유지)</label>
                <textarea id="advSourceHtml" class="modal-input" data-resetable
                    placeholder="여기에 소스코드를 붙여넣으세요..." style="height:80px;resize:none;"></textarea>
            </div>`;

        const mainButtonsContainer = modalElement.querySelector('.main-btns');
        mainButtonsContainer.style.cssText = 'display:flex;align-items:center;width:100%;';
        mainButtonsContainer.innerHTML = `
            <div class="tool-group" style="gap:5px;">
                <input type="text" id="advFromYM" class="modal-input input-date-ym"
                    value="${currentYearMonth}" data-resetable data-reset-value="${currentYearMonth}">
                <span style="color:#ccc">→</span>
                <input type="text" id="advToYM" class="modal-input input-date-ym"
                    value="${nextYearMonth}" data-resetable data-reset-value="${nextYearMonth}">
            </div>
            <button type="button" id="btnSave" class="btn-confirm">변환</button>`;

        modalElement.querySelector('#btnGenBase').addEventListener('click', () => {
            const yearMonth = modalElement.querySelector('#calBaseYM').value;
            if (!window.isValidYearMonth?.(yearMonth)) {
                window.showToast('01월부터 12월 사이로 입력해 주세요.');
                return;
            }
            const calendarHtml = generateBaseCalendar(yearMonth, {
                showHoliday: modalElement.querySelector('#calShowHoliday').checked,
                useId:       modalElement.querySelector('#calTargetId').checked,
            });
            if (typeof window.insertFormattedHtml === 'function') {
                window.insertFormattedHtml(calendarHtml);
                this.close('calendarModal');
            }
        });

        modalElement.querySelector('#btnSave').addEventListener('click', () => {
            const sourceHtml    = modalElement.querySelector('#advSourceHtml').value;
            const fromYearMonth = modalElement.querySelector('#advFromYM').value;
            const toYearMonth   = modalElement.querySelector('#advToYM').value;
            if (!sourceHtml.trim()) { window.showToast('변환할 소스코드를 입력해주세요.'); return; }
            if (!window.isValidYearMonth?.(fromYearMonth) || !window.isValidYearMonth?.(toYearMonth)) {
                window.showToast('1월부터 12월 사이로 입력해 주세요.');
                return;
            }
            try {
                if (typeof window.insertFormattedHtml === 'function') {
                    window.insertFormattedHtml(transformAdvancedCalendar(sourceHtml, fromYearMonth, toYearMonth));
                }
                this.close('calendarModal');
            } catch (_) {
                window.showToast('변환 중 오류가 발생했습니다. 소스코드를 확인해 주세요.');
            }
        });

        this._mount(modalElement);
    },

    openExtendRowModal() {
        let modalElement = document.getElementById('extendRowModal');
        if (modalElement) {
            modalElement.style.display = 'flex';
            this._refreshData(modalElement, 'extend_row_day_colors', 'modalTargetAttr'); 
            return;
        }

        modalElement = this.createBase('extendRowModal', '줄 확장', 'add_row_below');
        const todayDate    = new Date();
        const currentYearMonth = `${todayDate.getFullYear()}-${String(todayDate.getMonth() + 1).padStart(2, '0')}`;

        modalElement.querySelector('.modal-body').innerHTML = `
            <div class="modal-section">
                <label class="modal-label">기간 설정</label>
                <div class="tool-group" style="display:flex;align-items:center;gap:5px;">
                    <input type="text" id="modalExtendFrom" class="modal-input input-date-ym"
                        placeholder="17" style="width:100px;" data-resetable>
                    <span style="color:#ccc">→</span>
                    <input type="text" id="modalExtendTo" class="modal-input input-date-ym"
                        placeholder="31" style="width:100px;" data-resetable>
                </div>
                <p class="helper-text">* 17→31 경우 17일짜리 줄을 31일까지 확장합니다</p>
            </div>
            <div class="modal-section">
                <label class="modal-label">날짜 Id 자동 생성</label>
                <div class="id-input-wrapper">
                    <span class="id-prefix">${CONSTANTS.USER_CONTENT_PREFIX}</span>
                    <input type="text" id="modalDateId" class="id-inner-input"
                        placeholder="(입력)날짜 형태로 저장됩니다. 예)d → d1, day → day1 등." data-resetable>
                </div>
                <p class="helper-text">* 이미 샘플 코드에 Id 선택자가 있으면 사용할 필요 없음</p>
            </div>
            <div class="modal-section">
                <label class="modal-label">날짜칸 색상</label>
                <div class="input-row" style="display:flex;gap:10px;align-items:center;width:100%;">
                    <input type="text" id="modalTargetAttr" class="modal-input"
                        placeholder="#333333(평일), #0000FF(토), #FF0000(일)"
                        style="flex:7;min-width:0;" data-resetable>
                    <input type="month" id="modalBaseMonth" class="modal-input"
                        style="flex:3;min-width:0;cursor:pointer;text-align:center;"
                        value="${currentYearMonth}" data-resetable data-reset-value="${currentYearMonth}">
                </div>
                <p class="helper-text">* #헥스코드로만 입력해주세요(rgb 안 됨) 쉼표 구분 필수.</p>
                <p class="helper-text">* 1개만 입력하면 색 일괄 통일, 2개 입력 시 일요일만 색 구분</p>
            </div>`;

        modalElement.querySelector('.btn-confirm').onclick = () => {
            const fromDay = parseInt(document.getElementById('modalExtendFrom').value, 10);
            const toDay   = parseInt(document.getElementById('modalExtendTo').value, 10);
            if (fromDay > toDay) { window.showToast('종료일은 시작일보다 나중이어야 합니다.'); return; }

            const colorInputValue = (document.getElementById('modalTargetAttr')?.value || '').trim();
            if (colorInputValue) AppStore.set('extend_row_day_colors', colorInputValue);
            else AppStore.remove('extend_row_day_colors');

            if (typeof window.executeExtendRow === 'function') {
                window.executeExtendRow();
                ModalManager.close('extendRowModal');
            }
        };

        this._mount(modalElement);
        this._refreshData(modalElement, 'extend_row_day_colors', 'modalTargetAttr');
    },

    openLinkModal() {
        let modalElement = document.getElementById('linkModal');
        if (modalElement) {
            modalElement.style.display = 'flex';
            const targetBlankCheckbox = document.getElementById('modalTargetBlank');
            const underlineCheckbox   = document.getElementById('modalUnderline');
            if (targetBlankCheckbox) targetBlankCheckbox.checked = true;
            if (underlineCheckbox)   underlineCheckbox.checked   = true;
            return;
        }

        modalElement = this.createBase('linkModal', '링크 삽입/변경', 'link');
        modalElement.querySelector('.modal-body').innerHTML = `
            <div class="modal-section">
                <label class="modal-label">표시할 텍스트(또는 이미지 주소)</label>
                <input type="text" id="modalTextDisplay" class="modal-input"
                    placeholder="표시될 텍스트(입력 가능)" data-resetable>
                <div id="linkImgPreview">
                    <img id="modalPreviewImg" src="">
                </div>
            </div>
            <div class="modal-section">
                <label class="modal-label">URL</label>
                <input type="text" id="modalLinkHref" class="modal-input"
                    placeholder="https://..." data-resetable>
                <div class="checkbox-group">
                    <label class="checkbox-label">
                        <input type="checkbox" id="modalTargetBlank" checked
                            data-resetable data-default-checked="true"> 새 창에서 열기
                    </label>
                    <label class="checkbox-label">
                        <input type="checkbox" id="modalUnderline" checked
                            data-resetable data-default-checked="true"> 밑줄 제거
                    </label>
                </div>
            </div>
            <hr>
            <div class="modal-section">
                <label class="modal-label">Id (td 요소에 적용)</label>
                <div class="id-input-wrapper">
                    <span class="id-prefix">${CONSTANTS.USER_CONTENT_PREFIX}</span>
                    <input type="text" id="modalTdId" class="id-inner-input"
                        placeholder="Id값을 입력하세요" data-resetable>
                </div>
                <p class="helper-text">* td 칸 안에 있을 경우 해당 칸의 Id로 저장됩니다.</p>
            </div>`;

        modalElement.querySelector('.btn-confirm').onclick = () => {
            if (typeof window.applyLinkChanges === 'function') window.applyLinkChanges();
        };
        this._mount(modalElement);
    },
};

window.openModal = function (modalId) {
    const modalOpeners = {
        analysisModal:  () => ModalManager.openAnalysisModal(),
        ruleModal:      () => ModalManager.openRuleModal(),
        calendarModal:  () => ModalManager.openCalendarModal(),
        extendRowModal: () => ModalManager.openExtendRowModal(),
        linkModal:      () => {
            ModalManager.openLinkModal();
            if (typeof window.prepareLinkData === 'function') {
                ModalManager._savedRange = window.prepareLinkData(); 
            }
        },
    };
    const opener = modalOpeners[modalId];
    if (opener) opener();
};

window.closeModal = (modalId) => ModalManager.close(modalId);
