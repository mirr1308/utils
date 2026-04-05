/**
 * modal-factory.js
 * 역할: 공통 모달 레이아웃 생성 및 동적 주입
 */
const ModalManager = {
    currentDate: new Date(),
    createBase(id, title, icon = 'settings') {
        const overlay = document.createElement('div');
        overlay.id = id;
        overlay.className = 'modal-overlay';
        overlay.innerHTML = `
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
        overlay.querySelector('.btn-cancel').onclick = () => this.close(id);
        return overlay;
    },

    close(id) {
        const modal = document.getElementById(id);
        if (!modal) return;
        modal.querySelectorAll('[data-resetable]').forEach(el => {
            if (el.type === 'checkbox') {
                const def = el.dataset.defaultChecked;
                el.checked = def !== undefined ? (def === 'true') : el.defaultChecked;
            } else {
                el.value = el.dataset.resetValue ?? '';
            }
        });
        modal.style.display = 'none';
        if (typeof window.updatePreview === 'function') window.updatePreview();
        window.savedModalRange = null;
    },

    openAnalysisModal() {
		const modalId = 'analysisModal';
        let modal = document.getElementById(modalId);	
		if (modal) {
            modal.style.display = 'flex';
            const input = document.getElementById('analysisInput');
            const savedCode = AppStore.get('analysis_source_save');
            if (input) input.value = savedCode || ""; 
            return;
        }
        modal = this.createBase(modalId, '데이터 분석/적용', 'table_chart');
        modal.querySelector('.modal-body').innerHTML = `
            <div class="modal-section">
                <label class="modal-label">분석할 Table 소스코드 (Sample 추출)</label>
                <textarea id="analysisInput" class="modal-input" placeholder="여기에 소스코드를 붙여넣으세요..."></textarea>
                <p class="helper-text">* 붙여넣은 데이터를 분석하여 자동으로 삽입합니다.</p>
            </div>`;
        const input = modal.querySelector('#analysisInput');

        modal.querySelector('.btn-confirm').onclick = () => {
            if (typeof processAnalysis === 'function') processAnalysis();
        };
		document.getElementById('modalContainer').appendChild(modal);
        modal.style.display = 'flex';
		
        const savedCode = AppStore.get('analysis_source_save');
        if (input && savedCode) input.value = savedCode;
    },

    openRuleModal() {
        let modal = document.getElementById('ruleModal');
        if (modal) {
            modal.style.display = 'flex';
            if (typeof window.renderRules === 'function') window.renderRules();
            return;
        }
        modal = this.createBase('ruleModal', '커스텀 툴바 설정', 'playlist_add');
        modal.querySelector('.modal-body').innerHTML = `<div id="ruleGroupsContainer"></div>`;
        const addBtn = document.createElement('button');
        addBtn.className = 'btn-secondary';
        addBtn.innerText = '+ 새 그룹 추가';
        addBtn.onclick = () => addGroup();
        modal.querySelector('.left-btns').appendChild(addBtn);
        modal.querySelector('.btn-confirm').onclick = () => {
            if (typeof applyAndSaveRules === 'function') applyAndSaveRules();
        };
        document.getElementById('modalContainer').appendChild(modal);
        modal.style.display = 'flex';
        if (typeof window.renderRules === 'function') window.renderRules();
    },

    openCalendarModal() {
        let modal = document.getElementById('calendarModal');
        if (modal) { modal.style.display = 'flex'; return; }

        modal = this.createBase('calendarModal', '캘린더 생성/변환', 'calendar_month');
        const initYM = `${this.currentDate.getFullYear()}/${String(this.currentDate.getMonth() + 1).padStart(2, '0')}`;
        const nextDate = new Date(this.currentDate.getFullYear(), this.currentDate.getMonth() + 1, 1);
        const nextYM = `${nextDate.getFullYear()}/${String(nextDate.getMonth() + 1).padStart(2, '0')}`;

        modal.querySelector('.modal-body').innerHTML = `
            <div class="modal-section">
                <label class="modal-label">기본형 생성</label>
                <div class="tool-group" style="gap:10px;margin-bottom:10px;">
                    <input type="text" id="calBaseYM" class="modal-input input-date-ym"
                        value="${initYM}" data-resetable data-reset-value="${initYM}">
                    <button type="button" id="btnGenBase" class="btn-confirm">생성</button>
                </div>
                <div class="checkbox-group">
                    <label class="checkbox-label"><input type="checkbox" id="calShowHoliday" checked data-resetable> 휴일 표시</label>
                    <label class="checkbox-label"><input type="checkbox" id="calTargetId" data-resetable> Id 링크 생성</label>
                    <span style="color:#1a73e8;font-size:11px;">*#user_content_(날짜)</span>
                </div>
            </div>
            <hr>
            <div class="modal-section">
                <label class="modal-label">고급형 (스타일 유지)</label>
                <textarea id="advSourceHtml" class="modal-input" data-resetable
                    placeholder="여기에 소스코드를 붙여넣으세요..." style="height:80px;resize:none;"></textarea>
            </div>`;

        const mainBtns = modal.querySelector('.main-btns');
        mainBtns.style.cssText = 'display:flex;align-items:center;width:100%;';
        mainBtns.innerHTML = `
            <div class="tool-group" style="gap:5px;">
                <input type="text" id="advFromYM" class="modal-input input-date-ym"
                    value="${initYM}" data-resetable data-reset-value="${initYM}">
                <span style="color:#ccc">→</span>
                <input type="text" id="advToYM" class="modal-input input-date-ym"
                    value="${nextYM}" data-resetable data-reset-value="${nextYM}">
            </div>
            <button type="button" id="btnSave" class="btn-confirm">변환</button>`;

        modal.querySelector('#btnGenBase').addEventListener('click', () => {
            const ym = modal.querySelector('#calBaseYM').value;
            if (!isValidYearMonth(ym)) { window.showToast('01월부터 12월 사이로 입력해주세요.'); return; }
            const html = generateBaseCalendar(ym, {
                showHoliday: modal.querySelector('#calShowHoliday').checked,
                useId:       modal.querySelector('#calTargetId').checked,
            });
            if (typeof editor !== 'undefined') { insertFormattedHtml(html); modal.style.display = 'none'; }
        });

        modal.querySelector('#btnSave').addEventListener('click', () => {
            const src  = modal.querySelector('#advSourceHtml').value;
            const from = modal.querySelector('#advFromYM').value;
            const to   = modal.querySelector('#advToYM').value;
            if (!src.trim()) { window.showToast('변환할 소스코드를 입력해주세요.'); return; }
            if (!isValidYearMonth(from) || !isValidYearMonth(to)) { window.showToast('1월부터 12월 사이로 입력해주세요.'); return; }
            try {
                insertFormattedHtml(transformAdvancedCalendar(src, from, to));
                modal.style.display = 'none';
            } catch (e) {
                window.showToast('변환 중 오류가 발생했습니다. 소스코드를 확인해주세요.');
            }
        });

        document.body.appendChild(modal);
        modal.style.display = 'flex';
    },

    openExtendRowModal() {
        let modal = document.getElementById('extendRowModal');
        if (modal) { modal.style.display = 'flex'; return; }

        modal = this.createBase('extendRowModal', '줄 확장', 'add_row_below');
        const today = new Date();
        const currentYM = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;

        modal.querySelector('.modal-body').innerHTML = `
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
                    <span class="id-prefix">user_content_</span>
                    <input type="text" id="modalDateId" class="id-inner-input"
                        placeholder="(입력)날짜 형태로 저장됩니다. 예)d → d1, day → day1 등." data-resetable>
                </div>
                <p class="helper-text">* 이미 샘플 코드에 Id 선택자가 있으면 사용 할 필요 없음</p>
            </div>
            <div class="modal-section">
                <label class="modal-label">날짜칸 색상</label>
                <div class="input-row" style="display:flex;gap:10px;align-items:center;width:100%;">
                    <input type="text" id="modalTargetAttr" class="modal-input"
                        placeholder="#333333(평일), #0000FF(토), #FF0000(일)"
                        style="flex:7;min-width:0;" data-resetable>
                    <input type="month" id="modalBaseMonth" class="modal-input"
                        style="flex:3;min-width:0;cursor:pointer;text-align:center;"
                        value="${currentYM}" data-resetable data-reset-value="${currentYM}">
                </div>
                <p class="helper-text">* #헥스코드로만 입력해주세요(rgb 안 됨) 쉼표 구분 필수.</p>
                <p class="helper-text">* 1개만 입력하면 색 일괄 통일, 2개 입력 시 일요일만 색 구분</p>
            </div>`;

        modal.querySelector('.btn-confirm').onclick = () => {
            const from = parseInt(document.getElementById('modalExtendFrom').value, 10);
            const to   = parseInt(document.getElementById('modalExtendTo').value, 10);
            if (from > to) { window.showToast('종료일은 시작일보다 커야합니다.'); return; }
            if (typeof window.executeExtendRow === 'function') {
                window.executeExtendRow();
                ModalManager.close('extendRowModal');
            }
        };
        document.getElementById('modalContainer').appendChild(modal);
        modal.style.display = 'flex';
    },

    openLinkModal() {
        let modal = document.getElementById('linkModal');
        if (modal) {
            modal.style.display = 'flex';
            const targetBlank = document.getElementById('modalTargetBlank');
            const underline   = document.getElementById('modalUnderline');
            if (targetBlank) targetBlank.checked = true;
            if (underline)   underline.checked   = true;
            return;
        }
        modal = this.createBase('linkModal', '링크 삽입/변경', 'link');
        modal.querySelector('.modal-body').innerHTML = `
            <div class="modal-section">
                <label class="modal-label">보이는 글자 (또는 이미지 주소)</label>
                <input type="text" id="modalTextDisplay" class="modal-input"
                    placeholder="표시 될 텍스트(입력도 가능)" data-resetable>
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
                    <span class="id-prefix">user_content_</span>
                    <input type="text" id="modalTdId" class="id-inner-input"
                        placeholder="Id값을 입력하세요" data-resetable>
                </div>
                <p class="helper-text">* td 칸 안에 있을 경우 해당 칸의 Id로 저장됩니다.</p>
            </div>`;
        modal.querySelector('.btn-confirm').onclick = () => {
            if (typeof window.applyLinkChanges === 'function') window.applyLinkChanges();
        };
        document.getElementById('modalContainer').appendChild(modal);
        modal.style.display = 'flex';
    },
};

window.openModal = function (id) {
    const map = {
        analysisModal:  () => ModalManager.openAnalysisModal(),
        ruleModal:      () => ModalManager.openRuleModal(),
        calendarModal:  () => ModalManager.openCalendarModal(),
        extendRowModal: () => ModalManager.openExtendRowModal(),
        linkModal:      () => {
            ModalManager.openLinkModal();
            if (typeof window.prepareLinkData === 'function') window.prepareLinkData();
        },
    };
    const opener = map[id];
    if (opener) opener();
};
window.closeModal = (id) => ModalManager.close(id);