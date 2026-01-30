import { clear, div, toast, showExcelSavedDialog, chooseExcelMode } from '../core/dom.js';
import { ProgressDialog } from '../core/progress.js';
import { post, onHost } from '../core/bridge.js';
import { createRvtTable, renderRvtRows, getRvtName } from './rvtTable.js';

const DEFAULT_SUMMARY = { okCount: 0, failCount: 0, skipCount: 0 };

export function renderSharedParamBatch(root) {
  const target = root || document.getElementById('view-root') || document.getElementById('app');
  clear(target);
  const top = document.querySelector('#topbar-root .topbar') || document.querySelector('.topbar');
  if (top) top.classList.add('hub-topbar');

  const state = {
    spFilePath: '',
    groups: [],
    defsByGroup: {},
    categoryTree: [],
    paramGroups: [],
    selectedGroup: '',
    selectedParams: [],
    defSelection: new Set(),
    defSearch: '',
    rvtList: [],
    rvtChecked: new Set(),
    pendingFolderBrowse: false,
    options: {
      closeAllWorksetsOnOpen: true,
      syncComment: ''
    },
    summary: { ...DEFAULT_SUMMARY },
    logs: [],
    logTextPath: '',
    running: false,
    lastProgressPct: 0
  };

  const page = div('feature-shell sharedparambatch-page');
  const header = div('feature-header');
  const heading = div('feature-heading');
  heading.innerHTML = `
    <span class="feature-kicker">Project Parameter</span>
    <h2 class="feature-title">Project Parameter 추가 (Project/Shared)</h2>
    <p class="feature-sub">Project/Shared 파라미터를 여러 RVT에 일괄 추가/바인딩합니다.</p>`;

  const btnRun = cardBtn('실행', onRun);
  const btnExport = cardBtn('엑셀 내보내기', onExport, 'btn--secondary');
  btnExport.disabled = true;

  header.append(heading);
  page.append(header);

  const layout = div('sharedparambatch-layout');
  page.append(layout);

  const warningSection = div('section sharedparambatch-section spb-warning');
  const warningText = document.createElement('div');
  warningText.textContent = 'Shared Parameters TXT가 Revit에 설정되어 있지 않습니다. Manage > Shared Parameters에서 등록 후 새로고침하세요.';
  const warningActions = div('section-actions');
  warningActions.append(cardBtn('새로고침', () => post('sharedparambatch:init', {}), 'btn--secondary'));
  warningSection.append(warningText, warningActions);
  warningSection.style.display = 'none';

  const selectSection = div('section sharedparambatch-section');
  const paramPickerBtn = cardBtn('Parameter 선택하기', openParamPicker, 'btn--secondary');
  const selectHeader = buildSelectHeader('Shared Parameter 선택', paramPickerBtn, btnRun, btnExport);
  selectSection.append(selectHeader);

  const selectedSection = div('section sharedparambatch-section');
  const bulkApplyBtn = cardBtn('설정 일괄 적용', applyBulkSettings, 'btn--secondary');
  selectedSection.append(sectionHeader('Selected Parameters', [bulkApplyBtn]));
  const paramTable = document.createElement('table');
  paramTable.className = 'sharedparambatch-table';
  paramTable.innerHTML = '<thead><tr><th>Name</th><th>GUID</th><th>Binding</th><th>Group</th><th>Categories</th><th>Action</th></tr></thead><tbody></tbody>';
  const paramBody = paramTable.querySelector('tbody');
  const paramWrap = div('spb-tableWrap');
  paramWrap.append(paramTable);
  selectedSection.append(paramWrap);

  const rvtSection = div('section sharedparambatch-section');
  const rvtHeader = div('spb-rvtHeader');
  const rvtHeaderLeft = div('spb-rvtHeaderLeft');
  const rvtTitle = document.createElement('h3');
  rvtTitle.textContent = 'RVT 파일';
  const rvtActions = div('section-actions');
  rvtActions.append(
    cardBtn('추가', () => post('sharedparambatch:browse-rvts', {})),
    cardBtn('폴더 선택', onBrowseFolder),
    cardBtn('선택 삭제', removeSelectedRvts, 'btn--secondary'),
    cardBtn('전체 삭제', clearRvts, 'btn--secondary'),
    cardBtn('리스트 크게보기', openRvtModal, 'btn--secondary')
  );
  rvtHeaderLeft.append(rvtTitle, rvtActions);

  const rvtHeaderRight = div('spb-rvtHeaderRight');
  const closeWrap = document.createElement('label');
  closeWrap.className = 'spb-inlineCheck';
  const closeChk = document.createElement('input');
  closeChk.type = 'checkbox';
  closeChk.checked = true;
  closeChk.id = 'spb-close-worksets';
  closeChk.addEventListener('change', () => { state.options.closeAllWorksetsOnOpen = !!closeChk.checked; });
  const closeLbl = document.createElement('span');
  closeLbl.textContent = 'Workshared: Open CloseAllWorksets';
  closeWrap.append(closeChk, closeLbl);

  const syncInput = document.createElement('input');
  syncInput.type = 'text';
  syncInput.className = 'sharedparambatch-input spb-syncInput';
  syncInput.placeholder = 'Sync Comment';
  syncInput.addEventListener('input', () => { state.options.syncComment = syncInput.value || ''; });
  rvtHeaderRight.append(closeWrap, labelSpan('Sync Comment'), syncInput);

  rvtHeader.append(rvtHeaderLeft, rvtHeaderRight);
  rvtSection.append(rvtHeader);
  const { table: rvtTable, tbody: rvtBody, master: rvtMaster } = createRvtTable();
  const rvtListWrap = div('spb-rvtTableWrap');
  rvtListWrap.append(rvtTable);
  rvtSection.append(rvtListWrap);

  const rvtModal = buildRvtModal();
  const paramPickerModal = buildParamPickerModal();

  const resultSection = div('sharedparambatch-result');
  const resultHeader = sectionHeader('최근 실행 결과', []);
  resultSection.append(resultHeader);
  const summaryRow = div('sharedparambatch-summary');
  const badgeOk = summaryBadge('OK', '0');
  const badgeFail = summaryBadge('FAIL', '0');
  const badgeSkip = summaryBadge('SKIP', '0');
  summaryRow.append(badgeOk.wrap, badgeFail.wrap, badgeSkip.wrap);

  const logPathRow = div('sharedparambatch-logpath');
  const logPathText = document.createElement('span');
  logPathText.textContent = '로그 파일: -';
  const logOpenBtn = cardBtn('폴더 열기', () => {
    if (!state.logTextPath) { toast('로그 파일이 없습니다.', 'err'); return; }
    post('sharedparambatch:open-folder', { path: state.logTextPath });
  }, 'btn--secondary');
  logOpenBtn.disabled = true;
  logPathRow.append(logPathText, logOpenBtn);

  const logTable = document.createElement('table');
  logTable.className = 'sharedparambatch-table';
  logTable.innerHTML = '<thead><tr><th>Level</th><th>RVT</th><th>Message</th></tr></thead><tbody></tbody>';
  const logBody = logTable.querySelector('tbody');

  resultSection.append(summaryRow, logPathRow, logTable);
  resultSection.style.display = 'none';

  selectSection.append(resultSection);
  layout.append(warningSection, selectSection, selectedSection, rvtSection);
  page.append(layout);
  page.append(buildSettingsModal(), paramPickerModal, rvtModal);

  target.append(page);

  onHost('sharedparambatch:init', handleInit);
  onHost('sharedparambatch:rvts-picked', handleRvtsPicked);
  onHost('sharedparambatch:progress', handleProgress);
  onHost('sharedparambatch:done', handleDone);
  onHost('sharedparambatch:exported', handleExported);
  onHost('revit:error', ({ message }) => { finishRunning(false); toast(message || 'Revit 오류가 발생했습니다.', 'err', 3200); });
  onHost('host:error', ({ message }) => { finishRunning(false); toast(message || '호스트 오류가 발생했습니다.', 'err', 3200); });

  post('sharedparambatch:init', {});
  renderRvtList();

  function handleInit(payload) {
    if (!payload || !payload.ok) {
      warningSection.style.display = 'flex';
      toast(payload?.message || '초기화 실패', 'err');
      return;
    }
    warningSection.style.display = 'none';
    state.spFilePath = payload.spFilePath || '';
    state.groups = Array.isArray(payload.groups) ? payload.groups : [];
    state.defsByGroup = payload.defsByGroup || {};
    state.categoryTree = Array.isArray(payload.categoryTree) ? payload.categoryTree : [];
    state.paramGroups = Array.isArray(payload.paramGroups) ? payload.paramGroups : [];
    if (syncInput) syncInput.value = state.options.syncComment || '';
    closeChk.checked = !!state.options.closeAllWorksetsOnOpen;
    renderGroupOptions();
    renderDefinitionList();
    renderSelectedParams();
  }

  function renderGroupOptions() {
    const modal = buildParamPickerModal;
    const groupSelect = modal.groupSelect;
    if (!groupSelect) return;
    groupSelect.innerHTML = '';
    if (!state.groups.length) {
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = '그룹 없음';
      groupSelect.append(opt);
      state.selectedGroup = '';
      return;
    }
    state.groups.forEach((g, idx) => {
      const opt = document.createElement('option');
      opt.value = g;
      opt.textContent = g;
      groupSelect.append(opt);
      if (idx === 0) state.selectedGroup = g;
    });
    groupSelect.value = state.selectedGroup;
  }

  function renderDefinitionList() {
    const modal = buildParamPickerModal;
    const defList = modal.defList;
    if (!defList) return;
    defList.innerHTML = '';
    const defs = state.defsByGroup[state.selectedGroup] || [];
    const search = (state.defSearch || '').toLowerCase();
    const filtered = search
      ? defs.filter(d => (d.name || '').toLowerCase().includes(search))
      : defs;
    if (!filtered.length) {
      const opt = document.createElement('option');
      opt.value = '';
      opt.textContent = '정의 없음';
      defList.append(opt);
      return;
    }
    filtered.forEach((d) => {
      const opt = document.createElement('option');
      opt.value = d.guid;
      opt.textContent = `${d.name} (${d.paramTypeLabel || ''})`;
      opt.dataset.name = d.name;
      opt.dataset.group = state.selectedGroup;
      opt.dataset.guid = d.guid;
      opt.dataset.paramTypeLabel = d.paramTypeLabel || '';
      opt.dataset.desc = d.desc || '';
      if (state.defSelection.has(d.guid)) opt.selected = true;
      defList.append(opt);
    });
  }

  function openParamPicker() {
    const modal = buildParamPickerModal;
    if (modal.searchInput) modal.searchInput.value = state.defSearch || '';
    if (modal.groupSelect) modal.groupSelect.value = state.selectedGroup || '';
    renderDefinitionList();
    if (modal.overlay) modal.overlay.classList.add('is-open');
  }

  function onAddSelectedDefs() {
    const modal = buildParamPickerModal;
    const defList = modal.defList;
    if (!defList || !defList.options.length) return;
    const selected = Array.from(defList.selectedOptions || []);
    if (!selected.length) { toast('추가할 파라미터를 선택하세요.', 'err'); return; }

    const nextSelection = new Set(state.defSelection);
    selected.forEach((opt) => {
      const guid = opt.dataset.guid;
      if (!guid || state.selectedParams.some(p => p.guid === guid)) return;
      nextSelection.add(guid);
      state.selectedParams.push({
        groupName: opt.dataset.group || state.selectedGroup,
        name: opt.dataset.name || opt.textContent,
        guid,
        paramTypeLabel: opt.dataset.paramTypeLabel || '',
        desc: opt.dataset.desc || '',
        settings: {
          isInstanceBinding: true,
          paramGroup: state.paramGroups[0]?.id || 'PG_DATA',
          allowVaryBetweenGroups: false,
          categories: []
        }
      });
    });
    state.defSelection = nextSelection;
    renderSelectedParams();
  }

  function renderSelectedParams() {
    paramBody.innerHTML = '';
    if (!state.selectedParams.length) {
      const tr = document.createElement('tr');
      tr.className = 'empty-row';
      const td = document.createElement('td');
      td.className = 'empty-cell';
      td.colSpan = 6;
      td.textContent = '선택된 파라미터가 없습니다.';
      tr.append(td);
      paramBody.append(tr);
      updateButtons();
      return;
    }

    state.selectedParams.forEach((p, idx) => {
      const tr = document.createElement('tr');
      tr.append(tdText(p.name));
      tr.append(tdText(p.guid));
      tr.append(tdText(p.settings.isInstanceBinding ? 'Instance' : 'Type'));
      tr.append(tdText(formatParamGroup(p.settings.paramGroup)));
      tr.append(tdText(String(p.settings.categories.length)));

      const actionTd = document.createElement('td');
      const btnSetting = actionBtn('설정', () => openSettingsModal(idx));
      const btnRemove = actionBtn('삭제', () => {
        state.selectedParams.splice(idx, 1);
        renderSelectedParams();
      }, 'btn--ghost');
      actionTd.append(btnSetting, btnRemove);
      tr.append(actionTd);
      paramBody.append(tr);
    });
    updateButtons();
  }

  function renderRvtList() {
    const allChecked = state.rvtList.length > 0 && state.rvtList.every(f => state.rvtChecked.has(f));
    [rvtMaster, buildRvtModal.master].forEach((master) => {
      if (!master) return;
      master.checked = allChecked;
      master.disabled = state.rvtList.length === 0;
      master.onchange = () => {
        if (master.checked) state.rvtChecked = new Set(state.rvtList);
        else state.rvtChecked.clear();
        renderRvtList();
        updateButtons();
      };
    });
    const rows = state.rvtList.map((p, idx) => ({
      checked: state.rvtChecked.has(p),
      index: idx + 1,
      name: getRvtName(p),
      path: p,
      title: p,
      onToggle: (checked) => {
        if (checked) state.rvtChecked.add(p);
        else state.rvtChecked.delete(p);
        updateButtons();
      }
    }));
    renderRvtRows(rvtBody, rows);
    if (buildRvtModal.tbody) {
      renderRvtRows(buildRvtModal.tbody, rows);
    }
    updateButtons();
  }

  function handleRvtsPicked(payload) {
    if (state.pendingFolderBrowse) {
      ProgressDialog.hide();
      state.pendingFolderBrowse = false;
    }
    if (!payload || !payload.ok) {
      if (payload?.message) toast(payload.message, 'err');
      return;
    }
    const paths = Array.isArray(payload.rvtPaths) ? payload.rvtPaths : [];
    let changed = false;
    let added = 0;
    paths.forEach((p) => {
      if (!state.rvtList.includes(p)) {
        state.rvtList.push(p);
        state.rvtChecked.add(p);
        changed = true;
        added += 1;
      }
    });
    if (changed) renderRvtList();
    if (payload?.fromFolder) {
      toast(added ? `${added}개 추가됨` : '추가된 RVT가 없습니다.', added ? 'ok' : 'err');
    }
  }

  function removeSelectedRvts() {
    if (!state.rvtChecked.size) { toast('삭제할 RVT를 선택하세요.', 'err'); return; }
    state.rvtList = state.rvtList.filter(p => !state.rvtChecked.has(p));
    state.rvtChecked.clear();
    renderRvtList();
  }

  function clearRvts() {
    state.rvtList = [];
    state.rvtChecked.clear();
    renderRvtList();
  }

  function onBrowseFolder() {
    if (state.running) return;
    state.pendingFolderBrowse = true;
    ProgressDialog.show('RVT 폴더 선택', '폴더 내 RVT 파일을 찾는 중...');
    ProgressDialog.update(20, '폴더 내 RVT 파일을 찾는 중...', '');
    post('sharedparambatch:browse-folder', {});
  }

  function onRun() {
    if (state.running) return;
    if (!state.selectedParams.length) { toast('파라미터를 선택하세요.', 'err'); return; }
    if (!state.rvtList.length) { toast('RVT 파일을 추가하세요.', 'err'); return; }
    const missingCats = state.selectedParams.filter(p => !p.settings.categories.length);
    if (missingCats.length) {
      toast('카테고리를 지정하지 않은 파라미터가 있습니다.', 'err');
      return;
    }

    const payload = {
      spFilePath: state.spFilePath,
      rvtPaths: state.rvtList,
      parameters: state.selectedParams.map(p => ({
        groupName: p.groupName,
        paramName: p.name,
        guid: p.guid,
        paramTypeLabel: p.paramTypeLabel,
        description: p.desc,
        settings: {
          isInstanceBinding: p.settings.isInstanceBinding,
          paramGroup: p.settings.paramGroup,
          allowVaryBetweenGroups: p.settings.allowVaryBetweenGroups,
          categories: p.settings.categories
        }
      })),
      closeAllWorksetsOnOpen: state.options.closeAllWorksetsOnOpen,
      syncComment: state.options.syncComment
    };

    state.running = true;
    updateButtons();
    ProgressDialog.show('Project Parameter 추가', '작업 준비 중...');
    ProgressDialog.update(0, '작업 준비 중...', '');
    post('sharedparambatch:run', payload);
  }

  function handleProgress(payload) {
    if (!payload) return;
    const total = Number(payload.total || payload.Total || 0);
    const step = Number(payload.step || payload.Step || 0);
    const text = payload.text || payload.message || '';
    const percentFromStep = total > 0 ? (step / total) * 100 : 0;
    const percentFromPhase = Number(payload.phaseProgress) * 100;
    const percentRaw = Number.isFinite(payload.percent) ? Number(payload.percent) : (Number.isFinite(percentFromPhase) && percentFromPhase > 0 ? percentFromPhase : percentFromStep);
    const pct = Math.max(0, Math.min(100, percentRaw || 0));
    state.lastProgressPct = pct;
    ProgressDialog.show('Project Parameter 추가', text || '진행 중');
    ProgressDialog.update(pct, text || '진행 중', total ? `${step} / ${total}` : '');
    if (pct >= 100) ProgressDialog.hide();
  }

  function handleDone(payload) {
    finishRunning(true);
    if (!payload || !payload.ok) {
      toast(payload?.message || '실행 실패', 'err');
      return;
    }
    state.summary = payload.summary || { ...DEFAULT_SUMMARY };
    state.logs = Array.isArray(payload.logs) ? payload.logs : [];
    state.logTextPath = payload.logTextPath || '';
    renderSummary();
    renderLogs();
    if (state.logs.length) toast('완료되었습니다.', 'ok');
  }

  function finishRunning(resetProgress) {
    state.running = false;
    updateButtons();
    if (resetProgress) {
      state.lastProgressPct = 0;
      ProgressDialog.hide();
    }
  }

  function onExport() {
    if (!state.logs.length) { toast('내보낼 로그가 없습니다.', 'err'); return; }
    chooseExcelMode((mode) => post('sharedparambatch:export-excel', { excelMode: mode || 'fast' }));
  }

  function handleExported(payload) {
    if (!payload || !payload.ok) {
      toast(payload?.message || '엑셀 내보내기 실패', 'err');
      return;
    }
    const path = payload.filePath || payload.path;
    showExcelSavedDialog('엑셀로 내보냈습니다.', path, (p) => post('excel:open', { path: p }));
  }

  function renderSummary() {
    badgeOk.value.textContent = String(state.summary.okCount ?? 0);
    badgeFail.value.textContent = String(state.summary.failCount ?? 0);
    badgeSkip.value.textContent = String(state.summary.skipCount ?? 0);
    logPathText.textContent = state.logTextPath ? `로그 파일: ${state.logTextPath}` : '로그 파일: -';
    logOpenBtn.disabled = !state.logTextPath;
    btnExport.disabled = !state.logs.length;
    resultSection.style.display = state.logs.length ? 'flex' : 'none';
  }

  function renderLogs() {
    logBody.innerHTML = '';
    if (!state.logs.length) {
      const tr = document.createElement('tr');
      tr.className = 'empty-row';
      const td = document.createElement('td');
      td.className = 'empty-cell';
      td.colSpan = 3;
      td.textContent = '로그가 없습니다.';
      tr.append(td);
      logBody.append(tr);
      return;
    }
    state.logs.forEach((log) => {
      const tr = document.createElement('tr');
      tr.append(tdText(log.level || ''));
      tr.append(tdText(log.file || ''));
      tr.append(tdText(log.msg || ''));
      logBody.append(tr);
    });
  }

  function updateButtons() {
    const disabled = state.running;
    btnRun.disabled = disabled || !state.selectedParams.length || !state.rvtList.length;
    btnExport.disabled = disabled || !state.logs.length;
    const modal = buildParamPickerModal;
    if (modal.addBtn) modal.addBtn.disabled = disabled;
    bulkApplyBtn.disabled = disabled || state.selectedParams.length < 2;
  }

  function formatParamGroup(value) {
    const match = state.paramGroups.find(g => g.key === value || g.Key === value || g.id === value || g.Id === value);
    if (match) return match.label || match.Label || value || '';
    return value || '';
  }

  function applyBulkSettings() {
    if (state.selectedParams.length < 2) return;
    const template = state.selectedParams[0]?.settings;
    if (!template) return;
    const cloned = cloneSettings(template);
    state.selectedParams.forEach((p, idx) => {
      if (idx === 0 || !p) return;
      p.settings = cloneSettings(cloned);
    });
    renderSelectedParams();
  }

  function cloneSettings(src) {
    return {
      isInstanceBinding: !!src.isInstanceBinding,
      paramGroup: src.paramGroup,
      allowVaryBetweenGroups: !!src.allowVaryBetweenGroups,
      categories: Array.isArray(src.categories) ? src.categories.map(c => ({
        idInt: c.idInt,
        name: c.name,
        path: c.path
      })) : []
    };
  }

  function buildSettingsModal() {
    const overlay = div('sharedparambatch-modal-overlay');
    const modal = div('sharedparambatch-modal');
    const header = div('sharedparambatch-modal__header');
    const title = div('sharedparambatch-modal__title');
    const closeBtn = actionBtn('닫기', closeSettingsModal, 'btn--ghost');
    header.append(title, closeBtn);

    const body = div('sharedparambatch-modal__body');
    const footer = div('sharedparambatch-modal__footer');
    const saveBtn = cardBtn('저장', saveSettings);
    footer.append(saveBtn);

    modal.append(header, body, footer);
    overlay.append(modal);

    overlay.addEventListener('click', (ev) => { if (ev.target === overlay) closeSettingsModal(); });

    buildSettingsModal.overlay = overlay;
    buildSettingsModal.title = title;
    buildSettingsModal.body = body;
    buildSettingsModal.saveBtn = saveBtn;
    buildSettingsModal.currentIndex = -1;

    return overlay;
  }

  function buildSelectHeader(titleText, pickerButton, runButton, exportButton) {
    const header = div('spb-cardHeader');
    const left = div('spb-headerLeft');
    const title = document.createElement('h3');
    title.textContent = titleText;
    left.append(title, pickerButton);
    const right = div('spb-headerRight');
    right.append(runButton, exportButton);
    header.append(left, right);
    return header;
  }

  function buildParamPickerModal() {
    const overlay = div('sharedparambatch-modal-overlay');
    const modal = div('sharedparambatch-modal');
    const header = div('sharedparambatch-modal__header');
    const title = div('sharedparambatch-modal__title');
    title.textContent = '파라미터 선택';
    const closeBtn = actionBtn('닫기', closeParamPicker, 'btn--ghost');
    header.append(title, closeBtn);

    const body = div('sharedparambatch-modal__body');
    const footer = div('sharedparambatch-modal__footer');
    const addBtn = cardBtn('선택 추가', onAddSelectedDefs);
    const closeBtn2 = cardBtn('닫기', closeParamPicker, 'btn--secondary');
    footer.append(addBtn, closeBtn2);

    const groupSelect = document.createElement('select');
    groupSelect.className = 'sharedparambatch-select spb-groupSelect';
    groupSelect.addEventListener('change', () => {
      state.selectedGroup = groupSelect.value;
      renderDefinitionList();
    });

    const searchInput = document.createElement('input');
    searchInput.type = 'search';
    searchInput.className = 'sharedparambatch-input spb-search';
    searchInput.placeholder = '파라미터 검색…';
    searchInput.addEventListener('input', () => {
      state.defSearch = (searchInput.value || '').trim();
      renderDefinitionList();
    });

    const defHeader = div('spb-defHeader');
    defHeader.append(labelSpan('Group'), groupSelect, searchInput);

    const defList = document.createElement('select');
    defList.className = 'sharedparambatch-def-list spb-defList';
    defList.multiple = true;
    defList.addEventListener('change', () => {
      const next = new Set();
      Array.from(defList.selectedOptions || []).forEach((opt) => {
        const guid = opt.dataset.guid;
        if (guid) next.add(guid);
      });
      state.defSelection = next;
    });

    const defBox = div('sharedparambatch-card');
    defBox.append(defHeader, defList);

    body.append(defBox);
    modal.append(header, body, footer);
    overlay.append(modal);

    overlay.addEventListener('click', (ev) => { if (ev.target === overlay) closeParamPicker(); });

    buildParamPickerModal.overlay = overlay;
    buildParamPickerModal.groupSelect = groupSelect;
    buildParamPickerModal.defList = defList;
    buildParamPickerModal.searchInput = searchInput;
    buildParamPickerModal.addBtn = addBtn;

    return overlay;
  }

  function closeParamPicker() {
    const modal = buildParamPickerModal;
    if (modal.overlay) modal.overlay.classList.remove('is-open');
  }

  function buildRvtModal() {
    const overlay = div('sharedparambatch-modal-overlay');
    const modal = div('sharedparambatch-modal spb-modalLarge');
    const header = div('sharedparambatch-modal__header');
    const title = div('sharedparambatch-modal__title');
    title.textContent = 'RVT 파일 목록';
    const closeBtn = actionBtn('닫기', closeRvtModal, 'btn--ghost');
    header.append(title, closeBtn);

    const body = div('sharedparambatch-modal__body');
    const actions = div('section-actions');
    actions.append(
      cardBtn('추가', () => post('sharedparambatch:browse-rvts', {})),
      cardBtn('폴더 선택', onBrowseFolder),
      cardBtn('선택 삭제', removeSelectedRvts, 'btn--secondary'),
      cardBtn('전체 삭제', clearRvts, 'btn--secondary')
    );
    const { table, tbody, master } = createRvtTable();
    const wrap = div('spb-rvtTableWrap');
    wrap.append(table);
    body.append(actions, wrap);

    modal.append(header, body);
    overlay.append(modal);

    overlay.addEventListener('click', (ev) => { if (ev.target === overlay) closeRvtModal(); });

    buildRvtModal.overlay = overlay;
    buildRvtModal.tbody = tbody;
    buildRvtModal.master = master;

    return overlay;
  }

  function openRvtModal() {
    const modal = buildRvtModal;
    if (modal.overlay) modal.overlay.classList.add('is-open');
    renderRvtList();
  }

  function closeRvtModal() {
    const modal = buildRvtModal;
    if (modal.overlay) modal.overlay.classList.remove('is-open');
  }

  function openSettingsModal(index) {
    const modal = buildSettingsModal;
    if (!modal.body || !state.selectedParams[index]) return;

    const param = state.selectedParams[index];
    modal.currentIndex = index;
    modal.title.textContent = `${param.name} 설정`;
    modal.body.innerHTML = '';

    const bindingRow = div('sharedparambatch-modal-row');
    bindingRow.append(subTitle('Binding'));
    const bindingOpts = div('sharedparambatch-binding');
    const instRadio = radio('spb-binding', 'Instance', param.settings.isInstanceBinding, (checked) => {
      if (checked) param.settings.isInstanceBinding = true;
      toggleVary();
    });
    const typeRadio = radio('spb-binding', 'Type', !param.settings.isInstanceBinding, (checked) => {
      if (checked) param.settings.isInstanceBinding = false;
      toggleVary();
    });
    bindingOpts.append(instRadio.wrap, typeRadio.wrap);
    bindingRow.append(bindingOpts);

    const groupRow = div('sharedparambatch-modal-row');
    groupRow.append(subTitle('Parameter Group'));
    const groupSelectEl = document.createElement('select');
    groupSelectEl.className = 'sharedparambatch-select';
    state.paramGroups.forEach((g) => {
      const opt = document.createElement('option');
      opt.value = g.key || g.Key || g.id || g.Id;
      opt.textContent = g.label || g.Label || opt.value;
      groupSelectEl.append(opt);
    });
    groupSelectEl.value = param.settings.paramGroup;
    groupSelectEl.addEventListener('change', () => { param.settings.paramGroup = groupSelectEl.value; });
    groupRow.append(groupSelectEl);

    const varyRow = div('sharedparambatch-modal-row');
    const varyWrap = div('sharedparambatch-option');
    const varyChk = document.createElement('input');
    varyChk.type = 'checkbox';
    varyChk.checked = !!param.settings.allowVaryBetweenGroups;
    varyChk.id = 'spb-vary';
    varyChk.addEventListener('change', () => { param.settings.allowVaryBetweenGroups = !!varyChk.checked; });
    const varyLbl = document.createElement('label');
    varyLbl.setAttribute('for', 'spb-vary');
    varyLbl.textContent = 'Vary between groups';
    varyWrap.append(varyChk, varyLbl);
    varyRow.append(varyWrap);

    const categoryRow = div('sharedparambatch-modal-row');
    categoryRow.append(subTitle('Categories'));
    const categoryActions = div('sharedparambatch-category-actions');
    const btnAll = cardBtn('전체 선택', () => selectAllCategories(param), 'btn--secondary');
    const btnClear = cardBtn('전체 해제', () => clearAllCategories(param), 'btn--secondary');
    const btnClearChildren = cardBtn('서브카테고리 해제', () => clearSubCategories(param), 'btn--secondary');
    categoryActions.append(btnAll, btnClear, btnClearChildren);
    categoryRow.append(categoryActions);

    const treeWrap = div('sharedparambatch-category-tree');
    renderCategoryTree(treeWrap, param);
    categoryRow.append(treeWrap);

    modal.body.append(bindingRow, groupRow, varyRow, categoryRow);
    toggleVary();

    if (modal.overlay) modal.overlay.classList.add('is-open');

    function toggleVary() {
      varyChk.disabled = !param.settings.isInstanceBinding;
      if (!param.settings.isInstanceBinding) {
        param.settings.allowVaryBetweenGroups = false;
        varyChk.checked = false;
      }
    }
  }

  function closeSettingsModal() {
    const modal = buildSettingsModal;
    if (modal.overlay) modal.overlay.classList.remove('is-open');
    modal.currentIndex = -1;
  }

  function saveSettings() {
    closeSettingsModal();
    renderSelectedParams();
  }

  function renderCategoryTree(container, param) {
    container.innerHTML = '';
    if (!state.categoryTree.length) {
      container.textContent = '카테고리 정보가 없습니다.';
      return;
    }
    const selected = new Set(param.settings.categories.map(c => c.path || c.name));
    const list = document.createElement('ul');
    list.className = 'sharedparambatch-tree';
    state.categoryTree.forEach((node) => list.append(buildCategoryNode(node)));
    container.append(list);

    function buildCategoryNode(node) {
      const li = document.createElement('li');
      const row = div('sharedparambatch-tree-row');
      if (node.isBindable) {
        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.checked = selected.has(node.path || node.name);
        chk.addEventListener('change', () => {
          if (chk.checked) {
            addCategory(param, node);
          } else {
            removeCategory(param, node);
          }
        });
        row.append(chk);
      }
      const label = document.createElement('span');
      label.textContent = node.name || '';
      row.append(label);
      li.append(row);

      if (node.children && node.children.length) {
        const childList = document.createElement('ul');
        childList.className = 'sharedparambatch-tree';
        node.children.forEach((child) => childList.append(buildCategoryNode(child)));
        li.append(childList);
      }
      return li;
    }
  }

  function addCategory(param, node) {
    if (!param.settings.categories.find(c => c.path === node.path)) {
      param.settings.categories.push({
        idInt: node.idInt,
        name: node.name,
        path: node.path
      });
    }
  }

  function removeCategory(param, node) {
    param.settings.categories = param.settings.categories.filter(c => c.path !== node.path);
  }

  function selectAllCategories(param) {
    const all = [];
    collectBindable(state.categoryTree, all);
    param.settings.categories = all.map((n) => ({ idInt: n.idInt, name: n.name, path: n.path }));
    const modal = buildSettingsModal;
    if (modal.body) renderCategoryTree(modal.body.querySelector('.sharedparambatch-category-tree'), param);
  }

  function clearAllCategories(param) {
    param.settings.categories = [];
    const modal = buildSettingsModal;
    if (modal.body) renderCategoryTree(modal.body.querySelector('.sharedparambatch-category-tree'), param);
  }

  function clearSubCategories(param) {
    const targets = new Set();
    collectBindableWithDepth(state.categoryTree, targets, 0);
    param.settings.categories = param.settings.categories.filter(c => !targets.has(c.path || c.name));
    const modal = buildSettingsModal;
    if (modal.body) renderCategoryTree(modal.body.querySelector('.sharedparambatch-category-tree'), param);
  }

  function collectBindable(nodes, out) {
    nodes.forEach((n) => {
      if (n.isBindable) out.push(n);
      if (n.children && n.children.length) collectBindable(n.children, out);
    });
  }

  function collectBindableWithDepth(nodes, out, depth) {
    nodes.forEach((n) => {
      if (depth >= 1 && n.isBindable) out.add(n.path || n.name);
      if (n.children && n.children.length) collectBindableWithDepth(n.children, out, depth + 1);
    });
  }

  function sectionHeader(title, buttons) {
    const wrap = div('section-header');
    const h = document.createElement('h3');
    h.textContent = title;
    wrap.append(h);
    if (Array.isArray(buttons) && buttons.length) {
      const actions = div('section-actions');
      buttons.forEach(btn => actions.append(btn));
      wrap.append(actions);
    }
    return wrap;
  }

  function subTitle(text) {
    const el = document.createElement('div');
    el.className = 'sharedparambatch-subtitle';
    el.textContent = text;
    return el;
  }

  function summaryBadge(label, value) {
    const wrap = div('sharedparambatch-badge');
    const span = document.createElement('span');
    span.textContent = label;
    const val = document.createElement('strong');
    val.textContent = value;
    wrap.append(span, val);
    return { wrap, value: val };
  }

  function cardBtn(text, onClick, extraClass = '') {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = `btn ${extraClass}`.trim();
    btn.textContent = text;
    if (typeof onClick === 'function') btn.addEventListener('click', onClick);
    return btn;
  }

  function actionBtn(text, onClick, extraClass = '') {
    return cardBtn(text, onClick, `btn--small ${extraClass}`.trim());
  }

  function tdText(value) {
    const td = document.createElement('td');
    td.textContent = value == null ? '' : String(value);
    return td;
  }

  function labelSpan(text) {
    const span = document.createElement('span');
    span.className = 'sharedparambatch-label';
    span.textContent = text;
    return span;
  }

  function radio(name, label, checked, onChange) {
    const wrap = div('sharedparambatch-radio');
    const input = document.createElement('input');
    input.type = 'radio';
    input.name = name;
    input.checked = checked;
    const span = document.createElement('span');
    span.textContent = label;
    input.addEventListener('change', () => { if (input.checked && onChange) onChange(true); });
    wrap.append(input, span);
    return { wrap, input };
  }
}
