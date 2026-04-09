(() => {
  "use strict";

  const STORAGE_KEY = "allowanceRecordsV1";
  const PEOPLE_STORAGE_KEY = "allowancePeopleV1";
  const MEMO_TEMPLATES_STORAGE_KEY = "allowanceMemoTemplatesV1";
  const RECEIVER_COLORS_STORAGE_KEY = "allowanceReceiverColorsV1";
  const RECEIVER_PINS_STORAGE_KEY = "allowanceReceiverPinsV1";
  const APP_EXPORT_VERSION = 2;
  const RECEIVER_COLOR_PALETTE = [
    { accent: "#E53935", soft: "#FDECEA" },
    { accent: "#1E88E5", soft: "#EAF3FD" },
    { accent: "#FB8C00", soft: "#FFF3E7" },
    { accent: "#43A047", soft: "#EAF6ED" },
    { accent: "#8E24AA", soft: "#F5EAF9" },
    { accent: "#00ACC1", soft: "#E6F7FA" },
    { accent: "#D81B60", soft: "#FDEAF1" },
    { accent: "#3949AB", soft: "#EDEFFD" },
    { accent: "#7CB342", soft: "#F2F8E9" },
    { accent: "#F4511E", soft: "#FEEEE8" },
    { accent: "#6D4C41", soft: "#F2EEEC" },
    { accent: "#00897B", soft: "#E7F5F3" }
  ];

  /** @type {Array<Record<string, any>>} */
  let records = [];
  /** @type {Array<{name:string,canGive:boolean,canReceive:boolean}>} */
  let people = [];
  /** @type {string[]} */
  let memoTemplates = [];
  /** @type {Record<string, string>} */
  let receiverColors = {};
  /** @type {Record<string, {salt:string,hash:string,updatedAt:string}>} */
  let receiverPins = {};

  const form = document.getElementById("recordForm");
  const giverInput = document.getElementById("giver");
  const giverCandidates = document.getElementById("giverCandidates");
  const receiverInput = document.getElementById("receiver");
  const receiverCandidates = document.getElementById("receiverCandidates");
  const amountInput = document.getElementById("amount");
  const givenAtInput = document.getElementById("givenAt");
  const memoInput = document.getElementById("memo");
  const memoTemplateSelect = document.getElementById("memoTemplateSelect");
  const personForm = document.getElementById("personForm");
  const personNameInput = document.getElementById("personName");
  const personRoleSelect = document.getElementById("personRole");
  const giverPeopleList = document.getElementById("giverPeopleList");
  const receiverPeopleList = document.getElementById("receiverPeopleList");
  const memoTemplateForm = document.getElementById("memoTemplateForm");
  const memoTemplateInput = document.getElementById("memoTemplateInput");
  const memoTemplateList = document.getElementById("memoTemplateList");

  const unreceivedCountEl = document.getElementById("unreceivedCount");
  const receivedCountEl = document.getElementById("receivedCount");
  const totalAmountEl = document.getElementById("totalAmount");
  const monthlyTargetMonthInput = document.getElementById("monthlyTargetMonth");
  const monthlyReceiverFilterSelect = document.getElementById("monthlyReceiverFilter");
  const monthlyTotalAmountEl = document.getElementById("monthlyTotalAmount");
  const monthlyTotalCountEl = document.getElementById("monthlyTotalCount");
  const receiverColorForm = document.getElementById("receiverColorForm");
  const receiverColorTargetSelect = document.getElementById("receiverColorTarget");
  const receiverColorPicker = document.getElementById("receiverColorPicker");
  const resetReceiverColorBtn = document.getElementById("resetReceiverColorBtn");
  const receiverColorList = document.getElementById("receiverColorList");
  const receiverPinForm = document.getElementById("receiverPinForm");
  const receiverPinTargetSelect = document.getElementById("receiverPinTarget");
  const receiverPinInput = document.getElementById("receiverPinInput");
  const deleteReceiverPinBtn = document.getElementById("deleteReceiverPinBtn");
  const receiverPinList = document.getElementById("receiverPinList");
  const pinAuthDialog = document.getElementById("pinAuthDialog");
  const pinAuthForm = document.getElementById("pinAuthForm");
  const pinAuthReceiverText = document.getElementById("pinAuthReceiverText");
  const pinAuthInput = document.getElementById("pinAuthInput");
  const pinAuthError = document.getElementById("pinAuthError");
  const pinAuthCancelBtn = document.getElementById("pinAuthCancelBtn");

  const recordsContainer = document.getElementById("recordsContainer");
  const cardTemplate = document.getElementById("recordCardTemplate");
  const receiverFilterSelect = document.getElementById("receiverFilter");
  const receiverSummaryEl = document.getElementById("receiverSummary");
  const receiverRecordsContainer = document.getElementById("receiverRecordsContainer");
  const showListViewBtn = document.getElementById("showListViewBtn");
  const showCalendarViewBtn = document.getElementById("showCalendarViewBtn");
  const listView = document.getElementById("listView");
  const calendarView = document.getElementById("calendarView");
  const prevMonthBtn = document.getElementById("prevMonthBtn");
  const nextMonthBtn = document.getElementById("nextMonthBtn");
  const calendarMonthLabel = document.getElementById("calendarMonthLabel");
  const calendarGrid = document.getElementById("calendarGrid");
  const selectedDateLabel = document.getElementById("selectedDateLabel");
  const calendarRecordsContainer = document.getElementById("calendarRecordsContainer");

  const downloadJsonBtn = document.getElementById("downloadJsonBtn");
  const downloadCsvBtn = document.getElementById("downloadCsvBtn");
  const restoreFileInput = document.getElementById("restoreFileInput");
  const deleteAllBtn = document.getElementById("deleteAllBtn");
  const messageArea = document.getElementById("messageArea");
  let receiverFilterValue = "ALL";
  let currentView = "list";
  let calendarMonthCursor = startOfMonth(new Date());
  let selectedDateKey = toDateKey(new Date());
  let monthlyTargetMonthKey = toMonthKey(new Date());
  let monthlyReceiverFilterValue = "ALL";
  let receiverColorTargetValue = "";
  let receiverPinTargetValue = "";
  let recentAddedRecordId = "";
  /** @type {{receiverName:string,resolve:(value:boolean)=>void} | null} */
  let pinAuthContext = null;

  init();

  function init() {
    // 入力しやすいよう、初期値に現在日時(分単位)を設定
    givenAtInput.value = toDatetimeLocalValue(new Date());
    monthlyTargetMonthInput.value = monthlyTargetMonthKey;
    setupCollapsiblePanels();
    loadRecordsFromStorage();
    loadPeopleFromStorage();
    loadMemoTemplatesFromStorage();
    loadReceiverColorsFromStorage();
    loadReceiverPinsFromStorage();
    if (people.length === 0 && records.length > 0) {
      people = derivePeopleFromRecords(records);
      savePeopleToStorage();
    }
    renderAll();
    bindEvents();
    switchView(currentView);
  }

  function bindEvents() {
    form.addEventListener("submit", onSubmitRecord);
    downloadJsonBtn.addEventListener("click", exportBackupJson);
    downloadCsvBtn.addEventListener("click", exportCsv);
    restoreFileInput.addEventListener("change", onRestoreFileSelected);
    deleteAllBtn.addEventListener("click", onDeleteAll);
    receiverFilterSelect.addEventListener("change", onReceiverFilterChanged);
    personForm.addEventListener("submit", onSubmitPerson);
    giverPeopleList.addEventListener("click", onClickPersonChipDelete);
    receiverPeopleList.addEventListener("click", onClickPersonChipDelete);
    memoTemplateForm.addEventListener("submit", onSubmitMemoTemplate);
    memoTemplateList.addEventListener("click", onClickMemoTemplateDelete);
    memoTemplateSelect.addEventListener("change", onMemoTemplateSelected);
    monthlyTargetMonthInput.addEventListener("change", onMonthlyTargetMonthChanged);
    monthlyReceiverFilterSelect.addEventListener("change", onMonthlyReceiverFilterChanged);
    receiverColorForm.addEventListener("submit", onSubmitReceiverColor);
    receiverColorTargetSelect.addEventListener("change", onReceiverColorTargetChanged);
    resetReceiverColorBtn.addEventListener("click", onResetReceiverColor);
    receiverColorList.addEventListener("click", onClickReceiverColorList);
    receiverPinForm.addEventListener("submit", onSubmitReceiverPin);
    receiverPinTargetSelect.addEventListener("change", onReceiverPinTargetChanged);
    deleteReceiverPinBtn.addEventListener("click", onDeleteReceiverPin);
    receiverPinList.addEventListener("click", onClickReceiverPinList);
    pinAuthForm.addEventListener("submit", onSubmitPinAuth);
    pinAuthCancelBtn.addEventListener("click", onCancelPinAuth);
    pinAuthDialog.addEventListener("cancel", onPinAuthDialogCancel);
    showListViewBtn.addEventListener("click", () => switchView("list"));
    showCalendarViewBtn.addEventListener("click", () => switchView("calendar"));
    prevMonthBtn.addEventListener("click", () => moveCalendarMonth(-1));
    nextMonthBtn.addEventListener("click", () => moveCalendarMonth(1));
  }

  function setupCollapsiblePanels() {
    const panels = [...document.querySelectorAll("section.panel")];
    panels.forEach((panel) => {
      const sectionLabel = toSafeString(panel.getAttribute("aria-label"));
      panel.classList.add("is-collapsible");

      let heading = panel.querySelector(":scope > h2");
      if (!heading) {
        heading = document.createElement("h2");
        heading.textContent = sectionLabel || "セクション";
        panel.prepend(heading);
      }

      const header = document.createElement("div");
      header.className = "panel-header";
      panel.insertBefore(header, heading);
      header.appendChild(heading);

      const toggleBtn = document.createElement("button");
      toggleBtn.type = "button";
      toggleBtn.className = "panel-toggle-btn";
      header.appendChild(toggleBtn);

      const panelBody = document.createElement("div");
      panelBody.className = "panel-body";
      while (header.nextSibling) {
        panelBody.appendChild(header.nextSibling);
      }
      panel.appendChild(panelBody);

      const applyPanelState = (isCollapsed) => {
        panel.classList.toggle("panel-collapsed", isCollapsed);
        panelBody.hidden = isCollapsed;
        toggleBtn.textContent = isCollapsed ? "開く" : "閉じる";
        toggleBtn.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
      };

      const shouldStartCollapsed = sectionLabel !== "カレンダー" && sectionLabel !== "記録追加";
      applyPanelState(shouldStartCollapsed);
      toggleBtn.addEventListener("click", () => {
        applyPanelState(!panel.classList.contains("panel-collapsed"));
      });
    });
  }

  function onSubmitRecord(event) {
    event.preventDefault();

    const giver = sanitizeText(giverInput.value);
    const receiver = sanitizeText(receiverInput.value);
    const amount = Number.parseInt(amountInput.value, 10);
    const givenAtInputValue = givenAtInput.value;
    const memo = sanitizeText(memoInput.value);

    if (!giver || !receiver) {
      showMessage("渡した人・受け取る人を入力してください。", true);
      return;
    }
    if (!Number.isInteger(amount) || amount <= 0) {
      showMessage("金額は1円以上の整数で入力してください。", true);
      return;
    }

    const givenAtDate = new Date(givenAtInputValue);
    if (!givenAtInputValue || Number.isNaN(givenAtDate.getTime())) {
      showMessage("渡した日時を正しく入力してください。", true);
      return;
    }

    const previewGivenAt = givenAtDate.toISOString();
    const confirmationText = [
      "以上の内容で登録します。よろしいですか？",
      "",
      `渡した人: ${giver}`,
      `受け取る人: ${receiver}`,
      `金額: ${formatCurrency(amount)}`,
      `日時: ${formatDisplayDate(previewGivenAt)}`,
      `メモ: ${memo || "なし"}`
    ].join("\n");
    const ok = window.confirm(confirmationText);
    if (!ok) {
      showMessage("登録をキャンセルしました。入力内容は保持されています。", true);
      return;
    }

    const nowIso = new Date().toISOString();
    const newRecord = {
      id: createId(),
      giver,
      receiver,
      amount,
      givenAt: givenAtDate.toISOString(),
      memo,
      createdAt: nowIso,
      received: false,
      receivedAt: null,
      locked: false
    };

    recentAddedRecordId = newRecord.id;
    records.unshift(newRecord);
    persistAndRender();
    showMessage("記録を登録しました。一覧の先頭に追加されています。", false);

    form.reset();
    givenAtInput.value = toDatetimeLocalValue(new Date());
    giverInput.focus();
  }

  function onSubmitPerson(event) {
    event.preventDefault();

    const name = sanitizeText(personNameInput.value);
    const role = personRoleSelect.value;
    if (!name) {
      showMessage("登録する名前を入力してください。", true);
      return;
    }

    upsertPerson(name, role);
    savePeopleToStorage();
    renderPeopleManagement();
    showMessage("人を登録しました。", false);
    personForm.reset();
    personNameInput.focus();
  }

  function onClickPersonChipDelete(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (!target.matches("button[data-name][data-role]")) return;

    const name = target.dataset.name || "";
    const role = target.dataset.role || "";
    if (!name || !role) return;

    const ok = window.confirm(`「${name}」を${role === "giver" ? "渡す人候補" : "受け取る人候補"}から削除しますか？`);
    if (!ok) return;

    removePersonRole(name, role);
    savePeopleToStorage();
    renderPeopleManagement();
    showMessage("登録から削除しました。", false);
  }

  function onSubmitMemoTemplate(event) {
    event.preventDefault();

    const text = sanitizeText(memoTemplateInput.value);
    if (!text) {
      showMessage("定型文を入力してください。", true);
      return;
    }
    if (memoTemplates.includes(text)) {
      showMessage("同じ定型文がすでに登録されています。", true);
      return;
    }

    memoTemplates.push(text);
    memoTemplates.sort((a, b) => a.localeCompare(b, "ja"));
    saveMemoTemplatesToStorage();
    renderMemoTemplateManagement();
    showMessage("メモ定型文を登録しました。", false);
    memoTemplateForm.reset();
    memoTemplateInput.focus();
  }

  function onClickMemoTemplateDelete(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (!target.matches("button[data-memo-template]")) return;

    const template = target.dataset.memoTemplate || "";
    if (!template) return;

    const ok = window.confirm("このメモ定型文を削除しますか？");
    if (!ok) return;

    memoTemplates = memoTemplates.filter((item) => item !== template);
    saveMemoTemplatesToStorage();
    renderMemoTemplateManagement();
    showMessage("メモ定型文を削除しました。", false);
  }

  function onMemoTemplateSelected(event) {
    const value = event.target.value;
    if (!value) return;
    memoInput.value = value;
  }

  function onMonthlyTargetMonthChanged(event) {
    const value = toSafeString(event.target.value);
    monthlyTargetMonthKey = /^\d{4}-\d{2}$/.test(value) ? value : toMonthKey(new Date());
    monthlyTargetMonthInput.value = monthlyTargetMonthKey;
    renderMonthlySummary();
  }

  function onMonthlyReceiverFilterChanged(event) {
    monthlyReceiverFilterValue = event.target.value;
    renderMonthlySummary();
  }

  function onReceiverColorTargetChanged(event) {
    receiverColorTargetValue = event.target.value;
    syncReceiverColorPicker();
  }

  function onSubmitReceiverColor(event) {
    event.preventDefault();
    if (!receiverColorTargetValue) {
      showMessage("受取人を選択してください。", true);
      return;
    }

    const color = normalizeHexColor(receiverColorPicker.value);
    if (!color) {
      showMessage("色コードが不正です。", true);
      return;
    }

    receiverColors[receiverColorTargetValue] = color;
    saveReceiverColorsToStorage();
    renderAll();
    showMessage("受取人の色を保存しました。", false);
  }

  function onResetReceiverColor() {
    if (!receiverColorTargetValue) {
      showMessage("受取人を選択してください。", true);
      return;
    }
    if (!receiverColors[receiverColorTargetValue]) {
      showMessage("この受取人はすでに自動色です。", true);
      return;
    }

    const ok = window.confirm("この受取人の色設定を削除し、自動色に戻しますか？");
    if (!ok) return;

    delete receiverColors[receiverColorTargetValue];
    saveReceiverColorsToStorage();
    renderAll();
    showMessage("自動色に戻しました。", false);
  }

  function onClickReceiverColorList(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (!target.matches("button[data-receiver-color-reset]")) return;

    const receiverName = target.dataset.receiverColorReset || "";
    if (!receiverName) return;

    const ok = window.confirm(`「${receiverName}」の色設定を削除して自動色に戻しますか？`);
    if (!ok) return;

    delete receiverColors[receiverName];
    saveReceiverColorsToStorage();
    renderAll();
    showMessage("自動色に戻しました。", false);
  }

  function onReceiverPinTargetChanged(event) {
    receiverPinTargetValue = event.target.value;
    syncReceiverPinFormState();
  }

  async function onSubmitReceiverPin(event) {
    event.preventDefault();
    if (!receiverPinTargetValue) {
      showMessage("受取人を選択してください。", true);
      return;
    }

    const pin = toSafeString(receiverPinInput.value);
    if (!isValidPin(pin)) {
      showMessage("PINは4桁の数字で入力してください。", true);
      return;
    }

    receiverPins[receiverPinTargetValue] = await createReceiverPinEntry(pin);
    saveReceiverPinsToStorage();
    renderReceiverPinManagement();
    receiverPinInput.value = "";
    showMessage("PINを保存しました。", false);
  }

  function onDeleteReceiverPin() {
    if (!receiverPinTargetValue) {
      showMessage("受取人を選択してください。", true);
      return;
    }
    if (!receiverPins[receiverPinTargetValue]) {
      showMessage("この受取人にはPINが設定されていません。", true);
      return;
    }

    const ok = window.confirm(`「${receiverPinTargetValue}」のPIN設定を削除しますか？`);
    if (!ok) return;

    delete receiverPins[receiverPinTargetValue];
    saveReceiverPinsToStorage();
    renderReceiverPinManagement();
    showMessage("PIN設定を削除しました。", false);
  }

  function onClickReceiverPinList(event) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    if (!target.matches("button[data-receiver-pin-delete]")) return;

    const receiverName = target.dataset.receiverPinDelete || "";
    if (!receiverName || !receiverPins[receiverName]) return;

    const ok = window.confirm(`「${receiverName}」のPIN設定を削除しますか？`);
    if (!ok) return;

    delete receiverPins[receiverName];
    saveReceiverPinsToStorage();
    renderReceiverPinManagement();
    showMessage("PIN設定を削除しました。", false);
  }

  function onDeleteAll() {
    if (records.length === 0) {
      showMessage("削除する記録がありません。", true);
      return;
    }

    const ok = window.confirm("全件削除します。元に戻せません。よろしいですか？");
    if (!ok) return;

    records = [];
    persistAndRender();
    showMessage("全件削除しました。", false);
  }

  async function onConfirmReceived(recordId) {
    const target = records.find((item) => item.id === recordId);
    if (!target || target.locked || target.received) {
      return;
    }

    if (!receiverPins[target.receiver]) {
      showMessage("この受取人にはPINが設定されていません。先に設定してください。", true);
      return;
    }

    const verified = await requestPinVerification(target.receiver);
    if (!verified) return;

    target.received = true;
    target.locked = true;
    target.receivedAt = new Date().toISOString();

    persistAndRender();
    showMessage("受け取り確認を確定しました。", false);
  }

  function onDeleteOne(recordId) {
    const target = records.find((item) => item.id === recordId);
    if (!target) return;

    const ok = window.confirm("この記録を削除します。よろしいですか？");
    if (!ok) return;

    records = records.filter((item) => item.id !== recordId);
    persistAndRender();
    showMessage("記録を削除しました。", false);
  }

  function exportBackupJson() {
    try {
      const payload = {
        app: "allowance-tracker",
        version: APP_EXPORT_VERSION,
        exportedAt: new Date().toISOString(),
        records,
        people,
        memoTemplates,
        receiverColors,
        receiverPins
      };
      const jsonText = JSON.stringify(payload, null, 2);
      const fileName = `allowance-backup-${formatStampForFileName(new Date())}.json`;
      downloadTextFile(fileName, jsonText, "application/json;charset=utf-8");
      showMessage("JSONバックアップを書き出しました。", false);
    } catch (error) {
      showMessage("JSONバックアップの書き出しに失敗しました。", true);
    }
  }

  function exportCsv() {
    if (records.length === 0) {
      showMessage("書き出す記録がありません。", true);
      return;
    }

    const headers = [
      "ID",
      "渡した人",
      "受け取る人",
      "金額(円)",
      "渡した日時",
      "メモ",
      "作成日時",
      "受取状態",
      "受取確定日時",
      "ロック状態"
    ];

    const rows = records.map((item) => [
      item.id,
      item.giver,
      item.receiver,
      String(item.amount),
      formatDisplayDate(item.givenAt),
      item.memo || "",
      formatDisplayDate(item.createdAt),
      item.received ? "受取確定" : "未確認",
      item.receivedAt ? formatDisplayDate(item.receivedAt) : "",
      item.locked ? "true" : "false"
    ]);

    const csvBody = [headers, ...rows].map((row) => row.map(escapeCsvCell).join(",")).join("\r\n");

    // Excelでの文字化け対策としてUTF-8 BOMを付与
    const bom = "\uFEFF";
    const fileName = `allowance-records-${formatStampForFileName(new Date())}.csv`;
    downloadTextFile(fileName, bom + csvBody, "text/csv;charset=utf-8");
    showMessage("CSVを書き出しました。", false);
  }

  function onRestoreFileSelected(event) {
    const file = event.target.files && event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = () => {
      try {
        const text = String(reader.result || "");
        const parsed = JSON.parse(text);
        const restoredRecords = extractAndValidateBackup(parsed);
        const restoredPeople = extractAndValidatePeople(parsed, restoredRecords);
        const restoredMemoTemplates = extractAndValidateMemoTemplates(parsed);
        const restoredReceiverColors = extractAndValidateReceiverColors(parsed);
        const restoredReceiverPins = extractAndValidateReceiverPins(parsed);

        const ok = window.confirm(`バックアップから${restoredRecords.length}件を復元します。現在のデータは上書きされます。よろしいですか？`);
        if (!ok) {
          restoreFileInput.value = "";
          return;
        }

        records = restoredRecords;
        people = restoredPeople;
        memoTemplates = restoredMemoTemplates;
        receiverColors = restoredReceiverColors;
        receiverPins = restoredReceiverPins;
        persistAndRender();
        showMessage("復元が完了しました。", false);
      } catch (error) {
        showMessage("復元に失敗しました。JSON形式またはデータ内容を確認してください。", true);
      } finally {
        // 同じファイルを続けて選択できるように初期化
        restoreFileInput.value = "";
      }
    };

    reader.onerror = () => {
      showMessage("ファイルの読み込みに失敗しました。", true);
      restoreFileInput.value = "";
    };

    reader.readAsText(file, "utf-8");
  }

  function extractAndValidateBackup(payload) {
    // 互換性のため、配列直渡し形式とオブジェクト形式の両方を許容
    const candidateRecords = Array.isArray(payload) ? payload : payload && payload.records;
    if (!Array.isArray(candidateRecords)) {
      throw new Error("Invalid backup structure");
    }

    return candidateRecords.map(validateAndNormalizeRecord);
  }

  function extractAndValidatePeople(payload, fallbackRecords) {
    const rawPeople = payload && Array.isArray(payload.people) ? payload.people : null;
    if (!rawPeople) {
      return derivePeopleFromRecords(fallbackRecords);
    }

    return rawPeople.map(validateAndNormalizePerson);
  }

  function extractAndValidateMemoTemplates(payload) {
    const rawTemplates = payload && Array.isArray(payload.memoTemplates) ? payload.memoTemplates : null;
    if (!rawTemplates) {
      return [];
    }
    return rawTemplates.map(validateAndNormalizeMemoTemplate);
  }

  function extractAndValidateReceiverColors(payload) {
    const raw = payload && payload.receiverColors;
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
      return {};
    }

    /** @type {Record<string, string>} */
    const normalized = {};
    Object.entries(raw).forEach(([name, color]) => {
      const safeName = toSafeString(name);
      const safeColor = normalizeHexColor(color);
      if (!safeName || !safeColor) return;
      normalized[safeName] = safeColor;
    });
    return normalized;
  }

  function extractAndValidateReceiverPins(payload) {
    const raw = payload && payload.receiverPins;
    if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
      return {};
    }

    /** @type {Record<string, {salt:string,hash:string,updatedAt:string}>} */
    const normalized = {};
    Object.entries(raw).forEach(([name, entry]) => {
      const safeName = toSafeString(name);
      if (!safeName) return;
      try {
        normalized[safeName] = validateAndNormalizeReceiverPinEntry(entry);
      } catch (error) {
        // 不正なPIN設定エントリはスキップ
      }
    });
    return normalized;
  }

  function validateAndNormalizeRecord(raw) {
    if (!raw || typeof raw !== "object") {
      throw new Error("Record is not object");
    }

    const id = toSafeString(raw.id);
    const giver = toSafeString(raw.giver);
    const receiver = toSafeString(raw.receiver);
    const amount = Number.parseInt(raw.amount, 10);
    const givenAt = normalizeIsoDate(raw.givenAt);
    const memo = toSafeString(raw.memo);
    const createdAt = normalizeIsoDate(raw.createdAt || raw.givenAt);
    const received = Boolean(raw.received);
    const locked = Boolean(raw.locked);
    const receivedAt = raw.receivedAt ? normalizeIsoDate(raw.receivedAt) : null;

    if (!id || !giver || !receiver) {
      throw new Error("Missing required fields");
    }
    if (!Number.isInteger(amount) || amount <= 0) {
      throw new Error("Invalid amount");
    }

    return {
      id,
      giver,
      receiver,
      amount,
      givenAt,
      memo,
      createdAt,
      received,
      receivedAt,
      locked: locked || received
    };
  }

  function validateAndNormalizePerson(raw) {
    if (!raw || typeof raw !== "object") {
      throw new Error("Person is not object");
    }

    const name = toSafeString(raw.name);
    const canGive = Boolean(raw.canGive);
    const canReceive = Boolean(raw.canReceive);
    if (!name) {
      throw new Error("Person name is required");
    }
    if (!canGive && !canReceive) {
      throw new Error("Person role is required");
    }

    return { name, canGive, canReceive };
  }

  function validateAndNormalizeMemoTemplate(raw) {
    const text = toSafeString(raw);
    if (!text) {
      throw new Error("Memo template is required");
    }
    return text.slice(0, 240);
  }

  function validateAndNormalizeReceiverPinEntry(raw) {
    if (!raw || typeof raw !== "object") {
      throw new Error("PIN entry is not object");
    }
    const salt = toSafeString(raw.salt);
    const hash = toSafeString(raw.hash);
    const updatedAt = raw.updatedAt ? normalizeIsoDate(raw.updatedAt) : new Date().toISOString();
    if (!salt || !hash) {
      throw new Error("PIN entry is invalid");
    }
    return { salt, hash, updatedAt };
  }

  function loadRecordsFromStorage() {
    try {
      const raw = window.localStorage.getItem(STORAGE_KEY);
      if (!raw) {
        records = [];
        return;
      }
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        records = [];
        return;
      }
      records = parsed.map(validateAndNormalizeRecord);
    } catch (error) {
      records = [];
      showMessage("保存データの読み込みで問題が発生したため、初期化しました。", true);
    }
  }

  function loadPeopleFromStorage() {
    try {
      const raw = window.localStorage.getItem(PEOPLE_STORAGE_KEY);
      if (!raw) {
        people = [];
        return;
      }
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        people = [];
        return;
      }
      people = parsed.map(validateAndNormalizePerson);
    } catch (error) {
      people = [];
      showMessage("人の登録データの読み込みで問題が発生したため、初期化しました。", true);
    }
  }

  function loadMemoTemplatesFromStorage() {
    try {
      const raw = window.localStorage.getItem(MEMO_TEMPLATES_STORAGE_KEY);
      if (!raw) {
        memoTemplates = [];
        return;
      }
      const parsed = JSON.parse(raw);
      if (!Array.isArray(parsed)) {
        memoTemplates = [];
        return;
      }
      memoTemplates = parsed.map(validateAndNormalizeMemoTemplate);
    } catch (error) {
      memoTemplates = [];
      showMessage("メモ定型文の読み込みで問題が発生したため、初期化しました。", true);
    }
  }

  function loadReceiverColorsFromStorage() {
    try {
      const raw = window.localStorage.getItem(RECEIVER_COLORS_STORAGE_KEY);
      if (!raw) {
        receiverColors = {};
        return;
      }
      const parsed = JSON.parse(raw);
      receiverColors = extractAndValidateReceiverColors({ receiverColors: parsed });
    } catch (error) {
      receiverColors = {};
      showMessage("受取人色設定の読み込みで問題が発生したため、初期化しました。", true);
    }
  }

  function loadReceiverPinsFromStorage() {
    try {
      const raw = window.localStorage.getItem(RECEIVER_PINS_STORAGE_KEY);
      if (!raw) {
        receiverPins = {};
        return;
      }
      const parsed = JSON.parse(raw);
      receiverPins = extractAndValidateReceiverPins({ receiverPins: parsed });
    } catch (error) {
      receiverPins = {};
      showMessage("受取人PIN設定の読み込みで問題が発生したため、初期化しました。", true);
    }
  }

  function persistAndRender() {
    saveRecordsToStorage();
    savePeopleToStorage();
    saveMemoTemplatesToStorage();
    saveReceiverColorsToStorage();
    saveReceiverPinsToStorage();
    renderAll();
  }

  function saveRecordsToStorage() {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify(records));
  }

  function savePeopleToStorage() {
    window.localStorage.setItem(PEOPLE_STORAGE_KEY, JSON.stringify(people));
  }

  function saveMemoTemplatesToStorage() {
    window.localStorage.setItem(MEMO_TEMPLATES_STORAGE_KEY, JSON.stringify(memoTemplates));
  }

  function saveReceiverColorsToStorage() {
    window.localStorage.setItem(RECEIVER_COLORS_STORAGE_KEY, JSON.stringify(receiverColors));
  }

  function saveReceiverPinsToStorage() {
    window.localStorage.setItem(RECEIVER_PINS_STORAGE_KEY, JSON.stringify(receiverPins));
  }

  function renderAll() {
    renderStats();
    renderMonthlyReceiverFilterOptions();
    renderMonthlySummary();
    renderReceiverColorManagement();
    renderReceiverPinManagement();
    renderRecords();
    renderPeopleManagement();
    renderMemoTemplateManagement();
    renderReceiverView();
    renderCalendarView();
  }

  function renderPeopleManagement() {
    renderPersonCandidateDatalists();
    renderPeopleLists();
  }

  function renderMemoTemplateManagement() {
    renderMemoTemplateSelect();
    renderMemoTemplateList();
  }

  function renderReceiverColorManagement() {
    const receiverNames = getSortedReceiverNames();
    renderReceiverColorTargetOptions(receiverNames);
    renderReceiverColorList(receiverNames);
    syncReceiverColorPicker();
  }

  function renderReceiverPinManagement() {
    const receiverNames = getManagedReceiverNames();
    renderReceiverPinTargetOptions(receiverNames);
    renderReceiverPinList(receiverNames);
    syncReceiverPinFormState();
  }

  function renderReceiverColorTargetOptions(receiverNames) {
    const previous = receiverColorTargetValue;
    receiverColorTargetSelect.innerHTML = "";

    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = receiverNames.length > 0 ? "受取人を選択" : "受取人がいません";
    receiverColorTargetSelect.appendChild(placeholder);

    receiverNames.forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      receiverColorTargetSelect.appendChild(option);
    });

    const stillExists = receiverNames.includes(previous);
    receiverColorTargetValue = stillExists ? previous : (receiverNames[0] || "");
    receiverColorTargetSelect.value = receiverColorTargetValue;
    receiverColorTargetSelect.disabled = receiverNames.length === 0;
    receiverColorPicker.disabled = receiverNames.length === 0;
    resetReceiverColorBtn.disabled = receiverNames.length === 0;
  }

  function renderReceiverColorList(receiverNames) {
    receiverColorList.innerHTML = "";
    if (receiverNames.length === 0) {
      const emptyEl = document.createElement("p");
      emptyEl.className = "empty";
      emptyEl.textContent = "受取人の記録がありません。";
      receiverColorList.appendChild(emptyEl);
      return;
    }

    const fragment = document.createDocumentFragment();
    receiverNames.forEach((name) => {
      const item = document.createElement("div");
      item.className = "receiver-color-item";

      const main = document.createElement("div");
      main.className = "receiver-color-main";

      const swatch = document.createElement("span");
      swatch.className = "receiver-color-swatch";
      swatch.style.backgroundColor = getReceiverColor(name).accent;

      const nameEl = document.createElement("span");
      nameEl.className = "receiver-color-name";
      nameEl.textContent = name;

      const modeEl = document.createElement("span");
      modeEl.className = "receiver-color-mode";
      modeEl.textContent = receiverColors[name] ? "手動色" : "自動色";

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn danger";
      btn.textContent = "自動色";
      btn.dataset.receiverColorReset = name;
      btn.disabled = !receiverColors[name];

      main.appendChild(swatch);
      main.appendChild(nameEl);
      main.appendChild(modeEl);
      item.appendChild(main);
      item.appendChild(btn);
      fragment.appendChild(item);
    });

    receiverColorList.appendChild(fragment);
  }

  function syncReceiverColorPicker() {
    if (!receiverColorTargetValue) {
      receiverColorPicker.value = "#4bbcf4";
      resetReceiverColorBtn.disabled = true;
      return;
    }
    receiverColorPicker.value = getReceiverColor(receiverColorTargetValue).accent;
    resetReceiverColorBtn.disabled = !Boolean(receiverColors[receiverColorTargetValue]);
  }

  function renderReceiverPinTargetOptions(receiverNames) {
    const previous = receiverPinTargetValue;
    receiverPinTargetSelect.innerHTML = "";

    const placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = receiverNames.length > 0 ? "受取人を選択" : "受取人がいません";
    receiverPinTargetSelect.appendChild(placeholder);

    receiverNames.forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      receiverPinTargetSelect.appendChild(option);
    });

    receiverPinTargetValue = receiverNames.includes(previous) ? previous : (receiverNames[0] || "");
    receiverPinTargetSelect.value = receiverPinTargetValue;
  }

  function renderReceiverPinList(receiverNames) {
    receiverPinList.innerHTML = "";
    if (receiverNames.length === 0) {
      const emptyEl = document.createElement("p");
      emptyEl.className = "empty";
      emptyEl.textContent = "受取人がいません。";
      receiverPinList.appendChild(emptyEl);
      return;
    }

    const fragment = document.createDocumentFragment();
    receiverNames.forEach((name) => {
      const item = document.createElement("div");
      item.className = "receiver-pin-item";

      const main = document.createElement("div");
      main.className = "receiver-pin-main";

      const nameEl = document.createElement("span");
      nameEl.className = "receiver-pin-name";
      nameEl.textContent = name;

      const modeEl = document.createElement("span");
      modeEl.className = "receiver-pin-mode";
      modeEl.textContent = receiverPins[name] ? "PIN設定済み" : "PIN未設定";

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "btn danger";
      btn.textContent = "PIN削除";
      btn.dataset.receiverPinDelete = name;
      btn.disabled = !receiverPins[name];

      main.appendChild(nameEl);
      main.appendChild(modeEl);
      item.appendChild(main);
      item.appendChild(btn);
      fragment.appendChild(item);
    });

    receiverPinList.appendChild(fragment);
  }

  function syncReceiverPinFormState() {
    const hasTarget = Boolean(receiverPinTargetValue);
    receiverPinTargetSelect.disabled = receiverPinTargetSelect.options.length <= 1;
    receiverPinInput.disabled = !hasTarget;
    deleteReceiverPinBtn.disabled = !hasTarget || !Boolean(receiverPins[receiverPinTargetValue]);
  }

  function renderMemoTemplateSelect() {
    const selectedBefore = memoTemplateSelect.value;
    memoTemplateSelect.innerHTML = "";

    const defaultOption = document.createElement("option");
    defaultOption.value = "";
    defaultOption.textContent = "定型メモを選択（任意）";
    memoTemplateSelect.appendChild(defaultOption);

    memoTemplates.forEach((template) => {
      const option = document.createElement("option");
      option.value = template;
      option.textContent = template;
      memoTemplateSelect.appendChild(option);
    });

    memoTemplateSelect.value = memoTemplates.includes(selectedBefore) ? selectedBefore : "";
  }

  function renderMemoTemplateList() {
    memoTemplateList.innerHTML = "";

    if (memoTemplates.length === 0) {
      const emptyEl = document.createElement("span");
      emptyEl.className = "empty";
      emptyEl.textContent = "登録なし";
      memoTemplateList.appendChild(emptyEl);
      return;
    }

    memoTemplates.forEach((template) => {
      const chip = document.createElement("span");
      chip.className = "person-chip";

      const textEl = document.createElement("span");
      textEl.textContent = template;

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.textContent = "×";
      removeBtn.dataset.memoTemplate = template;
      removeBtn.setAttribute("aria-label", "メモ定型文を削除");

      chip.appendChild(textEl);
      chip.appendChild(removeBtn);
      memoTemplateList.appendChild(chip);
    });
  }

  function renderPersonCandidateDatalists() {
    giverCandidates.innerHTML = "";
    receiverCandidates.innerHTML = "";

    people
      .filter((item) => item.canGive)
      .sort((a, b) => a.name.localeCompare(b.name, "ja"))
      .forEach((item) => {
        const option = document.createElement("option");
        option.value = item.name;
        giverCandidates.appendChild(option);
      });

    people
      .filter((item) => item.canReceive)
      .sort((a, b) => a.name.localeCompare(b.name, "ja"))
      .forEach((item) => {
        const option = document.createElement("option");
        option.value = item.name;
        receiverCandidates.appendChild(option);
      });
  }

  function renderPeopleLists() {
    giverPeopleList.innerHTML = "";
    receiverPeopleList.innerHTML = "";

    const giverPeople = people
      .filter((item) => item.canGive)
      .sort((a, b) => a.name.localeCompare(b.name, "ja"));
    const receiverPeople = people
      .filter((item) => item.canReceive)
      .sort((a, b) => a.name.localeCompare(b.name, "ja"));

    if (giverPeople.length === 0) {
      const emptyEl = document.createElement("span");
      emptyEl.className = "empty";
      emptyEl.textContent = "登録なし";
      giverPeopleList.appendChild(emptyEl);
    } else {
      giverPeople.forEach((item) => giverPeopleList.appendChild(createPersonChip(item.name, "giver")));
    }

    if (receiverPeople.length === 0) {
      const emptyEl = document.createElement("span");
      emptyEl.className = "empty";
      emptyEl.textContent = "登録なし";
      receiverPeopleList.appendChild(emptyEl);
    } else {
      receiverPeople.forEach((item) => receiverPeopleList.appendChild(createPersonChip(item.name, "receiver")));
    }
  }

  function createPersonChip(name, role) {
    const chip = document.createElement("span");
    chip.className = "person-chip";

    const textEl = document.createElement("span");
    textEl.textContent = name;

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.textContent = "×";
    removeBtn.dataset.name = name;
    removeBtn.dataset.role = role;
    removeBtn.setAttribute("aria-label", `${name}を${role === "giver" ? "渡す人候補" : "受け取る人候補"}から削除`);

    chip.appendChild(textEl);
    chip.appendChild(removeBtn);
    return chip;
  }

  function renderStats() {
    const unreceivedCount = records.filter((item) => !item.received).length;
    const receivedCount = records.filter((item) => item.received).length;
    const totalAmount = records.reduce((sum, item) => sum + item.amount, 0);

    unreceivedCountEl.textContent = `${unreceivedCount}件`;
    receivedCountEl.textContent = `${receivedCount}件`;
    totalAmountEl.textContent = formatCurrency(totalAmount);
  }

  function renderMonthlySummary() {
    if (!/^\d{4}-\d{2}$/.test(monthlyTargetMonthKey)) {
      monthlyTargetMonthKey = toMonthKey(new Date());
      monthlyTargetMonthInput.value = monthlyTargetMonthKey;
    }

    const monthRecords = records.filter((item) => toMonthKey(new Date(item.givenAt)) === monthlyTargetMonthKey);
    const filteredMonthRecords = monthlyReceiverFilterValue === "ALL"
      ? monthRecords
      : monthRecords.filter((item) => item.receiver === monthlyReceiverFilterValue);
    const monthlyTotalAmount = filteredMonthRecords.reduce((sum, item) => sum + item.amount, 0);

    monthlyTotalAmountEl.textContent = formatCurrency(monthlyTotalAmount);
    monthlyTotalCountEl.textContent = `${filteredMonthRecords.length}件`;
  }

  function renderMonthlyReceiverFilterOptions() {
    const receiverNames = getSortedReceiverNames();
    const previousValue = monthlyReceiverFilterValue;

    monthlyReceiverFilterSelect.innerHTML = "";

    const allOption = document.createElement("option");
    allOption.value = "ALL";
    allOption.textContent = "全員";
    monthlyReceiverFilterSelect.appendChild(allOption);

    receiverNames.forEach((name) => {
      const option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      monthlyReceiverFilterSelect.appendChild(option);
    });

    const stillExists = previousValue === "ALL" || receiverNames.includes(previousValue);
    monthlyReceiverFilterValue = stillExists ? previousValue : "ALL";
    monthlyReceiverFilterSelect.value = monthlyReceiverFilterValue;
  }

  function renderRecords() {
    recordsContainer.innerHTML = "";

    if (records.length === 0) {
      const emptyEl = document.createElement("p");
      emptyEl.className = "empty";
      emptyEl.textContent = "記録はまだありません。上のフォームから追加してください。";
      recordsContainer.appendChild(emptyEl);
      return;
    }

    const fragment = document.createDocumentFragment();

    records.forEach((item) => {
      const node = cardTemplate.content.firstElementChild.cloneNode(true);
      node.dataset.recordId = item.id;
      const amountEl = node.querySelector(".amount");
      const statusBadgeEl = node.querySelector(".status-badge");
      const peopleEl = node.querySelector(".people");
      const datetimeEl = node.querySelector(".datetime");
      const memoEl = node.querySelector(".memo");
      const receivedInfoEl = node.querySelector(".received-info");
      const confirmBtn = node.querySelector(".confirm-btn");
      const deleteBtn = node.querySelector(".delete-btn");
      applyReceiverColorTheme(node, item.receiver);

      amountEl.textContent = formatCurrency(item.amount);

      if (item.received) {
        statusBadgeEl.textContent = "受取確定";
        statusBadgeEl.classList.add("status-done");
        receivedInfoEl.textContent = `受取者確認済み: ${formatDisplayDate(item.receivedAt)}`;
        confirmBtn.disabled = true;
        confirmBtn.textContent = "受取確定済み";
      } else {
        statusBadgeEl.textContent = "未確認";
        statusBadgeEl.classList.add("status-pending");
        receivedInfoEl.textContent = "受取者確認待ち";
      }

      peopleEl.textContent = `渡した人: ${item.giver} / 受け取る人: ${item.receiver}`;
      datetimeEl.textContent = `渡した日時: ${formatDisplayDate(item.givenAt)}`;
      memoEl.textContent = item.memo ? `メモ: ${item.memo}` : "メモ: なし";

      confirmBtn.addEventListener("click", () => onConfirmReceived(item.id));
      deleteBtn.addEventListener("click", () => onDeleteOne(item.id));

      fragment.appendChild(node);
    });

    recordsContainer.appendChild(fragment);

    if (recentAddedRecordId) {
      const addedCard = recordsContainer.querySelector(`[data-record-id="${recentAddedRecordId}"]`);
      if (addedCard) {
        addedCard.classList.add("new-record-flash");
        window.setTimeout(() => {
          addedCard.classList.remove("new-record-flash");
        }, 1800);
      }
      recentAddedRecordId = "";
    }
  }

  function onReceiverFilterChanged(event) {
    receiverFilterValue = event.target.value;
    renderReceiverRecords();
  }

  function renderReceiverView() {
    renderReceiverFilterOptions();
    renderReceiverSummary();
    renderReceiverRecords();
  }

  function renderReceiverFilterOptions() {
    const receiverNames = getSortedReceiverNames();
    const previousValue = receiverFilterValue;

    receiverFilterSelect.innerHTML = "";

    const allOption = document.createElement("option");
    allOption.value = "ALL";
    allOption.textContent = `全員（${records.length}件）`;
    receiverFilterSelect.appendChild(allOption);

    receiverNames.forEach((name) => {
      const receiverRecords = records.filter((item) => item.receiver === name);
      const option = document.createElement("option");
      option.value = name;
      option.textContent = `${name}（${receiverRecords.length}件）`;
      receiverFilterSelect.appendChild(option);
    });

    const stillExists = previousValue === "ALL" || receiverNames.includes(previousValue);
    receiverFilterValue = stillExists ? previousValue : "ALL";
    receiverFilterSelect.value = receiverFilterValue;
  }

  function renderReceiverSummary() {
    receiverSummaryEl.innerHTML = "";

    const summaryMap = buildReceiverSummaryMap();
    if (summaryMap.size === 0) {
      const emptyEl = document.createElement("p");
      emptyEl.className = "empty";
      emptyEl.textContent = "受け取る人のデータがありません。";
      receiverSummaryEl.appendChild(emptyEl);
      return;
    }

    const fragment = document.createDocumentFragment();
    const sortedEntries = [...summaryMap.entries()].sort((a, b) => a[0].localeCompare(b[0], "ja"));

    sortedEntries.forEach(([name, summary]) => {
      const card = document.createElement("article");
      card.className = "receiver-summary-card";
      applyReceiverColorTheme(card, name);

      const top = document.createElement("div");
      top.className = "receiver-summary-top";

      const nameEl = document.createElement("span");
      nameEl.className = "receiver-name";
      nameEl.textContent = name;

      const amountEl = document.createElement("strong");
      amountEl.textContent = formatCurrency(summary.totalAmount);

      top.appendChild(nameEl);
      top.appendChild(amountEl);

      const metaEl = document.createElement("p");
      metaEl.className = "receiver-meta";
      metaEl.textContent = `件数: ${summary.totalCount}件 / 未確認: ${summary.unreceivedCount}件 / 受取確定: ${summary.receivedCount}件`;

      card.appendChild(top);
      card.appendChild(metaEl);
      fragment.appendChild(card);
    });

    receiverSummaryEl.appendChild(fragment);
  }

  function renderReceiverRecords() {
    receiverRecordsContainer.innerHTML = "";

    const targetRecords = receiverFilterValue === "ALL"
      ? records
      : records.filter((item) => item.receiver === receiverFilterValue);

    if (targetRecords.length === 0) {
      const emptyEl = document.createElement("p");
      emptyEl.className = "empty";
      emptyEl.textContent = "該当する記録はありません。";
      receiverRecordsContainer.appendChild(emptyEl);
      return;
    }

    const fragment = document.createDocumentFragment();

    targetRecords.forEach((item) => {
      const card = document.createElement("article");
      card.className = "record-card";
      applyReceiverColorTheme(card, item.receiver);

      const top = document.createElement("div");
      top.className = "top-row";

      const amountEl = document.createElement("strong");
      amountEl.className = "amount";
      amountEl.textContent = formatCurrency(item.amount);

      const statusBadgeEl = document.createElement("span");
      statusBadgeEl.className = "status-badge";
      statusBadgeEl.textContent = item.received ? "受取確定" : "未確認";
      statusBadgeEl.classList.add(item.received ? "status-done" : "status-pending");

      top.appendChild(amountEl);
      top.appendChild(statusBadgeEl);

      const peopleEl = document.createElement("p");
      peopleEl.className = "people";
      peopleEl.textContent = `渡した人: ${item.giver} / 受け取る人: ${item.receiver}`;

      const dateEl = document.createElement("p");
      dateEl.className = "datetime";
      dateEl.textContent = `渡した日時: ${formatDisplayDate(item.givenAt)}`;

      const receivedInfoEl = document.createElement("p");
      receivedInfoEl.className = "received-info";
      receivedInfoEl.textContent = item.received
        ? `受取者確認済み: ${formatDisplayDate(item.receivedAt)}`
        : "受取者確認待ち";

      card.appendChild(top);
      card.appendChild(peopleEl);
      card.appendChild(dateEl);
      card.appendChild(receivedInfoEl);
      fragment.appendChild(card);
    });

    receiverRecordsContainer.appendChild(fragment);
  }

  function getSortedReceiverNames() {
    const names = new Set(records.map((item) => item.receiver));
    return [...names].sort((a, b) => a.localeCompare(b, "ja"));
  }

  function getManagedReceiverNames() {
    const names = new Set(getSortedReceiverNames());
    people.forEach((item) => {
      if (item.canReceive) names.add(item.name);
    });
    return [...names].sort((a, b) => a.localeCompare(b, "ja"));
  }

  function buildReceiverSummaryMap() {
    const map = new Map();

    records.forEach((item) => {
      if (!map.has(item.receiver)) {
        map.set(item.receiver, {
          totalCount: 0,
          unreceivedCount: 0,
          receivedCount: 0,
          totalAmount: 0
        });
      }

      const current = map.get(item.receiver);
      current.totalCount += 1;
      current.totalAmount += item.amount;
      if (item.received) {
        current.receivedCount += 1;
      } else {
        current.unreceivedCount += 1;
      }
    });

    return map;
  }

  function applyReceiverColorTheme(element, receiverName) {
    const { accent, soft } = getReceiverColor(receiverName);
    element.classList.add("receiver-colored");
    element.style.setProperty("--receiver-accent", accent);
    element.style.setProperty("--receiver-soft", soft);
  }

  function getReceiverColor(receiverName) {
    const safeName = toSafeString(receiverName);
    if (!safeName) return RECEIVER_COLOR_PALETTE[0];
    const custom = normalizeHexColor(receiverColors[safeName]);
    if (custom) {
      return { accent: custom, soft: makeSoftColor(custom, 0.86) };
    }
    const index = Math.abs(hashString(safeName)) % RECEIVER_COLOR_PALETTE.length;
    return RECEIVER_COLOR_PALETTE[index];
  }

  function normalizeHexColor(value) {
    const text = toSafeString(value).toLowerCase();
    return /^#[0-9a-f]{6}$/.test(text) ? text : "";
  }

  function makeSoftColor(hexColor, ratio) {
    const hex = normalizeHexColor(hexColor);
    if (!hex) return "#ffffff";
    const r = Number.parseInt(hex.slice(1, 3), 16);
    const g = Number.parseInt(hex.slice(3, 5), 16);
    const b = Number.parseInt(hex.slice(5, 7), 16);
    const mixedR = Math.round(r * (1 - ratio) + 255 * ratio);
    const mixedG = Math.round(g * (1 - ratio) + 255 * ratio);
    const mixedB = Math.round(b * (1 - ratio) + 255 * ratio);
    return `rgb(${mixedR}, ${mixedG}, ${mixedB})`;
  }

  function hashString(text) {
    let hash = 0;
    for (let i = 0; i < text.length; i += 1) {
      hash = (hash * 31 + text.charCodeAt(i)) | 0;
    }
    return hash;
  }

  function upsertPerson(name, role) {
    const index = people.findIndex((item) => item.name === name);
    if (index === -1) {
      people.push({
        name,
        canGive: role === "giver" || role === "both",
        canReceive: role === "receiver" || role === "both"
      });
      return;
    }

    const target = people[index];
    if (role === "giver" || role === "both") target.canGive = true;
    if (role === "receiver" || role === "both") target.canReceive = true;
  }

  function removePersonRole(name, role) {
    const target = people.find((item) => item.name === name);
    if (!target) return;

    if (role === "giver") target.canGive = false;
    if (role === "receiver") target.canReceive = false;

    people = people.filter((item) => item.canGive || item.canReceive);
  }

  function derivePeopleFromRecords(sourceRecords) {
    const map = new Map();
    sourceRecords.forEach((item) => {
      if (!map.has(item.giver)) {
        map.set(item.giver, { name: item.giver, canGive: true, canReceive: false });
      } else {
        map.get(item.giver).canGive = true;
      }

      if (!map.has(item.receiver)) {
        map.set(item.receiver, { name: item.receiver, canGive: false, canReceive: true });
      } else {
        map.get(item.receiver).canReceive = true;
      }
    });
    return [...map.values()];
  }

  function switchView(viewName) {
    currentView = viewName === "calendar" ? "calendar" : "list";
    const isCalendar = currentView === "calendar";

    listView.classList.toggle("hidden", isCalendar);
    calendarView.classList.toggle("hidden", !isCalendar);
    showListViewBtn.classList.toggle("active", !isCalendar);
    showCalendarViewBtn.classList.toggle("active", isCalendar);
  }

  function moveCalendarMonth(diffMonth) {
    calendarMonthCursor = new Date(
      calendarMonthCursor.getFullYear(),
      calendarMonthCursor.getMonth() + diffMonth,
      1
    );
    renderCalendarView();
  }

  function renderCalendarView() {
    renderCalendarHeader();
    renderCalendarGrid();
    renderSelectedDateRecords();
  }

  function renderCalendarHeader() {
    calendarMonthLabel.textContent = new Intl.DateTimeFormat("ja-JP", {
      year: "numeric",
      month: "long"
    }).format(calendarMonthCursor);
  }

  function renderCalendarGrid() {
    calendarGrid.innerHTML = "";

    const monthStart = startOfMonth(calendarMonthCursor);
    const gridStart = new Date(monthStart);
    gridStart.setDate(monthStart.getDate() - monthStart.getDay());

    const fragment = document.createDocumentFragment();

    for (let i = 0; i < 42; i += 1) {
      const currentDate = new Date(gridStart);
      currentDate.setDate(gridStart.getDate() + i);
      const dateKey = toDateKey(currentDate);
      const dayRecords = records.filter((item) => toDateKey(new Date(item.givenAt)) === dateKey);
      const dayTotal = dayRecords.reduce((sum, item) => sum + item.amount, 0);

      const dayButton = document.createElement("button");
      dayButton.type = "button";
      dayButton.className = "calendar-day";

      if (currentDate.getMonth() !== calendarMonthCursor.getMonth()) {
        dayButton.classList.add("other-month");
      }
      if (dateKey === selectedDateKey) {
        dayButton.classList.add("selected");
      }

      const dayHead = document.createElement("div");
      dayHead.className = "calendar-day-head";
      dayHead.textContent = String(currentDate.getDate());

      const dayInfo = document.createElement("div");
      dayInfo.className = "calendar-day-info";
      dayInfo.textContent = dayRecords.length > 0
        ? `${dayRecords.length}件 / ${formatCurrency(dayTotal)}`
        : "記録なし";

      dayButton.appendChild(dayHead);
      dayButton.appendChild(dayInfo);
      dayButton.addEventListener("click", () => onCalendarDaySelected(currentDate));
      fragment.appendChild(dayButton);
    }

    calendarGrid.appendChild(fragment);
  }

  function onCalendarDaySelected(date) {
    selectedDateKey = toDateKey(date);

    if (
      date.getFullYear() !== calendarMonthCursor.getFullYear() ||
      date.getMonth() !== calendarMonthCursor.getMonth()
    ) {
      calendarMonthCursor = startOfMonth(date);
    }

    renderCalendarView();
  }

  function renderSelectedDateRecords() {
    selectedDateLabel.textContent = `${formatDateKeyLabel(selectedDateKey)} の記録`;
    calendarRecordsContainer.innerHTML = "";

    const dayRecords = records
      .filter((item) => toDateKey(new Date(item.givenAt)) === selectedDateKey)
      .sort((a, b) => new Date(b.givenAt).getTime() - new Date(a.givenAt).getTime());

    if (dayRecords.length === 0) {
      const emptyEl = document.createElement("p");
      emptyEl.className = "empty";
      emptyEl.textContent = "この日の記録はありません。";
      calendarRecordsContainer.appendChild(emptyEl);
      return;
    }

    const fragment = document.createDocumentFragment();
    dayRecords.forEach((item) => {
      const card = document.createElement("article");
      card.className = "record-card";
      applyReceiverColorTheme(card, item.receiver);

      const top = document.createElement("div");
      top.className = "top-row";

      const amountEl = document.createElement("strong");
      amountEl.className = "amount";
      amountEl.textContent = formatCurrency(item.amount);

      const statusBadgeEl = document.createElement("span");
      statusBadgeEl.className = "status-badge";
      statusBadgeEl.classList.add(item.received ? "status-done" : "status-pending");
      statusBadgeEl.textContent = item.received ? "受取確定" : "未確認";

      top.appendChild(amountEl);
      top.appendChild(statusBadgeEl);

      const peopleEl = document.createElement("p");
      peopleEl.className = "people";
      peopleEl.textContent = `渡した人: ${item.giver} / 受け取る人: ${item.receiver}`;

      const timeEl = document.createElement("p");
      timeEl.className = "datetime";
      timeEl.textContent = `時刻: ${formatDisplayDate(item.givenAt)}`;

      const memoEl = document.createElement("p");
      memoEl.className = "memo";
      memoEl.textContent = item.memo ? `メモ: ${item.memo}` : "メモ: なし";

      card.appendChild(top);
      card.appendChild(peopleEl);
      card.appendChild(timeEl);
      card.appendChild(memoEl);
      fragment.appendChild(card);
    });

    calendarRecordsContainer.appendChild(fragment);
  }

  function downloadTextFile(fileName, text, mimeType) {
    const blob = new Blob([text], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  function sanitizeText(value) {
    return toSafeString(value).replace(/\s+/g, " ").trim();
  }

  function toSafeString(value) {
    if (value === null || value === undefined) return "";
    return String(value).trim();
  }

  function normalizeIsoDate(value) {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
      throw new Error("Invalid date");
    }
    return date.toISOString();
  }

  function startOfMonth(date) {
    return new Date(date.getFullYear(), date.getMonth(), 1);
  }

  function toDateKey(date) {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}`;
  }

  function toMonthKey(date) {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    return `${yyyy}-${mm}`;
  }

  function formatDateKeyLabel(dateKey) {
    const [y, m, d] = dateKey.split("-").map((v) => Number.parseInt(v, 10));
    if (!y || !m || !d) return dateKey;
    return `${y}年${m}月${d}日`;
  }

  function formatCurrency(amount) {
    return new Intl.NumberFormat("ja-JP", {
      style: "currency",
      currency: "JPY",
      maximumFractionDigits: 0
    }).format(amount);
  }

  function formatDisplayDate(isoString) {
    const date = new Date(isoString);
    if (Number.isNaN(date.getTime())) return "-";

    return new Intl.DateTimeFormat("ja-JP", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit"
    }).format(date);
  }

  function toDatetimeLocalValue(date) {
    const offset = date.getTimezoneOffset();
    const localDate = new Date(date.getTime() - offset * 60000);
    return localDate.toISOString().slice(0, 16);
  }

  function formatStampForFileName(date) {
    const yyyy = date.getFullYear();
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const dd = String(date.getDate()).padStart(2, "0");
    const hh = String(date.getHours()).padStart(2, "0");
    const mi = String(date.getMinutes()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd}-${hh}${mi}`;
  }

  function createId() {
    return `rec-${Date.now()}-${Math.random().toString(16).slice(2, 10)}`;
  }

  function escapeCsvCell(value) {
    const text = toSafeString(value);
    const escaped = text.replace(/"/g, '""');
    return `"${escaped}"`;
  }

  function isValidPin(pin) {
    return /^\d{4}$/.test(pin);
  }

  function createPinSalt() {
    const arr = new Uint8Array(16);
    if (window.crypto && typeof window.crypto.getRandomValues === "function") {
      window.crypto.getRandomValues(arr);
      return [...arr].map((v) => v.toString(16).padStart(2, "0")).join("");
    }
    return `${Date.now().toString(16)}${Math.random().toString(16).slice(2, 18)}`;
  }

  async function createReceiverPinEntry(pin) {
    const salt = createPinSalt();
    const hash = await hashPin(pin, salt);
    return {
      salt,
      hash,
      updatedAt: new Date().toISOString()
    };
  }

  // フロントエンドのみのため完全な秘匿にはならないが、平文保存は避けてハッシュ化して保持する。
  async function hashPin(pin, salt) {
    const payload = `${salt}:${pin}`;
    if (window.crypto && window.crypto.subtle && typeof TextEncoder !== "undefined") {
      const encoded = new TextEncoder().encode(payload);
      const digest = await window.crypto.subtle.digest("SHA-256", encoded);
      return [...new Uint8Array(digest)].map((v) => v.toString(16).padStart(2, "0")).join("");
    }
    return `fallback-${Math.abs(hashString(payload))}`;
  }

  async function verifyReceiverPin(receiverName, pin) {
    if (!isValidPin(pin)) return false;
    const entry = receiverPins[receiverName];
    if (!entry) return false;
    const hash = await hashPin(pin, entry.salt);
    return hash === entry.hash;
  }

  async function requestPinVerification(receiverName) {
    if (!(pinAuthDialog instanceof HTMLDialogElement)) {
      const entered = window.prompt(`${receiverName} のPIN（4桁）を入力してください。`);
      if (entered === null) return false;
      const pin = toSafeString(entered);
      if (!isValidPin(pin)) {
        showMessage("PINは4桁の数字で入力してください。", true);
        return false;
      }
      const valid = await verifyReceiverPin(receiverName, pin);
      if (!valid) showMessage("PINが一致しないため受取確定できません。", true);
      return valid;
    }

    pinAuthReceiverText.textContent = `受取人: ${receiverName}`;
    pinAuthInput.value = "";
    pinAuthError.textContent = "";
    pinAuthError.hidden = true;
    pinAuthDialog.showModal();

    return new Promise((resolve) => {
      pinAuthContext = { receiverName, resolve };
      pinAuthInput.focus();
    });
  }

  async function onSubmitPinAuth(event) {
    event.preventDefault();
    if (!pinAuthContext) return;

    const pin = toSafeString(pinAuthInput.value);
    if (!isValidPin(pin)) {
      showPinAuthError("PINは4桁の数字で入力してください。");
      return;
    }

    const valid = await verifyReceiverPin(pinAuthContext.receiverName, pin);
    if (!valid) {
      showPinAuthError("PINが一致しません。受取確定できません。");
      return;
    }

    const { resolve } = pinAuthContext;
    pinAuthContext = null;
    pinAuthDialog.close("verified");
    resolve(true);
  }

  function onCancelPinAuth() {
    if (!pinAuthContext) {
      if (pinAuthDialog.open) pinAuthDialog.close("cancel");
      return;
    }
    const { resolve } = pinAuthContext;
    pinAuthContext = null;
    if (pinAuthDialog.open) pinAuthDialog.close("cancel");
    resolve(false);
  }

  function onPinAuthDialogCancel(event) {
    event.preventDefault();
    onCancelPinAuth();
  }

  function showPinAuthError(text) {
    pinAuthError.hidden = false;
    pinAuthError.textContent = text;
  }

  function showMessage(text, isError) {
    messageArea.textContent = text;
    messageArea.classList.remove("is-error", "is-success", "is-pop");
    // 再アニメーションのために一度reflowを挟む
    void messageArea.offsetWidth;
    messageArea.classList.add(isError ? "is-error" : "is-success", "is-pop");
  }
})();

