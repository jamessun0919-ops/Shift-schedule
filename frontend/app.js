(() => {
  "use strict";

  const API = "/api";

  const ROW_TYPE_LABELS = {
    shift: "班別（起訖同列）",
    shift_start: "上班時間（此列僅上班）",
    shift_end: "下班時間（此列僅下班）",
    shift_string: "文字班別（如 0900-1900(1400-1500)）",
    metadata: "員編／其他資訊列",
  };

  // ---- App state ----
  let currentFile = null;
  let currentPreviewRows = [];
  let confirmedTemplate = null; // template used for the last successful /api/preview
  let lastEmployees = [];       // full employees array from last /api/preview

  // ---- DOM refs ----
  const $ = (id) => document.getElementById(id);

  const scopeIdInput = $("scope-id");
  const savedTemplateSelect = $("saved-template-select");
  const fileInput = $("file-input");
  const analyzeBtn = $("analyze-btn");
  const uploadStatus = $("upload-status");

  const modalOverlay = $("confirm-modal");
  const modalCloseBtn = $("modal-close-btn");
  const previewGridEl = $("preview-grid");

  const mapSheetName = $("map-sheet-name");
  const mapHeaderRow = $("map-header-row");
  const mapNameCol = $("map-name-col");
  const mapFirstDayCol = $("map-first-day-col");
  const mapColsPerDay = $("map-cols-per-day");
  const mapExpectedRows = $("map-expected-rows");
  const mapNameRowOffset = $("map-name-row-offset");

  const rowMeaningsEditor = $("row-meanings-editor");
  const templateNameInput = $("template-name-input");
  const saveTemplateBtn = $("save-template-btn");
  const saveTemplateStatus = $("save-template-status");

  const confirmPreviewBtn = $("confirm-preview-btn");
  const previewResultText = $("preview-result-text");

  const dayPreviewPanel = $("day-preview-panel");
  const healthBanner = $("health-banner");
  const anomaliesList = $("anomalies-list");
  const daySelect = $("day-select");
  const printBtn = $("print-btn");
  const dayTitle = $("day-title");
  const ganttContainer = $("gantt-container");
  const outputFormatSelect = $("output-format");
  const convertBtn = $("convert-btn");
  const convertStatus = $("convert-status");

  // ---- Helpers ----
  function setStatus(el, text, kind) {
    el.textContent = text || "";
    el.classList.remove("error", "ok");
    if (kind) el.classList.add(kind);
  }

  function safeScopeId() {
    return (scopeIdInput.value || "").trim();
  }

  async function fetchJson(url, options) {
    const res = await fetch(url, options);
    let body = null;
    try { body = await res.json(); } catch (e) { /* ignore */ }
    if (!res.ok) {
      const detail = body && body.detail ? body.detail : `HTTP ${res.status}`;
      throw new Error(detail);
    }
    return body;
  }

  // ---- Saved templates for scope ----
  async function refreshSavedTemplates() {
    const scopeId = safeScopeId();
    savedTemplateSelect.innerHTML = '<option value="">— 不使用已存範本，重新分析 —</option>';
    if (!scopeId) return;
    try {
      const list = await fetchJson(`${API}/templates?scope_id=${encodeURIComponent(scopeId)}`);
      for (const t of list) {
        const opt = document.createElement("option");
        opt.value = t.template_id;
        opt.textContent = t.template_name || t.template_id;
        savedTemplateSelect.appendChild(opt);
      }
    } catch (e) {
      // scope not found yet / no templates -- not an error worth surfacing
    }
  }
  scopeIdInput.addEventListener("blur", refreshSavedTemplates);

  // ---- Analyze ----
  analyzeBtn.addEventListener("click", async () => {
    const file = fileInput.files[0];
    if (!file) {
      setStatus(uploadStatus, "請先選擇檔案", "error");
      return;
    }
    currentFile = file;
    setStatus(uploadStatus, "分析中…", null);
    analyzeBtn.disabled = true;

    try {
      const fd = new FormData();
      fd.append("file", file);
      const analyzeResult = await fetchJson(`${API}/analyze`, { method: "POST", body: fd });
      currentPreviewRows = analyzeResult.preview_rows;

      let template = analyzeResult.suggested_template;

      const savedId = savedTemplateSelect.value;
      if (savedId) {
        const scopeId = safeScopeId();
        try {
          template = await fetchJson(`${API}/templates/${encodeURIComponent(scopeId)}/${encodeURIComponent(savedId)}`);
        } catch (e) {
          setStatus(uploadStatus, `已存範本讀取失敗，改用自動分析結果：${e.message}`, "error");
        }
      }

      setStatus(uploadStatus, "分析完成，請於確認視窗核對欄位。", "ok");
      openModal(template);
    } catch (e) {
      setStatus(uploadStatus, `分析失敗：${e.message}`, "error");
    } finally {
      analyzeBtn.disabled = false;
    }
  });

  // ---- Modal: open / close ----
  function openModal(template) {
    mapSheetName.value = template.sheet_name || "";
    mapHeaderRow.value = template.header_row_index || 1;
    mapNameCol.value = template.mapping?.name_col || 1;
    mapFirstDayCol.value = template.mapping?.first_day_col || 1;
    mapColsPerDay.value = template.mapping?.cols_per_day || 2;
    mapExpectedRows.value = template.block?.expected_rows || 1;
    mapNameRowOffset.value = template.block?.name_row_offset || 0;
    templateNameInput.value = template.template_name || "";

    renderPreviewGrid(currentPreviewRows);
    renderRowMeaningsEditor(template.block?.row_meanings || [], template.block?.expected_rows || 1);

    setStatus(previewResultText, "", null);
    modalOverlay.classList.remove("hidden");
  }

  function closeModal() {
    modalOverlay.classList.add("hidden");
  }
  modalCloseBtn.addEventListener("click", closeModal);
  modalOverlay.addEventListener("click", (e) => {
    if (e.target === modalOverlay) closeModal();
  });

  // ---- Raw preview grid ----
  function renderPreviewGrid(rows) {
    previewGridEl.innerHTML = "";
    if (!rows || !rows.length) return;
    const maxCols = Math.max(...rows.map((r) => r.length));

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    headRow.appendChild(document.createElement("th"));
    for (let c = 1; c <= maxCols; c++) {
      const th = document.createElement("th");
      th.textContent = c;
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    previewGridEl.appendChild(thead);

    const tbody = document.createElement("tbody");
    rows.forEach((row, rIdx) => {
      const tr = document.createElement("tr");
      const rowNumTd = document.createElement("td");
      rowNumTd.textContent = rIdx + 1;
      rowNumTd.style.fontWeight = "600";
      tr.appendChild(rowNumTd);
      for (let c = 0; c < maxCols; c++) {
        const td = document.createElement("td");
        const val = row[c];
        td.textContent = val === null || val === undefined ? "" : val;
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
    previewGridEl.appendChild(tbody);
  }

  // ---- Row meanings editor ----
  function renderRowMeaningsEditor(rowMeanings, expectedRows) {
    rowMeaningsEditor.innerHTML = "";
    const padded = [];
    for (let i = 0; i < expectedRows; i++) {
      padded.push(rowMeanings[i] || { type: "shift", index: i });
    }

    padded.forEach((meaning, offset) => {
      rowMeaningsEditor.appendChild(buildRowMeaningItem(offset, meaning));
    });
  }

  function buildRowMeaningItem(offset, meaning) {
    const item = document.createElement("div");
    item.className = "row-meaning-item";
    item.dataset.offset = String(offset);

    const label = document.createElement("span");
    label.className = "row-label";
    label.textContent = `第 ${offset + 1} 列`;
    item.appendChild(label);

    const select = document.createElement("select");
    select.dataset.role = "type";
    for (const [val, text] of Object.entries(ROW_TYPE_LABELS)) {
      const opt = document.createElement("option");
      opt.value = val;
      opt.textContent = text;
      if (val === meaning.type) opt.selected = true;
      select.appendChild(opt);
    }
    item.appendChild(select);

    const subFieldWrap = document.createElement("div");
    subFieldWrap.className = "sub-field-wrap";
    item.appendChild(subFieldWrap);

    function renderSubFields(type) {
      subFieldWrap.innerHTML = "";
      if (type === "shift" || type === "shift_start" || type === "shift_end") {
        const wrap = document.createElement("label");
        wrap.className = "sub-field";
        wrap.textContent = "班別序號：";
        const input = document.createElement("input");
        input.type = "number";
        input.min = "1";
        input.dataset.role = "shift-number";
        input.value = (typeof meaning.index === "number" ? meaning.index : offset) + 1;
        wrap.appendChild(input);
        subFieldWrap.appendChild(wrap);
      } else if (type === "metadata") {
        const nameWrap = document.createElement("label");
        nameWrap.className = "sub-field";
        nameWrap.textContent = "欄位名稱：";
        const nameInput = document.createElement("input");
        nameInput.type = "text";
        nameInput.dataset.role = "meta-name";
        nameInput.value = meaning.name || "employee_id";
        nameWrap.appendChild(nameInput);
        subFieldWrap.appendChild(nameWrap);

        const colWrap = document.createElement("label");
        colWrap.className = "sub-field";
        colWrap.textContent = "欄位位置（留空＝不辨識）：";
        const colInput = document.createElement("input");
        colInput.type = "number";
        colInput.min = "1";
        colInput.dataset.role = "meta-col";
        if (meaning.col !== undefined && meaning.col !== null) colInput.value = meaning.col;
        colWrap.appendChild(colInput);
        subFieldWrap.appendChild(colWrap);
      }
      // shift_string: no extra fields
    }

    renderSubFields(meaning.type);
    select.addEventListener("change", () => renderSubFields(select.value));

    return item;
  }

  mapExpectedRows.addEventListener("change", () => {
    const n = Math.max(1, Math.min(8, parseInt(mapExpectedRows.value, 10) || 1));
    mapExpectedRows.value = n;
    const existing = collectRowMeanings();
    renderRowMeaningsEditor(existing, n);
  });

  function collectRowMeanings() {
    const items = Array.from(rowMeaningsEditor.querySelectorAll(".row-meaning-item"));
    return items.map((item) => {
      const type = item.querySelector('[data-role="type"]').value;
      if (type === "shift" || type === "shift_start" || type === "shift_end") {
        const numEl = item.querySelector('[data-role="shift-number"]');
        const n = parseInt(numEl?.value, 10) || 1;
        return { type, index: n - 1 };
      }
      if (type === "metadata") {
        const nameEl = item.querySelector('[data-role="meta-name"]');
        const colEl = item.querySelector('[data-role="meta-col"]');
        const meaning = { type: "metadata", name: (nameEl?.value || "employee_id").trim() };
        if (colEl && colEl.value !== "") meaning.col = parseInt(colEl.value, 10);
        return meaning;
      }
      return { type: "shift_string", index: 0 };
    });
  }

  function collectTemplateFromForm() {
    return {
      template_name: templateNameInput.value.trim() || "未命名範本",
      sheet_name: mapSheetName.value,
      header_row_index: parseInt(mapHeaderRow.value, 10) || 1,
      mapping: {
        name_col: parseInt(mapNameCol.value, 10) || 1,
        first_day_col: parseInt(mapFirstDayCol.value, 10) || 1,
        cols_per_day: parseInt(mapColsPerDay.value, 10) || 2,
      },
      block: {
        expected_rows: parseInt(mapExpectedRows.value, 10) || 1,
        name_row_offset: parseInt(mapNameRowOffset.value, 10) || 0,
        row_meanings: collectRowMeanings(),
      },
    };
  }

  // ---- Save as template ----
  saveTemplateBtn.addEventListener("click", async () => {
    const scopeId = safeScopeId();
    if (!scopeId) {
      setStatus(saveTemplateStatus, "請先在上方填寫店家/品牌名稱", "error");
      return;
    }
    const template = collectTemplateFromForm();
    try {
      const result = await fetchJson(`${API}/templates/${encodeURIComponent(scopeId)}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(template),
      });
      setStatus(saveTemplateStatus, `已儲存（ID: ${result.template_id}）`, "ok");
      refreshSavedTemplates();
    } catch (e) {
      setStatus(saveTemplateStatus, `儲存失敗：${e.message}`, "error");
    }
  });

  // ---- Confirm & preview ----
  confirmPreviewBtn.addEventListener("click", async () => {
    if (!currentFile) {
      setStatus(previewResultText, "找不到原始檔案，請重新上傳", "error");
      return;
    }
    const template = collectTemplateFromForm();
    setStatus(previewResultText, "試解析中…", null);
    confirmPreviewBtn.disabled = true;

    try {
      const fd = new FormData();
      fd.append("file", currentFile);
      fd.append("template", JSON.stringify(template));
      const result = await fetchJson(`${API}/preview`, { method: "POST", body: fd });

      confirmedTemplate = template;
      lastEmployees = result.employees;

      setStatus(previewResultText, `解析出 ${result.employees_count} 位員工`, "ok");
      showDayPreviewPanel(result);
      closeModal();
    } catch (e) {
      setStatus(previewResultText, `試解析失敗：${e.message}`, "error");
    } finally {
      confirmPreviewBtn.disabled = false;
    }
  });

  // ---- Day preview panel ----
  function showDayPreviewPanel(result) {
    dayPreviewPanel.classList.remove("hidden");

    if (result.is_healthy) {
      healthBanner.textContent = "健康檢查通過";
      healthBanner.className = "health-banner ok";
    } else {
      healthBanner.textContent = "健康檢查警告：解析結果可能不正確，請檢查下方異常訊息與欄位對照";
      healthBanner.className = "health-banner warn";
    }

    anomaliesList.innerHTML = "";
    for (const a of result.anomalies || []) {
      const li = document.createElement("li");
      li.textContent = a;
      anomaliesList.appendChild(li);
    }

    daySelect.innerHTML = "";
    for (let d = 1; d <= 31; d++) {
      const opt = document.createElement("option");
      opt.value = d;
      opt.textContent = `${d} 日`;
      daySelect.appendChild(opt);
    }
    daySelect.value = "1";
    renderDayGantt(1);

    dayPreviewPanel.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  daySelect.addEventListener("change", () => renderDayGantt(parseInt(daySelect.value, 10)));
  printBtn.addEventListener("click", () => window.print());

  const SERIES_COLORS = ["var(--series-1)", "var(--series-2)", "var(--series-3)", "var(--series-4)"];
  const AXIS_START = 8;  // 08:00
  const AXIS_END = 24;   // 24:00

  function fmtHour(h) {
    if (h === null || h === undefined) return "";
    const hh = Math.floor(h);
    const mm = Math.round((h - hh) * 60);
    return `${String(hh).padStart(2, "0")}:${String(mm).padStart(2, "0")}`;
  }

  function renderDayGantt(day) {
    dayTitle.textContent = `${day} 日 排班預覽`;
    ganttContainer.innerHTML = "";

    const legend = document.createElement("div");
    legend.className = "gantt-legend";
    let maxShiftIdx = 0;
    for (const emp of lastEmployees) {
      const shifts = emp.days[day]?.shifts || [];
      maxShiftIdx = Math.max(maxShiftIdx, shifts.length);
    }
    for (let i = 0; i < Math.min(maxShiftIdx, SERIES_COLORS.length); i++) {
      const item = document.createElement("div");
      item.className = "gantt-legend-item";
      const swatch = document.createElement("span");
      swatch.className = "gantt-legend-swatch";
      swatch.style.background = SERIES_COLORS[i];
      item.appendChild(swatch);
      item.appendChild(document.createTextNode(`班別${i + 1}`));
      legend.appendChild(item);
    }
    ganttContainer.appendChild(legend);

    const axis = document.createElement("div");
    axis.className = "gantt-axis";
    for (let h = AXIS_START; h <= AXIS_END; h += 2) {
      const pct = ((h - AXIS_START) / (AXIS_END - AXIS_START)) * 100;
      const span = document.createElement("span");
      span.style.left = `${pct}%`;
      span.textContent = `${h}:00`;
      axis.appendChild(span);
    }
    ganttContainer.appendChild(axis);

    const employeesWithShifts = lastEmployees.filter((emp) =>
      (emp.days[day]?.shifts || []).some((s) => s.start !== null && s.end !== null)
    );

    if (!employeesWithShifts.length) {
      const empty = document.createElement("div");
      empty.className = "gantt-empty";
      empty.textContent = "當天無排班資料";
      ganttContainer.appendChild(empty);
      return;
    }

    for (const emp of employeesWithShifts) {
      const row = document.createElement("div");
      row.className = "gantt-row";

      const nameEl = document.createElement("div");
      nameEl.className = "gantt-name";
      nameEl.textContent = emp.name;
      nameEl.title = emp.name;
      row.appendChild(nameEl);

      const track = document.createElement("div");
      track.className = "gantt-track";

      const shifts = emp.days[day]?.shifts || [];
      shifts.forEach((s, idx) => {
        if (s.start === null || s.end === null) return;
        const startClamped = Math.max(AXIS_START, Math.min(AXIS_END, s.start));
        const endClamped = Math.max(AXIS_START, Math.min(AXIS_END, s.end));
        if (endClamped <= startClamped) return;

        const leftPct = ((startClamped - AXIS_START) / (AXIS_END - AXIS_START)) * 100;
        const widthPct = ((endClamped - startClamped) / (AXIS_END - AXIS_START)) * 100;

        const bar = document.createElement("div");
        bar.className = "gantt-bar";
        bar.style.left = `${leftPct}%`;
        bar.style.width = `${widthPct}%`;
        bar.style.background = SERIES_COLORS[idx % SERIES_COLORS.length];
        bar.textContent = `${fmtHour(s.start)}-${fmtHour(s.end)}`;
        bar.title = `${emp.name} 班別${idx + 1}：${fmtHour(s.start)}-${fmtHour(s.end)}`;
        track.appendChild(bar);
      });

      row.appendChild(track);
      ganttContainer.appendChild(row);
    }
  }

  // ---- Convert & download ----
  convertBtn.addEventListener("click", async () => {
    if (!currentFile || !confirmedTemplate) {
      setStatus(convertStatus, "請先完成上方「確認並試解析」步驟", "error");
      return;
    }
    setStatus(convertStatus, "轉換中…", null);
    convertBtn.disabled = true;

    try {
      const fd = new FormData();
      fd.append("file", currentFile);
      fd.append("template", JSON.stringify(confirmedTemplate));
      fd.append("output_format", outputFormatSelect.value);

      const res = await fetch(`${API}/convert`, { method: "POST", body: fd });
      if (!res.ok) {
        let detail = `HTTP ${res.status}`;
        try { detail = (await res.json()).detail || detail; } catch (e) { /* ignore */ }
        throw new Error(detail);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `converted.${outputFormatSelect.value}`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      setStatus(convertStatus, "下載完成", "ok");
    } catch (e) {
      setStatus(convertStatus, `轉換失敗：${e.message}`, "error");
    } finally {
      convertBtn.disabled = false;
    }
  });
})();
