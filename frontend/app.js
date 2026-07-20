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
  let currentHiddenCols = new Set(); // 原始檔案中被隱藏的欄（1-based），預覽格線不顯示
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
  const pickHint = $("pick-hint");

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
      currentHiddenCols = new Set(analyzeResult.hidden_cols || []);

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
    mapNameCol.value = colLetter(template.mapping?.name_col || 1);
    mapFirstDayCol.value =
      colLetter(template.mapping?.first_day_col || 1) +
      (template.first_data_row || (template.header_row_index || 1) + 1);
    mapColsPerDay.value = template.mapping?.cols_per_day || 2;
    mapExpectedRows.value = template.block?.expected_rows || 1;
    mapNameRowOffset.value = (template.block?.name_row_offset || 0) + 1;
    // 已存範本（有 template_id）沿用原本名稱；剛分析出的猜測範本則預設「品牌文字＋範本」
    const scopeId = safeScopeId();
    templateNameInput.value = template.template_id
      ? (template.template_name || "")
      : (scopeId ? `${scopeId}範本` : (template.template_name || ""));

    renderPreviewGrid(currentPreviewRows);
    renderRowMeaningsEditor(template.block?.row_meanings || [], template.block?.expected_rows || 1);

    setStatus(previewResultText, "", null);
    modalOverlay.classList.remove("hidden");
  }

  function closeModal() {
    disarm();
    modalOverlay.classList.add("hidden");
  }
  modalCloseBtn.addEventListener("click", closeModal);
  modalOverlay.addEventListener("click", (e) => {
    if (e.target === modalOverlay) closeModal();
  });

  // ---- Raw preview grid ----
  function colLetter(n) {
    let s = "";
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  // 從儲存格參照（如 "C5"）取出欄號（C -> 3）；只取欄，列僅供顯示確認
  function refToCol(ref) {
    const m = String(ref).toUpperCase().match(/^[A-Z]+/);
    if (!m) return NaN;
    let n = 0;
    for (const ch of m[0]) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n;
  }

  function renderPreviewGrid(rows) {
    previewGridEl.innerHTML = "";
    if (!rows || !rows.length) return;
    const maxCols = Math.max(...rows.map((r) => r.length));

    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    headRow.appendChild(document.createElement("th"));
    for (let c = 1; c <= maxCols; c++) {
      if (currentHiddenCols.has(c)) continue;
      const th = document.createElement("th");
      th.textContent = colLetter(c);
      headRow.appendChild(th);
    }
    thead.appendChild(headRow);
    previewGridEl.appendChild(thead);

    const tbody = document.createElement("tbody");
    rows.forEach((row, rIdx) => {
      const tr = document.createElement("tr");
      tr.dataset.row = String(rIdx + 1);
      const rowNumTd = document.createElement("td");
      rowNumTd.textContent = rIdx + 1;
      rowNumTd.style.fontWeight = "600";
      rowNumTd.dataset.row = String(rIdx + 1);
      tr.appendChild(rowNumTd);
      for (let c = 0; c < maxCols; c++) {
        if (currentHiddenCols.has(c + 1)) continue;
        const td = document.createElement("td");
        td.dataset.row = String(rIdx + 1);
        td.dataset.col = String(c + 1);
        const val = row[c];
        td.textContent = val === null || val === undefined ? "" : val;
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    });
    previewGridEl.appendChild(tbody);
    applyGridHighlight();
  }

  // ---- Click-to-fill: 點格子帶入列/欄號 ----
  let armed = null; // { input, axis, btn }

  function applyGridHighlight() {
    if (!previewGridEl.tBodies.length) return;
    const tbody = previewGridEl.tBodies[0];
    tbody.querySelectorAll("td.hl-col").forEach((el) => el.classList.remove("hl-col"));
    tbody.querySelectorAll("tr.hl-row").forEach((el) => el.classList.remove("hl-row"));
    const hr = parseInt(mapHeaderRow.value, 10);
    const cols = [refToCol(mapNameCol.value), refToCol(mapFirstDayCol.value)];
    if (hr) {
      const tr = tbody.querySelector(`tr[data-row="${hr}"]`);
      if (tr) tr.classList.add("hl-row");
    }
    cols.forEach((c) => {
      if (!c) return;
      tbody.querySelectorAll(`td[data-col="${c}"]`).forEach((td) => td.classList.add("hl-col"));
    });
  }

  function clearPickHover() {
    previewGridEl.querySelectorAll("td.pick-col").forEach((el) => el.classList.remove("pick-col"));
    previewGridEl.querySelectorAll("tr.pick-row").forEach((el) => el.classList.remove("pick-row"));
  }

  function disarm() {
    if (armed) armed.btn.classList.remove("armed");
    armed = null;
    previewGridEl.classList.remove("picking");
    clearPickHover();
    pickHint.classList.add("hidden");
  }

  function armPick(btn) {
    if (armed && armed.btn === btn) { disarm(); return; }
    disarm();
    armed = { input: $(btn.dataset.target), axis: btn.dataset.axis, btn };
    btn.classList.add("armed");
    previewGridEl.classList.add("picking");
    const unit = { row: "列號", col: "欄號", cell: "位置" }[armed.axis];
    pickHint.textContent = `點選左側表格中的儲存格，將其${unit}帶入「${btn.dataset.label}」`;
    pickHint.classList.remove("hidden");
  }

  document.querySelectorAll(".pick-btn").forEach((btn) => {
    btn.addEventListener("click", () => armPick(btn));
  });

  [mapHeaderRow, mapNameCol, mapFirstDayCol].forEach((inp) => {
    inp.addEventListener("input", applyGridHighlight);
  });

  previewGridEl.addEventListener("mousemove", (e) => {
    if (!armed) return;
    clearPickHover();
    const td = e.target.closest("td");
    if (!td) return;
    if (armed.axis === "row" && td.dataset.row) {
      td.parentElement.classList.add("pick-row");
    } else if (armed.axis === "col" && td.dataset.col) {
      previewGridEl.tBodies[0]
        .querySelectorAll(`td[data-col="${td.dataset.col}"]`)
        .forEach((x) => x.classList.add("pick-col"));
    } else if (armed.axis === "cell" && td.dataset.col && td.dataset.row) {
      td.classList.add("pick-col");
    }
  });
  previewGridEl.addEventListener("mouseleave", () => { if (armed) clearPickHover(); });

  previewGridEl.addEventListener("click", (e) => {
    if (!armed) return;
    const td = e.target.closest("td");
    if (!td) return;
    let val;
    if (armed.axis === "row") val = td.dataset.row;
    else if (armed.axis === "col") val = colLetter(parseInt(td.dataset.col, 10));
    else if (armed.axis === "cell" && td.dataset.col && td.dataset.row) {
      val = colLetter(parseInt(td.dataset.col, 10)) + td.dataset.row;
    }
    if (val === undefined) return;
    armed.input.value = val;
    td.classList.add("cell-flash");
    setTimeout(() => td.classList.remove("cell-flash"), 500);
    disarm();
    applyGridHighlight();
  });

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
        const colWrap = document.createElement("label");
        colWrap.className = "sub-field";
        colWrap.textContent = "欄位位置（留空＝不辨識）：";
        const colInput = document.createElement("input");
        colInput.type = "text";
        colInput.placeholder = "例如 B";
        colInput.dataset.role = "meta-col";
        if (meaning.col !== undefined && meaning.col !== null) colInput.value = colLetter(meaning.col);
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
        const colEl = item.querySelector('[data-role="meta-col"]');
        const meaning = { type: "metadata", name: "employee_id" };
        const col = colEl && colEl.value.trim() !== "" ? refToCol(colEl.value) : NaN;
        if (!isNaN(col)) meaning.col = col;
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
        name_col: refToCol(mapNameCol.value) || 1,
        first_day_col: refToCol(mapFirstDayCol.value) || 1,
        cols_per_day: parseInt(mapColsPerDay.value, 10) || 2,
      },
      block: {
        expected_rows: parseInt(mapExpectedRows.value, 10) || 1,
        name_row_offset: (parseInt(mapNameRowOffset.value, 10) || 1) - 1,
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

    const axis = document.createElement("div");
    axis.className = "gantt-axis";
    for (let h = AXIS_START; h <= AXIS_END; h += 1) {
      const pct = ((h - AXIS_START) / (AXIS_END - AXIS_START)) * 100;
      const span = document.createElement("span");
      span.style.left = `${pct}%`;
      span.textContent = `${h}`;
      axis.appendChild(span);
    }
    ganttContainer.appendChild(axis);

    if (!lastEmployees.length) {
      const empty = document.createElement("div");
      empty.className = "gantt-empty";
      empty.textContent = "當天無排班資料";
      ganttContainer.appendChild(empty);
      return;
    }

    // 顯示所有員工（含當天沒有可辨識時間段的人，如記錄方式特殊的營運經理），
    // 只是沒有時間段的人不會畫出長條，只留姓名列——與下載檔案的既有行為一致。
    for (const emp of lastEmployees) {
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
        // 目前階段：前後段班別性質上無差異，故不用顏色區分，統一用同一色；
        // SERIES_COLORS 其餘色階保留供後續需求（例如真的需要區分班別時）使用。
        bar.style.background = SERIES_COLORS[0];
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
