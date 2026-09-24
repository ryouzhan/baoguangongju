document.addEventListener("DOMContentLoaded", () => {
  initTabs();
  initSettingsModal();
  initTab1CustomsDoc();
  initTab2FbaMerger();
  initTab3Declaration();
  fetchInitialConfig();
});

function showToast(message, type = "info") {
  const container = document.getElementById("toast-container");
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;

  let iconSvg = "";
  if (type === "success") {
    iconSvg = `<svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path></svg>`;
  } else if (type === "error") {
    iconSvg = `<svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>`;
  } else {
    iconSvg = `<svg width="20" height="20" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>`;
  }

  toast.innerHTML = `${iconSvg}<span>${escapeHtml(message)}</span>`;
  container.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = "0";
    toast.style.transform = "translateX(100%)";
    toast.style.transition = "all 0.3s ease";
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

function escapeHtml(str) {
  if (!str) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function formatBytes(bytes) {
  if (bytes === 0) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + " " + sizes[i];
}

function addLog(terminalId, msg, type = "info") {
  const terminal = document.getElementById(terminalId);
  if (!terminal) return;
  const entry = document.createElement("div");
  entry.className = `log-entry ${type}`;
  const now = new Date().toTimeString().split(" ")[0];
  entry.textContent = `[${now}] ${msg}`;
  terminal.appendChild(entry);
  terminal.scrollTop = terminal.scrollHeight;
}

function initTabs() {
  const tabBtns = document.querySelectorAll(".tab-btn");
  const tabPanels = document.querySelectorAll(".tab-panel");

  tabBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      tabBtns.forEach(b => b.classList.remove("active"));
      tabPanels.forEach(p => p.classList.remove("active"));

      btn.classList.add("active");
      const targetId = btn.getAttribute("data-tab");
      const targetPanel = document.getElementById(targetId);
      if (targetPanel) {
        targetPanel.classList.add("active");
      }
    });
  });
}

let currentSystemConfig = null;

function initSettingsModal() {
  const modal = document.getElementById("settingsModal");
  const openBtn = document.getElementById("openSettingsBtn");
  const closeBtn = document.getElementById("closeSettingsBtn");
  const saveBtn = document.getElementById("saveConfigBtn");
  const resetBtn = document.getElementById("resetDefaultConfigBtn");

  openBtn.addEventListener("click", () => {
    loadSettingsIntoForm();
    modal.classList.add("show");
  });

  closeBtn.addEventListener("click", () => {
    modal.classList.remove("show");
  });

  modal.addEventListener("click", (e) => {
    if (e.target === modal) modal.classList.remove("show");
  });

  saveBtn.addEventListener("click", async () => {
    try {
      let parsedCfg = {};
      try {
        parsedCfg = JSON.parse(document.getElementById("cfgJsonEditor").value);
      } catch (err) {
        showToast("JSON 语法解析错误，请检查格式！", "error");
        return;
      }

      parsedCfg.airscript = parsedCfg.airscript || {};
      parsedCfg.airscript.webhook_url = document.getElementById("cfgWebhookUrl").value.trim();
      parsedCfg.airscript.token = document.getElementById("cfgToken").value.trim();

      const newRate = parseFloat(document.getElementById("cfgExchangeRate").value) || 7.2;

      saveBtn.disabled = true;
      saveBtn.textContent = "保存中...";

      const res = await fetch("/api/config", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ config: parsedCfg, exchange_rate: newRate })
      });
      const data = await res.json();
      if (data.status === "ok") {
        currentSystemConfig = data.config;
        showToast("系统配置已成功保存更新！", "success");
        modal.classList.remove("show");
        document.getElementById("usdExchangeRate").value = newRate.toFixed(2);
      } else {
        showToast(data.message || "保存失败", "error");
      }
    } catch (e) {
      showToast("网络请求异常: " + e.message, "error");
    } finally {
      saveBtn.disabled = false;
      saveBtn.textContent = "保存所有更改";
    }
  });

  resetBtn.addEventListener("click", () => {
    if (confirm("确定恢复默认系统配置吗？")) {
      fetchInitialConfig(true);
    }
  });
}

async function fetchInitialConfig(isReset = false) {
  try {
    const res = await fetch("/api/config");
    const data = await res.json();
    if (data.status === "ok") {
      currentSystemConfig = data.config;
      if (data.exchange_rate) {
        document.getElementById("usdExchangeRate").value = parseFloat(data.exchange_rate).toFixed(2);
      }
      if (isReset) {
        loadSettingsIntoForm();
        showToast("已重置为当前配置文件内容", "info");
      }
    }
  } catch (e) {
    console.error("加载配置失败:", e);
  }
}

function loadSettingsIntoForm() {
  if (!currentSystemConfig) return;
  const air = currentSystemConfig.airscript || {};
  document.getElementById("cfgWebhookUrl").value = air.webhook_url || "";
  document.getElementById("cfgToken").value = air.token || "";
  document.getElementById("cfgExchangeRate").value = document.getElementById("usdExchangeRate").value;
  document.getElementById("cfgJsonEditor").value = JSON.stringify(currentSystemConfig, null, 2);
}

// ==================== 1. 报关资料在线生成 ====================
function initTab1CustomsDoc() {
  const dropzone = document.getElementById("deliveryDropzone");
  const fileInput = document.getElementById("deliveryFileInput");
  const fileBadge = document.getElementById("deliveryFileBadge");
  const fileNameSpan = document.getElementById("deliveryFileName");
  const fileSizeSpan = document.getElementById("deliveryFileSize");
  const removeBtn = document.getElementById("removeDeliveryFileBtn");
  const startBtn = document.getElementById("startCustomsDocBtn");
  const rateInput = document.getElementById("usdExchangeRate");
  const rateChips = document.querySelectorAll(".rate-chip[data-rate]");
  const testAirscriptBtn = document.getElementById("testAirscriptBtn");
  const clearLogBtn = document.getElementById("terminalClearBtn");

  let selectedFile = null;

  rateChips.forEach(chip => {
    chip.addEventListener("click", () => {
      rateInput.value = parseFloat(chip.getAttribute("data-rate")).toFixed(2);
    });
  });

  testAirscriptBtn.addEventListener("click", async () => {
    testAirscriptBtn.disabled = true;
    testAirscriptBtn.textContent = "正在连接云端...";
    addLog("customsTerminalBody", "正在连接 AirScript Webhook 验证采购单同步状态...", "info");

    try {
      const res = await fetch("/api/airscript/test", { method: "POST" });
      const data = await res.json();
      if (data.status === "ok") {
        showToast(data.message, "success");
        addLog("customsTerminalBody", `✅ 云端同步成功！获取到 ${data.count} 条有效采购单数据`, "success");
        document.getElementById("cloudStatusText").textContent = `AirScript 云端：已同步 (${data.count}条)`;
      } else {
        showToast(data.message, "error");
        addLog("customsTerminalBody", `❌ ${data.message}`, "error");
      }
    } catch (e) {
      showToast("连接失败: " + e.message, "error");
      addLog("customsTerminalBody", `❌ 网络请求异常: ${e.message}`, "error");
    } finally {
      testAirscriptBtn.disabled = false;
      testAirscriptBtn.textContent = "测试云端同步";
    }
  });

  dropzone.addEventListener("click", () => fileInput.click());

  dropzone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });

  dropzone.addEventListener("dragleave", () => {
    dropzone.classList.remove("dragover");
  });

  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleDeliveryFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleDeliveryFile(e.target.files[0]);
    }
  });

  function handleDeliveryFile(file) {
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      showToast("发货单格式必须为 .xlsx 文件！", "error");
      return;
    }
    selectedFile = file;
    fileNameSpan.textContent = file.name;
    fileSizeSpan.textContent = `(${formatBytes(file.size)})`;
    fileBadge.classList.add("show");
    startBtn.disabled = false;
    addLog("customsTerminalBody", `已加载发货单文件: ${file.name} (${formatBytes(file.size)})`, "info");
  }

  removeBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    selectedFile = null;
    fileInput.value = "";
    fileBadge.classList.remove("show");
    startBtn.disabled = true;
    addLog("customsTerminalBody", "发货单已移除", "warn");
  });

  clearLogBtn.addEventListener("click", () => {
    document.getElementById("customsTerminalBody").innerHTML = "";
  });

  startBtn.addEventListener("click", async () => {
    if (!selectedFile) return;
    const rate = parseFloat(rateInput.value);
    if (!rate || rate <= 0) {
      showToast("请输入有效的美元汇率（必须大于0）", "error");
      return;
    }

    startBtn.disabled = true;
    startBtn.innerHTML = `<span class="status-dot"></span> 正在生成中...`;
    document.getElementById("customsStatsContainer").style.display = "none";
    document.getElementById("customsDownloadBanner").classList.remove("show");

    addLog("customsTerminalBody", "================== 开始报关资料生成任务 ==================", "info");

    const formData = new FormData();
    formData.append("delivery_file", selectedFile);
    formData.append("exchange_rate", rate);

    try {
      const res = await fetch("/api/customs_doc/generate", {
        method: "POST",
        body: formData
      });
      const data = await res.json();

      if (data.logs && Array.isArray(data.logs)) {
        data.logs.forEach(l => {
          const type = l.includes("❌") ? "error" : l.includes("⚠️") ? "warn" : l.includes("✅") || l.includes("🎉") ? "success" : "info";
          addLog("customsTerminalBody", l, type);
        });
      }

      if (data.status === "ok") {
        const stats = data.stats;
        showToast("🎉 报关资料生成成功！", "success");

        document.getElementById("statCny").textContent = `¥${stats.total_cny.toLocaleString(undefined, {minimumFractionDigits: 2})}`;
        document.getElementById("statUsd").textContent = `$${stats.total_usd.toLocaleString(undefined, {minimumFractionDigits: 2})}`;
        document.getElementById("statWeight").textContent = `${stats.total_gw} / ${stats.total_nw}`;
        document.getElementById("statCbm").textContent = `${stats.total_cbm}`;
        document.getElementById("statCounts").textContent = `${stats.detail_count} 行 / ${stats.summary_count} 行`;
        document.getElementById("statUnmatch").textContent = `${stats.unmatch_count}`;

        document.getElementById("customsStatsContainer").style.display = "block";

        document.getElementById("customsDownloadFilename").textContent = stats.filename;
        const dlLink = document.getElementById("customsDownloadLink");
        dlLink.href = data.download_url;
        dlLink.setAttribute("download", stats.filename);
        document.getElementById("customsDownloadBanner").classList.add("show");

      } else {
        showToast(data.message || "生成失败！", "error");
        addLog("customsTerminalBody", `❌ 生成失败: ${data.message}`, "error");
      }
    } catch (e) {
      showToast("请求失败: " + e.message, "error");
      addLog("customsTerminalBody", `❌ 网络或服务异常: ${e.message}`, "error");
    } finally {
      startBtn.disabled = false;
      startBtn.innerHTML = `
        <svg width="18" height="18" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 10V3L4 14h7v7l9-11h-7z"></path>
        </svg>
        生成报关资料
      `;
    }
  });
}

// ==================== 2. FBA 报关数据合并 ====================
function initTab2FbaMerger() {
  const dropzone = document.getElementById("fbaSourceDropzone");
  const fileInput = document.getElementById("fbaSourceFileInput");
  const hintText = document.getElementById("fbaSourceHint");
  const transferLayout = document.getElementById("fbaTransferLayout");
  const exportBtnContainer = document.getElementById("fbaExportBtnContainer");

  const allListUl = document.getElementById("fbaAllList");
  const curGroupUl = document.getElementById("fbaCurGroupList");
  const savedGroupsDiv = document.getElementById("fbaSavedGroupsList");

  const addrCountSpan = document.getElementById("fbaAddrCount");
  const curGroupCountSpan = document.getElementById("fbaCurGroupCount");
  const savedGroupCountSpan = document.getElementById("fbaSavedGroupCount");

  const searchInput = document.getElementById("fbaSearchInput");
  const selectAllBtn = document.getElementById("fbaSelectAllBtn");
  const deselectAllBtn = document.getElementById("fbaDeselectAllBtn");

  const addBtn = document.getElementById("fbaAddToGroupBtn");
  const removeBtn = document.getElementById("fbaRemoveFromGroupBtn");
  const clearCurBtn = document.getElementById("fbaClearCurGroupBtn");
  const saveGroupBtn = document.getElementById("fbaSaveGroupBtn");

  const batchTextarea = document.getElementById("fbaBatchTextarea");
  const batchAddBtn = document.getElementById("fbaBatchAddBtn");
  const clearAllSavedBtn = document.getElementById("fbaClearAllSavedGroupsBtn");

  const exportBtn = document.getElementById("fbaExportBtn");

  let sessionId = null;
  let allAddresses = [];
  let fbaDisplayMap = {};
  let currentGroup = [];
  let savedGroups = [];

  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("dragover", (e) => { e.preventDefault(); dropzone.classList.add("dragover"); });
  dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
  dropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      loadFbaFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files.length > 0) {
      loadFbaFile(e.target.files[0]);
    }
  });

  async function loadFbaFile(file) {
    hintText.textContent = `正在解析 ${file.name}...`;
    const formData = new FormData();
    formData.append("source_file", file);

    try {
      const res = await fetch("/api/fba_merger/load", { method: "POST", body: formData });
      const data = await res.json();
      if (data.status === "ok") {
        sessionId = data.session_id;
        allAddresses = data.address_list || [];
        fbaDisplayMap = {};
        allAddresses.forEach(item => {
          if (item.includes("(") && item.includes(")")) {
            const fba = item.substring(item.lastIndexOf("(") + 1, item.lastIndexOf(")")).trim();
            if (fba) fbaDisplayMap[fba] = item;
          }
        });

        hintText.textContent = `已成功解析: ${file.name} | 共 ${data.row_count} 行数据，已提取 ${data.address_count} 个去重配送地址`;
        addrCountSpan.textContent = allAddresses.length;

        currentGroup = [];
        savedGroups = [];
        renderAllList();
        renderCurGroup();
        renderSavedGroups();

        transferLayout.style.display = "grid";
        exportBtnContainer.style.display = "block";
        document.getElementById("fbaDownloadBanner").classList.remove("show");
        showToast(`已成功载入 ${data.address_count} 个地址`, "success");
      } else {
        hintText.textContent = `解析失败: ${data.message}`;
        showToast(data.message, "error");
      }
    } catch (e) {
      hintText.textContent = `解析异常: ${e.message}`;
      showToast("文件上传解析异常: " + e.message, "error");
    }
  }

  function renderAllList() {
    const query = searchInput.value.trim().toLowerCase();
    allListUl.innerHTML = "";
    allAddresses.forEach(item => {
      if (!query || item.toLowerCase().includes(query)) {
        const li = document.createElement("li");
        li.className = "transfer-item";
        li.textContent = item;
        li.addEventListener("click", () => {
          li.classList.toggle("selected");
        });
        allListUl.appendChild(li);
      }
    });
  }

  searchInput.addEventListener("input", renderAllList);

  selectAllBtn.addEventListener("click", () => {
    allListUl.querySelectorAll(".transfer-item").forEach(li => li.classList.add("selected"));
  });

  deselectAllBtn.addEventListener("click", () => {
    allListUl.querySelectorAll(".transfer-item").forEach(li => li.classList.remove("selected"));
  });

  addBtn.addEventListener("click", () => {
    const selectedLis = allListUl.querySelectorAll(".transfer-item.selected");
    selectedLis.forEach(li => {
      const val = li.textContent.trim();
      if (!currentGroup.includes(val)) {
        currentGroup.push(val);
      }
      li.classList.remove("selected");
    });
    renderCurGroup();
  });

  removeBtn.addEventListener("click", () => {
    const selectedLis = curGroupUl.querySelectorAll(".transfer-item.selected");
    const toRemove = Array.from(selectedLis).map(li => li.textContent.trim());
    currentGroup = currentGroup.filter(item => !toRemove.includes(item));
    renderCurGroup();
  });

  clearCurBtn.addEventListener("click", () => {
    currentGroup = [];
    renderCurGroup();
  });

  function renderCurGroup() {
    curGroupUl.innerHTML = "";
    currentGroup.forEach(item => {
      const li = document.createElement("li");
      li.className = "transfer-item";
      li.textContent = item;
      li.addEventListener("click", () => {
        li.classList.toggle("selected");
      });
      curGroupUl.appendChild(li);
    });
    curGroupCountSpan.textContent = currentGroup.length;
  }

  batchAddBtn.addEventListener("click", () => {
    const text = batchTextarea.value.trim();
    if (!text) {
      showToast("请在输入框粘贴待添加的 FBA 编号！", "warn");
      return;
    }
    const lines = text.split("\n").map(l => l.trim()).filter(l => l);
    let addedCount = 0;
    lines.forEach(fba => {
      if (fbaDisplayMap[fba]) {
        const disp = fbaDisplayMap[fba];
        if (!currentGroup.includes(disp)) {
          currentGroup.push(disp);
          addedCount++;
        }
      }
    });
    renderCurGroup();
    showToast(`根据 FBA 编号匹配到并添加了 ${addedCount} 个地址`, "success");
    batchTextarea.value = "";
  });

  saveGroupBtn.addEventListener("click", () => {
    if (currentGroup.length === 0) {
      showToast("当前待合并组为空，不能保存！", "warn");
      return;
    }
    savedGroups.push([...currentGroup]);
    currentGroup = [];
    renderCurGroup();
    renderSavedGroups();
    showToast("已成功保存一个合并分组！", "success");
  });

  function renderSavedGroups() {
    savedGroupsDiv.innerHTML = "";
    savedGroups.forEach((grp, idx) => {
      const chip = document.createElement("div");
      chip.className = "saved-group-chip";
      const preview = grp.slice(0, 2).join(" + ") + (grp.length > 2 ? ` ...等 ${grp.length} 个地址` : "");
      chip.innerHTML = `
        <div>
          <strong>组 ${idx + 1} (${grp.length}项):</strong>
          <span style="font-size: 0.8rem; margin-left: 6px; color: #cbd5e1;">${escapeHtml(preview)}</span>
        </div>
        <button class="btn-remove-file" title="删除此组">✕</button>
      `;
      chip.querySelector("button").addEventListener("click", () => {
        savedGroups.splice(idx, 1);
        renderSavedGroups();
      });
      savedGroupsDiv.appendChild(chip);
    });
    savedGroupCountSpan.textContent = savedGroups.length;
    exportBtn.disabled = savedGroups.length === 0;
  }

  clearAllSavedBtn.addEventListener("click", () => {
    savedGroups = [];
    renderSavedGroups();
  });

  exportBtn.addEventListener("click", async () => {
    if (!sessionId) {
      showToast("会话失效，请重新上传文件！", "error");
      return;
    }
    if (savedGroups.length === 0) {
      showToast("请至少保存一个合并分组！", "warn");
      return;
    }

    exportBtn.disabled = true;
    exportBtn.textContent = "正在合并计算...";
    document.getElementById("fbaDownloadBanner").classList.remove("show");

    try {
      const res = await fetch("/api/fba_merger/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          session_id: sessionId,
          groups: savedGroups
        })
      });
      const data = await res.json();
      if (data.status === "ok") {
        showToast("FBA 数据智能合并完成！", "success");
        document.getElementById("fbaDownloadFilename").textContent = data.stats.filename;
        document.getElementById("fbaDownloadDesc").textContent = `已聚合 ${data.stats.groups_count} 个合并组，共计追加 ${data.stats.merged_rows_count} 行合并数据到原工作表尾部`;
        const dlLink = document.getElementById("fbaDownloadLink");
        dlLink.href = data.download_url;
        dlLink.setAttribute("download", data.stats.filename);
        document.getElementById("fbaDownloadBanner").classList.add("show");
      } else {
        showToast(data.message || "合并失败", "error");
      }
    } catch (e) {
      showToast("导出请求失败: " + e.message, "error");
    } finally {
      exportBtn.disabled = false;
      exportBtn.textContent = "执行分组智能合并并导出 Excel";
    }
  });
}

// ==================== 3. 报关单套版导出 (集成智能清理与重命名) ====================
function initTab3Declaration() {
  const tplDropzone = document.getElementById("declTemplateDropzone");
  const tplInput = document.getElementById("declTemplateFileInput");
  const tplText = document.getElementById("declTemplateName");

  const dataDropzone = document.getElementById("declDataDropzone");
  const dataInput = document.getElementById("declDataFileInput");
  const dataText = document.getElementById("declDataName");

  const startBtn = document.getElementById("startDeclBtn");

  let tplFile = null;
  let dataFile = null;

  tplDropzone.addEventListener("click", () => tplInput.click());
  tplDropzone.addEventListener("dragover", (e) => { e.preventDefault(); tplDropzone.classList.add("dragover"); });
  tplDropzone.addEventListener("dragleave", () => tplDropzone.classList.remove("dragover"));
  tplDropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    tplDropzone.classList.remove("dragover");
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleTplFile(e.dataTransfer.files[0]);
    }
  });
  tplInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleTplFile(e.target.files[0]);
    }
  });

  function handleTplFile(file) {
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      showToast("模板文件必须为 .xlsx！", "error");
      return;
    }
    tplFile = file;
    tplText.textContent = `已选: ${file.name}`;
    checkReady();
    addLog("declTerminalBody", `已加载模板文件: ${file.name}`, "info");
  }

  dataDropzone.addEventListener("click", () => dataInput.click());
  dataDropzone.addEventListener("dragover", (e) => { e.preventDefault(); dataDropzone.classList.add("dragover"); });
  dataDropzone.addEventListener("dragleave", () => dataDropzone.classList.remove("dragover"));
  dataDropzone.addEventListener("drop", (e) => {
    e.preventDefault();
    dataDropzone.classList.remove("dragover");
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleDataFile(e.dataTransfer.files[0]);
    }
  });
  dataInput.addEventListener("change", (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleDataFile(e.target.files[0]);
    }
  });

  function handleDataFile(file) {
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      showToast("数据源清单必须为 .xlsx！", "error");
      return;
    }
    dataFile = file;
    dataText.textContent = `已选: ${file.name}`;
    checkReady();
    addLog("declTerminalBody", `已加载数据源清单: ${file.name}`, "info");
  }

  function checkReady() {
    startBtn.disabled = !(tplFile && dataFile);
  }

  startBtn.addEventListener("click", async () => {
    if (!tplFile || !dataFile) return;

    startBtn.disabled = true;
    startBtn.textContent = "正在批量生成成品报关单...";
    document.getElementById("declDownloadBanner").classList.remove("show");

    addLog("declTerminalBody", "================== 开始套版生成与直接清理重命名任务 ==================", "info");

    const formData = new FormData();
    formData.append("template_file", tplFile);
    formData.append("data_file", dataFile);

    try {
      const res = await fetch("/api/declaration/generate", {
        method: "POST",
        body: formData
      });
      const data = await res.json();

      if (data.logs && Array.isArray(data.logs)) {
        data.logs.forEach(l => {
          const type = l.includes("❌") ? "error" : l.includes("⚠️") ? "warn" : l.includes("✅") ? "success" : "info";
          addLog("declTerminalBody", l, type);
        });
      }

      if (data.status === "ok") {
        showToast(`🎉 成功直接生成 ${data.success_count} 份成品报关单！`, "success");
        const isZip = data.download_url.endsWith(".zip");
        document.getElementById("declDownloadFilename").textContent = isZip 
          ? `成品报关单打包文件 (${data.success_count}份)` 
          : (data.file_names && data.file_names[0] ? data.file_names[0] : "成品报关单.xlsx");
        document.getElementById("declDownloadDesc").textContent = `共处理 ${data.total_fba} 个 FBA 分组，已自动清理未填空白行并按合同号重命名，生成即成品`;
        const dlLink = document.getElementById("declDownloadLink");
        dlLink.href = data.download_url;
        document.getElementById("declDownloadBanner").classList.add("show");
      } else {
        showToast(data.message || "生成失败", "error");
        addLog("declTerminalBody", `❌ 生成异常: ${data.message}`, "error");
      }
    } catch (e) {
      showToast("生成异常: " + e.message, "error");
      addLog("declTerminalBody", `❌ 网络或服务错误: ${e.message}`, "error");
    } finally {
      startBtn.disabled = false;
      startBtn.textContent = "启动套版直接生成成品报关单";
    }
  });
}
