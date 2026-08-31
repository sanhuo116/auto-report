const templateCatalog = {
  "saas-general": {
    id: "saas-general",
    scenario: "语音 - SaaS",
    name: "通用模板",
    skill: "saas-xunjian-tongyong",
    taskName: "语音saas巡检",
  },
  "hangzhou-shebao": {
    id: "hangzhou-shebao",
    scenario: "文本",
    name: "杭州社保",
    skill: "hangzhoushebao-report",
    taskName: "杭州社保巡检",
  },
  "text-project": {
    id: "text-project",
    scenario: "文本",
    name: "文本项目通用模板",
    skill: "chatbot-xjbg",
    taskName: "文本项目巡检报告",
  },
  "yuhang-shebao": {
    id: "yuhang-shebao",
    scenario: "语音 - 私有化",
    name: "余杭社保",
    skill: "yuhangshebao-report",
    taskName: "余杭社保巡检",
  },
};

const apiConfig = window.INSPECTION_API_CONFIG || {};
const apiBaseUrl = apiConfig.baseUrl || "";
const apiEndpoints = {
  submit: `${apiBaseUrl}${apiConfig.submitPath || "/api/inspection-requests"}`,
  reports: `${apiBaseUrl}${apiConfig.reportsPath || "/api/inspection-reports"}`,
  retry: `${apiBaseUrl}${apiConfig.retryPath || "/api/inspection-requests/:id/retry"}`,
  archive: `${apiBaseUrl}${apiConfig.archivePath || "/api/inspection-reports/:id/archive"}`,
  download: `${apiBaseUrl}${apiConfig.downloadPath || "/api/inspection-reports/:id/download"}`,
};
const uploadFieldName = apiConfig.fileFieldName || "files";
const tokenStorageKey = apiConfig.tokenStorageKey || "inspectionApiToken";
const feishuDestination = apiConfig.feishuDestination || null;

const state = {
  files: [],
  filter: "all",
  reports: [],
};

const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => Array.from(document.querySelectorAll(selector));

const navItems = $$(".nav-item");
const pages = $$("[data-page-view]");
const breadcrumbCurrent = $("#breadcrumb-current");
const requestForm = $("#request-form");
const fileInput = $("#file-input");
const uploadZone = $("#upload-zone");
const fileList = $("#file-list");
const remarks = $("#remarks");
const charCount = $("#char-count");
const toast = $("#toast");
const toastMessage = $("#toast-message");
const submitButton = requestForm.querySelector('button[type="submit"]');
const authModal = $("#auth-modal");
const authTokenInput = $("#auth-token");
const saveAuthTokenButton = $("#save-auth-token");
let toastTimer;
let isSubmitting = false;
let authRetry = null;

function switchPage(pageName) {
  navItems.forEach((item) => {
    item.classList.toggle("active", item.dataset.page === pageName);
  });
  pages.forEach((page) => {
    page.classList.toggle("active", page.dataset.pageView === pageName);
  });
  breadcrumbCurrent.textContent = pageName === "submit" ? "需求提交" : "报告提取";
  if (pageName === "reports") loadReports();
}

navItems.forEach((item) => {
  item.addEventListener("click", () => switchPage(item.dataset.page));
});

$$(".scenario-option").forEach((option) => {
  option.addEventListener("click", () => {
    $$(".scenario-option").forEach((item) => item.classList.remove("selected"));
    option.classList.add("selected");
    option.querySelector("input").checked = true;
    syncTemplates(option.querySelector("input").value);
  });
});

$$(".template-option").forEach((option) => {
  option.addEventListener("click", () => {
    const scenario = $('input[name="scenario"]:checked').value;
    if (option.dataset.scenario !== scenario) return;
    $$(".template-option").forEach((item) => item.classList.remove("selected"));
    option.classList.add("selected");
    option.querySelector("input").checked = true;
  });
});

function syncTemplates(scenario) {
  const matchingOptions = $$(".template-option").filter((option) => option.dataset.scenario === scenario);
  const checkedOption = matchingOptions.find((option) => option.querySelector("input").checked);
  const activeTemplateId = checkedOption?.dataset.templateId || matchingOptions[0]?.dataset.templateId;
  $$(".template-option").forEach((option) => {
    const isMatch = option.dataset.scenario === scenario;
    const isSelected = isMatch && option.dataset.templateId === activeTemplateId;
    const input = option.querySelector("input");
    const stateLabel = option.querySelector(".template-state");
    option.hidden = !isMatch;
    option.classList.toggle("selected", isSelected);
    option.classList.toggle("muted", false);
    input.disabled = !isMatch;
    input.checked = isSelected;
    stateLabel.textContent = isSelected ? "当前选择" : "可选择";
  });
}

$("#choose-file").addEventListener("click", () => fileInput.click());
uploadZone.addEventListener("click", (event) => {
  if (!event.target.closest("button")) fileInput.click();
});
fileInput.addEventListener("change", (event) => addFiles(event.target.files));

["dragenter", "dragover"].forEach((eventName) => {
  uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadZone.classList.add("dragover");
  });
});
["dragleave", "drop"].forEach((eventName) => {
  uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadZone.classList.remove("dragover");
  });
});
uploadZone.addEventListener("drop", (event) => addFiles(event.dataTransfer.files));

function addFiles(fileCollection) {
  const incoming = Array.from(fileCollection || []);
  const existingKeys = new Set(state.files.map((file) => `${file.name}-${file.size}`));
  incoming.forEach((file) => {
    if (!existingKeys.has(`${file.name}-${file.size}`)) state.files.push(file);
  });
  renderFiles();
  fileInput.value = "";
}

function renderFiles() {
  fileList.innerHTML = state.files
    .map(
      (file, index) => `
        <div class="file-item">
          <span class="file-type">${getFileExtension(file.name)}</span>
          <span class="file-name">${escapeHtml(file.name)}</span>
          <span>${formatBytes(file.size)}</span>
          <button class="remove-file" type="button" data-file-index="${index}" aria-label="移除 ${escapeHtml(file.name)}">×</button>
        </div>
      `,
    )
    .join("");
  $$(".remove-file").forEach((button) => {
    button.addEventListener("click", () => {
      state.files.splice(Number(button.dataset.fileIndex), 1);
      renderFiles();
    });
  });
}

function getFileExtension(name) {
  const extension = name.split(".").pop().toUpperCase();
  return extension.length > 4 ? "FILE" : extension;
}

function formatBytes(bytes) {
  if (!bytes) return "0 KB";
  if (bytes < 1024 * 1024) return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

remarks.addEventListener("input", () => {
  if (remarks.value.length > 500) remarks.value = remarks.value.slice(0, 500);
  charCount.textContent = `${remarks.value.length} / 500`;
});

requestForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (isSubmitting) return;
  if (!state.files.length) {
    showToast("请先上传至少一份通话记录", "error");
    uploadZone.animate(
      [
        { transform: "translateX(0)" },
        { transform: "translateX(-4px)" },
        { transform: "translateX(4px)" },
        { transform: "translateX(0)" },
      ],
      { duration: 220 },
    );
    return;
  }

  const tenant = $("#tenant-name").value.trim();
  const startDate = $("#start-date").value;
  const endDate = $("#end-date").value;
  const scenario = $('input[name="scenario"]:checked').value;
  const templateId = $('input[name="template"]:checked')?.value;
  const template = templateCatalog[templateId];
  if (!template || template.scenario !== scenario) {
    showToast("请选择当前业务场景对应的巡检模板", "error");
    return;
  }
  if (!tenant || !startDate || !endDate) return;

  const remarksValue = remarks.value.trim();
  const payload = {
    tenant,
    scenario,
    templateId: template.id,
    template: template.name,
    skill: template.skill,
    executor: "codex",
    taskName: template.taskName,
    periodStart: startDate,
    periodEnd: endDate,
    remarks: remarksValue,
    destination: feishuDestination,
  };
  const formData = new FormData();
  Object.entries(payload).forEach(([key, value]) => {
    formData.append(key, typeof value === "object" && value !== null ? JSON.stringify(value) : value);
  });
  formData.append("payload", JSON.stringify(payload));
  state.files.forEach((file) => formData.append(uploadFieldName, file, file.name));

  setSubmitLoading(true);
  try {
    const response = await fetchApi(apiEndpoints.submit, {
      method: "POST",
      body: formData,
    });
    const responseData = await parseResponse(response);
    if (!response.ok) {
      throw new Error(responseData?.message || `提交接口返回 ${response.status}`);
    }

    const report = normalizeReport(responseData?.report || responseData?.task || responseData?.data || responseData, payload);
    state.reports.unshift(report);
    requestForm.reset();
    state.files = [];
    renderFiles();
    charCount.textContent = "0 / 500";
    const defaultScenarioOption = $('.scenario-option input[value="文本"]').closest(".scenario-option");
    $$(".scenario-option").forEach((item) => item.classList.remove("selected"));
    defaultScenarioOption.classList.add("selected");
    defaultScenarioOption.querySelector("input").checked = true;
    syncTemplates("文本");
    showToast(`需求已提交，任务编号 ${report.id}`);
    switchPage("reports");
  } catch (error) {
    console.error("inspection request submission failed", error);
    handleApiFailure(error, `提交失败：${error.message || "接口不可用"}`, () => requestForm.requestSubmit());
  } finally {
    setSubmitLoading(false);
  }
});

$("#new-request").addEventListener("click", () => switchPage("submit"));

function setSubmitLoading(loading) {
  isSubmitting = loading;
  submitButton.disabled = loading;
  submitButton.setAttribute("aria-busy", String(loading));
  submitButton.querySelector("span:first-child").textContent = loading ? "正在提交..." : "提交巡检需求";
  submitButton.querySelector(".button-arrow").textContent = loading ? "…" : "→";
}

function getAuthToken() {
  if (apiConfig.authToken) return apiConfig.authToken;
  try {
    return window.localStorage.getItem(tokenStorageKey) || window.sessionStorage.getItem(tokenStorageKey) || "";
  } catch {
    return "";
  }
}

function openAuthModal(retryAction) {
  authRetry = retryAction || null;
  authTokenInput.value = "";
  authModal.hidden = false;
  window.setTimeout(() => authTokenInput.focus(), 0);
}

function closeAuthModal() {
  authModal.hidden = true;
  authRetry = null;
}

function handleApiFailure(error, message, retryAction) {
  if (error.status === 401) openAuthModal(retryAction);
  showToast(message, "error");
}

authModal.addEventListener("click", (event) => {
  if (event.target === authModal) closeAuthModal();
});

$(".auth-close").addEventListener("click", closeAuthModal);
saveAuthTokenButton.addEventListener("click", () => {
  const token = authTokenInput.value.trim().replace(/^Bearer\s+/i, "");
  if (!token) {
    showToast("请输入有效的 Bearer Token", "error");
    authTokenInput.focus();
    return;
  }
  try {
    window.localStorage.setItem(tokenStorageKey, token);
  } catch {
    showToast("浏览器无法保存认证信息", "error");
    return;
  }
  const retryAction = authRetry;
  closeAuthModal();
  showToast("认证信息已保存，正在重试");
  if (retryAction) retryAction();
});

authTokenInput.addEventListener("keydown", (event) => {
  if (event.key === "Enter") saveAuthTokenButton.click();
});

async function fetchApi(url, options = {}) {
  const headers = new Headers(options.headers || {});
  const token = getAuthToken();
  if (token) headers.set("Authorization", `Bearer ${token}`);

  const response = await fetch(url, {
    ...options,
    headers,
    credentials: "include",
  });
  if (response.status === 501) {
    const error = new Error("当前页面连接的是静态文件服务，未配置可接收 POST 的生产 API");
    error.status = 501;
    throw error;
  }
  if (response.status === 404) {
    const error = new Error("生产 API 地址不存在，请检查接口地址配置");
    error.status = 404;
    throw error;
  }
  if (response.status === 401) {
    const error = new Error("登录凭证无效或已过期，请重新登录后再试");
    error.status = 401;
    throw error;
  }
  return response;
}

async function parseResponse(response) {
  const text = await response.text();
  if (!text) return {};
  const contentType = response.headers.get("content-type") || "";
  if (contentType.includes("text/html") || text.trimStart().startsWith("<")) {
    return { message: `服务端未提供该接口（HTTP ${response.status}）` };
  }
  try {
    return JSON.parse(text);
  } catch {
    return { message: text };
  }
}

function normalizeStatus(value) {
  const status = String(value || "").toLowerCase();
  if (status === "success" || status === "succeeded" || status.includes("成功")) return "success";
  if (status === "failed" || status === "failure" || status.includes("失败")) return "failed";
  return "processing";
}

function normalizeReport(raw = {}, fallback = {}) {
  const tenant = raw.tenant || raw.tenantName || fallback.tenant || "";
  const scenario = raw.scenario || fallback.scenario || "";
  const templateId = raw.templateId || fallback.templateId || "";
  const templateDefinition = templateCatalog[templateId];
  const template =
    raw.template ||
    raw.templateName ||
    fallback.template ||
    templateDefinition?.name ||
    "";
  const skill =
    raw.skill ||
    raw.skillName ||
    fallback.skill ||
    templateDefinition?.skill ||
    "";
  const status = normalizeStatus(raw.status || raw.state || fallback.status);
  return {
    id: raw.id || raw.taskId || raw.requestId || `task-${Date.now()}`,
    tenant,
    name: raw.name || raw.reportName || `${tenant}巡检报告`,
    scenario,
    templateId,
    template,
    skill,
    period:
      raw.period ||
      `${formatDate(raw.periodStart || fallback.periodStart || "")} - ${formatDate(raw.periodEnd || fallback.periodEnd || "")}`,
    generatedAt: raw.generatedAt || raw.createdAt || "刚刚",
    status,
    statusText: raw.statusText || (status === "success" ? "成功" : status === "failed" ? "失败" : "制作中"),
    taskName: raw.taskName || raw.task || "",
    downloadUrl: raw.downloadUrl || raw.reportUrl || "",
    feishuSyncStatus: raw.feishuSyncStatus || raw.archiveStatus || "",
    feishuRecordId: raw.feishuRecordId || raw.archiveRecordId || "",
    feishuError: raw.feishuError || raw.archiveError || "",
  };
}

async function loadReports() {
  try {
    const response = await fetchApi(apiEndpoints.reports, {
      method: "GET",
    });
    const responseData = await parseResponse(response);
    if (!response.ok) {
      throw new Error(responseData?.message || `报告列表接口返回 ${response.status}`);
    }
    const records = Array.isArray(responseData)
      ? responseData
      : responseData.reports || responseData.items || responseData.data || [];
    state.reports = records.map((record) => normalizeReport(record));
    renderReports();
  } catch (error) {
    console.error("inspection reports loading failed", error);
    renderReports();
    handleApiFailure(error, `报告列表加载失败：${error.message || "接口不可用"}`, loadReports);
  }
}

$$(".filter-tab").forEach((tab) => {
  tab.addEventListener("click", () => {
    $$(".filter-tab").forEach((item) => item.classList.remove("active"));
    tab.classList.add("active");
    state.filter = tab.dataset.filter;
    renderReports();
  });
});

$("#report-search").addEventListener("input", renderReports);
$("#refresh-reports").addEventListener("click", () => {
  loadReports();
});

function renderReports() {
  const keyword = $("#report-search").value.trim().toLowerCase();
  const filtered = state.reports.filter((report) => {
    const matchesFilter = state.filter === "all" || report.status === state.filter;
    const matchesKeyword =
      !keyword ||
      report.name.toLowerCase().includes(keyword) ||
      report.tenant.toLowerCase().includes(keyword) ||
      report.template.toLowerCase().includes(keyword) ||
      report.skill.toLowerCase().includes(keyword);
    return matchesFilter && matchesKeyword;
  });

  $("#report-table-body").innerHTML = filtered
    .map(
      (report) => `
        <tr>
          <td>
            <div class="report-name">
              <span class="report-file-icon">PDF</span>
              <div>
                <strong title="${escapeHtml(report.name)}">${escapeHtml(report.name)}</strong>
                <span>${escapeHtml(report.tenant)}</span>
              </div>
            </div>
          </td>
          <td>${escapeHtml(report.scenario)}</td>
          <td>
            <span class="template-cell">${escapeHtml(report.template)}</span>
            <small class="skill-cell">${escapeHtml(report.skill)}</small>
            ${
              report.feishuSyncStatus
                ? `<small class="archive-cell ${report.feishuSyncStatus === "已归档" ? "success" : report.feishuSyncStatus === "归档失败" ? "failed" : ""}" title="${escapeHtml(report.feishuError || "")}">飞书：${escapeHtml(report.feishuSyncStatus)}</small>`
                : ""
            }
          </td>
          <td>${escapeHtml(report.period)}</td>
          <td>${escapeHtml(report.generatedAt)}</td>
          <td><span class="status-badge ${report.status}">${report.statusText}</span></td>
          <td class="align-right">
            ${
              report.status === "success"
                ? `<button class="action-button" type="button" data-download-id="${report.id}"><span class="download-symbol">↓</span> 下载</button>${
                    report.feishuSyncStatus && report.feishuSyncStatus !== "已归档"
                      ? `<button class="action-button" type="button" data-archive-id="${report.id}">重新归档</button>`
                      : ""
                  }`
                : report.status === "failed"
                  ? `<button class="action-button" type="button" data-retry-id="${report.id}">重新生成</button>`
                  : `<span class="action-button disabled">制作中</span>`
            }
          </td>
        </tr>
      `,
    )
    .join("");

  $("#empty-state").hidden = filtered.length > 0;
  $("#report-table-body").hidden = filtered.length === 0;
  $("#table-result-count").textContent = `共 ${filtered.length} 条报告`;
  $$("[data-download-id]").forEach((button) => {
    button.addEventListener("click", () => downloadReport(button.dataset.downloadId));
  });
  $$("[data-retry-id]").forEach((button) => {
    button.addEventListener("click", () => retryReport(button.dataset.retryId));
  });
  $$("[data-archive-id]").forEach((button) => {
    button.addEventListener("click", () => archiveReport(button.dataset.archiveId));
  });
  updateCounts();
}

async function retryReport(id) {
  const report = state.reports.find((item) => item.id === id);
  if (!report) return;
  const retryUrl = apiEndpoints.retry.replace(":id", encodeURIComponent(report.id));
  try {
    const response = await fetchApi(retryUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ taskId: report.id }),
    });
    const responseData = await parseResponse(response);
    if (!response.ok) {
      throw new Error(responseData?.message || `重新生成接口返回 ${response.status}`);
    }
    const updatedReport = normalizeReport(responseData?.report || responseData?.task || responseData?.data || responseData, report);
    const reportIndex = state.reports.findIndex((item) => item.id === report.id);
    state.reports[reportIndex] = updatedReport;
    renderReports();
    showToast("已重新提交制作");
  } catch (error) {
    console.error("inspection report retry failed", error);
    handleApiFailure(error, `重新生成失败：${error.message || "接口不可用"}`, () => retryReport(id));
  }
}

async function downloadReport(id) {
  const report = state.reports.find((item) => item.id === id);
  if (!report) return;
  const downloadUrl = report.downloadUrl || apiEndpoints.download.replace(":id", encodeURIComponent(report.id));
  try {
    const response = await fetchApi(downloadUrl);
    if (!response.ok) {
      const responseData = await parseResponse(response);
      throw new Error(responseData?.message || `下载接口返回 ${response.status}`);
    }
    const blob = await response.blob();
    const objectUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = objectUrl;
    link.download = `${report.name}.docx`;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    window.setTimeout(() => {
      URL.revokeObjectURL(objectUrl);
      link.remove();
    }, 1000);
    showToast("报告下载已开始");
  } catch (error) {
    console.error("inspection report download failed", error);
    handleApiFailure(error, `下载失败：${error.message || "接口不可用"}`, () => downloadReport(id));
  }
}

async function archiveReport(id) {
  const archiveUrl = apiEndpoints.archive.replace(":id", encodeURIComponent(id));
  try {
    const response = await fetchApi(archiveUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ taskId: id }),
    });
    const responseData = await parseResponse(response);
    if (!response.ok) {
      throw new Error(responseData?.message || `归档接口返回 ${response.status}`);
    }
    const updatedReport = normalizeReport(responseData?.report || responseData?.task || responseData?.data || responseData);
    const reportIndex = state.reports.findIndex((item) => item.id === id);
    if (reportIndex >= 0) state.reports[reportIndex] = updatedReport;
    renderReports();
    showToast("已重新提交飞书归档");
  } catch (error) {
    console.error("inspection report archive failed", error);
    handleApiFailure(error, `归档失败：${error.message || "接口不可用"}`, () => archiveReport(id));
  }
}

function updateCounts() {
  const counts = state.reports.reduce(
    (result, report) => {
      result.all += 1;
      result[report.status] += 1;
      return result;
    },
    { all: 0, processing: 0, success: 0, failed: 0 },
  );
  $("#all-count").textContent = counts.all;
  $("#processing-count").textContent = counts.processing;
  $("#success-count").textContent = counts.success;
  $("#failed-count").textContent = counts.failed;
  $("[data-filter='all'] span").textContent = counts.all;
  $("[data-filter='processing'] span").textContent = counts.processing;
  $("[data-filter='success'] span").textContent = counts.success;
  $("[data-filter='failed'] span").textContent = counts.failed;
}

function formatDate(value) {
  return value.replaceAll("-", ".");
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (character) => {
    const entities = { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;" };
    return entities[character];
  });
}

function showToast(message, type = "success") {
  toastMessage.textContent = message;
  toast.classList.toggle("error", type === "error");
  toast.classList.add("show");
  window.clearTimeout(toastTimer);
  toastTimer = window.setTimeout(() => toast.classList.remove("show"), 2800);
}

renderReports();
syncTemplates("文本");
