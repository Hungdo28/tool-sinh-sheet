/**
 * Web Application Logic - Test Case Sheet Generator
 * Xử lý đọc file Excel / Google Sheets, phân tích nhóm, và sinh file Excel mới đa sheet trên trình duyệt.
 */

// ================= GLOBAL STATE =================
let currentWorkbook = null;
let currentFileName = "sample.xlsx";
let currentSheetName = "";
let currentSheetData = [];
let detectedGroups = {};
let selectedGroupNames = new Set();
let lastGeneratedBlob = null;

// ================= DOM ELEMENTS =================
const fileInput = document.getElementById("file-input");
const dropZone = document.getElementById("drop-zone");
const fileInfo = document.getElementById("file-info");
const fileNameDisplay = document.getElementById("file-name-display");
const btnRemoveFile = document.getElementById("btn-remove-file");

const tabFile = document.getElementById("tab-file");
const tabGsheet = document.getElementById("tab-gsheet");
const paneFile = document.getElementById("source-file-pane");
const paneGsheet = document.getElementById("source-gsheet-pane");
const gsheetUrlInput = document.getElementById("gsheet-url");
const btnFetchGsheet = document.getElementById("btn-fetch-gsheet");

const sheetSelect = document.getElementById("sheet-select");
const headerRowInput = document.getElementById("header-row");
const groupColSelect = document.getElementById("group-col-select");
const chkForwardFill = document.getElementById("chk-forward-fill");

const chkOverview = document.getElementById("chk-overview");
const outputFilenameInput = document.getElementById("output-filename");

const groupsTableBody = document.getElementById("groups-table-body");
const previewSummaryBadge = document.getElementById("preview-summary-badge");
const btnSelectAll = document.getElementById("btn-select-all");
const btnDeselectAll = document.getElementById("btn-deselect-all");

const btnGenerate = document.getElementById("btn-generate");
const progressBox = document.getElementById("progress-box");
const progressFill = document.getElementById("progress-fill");
const progressText = document.getElementById("progress-text");

const successModal = document.getElementById("success-modal");
const btnCloseModal = document.getElementById("btn-close-modal");
const btnDownloadAgain = document.getElementById("btn-download-again");
const btnLoadSample = document.getElementById("btn-load-sample");
const statSheetsCount = document.getElementById("stat-sheets-count");
const statTcCount = document.getElementById("stat-tc-count");

const ruleEachCase = document.getElementById("rule-each-case");
const ruleByGroup = document.getElementById("rule-by-group");
const cardRuleEachCase = document.getElementById("card-rule-each-case");
const cardRuleByGroup = document.getElementById("card-rule-by-group");
const lblColSelect = document.getElementById("lbl-col-select");
const groupColHint = document.getElementById("group-col-hint");

// ================= INIT & EVENT LISTENERS =================
document.addEventListener("DOMContentLoaded", () => {
    setupTabSwitching();
    setupDragAndDrop();
    setupFormEvents();
    setupSampleLoader();
    setupPreviewTabSwitching();

    // Tự động nạp mẫu 20 cột chuẩn ngay khi mở để người dùng xem luôn
    setTimeout(() => {
        if (!currentWorkbook && btnLoadSample) {
            btnLoadSample.click();
        }
    }, 200);
});

// Tab switching
function setupTabSwitching() {
    tabFile.addEventListener("click", () => {
        tabFile.classList.add("active");
        tabGsheet.classList.remove("active");
        paneFile.classList.remove("hidden");
        paneGsheet.classList.add("hidden");
    });

    tabGsheet.addEventListener("click", () => {
        tabGsheet.classList.add("active");
        tabFile.classList.remove("active");
        paneGsheet.classList.remove("hidden");
        paneFile.classList.add("hidden");
    });
}

// Drag & drop file handling
function setupDragAndDrop() {
    dropZone.addEventListener("click", () => fileInput.click());

    fileInput.addEventListener("change", (e) => {
        if (e.target.files && e.target.files.length > 0) {
            handleFileSelected(e.target.files[0]);
        }
    });

    ["dragenter", "dragover"].forEach(name => {
        dropZone.addEventListener(name, (e) => {
            e.preventDefault();
            dropZone.classList.add("dragover");
        });
    });

    ["dragleave", "drop"].forEach(name => {
        dropZone.addEventListener(name, (e) => {
            e.preventDefault();
            dropZone.classList.remove("dragover");
        });
    });

    dropZone.addEventListener("drop", (e) => {
        if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            handleFileSelected(e.dataTransfer.files[0]);
        }
    });

    btnRemoveFile.addEventListener("click", () => {
        currentWorkbook = null;
        fileInput.value = "";
        dropZone.classList.remove("hidden");
        fileInfo.classList.add("hidden");
        sheetSelect.innerHTML = '<option value="">(Chưa chọn file)</option>';
        sheetSelect.disabled = true;
        groupColSelect.innerHTML = '<option value="">(Chọn sheet trước)</option>';
        groupColSelect.disabled = true;
        btnGenerate.disabled = true;
        renderEmptyTable();
    });
}

// Google Sheets Fetching
btnFetchGsheet.addEventListener("click", async () => {
    const url = gsheetUrlInput.value.trim();
    if (!url) {
        alert("Vui lòng nhập đường dẫn Google Sheets.");
        return;
    }

    const match = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) {
        alert("Đường dẫn Google Sheets không đúng định dạng. Cần có dạng:\nhttps://docs.google.com/spreadsheets/d/.../edit");
        return;
    }

    const sheetId = match[1];
    const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;

    btnFetchGsheet.disabled = true;
    btnFetchGsheet.textContent = "⏳ Đang tải...";

    try {
        const response = await fetch(exportUrl);
        if (!response.ok) {
            throw new Error(`HTTP Error: ${response.status}. Hãy chắc chắn Google Sheet đã bật chia sẻ 'Bất kỳ ai có liên kết'.`);
        }
        const arrayBuffer = await response.arrayBuffer();
        currentFileName = "GoogleSheet_Export.xlsx";
        processWorkbookData(arrayBuffer, currentFileName);
    } catch (err) {
        alert(`Không thể tải Google Sheet trực tiếp qua trình duyệt (do bảo mật CORS của Google).\n\n💡 Cách khắc phục đơn giản:\n1. Vào Google Sheet, chọn File > Tải xuống > Microsoft Excel (.xlsx)\n2. Sau đó kéo thả file tải về vào tab 'File Excel từ máy' là xong ngay!`);
    } finally {
        btnFetchGsheet.disabled = false;
        btnFetchGsheet.textContent = "📥 Tải Sheet";
    }
});

// Setup Form Events
function setupFormEvents() {
    ruleEachCase.addEventListener("change", () => onRuleChanged());
    ruleByGroup.addEventListener("change", () => onRuleChanged());

    sheetSelect.addEventListener("change", () => onSheetSelected());
    headerRowInput.addEventListener("change", () => onSheetSelected());
    groupColSelect.addEventListener("change", () => analyzeGroups());
    chkForwardFill.addEventListener("change", () => analyzeGroups());

    btnSelectAll.addEventListener("click", () => {
        selectedGroupNames = new Set(Object.keys(detectedGroups));
        updateTableCheckboxes();
    });

    btnDeselectAll.addEventListener("click", () => {
        selectedGroupNames.clear();
        updateTableCheckboxes();
    });

    btnGenerate.addEventListener("click", () => executeSheetGeneration());

    btnCloseModal.addEventListener("click", () => successModal.classList.add("hidden"));
    btnDownloadAgain.addEventListener("click", () => {
        if (lastGeneratedBlob) {
            triggerDownload(lastGeneratedBlob, outputFilenameInput.value.trim() || "TestCases_Tach_Sheet.xlsx");
        }
    });
}

function onRuleChanged() {
    if (ruleEachCase.checked) {
        cardRuleEachCase.classList.add("active");
        cardRuleByGroup.classList.remove("active");
        lblColSelect.textContent = "Cột định danh Test Case (dùng làm Tên Sheet):";
        groupColHint.textContent = "💡 Tool sẽ sinh mỗi testcase thành 1 sheet mới riêng biệt với template giống hệt file tổng.";
    } else {
        cardRuleByGroup.classList.add("active");
        cardRuleEachCase.classList.remove("active");
        lblColSelect.textContent = "Cột dùng để phân nhóm gom testcase:";
        groupColHint.textContent = "💡 Gom các testcase cùng nhóm (Module, Feature...) vào chung 1 sheet.";
    }
    onSheetSelected();
}

// ================= FILE PROCESSING =================
function handleFileSelected(file) {
    currentFileName = file.name;
    const reader = new FileReader();

    reader.onload = (e) => {
        processWorkbookData(e.target.result, file.name);
    };

    reader.readAsArrayBuffer(file);
}

function processWorkbookData(arrayBuffer, fileName) {
    try {
        currentWorkbook = XLSX.read(arrayBuffer, { type: "array", cellStyles: true, cellNF: true });
        
        dropZone.classList.add("hidden");
        fileInfo.classList.remove("hidden");
        fileNameDisplay.textContent = fileName;

        // Cập nhật tên file xuất gợi ý
        const base = fileName.substring(0, fileName.lastIndexOf(".")) || fileName;
        outputFilenameInput.value = `${base}_Tach_Sheet.xlsx`;

        // Đổ danh sách sheet vào combobox
        sheetSelect.innerHTML = "";
        currentWorkbook.SheetNames.forEach(sName => {
            const opt = document.createElement("option");
            opt.value = sName;
            opt.textContent = sName;
            sheetSelect.appendChild(opt);
        });
        sheetSelect.disabled = false;

        onSheetSelected();

    } catch (err) {
        alert("Lỗi khi đọc file Excel: " + err.message);
    }
}

function onSheetSelected() {
    if (!currentWorkbook) return;
    const sName = sheetSelect.value;
    if (!sName) return;

    currentSheetName = sName;
    const worksheet = currentWorkbook.Sheets[sName];
    
    // Đọc dạng mảng 2 chiều (header: 1)
    currentSheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    const headerIdx = parseInt(headerRowInput.value || "1", 10) - 1;
    if (headerIdx < 0 || headerIdx >= currentSheetData.length) {
        alert("Dòng tiêu đề vượt quá số dòng dữ liệu trong sheet.");
        return;
    }

    const headerRow = currentSheetData[headerIdx] || [];
    groupColSelect.innerHTML = "";

    const isEachCase = ruleEachCase.checked;
    let bestColIdx = 0;
    
    // Nếu tách từng case: ưu tiên cột mã testcase con, mã tc, id, testcase id
    // Nếu gom theo module: ưu tiên module, phân hệ, feature, cha
    const caseKeywords = ["con", "mã test case", "mã tc", "sub-id", "sub id", "test case id", "mã", "id", "tên test case"];
    const groupKeywords = ["module", "phân hệ", "chức năng", "feature", "cha", "suite", "nhóm", "category"];
    const targetKeywords = isEachCase ? caseKeywords : groupKeywords;

    headerRow.forEach((colName, idx) => {
        const colLetter = XLSX.utils.encode_col(idx);
        const displayName = colName ? String(colName).trim() : `Cột ${colLetter}`;
        
        const opt = document.createElement("option");
        opt.value = idx;
        opt.textContent = `Cột ${colLetter}: ${displayName}`;
        groupColSelect.appendChild(opt);

        const lower = displayName.toLowerCase();
        if (targetKeywords.some(k => lower.includes(k))) {
            bestColIdx = idx;
        }
    });

    groupColSelect.disabled = false;
    if (headerRow.length > 0) {
        groupColSelect.value = bestColIdx;
    }

    analyzeGroups();
}

// ================= ANALYZE GROUPS =================
function sanitizeSheetName(name, existingNames = []) {
    if (!name || String(name).trim() === "") {
        name = "Chua_Phan_Nhom";
    }
    // Loại bỏ ký tự cấm: \ / ? * [ ] :
    let clean = String(name).trim().replace(/[\\/*?:\[\]]/g, "_").replace(/^'+|'+$/g, "");
    if (!clean) clean = "Sheet";
    clean = clean.substring(0, 31);

    const existingLower = existingNames.map(n => n.toLowerCase());
    if (!existingLower.includes(clean.toLowerCase())) {
        return clean;
    }

    let counter = 1;
    while (true) {
        const suffix = `_${counter}`;
        const avail = 31 - suffix.length;
        const candidate = `${clean.substring(0, avail)}${suffix}`;
        if (!existingLower.includes(candidate.toLowerCase())) {
            return candidate;
        }
        counter++;
    }
}

function setupPreviewTabSwitching() {
    const tabData = document.getElementById("p-tab-data");
    const tabSheets = document.getElementById("p-tab-sheets");
    const tabSingle = document.getElementById("p-tab-single");

    const paneData = document.getElementById("pane-tab-data");
    const paneSheets = document.getElementById("pane-tab-sheets");
    const paneSingle = document.getElementById("pane-tab-single");

    const searchInput = document.getElementById("data-search-input");
    const selectMockTc = document.getElementById("select-mock-tc");

    if (tabData && tabSheets && tabSingle) {
        tabData.addEventListener("click", () => {
            tabData.classList.add("active");
            tabSheets.classList.remove("active");
            tabSingle.classList.remove("active");
            paneData.classList.remove("hidden");
            paneSheets.classList.add("hidden");
            paneSingle.classList.add("hidden");
        });

        tabSheets.addEventListener("click", () => {
            tabSheets.classList.add("active");
            tabData.classList.remove("active");
            tabSingle.classList.remove("active");
            paneSheets.classList.remove("hidden");
            paneData.classList.add("hidden");
            paneSingle.classList.add("hidden");
        });

        tabSingle.addEventListener("click", () => {
            tabSingle.classList.add("active");
            tabData.classList.remove("active");
            tabSheets.classList.remove("active");
            paneSingle.classList.remove("hidden");
            paneData.classList.add("hidden");
            paneSheets.classList.add("hidden");
            renderSingleSheetMockup();
        });
    }

    if (searchInput) {
        searchInput.addEventListener("input", () => {
            filterFullDataTable(searchInput.value.trim().toLowerCase());
        });
    }

    if (selectMockTc) {
        selectMockTc.addEventListener("change", () => {
            renderSingleSheetMockup();
        });
    }
}

function analyzeGroups() {
    if (!currentSheetData || currentSheetData.length === 0) return;

    const headerIdx = parseInt(headerRowInput.value || "1", 10) - 1;
    const groupColIdx = parseInt(groupColSelect.value || "0", 10);
    const forwardFill = chkForwardFill.checked;

    detectedGroups = {};
    const existingNames = ["Tổng quan", "Overview", "Index"];
    let lastGroup = "";

    for (let r = headerIdx + 1; r < currentSheetData.length; r++) {
        const row = currentSheetData[r];
        if (!row || row.every(val => val === "" || val === null || val === undefined)) {
            continue; // Bỏ qua dòng trống
        }

        let rawVal = row[groupColIdx];
        let valClean = rawVal !== undefined && rawVal !== null ? String(rawVal).trim() : "";
        const lowerVal = valClean.toLowerCase();

        // Bỏ qua nếu là lặp lại dòng header
        if (["id", "mã test case", "stt", "mã tc", "testcase id", "tên test case", "mục đích kiểm thử"].includes(lowerVal)) {
            continue;
        }

        let currentGroup = "";
        if (valClean !== "") {
            currentGroup = valClean;
            lastGroup = currentGroup;
        } else if (forwardFill && lastGroup !== "") {
            currentGroup = lastGroup;
        } else {
            // Bỏ qua các dòng banner trước khi có test case đầu tiên
            continue;
        }

        if (!detectedGroups[currentGroup]) {
            const clean = sanitizeSheetName(currentGroup, existingNames);
            existingNames.push(clean);
            detectedGroups[currentGroup] = {
                groupName: currentGroup,
                sheetName: clean,
                count: 0,
                rows: []
            };
        }

        detectedGroups[currentGroup].count++;
        detectedGroups[currentGroup].rows.push(row);
    }

    selectedGroupNames = new Set(Object.keys(detectedGroups));
    renderGroupsTable();
    renderFullDataTable();
    updateMockTcSelect();
}

function renderFullDataTable() {
    const head = document.getElementById("full-table-head");
    const body = document.getElementById("full-table-body");
    const stats = document.getElementById("data-table-stats");
    if (!head || !body) return;

    if (!currentSheetData || currentSheetData.length === 0) {
        body.innerHTML = '<tr><td class="empty-state">Chưa có dữ liệu</td></tr>';
        return;
    }

    const headerIdx = parseInt(headerRowInput.value || "1", 10) - 1;
    const headerRow = currentSheetData[headerIdx] || [];

    // Header đa tầng
    let parentHeaderRow = headerIdx > 0 ? (currentSheetData[0] || []) : null;

    let headHtml = "<tr>";
    headerRow.forEach((colName, cIdx) => {
        const colLetter = XLSX.utils.encode_col(cIdx);
        let name = colName ? String(colName).trim() : `Cột ${colLetter}`;
        if (parentHeaderRow && parentHeaderRow[cIdx] && !name.includes(parentHeaderRow[cIdx])) {
            name = `${parentHeaderRow[cIdx]} - ${name}`;
        }
        headHtml += `<th title="${name}">${colLetter}: ${name}</th>`;
    });
    headHtml += "</tr>";
    head.innerHTML = headHtml;

    // Render Data Rows
    let bodyHtml = "";
    let dataRowCount = 0;
    for (let r = headerIdx + 1; r < currentSheetData.length; r++) {
        const row = currentSheetData[r];
        if (!row || row.every(val => val === "" || val === null || val === undefined)) continue;

        dataRowCount++;
        bodyHtml += `<tr class="data-row-item" data-text="${row.join(' ').toLowerCase()}">`;
        headerRow.forEach((_, cIdx) => {
            const cellVal = (row[cIdx] !== undefined && row[cIdx] !== null) ? String(row[cIdx]).trim() : "";
            bodyHtml += `<td title="${cellVal}">${cellVal || '<span style="color:#64748B">-</span>'}</td>`;
        });
        bodyHtml += "</tr>";
    }

    body.innerHTML = bodyHtml || '<tr><td class="empty-state">Không có dòng dữ liệu nào</td></tr>';
    if (stats) {
        stats.textContent = `📊 Hiển thị: ${dataRowCount} dòng dữ liệu | ${headerRow.length} cột chuẩn QA`;
    }
}

function filterFullDataTable(query) {
    const rows = document.querySelectorAll(".data-row-item");
    let visibleCount = 0;
    rows.forEach(r => {
        const text = r.getAttribute("data-text") || "";
        if (!query || text.includes(query)) {
            r.style.display = "";
            visibleCount++;
        } else {
            r.style.display = "none";
        }
    });
    const stats = document.getElementById("data-table-stats");
    if (stats) {
        stats.textContent = `📊 Tìm thấy: ${visibleCount} dòng phù hợp với '${query}'`;
    }
}

function updateMockTcSelect() {
    const sel = document.getElementById("select-mock-tc");
    if (!sel) return;
    sel.innerHTML = "";
    const keys = Object.keys(detectedGroups);
    keys.forEach(k => {
        const opt = document.createElement("option");
        opt.value = k;
        opt.textContent = `${k} (${detectedGroups[k].sheetName})`;
        sel.appendChild(opt);
    });
    renderSingleSheetMockup();
}

function renderSingleSheetMockup() {
    const sel = document.getElementById("select-mock-tc");
    const titleBadge = document.getElementById("mockup-tab-title");
    const content = document.getElementById("mockup-sheet-content");
    if (!content || !sel) return;

    const selectedCase = sel.value || Object.keys(detectedGroups)[0];
    if (!selectedCase || !detectedGroups[selectedCase]) {
        content.innerHTML = "<p>Vui lòng chọn test case để xem trước</p>";
        return;
    }

    const groupInfo = detectedGroups[selectedCase];
    if (titleBadge) {
        titleBadge.textContent = `📄 Sheet: ${groupInfo.sheetName}`;
    }

    const headerIdx = parseInt(headerRowInput.value || "1", 10) - 1;
    const headerRows = currentSheetData.slice(0, headerIdx + 1);

    let html = `
        <div class="mockup-header-cyan">
            [TIÊU ĐỀ 2 TẦNG XANH CYAN (#00B0F0) - 20 CỘT ĐẦY ĐỦ]
        </div>
        <div class="mockup-banner-yellow">
            Job quét và tổng hợp danh sách nợ (Merged A đến T)
        </div>
    `;

    groupInfo.rows.forEach((r, idx) => {
        const tcId = r[0] || selectedCase;
        const purpose = r[1] || "";
        const steps = r[2] || "";
        const expected = r[3] || "";
        html += `
            <div class="mockup-data-row">
                <strong>ID:</strong> <span style="color: #0284C7; font-weight: bold;">${tcId}</span> | 
                <strong>Mục đích:</strong> ${purpose}<br>
                <strong>Các bước thực hiện:</strong> ${steps}<br>
                <strong>Kết quả mong muốn:</strong> ${expected}
            </div>
        `;
    });

    content.innerHTML = html;
}

function renderGroupsTable() {
    groupsTableBody.innerHTML = "";
    const groupKeys = Object.keys(detectedGroups);

    if (groupKeys.length === 0) {
        renderEmptyTable();
        btnGenerate.disabled = true;
        previewSummaryBadge.textContent = "Không tìm thấy nhóm";
        return;
    }

    let totalTc = 0;
    groupKeys.forEach((key, idx) => {
        const item = detectedGroups[key];
        totalTc += item.count;

        const tr = document.createElement("tr");
        tr.id = `row-group-${idx}`;
        tr.innerHTML = `
            <td style="text-align: center;">
                <input type="checkbox" class="group-chk" data-group="${key}" ${selectedGroupNames.has(key) ? "checked" : ""}>
            </td>
            <td style="text-align: center; color: var(--text-muted);">${idx + 1}</td>
            <td style="font-weight: 600;">${item.groupName}</td>
            <td style="color: var(--color-accent); font-family: monospace;">📑 ${item.sheetName}</td>
            <td style="text-align: center; font-weight: 700; color: #FFFFFF;">${item.count}</td>
            <td style="text-align: center; color: #34D399; font-weight: 600;">Sẵn sàng</td>
        `;

        tr.addEventListener("click", (e) => {
            if (e.target.tagName !== "INPUT") {
                const chk = tr.querySelector(".group-chk");
                chk.checked = !chk.checked;
                toggleGroup(key, chk.checked);
            }
        });

        const chk = tr.querySelector(".group-chk");
        chk.addEventListener("change", (e) => {
            toggleGroup(key, e.target.checked);
        });

        groupsTableBody.appendChild(tr);
    });

    btnGenerate.disabled = false;
    updateBadge();
}

function toggleGroup(key, isChecked) {
    if (isChecked) {
        selectedGroupNames.add(key);
    } else {
        selectedGroupNames.remove(key);
    }
    updateBadge();
}

function updateTableCheckboxes() {
    const checkboxes = document.querySelectorAll(".group-chk");
    checkboxes.forEach(chk => {
        const key = chk.getAttribute("data-group");
        chk.checked = selectedGroupNames.has(key);
    });
    updateBadge();
}

function updateBadge() {
    const totalGroups = Object.keys(detectedGroups).length;
    const selCount = selectedGroupNames.size;
    let selTc = 0;
    selectedGroupNames.forEach(k => {
        if (detectedGroups[k]) selTc += detectedGroups[k].count;
    });

    previewSummaryBadge.textContent = `Đã chọn: ${selCount}/${totalGroups} nhóm | ${selTc} testcase`;
    btnGenerate.disabled = (selCount === 0);
}

function renderEmptyTable() {
    groupsTableBody.innerHTML = `
        <tr>
            <td colspan="5" class="empty-state">
                <div class="empty-box">
                    <div class="empty-icon">📊</div>
                    <p>Vui lòng chọn file Excel hoặc dán link Google Sheet để bắt đầu quét nhóm</p>
                </div>
            </td>
        </tr>
    `;
    previewSummaryBadge.textContent = "Chưa có dữ liệu";
}

// ================= GENERATE SHEETS =================
async function executeSheetGeneration() {
    if (!currentWorkbook || selectedGroupNames.size === 0) return;

    btnGenerate.disabled = true;
    progressBox.classList.remove("hidden");
    setProgress(10, "Đang khởi tạo workbook mới...");

    await sleep(80);

    try {
        const newWb = XLSX.utils.book_new();
        const headerIdx = parseInt(headerRowInput.value || "1", 10) - 1;
        const headerRows = currentSheetData.slice(0, headerIdx + 1);

        const includeOverview = chkOverview.checked;
        const addBackLink = chkBackLink.checked;
        const groupKeys = Array.from(selectedGroupNames);
        const totalGroups = groupKeys.length;
        let totalTcExported = 0;

        // 1. Tạo sheet Tổng quan trước (nếu bật)
        if (includeOverview) {
            setProgress(25, "Đang tạo trang Tổng quan...");
            const overviewData = [
                ["", "", "", "", ""],
                ["", "BẢNG TỔNG QUAN TEST CASES THEO MODULE / NHÓM", "", "", ""],
                ["", `Thời gian tạo: ${new Date().toLocaleString()} | Tổng số nhóm: ${totalGroups}`, "", "", ""],
                ["", "", "", "", ""],
                ["", "STT", "Tên Phân Nhóm / Module", "Tên Sheet Chi Tiết", "Số Test Case Con"]
            ];

            let tcSum = 0;
            groupKeys.forEach(k => tcSum += detectedGroups[k].count);

            groupKeys.forEach((key, stt) => {
                const item = detectedGroups[key];
                overviewData.push([
                    "",
                    stt + 1,
                    item.groupName,
                    item.sheetName,
                    item.count
                ]);
            });

            // Tổng cộng
            overviewData.push(["", "TỔNG CỘNG", "", "", tcSum]);

            const wsOverview = XLSX.utils.aoa_to_sheet(overviewData);
            // Thiết lập độ rộng cột
            wsOverview["!cols"] = [
                { wch: 3 },
                { wch: 8 },
                { wch: 38 },
                { wch: 35 },
                { wch: 18 }
            ];

            XLSX.utils.book_append_sheet(newWb, wsOverview, "Tổng quan");
        }

        // 2. Tạo từng sheet con
        for (let i = 0; i < groupKeys.length; i++) {
            const key = groupKeys[i];
            const item = detectedGroups[key];
            const pct = 30 + Math.round((i / totalGroups) * 60);
            setProgress(pct, `Đang tạo sheet [${i + 1}/${totalGroups}]: ${item.sheetName}...`);
            await sleep(30);

            const sheetRows = [];

            // Dòng Header (Bắt đầu ngay từ dòng 1)
            headerRows.forEach(hRow => {
                sheetRows.push([...hRow]);
            });

            // Các dòng test case con
            item.rows.forEach(r => {
                sheetRows.push([...r]);
                totalTcExported++;
            });

            const wsChild = XLSX.utils.aoa_to_sheet(sheetRows);

            // Giữ nguyên 100% độ rộng cột từ template nguồn
            const srcSheet = currentWorkbook.Sheets[currentSheetName];
            if (srcSheet && srcSheet["!cols"]) {
                wsChild["!cols"] = srcSheet["!cols"];
            } else if (sheetRows.length > 0) {
                const maxCols = Math.max(...sheetRows.map(r => r.length));
                const colWidths = [];
                for (let c = 0; c < maxCols; c++) {
                    colWidths.push({ wch: 22 });
                }
                colWidths[0] = { wch: 14 };
                wsChild["!cols"] = colWidths;
            }

            // Tái tạo các ô gộp (Merged Cells) cho phần Header template trên sheet con (offset = 0)
            if (srcSheet && srcSheet["!merges"]) {
                const childMerges = [];
                srcSheet["!merges"].forEach(m => {
                    if (m.e.r <= headerIdx) {
                        childMerges.push({
                            s: { r: m.s.r, c: m.s.c },
                            e: { r: m.e.r, c: m.e.c }
                        });
                    }
                });
                if (childMerges.length > 0) {
                    wsChild["!merges"] = childMerges;
                }
            }

            XLSX.utils.book_append_sheet(newWb, wsChild, item.sheetName);
        }

        setProgress(95, "Đang đóng gói và xuất file...");
        await sleep(50);

        // Xuất file
        const outFileName = outputFilenameInput.value.trim() || "TestCases_Tach_Sheet.xlsx";
        XLSX.writeFile(newWb, outFileName);

        setProgress(100, "Hoàn tất thành công!");
        await sleep(150);

        // Hiển thị Modal
        progressBox.classList.add("hidden");
        btnGenerate.disabled = false;

        statSheetsCount.textContent = totalGroups;
        statTcCount.textContent = totalTcExported;
        document.getElementById("modal-success-msg").textContent = `File '${outFileName}' đã được tải về thành công với ${totalGroups} sheet con (giữ nguyên 100% template và màu sắc)!`;
        successModal.classList.remove("hidden");

    } catch (err) {
        alert("Lỗi khi sinh sheet: " + err.message);
        btnGenerate.disabled = false;
        progressBox.classList.add("hidden");
    }
}

function setProgress(percent, text) {
    progressFill.style.width = `${percent}%`;
    progressText.textContent = text;
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// ================= DÙNG THỬ FILE MẪU CHUẨN 100% =================
function setupSampleLoader() {
    btnLoadSample.addEventListener("click", () => {
        // Template 2 tầng màu xanh Cyan chuẩn 100% theo ảnh của bạn
        const row1 = [
            "ID", "Mục đích kiểm thử", "Các bước thực hiện", "Kết quả mong muốn",
            "thực tế (UAT)", "thực tế (Prod)",
            "Data test", "", "",
            "UAT", "", "",
            "PROD", "", "",
            "Kết quả hiện tại", "(ID file log bug/ Jira)", "Ghi chú", "Ngày test UAT", "Ngày test PROD"
        ];

        const row2 = [
            "", "", "", "",
            "", "",
            "Lần 1", "Lần 2", "Lần 3",
            "Lần 1", "Lần 2", "Lần 3",
            "Lần 1", "Lần 2", "Lần 3",
            "", "", "", "", ""
        ];

        const row3 = [
            "", "Job quét và tổng hợp danh sách nợ", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
        ];

        const sampleRows = [
            [
                "[Tên_TC]_1",
                "Xác minh hệ thống quét bản ghi thỏa mãn điều kiện [STATUS != 1, FAILURE_REASON = 1, IS_SCAN > 0]",
                "1. Bảng SETTLEMENT.PRE_TRANS có bản ghi với STATUS=2, FAILURE_REASON=1, IS_SCAN=1",
                "2. - Bản ghi được chọn để xử lý (được đánh dấu hoặc chuyển trạng thái theo quy trình nội bộ)",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            ],
            [
                "[Tên_TC]_2",
                "Xác minh hệ thống bỏ qua bản ghi không thỏa mãn điều kiện quét [STATUS = 1]",
                "1. Bảng SETTLEMENT.PRE_TRANS có bản ghi với STATUS=1, FAILURE_REASON=1, IS_SCAN=1",
                "2. - Bản ghi bị bỏ qua, không được chọn đưa vào danh sách xử lý nợ",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            ],
            [
                "[Tên_TC]_3",
                "Xác minh hệ thống bỏ qua bản ghi có FAILURE_REASON khác 1",
                "1. Bảng SETTLEMENT.PRE_TRANS có bản ghi với STATUS=2, FAILURE_REASON=0, IS_SCAN=1",
                "2. - Hệ thống không quét bản ghi này",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            ],
            [
                "[Tên_TC]_4",
                "Xác minh hệ thống xử lý khi IS_SCAN = 0",
                "1. Bảng SETTLEMENT.PRE_TRANS có bản ghi với STATUS=2, FAILURE_REASON=1, IS_SCAN=0",
                "2. - Bản ghi không thỏa mãn điều kiện IS_SCAN > 0, bị bỏ qua",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            ],
            [
                "[Tên_TC]_5",
                "Kiểm tra hiệu năng khi bảng SETTLEMENT.PRE_TRANS có 100,000 bản ghi",
                "1. Chuẩn bị 100k bản ghi thỏa mãn điều kiện quét\n2. Kích hoạt Job quét danh sách nợ",
                "2. - Thời gian thực thi Job < 5 phút, không bị timeout hay khóa bảng",
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            ]
        ];

        const sampleWb = XLSX.utils.book_new();
        const sampleWs = XLSX.utils.aoa_to_sheet([row1, row2, row3, ...sampleRows]);

        // Cấu hình Merges cho Header 2 tầng chuẩn
        sampleWs["!merges"] = [
            { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },   // ID
            { s: { r: 0, c: 1 }, e: { r: 1, c: 1 } },   // Mục đích kiểm thử
            { s: { r: 0, c: 2 }, e: { r: 1, c: 2 } },   // Các bước thực hiện
            { s: { r: 0, c: 3 }, e: { r: 1, c: 3 } },   // Kết quả mong muốn
            { s: { r: 0, c: 4 }, e: { r: 1, c: 4 } },   // thực tế (UAT)
            { s: { r: 0, c: 5 }, e: { r: 1, c: 5 } },   // thực tế (Prod)
            { s: { r: 0, c: 6 }, e: { r: 0, c: 8 } },   // Data test (Lần 1, 2, 3)
            { s: { r: 0, c: 9 }, e: { r: 0, c: 11 } },  // UAT (Lần 1, 2, 3)
            { s: { r: 0, c: 12 }, e: { r: 0, c: 14 } }, // PROD (Lần 1, 2, 3)
            { s: { r: 0, c: 15 }, e: { r: 1, c: 15 } }, // Kết quả hiện tại
            { s: { r: 0, c: 16 }, e: { r: 1, c: 16 } }, // (ID file log bug/ Jira)
            { s: { r: 0, c: 17 }, e: { r: 1, c: 17 } }, // Ghi chú
            { s: { r: 0, c: 18 }, e: { r: 1, c: 18 } }, // Ngày test UAT
            { s: { r: 0, c: 19 }, e: { r: 1, c: 19 } }  // Ngày test PROD
        ];

        sampleWs["!cols"] = [
            { wch: 14 }, // ID
            { wch: 38 }, // Mục đích
            { wch: 42 }, // Các bước
            { wch: 42 }, // Kết quả mong muốn
            { wch: 14 }, { wch: 14 },
            { wch: 10 }, { wch: 10 }, { wch: 10 },
            { wch: 10 }, { wch: 10 }, { wch: 10 },
            { wch: 10 }, { wch: 10 }, { wch: 10 },
            { wch: 16 }, { wch: 22 }, { wch: 16 },
            { wch: 16 }, { wch: 16 }
        ];

        XLSX.utils.book_append_sheet(sampleWb, sampleWs, "Master_TestCases");

        const sampleBuffer = XLSX.write(sampleWb, { type: "array", bookType: "xlsx" });
        currentFileName = "sample_master_testcases.xlsx";
        
        // Thiết lập dòng header là dòng 3 (để lấy cả banner vàng) hoặc dòng 2
        headerRowInput.value = "3";
        processWorkbookData(sampleBuffer, currentFileName);
    });
}
