/**
 * Word Table Formatter for Thesis / Scientific Papers
 * Uses Office JavaScript API (Word.run)
 */

/* global Office, Word */

Office.onReady((info) => {
    initPaneZoom();
    initAddinReloadListener();

    try {
        if (Office.actions && typeof Office.actions.associate === "function") {
            Office.actions.associate("refreshAddin", refreshAddin);
        }
    } catch (e) {}

    if (info.host === Office.HostType.Word) {
        document.getElementById("btn-format-selected").onclick = formatSelectedTable;
        document.getElementById("btn-format-all").onclick = formatAllTables;
        document.getElementById("btn-format-quochieu").onclick = formatQuocHieuButton;

        // Xử lý nút tăng / giảm cỡ chữ
        const btnDec = document.getElementById("btn-decrease-font");
        const btnInc = document.getElementById("btn-increase-font");
        if (btnDec) btnDec.onclick = () => setFontSize(getSelectedFontSize() - 0.5);
        if (btnInc) btnInc.onclick = () => setFontSize(getSelectedFontSize() + 0.5);

        // Xử lý các chip gợi ý cỡ chữ (9pt, 10pt, 11pt, 12pt)
        document.querySelectorAll(".btn-chip").forEach(chip => {
            chip.onclick = () => {
                const size = parseFloat(chip.getAttribute("data-size"));
                setFontSize(size);
            };
        });

        // Xử lý khi người dùng nhập số trực tiếp vào ô input cỡ chữ
        const inputFont = document.getElementById("input-font-size");
        if (inputFont) {
            inputFont.oninput = (e) => {
                const val = parseFloat(e.target.value);
                if (!isNaN(val)) updatePresetActive(val);
            };
        }

        // Xử lý nút tăng / giảm 4 cạnh lề ô (Trên, Dưới, Trái, Phải)
        const bindMarginStepper = (decId, incId, inputId) => {
            const btnDec = document.getElementById(decId);
            const btnInc = document.getElementById(incId);
            const input = document.getElementById(inputId);
            if (btnDec) btnDec.onclick = () => setIndividualMargin(inputId, getMarginVal(inputId) - 0.05);
            if (btnInc) btnInc.onclick = () => setIndividualMargin(inputId, getMarginVal(inputId) + 0.05);
            if (input) {
                input.oninput = () => checkMarginPresets();
            }
        };

        bindMarginStepper("btn-dec-top", "btn-inc-top", "input-margin-top");
        bindMarginStepper("btn-dec-bottom", "btn-inc-bottom", "input-margin-bottom");
        bindMarginStepper("btn-dec-left", "btn-inc-left", "input-margin-left");
        bindMarginStepper("btn-dec-right", "btn-inc-right", "input-margin-right");

        // Xử lý các chip gợi ý lề ô
        document.querySelectorAll(".btn-chip-margin").forEach(chip => {
            chip.onclick = () => {
                const top = parseFloat(chip.getAttribute("data-top"));
                const bottom = parseFloat(chip.getAttribute("data-bottom"));
                const left = parseFloat(chip.getAttribute("data-left"));
                const right = parseFloat(chip.getAttribute("data-right"));
                setAllMargins(top, bottom, left, right);
            };
        });

        // Xử lý nút tăng / giảm / bỏ trống cho độ cao hàng & độ rộng bảng (tùy chọn)
        const bindOptionalDimStepper = (decId, incId, inputId, step, min, max, fallbackInit) => {
            const btnDec = document.getElementById(decId);
            const btnInc = document.getElementById(incId);
            const input = document.getElementById(inputId);
            const adjust = (delta) => {
                if (!input) return;
                const raw = input.value.trim();
                const current = raw === "" ? fallbackInit : parseFloat(raw);
                const next = isNaN(current) ? fallbackInit : (raw === "" ? fallbackInit : current + delta);
                const clamped = Math.max(min, Math.min(max, Math.round(next * 100) / 100));
                input.value = clamped;
            };
            if (btnDec) btnDec.onclick = () => adjust(-step);
            if (btnInc) btnInc.onclick = () => adjust(step);
        };

        bindOptionalDimStepper("btn-dec-height", "btn-inc-height", "input-row-height", 0.1, 0.1, 20, 0.7);
        bindOptionalDimStepper("btn-dec-width", "btn-inc-width", "input-table-width", 0.5, 1, 30, 16.5);

        const btnClearDims = document.getElementById("btn-clear-dimensions");
        if (btnClearDims) {
            btnClearDims.onclick = () => {
                const h = document.getElementById("input-row-height");
                const w = document.getElementById("input-table-width");
                if (h) h.value = "";
                if (w) w.value = "";
            };
        }

        // Xử lý các tùy chọn Format Paragraph
        initParagraphControls();
    }
});

/**
 * Lấy cỡ chữ người dùng đã chọn (mặc định 11pt)
 * @returns {number}
 */
function getSelectedFontSize() {
    const input = document.getElementById("input-font-size");
    const val = parseFloat(input ? input.value : 11);
    return (isNaN(val) || val <= 0) ? 11 : val;
}

/**
 * Cập nhật trạng thái active cho các nút chip gợi ý cỡ chữ
 * @param {number} size
 */
function updatePresetActive(size) {
    document.querySelectorAll(".btn-chip").forEach(chip => {
        const chipSize = parseFloat(chip.getAttribute("data-size"));
        if (chipSize === size) {
            chip.classList.add("active");
        } else {
            chip.classList.remove("active");
        }
    });
}

/**
 * Đặt cỡ chữ mới cho ô input và cập nhật giao diện
 * @param {number} newSize
 */
function setFontSize(newSize) {
    newSize = Math.max(6, Math.min(16, Math.round(newSize * 2) / 2));
    const input = document.getElementById("input-font-size");
    if (input) input.value = newSize;
    updatePresetActive(newSize);
}

/**
 * Lấy giá trị lề ô từ 1 ô input (đơn vị cm)
 * @param {string} id
 * @param {number} defaultVal
 * @returns {number}
 */
function getMarginVal(id, defaultVal = 0.2) {
    const input = document.getElementById(id);
    const val = parseFloat(input ? input.value : defaultVal);
    return (isNaN(val) || val < 0) ? defaultVal : val;
}

/**
 * Lấy đầy đủ 4 cạnh lề ô (Cell margins) theo cm
 * @returns {{top: number, bottom: number, left: number, right: number}}
 */
function getSelectedMargins() {
    return {
        top: getMarginVal("input-margin-top", 0.2),
        bottom: getMarginVal("input-margin-bottom", 0.2),
        left: getMarginVal("input-margin-left", 0.2),
        right: getMarginVal("input-margin-right", 0.2),
    };
}

/**
 * Cập nhật trạng thái active cho các nút chip gợi ý lề ô
 */
function checkMarginPresets() {
    const cur = getSelectedMargins();
    document.querySelectorAll(".btn-chip-margin").forEach(chip => {
        const t = parseFloat(chip.getAttribute("data-top"));
        const b = parseFloat(chip.getAttribute("data-bottom"));
        const l = parseFloat(chip.getAttribute("data-left"));
        const r = parseFloat(chip.getAttribute("data-right"));
        if (
            Math.abs(cur.top - t) < 0.01 &&
            Math.abs(cur.bottom - b) < 0.01 &&
            Math.abs(cur.left - l) < 0.01 &&
            Math.abs(cur.right - r) < 0.01
        ) {
            chip.classList.add("active");
        } else {
            chip.classList.remove("active");
        }
    });
}

/**
 * Đặt giá trị lề ô cho 1 cạnh cụ thể
 * @param {string} id
 * @param {number} newVal
 */
function setIndividualMargin(id, newVal) {
    newVal = Math.max(0, Math.min(2, Math.round(newVal * 100) / 100));
    const input = document.getElementById(id);
    if (input) input.value = newVal;
    checkMarginPresets();
}

/**
 * Đặt đồng thời cả 4 cạnh lề ô
 * @param {number} top
 * @param {number} bottom
 * @param {number} left
 * @param {number} right
 */
function setAllMargins(top, bottom, left, right) {
    const setVal = (id, v) => {
        const input = document.getElementById(id);
        if (input) input.value = Math.max(0, Math.min(2, Math.round(v * 100) / 100));
    };
    setVal("input-margin-top", top);
    setVal("input-margin-bottom", bottom);
    setVal("input-margin-left", left);
    setVal("input-margin-right", right);
    checkMarginPresets();
}

/**
 * Trả về chuỗi mô tả ngắn gọn về 4 cạnh lề ô
 * @returns {string}
 */
function getMarginsDescription() {
    const m = getSelectedMargins();
    if (m.top === m.bottom && m.bottom === m.left && m.left === m.right) {
        return `${m.top}cm`;
    }
    return `T:${m.top}/D:${m.bottom}/L:${m.left}/R:${m.right}cm`;
}

/**
 * Lấy giá trị độ cao hàng & độ rộng bảng tùy chọn (đơn vị cm).
 * Trả về null nếu người dùng bỏ trống (không nhập).
 * @returns {{rowHeightCm: number | null, tableWidthCm: number | null}}
 */
function getOptionalDimensions() {
    const parseOptional = (id) => {
        const input = document.getElementById(id);
        if (!input) return null;
        const raw = input.value.trim();
        if (raw === "") return null;
        const val = parseFloat(raw);
        return (!isNaN(val) && val > 0) ? val : null;
    };

    return {
        rowHeightCm: parseOptional("input-row-height"),
        tableWidthCm: parseOptional("input-table-width"),
    };
}

/**
 * Khởi tạo các sự kiện cho thẻ Tùy chỉnh Đoạn văn (Paragraph)
 */
function initParagraphControls() {
    const bindStepper = (decId, incId, inputId, step, min, max, decimals = 2) => {
        const btnDec = document.getElementById(decId);
        const btnInc = document.getElementById(incId);
        const input = document.getElementById(inputId);
        const factor = Math.pow(10, decimals);
        const adjust = (delta) => {
            if (!input) return;
            const cur = parseFloat(input.value);
            const base = isNaN(cur) ? 0 : cur;
            const next = Math.max(min, Math.min(max, Math.round((base + delta) * factor) / factor));
            input.value = next;
            syncParagraphChips();
        };
        if (btnDec) btnDec.onclick = () => adjust(-step);
        if (btnInc) btnInc.onclick = () => adjust(step);
        if (input) input.oninput = syncParagraphChips;
    };

    bindStepper("btn-dec-line-spacing", "btn-inc-line-spacing", "input-para-line-spacing", 0.05, 0.5, 5, 2);
    bindStepper("btn-dec-space-before", "btn-inc-space-before", "input-para-space-before", 1, 0, 100, 1);
    bindStepper("btn-dec-space-after", "btn-inc-space-after", "input-para-space-after", 1, 0, 100, 1);
    bindStepper("btn-dec-indent-first", "btn-inc-indent-first", "input-para-indent-first", 0.1, -5, 10, 2);
    bindStepper("btn-dec-indent-left", "btn-inc-indent-left", "input-para-indent-left", 0.1, 0, 10, 2);
    bindStepper("btn-dec-indent-right", "btn-inc-indent-right", "input-para-indent-right", 0.1, 0, 10, 2);

    document.querySelectorAll(".btn-chip-align").forEach(chip => {
        chip.onclick = () => {
            document.querySelectorAll(".btn-chip-align").forEach(c => c.classList.remove("active"));
            chip.classList.add("active");
        };
    });

    document.querySelectorAll(".btn-chip-ls").forEach(chip => {
        chip.onclick = () => {
            const input = document.getElementById("input-para-line-spacing");
            if (input) input.value = parseFloat(chip.getAttribute("data-ls"));
            syncParagraphChips();
        };
    });

    document.querySelectorAll(".btn-chip-spacing").forEach(chip => {
        chip.onclick = () => {
            const beforeIn = document.getElementById("input-para-space-before");
            const afterIn = document.getElementById("input-para-space-after");
            if (beforeIn) beforeIn.value = parseFloat(chip.getAttribute("data-before"));
            if (afterIn) afterIn.value = parseFloat(chip.getAttribute("data-after"));
            syncParagraphChips();
        };
    });

    document.querySelectorAll(".btn-chip-indent").forEach(chip => {
        chip.onclick = () => {
            const firstIn = document.getElementById("input-para-indent-first");
            const leftIn = document.getElementById("input-para-indent-left");
            const rightIn = document.getElementById("input-para-indent-right");
            if (firstIn) firstIn.value = parseFloat(chip.getAttribute("data-first"));
            if (leftIn) leftIn.value = parseFloat(chip.getAttribute("data-left"));
            if (rightIn) rightIn.value = parseFloat(chip.getAttribute("data-right"));
            syncParagraphChips();
        };
    });

    const btnReset = document.getElementById("btn-para-preset-reset");
    if (btnReset) {
        btnReset.onclick = () => {
            document.querySelectorAll(".btn-chip-align").forEach(c => {
                c.classList.toggle("active", c.getAttribute("data-align") === "auto");
            });
            const setVal = (id, val) => {
                const el = document.getElementById(id);
                if (el) el.value = val;
            };
            setVal("input-para-line-spacing", 1);
            setVal("input-para-space-before", 0);
            setVal("input-para-space-after", 0);
            setVal("input-para-indent-first", 0);
            setVal("input-para-indent-left", 0);
            setVal("input-para-indent-right", 0);
            syncParagraphChips();
        };
    }
}

/**
 * Đồng bộ trạng thái active cho các nút gợi ý Paragraph
 */
function syncParagraphChips() {
    const opts = getParagraphOptions();
    document.querySelectorAll(".btn-chip-ls").forEach(c => {
        c.classList.toggle("active", Math.abs(parseFloat(c.getAttribute("data-ls")) - opts.lineSpacing) < 0.01);
    });
    document.querySelectorAll(".btn-chip-spacing").forEach(c => {
        const b = parseFloat(c.getAttribute("data-before"));
        const a = parseFloat(c.getAttribute("data-after"));
        c.classList.toggle("active", Math.abs(b - opts.spaceBefore) < 0.01 && Math.abs(a - opts.spaceAfter) < 0.01);
    });
    document.querySelectorAll(".btn-chip-indent").forEach(c => {
        const f = parseFloat(c.getAttribute("data-first"));
        const l = parseFloat(c.getAttribute("data-left"));
        const r = parseFloat(c.getAttribute("data-right"));
        c.classList.toggle(
            "active",
            Math.abs(f - opts.firstLineIndentCm) < 0.01 &&
            Math.abs(l - opts.leftIndentCm) < 0.01 &&
            Math.abs(r - opts.rightIndentCm) < 0.01
        );
    });
}

/**
 * Lấy toàn bộ cấu hình định dạng Đoạn văn (Paragraph) cho nội dung trong bảng
 */
function getParagraphOptions() {
    const numVal = (id, def, minVal = -Infinity) => {
        const el = document.getElementById(id);
        const v = parseFloat(el ? el.value : def);
        return (isNaN(v) || v < minVal) ? def : v;
    };
    const activeAlign = document.querySelector(".btn-chip-align.active");

    return {
        alignment: activeAlign ? activeAlign.getAttribute("data-align") : "auto",
        lineSpacing: numVal("input-para-line-spacing", 1, 0.5),
        spaceBefore: numVal("input-para-space-before", 0, 0),
        spaceAfter: numVal("input-para-space-after", 0, 0),
        firstLineIndentCm: numVal("input-para-indent-first", 0),
        leftIndentCm: numVal("input-para-indent-left", 0, 0),
        rightIndentCm: numVal("input-para-indent-right", 0, 0),
    };
}

/**
 * Hiển thị thông báo trạng thái trên Task Pane
 * @param {string} message - Nội dung thông báo
 * @param {"info" | "success" | "warning" | "error"} type - Loại thông báo
 */
function showStatus(message, type = "info") {
    const statusBox = document.getElementById("status-box");
    const statusIcon = document.getElementById("status-icon");
    const statusMessage = document.getElementById("status-message");

    const icons = {
        info: "⏳",
        success: "✅",
        warning: "⚠️",
        error: "❌"
    };

    statusBox.className = `status-box ${type}`;
    statusIcon.textContent = icons[type] || "ℹ️";
    statusMessage.textContent = message;
    statusBox.style.display = "flex";
}

/**
 * Bật/tắt trạng thái vô hiệu hóa các nút bấm trong khi xử lý
 * @param {boolean} disabled
 */
function setButtonsDisabled(disabled) {
    document.getElementById("btn-format-selected").disabled = disabled;
    document.getElementById("btn-format-all").disabled = disabled;
    const btnQH = document.getElementById("btn-format-quochieu");
    if (btnQH) btnQH.disabled = disabled;
}

/**
 * Chuẩn hóa chuỗi tiếng Việt (bỏ dấu, chuyển chữ thường) để đối chiếu chính xác
 * @param {string} str
 * @returns {string}
 */
function normalizeVietnamese(str) {
    if (!str) return "";
    return str
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/đ/g, "d")
        .replace(/Đ/g, "d")
        .replace(/\s+/g, " ")
        .trim();
}

/**
 * Kiểm tra xem bảng có phải là bảng tiêu đề / Quốc hiệu - Tiêu ngữ không
 * @param {string[][]} values
 * @returns {boolean}
 */
function isQuocHieuTable(values) {
    if (!values || !Array.isArray(values)) return false;
    const combined = values.flat().join(" ");
    const norm = normalizeVietnamese(combined);

    return norm.includes("cong hoa xa hoi chu nghia viet nam") ||
        (norm.includes("doc lap") && norm.includes("tu do") && norm.includes("hanh phuc"));
}

/**
 * Đo chiều rộng văn bản (đơn vị pt) để căn cột số liệu vừa đúng 1 dòng
 * @param {string} text
 * @param {number} fontSizePt
 * @param {boolean} isBold
 * @returns {number}
 */
function measureTextWidthPt(text, fontSizePt, isBold) {
    if (!text) return 0;
    if (!measureTextWidthPt._ctx) {
        const canvas = document.createElement("canvas");
        measureTextWidthPt._ctx = canvas.getContext("2d");
    }
    const ctx = measureTextWidthPt._ctx;
    if (!ctx) return text.length * fontSizePt * 0.55;
    ctx.font = `${isBold ? "bold " : ""}${fontSizePt}pt "Times New Roman", Times, serif`;
    return ctx.measureText(text).width * 0.75;
}

/**
 * 1. ĐỊNH DẠNG BẢNG DỮ LIỆU KHÓA LUẬN (Data Table)
 * - Cỡ chữ linh hoạt theo ô nhập (mặc định 11pt, có thể giảm xuống 10pt, 9pt...)
 * - Căn đều kích thước các cột số liệu (Cột 2 trở đi) để số không bị nhảy dòng
 * - Table Style: Table Grid
 * - Căn lề dọc ô: Center
 * - Hàng 1: In đậm, căn giữa
 * - Từ hàng 2: Cột 1 căn trái, Cột 2+ căn phải
 * - AutoFit Window
 * @param {Word.Table} table
 * @param {Word.RequestContext} context
 */
async function formatSingleTable(table, context) {
    const fontSize = getSelectedFontSize();
    const chkEqual = document.getElementById("chk-equal-numeric-cols");
    const shouldEqualCols = chkEqual ? chkEqual.checked : true;
    const dims = getOptionalDimensions();
    const PT_PER_CM = 72 / 2.54;

    // 1. Áp dụng Table Style: "Table Grid"
    try {
        table.styleBuiltIn = "TableGrid";
    } catch {
        try {
            table.style = "Table Grid";
        } catch {
            // Dự phòng
        }
    }

    // 2. Luôn chạy AutoFit Window TRƯỚC TIÊN rồi mới điều chỉnh các thông số còn lại sau
    let autoFitTableWidthPt = null;
    try {
        if (typeof table.autoFitWindow === "function") {
            table.autoFitWindow();
        } else if (typeof table.autoFitBehavior === "function") {
            table.autoFitBehavior("Window");
        }
        table.load("width");
    } catch (e) {
        console.warn("AutoFit Window không được hỗ trợ hoặc gặp lỗi:", e);
    }

    // Load các hàng (rows) của bảng đồng thời sync AutoFit Window
    const rows = table.rows;
    rows.load("items");
    await context.sync();

    try {
        if (typeof table.width === "number" && table.width > 0) {
            autoFitTableWidthPt = table.width;
        }
    } catch (e) {}

    if (rows.items.length === 0) {
        return;
    }

    // Quy chuẩn Cell Margin: lấy theo 4 ô tùy chỉnh riêng của người dùng (cm -> pt)
    const margins = getSelectedMargins();
    // 1 cm = 72 / 2.54 pt ≈ 28.3465 pt
    const topPt = margins.top * (72 / 2.54);
    const bottomPt = margins.bottom * (72 / 2.54);
    const leftPt = margins.left * (72 / 2.54);
    const rightPt = margins.right * (72 / 2.54);

    try {
        table.setCellPadding("Top", topPt);
        table.setCellPadding("Bottom", bottomPt);
        table.setCellPadding("Left", leftPt);
        table.setCellPadding("Right", rightPt);
    } catch (e) {}

    // Áp dụng độ cao hàng nếu người dùng có nhập giá trị (bỏ trống -> giữ nguyên)
    if (dims.rowHeightCm !== null) {
        const rowHeightPt = dims.rowHeightCm * PT_PER_CM;
        for (let r = 0; r < rows.items.length; r++) {
            try {
                rows.items[r].preferredHeight = rowHeightPt;
            } catch (e) {}
        }
    }

    // 3. Load các ô (cells) cho từng hàng (và columnWidth hàng 0 phòng trường hợp table.width chưa trả về số dương)
    for (let r = 0; r < rows.items.length; r++) {
        rows.items[r].cells.load("items");
    }
    await context.sync();

    if (autoFitTableWidthPt === null && rows.items[0].cells.items.length > 0) {
        rows.items[0].cells.items.forEach((cell) => cell.load("columnWidth"));
        await context.sync();
        const sumColWidths = rows.items[0].cells.items.reduce(
            (sum, cell) => sum + (typeof cell.columnWidth === "number" && cell.columnWidth > 0 ? cell.columnWidth : 0),
            0
        );
        if (sumColWidths > 0) {
            autoFitTableWidthPt = sumColWidths;
        }
    }

    // 4. Căn lề dọc ô (Center), đặt lề ô và load đoạn văn bản qua cell.body
    for (let r = 0; r < rows.items.length; r++) {
        const cells = rows.items[r].cells.items;
        for (let c = 0; c < cells.length; c++) {
            const cell = cells[c];
            cell.verticalAlignment = "Center";
            try {
                cell.setCellPadding("Top", topPt);
                cell.setCellPadding("Bottom", bottomPt);
                cell.setCellPadding("Left", leftPt);
                cell.setCellPadding("Right", rightPt);
            } catch (err) {}
            cell.body.paragraphs.load("items, text");
        }
    }
    await context.sync();

    // 5. Duyệt từng ô và áp dụng định dạng đoạn văn & font chữ
    const paraOpts = getParagraphOptions();
    for (let r = 0; r < rows.items.length; r++) {
        const isHeaderRow = (r === 0);
        const cells = rows.items[r].cells.items;

        for (let c = 0; c < cells.length; c++) {
            const cell = cells[c];
            const isFirstCol = (c === 0);
            const paragraphs = cell.body.paragraphs.items;

            for (let p = 0; p < paragraphs.length; p++) {
                const para = paragraphs[p];

                // Font chữ toàn bảng theo kích thước người dùng chọn
                para.font.name = "Times New Roman";
                para.font.size = fontSize;
                para.font.color = "#000000";

                // Áp dụng cấu hình Paragraph cho nội dung trong bảng
                para.spaceBefore = paraOpts.spaceBefore;
                para.spaceAfter = paraOpts.spaceAfter;
                para.lineSpacing = (paraOpts.lineSpacing === 1) ? fontSize : (paraOpts.lineSpacing * 12);

                para.firstLineIndent = paraOpts.firstLineIndentCm * PT_PER_CM;
                para.leftIndent = paraOpts.leftIndentCm * PT_PER_CM;
                para.rightIndent = paraOpts.rightIndentCm * PT_PER_CM;

                // Quy tắc cho từng hàng/cột:
                if (isHeaderRow) {
                    para.font.bold = true;
                    para.alignment = "Centered";
                } else {
                    para.font.bold = false;

                    if (paraOpts.alignment && paraOpts.alignment !== "auto") {
                        para.alignment = paraOpts.alignment;
                    } else if (isFirstCol) {
                        para.alignment = "Left";
                    } else {
                        para.alignment = "Right";
                    }
                }
            }
        }
    }

    // 6. Điều chỉnh độ rộng tùy chỉnh (nếu có) & căn chỉnh kích thước các cột
    const customWidthPt = (dims.tableWidthCm !== null) ? (dims.tableWidthCm * PT_PER_CM) : null;
    if (customWidthPt !== null) {
        try {
            table.width = customWidthPt;
        } catch (e) {
            console.warn("Lỗi khi đặt độ rộng bảng:", e);
        }
    }

    if (shouldEqualCols && rows.items.length > 0) {
        try {
            const colCount = rows.items[0].cells.items.length;

            if (colCount >= 2) {
                // Tổng chiều rộng bảng: ưu tiên độ rộng tùy chỉnh -> độ rộng AutoFit Window đã đo ở bước 2 -> chuẩn A4 (~468pt)
                const totalWidth = customWidthPt !== null
                    ? customWidthPt
                    : (autoFitTableWidthPt !== null ? autoFitTableWidthPt : 468);

                const cellExtraPt = leftPt + rightPt + (paraOpts.leftIndentCm + paraOpts.rightIndentCm) * PT_PER_CM + 1;
                const colRequiredWidths = [];

                // Xét các cột số liệu từ phải sang trái để đo độ rộng nhỏ nhất vừa đúng 1 hàng của ô
                for (let c = colCount - 1; c >= 1; c--) {
                    let maxColWidth = 12;
                    for (let r = 0; r < rows.items.length; r++) {
                        const rowCells = rows.items[r].cells.items;
                        if (c >= rowCells.length) continue;
                        const isHeader = (r === 0);
                        const paras = rowCells[c].body.paragraphs.items;
                        for (let p = 0; p < paras.length; p++) {
                            const lines = (paras[p].text || "").split(/[\r\n]+/);
                            for (let k = 0; k < lines.length; k++) {
                                const lineText = lines[k].trim();
                                if (!lineText) continue;
                                const textPt = measureTextWidthPt(lineText, fontSize, isHeader);
                                const firstIndentPt = isHeader ? 0 : Math.max(0, paraOpts.firstLineIndentCm) * PT_PER_CM;
                                const needed = textPt + cellExtraPt + firstIndentPt;
                                if (needed > maxColWidth) {
                                    maxColWidth = needed;
                                }
                            }
                        }
                    }
                    colRequiredWidths.push(maxColWidth);
                }

                // Distribute đều các cột số liệu theo độ rộng nhỏ nhất có thể mà vẫn vừa 1 hàng
                let numColWidth = Math.ceil(Math.max(...colRequiredWidths) * 2) / 2;
                const minLabelWidth = Math.max(72, Math.round(totalWidth * 0.2));
                if (totalWidth - numColWidth * (colCount - 1) < minLabelWidth) {
                    numColWidth = Math.floor(((totalWidth - minLabelWidth) / (colCount - 1)) * 2) / 2;
                }

                // Cột chỉ tiêu (cột đầu tiên) nhận toàn bộ phần chiều rộng còn lại
                const labelWidth = Math.max(minLabelWidth, Math.round((totalWidth - numColWidth * (colCount - 1)) * 10) / 10);

                // Gán columnWidth trực tiếp cho từng ô trong bảng (WordApi 1.3+ tương thích 100% Word for Mac & Windows)
                for (let r = 0; r < rows.items.length; r++) {
                    const rowCells = rows.items[r].cells.items;
                    if (rowCells.length !== colCount) continue;
                    for (let c = 0; c < rowCells.length; c++) {
                        rowCells[c].columnWidth = (c === 0) ? labelWidth : numColWidth;
                    }
                }
                await context.sync();
            }
        } catch (e) {
            console.warn("Lỗi khi điều chỉnh độ rộng cột:", e);
        }
    }
}

/**
 * 2. ĐỊNH DẠNG BẢNG QUỐC HIỆU - TIÊU NGỮ (Header Table)
 * - Xóa sạch mọi đường viền bảng (PlainTable1 + No border)
 * - Căn lề dọc ô: Top (để 2 bên cân đối từ dòng đầu tiên)
 * - Cả 2 cột: Căn giữa ngang (Centered)
 * - Cột 1 (Đơn vị/Trường): Times New Roman 12pt, Trường không đậm, Viện/Khoa in đậm
 * - Cột 2 (Quốc hiệu):
 *   + Dòng "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM": In hoa, in đậm, 12pt
 *   + Dòng "Độc lập – Tự do – Hạnh phúc": In đậm, 12pt
 *   + Dòng ngày tháng: In nghiêng, không đậm, 12pt
 * - AutoFit Window
 * @param {Word.Table} table
 * @param {Word.RequestContext} context
 */
async function formatQuocHieuTable(table, context) {
    // 1. Áp dụng Style bảng trơn không viền của Word (PlainTable1)
    try {
        table.styleBuiltIn = "PlainTable1";
    } catch {
        try {
            table.style = "Plain Table 1";
        } catch {}
    }

    // 2. Luôn AutoFit Window trước khi điều chỉnh các thông số còn lại
    try {
        if (typeof table.autoFitWindow === "function") {
            table.autoFitWindow();
        } else if (typeof table.autoFitBehavior === "function") {
            table.autoFitBehavior("Window");
        }
    } catch {}

    // 2. Xóa các đường viền của bảng (Dùng các vị trí an toàn)
    try {
        table.getBorder("All").type = "None";
    } catch {}

    const safeBorders = ["Top", "Bottom", "Left", "Right"];
    for (const b of safeBorders) {
        try {
            table.getBorder(b).type = "None";
        } catch {}
    }

    // 3. Load các hàng và các ô
    const rows = table.rows;
    rows.load("items");
    await context.sync();

    if (rows.items.length === 0) return;

    // Nếu bảng có nhiều hàng/cột thì mới xóa viền bên trong
    if (rows.items.length > 1) {
        try {
            table.getBorder("InsideHorizontal").type = "None";
        } catch {}
    }
    try {
        table.getBorder("InsideVertical").type = "None";
    } catch {}

    for (let r = 0; r < rows.items.length; r++) {
        rows.items[r].cells.load("items");
    }
    await context.sync();

    // 4. Căn lề dọc ô: Top & xóa sạch viền ở cấp độ từng ô (cell borders)
    for (let r = 0; r < rows.items.length; r++) {
        const cells = rows.items[r].cells.items;
        for (let c = 0; c < cells.length; c++) {
            const cell = cells[c];
            cell.verticalAlignment = "Top";

            const cellBorders = ["Top", "Bottom", "Left", "Right"];
            for (const cb of cellBorders) {
                try {
                    cell.getBorder(cb).type = "None";
                } catch {}
            }

            cell.body.paragraphs.load("items, text");
        }
    }
    await context.sync();

    // 5. Áp dụng định dạng chi tiết từng ô
    for (let r = 0; r < rows.items.length; r++) {
        const cells = rows.items[r].cells.items;

        for (let c = 0; c < cells.length; c++) {
            const cell = cells[c];
            const isRightColumn = (c >= 1);
            const paragraphs = cell.body.paragraphs.items;

            for (let p = 0; p < paragraphs.length; p++) {
                const para = paragraphs[p];
                const rawText = (para.text || "").trim();
                const normText = normalizeVietnamese(rawText);

                // Quy chuẩn font và khoảng cách đoạn
                para.font.name = "Times New Roman";
                para.font.color = "#000000";
                para.spaceBefore = 0;
                para.spaceAfter = 0;
                para.lineSpacing = 12;

                // Quy chuẩn: Loại bỏ hoàn toàn lùi đầu dòng và thụt lề trong ô bảng
                para.firstLineIndent = 0;
                para.leftIndent = 0;
                para.rightIndent = 0;

                // Cả hai cột đều CĂN GIỮA (Centered)
                para.alignment = "Centered";

                if (isRightColumn) {
                    // CỘT PHẢI: QUỐC HIỆU - TIÊU NGỮ & NGÀY THÁNG
                    if (normText.includes("cong hoa xa hoi chu nghia viet nam")) {
                        // Dòng 1: CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM (12pt, In đậm, In hoa)
                        para.font.size = 12;
                        para.font.bold = true;
                        para.font.italic = false;
                    } else if (normText.includes("doc lap") && normText.includes("tu do")) {
                        // Dòng 2: Độc lập – Tự do – Hạnh phúc (12pt, In đậm)
                        para.font.size = 12;
                        para.font.bold = true;
                        para.font.italic = false;
                    } else if (normText.includes("ngay") || normText.includes("thang") || normText.includes("nam") || normText.includes("...")) {
                        // Dòng địa danh / ngày tháng: In nghiêng, không đậm (12pt)
                        para.font.size = 12;
                        para.font.bold = false;
                        para.font.italic = true;
                    } else if (rawText.length > 0) {
                        para.font.size = 12;
                        para.font.italic = false;
                    }
                } else {
                    // CỘT TRÁI: TÊN CƠ QUAN / ĐƠN VỊ
                    para.font.size = 12;
                    para.font.italic = false;

                    if (paragraphs.length <= 2) {
                        para.font.bold = (p === paragraphs.length - 1);
                    } else {
                        if (p >= 2 || normText.includes("vien") || normText.includes("khoa") || normText.includes("trung tam") || normText.includes("bo mon")) {
                            para.font.bold = true;
                        } else {
                            para.font.bold = false;
                        }
                    }
                }
            }
        }
    }
}

/**
 * Xử lý: Định dạng bảng đang được chọn (Tự động nhận diện Quốc hiệu hoặc Bảng dữ liệu)
 */
async function formatSelectedTable() {
    setButtonsDisabled(true);
    showStatus("Đang kiểm tra bảng được chọn...", "info");

    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();

            const parentTable = selection.parentTableOrNullObject;
            const parentCell = selection.parentTableCellOrNullObject;

            parentTable.load("isNullObject");
            parentCell.load("isNullObject");
            await context.sync();

            let targetTable = null;
            if (!parentTable.isNullObject) {
                targetTable = parentTable;
            } else if (!parentCell.isNullObject) {
                targetTable = parentCell.parentTable;
            }

            if (!targetTable) {
                showStatus("Vui lòng đặt con trỏ chuột vào trong bảng cần định dạng!", "warning");
                return;
            }

            targetTable.load("values");
            await context.sync();

            if (isQuocHieuTable(targetTable.values)) {
                showStatus("Phát hiện phần Quốc hiệu - Tiêu ngữ! Đang định dạng chuẩn...", "info");
                await formatQuocHieuTable(targetTable, context);
                await context.sync();
                showStatus("Đã định dạng phần Quốc hiệu - Tiêu ngữ thành công (xóa viền, căn giữa 2 bên, in đậm/nghiêng đúng chuẩn)!", "success");
            } else {
                const fontSize = getSelectedFontSize();
                const marginDesc = getMarginsDescription();
                showStatus(`Đang định dạng bảng dữ liệu (Font ${fontSize}pt, Lề [${marginDesc}], căn đều cột số)...`, "info");
                await formatSingleTable(targetTable, context);
                await context.sync();
                showStatus(`Đã định dạng bảng dữ liệu thành công (Font ${fontSize}pt, Lề [${marginDesc}])!`, "success");
            }
        });
    } catch (error) {
        console.error("Lỗi:", error);
        showStatus(`Lỗi: ${error.message || error}`, "error");
    } finally {
        setButtonsDisabled(false);
    }
}

/**
 * Xử lý: Định dạng tất cả các bảng trong file
 */
async function formatAllTables() {
    setButtonsDisabled(true);
    const fontSize = getSelectedFontSize();
    const marginDesc = getMarginsDescription();
    showStatus(`Đang tìm và phân loại tất cả các bảng (Font ${fontSize}pt, Lề [${marginDesc}])...`, "info");

    try {
        await Word.run(async (context) => {
            const tables = context.document.body.tables;
            tables.load("items");
            await context.sync();

            const count = tables.items.length;
            if (count === 0) {
                showStatus("Không tìm thấy bảng nào trong tài liệu!", "warning");
                return;
            }

            // Load values cho tất cả các bảng để phân loại
            for (let i = 0; i < count; i++) {
                tables.items[i].load("values");
            }
            await context.sync();

            let dataTablesCount = 0;
            let quocHieuCount = 0;

            for (let i = 0; i < count; i++) {
                const tbl = tables.items[i];
                if (isQuocHieuTable(tbl.values)) {
                    showStatus(`Đang định dạng phần Quốc hiệu - Tiêu ngữ (${i + 1}/${count})...`, "info");
                    await formatQuocHieuTable(tbl, context);
                    quocHieuCount++;
                } else {
                    showStatus(`Đang định dạng bảng dữ liệu (${i + 1}/${count}) với Font ${fontSize}pt, Lề [${marginDesc}]...`, "info");
                    await formatSingleTable(tbl, context);
                    dataTablesCount++;
                }
            }

            await context.sync();
            let msg = `Hoàn thành! Đã định dạng ${dataTablesCount} bảng dữ liệu (Font ${fontSize}pt, Lề [${marginDesc}])`;
            if (quocHieuCount > 0) {
                msg += ` và ${quocHieuCount} bảng Quốc hiệu - Tiêu ngữ`;
            }
            msg += ".";
            showStatus(msg, "success");
        });
    } catch (error) {
        console.error("Lỗi khi định dạng tất cả các bảng:", error);
        showStatus(`Lỗi: ${error.message || error}`, "error");
    } finally {
        setButtonsDisabled(false);
    }
}

/**
 * Nút bấm riêng: Chủ động định dạng phần Quốc hiệu - Tiêu ngữ
 */
async function formatQuocHieuButton() {
    setButtonsDisabled(true);
    showStatus("Đang xử lý phần Quốc hiệu - Tiêu ngữ...", "info");

    try {
        await Word.run(async (context) => {
            const selection = context.document.getSelection();
            const parentTable = selection.parentTableOrNullObject;
            const parentCell = selection.parentTableCellOrNullObject;
            parentTable.load("isNullObject");
            parentCell.load("isNullObject");
            await context.sync();

            let targetTable = null;
            if (!parentTable.isNullObject) {
                targetTable = parentTable;
            } else if (!parentCell.isNullObject) {
                targetTable = parentCell.parentTable;
            }

            if (targetTable) {
                await formatQuocHieuTable(targetTable, context);
                await context.sync();
                showStatus("Đã định dạng bảng đang chọn theo đúng chuẩn Quốc hiệu - Tiêu ngữ!", "success");
                return;
            }

            const tables = context.document.body.tables;
            tables.load("items");
            await context.sync();

            for (let i = 0; i < tables.items.length; i++) {
                tables.items[i].load("values");
            }
            await context.sync();

            let found = false;
            for (let i = 0; i < tables.items.length; i++) {
                if (isQuocHieuTable(tables.items[i].values)) {
                    await formatQuocHieuTable(tables.items[i], context);
                    found = true;
                }
            }

            if (found) {
                await context.sync();
                showStatus("Đã tìm thấy và định dạng bảng Quốc hiệu - Tiêu ngữ theo đúng mẫu chuẩn!", "success");
            } else {
                showStatus("Không tìm thấy bảng Quốc hiệu - Tiêu ngữ. Vui lòng đặt con trỏ chuột vào bảng và thử lại!", "warning");
            }
        });
    } catch (error) {
        console.error("Lỗi:", error);
        showStatus(`Lỗi: ${error.message || error}`, "error");
    } finally {
        setButtonsDisabled(false);
    }
}

/**
 * Khởi tạo tính năng Zoom In / Zoom Out / Reset cho Task Pane bằng Cmd/Ctrl + +/-/0
 */
function initPaneZoom() {
    const STORAGE_KEY = "thesis_taskpane_zoom";
    const MIN_ZOOM = 0.6;
    const MAX_ZOOM = 1.8;
    const STEP = 0.1;

    let currentZoom = 1;
    try {
        const saved = parseFloat(localStorage.getItem(STORAGE_KEY));
        if (!isNaN(saved) && saved >= MIN_ZOOM && saved <= MAX_ZOOM) {
            currentZoom = saved;
        }
    } catch {}

    let toastTimer = null;
    const showZoomToast = (zoomVal) => {
        let toast = document.getElementById("zoom-toast");
        if (!toast) {
            toast = document.createElement("div");
            toast.id = "zoom-toast";
            toast.className = "zoom-toast";
            document.body.appendChild(toast);
        }
        toast.textContent = `${Math.round(zoomVal * 100)}%`;
        toast.classList.add("visible");
        if (toastTimer) clearTimeout(toastTimer);
        toastTimer = setTimeout(() => {
            toast.classList.remove("visible");
        }, 1200);
    };

    const applyZoom = (newZoom, showIndicator = true) => {
        currentZoom = Math.max(MIN_ZOOM, Math.min(MAX_ZOOM, Math.round(newZoom * 10) / 10));
        document.documentElement.style.zoom = String(currentZoom);
        try {
            localStorage.setItem(STORAGE_KEY, String(currentZoom));
        } catch {}
        if (showIndicator) {
            showZoomToast(currentZoom);
        }
    };

    applyZoom(currentZoom, false);

    if (window.__thesisZoomBound) return;
    window.__thesisZoomBound = true;

    document.addEventListener("keydown", (e) => {
        if (!(e.metaKey || e.ctrlKey) || e.altKey) return;

        const key = e.key;
        const code = e.code;

        if (key === "+" || key === "=" || code === "Equal" || code === "NumpadAdd") {
            e.preventDefault();
            e.stopPropagation();
            applyZoom(currentZoom + STEP);
        } else if (key === "-" || key === "_" || code === "Minus" || code === "NumpadSubtract") {
            e.preventDefault();
            e.stopPropagation();
            applyZoom(currentZoom - STEP);
        } else if (key === "0" || code === "Digit0" || code === "Numpad0") {
            e.preventDefault();
            e.stopPropagation();
            applyZoom(1);
        }
    });
}

/**
 * Tải lại Task Pane kèm cache-busting để nhận bản cập nhật mới nhất
 */
async function forceReloadTaskpane() {
    try {
        if ("caches" in window) {
            const keys = await caches.keys();
            await Promise.all(keys.map((k) => caches.delete(k)));
        }
    } catch {}

    try {
        const url = new URL(window.location.href);
        url.searchParams.set("_v", String(Date.now()));
        window.location.replace(url.toString());
    } catch {
        window.location.reload();
    }
}

/**
 * Lắng nghe tín hiệu cập nhật từ nút trên Ribbon Bar
 */
function initAddinReloadListener() {
    if (window.__thesisReloadBound) return;
    window.__thesisReloadBound = true;

    window.addEventListener("storage", (e) => {
        if (e.key === "thesis_addin_force_reload") {
            forceReloadTaskpane();
        }
    });

    try {
        if (typeof BroadcastChannel !== "undefined") {
            const ch = new BroadcastChannel("thesis_addin_channel");
            ch.onmessage = (msg) => {
                if (msg && msg.data === "force_reload") {
                    forceReloadTaskpane();
                }
            };
        }
    } catch {}
}

/**
 * Hàm ExecuteFunction được gọi khi bấm nút "Cập nhật Add-in" trên Ribbon Bar
 * @param {Office.AddinCommands.Event} event
 */
async function refreshAddin(event) {
    try {
        try {
            localStorage.setItem("thesis_addin_force_reload", String(Date.now()));
        } catch {}

        try {
            if (typeof BroadcastChannel !== "undefined") {
                const ch = new BroadcastChannel("thesis_addin_channel");
                ch.postMessage("force_reload");
                ch.close();
            }
        } catch {}

        try {
            if ("caches" in window) {
                const keys = await caches.keys();
                await Promise.all(keys.map((k) => caches.delete(k)));
            }
        } catch {}
    } finally {
        try {
            if (event && typeof event.completed === "function") {
                event.completed();
            }
        } catch {}
    }
}

if (typeof window !== "undefined") {
    window.refreshAddin = refreshAddin;
}
