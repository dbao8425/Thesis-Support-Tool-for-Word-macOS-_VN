/**
 * Word Table Formatter for Thesis / Scientific Papers
 * Uses Office JavaScript API (Word.run)
 */

/* global Office, Word */

Office.onReady((info) => {
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

    // 2. Load các hàng (rows) của bảng
    const rows = table.rows;
    rows.load("items");
    await context.sync();

    if (rows.items.length === 0) {
        return;
    }

    // 3. Load các ô (cells) cho từng hàng
    for (let r = 0; r < rows.items.length; r++) {
        rows.items[r].cells.load("items");
    }
    await context.sync();

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

                // Khoảng cách đoạn: Before = 0, After = 0, Line spacing = dòng đơn theo cỡ chữ
                para.spaceBefore = 0;
                para.spaceAfter = 0;
                para.lineSpacing = fontSize;

                // Quy tắc cho từng hàng/cột:
                if (isHeaderRow) {
                    para.font.bold = true;
                    para.alignment = "Centered";
                } else {
                    para.font.bold = false;

                    if (isFirstCol) {
                        para.alignment = "Left";
                    } else {
                        para.alignment = "Right";
                    }
                }
            }
        }
    }

    // 6. Căn đều kích thước các cột số liệu (Cột 2 trở đi) bằng Word.TableCell.columnWidth
    if (shouldEqualCols && rows.items.length > 0) {
        try {
            const colCount = rows.items[0].cells.items.length;

            if (colCount >= 2) {
                // Chiều rộng trang chuẩn A4 (lề 2.5cm) là ~468pt
                const totalWidth = 468;

                // Tỷ lệ phân bổ cho Cột 1 (nhãn/chỉ tiêu):
                // Bảng càng nhiều cột số thì cột nhãn càng co lại hợp lý
                let labelRatio = 0.36;
                if (colCount >= 7) labelRatio = 0.25;
                else if (colCount >= 6) labelRatio = 0.28;
                else if (colCount >= 5) labelRatio = 0.30;
                else if (colCount >= 4) labelRatio = 0.33;

                const labelWidth = Math.round(totalWidth * labelRatio);
                const numColWidth = Math.floor((totalWidth - labelWidth) / (colCount - 1));

                // Gán columnWidth trực tiếp cho từng ô trong bảng (WordApi 1.3+ tương thích 100% Word for Mac)
                for (let r = 0; r < rows.items.length; r++) {
                    const rowCells = rows.items[r].cells.items;
                    for (let c = 0; c < rowCells.length; c++) {
                        rowCells[c].columnWidth = (c === 0) ? labelWidth : numColWidth;
                    }
                }
                await context.sync();
            }
        } catch (e) {
            console.warn("Lỗi khi điều chỉnh độ rộng cột:", e);
        }
    } else {
        // Chỉ chạy AutoFit Window tự do nếu người dùng tắt tính năng căn đều cột
        try {
            if (typeof table.autoFitWindow === "function") {
                table.autoFitWindow();
            } else if (typeof table.autoFitBehavior === "function") {
                table.autoFitBehavior("Window");
            }
        } catch (e) {
            console.warn("AutoFit Window không được hỗ trợ hoặc gặp lỗi:", e);
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

    // 6. AutoFit Window
    try {
        if (typeof table.autoFitWindow === "function") {
            table.autoFitWindow();
        } else if (typeof table.autoFitBehavior === "function") {
            table.autoFitBehavior("Window");
        }
    } catch {}
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
