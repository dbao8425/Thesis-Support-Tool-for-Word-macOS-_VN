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
    }
});

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
 * - Table Style: Table Grid
 * - Font: Times New Roman 11pt, #000000
 * - Giãn đoạn: 0pt Before/After, Single line
 * - Căn lề dọc ô: Center
 * - Hàng 1: In đậm, căn giữa
 * - Từ hàng 2: Cột 1 căn trái, Cột 2+ căn phải
 * - AutoFit Window
 * @param {Word.Table} table
 * @param {Word.RequestContext} context
 */
async function formatSingleTable(table, context) {
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

    // 4. Căn lề dọc ô (Center) và load đoạn văn bản qua cell.body
    for (let r = 0; r < rows.items.length; r++) {
        const cells = rows.items[r].cells.items;
        for (let c = 0; c < cells.length; c++) {
            const cell = cells[c];
            cell.verticalAlignment = "Center";
            cell.body.paragraphs.load("items");
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

                // Font chữ toàn bảng: Times New Roman, 11pt, #000000
                para.font.name = "Times New Roman";
                para.font.size = 11;
                para.font.color = "#000000";

                // Khoảng cách đoạn: Before = 0, After = 0, Line spacing = Single (11pt)
                para.spaceBefore = 0;
                para.spaceAfter = 0;
                para.lineSpacing = 11;

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

    // 6. Tự động co giãn bảng theo độ rộng trang (AutoFit Window)
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

/**
 * 2. ĐỊNH DẠNG BẢNG QUỐC HIỆU - TIÊU NGỮ (Header Table)
 * - Xóa sạch mọi đường viền bảng (No border)
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

    // 2. Xóa các đường viền của bảng (Dùng các vị trí an toàn cho bảng 1 hàng)
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

            // Xóa triệt để viền từng ô để không còn đường kẻ nào sót lại
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

    // 4. Áp dụng định dạng chi tiết từng ô
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

                    // Theo mẫu: 2 dòng đầu (TRƯỜNG ĐẠI HỌC TÀI CHÍNH - MARKETING) là chữ thường/không đậm
                    // Các dòng đơn vị trực tiếp (VIỆN CÔNG NGHỆ TÀI CHÍNH NGÂN HÀNG) là in đậm
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

    // 5. AutoFit Window
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
                showStatus("Đang định dạng bảng dữ liệu theo chuẩn khóa luận...", "info");
                await formatSingleTable(targetTable, context);
                await context.sync();
                showStatus("Đã định dạng bảng dữ liệu được chọn thành công!", "success");
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
 * - Tự động nhận diện bảng Quốc hiệu - Tiêu ngữ để định dạng không viền
 * - Định dạng các bảng dữ liệu còn lại theo chuẩn bảng số liệu khóa luận
 */
async function formatAllTables() {
    setButtonsDisabled(true);
    showStatus("Đang tìm và phân loại tất cả các bảng trong tài liệu...", "info");

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
                    showStatus(`Đang định dạng bảng dữ liệu (${i + 1}/${count})...`, "info");
                    await formatSingleTable(tbl, context);
                    dataTablesCount++;
                }
            }

            await context.sync();
            let msg = `Hoàn thành! Đã định dạng ${dataTablesCount} bảng dữ liệu`;
            if (quocHieuCount > 0) {
                msg += ` và ${quocHieuCount} bảng Quốc hiệu - Tiêu ngữ (đã xóa viền và căn chuẩn)`;
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
            // Kiểm tra xem người dùng có đang đặt trỏ ở bảng nào không
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

            // Nếu người dùng đã chọn một bảng cụ thể
            if (targetTable) {
                await formatQuocHieuTable(targetTable, context);
                await context.sync();
                showStatus("Đã định dạng bảng đang chọn theo đúng chuẩn Quốc hiệu - Tiêu ngữ!", "success");
                return;
            }

            // Nếu không chọn bảng, tự động tìm bảng Quốc hiệu trong tài liệu
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
