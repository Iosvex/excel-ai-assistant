// ============================================================
// CONFIGURATION — update this to your Render backend URL
// ============================================================
const BACKEND_URL = "https://surface-antivirus-footage.ngrok-free.dev";
// If frontend and backend are on same server (FastAPI serves static),
// change BACKEND_URL to "" (empty string) to use relative URLs.

// ============================================================
// GLOBAL STATE
// ============================================================
let chatHistory = [];
let officeReady = false;

// ============================================================
// OFFICE INITIALIZATION
// ============================================================
Office.onReady((info) => {
    console.log("[INIT] Office.js ready:", info);
    officeReady = true;

    const sendBtn = document.getElementById("sendBtn");
    const userInput = document.getElementById("userInput");

    if (sendBtn) sendBtn.addEventListener("click", sendMessage);
    if (userInput) {
        userInput.addEventListener("keypress", (e) => {
            if (e.key === "Enter" && !e.shiftKey) {
                e.preventDefault();
                sendMessage();
            }
        });
    }

    addMessage("assistant",
        "Hi! I'm your Excel AI Assistant. I can add rows, edit data, insert formulas, format cells, clean columns, sort sheets, create new sheets, and answer questions about your data. Just tell me what you need.",
        false
    );
});

window.addEventListener("load", () => {
    setTimeout(() => {
        if (!officeReady) {
            console.warn("[WARN] Office.js did not initialize");
            addMessage("assistant", "Office.js failed to load. Please close and reopen the add-in.", false, true);
        }
    }, 3000);
});

// ============================================================
// UI HELPERS
// ============================================================
function addMessage(role, content, addToHistory = true, isError = false) {
    const container = document.getElementById("chatContainer");
    if (!container) return;

    const wrap = document.createElement("div");
    wrap.className = `msg-wrap ${role === "user" ? "user-wrap" : "ai-wrap"}`;

    const bubble = document.createElement("div");
    bubble.className = `bubble ${role === "user" ? "user-bubble" : "ai-bubble"}${isError ? " error-bubble" : ""}`;
    bubble.textContent = content;

    wrap.appendChild(bubble);
    container.appendChild(wrap);
    container.scrollTop = container.scrollHeight;

    if (addToHistory && (role === "user" || role === "assistant")) {
        chatHistory.push({ role, content });
    }
}

function addFeedback(text, isError = false) {
    const container = document.getElementById("chatContainer");
    if (!container) return;

    const fb = document.createElement("div");
    fb.className = `feedback ${isError ? "feedback-error" : "feedback-ok"}`;
    fb.textContent = text;
    container.appendChild(fb);
    container.scrollTop = container.scrollHeight;
}

let loadingEl = null;
function showLoading() {
    const container = document.getElementById("chatContainer");
    if (!container) return;

    const wrap = document.createElement("div");
    wrap.className = "msg-wrap ai-wrap";
    wrap.id = "loadingWrap";

    const dots = document.createElement("div");
    dots.className = "loading-dots";
    dots.innerHTML = "<span></span><span></span><span></span>";

    wrap.appendChild(dots);
    container.appendChild(wrap);
    container.scrollTop = container.scrollHeight;
    loadingEl = wrap;
}

function hideLoading() {
    if (loadingEl) {
        loadingEl.remove();
        loadingEl = null;
    }
}

// ============================================================
// SHEET SCHEMA READER
// ============================================================
async function getAllSheetsSchema() {
    if (!officeReady) return {};
    console.log("[SCHEMA] Reading all sheet headers...");

    return Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();

        const schema = {};
        for (const sheet of sheets.items) {
            const usedRange = sheet.getUsedRangeOrNullObject();
            usedRange.load("rowCount, values");
            await context.sync();

            if (!usedRange.isNullObject && usedRange.rowCount > 0 && usedRange.values[0]) {
                schema[sheet.name] = usedRange.values[0].filter(h => h !== null && h !== "");
            } else {
                schema[sheet.name] = [];
            }
        }

        console.log("[SCHEMA] Sheet schema:", schema);
        return schema;
    }).catch(err => {
        console.error("[SCHEMA ERROR]", err);
        return {};
    });
}

// ============================================================
// SEND MESSAGE — MAIN FLOW
// ============================================================
async function sendMessage() {
    const inputField = document.getElementById("userInput");
    if (!inputField) return;

    const userText = inputField.value.trim();
    if (!userText) return;

    console.log("[SEND] User input:", userText);
    addMessage("user", userText, true);
    inputField.value = "";
    showLoading();

    if (!officeReady) {
        hideLoading();
        addMessage("assistant", "Office.js is not ready. Please reload the add-in.", false, true);
        return;
    }

    try {
        // STEP 1: Read sheet schema
        console.log("[STEP 1] Reading sheet schema");
        const sheetSchema = await getAllSheetsSchema();

        // STEP 2: Send to backend
        console.log("[STEP 2] Sending to backend:", BACKEND_URL + "/process");
        const response = await fetch(`${BACKEND_URL}/process`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                user_input: userText,
                sheet_schema: sheetSchema,
                chat_history: chatHistory.slice(-30)
            })
        });

        if (!response.ok) {
            throw new Error(`Backend returned HTTP ${response.status}`);
        }

        const data = await response.json();
        console.log("[STEP 2] Backend response:", data);

        hideLoading();

        // STEP 3: Show AI reply
        if (data.reply) {
            console.log("[STEP 3] Showing reply:", data.reply);
            addMessage("assistant", data.reply, true);
        }

        // STEP 4: Execute actions
        if (data.actions && data.actions.length > 0) {
            console.log(`[STEP 4] Executing ${data.actions.length} action(s)`);
            for (const action of data.actions) {
                await executeAction(action);
            }
        } else {
            console.log("[STEP 4] No actions to execute (query or error)");
        }

        // STEP 5: Show warnings
        if (data.warnings && data.warnings.length > 0) {
            data.warnings.forEach(w => {
                console.warn("[WARNING]", w);
                addFeedback("⚠ " + w, true);
            });
        }

    } catch (err) {
        hideLoading();
        console.error("[ERROR] sendMessage failed:", err);
        addMessage("assistant", `Error: ${err.message}. Check that your backend is running and BACKEND_URL is correct.`, false, true);
    }
}

// ============================================================
// ACTION DISPATCHER
// ============================================================
async function executeAction(action) {
    console.log("[ACTION] Dispatching:", action.type, action);

    try {
        switch (action.type) {
            case "add_row":
                await actionAddRow(action.sheet, action.data);
                break;
            case "edit_row":
                await actionEditRow(action.sheet, action.search, action.new_values);
                break;
            case "delete_row":
                await actionDeleteRow(action.sheet, action.search, action.row_index);
                break;
            case "formula":
                await actionInsertFormula(action.sheet, action.cell, action.formula);
                break;
            case "format":
                await actionFormatRange(action.sheet, action.range, action.formatting);
                break;
            case "clean":
                await actionCleanColumn(action.sheet, action.column, action.clean_action || "trim");
                break;
            case "create_sheet":
                await actionCreateSheet(action.new_sheet, action.headers || []);
                break;
            case "sort":
                await actionSort(action.sheet, action.sort_by, action.order || "asc");
                break;
            case "query":
                console.log("[ACTION] Query — reply already shown");
                break;
            default:
                console.warn("[ACTION] Unknown type:", action.type);
                addFeedback(`Unknown action: ${action.type}`, true);
        }
    } catch (err) {
        console.error(`[ACTION ERROR] ${action.type}:`, err);
        addFeedback(`Failed (${action.type}): ${err.message}`, true);
    }
}

// ============================================================
// ADD ROW
// ============================================================
async function actionAddRow(sheetName, dataObject) {
    console.log("[ADD ROW] Sheet:", sheetName, "Data:", dataObject);

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("isNullObject");
        await context.sync();

        console.log("[ADD ROW] Sheet isNullObject:", sheet.isNullObject);
        if (sheet.isNullObject) throw new Error(`Sheet "${sheetName}" not found`);

        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("rowCount, values");
        await context.sync();

        console.log("[ADD ROW] rowCount:", usedRange.rowCount);

        let headers = [];
        let newRowIndex = 0;

        if (!usedRange.isNullObject && usedRange.rowCount > 0) {
            headers = usedRange.values[0].filter(h => h !== null && h !== "");
            newRowIndex = usedRange.rowCount;
        } else {
            // Sheet is empty — create headers from data keys
            headers = Object.keys(dataObject);
            const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
            headerRange.values = [headers];
            headerRange.format.font.bold = true;
            headerRange.format.fill.color = "#1e3a5f";
            headerRange.format.font.color = "#ffffff";
            await context.sync();
            newRowIndex = 1;
            console.log("[ADD ROW] Created headers:", headers);
        }

        const rowValues = headers.map(h => {
            // Case-insensitive key matching
            const key = Object.keys(dataObject).find(k => k.toLowerCase() === h.toLowerCase());
            return key !== undefined ? dataObject[key] : "";
        });

        console.log("[ADD ROW] Writing row:", rowValues, "at index:", newRowIndex);

        const targetRange = sheet.getRangeByIndexes(newRowIndex, 0, 1, headers.length);
        targetRange.values = [rowValues];
        targetRange.format.autofitColumns();
        await context.sync();

        console.log("[ADD ROW] Success");
        addFeedback(`✓ Row added to "${sheetName}"`);
    });
}

// ============================================================
// EDIT ROW
// ============================================================
async function actionEditRow(sheetName, searchCriteria, newValues) {
    console.log("[EDIT ROW] Sheet:", sheetName, "Search:", searchCriteria, "New values:", newValues);

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("isNullObject");
        await context.sync();
        if (sheet.isNullObject) throw new Error(`Sheet "${sheetName}" not found`);

        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("values, rowCount");
        await context.sync();
        if (usedRange.isNullObject || usedRange.rowCount <= 1) throw new Error("No data in sheet");

        const headers = usedRange.values[0];

        // Find target row by search criteria
        let targetRow = -1;
        if (searchCriteria) {
            const [searchKey, searchVal] = Object.entries(searchCriteria)[0];
            const colIdx = headers.findIndex(h => h && h.toString().toLowerCase() === searchKey.toLowerCase());
            if (colIdx === -1) throw new Error(`Column "${searchKey}" not found`);

            for (let i = 1; i < usedRange.rowCount; i++) {
                if (usedRange.values[i][colIdx] != null &&
                    usedRange.values[i][colIdx].toString().toLowerCase() === searchVal.toString().toLowerCase()) {
                    targetRow = i;
                    break;
                }
            }
        }

        if (targetRow === -1) throw new Error(`Row not found matching ${JSON.stringify(searchCriteria)}`);
        console.log("[EDIT ROW] Found row at index:", targetRow);

        for (const [colName, newVal] of Object.entries(newValues)) {
            const colIdx = headers.findIndex(h => h && h.toString().toLowerCase() === colName.toLowerCase());
            if (colIdx !== -1) {
                const cell = sheet.getRangeByIndexes(targetRow, colIdx, 1, 1);
                cell.values = [[newVal]];
                console.log(`[EDIT ROW] Updated "${colName}" to "${newVal}"`);
            } else {
                console.warn(`[EDIT ROW] Column "${colName}" not found, skipping`);
            }
        }

        await context.sync();
        console.log("[EDIT ROW] Success");
        addFeedback(`✓ Row updated in "${sheetName}"`);
    });
}

// ============================================================
// DELETE ROW
// ============================================================
async function actionDeleteRow(sheetName, searchCriteria, rowIndex) {
    console.log("[DELETE ROW] Sheet:", sheetName, "Search:", searchCriteria);

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("isNullObject");
        await context.sync();
        if (sheet.isNullObject) throw new Error(`Sheet "${sheetName}" not found`);

        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("values, rowCount");
        await context.sync();
        if (usedRange.isNullObject || usedRange.rowCount <= 1) throw new Error("No data in sheet");

        const headers = usedRange.values[0];
        let targetRow = -1;

        if (rowIndex !== undefined && rowIndex >= 0) {
            targetRow = rowIndex;
        } else if (searchCriteria) {
            const [searchKey, searchVal] = Object.entries(searchCriteria)[0];
            const colIdx = headers.findIndex(h => h && h.toString().toLowerCase() === searchKey.toLowerCase());
            if (colIdx === -1) throw new Error(`Column "${searchKey}" not found`);

            for (let i = 1; i < usedRange.rowCount; i++) {
                if (usedRange.values[i][colIdx] != null &&
                    usedRange.values[i][colIdx].toString().toLowerCase() === searchVal.toString().toLowerCase()) {
                    targetRow = i;
                    break;
                }
            }
        }

        if (targetRow === -1) throw new Error("Row not found");
        console.log("[DELETE ROW] Deleting row at index:", targetRow);

        const rowRange = sheet.getRangeByIndexes(targetRow, 0, 1, headers.length);
        rowRange.delete(Excel.DeleteShiftDirection.up);
        await context.sync();

        console.log("[DELETE ROW] Success");
        addFeedback(`✓ Row deleted from "${sheetName}"`);
    });
}

// ============================================================
// INSERT FORMULA
// ============================================================
async function actionInsertFormula(sheetName, cellAddress, formula) {
    console.log("[FORMULA] Sheet:", sheetName, "Cell:", cellAddress, "Formula:", formula);

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("isNullObject");
        await context.sync();
        if (sheet.isNullObject) throw new Error(`Sheet "${sheetName}" not found`);

        const range = sheet.getRange(cellAddress);
        range.formulas = [[formula]];
        await context.sync();

        console.log("[FORMULA] Success");
        addFeedback(`✓ Formula "${formula}" inserted at ${sheetName}!${cellAddress}`);
    });
}

// ============================================================
// FORMAT RANGE
// ============================================================
async function actionFormatRange(sheetName, rangeAddress, formatting) {
    console.log("[FORMAT] Sheet:", sheetName, "Range:", rangeAddress, "Formatting:", formatting);

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("isNullObject");
        await context.sync();
        if (sheet.isNullObject) throw new Error(`Sheet "${sheetName}" not found`);

        const range = sheet.getRange(rangeAddress);

        if (!formatting) throw new Error("No formatting options provided");
        if (formatting.bold !== undefined) range.format.font.bold = formatting.bold;
        if (formatting.color) range.format.font.color = formatting.color;
        if (formatting.fillColor) range.format.fill.color = formatting.fillColor;
        if (formatting.fontSize) range.format.font.size = formatting.fontSize;
        if (formatting.italic !== undefined) range.format.font.italic = formatting.italic;
        if (formatting.numberFormat) range.numberFormat = [[formatting.numberFormat]];
        if (formatting.horizontalAlignment) range.format.horizontalAlignment = formatting.horizontalAlignment;
        if (formatting.wrapText !== undefined) range.format.wrapText = formatting.wrapText;

        await context.sync();
        console.log("[FORMAT] Success");
        addFeedback(`✓ Formatted ${sheetName}!${rangeAddress}`);
    });
}

// ============================================================
// CLEAN COLUMN
// ============================================================
async function actionCleanColumn(sheetName, columnName, cleanAction) {
    console.log("[CLEAN] Sheet:", sheetName, "Column:", columnName, "Action:", cleanAction);

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("isNullObject");
        await context.sync();
        if (sheet.isNullObject) throw new Error(`Sheet "${sheetName}" not found`);

        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("values, rowCount");
        await context.sync();
        if (usedRange.isNullObject || usedRange.rowCount <= 1) throw new Error("No data in sheet");

        const headers = usedRange.values[0];
        const colIdx = headers.findIndex(h => h && h.toString().toLowerCase() === columnName.toLowerCase());
        if (colIdx === -1) throw new Error(`Column "${columnName}" not found`);

        let changedCount = 0;
        for (let i = 1; i < usedRange.rowCount; i++) {
            let val = usedRange.values[i][colIdx];
            if (typeof val !== "string") continue;

            let cleaned = val;
            if (cleanAction === "trim") cleaned = val.trim().replace(/\s+/g, " ");
            else if (cleanAction === "uppercase") cleaned = val.toUpperCase();
            else if (cleanAction === "lowercase") cleaned = val.toLowerCase();
            else if (cleanAction === "capitalize") cleaned = val.charAt(0).toUpperCase() + val.slice(1).toLowerCase();

            if (cleaned !== val) {
                const cell = sheet.getRangeByIndexes(i, colIdx, 1, 1);
                cell.values = [[cleaned]];
                changedCount++;
            }
        }

        await context.sync();
        console.log(`[CLEAN] Success — ${changedCount} cells updated`);
        addFeedback(`✓ Cleaned "${columnName}" in "${sheetName}" — ${changedCount} cells updated`);
    });
}

// ============================================================
// CREATE SHEET
// ============================================================
async function actionCreateSheet(sheetName, headers) {
    console.log("[CREATE SHEET] Name:", sheetName, "Headers:", headers);

    await Excel.run(async (context) => {
        const existing = context.workbook.worksheets.getItemOrNullObject(sheetName);
        existing.load("isNullObject");
        await context.sync();

        if (!existing.isNullObject) throw new Error(`Sheet "${sheetName}" already exists`);

        const newSheet = context.workbook.worksheets.add(sheetName);

        if (headers && headers.length > 0) {
            const headerRange = newSheet.getRangeByIndexes(0, 0, 1, headers.length);
            headerRange.values = [headers];
            headerRange.format.font.bold = true;
            headerRange.format.fill.color = "#1e3a5f";
            headerRange.format.font.color = "#ffffff";
            headerRange.format.autofitColumns();
        }

        await context.sync();
        console.log("[CREATE SHEET] Success");
        addFeedback(`✓ Created sheet "${sheetName}"${headers.length > 0 ? " with " + headers.length + " columns" : ""}`);
    });
}

// ============================================================
// SORT
// ============================================================
async function actionSort(sheetName, sortColumn, order) {
    console.log("[SORT] Sheet:", sheetName, "Column:", sortColumn, "Order:", order);

    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
        sheet.load("isNullObject");
        await context.sync();
        if (sheet.isNullObject) throw new Error(`Sheet "${sheetName}" not found`);

        const usedRange = sheet.getUsedRangeOrNullObject();
        usedRange.load("values, rowCount, columnCount");
        await context.sync();
        if (usedRange.isNullObject || usedRange.rowCount <= 1) throw new Error("No data to sort");

        const headers = usedRange.values[0];
        const colIdx = headers.findIndex(h => h && h.toString().toLowerCase() === sortColumn.toLowerCase());
        if (colIdx === -1) throw new Error(`Column "${sortColumn}" not found`);

        // Sort data rows only (exclude header row)
        const dataRange = sheet.getRangeByIndexes(1, 0, usedRange.rowCount - 1, usedRange.columnCount);
        dataRange.sort.apply([{
            key: colIdx,
            ascending: order !== "desc"
        }]);

        await context.sync();
        console.log("[SORT] Success");
        addFeedback(`✓ Sorted "${sheetName}" by "${sortColumn}" (${order}ending)`);
    });
}