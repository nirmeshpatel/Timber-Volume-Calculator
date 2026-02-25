const STORAGE_KEY = "cylinderVolumeHistory";
const CUSTOMER_HISTORY_KEY = "customerVolumeHistory";
const EXCEL_SHEET_NAME = "CustomerData";
const EXCEL_HEADERS = ["Date", "Name", "WhatsApp Number", "Address", "Total Volume (ft^3)"];
const HANDLE_DB_NAME = "excelFileHandleDb";
const HANDLE_STORE_NAME = "handles";
const HANDLE_KEY = "customerExcelHandle";

function setExcelStatus(message, isError = false) {
    const status = document.getElementById("excelStatus");
    if (!status) return;
    status.textContent = `Excel: ${message}`;
    status.classList.toggle("error", isError);
}

function openHandleDb() {
    return new Promise((resolve, reject) => {
        const request = indexedDB.open(HANDLE_DB_NAME, 1);

        request.onupgradeneeded = () => {
            const db = request.result;
            if (!db.objectStoreNames.contains(HANDLE_STORE_NAME)) {
                db.createObjectStore(HANDLE_STORE_NAME);
            }
        };

        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
    });
}

async function saveFileHandle(handle) {
    const db = await openHandleDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(HANDLE_STORE_NAME, "readwrite");
        tx.objectStore(HANDLE_STORE_NAME).put(handle, HANDLE_KEY);
        tx.oncomplete = () => resolve();
        tx.onerror = () => reject(tx.error);
    });
}

async function getSavedFileHandle() {
    const db = await openHandleDb();
    return new Promise((resolve, reject) => {
        const tx = db.transaction(HANDLE_STORE_NAME, "readonly");
        const request = tx.objectStore(HANDLE_STORE_NAME).get(HANDLE_KEY);
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error);
    });
}

async function ensureReadWritePermission(handle) {
    if (!handle) return false;

    const current = await handle.queryPermission({ mode: "readwrite" });
    if (current === "granted") return true;

    const requested = await handle.requestPermission({ mode: "readwrite" });
    return requested === "granted";
}

async function hasReadWritePermission(handle) {
    if (!handle) return false;
    try {
        return (await handle.queryPermission({ mode: "readwrite" })) === "granted";
    } catch (error) {
        return false;
    }
}

function createWorkbookWithHeaders() {
    const workbook = XLSX.utils.book_new();
    const sheet = XLSX.utils.aoa_to_sheet([EXCEL_HEADERS]);
    XLSX.utils.book_append_sheet(workbook, sheet, EXCEL_SHEET_NAME);
    return workbook;
}

async function writeWorkbookToHandle(workbook, handle) {
    const data = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
    const writable = await handle.createWritable();
    await writable.write(data);
    await writable.close();
}

function appendRowToSheet(sheet, row) {
    const allRows = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
    if (allRows.length === 0) {
        XLSX.utils.sheet_add_aoa(sheet, [EXCEL_HEADERS], { origin: "A1" });
    }
    XLSX.utils.sheet_add_aoa(sheet, [row], { origin: -1 });
}

async function connectExcelFile() {
    if (!window.showSaveFilePicker || !window.indexedDB) {
        setExcelStatus("Not supported in this browser. Use latest Chrome/Edge.", true);
        return false;
    }

    if (!window.XLSX) {
        setExcelStatus("Excel library not loaded. Check internet and reload.", true);
        return false;
    }

    try {
        const handle = await window.showSaveFilePicker({
            suggestedName: "customer_data.xlsx",
            types: [
                {
                    description: "Excel Workbook",
                    accept: {
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"]
                    }
                }
            ]
        });

        const hasPermission = await ensureReadWritePermission(handle);
        if (!hasPermission) {
            setExcelStatus("Permission denied for Excel file.", true);
            return false;
        }

        await saveFileHandle(handle);

        const existingFile = await handle.getFile();
        if (existingFile.size === 0) {
            const workbook = createWorkbookWithHeaders();
            await writeWorkbookToHandle(workbook, handle);
        }

        setExcelStatus(`Connected (${handle.name})`);
        return true;
    } catch (error) {
        if (error && error.name === "AbortError") {
            setExcelStatus("Connection cancelled.");
            return false;
        }
        setExcelStatus("Failed to connect Excel file.", true);
        return false;
    }
}

async function appendCustomerToExcel(entry, options = {}) {
    const autoConnect = Boolean(options.autoConnect);

    if (!window.XLSX) {
        setExcelStatus("Excel library not loaded. Check internet and reload.", true);
        return false;
    }

    if (!window.indexedDB || !window.showSaveFilePicker) {
        setExcelStatus("Browser does not support direct Excel autosave.", true);
        return false;
    }

    try {
        let handle = await getSavedFileHandle();

        if (!handle) {
            if (autoConnect) {
                const connected = await connectExcelFile();
                if (!connected) return false;
                handle = await getSavedFileHandle();
            } else {
                setExcelStatus("Not connected. Click Connect Excel.", true);
                return false;
            }
        }

        const hasPermission = await ensureReadWritePermission(handle);
        if (!hasPermission) {
            setExcelStatus("No permission to update Excel file.", true);
            return false;
        }

        const file = await handle.getFile();
        let workbook;

        if (file.size > 0) {
            const arrayBuffer = await file.arrayBuffer();
            workbook = XLSX.read(arrayBuffer, { type: "array" });
        } else {
            workbook = createWorkbookWithHeaders();
        }

        let sheet = workbook.Sheets[EXCEL_SHEET_NAME];
        if (!sheet) {
            sheet = XLSX.utils.aoa_to_sheet([EXCEL_HEADERS]);
            workbook.Sheets[EXCEL_SHEET_NAME] = sheet;
            workbook.SheetNames.push(EXCEL_SHEET_NAME);
        }

        appendRowToSheet(sheet, [
            entry.date,
            entry.name,
            entry.whatsapp,
            entry.address,
            entry.totalVolume.toFixed(3)
        ]);

        await writeWorkbookToHandle(workbook, handle);
        setExcelStatus(`Updated at ${new Date().toLocaleTimeString()}`);
        return true;
    } catch (error) {
        setExcelStatus("Could not update Excel file.", true);
        return false;
    }
}

async function restoreExcelConnectionStatus() {
    if (!window.indexedDB || !window.showSaveFilePicker) {
        setExcelStatus("Feature requires latest Chrome/Edge.", true);
        return;
    }

    try {
        const handle = await getSavedFileHandle();
        if (!handle) {
            setExcelStatus("Not connected");
            return;
        }

        const hasPermission = await hasReadWritePermission(handle);
        if (!hasPermission) {
            setExcelStatus("Re-connect needed (permission not granted)", true);
            return;
        }

        setExcelStatus(`Connected (${handle.name})`);
    } catch (error) {
        setExcelStatus("Connection unavailable. Re-connect Excel.", true);
    }
}

function getHistory() {
    const raw = JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]");
    if (!Array.isArray(raw)) return [];
    return raw.filter(item =>
        item &&
        typeof item.circumference === "string" &&
        typeof item.height === "string" &&
        typeof item.volume === "number" &&
        Number.isFinite(item.volume)
    );
}

function saveHistory(history) {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(history));
}

function getCustomerHistory() {
    const raw = JSON.parse(localStorage.getItem(CUSTOMER_HISTORY_KEY) || "[]");
    if (!Array.isArray(raw)) return [];
    return raw.filter(item =>
        item &&
        typeof item.id === "string" &&
        typeof item.date === "string" &&
        typeof item.name === "string" &&
        typeof item.whatsapp === "string" &&
        typeof item.address === "string" &&
        typeof item.totalVolume === "number" &&
        Number.isFinite(item.totalVolume)
    );
}

function saveCustomerHistory(history) {
    localStorage.setItem(CUSTOMER_HISTORY_KEY, JSON.stringify(history));
}

function renderHistory() {
    const historyBody = document.getElementById("historyBody");
    const totalVolume = document.getElementById("totalVolume");
    const history = getHistory();
    let total = 0;

    if (history.length === 0) {
        historyBody.innerHTML = '<tr><td class="empty" colspan="4">No saved calculations yet.</td></tr>';
        totalVolume.textContent = "Total Volume: 0.000 ft^3";
        return;
    }

    historyBody.innerHTML = history
        .map((item, index) => {
            total += item.volume;
            return `<tr>
                <td>${item.circumference}</td>
                <td>${item.height}</td>
                <td>${item.volume.toFixed(3)}</td>
                <td><button class="delete-btn" onclick="deleteEntry(${index})">Delete</button></td>
            </tr>`;
        })
        .join("");
    totalVolume.textContent = `Total Volume: ${total.toFixed(3)} ft^3`;
}

function renderCustomerHistory() {
    const body = document.getElementById("customerHistoryBody");
    const history = getCustomerHistory();

    if (history.length === 0) {
        body.innerHTML = '<tr><td class="empty" colspan="6">No customer history yet.</td></tr>';
        return;
    }

    body.innerHTML = history
        .map(item => `<tr>
            <td>${item.date}</td>
            <td>${item.name}</td>
            <td>${item.whatsapp}</td>
            <td>${item.address}</td>
            <td>${item.totalVolume.toFixed(3)}</td>
            <td><button class="delete-btn" onclick="deleteCustomerHistory('${item.id}')">Delete</button></td>
        </tr>`)
        .join("");
}

function getTotalVolume(history) {
    return history.reduce((sum, item) => sum + item.volume, 0);
}

function createCustomerRecord(totalVolume) {
    const customerName = document.getElementById("customerName").value.trim() || "Customer";
    const customerPhone = document.getElementById("customerPhone").value.trim();
    const customerAddress = document.getElementById("customerAddress").value.trim() || "Not provided";
    const cleanPhone = customerPhone.replace(/[^\d]/g, "");

    return {
        customerName,
        customerAddress,
        cleanPhone,
        customerRecord: {
            id: `${Date.now()}_${Math.random().toString(36).slice(2, 8)}`,
            date: new Date().toLocaleString(),
            name: customerName,
            whatsapp: cleanPhone || customerPhone || "Not provided",
            address: customerAddress,
            totalVolume
        }
    };
}

async function saveCustomerRecord(entry) {
    const customerHistory = getCustomerHistory();
    customerHistory.unshift(entry);
    saveCustomerHistory(customerHistory);
    renderCustomerHistory();
    return appendCustomerToExcel(entry, { autoConnect: true });
}

async function saveCustomerDataOnly() {
    const history = getHistory();
    const result = document.getElementById("result");

    if (history.length === 0) {
        result.textContent = "No data to save. Add at least one calculation.";
        return;
    }

    const total = getTotalVolume(history);
    const { customerRecord } = createCustomerRecord(total);
    const isSaved = await saveCustomerRecord(customerRecord);

    if (isSaved) {
        result.textContent = "Customer data saved to Excel successfully.";
    } else {
        result.textContent = "Saved in app history. Excel save failed; check Excel status.";
    }
}

async function shareOnWhatsApp() {
    const history = getHistory();
    const result = document.getElementById("result");

    if (history.length === 0) {
        result.textContent = "No data to share. Add at least one calculation.";
        return;
    }

    const total = getTotalVolume(history);
    const { customerName, customerAddress, cleanPhone, customerRecord } = createCustomerRecord(total);

    const lines = history.map((item, index) =>
        `${index + 1}. Circumference: ${item.circumference}, Height: ${item.height}, Volume: ${item.volume.toFixed(3)} ft^3`
    );

    const message = [
        `Hello ${customerName},`,
        "Cylinder Volume Report",
        ...lines,
        `Total Volume: ${total.toFixed(3)} ft^3`,
        `Address: ${customerAddress}`,
        "",
        "Thank you for being a valued customer of my business."
    ].join("\n");

    const encodedMessage = encodeURIComponent(message);
    const url = cleanPhone
        ? `https://wa.me/${cleanPhone}?text=${encodedMessage}`
        : `https://wa.me/?text=${encodedMessage}`;

    await saveCustomerRecord(customerRecord);

    window.open(url, "_blank");
}

function deleteEntry(index) {
    const history = getHistory();
    if (index < 0 || index >= history.length) return;
    history.splice(index, 1);
    saveHistory(history);
    renderHistory();
}

function clearHistory() {
    localStorage.removeItem(STORAGE_KEY);
    renderHistory();
}

function deleteCustomerHistory(id) {
    const history = getCustomerHistory().filter(item => item.id !== id);
    saveCustomerHistory(history);
    renderCustomerHistory();
}

function calculateVolume() {
    let cFeet = parseFloat(document.getElementById("cFeet").value) || 0;
    let cInches = parseFloat(document.getElementById("cInches").value) || 0;
    let hFeet = parseFloat(document.getElementById("hFeet").value) || 0;
    let hInches = parseFloat(document.getElementById("hInches").value) || 0;

    let circumference = cFeet + (cInches / 12);
    let height = hFeet + (hInches / 12);

    if (circumference <= 0 || height <= 0) {
        document.getElementById("result").innerHTML = "Please enter valid values.";
        return;
    }

    let radius = circumference / (2 * Math.PI);
    let volume = Math.PI * radius * radius * height;

    const output = "Volume = " + volume.toFixed(3) + " cubic feet";
    document.getElementById("result").innerHTML = output;

    const entry = {
        circumference: `${cFeet}ft ${cInches}in`,
        height: `${hFeet}ft ${hInches}in`,
        volume
    };
    const history = getHistory();
    history.unshift(entry);
    saveHistory(history);
    renderHistory();
}

renderHistory();
renderCustomerHistory();
restoreExcelConnectionStatus();
