
"use strict";

const state = {
    currentStep: 0,
    file: null,
    headers: [],
    records: [],
    selectedObjects: []
};

const panels = [...document.querySelectorAll(".panel")];
const steps = [...document.querySelectorAll(".step")];

const csvFile = document.getElementById("csvFile");
const fileInfo = document.getElementById("fileInfo");
const filePreview = document.getElementById("filePreview");
const previewTable = document.getElementById("previewTable");
const previewInfo = document.getElementById("previewInfo");
const nextDatasetButton = document.getElementById("nextDatasetButton");
const nextObjectButton = document.getElementById("nextObjectButton");
const objectList = document.getElementById("objectList");
const selectionCount = document.getElementById("selectionCount");
const parameterTable = document.getElementById("parameterTable");
const parameterCount = document.getElementById("parameterCount");
const selectedObject = document.getElementById("selectedObject");

function showStep(step) {
    if (step < 0 || step >= panels.length) return;

    state.currentStep = step;

    panels.forEach((panel, index) => {
        panel.classList.toggle("active", index === step);
    });

    steps.forEach((button, index) => {
        button.classList.toggle("active", index === step);
        button.classList.toggle("done", index < step);
    });

    window.scrollTo({
        top: 0,
        behavior: "smooth"
    });
}

steps.forEach(button => {
    button.addEventListener("click", () => {
        const target = Number(button.dataset.step);

        if (target <= state.currentStep) {
            showStep(target);
        }
    });
});

function clearError(container) {
    container.querySelector(".error")?.remove();
}

function showError(container, message) {
    clearError(container);

    const error = document.createElement("div");
    error.className = "error";
    error.textContent = message;

    container.prepend(error);

    error.scrollIntoView({
        behavior: "smooth",
        block: "center"
    });
}

// Membaca file CSV
function parseCSV(text) {
    const rows = [];
    let row = [];
    let value = "";
    let quoted = false;

    text = text.replace(/^\uFEFF/, "");

    for (let i = 0; i < text.length; i++) {
        const char = text[i];

        if (char === '"') {
            if (quoted && text[i + 1] === '"') {
                value += '"';
                i++;
            } else {
                quoted = !quoted;
            }
        } else if (char === "," && !quoted) {
            row.push(value.trim());
            value = "";
        } else if (
            (char === "\n" || char === "\r") &&
            !quoted
        ) {
            if (char === "\r" && text[i + 1] === "\n") {
                i++;
            }

            row.push(value.trim());

            if (row.some(cell => cell !== "")) {
                rows.push(row);
            }

            row = [];
            value = "";
        } else {
            value += char;
        }
    }

    if (quoted) {
        throw new Error(
            "Format CSV tidak valid: tanda kutip belum ditutup."
        );
    }

    row.push(value.trim());

    if (row.some(cell => cell !== "")) {
        rows.push(row);
    }

    return rows;
}

// Membaca file Excel dan mengubahnya menjadi array
function parseExcel(arrayBuffer) {
    if (typeof XLSX === "undefined") {
        throw new Error(
            "Library Excel belum dimuat. Periksa koneksi internet dan library SheetJS."
        );
    }

    const workbook = XLSX.read(arrayBuffer, {
        type: "array"
    });

    if (!workbook.SheetNames.length) {
        throw new Error(
            "File Excel tidak memiliki lembar kerja."
        );
    }

    // Mengambil lembar kerja pertama
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    // Mengubah isi lembar kerja menjadi array dua dimensi
    let rows = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: false,
        blankrows: false
    });

    // Menghapus baris yang sepenuhnya kosong
    rows = rows.filter(row =>
        row.some(value => String(value).trim() !== "")
    );

    return rows;
}

// Validasi dataset CSV maupun Excel
function validateCSV(rows) {
    if (rows.length < 2) {
        throw new Error(
            "Dataset harus memiliki header dan minimal satu objek."
        );
    }

    const headers = rows[0].map(header =>
        String(header).trim()
    );

    if (headers.length < 2) {
        throw new Error(
            "Dataset harus memiliki kolom objek dan minimal satu parameter."
        );
    }

    if (headers.some(header => !header)) {
        throw new Error(
            "Nama kolom tidak boleh kosong."
        );
    }

    if (new Set(headers).size !== headers.length) {
        throw new Error(
            "Nama kolom tidak boleh duplikat."
        );
    }

    const records = rows.slice(1).map((row, index) => {
        if (row.length > headers.length) {
            throw new Error(
                `Baris ${index + 2} memiliki kolom lebih banyak daripada header.`
            );
        }

        return headers.map((_, columnIndex) =>
            String(row[columnIndex] ?? "").trim()
        );
    });

    const objects = records.map(record => record[0]);

    if (objects.some(object => !object)) {
        throw new Error(
            "Nama objek pada kolom pertama tidak boleh kosong."
        );
    }

    if (new Set(objects).size !== objects.length) {
        throw new Error(
            "Nama objek pada kolom pertama harus unik."
        );
    }

    return { headers, records };
}

// Menampilkan pratinjau dataset
function renderPreview() {
    const thead = previewTable.querySelector("thead");
    const tbody = previewTable.querySelector("tbody");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    const headerRow = document.createElement("tr");

    state.headers.forEach(header => {
        const th = document.createElement("th");
        th.textContent = header;
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);

    state.records.forEach(record => {
        const tr = document.createElement("tr");

        record.forEach(value => {
            const td = document.createElement("td");
            td.textContent = value;
            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    previewInfo.textContent =
        `${state.records.length} objek × ${state.headers.length} kolom`;

    filePreview.classList.remove("hidden");
}

// Upload CSV atau Excel
csvFile.addEventListener("change", () => {
    const panel = document.querySelector('[data-panel="0"]');

    clearError(panel);

    state.file = csvFile.files[0] || null;
    state.headers = [];
    state.records = [];
    state.selectedObjects = [];

    nextDatasetButton.disabled = true;
    fileInfo.classList.add("hidden");
    filePreview.classList.add("hidden");
    objectList.innerHTML = "";

    selectionCount.textContent = "Belum ada objek dipilih.";

    if (!state.file) return;

    const fileName = state.file.name.toLowerCase();

    const isCSV = fileName.endsWith(".csv");

    const isExcel =
        fileName.endsWith(".xlsx") ||
        fileName.endsWith(".xls");

    if (!isCSV && !isExcel) {
        showError(
            panel,
            "Format file harus CSV, XLSX, atau XLS."
        );

        csvFile.value = "";
        state.file = null;
        return;
    }

    const reader = new FileReader();

    reader.onload = event => {
        try {
            let rows;

            if (isCSV) {
                rows = parseCSV(event.target.result);
            } else {
                rows = parseExcel(event.target.result);
            }

            const dataset = validateCSV(rows);

            state.headers = dataset.headers;
            state.records = dataset.records;

            fileInfo.textContent =
                `File dipilih: ${state.file.name} · ` +
                `${state.records.length} objek · ` +
                `${state.headers.length - 1} parameter`;

            fileInfo.classList.remove("hidden");

            renderPreview();
            renderObjects();

            nextDatasetButton.disabled = false;

        } catch (error) {
            showError(panel, error.message);

            state.headers = [];
            state.records = [];
            state.selectedObjects = [];

            objectList.innerHTML = "";
            filePreview.classList.add("hidden");
            nextDatasetButton.disabled = true;
        }
    };

    reader.onerror = () => {
        showError(panel, "File gagal dibaca.");
        nextDatasetButton.disabled = true;
    };

    if (isExcel) {
        reader.readAsArrayBuffer(state.file);
    } else {
        reader.readAsText(state.file, "UTF-8");
    }
});

// Menampilkan daftar objek yang dapat dipilih
function renderObjects() {
    objectList.innerHTML = "";

    state.records.forEach((record, index) => {
        const label = document.createElement("label");
        label.className = "object-card";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.className = "object-checkbox";
        checkbox.value = String(index);

        checkbox.checked =
            state.selectedObjects.includes(index);

        const name = document.createElement("span");
        name.className = "object-name";
        name.textContent = record[0];

        checkbox.addEventListener("change", () => {
            if (checkbox.checked) {
                if (!state.selectedObjects.includes(index)) {
                    state.selectedObjects.push(index);
                }
            } else {
                state.selectedObjects =
                    state.selectedObjects.filter(
                        selectedIndex => selectedIndex !== index
                    );
            }

            updateSelectionCount();
        });

        label.append(checkbox, name);
        objectList.appendChild(label);
    });

    updateSelectionCount();
}

function updateSelectionCount() {
    const count = state.selectedObjects.length;

    selectionCount.textContent = count
        ? `${count} objek dipilih.`
        : "Belum ada objek dipilih.";
}

// Mengubah nilai menjadi angka jika memungkinkan
function parseNumericValue(raw) {
    const text = String(raw ?? "").trim();

    if (!text) return null;

    const match = text.match(
        /^(Rp\s*)?([-+]?\d[\d.,]*)(.*?)$/i
    );

    if (!match) return null;

    const prefix = match[1] ? "Rp" : "";
    let numberText = match[2];
    const suffix = match[3].trim();

    if (
        numberText.includes(".") &&
        numberText.includes(",")
    ) {
        if (
            numberText.lastIndexOf(",") >
            numberText.lastIndexOf(".")
        ) {
            numberText = numberText
                .replace(/\./g, "")
                .replace(",", ".");
        } else {
            numberText = numberText.replace(/,/g, "");
        }
    } else if (numberText.includes(",")) {
        const parts = numberText.split(",");

        numberText =
            parts.length === 2 && parts[1].length !== 3
                ? parts[0] + "." + parts[1]
                : parts.join("");
    } else if (numberText.includes(".")) {
        const parts = numberText.split(".");

        numberText =
            parts.length > 1 &&
            parts.slice(1).every(part => part.length === 3)
                ? parts.join("")
                : numberText;
    }

    const number = Number(numberText);

    if (!Number.isFinite(number)) return null;

    return { number, prefix, suffix };
}

function formatNumber(number) {
    return new Intl.NumberFormat("id-ID", {
        maximumFractionDigits: 4
    }).format(number);
}

// Merangkum nilai parameter dari objek terpilih
function summarizeValues(values) {
    const nonEmpty = values
        .map(value => String(value ?? "").trim())
        .filter(value => value !== "");

    if (!nonEmpty.length) return "—";

    const parsed = nonEmpty.map(parseNumericValue);

    const allNumeric = parsed.every(
        value => value !== null
    );

    if (!allNumeric) {
        return [...new Set(nonEmpty)].join(", ");
    }

    const first = parsed[0];

    const sameUnit = parsed.every(value =>
        value.prefix === first.prefix &&
        value.suffix === first.suffix
    );

    if (!sameUnit) {
        return [...new Set(nonEmpty)].join(", ");
    }

    const min = Math.min(
        ...parsed.map(value => value.number)
    );

    const max = Math.max(
        ...parsed.map(value => value.number)
    );

    const prefix = first.prefix;
    const suffix = first.suffix
        ? ` ${first.suffix}`
        : "";

    if (min === max) {
        return `${prefix}${formatNumber(min)}${suffix}`;
    }

    return (
        `${prefix}${formatNumber(min)}–` +
        `${prefix}${formatNumber(max)}${suffix}`
    );
}

// Menghasilkan ringkasan parameter
function generateResult() {
    const panel = document.querySelector('[data-panel="1"]');

    clearError(panel);

    if (!state.selectedObjects.length) {
        showError(
            panel,
            "Pilih minimal satu objek terlebih dahulu."
        );

        return false;
    }

    const selectedRecords = state.selectedObjects.map(
        index => state.records[index]
    );

    parameterTable.innerHTML = "";

    state.headers.slice(1).forEach((parameter, index) => {
        const tr = document.createElement("tr");

        const numberCell = document.createElement("td");
        numberCell.textContent = String(index + 1);

        const parameterCell = document.createElement("td");
        parameterCell.textContent = parameter;

        const valueCell = document.createElement("td");

        const values = selectedRecords.map(
            record => record[index + 1]
        );

        valueCell.textContent = summarizeValues(values);

        tr.append(numberCell, parameterCell, valueCell);
        parameterTable.appendChild(tr);
    });

    selectedObject.textContent = selectedRecords
        .map(record => record[0])
        .join(", ");

    parameterCount.textContent =
        `${state.headers.length - 1} parameter`;

    return true;
}

// Tombol menuju tahap pemilihan objek
nextDatasetButton.addEventListener("click", () => {
    const panel = document.querySelector('[data-panel="0"]');

    clearError(panel);

    if (!state.records.length) {
        showError(
            panel,
            "Unggah dataset terlebih dahulu."
        );

        return;
    }

    showStep(1);
});

// Tombol menuju tahap ringkasan parameter
nextObjectButton.addEventListener("click", () => {
    if (generateResult()) {
        showStep(2);
    }
});

// Tombol kembali ke tahap sebelumnya
document.querySelectorAll(".prev").forEach(button => {
    button.addEventListener("click", () => {
        showStep(state.currentStep - 1);
    });
});

// Mengatur ulang aplikasi
document.getElementById("restartBtn").addEventListener("click", () => {
    state.currentStep = 0;
    state.file = null;
    state.headers = [];
    state.records = [];
    state.selectedObjects = [];

    csvFile.value = "";

    fileInfo.textContent = "";
    fileInfo.classList.add("hidden");

    filePreview.classList.add("hidden");
    previewInfo.textContent = "";

    previewTable.querySelector("thead").innerHTML = "";
    previewTable.querySelector("tbody").innerHTML = "";

    objectList.innerHTML = "";
    parameterTable.innerHTML = "";

    selectedObject.textContent = "-";
    parameterCount.textContent = "0 parameter";
    selectionCount.textContent = "Belum ada objek dipilih.";

    nextDatasetButton.disabled = true;

    panels.forEach(panel => clearError(panel));

    showStep(0);
});

// Menampilkan tahap pertama saat aplikasi dibuka
showStep(0);