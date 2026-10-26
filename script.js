const state = {
    currentStep: 0,
    file: null,
    data: null,
    selectedParameters: [],
    rules: {},
    softSet: {},
    weights: {},
    result: [],
    best: [],
    maxScore: 0
};

const panels = [...document.querySelectorAll(".panel")];
const steps = [...document.querySelectorAll(".step")];

const csvFile = document.getElementById("csvFile");
const uploadBtn = document.getElementById("uploadBtn");
const fileInfo = document.getElementById("fileInfo");
const filePreview = document.getElementById("filePreview");
const previewTable = document.getElementById("previewTable");
const previewInfo = document.getElementById("previewInfo");

function toggleImage() {
    const image = document.getElementById("csvImage");

    if (!image) return;

    if (image.style.display === "none") {
        image.style.display = "block";
    } else {
        image.style.display = "none";
    }
}

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

    window.scrollTo({ top: 0, behavior: "smooth" });
}

steps.forEach(button => {
    button.addEventListener("click", () => {
        const target = Number(button.dataset.step);
        if (target <= state.currentStep) showStep(target);
    });
});

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function formatScore(score) {
    return `${(score * 100).toFixed(1)}%`;
}

function clearError(container) {
    container?.querySelector(".error")?.remove();
}

function showError(container, message) {
    if (!container) return;

    clearError(container);

    const error = document.createElement("div");
    error.className = "error";
    error.textContent = message;
    container.prepend(error);
}

function parseCSV(text) {
    text = text.replace(/^\uFEFF/, "");

    const rows = [];
    let row = [];
    let value = "";
    let quoted = false;

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
        } else if ((char === "\n" || char === "\r") && !quoted) {
            if (char === "\r" && text[i + 1] === "\n") i++;

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
        throw new Error("Format CSV tidak valid: tanda kutip belum ditutup.");
    }

    row.push(value.trim());

    if (row.some(cell => cell !== "")) {
        rows.push(row);
    }

    return rows;
}

function parseExcel(arrayBuffer) {
    if (typeof XLSX === "undefined") {
        throw new Error(
            "Pustaka Excel belum dimuat. Periksa koneksi internet."
        );
    }

    const workbook = XLSX.read(arrayBuffer, { type: "array" });

    if (!workbook.SheetNames.length) {
        throw new Error("File Excel tidak memiliki lembar kerja.");
    }

    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    return XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: "",
        raw: false,
        blankrows: false
    }).filter(row =>
        row.some(value => String(value).trim() !== "")
    );
}

function validateRows(rows) {
    if (!Array.isArray(rows) || rows.length < 2) {
        throw new Error(
            "Dataset harus memiliki header dan minimal satu baris data."
        );
    }

    const headers = rows[0].map(value =>
        String(value ?? "").trim()
    );

    if (headers.length < 2) {
        throw new Error(
            "Dataset harus memiliki kolom objek dan minimal satu parameter."
        );
    }

    if (headers.some(header => !header)) {
        throw new Error("Semua kolom harus memiliki nama.");
    }

    if (new Set(headers).size !== headers.length) {
        throw new Error("Nama kolom tidak boleh duplikat.");
    }

    const records = rows.slice(1).map((row, index) => {
        if (row.length > headers.length) {
            throw new Error(
                `Baris ${index + 2} memiliki kolom melebihi jumlah header.`
            );
        }

        return headers.map((_, columnIndex) =>
            String(row[columnIndex] ?? "").trim()
        );
    });

    if (!records.length) {
        throw new Error("Dataset tidak memiliki objek.");
    }

    const objects = records.map(record => record[0]);

    if (objects.some(object => !object)) {
        throw new Error("Nama objek pada kolom pertama tidak boleh kosong.");
    }

    if (new Set(objects).size !== objects.length) {
        throw new Error("Nama objek pada kolom pertama harus unik.");
    }

    const data = records.map(record =>
        Object.fromEntries(
            headers.map((header, index) => [header, record[index]])
        )
    );

    const parameters = headers.slice(1).map(name => {
        const values = records
            .map(record => record[headers.indexOf(name)])
            .filter(value => value !== "");

        const numbers = values.map(Number);
        const numeric = values.length > 0 &&
            numbers.every(value => Number.isFinite(value));

        return {
            name,
            type: numeric ? "numeric" : "categorical",
            min: numeric ? Math.min(...numbers) : null,
            max: numeric ? Math.max(...numbers) : null,
            values: numeric ? [] : [...new Set(values)]
        };
    });

    return {
        data,
        parameters,
        objects,
        rows: records.length,
        columns: parameters.length,
        missing: records.reduce(
            (sum, record) =>
                sum + record.filter(value => value === "").length,
            0
        ),
        object_column: headers[0],
        headers
    };
}

function renderPreview(rows) {
    const thead = previewTable.querySelector("thead");
    const tbody = previewTable.querySelector("tbody");

    thead.innerHTML = "";
    tbody.innerHTML = "";

    if (!rows.length) {
        previewInfo.textContent = "Dataset kosong.";
        return;
    }

    const headers = rows[0];
    const headerRow = document.createElement("tr");

    headers.forEach(header => {
        const th = document.createElement("th");
        th.textContent = String(header ?? "");
        headerRow.appendChild(th);
    });

    thead.appendChild(headerRow);

    rows.slice(1, 11).forEach(row => {
        const tr = document.createElement("tr");

        headers.forEach((_, index) => {
            const td = document.createElement("td");
            td.textContent = row[index] ?? "";
            tr.appendChild(td);
        });

        tbody.appendChild(tr);
    });

    previewInfo.textContent =
        `Menampilkan ${Math.min(10, rows.length - 1)} dari ${rows.length - 1} objek.`;
}

csvFile.addEventListener("change", async () => {
    const panel = document.querySelector('[data-panel="0"]');

    clearError(panel);

    state.file = null;
    state.data = null;
    uploadBtn.disabled = true;
    fileInfo.textContent = "";
    fileInfo.classList.add("hidden");
    filePreview.classList.add("hidden");

    const file = csvFile.files?.[0];
    if (!file) return;

    const extension = file.name.split(".").pop().toLowerCase();

    if (!["csv", "xlsx", "xls"].includes(extension)) {
        showError(panel, "Pilih file dengan format CSV, XLSX, atau XLS.");
        csvFile.value = "";
        return;
    }

    try {
        let rows;

        if (extension === "csv") {
            rows = parseCSV(await file.text());
        } else {
            rows = parseExcel(await file.arrayBuffer());
        }

        const dataset = validateRows(rows);

        state.file = file;
        state.data = dataset;

        renderPreview(rows);
        filePreview.classList.remove("hidden");

        fileInfo.textContent =
            `${file.name} — ${dataset.rows} objek, ` +
            `${dataset.columns} parameter`;

        fileInfo.classList.remove("hidden");
        uploadBtn.disabled = false;
    } catch (error) {
        state.file = null;
        state.data = null;
        showError(panel, error.message || "File gagal dibaca.");
    }
});

uploadBtn.addEventListener("click", () => {
    const panel = document.querySelector('[data-panel="0"]');

    clearError(panel);

    if (!state.file || !state.data) {
        showError(panel, "Pilih file dan pastikan dataset berhasil dibaca.");
        return;
    }

    state.selectedParameters = [];
    state.rules = {};
    state.softSet = {};
    state.weights = {};
    state.result = [];
    state.best = [];
    state.maxScore = 0;

    renderDatasetInfo();
    renderParameters();
    showStep(1);
});

function renderDatasetInfo() {
    const stats = document.getElementById("datasetStats");

    stats.innerHTML = `
        <div class="stat">
            <div class="number">${state.data.rows}</div>
            <div class="label">Jumlah Objek</div>
        </div>
        <div class="stat">
            <div class="number">${state.data.columns}</div>
            <div class="label">Jumlah Parameter</div>
        </div>
        <div class="stat">
            <div class="number">${state.data.missing}</div>
            <div class="label">Data Kosong</div>
        </div>
    `;

    document.getElementById("universeList").innerHTML =
        state.data.objects.map(object => `
            <span class="tag">${escapeHtml(object)}</span>
        `).join("");
}

function renderParameters() {
    const container = document.getElementById("parameterList");

    container.innerHTML = state.data.parameters.map(parameter => `
        <div class="parameter-card">
            <label class="check-row">
                <input
                    type="checkbox"
                    class="parameter-checkbox"
                    value="${escapeHtml(parameter.name)}"
                >
                <strong>${escapeHtml(parameter.name)}</strong>
            </label>
            <div class="type">
                Tipe:
                ${parameter.type === "numeric" ? "Numerik" : "Kategorikal"}
            </div>
            <div class="type">
                ${
                    parameter.type === "numeric"
                        ? `Rentang ${parameter.min}–${parameter.max}`
                        : `Kategori: ${parameter.values.map(escapeHtml).join(", ")}`
                }
            </div>
        </div>
    `).join("");
}

function renderRules() {
    const container = document.getElementById("rulesContainer");
    container.innerHTML = "";

    state.selectedParameters.forEach(name => {
        const parameter = state.data.parameters.find(
            item => item.name === name
        );

        if (!parameter) return;

        if (parameter.type === "numeric") {
            container.insertAdjacentHTML("beforeend", `
                <div class="rule-card"
                     data-parameter="${escapeHtml(parameter.name)}">
                    <h3>${escapeHtml(parameter.name)}</h3>

                    <div class="rule-controls">
                        <select class="operator">
                            <option value="≥">≥</option>
                            <option value="≤">≤</option>
                            <option value=">">&gt;</option>
                            <option value="<">&lt;</option>
                        </select>

                        <input
                            type="number"
                            class="numeric-value"
                            min="${parameter.min}"
                            max="${parameter.max}"
                            value="${parameter.min}"
                            step="any"
                            required
                        >
                    </div>

                    <div class="type">
                        Rentang data: ${parameter.min}–${parameter.max}
                    </div>
                </div>
            `);
        } else {
            container.insertAdjacentHTML("beforeend", `
                <div class="rule-card"
                     data-parameter="${escapeHtml(parameter.name)}">
                    <h3>${escapeHtml(parameter.name)}</h3>

                    <div class="category-values">
                        ${parameter.values.map(value => `
                            <label>
                                <input
                                    type="checkbox"
                                    class="category-checkbox"
                                    value="${escapeHtml(value)}"
                                >
                                ${escapeHtml(value)}
                            </label>
                        `).join("")}
                    </div>
                </div>
            `);
        }
    });
}

function collectRules() {
    const rules = {};
    const cards = [
        ...document.querySelectorAll('[data-panel="3"] .rule-card')
    ];

    for (const card of cards) {
        const parameter = card.dataset.parameter;
        const numeric = card.querySelector(".numeric-value");

        if (numeric) {
            if (
                numeric.value.trim() === "" ||
                !Number.isFinite(Number(numeric.value))
            ) {
                throw new Error(
                    `Masukkan batas numerik untuk parameter ${parameter}.`
                );
            }

            const value = Number(numeric.value);
            const min = Number(numeric.min);
            const max = Number(numeric.max);

            if (value < min || value > max) {
                throw new Error(
                    `Batas ${parameter} harus berada antara ${min} dan ${max}.`
                );
            }

            rules[parameter] = {
                type: "numeric",
                operator: card.querySelector(".operator").value,
                value
            };
        } else {
            const values = [
                ...card.querySelectorAll(".category-checkbox:checked")
            ].map(input => input.value);

            if (!values.length) {
                throw new Error(
                    `Pilih minimal satu kategori untuk ${parameter}.`
                );
            }

            rules[parameter] = {
                type: "categorical",
                values
            };
        }
    }

    return rules;
}

function satisfiesRule(value, rule) {
    if (
        value === null ||
        value === undefined ||
        String(value).trim() === ""
    ) {
        return false;
    }

    if (rule.type === "numeric") {
        const number = Number(value);

        if (!Number.isFinite(number)) return false;

        switch (rule.operator) {
            case "≥":
                return number >= rule.value;
            case "≤":
                return number <= rule.value;
            case ">":
                return number > rule.value;
            case "<":
                return number < rule.value;
            default:
                return false;
        }
    }

    return rule.values.includes(String(value).trim());
}

function buildSoftSet() {
    const objectColumn = state.data.object_column;
    const rows = state.data.data;

    if (!Array.isArray(rows) || !rows.length) {
        throw new Error("Dataset tidak memiliki objek.");
    }

    const softSet = {};

    for (const parameter of state.selectedParameters) {
        const rule = state.rules[parameter];

        if (!rule) {
            throw new Error(
                `Aturan parameter ${parameter} belum tersedia.`
            );
        }

        softSet[parameter] = rows
            .filter(row => satisfiesRule(row[parameter], rule))
            .map(row => String(row[objectColumn]));
    }

    state.softSet = softSet;
}

function renderSoftSet() {
    document.getElementById("softSetTable").innerHTML =
        Object.entries(state.softSet).map(([parameter, members]) => `
            <tr>
                <td>${escapeHtml(parameter)}</td>
                <td>
                    ${
                        members.length
                            ? members.map(escapeHtml).join(", ")
                            : "∅"
                    }
                </td>
            </tr>
        `).join("");
}

function renderWeights() {
    const container = document.getElementById("weightContainer");

    container.innerHTML = `
        <div class="card">
            <h3>Bobot Parameter</h3>
            <p>
                Masukkan bobot setiap parameter dalam rentang 0 sampai 100.
                Bobot akan dinormalisasi saat menghitung score.
            </p>
            <div id="weightList" class="parameter-grid"></div>
            <div id="weightTotal" class="type"></div>
        </div>
    `;

    const weightList = container.querySelector("#weightList");

    weightList.innerHTML = state.selectedParameters.map(parameter => `
        <div class="parameter-card">
            <label>
                <strong>${escapeHtml(parameter)}</strong>
            </label>

            <div class="rule-controls">
                <input
                    type="number"
                    class="weight-input"
                    data-parameter="${escapeHtml(parameter)}"
                    min="0"
                    max="100"
                    step="any"
                    value="${state.weights[parameter] ?? 1}"
                    required
                >
            </div>
        </div>
    `).join("");

    weightList.querySelectorAll(".weight-input").forEach(input => {
        input.addEventListener("input", updateWeightTotal);
    });

    updateWeightTotal();
}

function updateWeightTotal() {
    const inputs = [
        ...document.querySelectorAll('[data-panel="5"] .weight-input')
    ];

    const total = inputs.reduce((sum, input) => {
        const value = input.value.trim() === ""
            ? 0
            : Number(input.value);

        return sum + (Number.isFinite(value) ? value : 0);
    }, 0);

    const totalElement = document.getElementById("weightTotal");

    if (totalElement) {
        totalElement.textContent = `Total bobot: ${total.toFixed(2)}`;
    }
}

function collectWeights() {
    const weights = {};
    const inputs = [
        ...document.querySelectorAll('[data-panel="5"] .weight-input')
    ];

    let total = 0;

    for (const input of inputs) {
        const parameter = input.dataset.parameter;
        const rawValue = input.value.trim();
        const value = Number(rawValue);

        if (
            rawValue === "" ||
            !Number.isFinite(value) ||
            value < 0 ||
            value > 100
        ) {
            throw new Error(
                `Bobot ${parameter} harus berada antara 0 dan 100.`
            );
        }

        weights[parameter] = value;
        total += value;
    }

    if (total <= 0) {
        throw new Error("Total bobot harus lebih besar dari 0.");
    }

    state.weights = weights;
}

function calculateRanking() {
    const rows = state.data.data;
    const objectColumn = state.data.object_column;

    const totalWeight = state.selectedParameters.reduce(
        (sum, parameter) => sum + state.weights[parameter],
        0
    );

    if (totalWeight <= 0) {
        throw new Error("Total bobot harus lebih besar dari 0.");
    }

    const result = rows.map((row, index) => {
        const name = String(row[objectColumn]);

        const satisfied = state.selectedParameters.filter(
            parameter =>
                state.weights[parameter] > 0 &&
                state.softSet[parameter].includes(name)
        );

        const weightedScore = satisfied.reduce(
            (sum, parameter) => sum + state.weights[parameter],
            0
        );

        return {
            Nama: name,
            Score: weightedScore / totalWeight,
            "Parameter Terpenuhi": satisfied,
            _originalIndex: index
        };
    });

    result.sort((a, b) =>
        b.Score - a.Score ||
        a._originalIndex - b._originalIndex
    );

    state.result = result.map((item, index) => ({
        Peringkat: index + 1,
        Nama: item.Nama,
        Score: item.Score,
        "Parameter Terpenuhi": item["Parameter Terpenuhi"]
    }));

    state.maxScore = state.result.length
        ? state.result[0].Score
        : 0;

    state.best = state.result.filter(
        item => Math.abs(item.Score - state.maxScore) < 1e-10
    );
}

function renderRanking() {
    document.getElementById("rankingTable").innerHTML =
        state.result.map(item => `
            <tr>
                <td>${item.Peringkat}</td>
                <td><strong>${escapeHtml(item.Nama)}</strong></td>
                <td>${formatScore(item.Score)}</td>
                <td>
                    ${
                        item["Parameter Terpenuhi"].length
                            ? item["Parameter Terpenuhi"]
                                .map(escapeHtml)
                                .join(", ")
                            : "Tidak ada"
                    }
                </td>
            </tr>
        `).join("");
}

function renderRecommendations() {
    const summary = document.getElementById("recommendationSummary");
    const cards = document.getElementById("recommendationCards");

    summary.innerHTML = `
        <div class="summary">
            <div>Score tertinggi</div>
            <div class="score">${formatScore(state.maxScore)}</div>
            <div>
                Terdapat <strong>${state.best.length}</strong>
                rekomendasi dengan score tertinggi.
            </div>
        </div>
    `;

    cards.innerHTML = state.best.map(item => {
        const originalData = state.data.data.find(
            row =>
                String(row[state.data.object_column]) ===
                String(item.Nama)
        );

        const specifications = originalData
            ? Object.entries(originalData)
                .filter(([key]) => key !== state.data.object_column)
                .map(([key, value]) => `
                    <div class="specification">
                        <span>${escapeHtml(key)}</span>
                        <span>${escapeHtml(value)}</span>
                    </div>
                `)
                .join("")
            : "";

        return `
            <div class="recommendation-card">
                <div class="rank">PERINGKAT ${item.Peringkat}</div>
                <h3>${escapeHtml(item.Nama)}</h3>
                <div class="score">${formatScore(item.Score)}</div>

                <p>
                    Parameter terpenuhi:
                    ${
                        item["Parameter Terpenuhi"].length
                            ? item["Parameter Terpenuhi"]
                                .map(escapeHtml)
                                .join(", ")
                            : "Tidak ada"
                    }
                </p>

                <div class="specifications">
                    <h4>Spesifikasi</h4>
                    ${specifications}
                </div>
            </div>
        `;
    }).join("");
}

document.querySelectorAll(".next").forEach(button => {
    button.addEventListener("click", () => {
        const panel = button.closest(".panel");
        const step = Number(panel.dataset.panel);

        clearError(panel);

        try {
            if (step === 0) {
                if (!state.data) {
                    throw new Error("Upload dataset terlebih dahulu.");
                }

                showStep(1);
                return;
            }

            if (step === 1) {
                showStep(2);
                return;
            }

            if (step === 2) {
                const selected = [
                    ...panel.querySelectorAll(
                        ".parameter-checkbox:checked"
                    )
                ].map(input => input.value);

                if (!selected.length) {
                    throw new Error("Pilih minimal satu parameter.");
                }

                state.selectedParameters = selected;
                state.rules = {};
                state.weights = {};

                renderRules();
                showStep(3);
                return;
            }

            if (step === 3) {
                state.rules = collectRules();
                buildSoftSet();
                renderSoftSet();
                showStep(4);
                return;
            }

            if (step === 4) {
                state.weights = Object.fromEntries(
                    state.selectedParameters.map(parameter => [
                        parameter,
                        state.weights[parameter] ?? 1
                    ])
                );

                renderWeights();
                showStep(5);
                return;
            }

            if (step === 5) {
                collectWeights();
                calculateRanking();
                renderRanking();
                showStep(6);
                return;
            }

            if (step === 6) {
                renderRecommendations();
                showStep(7);
            }
        } catch (error) {
            showError(panel, error.message || "Terjadi kesalahan.");
        }
    });
});

document.querySelectorAll(".prev").forEach(button => {
    button.addEventListener("click", () => {
        showStep(state.currentStep - 1);
    });
});

document.getElementById("restartBtn").addEventListener("click", () => {
    state.currentStep = 0;
    state.file = null;
    state.data = null;
    state.selectedParameters = [];
    state.rules = {};
    state.softSet = {};
    state.weights = {};
    state.result = [];
    state.best = [];
    state.maxScore = 0;

    csvFile.value = "";
    uploadBtn.disabled = true;

    fileInfo.textContent = "";
    fileInfo.classList.add("hidden");

    filePreview.classList.add("hidden");
    previewInfo.textContent = "";
    previewTable.querySelector("thead").innerHTML = "";
    previewTable.querySelector("tbody").innerHTML = "";

    document.getElementById("datasetStats").innerHTML = "";
    document.getElementById("universeList").innerHTML = "";
    document.getElementById("parameterList").innerHTML = "";
    document.getElementById("rulesContainer").innerHTML = "";
    document.getElementById("softSetTable").innerHTML = "";
    document.getElementById("weightContainer").innerHTML = "";
    document.getElementById("rankingTable").innerHTML = "";
    document.getElementById("recommendationSummary").innerHTML = "";
    document.getElementById("recommendationCards").innerHTML = "";

    panels.forEach(panel => clearError(panel));

    showStep(0);
});

showStep(0);
