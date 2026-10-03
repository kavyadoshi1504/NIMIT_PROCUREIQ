// ─── DOM References ────────────────────────────────────────────────────────────

// Purchase Requests
const manualButton        = document.getElementById("manualButton");
const excelButton         = document.getElementById("excelButton");
const manualSection       = document.getElementById("manualSection");
const excelSection        = document.getElementById("excelSection");
const manualForm          = document.getElementById("manualForm");
const qtyInput            = document.getElementById("qty");
const salesRateInput      = document.getElementById("sales_rate");
const taxableValueInput   = document.getElementById("taxable_value");
const excelFile           = document.getElementById("excelFile");
const excelFileName       = document.getElementById("excelFileName");
const importButton        = document.getElementById("importButton");
const excelPreview        = document.getElementById("excelPreview");
const previewBody         = document.getElementById("previewBody");
const saveImportedButton  = document.getElementById("saveImportedButton");
const message             = document.getElementById("message");

// Vendor Masters
const vendorManualButton    = document.getElementById("vendorManualButton");
const vendorExcelButton     = document.getElementById("vendorExcelButton");
const vendorManualSection   = document.getElementById("vendorManualSection");
const vendorExcelSection    = document.getElementById("vendorExcelSection");
const vendorManualForm      = document.getElementById("vendorManualForm");
const vendorExcelFile       = document.getElementById("vendorExcelFile");
const vendorExcelFileName   = document.getElementById("vendorExcelFileName");
const vendorImportButton    = document.getElementById("vendorImportButton");
const vendorExcelPreview    = document.getElementById("vendorExcelPreview");
const vendorPreviewForm     = document.getElementById("vendorPreviewForm");
const saveVendorExcelButton = document.getElementById("saveVendorExcelButton");
const vendorMessage         = document.getElementById("vendorMessage");

// Vendor Inquiries
const vendorInquiriesPage = document.getElementById("vendorInquiriesPage");   // FIX: explicit lookup
const vendorInquiryList   = document.getElementById("vendorInquiryList");

// Purchase Orders
const purchaseOrdersList = document.getElementById("purchaseOrdersList");

// Order Tracking
const orderTrackingList      = document.getElementById("orderTrackingList");
const orderTrackingPrev      = document.getElementById("orderTrackingPrev");
const orderTrackingNext      = document.getElementById("orderTrackingNext");
const orderTrackingPageLabel = document.getElementById("orderTrackingPageLabel");

// Goods Received
const goodsReceivedList    = document.getElementById("goodsReceivedList");
const goodsReceivedMessage = document.getElementById("goodsReceivedMessage");

// Reports & Audits
const reportLogsTabButton = document.getElementById("reportLogsTabButton");
const auditLogsTabButton  = document.getElementById("auditLogsTabButton");
const reportLogsSection   = document.getElementById("reportLogsSection");
const auditLogsSection    = document.getElementById("auditLogsSection");
const reportLogsBody      = document.getElementById("reportLogsBody");
const auditLogsBody       = document.getElementById("auditLogsBody");
const logUsernameFilter   = document.getElementById("logUsernameFilter");
const logFromDate         = document.getElementById("logFromDate");
const logToDate           = document.getElementById("logToDate");
const applyLogFilters     = document.getElementById("applyLogFilters");
const logsMessage         = document.getElementById("logsMessage");
const logsPrev            = document.getElementById("logsPrev");
const logsNext            = document.getElementById("logsNext");
const logsPageLabel       = document.getElementById("logsPageLabel");

// Logout
const logoutButton = document.getElementById("logoutButton");

// ─── State ─────────────────────────────────────────────────────────────────────

let importedRows = [];
let orderTrackingPage = 1;
let orderTrackingTotalPages = 1;
let goodsReceivedOrders = [];
let showCompletedGoods = false;
let activeLogTab    = "report";   // "report" | "audit"
let logsPage         = 1;
let logsTotalPages   = 1;

// ─── Utilities ─────────────────────────────────────────────────────────────────

const AUTH_URL = "http://localhost:5000"; // same value as authServiceUrl on the backend

async function apiFetch(url, options = {}) {
    const response = await fetch(url, { credentials: "include", ...options });
    if (response.status === 401) {
        window.location.href = AUTH_URL;
        return null; // redirect already started
    }
    return response;
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

function showMessage(text) {
    message.textContent = text;
}

function showVendorMessage(text) {
    vendorMessage.textContent = text;
}

function formatDate(value) {
    if (!value) return "-";
    const d = new Date(value);
    return isNaN(d.getTime()) ? String(value) : d.toLocaleDateString("en-IN");
}

function formatDateTime(value) {
    if (!value) return "-";
    const d = new Date(value);
    if (isNaN(d.getTime())) return String(value);
    return d.toLocaleString("en-IN", {
        day: "2-digit", month: "short", year: "numeric",
        hour: "2-digit", minute: "2-digit"
    });
}

function formatCurrency(value) {
    const n = Number(value);
    return isNaN(n) ? "-" : n.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// FIX: sign-aware rupee format  (-500 -> "-₹500.00" instead of "₹-500.00")
function formatRupee(value) {
    const n = Number(value);
    if (isNaN(n)) return "-";
    return `${n < 0 ? "-" : ""}₹${formatCurrency(Math.abs(n))}`;
}

function todayISO() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

// FIX: read username at time of use, not once at script load
function currentUser() {
    return window.currentUsername || "";
}

function prefillOfficeUse() {
    const regDate = document.getElementById("registration_date");
    const rec     = document.getElementById("recommended_by");
    const app     = document.getElementById("approved_by");
    if (regDate && !regDate.value) regDate.value = todayISO();
    if (rec && !rec.value)         rec.value     = currentUser();
    if (app && !app.value)         app.value     = currentUser();
}

// Shared PR info grid used by inquiries / quotations / order tracking / purchase orders
function prInfoItems(x, extraFirst = "") {
    const item = (label, value) => `
        <div class="inquiry-card-item">
            <span class="inquiry-card-label">${label}</span>
            <span class="inquiry-card-value">${value}</span>
        </div>`;

    return `
        ${extraFirst}
        ${item("PR Number", escapeHtml(x.pr_number || "-"))}
        ${item("PR Date", formatDate(x.pr_date))}
        ${item("Party Name", escapeHtml(x.party_name || "-"))}
        ${item("Location", escapeHtml(x.location || "-"))}
        ${item("Territory", escapeHtml(x.territory || "-"))}
        ${item("Product Category", escapeHtml(x.product_category || "-"))}
        ${item("Item Name", escapeHtml(x.item_name || "-"))}
        ${item("Make / Model", `${escapeHtml(x.make || "-")} / ${escapeHtml(x.model || "-")}`)}
        ${item("Quantity", `${escapeHtml(x.qty ?? 0)} ${escapeHtml(x.unit || "")}`)}
        ${item("Sales Rate", formatCurrency(x.sales_rate))}
        ${item("Taxable Value", formatCurrency(x.taxable_value))}
        ${item("Product Remarks", escapeHtml(x.product_remarks || "-"))}
    `;
}

// Prevent scroll from changing number input values
document.addEventListener("wheel", event => {
    if (document.activeElement?.type === "number") {
        document.activeElement.blur();
    }
}, { passive: true });

// ─── Purchase Request: Taxable Value ───────────────────────────────────────────

// FIX: allow rates below 1 (e.g. 0.75)
salesRateInput.min = "0.01";

function calculateTaxableValue() {
    const qty  = Number(qtyInput.value)       || 0;
    const rate = Number(salesRateInput.value) || 0;
    taxableValueInput.value = (qty * rate).toFixed(2);
}

qtyInput.addEventListener("input", calculateTaxableValue);
salesRateInput.addEventListener("input", calculateTaxableValue);

// ─── Sidebar Navigation ────────────────────────────────────────────────────────

document.querySelectorAll(".main-menu").forEach(button => {
    button.addEventListener("click", () => {
        const sectionId = button.dataset.section;
        const section   = document.getElementById(sectionId);

        document.querySelectorAll(".main-menu").forEach(b => b.classList.remove("active"));
        document.querySelectorAll(".submenu").forEach(s => s.classList.remove("open"));

        button.classList.add("active");
        if (section) section.classList.add("open");
    });
});

document.querySelectorAll(".submenu-item").forEach(button => {
    button.addEventListener("click", () => {
        if (button.classList.contains("disabled")) return;

        document.querySelectorAll(".submenu-item").forEach(b => b.classList.remove("active"));
        document.querySelectorAll(".page").forEach(p => p.classList.add("hidden"));

        button.classList.add("active");

        const page = button.dataset.page;

        if (page === "purchase-requests") {
            document.getElementById("purchaseRequestsPage").classList.remove("hidden");
        }

        if (page === "vendor-masters") {
            document.getElementById("vendorMastersPage").classList.remove("hidden");
            setTurnoverYears();
            prefillOfficeUse();
        }

        if (page === "vendor-inquiries") {
            vendorInquiriesPage.classList.remove("hidden");
            loadVendorInquiries();
        }

        if (page === "quotation-comparisons") {
            document.getElementById("quotationComparisonsPage").classList.remove("hidden");
            loadQuotationComparisons();
        }

        if (page === "order-tracking") {
            document.getElementById("orderTrackingPage").classList.remove("hidden");
            orderTrackingPage = 1;
            loadOrderTracking();
        }

        if (page === "purchase-orders") {
            document.getElementById("purchaseOrdersPage").classList.remove("hidden");
            loadPurchaseOrders();
        }

        if (page === "goods-received") {
            document.getElementById("goodsReceivedPage").classList.remove("hidden");
            loadGoodsReceived();
        }

        if (page === "vendor-performance") {
            document.getElementById("vendorPerformancePage").classList.remove("hidden");
            loadVendorPerformance();
        }

        if (page === "past-price-reference") {
            document.getElementById("pastPriceReferencePage").classList.remove("hidden");
            initPastPriceReference();
        }

        if (page === "management-insights") {
            document.getElementById("managementInsightsPage").classList.remove("hidden");
            loadManagementInsights();
        }

        if (page === "reports-audits") {
            document.getElementById("reportsAuditsPage").classList.remove("hidden");
            logsPage = 1;
            loadLogs();
        }
    });
});

// ─── PR Toggle: Manual / Excel ─────────────────────────────────────────────────

manualButton.addEventListener("click", () => {
    manualButton.classList.add("active");
    excelButton.classList.remove("active");
    manualSection.classList.remove("hidden");
    excelSection.classList.add("hidden");
    showMessage("");
});

excelButton.addEventListener("click", () => {
    excelButton.classList.add("active");
    manualButton.classList.remove("active");
    excelSection.classList.remove("hidden");
    manualSection.classList.add("hidden");
    showMessage("");
});

// ─── PR Manual Form Submit ─────────────────────────────────────────────────────

manualForm.addEventListener("submit", async event => {
    event.preventDefault();

    const data = {
        pr_date:          document.getElementById("pr_date").value,
        party_name:       document.getElementById("party_name").value,
        location:         document.getElementById("location").value,
        territory:        document.getElementById("territory").value,
        product_category: document.getElementById("product_category").value,
        item_name:        document.getElementById("item_name").value,
        product_remarks:  document.getElementById("product_remarks").value,
        make:             document.getElementById("make").value,
        model:            document.getElementById("model").value,
        qty:              document.getElementById("qty").value,
        unit:             document.getElementById("unit").value,
        sales_rate:       document.getElementById("sales_rate").value
    };

    const submitBtn = manualForm.querySelector("button[type='submit']");

    try {
        submitBtn.disabled = true;
        showMessage("Saving purchase request...");

        const response = await apiFetch("/purchase-requests", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(data)
        });
        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            showMessage(result.message || "Failed to add purchase request");
            return;
        }

        alert(`PR for "${data.item_name}" created successfully!\nPR Number: ${result.pr_number}`);
        showMessage("");
        manualForm.reset();
        calculateTaxableValue();

    } catch {
        showMessage("Failed to connect to procurement service");
    } finally {
        submitBtn.disabled = false;
    }
});

// ─── PR Excel: File Chosen ─────────────────────────────────────────────────────

excelFile.addEventListener("change", () => {
    excelFileName.textContent = excelFile.files.length
        ? excelFile.files[0].name
        : "No file chosen";
});

// ─── PR Excel: Import Preview ──────────────────────────────────────────────────

importButton.addEventListener("click", async () => {
    if (!excelFile.files.length) {
        showMessage("Please select an Excel file");
        return;
    }

    const formData = new FormData();
    formData.append("file", excelFile.files[0]);

    try {
        importButton.disabled = true;
        showMessage("Processing Excel file...");

        const response = await apiFetch("/purchase-requests/import-preview", {
            method: "POST",
            body: formData
        });
        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            showMessage(result.message || "Failed to process Excel file");
            return;
        }

        importedRows = result.rows || [];

        // FIX: no longer shadows the todayISO() helper
        const today = todayISO();
        importedRows.forEach(row => {
            if (!row.pr_date || String(row.pr_date).trim() === "") {
                row.pr_date = today;
            }
        });

        renderPreview();
        showMessage(`${importedRows.length} rows ready for review. Edit if needed, then click Save.`);

    } catch {
        showMessage("Failed to connect to procurement manager service");
    } finally {
        importButton.disabled = false;
    }
});

// ─── PR Excel: Render Editable Preview Table ───────────────────────────────────

function renderPreview() {
    previewBody.innerHTML = "";

    importedRows.forEach((row, index) => {
        const tr = document.createElement("tr");

        // FIX: pr_date is now a real date input
        tr.innerHTML = `
            <td><input type="date" data-index="${index}" data-field="pr_date" value="${escapeHtml(String(row.pr_date || "").slice(0, 10))}"></td>
            <td><input data-index="${index}" data-field="party_name"       value="${escapeHtml(row.party_name || "")}"></td>
            <td><input data-index="${index}" data-field="location"         value="${escapeHtml(row.location || "")}"></td>
            <td><input data-index="${index}" data-field="territory"        value="${escapeHtml(row.territory || "")}"></td>
            <td><input data-index="${index}" data-field="product_category" value="${escapeHtml(row.product_category || "")}"></td>
            <td><input data-index="${index}" data-field="item_name"        value="${escapeHtml(row.item_name || "")}"></td>
            <td><input data-index="${index}" data-field="product_remarks"  value="${escapeHtml(row.product_remarks || "")}"></td>
            <td><input data-index="${index}" data-field="make"             value="${escapeHtml(row.make || "")}"></td>
            <td><input data-index="${index}" data-field="model"            value="${escapeHtml(row.model || "")}"></td>
            <td><input type="number" data-index="${index}" data-field="qty"        value="${row.qty ?? 0}"        min="0.01" step="0.01"></td>
            <td>
                <select data-index="${index}" data-field="unit">
                    <option value="number" ${(row.unit || "number") === "number" ? "selected" : ""}>Number</option>
                    <option value="set"    ${row.unit === "set"    ? "selected" : ""}>Set</option>
                    <option value="meter"  ${row.unit === "meter"  ? "selected" : ""}>Meter</option>
                    <option value="lot"    ${row.unit === "lot"    ? "selected" : ""}>Lot</option>
                </select>
            </td>
            <td><input type="number" data-index="${index}" data-field="sales_rate" value="${row.sales_rate ?? 0}" min="0.01" step="0.01"></td>
            <td><input class="taxable-input" type="number" value="${calculateRowTaxableValue(row)}" readonly></td>
        `;

        previewBody.appendChild(tr);
    });

    excelPreview.classList.remove("hidden");

    previewBody.querySelectorAll("[data-field]").forEach(input => {
        input.addEventListener("change", updateImportedRow);
        input.addEventListener("input", updateImportedRow);
    });
}

function updateImportedRow(event) {
    const input = event.target;
    const index = Number(input.dataset.index);
    const field = input.dataset.field;

    if (!field) return;

    importedRows[index][field] = input.value;

    if (field === "qty" || field === "sales_rate") {
        const qty    = Number(importedRows[index].qty)        || 0;
        const rate   = Number(importedRows[index].sales_rate) || 0;
        const taxVal = qty * rate;

        importedRows[index].taxable_value = taxVal;

        const taxInput = input.closest("tr").querySelector(".taxable-input");
        if (taxInput) taxInput.value = taxVal.toFixed(2);
    }
}

function calculateRowTaxableValue(row) {
    const qty  = Number(row.qty)        || 0;
    const rate = Number(row.sales_rate) || 0;
    return (qty * rate).toFixed(2);
}

// FIX: validate imported rows before saving; returns error text or ""
function validateImportedRows(rows) {
    const dateRe = /^\d{4}-\d{2}-\d{2}$/;
    const required = ["party_name", "location", "territory", "product_category", "item_name", "make", "model", "unit"];

    for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const n = i + 1;

        if (!dateRe.test(String(r.pr_date || "").trim())) return `Row ${n}: PR date must be YYYY-MM-DD.`;

        for (const f of required) {
            if (!String(r[f] ?? "").trim()) return `Row ${n}: ${f.replace(/_/g, " ")} is required.`;
        }

        if (!(Number(r.qty) > 0))        return `Row ${n}: quantity must be greater than 0.`;
        if (!(Number(r.sales_rate) > 0)) return `Row ${n}: sales rate must be greater than 0.`;
    }
    return "";
}

// ─── PR Excel: Save Imported Rows ─────────────────────────────────────────────

saveImportedButton.addEventListener("click", async () => {
    if (!importedRows.length) {
        showMessage("No rows available to save");
        return;
    }

    const validationError = validateImportedRows(importedRows);
    if (validationError) {
        showMessage(validationError);
        return;
    }

    // FIX: send real numbers, not strings
    const rowsToSend = importedRows.map(r => ({
        ...r,
        qty:           Number(r.qty),
        sales_rate:    Number(r.sales_rate),
        taxable_value: Number(r.qty) * Number(r.sales_rate)
    }));

    try {
        saveImportedButton.disabled = true;
        showMessage("Saving purchase requests...");

        const response = await apiFetch("/purchase-requests/import", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ rows: rowsToSend })
        });
        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            showMessage(result.message || "Failed to save purchase requests");
            return;
        }

        const savedRows = result.rows || [];
        const count     = savedRows.length || importedRows.length;

        if (savedRows.length >= 2) {
            const first = savedRows[0].pr_number;
            const last  = savedRows[savedRows.length - 1].pr_number;
            alert(`${count} items added successfully!\nPR Numbers: ${first} to ${last}`);
        } else if (savedRows.length === 1) {
            alert(`1 item added successfully!\nPR Number: ${savedRows[0].pr_number}`);
        } else {
            alert(`${count} purchase request(s) saved successfully.`);
        }

        showMessage("");
        importedRows = [];
        previewBody.innerHTML = "";
        excelPreview.classList.add("hidden");
        excelFile.value = "";
        excelFileName.textContent = "No file chosen";

    } catch {
        showMessage("Failed to connect to procurement service");
    } finally {
        saveImportedButton.disabled = false;
    }
});

// ─── Vendor Toggle: Manual / Excel ────────────────────────────────────────────

vendorManualButton.addEventListener("click", () => {
    vendorManualButton.classList.add("active");
    vendorExcelButton.classList.remove("active");
    vendorManualSection.classList.remove("hidden");
    vendorExcelSection.classList.add("hidden");
    showVendorMessage("");
});

vendorExcelButton.addEventListener("click", () => {
    vendorExcelButton.classList.add("active");
    vendorManualButton.classList.remove("active");
    vendorExcelSection.classList.remove("hidden");
    vendorManualSection.classList.add("hidden");
    showVendorMessage("");
});

// ─── Vendor Masters: Turnover Year Labels ──────────────────────────────────────

function setTurnoverYears() {
    const now       = new Date();
    const month     = now.getMonth() + 1;
    const year      = now.getFullYear();
    const startYear = month >= 4 ? year : year - 1;

    [
        { id: "turnover_value_1", text: `${String(startYear).slice(-2)}-${String(startYear + 1).slice(-2)} Turnover` },
        { id: "turnover_value_2", text: `${String(startYear - 1).slice(-2)}-${String(startYear).slice(-2)} Turnover` },
        { id: "turnover_value_3", text: `${String(startYear - 2).slice(-2)}-${String(startYear - 1).slice(-2)} Turnover` }
    ].forEach(({ id, text }) => {
        const input = document.getElementById(id);
        if (!input) return;
        const label = input.closest(".form-group")?.querySelector("label");
        if (label) label.textContent = text;
    });
}

setTurnoverYears();
prefillOfficeUse();

// ─── Vendor Masters: GST Duplicate Check ──────────────────────────────────────

async function checkVendorGST(gstNumber) {
    if (!gstNumber?.trim()) return true;

    try {
        const response = await apiFetch(`/vendors/check-gst?gst_number=${encodeURIComponent(gstNumber.trim())}`);
        if (!response) return false;
        const result   = await response.json();

        if (!response.ok) {
            showVendorMessage(result.message || "Failed to check GST number");
            return false;
        }

        if (result.exists) {
            alert(`A company with this GST number already exists: ${result.vendor_name}`);
            return false;
        }

        return true;

    } catch {
        showVendorMessage("Failed to connect to procurement manager service");
        return false;
    }
}

document.getElementById("gst_number").addEventListener("blur", async event => {
    await checkVendorGST(event.target.value);
});

// ─── Vendor Masters: Excel File Chosen ────────────────────────────────────────

vendorExcelFile.addEventListener("change", () => {
    vendorExcelFileName.textContent = vendorExcelFile.files.length
        ? vendorExcelFile.files[0].name
        : "No file chosen";
});

// ─── Vendor Masters: required-field validation (matches NOT NULL columns) ─────

const VENDOR_REQUIRED_INPUTS = {
    vendor_name: "Vendor name", legal_entity: "Legal entity", commercial_role: "Commercial role",
    year_of_incorporation: "Year of incorporation",
    office_address: "Office address", office_name: "Office contact name", office_phone: "Office contact number",
    director_name: "Director name", director_designation: "Director designation",
    director_mobile: "Director mobile", director_email: "Director email",
    sales_name: "Sales team name", sales_contact: "Sales team contact", sales_email: "Sales team email",
    accounts_name: "Accounts team name", accounts_contact: "Accounts team contact", accounts_email: "Accounts team email",
    gst_number: "GST number", pan_number: "PAN number",
    bank_name: "Bank name", account_no: "Bank account number", bank_branch: "Bank branch",
    account_type: "Account type", ifsc_rtgs: "IFSC code",
    branch1_address: "Branch 1 address", turnover_value_1: "Latest year turnover",
    recommended_by: "Recommended by", approved_by: "Approved by"
};

const VENDOR_REQUIRED_DOCS = {
    gst_document: "GST document",
    pan_document: "PAN document",
    itr_last_year_document: "Last year ITR"
};

function validateVendorForm() {
    for (const [id, label] of Object.entries(VENDOR_REQUIRED_INPUTS)) {
        if (!String(document.getElementById(id)?.value ?? "").trim()) return `${label} is required.`;
    }
    for (const [id, label] of Object.entries(VENDOR_REQUIRED_DOCS)) {
        if (!document.getElementById(id)?.files?.length) return `${label} is required.`;
    }
    return "";
}

// Field names (DB column names) that must be filled in the Excel preview
const VENDOR_REQUIRED_FIELDS = [
    "vendor_name", "legal_entity", "commercial_role", "year_of_incorporation",
    "office_address", "office_contact_name", "office_contact_number",
    "director_or_ceo_or_management_name", "director_or_ceo_or_management_designation",
    "director_or_ceo_or_management_mobile_no", "director_or_ceo_or_management_email",
    "sales_team_name", "sales_team_contact", "sales_team_email",
    "accounts_team_name", "accounts_team_contact", "accounts_team_email",
    "gst_number", "pan_number",
    "bank_name", "bank_account_no", "bank_branch", "bank_account_type", "bank_ifsc",
    "branch_office_1_address", "turnover_year_1", "turnover_value_1",
    "recommended_by", "approved_by"
];

function validateVendorPreviewData(data) {
    for (const f of VENDOR_REQUIRED_FIELDS) {
        if (!String(data[f] ?? "").trim()) return `${formatVendorFieldName(f)} is required.`;
    }
    for (const [id, label] of Object.entries(VENDOR_REQUIRED_DOCS)) {
        if (!document.getElementById(`excel_${id}`)?.files?.length) return `${label} is required.`;
    }
    return "";
}

// ─── Vendor Masters: Manual Form Submit ───────────────────────────────────────

vendorManualForm.addEventListener("submit", async event => {
    event.preventDefault();

    const validationError = validateVendorForm();
    if (validationError) {
        showVendorMessage(validationError);
        return;
    }

    if (!(await checkVendorGST(document.getElementById("gst_number").value))) return;

    const turnoverLabel = (inputId) => {
        const input = document.getElementById(inputId);
        const label = input?.closest(".form-group")?.querySelector("label");
        return label ? label.textContent.replace(" Turnover", "").trim() : "";
    };

    const val = id => document.getElementById(id)?.value ?? "";

    const data = {
        registration_date:                         val("registration_date") || todayISO(),
        vendor_name:                               val("vendor_name"),

        office_address:                            val("office_address"),
        office_contact_name:                       val("office_name"),
        office_contact_number:                     val("office_phone"),
        factory_address:                           val("factory_address"),
        factory_contact_name:                      val("factory_name"),
        factory_contact_number:                    val("factory_phone"),
        warehouse_address:                         val("warehouse_address"),
        warehouse_contact_name:                    val("warehouse_name"),
        warehouse_contact_number:                  val("warehouse_phone"),
        workshop_address:                          val("workshop_address"),
        workshop_contact_name:                     val("workshop_name"),
        workshop_contact_number:                   val("workshop_phone"),

        legal_entity:                              val("legal_entity"),
        commercial_role:                           val("commercial_role"),
        year_of_incorporation:                     val("year_of_incorporation"),

        director_or_ceo_or_management_name:        val("director_name"),
        director_or_ceo_or_management_designation: val("director_designation"),
        director_or_ceo_or_management_mobile_no:   val("director_mobile"),
        director_or_ceo_or_management_email:       val("director_email"),
        director_or_ceo_or_management_web_address: val("director_web"),

        sales_team_name:                           val("sales_name"),
        sales_team_contact:                        val("sales_contact"),
        sales_team_email:                          val("sales_email"),
        accounts_team_name:                        val("accounts_name"),
        accounts_team_contact:                     val("accounts_contact"),
        accounts_team_email:                       val("accounts_email"),

        gst_number:                                val("gst_number"),
        pan_number:                                val("pan_number"),
        msme_number:                               val("msme_number"),

        bank_name:                                 val("bank_name"),
        bank_account_no:                           val("account_no"),
        bank_branch:                               val("bank_branch"),
        bank_account_type:                         val("account_type"),
        bank_ifsc:                                 val("ifsc_rtgs"),

        branch_office_1_address:                   val("branch1_address"),
        branch_office_1_contact_name:              val("branch1_name"),
        branch_office_1_contact_number:            val("branch1_contact"),
        branch_office_2_address:                   val("branch2_address"),
        branch_office_2_contact_name:              val("branch2_name"),
        branch_office_2_contact_number:            val("branch2_contact"),
        branch_office_3_address:                   val("branch3_address"),
        branch_office_3_contact_name:              val("branch3_name"),
        branch_office_3_contact_number:            val("branch3_contact"),

        turnover_year_1:                           turnoverLabel("turnover_value_1"),
        turnover_value_1:                          val("turnover_value_1"),
        turnover_year_2:                           turnoverLabel("turnover_value_2"),
        turnover_value_2:                          val("turnover_value_2"),
        turnover_year_3:                           turnoverLabel("turnover_value_3"),
        turnover_value_3:                          val("turnover_value_3"),

        recommended_by:                            val("recommended_by"),
        approved_by:                               val("approved_by")
    };

    const formData = new FormData();
    formData.append("vendor_data", JSON.stringify(data));

    ["gst_document", "pan_document", "msme_document",
     "itr_last_year_document", "itr_second_last_year_document", "itr_third_last_year_document"
    ].forEach(docId => {
        const file = document.getElementById(docId)?.files[0];
        if (file) formData.append(docId, file);
    });

    const submitBtn = vendorManualForm.querySelector("button[type='submit']");

    try {
        submitBtn.disabled = true;
        showVendorMessage("Saving vendor...");

        const response = await apiFetch("/vendors", {
            method: "POST",
            body: formData
        });
        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            if (response.status === 409) {
                alert(result.message);
            } else {
                showVendorMessage(result.message || "Failed to add vendor");
            }
            return;
        }

        alert(`Vendor "${data.vendor_name}" added successfully!\nVendor Code: ${result.vendor_code}`);
        showVendorMessage("");
        vendorManualForm.reset();
        setTurnoverYears();
        prefillOfficeUse();

    } catch {
        showVendorMessage("Failed to connect to procurement service");
    } finally {
        submitBtn.disabled = false;
    }
});

// ─── Vendor Masters: Excel Import Preview ─────────────────────────────────────

vendorImportButton.addEventListener("click", async () => {
    if (!vendorExcelFile.files.length) {
        showVendorMessage("Please select an Excel file");
        return;
    }

    const formData = new FormData();
    formData.append("file", vendorExcelFile.files[0]);

    try {
        vendorImportButton.disabled = true;
        showVendorMessage("Processing Excel file...");

        const response = await apiFetch("/vendors/import-preview", {
            method: "POST",
            body: formData
        });
        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            showVendorMessage(result.message || "Failed to process Excel file");
            return;
        }

        await renderVendorPreview(result.vendor);
        showVendorMessage("Vendor preview ready. Upload documents, then click Save Vendor.");

    } catch {
        showVendorMessage("Failed to connect to procurement manager service");
    } finally {
        vendorImportButton.disabled = false;
    }
});

const VENDOR_SECTIONS = [
    {
        title: "Basic Details",
        fields: [
            ["registration_date", "Registration Date"],
            ["vendor_name", "Vendor Name"],
            ["legal_entity", "Legal Entity"],
            ["commercial_role", "Commercial Role"],
            ["year_of_incorporation", "Year of Incorporation"]
        ]
    },
    {
        title: "Address with Phone Numbers",
        fields: [
            ["office_address", "Office Address"],
            ["office_contact_name", "Office Contact Name"],
            ["office_contact_number", "Office Contact Number"],
            ["factory_address", "Factory Address"],
            ["factory_contact_name", "Factory Contact Name"],
            ["factory_contact_number", "Factory Contact Number"],
            ["warehouse_address", "Warehouse Address"],
            ["warehouse_contact_name", "Warehouse Contact Name"],
            ["warehouse_contact_number", "Warehouse Contact Number"],
            ["workshop_address", "Workshop Address"],
            ["workshop_contact_name", "Workshop Contact Name"],
            ["workshop_contact_number", "Workshop Contact Number"]
        ]
    },
    {
        title: "Director / CEO / Management Team",
        fields: [
            ["director_or_ceo_or_management_name", "Name"],
            ["director_or_ceo_or_management_designation", "Designation"],
            ["director_or_ceo_or_management_mobile_no", "Mobile No."],
            ["director_or_ceo_or_management_email", "Email"],
            ["director_or_ceo_or_management_web_address", "Web Address"]
        ]
    },
    {
        title: "Sales Team",
        fields: [
            ["sales_team_name", "Name"],
            ["sales_team_contact", "Contact"],
            ["sales_team_email", "Email"]
        ]
    },
    {
        title: "Accounts Team",
        fields: [
            ["accounts_team_name", "Name"],
            ["accounts_team_contact", "Contact"],
            ["accounts_team_email", "Email"]
        ]
    },
    {
        title: "Tax & Registration",
        fields: [
            ["gst_number", "GST Number"],
            ["pan_number", "PAN Number"],
            ["msme_number", "MSME Number"]
        ]
    },
    {
        title: "Bank Details",
        fields: [
            ["bank_name", "Bank Name"],
            ["bank_account_no", "Account No."],
            ["bank_branch", "Branch Details"],
            ["bank_account_type", "Type of Account"],
            ["bank_ifsc", "IFSC/RTGS Code"]
        ]
    },
    {
        title: "Branch Offices",
        fields: [
            ["branch_office_1_address", "Branch 1 Address"],
            ["branch_office_1_contact_name", "Branch 1 Contact Name"],
            ["branch_office_1_contact_number", "Branch 1 Contact Number"],
            ["branch_office_2_address", "Branch 2 Address"],
            ["branch_office_2_contact_name", "Branch 2 Contact Name"],
            ["branch_office_2_contact_number", "Branch 2 Contact Number"],
            ["branch_office_3_address", "Branch 3 Address"],
            ["branch_office_3_contact_name", "Branch 3 Contact Name"],
            ["branch_office_3_contact_number", "Branch 3 Contact Number"]
        ]
    },
    {
        title: "Turnover of Last Three Years",
        fields: [
            ["turnover_year_1", "Year 1"],
            ["turnover_value_1", "Value 1"],
            ["turnover_year_2", "Year 2"],
            ["turnover_value_2", "Value 2"],
            ["turnover_year_3", "Year 3"],
            ["turnover_value_3", "Value 3"]
        ]
    },
    {
        title: "For Office Use Only",
        fields: [
            ["recommended_by", "Recommended By"],
            ["approved_by", "Approved By"]
        ]
    }
];

async function renderVendorPreview(vendor) {
    vendorPreviewForm.innerHTML = "";

    const used = new Set();

    const buildSection = (title, fields) => {
        const section = document.createElement("div");
        section.className = "form-section";

        const heading = document.createElement("h3");
        heading.className = "form-section-title";
        heading.textContent = title;
        section.appendChild(heading);

        const grid = document.createElement("div");
        grid.className = "form-grid";

        fields.forEach(([field, label]) => {
            used.add(field);

            const isDate = field === "registration_date";
            const value  = vendor[field];
            const shown  = isDate ? (value || todayISO()) : (value || "");

            const group = document.createElement("div");
            group.className = "form-group";
            group.innerHTML = `
                <label>${escapeHtml(label)}</label>
                <input id="preview_${field}" type="${isDate ? "date" : "text"}" value="${escapeHtml(shown)}">
            `;
            grid.appendChild(group);
        });

        section.appendChild(grid);
        vendorPreviewForm.appendChild(section);
    };

    VENDOR_SECTIONS.forEach(section => buildSection(section.title, section.fields));

    // Safety net: any field the backend returns that is not listed above
    const extraFields = Object.keys(vendor)
        .filter(field => !used.has(field))
        .map(field => [field, formatVendorFieldName(field)]);

    if (extraFields.length) buildSection("Other Details", extraFields);

    const recommendedInput = document.getElementById("preview_recommended_by");
    const approvedInput    = document.getElementById("preview_approved_by");

    if (recommendedInput && !recommendedInput.value) recommendedInput.value = currentUser();
    if (approvedInput    && !approvedInput.value)    approvedInput.value    = currentUser();

    vendorExcelPreview.classList.remove("hidden");

    const gstInput = document.getElementById("preview_gst_number");

    if (gstInput) {
        const gstValid = await checkVendorGST(gstInput.value);
        saveVendorExcelButton.disabled = !gstValid;
        return;
    }

    saveVendorExcelButton.disabled = false;
}

function formatVendorFieldName(field) {
    return field
        .replace(/_/g, " ")
        .replace(/\b\w/g, c => c.toUpperCase());
}

// ─── Vendor Masters: Save Excel-Imported Vendor ────────────────────────────────

saveVendorExcelButton.addEventListener("click", async () => {
    const gstInput = document.getElementById("preview_gst_number");

    if (gstInput && !(await checkVendorGST(gstInput.value))) return;

    const data = {};
    vendorPreviewForm.querySelectorAll("input").forEach(input => {
        data[input.id.replace("preview_", "")] = input.value;
    });

    if (!data.registration_date) data.registration_date = todayISO();
    if (!data.recommended_by)    data.recommended_by    = currentUser();
    if (!data.approved_by)       data.approved_by       = currentUser();

    const validationError = validateVendorPreviewData(data);
    if (validationError) {
        showVendorMessage(validationError);
        return;
    }

    const vendorName = data.vendor_name || "Vendor";

    const formData = new FormData();
    formData.append("vendor_data", JSON.stringify(data));

    ["gst_document", "pan_document", "msme_document",
     "itr_last_year_document", "itr_second_last_year_document", "itr_third_last_year_document"
    ].forEach(docId => {
        const file = document.getElementById(`excel_${docId}`)?.files[0];
        if (file) formData.append(docId, file);
    });

    try {
        saveVendorExcelButton.disabled = true;
        showVendorMessage("Saving vendor...");

        const response = await apiFetch("/vendors/import", {
            method: "POST",
            body: formData
        });
        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            if (response.status === 409) {
                alert(result.message);
            } else {
                showVendorMessage(result.message || "Failed to save vendor");
            }
            return;
        }

        alert(`Vendor "${vendorName}" added successfully!\nVendor Code: ${result.vendor_code}`);
        showVendorMessage("");

        vendorExcelPreview.classList.add("hidden");
        vendorPreviewForm.innerHTML = "";
        vendorExcelFile.value = "";
        vendorExcelFileName.textContent = "No file chosen";

        ["excel_gst_document", "excel_pan_document", "excel_msme_document",
         "excel_itr_last_year_document", "excel_itr_second_last_year_document", "excel_itr_third_last_year_document"
        ].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = "";
        });

    } catch {
        showVendorMessage("Failed to connect to procurement service");
    } finally {
        saveVendorExcelButton.disabled = false;
    }
});

// ─── Vendor Inquiries: Load List ───────────────────────────────────────────────

async function loadVendorInquiries() {
    vendorInquiryList.innerHTML = "Loading inquiries...";

    try {
        const response = await apiFetch("/vendor-inquiries");
        if (!response) return;
        const result   = await response.json();

        if (!response.ok || !result.success) {
            vendorInquiryList.textContent = result.message || "Failed to load inquiries";
            return;
        }

        renderVendorInquiries(result.inquiries || []);

    } catch {
        vendorInquiryList.textContent = "Failed to connect to procurement manager service";
    }
}

// ─── Vendor Inquiries: Render Cards ───────────────────────────────────────────

function renderVendorInquiries(inquiries) {
    vendorInquiryList.innerHTML = "";

    if (!inquiries.length) {
        vendorInquiryList.textContent = "No open vendor inquiries available";
        return;
    }

    inquiries.forEach(inquiry => {
        const id   = inquiry.inquiry_id;
        const card = document.createElement("div");
        card.className = "inquiry-card";

        card.dataset.prNumber      = inquiry.pr_number || "";
        card.dataset.itemName      = inquiry.item_name || "";
        card.dataset.quotedVendors = "[]";

        card.innerHTML = `
            <div class="inquiry-card-info">
                ${prInfoItems(inquiry)}
            </div>

            <div class="inquiry-card-footer">
                <button
                    type="button"
                    class="primary-button add-vendor-btn"
                    data-inquiry-id="${id}"
                    data-qty="${inquiry.qty ?? 0}">
                    + Add Vendor
                </button>
                <span class="inquiry-status">OPEN</span>
            </div>

            <div class="add-vendor-form" id="addVendorForm-${id}" style="display:none">
                <h3>Add Vendor Quotation</h3>

                <div class="form-group">
                    <label>Vendor</label>
                    <select class="vendor-select" required>
                        <option value="">Select</option>
                    </select>
                </div>

                <div style="display:flex;gap:16px;align-items:flex-end;">
                    <div class="form-group" style="flex:1;">
                        <label>Price Per Unit</label>
                        <input type="number" class="vendor-price" min="0.01" step="0.01" placeholder="Enter price per unit" required>
                    </div>
                    <div class="form-group" style="flex:1;">
                        <label>Total Price</label>
                        <input type="text" class="vendor-total-display" readonly style="background:#f5f5f5;">
                    </div>
                </div>

                <div class="form-group">
                    <label>Expected Delivery Date</label>
                    <input type="date" class="vendor-delivery-date">
                </div>

                <div class="form-group">
                    <label>Advance (%)</label>
                    <input type="number" class="vendor-advance-pct" min="0" max="100" step="0.01" placeholder="e.g. 30">
                </div>

                <div class="form-group vendor-advance-preview-group" style="display:none">
                    <label>Advance Amount</label>
                    <input type="text" class="vendor-advance-preview" readonly>
                </div>

                <div class="form-group balance-days-group" style="display:none">
                    <label>Remaining Payment Due (days)</label>
                    <input type="number" class="vendor-balance-days" min="1" step="1" placeholder="e.g. 30, 45, 60">
                </div>

                <div class="form-group">
                    <label>Payment Terms Remarks <span style="font-weight:400;color:#888;">(optional)</span></label>
                    <textarea class="vendor-payment-remarks" placeholder="Any additional payment notes"></textarea>
                </div>

                <div class="add-vendor-form-actions">
                    <button type="button" class="primary-button save-vendor-btn" data-inquiry-id="${id}">
                        Save Vendor
                    </button>
                    <button type="button" class="cancel-vendor-btn" data-inquiry-id="${id}">
                        Cancel
                    </button>
                </div>

                <p class="vendor-form-message"></p>
            </div>

            <div class="inquiry-vendor-section" id="vendorsSection-${id}" hidden></div>
        `;

        vendorInquiryList.appendChild(card);

        // FIX: show quotations that already exist when the tab is opened
        loadInquiryVendors(id);
    });
}

// FIX: single shared calculator for total + advance preview + balance-days visibility
function wireQuotationForm(form, qty) {
    if (form.dataset.wired === "1") return;
    form.dataset.wired = "1";

    const priceEl          = form.querySelector(".vendor-price");
    const totalDisplayEl   = form.querySelector(".vendor-total-display");
    const advancePctEl     = form.querySelector(".vendor-advance-pct");
    const balanceDaysGroup = form.querySelector(".balance-days-group");
    const previewGroup     = form.querySelector(".vendor-advance-preview-group");
    const previewInput     = form.querySelector(".vendor-advance-preview");

    const fmt = n => `₹${n.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;

    const update = () => {
        const price = Number(priceEl.value) || 0;
        const pct   = Number(advancePctEl.value) || 0;

        totalDisplayEl.value = (price > 0 && qty > 0) ? fmt(price * qty) : "";

        balanceDaysGroup.style.display = (pct > 0 && pct < 100) ? "" : "none";
        if (pct >= 100 || pct === 0) {
            balanceDaysGroup.querySelector(".vendor-balance-days").value = "";
        }

        if (pct > 0 && price > 0) {
            previewInput.value         = fmt((pct / 100) * price * qty);
            previewGroup.style.display = "";
        } else {
            previewGroup.style.display = "none";
            previewInput.value         = "";
        }
    };

    priceEl.addEventListener("input", update);
    advancePctEl.addEventListener("input", update);
}

// ─── Vendor Inquiries: Delegated Click Handler ─────────────────────────────────

vendorInquiryList.addEventListener("click", async event => {

    // Add Vendor → open (or close if already open)
    const addBtn = event.target.closest(".add-vendor-btn");
    if (addBtn) {
        const inquiryId = addBtn.dataset.inquiryId;
        const form      = document.getElementById(`addVendorForm-${inquiryId}`);

        // FIX: clicking again closes instead of wiping the form
        if (form.style.display !== "none") {
            form.style.display = "none";
            return;
        }

        form.style.display = "";

        wireQuotationForm(form, Number(addBtn.dataset.qty) || 0);

        // FIX: hide vendors that already have a quotation on this inquiry
        const card   = form.closest(".inquiry-card");
        const quoted = JSON.parse(card.dataset.quotedVendors || "[]").map(String);
        await loadVendorsIntoDropdown(form.querySelector(".vendor-select"), quoted);
        return;
    }

    // Cancel → hide form
    const cancelBtn = event.target.closest(".cancel-vendor-btn");
    if (cancelBtn) {
        const inquiryId = cancelBtn.dataset.inquiryId;
        document.getElementById(`addVendorForm-${inquiryId}`).style.display = "none";
        return;
    }

    // Save Vendor → POST quotation
    const saveBtn = event.target.closest(".save-vendor-btn");
    if (saveBtn) {
        await saveVendorForInquiry(saveBtn.dataset.inquiryId);
    }
});

// ─── Vendor Inquiries: Populate Vendor Dropdown ────────────────────────────────

async function loadVendorsIntoDropdown(selectElement, excludeIds = []) {
    const previous = selectElement.value;
    selectElement.innerHTML = '<option value="">Loading vendors...</option>';

    try {
        const response = await apiFetch("/vendors");
        if (!response) return;
        const result   = await response.json();

        if (!response.ok || !result.success) {
            throw new Error(result.message || "Failed to load vendors");
        }

        const vendors = (result.vendors || [])
            .filter(v => !excludeIds.includes(String(v.vendor_id)))
            .sort((a, b) => a.vendor_id - b.vendor_id);

        if (!vendors.length) {
            selectElement.innerHTML = '<option value="">No eligible vendors available</option>';
            return;
        }

        selectElement.innerHTML = '<option value="">Select Vendor</option>';

        vendors.forEach(vendor => {
            const option       = document.createElement("option");
            option.value       = vendor.vendor_id;
            option.textContent = `${vendor.vendor_name} (${vendor.vendor_code || vendor.vendor_id})`;
            selectElement.appendChild(option);
        });

        if (previous && vendors.some(v => String(v.vendor_id) === previous)) {
            selectElement.value = previous;
        }

    } catch {
        selectElement.innerHTML = '<option value="">Failed to load vendors</option>';
    }
}

// ─── Vendor Inquiries: Save Quotation ─────────────────────────────────────────

async function saveVendorForInquiry(inquiryId) {
    const form             = document.getElementById(`addVendorForm-${inquiryId}`);
    const select           = form.querySelector(".vendor-select");
    const priceEl          = form.querySelector(".vendor-price");
    const deliveryDateEl   = form.querySelector(".vendor-delivery-date");
    const advancePctEl     = form.querySelector(".vendor-advance-pct");
    const balanceDaysEl    = form.querySelector(".vendor-balance-days");
    const paymentRemarksEl = form.querySelector(".vendor-payment-remarks");
    const remarksEl        = form.querySelector(".vendor-remarks");
    const msgEl            = form.querySelector(".vendor-form-message");
    const saveBtn          = form.querySelector(".save-vendor-btn");

    const vendorId   = select.value;
    const vendorName = select.options[select.selectedIndex]?.text || "";
    const price      = Number(priceEl.value);
    const advancePct = advancePctEl.value !== "" ? Number(advancePctEl.value) : null;

    if (!vendorId) { msgEl.textContent = "Please select a vendor."; return; }
    if (!price || price <= 0) { msgEl.textContent = "Please enter a valid price per unit."; return; }
    if (advancePct !== null && (advancePct < 0 || advancePct > 100)) {
        msgEl.textContent = "Advance % must be between 0 and 100."; return;
    }

    const needsBalanceDays = advancePct !== null && advancePct > 0 && advancePct < 100;
    const balanceDays      = balanceDaysEl.value ? Number(balanceDaysEl.value) : null;

    if (needsBalanceDays && (!balanceDays || balanceDays < 1)) {
        msgEl.textContent = "Please enter how many days for the remaining payment."; return;
    }

    let payment_type, advance_type, advance_value;

    if (advancePct === null || advancePct === 0) {
        payment_type  = "CREDIT";
        advance_type  = null;
        advance_value = null;
    } else if (advancePct >= 100) {
        payment_type  = "ADVANCE";
        advance_type  = "PERCENTAGE";
        advance_value = 100;
    } else {
        payment_type  = "ADVANCE_PLUS_BALANCE";
        advance_type  = "PERCENTAGE";
        advance_value = advancePct;
    }

    const card     = form.closest(".inquiry-card");
    const prNumber = card?.dataset.prNumber || "";
    const itemName = card?.dataset.itemName || "";

    try {
        saveBtn.disabled  = true;
        msgEl.textContent = "Saving...";

        const response = await apiFetch(`/vendor-inquiries/${inquiryId}/vendors`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                vendor_id:              Number(vendorId),
                price_per_unit:         price,
                expected_delivery_date: deliveryDateEl.value || null,
                payment_type,
                advance_type,
                advance_value,
                balance_due_days:       balanceDays,
                payment_terms_remarks:  paymentRemarksEl.value.trim() || null,
                remarks:                remarksEl?.value.trim() || null
            })
        });
        if (!response) return;

        const result = await response.json();
        if (!response.ok || !result.success) throw new Error(result.message || "Failed to add vendor");

        alert(
            `Vendor added successfully!\n\n` +
            `PR Number: ${prNumber}\n` +
            `Item: ${itemName}\n` +
            `Vendor: ${vendorName}\n` +
            `Total Price: ₹${formatCurrency(result.total_price)}`
        );

        msgEl.textContent = "";

        select.value           = "";
        priceEl.value          = "";
        deliveryDateEl.value   = "";
        advancePctEl.value     = "";
        balanceDaysEl.value    = "";
        paymentRemarksEl.value = "";
        if (remarksEl) remarksEl.value = "";
        form.querySelector(".balance-days-group").style.display           = "none";
        form.querySelector(".vendor-advance-preview-group").style.display = "none";
        form.querySelector(".vendor-advance-preview").value               = "";
        form.querySelector(".vendor-total-display").value                 = "";

        await loadInquiryVendors(inquiryId);

        // refresh dropdown so the vendor just quoted disappears
        const quoted = JSON.parse(card.dataset.quotedVendors || "[]").map(String);
        await loadVendorsIntoDropdown(select, quoted);

    } catch (error) {
        msgEl.textContent = error.message || "Failed to connect to procurement service";
    } finally {
        saveBtn.disabled = false;
    }
}

// ─── Vendor Inquiries: Load Vendor Table for One Card ─────────────────────────

async function loadInquiryVendors(inquiryId) {
    const section = document.getElementById(`vendorsSection-${inquiryId}`);
    if (!section) return;

    section.innerHTML = "<em style='font-size:13px;color:#777;'>Loading vendors...</em>";

    try {
        const response = await apiFetch(`/vendor-inquiries/${inquiryId}`);
        if (!response) return;
        const result   = await response.json();

        if (!response.ok || !result.success) {
            section.innerHTML = `<em style='font-size:13px;color:#c00;'>${escapeHtml(result.message || "Failed to load vendors")}</em>`;
            return;
        }

        const vendors = result.vendors || [];

        // remember who is already quoted (assumes each row carries vendor_id)
        const card = section.closest(".inquiry-card");
        if (card) card.dataset.quotedVendors = JSON.stringify(vendors.map(v => v.vendor_id).filter(v => v != null));

        renderInquiryVendorTable(section, vendors);

    } catch {
        section.innerHTML = "<em style='font-size:13px;color:#c00;'>Failed to connect to procurement manager service</em>";
    }
}

// FIX: 9 headers, 9 cells, same order; removed the empty Action column
function renderInquiryVendorTable(container, vendors) {
    if (!vendors.length) {
        container.innerHTML = "";
        container.hidden = true;
        return;
    }

    container.hidden = false;

    container.innerHTML = `
        <div class="table-container">
            <table class="inquiry-vendor-table">
                <thead>
                    <tr>
                        <th>Vendor Code</th>
                        <th>Vendor Name</th>
                        <th>Price / Unit</th>
                        <th>Total Price</th>
                        <th>Advance</th>
                        <th>Advance Amount</th>
                        <th>Delivery Date</th>
                        <th>Balance Due (days)</th>
                        <th>Remarks</th>
                    </tr>
                </thead>
                <tbody>
                    ${vendors.map(v => `
                        <tr>
                            <td>${escapeHtml(v.vendor_code || "-")}</td>
                            <td>${escapeHtml(v.vendor_name || "-")}</td>
                            <td>${formatCurrency(v.price_per_unit)}</td>
                            <td>${formatCurrency(v.total_price)}</td>
                            <td>${v.advance_value != null ? `${escapeHtml(String(v.advance_value))}%` : "-"}</td>
                            <td>${v.advance_amount != null ? `₹${formatCurrency(v.advance_amount)}` : "-"}</td>
                            <td>${formatDate(v.expected_delivery_date)}</td>
                            <td>${v.balance_due_days != null ? escapeHtml(String(v.balance_due_days)) : "-"}</td>
                            <td>${escapeHtml(v.remarks || "-")}</td>
                        </tr>
                    `).join("")}
                </tbody>
            </table>
        </div>
    `;
}

// ─── Logout ────────────────────────────────────────────────────────────────────

logoutButton.addEventListener("click", async () => {
    try {
        logoutButton.disabled = true;

        const response = await apiFetch("/logout", { method: "POST" });
        if (!response) return;
        const result   = await response.json();

        if (result.success) {
            window.location.href = result.redirect_url;
            return;
        }

        showMessage(result.message || "Failed to logout");

    } catch {
        showMessage("Failed to logout");
    } finally {
        logoutButton.disabled = false;
    }
});

// ─── Quotation & Comparisons ───────────────────────────────────────────────────

async function loadQuotationComparisons() {
    const list = document.getElementById("quotationList");
    list.innerHTML = "Loading...";

    try {
        const response = await apiFetch("/quotation-comparisons");
        if (!response) return;
        const result   = await response.json();

        if (!response.ok || !result.success) {
            list.textContent = result.message || "Failed to load quotations";
            return;
        }

        renderQuotationComparisons(result.inquiries || []);

    } catch {
        list.textContent = "Failed to connect to procurement manager service";
    }
}

function renderQuotationComparisons(inquiries) {
    const list = document.getElementById("quotationList");
    list.innerHTML = "";

    if (!inquiries.length) {
        list.textContent = "No open inquiries found";
        return;
    }

    inquiries.forEach(inquiry => {
        const card = document.createElement("div");
        card.className = "inquiry-card";

        const prices      = inquiry.vendors.map(v => Number(v.price_per_unit)).filter(p => p > 0);
        const lowestPrice = prices.length ? Math.min(...prices) : null;

        card.innerHTML = `
            <div class="inquiry-card-info">
                ${prInfoItems(inquiry)}
            </div>

            <div class="inquiry-card-footer">
                <span class="inquiry-status">OPEN</span>
                <span style="font-size:13px;color:#777;">
                    ${inquiry.vendors.length} vendor${inquiry.vendors.length !== 1 ? "s" : ""} quoted
                </span>
            </div>

            ${inquiry.vendors.length ? `
                <div class="inquiry-vendor-section">
                    <div class="table-container">
                        <table class="inquiry-vendor-table">
                            <thead>
                                <tr>
                                    <th>Vendor Code</th>
                                    <th>Vendor Name</th>
                                    <th>Price / Unit</th>
                                    <th>Total Price</th>
                                    <th>Advance</th>
                                    <th>Advance Amount</th>
                                    <th>Delivery Date</th>
                                    <th>Balance Due (days)</th>
                                    <th>Remarks</th>
                                    <th>Action</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${inquiry.vendors.map(v => {
                                    const isLowest =
                                        lowestPrice !== null &&
                                        Number(v.price_per_unit) === lowestPrice;

                                    return `
                                        <tr>
                                            <td>${escapeHtml(v.vendor_code || "-")}</td>
                                            <td>${escapeHtml(v.vendor_name || "-")}</td>
                                            <td class="${isLowest ? "lowest-price" : ""}">${formatCurrency(v.price_per_unit)}</td>
                                            <td>${formatCurrency(v.total_price)}</td>
                                            <td>${v.advance_value != null ? `${escapeHtml(String(v.advance_value))}%` : "-"}</td>
                                            <td>${v.advance_amount != null ? `₹${formatCurrency(v.advance_amount)}` : "-"}</td>
                                            <td>${formatDate(v.expected_delivery_date)}</td>
                                            <td>${v.balance_due_days != null ? escapeHtml(String(v.balance_due_days)) : "-"}</td>
                                            <td>${escapeHtml(v.remarks || "-")}</td>
                                            <td>
                                                ${v.is_selected
                                                    ? `<span style="color:#16803c;font-weight:600;">✓ Selected</span>`
                                                    : `<button
                                                            type="button"
                                                            class="primary-button select-vendor-btn"
                                                            data-inquiry-id="${inquiry.inquiry_id}"
                                                            data-inquiry-vendor-id="${v.inquiry_vendor_id}">
                                                            Select
                                                       </button>`
                                                }
                                            </td>
                                        </tr>
                                    `;
                                }).join("")}
                            </tbody>
                        </table>
                    </div>
                </div>
            ` : `
                <p class="no-vendors-note" style="margin-top:14px;">
                    No vendor quotations added yet for this inquiry
                </p>
            `}
        `;

        list.appendChild(card);
    });
}

// ─── Quotation & Comparisons: Select Vendor ────────────────────────────────────

document.getElementById("quotationList").addEventListener("click", async event => {
    const btn = event.target.closest(".select-vendor-btn");
    if (!btn) return;

    const inquiryId       = btn.dataset.inquiryId;
    const inquiryVendorId = btn.dataset.inquiryVendorId;

    const confirmed = confirm("Select this vendor for the purchase request? This cannot be undone from here.");
    if (!confirmed) return;

    btn.disabled = true;
    btn.textContent = "Selecting...";

    try {
        const response = await apiFetch(`/vendor-inquiries/${inquiryId}/select-vendor`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ inquiry_vendor_id: Number(inquiryVendorId) })
        });
        if (!response) return;

        const result = await response.json();

        if (!response.ok || !result.success) {
            alert(result.message || "Failed to select vendor");
            btn.disabled = false;
            btn.textContent = "Select";
            return;
        }

        await loadQuotationComparisons();

    } catch {
        alert("Failed to connect to procurement manager service");
        btn.disabled = false;
        btn.textContent = "Select";
    }
});

// ─── Order Tracking: Load a Page ───────────────────────────────────────────────

async function loadOrderTracking() {
    orderTrackingList.innerHTML = "Loading...";
    orderTrackingPrev.disabled = true;
    orderTrackingNext.disabled = true;

    try {
        const response = await apiFetch(`/order-tracking?page=${orderTrackingPage}`);
        if (!response) return;

        const result = await response.json();

        if (!response.ok || !result.success) {
            orderTrackingList.textContent = result.message || "Failed to load order tracking data";
            return;
        }

        renderOrderTracking(result.rows || []);

        const { page, total_pages, has_prev, has_next } = result.pagination;
        orderTrackingPage       = page;
        orderTrackingTotalPages = total_pages;

        orderTrackingPageLabel.textContent = `Page ${page} of ${total_pages}`;
        orderTrackingPrev.disabled = !has_prev;
        orderTrackingNext.disabled = !has_next;

    } catch {
        orderTrackingList.textContent = "Failed to connect to procurement manager service";
    }
}

function statusLabel(status) {
    if (!status) return "NO INQUIRY";
    return status.replace(/_/g, " ");
}

function statusClass(status) {
    return `status-${(status || "none").toLowerCase()}`;
}

function renderOrderTracking(rows) {
    orderTrackingList.innerHTML = "";

    if (!rows.length) {
        orderTrackingList.textContent = "No purchase requests found";
        return;
    }

    rows.forEach(row => {
        const card = document.createElement("div");
        card.className = "po-row";

        const poNumberDisplay = row.po_number
            ? `<div class="po-row-field">
                   <span class="inquiry-card-label">PO Number</span>
                   <span class="inquiry-card-value">${escapeHtml(row.po_number)}</span>
               </div>`
            : "";

        const poNumberItem = row.po_number
            ? `<div class="inquiry-card-item">
                   <span class="inquiry-card-label">PO Number</span>
                   <span class="inquiry-card-value">${escapeHtml(row.po_number)}</span>
               </div>`
            : "";

        card.innerHTML = `
            <div class="po-row-summary">
                <div class="po-row-field">
                    <span class="inquiry-card-label">PR Number</span>
                    <span class="inquiry-card-value">${escapeHtml(row.pr_number || "-")}</span>
                </div>
                ${poNumberDisplay}
                <div class="po-row-field">
                    <span class="inquiry-card-label">Party Name</span>
                    <span class="inquiry-card-value">${escapeHtml(row.party_name || "-")}</span>
                </div>
                <div class="po-row-field">
                    <span class="inquiry-card-label">Item Name</span>
                    <span class="inquiry-card-value">${escapeHtml(row.item_name || "-")}</span>
                </div>
                <div class="po-row-field">
                    <span class="inquiry-card-label">Sales Rate</span>
                    <span class="inquiry-card-value">${formatCurrency(row.sales_rate)}</span>
                </div>
                <span class="inquiry-status ${statusClass(row.status)}">${statusLabel(row.status)}</span>
                <button type="button" class="expand-po-btn" data-ot-id="${row.pr_id}">
                    Expand ▾
                </button>
            </div>

            <div class="po-row-detail" id="otDetail-${row.pr_id}" style="display:none">
                <div class="inquiry-card-info">
                    ${prInfoItems(row, poNumberItem)}
                </div>
            </div>
        `;

        orderTrackingList.appendChild(card);
    });
}

// ─── Order Tracking: Pagination Buttons ────────────────────────────────────────

orderTrackingPrev.addEventListener("click", () => {
    if (orderTrackingPage > 1) {
        orderTrackingPage -= 1;
        loadOrderTracking();
    }
});

orderTrackingNext.addEventListener("click", () => {
    if (orderTrackingPage < orderTrackingTotalPages) {
        orderTrackingPage += 1;
        loadOrderTracking();
    }
});

// ─── Order Tracking: Expand Row ────────────────────────────────────────────────

orderTrackingList.addEventListener("click", event => {
    const expandBtn = event.target.closest(".expand-po-btn");
    if (!expandBtn) return;

    const otId   = expandBtn.dataset.otId;
    const detail = document.getElementById(`otDetail-${otId}`);
    const isOpen = detail.style.display !== "none";
    detail.style.display = isOpen ? "none" : "block";
    expandBtn.textContent = isOpen ? "Expand ▾" : "Collapse ▴";
});

// ─── Purchase Orders: Load List ────────────────────────────────────────────────

async function loadPurchaseOrders() {
    purchaseOrdersList.innerHTML = "Loading...";

    try {
        const response = await apiFetch("/purchase-orders");
        if (!response) return;
        const result   = await response.json();

        if (!response.ok || !result.success) {
            purchaseOrdersList.textContent = result.message || "Failed to load purchase orders";
            return;
        }

        renderPurchaseOrders(result.purchase_orders || []);

    } catch {
        purchaseOrdersList.textContent = "Failed to connect to procurement service";
    }
}

function vendorDetailField(label, value) {
    return `
        <div class="vendor-details-item">
            <span class="vendor-details-label">${escapeHtml(label)}</span>
            <span class="vendor-details-value">${escapeHtml(value || "-")}</span>
        </div>
    `;
}

function renderPurchaseOrders(orders) {
    purchaseOrdersList.innerHTML = "";

    if (!orders.length) {
        purchaseOrdersList.textContent = "No purchase orders available";
        return;
    }

    orders.forEach(po => {
        const card = document.createElement("div");
        card.className = "po-row";

        card.innerHTML = `
            <div class="po-row-summary">
                <div class="po-row-field">
                    <span class="inquiry-card-label">PO Number</span>
                    <span class="inquiry-card-value">${escapeHtml(po.po_number || "-")}</span>
                </div>
                <div class="po-row-field">
                    <span class="inquiry-card-label">Party Name</span>
                    <span class="inquiry-card-value">${escapeHtml(po.party_name || "-")}</span>
                </div>
                <div class="po-row-field">
                    <span class="inquiry-card-label">Item Name</span>
                    <span class="inquiry-card-value">${escapeHtml(po.item_name || "-")}</span>
                </div>
                <div class="po-row-field">
                    <span class="inquiry-card-label">Total Price</span>
                    <span class="inquiry-card-value">${formatCurrency(po.total_price)}</span>
                </div>
                <button type="button" class="expand-po-btn" data-po-id="${po.po_id}">
                    Expand ▾
                </button>
            </div>

            <div class="po-row-detail" id="poDetail-${po.po_id}" style="display:none">
                <div class="inquiry-card-info">
                    ${prInfoItems(po)}
                </div>

                <div class="inquiry-card-footer">
                    <button type="button" class="primary-button vendor-details-toggle" data-po-id="${po.po_id}">
                        Show Vendor Details
                    </button>
                    ${po.status === "DRAFT" ? `
                        <button type="button" class="primary-button issue-po-btn" data-po-id="${po.po_id}">
                            Issue PO
                        </button>
                    ` : ""}
                    ${(po.status === "ISSUED" || po.status === "COMPLETED") ? `
                        <button type="button" class="secondary-button download-po-btn" data-po-id="${po.po_id}">
                            ⬇ Download PO
                        </button>
                    ` : ""}
                    <div class="pi-actions" id="piActions-${po.po_id}"></div>
                    <span class="inquiry-status status-${(po.status || "").toLowerCase()}">${escapeHtml(po.status || "-")}</span>
                </div>

                <div class="vendor-details-panel" id="vendorDetailsPanel-${po.po_id}">
                    <div class="vendor-details-grid">
                        ${vendorDetailField("Vendor Code", po.vendor_code)}
                        ${vendorDetailField("Vendor Name", po.vendor_name)}
                        ${vendorDetailField("Legal Entity", po.legal_entity)}
                        ${vendorDetailField("Commercial Role", po.commercial_role)}
                        ${vendorDetailField("Year of Incorporation", po.year_of_incorporation)}
                        ${vendorDetailField("Registration Date", formatDate(po.registration_date))}
                        ${vendorDetailField("Office Address & Phone", po.office_address_and_phone)}
                        ${vendorDetailField("Factory Address & Phone", po.factory_address_and_phone)}
                        ${vendorDetailField("Warehouse Address & Phone", po.warehouse_address_and_phone)}
                        ${vendorDetailField("Workshop Address & Phone", po.workshop_address_and_phone)}
                        ${vendorDetailField("Director/CEO Name", po.director_or_ceo_or_management_name)}
                        ${vendorDetailField("Designation", po.director_or_ceo_or_management_designation)}
                        ${vendorDetailField("Mobile No.", po.director_or_ceo_or_management_mobile_no)}
                        ${vendorDetailField("Email", po.director_or_ceo_or_management_email)}
                        ${vendorDetailField("Web Address", po.director_or_ceo_or_management_web_address)}
                        ${vendorDetailField("Sales Team Name", po.sales_team_name)}
                        ${vendorDetailField("Sales Team Contact", po.sales_team_contact)}
                        ${vendorDetailField("Sales Team Email", po.sales_team_email)}
                        ${vendorDetailField("Accounts Team Name", po.accounts_team_name)}
                        ${vendorDetailField("Accounts Team Contact", po.accounts_team_contact)}
                        ${vendorDetailField("Accounts Team Email", po.accounts_team_email)}
                        ${vendorDetailField("GST Number", po.gst_number)}
                        ${vendorDetailField("PAN Number", po.pan_number)}
                        ${vendorDetailField("MSME Number", po.msme_number)}
                        ${vendorDetailField("Bank Details", po.bank_details)}
                        ${vendorDetailField("Branch Office 1", po.branch_office_1)}
                        ${vendorDetailField("Branch Office 2", po.branch_office_2)}
                        ${vendorDetailField("Branch Office 3", po.branch_office_3)}
                        ${vendorDetailField(po.turnover_year_1 || "Turnover 1", po.turnover_value_1)}
                        ${vendorDetailField(po.turnover_year_2 || "Turnover 2", po.turnover_value_2)}
                        ${vendorDetailField(po.turnover_year_3 || "Turnover 3", po.turnover_value_3)}
                        ${vendorDetailField("Quoted Price / Unit", formatCurrency(po.price_per_unit))}
                        ${vendorDetailField("Quoted Total Price", formatCurrency(po.total_price))}
                        ${vendorDetailField("Quotation Remarks", po.quotation_remarks)}
                    </div>
                </div>
            </div>
        `;

        purchaseOrdersList.appendChild(card);

        if (po.status !== "DRAFT") loadProformaInvoiceStatus(po.po_id);
    });
}

// ─── Purchase Orders: Proforma Invoice (Upload / Download) ────────────────────

async function loadProformaInvoiceStatus(poId) {
    const container = document.getElementById(`piActions-${poId}`);
    if (!container) return;

    try {
        const response = await apiFetch(`/purchase-orders/${poId}/proforma-invoice/status`);
        if (!response) return;

        const result = await response.json();
        if (!response.ok || !result.success) return;

        renderProformaInvoiceButton(container, poId, result.exists);

    } catch {
        // Silently ignore — the button just won't render if the check fails.
    }
}

function renderProformaInvoiceButton(container, poId, exists) {
    container.innerHTML = exists
        ? `<button type="button" class="secondary-button download-pi-btn" data-po-id="${poId}">⬇ Download PI</button>`
        : `<label class="secondary-button upload-pi-label">
               Upload PI
               <input type="file" class="upload-pi-input" data-po-id="${poId}" hidden>
           </label>`;
}

// ─── Purchase Orders: Expand Row + Toggle Vendor Details ──────────────────────

purchaseOrdersList.addEventListener("click", async event => {

    const expandBtn = event.target.closest(".expand-po-btn");
    if (expandBtn) {
        const poId   = expandBtn.dataset.poId;
        const detail = document.getElementById(`poDetail-${poId}`);
        const isOpen = detail.style.display !== "none";
        detail.style.display = isOpen ? "none" : "block";
        expandBtn.textContent = isOpen ? "Expand ▾" : "Collapse ▴";
        return;
    }

    const vendorBtn = event.target.closest(".vendor-details-toggle");
    if (vendorBtn) {
        const poId   = vendorBtn.dataset.poId;
        const panel  = document.getElementById(`vendorDetailsPanel-${poId}`);
        const isOpen = panel.classList.toggle("open");
        vendorBtn.textContent = isOpen ? "Hide Vendor Details" : "Show Vendor Details";
        return;
    }

    const issuePOBtn = event.target.closest(".issue-po-btn");
    if (issuePOBtn) {
        const poId = issuePOBtn.dataset.poId;

        if (!confirm("Issue this Purchase Order? It will become active for goods receipt.")) return;

        issuePOBtn.disabled = true;
        issuePOBtn.textContent = "Issuing...";

        try {
            const response = await apiFetch(`/purchase-orders/${poId}/issue`, { method: "POST" });
            if (!response) return;
            if (!response.ok) {
                const result = await response.json();
                alert(result.message || "Failed to issue PO.");
                issuePOBtn.disabled = false;
                issuePOBtn.textContent = "Issue PO";
                return;
            }
            const blob = await response.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;

            const disposition = response.headers.get("Content-Disposition");
            const match = disposition?.match(/filename="?([^"]+)"?/i);

            a.download = match ? match[1] : `PO_${poId}.pdf`;

            a.click();
            URL.revokeObjectURL(url);
            alert("PO issued successfully.");
            await loadPurchaseOrders();

        } catch (error) {
            console.error(error);
            alert("Failed to connect to procurement service.");
            issuePOBtn.disabled = false;
            issuePOBtn.textContent = "Issue PO";
        }
        return;
    }

    const downloadPoBtn = event.target.closest(".download-po-btn");
    if (downloadPoBtn) {
        window.open(`/purchase-orders/${downloadPoBtn.dataset.poId}/download`, "_blank");
        return;
    }

    const downloadPiBtn = event.target.closest(".download-pi-btn");
    if (downloadPiBtn) {
        window.open(`/purchase-orders/${downloadPiBtn.dataset.poId}/proforma-invoice`, "_blank");
        return;
    }
});

purchaseOrdersList.addEventListener("change", async event => {
    const uploadInput = event.target.closest(".upload-pi-input");
    if (!uploadInput) return;

    const file = uploadInput.files[0];
    if (!file) return;

    const poId      = uploadInput.dataset.poId;
    const formData  = new FormData();
    formData.append("proforma_invoice", file);

    try {
        const response = await apiFetch(`/purchase-orders/${poId}/proforma-invoice`, {
            method: "POST",
            body: formData
        });

        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            alert(result.message || "Failed to upload Proforma Invoice");
            return;
        }

        alert("Proforma Invoice uploaded successfully.");
        await loadProformaInvoiceStatus(poId);

    } catch {
        alert("Failed to connect to procurement service.");
    }
});

// ─── Goods Received ─────────────────────────────────────────────────────────

// FIX: "Show completed" toggle, created here so index.html needs no change
function ensureCompletedToggle() {
    if (document.getElementById("goodsShowCompletedWrap")) return;

    const wrap = document.createElement("label");
    wrap.id = "goodsShowCompletedWrap";
    wrap.style.cssText = "display:inline-flex;align-items:center;gap:8px;margin-bottom:12px;font-size:14px;cursor:pointer;";
    wrap.innerHTML = `<input type="checkbox" id="goodsShowCompleted"> Show completed orders`;

    goodsReceivedList.parentNode.insertBefore(wrap, goodsReceivedList);

    wrap.querySelector("input").addEventListener("change", event => {
        showCompletedGoods = event.target.checked;
        renderGoodsReceived(goodsReceivedOrders);
    });
}

async function loadGoodsReceived() {
    ensureCompletedToggle();

    try {
        goodsReceivedMessage.textContent = "Loading goods received...";

        const response = await apiFetch("/goods-received");
        if (!response) return;
        const result = await response.json();

        if (!response.ok) {
            goodsReceivedMessage.textContent =
                result.message || "Failed to load goods received";
            return;
        }

        goodsReceivedOrders = result.orders || [];
        renderGoodsReceived(goodsReceivedOrders);
        goodsReceivedMessage.textContent = "";

    } catch (error) {
        console.error(error);
        goodsReceivedMessage.textContent =
            "Failed to connect to procurement service";
    }
}

function renderGoodsReceived(allOrders) {
    const rows = showCompletedGoods
        ? allOrders
        : allOrders.filter(o => o.po_status !== "COMPLETED");

    if (!rows.length) {
        goodsReceivedList.innerHTML = `
            <div class="message">
                No purchase orders available for goods receipt.
            </div>
        `;
        return;
    }

    goodsReceivedList.innerHTML = rows.map(po => {

        const ordered   = Number(po.ordered_quantity)   || 0;
        const received  = Number(po.received_quantity)  || 0;
        const returned  = Number(po.returned_quantity)  || 0;   // used if backend sends it
        const remaining = Number(po.remaining_quantity) || 0;
        const progress  = Number(po.progress)           || 0;
        const isDone    = po.po_status === "COMPLETED";

        // FIX: completion is based on net received (received - returned)
        const netReceived = received - returned;

        return `
            <div class="goods-received-card">

                <div class="goods-received-stats">
                    <div>
                        <span class="goods-received-label">PO Number</span>
                        <strong>${escapeHtml(po.po_number)}</strong>
                    </div>
                    <div>
                        <span class="goods-received-label">Vendor</span>
                        <strong>${escapeHtml(po.vendor_name)}</strong>
                    </div>
                    <div>
                        <span class="goods-received-label">Item</span>
                        <strong>${escapeHtml(po.item_name)}</strong>
                    </div>
                    <div>
                        <span class="goods-received-label">Ordered Qty</span>
                        <strong>${ordered} ${escapeHtml(po.unit || "")}</strong>
                    </div>
                    <div>
                        <span class="goods-received-label">Received Qty</span>
                        <strong>${received}</strong>
                    </div>
                    <div>
                        <span class="goods-received-label">Remaining Qty</span>
                        <strong>${remaining}</strong>
                    </div>
                </div>
                <div class="goods-progress-section">
                    <div class="goods-progress-header">
                        <span>Receipt Progress</span>
                        <strong>${progress.toFixed(0)}%</strong>
                    </div>
                    <div class="goods-progress-bar">
                        <div class="goods-progress-fill" style="width: ${Math.min(progress, 100)}%"></div>
                    </div>
                </div>

                <div class="goods-received-actions">
                    ${isDone ? `<span class="inquiry-status status-completed">COMPLETED</span>` : `
                        <button class="primary-button" onclick="openReceiveForm(${po.po_id})">
                            Add Received
                        </button>
                        <button class="secondary-button" onclick="openReturnForm(${po.po_id})">
                            Return
                        </button>
                        ${netReceived >= ordered ? `
                            <button class="secondary-button" onclick="completePO(${po.po_id}, this)">
                                Mark Complete
                            </button>
                        ` : ""}
                    `}
                    <button class="secondary-button" onclick="viewGoodsHistory(${po.po_id})">
                        History
                    </button>
                </div>

                <div id="receive-form-${po.po_id}" class="goods-inline-form hidden"></div>
                <div id="return-form-${po.po_id}" class="goods-inline-form hidden"></div>
                <div id="history-${po.po_id}" class="goods-history hidden"></div>

            </div>
        `;

    }).join("");
}

function openReceiveForm(poId) {
    const form       = document.getElementById(`receive-form-${poId}`);
    const returnForm = document.getElementById(`return-form-${poId}`);

    returnForm.classList.add("hidden");

    form.innerHTML = `
        <h3>Add Received Quantity</h3>
        <div class="form-grid">
            <div class="form-group">
                <label>Received Quantity</label>
                <input type="number" id="received-quantity-${poId}" min="0.01" step="0.01" required>
            </div>
            <div class="form-group">
                <label>Receipt Date</label>
                <input type="date" id="receipt-date-${poId}" value="${todayISO()}" required>
            </div>
            <div class="form-group full-width">
                <label>Remarks <span style="font-weight:400;color:#888;">(optional)</span></label>
                <textarea id="receipt-remarks-${poId}" placeholder="Optional remarks"></textarea>
            </div>
        </div>
        <div class="goods-form-actions">
            <button class="primary-button" onclick="saveGoodsReceipt(${poId}, this)">Save Receipt</button>
            <button class="cancel-vendor-btn" onclick="closeGoodsForm(${poId}, 'receive')">Cancel</button>
        </div>
    `;

    form.classList.remove("hidden");
}

function openReturnForm(poId) {
    const form        = document.getElementById(`return-form-${poId}`);
    const receiveForm = document.getElementById(`receive-form-${poId}`);

    receiveForm.classList.add("hidden");

    form.innerHTML = `
        <h3>Return Goods</h3>
        <div class="form-grid">
            <div class="form-group">
                <label>Return Quantity</label>
                <input type="number" id="return-quantity-${poId}" min="0.01" step="0.01" required>
            </div>
            <div class="form-group">
                <label>Return Date</label>
                <input type="date" id="return-date-${poId}" value="${todayISO()}" required>
            </div>
            <div class="form-group full-width">
                <label>Reason <span style="color:#c00;">*</span></label>
                <textarea id="return-reason-${poId}" required placeholder="Enter reason for return"></textarea>
            </div>
        </div>
        <div class="goods-form-actions">
            <button class="primary-button" onclick="saveGoodsReturn(${poId}, this)">Save Return</button>
            <button class="cancel-vendor-btn" onclick="closeGoodsForm(${poId}, 'return')">Cancel</button>
        </div>
    `;

    form.classList.remove("hidden");
}

// FIX: buttons disabled while the request is in flight (no double posting)
async function saveGoodsReceipt(poId, btn) {
    const receivedQuantity  = Number(document.getElementById(`received-quantity-${poId}`).value);
    const receiptDate       = document.getElementById(`receipt-date-${poId}`).value;
    const remarks           = document.getElementById(`receipt-remarks-${poId}`).value.trim();

    if (!receivedQuantity || receivedQuantity <= 0) {
        alert("Received quantity must be greater than 0."); return;
    }
    if (!receiptDate) {
        alert("Receipt date is required."); return;
    }

    if (btn) btn.disabled = true;

    try {
        const response = await apiFetch("/goods-received", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                po_id:              poId,
                receipt_date:       receiptDate,
                received_date:      receiptDate,     // column name, in case the backend expects it
                received_quantity:  receivedQuantity,
                remarks
            })
        });
        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            alert(result.message || "Failed to record receipt.");
            if (btn) btn.disabled = false;
            return;
        }

        alert("Goods received successfully.");
        await loadGoodsReceived();

    } catch (error) {
        console.error(error);
        alert("Failed to connect to procurement service.");
        if (btn) btn.disabled = false;
    }
}

async function saveGoodsReturn(poId, btn) {
    const returnQuantity = Number(document.getElementById(`return-quantity-${poId}`).value);
    const returnDate     = document.getElementById(`return-date-${poId}`).value;
    const returnReason   = document.getElementById(`return-reason-${poId}`).value.trim();

    if (!returnQuantity || returnQuantity <= 0) {
        alert("Return quantity must be greater than 0."); return;
    }
    if (!returnDate) {
        alert("Return date is required."); return;
    }
    if (!returnReason) {
        alert("Return reason is required."); return;
    }

    if (btn) btn.disabled = true;

    try {
        const response = await apiFetch("/goods-returns", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                po_id:           poId,
                return_quantity: returnQuantity,
                return_date:     returnDate,
                return_reason:   returnReason
            })
        });
        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            alert(result.message || "Failed to record return.");
            if (btn) btn.disabled = false;
            return;
        }

        alert("Return recorded successfully.");
        await loadGoodsReceived();

    } catch (error) {
        console.error(error);
        alert("Failed to connect to procurement service.");
        if (btn) btn.disabled = false;
    }
}

function closeGoodsForm(poId, type) {
    const form = document.getElementById(
        `${type === "receive" ? "receive" : "return"}-form-${poId}`
    );
    if (form) {
        form.classList.add("hidden");
        form.innerHTML = "";
    }
}

async function viewGoodsHistory(poId) {
    const history = document.getElementById(`history-${poId}`);

    if (!history.classList.contains("hidden")) {
        history.classList.add("hidden");
        return;
    }

    try {
        history.innerHTML = "Loading history...";
        history.classList.remove("hidden");

        const response = await apiFetch(`/goods-received/${poId}`);
        if (!response) return;
        const result   = await response.json();

        if (!response.ok) {
            history.innerHTML = escapeHtml(result.message || "Failed to load history");
            return;
        }

        const receipts = result.receipts || [];
        const returns  = result.returns  || [];

        history.innerHTML = `
            <h3>Receipt History</h3>
            ${receipts.length ? `
                <div class="table-container">
                    <table class="goods-history-table">
                        <thead>
                            <tr>
                                <th>Receipt ID</th>
                                <th>Receipt Date</th>
                                <th>Received Qty</th>
                                <th>Remarks</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${receipts.map(r => `
                                <tr>
                                    <td>${escapeHtml(String(r.receipt_id))}</td>
                                    <td>${formatDate(r.received_date)}</td>
                                    <td>${r.received_quantity}</td>
                                    <td>${escapeHtml(r.remarks || "-")}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>
                </div>
            ` : `<p>No receipts recorded.</p>`}

            <h3 class="return-history-title">Return History</h3>
            ${returns.length ? `
                <div class="table-container">
                    <table class="goods-history-table">
                        <thead>
                            <tr>
                                <th>Return ID</th>
                                <th>Return Date</th>
                                <th>Quantity</th>
                                <th>Reason</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${returns.map(r => `
                                <tr>
                                    <td>${escapeHtml(String(r.return_id))}</td>
                                    <td>${formatDate(r.return_date)}</td>
                                    <td>${r.return_quantity}</td>
                                    <td>${escapeHtml(r.return_reason)}</td>
                                </tr>
                            `).join("")}
                        </tbody>
                    </table>
                </div>
            ` : `<p>No returns recorded.</p>`}
        `;

    } catch (error) {
        console.error(error);
        history.innerHTML = "Failed to connect to procurement service.";
    }
}

async function completePO(poId, btn) {
    if (!confirm("Mark this Purchase Order as completed? This cannot be undone.")) return;

    if (btn) btn.disabled = true;

    try {
        const response = await apiFetch(`/purchase-orders/${poId}/complete`, {
            method: "POST",
            headers: { "Content-Type": "application/json" }
        });

        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            alert(result.message || "Failed to complete Purchase Order.");
            if (btn) btn.disabled = false;
            return;
        }

        alert("Purchase Order marked as completed.");
        await loadGoodsReceived();

    } catch (error) {
        console.error(error);
        alert("Failed to connect to procurement service.");
        if (btn) btn.disabled = false;
    }
}

// ─── Vendor Performance ────────────────────────────────────────────────────────

async function loadVendorPerformance() {
    const list = document.getElementById("vendorPerformanceList");
    const msg  = document.getElementById("vendorPerformanceMessage");

    list.innerHTML = "";
    msg.textContent = "Loading vendor performance...";

    try {
        const response = await apiFetch("/vendor-performance");
        if (!response) return;
        const result = await response.json();

        if (!response.ok || !result.success) {
            msg.textContent = result.message || "Failed to load vendor performance";
            return;
        }

        msg.textContent = "";
        renderVendorPerformance(result.vendors || []);

    } catch {
        msg.textContent = "Failed to connect to procurement service";
    }
}

// FIX: single handler for both blacklisting and un-blacklisting a vendor
async function toggleVendorBlacklist(vendorId, vendorName, isBlacklisted, btn) {
    const action = isBlacklisted ? "unblacklist" : "blacklist";
    const verb   = isBlacklisted ? "remove this vendor from the blacklist" : "blacklist this vendor";
 
    const reason = prompt(`Enter a reason to ${verb} (optional):`, "");
    if (reason === null) return;   // user cancelled the prompt
 
    if (!confirm(`Are you sure you want to ${verb}?\n\nVendor: ${vendorName}`)) return;
 
    if (btn) btn.disabled = true;
 
    try {
        const response = await apiFetch(`/vendor-performance/${vendorId}/${action}`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ reason: reason.trim() || null })
        });
        if (!response) return;
 
        const result = await response.json();
 
        if (!response.ok || !result.success) {
            alert(result.message || `Failed to ${action} vendor.`);
            if (btn) btn.disabled = false;
            return;
        }
 
        alert(isBlacklisted ? "Vendor removed from blacklist successfully." : "Vendor blacklisted successfully.");
        await loadVendorPerformance();
 
    } catch {
        alert("Failed to connect to procurement service.");
        if (btn) btn.disabled = false;
    }
}

// FIX: the click listener is attached ONCE here, not on every render
document.getElementById("vendorPerformanceList").addEventListener("click", async event => {
    const list = event.currentTarget;
 
    // Blacklist / Unblacklist button
    const blacklistBtn = event.target.closest(".vp-blacklist-btn");
    if (blacklistBtn) {
        const vendorId      = blacklistBtn.dataset.vendorId;
        const vendorName    = blacklistBtn.dataset.vendorName;
        const isBlacklisted = blacklistBtn.dataset.blacklisted === "1";
        await toggleVendorBlacklist(vendorId, vendorName, isBlacklisted, blacklistBtn);
        return;
    }
 
    const btn  = event.target.closest(".vp-drill-btn");
    if (!btn) return;
 
    const vendorId  = btn.dataset.vendorId;
    const detailRow = document.getElementById(`vpDetail-${vendorId}`);
    const content   = document.getElementById(`vpDetailContent-${vendorId}`);
    const isOpen    = detailRow.style.display !== "none";
 
    list.querySelectorAll(".vp-detail-row").forEach(r => r.style.display = "none");
    list.querySelectorAll(".vp-drill-btn").forEach(b => b.textContent = "Details ▾");
 
    if (isOpen) return;
 
    btn.textContent = "Details ▴";
    detailRow.style.display = "";
    content.innerHTML = "Loading...";
 
    try {
        const response = await apiFetch(`/vendor-performance/${vendorId}`);
        if (!response) return;
        const result   = await response.json();
 
        if (!response.ok || !result.success) {
            content.innerHTML = escapeHtml(result.message || "Failed to load details");
            return;
        }
 
        renderVendorPerformanceDetail(content, result);
 
    } catch {
        content.innerHTML = "Failed to connect to procurement service";
    }
});
 
function renderVendorPerformance(vendors) {
    const list = document.getElementById("vendorPerformanceList");
    list.innerHTML = "";
 
    if (!vendors.length) {
        list.innerHTML = `<div class="message">No vendors found.</div>`;
        return;
    }
 
    vendors.sort((a, b) => {
        if (!a.vendor_code && !b.vendor_code) return 0;
        if (!a.vendor_code) return 1;
        if (!b.vendor_code) return -1;
        return a.vendor_code.localeCompare(b.vendor_code, undefined, { numeric: true });
    });
 
    const table = document.createElement("div");
    table.className = "table-container";
    table.innerHTML = `
        <table>
            <thead>
                <tr>
                    <th>Vendor Code</th>
                    <th>Vendor Name</th>
                    <th>Total Orders</th>
                    <th>Order Value (₹)</th>
                    <th>Avg Lead Time</th>
                    <th>On-Time Delivery</th>
                    <th>Open Orders</th>
                    <th></th>
                </tr>
            </thead>
            <tbody>
                ${vendors.map(v => `
                    <tr${v.is_blacklisted ? ' class="vp-row-blacklisted"' : ""}>
                        <td>${escapeHtml(v.vendor_code || "-")}</td>
                        <td>
                            ${escapeHtml(v.vendor_name)}
                            ${v.is_blacklisted ? '<span class="vp-blacklisted-badge">Blacklisted</span>' : ""}
                        </td>
                        <td>${v.total_orders}</td>
                        <td>${formatCurrency(v.order_value)}</td>
                        <td>${v.avg_lead_time_days !== null ? `${v.avg_lead_time_days} days` : "-"}</td>
                        <td>
                            ${v.on_time_pct !== null
                                ? `<span class="${v.on_time_pct >= 80 ? "status-open" : v.on_time_pct >= 50 ? "status-vendor_selected" : "status-cancelled"}">
                                       ${v.on_time_pct}%
                                   </span>
                                   <span style="font-size:12px;color:#888;margin-left:6px;">
                                       (${v.on_time_deliveries}/${v.total_deliveries})
                                   </span>`
                                : "-"
                            }
                        </td>
                        <td>${v.open_orders > 0
                            ? `<span class="inquiry-status status-open">${v.open_orders}</span>`
                            : `<span style="color:#888;">0</span>`
                        }</td>
                        <td>
                            <div class="vp-actions">
                                <button
                                    type="button"
                                    class="option-button vp-drill-btn"
                                    data-vendor-id="${v.vendor_id}"
                                    data-vendor-name="${escapeHtml(v.vendor_name)}">
                                    Details ▾
                                </button>
                                <button
                                    type="button"
                                    class="option-button vp-blacklist-btn${v.is_blacklisted ? " vp-blacklist-btn--active" : ""}"
                                    data-vendor-id="${v.vendor_id}"
                                    data-vendor-name="${escapeHtml(v.vendor_name)}"
                                    data-blacklisted="${v.is_blacklisted ? "1" : "0"}">
                                    ${v.is_blacklisted ? "Unblacklist" : "Blacklist"}
                                </button>
                            </div>
                        </td>
                    </tr>
                    <tr class="vp-detail-row" id="vpDetail-${v.vendor_id}" style="display:none">
                        <td colspan="8">
                            <div class="vp-detail-container" id="vpDetailContent-${v.vendor_id}">
                                Loading...
                            </div>
                        </td>
                    </tr>
                `).join("")}
            </tbody>
        </table>
    `;
 
    list.appendChild(table);
}

function renderVendorPerformanceDetail(container, { profile, kpis, orders }) {
    const kpiCard = (label, value) => `
        <div class="inquiry-card-item">
            <span class="inquiry-card-label">${label}</span>
            <span class="inquiry-card-value">${value}</span>
        </div>
    `;

    container.innerHTML = `
        <div class="inquiry-card-info" style="margin-bottom:16px;">
            ${kpiCard("Total Orders",      kpis.total_orders)}
            ${kpiCard("Order Value",       `₹${formatCurrency(kpis.order_value)}`)}
            ${kpiCard("Avg Lead Time",     kpis.avg_lead_time_days !== null ? `${kpis.avg_lead_time_days} days` : "-")}
            ${kpiCard("On-Time Delivery",  kpis.on_time_pct !== null ? `${kpis.on_time_pct}% (${kpis.on_time_deliveries}/${kpis.total_deliveries})` : "-")}
            ${kpiCard("Open Orders",       kpis.open_orders)}
        </div>

        <div class="inquiry-card-info" style="margin-bottom:16px;">
            ${kpiCard("GST",           escapeHtml(profile.gst_number  || "-"))}
            ${kpiCard("PAN",           escapeHtml(profile.pan_number  || "-"))}
            ${kpiCard("Office",        escapeHtml(profile.office_address || "-"))}
            ${kpiCard("Sales Contact", escapeHtml(profile.sales_team_email || "-"))}
        </div>

        <div class="table-container">
            <table class="inquiry-vendor-table">
                <thead>
                    <tr>
                        <th>PO Number</th>
                        <th>PO Date</th>
                        <th>Item</th>
                        <th>Qty</th>
                        <th>Total Price</th>
                        <th>Expected Delivery</th>
                        <th>First Receipt</th>
                        <th>Lead Time</th>
                        <th>On Time?</th>
                        <th>Status</th>
                    </tr>
                </thead>
                <tbody>
                    ${orders.length ? orders.map(o => `
                        <tr>
                            <td>${escapeHtml(o.po_number  || "-")}</td>
                            <td>${formatDate(o.po_date)}</td>
                            <td>${escapeHtml(o.item_name  || "-")}</td>
                            <td>${o.qty} ${escapeHtml(o.unit || "")}</td>
                            <td>${formatCurrency(o.total_price)}</td>
                            <td>${formatDate(o.expected_delivery_date)}</td>
                            <td>${formatDate(o.first_received_date)}</td>
                            <td>${o.lead_time_days !== null ? `${o.lead_time_days} days` : "-"}</td>
                            <td>
                                ${o.first_received_date === null
                                    ? `<span style="color:#888;">Pending</span>`
                                    : o.delivered_on_time
                                        ? `<span class="status-open">✓ Yes</span>`
                                        : `<span class="status-cancelled">✗ No</span>`
                                }
                            </td>
                            <td>
                                <span class="inquiry-status status-${(o.status || "").toLowerCase()}">
                                    ${escapeHtml(o.status || "-")}
                                </span>
                            </td>
                        </tr>
                    `).join("") : `<tr><td colspan="10" style="text-align:center;color:#888;">No orders yet</td></tr>`}
                </tbody>
            </table>
        </div>
    `;
}

// ─── Past Price Reference ──────────────────────────────────────────────────────

let pprModelList     = [];
let pprSelected      = null;
let pprListenersBound = false;   // FIX: bind input/outside-click listeners only once

async function loadPPRModels() {
    try {
        const response = await apiFetch("/past-price-reference/models");
        if (!response) return;
        const result   = await response.json();

        if (!response.ok || !result.success) return;

        pprModelList = result.models || [];

    } catch {
        // silently fail — search will just show no results
    }
}

function initPastPriceReference() {
    const input    = document.getElementById("pprSearchInput");
    const dropdown = document.getElementById("pprDropdown");

    if (!pprModelList.length) loadPPRModels();

    input.value = "";
    document.getElementById("pprResults").innerHTML   = "";
    document.getElementById("pprMessage").textContent = "";
    document.getElementById("pprSelectedLabel").classList.add("hidden");
    dropdown.classList.add("hidden");
    pprSelected = null;

    if (!pprListenersBound) {
        input.addEventListener("input", onPPRInput);
        document.addEventListener("click", onPPROutsideClick);
        pprListenersBound = true;
    }
}

function onPPRInput() {
    const input    = document.getElementById("pprSearchInput");
    const dropdown = document.getElementById("pprDropdown");
    const query    = input.value.trim().toLowerCase();

    if (!query) {
        dropdown.classList.add("hidden");
        dropdown.innerHTML = "";
        return;
    }

    const matches = pprModelList.filter(m =>
        (m.model     || "").toLowerCase().includes(query) ||
        (m.item_name || "").toLowerCase().includes(query) ||
        (m.make      || "").toLowerCase().includes(query)
    );

    if (!matches.length) {
        dropdown.innerHTML = `<div class="ppr-dropdown-empty">No product found</div>`;
        dropdown.classList.remove("hidden");
        return;
    }

    dropdown.innerHTML = matches.slice(0, 30).map(m => `
        <div
            class="ppr-dropdown-item"
            data-model="${escapeHtml(m.model)}"
        >
            <span class="ppr-dropdown-model">${escapeHtml(m.model)}</span>
            <span class="ppr-dropdown-meta">
                ${escapeHtml(m.item_name || "")}
                ${m.make ? `· ${escapeHtml(m.make)}` : ""}
                ${m.product_category ? `· ${escapeHtml(m.product_category)}` : ""}
            </span>
        </div>
    `).join("");

    dropdown.classList.remove("hidden");

    dropdown.querySelectorAll(".ppr-dropdown-item").forEach(item => {
        item.addEventListener("click", () => {
            const model = item.dataset.model;
            document.getElementById("pprSearchInput").value = model;
            dropdown.classList.add("hidden");
            loadPPRDetail(model);
        });
    });
}

function onPPROutsideClick(e) {
    const dropdown = document.getElementById("pprDropdown");
    const input    = document.getElementById("pprSearchInput");
    if (dropdown && !dropdown.contains(e.target) && e.target !== input) {
        dropdown.classList.add("hidden");
    }
}

async function loadPPRDetail(model) {
    const results = document.getElementById("pprResults");
    const msg     = document.getElementById("pprMessage");
    const label   = document.getElementById("pprSelectedLabel");

    results.innerHTML = "";
    msg.textContent   = "Loading...";
    label.classList.add("hidden");

    try {
        const response = await apiFetch(`/past-price-reference?model=${encodeURIComponent(model)}`);
        if (!response) return;
        const result   = await response.json();

        if (!response.ok || !result.success) {
            msg.textContent = result.message || "Failed to fetch price history";
            return;
        }

        msg.textContent = "";

        if (!result.purchases.length) {
            msg.textContent = "No purchase history found for this model";
            return;
        }

        label.textContent = `Showing results for: ${result.model}${result.item_name ? ` — ${result.item_name}` : ""}${result.make ? ` (${result.make})` : ""}`;
        label.classList.remove("hidden");

        renderPPRResults(result);

    } catch {
        msg.textContent = "Failed to connect to procurement service";
    }
}

function renderPPRResults({ last3, byVendor }) {

    const results = document.getElementById("pprResults");

    const summaryHtml = `
        <div class="ppr-section">
            <h3 class="ppr-section-title">Last 3 Purchases Summary</h3>
            <div class="ppr-summary-strip">
                ${last3.map((p, i) => `
                    <div class="ppr-summary-card">
                        <div class="ppr-summary-rank">#${i + 1}</div>
                        <div class="ppr-summary-field">
                            <span class="ppr-summary-label">Date</span>
                            <span class="ppr-summary-value">${formatDate(p.po_date)}</span>
                        </div>
                        <div class="ppr-summary-field">
                            <span class="ppr-summary-label">Vendor</span>
                            <span class="ppr-summary-value">${escapeHtml(p.vendor_name)}</span>
                        </div>
                        <div class="ppr-summary-field">
                            <span class="ppr-summary-label">Qty</span>
                            <span class="ppr-summary-value">${p.qty} ${escapeHtml(p.unit || "")}</span>
                        </div>
                        <div class="ppr-summary-field">
                            <span class="ppr-summary-label">Unit Rate</span>
                            <span class="ppr-summary-value ppr-rate">₹${formatCurrency(p.price_per_unit)}</span>
                        </div>
                        <div class="ppr-summary-field">
                            <span class="ppr-summary-label">PO Number</span>
                            <span class="ppr-summary-value">${escapeHtml(p.po_number)}</span>
                        </div>
                    </div>
                `).join("")}
            </div>
        </div>
    `;

    const vendorSections = byVendor.map(v => `
        <div class="ppr-vendor-block">
            <div class="ppr-vendor-header">
                <span class="ppr-vendor-name">${escapeHtml(v.vendor_name)}</span>
                ${v.vendor_code
                    ? `<span class="ppr-vendor-code">${escapeHtml(v.vendor_code)}</span>`
                    : ""
                }
                <span class="ppr-vendor-count">${v.purchases.length} order${v.purchases.length !== 1 ? "s" : ""}</span>
            </div>
            <div class="table-container">
                <table class="inquiry-vendor-table">
                    <thead>
                        <tr>
                            <th>PO Number</th>
                            <th>Date</th>
                            <th>Qty</th>
                            <th>Unit</th>
                            <th>Unit Rate</th>
                            <th>Total Price</th>
                            <th>Status</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${v.purchases.map(p => `
                            <tr>
                                <td>${escapeHtml(p.po_number)}</td>
                                <td>${formatDate(p.po_date)}</td>
                                <td>${p.qty}</td>
                                <td>${escapeHtml(p.unit || "-")}</td>
                                <td class="ppr-rate-cell">₹${formatCurrency(p.price_per_unit)}</td>
                                <td>₹${formatCurrency(p.total_price)}</td>
                                <td>
                                    <span class="inquiry-status status-${(p.status || "").toLowerCase()}">
                                        ${escapeHtml(p.status || "-")}
                                    </span>
                                </td>
                            </tr>
                        `).join("")}
                    </tbody>
                </table>
            </div>
        </div>
    `).join("");

    results.innerHTML = summaryHtml + `
        <div class="ppr-section">
            <h3 class="ppr-section-title">Full History by Vendor</h3>
            ${vendorSections}
        </div>
    `;
}

// =========================================================
// MANAGEMENT INSIGHTS
// =========================================================

let procurementTrendChart = null;
let prPoTrendChart = null;
let poStatusChart = null;
let insightsRequestId = 0;   // FIX: ignore out-of-order responses

const insightsPeriod          = document.getElementById("insightsPeriod");
const insightsCustomRange     = document.getElementById("insightsCustomRange");
const insightsFromDate        = document.getElementById("insightsFromDate");
const insightsToDate          = document.getElementById("insightsToDate");
const applyInsightsDate       = document.getElementById("applyInsightsDate");
const managementInsightsMessage = document.getElementById("managementInsightsMessage");

// ---------------------------------------------------------
// PERIOD SELECT
// ---------------------------------------------------------

if (insightsPeriod) {
    insightsPeriod.addEventListener("change", () => {
        if (insightsPeriod.value === "custom") {
            insightsCustomRange.classList.remove("hidden");
            return;
        }
        insightsCustomRange.classList.add("hidden");
        loadManagementInsights();
    });
}

if (applyInsightsDate) {
    applyInsightsDate.addEventListener("click", () => {
        const from = insightsFromDate.value;
        const to   = insightsToDate.value;

        if (!from || !to) {
            managementInsightsMessage.textContent = "Please select both From and To dates.";
            return;
        }
        if (from > to) {
            managementInsightsMessage.textContent = "From date cannot be greater than To date.";
            return;
        }

        managementInsightsMessage.textContent = "";
        loadManagementInsights();
    });
}

// ---------------------------------------------------------
// LOAD  (fetch and render are separate so errors are labelled correctly)
// ---------------------------------------------------------

async function loadManagementInsights() {
    if (!managementInsightsMessage) return;

    const period = insightsPeriod.value;
    let url = `/management-insights?period=${encodeURIComponent(period)}`;

    if (period === "custom") {
        const from = insightsFromDate.value;
        const to   = insightsToDate.value;

        if (!from || !to) {
            managementInsightsMessage.textContent = "Please select a custom date range.";
            return;
        }

        url += `&from=${encodeURIComponent(from)}&to=${encodeURIComponent(to)}`;
    }

    const requestId = ++insightsRequestId;
    managementInsightsMessage.textContent = "Loading management insights...";

    let result;

    try {
        const response = await apiFetch(url);
        if (!response) return;
        result = await response.json();

        if (requestId !== insightsRequestId) return;   // a newer request superseded this one

        if (!response.ok || !result.success) {
            resetManagementInsights();
            managementInsightsMessage.textContent = result.message || "Failed to load management insights.";
            return;
        }
    } catch (error) {
        if (requestId !== insightsRequestId) return;
        console.error("Management Insights fetch error:", error);
        resetManagementInsights();
        managementInsightsMessage.textContent = "Failed to connect to procurement service.";
        return;
    }

    try {
        renderManagementInsights(result);
        managementInsightsMessage.textContent = "";
    } catch (error) {
        console.error("Management Insights render error:", error);
        managementInsightsMessage.textContent = "Loaded data but failed to display it. Check the console.";
    }
}

// FIX: clear stale numbers / charts when a load fails
function resetManagementInsights() {
    renderManagementInsights({});
}

function renderManagementInsights(data) {
    renderInsightOverview(data);
    renderInsightFinancials(data);
    renderInsightOperational(data);
    renderTopPerformers(data.top_performers || {});
    renderPaymentInsights(data.payment_types || []);
    renderProcurementTrend(data.trends?.daily_procurement || []);
    renderPRPOTrend(data.trends?.monthly_procurement || []);
    renderPOStatus(data.po_status || []);
}

const setText = (id, text) => {
    const el = document.getElementById(id);
    if (el) el.textContent = text;
};

const formatCount = (n) => Number(n || 0).toLocaleString("en-IN");

// Shows "No data" instead of a blank chart
function setChartEmpty(canvas, isEmpty) {
    const holder = canvas.parentElement;
    let note = holder.querySelector(".chart-empty-note");

    if (isEmpty) {
        if (!note) {
            note = document.createElement("div");
            note.className = "chart-empty-note";
            note.style.cssText = "position:absolute;inset:0;display:flex;align-items:center;justify-content:center;color:#888;font-size:14px;pointer-events:none;";
            note.textContent = "No data for this period";
            if (getComputedStyle(holder).position === "static") holder.style.position = "relative";
            holder.appendChild(note);
        }
    } else if (note) {
        note.remove();
    }
}

// "120 nos, 30 kg" from [{unit, <key>}]
function unitBreakdown(rows, key) {
    const parts = (rows || [])
        .filter(r => Number(r[key]) > 0)
        .map(r => `${Number(r[key]).toLocaleString("en-IN")} ${r.unit || ""}`.trim());
    return parts.length ? parts.join(", ") : "0";
}

// ---------------------------------------------------------
// OVERVIEW
// ---------------------------------------------------------

function renderInsightOverview(data) {
    const overview = data.overview || {};

    // FIX: PR card reads the real PR count, not PRs that happen to have a PO
    setText("insightTotalPRs",      formatCount(data.purchase_requests?.total));
    setText("insightTotalPOs",      formatCount(overview.total_pos));
    setText("insightDraftPOs",      formatCount(overview.draft_pos));
    setText("insightIssuedPOs",     formatCount(overview.issued_pos));
    setText("insightCompletedPOs",  formatCount(overview.completed_pos));
    setText("insightCancelledPOs",  formatCount(overview.cancelled_pos));
}

// ---------------------------------------------------------
// FINANCIAL
// ---------------------------------------------------------

function renderInsightFinancials(data) {
    const financial = data.financial || {};
    const overview  = data.overview  || {};

    setText("insightSalesValue",       formatRupee(financial.total_sales_value || 0));
    setText("insightProcurementValue", formatRupee(financial.total_procurement_value || 0));
    setText("insightGrossProfit",      formatRupee(financial.gross_profit || 0));
    setText("insightGrossMargin",      `${Number(financial.gross_margin_percentage || 0).toFixed(2)}%`);
    setText("insightAveragePO",        formatRupee(overview.average_po_value || 0));
    setText("insightHighestPO",        formatRupee(overview.highest_po_value || 0));

    // Red for a loss, default otherwise
    const isLoss = Number(financial.gross_profit || 0) < 0;
    ["insightGrossProfit", "insightGrossMargin"].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.style.color = isLoss ? "#c00" : "";
    });
}

// ---------------------------------------------------------
// OPERATIONAL
// ---------------------------------------------------------

function renderInsightOperational(data) {
    const eff = data.efficiency || {};

    // FIX: null from the backend means "no data", show "-" not "0 days"
    const days = (v, digits = 0) => (v === null || v === undefined ? "-" : `${Number(v).toFixed(digits)} days`);

    setText("insightAveragePRPO", days(eff.average_pr_to_po_days, 1));
    setText("insightFastestPRPO", days(eff.fastest_pr_to_po_days));
    setText("insightSlowestPRPO", days(eff.slowest_pr_to_po_days));

    // FIX: quantities shown per unit instead of one meaningless sum
    setText("insightGoodsReceived",     unitBreakdown(data.goods_received?.by_unit, "received"));
}

// ---------------------------------------------------------
// PROCUREMENT TREND CHART
// ---------------------------------------------------------

function renderProcurementTrend(rows) {
    const canvas = document.getElementById("procurementTrendChart");
    if (!canvas) return;

    if (procurementTrendChart) procurementTrendChart.destroy();

    setChartEmpty(canvas, !rows.length || rows.every(r => !r.sales_value && !r.procurement_value));

    const labels  = rows.map(r => r.day);
    const compact = new Intl.NumberFormat("en-IN", { notation: "compact", maximumFractionDigits: 1 });
    const spansYears = labels.length > 1 && labels[0].slice(0, 4) !== labels[labels.length - 1].slice(0, 4);

    const dayLabel = (iso, full) => {
        const d = new Date(`${iso}T00:00:00`);
        return d.toLocaleDateString("en-IN", {
            day: "2-digit",
            month: "short",
            ...(full || spansYears ? { year: "numeric" } : {})
        });
    };

    // Compact, stock-chart look: thin lines, no points, crosshair tooltip, value axis on the right
    const line = (label, key, color, fill) => ({
        label,
        data: rows.map(r => Number(r[key] || 0)),
        borderColor: color,
        backgroundColor: color + "22",
        borderWidth: 1.5,
        pointRadius: 0,
        pointHoverRadius: 4,
        tension: 0.25,
        fill
    });

    procurementTrendChart = new Chart(canvas, {
        type: "line",
        data: {
            labels,
            datasets: [
                line("Sales Value",       "sales_value",       "#2563eb", true),
                line("Procurement Value", "procurement_value", "#f59e0b", false),
                line("Gross Profit",      "gross_profit",      "#16a34a", false)
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            interaction: { mode: "index", intersect: false },
            plugins: {
                legend: {
                    position: "top",
                    align: "start",
                    labels: { usePointStyle: true, boxWidth: 8, font: { size: 11 } }
                },
                tooltip: {
                    callbacks: {
                        title: items => dayLabel(labels[items[0].dataIndex], true),
                        label: ctx => `${ctx.dataset.label}: ${formatRupee(ctx.raw)}`
                    }
                }
            },
            scales: {
                x: {
                    grid: { display: false },
                    ticks: {
                        autoSkip: true,
                        maxTicksLimit: 8,
                        maxRotation: 0,
                        font: { size: 11 },
                        callback: function (value) { return dayLabel(this.getLabelForValue(value)); }
                    }
                },
                y: {
                    position: "right",
                    grid: { color: "rgba(128,128,128,0.15)" },
                    ticks: {
                        maxTicksLimit: 5,
                        font: { size: 11 },
                        callback: value => (Number(value) < 0 ? "-" : "") + "₹" + compact.format(Math.abs(Number(value)))
                    }
                }
            }
        }
    });
}


// ---------------------------------------------------------
// PR VS PO CHART
// ---------------------------------------------------------

function renderPRPOTrend(rows) {
    const canvas = document.getElementById("prPoTrendChart");
    if (!canvas) return;

    if (prPoTrendChart) prPoTrendChart.destroy();

    setChartEmpty(canvas, !rows.length);

    prPoTrendChart = new Chart(canvas, {
        type: "bar",
        data: {
            labels: rows.map(r => r.month),
            datasets: [
                { label: "Purchase Requests", data: rows.map(r => Number(r.pr_count || 0)) },
                { label: "Purchase Orders",   data: rows.map(r => Number(r.po_count || 0)) }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            scales: {
                y: { beginAtZero: true, ticks: { precision: 0 } }
            }
        }
    });
}

// ---------------------------------------------------------
// PO STATUS CHART
// ---------------------------------------------------------

function renderPOStatus(rows) {
    const canvas = document.getElementById("poStatusChart");
    if (!canvas) return;

    if (poStatusChart) poStatusChart.destroy();

    setChartEmpty(canvas, !rows.length);

    poStatusChart = new Chart(canvas, {
        type: "doughnut",
        data: {
            labels: rows.map(r => String(r.status || "").replace(/_/g, " ")),
            datasets: [{ data: rows.map(r => Number(r.po_count || 0)) }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: { legend: { position: "bottom" } }
        }
    });
}

// ---------------------------------------------------------
// TABLES
// ---------------------------------------------------------

function renderInsightRows(bodyId, rows, colspan, emptyText, rowHtml) {
    const body = document.getElementById(bodyId);
    if (!body) return;

    if (!rows.length) {
        body.innerHTML = `<tr><td colspan="${colspan}">${emptyText}</td></tr>`;
        return;
    }

    body.innerHTML = rows.map(rowHtml).join("");
}

function renderTopPerformers(top) {
    const fill = (nameId, metaId, item, nameText, metaText) => {
        setText(nameId, item ? nameText : "-");
        setText(metaId, item ? metaText : "No orders in this period");
    };

    const meta = item => `${formatCount(item.po_count)} PO${item.po_count === 1 ? "" : "s"} · ${formatRupee(item.procurement_value)}`;

    fill("insightTopMake", "insightTopMakeMeta", top.make,
        top.make?.name, top.make && meta(top.make));

    fill("insightTopModel", "insightTopModelMeta", top.model,
        top.model?.name,
        top.model && `${top.model.make ? top.model.make + " · " : ""}${meta(top.model)}`);

    fill("insightTopVendor", "insightTopVendorMeta", top.vendor,
        top.vendor && (top.vendor.code ? `${top.vendor.name} (${top.vendor.code})` : top.vendor.name),
        top.vendor && meta(top.vendor));
}

function renderPaymentInsights(rows) {
    renderInsightRows("insightPaymentBody", rows, 3, "No payment data available", row => `
        <tr>
            <td>${escapeHtml(row.payment_type || "-")}</td>
            <td>${formatCount(row.po_count)}</td>
            <td>${formatRupee(row.procurement_value || 0)}</td>
        </tr>
    `);
}

// ─── Reports & Audits ──────────────────────────────────────────────────────

reportLogsTabButton.addEventListener("click", () => {
    if (activeLogTab === "report") return;
    activeLogTab = "report";
    reportLogsTabButton.classList.add("active");
    auditLogsTabButton.classList.remove("active");
    reportLogsSection.classList.remove("hidden");
    auditLogsSection.classList.add("hidden");
    logsPage = 1;
    loadLogs();
});

auditLogsTabButton.addEventListener("click", () => {
    if (activeLogTab === "audit") return;
    activeLogTab = "audit";
    auditLogsTabButton.classList.add("active");
    reportLogsTabButton.classList.remove("active");
    auditLogsSection.classList.remove("hidden");
    reportLogsSection.classList.add("hidden");
    logsPage = 1;
    loadLogs();
});

applyLogFilters.addEventListener("click", () => {
    logsPage = 1;
    loadLogs();
});

logsPrev.addEventListener("click", () => {
    if (logsPage > 1) {
        logsPage -= 1;
        loadLogs();
    }
});

logsNext.addEventListener("click", () => {
    if (logsPage < logsTotalPages) {
        logsPage += 1;
        loadLogs();
    }
});

async function loadLogs() {
    logsMessage.textContent = "Loading...";
    logsPrev.disabled = true;
    logsNext.disabled = true;

    const endpoint = activeLogTab === "report" ? "/report-logs" : "/audit-logs";

    const params = new URLSearchParams({ page: logsPage, limit: 25 });
    if (logUsernameFilter.value.trim()) params.set("username", logUsernameFilter.value.trim());
    if (logFromDate.value)              params.set("from", logFromDate.value);
    if (logToDate.value)                params.set("to", logToDate.value);

    try {
        const response = await apiFetch(`${endpoint}?${params.toString()}`);
        if (!response) return;
        const result   = await response.json();

        if (!response.ok || !result.success) {
            logsMessage.textContent = result.message || "Failed to load logs";
            return;
        }

        logsMessage.textContent = "";

        if (activeLogTab === "report") {
            renderReportLogs(result.rows || []);
        } else {
            renderAuditLogs(result.rows || []);
        }

        const { page, total_pages, has_prev, has_next } = result.pagination;
        logsPage       = page;
        logsTotalPages = total_pages;

        logsPageLabel.textContent = `Page ${page} of ${total_pages}`;
        logsPrev.disabled = !has_prev;
        logsNext.disabled = !has_next;

    } catch {
        logsMessage.textContent = "Failed to connect to procurement service";
    }
}

function renderReportLogs(rows) {
    if (!rows.length) {
        reportLogsBody.innerHTML = `<tr class="logs-empty-row"><td colspan="4">No report logs found</td></tr>`;
        return;
    }

    reportLogsBody.innerHTML = rows.map(row => `
        <tr>
            <td>${formatDateTime(row.log_timestamp)}</td>
            <td>${escapeHtml(row.username)}</td>
            <td>${escapeHtml(row.action)}</td>
            <td class="logs-report-cell">${escapeHtml(row.report)}</td>
        </tr>
    `).join("");
}

function renderAuditLogs(rows) {
    if (!rows.length) {
        auditLogsBody.innerHTML = `<tr class="logs-empty-row"><td colspan="5">No audit logs found</td></tr>`;
        return;
    }

    auditLogsBody.innerHTML = rows.map(row => `
        <tr>
            <td>${formatDateTime(row.log_timestamp)}</td>
            <td>${escapeHtml(row.username)}</td>
            <td>${escapeHtml(row.action)}</td>
            <td class="logs-value-cell audit-old-value">${escapeHtml(row.old_value || "-")}</td>
            <td class="logs-value-cell audit-new-value">${escapeHtml(row.new_value || "-")}</td>
        </tr>
    `).join("");
}

// Inline onclick handlers need these on window
window.completePO       = completePO;
window.openReceiveForm  = openReceiveForm;
window.openReturnForm   = openReturnForm;
window.saveGoodsReceipt = saveGoodsReceipt;
window.saveGoodsReturn  = saveGoodsReturn;
window.closeGoodsForm   = closeGoodsForm;
window.viewGoodsHistory = viewGoodsHistory;