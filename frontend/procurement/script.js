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
const vendorInquiryList = document.getElementById("vendorInquiryList");

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

// Vendor Inquiries page — browser exposes element IDs as globals, but we
// declare it explicitly so the reference is clear and lint-safe.
const vendorInquiriesPage = document.getElementById("vendorInquiriesPage");

// Logout
const logoutButton = document.getElementById("logoutButton");

// ─── State ─────────────────────────────────────────────────────────────────────

let importedRows = [];
let orderTrackingPage = 1;
let orderTrackingTotalPages = 1;

// ─── Utilities ─────────────────────────────────────────────────────────────────

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
    // Server returns MySQL DATE columns as plain YYYY-MM-DD strings (dateStrings: true).
    // Parse the date parts directly to avoid UTC-vs-local timezone shifts.
    const parts = String(value).match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (parts) {
        const [, y, m, d] = parts;
        return `${d}/${m}/${y}`;
    }
    // Fallback for unexpected formats
    const dt = new Date(value);
    return isNaN(dt.getTime()) ? String(value) : dt.toLocaleDateString("en-IN");
}

function formatCurrency(value) {
    const n = Number(value);
    return isNaN(n) ? "-" : n.toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// FIX: always derive today's date in IST, not UTC.
// new Date().toISOString() is UTC and gives yesterday's date in IST between
// midnight and 05:30 IST.
function todayISO() {
    const d = new Date();
    const ist = new Date(d.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }));
    return `${ist.getFullYear()}-${String(ist.getMonth() + 1).padStart(2, "0")}-${String(ist.getDate()).padStart(2, "0")}`;
}

// ─── Fetch wrapper: handle 401 / non-JSON (session expiry) ────────────────────
// FIX: instead of silently swallowing "Failed to connect" on every non-JSON
// response, redirect to the auth service when the session has expired.

const AUTH_URL = "http://localhost:5000"; // must match authServiceUrl on the backend

async function apiFetch(url, options = {}) {
    const response = await fetch(url, { credentials: "include", ...options });
    if (response.status === 401) {
        window.location.href = AUTH_URL;
        return null; // redirect already happened
    }
    return response;
}

// Prevent scroll from changing number input values
document.addEventListener("wheel", event => {
    if (document.activeElement?.type === "number") {
        document.activeElement.blur();
    }
}, { passive: true });

// ─── Purchase Request: Taxable Value ───────────────────────────────────────────

function calculateTaxableValue() {
    const qty  = Number(qtyInput.value)       || 0;
    const rate = Number(salesRateInput.value) || 0;
    taxableValueInput.value = (qty * rate).toFixed(2);
}

// Guard: inputs may not exist on every page variant.
if (qtyInput)       qtyInput.addEventListener("input", calculateTaxableValue);
if (salesRateInput) salesRateInput.addEventListener("input", calculateTaxableValue);

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

        const todayStr = todayISO(); // FIX: use IST-safe todayISO()
        importedRows.forEach(row => {
            if (!row.pr_date || String(row.pr_date).trim() === "") {
                row.pr_date = todayStr;
            }
        });

        renderPreview();
        showMessage(`${importedRows.length} rows ready for review. Edit if needed, then click Save.`);

    } catch {
        showMessage("Failed to connect to procurement service");
    } finally {
        importButton.disabled = false;
    }
});

// ─── PR Excel: Render Editable Preview Table ───────────────────────────────────

function renderPreview() {
    previewBody.innerHTML = "";

    importedRows.forEach((row, index) => {
        const tr = document.createElement("tr");

        // FIX: pr_date uses type="date" so the browser validates the format
        // and the user cannot type free-text like "28/09/2026".
        tr.innerHTML = `
            <td><input type="date" data-index="${index}" data-field="pr_date" value="${escapeHtml(row.pr_date || "")}"></td>
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

// ─── PR Excel: Save Imported Rows ─────────────────────────────────────────────

saveImportedButton.addEventListener("click", async () => {
    if (!importedRows.length) {
        showMessage("No rows available to save");
        return;
    }

    // FIX: validate every row before sending to the server.
    const dateRe = /^\d{4}-\d{2}-\d{2}$/;
    for (const [i, r] of importedRows.entries()) {
        const rowNum = i + 1;
        if (!dateRe.test(String(r.pr_date || "").trim())) {
            showMessage(`Row ${rowNum}: date must be YYYY-MM-DD (use the date picker)`);
            return;
        }
        if (!(Number(r.qty) > 0)) {
            showMessage(`Row ${rowNum}: quantity must be greater than 0`);
            return;
        }
        if (!(Number(r.sales_rate) > 0)) {
            showMessage(`Row ${rowNum}: sales rate must be greater than 0`);
            return;
        }
    }

    try {
        saveImportedButton.disabled = true;
        showMessage("Saving purchase requests...");

        const response = await apiFetch("/purchase-requests/import", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ rows: importedRows })
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

    // FIX: "last three years" should be the three completed financial years.
    // In September 2026, the current FY 26-27 is not finished, so start with
    // 25-26 (startYear = 2025 when month < 4, else year-1).
    const currentFYStart = month >= 4 ? year : year - 1;
    const startYear      = currentFYStart - 1; // most recent *completed* FY

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

// ─── Vendor Masters: Prefill recommended_by / approved_by ─────────────────────
// FIX: read currentUsername at the time the form initialises, not at script
// load time (it may not be set yet when the script first runs).

function prefillVendorFormMeta() {
    const regDate = document.getElementById("registration_date");
    const recBy   = document.getElementById("recommended_by");
    const appBy   = document.getElementById("approved_by");

    if (regDate && !regDate.value) regDate.value = todayISO();
    if (recBy   && !recBy.value)   recBy.value   = window.currentUsername || "";
    if (appBy   && !appBy.value)   appBy.value   = window.currentUsername || "";
}

prefillVendorFormMeta();

// ─── Vendor Masters: GST Duplicate Check ──────────────────────────────────────

// FIX: track whether a GST alert has already been shown in this blur cycle so
// submit doesn't pop a second popup.
let _lastGSTChecked = "";
let _lastGSTValid   = true;

async function checkVendorGST(gstNumber) {
    const trimmed = (gstNumber || "").trim();
    if (!trimmed) return true;

    // Return cached result for the same value so blur + submit don't double-alert.
    if (trimmed === _lastGSTChecked) return _lastGSTValid;

    try {
        const response = await apiFetch(`/vendors/check-gst?gst_number=${encodeURIComponent(trimmed)}`);
        if (!response) return false;

        const result = await response.json();

        if (!response.ok) {
            showVendorMessage(result.message || "Failed to check GST number");
            _lastGSTChecked = trimmed;
            _lastGSTValid   = false;
            return false;
        }

        if (result.exists) {
            alert(`A company with this GST number already exists: ${result.vendor_name}`);
            _lastGSTChecked = trimmed;
            _lastGSTValid   = false;
            return false;
        }

        _lastGSTChecked = trimmed;
        _lastGSTValid   = true;
        return true;

    } catch {
        showVendorMessage("Failed to connect to procurement service");
        _lastGSTChecked = "";
        _lastGSTValid   = false;
        return false;
    }
}

document.getElementById("gst_number").addEventListener("blur", async event => {
    await checkVendorGST(event.target.value);
});

// ─── Vendor Masters: Manual Form Validation ────────────────────────────────────

// FIX: validate required fields on the frontend before the fetch so users get
// a clear field-level message instead of a generic backend error.
const VENDOR_REQUIRED_FIELDS = [
    ["registration_date",                         "Registration Date"],
    ["vendor_name",                               "Vendor Name"],
    ["office_address",                            "Office Address"],
    ["office_name",                               "Office Contact Name"],
    ["office_phone",                              "Office Contact Number"],
    ["legal_entity",                              "Legal Entity"],
    ["commercial_role",                           "Commercial Role"],
    ["year_of_incorporation",                     "Year of Incorporation"],
    ["director_name",                             "Director / CEO Name"],
    ["director_designation",                      "Director / CEO Designation"],
    ["director_mobile",                           "Director / CEO Mobile"],
    ["director_email",                            "Director / CEO Email"],
    ["sales_name",                                "Sales Team Name"],
    ["sales_contact",                             "Sales Team Contact"],
    ["sales_email",                               "Sales Team Email"],
    ["accounts_name",                             "Accounts Team Name"],
    ["accounts_contact",                          "Accounts Team Contact"],
    ["accounts_email",                            "Accounts Team Email"],
    ["gst_number",                                "GST Number"],
    ["pan_number",                                "PAN Number"],
    ["bank_name",                                 "Bank Name"],
    ["account_no",                                "Bank Account Number"],
    ["bank_branch",                               "Bank Branch"],
    ["account_type",                              "Account Type"],
    ["ifsc_rtgs",                                 "IFSC / RTGS Code"],
    ["branch1_address",                           "Branch Office 1 Address"],
    ["turnover_value_1",                          "Turnover Year 1 Value"],
    ["recommended_by",                            "Recommended By"],
    ["approved_by",                               "Approved By"]
];

const VENDOR_REQUIRED_DOCUMENTS = [
    ["gst_document",          "GST Document"],
    ["pan_document",          "PAN Document"],
    ["itr_last_year_document","ITR (Last Year)"]
];

function validateVendorForm() {
    for (const [id, label] of VENDOR_REQUIRED_FIELDS) {
        const el = document.getElementById(id);
        if (!el || !el.value.trim()) {
            showVendorMessage(`${label} is required`);
            el?.focus();
            return false;
        }
    }

    for (const [id, label] of VENDOR_REQUIRED_DOCUMENTS) {
        const el = document.getElementById(id);
        if (!el || !el.files?.[0]) {
            showVendorMessage(`${label} document is required`);
            return false;
        }
    }

    return true;
}

// ─── Vendor Masters: Excel File Chosen ────────────────────────────────────────

vendorExcelFile.addEventListener("change", () => {
    vendorExcelFileName.textContent = vendorExcelFile.files.length
        ? vendorExcelFile.files[0].name
        : "No file chosen";
});

// ─── Vendor Masters: Manual Form Submit ───────────────────────────────────────

vendorManualForm.addEventListener("submit", async event => {
    event.preventDefault();

    // FIX: run frontend validation first.
    if (!validateVendorForm()) return;
    if (!(await checkVendorGST(document.getElementById("gst_number").value))) return;

    const turnoverLabel = (inputId) => {
        const input = document.getElementById(inputId);
        const label = input?.closest(".form-group")?.querySelector("label");
        return label ? label.textContent.replace(" Turnover", "").trim() : "";
    };

    const field = id => document.getElementById(id)?.value ?? "";

    const data = {
        // FIX: read currentUsername at save time so it's never blank.
        registration_date:                         field("registration_date") || todayISO(),
        vendor_name:                               field("vendor_name"),

        office_address:                            field("office_address"),
        office_contact_name:                       field("office_name"),
        office_contact_number:                     field("office_phone"),
        factory_address:                           field("factory_address"),
        factory_contact_name:                      field("factory_name"),
        factory_contact_number:                    field("factory_phone"),
        warehouse_address:                         field("warehouse_address"),
        warehouse_contact_name:                    field("warehouse_name"),
        warehouse_contact_number:                  field("warehouse_phone"),
        workshop_address:                          field("workshop_address"),
        workshop_contact_name:                     field("workshop_name"),
        workshop_contact_number:                   field("workshop_phone"),

        legal_entity:                              field("legal_entity"),
        commercial_role:                           field("commercial_role"),
        year_of_incorporation:                     field("year_of_incorporation"),

        director_or_ceo_or_management_name:        field("director_name"),
        director_or_ceo_or_management_designation: field("director_designation"),
        director_or_ceo_or_management_mobile_no:   field("director_mobile"),
        director_or_ceo_or_management_email:       field("director_email"),
        director_or_ceo_or_management_web_address: field("director_web"),

        sales_team_name:                           field("sales_name"),
        sales_team_contact:                        field("sales_contact"),
        sales_team_email:                          field("sales_email"),
        accounts_team_name:                        field("accounts_name"),
        accounts_team_contact:                     field("accounts_contact"),
        accounts_team_email:                       field("accounts_email"),

        gst_number:                                field("gst_number"),
        pan_number:                                field("pan_number"),
        msme_number:                               field("msme_number"),

        bank_name:                                 field("bank_name"),
        bank_account_no:                           field("account_no"),
        bank_branch:                               field("bank_branch"),
        bank_account_type:                         field("account_type"),
        bank_ifsc:                                 field("ifsc_rtgs"),

        branch_office_1_address:                   field("branch1_address"),
        branch_office_1_contact_name:              field("branch1_name"),
        branch_office_1_contact_number:            field("branch1_contact"),
        branch_office_2_address:                   field("branch2_address"),
        branch_office_2_contact_name:              field("branch2_name"),
        branch_office_2_contact_number:            field("branch2_contact"),
        branch_office_3_address:                   field("branch3_address"),
        branch_office_3_contact_name:              field("branch3_name"),
        branch_office_3_contact_number:            field("branch3_contact"),

        turnover_year_1:                           turnoverLabel("turnover_value_1"),
        turnover_value_1:                          field("turnover_value_1"),
        turnover_year_2:                           turnoverLabel("turnover_value_2"),
        turnover_value_2:                          field("turnover_value_2"),
        turnover_year_3:                           turnoverLabel("turnover_value_3"),
        turnover_value_3:                          field("turnover_value_3"),

        recommended_by:                            field("recommended_by") || window.currentUsername || "",
        approved_by:                               field("approved_by")    || window.currentUsername || ""
    };

    const formData = new FormData();
    formData.append("vendor_data", JSON.stringify(data));

    ["gst_document", "pan_document", "msme_document",
     "itr_last_year_document", "itr_second_last_year_document", "itr_third_last_year_document"
    ].forEach(f => {
        const file = document.getElementById(f)?.files[0];
        if (file) formData.append(f, file);
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
        prefillVendorFormMeta();
        _lastGSTChecked = "";
        _lastGSTValid   = true;

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
        showVendorMessage("Failed to connect to procurement service");
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

        fields.forEach(([f, label]) => {
            used.add(f);

            const isDate = f === "registration_date";
            const value  = vendor[f];
            const shown  = isDate ? (value || todayISO()) : (value || "");

            const group = document.createElement("div");
            group.className = "form-group";
            group.innerHTML = `
                <label>${escapeHtml(label)}</label>
                <input id="preview_${f}" type="${isDate ? "date" : "text"}" value="${escapeHtml(shown)}">
            `;
            grid.appendChild(group);
        });

        section.appendChild(grid);
        vendorPreviewForm.appendChild(section);
    };

    VENDOR_SECTIONS.forEach(section => buildSection(section.title, section.fields));

    const extraFields = Object.keys(vendor)
        .filter(f => !used.has(f))
        .map(f => [f, formatVendorFieldName(f)]);

    if (extraFields.length) buildSection("Other Details", extraFields);

    const recommendedInput = document.getElementById("preview_recommended_by");
    const approvedInput    = document.getElementById("preview_approved_by");

    // FIX: read currentUsername at render time.
    if (recommendedInput && !recommendedInput.value) recommendedInput.value = window.currentUsername || "";
    if (approvedInput    && !approvedInput.value)    approvedInput.value    = window.currentUsername || "";

    vendorExcelPreview.classList.remove("hidden");

    const gstInput = document.getElementById("preview_gst_number");

    if (gstInput) {
        // Reset GST cache so the Excel preview gets a fresh check.
        _lastGSTChecked = "";
        const gstValid = await checkVendorGST(gstInput.value);
        saveVendorExcelButton.disabled = !gstValid;
        return;
    }

    saveVendorExcelButton.disabled = false;
}

function formatVendorFieldName(f) {
    return f
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
    // FIX: read currentUsername at save time.
    if (!data.recommended_by) data.recommended_by = window.currentUsername || "";
    if (!data.approved_by)    data.approved_by    = window.currentUsername || "";

    const vendorName = data.vendor_name || "Vendor";

    const formData = new FormData();
    formData.append("vendor_data", JSON.stringify(data));

    ["gst_document", "pan_document", "msme_document",
     "itr_last_year_document", "itr_second_last_year_document", "itr_third_last_year_document"
    ].forEach(f => {
        const file = document.getElementById(`excel_${f}`)?.files[0];
        if (file) formData.append(f, file);
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
                // FIX: re-enable the button so the user can correct and retry.
                saveVendorExcelButton.disabled = false;
            } else {
                showVendorMessage(result.message || "Failed to save vendor");
                saveVendorExcelButton.disabled = false;
            }
            return;
        }

        alert(`Vendor "${vendorName}" added successfully!\nVendor Code: ${result.vendor_code}`);
        showVendorMessage("");

        vendorExcelPreview.classList.add("hidden");
        vendorPreviewForm.innerHTML = "";
        vendorExcelFile.value = "";
        vendorExcelFileName.textContent = "No file chosen";
        _lastGSTChecked = "";
        _lastGSTValid   = true;

        ["excel_gst_document", "excel_pan_document", "excel_msme_document",
         "excel_itr_last_year_document", "excel_itr_second_last_year_document", "excel_itr_third_last_year_document"
        ].forEach(id => {
            const el = document.getElementById(id);
            if (el) el.value = "";
        });

    } catch {
        showVendorMessage("Failed to connect to procurement service");
        saveVendorExcelButton.disabled = false;
    }
});

// ─── Vendor Inquiries: Load List ───────────────────────────────────────────────

async function loadVendorInquiries() {
    vendorInquiryList.innerHTML = "Loading inquiries...";

    try {
        const response = await apiFetch("/vendor-inquiries");
        if (!response) return;

        const result = await response.json();

        if (!response.ok || !result.success) {
            vendorInquiryList.textContent = result.message || "Failed to load inquiries";
            return;
        }

        renderVendorInquiries(result.inquiries || []);

    } catch {
        vendorInquiryList.textContent = "Failed to connect to procurement service";
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

        card.dataset.prNumber = inquiry.pr_number  || "";
        card.dataset.itemName = inquiry.item_name  || "";

        card.innerHTML = `
            <div class="inquiry-card-info">
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">PR Number</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.pr_number || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">PR Date</span>
                    <span class="inquiry-card-value">${formatDate(inquiry.pr_date)}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Party Name</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.party_name || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Location</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.location || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Territory</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.territory || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Product Category</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.product_category || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Item Name</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.item_name || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Make / Model</span>
                    <span class="inquiry-card-value">
                        ${escapeHtml(inquiry.make || "-")} / ${escapeHtml(inquiry.model || "-")}
                    </span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Quantity</span>
                    <span class="inquiry-card-value">
                        ${escapeHtml(String(inquiry.qty ?? 0))} ${escapeHtml(inquiry.unit || "")}
                    </span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Sales Rate</span>
                    <span class="inquiry-card-value">${formatCurrency(inquiry.sales_rate)}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Taxable Value</span>
                    <span class="inquiry-card-value">${formatCurrency(inquiry.taxable_value)}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Product Remarks</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.product_remarks || "-")}</span>
                </div>
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

            <!-- Inline add-vendor form (hidden by default) -->
            <div class="add-vendor-form" id="addVendorForm-${id}" style="display:none">
                <h3>Add Vendor Quotation</h3>

                <div class="form-group">
                    <label>Vendor</label>
                    <select class="vendor-select" required>
                        <option value="">Select Vendor</option>
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
                    <input type="date" class="vendor-delivery-date" min="${todayISO()}">
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

            <!-- Vendor quotation table -->
            <div class="inquiry-vendor-section" id="vendorsSection-${id}" hidden></div>
        `;

        vendorInquiryList.appendChild(card);

        // FIX: load existing quotations immediately so returning users see them.
        loadInquiryVendors(id);

        // FIX: wire up the form's live-calculation listeners once, here, so
        // they are never duplicated no matter how many times the form is
        // opened or closed.
        wireVendorForm(id, Number(inquiry.qty ?? 0));
    });
}

// ─── Vendor Inquiries: Wire Form Listeners (once per card) ────────────────────
// FIX: all input listeners are attached once at render time, not inside the
// click handler. This removes stacked listener accumulation and ensures the
// advance preview recalculates correctly when either price OR advance % changes.

function wireVendorForm(inquiryId, qty) {
    const form         = document.getElementById(`addVendorForm-${inquiryId}`);
    if (!form) return;

    const priceEl          = form.querySelector(".vendor-price");
    const totalDisplayEl   = form.querySelector(".vendor-total-display");
    const advancePctEl     = form.querySelector(".vendor-advance-pct");
    const advancePreviewGp = form.querySelector(".vendor-advance-preview-group");
    const advancePreviewEl = form.querySelector(".vendor-advance-preview");
    const balanceDaysGroup = form.querySelector(".balance-days-group");
    const balanceDaysEl    = form.querySelector(".vendor-balance-days");

    function recalcTotals() {
        const price = Number(priceEl.value) || 0;
        const pct   = Number(advancePctEl.value);

        // Total price
        if (price > 0 && qty > 0) {
            const total = price * qty;
            totalDisplayEl.value = `₹${total.toLocaleString("en-IN", {
                minimumFractionDigits: 2,
                maximumFractionDigits: 2
            })}`;

            // Advance amount — recalculate whenever price changes too
            if (pct > 0) {
                const amount = (pct / 100) * total;
                advancePreviewEl.value = `₹${amount.toLocaleString("en-IN", {
                    minimumFractionDigits: 2,
                    maximumFractionDigits: 2
                })}`;
                advancePreviewGp.style.display = "";
            } else {
                advancePreviewGp.style.display = "none";
                advancePreviewEl.value = "";
            }
        } else {
            totalDisplayEl.value   = "";
            advancePreviewGp.style.display = "none";
            advancePreviewEl.value = "";
        }

        // Balance days visibility
        balanceDaysGroup.style.display = (pct > 0 && pct < 100) ? "" : "none";
        if (pct >= 100 || pct === 0) balanceDaysEl.value = "";
    }

    priceEl.addEventListener("input", recalcTotals);
    advancePctEl.addEventListener("input", recalcTotals);
}

// ─── Vendor Inquiries: Delegated Click Handler ─────────────────────────────────
// FIX: this listener is attached once (not inside renderVendorInquiries).

vendorInquiryList.addEventListener("click", async event => {

    // ── Add Vendor button → open form + load dropdown ──
    const addBtn = event.target.closest(".add-vendor-btn");
    if (addBtn) {
        const inquiryId = addBtn.dataset.inquiryId;
        const form = document.getElementById(`addVendorForm-${inquiryId}`);
        if (!form) return;

        form.style.display = "";

        // FIX: load vendors and filter out already-quoted ones.
        // The /vendor-inquiries/:id response includes vendor_id on each row.
        const select = form.querySelector(".vendor-select");
        await loadVendorsIntoDropdown(select, inquiryId);

        return;
    }

    // ── Cancel → hide form ──
    const cancelBtn = event.target.closest(".cancel-vendor-btn");
    if (cancelBtn) {
        const inquiryId = cancelBtn.dataset.inquiryId;
        const form = document.getElementById(`addVendorForm-${inquiryId}`);
        if (form) form.style.display = "none";
        return;
    }

    // ── Save Vendor ──
    const saveBtn = event.target.closest(".save-vendor-btn");
    if (saveBtn) {
        await saveVendorForInquiry(saveBtn.dataset.inquiryId);
        return;
    }
});

// ─── Vendor Inquiries: Populate Vendor Dropdown ────────────────────────────────
// FIX: fetch already-quoted vendor IDs and exclude them from the dropdown so
// the user cannot accidentally pick a duplicate (which would get a 409 error
// from the backend).

async function loadVendorsIntoDropdown(selectElement, inquiryId) {
    selectElement.innerHTML = '<option value="">Loading vendors...</option>';

    try {
        // Fetch all eligible vendors and already-quoted vendors in parallel.
        const [vendorsRes, inquiryRes] = await Promise.all([
            apiFetch("/vendors"),
            apiFetch(`/vendor-inquiries/${inquiryId}`)
        ]);

        if (!vendorsRes || !inquiryRes) return;

        const [vendorsResult, inquiryResult] = await Promise.all([
            vendorsRes.json(),
            inquiryRes.json()
        ]);

        if (!vendorsRes.ok || !vendorsResult.success) {
            throw new Error(vendorsResult.message || "Failed to load vendors");
        }

        // Build a Set of already-quoted vendor IDs.
        const quotedIds = new Set(
            (inquiryResult.vendors || []).map(v => v.vendor_id)
        );

        const vendors = (vendorsResult.vendors || [])
            .filter(v => !quotedIds.has(v.vendor_id))
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
    const msgEl            = form.querySelector(".vendor-form-message");
    const saveBtn          = form.querySelector(".save-vendor-btn");

    const vendorId   = select.value;
    const vendorName = select.options[select.selectedIndex]?.text || "";
    const price      = Number(priceEl.value);
    const advancePct = advancePctEl.value !== "" ? Number(advancePctEl.value) : null;

    if (!vendorId)                  { msgEl.textContent = "Please select a vendor."; return; }
    if (!price || price <= 0)       { msgEl.textContent = "Please enter a valid price per unit."; return; }
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
        // FIX: disable button immediately to prevent double-submit.
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
                payment_terms_remarks:  paymentRemarksEl.value.trim() || null
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

        // Reset form fields.
        select.value           = "";
        priceEl.value          = "";
        deliveryDateEl.value   = "";
        advancePctEl.value     = "";
        balanceDaysEl.value    = "";
        paymentRemarksEl.value = "";
        form.querySelector(".balance-days-group").style.display           = "none";
        form.querySelector(".vendor-advance-preview-group").style.display = "none";
        form.querySelector(".vendor-advance-preview").value               = "";
        form.querySelector(".vendor-total-display").value                 = "";

        form.style.display = "none";

        // Reload the quotations table for this card.
        await loadInquiryVendors(inquiryId);

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
    section.hidden = false;

    try {
        const response = await apiFetch(`/vendor-inquiries/${inquiryId}`);
        if (!response) return;

        const result = await response.json();

        if (!response.ok || !result.success) {
            section.innerHTML = `<em style='font-size:13px;color:#c00;'>${escapeHtml(result.message || "Failed to load vendors")}</em>`;
            return;
        }

        renderInquiryVendorTable(section, result.vendors || []);

    } catch {
        section.innerHTML = "<em style='font-size:13px;color:#c00;'>Failed to connect to procurement service</em>";
    }
}

function renderInquiryVendorTable(container, vendors) {
    if (!vendors.length) { container.innerHTML = ""; container.hidden = true; return; }

    // FIX: the backend returns payment_terms_remarks, not remarks, for the
    // payment notes column. Map the right field.
    container.innerHTML = `
        <div class="table-container">
            <table class="inquiry-vendor-table">
                <thead>
                    <tr>
                        <th>Vendor Code</th>
                        <th>Vendor Name</th>
                        <th>Price / Unit</th>
                        <th>Total Price</th>
                        <th>Delivery Date</th>
                        <th>Payment Type</th>
                        <th>Advance</th>
                        <th>Payment Remarks</th>
                    </tr>
                </thead>
                <tbody>
                    ${vendors.map(v => `
                        <tr>
                            <td>${escapeHtml(v.vendor_code || "-")}</td>
                            <td>${escapeHtml(v.vendor_name || "-")}</td>
                            <td>${formatCurrency(v.price_per_unit)}</td>
                            <td>${formatCurrency(v.total_price)}</td>
                            <td>${formatDate(v.expected_delivery_date)}</td>
                            <td>${escapeHtml(v.payment_type || "-")}</td>
                            <td>
                                ${v.advance_value != null
                                    ? `${escapeHtml(String(v.advance_value))}%`
                                    + (v.balance_due_days != null
                                        ? `<br><span style="font-size:12px;color:#777;">Balance in ${escapeHtml(String(v.balance_due_days))} days</span>`
                                        : "")
                                    : "-"
                                }
                            </td>
                            <td>${escapeHtml(v.payment_terms_remarks || "-")}</td>
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

        const result = await response.json();

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

        const result = await response.json();

        if (!response.ok || !result.success) {
            list.textContent = result.message || "Failed to load quotations";
            return;
        }

        renderQuotationComparisons(result.inquiries || []);

    } catch {
        list.textContent = "Failed to connect to procurement service";
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
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">PR Number</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.pr_number || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">PR Date</span>
                    <span class="inquiry-card-value">${formatDate(inquiry.pr_date)}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Party Name</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.party_name || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Location</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.location || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Territory</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.territory || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Product Category</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.product_category || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Item Name</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.item_name || "-")}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Make / Model</span>
                    <span class="inquiry-card-value">
                        ${escapeHtml(inquiry.make || "-")} / ${escapeHtml(inquiry.model || "-")}
                    </span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Quantity</span>
                    <span class="inquiry-card-value">
                        ${escapeHtml(String(inquiry.qty ?? 0))} ${escapeHtml(inquiry.unit || "")}
                    </span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Sales Rate</span>
                    <span class="inquiry-card-value">${formatCurrency(inquiry.sales_rate)}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Taxable Value</span>
                    <span class="inquiry-card-value">${formatCurrency(inquiry.taxable_value)}</span>
                </div>
                <div class="inquiry-card-item">
                    <span class="inquiry-card-label">Product Remarks</span>
                    <span class="inquiry-card-value">${escapeHtml(inquiry.product_remarks || "-")}</span>
                </div>
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
                                </tr>
                            </thead>
                            <tbody>
                                ${inquiry.vendors.map(v => {
                                    const isLowest = lowestPrice !== null && Number(v.price_per_unit) === lowestPrice;
                                    return `
                                        <tr>
                                            <td>${escapeHtml(v.vendor_code || "-")}</td>
                                            <td>${escapeHtml(v.vendor_name || "-")}</td>
                                            <td class="${isLowest ? "lowest-price" : ""}">
                                                ${formatCurrency(v.price_per_unit)}
                                            </td>
                                            <td>${formatCurrency(v.total_price)}</td>
                                            <td>${v.advance_value != null ? `${escapeHtml(String(v.advance_value))}%` : "-"}</td>
                                            <td>${v.advance_amount != null ? `₹${formatCurrency(v.advance_amount)}` : "-"}</td>
                                            <td>${formatDate(v.expected_delivery_date)}</td>
                                            <td>${v.balance_due_days != null ? escapeHtml(String(v.balance_due_days)) : "-"}</td>
                                            <td>${escapeHtml(v.remarks || "-")}</td>
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
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">PR Number</span>
                        <span class="inquiry-card-value">${escapeHtml(row.pr_number || "-")}</span>
                    </div>
                    ${row.po_number ? `
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">PO Number</span>
                        <span class="inquiry-card-value">${escapeHtml(row.po_number)}</span>
                    </div>` : ""}
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">PR Date</span>
                        <span class="inquiry-card-value">${formatDate(row.pr_date)}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Party Name</span>
                        <span class="inquiry-card-value">${escapeHtml(row.party_name || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Location</span>
                        <span class="inquiry-card-value">${escapeHtml(row.location || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Territory</span>
                        <span class="inquiry-card-value">${escapeHtml(row.territory || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Product Category</span>
                        <span class="inquiry-card-value">${escapeHtml(row.product_category || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Item Name</span>
                        <span class="inquiry-card-value">${escapeHtml(row.item_name || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Make / Model</span>
                        <span class="inquiry-card-value">${escapeHtml(row.make || "-")} / ${escapeHtml(row.model || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Quantity</span>
                        <span class="inquiry-card-value">${escapeHtml(String(row.qty ?? 0))} ${escapeHtml(row.unit || "")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Sales Rate</span>
                        <span class="inquiry-card-value">${formatCurrency(row.sales_rate)}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Taxable Value</span>
                        <span class="inquiry-card-value">${formatCurrency(row.taxable_value)}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Product Remarks</span>
                        <span class="inquiry-card-value">${escapeHtml(row.product_remarks || "-")}</span>
                    </div>
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

        const result = await response.json();

        if (!response.ok || !result.success) {
            purchaseOrdersList.textContent = result.message || "Failed to load purchase orders";
            return;
        }

        renderPurchaseOrders(result.purchase_orders || []);

    } catch {
        purchaseOrdersList.textContent = "Failed to connect to procurement service";
    }
}

// FIX: escape the label parameter — turnover_year_1/2/3 are free-text vendor
// fields and were previously rendered unescaped, creating a stored XSS vector.
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
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">PR Number</span>
                        <span class="inquiry-card-value">${escapeHtml(po.pr_number || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">PR Date</span>
                        <span class="inquiry-card-value">${formatDate(po.pr_date)}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Party Name</span>
                        <span class="inquiry-card-value">${escapeHtml(po.party_name || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Location</span>
                        <span class="inquiry-card-value">${escapeHtml(po.location || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Territory</span>
                        <span class="inquiry-card-value">${escapeHtml(po.territory || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Product Category</span>
                        <span class="inquiry-card-value">${escapeHtml(po.product_category || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Item Name</span>
                        <span class="inquiry-card-value">${escapeHtml(po.item_name || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Make / Model</span>
                        <span class="inquiry-card-value">${escapeHtml(po.make || "-")} / ${escapeHtml(po.model || "-")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Quantity</span>
                        <span class="inquiry-card-value">${escapeHtml(String(po.qty ?? 0))} ${escapeHtml(po.unit || "")}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Sales Rate</span>
                        <span class="inquiry-card-value">${formatCurrency(po.sales_rate)}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Taxable Value</span>
                        <span class="inquiry-card-value">${formatCurrency(po.taxable_value)}</span>
                    </div>
                    <div class="inquiry-card-item">
                        <span class="inquiry-card-label">Product Remarks</span>
                        <span class="inquiry-card-value">${escapeHtml(po.product_remarks || "-")}</span>
                    </div>
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

                    <span class="inquiry-status status-${(po.status || "").toLowerCase()}">
                        ${escapeHtml(po.status || "-")}
                    </span>
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
        const poId  = vendorBtn.dataset.poId;
        const panel = document.getElementById(`vendorDetailsPanel-${poId}`);
        const isOpen = panel.classList.toggle("open");
        vendorBtn.textContent = isOpen ? "Hide Vendor Details" : "Show Vendor Details";
        return;
    }

    // ── Issue PO ──────────────────────────────────────────────────────────────

    const issuePOBtn = event.target.closest(".issue-po-btn");
    if (issuePOBtn) {
        const poId = issuePOBtn.dataset.poId;
        if (!confirm("Issue this Purchase Order?")) return;

        issuePOBtn.disabled    = true;
        issuePOBtn.textContent = "Issuing...";

        try {
            const response = await apiFetch(`/purchase-orders/${poId}/issue`, { method: "POST" });
            if (!response) return;

            if (!response.ok) {
                const result = await response.json();
                alert(result.message || "Failed to issue PO.");
                return;
            }
            const blob = await response.blob();
            const url  = URL.createObjectURL(blob);
            const a    = document.createElement("a");
            a.href     = url;

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
        } finally {
            // Button is stale after loadPurchaseOrders re-renders; no need to re-enable.
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

// ─── Goods Received ───────────────────────────────────────────────────────────

async function loadGoodsReceived() {
    try {
        goodsReceivedMessage.textContent = "Loading goods received...";

        const response = await apiFetch("/goods-received");
        if (!response) return;

        const result = await response.json();

        if (!response.ok) {
            goodsReceivedMessage.textContent = result.message || "Failed to load goods received";
            return;
        }

        const activeOrders = (result.orders || []).filter(o => o.po_status !== "COMPLETED");

        renderGoodsReceived(activeOrders);
        goodsReceivedMessage.textContent = "";

    } catch (error) {
        console.error(error);
        goodsReceivedMessage.textContent = "Failed to connect to procurement service";
    }
}

function renderGoodsReceived(rows) {
    if (!rows.length) {
        goodsReceivedList.innerHTML = `
            <div class="message">No purchase orders available for goods receipt.</div>
        `;
        return;
    }

    goodsReceivedList.innerHTML = rows.map(po => {
        const ordered   = Number(po.ordered_quantity)   || 0;
        const received  = Number(po.received_quantity)  || 0;
        const remaining = Number(po.remaining_quantity) || 0;
        const progress  = Number(po.progress)           || 0;

        // FIX: "Mark Complete" should compare net received (after returns)
        // against ordered quantity.  The backend already does this calculation
        // and returns received_quantity as the net figure.
        const canComplete = received >= ordered && ordered > 0;

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
                    <button class="primary-button" onclick="openReceiveForm(${po.po_id}, this)">
                        Add Received
                    </button>
                    <button class="secondary-button" onclick="openReturnForm(${po.po_id}, this)">
                        Return
                    </button>
                    ${canComplete ? `
                        <button class="secondary-button" onclick="completePO(${po.po_id}, this)">
                            Mark Complete
                        </button>
                    ` : ""}
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
                <input type="date" id="receipt-date-${poId}" value="${todayISO()}" max="${todayISO()}" required>
            </div>
            <div class="form-group full-width">
                <label>Remarks <span style="font-weight:400;color:#888;">(optional)</span></label>
                <textarea id="receipt-remarks-${poId}" placeholder="Optional remarks"></textarea>
            </div>
        </div>
        <div class="goods-form-actions">
            <button class="primary-button" id="save-receipt-btn-${poId}" onclick="saveGoodsReceipt(${poId})">Save Receipt</button>
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
                <input type="date" id="return-date-${poId}" value="${todayISO()}" max="${todayISO()}" required>
            </div>
            <div class="form-group full-width">
                <label>Reason <span style="color:#c00;">*</span></label>
                <textarea id="return-reason-${poId}" required placeholder="Enter reason for return"></textarea>
            </div>
        </div>
        <div class="goods-form-actions">
            <button class="primary-button" id="save-return-btn-${poId}" onclick="saveGoodsReturn(${poId})">Save Return</button>
            <button class="cancel-vendor-btn" onclick="closeGoodsForm(${poId}, 'return')">Cancel</button>
        </div>
    `;

    form.classList.remove("hidden");
}

// FIX: saveGoodsReceipt and saveGoodsReturn disable their submit button
// immediately to prevent double-posting of stock data.
async function saveGoodsReceipt(poId) {
    const receivedQuantity = Number(document.getElementById(`received-quantity-${poId}`).value);
    const receiptDate      = document.getElementById(`receipt-date-${poId}`).value;
    const remarks          = document.getElementById(`receipt-remarks-${poId}`).value.trim();
    const btn              = document.getElementById(`save-receipt-btn-${poId}`);

    if (!receivedQuantity || receivedQuantity <= 0) { alert("Received quantity must be greater than 0."); return; }
    if (!receiptDate)                               { alert("Receipt date is required."); return; }

    try {
        if (btn) btn.disabled = true;

        const response = await apiFetch("/goods-received", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                po_id:             poId,
                receipt_date:      receiptDate,
                received_quantity: receivedQuantity,
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

async function saveGoodsReturn(poId) {
    const returnQuantity = Number(document.getElementById(`return-quantity-${poId}`).value);
    const returnDate     = document.getElementById(`return-date-${poId}`).value;
    const returnReason   = document.getElementById(`return-reason-${poId}`).value.trim();
    const btn            = document.getElementById(`save-return-btn-${poId}`);

    if (!returnQuantity || returnQuantity <= 0) { alert("Return quantity must be greater than 0."); return; }
    if (!returnDate)                            { alert("Return date is required."); return; }
    if (!returnReason)                          { alert("Return reason is required."); return; }

    try {
        if (btn) btn.disabled = true;

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

        const result = await response.json();

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

// FIX: completePO disables the button immediately to prevent double-submit.
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
// FIX: the click listener for drill-down is attached once here at module level,
// NOT inside loadVendorPerformance/renderVendorPerformance. The original code
// stacked a new listener on every visit to the page, causing the row to
// open-and-immediately-close on the second visit.

const vendorPerformanceList = document.getElementById("vendorPerformanceList");

vendorPerformanceList.addEventListener("click", async event => {
    const btn = event.target.closest(".vp-drill-btn");
    if (!btn) return;

    const vendorId  = btn.dataset.vendorId;
    const detailRow = document.getElementById(`vpDetail-${vendorId}`);
    const content   = document.getElementById(`vpDetailContent-${vendorId}`);
    const isOpen    = detailRow.style.display !== "none";

    // Collapse all others.
    vendorPerformanceList.querySelectorAll(".vp-detail-row").forEach(r => r.style.display = "none");
    vendorPerformanceList.querySelectorAll(".vp-drill-btn").forEach(b => b.textContent = "Details ▾");

    if (isOpen) return;

    btn.textContent        = "Details ▴";
    detailRow.style.display = "";
    content.innerHTML      = "Loading...";

    try {
        const response = await apiFetch(`/vendor-performance/${vendorId}`);
        if (!response) return;

        const result = await response.json();

        if (!response.ok || !result.success) {
            content.innerHTML = escapeHtml(result.message || "Failed to load details");
            return;
        }

        renderVendorPerformanceDetail(content, result);

    } catch {
        content.innerHTML = "Failed to connect to procurement service";
    }
});

async function loadVendorPerformance() {
    const msg = document.getElementById("vendorPerformanceMessage");

    vendorPerformanceList.innerHTML = "";
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

function renderVendorPerformance(vendors) {
    if (!vendors.length) {
        vendorPerformanceList.innerHTML = `<div class="message">No vendors found.</div>`;
        return;
    }

    vendors.sort((a, b) => {
        if (!a.vendor_code && !b.vendor_code) return 0;
        if (!a.vendor_code) return 1;
        if (!b.vendor_code) return -1;
        return a.vendor_code.localeCompare(b.vendor_code, undefined, { numeric: true });
    });

    // FIX: no addEventListener call here — the listener lives at module level above.
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
                    <tr>
                        <td>${escapeHtml(v.vendor_code || "-")}</td>
                        <td>${escapeHtml(v.vendor_name)}</td>
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
                            <button
                                type="button"
                                class="option-button vp-drill-btn"
                                data-vendor-id="${v.vendor_id}"
                                data-vendor-name="${escapeHtml(v.vendor_name)}">
                                Details ▾
                            </button>
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

    vendorPerformanceList.appendChild(table);
}

function renderVendorPerformanceDetail(container, { profile, kpis, orders }) {
    const kpiCard = (label, value) => `
        <div class="inquiry-card-item">
            <span class="inquiry-card-label">${escapeHtml(label)}</span>
            <span class="inquiry-card-value">${escapeHtml(String(value ?? "-"))}</span>
        </div>
    `;

    container.innerHTML = `
        <div class="inquiry-card-info" style="margin-bottom:16px;">
            ${kpiCard("Total Orders",     kpis.total_orders)}
            ${kpiCard("Order Value",      `₹${formatCurrency(kpis.order_value)}`)}
            ${kpiCard("Avg Lead Time",    kpis.avg_lead_time_days !== null ? `${kpis.avg_lead_time_days} days` : "-")}
            ${kpiCard("On-Time Delivery", kpis.on_time_pct !== null ? `${kpis.on_time_pct}% (${kpis.on_time_deliveries}/${kpis.total_deliveries})` : "-")}
            ${kpiCard("Open Orders",      kpis.open_orders)}
        </div>

        <div class="inquiry-card-info" style="margin-bottom:16px;">
            ${kpiCard("GST",           profile.gst_number  || "-")}
            ${kpiCard("PAN",           profile.pan_number  || "-")}
            ${kpiCard("Office",        profile.office_address || "-")}
            ${kpiCard("Sales Contact", profile.sales_team_email || "-")}
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
// FIX: listeners are attached once (pprListenersAttached guard) instead of on
// every call to initPastPriceReference, preventing accumulation across visits.
// FIX: the model list is refreshed on every page open so newly added products appear.
// FIX: loadPPRDetail guards against a blank model string before calling the API.

let pprModelList = [];
let pprListenersAttached = false;

async function loadPPRModels() {
    try {
        const response = await apiFetch("/past-price-reference/models");
        if (!response) return;

        const result = await response.json();
        if (result.success) pprModelList = result.models || [];

    } catch {
        // Silently fail — the search box will simply yield no suggestions.
    }
}

function initPastPriceReference() {
    const input    = document.getElementById("pprSearchInput");
    const dropdown = document.getElementById("pprDropdown");

    // Refresh the model list on every page open so new products appear.
    loadPPRModels();

    // Reset UI state.
    input.value = "";
    document.getElementById("pprResults").innerHTML   = "";
    document.getElementById("pprMessage").textContent = "";
    document.getElementById("pprSelectedLabel").classList.add("hidden");
    dropdown.classList.add("hidden");

    // FIX: only attach listeners once.
    if (!pprListenersAttached) {
        input.addEventListener("input", onPPRInput);
        document.addEventListener("click", onPPROutsideClick);
        pprListenersAttached = true;
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
        (m.item_name || "").toLowerCase().includes(query) ||
        (m.model     || "").toLowerCase().includes(query) ||
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
            data-model="${escapeHtml(m.model || "")}"
            data-item-name="${escapeHtml(m.item_name || "")}">
            <span class="ppr-dropdown-model">${escapeHtml(m.item_name || "-")}</span>
            <span class="ppr-dropdown-meta">
                ${m.model ? `Model: ${escapeHtml(m.model)}` : ""}
                ${m.make ? ` · ${escapeHtml(m.make)}` : ""}
                ${m.product_category ? ` · ${escapeHtml(m.product_category)}` : ""}
            </span>
        </div>
    `).join("");

    dropdown.classList.remove("hidden");

    dropdown.querySelectorAll(".ppr-dropdown-item").forEach(item => {
        item.addEventListener("click", () => {
            const model    = item.dataset.model;
            const itemName = item.dataset.itemName;

            document.getElementById("pprSearchInput").value = itemName;
            dropdown.classList.add("hidden");

            // FIX: only call the API when model is non-empty.
            if (model) {
                loadPPRDetail(model);
            } else {
                document.getElementById("pprMessage").textContent = "This item has no model number on record.";
            }
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

    results.innerHTML   = "";
    msg.textContent     = "Loading...";
    label.classList.add("hidden");

    try {
        const response = await apiFetch(`/past-price-reference?model=${encodeURIComponent(model)}`);
        if (!response) return;

        const result = await response.json();

        if (!response.ok || !result.success) {
            msg.textContent = result.message || "Failed to fetch price history";
            return;
        }

        msg.textContent = "";

        if (!result.purchases?.length) {
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
                ${v.vendor_code ? `<span class="ppr-vendor-code">${escapeHtml(v.vendor_code)}</span>` : ""}
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

// ─── Window exports (called from inline onclick attributes in goods-received) ──

window.completePO       = completePO;
window.openReceiveForm  = openReceiveForm;
window.openReturnForm   = openReturnForm;
window.saveGoodsReceipt = saveGoodsReceipt;
window.saveGoodsReturn  = saveGoodsReturn;
window.closeGoodsForm   = closeGoodsForm;
window.viewGoodsHistory = viewGoodsHistory;