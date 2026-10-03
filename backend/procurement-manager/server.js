require("dotenv").config();
const express = require("express");
const mysql = require("mysql2/promise");
const path = require("path");
const fs = require("fs");
const multer = require("multer");
const { readSheet } = require("read-excel-file/node");
const PDFDocument = require("pdfkit");

const app = express();
const PORT = process.env.PORT;

/* ==========================================================================
   TECHNICAL LOGGER (console only - separate from business Report Logs)
   ========================================================================== */

function log(message) {
    console.log(`[${new Date().toLocaleString("en-IN", { timeZone: "Asia/Kolkata" })}] ${message}`);
}

/* ==========================================================================
   DATABASE CONNECTIONS
   ========================================================================== */

const dbConfig = {
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME
};

// Main procurement database pool
const db = mysql.createPool({
    ...dbConfig,
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0
});

// Logging database pool (same server and credentials, database = DB2_NAME) - holds report_logs
const logDb = mysql.createPool({
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB2_NAME,
    waitForConnections: true,
    connectionLimit: 5,
    queueLimit: 0
});

const authServiceUrl = process.env.AUTH_SERVICE_URL;

const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 10 * 1024 * 1024
    }
});

const vendorFolder = path.join(__dirname, "..", "vendor");
const poFolder      = path.join(__dirname, "..", "purchase-orders");
const piFolder      = path.join(__dirname, "..", "proforma-invoice");
const headerImgPath = path.join(__dirname, "..", "PO header.png");

app.use(express.json());
app.use(express.urlencoded({ extended: false }));

/* ==========================================================================
   AUTHENTICATION
   ========================================================================== */

// Middleware: allows only logged-in PROCUREMENT_MANAGER users, otherwise redirects to the auth service
async function verifyManager(req, res, next) {
    try {
        if (!req.headers.cookie) {
            return res.status(401).json({ success: false, message: "Unauthenticated" });
        }
        const response = await fetch(`${authServiceUrl}/verify`, {
            method: "GET",
            headers: { Cookie: req.headers.cookie }
        });
        if (!response.ok) {
            return res.status(401).json({ success: false, message: "Unauthenticated" });
        }
        const data = await response.json();
        if (!data.authenticated || data.role !== "PROCUREMENT_MANAGER") {
            return res.status(401).json({ success: false, message: "Unauthenticated" });
        }
        req.user = { user_id: data.user_id, username: data.username, role: data.role };
        next();
    } catch (error) {
        console.error("Authentication verification failed:", error);
        return res.status(401).json({ success: false, message: "Unauthenticated" });
    }
}

/* ==========================================================================
   REPORT LOG HELPERS (business activity history stored in DB2_NAME.report_logs)
   ========================================================================== */

const actorCache = new Map();
const ACTOR_CACHE_TTL_MS = 5 * 60 * 1000;

// Resolves the username performing the current request (from req.user, or from the auth service using the session cookie)
async function getActor(req) {
    if (req.user?.username) return req.user.username;
    const cookie = req.headers.cookie;
    if (!cookie) return "UNKNOWN";
    const cached = actorCache.get(cookie);
    if (cached && cached.expires > Date.now()) return cached.username;
    try {
        const response = await fetch(`${authServiceUrl}/verify`, {method: "GET",headers: { Cookie: cookie }});
        if (response.ok) {
            const data = await response.json();
            if (data.authenticated && data.username) {
                if (actorCache.size > 500) {
                    for (const [key, value] of actorCache) {
                        if (value.expires <= Date.now()) actorCache.delete(key);
                    }
                }
                actorCache.set(cookie, { username: data.username, expires: Date.now() + ACTOR_CACHE_TTL_MS });
                return data.username;
            }
        }
    } catch (error) {
        log(`ERROR resolving username for report log: ${error.message}`);
    }
    return "UNKNOWN";
}

// Writes one business-activity entry to report_logs. Never throws, so a logging failure can never break a business action.
async function writeReportLog(req, action, report) {
    try {
        const username = await getActor(req);
        await logDb.execute(
            `INSERT INTO report_logs (username, action, report) VALUES (?, ?, ?)`,
            [String(username).slice(0, 100), String(action).slice(0, 100), String(report)]
        );
    } catch (error) {
        log(`ERROR writing report log (${action}): ${error.message}`);
    }
}

// Formats a number as an Indian-style currency amount for report text
function money(value) {
    return Number(value || 0).toLocaleString("en-IN", {minimumFractionDigits: 2,maximumFractionDigits: 2});
}

// Returns "-" for empty values so report text never contains blanks or "null"
function orDash(value) {
    if (value === null || value === undefined) return "-";
    const text = String(value).trim();
    return text === "" ? "-" : text;
}

// Formats a Date / date string as YYYY-MM-DD for report text
function dateText(value) {
    if (!value) return "-";
    if (value instanceof Date) return value.toISOString().split("T")[0];
    return String(value).split("T")[0];
}

// Builds the human-readable description of a purchase request used in several report entries
function describePr(pr) {
    return `PR ${pr.pr_number} dated ${dateText(pr.pr_date)} for party "${orDash(pr.party_name)}" ` +
        `(location: ${orDash(pr.location)}, territory: ${orDash(pr.territory)}). ` +
        `Product: ${orDash(pr.item_name)} [category: ${orDash(pr.product_category)}], make: ${orDash(pr.make)}, model: ${orDash(pr.model)}. ` +
        `Quantity: ${Number(pr.qty || 0)} ${orDash(pr.unit)} at sales rate ${money(pr.sales_rate)}, taxable value ${money(pr.taxable_value)}. ` +
        `Remarks: ${orDash(pr.product_remarks)}.`;
}

// Builds the human-readable payment-terms text used in quotation / PO report entries
function describePayment(q) {
    let text = `Payment type: ${orDash(q.payment_type)}`;
    if (q.advance_type) text += `, advance: ${q.advance_value} (${q.advance_type})`;
    if (q.balance_due_days) text += `, balance due in ${q.balance_due_days} days`;
    if (q.payment_terms_remarks) text += `, terms: ${q.payment_terms_remarks}`;
    return text;
}

/* ==========================================================================
   GENERAL HELPERS
   ========================================================================== */

// Trims strings and converts empty / "-" values to null
function cleanValue(value) {
    if (value === null || value === undefined) return null;
    if (typeof value === "string") {
        const cleaned = value.trim();
        if (cleaned === "" || cleaned === "-") return null;
        return cleaned;
    }
    return value;
}

// Converts a value to a finite number, defaulting to 0
function cleanNumber(value) {
    if (value === null || value === undefined || value === "") return 0;
    if (typeof value === "string" && value.trim() === "-") return 0;
    const number = Number(value);
    return Number.isFinite(number) ? number : 0;
}

// Returns the Indian financial year (April-March) as "YY-YY" for the given date
function getFinancialYear(date = new Date()) {
    const month = date.getMonth() + 1;
    const year = date.getFullYear();
    if (month >= 4)return `${String(year).slice(-2)}-${String(year + 1).slice(-2)}`;
    return `${String(year - 1).slice(-2)}-${String(year).slice(-2)}`;
}

// Generates the next sequential PR number for the current financial year inside the caller's transaction
async function generatePrNumber(connection) {
    const financialYear = getFinancialYear();
    await connection.execute(
        `INSERT INTO pr_sequences (financial_year, last_number)
         VALUES (?, 0)
         ON DUPLICATE KEY UPDATE financial_year = financial_year`,
        [financialYear]
    );
    await connection.execute(
        `UPDATE pr_sequences
         SET last_number = LAST_INSERT_ID(last_number + 1)
         WHERE financial_year = ?`,
        [financialYear]
    );
    const [rows] = await connection.execute(`SELECT LAST_INSERT_ID() AS sequence_number`);
    const sequenceNumber = rows[0].sequence_number;
    return `NEE/${financialYear}/PR/${String(sequenceNumber).padStart(4, "0")}`;
}

// Normalizes an Excel header into a snake_case key
function normalizeHeader(value) {
    return String(value || "")
        .trim()
        .toLowerCase()
        .replace(/\s+/g, "_")
        .replace(/[\/()-]+/g, "_")
        .replace(/_+/g, "_")
        .replace(/^_|_$/g, "");
}

// Converts Date objects and dd-mm-yyyy style text into YYYY-MM-DD
function formatDate(value) {
    if (!value) return null;
    if (value instanceof Date) return value.toISOString().split("T")[0];
    const text = String(value).trim();
    if (text === "" || text === "-") return null;
    const match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
    if (match)return `${match[3]}-${match[2].padStart(2, "0")}-${match[1].padStart(2, "0")}`;
    return text;
}

// Makes a company name safe to use as a folder name
function sanitizeFolderName(name) {
    return String(name)
        .trim()
        .replace(/[<>:"/\\|?*]/g, "_")
        .replace(/\s+/g, " ");
}

// Returns the lower-case file extension of an upload, defaulting to .pdf
function getFileExtension(filename) {
    const extension = path.extname(filename || "").toLowerCase();
    return extension || ".pdf";
}

// Base file name for a PO's Proforma Invoice: the PO number with "PO" replaced
// by "PI", sanitized for the filesystem. Shared by upload/status/download so
// all three always agree on where a PO's PI lives.
function getPiBaseName(poNumber) {
    const piNumber = String(poNumber || "PI").replace(/PO/gi, "PI");
    return piNumber.replace(/[\/\\:*?"<>|]/g, "_");
}

// Full saved file name (base name + the uploaded file's own extension).
function getPiFileName(poNumber, originalName) {
    return `${getPiBaseName(poNumber)}${getFileExtension(originalName)}`;
}

// Finds the previously uploaded PI for a PO, regardless of its extension.
// Returns the full path, or null if nothing has been uploaded yet.
function findPiFilePath(poNumber) {
    const baseName = getPiBaseName(poNumber);
    if (!fs.existsSync(piFolder)) return null;
    const match = fs.readdirSync(piFolder).find(fileName => path.parse(fileName).name === baseName);
    return match ? path.join(piFolder, match) : null;
}

// Derives the PO number from the PR number by swapping /PR/ for /PO/
async function generatePoNumber(connection, prNumber) {
    return prNumber.replace("/PR/", "/PO/");
}

/* ==========================================================================
   VENDOR HELPERS
   ========================================================================== */

// Writes the uploaded vendor documents into vendor/<company name>/ and returns their locations
function saveVendorDocuments(files, companyName) {
    const folderName = sanitizeFolderName(companyName);
    const companyFolder = path.join(vendorFolder, folderName);
    fs.mkdirSync(companyFolder, { recursive: true });
    const documentLocations = {};
    const createdFiles = [];
    const documentNames = {
        gst_document: "GST",
        pan_document: "PAN",
        msme_document: "MSME",
        itr_last_year_document: "ITR_Last_Year",
        itr_second_last_year_document: "ITR_Second_Last_Year",
        itr_third_last_year_document: "ITR_Third_Last_Year"
    };
    for (const [fieldName, fileList] of Object.entries(files || {})) {
        const file = fileList?.[0];
        if (!file) continue;
        const fileName = `${documentNames[fieldName]}${getFileExtension(file.originalname)}`;
        const filePath = path.join(companyFolder, fileName);
        fs.writeFileSync(filePath, file.buffer);
        createdFiles.push(filePath);
        documentLocations[fieldName] = path.relative(__dirname, filePath).replace(/\\/g, "/");
    }
    return { documentLocations, createdFiles, companyFolder };
}

// Deletes files written during a failed vendor save and removes the company folder if it is empty
function cleanupFiles(files, companyFolder) {
    for (const filePath of files || []) {
        try {
            if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
        } catch (error) {
            log(`ERROR deleting file: ${error.message}`);
        }
    }
    try {
        if (companyFolder && fs.existsSync(companyFolder) && fs.readdirSync(companyFolder).length === 0) {
            fs.rmdirSync(companyFolder);
        }
    } catch (error) {
        log(`ERROR removing vendor folder: ${error.message}`);
    }
}

// Multer middleware accepting the six vendor document uploads
function getVendorDocuments() {
    return upload.fields([
        { name: "gst_document", maxCount: 1 },
        { name: "pan_document", maxCount: 1 },
        { name: "msme_document", maxCount: 1 },
        { name: "itr_last_year_document", maxCount: 1 },
        { name: "itr_second_last_year_document", maxCount: 1 },
        { name: "itr_third_last_year_document", maxCount: 1 }
    ]);
}

// Parses the vendor_data JSON sent alongside the vendor document uploads
function parseVendorData(req) {
    if (!req.body.vendor_data) return null;
    try {
        return typeof req.body.vendor_data === "string" ? JSON.parse(req.body.vendor_data) : req.body.vendor_data;
    } catch {
        return null;
    }
}

// Database columns of vendor_oem_masters that hold vendor details
const VENDOR_FIELDS = [
    "registration_date",
    "vendor_name",
    "office_address",
    "office_contact_name",
    "office_contact_number",
    "factory_address",
    "factory_contact_name",
    "factory_contact_number",
    "warehouse_address",
    "warehouse_contact_name",
    "warehouse_contact_number",
    "workshop_address",
    "workshop_contact_name",
    "workshop_contact_number",
    "legal_entity",
    "commercial_role",
    "year_of_incorporation",
    "director_or_ceo_or_management_name",
    "director_or_ceo_or_management_designation",
    "director_or_ceo_or_management_mobile_no",
    "director_or_ceo_or_management_email",
    "director_or_ceo_or_management_web_address",
    "sales_team_name",
    "sales_team_contact",
    "sales_team_email",
    "accounts_team_name",
    "accounts_team_contact",
    "accounts_team_email",
    "gst_number",
    "pan_number",
    "msme_number",
    "bank_name",
    "bank_account_no",
    "bank_branch",
    "bank_account_type",
    "bank_ifsc",
    "branch_office_1_address",
    "branch_office_1_contact_name",
    "branch_office_1_contact_number",
    "branch_office_2_address",
    "branch_office_2_contact_name",
    "branch_office_2_contact_number",
    "branch_office_3_address",
    "branch_office_3_contact_name",
    "branch_office_3_contact_number",
    "turnover_year_1",
    "turnover_value_1",
    "turnover_year_2",
    "turnover_value_2",
    "turnover_year_3",
    "turnover_value_3",
    "recommended_by",
    "approved_by"
];

// Database columns of vendor_oem_masters that hold uploaded document paths
const VENDOR_DOCUMENT_FIELDS = [
    "gst_document",
    "pan_document",
    "msme_document",
    "itr_last_year_document",
    "itr_second_last_year_document",
    "itr_third_last_year_document"
];

// Checks that all mandatory vendor fields and documents are present
function validateVendorData(data, files) {
    const requiredFields = [
        "registration_date",
        "vendor_name",
        "office_address",
        "office_contact_name",
        "office_contact_number",
        "legal_entity",
        "commercial_role",
        "year_of_incorporation",
        "director_or_ceo_or_management_name",
        "director_or_ceo_or_management_designation",
        "director_or_ceo_or_management_mobile_no",
        "director_or_ceo_or_management_email",
        "sales_team_name",
        "sales_team_contact",
        "sales_team_email",
        "accounts_team_name",
        "accounts_team_contact",
        "accounts_team_email",
        "gst_number",
        "pan_number",
        "bank_name",
        "bank_account_no",
        "bank_branch",
        "bank_account_type",
        "bank_ifsc",
        "branch_office_1_address",
        "turnover_year_1",
        "turnover_value_1",
        "recommended_by",
        "approved_by"
    ];
    const missingFields = requiredFields.filter(field => !cleanValue(data?.[field]));
    const requiredDocuments = [
        "gst_document",
        "pan_document",
        "itr_last_year_document"
    ];
    const missingDocuments = requiredDocuments.filter(field => !files?.[field]?.[0]);
    return { missingFields, missingDocuments };
}

// Cleans the raw vendor data into the exact set of columns that get stored
function prepareVendorData(data) {
    const vendor = {};
    for (const field of VENDOR_FIELDS)vendor[field] = cleanValue(data?.[field]);
    vendor.registration_date = formatDate(data?.registration_date);
    return vendor;
}

// Inserts a vendor (rejecting duplicate GST numbers), stores documents and assigns the vendor code
async function saveVendor(connection, data, files) {
    const vendor = prepareVendorData(data);
    const [existingVendor] = await connection.execute(
        `SELECT vendor_id, vendor_name FROM vendor_oem_masters WHERE gst_number = ? LIMIT 1`,
        [vendor.gst_number]
    );
    if (existingVendor.length > 0)throw new Error(`A vendor named "${existingVendor[0].vendor_name}" already exists with this GST number`);
    const { documentLocations, createdFiles, companyFolder } = saveVendorDocuments(files, vendor.vendor_name);
    try {
        const columns = [...VENDOR_FIELDS, ...VENDOR_DOCUMENT_FIELDS];
        const values = [
            ...VENDOR_FIELDS.map(field => vendor[field] ?? null),
            ...VENDOR_DOCUMENT_FIELDS.map(field => documentLocations[field] || null)
        ];
        const placeholders = columns.map(() => "?").join(", ");
        const [result] = await connection.execute(
            `INSERT INTO vendor_oem_masters (${columns.join(", ")})
             VALUES (${placeholders})`,
            values
        );
        const vendorId = result.insertId;
        const financialYear = getFinancialYear().replace("-", "");
        const vendorCode = `NEE${financialYear}V${String(vendorId).padStart(3, "0")}`;
        await connection.execute(
            `UPDATE vendor_oem_masters SET vendor_code = ? WHERE vendor_id = ?`,
            [vendorCode, vendorId]
        );
        return { vendorId, vendorCode, createdFiles, companyFolder, documentLocations };
    } catch (error) {
        cleanupFiles(createdFiles, companyFolder);
        throw error;
    }
}

// Builds the report text for a newly registered vendor
function describeVendor(data, result, source) {
    const vendor = prepareVendorData(data);
    const documents = Object.keys(result.documentLocations || {}).join(", ");
    return `Vendor "${orDash(vendor.vendor_name)}" registered ${source} with vendor code ${result.vendorCode} (vendor ID ${result.vendorId}). ` +
        `Registration date: ${orDash(vendor.registration_date)}. Legal entity: ${orDash(vendor.legal_entity)}, commercial role: ${orDash(vendor.commercial_role)}, ` +
        `year of incorporation: ${orDash(vendor.year_of_incorporation)}. ` +
        `GST: ${orDash(vendor.gst_number)}, PAN: ${orDash(vendor.pan_number)}, MSME: ${orDash(vendor.msme_number)}. ` +
        `Office: ${orDash(vendor.office_address)} (contact: ${orDash(vendor.office_contact_name)}, ${orDash(vendor.office_contact_number)}). ` +
        `Management contact: ${orDash(vendor.director_or_ceo_or_management_name)} (${orDash(vendor.director_or_ceo_or_management_designation)}), ` +
        `${orDash(vendor.director_or_ceo_or_management_mobile_no)}, ${orDash(vendor.director_or_ceo_or_management_email)}. ` +
        `Sales contact: ${orDash(vendor.sales_team_name)}, ${orDash(vendor.sales_team_contact)}. ` +
        `Bank: ${orDash(vendor.bank_name)}, ${orDash(vendor.bank_branch)}, IFSC ${orDash(vendor.bank_ifsc)}. ` +
        `Recommended by: ${orDash(vendor.recommended_by)}, approved by: ${orDash(vendor.approved_by)}. ` +
        `Documents uploaded: ${orDash(documents)}.`;
}

const formStr = v => {
    const c = cleanValue(v);
    return c === null ? null : String(c).trim();
};

const formKey = v => (formStr(v) || "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();

// Maps form labels to database fields for each labelled section of the Vendor Registration Form
const VENDOR_FORM_SECTIONS = {
    director: {
        "name": "director_or_ceo_or_management_name",
        "designation": "director_or_ceo_or_management_designation",
        "mobile no": "director_or_ceo_or_management_mobile_no",
        "e mail": "director_or_ceo_or_management_email",
        "web address": "director_or_ceo_or_management_web_address"
    },
    sales: {
        "name": "sales_team_name",
        "contact detail": "sales_team_contact",
        "e mail": "sales_team_email"
    },
    accounts: {
        "name": "accounts_team_name",
        "contact detail": "accounts_team_contact",
        "e mail": "accounts_team_email"
    },
    company: {
        "gst no": "gst_number",
        "pan no": "pan_number",
        "bank name": "bank_name",
        "account no": "bank_account_no",
        "branch details": "bank_branch",
        "type of account": "bank_account_type",
        "ifsc rtgs code": "bank_ifsc"
    }
};

// Maps section heading text in the Vendor Registration Form to a section key
const SECTION_HEADERS = {
    "director ceo management team": "director",
    "sales team": "sales",
    "account team": "accounts",
    "company s detail": "company",
    "client details": "client",
    "turnover of last three years": "turnover",
    "documents to be submitted": "documents",
    "for office use only": "office_use"
};

// Detects whether an uploaded sheet is the formatted Vendor Registration Form
function isVendorForm(rows) {
    return rows.some(row => formKey(row[0]) === "vendor registration form");
}

// Parses the formatted Vendor Registration Form sheet into vendor fields
function parseVendorForm(rows) {
    const vendor = Object.fromEntries(VENDOR_FIELDS.map(field => [field, null]));
    const set = (field, value) => {
        const text = formStr(value);
        if (text) vendor[field] = text;
    };
    const locationBySide = {};
    let section = null;
    let branchNo = null;
    let turnoverNo = 0;
    for (const row of rows) {
        const a = formKey(row[0]);
        const b = formKey(row[1]);
        if (SECTION_HEADERS[a]) {
            section = SECTION_HEADERS[a];
            continue;
        }
        if (a.startsWith("details of branch office")) {
            section = "branch";
            continue;
        }
        if (a === "status") {
            set("legal_entity", row[1]);
            section = "status";
            continue;
        }
        if (section === "status" && !a && b) {
            set("commercial_role", row[1]);
            section = null;
            continue;
        }
        if (a === "year of incorporation") {
            set("year_of_incorporation", row[1]);
            continue;
        }
        if (section === null) {
            if (a === "name") set("vendor_name", row[1]);
            if (formKey(row[5]) === "date" && row[6])vendor.registration_date = formatDate(row[6]);
            for (const [labelCol, subCol, valueCol] of [[1, 2, 3], [4, 5, 6]]) {
                const loc = formKey(row[labelCol]);
                if (["office", "factory", "warehouse", "workshop"].includes(loc))locationBySide[subCol] = loc;
                const prefix = locationBySide[subCol];
                if (!prefix) continue;
                const sub = formKey(row[subCol]);
                if (sub === "address") set(`${prefix}_address`, row[valueCol]);
                if (sub === "contact name") set(`${prefix}_contact_name`, row[valueCol]);
                if (sub === "contact number") set(`${prefix}_contact_number`, row[valueCol]);
            }
            continue;
        }
        if (VENDOR_FORM_SECTIONS[section]) {
            if (section === "company" && a.startsWith("micro small medium")) {
                set("msme_number", row[1]);
                continue;
            }
            const field = VENDOR_FORM_SECTIONS[section][a];
            if (field) set(field, row[1]);
            continue;
        }
        if (section === "branch") {
            if (/^[1-3]$/.test(a)) branchNo = Number(a);
            if (!branchNo) continue;
            if (b === "address") set(`branch_office_${branchNo}_address`, row[2]);
            if (b === "contact name") set(`branch_office_${branchNo}_contact_name`, row[2]);
            if (b === "contact number") set(`branch_office_${branchNo}_contact_number`, row[2]);
            continue;
        }
        if (section === "turnover") {
            const year = formStr(row[0]);
            if (a === "year" || !year || turnoverNo >= 3) continue;
            turnoverNo++;
            set(`turnover_year_${turnoverNo}`, year);
            set(`turnover_value_${turnoverNo}`, row[1]);
            continue;
        }
        if (section === "office_use") {
            if (a.startsWith("recommended by")) set("recommended_by", row[1]);
            if (a.startsWith("approved by")) set("approved_by", row[1]);
        }
    }
    return vendor;
}

// Parses a one-row tabular vendor sheet (header row + one vendor row) into vendor fields
function parseVendorTable(excelRows) {
    const dataRows = excelRows.slice(1).filter(row => row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ""));
    if (dataRows.length !== 1) {
        const error = new Error("Excel file must contain exactly one vendor");
        error.status = 400;
        throw error;
    }
    const headers = excelRows[0].map(normalizeHeader);
    const columnIndex = {};
    headers.forEach((header, index) => {
        if (header) columnIndex[header] = index;
    });
    const row = dataRows[0];
    const getCell = field => {
        if (columnIndex[field] === undefined) return null;
        return cleanValue(row[columnIndex[field]]);
    };
    const vendor = {};
    for (const field of VENDOR_FIELDS) {
        const value = getCell(field);
        vendor[field] = typeof value === "number" ? String(value) : value;
    }
    vendor.registration_date = formatDate(getCell("registration_date"));
    return vendor;
}

/* ==========================================================================
   PURCHASE ORDER PDF GENERATION
   ========================================================================== */

// Renders the purchase order onto the company letterhead layout and saves it in the purchase-orders folder
function generatePoPdf(po) {
    return new Promise((resolve, reject) => {
        try {
            fs.mkdirSync(poFolder, { recursive: true });
            const safeFileName = String(po.po_number || "PO").replace(/[\/\\:*?"<>|]/g, "_") + ".pdf";
            const outputPath = path.join(poFolder, safeFileName);
            const doc = new PDFDocument({ size: [612, 792], margin: 0 });
            const stream = fs.createWriteStream(outputPath);
            doc.pipe(stream);
            const fontDir = path.join(__dirname,".." ,"fonts");
            doc.registerFont("Arial", path.join(fontDir, "ARIAL.TTF"));
            doc.registerFont("Arial-Bold", path.join(fontDir, "ARIALBD.TTF"));
            doc.registerFont("Calibri", path.join(fontDir, "CALIBRI.TTF"));
            doc.registerFont("Calibri-Bold", path.join(fontDir, "CALIBRIB.TTF"));
            function fmtDate(val) {
                if (!val) return "";
                const d = new Date(val);
                if (isNaN(d.getTime())) return String(val);
                return `${String(d.getDate()).padStart(2, "0")}.${String(d.getMonth() + 1).padStart(2, "0")}.${d.getFullYear()}`;
            }
            function fmtCur(val) {
                return Number(val || 0).toLocaleString("en-IN", {minimumFractionDigits: 2,maximumFractionDigits: 2});
            }
            function line(x1, y1, x2, y2) {
                doc.moveTo(x1, y1).lineTo(x2, y2).stroke();
            }
            function rect(x, y, w, h) {
                doc.rect(x, y, w, h).stroke();
            }
            doc.lineWidth(1);
            // Company letterhead image
            if (fs.existsSync(headerImgPath)) {
                doc.image(headerImgPath, 50.4, 54, {
                    width: 504.8,
                    height: 79
                });
            }
            const ML = 85;
            const MR = 524;
            const MW = MR - ML;
            // Title box
            const titleY = 192.2;
            const titleH = 21.9;
            rect(ML, titleY, MW, titleH);
            doc.font("Arial-Bold")
                .fontSize(15.24)
                .text("PURCHASE ORDER", ML, titleY + 1.5, {
                    width: MW,
                    align: "center",
                    lineBreak: false
                });
            // Vendor block (left) and PO reference block (right)
            const vendorY = 214.1;
            const vendorH = 84.15;
            rect(ML, vendorY, MW, vendorH);
            const vendorSplit = 316;
            const labelSplit = 350;
            line(vendorSplit, vendorY, vendorSplit, vendorY + vendorH);
            line(labelSplit, vendorY, labelSplit, vendorY + vendorH);
            line(labelSplit, 225.2, MR, 225.2);
            line(labelSplit, 237.15, MR, 237.15);
            line(labelSplit, 249.05, MR, 249.05);
            line(labelSplit, 260.9, MR, 260.9);
            line(labelSplit, 272.7, MR, 272.7);
            doc.font("Arial-Bold").fontSize(9).text("TO,", ML + 1.5, vendorY + 0.3, {lineBreak: false});
            const vendorName = String(po.vendor_name || "").trim();
            const vendorAddress = String(po.vendor_address || "").trim();
            doc.font("Arial").fontSize(8.5);
            let vendorTextY = vendorY + 12;
            if (vendorName) {
                doc.text(vendorName, ML + 1.5, vendorTextY, {width: vendorSplit - ML - 5,lineBreak: false});
                vendorTextY += 11;
            }
            if (vendorAddress) {
                const addressWidth = vendorSplit - ML - 5;
                const addressHeight = doc.heightOfString(vendorAddress, {width: addressWidth,lineGap: 0});
                doc.text(vendorAddress, ML + 1.5, vendorTextY, {width: addressWidth,lineGap: 0});
                vendorTextY += addressHeight + 3;
            }
            if (po.vendor_gst)doc.font("Arial-Bold").fontSize(8.5).text(`GST: ${po.vendor_gst}`, ML + 1.5, vendorTextY, {width: vendorSplit - ML - 5,lineBreak: false});
            const labelX = labelSplit + 1.5;
            const valueX = 430;
            doc.font("Arial-Bold").fontSize(9);
            doc.text("REF.P.O.NO.", labelX, vendorY + 0.3, {lineBreak: false});
            doc.text("DATE", labelX, 226.2, {lineBreak: false});
            doc.text("GST NO", labelX, 250, {lineBreak: false});
            doc.font("Arial-Bold").fontSize(7.56).text(po.po_number || "", valueX, vendorY + 1, {width: MR - valueX - 3,lineBreak: false});
            doc.font("Arial").fontSize(9).text(fmtDate(po.po_date), valueX, 226.2, {lineBreak: false});
            doc.text(process.env.GST_NUMBER || "", valueX, 250, {lineBreak: false});
            // Attention line
            const attnY = 297;
            const attnH = 21.3;
            rect(ML, attnY, MW, attnH);
            doc.font("Arial-Bold").fontSize(9).text(`ATTN. : ${po.vendor_name || ""}`, ML, attnY + 4, {width: MW,align: "center",lineBreak: false});
            // Item table frame and column headers
            const tableHeaderY = 319.6;
            const headerH = 10.8;
            const x0 = 85.6;
            const x1 = 147.6;
            const x2 = 316.3;
            const x3 = 350.3;
            const x4 = 428.4;
            const x5 = 524;
            const totalY = 509.7;
            const totalH = 11.5;
            const tableBottom = totalY + totalH;
            rect(x0, tableHeaderY, x5 - x0, tableBottom - tableHeaderY);
            line(x1, tableHeaderY, x1, tableBottom);
            line(x2, tableHeaderY, x2, tableBottom);
            line(x3, tableHeaderY, x3, tableBottom);
            line(x4, tableHeaderY, x4, tableBottom);
            line(x0, tableHeaderY + headerH, x5, tableHeaderY + headerH);
            doc.font("Calibri-Bold").fontSize(7.56);
            doc.text("SR. NO.", x0, tableHeaderY + 1.3, {width: x1 - x0,align: "center",lineBreak: false});
            doc.text("Model Number", x1, tableHeaderY + 1.3, {width: x2 - x1,align: "center",lineBreak: false});
            doc.text("Qty", x2, tableHeaderY + 1.3, {width: x3 - x2,align: "center",lineBreak: false});
            doc.text("Unit Rate", x3, tableHeaderY + 1.3, {width: x4 - x3,align: "center",lineBreak: false});
            doc.text("Total", x4, tableHeaderY + 1.3, {width: x5 - x4,align: "center",lineBreak: false});
            // Item row
            const itemTop = tableHeaderY + headerH + 1;
            const itemDesc = [
                po.item_name,
                po.make,
                po.model
            ].filter(Boolean).join(" / ");
            doc.font("Calibri").fontSize(7.56);
            doc.text("1", x0, itemTop + 2, {width: x1 - x0,align: "center",lineBreak: false});
            doc.text(itemDesc, x1 + 3, itemTop + 2, {width: x2 - x1 - 6,align: "left",lineBreak: false});
            doc.text(Number(po.qty || 0).toFixed(2), x2, itemTop + 2, {width: x3 - x2,align: "center",lineBreak: false});
            doc.text(fmtCur(po.price_per_unit), x3, itemTop + 2, {width: x4 - x3,align: "center",lineBreak: false});
            doc.text(fmtCur(po.total_price), x4, itemTop + 2, {width: x5 - x4,align: "center",lineBreak: false});
            line(x4, 340.6, x5, 340.6);
            // Table total row
            line(x0, totalY, x5, totalY);
            doc.font("Calibri").fontSize(7.56).text("TOTAL:", x0, totalY + 1.8, {width: x4 - x0,align: "center",lineBreak: false});
            doc.text(fmtCur(po.total_price), x4, totalY + 1.8, {width: x5 - x4,align: "center",lineBreak: false});
            // Footer frame: terms & conditions (left) and signature block (right)
            const footerY = tableBottom;
            const footerBottom = 618.9;
            rect(x0, footerY, x5 - x0, footerBottom - footerY);
            const footerSplit = x3;
            line(footerSplit, footerY, footerSplit, footerBottom);
            line(x0, 531.4, footerSplit, 531.4);
            line(x1, 531.4, x1, 607.5);
            line(x2, 531.4, x2, 607.5);
            line(x0, 607.5, x5, 607.5);
            doc.font("Calibri").fontSize(7.56).text("TERMS & CONDITIONS :", x0, footerY + 1, {width: footerSplit - x0,align: "center",lineBreak: false});
            doc.font("Calibri").fontSize(7.56);
            doc.text("1", x0, 533, {width: x1 - x0,align: "center",lineBreak: false});
            doc.text("2", x0, 543.6, {width: x1 - x0,align: "center",lineBreak: false});
            if (po.payment_terms_remarks)doc.text("3", x0, 554.2, {width: x1 - x0,align: "center",lineBreak: false});
            doc.text("DELIVERY : AT OUR OFFICE.", x1 + 3, 533, {width: x2 - x1 - 6,lineBreak: false});
            doc.text("TAX : EXTRA", x1 + 3, 543.6, {width: x2 - x1 - 6,lineBreak: false});
            if (po.payment_terms_remarks)doc.text(po.payment_terms_remarks, x1 + 3, 554.2, {width: x2 - x1 - 6,lineBreak: false});
            // Authorised signatory block
            const signX = footerSplit + 1.5;
            doc.font("Arial-Bold").fontSize(6.96);
            doc.text("FOR,", signX, footerY + 0.5, {lineBreak: false});
            doc.text("NIMIT ELECTRONICS AND EQUIPMENTS,", signX, footerY + 11.5, {width: x5 - signX - 3,lineBreak: false});
            doc.text("AUTHORISED SIGNATORY", signX, 597, {width: x5 - signX - 3,lineBreak: false});
            doc.end();
            stream.on("finish", () => resolve(outputPath));
            stream.on("error", reject);
        } catch (err) {
            reject(err);
        }
    });
}

/* ==========================================================================
   INSIGHTS HELPERS (date ranges and gap filling for Management Insights)
   ========================================================================== */

const pad2 = (n) => String(n).padStart(2, "0");
const ymd = (d) => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

// Checks that a YYYY-MM-DD string is a real calendar date
function isRealDate(str) {
    if (!DATE_RE.test(str)) return false;
    const [y, m, d] = str.split("-").map(Number);
    const dt = new Date(y, m - 1, d);
    return dt.getFullYear() === y && dt.getMonth() === m - 1 && dt.getDate() === d;
}

// Converts a period keyword (or custom from/to) into a start and end date
function resolveInsightsRange(period, from, to) {
    const now = new Date();
    const y = now.getFullYear();
    const m = now.getMonth();
    const d = now.getDate();
    switch (period) {
        case "current_month":
            return { start: ymd(new Date(y, m, 1)), end: ymd(now) };
        case "last_30_days":
            return { start: ymd(new Date(y, m, d - 29)), end: ymd(now) };
        case "last_month":
            return { start: ymd(new Date(y, m - 1, 1)), end: ymd(new Date(y, m, 0)) };
        case "last_3_months":
            return { start: ymd(new Date(y, m - 2, 1)), end: ymd(now) };
        case "last_6_months":
            return { start: ymd(new Date(y, m - 5, 1)), end: ymd(now) };
        case "last_1_year":
            return { start: ymd(new Date(y - 1, m, d)), end: ymd(now) };
        case "last_2_years":
            return { start: ymd(new Date(y - 2, m, d)), end: ymd(now) };
        case "all_time":
            return { start: null, end: null };
        case "custom":
            if (!from || !to) return { error: "Custom period requires from and to dates" };
            if (!isRealDate(from) || !isRealDate(to)) return { error: "Invalid custom date range" };
            if (from > to) return { error: "From date cannot be greater than to date" };
            return { start: from, end: to };
        default:
            return { error: "Invalid period" };
    }
}

// Lists every YYYY-MM between the first and last month present in the given keys
function monthsBetween(keys) {
    if (!keys.length) return [];
    const sorted = [...keys].sort();
    let [y, m] = sorted[0].split("-").map(Number);
    const [ly, lm] = sorted[sorted.length - 1].split("-").map(Number);
    const out = [];
    while (y < ly || (y === ly && m <= lm)) {
        out.push(`${y}-${pad2(m)}`);
        m += 1;
        if (m > 12) { m = 1; y += 1; }
    }
    return out;
}

// Lists every calendar day from a to b (YYYY-MM-DD, inclusive)
function daysBetween(a, b) {
    const [y, m, d] = a.split("-").map(Number);
    const out = [];
    for (let i = 0; i < 3700; i++) {
        const key = ymd(new Date(y, m - 1, d + i));
        if (key > b) break;
        out.push(key);
    }
    return out;
}

const num = (v) => Number(v || 0);

/* ==========================================================================
   STATIC FILES AND PAGE
   ========================================================================== */

app.use("/logo", express.static(path.join(__dirname, "..", "NIMIT LOGO.png")));
app.use(express.static(path.join(__dirname, "../../frontend/procurement-manager"), { index: false }));

// GET / - Serves the Procurement Manager dashboard page (manager login required) with the username injected
app.get("/", verifyManager, (req, res) => {
    log(`Authenticated - User: ${req.user?.username}`);
    const htmlPath = path.join(__dirname, "../../frontend/procurement-manager/index.html");
    let html = fs.readFileSync(htmlPath, "utf8");
    html = html.replace("</head>", `<script>window.currentUsername=${JSON.stringify(req.user.username)};</script></head>`);
    return res.send(html);
});

// POST /logout - Logs the Procurement Manager out through the auth service and clears the session cookie
app.post("/logout", async (req, res) => {
    log("POST /logout - Procurement Manager logout requested");
    try {
        const response = await fetch(`${authServiceUrl}/logout`, {
            method: "POST",
            headers: {
                Cookie: req.headers.cookie || ""
            }
        });
        if (!response.ok) {
            log("Auth service failed to logout procurement manager user");
            return res.status(500).json({
                success: false,
                message: "Failed to logout"
            });
        }
        res.clearCookie("login_session");
        log("Procurement Manager logout successful");
        return res.json({
            success: true,
            redirect_url: authServiceUrl
        });
    } catch (error) {
        log(`ERROR during procurement manager logout: ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to logout"
        });
    }
});

/* ==========================================================================
   TAB: PURCHASE REQUESTS
   ========================================================================== */

// POST /purchase-requests - Creates a single purchase request with the next PR number and opens a vendor inquiry for it
app.post("/purchase-requests", async (req, res) => {
    const {
        pr_date,
        party_name,
        location,
        territory,
        product_category,
        item_name,
        product_remarks,
        make,
        model,
        qty,
        unit,
        sales_rate
    } = req.body;
    const requiredFields = {
        party_name,
        location,
        territory,
        product_category,
        item_name,
        make,
        model,
        qty,
        unit,
        sales_rate
    };
    for (const [field, value] of Object.entries(requiredFields)) {
        if (value === undefined || value === null || String(value).trim() === "") {
            return res.status(400).json({
                success: false,
                message: `${field.replace(/_/g, " ")} is required`
            });
        }
    }
    if (Number(qty) <= 0) {
        return res.status(400).json({
            success: false,
            message: "Quantity must be greater than 0"
        });
    }
    if (Number(sales_rate) <= 0) {
        return res.status(400).json({
            success: false,
            message: "Sales rate must be greater than 0"
        });
    }
    const pr_date_final = pr_date && String(pr_date).trim() !== "" ? pr_date : new Date().toISOString().split("T")[0];
    const taxable_value = Number(qty) * Number(sales_rate);
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();
        const pr_number = await generatePrNumber(connection);
        const [result] = await connection.execute(
            `INSERT INTO purchase_requests
            (
                pr_number,
                pr_date,
                party_name,
                location,
                territory,
                product_category,
                item_name,
                product_remarks,
                make,
                model,
                qty,
                unit,
                sales_rate,
                taxable_value
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [
                pr_number,
                pr_date_final,
                party_name,
                location,
                territory,
                product_category,
                item_name,
                product_remarks || null,
                make,
                model,
                qty,
                unit,
                sales_rate,
                taxable_value
            ]
        );
        const [inquiryResult] = await connection.execute(
            `INSERT INTO vendor_inquiries
            (
                pr_id,
                status,
                remarks
            )
            VALUES (?, 'OPEN', ?)`,
            [
                result.insertId,
                product_remarks || null
            ]
        );
        await connection.commit();
        await writeReportLog(req, "PR_CREATED",
            `Purchase request raised. ${describePr({
                pr_number, pr_date: pr_date_final, party_name, location, territory, product_category,
                item_name, make, model, qty, unit, sales_rate, taxable_value, product_remarks
            })}`
        );
        await writeReportLog(req, "VENDOR_INQUIRY_CREATED",
            `Vendor inquiry (inquiry ID ${inquiryResult.insertId}) opened for ${pr_number}. ` +
            `Item: ${item_name}, make: ${make}, model: ${model}, quantity: ${Number(qty)} ${unit}. Status: OPEN, awaiting vendor quotations.`
        );
        return res.json({
            success: true,
            message: "Purchase request created successfully",
            pr_number,
            pr_id: result.insertId
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        log(`Purchase request creation failed - ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to create purchase request"
        });
    } finally {
        connection.release();
    }
});

// POST /purchase-requests/import-preview - Reads an uploaded Excel sheet and returns editable PR rows without saving anything
app.post("/purchase-requests/import-preview", upload.single("file"), async (req, res) => {
    log("POST /purchase-requests/import-preview - Excel preview requested");
    try {
        if (!req.file) {
            return res.status(400).json({
                success: false,
                message: "Excel file is required"
            });
        }
        const excelRows = await readSheet(req.file.buffer);
        if (!excelRows || excelRows.length <= 1) {
            return res.status(400).json({
                success: false,
                message: "Excel file contains no data"
            });
        }
        const headers = excelRows[0].map(normalizeHeader);
        const columnIndex = {};
        headers.forEach((header, index) => {
            if (header) columnIndex[header] = index;
        });
        const getCell = (row, field) => {
            if (columnIndex[field] === undefined) return null;
            return cleanValue(row[columnIndex[field]]);
        };
        const rows = excelRows.slice(1).filter(row => row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== "")).map(row => {
            const quantity = cleanNumber(getCell(row, "qty"));
            const rate = cleanNumber(getCell(row, "sales_rate"));
            return {
                pr_date: formatDate(getCell(row, "pr_date")),
                party_name: getCell(row, "party_name"),
                location: getCell(row, "location"),
                territory: getCell(row, "territory"),
                product_category: getCell(row, "product_category"),
                item_name: getCell(row, "item_name"),
                product_remarks: getCell(row, "product_remarks"),
                make: getCell(row, "make"),
                model: getCell(row, "model"),
                qty: quantity,
                unit: getCell(row, "unit"),
                sales_rate: rate,
                taxable_value: quantity * rate
            };
        });
        if (rows.length === 0) {
            return res.status(400).json({
                success: false,
                message: "Excel file contains no data"
            });
        }
        log(`Excel preview successful - ${rows.length} rows ready for editing`);
        return res.json({
            success: true,
            rows
        });
    } catch (error) {
        log(`ERROR creating Excel preview: ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to process Excel file"
        });
    }
});

// POST /purchase-requests/import - Saves the reviewed Excel rows as purchase requests (one PR number and one vendor inquiry per row) in a single transaction
app.post("/purchase-requests/import", async (req, res) => {
    const connection = await db.getConnection();
    try {
        const rows = req.body.rows || [];
        if (!Array.isArray(rows) || rows.length === 0) {
            return res.status(400).json({
                success: false,
                message: "No rows available for import"
            });
        }
        await connection.beginTransaction();
        const insertedRows = [];
        const reportItems = [];
        for (const row of rows) {
            const pr_number = await generatePrNumber(connection);
            const qty = Number(row.qty || 0);
            const sales_rate = Number(row.sales_rate || 0);
            const taxable_value = qty * sales_rate;
            const [result] = await connection.execute(
                `INSERT INTO purchase_requests
                (pr_number, pr_date, party_name, location, territory,
                 product_category, item_name, product_remarks, make, model,
                 qty, unit, sales_rate, taxable_value)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                [
                    pr_number,
                    row.pr_date || null,
                    row.party_name || null,
                    row.location || null,
                    row.territory || null,
                    row.product_category || null,
                    row.item_name || null,
                    row.product_remarks || null,
                    row.make || null,
                    row.model || null,
                    qty,
                    row.unit || null,
                    sales_rate,
                    taxable_value
                ]
            );
            const prId = result.insertId;
            const [inquiryResult] = await connection.execute(
                `INSERT INTO vendor_inquiries (pr_id, status, remarks)
                 VALUES (?, 'OPEN', ?)`,
                [prId, row.product_remarks || null]
            );
            insertedRows.push({ id: prId, pr_number });
            reportItems.push({
                inquiry_id: inquiryResult.insertId,
                pr: {
                    pr_number,
                    pr_date: row.pr_date,
                    party_name: row.party_name,
                    location: row.location,
                    territory: row.territory,
                    product_category: row.product_category,
                    item_name: row.item_name,
                    make: row.make,
                    model: row.model,
                    qty,
                    unit: row.unit,
                    sales_rate,
                    taxable_value,
                    product_remarks: row.product_remarks
                }
            });
        }
        await connection.commit();
        log(`Purchase Requests imported successfully - ${insertedRows.length} rows`);
        for (const item of reportItems) {
            await writeReportLog(req, "PR_CREATED",
                `Purchase request raised through Excel import. ${describePr(item.pr)}`
            );
            await writeReportLog(req, "VENDOR_INQUIRY_CREATED",
                `Vendor inquiry (inquiry ID ${item.inquiry_id}) opened for ${item.pr.pr_number} through Excel import. ` +
                `Item: ${orDash(item.pr.item_name)}, make: ${orDash(item.pr.make)}, model: ${orDash(item.pr.model)}, ` +
                `quantity: ${item.pr.qty} ${orDash(item.pr.unit)}. Status: OPEN, awaiting vendor quotations.`
            );
        }
        return res.json({
            success: true,
            message: "Purchase Requests imported successfully",
            rows: insertedRows
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        log(`Purchase Request Excel import failed - ${error.message}`);
        return res.status(500).json({
            success: false,
            message: error.message
        });
    } finally {
        connection.release();
    }
});

/* ==========================================================================
   TAB: ORDER TRACKING
   ========================================================================== */

// GET /order-tracking - Paginated list (10 per page) of all purchase requests with their inquiry status and PO number
app.get("/order-tracking", async (req, res) => {
    const page = Math.max(1, Number(req.query.page) || 1);
    const limit = 10;
    const offset = (page - 1) * limit;
    try {
        const [[{ total }]] = await db.execute(
            `SELECT COUNT(*) AS total FROM purchase_requests`
        );
        const totalPages = Math.max(1, Math.ceil(total / limit));
        const [rows] = await db.query(`
            SELECT
                pr.id AS pr_id,
                pr.pr_number,
                pr.pr_date,
                pr.party_name,
                pr.location,
                pr.territory,
                pr.product_category,
                pr.item_name,
                pr.make,
                pr.model,
                pr.qty,
                pr.unit,
                pr.sales_rate,
                pr.taxable_value,
                pr.product_remarks,
                vi.inquiry_id,
                po.po_number,
                CASE
                    WHEN po.status = 'COMPLETED' THEN 'CLOSED'
                    WHEN po.status = 'CANCELLED' THEN 'CANCELLED'
                    ELSE vi.status
                END AS status
            FROM purchase_requests pr
            LEFT JOIN vendor_inquiries vi
                ON vi.pr_id = pr.id
            LEFT JOIN purchase_orders po
                ON po.pr_id = pr.id
            ORDER BY pr.id DESC
            LIMIT ${limit} OFFSET ${offset}
        `);
        return res.json({
            success: true,
            rows,
            pagination: {
                page,
                limit,
                total,
                total_pages: totalPages,
                has_prev: page > 1,
                has_next: page < totalPages
            }
        });
    } catch (error) {
        console.error(error);
        log(`Order tracking fetch failed - ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to fetch order tracking data"
        });
    }
});

/* ==========================================================================
   TAB: VENDOR MASTER (registration, import, lookups)
   ========================================================================== */

// POST /vendors - Registers a vendor entered manually, together with its uploaded documents
app.post("/vendors", getVendorDocuments(), async (req, res) => {
    log("POST /vendors - Manual vendor entry");
    const connection = await db.getConnection();
    try {
        const data = parseVendorData(req);
        const files = req.files || {};
        const validation = validateVendorData(data, files);
        if (validation.missingFields.length || validation.missingDocuments.length) {
            return res.status(400).json({
                success: false,
                message: "Vendor data validation failed",
                missing_fields: validation.missingFields,
                missing_documents: validation.missingDocuments
            });
        }
        await connection.beginTransaction();
        const result = await saveVendor(connection, data, files);
        await connection.commit();
        log(`Vendor created successfully - ${result.vendorCode}`);
        await writeReportLog(req, "VENDOR_CREATED", describeVendor(data, result, "through manual entry"));
        return res.status(201).json({
            success: true,
            message: "Vendor created successfully",
            vendor_id: result.vendorId,
            vendor_code: result.vendorCode
        });
    } catch (error) {
        await connection.rollback();
        log(`ERROR creating vendor: ${error.message}`);
        return res.status(error.message.includes("already exists") ? 409 : 500).json({
            success: false,
            message: error.message.includes("already exists") ? error.message : "Failed to create vendor"
        });
    } finally {
        connection.release();
    }
});

// GET /vendors/check-gst - Checks whether a vendor with the given GST number is already registered
app.get("/vendors/check-gst", async (req, res) => {
    const gstNumber = cleanValue(req.query.gst_number);
    if (!gstNumber) {
        return res.status(400).json({
            success: false,
            message: "GST number is required"
        });
    }
    try {
        const [rows] = await db.execute(
            `SELECT vendor_name FROM vendor_oem_masters WHERE gst_number = ? LIMIT 1`,
            [gstNumber]
        );
        if (rows.length > 0) {
            return res.json({
                success: true,
                exists: true,
                vendor_name: rows[0].vendor_name
            });
        }
        return res.json({
            success: true,
            exists: false
        });
    } catch (error) {
        log(`ERROR checking GST number: ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to check GST number"
        });
    }
});

// POST /vendors/import-preview - Reads an uploaded vendor Excel (registration form or one-row table) and returns the parsed vendor without saving
app.post("/vendors/import-preview", upload.single("file"), async (req, res) => {
    log("POST /vendors/import-preview - Vendor Excel preview requested");
    try {
        if (!req.file) {
            return res.status(400).json({
                success: false,
                message: "Excel file is required"
            });
        }
        const excelRows = await readSheet(req.file.buffer);
        if (!excelRows || excelRows.length <= 1) {
            return res.status(400).json({
                success: false,
                message: "Excel file contains no vendor data"
            });
        }
        const vendor = isVendorForm(excelRows) ? parseVendorForm(excelRows) : parseVendorTable(excelRows);
        log("Vendor Excel preview successful");
        return res.json({
            success: true,
            vendor
        });
    } catch (error) {
        log(`ERROR creating vendor Excel preview: ${error.message}`);
        return res.status(error.status || 500).json({
            success: false,
            message: error.status ? error.message : "Failed to process vendor Excel file"
        });
    }
});

// POST /vendors/import - Saves the reviewed vendor from an Excel import together with its uploaded documents
app.post("/vendors/import", getVendorDocuments(), async (req, res) => {
    log("POST /vendors/import - Saving imported vendor");
    const connection = await db.getConnection();
    try {
        const data = parseVendorData(req);
        const files = req.files || {};
        const validation = validateVendorData(data, files);
        if (validation.missingFields.length || validation.missingDocuments.length) {
            return res.status(400).json({
                success: false,
                message: "Vendor data validation failed",
                missing_fields: validation.missingFields,
                missing_documents: validation.missingDocuments
            });
        }
        await connection.beginTransaction();
        const result = await saveVendor(connection, data, files);
        await connection.commit();
        log(`Imported vendor saved successfully - ${result.vendorCode}`);
        await writeReportLog(req, "VENDOR_IMPORTED", describeVendor(data, result, "through Excel import"));
        return res.status(201).json({
            success: true,
            message: "Vendor imported successfully",
            vendor_id: result.vendorId,
            vendor_code: result.vendorCode
        });
    } catch (error) {
        await connection.rollback();
        log(`ERROR saving imported vendor: ${error.message}`);
        return res.status(error.message.includes("already exists") ? 409 : 500).json({
            success: false,
            message: error.message.includes("already exists") ? error.message : "Failed to import vendor"
        });
    } finally {
        connection.release();
    }
});

// GET /vendors - Lists all non-blacklisted vendors (oldest first) for vendor dropdowns
app.get("/vendors", async (req, res) => {
    try {
        const [rows] = await db.execute(
            `SELECT
                vendor_id,
                vendor_code,
                vendor_name,
                gst_number
             FROM vendor_oem_masters
             WHERE is_blacklisted = FALSE
             ORDER BY vendor_id ASC`
        );
        return res.json({
            success: true,
            vendors: rows
        });
    } catch (error) {
        console.error(error);
        log(`Vendor list fetch failed - ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to fetch vendors"
        });
    }
});

/* ==========================================================================
   TAB: VENDOR INQUIRIES AND QUOTATIONS
   ========================================================================== */

// GET /vendor-inquiries - Lists all OPEN vendor inquiries with their purchase request details
app.get("/vendor-inquiries", async (req, res) => {
    try {
        const [inquiries] = await db.execute(`
            SELECT
                vi.inquiry_id,
                pr.id AS pr_id,
                pr.pr_number,
                pr.pr_date,
                pr.party_name,
                pr.location,
                pr.territory,
                pr.product_category,
                pr.item_name,
                pr.make,
                pr.model,
                pr.qty,
                pr.unit,
                pr.sales_rate,
                pr.taxable_value,
                pr.product_remarks
            FROM vendor_inquiries vi
            INNER JOIN purchase_requests pr
                ON pr.id = vi.pr_id
            WHERE vi.status = 'OPEN'
            ORDER BY pr.id DESC
        `);
        return res.json({
            success: true,
            inquiries
        });
    } catch (error) {
        console.error(error);
        log(`Vendor Inquiry fetch failed - ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to fetch open vendor inquiries"
        });
    }
});

// GET /vendor-inquiries/:inquiry_id - Returns one OPEN inquiry with all vendor quotations added to it so far
app.get("/vendor-inquiries/:inquiry_id", async (req, res) => {
    const inquiryId = Number(req.params.inquiry_id);
    if (!Number.isInteger(inquiryId) || inquiryId <= 0) {
        return res.status(400).json({ success: false, message: "Invalid inquiry ID" });
    }
    try {
        const [inquiries] = await db.execute(`
            SELECT
                vi.inquiry_id,
                vi.pr_id,
                pr.pr_number,
                pr.party_name,
                pr.item_name,
                pr.make,
                pr.model,
                pr.qty,
                pr.unit,
                vi.remarks
            FROM vendor_inquiries vi
            INNER JOIN purchase_requests pr ON vi.pr_id = pr.id
            WHERE vi.inquiry_id = ? AND vi.status = 'OPEN'
            LIMIT 1
        `, [inquiryId]);
        if (inquiries.length === 0)return res.status(404).json({ success: false, message: "Open Vendor Inquiry not found" });
        const [vendors] = await db.execute(`
            SELECT
                viv.inquiry_vendor_id,
                viv.vendor_id,
                vom.vendor_code,
                vom.vendor_name,
                viv.price_per_unit,
                viv.total_price,
                viv.expected_delivery_date,
                viv.payment_type,
                viv.advance_type,
                viv.advance_value,
                viv.balance_due_days,
                viv.payment_terms_remarks,
                viv.remarks
            FROM vendor_inquiry_vendors viv
            INNER JOIN vendor_oem_masters vom ON viv.vendor_id = vom.vendor_id
            WHERE viv.inquiry_id = ?
            ORDER BY viv.inquiry_vendor_id DESC
        `, [inquiryId]);
        return res.json({ success: true, inquiry: inquiries[0], vendors });
    } catch (error) {
        console.error(error);
        log(`Vendor Inquiry details fetch failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to fetch Vendor Inquiry details" });
    }
});

// POST /vendor-inquiries/:inquiry_id/vendors - Adds a vendor's quotation (price, delivery, payment terms) to an OPEN inquiry
app.post("/vendor-inquiries/:inquiry_id/vendors", async (req, res) => {
    const inquiryId = Number(req.params.inquiry_id);
    const {
        vendor_id,
        price_per_unit,
        expected_delivery_date,
        payment_type,
        advance_type,
        advance_value,
        balance_due_days,
        payment_terms_remarks,
        remarks
    } = req.body;
    if (!Number.isInteger(inquiryId) || inquiryId <= 0)return res.status(400).json({ success: false, message: "Invalid inquiry ID" });
    const vendorId = Number(vendor_id);
    const price = Number(price_per_unit);
    if (!Number.isInteger(vendorId) || vendorId <= 0)return res.status(400).json({ success: false, message: "Please select a valid vendor" });
    if (price_per_unit === undefined || price_per_unit === null || String(price_per_unit).trim() === "" || !Number.isFinite(price) || price <= 0)return res.status(400).json({ success: false, message: "Price per unit must be greater than 0" });
    const validPaymentTypes = ["ADVANCE", "CREDIT", "ADVANCE_PLUS_BALANCE", "CUSTOM"];
    if (!payment_type || !validPaymentTypes.includes(payment_type)) {
        return res.status(400).json({ success: false, message: "Please select a valid payment type" });
    }
    if (["ADVANCE", "ADVANCE_PLUS_BALANCE"].includes(payment_type)) {
        if (!advance_type || !["PERCENTAGE", "FIXED_AMOUNT"].includes(advance_type))return res.status(400).json({ success: false, message: "Advance type is required for this payment type" });
        if (!advance_value || Number(advance_value) <= 0)return res.status(400).json({ success: false, message: "Advance value is required for this payment type" });
    }
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();
        const [inquiryRows] = await connection.execute(`
            SELECT vi.inquiry_id, pr.qty, pr.pr_number, pr.item_name, pr.make, pr.model, pr.unit
            FROM vendor_inquiries vi
            INNER JOIN purchase_requests pr ON vi.pr_id = pr.id
            WHERE vi.inquiry_id = ? AND vi.status = 'OPEN'
            FOR UPDATE
        `, [inquiryId]);
        if (inquiryRows.length === 0) {
            await connection.rollback();
            return res.status(404).json({ success: false, message: "Open Vendor Inquiry not found" });
        }
        const [vendorRows] = await connection.execute(`
            SELECT vendor_id, vendor_name, vendor_code
            FROM vendor_oem_masters
            WHERE vendor_id = ? AND is_blacklisted = FALSE
            LIMIT 1
        `, [vendorId]);
        if (vendorRows.length === 0) {
            await connection.rollback();
            return res.status(400).json({ success: false, message: "Vendor does not exist or is blacklisted" });
        }
        const [existingRows] = await connection.execute(`
            SELECT inquiry_vendor_id
            FROM vendor_inquiry_vendors
            WHERE inquiry_id = ? AND vendor_id = ?
            LIMIT 1
        `, [inquiryId, vendorId]);
        if (existingRows.length > 0) {
            await connection.rollback();
            return res.status(409).json({ success: false, message: "This vendor has already been added to the inquiry" });
        }
        const qty = Number(inquiryRows[0].qty);
        const totalPrice = qty * price;
        let advanceAmount = null;
        if (advance_value && Number(advance_value) > 0) {
            advanceAmount = parseFloat(((Number(advance_value) / 100) * totalPrice).toFixed(2));
        }
        if (!Number.isFinite(totalPrice)) {
            await connection.rollback();
            return res.status(400).json({ success: false, message: "Calculated total price is invalid" });
        }
        const [result] = await connection.execute(`
            INSERT INTO vendor_inquiry_vendors
            (
                inquiry_id,
                vendor_id,
                price_per_unit,
                total_price,
                expected_delivery_date,
                payment_type,
                advance_type,
                advance_value,
                advance_amount,
                balance_due_days,
                payment_terms_remarks,
                is_selected,
                remarks
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, FALSE, ?)
        `, [
            inquiryId,
            vendorId,
            price,
            totalPrice,
            expected_delivery_date || null,
            payment_type,
            advance_type || null,
            advance_value ? Number(advance_value) : null,
            advanceAmount,
            balance_due_days ? Number(balance_due_days) : null,
            payment_terms_remarks?.trim() || null,
            remarks?.trim() || null
        ]);
        await connection.commit();
        log(`Vendor added to Inquiry - Inquiry ID: ${inquiryId}, Vendor ID: ${vendorId}`);
        const inq = inquiryRows[0];
        const vendorRow = vendorRows[0];
        await writeReportLog(req, "VENDOR_QUOTATION_ADDED",
            `Vendor "${vendorRow.vendor_name}" (${orDash(vendorRow.vendor_code)}) added to vendor inquiry ${inquiryId} for ${inq.pr_number} with a quotation. ` +
            `Item: ${orDash(inq.item_name)}, make: ${orDash(inq.make)}, model: ${orDash(inq.model)}, quantity: ${qty} ${orDash(inq.unit)}. ` +
            `Quoted price per unit: ${money(price)}, total price: ${money(totalPrice)}. ` +
            `Expected delivery: ${dateText(expected_delivery_date)}. ` +
            `${describePayment({ payment_type, advance_type, advance_value, balance_due_days, payment_terms_remarks })}` +
            `${advanceAmount !== null ? `, advance amount: ${money(advanceAmount)}` : ""}. ` +
            `Remarks: ${orDash(remarks)}.`
        );
        return res.status(201).json({
            success: true,
            message: "Vendor quotation added successfully",
            inquiry_vendor_id: result.insertId,
            total_price: totalPrice
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        log(`Vendor quotation addition failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to add vendor quotation" });
    } finally {
        connection.release();
    }
});

// GET /quotation-comparisons - Lists all OPEN inquiries with every vendor quotation (cheapest first) for side-by-side comparison
app.get("/quotation-comparisons", async (req, res) => {
    try {
        const [inquiries] = await db.execute(`
            SELECT
                vi.inquiry_id,
                pr.id AS pr_id,
                pr.pr_number,
                pr.pr_date,
                pr.party_name,
                pr.location,
                pr.territory,
                pr.product_category,
                pr.item_name,
                pr.make,
                pr.model,
                pr.qty,
                pr.unit,
                pr.sales_rate,
                pr.taxable_value,
                pr.product_remarks
            FROM vendor_inquiries vi
            INNER JOIN purchase_requests pr
                ON pr.id = vi.pr_id
            WHERE vi.status = 'OPEN'
            ORDER BY pr.id DESC
        `);
        if (!inquiries.length)return res.json({ success: true, inquiries: [] });
        const inquiryIds = inquiries.map(i => i.inquiry_id);
        const [vendors] = await db.execute(`
            SELECT
                viv.inquiry_vendor_id,
                viv.inquiry_id,
                viv.vendor_id,
                vom.vendor_code,
                vom.vendor_name,
                viv.price_per_unit,
                viv.total_price,
                viv.advance_type,
                viv.advance_value,
                viv.advance_amount,
                viv.expected_delivery_date,
                viv.balance_due_days,
                viv.payment_type,
                viv.is_selected,
                viv.remarks
            FROM vendor_inquiry_vendors viv
            INNER JOIN vendor_oem_masters vom
                ON viv.vendor_id = vom.vendor_id
            WHERE viv.inquiry_id IN (${inquiryIds.map(() => "?").join(",")})
            ORDER BY viv.inquiry_id, viv.price_per_unit ASC
        `, inquiryIds);
        const vendorMap = {};
        vendors.forEach(v => {
            if (!vendorMap[v.inquiry_id]) vendorMap[v.inquiry_id] = [];
            vendorMap[v.inquiry_id].push(v);
        });
        const result = inquiries.map(inq => ({...inq,vendors: vendorMap[inq.inquiry_id] || []}));
        return res.json({ success: true, inquiries: result });
    } catch (error) {
        console.error(error);
        log(`Quotation comparison fetch failed - ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to fetch quotation comparisons"
        });
    }
});

// POST /vendor-inquiries/:inquiry_id/select-vendor - Finalizes the chosen vendor quotation and creates a DRAFT purchase order from it
app.post("/vendor-inquiries/:inquiry_id/select-vendor", async (req, res) => {
    const inquiryId = Number(req.params.inquiry_id);
    const { inquiry_vendor_id } = req.body;
    if (!Number.isInteger(inquiryId) || inquiryId <= 0)return res.status(400).json({ success: false, message: "Invalid inquiry ID" });
    const inquiryVendorId = Number(inquiry_vendor_id);
    if (!Number.isInteger(inquiryVendorId) || inquiryVendorId <= 0)return res.status(400).json({ success: false, message: "Please select a valid vendor quotation" });
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();
        const [inquiryRows] = await connection.execute(`
            SELECT vi.inquiry_id, vi.status, vi.pr_id
            FROM vendor_inquiries vi
            WHERE vi.inquiry_id = ?
            FOR UPDATE
        `, [inquiryId]);
        if (inquiryRows.length === 0) {
            await connection.rollback();
            return res.status(404).json({ success: false, message: "Vendor Inquiry not found" });
        }
        if (inquiryRows[0].status !== "OPEN") {
            await connection.rollback();
            return res.status(409).json({ success: false, message: "A vendor has already been selected for this inquiry" });
        }
        const prId = inquiryRows[0].pr_id;
        const [vendorRows] = await connection.execute(`
            SELECT
                viv.inquiry_vendor_id,
                viv.vendor_id,
                viv.price_per_unit,
                viv.total_price,
                viv.expected_delivery_date,
                viv.payment_type,
                viv.advance_type,
                viv.advance_value,
                viv.balance_due_days,
                viv.payment_terms_remarks,
                viv.remarks,
                vom.vendor_code,
                vom.vendor_name,
                vom.office_address,
                vom.office_contact_name,
                vom.office_contact_number
            FROM vendor_inquiry_vendors viv
            INNER JOIN vendor_oem_masters vom ON vom.vendor_id = viv.vendor_id
            WHERE viv.inquiry_vendor_id = ? AND viv.inquiry_id = ?
            LIMIT 1
        `, [inquiryVendorId, inquiryId]);
        if (vendorRows.length === 0) {
            await connection.rollback();
            return res.status(404).json({ success: false, message: "Vendor quotation not found for this inquiry" });
        }
        const [prRows] = await connection.execute(`
            SELECT * FROM purchase_requests WHERE id = ? LIMIT 1
        `, [prId]);
        if (prRows.length === 0) {
            await connection.rollback();
            return res.status(404).json({ success: false, message: "Purchase request not found" });
        }
        const [quotationRows] = await connection.execute(`
            SELECT vom.vendor_name, viv.price_per_unit, viv.total_price
            FROM vendor_inquiry_vendors viv
            INNER JOIN vendor_oem_masters vom ON vom.vendor_id = viv.vendor_id
            WHERE viv.inquiry_id = ?
            ORDER BY viv.price_per_unit ASC
        `, [inquiryId]);
        const viv = vendorRows[0];
        const pr  = prRows[0];
        const contactLine = [viv.office_contact_name, viv.office_contact_number].filter(Boolean).join(" - ");
        const vendorAddress = [viv.office_address, contactLine].filter(Boolean).join("\n");
        await connection.execute(`
            UPDATE vendor_inquiry_vendors SET is_selected = FALSE WHERE inquiry_id = ?
        `, [inquiryId]);
        await connection.execute(`
            UPDATE vendor_inquiry_vendors SET is_selected = TRUE WHERE inquiry_vendor_id = ?
        `, [inquiryVendorId]);
        await connection.execute(`
            UPDATE vendor_inquiries SET status = 'VENDOR_SELECTED' WHERE inquiry_id = ?
        `, [inquiryId]);
        const poNumber = await generatePoNumber(connection, pr.pr_number);
        const poDate   = new Date().toISOString().split("T")[0];
        await connection.execute(`
            INSERT INTO purchase_orders (
                po_number,
                po_date,
                pr_id,
                inquiry_id,
                inquiry_vendor_id,
                vendor_id,
                pr_number,
                pr_date,
                party_name,
                location,
                territory,
                product_category,
                item_name,
                product_remarks,
                make,
                model,
                qty,
                unit,
                sales_rate,
                taxable_value,
                vendor_code,
                vendor_name,
                vendor_address,
                price_per_unit,
                total_price,
                expected_delivery_date,
                payment_type,
                advance_type,
                advance_value,
                balance_due_days,
                payment_terms_remarks,
                status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'DRAFT')
        `, [
            poNumber,
            poDate,
            prId,
            inquiryId,
            inquiryVendorId,
            viv.vendor_id,
            pr.pr_number,
            pr.pr_date,
            pr.party_name,
            pr.location,
            pr.territory,
            pr.product_category,
            pr.item_name,
            pr.product_remarks,
            pr.make,
            pr.model,
            pr.qty,
            pr.unit,
            pr.sales_rate,
            pr.taxable_value,
            viv.vendor_code,
            viv.vendor_name,
            vendorAddress,
            viv.price_per_unit,
            viv.total_price,
            viv.expected_delivery_date || null,
            viv.payment_type,
            viv.advance_type || null,
            viv.advance_value || null,
            viv.balance_due_days || null,
            viv.payment_terms_remarks || null
        ]);
        await connection.commit();
        log(`Vendor selected and PO created - Inquiry ID: ${inquiryId}, PO: ${poNumber}`);
        const comparison = quotationRows.map((q, i) => `${i + 1}. ${q.vendor_name} - ${money(q.price_per_unit)} per unit (total ${money(q.total_price)})`).join("; ");
        const grossProfit = Number(pr.taxable_value || 0) - Number(viv.total_price || 0);
        await writeReportLog(req, "VENDOR_SELECTED",
            `Quotation finalized for vendor inquiry ${inquiryId} (${pr.pr_number}). ` +
            `${quotationRows.length} quotation(s) compared, cheapest first: ${comparison}. ` +
            `Selected vendor: "${viv.vendor_name}" (${orDash(viv.vendor_code)}) at ${money(viv.price_per_unit)} per unit, total ${money(viv.total_price)}. ` +
            `Item: ${orDash(pr.item_name)}, make: ${orDash(pr.make)}, model: ${orDash(pr.model)}, quantity: ${Number(pr.qty)} ${orDash(pr.unit)}. ` +
            `Sales taxable value: ${money(pr.taxable_value)}, expected gross profit: ${money(grossProfit)}. ` +
            `Inquiry status changed from OPEN to VENDOR_SELECTED.`
        );
        await writeReportLog(req, "PO_CREATED",
            `Draft purchase order ${poNumber} dated ${poDate} created from ${pr.pr_number} (inquiry ID ${inquiryId}). ` +
            `Vendor: "${viv.vendor_name}" (${orDash(viv.vendor_code)}). Party: "${orDash(pr.party_name)}" (${orDash(pr.location)}, ${orDash(pr.territory)}). ` +
            `Item: ${orDash(pr.item_name)}, make: ${orDash(pr.make)}, model: ${orDash(pr.model)}, quantity: ${Number(pr.qty)} ${orDash(pr.unit)}. ` +
            `Price per unit: ${money(viv.price_per_unit)}, PO total: ${money(viv.total_price)}. ` +
            `Expected delivery: ${dateText(viv.expected_delivery_date)}. ${describePayment(viv)}. Status: DRAFT.`
        );
        return res.json({
            success: true,
            message: "Vendor selected successfully",
            vendor_id: viv.vendor_id,
            po_number: poNumber
        });

    } catch (error) {
        await connection.rollback();
        console.error(error);
        log(`Vendor selection failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to select vendor" });
    } finally {
        connection.release();
    }
});

/* ==========================================================================
   TAB: PURCHASE ORDERS
   ========================================================================== */

// GET /purchase-orders - Lists all purchase orders (newest first) with full vendor master details and quotation remarks
app.get("/purchase-orders", async (req, res) => {
    try {
        const [rows] = await db.execute(`
            SELECT
                po.po_id,
                po.po_number,
                po.po_date,
                po.pr_number,
                po.pr_date,
                po.party_name,
                po.location,
                po.territory,
                po.product_category,
                po.item_name,
                po.product_remarks,
                po.make,
                po.model,
                po.qty,
                po.unit,
                po.sales_rate,
                po.taxable_value,
                po.vendor_id,
                vom.vendor_code,
                vom.vendor_name,
                po.vendor_address,
                po.price_per_unit,
                po.total_price,
                po.expected_delivery_date,
                po.payment_type,
                po.advance_type,
                po.advance_value,
                po.balance_due_days,
                po.payment_terms_remarks,
                po.status,
                vom.registration_date,
                vom.legal_entity,
                vom.commercial_role,
                vom.year_of_incorporation,
                vom.office_address,
                vom.office_contact_name,
                vom.office_contact_number,
                vom.factory_address,
                vom.factory_contact_name,
                vom.factory_contact_number,
                vom.warehouse_address,
                vom.warehouse_contact_name,
                vom.warehouse_contact_number,
                vom.workshop_address,
                vom.workshop_contact_name,
                vom.workshop_contact_number,
                CONCAT_WS(' - ',
                    vom.office_address,
                    vom.office_contact_name,
                    vom.office_contact_number
                ) AS office_address_and_phone,
                CONCAT_WS(' - ',
                    vom.factory_address,
                    vom.factory_contact_name,
                    vom.factory_contact_number
                ) AS factory_address_and_phone,
                CONCAT_WS(' - ',
                    vom.warehouse_address,
                    vom.warehouse_contact_name,
                    vom.warehouse_contact_number
                ) AS warehouse_address_and_phone,
                CONCAT_WS(' - ',
                    vom.workshop_address,
                    vom.workshop_contact_name,
                    vom.workshop_contact_number
                ) AS workshop_address_and_phone,
                vom.director_or_ceo_or_management_name,
                vom.director_or_ceo_or_management_designation,
                vom.director_or_ceo_or_management_mobile_no,
                vom.director_or_ceo_or_management_email,
                vom.director_or_ceo_or_management_web_address,
                vom.sales_team_name,
                vom.sales_team_contact,
                vom.sales_team_email,
                vom.accounts_team_name,
                vom.accounts_team_contact,
                vom.accounts_team_email,
                vom.gst_number,
                vom.pan_number,
                vom.msme_number,
                vom.bank_name,
                vom.bank_account_no,
                vom.bank_branch,
                vom.bank_account_type,
                vom.bank_ifsc,
                CONCAT_WS(' - ',
                    vom.bank_name,
                    vom.bank_account_no,
                    vom.bank_branch,
                    vom.bank_account_type,
                    vom.bank_ifsc
                ) AS bank_details,
                vom.branch_office_1_address,
                vom.branch_office_1_contact_name,
                vom.branch_office_1_contact_number,
                CONCAT_WS(' - ',
                    vom.branch_office_1_address,
                    vom.branch_office_1_contact_name,
                    vom.branch_office_1_contact_number
                ) AS branch_office_1,
                vom.branch_office_2_address,
                vom.branch_office_2_contact_name,
                vom.branch_office_2_contact_number,
                CONCAT_WS(' - ',
                    vom.branch_office_2_address,
                    vom.branch_office_2_contact_name,
                    vom.branch_office_2_contact_number
                ) AS branch_office_2,
                vom.branch_office_3_address,
                vom.branch_office_3_contact_name,
                vom.branch_office_3_contact_number,
                CONCAT_WS(' - ',
                    vom.branch_office_3_address,
                    vom.branch_office_3_contact_name,
                    vom.branch_office_3_contact_number
                ) AS branch_office_3,
                vom.turnover_year_1,
                vom.turnover_value_1,
                vom.turnover_year_2,
                vom.turnover_value_2,
                vom.turnover_year_3,
                vom.turnover_value_3,
                vom.recommended_by,
                vom.approved_by,
                vom.client_details,
                vom.is_blacklisted,
                viv.remarks AS quotation_remarks
            FROM purchase_orders po
            LEFT JOIN vendor_oem_masters vom
                ON vom.vendor_id = po.vendor_id
            LEFT JOIN vendor_inquiry_vendors viv
                ON viv.inquiry_vendor_id = po.inquiry_vendor_id
            ORDER BY po.po_id DESC
        `);
        return res.json({
            success: true,
            purchase_orders: rows
        });
    } catch (error) {
        console.error(error);
        log(`Purchase order fetch failed - ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to fetch purchase orders"
        });
    }
});

// POST /purchase-orders/:po_id/issue - Issues a DRAFT purchase order: generates the PDF, marks it ISSUED and returns the PDF as a download
app.post("/purchase-orders/:po_id/issue", async (req, res) => {
    const poId = Number(req.params.po_id);
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();
        const [rows] = await connection.execute(`
            SELECT
                po.*,
                vom.office_address AS vendor_address,
                vom.gst_number     AS vendor_gst
            FROM purchase_orders po
            LEFT JOIN vendor_oem_masters vom ON vom.vendor_id = po.vendor_id
            WHERE po.po_id = ?
            FOR UPDATE
        `, [poId]);
        if (!rows.length) {
            await connection.rollback();
            return res.status(404).json({ success: false, message: "PO not found" });
        }
        if (rows[0].status !== "DRAFT") {
            await connection.rollback();
            return res.status(409).json({ success: false, message: "PO is not in DRAFT status" });
        }
        const po = rows[0];
        const pdfPath = await generatePoPdf(po);
        const relPath = path.relative(__dirname, pdfPath).replace(/\\/g, "/");
        await connection.execute(
            `UPDATE purchase_orders SET status = 'ISSUED', issued_po_path = ? WHERE po_id = ?`,
            [relPath, poId]
        );
        await connection.commit();
        log(`PO issued - PO ID: ${poId}, Number: ${po.po_number}`);
        await writeReportLog(req, "PO_ISSUED",
            `Purchase order ${po.po_number} issued to vendor "${orDash(po.vendor_name)}" (${orDash(po.vendor_code)}, GST: ${orDash(po.vendor_gst)}). ` +
            `Reference ${po.pr_number} for party "${orDash(po.party_name)}". ` +
            `Item: ${orDash(po.item_name)}, make: ${orDash(po.make)}, model: ${orDash(po.model)}, quantity: ${Number(po.qty)} ${orDash(po.unit)}. ` +
            `Price per unit: ${money(po.price_per_unit)}, PO total: ${money(po.total_price)}. ` +
            `Expected delivery: ${dateText(po.expected_delivery_date)}. ${describePayment(po)}. ` +
            `Status changed from DRAFT to ISSUED and the PO PDF was generated.`
        );
        const safeFileName = po.po_number.replace(/\//g, "_") + ".pdf";
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", `attachment; filename="${safeFileName}"`);
        res.setHeader("Content-Length", fs.statSync(pdfPath).size);
        fs.createReadStream(pdfPath).pipe(res);
    } catch (error) {
        await connection.rollback();
        console.error(error);
        log(`PO issue failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to issue PO" });
    } finally {
        connection.release();
    }
});

// GET /purchase-orders/:po_id/download - Downloads the previously generated PDF of an issued purchase order
app.get("/purchase-orders/:po_id/download", async (req, res) => {
    const poId = Number(req.params.po_id);
    if (!Number.isInteger(poId) || poId <= 0)return res.status(400).json({ success: false, message: "Invalid PO ID" });
    try {
        const [rows] = await db.execute(
            `SELECT po_number, issued_po_path FROM purchase_orders WHERE po_id = ? LIMIT 1`,
            [poId]
        );
        if (!rows.length)return res.status(404).json({ success: false, message: "Purchase Order not found" });
        const { po_number, issued_po_path } = rows[0];
        if (!issued_po_path)return res.status(404).json({ success: false, message: "PDF not yet generated for this PO" });
        const absPath = path.join(__dirname, issued_po_path);
        if (!fs.existsSync(absPath))return res.status(404).json({ success: false, message: "PDF file not found on server" });
        const safeFileName = po_number.replace(/\//g, "_") + ".pdf";
        res.setHeader("Content-Type", "application/pdf");
        res.setHeader("Content-Disposition", `attachment; filename="${safeFileName}"`);
        res.setHeader("Content-Length", fs.statSync(absPath).size);
        fs.createReadStream(absPath).pipe(res);
    } catch (error) {
        console.error(error);
        log(`PO download failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to download PO" });
    }
});

// POST /purchase-orders/:po_id/complete - Manually marks a purchase order COMPLETED (needs at least one receipt) and closes its inquiry
app.post("/purchase-orders/:po_id/complete", async (req, res) => {
    const poId = Number(req.params.po_id);
    if (!Number.isInteger(poId) || poId <= 0)return res.status(400).json({ success: false, message: "Invalid PO ID" });
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();
        const [poRows] = await connection.execute(`
            SELECT po_id, qty, status, po_number, pr_number, vendor_name, vendor_code,
                   item_name, make, model, unit, total_price, inquiry_id
            FROM purchase_orders
            WHERE po_id = ?
            FOR UPDATE
        `, [poId]);
        if (poRows.length === 0) {
            await connection.rollback();
            return res.status(404).json({ success: false, message: "Purchase Order not found" });
        }
        const po = poRows[0];
        if (po.status === "CANCELLED") {
            await connection.rollback();
            return res.status(409).json({ success: false, message: "Cancelled PO cannot be completed" });
        }
        if (po.status === "COMPLETED") {
            await connection.rollback();
            return res.status(409).json({ success: false, message: "PO is already completed" });
        }
        const [receivedRows] = await connection.execute(`
            SELECT COALESCE(SUM(received_quantity), 0) AS total_received
            FROM goods_received
            WHERE po_id = ?
        `, [poId]);
        const totalReceived = Number(receivedRows[0].total_received);
        if (totalReceived === 0) {
            await connection.rollback();
            return res.status(400).json({ success: false, message: "Cannot complete a PO with no goods received" });
        }
        await connection.execute(`
            UPDATE purchase_orders SET status = 'COMPLETED' WHERE po_id = ?
        `, [poId]);
        await connection.execute(`
            UPDATE vendor_inquiries
            SET status = 'CLOSED'
            WHERE inquiry_id = (
                SELECT inquiry_id FROM purchase_orders WHERE po_id = ?
            )
        `, [poId]);
        await connection.commit();
        log(`PO completed manually - PO: ${poId}`);
        await writeReportLog(req, "PO_COMPLETED",
            `Purchase order ${po.po_number} (${po.pr_number}) marked as completed. ` +
            `Vendor: "${orDash(po.vendor_name)}" (${orDash(po.vendor_code)}). ` +
            `Item: ${orDash(po.item_name)}, make: ${orDash(po.make)}, model: ${orDash(po.model)}. ` +
            `Ordered quantity: ${Number(po.qty)} ${orDash(po.unit)}, total received before completion: ${totalReceived} ${orDash(po.unit)}` +
            `${totalReceived < Number(po.qty) ? ` (short by ${Number(po.qty) - totalReceived} ${orDash(po.unit)})` : ""}. ` +
            `PO value: ${money(po.total_price)}. Status changed to COMPLETED and vendor inquiry ${orDash(po.inquiry_id)} closed.`
        );
        return res.json({ success: true, message: "Purchase Order marked as completed" });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        log(`PO complete failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to complete Purchase Order" });
    } finally {
        connection.release();
    }
});

// POST /purchase-orders/:po_id/proforma-invoice - Uploads (or replaces) the Proforma
// Invoice for an issued/completed PO. Stored at backend/proforma-invoice/<PO number
// with PO swapped for PI>.<ext>. Report log: PI_UPLOADED.
app.post("/purchase-orders/:po_id/proforma-invoice", upload.single("proforma_invoice"), async (req, res) => {
    const poId = Number(req.params.po_id);
    if (!Number.isInteger(poId) || poId <= 0)return res.status(400).json({ success: false, message: "Invalid PO ID" });
    if (!req.file)return res.status(400).json({ success: false, message: "Proforma Invoice file is required" });
    try {
        const [rows] = await db.execute(
            `SELECT po_id, po_number, status FROM purchase_orders WHERE po_id = ? LIMIT 1`,
            [poId]
        );
        if (!rows.length)return res.status(404).json({ success: false, message: "Purchase Order not found" });
        const po = rows[0];
        if (po.status === "DRAFT")return res.status(409).json({ success: false, message: "PO must be issued before uploading a Proforma Invoice" });
        fs.mkdirSync(piFolder, { recursive: true });
        // Replace any previously uploaded PI (possibly with a different extension).
        const existing = findPiFilePath(po.po_number);
        if (existing) fs.unlinkSync(existing);
        const fileName = getPiFileName(po.po_number, req.file.originalname);
        const filePath = path.join(piFolder, fileName);
        fs.writeFileSync(filePath, req.file.buffer);
        log(`Proforma Invoice uploaded - PO ID: ${poId}, File: ${fileName}`);
        await writeReportLog(
            req,
            "PI_UPLOADED",
            `Proforma Invoice uploaded for Purchase Order ${po.po_number} (File: ${fileName}).`
        );
        return res.json({
            success: true,
            message: "Proforma Invoice uploaded successfully",
            file_name: fileName
        });
    } catch (error) {
        console.error(error);
        log(`Proforma Invoice upload failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to upload Proforma Invoice" });
    }
});

// GET /purchase-orders/:po_id/proforma-invoice/status - Tells the frontend whether a
// Proforma Invoice already exists for this PO, so it can show Upload vs Download.
app.get("/purchase-orders/:po_id/proforma-invoice/status", async (req, res) => {
    const poId = Number(req.params.po_id);
    if (!Number.isInteger(poId) || poId <= 0)return res.status(400).json({ success: false, message: "Invalid PO ID" });
    try {
        const [rows] = await db.execute(
            `SELECT po_number FROM purchase_orders WHERE po_id = ? LIMIT 1`,
            [poId]
        );
        if (!rows.length)return res.status(404).json({ success: false, message: "Purchase Order not found" });
        return res.json({ success: true, exists: !!findPiFilePath(rows[0].po_number) });
    } catch (error) {
        console.error(error);
        log(`Proforma Invoice status check failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to check Proforma Invoice status" });
    }
});

// GET /purchase-orders/:po_id/proforma-invoice - Downloads the uploaded Proforma
// Invoice for a PO. Report log: PI_DOWNLOADED.
app.get("/purchase-orders/:po_id/proforma-invoice", async (req, res) => {
    const poId = Number(req.params.po_id);
    if (!Number.isInteger(poId) || poId <= 0)return res.status(400).json({ success: false, message: "Invalid PO ID" });
    try {
        const [rows] = await db.execute(
            `SELECT po_number FROM purchase_orders WHERE po_id = ? LIMIT 1`,
            [poId]
        );
        if (!rows.length)return res.status(404).json({ success: false, message: "Purchase Order not found" });
        const po_number = rows[0].po_number;
        const filePath  = findPiFilePath(po_number);
        if (!filePath)return res.status(404).json({ success: false, message: "Proforma Invoice not yet uploaded for this PO" });
        await writeReportLog(req,"PI_DOWNLOADED",`Proforma Invoice downloaded for Purchase Order ${po_number}.`);
        res.setHeader("Content-Disposition", `attachment; filename="${path.basename(filePath)}"`);
        res.setHeader("Content-Length", fs.statSync(filePath).size);
        fs.createReadStream(filePath).pipe(res);
    } catch (error) {
        console.error(error);
        log(`Proforma Invoice download failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to download Proforma Invoice" });
    }
});

/* ==========================================================================
   TAB: GOODS RECEIVED AND RETURNS
   ========================================================================== */

// GET /goods-received - Lists ISSUED purchase orders with ordered / received / remaining quantities and progress
app.get("/goods-received", async (req, res) => {
    try {
        const [rows] = await db.execute(`
            SELECT
                po.po_id,
                po.po_number,
                po.po_date,
                po.pr_number,
                po.pr_date,
                po.party_name,
                po.location,
                po.territory,
                po.product_category,
                po.item_name,
                po.product_remarks,
                po.make,
                po.model,
                po.qty AS ordered_quantity,
                po.unit,
                po.vendor_code,
                po.vendor_name,
                po.vendor_address,
                po.expected_delivery_date,
                po.status AS po_status,
                COALESCE(SUM(gr.received_quantity), 0) AS received_quantity,
                COALESCE((
                    SELECT SUM(ret.return_quantity)
                    FROM goods_returns ret
                    WHERE ret.po_id = po.po_id
                ), 0) AS returned_quantity
            FROM purchase_orders po
            LEFT JOIN goods_received gr ON gr.po_id = po.po_id
            WHERE po.status = 'ISSUED'
            GROUP BY
                po.po_id, po.po_number, po.po_date, po.pr_number, po.pr_date,
                po.party_name, po.location, po.territory, po.product_category,
                po.item_name, po.product_remarks, po.make, po.model, po.qty,
                po.unit, po.vendor_code, po.vendor_name, po.vendor_address,
                po.expected_delivery_date, po.status
            ORDER BY po.po_id DESC
        `);
        const result = rows.map(row => {
            const orderedQuantity   = Number(row.ordered_quantity);
            const totalReceived     = Number(row.received_quantity);
            const returnedQuantity  = Number(row.returned_quantity);
            const receivedQuantity  = Math.max(totalReceived - returnedQuantity, 0);
            const remainingQuantity = Math.max(orderedQuantity - receivedQuantity, 0);
            const progress = orderedQuantity > 0 ? Math.min((receivedQuantity / orderedQuantity) * 100, 100) : 0;
            return {
                ...row,
                ordered_quantity:   orderedQuantity,
                received_quantity:  receivedQuantity,
                remaining_quantity: remainingQuantity,
                progress: Number(progress.toFixed(2))
            };
        });
        return res.json({ success: true, orders: result });
    } catch (error) {
        console.error(error);
        log(`Goods Received fetch failed - ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to fetch goods received data"
        });
    }
});

// GET /goods-received/:po_id - Returns one purchase order with its full receipt and return history and stock figures
app.get("/goods-received/:po_id", async (req, res) => {
    const poId = Number(req.params.po_id);
    if (!Number.isInteger(poId) || poId <= 0)return res.status(400).json({ success: false, message: "Invalid PO ID" });
    try {
        const [poRows] = await db.execute(`
            SELECT
                po_id, po_number, po_date,
                vendor_code, vendor_name,
                item_name, make, model,
                qty AS ordered_quantity, unit,
                expected_delivery_date, status
            FROM purchase_orders
            WHERE po_id = ?
            LIMIT 1
        `, [poId]);
        if (poRows.length === 0)return res.status(404).json({ success: false, message: "Purchase Order not found" });
        const [receipts] = await db.execute(`
            SELECT receipt_id, received_date, received_quantity, remarks
            FROM goods_received
            WHERE po_id = ?
            ORDER BY received_date DESC, receipt_id DESC
        `, [poId]);
        const [returns] = await db.execute(`
            SELECT return_id, return_date, return_quantity, return_reason
            FROM goods_returns
            WHERE po_id = ?
            ORDER BY return_date DESC, return_id DESC
        `, [poId]);
        const po               = poRows[0];
        const orderedQuantity  = Number(po.ordered_quantity);
        const totalReceived    = receipts.reduce((sum, r) => sum + Number(r.received_quantity), 0);
        const totalReturned    = returns.reduce((sum, r)  => sum + Number(r.return_quantity), 0);
        const stockInHand      = Math.max(totalReceived - totalReturned, 0);
        const remainingQuantity = Math.max(orderedQuantity - totalReceived, 0);
        const progress         = orderedQuantity > 0 ? Math.min((totalReceived / orderedQuantity) * 100, 100) : 0;
        return res.json({
            success: true,
            order: {
                ...po,
                ordered_quantity:   orderedQuantity,
                received_quantity:  totalReceived,
                returned_quantity:  totalReturned,
                stock_in_hand:      stockInHand,
                remaining_quantity: remainingQuantity,
                progress: Number(progress.toFixed(2))
            },
            receipts,
            returns
        });
    } catch (error) {
        console.error(error);
        log(`Goods Received history fetch failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to fetch goods received history" });
    }
});

// POST /goods-received - Records a goods receipt against a purchase order (cannot exceed the remaining quantity)
app.post("/goods-received", async (req, res) => {
    const { po_id, received_quantity, receipt_date, remarks } = req.body;
    const poId        = Number(po_id);
    const receivedQty = Number(received_quantity);
    if (!Number.isInteger(poId) || poId <= 0)return res.status(400).json({ success: false, message: "Invalid PO ID" });
    if (!Number.isFinite(receivedQty) || receivedQty <= 0)return res.status(400).json({ success: false, message: "Received quantity must be greater than 0" });
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();
        const [poRows] = await connection.execute(`
            SELECT po_id, qty, unit, status, po_number, vendor_name, vendor_code, item_name, make, model
            FROM purchase_orders
            WHERE po_id = ?
            FOR UPDATE
        `, [poId]);
        if (poRows.length === 0) {
            await connection.rollback();
            return res.status(404).json({ success: false, message: "Purchase Order not found" });
        }
        const po = poRows[0];
        if (po.status === "CANCELLED") {
            await connection.rollback();
            return res.status(409).json({ success: false, message: "Cancelled Purchase Order cannot receive goods" });
        }
        if (po.status === "COMPLETED") {
            await connection.rollback();
            return res.status(409).json({ success: false, message: "Purchase Order is already completed" });
        }
        const orderedQuantity = Number(po.qty);
        const [receivedRows] = await connection.execute(`
            SELECT COALESCE(SUM(received_quantity), 0) AS total_received
            FROM goods_received
            WHERE po_id = ?
        `, [poId]);
        const [returnRows] = await connection.execute(`
            SELECT COALESCE(SUM(return_quantity), 0) AS total_returned
            FROM goods_returns
            WHERE po_id = ?
        `, [poId]);
        const alreadyReceived   = Number(receivedRows[0].total_received);
        const totalReturned     = Number(returnRows[0].total_returned);
        const remainingQuantity = orderedQuantity - alreadyReceived + totalReturned;
        if (receivedQty > remainingQuantity) {
            await connection.rollback();
            return res.status(400).json({
                success: false,
                message: `Cannot receive ${receivedQty}. Only ${remainingQuantity} ${po.unit || "units"} remaining.`
            });
        }
        const finalReceivedQuantity = alreadyReceived + receivedQty;
        const receivedDate = receipt_date && String(receipt_date).trim() !== "" ? receipt_date : new Date().toISOString().split("T")[0];
        const [result] = await connection.execute(`
            INSERT INTO goods_received (po_id, received_quantity, received_date, remarks)
            VALUES (?, ?, ?, ?)
        `, [poId, receivedQty, receivedDate, remarks?.trim() || null]);
        await connection.commit();
        log(`Goods received - PO: ${poId}, Received: ${receivedQty}`);
        await writeReportLog(req, "GOODS_RECEIVED",
            `Goods received against purchase order ${po.po_number} from vendor "${orDash(po.vendor_name)}" (${orDash(po.vendor_code)}) on ${dateText(receivedDate)} (receipt ID ${result.insertId}). ` +
            `Item: ${orDash(po.item_name)}, make: ${orDash(po.make)}, model: ${orDash(po.model)}. ` +
            `Quantity received now: ${receivedQty} ${orDash(po.unit)}. ` +
            `Total received so far: ${finalReceivedQuantity} of ${orderedQuantity} ${orDash(po.unit)} ordered, ` +
            `remaining: ${Math.max(orderedQuantity - finalReceivedQuantity, 0)} ${orDash(po.unit)}. ` +
            `Remarks: ${orDash(remarks)}.`
        );
        return res.status(201).json({
            success: true,
            message: "Goods received successfully",
            receipt_id:         result.insertId,
            received_quantity:  receivedQty,
            total_received:     finalReceivedQuantity,
            remaining_quantity: Math.max(orderedQuantity - finalReceivedQuantity, 0),
            progress: orderedQuantity > 0 ? Number(Math.min((finalReceivedQuantity / orderedQuantity) * 100, 100).toFixed(2)) : 0,
            po_status: po.status
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        log(`Goods Received creation failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to record goods received" });
    } finally {
        connection.release();
    }
});

// POST /goods-returns - Records goods returned to the vendor against a purchase order (cannot exceed the quantity in hand)
app.post("/goods-returns", async (req, res) => {
    const { po_id, receipt_id, return_quantity, return_reason, return_date } = req.body;
    const poId      = Number(po_id);
    const receiptId = receipt_id ? Number(receipt_id) : null;
    const returnQty = Number(return_quantity);
    if (!Number.isInteger(poId) || poId <= 0)return res.status(400).json({ success: false, message: "Invalid PO ID" });
    if (receiptId !== null && (!Number.isInteger(receiptId) || receiptId <= 0))return res.status(400).json({ success: false, message: "Invalid receipt ID" });
    if (!Number.isFinite(returnQty) || returnQty <= 0)return res.status(400).json({ success: false, message: "Return quantity must be greater than 0" });
    if (!return_reason || String(return_reason).trim() === "")return res.status(400).json({ success: false, message: "Return reason is required" });
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();
        const [poRows] = await connection.execute(`
            SELECT po_id, qty, status, unit, po_number, vendor_name, vendor_code, item_name, make, model
            FROM purchase_orders
            WHERE po_id = ?
            FOR UPDATE
        `, [poId]);
        if (poRows.length === 0) {
            await connection.rollback();
            return res.status(404).json({ success: false, message: "Purchase Order not found" });
        }
        const [receivedRows] = await connection.execute(`
            SELECT COALESCE(SUM(received_quantity), 0) AS total_received
            FROM goods_received
            WHERE po_id = ?
        `, [poId]);
        const totalReceived = Number(receivedRows[0].total_received);
        const [returnRows] = await connection.execute(`
            SELECT COALESCE(SUM(return_quantity), 0) AS total_returned
            FROM goods_returns
            WHERE po_id = ?
        `, [poId]);
        const totalReturned    = Number(returnRows[0].total_returned);
        const availableToReturn = totalReceived - totalReturned;
        if (returnQty > availableToReturn) {
            await connection.rollback();
            return res.status(400).json({
                success: false,
                message: `Cannot return ${returnQty}. Only ${availableToReturn} units available to return.`
            });
        }
        if (receiptId !== null) {
            const [receiptRows] = await connection.execute(`
                SELECT receipt_id
                FROM goods_received
                WHERE receipt_id = ? AND po_id = ?
                LIMIT 1
            `, [receiptId, poId]);
            if (receiptRows.length === 0) {
                await connection.rollback();
                return res.status(400).json({
                    success: false,
                    message: "Receipt does not belong to this Purchase Order"
                });
            }
        }
        const returnDate = return_date && String(return_date).trim() !== "" ? return_date : new Date().toISOString().split("T")[0];
        const [result] = await connection.execute(`
            INSERT INTO goods_returns (po_id, receipt_id, return_quantity, return_date, return_reason)
            VALUES (?, ?, ?, ?, ?)
        `, [poId, receiptId, returnQty, returnDate, String(return_reason).trim()]);
        await connection.commit();
        log(`Goods returned - PO: ${poId}, Return: ${returnQty}, Reason: ${String(return_reason).trim()}`);
        const newTotalReturned = totalReturned + returnQty;
        const po = poRows[0];
        await writeReportLog(req, "GOODS_RETURNED",
            `Goods returned against purchase order ${po.po_number} to vendor "${orDash(po.vendor_name)}" (${orDash(po.vendor_code)}) on ${dateText(returnDate)} (return ID ${result.insertId}` +
            `${receiptId !== null ? `, against receipt ID ${receiptId}` : ""}). ` +
            `Item: ${orDash(po.item_name)}, make: ${orDash(po.make)}, model: ${orDash(po.model)}. ` +
            `Quantity returned: ${returnQty} ${orDash(po.unit)}. Reason: ${String(return_reason).trim()}. ` +
            `Total received: ${totalReceived} ${orDash(po.unit)}, total returned: ${newTotalReturned} ${orDash(po.unit)}, ` +
            `stock in hand: ${Math.max(totalReceived - newTotalReturned, 0)} ${orDash(po.unit)}.`
        );
        return res.status(201).json({
            success: true,
            message: "Goods returned successfully",
            return_id:        result.insertId,
            return_quantity:  returnQty,
            total_received:   totalReceived,
            total_returned:   newTotalReturned,
            stock_in_hand:    Math.max(totalReceived - newTotalReturned, 0),
            available_to_return: availableToReturn - returnQty
        });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        log(`Goods Return creation failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to record goods return" });
    } finally {
        connection.release();
    }
});

/* ==========================================================================
   TAB: VENDOR PERFORMANCE
   ========================================================================== */

// GET /vendor-performance - Scorecard for every vendor: order count and value, average lead time, on-time delivery %, open orders
app.get("/vendor-performance", async (req, res) => {
    try {
        const [rows] = await db.execute(`
            SELECT
                vom.vendor_id,
                COALESCE(vom.vendor_code, '')          AS vendor_code,
                vom.vendor_name,
                vom.is_blacklisted,
                COUNT(DISTINCT po.po_id)               AS total_orders,
                COALESCE(SUM(po.total_price), 0)       AS order_value,
                ROUND(
                    AVG(
                        CASE
                            WHEN gr_first.first_received_date IS NOT NULL
                            THEN DATEDIFF(gr_first.first_received_date, po.po_date)
                        END
                    ), 1
                )                                       AS avg_lead_time_days,
                COUNT(
                    CASE
                        WHEN gr_first.first_received_date IS NOT NULL
                         AND po.expected_delivery_date IS NOT NULL
                         AND gr_first.first_received_date <= po.expected_delivery_date
                        THEN 1
                    END
                )                                       AS on_time_deliveries,
                COUNT(
                    CASE
                        WHEN gr_first.first_received_date IS NOT NULL
                        THEN 1
                    END
                )                                       AS total_deliveries,
                COUNT(
                    CASE
                        WHEN po.status IN ('DRAFT', 'ISSUED')
                        THEN 1
                    END
                )                                       AS open_orders
            FROM vendor_oem_masters vom
            LEFT JOIN purchase_orders po
                ON po.vendor_id = vom.vendor_id
            LEFT JOIN (
                SELECT po_id, MIN(received_date) AS first_received_date
                FROM goods_received
                GROUP BY po_id
            ) gr_first
                ON gr_first.po_id = po.po_id
            GROUP BY vom.vendor_id, vom.vendor_code, vom.vendor_name
            ORDER BY vom.vendor_name ASC
        `);
        const result = rows.map(row => {
            const onTime  = Number(row.on_time_deliveries);
            const total   = Number(row.total_deliveries);
            return {
                vendor_id:          row.vendor_id,
                vendor_code:        row.vendor_code,
                vendor_name:        row.vendor_name,
                is_blacklisted:     Boolean(row.is_blacklisted),
                total_orders:       Number(row.total_orders),
                order_value:        Number(row.order_value),
                avg_lead_time_days: row.avg_lead_time_days !== null ? Number(row.avg_lead_time_days) : null,
                on_time_deliveries: onTime,
                total_deliveries:   total,
                on_time_pct:        total > 0 ? Number(((onTime / total) * 100).toFixed(1)) : null,
                open_orders:        Number(row.open_orders)
            };
        });
        return res.json({ success: true, vendors: result });
    } catch (error) {
        console.error(error);
        log(`Vendor performance fetch failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to fetch vendor performance" });
    }
});

// GET /vendor-performance/:vendor_id - Detailed performance of one vendor: profile, KPIs and per-PO delivery breakdown
app.get("/vendor-performance/:vendor_id", async (req, res) => {
    const vendorId = Number(req.params.vendor_id);
    if (!Number.isInteger(vendorId) || vendorId <= 0)return res.status(400).json({ success: false, message: "Invalid vendor ID" });
    try {
        const [profileRows] = await db.execute(`
            SELECT
                vendor_id, vendor_code, vendor_name,
                gst_number, pan_number, msme_number,
                office_address, office_contact_name, office_contact_number,
                sales_team_name, sales_team_contact, sales_team_email,
                is_blacklisted
            FROM vendor_oem_masters
            WHERE vendor_id = ?
            LIMIT 1
        `, [vendorId]);
        if (profileRows.length === 0)return res.status(404).json({ success: false, message: "Vendor not found" });
        const [orderRows] = await db.execute(`
            SELECT
                po.po_id,
                po.po_number,
                po.po_date,
                po.item_name,
                po.make,
                po.model,
                po.qty,
                po.unit,
                po.price_per_unit,
                po.total_price,
                po.expected_delivery_date,
                po.status,
                gr_first.first_received_date,
                CASE
                    WHEN gr_first.first_received_date IS NOT NULL
                    THEN DATEDIFF(gr_first.first_received_date, po.po_date)
                END                                         AS lead_time_days,
                CASE
                    WHEN gr_first.first_received_date IS NOT NULL
                     AND po.expected_delivery_date IS NOT NULL
                     AND gr_first.first_received_date <= po.expected_delivery_date
                    THEN 1
                    ELSE 0
                END                                         AS delivered_on_time
            FROM purchase_orders po
            LEFT JOIN (
                SELECT po_id, MIN(received_date) AS first_received_date
                FROM goods_received
                GROUP BY po_id
            ) gr_first ON gr_first.po_id = po.po_id
            WHERE po.vendor_id = ?
            ORDER BY po.po_id DESC
        `, [vendorId]);
        const orders        = orderRows.map(r => ({
            po_id:                  r.po_id,
            po_number:              r.po_number,
            po_date:                r.po_date,
            item_name:              r.item_name,
            make:                   r.make,
            model:                  r.model,
            qty:                    Number(r.qty),
            unit:                   r.unit,
            price_per_unit:         Number(r.price_per_unit),
            total_price:            Number(r.total_price),
            expected_delivery_date: r.expected_delivery_date,
            first_received_date:    r.first_received_date,
            lead_time_days:         r.lead_time_days !== null ? Number(r.lead_time_days) : null,
            delivered_on_time:      Boolean(r.delivered_on_time),
            status:                 r.status
        }));
        const totalOrders       = orders.length;
        const orderValue        = orders.reduce((s, o) => s + o.total_price, 0);
        const openOrders        = orders.filter(o => ["DRAFT", "ISSUED"].includes(o.status)).length;
        const delivered         = orders.filter(o => o.lead_time_days !== null);
        const avgLeadTime       = delivered.length ? Number((delivered.reduce((s, o) => s + o.lead_time_days, 0) / delivered.length).toFixed(1)) : null;
        const onTimeCount       = delivered.filter(o => o.delivered_on_time).length;
        const onTimePct         = delivered.length ? Number(((onTimeCount / delivered.length) * 100).toFixed(1)) : null;
        return res.json({
            success: true,
            profile: profileRows[0],
            kpis: {
                total_orders:       totalOrders,
                order_value:        orderValue,
                avg_lead_time_days: avgLeadTime,
                on_time_deliveries: onTimeCount,
                total_deliveries:   delivered.length,
                on_time_pct:        onTimePct,
                open_orders:        openOrders
            },
            orders
        });
    } catch (error) {
        console.error(error);
        log(`Vendor performance detail fetch failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to fetch vendor performance details" });
    }
});

// POST /vendor-performance/:vendor_id/blacklist - Blacklists a vendor so it no longer appears in vendor dropdowns or quotations
app.post("/vendor-performance/:vendor_id/blacklist", async (req, res) => {
    const vendorId = Number(req.params.vendor_id);
    const { reason } = req.body;
    if (!Number.isInteger(vendorId) || vendorId <= 0)return res.status(400).json({ success: false, message: "Invalid vendor ID" });
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();
        const [rows] = await connection.execute(`
            SELECT vendor_id, vendor_code, vendor_name, is_blacklisted
            FROM vendor_oem_masters
            WHERE vendor_id = ?
            FOR UPDATE
        `, [vendorId]);
        if (rows.length === 0) {
            await connection.rollback();
            return res.status(404).json({ success: false, message: "Vendor not found" });
        }
        const vendor = rows[0];
        if (vendor.is_blacklisted) {
            await connection.rollback();
            return res.status(409).json({ success: false, message: "Vendor is already blacklisted" });
        }
        await connection.execute(`
            UPDATE vendor_oem_masters SET is_blacklisted = TRUE WHERE vendor_id = ?
        `, [vendorId]);
        await connection.commit();
        log(`Vendor blacklisted - Vendor ID: ${vendorId}`);
        await writeReportLog(req, "VENDOR_BLACKLISTED",
            `Vendor "${vendor.vendor_name}" (${orDash(vendor.vendor_code)}, vendor ID ${vendorId}) blacklisted. ` +
            `Reason: ${orDash(reason)}. Vendor will no longer appear in vendor dropdowns or be selectable for new quotations.`
        );
        return res.json({ success: true, message: "Vendor blacklisted successfully" });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        log(`Vendor blacklist failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to blacklist vendor" });
    } finally {
        connection.release();
    }
});
 
// POST /vendor-performance/:vendor_id/unblacklist - Removes a vendor from the blacklist, restoring it to normal use
app.post("/vendor-performance/:vendor_id/unblacklist", async (req, res) => {
    const vendorId = Number(req.params.vendor_id);
    const { reason } = req.body;
    if (!Number.isInteger(vendorId) || vendorId <= 0)return res.status(400).json({ success: false, message: "Invalid vendor ID" });
    const connection = await db.getConnection();
    try {
        await connection.beginTransaction();
        const [rows] = await connection.execute(`
            SELECT vendor_id, vendor_code, vendor_name, is_blacklisted
            FROM vendor_oem_masters
            WHERE vendor_id = ?
            FOR UPDATE
        `, [vendorId]);
        if (rows.length === 0) {
            await connection.rollback();
            return res.status(404).json({ success: false, message: "Vendor not found" });
        }
        const vendor = rows[0];
        if (!vendor.is_blacklisted) {
            await connection.rollback();
            return res.status(409).json({ success: false, message: "Vendor is not blacklisted" });
        }
        await connection.execute(`
            UPDATE vendor_oem_masters SET is_blacklisted = FALSE WHERE vendor_id = ?
        `, [vendorId]);
        await connection.commit();
        log(`Vendor unblacklisted - Vendor ID: ${vendorId}`);
        await writeReportLog(req, "VENDOR_UNBLACKLISTED",
            `Vendor "${vendor.vendor_name}" (${orDash(vendor.vendor_code)}, vendor ID ${vendorId}) removed from blacklist. ` +
            `Reason: ${orDash(reason)}. Vendor is available again in vendor dropdowns and for new quotations.`
        );
        return res.json({ success: true, message: "Vendor removed from blacklist successfully" });
    } catch (error) {
        await connection.rollback();
        console.error(error);
        log(`Vendor unblacklist failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to unblacklist vendor" });
    } finally {
        connection.release();
    }
});

/* ==========================================================================
   TAB: PAST PRICE REFERENCE
   ========================================================================== */

// GET /past-price-reference/models - Lists the distinct models ever purchased (with make, item and category) for the model picker
app.get("/past-price-reference/models", async (req, res) => {
    try {
        const [rows] = await db.execute(`
            SELECT DISTINCT
                model,
                make,
                item_name,
                product_category
            FROM purchase_orders
            WHERE model IS NOT NULL AND model != ''
            ORDER BY model ASC
        `);
        return res.json({ success: true, models: rows });
    } catch (error) {
        console.error(error);
        log(`Past price reference model list failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to fetch model list" });
    }
});

// GET /past-price-reference - Purchase price history of one model: all purchases, the last three, and purchases grouped by vendor
app.get("/past-price-reference", async (req, res) => {
    const model = cleanValue(req.query.model);
    if (!model)return res.status(400).json({ success: false, message: "Model name is required" });
    try {
        const [purchases] = await db.execute(`
            SELECT
                po.po_number,
                po.po_date,
                po.item_name,
                po.make,
                po.model,
                po.product_category,
                po.qty,
                po.unit,
                po.price_per_unit,
                po.total_price,
                po.vendor_code,
                po.vendor_name,
                po.status
            FROM purchase_orders po
            WHERE po.model = ?
            ORDER BY po.po_date DESC, po.po_id DESC
        `, [model]);
        if (!purchases.length)return res.json({ success: true, model, purchases: [], summary: [] });
        const last3 = purchases.slice(0, 3);
        const vendorMap = {};
        purchases.forEach(p => {
            const key = p.vendor_code || p.vendor_name;
            if (!vendorMap[key]) {
                vendorMap[key] = {
                    vendor_code: p.vendor_code,
                    vendor_name: p.vendor_name,
                    purchases:   []
                };
            }
            vendorMap[key].purchases.push(p);
        });
        const byVendor = Object.values(vendorMap);
        return res.json({
            success:   true,
            model,
            item_name: purchases[0].item_name,
            make:      purchases[0].make,
            purchases,
            last3,
            byVendor
        });
    } catch (error) {
        console.error(error);
        log(`Past price reference fetch failed - ${error.message}`);
        return res.status(500).json({ success: false, message: "Failed to fetch price history" });
    }
});

/* ==========================================================================
   TAB: MANAGEMENT INSIGHTS
   ========================================================================== */

// GET /management-insights - Dashboard analytics for a chosen period: PR/PO overview, financials, trends, top performers, payments, receipts and PR-to-PO efficiency
app.get("/management-insights", async (req, res) => {
    try {
        const { period = "current_month", from, to } = req.query;
        const range = resolveInsightsRange(period, from, to);
        if (range.error)return res.status(400).json({ success: false, message: range.error });
        const { start, end } = range;
        const params = start ? [start, end] : [];
        const inRange = (col) => (start ? `${col} BETWEEN ? AND ?` : "1=1");
        const FIN = "po.status IN ('ISSUED','COMPLETED')";
        const run = (sql) => db.execute(sql, params);
        const UNIT = "COALESCE(po.unit, '')";
        const [
            [financialRows],
            [prRows],
            [monthlyFinRows],
            [monthlyPrRows],
            [monthlyStatusRows],
            [statusRows],
            [topMakeRows],
            [topModelRows],
            [topVendorRows],
            [dailyFinRows],
            [paymentRows],
            [receiptUnitRows],
            [efficiencyRows]
        ] = await Promise.all([
            run(`
                SELECT
                    COALESCE(SUM(pr.taxable_value), 0) AS total_sales_value,
                    COALESCE(SUM(po.total_price), 0)   AS total_procurement_value,
                    COALESCE(SUM(pr.taxable_value - po.total_price), 0) AS gross_profit,
                    COALESCE(SUM(pr.taxable_value - po.total_price)
                             / NULLIF(SUM(pr.taxable_value), 0) * 100, 0) AS gross_margin_percentage,
                    COALESCE(AVG(po.total_price), 0) AS average_po_value,
                    COALESCE(MAX(po.total_price), 0) AS highest_po_value,
                    COALESCE(MAX(pr.taxable_value - po.total_price), 0) AS highest_gross_profit
                FROM purchase_orders po
                INNER JOIN purchase_requests pr ON pr.id = po.pr_id
                WHERE ${inRange("po.po_date")} AND ${FIN}
            `),
            run(`
                SELECT
                    COUNT(*) AS total_prs,
                    COALESCE(SUM(pr.taxable_value), 0) AS total_sales_value,
                    COALESCE(AVG(pr.taxable_value), 0) AS average_pr_value,
                    COALESCE(MAX(pr.taxable_value), 0) AS highest_pr_value
                FROM purchase_requests pr
                WHERE ${inRange("pr.pr_date")}
            `),
            run(`
                SELECT
                    DATE_FORMAT(po.po_date, '%Y-%m') AS month,
                    COALESCE(SUM(pr.taxable_value), 0) AS sales_value,
                    COALESCE(SUM(po.total_price), 0)   AS procurement_value,
                    COALESCE(SUM(pr.taxable_value - po.total_price), 0) AS gross_profit,
                    COALESCE(SUM(pr.taxable_value - po.total_price)
                             / NULLIF(SUM(pr.taxable_value), 0) * 100, 0) AS gross_margin_percentage
                FROM purchase_orders po
                INNER JOIN purchase_requests pr ON pr.id = po.pr_id
                WHERE ${inRange("po.po_date")} AND ${FIN}
                GROUP BY DATE_FORMAT(po.po_date, '%Y-%m')
            `),
            run(`
                SELECT DATE_FORMAT(pr.pr_date, '%Y-%m') AS month, COUNT(*) AS pr_count
                FROM purchase_requests pr
                WHERE pr.pr_date IS NOT NULL AND ${inRange("pr.pr_date")}
                GROUP BY DATE_FORMAT(pr.pr_date, '%Y-%m')
            `),
            run(`
                SELECT
                    DATE_FORMAT(po.po_date, '%Y-%m') AS month,
                    COUNT(*) AS total_pos,
                    SUM(po.status = 'COMPLETED') AS completed_pos,
                    SUM(po.status = 'ISSUED')    AS issued_pos,
                    SUM(po.status = 'DRAFT')     AS draft_pos,
                    SUM(po.status = 'CANCELLED') AS cancelled_pos
                FROM purchase_orders po
                WHERE ${inRange("po.po_date")}
                GROUP BY DATE_FORMAT(po.po_date, '%Y-%m')
            `),
            run(`
                SELECT po.status, COUNT(*) AS po_count,
                       COALESCE(SUM(po.total_price), 0) AS procurement_value
                FROM purchase_orders po
                WHERE ${inRange("po.po_date")}
                GROUP BY po.status
                ORDER BY po_count DESC
            `),
            run(`
                SELECT TRIM(po.make) AS make,
                       COUNT(*) AS po_count,
                       COALESCE(SUM(po.total_price), 0) AS procurement_value
                FROM purchase_orders po
                WHERE ${inRange("po.po_date")} AND ${FIN}
                  AND po.make IS NOT NULL AND TRIM(po.make) <> ''
                GROUP BY TRIM(po.make)
                ORDER BY procurement_value DESC, po_count DESC
                LIMIT 1
            `),
            run(`
                SELECT TRIM(po.model) AS model,
                       MAX(TRIM(po.make)) AS make,
                       COUNT(*) AS po_count,
                       COALESCE(SUM(po.total_price), 0) AS procurement_value
                FROM purchase_orders po
                WHERE ${inRange("po.po_date")} AND ${FIN}
                  AND po.model IS NOT NULL AND TRIM(po.model) <> ''
                GROUP BY TRIM(po.model)
                ORDER BY procurement_value DESC, po_count DESC
                LIMIT 1
            `),
            run(`
                SELECT po.vendor_id,
                       MAX(po.vendor_name) AS vendor_name,
                       MAX(po.vendor_code) AS vendor_code,
                       COUNT(*) AS po_count,
                       COALESCE(SUM(po.total_price), 0) AS procurement_value
                FROM purchase_orders po
                WHERE ${inRange("po.po_date")} AND ${FIN}
                GROUP BY po.vendor_id
                ORDER BY procurement_value DESC, po_count DESC
                LIMIT 1
            `),
            run(`
                SELECT
                    DATE_FORMAT(po.po_date, '%Y-%m-%d') AS day,
                    COALESCE(SUM(pr.taxable_value), 0) AS sales_value,
                    COALESCE(SUM(po.total_price), 0)   AS procurement_value,
                    COALESCE(SUM(pr.taxable_value - po.total_price), 0) AS gross_profit
                FROM purchase_orders po
                INNER JOIN purchase_requests pr ON pr.id = po.pr_id
                WHERE ${inRange("po.po_date")} AND ${FIN}
                GROUP BY DATE_FORMAT(po.po_date, '%Y-%m-%d')
            `),
            run(`
                SELECT COALESCE(po.payment_type, 'Not Specified') AS payment_type,
                       COUNT(*) AS po_count,
                       COALESCE(SUM(po.total_price), 0) AS procurement_value
                FROM purchase_orders po
                WHERE ${inRange("po.po_date")} AND ${FIN}
                GROUP BY COALESCE(po.payment_type, 'Not Specified')
                ORDER BY procurement_value DESC
            `),
            run(`
                SELECT ${UNIT} AS unit,
                       COUNT(*) AS receipts,
                       COALESCE(SUM(gr.received_quantity), 0) AS received
                FROM goods_received gr
                INNER JOIN purchase_orders po ON po.po_id = gr.po_id
                WHERE ${inRange("gr.received_date")}
                GROUP BY ${UNIT}
            `),
            run(`
                SELECT
                    AVG(DATEDIFF(po.po_date, pr.pr_date)) AS average_pr_to_po_days,
                    MIN(DATEDIFF(po.po_date, pr.pr_date)) AS fastest_pr_to_po_days,
                    MAX(DATEDIFF(po.po_date, pr.pr_date)) AS slowest_pr_to_po_days,
                    COUNT(*) AS sample_size
                FROM purchase_orders po
                INNER JOIN purchase_requests pr ON pr.id = po.pr_id
                WHERE ${inRange("po.po_date")}
                  AND po.status <> 'CANCELLED'
                  AND pr.pr_date IS NOT NULL
                  AND po.po_date >= pr.pr_date
            `)
        ]);
        const fin = financialRows[0] || {};
        const prs = prRows[0] || {};
        const eff = efficiencyRows[0] || {};
        const statusCount = (s) => num(statusRows.find(r => r.status === s)?.po_count);
        const totalPos = statusRows.reduce((sum, r) => sum + num(r.po_count), 0);
        const finMap = new Map(monthlyFinRows.map(r => [r.month, r]));
        const prMap = new Map(monthlyPrRows.map(r => [r.month, r]));
        const stMap = new Map(monthlyStatusRows.map(r => [r.month, r]));
        const months = monthsBetween([...finMap.keys(), ...prMap.keys(), ...stMap.keys()]);
        const monthlyProcurement = months.map(month => {
            const f = finMap.get(month);
            return {
                month,
                pr_count: num(prMap.get(month)?.pr_count),
                po_count: num(stMap.get(month)?.total_pos),
                sales_value: num(f?.sales_value),
                procurement_value: num(f?.procurement_value),
                gross_profit: num(f?.gross_profit),
                gross_margin_percentage: num(f?.gross_margin_percentage)
            };
        });
        const monthlyCompletion = months.map(month => {
            const s = stMap.get(month) || {};
            return {
                month,
                total_pos: num(s.total_pos),
                completed_pos: num(s.completed_pos),
                issued_pos: num(s.issued_pos),
                draft_pos: num(s.draft_pos),
                cancelled_pos: num(s.cancelled_pos)
            };
        });
        const dayMap = new Map(dailyFinRows.map(r => [r.day, r]));
        const dayKeys = [...dayMap.keys()].sort();
        const dayFrom = start || dayKeys[0];
        const dayTo = end || dayKeys[dayKeys.length - 1];
        const dailyProcurement = (dayFrom && dayTo) ? daysBetween(dayFrom, dayTo).map(day => {
            const r = dayMap.get(day);
            return {
                day,
                sales_value: num(r?.sales_value),
                procurement_value: num(r?.procurement_value),
                gross_profit: num(r?.gross_profit)
            };
        }) : [];
        const topMake = topMakeRows[0] ? { name: topMakeRows[0].make, po_count: num(topMakeRows[0].po_count), procurement_value: num(topMakeRows[0].procurement_value) } : null;
        const topModel = topModelRows[0] ? { name: topModelRows[0].model, make: topModelRows[0].make || "", po_count: num(topModelRows[0].po_count), procurement_value: num(topModelRows[0].procurement_value) } : null;
        const topVendor = topVendorRows[0] ? { name: topVendorRows[0].vendor_name, code: topVendorRows[0].vendor_code || "", po_count: num(topVendorRows[0].po_count), procurement_value: num(topVendorRows[0].procurement_value) } : null;
        const nullableNum = (v) => (v === null || v === undefined ? null : Number(v));
        return res.json({
            success: true,
            period: { type: period, from: start, to: end },
            overview: {
                total_prs: num(prs.total_prs),
                total_pos: totalPos,
                draft_pos: statusCount("DRAFT"),
                issued_pos: statusCount("ISSUED"),
                completed_pos: statusCount("COMPLETED"),
                cancelled_pos: statusCount("CANCELLED"),
                average_po_value: num(fin.average_po_value),
                highest_po_value: num(fin.highest_po_value)
            },
            financial: {
                total_sales_value: num(fin.total_sales_value),
                total_procurement_value: num(fin.total_procurement_value),
                gross_profit: num(fin.gross_profit),
                gross_margin_percentage: num(fin.gross_margin_percentage),
                highest_gross_profit: num(fin.highest_gross_profit)
            },
            purchase_requests: {
                total: num(prs.total_prs),
                total_sales_value: num(prs.total_sales_value),
                average_pr_value: num(prs.average_pr_value),
                highest_pr_value: num(prs.highest_pr_value)
            },
            trends: {
                daily_procurement: dailyProcurement,
                monthly_procurement: monthlyProcurement,
                monthly_completion: monthlyCompletion
            },
            top_performers: { make: topMake, model: topModel, vendor: topVendor },
            payment_types: paymentRows,
            po_status: statusRows,
            goods_received: {
                total_receipts: receiptUnitRows.reduce((s, r) => s + num(r.receipts), 0),
                by_unit: receiptUnitRows.map(r => ({
                    unit: r.unit,
                    received: num(r.received)
                }))
            },
            efficiency: {
                average_pr_to_po_days: nullableNum(eff.average_pr_to_po_days),
                fastest_pr_to_po_days: nullableNum(eff.fastest_pr_to_po_days),
                slowest_pr_to_po_days: nullableNum(eff.slowest_pr_to_po_days),
                sample_size: num(eff.sample_size)
            }
        });
    } catch (error) {
        console.error("Management Insights Error:", error);
        return res.status(500).json({
            success: false,
            message: "Failed to load management insights"
        });
    }
});

// GET /report-logs - Paginated business-activity history from report_logs, optionally filtered by username, action, and date range
app.get("/report-logs", async (req, res) => {
    const page  = Math.max(1, Number(req.query.page) || 1);
    const limit = Math.min(100, Math.max(1, Number(req.query.limit) || 25));
    const offset = (page - 1) * limit;
    const username = cleanValue(req.query.username);
    const action   = cleanValue(req.query.action);
    const from     = cleanValue(req.query.from);
    const to       = cleanValue(req.query.to);
    const conditions = [];
    const params = [];
    if (username) { conditions.push("username = ?"); params.push(username); }
    if (action)   { conditions.push("action = ?"); params.push(action); }
    if (from)     { conditions.push("log_timestamp >= ?"); params.push(`${from} 00:00:00`); }
    if (to)       { conditions.push("log_timestamp <= ?"); params.push(`${to} 23:59:59`); }
    const whereClause = conditions.length ? `WHERE ${conditions.join(" AND ")}` : "";
    try {
        const [[{ total }]] = await logDb.execute(
            `SELECT COUNT(*) AS total FROM report_logs ${whereClause}`,
            params
        );
        const totalPages = Math.max(1, Math.ceil(total / limit));
        const [rows] = await logDb.query(
            `SELECT report_log_id, log_timestamp, username, action, report
            FROM report_logs
            ${whereClause}
            ORDER BY log_timestamp DESC, report_log_id DESC
            LIMIT ${limit} OFFSET ${offset}`,
            params
        );
        return res.json({
            success: true,
            rows,
            pagination: {
                page,
                limit,
                total,
                total_pages: totalPages,
                has_prev: page > 1,
                has_next: page < totalPages
            }
        });
    } catch (error) {
        console.error(error);
        log(`Report log fetch failed - ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to fetch report logs"
        });
    }
});

// GET /audit-logs - Paginated administrative/system change history from audit_logs, optionally filtered by username, action, and date range
app.get("/audit-logs", async (req, res) => {
    const page  = Math.max(1, Number(req.query.page) || 1);
    const limit = Math.min(100, Math.max(1, Number(req.query.limit) || 25));
    const offset = (page - 1) * limit;
    const username = cleanValue(req.query.username);
    const action   = cleanValue(req.query.action);
    const from     = cleanValue(req.query.from);
    const to       = cleanValue(req.query.to);
    const conditions = [];
    const params = [];
    if (username) { conditions.push("username = ?"); params.push(username); }
    if (action)   { conditions.push("action = ?"); params.push(action); }
    if (from)     { conditions.push("log_timestamp >= ?"); params.push(`${from} 00:00:00`); }
    if (to)       { conditions.push("log_timestamp <= ?"); params.push(`${to} 23:59:59`); }
    const whereClause = conditions.length ? `WHERE ${conditions.join(" AND ")}` : "";
    try {
        const [[{ total }]] = await logDb.execute(
            `SELECT COUNT(*) AS total FROM audit_logs ${whereClause}`,
            params
        );
        const totalPages = Math.max(1, Math.ceil(total / limit));
        const [rows] = await logDb.query(
            `SELECT audit_log_id, log_timestamp, username, action, old_value, new_value
            FROM audit_logs
            ${whereClause}
            ORDER BY log_timestamp DESC, audit_log_id DESC
            LIMIT ${limit} OFFSET ${offset}`,
            params
        );
        return res.json({
            success: true,
            rows,
            pagination: {
                page,
                limit,
                total,
                total_pages: totalPages,
                has_prev: page > 1,
                has_next: page < totalPages
            }
        });
    } catch (error) {
        console.error(error);
        log(`Audit log fetch failed - ${error.message}`);
        return res.status(500).json({
            success: false,
            message: "Failed to fetch audit logs"
        });
    }
});

/* ==========================================================================
   SERVER STARTUP
   ========================================================================== */

   app.listen(PORT, async () => {
    log(`Starting procurement Manager service on port ${PORT}`);
    try {
        await db.query("SELECT 1");
        log("Database connection successful");
    } catch (error) {
        log(`Database connection failed: ${error.message}`);
    }
    try {
        await logDb.query("SELECT 1");
        log("Logging database connection successful");
    } catch (error) {
        log(`Logging database connection failed: ${error.message}`);
    }
    log(`Procurement Manager service running at http://localhost:${PORT}`);
});