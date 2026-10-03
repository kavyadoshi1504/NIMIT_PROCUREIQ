CREATE DATABASE IF NOT EXISTS ProcureIQ;
USE ProcureIQ;
CREATE TABLE users (
    user_id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    username VARCHAR(100) NOT NULL,
    password_hash VARCHAR(255) NOT NULL,
    role ENUM('ADMIN', 'PROCUREMENT_MANAGER', 'PROCUREMENT') NOT NULL,
    is_active TINYINT(1) NOT NULL DEFAULT 1,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP,

    PRIMARY KEY (user_id),
    UNIQUE KEY username (username)
);
CREATE TABLE login_logs (
    login_log_id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    user_id BIGINT UNSIGNED NOT NULL,
    login_time TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    login_status ENUM('SUCCESS', 'FAILED') NOT NULL,

    PRIMARY KEY (login_log_id),
    KEY user_id (user_id),

    CONSTRAINT login_logs_ibfk_1purchase_orders
        FOREIGN KEY (user_id)
        REFERENCES users (user_id)
);
CREATE TABLE pr_sequences (
    financial_year VARCHAR(9) NOT NULL,
    last_number INT UNSIGNED NOT NULL DEFAULT 0,

    PRIMARY KEY (financial_year)
);
CREATE TABLE purchase_requests (
    id INT UNSIGNED NOT NULL AUTO_INCREMENT,
    pr_number VARCHAR(100) NOT NULL,
    pr_date DATE DEFAULT NULL,
    party_name VARCHAR(255) DEFAULT NULL,
    location VARCHAR(100) DEFAULT NULL,
    territory VARCHAR(100) DEFAULT NULL,
    product_category VARCHAR(150) DEFAULT NULL,
    item_name VARCHAR(255) DEFAULT NULL,
    product_remarks TEXT,
    make VARCHAR(100) DEFAULT NULL,
    model VARCHAR(150) DEFAULT NULL,
    qty DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    unit VARCHAR(50) DEFAULT NULL,
    sales_rate DECIMAL(15,2) NOT NULL DEFAULT 0.00,
    taxable_value DECIMAL(15,2) NOT NULL DEFAULT 0.00,
    created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP,

    PRIMARY KEY (id),
    UNIQUE KEY uq_pr_number (pr_number)
);
CREATE TABLE vendor_oem_masters (
    vendor_id INT UNSIGNED NOT NULL AUTO_INCREMENT,
    vendor_code VARCHAR(50) DEFAULT NULL,
    registration_date DATE NOT NULL,
    vendor_name VARCHAR(255) NOT NULL,
    legal_entity TEXT NOT NULL,
    commercial_role TEXT NOT NULL,
    year_of_incorporation VARCHAR(20) NOT NULL,

    director_or_ceo_or_management_name TEXT NOT NULL,
    director_or_ceo_or_management_designation TEXT NOT NULL,
    director_or_ceo_or_management_mobile_no TEXT NOT NULL,
    director_or_ceo_or_management_email TEXT NOT NULL,
    director_or_ceo_or_management_web_address TEXT,

    sales_team_name TEXT NOT NULL,
    sales_team_contact TEXT NOT NULL,
    sales_team_email TEXT NOT NULL,

    accounts_team_name TEXT NOT NULL,
    accounts_team_contact TEXT NOT NULL,
    accounts_team_email TEXT NOT NULL,

    gst_number TEXT NOT NULL,
    pan_number TEXT NOT NULL,
    msme_number TEXT,

    turnover_year_1 VARCHAR(20) NOT NULL,
    turnover_value_1 VARCHAR(100) NOT NULL,
    turnover_year_2 VARCHAR(20) DEFAULT NULL,
    turnover_value_2 VARCHAR(100) DEFAULT NULL,
    turnover_year_3 VARCHAR(20) DEFAULT NULL,
    turnover_value_3 VARCHAR(100) DEFAULT NULL,

    gst_document TEXT NOT NULL,
    pan_document TEXT NOT NULL,
    msme_document TEXT,
    itr_last_year_document TEXT NOT NULL,
    itr_second_last_year_document TEXT,
    itr_third_last_year_document TEXT,

    recommended_by TEXT NOT NULL,
    approved_by TEXT NOT NULL,
    is_blacklisted TINYINT(1) NOT NULL DEFAULT 0,

    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP,

    office_address TEXT NOT NULL,
    office_contact_name VARCHAR(150) NOT NULL,
    office_contact_number VARCHAR(30) NOT NULL,

    factory_address TEXT,
    factory_contact_name VARCHAR(150),
    factory_contact_number VARCHAR(30),

    warehouse_address TEXT,
    warehouse_contact_name VARCHAR(150),
    warehouse_contact_number VARCHAR(30),

    workshop_address TEXT,
    workshop_contact_name VARCHAR(150),
    workshop_contact_number VARCHAR(30),

    bank_name VARCHAR(150) NOT NULL,
    bank_account_no VARCHAR(50) NOT NULL,
    bank_branch VARCHAR(200) NOT NULL,
    bank_account_type VARCHAR(30) NOT NULL,
    bank_ifsc VARCHAR(20) NOT NULL,

    branch_office_1_address TEXT NOT NULL,
    branch_office_1_contact_name VARCHAR(150),
    branch_office_1_contact_number VARCHAR(30),

    branch_office_2_address TEXT,
    branch_office_2_contact_name VARCHAR(150),
    branch_office_2_contact_number VARCHAR(30),

    branch_office_3_address TEXT,
    branch_office_3_contact_name VARCHAR(150),
    branch_office_3_contact_number VARCHAR(30),

    client_details TEXT,

    PRIMARY KEY (vendor_id),
    UNIQUE KEY vendor_code (vendor_code)
);
CREATE TABLE vendor_inquiries (
    inquiry_id INT UNSIGNED NOT NULL AUTO_INCREMENT,
    pr_id INT UNSIGNED NOT NULL,
    status ENUM('OPEN', 'VENDOR_SELECTED', 'CLOSED', 'CANCELLED') NOT NULL DEFAULT 'OPEN',
    remarks TEXT,
    created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP,

    PRIMARY KEY (inquiry_id),
    UNIQUE KEY uq_vendor_inquiry_pr (pr_id),

    CONSTRAINT vendor_inquiries_ibfk_1
        FOREIGN KEY (pr_id)
        REFERENCES purchase_requests (id)
);
CREATE TABLE vendor_inquiry_vendors (
    inquiry_vendor_id INT UNSIGNED NOT NULL AUTO_INCREMENT,
    inquiry_id INT UNSIGNED NOT NULL,
    vendor_id INT UNSIGNED NOT NULL,

    price_per_unit DECIMAL(15,2) NOT NULL DEFAULT 0.00,
    total_price DECIMAL(15,2) NOT NULL DEFAULT 0.00,
    is_selected TINYINT(1) NOT NULL DEFAULT 0,
    remarks TEXT,

    created_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP,

    expected_delivery_date DATE DEFAULT NULL,

    payment_type ENUM(
        'ADVANCE',
        'CREDIT',
        'ADVANCE_PLUS_BALANCE',
        'CUSTOM'
    ) NOT NULL DEFAULT 'CREDIT',

    advance_type ENUM(
        'PERCENTAGE',
        'FIXED_AMOUNT'
    ) DEFAULT NULL,

    advance_value DECIMAL(15,2) DEFAULT NULL,
    advance_amount DECIMAL(15,2) DEFAULT NULL,
    balance_due_days SMALLINT UNSIGNED DEFAULT NULL,
    payment_terms_remarks TEXT,

    PRIMARY KEY (inquiry_vendor_id),

    UNIQUE KEY uq_inquiry_vendor (inquiry_id, vendor_id),

    KEY vendor_id (vendor_id),

    CONSTRAINT vendor_inquiry_vendors_ibfk_1
        FOREIGN KEY (inquiry_id)
        REFERENCES vendor_inquiries (inquiry_id),

    CONSTRAINT vendor_inquiry_vendors_ibfk_2
        FOREIGN KEY (vendor_id)
        REFERENCES vendor_oem_masters (vendor_id)
);
CREATE TABLE purchase_orders (
    po_id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    po_number VARCHAR(100) NOT NULL,
    po_date DATE NOT NULL,

    pr_id INT UNSIGNED NOT NULL,
    inquiry_id INT UNSIGNED NOT NULL,
    inquiry_vendor_id INT UNSIGNED NOT NULL,
    vendor_id INT UNSIGNED NOT NULL,

    pr_number VARCHAR(100) DEFAULT NULL,
    pr_date DATE DEFAULT NULL,
    party_name VARCHAR(255) DEFAULT NULL,
    location VARCHAR(100) DEFAULT NULL,
    territory VARCHAR(100) DEFAULT NULL,

    product_category VARCHAR(150) DEFAULT NULL,
    item_name VARCHAR(255) DEFAULT NULL,
    product_remarks TEXT,
    make VARCHAR(100) DEFAULT NULL,
    model VARCHAR(150) DEFAULT NULL,

    qty DECIMAL(12,2) NOT NULL DEFAULT 0.00,
    unit VARCHAR(50) DEFAULT NULL,
    sales_rate DECIMAL(15,2) NOT NULL DEFAULT 0.00,
    taxable_value DECIMAL(15,2) NOT NULL DEFAULT 0.00,

    vendor_code VARCHAR(50) DEFAULT NULL,
    vendor_name VARCHAR(255) NOT NULL,
    vendor_address TEXT,

    price_per_unit DECIMAL(15,2) NOT NULL,
    total_price DECIMAL(15,2) NOT NULL,

    expected_delivery_date DATE DEFAULT NULL,

    payment_type ENUM(
        'ADVANCE',
        'CREDIT',
        'ADVANCE_PLUS_BALANCE',
        'CUSTOM'
    ) NOT NULL,

    advance_type ENUM(
        'PERCENTAGE',
        'FIXED_AMOUNT'
    ) DEFAULT NULL,

    advance_value DECIMAL(15,2) DEFAULT NULL,
    balance_due_days SMALLINT UNSIGNED DEFAULT NULL,
    payment_terms_remarks TEXT,

    status ENUM(
        'DRAFT',
        'ISSUED',
        'CANCELLED',
        'COMPLETED'
    ) NOT NULL DEFAULT 'DRAFT',
    issued_po_path TEXT NULL,

    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
        ON UPDATE CURRENT_TIMESTAMP,

    PRIMARY KEY (po_id),

    UNIQUE KEY po_number (po_number),
    UNIQUE KEY uq_po_inquiry (inquiry_id),

    KEY pr_id (pr_id),
    KEY inquiry_vendor_id (inquiry_vendor_id),
    KEY vendor_id (vendor_id),

    CONSTRAINT purchase_orders_ibfk_1
        FOREIGN KEY (pr_id)
        REFERENCES purchase_requests (id),

    CONSTRAINT purchase_orders_ibfk_2
        FOREIGN KEY (inquiry_id)
        REFERENCES vendor_inquiries (inquiry_id),

    CONSTRAINT purchase_orders_ibfk_3
        FOREIGN KEY (inquiry_vendor_id)
        REFERENCES vendor_inquiry_vendors (inquiry_vendor_id),

    CONSTRAINT purchase_orders_ibfk_4
        FOREIGN KEY (vendor_id)
        REFERENCES vendor_oem_masters (vendor_id)
);
CREATE TABLE goods_received (
    receipt_id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,

    po_id BIGINT UNSIGNED NOT NULL,

    received_quantity DECIMAL(12,2) NOT NULL,
    defective_quantity DECIMAL(12,2) NOT NULL DEFAULT 0.00,

    received_date DATE NOT NULL,

    received_by BIGINT UNSIGNED NOT NULL,

    remarks TEXT,

    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,

    PRIMARY KEY (receipt_id),

    KEY po_id (po_id),
    KEY received_by (received_by),

    CONSTRAINT goods_received_ibfk_1
        FOREIGN KEY (po_id)
        REFERENCES purchase_orders (po_id),

    CONSTRAINT goods_received_ibfk_2
        FOREIGN KEY (received_by)
        REFERENCES users (user_id),

    CHECK (received_quantity > 0),
    CHECK (defective_quantity >= 0),
    CHECK (defective_quantity <= received_quantity)
);
CREATE TABLE goods_returns (
    return_id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,

    po_id BIGINT UNSIGNED NOT NULL,
    receipt_id BIGINT UNSIGNED DEFAULT NULL,

    return_quantity DECIMAL(12,2) NOT NULL,

    return_date DATE NOT NULL,

    return_reason TEXT NOT NULL,

    returned_by BIGINT UNSIGNED NOT NULL,

    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,

    PRIMARY KEY (return_id),

    KEY po_id (po_id),
    KEY receipt_id (receipt_id),
    KEY returned_by (returned_by),

    CONSTRAINT goods_returns_ibfk_1
        FOREIGN KEY (po_id)
        REFERENCES purchase_orders (po_id),

    CONSTRAINT goods_returns_ibfk_2
        FOREIGN KEY (receipt_id)
        REFERENCES goods_received (receipt_id),

    CONSTRAINT goods_returns_ibfk_3
        FOREIGN KEY (returned_by)
        REFERENCES users (user_id),

    CHECK (return_quantity > 0)
);
