CREATE DATABASE IF NOT EXISTS ProcureIQ_Logs;
USE ProcureIQ_Logs;
CREATE TABLE report_logs (
    report_log_id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    log_timestamp TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    username VARCHAR(100) NOT NULL,
    action VARCHAR(100) NOT NULL,
    report TEXT NOT NULL,
    PRIMARY KEY (report_log_id),
    KEY idx_report_timestamp (log_timestamp),
    KEY idx_report_username_timestamp (username, log_timestamp),
    KEY idx_report_action_timestamp (action, log_timestamp)
);
CREATE TABLE audit_logs (
    audit_log_id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT,
    log_timestamp TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    username VARCHAR(100) NOT NULL,
    action VARCHAR(100) NOT NULL,
    old_value TEXT DEFAULT NULL,
    new_value TEXT DEFAULT NULL,
    PRIMARY KEY (audit_log_id),
    KEY idx_audit_timestamp (log_timestamp),
    KEY idx_audit_username_timestamp (username, log_timestamp),
    KEY idx_audit_action_timestamp (action, log_timestamp)
);