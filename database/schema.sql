CREATE DATABASE IF NOT EXISTS document_automation CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci;
USE document_automation;

-- 1. Bảng văn phòng công chứng
CREATE TABLE IF NOT EXISTS notary_offices (
    id CHAR(36) PRIMARY KEY,
    name VARCHAR(255) NOT NULL,
    email VARCHAR(191) NOT NULL UNIQUE,
    password_hash VARCHAR(255) NOT NULL,
    phone VARCHAR(20) DEFAULT NULL,
    status ENUM('active', 'inactive') DEFAULT 'active',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP
) ENGINE=InnoDB;

-- 2. Bảng quản lý template docx gốc
CREATE TABLE IF NOT EXISTS template_categories (
    id CHAR(36) PRIMARY KEY,
    name VARCHAR(255) NOT NULL,
    parent_id CHAR(36) DEFAULT NULL,
    sort_order INT DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (parent_id) REFERENCES template_categories(id) ON DELETE SET NULL
) ENGINE=InnoDB;

-- 3. Bảng quản lý template docx gốc
CREATE TABLE IF NOT EXISTS templates (
    id CHAR(36) PRIMARY KEY,
    office_id CHAR(36) NOT NULL,
    category_id CHAR(36) DEFAULT NULL,
    name VARCHAR(255) NOT NULL,
    file_path VARCHAR(500) NOT NULL,
    status ENUM('active', 'draft') DEFAULT 'draft',
    parent_template_id CHAR(36) DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (office_id) REFERENCES notary_offices(id) ON DELETE CASCADE,
    FOREIGN KEY (category_id) REFERENCES template_categories(id) ON DELETE SET NULL,
    FOREIGN KEY (parent_template_id) REFERENCES templates(id) ON DELETE SET NULL
) ENGINE=InnoDB;

-- 4. Bảng quản lý các biến động quét được từ file Word
CREATE TABLE IF NOT EXISTS template_fields (
    id CHAR(36) PRIMARY KEY,
    template_id CHAR(36) NOT NULL,
    key_name VARCHAR(100) NOT NULL,
    field_type ENUM('text', 'date', 'number', 'boolean') DEFAULT 'text',
    label VARCHAR(255) NOT NULL,
    is_required TINYINT(1) DEFAULT 1,
    order_index INT DEFAULT 0,
    parent_field_key VARCHAR(100) DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE CASCADE,
    UNIQUE KEY uk_template_key (template_id, key_name)
) ENGINE=InnoDB;

-- 5. Bảng lưu trữ hồ sơ nộp và dữ liệu JSON phẳng
CREATE TABLE IF NOT EXISTS document_submissions (
    id CHAR(36) PRIMARY KEY,
    template_id CHAR(36) NOT NULL,
    customer_name VARCHAR(255) DEFAULT NULL,
    customer_phone VARCHAR(20) DEFAULT NULL,
    status ENUM('pending', 'completed', 'failed') DEFAULT 'pending',
    values_json JSON NOT NULL,
    output_file_path VARCHAR(500) DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    completed_at DATETIME DEFAULT NULL,
    FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE CASCADE
) ENGINE=InnoDB;
