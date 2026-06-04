-- PostgreSQL không dùng CREATE DATABASE IF NOT EXISTS trong cùng script app
-- Hãy tạo database trên Supabase/Neon trước, rồi chạy phần dưới

-- 1. Enum types
DO $$ BEGIN
    CREATE TYPE office_status AS ENUM ('active', 'inactive');
EXCEPTION
    WHEN duplicate_object THEN null;
END $$;

DO $$ BEGIN
    CREATE TYPE template_status AS ENUM ('active', 'draft');
EXCEPTION
    WHEN duplicate_object THEN null;
END $$;

DO $$ BEGIN
    CREATE TYPE field_type_enum AS ENUM ('text', 'date', 'number', 'boolean');
EXCEPTION
    WHEN duplicate_object THEN null;
END $$;

DO $$ BEGIN
    CREATE TYPE submission_status AS ENUM ('pending', 'completed', 'failed');
EXCEPTION
    WHEN duplicate_object THEN null;
END $$;


-- 2. Bảng văn phòng công chứng
CREATE TABLE IF NOT EXISTS notary_offices (
    id UUID PRIMARY KEY,
    name VARCHAR(255) NOT NULL,
    email VARCHAR(191) NOT NULL UNIQUE,
    password_hash VARCHAR(255) NOT NULL,
    phone VARCHAR(20) DEFAULT NULL,
    status office_status DEFAULT 'active',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);


-- 3. Bảng danh mục template
CREATE TABLE IF NOT EXISTS template_categories (
    id UUID PRIMARY KEY,
    name VARCHAR(255) NOT NULL,
    parent_id UUID DEFAULT NULL,
    sort_order INT DEFAULT 0,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT fk_template_categories_parent
        FOREIGN KEY (parent_id)
        REFERENCES template_categories(id)
        ON DELETE SET NULL
);


-- 4. Bảng quản lý template docx gốc
CREATE TABLE IF NOT EXISTS templates (
    id UUID PRIMARY KEY,
    office_id UUID NOT NULL,
    category_id UUID DEFAULT NULL,
    name VARCHAR(255) NOT NULL,
    file_path VARCHAR(500) NOT NULL,
    status template_status DEFAULT 'draft',
    parent_template_id UUID DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT fk_templates_office
        FOREIGN KEY (office_id)
        REFERENCES notary_offices(id)
        ON DELETE CASCADE,

    CONSTRAINT fk_templates_category
        FOREIGN KEY (category_id)
        REFERENCES template_categories(id)
        ON DELETE SET NULL,

    CONSTRAINT fk_templates_parent_template
        FOREIGN KEY (parent_template_id)
        REFERENCES templates(id)
        ON DELETE SET NULL
);


-- 5. Bảng quản lý các biến động quét được từ file Word
CREATE TABLE IF NOT EXISTS template_fields (
    id UUID PRIMARY KEY,
    template_id UUID NOT NULL,
    key_name VARCHAR(100) NOT NULL,
    field_type field_type_enum DEFAULT 'text',
    label VARCHAR(255) NOT NULL,
    is_required BOOLEAN DEFAULT TRUE,
    order_index INT DEFAULT 0,
    parent_field_key VARCHAR(100) DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT fk_template_fields_template
        FOREIGN KEY (template_id)
REFERENCES templates(id)
        ON DELETE CASCADE,

    CONSTRAINT uk_template_key
        UNIQUE (template_id, key_name)
);


-- 6. Bảng lưu trữ hồ sơ nộp và dữ liệu JSON phẳng
CREATE TABLE IF NOT EXISTS document_submissions (
    id UUID PRIMARY KEY,
    template_id UUID NOT NULL,
    customer_name VARCHAR(255) DEFAULT NULL,
    customer_phone VARCHAR(20) DEFAULT NULL,
    status submission_status DEFAULT 'pending',
    values_json JSONB NOT NULL,
    output_file_path VARCHAR(500) DEFAULT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    completed_at TIMESTAMP DEFAULT NULL,

    CONSTRAINT fk_document_submissions_template
        FOREIGN KEY (template_id)
        REFERENCES templates(id)
        ON DELETE CASCADE
);