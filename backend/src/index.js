const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');
require('dotenv').config();

const { pool, testConnection } = require('./db');
const categoriesRouter = require('./routes/categories');
const templatesRouter = require('./routes/templates');
const submissionsRouter = require('./routes/submissions');

const DEFAULT_TEMPLATE_CATEGORIES = [
  {
    name: 'Hợp đồng',
    children: ['Chuyển nhượng', 'Tặng cho', 'Thế chấp', 'Ủy quyền']
  },
  {
    name: 'Đất đai',
    children: ['Sang tên', 'Biến động', 'Đăng ký']
  },
  {
    name: 'Thuế',
    children: ['TNCN', 'Lệ phí trước bạ', 'Phi nông nghiệp']
  },
  {
    name: 'Biểu mẫu nội bộ',
    children: []
  }
];

const DEFAULT_OFFICE_ID = 'd3b07384-d113-4ec6-a5d6-c0c2a05d2ed1';

const app = express();
const PORT = process.env.PORT || 5000;

// Middleware
app.use(cors());
app.use(express.json());

// Ensure upload directories exist
const UPLOADS_DIR = path.join(__dirname, '..', 'uploads');
const TEMPLATES_DIR = path.join(UPLOADS_DIR, 'templates');
const OUTPUTS_DIR = path.join(UPLOADS_DIR, 'outputs');

[UPLOADS_DIR, TEMPLATES_DIR, OUTPUTS_DIR].forEach(dir => {
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
});

// Seed data Văn phòng công chứng mặc định
async function seedDefaultOffice() {
  try {
    const [rows] = await pool.query('SELECT id FROM notary_offices WHERE id = ?', [DEFAULT_OFFICE_ID]);
    if (rows.length === 0) {
      await pool.query(
        `INSERT INTO notary_offices (id, name, email, password_hash, phone, status) 
         VALUES (?, ?, ?, ?, ?, ?)`,
        [
          DEFAULT_OFFICE_ID,
          'Văn phòng Công chứng Trung tâm',
          'admin@congchungso.vn',
          '$2b$10$NotaryOfficeDefaultPasswordHashHere', // password giả định
          '0901234567',
          'active'
        ]
      );
      console.log('🌱 Đã tự động tạo văn phòng công chứng mặc định.');
    }
  } catch (error) {
    console.error('Không thể tạo dữ liệu văn phòng công chứng mặc định:', error.message);
  }
}

async function seedDefaultTemplateCategories() {
  try {
    const [rows] = await pool.query('SELECT COUNT(*) AS cnt FROM template_categories');
    if (rows[0].cnt > 0) {
      return;
    }

    for (let i = 0; i < DEFAULT_TEMPLATE_CATEGORIES.length; i++) {
      const root = DEFAULT_TEMPLATE_CATEGORIES[i];
      const rootId = uuidv4();

      await pool.query(
        'INSERT INTO template_categories (id, name, parent_id, sort_order) VALUES (?, ?, ?, ?)',
        [rootId, root.name, null, i]
      );

      for (let j = 0; j < root.children.length; j++) {
        await pool.query(
          'INSERT INTO template_categories (id, name, parent_id, sort_order) VALUES (?, ?, ?, ?)',
          [uuidv4(), root.children[j], rootId, j]
        );
      }
    }

    console.log('✅ Đã tạo cây danh mục biểu mẫu mặc định.');
  } catch (error) {
    console.error('Không thể tạo danh mục biểu mẫu mặc định:', error.message);
  }
}

// API Routes
app.use('/api/categories', categoriesRouter);
app.use('/api/templates', templatesRouter);
app.use('/api/submissions', submissionsRouter);

// Database migration logic
async function migrateDatabase() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS notary_offices (
        id CHAR(36) PRIMARY KEY,
        name VARCHAR(255) NOT NULL,
        email VARCHAR(191) NOT NULL UNIQUE,
        password_hash VARCHAR(255) NOT NULL,
        phone VARCHAR(20),
        status VARCHAR(20) NOT NULL DEFAULT 'active',
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS template_categories (
        id CHAR(36) PRIMARY KEY,
        name VARCHAR(255) NOT NULL,
        parent_id CHAR(36),
        sort_order INTEGER DEFAULT 0,
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        CONSTRAINT fk_template_categories_parent
          FOREIGN KEY (parent_id) REFERENCES template_categories(id) ON DELETE SET NULL
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS templates (
        id CHAR(36) PRIMARY KEY,
        office_id CHAR(36) NOT NULL,
        category_id CHAR(36),
        name VARCHAR(255) NOT NULL,
        file_path VARCHAR(500) NOT NULL,
        status VARCHAR(20) NOT NULL DEFAULT 'draft',
        parent_template_id CHAR(36),
        is_repeated BOOLEAN NOT NULL DEFAULT FALSE,
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        CONSTRAINT fk_templates_office
          FOREIGN KEY (office_id) REFERENCES notary_offices(id) ON DELETE CASCADE,
        CONSTRAINT fk_templates_category
          FOREIGN KEY (category_id) REFERENCES template_categories(id) ON DELETE SET NULL,
        CONSTRAINT fk_templates_parent
          FOREIGN KEY (parent_template_id) REFERENCES templates(id) ON DELETE SET NULL
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS template_fields (
        id CHAR(36) PRIMARY KEY,
        template_id CHAR(36) NOT NULL,
        key_name VARCHAR(100) NOT NULL,
        field_type VARCHAR(20) NOT NULL DEFAULT 'text',
        label VARCHAR(255) NOT NULL,
        is_required BOOLEAN NOT NULL DEFAULT TRUE,
        order_index INTEGER DEFAULT 0,
        replace_text VARCHAR(500),
        paragraph_context TEXT,
        parent_field_key VARCHAR(100),
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        CONSTRAINT fk_template_fields_template
          FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE CASCADE,
        CONSTRAINT uk_template_fields_template_key UNIQUE (template_id, key_name)
      )
    `);

    await pool.query(`
      CREATE TABLE IF NOT EXISTS document_submissions (
        id CHAR(36) PRIMARY KEY,
        template_id CHAR(36) NOT NULL,
        customer_name VARCHAR(255),
        customer_phone VARCHAR(20),
        status VARCHAR(20) NOT NULL DEFAULT 'pending',
        values_json JSONB NOT NULL,
        output_file_path VARCHAR(500),
        created_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
        completed_at TIMESTAMPTZ,
        CONSTRAINT fk_document_submissions_template
          FOREIGN KEY (template_id) REFERENCES templates(id) ON DELETE CASCADE
      )
    `);

    const [templatesColumns] = await pool.query(`
      SELECT column_name
      FROM information_schema.columns
      WHERE table_schema = 'public' AND table_name = 'templates'
    `);
    const templatesColumnNames = templatesColumns.map(c => c.column_name);

    if (!templatesColumnNames.includes('parent_template_id')) {
      await pool.query('ALTER TABLE templates ADD COLUMN parent_template_id CHAR(36)');
      await pool.query(`
        ALTER TABLE templates
        ADD CONSTRAINT fk_templates_parent
        FOREIGN KEY (parent_template_id) REFERENCES templates(id) ON DELETE SET NULL
      `);
    }

    if (!templatesColumnNames.includes('is_repeated')) {
      await pool.query('ALTER TABLE templates ADD COLUMN is_repeated BOOLEAN NOT NULL DEFAULT FALSE');
    }

    if (!templatesColumnNames.includes('category_id')) {
      await pool.query('ALTER TABLE templates ADD COLUMN category_id CHAR(36)');
      await pool.query(`
        ALTER TABLE templates
        ADD CONSTRAINT fk_templates_category
        FOREIGN KEY (category_id) REFERENCES template_categories(id) ON DELETE SET NULL
      `);
    }

    const [fieldColumns] = await pool.query(`
      SELECT column_name
      FROM information_schema.columns
      WHERE table_schema = 'public' AND table_name = 'template_fields'
    `);
    const fieldColumnNames = fieldColumns.map(c => c.column_name);

    if (!fieldColumnNames.includes('replace_text')) {
      await pool.query('ALTER TABLE template_fields ADD COLUMN replace_text VARCHAR(500)');
    }

    if (!fieldColumnNames.includes('paragraph_context')) {
      await pool.query('ALTER TABLE template_fields ADD COLUMN paragraph_context TEXT');
    }

    if (!fieldColumnNames.includes('parent_field_key')) {
      await pool.query('ALTER TABLE template_fields ADD COLUMN parent_field_key VARCHAR(100)');
    }

    console.log('✅ Database migration hoàn tất trên PostgreSQL.');
  } catch (error) {
    console.error('❌ Lỗi khi chạy migration database:', error.message);
  }
}

// Khởi chạy server
async function startServer() {
  await testConnection();
  await migrateDatabase();
  await seedDefaultOffice();
  await seedDefaultTemplateCategories();
  
  app.listen(PORT, () => {
    console.log(`🚀 Backend server running at http://localhost:${PORT}`);
  });
}

startServer();
