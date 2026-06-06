const { v4: uuidv4 } = require('uuid');
const { pool } = require('../db');

async function getCategories(req, res) {
  try {
    const [rows] = await pool.query(`
      SELECT c.id, c.name, c.parent_id, c.sort_order, c.created_at,
             (SELECT COUNT(*) FROM template_categories cc WHERE cc.parent_id = c.id) AS children_count,
             (SELECT COUNT(*) FROM templates t WHERE t.category_id = c.id) AS templates_count
      FROM template_categories c
      ORDER BY c.parent_id IS NOT NULL, c.sort_order ASC, c.name ASC
    `);
    res.json(rows);
  } catch (error) {
    console.error('Lỗi API GET /api/template-categories:', error);
    res.status(500).json({ error: 'Không thể tải danh mục biểu mẫu.' });
  }
}

async function createCategory(req, res) {
  try {
    const name = (req.body.name || '').trim();
    const parentId = req.body.parentId || null;

    if (!name) {
      return res.status(400).json({ error: 'Tên danh mục không được để trống.' });
    }

    if (parentId) {
      const [parents] = await pool.query('SELECT id FROM template_categories WHERE id = ?', [parentId]);
      if (parents.length === 0) {
        return res.status(400).json({ error: 'Danh mục cha không tồn tại.' });
      }
    }

    const [existing] = await pool.query(
      'SELECT id FROM template_categories WHERE name = ? AND ((parent_id IS NULL AND ? IS NULL) OR parent_id = ?)',
      [name, parentId, parentId]
    );

    if (existing.length > 0) {
      return res.status(409).json({ error: 'Danh mục này đã tồn tại trong cùng cấp.' });
    }

    const [[{ nextSortOrder }]] = await pool.query(
      'SELECT COALESCE(MAX(sort_order), -1) + 1 AS nextSortOrder FROM template_categories WHERE ((parent_id IS NULL AND ? IS NULL) OR parent_id = ?)',
      [parentId, parentId]
    );

    const category = {
      id: uuidv4(),
      name,
      parent_id: parentId,
      sort_order: nextSortOrder
    };

    await pool.query(
      'INSERT INTO template_categories (id, name, parent_id, sort_order) VALUES (?, ?, ?, ?)',
      [category.id, category.name, category.parent_id, category.sort_order]
    );

    res.status(201).json(category);
  } catch (error) {
    console.error('Lỗi API POST /api/template-categories:', error);
    res.status(500).json({ error: 'Không thể tạo danh mục biểu mẫu.' });
  }
}

module.exports = {
  getCategories,
  createCategory
};
