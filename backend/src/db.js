const mysql = require('mysql2/promise');
require('dotenv').config();

const pool = mysql.createPool({
  host: process.env.DB_HOST || 'localhost',
  port: parseInt(process.env.DB_PORT || '3306'),
  user: process.env.DB_USER || 'root',
  password: process.env.DB_PASSWORD || '',
  database: process.env.DB_NAME || 'document_automation',
  waitForConnections: true,
  connectionLimit: 10,
  queueLimit: 0,
  // Tự động chuyển đổi các kiểu dữ liệu DATE/DATETIME thành chuỗi để tránh múi giờ bị lệch
  dateStrings: true
});

// Hàm kiểm tra kết nối CSDL khi khởi động ứng dụng
async function testConnection() {
  try {
    const connection = await pool.getConnection();
    console.log('✅ Kết nối database MySQL thành công!');
    connection.release();
  } catch (error) {
    console.error('❌ Kết nối database MySQL thất bại:', error.message);
    console.log('⚠️ Vui lòng cấu hình file backend/.env hoặc đảm bảo MySQL đang chạy.');
  }
}

module.exports = {
  pool,
  testConnection
};
