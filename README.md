# Công Chứng Số

Nền tảng quản lý biểu mẫu và tự động hóa tài liệu cho văn phòng công chứng.

Dự án hiện gồm 3 thành phần chính:

- `frontend`: giao diện người dùng, xây dựng bằng React + Vite
- `backend`: API chính, xây dựng bằng Node.js + Express + MySQL
- `doc-converter-service`: microservice .NET 8 chạy trên Windows để chuyển file `.doc` sang `.docx` bằng Microsoft Word COM

## Tính năng chính

- Quản lý thư viện biểu mẫu theo danh mục dạng cây
- Tải lên nhiều biểu mẫu Word cùng lúc
- Hỗ trợ cả `.docx` và `.doc`
- Tự động quét placeholder dạng `{{ten_bien}}`
- Cấu hình field động cho từng biểu mẫu
- Liên kết biểu mẫu cha/con
- Điền form và xuất bộ hồ sơ `.docx`
- Batch convert các file `.doc` cũ sang `.docx`

## Kiến trúc

```text
Frontend (React/Vite)
        |
        v
Backend API (Node.js/Express)
        |
        +--> parse .docx, quét placeholder, lưu DB
        |
        +--> gọi Doc Converter Service khi upload file .doc
                     |
                     v
          Microsoft Word COM trên Windows
```

## Cấu trúc thư mục

```text
congchung/
├─ frontend/
├─ backend/
├─ doc-converter-service/
├─ database/
└─ template-linking.md
```

## Yêu cầu môi trường

### Bắt buộc

- Windows
- Node.js 18 trở lên
- MySQL
- .NET 8 SDK
- Microsoft Word hoặc Microsoft Office nếu muốn xử lý file `.doc`

### Lưu ý quan trọng

- File `.docx` vẫn có thể xử lý được khi `doc-converter-service` không chạy.
- File `.doc` chỉ xử lý được khi:
  - `doc-converter-service` đang chạy
  - máy có cài Microsoft Word

## Cài đặt

### 1. Backend

```bash
cd backend
npm install
```

Tạo file `.env` trong `backend/` nếu chưa có:

```env
PORT=5000
DB_HOST=127.0.0.1
DB_PORT=3306
DB_USER=root
DB_PASSWORD=your_password
DB_NAME=congchung

DOC_CONVERTER_URL=http://127.0.0.1:5051
DOC_CONVERTER_TIMEOUT_MS=20000
```

### 2. Frontend

```bash
cd frontend
npm install
```

### 3. Doc Converter Service

```bash
cd doc-converter-service
dotnet restore
```

## Khởi tạo database

Import schema:

```bash
mysql -u root -p congchung < database/schema.sql
```

Backend cũng có một số migration runtime để bổ sung cột nếu cần.

## Cách chạy dự án

Hiện tại dự án cần chạy riêng 3 tiến trình.

### Terminal 1: Doc Converter Service

```bash
cd doc-converter-service
dotnet run
```

Mặc định service chạy tại:

```text
http://127.0.0.1:5051
```

Kiểm tra nhanh:

```bash
curl http://127.0.0.1:5051/health
```

### Terminal 2: Backend

```bash
cd backend
npm run dev
```

Mặc định backend chạy tại:

```text
http://localhost:5000
```

### Terminal 3: Frontend

```bash
cd frontend
npm run dev
```

Mặc định frontend chạy tại:

```text
http://localhost:5173
```

## Luồng upload biểu mẫu

### Trường hợp `.docx`

1. Frontend gửi file lên backend
2. Backend lưu file tạm
3. Backend quét placeholder
4. Backend lưu metadata vào MySQL
5. Frontend mở màn hình chi tiết để cấu hình

### Trường hợp `.doc`

1. Frontend gửi file lên backend
2. Backend gửi file sang `doc-converter-service`
3. Service .NET mở Word ngầm và `Save As` sang `.docx`
4. Service trả file `.docx` về backend
5. Backend tiếp tục quét placeholder và lưu DB

## Multi-upload

Hệ thống hiện hỗ trợ:

- tối đa `10 file / lần`
- mỗi file tối đa `1MB`
- xử lý độc lập từng file
- file lỗi không làm hỏng cả batch

Kết quả trả về theo từng file:

- `success`
- `failed`
- thông báo lỗi cụ thể

## Script hữu ích

### Backend

```bash
npm run dev
npm start
npm run convert:doc -- "D:\source-doc-folder" "D:\output-docx-folder"
```

### Frontend

```bash
npm run dev
npm run build
```

### Converter Service

```bash
dotnet build
dotnet run
```

## Troubleshooting

### 1. File `.doc` báo lỗi không kết nối microservice

Nguyên nhân:

- `doc-converter-service` chưa chạy
- sai `DOC_CONVERTER_URL`

Kiểm tra:

```bash
curl http://127.0.0.1:5051/health
```

### 2. File `.doc` không convert được dù đã chạy service

Kiểm tra:

- Microsoft Word đã cài trên máy chưa
- Word có mở được file `.doc` thủ công không
- service có quyền tạo file trong thư mục `temp/` không

### 3. Upload thành công nhưng `0 biến`

Nguyên nhân thường gặp:

- file không chứa placeholder theo đúng mẫu `{{ten_bien}}`
- placeholder bị tách vỡ cấu trúc mà parser hiện tại không nhận diện được

### 4. Tên file tiếng Việt bị lỗi

Backend đã có bước chuẩn hóa tên file upload. Nếu vẫn gặp, cần kiểm tra tên file gốc trên máy người dùng và cách browser gửi multipart.

## Tài liệu liên quan

- `template-linking.md`
- `doc-converter-service/README.md`

## Hướng phát triển tiếp

- đóng gói script chạy cùng lúc frontend/backend/converter
- thêm progress tracking cho batch upload
- bổ sung health check tổng hợp cho cả 3 thành phần
- thêm deploy guide cho máy văn phòng
