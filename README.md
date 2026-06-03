# Cong Chung So

Nen tang quan ly bieu mau va tu dong hoa tai lieu cho van phong cong chung.

Du an gom 3 thanh phan:

- `frontend`: giao dien React + Vite
- `backend`: API Node.js + Express + MySQL
- `doc-converter-service`: microservice .NET 8 tren Windows de chuyen `.doc -> .docx` bang Microsoft Word COM

## Tinh nang chinh

- Quan ly thu vien bieu mau theo danh muc dang cay
- Upload nhieu bieu mau Word cung luc
- Ho tro ca `.docx` va `.doc`
- Tu dong quet placeholder `{{ten_bien}}`
- Cau hinh field dong cho tung bieu mau
- Lien ket bieu mau cha/con
- Dien form va xuat bo ho so `.docx`
- Batch convert cac file `.doc` cu sang `.docx`

## Kien truc

```text
Frontend (React/Vite)
        |
        v
Backend API (Node.js/Express)
        |
        +--> parse .docx, scan placeholders, luu DB
        |
        +--> goi Doc Converter Service khi upload file .doc
                     |
                     v
          Microsoft Word COM tren Windows
```

## Cau truc thu muc

```text
congchung/
|- frontend/
|- backend/
|- doc-converter-service/
|- database/
|- service-flow.html
`- template-linking.md
```

## Yeu cau moi truong

### Bat buoc

- Windows
- Node.js 18+ hoac moi hon
- MySQL
- .NET 8 SDK
- Microsoft Word / Microsoft Office da cai tren may neu muon xu ly file `.doc`

### Luu y quan trong

- File `.docx` co the xu ly duoc khi `doc-converter-service` khong chay.
- File `.doc` chi xu ly duoc khi:
  - `doc-converter-service` dang chay
  - may co Microsoft Word

## Cai dat

### 1. Backend

```bash
cd backend
npm install
```

Tao file `.env` trong `backend/` neu chua co:

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

## Khoi tao database

Import schema:

```bash
mysql -u root -p congchung < database/schema.sql
```

Backend cung co mot so migration runtime de bo sung cot neu can.

## Cach chay du an

Du an hien tai can chay rieng 3 tien trinh.

### Terminal 1: Doc Converter Service

```bash
cd doc-converter-service
dotnet run
```

Mac dinh service chay tai:

```text
http://127.0.0.1:5051
```

Kiem tra nhanh:

```bash
curl http://127.0.0.1:5051/health
```

### Terminal 2: Backend

```bash
cd backend
npm run dev
```

Mac dinh backend chay tai:

```text
http://localhost:5000
```

### Terminal 3: Frontend

```bash
cd frontend
npm run dev
```

Mac dinh frontend chay tai:

```text
http://localhost:5173
```

## Luong upload bieu mau

### Truong hop `.docx`

1. Frontend gui file len backend
2. Backend luu file tam
3. Backend quet placeholder
4. Backend luu metadata vao MySQL
5. Frontend mo man hinh chi tiet cau hinh

### Truong hop `.doc`

1. Frontend gui file len backend
2. Backend gui file sang `doc-converter-service`
3. Service .NET mo Word ngam va `Save As` sang `.docx`
4. Service tra file `.docx` ve backend
5. Backend tiep tuc quet placeholder va luu DB

## Multi-upload

He thong hien ho tro:

- toi da `10 file / lan`
- moi file toi da `1MB`
- xu ly doc lap tung file
- file loi khong lam hong ca batch

Ket qua tra ve theo tung file:

- `success`
- `failed`
- thong bao loi cu the

## Script huu ich

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

### 1. File `.doc` bao loi khong ket noi microservice

Nguyen nhan:

- `doc-converter-service` chua chay
- sai `DOC_CONVERTER_URL`

Kiem tra:

```bash
curl http://127.0.0.1:5051/health
```

### 2. File `.doc` khong convert duoc du da chay service

Kiem tra:

- Microsoft Word da cai tren may chua
- Word co mo duoc file `.doc` thu cong khong
- service co quyen tao file trong thu muc `temp/` khong

### 3. Upload thanh cong nhung `0 bien`

Nguyen nhan thuong gap:

- file khong chua placeholder theo dung mau `{{ten_bien}}`
- placeholder bi tach vo cau truc ma parser hien tai khong nhan dien duoc

### 4. Ten file tieng Viet bi loi

Backend da co buoc chuan hoa ten file upload. Neu van gap, can kiem tra ten file goc tren may nguoi dung va cach browser gui multipart.

## Tai lieu lien quan

- [service-flow.html](D:/MyProject/congchung/service-flow.html)
- [template-linking.md](D:/MyProject/congchung/template-linking.md)
- [doc-converter-service/README.md](D:/MyProject/congchung/doc-converter-service/README.md)

## Huong phat trien tiep

- dong goi script chay cung luc frontend/backend/converter
- them progress tracking cho batch upload
- bo sung health check tong hop cho ca 3 thanh phan
- them deploy guide cho may van phong
