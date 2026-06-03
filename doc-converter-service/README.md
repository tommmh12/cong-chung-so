# Doc Converter Service

Microservice `.NET 8` chạy trên Windows để chuyển file `.doc` sang `.docx` bằng `Microsoft.Office.Interop.Word`.

## Yêu cầu môi trường

- Windows
- Microsoft Word đã cài trên máy
- .NET 8 SDK hoặc Runtime

## NuGet

Project dùng gói:

- `Microsoft.Office.Interop.Word`

Nếu cần cài thủ công:

```powershell
dotnet add package Microsoft.Office.Interop.Word
```

## Chạy service

```powershell
cd doc-converter-service
dotnet restore
dotnet run
```

Mặc định service chạy tại:

```text
http://127.0.0.1:5051
```

## API

### `GET /health`

Kiểm tra service còn sống.

### `POST /api/convert-doc`

- Content-Type: `multipart/form-data`
- Field file: `file`
- Chỉ nhận file `.doc`
- Trả về file `.docx` trực tiếp trong response

Ví dụ bằng PowerShell:

```powershell
curl.exe -X POST "http://127.0.0.1:5051/api/convert-doc" `
  -F "file=@C:\docs\hop-dong.doc"
```

## Cấu hình

Trong `appsettings.json`:

```json
{
  "Converter": {
    "TempRoot": "temp",
    "TimeoutSeconds": 15
  }
}
```

## Tích hợp với backend Node

Backend hiện tại gọi service này qua biến môi trường:

- `DOC_CONVERTER_URL` mặc định `http://127.0.0.1:5051`
- `DOC_CONVERTER_TIMEOUT_MS` mặc định `20000`

Ví dụ:

```powershell
$env:DOC_CONVERTER_URL="http://127.0.0.1:5051"
$env:DOC_CONVERTER_TIMEOUT_MS="20000"
```
