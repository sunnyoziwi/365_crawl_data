# 365 Crawl Data

Script Python dùng Claude API (Anthropic) để tự động đọc file hợp đồng khách sạn (PDF, DOCX, DOC, TXT) và trích xuất thông tin có cấu trúc (tên khách sạn, loại phòng, bảng giá theo mùa, chính sách...) ra file JSON.

## Cấu trúc thư mục

```
365_crawl_data/
├── crawl.py            # Script chính
├── requirements.txt    # Danh sách thư viện cần cài
├── data/                # Dữ liệu đầu vào (hợp đồng gốc) - KHÔNG đẩy lên Git
│   └── HOTEL/
│       └── HANOI/
│           └── <Tên khách sạn>/
│               └── ... (file .pdf, .docx, .doc, .txt)
└── data_result/         # Dữ liệu đầu ra (JSON đã trích xuất) - KHÔNG đẩy lên Git
```

> Thư mục `data/` và `data_result/` không được đưa lên GitHub (xem `.gitignore`) vì chứa dữ liệu hợp đồng/giá cả nội bộ và dung lượng lớn. Bạn tự chuẩn bị dữ liệu này ở máy chạy script.

## Yêu cầu

- Python 3.10 trở lên
- Một API key của Anthropic ([console.anthropic.com](https://console.anthropic.com))

## Cài đặt

### 1. Tạo môi trường ảo (venv)

```powershell
python -m venv .venv
.\.venv\Scripts\activate
```

### 2. Cài thư viện

```powershell
pip install -r requirements.txt
```

## Cấu hình API key

Script đọc API key từ biến môi trường `ANTHROPIC_API_KEY` (không hard-code key trong code). Đặt biến này trước khi chạy:

```powershell
$env:ANTHROPIC_API_KEY = "sk-ant-..."
```

> Lưu ý: mỗi lần mở cửa sổ PowerShell mới bạn cần đặt lại biến này, trừ khi bạn set nó ở cấp hệ thống (System Environment Variables).

## Cách chạy

```powershell
python crawl.py
```

Script sẽ:
1. Duyệt qua từng thư mục khách sạn trong `SOURCE_DIR` (mặc định `data/HANOI`)
2. Đọc toàn bộ file `.pdf`, `.docx`, `.doc`, `.txt` bên trong (kể cả thư mục con)
3. Gửi nội dung cho Claude để trích xuất thông tin theo schema định sẵn
4. Ghi kết quả JSON tương ứng vào `OUTPUT_BASE_DIR` (mặc định `data_result/HANOI`), giữ nguyên cấu trúc thư mục

Nếu file JSON kết quả đã tồn tại, script sẽ bỏ qua (không xử lý lại) để tiết kiệm chi phí gọi API.

## Các tuỳ chọn cấu hình (đầu file `crawl.py`)

| Biến | Ý nghĩa | Mặc định |
|---|---|---|
| `SOURCE_DIR` | Thư mục chứa dữ liệu gốc cần xử lý | `data/HANOI` |
| `OUTPUT_BASE_DIR` | Thư mục ghi kết quả JSON | `data_result/HANOI` |
| `LIMIT_HOTELS` | Giới hạn số khách sạn xử lý (để test). Đặt `None` để chạy toàn bộ | `5` |
| `MODEL_NAME` | Model Claude sử dụng | `claude-haiku-4-5` |

## Lưu ý quan trọng

- **Đường dẫn dữ liệu thực tế**: dữ liệu mẫu trong `data/` hiện đang nằm ở `data/HOTEL/HANOI/...`, trong khi `SOURCE_DIR` mặc định trong code là `data/HANOI`. Hãy sửa `SOURCE_DIR` (và `OUTPUT_BASE_DIR` nếu cần) cho khớp với vị trí dữ liệu thực tế trước khi chạy, nếu không script sẽ báo "Không tìm thấy thư mục".
- **Chi phí API**: mỗi file được xử lý sẽ tốn 1 lượt gọi Claude API, tính phí theo tài khoản Anthropic của bạn.
- **Không commit API key**: tuyệt đối không hard-code hoặc commit API key vào code/Git.
