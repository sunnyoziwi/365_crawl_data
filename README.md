# 365 Crawl Data

Script Python dùng Claude API (Anthropic) để tự động đọc file hợp đồng khách sạn (PDF, DOCX, DOC, TXT) và trích xuất thông tin có cấu trúc (tên khách sạn, loại phòng, bảng giá theo mùa, chính sách...) ra file Excel.

## Cấu trúc thư mục

```
365_crawl_data/
├── crawl.py            # Script chính
├── requirements.txt    # Danh sách thư viện cần cài
├── .env                 # API key (tự tạo, KHÔNG đẩy lên Git)
├── .env.example         # Mẫu file .env
├── data/                # Dữ liệu đầu vào (hợp đồng gốc)
│   └── HOTEL/
│       └── HANOI/
│           └── <Tên khách sạn>/
│               └── ... (file .pdf, .docx, .doc, .txt)
└── data_result/         # Dữ liệu đầu ra (Excel đã trích xuất) - KHÔNG đẩy lên Git
```

> Thư mục `data_result/` không được đưa lên GitHub (xem `.gitignore`) vì là dữ liệu sinh ra từ script, người dùng nào cũng tự tạo lại được.

## Yêu cầu

- Python 3.10 trở lên
- Một API key của Anthropic ([console.anthropic.com](https://console.anthropic.com))

## Cài đặt

### 0. Extract file zip (_Tải tại ô Code (Xanh lá cây)_ -> _Ấn "Download Zip"_)

### 0,5. Mở file explore, tại folder 365_crawl_data_main, click chuột ở thanh địa chỉ (hoặc bấm Ctrl + L), Gõ cmd/Powershell và nhấn enter

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

Tạo file `.env` (copy từ `.env.example`) trong thư mục gốc project, điền key thật:

```
ANTHROPIC_API_KEY=sk-ant-...
```

Script tự động đọc key từ file `.env` này khi chạy (dùng `python-dotenv`), không cần set biến môi trường Windows thủ công. File `.env` đã được `.gitignore` chặn nên không bao giờ bị đẩy lên Git.

## Cách chạy

```powershell
python crawl.py
```

Script sẽ:
1. Duyệt qua từng thư mục khách sạn trong `SOURCE_DIR` (mặc định `data/HOTEL/HANOI`)
2. Đọc toàn bộ file `.pdf`, `.docx`, `.doc`, `.txt` bên trong (kể cả thư mục con)
3. Gửi nội dung cho Claude để trích xuất thông tin theo schema định sẵn
4. Ghi kết quả ra file **Excel (`.xlsx`)** tương ứng vào `OUTPUT_BASE_DIR` (mặc định `data_result/HANOI`), giữ nguyên cấu trúc thư mục

Nếu file Excel kết quả đã tồn tại, script sẽ bỏ qua (không xử lý lại) để tiết kiệm chi phí gọi API.

### Định dạng file Excel kết quả

Mỗi file `.xlsx` có 2 cột **Key / Value**. Các trường lồng nhau trong dữ liệu gốc được làm phẳng (flatten) bằng cách nối tên bằng dấu `_`, phần tử trong danh sách (array) được đánh số bắt đầu từ 0. Ví dụ:

| Key | Value |
|---|---|
| hotel_information_hotel_name | Apricot |
| contract_information_contract_name | Hợp đồng 2024 |
| room_information_0_room_name | Deluxe |
| room_information_1_room_name | Suite |
| policies_raw_text_child_policy_0 | Trẻ dưới 6 tuổi miễn phí |

## Các tuỳ chọn cấu hình (đầu file `crawl.py`)

| Biến | Ý nghĩa | Mặc định |
|---|---|---|
| `SOURCE_DIR` | Thư mục chứa dữ liệu gốc cần xử lý | `data/HOTEL/HANOI` |
| `OUTPUT_BASE_DIR` | Thư mục ghi kết quả Excel | `data_result/HANOI` |
| `LIMIT_HOTELS` | Giới hạn số khách sạn xử lý (để test). Đặt `None` để chạy toàn bộ | `5` |
| `MODEL_NAME` | Model Claude sử dụng | `claude-haiku-4-5` |

## Lưu ý quan trọng

- **Chi phí API**: mỗi file được xử lý sẽ tốn 1 lượt gọi Claude API, tính phí theo tài khoản Anthropic của bạn.
- **Không commit API key**: tuyệt đối không hard-code hoặc commit key thật vào code/Git — chỉ dùng file `.env` (đã bị `.gitignore` chặn).
