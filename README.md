# 365 Crawl Data

Script Python dùng Claude API (Anthropic) để tự động đọc file hợp đồng khách sạn (PDF, DOCX, DOC, TXT) và trích xuất bảng giá phòng theo mùa ra file Excel với các cột cố định. Có 2 cách dùng: chạy hàng loạt theo thư mục (`crawl.py`) hoặc tải file trực tiếp qua giao diện web (`app.py`, dùng Gradio).

## Cấu trúc thư mục

```
365_crawl_data/
├── crawl.py             # Xử lý hàng loạt theo thư mục data/HOTEL/...
├── app.py                # Giao diện Gradio: tải file lên, tải file Excel kết quả về
├── extractor.py          # Logic dùng chung (đọc file, gọi Claude, ghi Excel)
├── requirements.txt      # Danh sách thư viện cần cài
├── .env                  # API key (tự tạo, KHÔNG đẩy lên Git)
├── .env.example          # Mẫu file .env
├── data/                 # Dữ liệu đầu vào (hợp đồng gốc, dùng cho crawl.py)
│   └── HOTEL/
│       └── HANOI/
│           └── <Tên khách sạn>/
│               └── ... (file .pdf, .docx, .doc, .txt)
└── data_result/          # Dữ liệu đầu ra (Excel đã trích xuất) - KHÔNG đẩy lên Git
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

## Cách chạy (1) — Xử lý hàng loạt theo thư mục

```powershell
python crawl.py
```

Script sẽ:
1. Duyệt qua từng thư mục khách sạn trong `SOURCE_DIR` (mặc định `data/HOTEL/HANOI`)
2. Đọc toàn bộ file `.pdf`, `.docx`, `.doc`, `.txt` bên trong (kể cả thư mục con)
3. Gửi nội dung cho Claude để trích xuất bảng giá theo schema định sẵn
4. Ghi kết quả ra file **Excel (`.xlsx`)** tương ứng vào `OUTPUT_BASE_DIR` (mặc định `data_result/HANOI`), giữ nguyên cấu trúc thư mục

Nếu file Excel kết quả đã tồn tại, script sẽ bỏ qua (không xử lý lại) để tiết kiệm chi phí gọi API.

## Cách chạy (2) — Giao diện web (Gradio)

```powershell
python app.py
```

Mở địa chỉ hiện ra trong terminal (mặc định `http://127.0.0.1:7860`), sau đó:
1. Tải lên một hoặc nhiều file hợp đồng (PDF, DOCX, DOC, TXT)
2. Bấm **Trích xuất & Tạo Excel**
3. Tải kết quả về — **mỗi file input trả về 1 file Excel riêng** (cùng tên với file gốc)

Phù hợp khi cần xử lý nhanh vài file lẻ, không cần sắp xếp vào cấu trúc thư mục `data/HOTEL/...`.

### Tự host cho máy khác cùng wifi/LAN truy cập

Mặc định app chỉ chạy ở `127.0.0.1` — chỉ máy đang chạy `python app.py` mới mở được. Để các máy khác **cùng wifi/LAN** cũng vào được:

1. Chạy app với `server_name` mở ra toàn bộ mạng (thay vì chỉ localhost):

   ```powershell
   # Windows (PowerShell)
   $env:GRADIO_SERVER_NAME="0.0.0.0"
   $env:GRADIO_SERVER_PORT="7860"
   python app.py
   ```

   ```bash
   # macOS/Linux
   GRADIO_SERVER_NAME=0.0.0.0 GRADIO_SERVER_PORT=7860 python app.py
   ```

2. Lấy IP LAN của máy đang chạy app:
   - macOS: `ipconfig getifaddr en0` (wifi) hoặc mở System Settings → Wi-Fi → Details
   - Windows: chạy `ipconfig`, xem dòng **IPv4 Address**

3. Từ máy khác **cùng wifi**, mở trình duyệt vào `http://<IP-LAN-máy-host>:<PORT>`, ví dụ `http://192.168.1.5:7860`

Lưu ý:
- Cách này chỉ hoạt động cho máy **cùng mạng LAN/wifi**. Máy ở wifi/mạng khác sẽ **không** vào được, vì `192.168.x.x` là IP nội bộ, router không tự chuyển tiếp traffic từ ngoài vào. Muốn máy khác wifi vào được, dùng `demo.launch(share=True)` (Gradio tạo link public tạm qua tunnel của họ) hoặc deploy app lên server/cloud có IP public.
- Nếu đúng IP mà vẫn không vào được: kiểm tra Firewall của máy host có chặn cổng đó không, hoặc wifi đang bật **Client/AP Isolation** (chặn các thiết bị thấy nhau) — thường gặp ở wifi quán cà phê, khách sạn, mạng khách (guest wifi).

### Định dạng file Excel kết quả

Mỗi file `.xlsx` có các cột cố định sau, mỗi dòng là 1 loại phòng trong 1 khoảng ngày (mùa giá):

| City | Hotel Name | Room type | Capacity | From | Until | Single | Double | Extra Bed | Triple | Quad |
|---|---|---|---|---|---|---|---|---|---|---|
| Hanoi | Apricot | Deluxe | 2 | 01-Nov-26 | 03-May-27 | | 1,800,000 | 500,000 | 2,300,000 | |

Quy tắc trích xuất:
- Chỉ lấy giá **FIT** (khách lẻ), bỏ qua giá **GIT/Group**.
- Trường không có dữ liệu sẽ để **trống** (không điền "N/A", "unknown"...).
- Nếu hợp đồng không có sẵn giá **Triple**, hệ thống tự suy ra = Double + Extra Bed (khi có đủ 2 giá này). Giá **Quad** không bao giờ tự suy ra, chỉ lấy khi hợp đồng ghi rõ.

## Các tuỳ chọn cấu hình (đầu file `extractor.py`)

| Biến | Ý nghĩa | Mặc định |
|---|---|---|
| `MODEL_NAME` | Model Claude sử dụng | `claude-haiku-4-5` |

### Tuỳ chọn riêng của `app.py` (giao diện Gradio, đặt qua biến môi trường)

| Biến môi trường | Ý nghĩa | Mặc định |
|---|---|---|
| `GRADIO_SERVER_NAME` | Địa chỉ IP để lắng nghe. `127.0.0.1` = chỉ máy này; `0.0.0.0` = mở cho cả mạng LAN/wifi | `127.0.0.1` |
| `GRADIO_SERVER_PORT` | Cổng chạy app | `7860` |

### Tuỳ chọn riêng của `crawl.py` (xử lý hàng loạt)

| Biến | Ý nghĩa | Mặc định |
|---|---|---|
| `SOURCE_DIR` | Thư mục chứa dữ liệu gốc cần xử lý | `data/HOTEL/HANOI` |
| `OUTPUT_BASE_DIR` | Thư mục ghi kết quả Excel | `data_result/HANOI` |
| `LIMIT_HOTELS` | Giới hạn số khách sạn xử lý (để test). Đặt `None` để chạy toàn bộ | `5` |

## Lưu ý quan trọng

- **Chi phí API**: mỗi file được xử lý sẽ tốn 1 lượt gọi Claude API, tính phí theo tài khoản Anthropic của bạn.
- **Không commit API key**: tuyệt đối không hard-code hoặc commit key thật vào code/Git — chỉ dùng file `.env` (đã bị `.gitignore` chặn).
