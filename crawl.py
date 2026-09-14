import os
from pathlib import Path
from anthropic import Anthropic
from dotenv import load_dotenv
import openpyxl
import pypdf
import docx

load_dotenv()

# =============================================================
# 1. CẤU HÌNH ĐƯỜNG DẪN VÀ THIẾT LẬP
# =============================================================
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
SOURCE_DIR = Path("data/HOTEL/HANOI")
OUTPUT_BASE_DIR = Path("data_result/HANOI")

# Đặt số lượng folder khách sạn cần test (None nếu muốn xử lý toàn bộ)
LIMIT_HOTELS = 5  

# Dùng claude-3-haiku-20240307 (hoặc "claude-3-5-haiku-latest") để tránh lỗi 404
MODEL_NAME = "claude-haiku-4-5"

if not ANTHROPIC_API_KEY:
    raise RuntimeError("Thiếu ANTHROPIC_API_KEY. Hãy đặt biến môi trường ANTHROPIC_API_KEY trước khi chạy.")

client = Anthropic(api_key=ANTHROPIC_API_KEY.strip())

# =============================================================
# 2. SCHEMA TRÍCH XUẤT DỮ LIỆU CHUẨN JSON
# =============================================================
EXTRACTION_TOOL = {
    "name": "save_contract_data",
    "description": "Extract structured hotel contract, rates, rooms, and policies.",
    "input_schema": {
        "type": "object",
        "properties": {
            "hotel_information": {
                "type": "object",
                "properties": {
                    "hotel_name": {"type": "string"}
                },
                "required": ["hotel_name"]
            },
            "contract_information": {
                "type": "object",
                "properties": {
                    "contract_name": {"type": "string"},
                    "validity_start_date": {"type": "string"},
                    "validity_end_date": {"type": "string"}
                },
                "required": ["contract_name"]
            },
            "room_information": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "room_name": {"type": "string"},
                        "room_category": {"type": "string"},
                        "size_sqm": {"type": "string"},
                        "number_of_rooms": {"type": "string"},
                        "bed_types": {"type": "array", "items": {"type": "string"}},
                        "max_occupancy": {"type": "string"},
                        "extra_bed_allowed": {"type": "string"}
                    },
                    "required": ["room_name"]
                }
            },
            "seasonal_rates": {
                "type": "array",
                "items": {
                    "type": "object",
                    "properties": {
                        "season_name": {"type": "string"},
                        "date_range": {"type": "string"},
                        "rates": {
                            "type": "array",
                            "items": {
                                "type": "object",
                                "properties": {
                                    "room_name": {"type": "string"},
                                    "market_type": {"type": "string"},
                                    "applicable_nationality": {"type": "array", "items": {"type": "string"}},
                                    "applicable_days": {"type": "array", "items": {"type": "string"}},
                                    "currency": {"type": "string"},
                                    "price": {"type": "string"}
                                }
                            }
                        }
                    }
                }
            },
            "policies_raw_text": {
                "type": "object",
                "properties": {
                    "extra_bed_policy": {"type": "array", "items": {"type": "string"}},
                    "child_policy": {"type": "array", "items": {"type": "string"}},
                    "meal_and_beverage_policy": {"type": "array", "items": {"type": "string"}},
                    "checkin_checkout_policy": {"type": "array", "items": {"type": "string"}},
                    "cancellation_and_payment_policy": {"type": "array", "items": {"type": "string"}},
                    "surcharges_and_other_notes": {"type": "array", "items": {"type": "string"}}
                }
            }
        },
        "required": [
            "hotel_information",
            "contract_information",
            "room_information",
            "seasonal_rates",
            "policies_raw_text"
        ]
    }
}

# =============================================================
# 3. LÀM PHẲNG DỮ LIỆU LỒNG NHAU VÀ GHI FILE EXCEL
# =============================================================
def flatten_data(obj, prefix=""):
    """Làm phẳng dict/list lồng nhau thành dict 1 cấp, key nối bằng dấu '_' (vd: room_information_0_room_name)."""
    flat = {}
    if isinstance(obj, dict):
        for key, value in obj.items():
            new_key = f"{prefix}_{key}" if prefix else key
            flat.update(flatten_data(value, new_key))
    elif isinstance(obj, list):
        for idx, item in enumerate(obj):
            new_key = f"{prefix}_{idx}"
            flat.update(flatten_data(item, new_key))
    else:
        flat[prefix] = obj
    return flat


def write_excel(data: dict, output_path: Path):
    flat_data = flatten_data(data)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["Key", "Value"])
    for key, value in flat_data.items():
        ws.append([key, value])
    wb.save(output_path)

# =============================================================
# 4. TRÍCH XUẤT VĂN BẢN TỪ FILE
# =============================================================
def read_text_from_file(file_path: Path) -> str:
    ext = file_path.suffix.lower()
    
    if ext == ".txt":
        with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()

    elif ext == ".docx":
        try:
            doc = docx.Document(file_path)
            full_text = [p.text for p in doc.paragraphs]
            for table in doc.tables:
                for row in table.rows:
                    full_text.append(" | ".join(cell.text.strip() for cell in row.cells))
            return "\n".join(full_text)
        except Exception as e:
            print(f"      [!] Lỗi mở docx: {e}")
            return ""

    elif ext == ".pdf":
        try:
            reader = pypdf.PdfReader(str(file_path))
            text = []
            for page in reader.pages:
                t = page.extract_text()
                if t:
                    text.append(t)
            return "\n".join(text)
        except Exception as e:
            print(f"      [!] Lỗi mở pdf: {e}")
            return ""

    elif ext == ".doc":
        try:
            with open(file_path, "rb") as f:
                content = f.read()
            return "".join([chr(b) for b in content if 32 <= b <= 126 or b in (10, 13)])
        except Exception:
            return ""

    return ""

# =============================================================
# 5. GỬI PROMPT VÀ GHI FILE KẾT QUẢ
# =============================================================
def process_file(source_file_path: Path, output_file_path: Path):
    output_file_path.parent.mkdir(parents=True, exist_ok=True)
    
    if output_file_path.exists():
        print(f"      [SKIP] Đã tồn tại: {output_file_path.name}")
        return

    doc_text = read_text_from_file(source_file_path)
    if not doc_text.strip():
        print(f"      [!] Không bóc tách được văn bản: {source_file_path.name}")
        return

    prompt = (
        "Trích xuất toàn bộ thông tin hợp đồng khách sạn, bảng giá, loại phòng và chính sách "
        "từ văn bản dưới đây và điền chính xác vào hàm save_contract_data:\n\n"
        f"--- BẮT ĐẦU TÀI LIỆU ---\n{doc_text}\n--- KẾT THÚC TÀI LIỆU ---"
    )

    try:
        response = client.messages.create(
            model=MODEL_NAME,
            max_tokens=4096,
            tools=[EXTRACTION_TOOL],
            tool_choice={"type": "tool", "name": "save_contract_data"},
            messages=[{"role": "user", "content": prompt}]
        )

        for block in response.content:
            if block.type == "tool_use" and block.name == "save_contract_data":
                write_excel(block.input, output_file_path)
                print(f"      [✓] Hoàn thành: {output_file_path.name}")
                return

    except Exception as e:
        print(f"      [X] Lỗi API tại {source_file_path.name}: {e}")

# =============================================================
# 6. CHẠY VÒNG LẶP TOÀN BỘ CẤU TRÚC THƯ MỤC
# =============================================================
def main():
    if not SOURCE_DIR.exists():
        print(f"Không tìm thấy thư mục: {SOURCE_DIR.resolve()}")
        return

    hotel_folders = sorted([d for d in SOURCE_DIR.iterdir() if d.is_dir()])
    
    if LIMIT_HOTELS:
        hotel_folders = hotel_folders[:LIMIT_HOTELS]

    print(f"=== Bắt đầu xử lý {len(hotel_folders)} khách sạn ===")
    valid_exts = {".pdf", ".docx", ".doc", ".txt"}

    for idx, hotel_dir in enumerate(hotel_folders, 1):
        hotel_name = hotel_dir.name
        print(f"\n[{idx}/{len(hotel_folders)}] Khách sạn: {hotel_name}")

        # Duyệt đệ quy tìm tất cả file kể cả trong thư mục con như 2024-2025, 2025...
        all_files = [f for f in hotel_dir.rglob("*") if f.is_file() and f.suffix.lower() in valid_exts]

        if not all_files:
            print(f"    (Không tìm thấy file hợp lệ trong thư mục hoặc thư mục con)")
            continue

        for f in all_files:
            relative_path = f.relative_to(hotel_dir)
            target_xlsx_path = (OUTPUT_BASE_DIR / hotel_name / relative_path).with_suffix(".xlsx")

            print(f"    -> Đang xử lý: {relative_path}")
            process_file(f, target_xlsx_path)

    print("\n=== Hoàn thành! ===")

if __name__ == "__main__":
    main()