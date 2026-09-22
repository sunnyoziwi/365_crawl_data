from pathlib import Path

from extractor import VALID_EXTS, extract_rates, read_text_from_file, write_excel

# =============================================================
# CẤU HÌNH ĐƯỜNG DẪN VÀ THIẾT LẬP
# =============================================================
SOURCE_DIR = Path("data/HOTEL/HANOI")
OUTPUT_BASE_DIR = Path("data_result/HANOI")

LIMIT_HOTELS = 5


def process_file(source_file_path: Path, output_file_path: Path, hotel_name_hint: str):
    output_file_path.parent.mkdir(parents=True, exist_ok=True)

    if output_file_path.exists():
        print(f"      [SKIP] Đã tồn tại: {output_file_path.name}")
        return

    doc_text = read_text_from_file(source_file_path)
    if not doc_text.strip():
        print(f"      [!] Không bóc tách được văn bản: {source_file_path.name}")
        return

    try:
        rates = extract_rates(doc_text, hotel_name_hint)
        if rates:
            write_excel(rates, output_file_path)
            print(f"      [✓] Hoàn thành: {output_file_path.name}")
        else:
            print(f"      [-] Không có dữ liệu rates: {source_file_path.name}")
    except Exception as e:
        print(f"      [X] Lỗi API tại {source_file_path.name}: {e}")


# =============================================================
# QUÉT VÀ CHẠY TOÀN BỘ FOLDER
# =============================================================
def main():
    if not SOURCE_DIR.exists():
        print(f"Không tìm thấy thư mục nguồn: {SOURCE_DIR.resolve()}")
        return

    hotel_folders = sorted([d for d in SOURCE_DIR.iterdir() if d.is_dir()])

    if LIMIT_HOTELS:
        hotel_folders = hotel_folders[:LIMIT_HOTELS]

    print(f"=== Bắt đầu xử lý {len(hotel_folders)} thư mục khách sạn ===")

    for idx, hotel_dir in enumerate(hotel_folders, 1):
        hotel_name = hotel_dir.name
        print(f"\n[{idx}/{len(hotel_folders)}] Khách sạn: {hotel_name}")

        all_files = [f for f in hotel_dir.rglob("*") if f.is_file() and f.suffix.lower() in VALID_EXTS]

        if not all_files:
            print("    (Không có tệp tài liệu hợp lệ)")
            continue

        for f in all_files:
            relative_path = f.relative_to(hotel_dir)
            target_xlsx_path = (OUTPUT_BASE_DIR / hotel_name / relative_path).with_suffix(".xlsx")

            print(f"    -> Đang xử lý: {relative_path}")
            process_file(f, target_xlsx_path, hotel_name)

    print("\n=== Hoàn thành toàn bộ tiến trình! ===")


if __name__ == "__main__":
    main()
