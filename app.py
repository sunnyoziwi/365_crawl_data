"""Giao diện Gradio: tải lên nhiều file hợp đồng (PDF/DOCX/DOC/TXT),
trích xuất bảng giá bằng Claude và xuất ra 1 file Excel duy nhất
theo bộ cột cố định (xem extractor.COLUMNS).
"""
from pathlib import Path

import gradio as gr

from extractor import VALID_EXTS, extract_rates, read_text_from_file, write_excel

OUTPUT_DIR = Path("data_result/gradio")


def _file_path(f) -> Path:
    return Path(f if isinstance(f, str) else f.name)


def process_files(files, progress=gr.Progress()):
    if not files:
        raise gr.Error("Vui lòng tải lên ít nhất 1 file (PDF, DOCX, DOC, TXT).")

    all_rows = []
    processed_stems = []
    log_lines = []

    for f in progress.tqdm(files, desc="Đang trích xuất dữ liệu..."):
        path = _file_path(f)

        if path.suffix.lower() not in VALID_EXTS:
            log_lines.append(f"⏭️ {path.name}: định dạng không được hỗ trợ, bỏ qua.")
            continue

        text = read_text_from_file(path)
        if not text.strip():
            log_lines.append(f"⚠️ {path.name}: không đọc được nội dung văn bản.")
            continue

        try:
            rates = extract_rates(text, hotel_name_hint=path.stem)
        except Exception as e:
            log_lines.append(f"❌ {path.name}: lỗi khi gọi Claude API ({e}).")
            continue

        if not rates:
            log_lines.append(f"➖ {path.name}: không tìm thấy dòng giá nào.")
            continue

        all_rows.extend(rates)
        processed_stems.append(path.stem)
        log_lines.append(f"✅ {path.name}: trích xuất {len(rates)} dòng giá.")

    if not all_rows:
        raise gr.Error("Không trích xuất được dữ liệu nào từ các file đã tải lên.\n\n" + "\n".join(log_lines))

    if len(processed_stems) == 1:
        output_name = f"{processed_stems[0]}.xlsx"
    else:
        output_name = f"{processed_stems[0]}_gop_{len(processed_stems)}_file.xlsx"

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    output_path = OUTPUT_DIR / output_name
    write_excel(all_rows, output_path)

    log_lines.append(f"\n🎉 Hoàn thành: {len(all_rows)} dòng giá, {len(files)} file đã xử lý.")
    return str(output_path), "\n".join(log_lines)


with gr.Blocks(title="Trích xuất bảng giá khách sạn") as demo:
    gr.Markdown(
        "# 🏨 Trích xuất bảng giá khách sạn → Excel\n"
        "Tải lên một hoặc nhiều file hợp đồng (**PDF, DOCX, DOC, TXT**). "
        "Claude sẽ đọc và trích xuất bảng giá FIT theo mùa, sau đó gộp vào **1 file Excel** "
        "với các cột cố định: City, Hotel Name, Room type, Capacity, From, Until, Single, Double, Extra Bed, Triple, Quad."
    )

    with gr.Row():
        file_input = gr.File(
            label="File hợp đồng",
            file_count="multiple",
            file_types=[".pdf", ".docx", ".doc", ".txt"],
        )

    run_btn = gr.Button("Trích xuất & Tạo Excel", variant="primary")

    with gr.Row():
        output_file = gr.File(label="File Excel kết quả")

    log_output = gr.Textbox(label="Nhật ký xử lý", lines=10, interactive=False)

    run_btn.click(fn=process_files, inputs=file_input, outputs=[output_file, log_output])

if __name__ == "__main__":
    demo.launch()
