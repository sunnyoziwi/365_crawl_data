"""Giao diện Gradio: tải lên nhiều file hợp đồng (PDF/DOCX/DOC/TXT),
trích xuất bảng giá bằng Claude và xuất ra mỗi file input 1 file Excel
tương ứng (cùng tên), theo bộ cột cố định (xem extractor.COLUMNS).
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

    output_paths = []
    log_lines = []

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

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

        output_path = OUTPUT_DIR / f"{path.stem}.xlsx"
        write_excel(rates, output_path)
        output_paths.append(str(output_path))
        log_lines.append(f"✅ {path.name}: trích xuất {len(rates)} dòng giá → {output_path.name}")

    if not output_paths:
        raise gr.Error("Không trích xuất được dữ liệu nào từ các file đã tải lên.\n\n" + "\n".join(log_lines))

    log_lines.append(f"\n🎉 Hoàn thành: {len(output_paths)}/{len(files)} file đã xử lý.")
    return output_paths, "\n".join(log_lines)


with gr.Blocks(title="Trích xuất bảng giá khách sạn") as demo:
    gr.Markdown(
        "# 🏨 Trích xuất bảng giá khách sạn → Excel\n"
        "Tải lên một hoặc nhiều file hợp đồng (**PDF, DOCX, DOC, TXT**). "
        "Claude sẽ đọc và trích xuất bảng giá FIT theo mùa, mỗi file input trả về **1 file Excel riêng** (cùng tên) "
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
        output_file = gr.File(label="File Excel kết quả", file_count="multiple")

    log_output = gr.Textbox(label="Nhật ký xử lý", lines=10, interactive=False)

    run_btn.click(fn=process_files, inputs=file_input, outputs=[output_file, log_output])

if __name__ == "__main__":
    demo.launch()
