"""Giao diện Gradio: tải lên nhiều file hợp đồng (PDF/DOCX/DOC/TXT),
trích xuất bảng giá bằng Claude và xuất ra mỗi file input 1 file Excel
tương ứng (cùng tên), theo bộ cột cố định (xem extractor.COLUMNS).

Khách sạn và villa dùng chung 1 logic: chỉ villa từ 2 phòng ngủ trở lên
mới điền giá vào cột Quad, các phòng khác điền theo occupancy.
"""
import os
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

        if not read_text_from_file(path).strip():
            log_lines.append(f"⚠️ {path.name}: không đọc được nội dung văn bản.")
            continue

        try:
            rates = extract_rates(path, hotel_name_hint=path.stem)
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
        "# 🏨 Trích xuất bảng giá → Excel\n"
        "- File hỗ trợ: **PDF, DOCX, DOC, TXT**, mỗi file trả về **1 file Excel** cùng tên\n\n"
        "**🏨 Rule khách sạn** (áp dụng cho mọi phòng)\n"
        "- Chỉ lấy giá **FIT**, bỏ giá GIT/Group\n"
        "- Thiếu thông tin → để trống\n"
        "- Quad chỉ điền khi hợp đồng có giá, không tự tính\n"
        "- **Single = Double**\n"
        "- Không có Triple → **Triple = Double + Extra Bed**\n\n"
        "**🏡 Rule villa**\n"
        "- Villa từ **2 phòng ngủ trở lên** (vd 'Two-Bedroom Pool Villa') → giá chỉ điền cột **Quad**\n"
        "- Villa còn lại → dùng rule khách sạn"
    )

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
    # Mặc định chỉ máy này truy cập được (127.0.0.1). Đặt GRADIO_SERVER_NAME=0.0.0.0
    # để máy khác cùng mạng LAN/wifi vào được qua IP LAN của máy này (xem README).
    demo.launch(
        server_name=os.environ.get("GRADIO_SERVER_NAME", "127.0.0.1"),
        server_port=int(os.environ.get("GRADIO_SERVER_PORT", "7861")),
    )
