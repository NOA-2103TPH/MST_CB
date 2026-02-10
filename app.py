import os
import io
import tempfile
import pandas as pd
import streamlit as st

from save_code import run_lookup

st.set_page_config(page_title="Tra cứu MST theo CCCD", layout="wide")
st.markdown(
    "<h1 style='text-align:center;'>Tra cứu MST theo CCCD</h1>",
    unsafe_allow_html=True,
)

if "logs" not in st.session_state:
    st.session_state["logs"] = []
if "output_path" not in st.session_state:
    st.session_state["output_path"] = None

progress_bar = st.progress(0, text="Chưa chạy")
log_box = st.empty()


uploaded = st.file_uploader("Chọn file Excel đầu vào (cột CCCD bắt buộc)", type=["xlsx", "xls", "csv"])

show_browser = st.checkbox("Mở cửa sổ Chrome khi chạy (tắt nếu chạy trên server không có GUI)", value=True)
show_logs = st.checkbox("Hiển thị log chi tiết", value=False)

col1, col2 = st.columns([1, 1])
run_clicked = col1.button("Run tra cứu", type="primary", use_container_width=True, disabled=uploaded is None)
clear_clicked = col2.button("Xóa log", use_container_width=True)

if clear_clicked:
    st.session_state["logs"] = []
    log_box.text("")
    progress_bar.progress(0, text="Chưa chạy")

if run_clicked and uploaded:
    # Lưu file upload vào temp
    workdir = tempfile.mkdtemp(prefix="mst_lookup_")
    input_path = os.path.join(workdir, "input.xlsx")
    with open(input_path, "wb") as f:
        f.write(uploaded.read())

    output_path = os.path.join(workdir, "output.xlsx")
    st.session_state["output_path"] = output_path

    st.info("Bắt đầu chạy, vui lòng chờ...")
    def update_progress(done, total):
        ratio = done / total if total else 0
        progress_bar.progress(ratio, text=f"Đang chạy: {done}/{total}")

    def log_ui(msg: str):
        if not show_logs:
            return
        st.session_state["logs"].append(str(msg))
        log_box.text("\n".join(st.session_state["logs"][-200:]))  # keep last 200 lines

    run_lookup(
        input_path,
        output_path,
        log_fn=log_ui,
        headless=not show_browser,
        progress_fn=update_progress,
    )
    progress_bar.progress(1.0, text="Hoàn tất")
    st.success("Đã chạy xong.")

if st.session_state.get("output_path") and os.path.exists(st.session_state["output_path"]):
    st.subheader("Xem trước kết quả")
    excel_bytes = None
    try:
        with open(st.session_state["output_path"], "rb") as f:
            excel_bytes = f.read()
        # preview từ buffer để tránh lỗi handle
        df_preview = pd.read_excel(io.BytesIO(excel_bytes), dtype=str)
        st.dataframe(df_preview.head(50), use_container_width=True)
    except Exception as e:
        st.info(f"Không đọc được preview: {e}")

    if excel_bytes:
        try:
            st.download_button(
                "Tải kết quả Excel",
                data=excel_bytes,
                file_name="ket_qua_mst.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.warning(f"Không tạo được nút tải: {e}")

st.caption("Cần thư viện: streamlit, selenium, webdriver-manager, pandas.")
