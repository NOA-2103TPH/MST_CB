# Tra cứu MST theo CCCD (Selenium + Streamlit)

Ứng dụng này tự động tra cứu mã số thuế (MST) từ số CCCD trên trang masothue.com và xuất kết quả ra Excel. Có thể chạy dạng script CLI hoặc giao diện web Streamlit.

## Yêu cầu môi trường
- Python 3.9+ (đã thử trên 3.12)
- Google Chrome và chromedriver tương ứng (webdriver-manager sẽ tải giúp)
- Các thư viện Python:
  ```bash
  pip install -r requirements.txt
  ```
  hoặc ngắn gọn:
  ```bash
  pip install streamlit selenium webdriver-manager pandas openpyxl
  ```

## Chuẩn bị dữ liệu
- File Excel đầu vào cần có cột `CCCD`. Các cột khác (`MST`, `Tên`, `Trạng thái`) có thể để trống, code sẽ tự thêm.
- Mặc định đường dẫn:
  - Input: `data/Book2.xlsx` (có thể đổi bằng biến môi trường `INPUT_FILE`)
  - Output: `result/kq_loc1.xlsx` (có thể đổi bằng biến môi trường `OUTPUT_FILE`)

## Chạy dạng script CLI
```bash
python save_code.py
```
- Script sẽ mở Chrome (hiển thị). Muốn ẩn (headless) thì đặt biến `HEADLESS=1` và chỉnh trong code nếu cần.
- Kết quả ghi dần vào file output sau mỗi dòng tra cứu.

## Chạy giao diện Streamlit
```bash
streamlit run app.py
```
Trong giao diện:
1. Upload file Excel/CSV.
2. Chọn tùy chọn:
   - **Mở cửa sổ Chrome khi chạy**: bật để xem trình duyệt tự chạy; tắt nếu chạy trên server không có GUI.
   - **Hiển thị log chi tiết**: bật khi cần debug, tắt để giao diện gọn.
3. Bấm **Run tra cứu**.
4. Sau khi xong:
   - Khu vực **Xem trước kết quả** hiển thị 50 dòng đầu (đọc dạng chuỗi nên không mất số 0 ở đầu).
   - Nút **Tải kết quả Excel** để tải toàn bộ file.

## Thông số mặc định (có thể đổi trong `save_code.py`)
- Nghỉ 2 phút sau mỗi 120 bản ghi (`batch_size`, `rest_seconds`).
- Trễ lịch sự 2 giây giữa các request.
- Tự thử đóng popup quảng cáo và refresh mỗi 10 lần chờ.
- Retry 1 lần với driver mới nếu gặp lỗi tương tác/kết nối.

## Mẹo & xử lý sự cố
- Nếu bị chặn quảng cáo làm kẹt ô tìm kiếm: để Chrome hiển thị và đóng thủ công; hoặc chạy lại với tùy chọn hiển thị.
- Nếu mất số 0 ở đầu trong Excel: luôn mở bằng chế độ text hoặc kiểm tra preview trong app (đã hiển thị dạng chuỗi).
- Nếu không xem được preview, kiểm tra log trên Streamlit (bật “Hiển thị log chi tiết”) và đảm bảo file output có thực sự được ghi.

## Cấu trúc repo chính
- `save_code.py`: lõi tra cứu, có hàm `run_lookup(input_path, output_path, ...)`.
- `app.py`: giao diện Streamlit.
- `result/`, `data/`: thư mục mặc định chứa file đầu vào/đầu ra.
- `lookup.log`, `page_dump*.html`: log và snapshot khi lỗi.

## Deploy lên Streamlit Cloud
1. Đảm bảo đã commit hai file:
   - `requirements.txt` (streamlit, selenium, webdriver-manager, pandas, openpyxl, numpy)
   - `packages.txt` (chromium, chromium-driver) để có sẵn trình duyệt/driver.
2. Trên Streamlit Cloud, chọn repo và file khởi chạy `app.py`.
3. Cloud sẽ tự cài gói pip và apt theo 2 file trên. Nếu cần, đặt biến môi trường:
   - `STREAMLIT_SERVER_HEADLESS=true` (mặc định có sẵn)
   - `CHROME_BIN=/usr/bin/chromium` (thường đúng với packages.txt)
   - `CHROMEDRIVER_PATH=/usr/bin/chromedriver`
4. Sau khi build xong, chạy app. Nếu thấy lỗi ModuleNotFoundError (selenium...), kiểm tra `requirements.txt` đã push chưa và bấm “Clear cache and rerun” trên Streamlit Cloud.

## Giấy phép
Bạn tùy ý chọn giấy phép khi public; file này không gán license.***
