# TRỘN ĐỀ PRO (TRON-DE-PRO)
### Hệ Thống Tự Động Hóa Trộn Đề Thi Trắc Nghiệm Chuẩn GDPT 2018 & CĐN01

> **Tác giả**: ThS. Nguyễn Văn Sang  
> **Đơn vị**: Khoa Cơ bản — Trường Cao đẳng Nghề số 1 — BQP  
> **Repository**: [https://github.com/sanginnova/tron-de-pro](https://github.com/sanginnova/tron-de-pro)

---

## 🌟 ĐẶC ĐIỂM NỔI BẬT

Khắc phục triệt để toàn bộ nhược điểm của các phần mềm trộn đề truyền thống:

1. **Bảo tồn 100% MathType OLE & OMML Native (Zero Loss)**:
   - Sử dụng cơ chế sao chép sâu và clone package DOCX gốc.
   - Không bị mất công thức, không bị biến dạng hay mờ như các tool chuyển sang ảnh pixel.
   - Người dùng có thể click đúp chuột vào công thức để chỉnh sửa trực tiếp bằng MathType bình thường.

2. **Dàn TabStop thông minh 1 – 2 – 4 cột chuẩn Sư phạm**:
   - Tự động đo kích thước vật lý thực tế (cm & pt) của thẻ `<v:shape>` và văn bản:
     * **4 Cột (1 dòng)**: Khi độ rộng $\le 3.65\text{ cm}$ (đáp án số, chữ, công thức ngắn), tiết kiệm tối đa diện tích giấy in.
     * **2 Cột (2 dòng)**: Khi độ rộng từ $3.65\text{ cm} - 7.5\text{ cm}$.
     * **1 Cột (4 dòng)**: Khi phương án dài $> 7.5\text{ cm}$.
   - Khử sạch tab thừa cũ mà không làm dính chữ giữa các từ.

3. **Chống xé lẻ câu hỏi qua 2 trang (`keep_with_next`)**:
   - Gán thuộc tính neo giữ thông minh: Word sẽ luôn giữ trọn vẹn cả câu hỏi cùng 4 đáp án trên cùng một trang, triệt tiêu lỗi bị rớt đáp án A ở trang trước, B-C-D ở trang sau.

4. **Đánh số trang động theo Section (`SECTIONPAGES`)**:
   - Phần Đề thi hiển thị chuẩn Bộ GD&ĐT: `Mã đề thi: 101 — Trang 1 / 2`, `Trang 2 / 2`.
   - **Bảng đáp án tách riêng Section 2**: Bắt đầu bằng trang mới hoàn toàn và KHÔNG tính vào tổng số trang của đề thi học sinh.

5. **Thuật toán cân bằng phân bố đáp án (Anti-Bias Balancing)**:
   - Ngăn chặn học sinh đánh lụi 1 cột trúng điểm cao.
   - Tỷ lệ A, B, C, D cân bằng gần như tuyệt đối; chặn trùng quá 2 câu liên tiếp cùng 1 đáp án.

6. **Hệ sinh thái xuất file toàn diện**:
   - Đề thi học sinh (`De_101.docx`, `De_102.docx`...).
   - Đề thi có lời giải chi tiết và gạch chân đáp án đỏ (`De_101_LoiGiai.docx`...).
   - Bảng ma trận đáp án Excel & Sheet đối chiếu câu gốc (`DapAn.xlsx`).
   - File cấu hình chấm thi cho **TnMaker** (`TnMaker_Import.xlsx`), **ZipGrade** (`ZipGrade_Import.csv`), **Azota** (`Azota_Import.txt`).

---

## 🚀 CÀI ĐẶT & YÊU CẦU

Yêu cầu môi trường **Python 3.9+**:

```bash
git clone https://github.com/sanginnova/tron-de-pro.git
cd tron-de-pro
pip install -r requirements.txt
```

---

## 📖 HƯỚNG DẪN SỬ DỤNG

### 1. Sử dụng giao diện dòng lệnh (CLI)

```bash
# Trộn 4 mã đề (101, 102, 103, 104), thời gian 45 phút:
python cli.py "duong_dan_file_de.docx" --codes 4 --start 101 --title "ĐỀ KIỂM TRA GIỮA HỌC KỲ I" --time 45

# Chế độ rút riêng bộ câu cho từng mã đề (Independent Subsets):
python cli.py "duong_dan_file_de.docx" --separate --codes 4
```

### 2. Click chuột 1-Click (Windows)
Chạy file `CHAY_TRON_DE_PRO.bat`, kéo thả file Word đề thi vào cửa sổ cmd và nhấn Enter.

---

## 📁 CẤU TRÚC MÃ NGUỒN

- `builder_docx.py`: Module dựng file Word sư phạm (Hanging Indent 1.6cm, TabStop 1-2-4 cột, Section break, Dynamic PAGE).
- `parser.py`: Bộ phân tích và trích xuất câu hỏi từ file Word, nhận diện 4 phần GDPT 2018, bảo toàn 100% MathType OLE.
- `shuffler.py`: Thuật toán xáo trộn câu hỏi, xáo trộn phương án và cân bằng Anti-bias.
- `exporter_omr.py`: Module xuất file Excel ma trận đáp án và các định dạng nạp app TnMaker/ZipGrade/Azota.
- `engine.py`: Bộ điều phối toàn bộ luồng xử lý (pipeline).
- `cli.py`: Giao diện dòng lệnh.
- `schema.py`: Định nghĩa cấu trúc dữ liệu.

---

## 📄 GIẤY PHÉP (LICENSE)
Phát triển phục vụ công tác giảng dạy Toán học THPT & CĐN. Bản quyền © 2026 ThS. Nguyễn Văn Sang.
