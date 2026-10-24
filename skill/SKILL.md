---
name: tron-de-pro
description: >-
  Kỹ năng và Sub-agent chuyên nghiệp tự động hóa Trộn đề thi trắc nghiệm (THPT, THCS, CĐN01) từ file Word (.docx) chuẩn GDPT 2018 (Phần I: ABCD, Phần II: Đúng/Sai, Phần III: Trả lời ngắn, Phần IV: Tự luận). Hỗ trợ rút ngẫu nhiên bộ câu riêng cho từng mã đề theo ma trận NB/TH/VD/VDC hoặc dùng chung bộ câu; bảo toàn 100% công thức MathType/OMML Cambria Math và hình ảnh; tự động cân bằng phân bố đáp án Anti-Bias; dóng lề treo Hanging Indent 1.6cm thẳng trục; dàn TabStop thông minh 1-2-4 cột theo kích thước vật lý cm; đánh số trang động PAGE/NUMPAGES; chống xé câu hỏi qua 2 trang; tách riêng trang đáp án độc lập full 3 phần; và xuất trọn bộ file Word đề thi, đề kèm lời giải, Excel ma trận đáp án, sheet đối chiếu câu gốc, và file cấu hình chấm thi cho TnMaker/ZipGrade/Azota. Kích hoạt khi người dùng gõ /tron-de hoặc yêu cầu "trộn đề", "trộn ngân hàng câu hỏi", "sinh các mã đề thi".
---

# SUB-AGENT CHUYÊN TRỘN ĐỀ TOÁN CHUYÊN NGHIỆP (TRỘN ĐỀ PRO)
> **Đơn vị chủ quản**: ThS. Nguyễn Văn Sang - Khoa Cơ bản - Trường Cao đẳng Nghề số 1 - BQP.  
> **Lệnh tắt kích hoạt**: `/tron-de` hoặc `/tron-de-pro`.

---

## 1. MỤC TIÊU VÀ SỨ MỆNH
Sub-agent này giải quyết triệt để các hạn chế của các phần mềm trộn đề truyền thống (như vỡ công thức MathType/Equation, lệch đáp án, câu hỏi bị chia cắt vụn qua 2 trang, đáp án A/B/C/D bị lệch tỷ lệ, hoặc không tích hợp app chấm trắc nghiệm điện thoại).

Hệ thống hỗ trợ đầy đủ **4 dạng thức câu hỏi GDPT 2018**:
1. **Phần I**: Câu trắc nghiệm 4 phương án lựa chọn ($A, B, C, D$).
2. **Phần II**: Câu trắc nghiệm Đúng / Sai ($a, b, c, d$).
3. **Phần III**: Câu trắc nghiệm Trả lời ngắn.
4. **Phần IV**: Câu hỏi Tự luận / Câu chùm ngữ cảnh chung `[NHÓM] ... [HẾT NHÓM]`.

---

## 2. NGUYÊN TẮC VÀNG VÀ CÔNG NGHỆ CỐT LÕI (ĐÃ ĐƯỢC NÂNG CẤP HOÀN TOÀN)

1. **Bảo tồn nguyên vẹn 100% MathType OLE & Image (Zero Loss via Template Cloning)**:
   - Thay vì tạo file Word trắng (làm mất toàn bộ `relationships`, `embeddings`, `media`), engine khởi tạo trực tiếp từ file DOCX gốc, chỉ xóa nội dung `<w:body>` (giữ nguyên thẻ `<w:sectPr>` và các ID liên kết OLE).
   - Công thức MathType hiển thị sắc nét, mở lên chỉnh sửa trực tiếp bằng MathType bình thường.

2. **Chuẩn hóa Typography Sư phạm**:
   - **Font mặc định**: `Times New Roman`, cỡ chữ chuẩn `12pt` cho toàn bộ các run và văn bản.
   - **Lề trang A4**: Trái 2.0 cm, Phải 1.5 cm, Trên 1.5 cm, Dưới 1.5 cm (chiều rộng khả dụng: 17.5 cm).
   - **Hanging Indent 1.6 cm**: Thụt lề treo chuẩn trục cho cả câu hỏi và các phương án.

3. **Dàn TabStop Thông minh 1 – 2 – 4 cột theo kích thước vật lý thực tế (cm & pt)**:
   - Đọc trực tiếp thuộc tính `style="width: ...pt"` của thẻ hình vector MathType `<v:shape>` và độ dài văn bản để tính kích thước vật lý chuẩn.
   - **4 Cột (1 dòng)**: Áp dụng khi độ rộng tối đa $\le 3.65\text{ cm}$. Đặt 3 TabStops cố định tại `5.6cm` (B), `9.6cm` (C), `13.5cm` (D). Mỗi cột rộng đều ~4.0cm, tiết kiệm tối đa giấy in.
   - **2 Cột (2 dòng)**: Áp dụng khi độ rộng trong khoảng $3.65\text{ cm} - 7.5\text{ cm}$. Dòng 1 gồm A và B (nhảy tab tới `9.6cm`), Dòng 2 gồm C và D (nhảy tab tới `9.6cm`). Mỗi cột rộng đều ~8.0cm.
   - **1 Cột (4 dòng)**: Áp dụng khi độ rộng $> 7.5\text{ cm}$ (mệnh đề dài, công thức cồng kềnh). Mỗi phương án nằm trên 1 dòng riêng tại lề treo `1.6cm`.
   - **Khử sạch thẻ `<w:tab/>` và khoảng trắng thừa** ở cuối phương án mà không làm mất khoảng cách giữa các từ.

4. **Đánh số trang động tự động (PAGE / NUMPAGES)**:
   - Định dạng chân trang (Footer) chuẩn Bộ GD&ĐT: `Mã đề thi: [Mã] — Trang {PAGE} / {NUMPAGES}`.
   - Tự động nhảy số trang theo thời gian thực và tự tính tổng số trang khi in ấn hoặc xuất PDF.

5. **Chống xé lẻ câu hỏi qua 2 trang (`keep_with_next` thông minh)**:
   - Gán `keep_with_next = True` cho đề bài và các phương án đầu ($A, B, C$).
   - Gán `keep_with_next = False` cho phương án cuối cùng ($D$). Nếu trang giấy không đủ diện tích, Word sẽ tự động chuyển trọn vẹn cả câu hỏi cùng 4 đáp án sang trang sau, không bao giờ bị rớt lẻ loi $A$ ở trang trên, $B, C, D$ ở trang dưới.

6. **Tách riêng Bảng đáp án sang trang mới độc lập (Full 3 Phần)**:
   - Đề thi của học sinh kết thúc trang trọng tại `------------------ HẾT ------------------`.
   - Ngắt trang `Page Break` tự động đẩy Bảng đáp án sang một trang mới hoàn toàn, gồm:
     * **Phần I**: Bảng Ma trận 10 cột trắc nghiệm ABCD.
     * **Phần II**: Bảng tra kết quả trắc nghiệm Đúng / Sai ($a, b, c, d$ phân biệt xanh `Đ` và đỏ `S`).
     * **Phần III**: Bảng kết quả trả lời ngắn.

7. **Thuật toán Cân bằng Phân bố Đáp án (Anti-Bias Shuffling Engine)**:
   - Triệt tiêu hoàn toàn hiện tượng học sinh "đánh lụi 1 cột trúng điểm cao":
     * Tỷ lệ các chữ cái $A, B, C, D$ trong toàn đề được cân bằng đều nhau (chênh lệch tối đa $\le 1 - 2$ câu).
     * Ngăn chặn triệt để tình trạng **quá 2 câu liên tiếp trùng một chữ cái**.

8. **Hai chế độ Trộn linh hoạt**:
   - **Chế độ 1 (Dùng chung bộ câu)**: Mọi mã đề (ví dụ 101, 102, 103, 104) dùng chung một danh sách câu hỏi, chỉ đảo thứ tự câu và đảo vị trí các phương án.
   - **Chế độ 2 (Mỗi mã rút bộ câu riêng - Independent Subsets)**: Mỗi mã rút ngẫu nhiên bộ câu khác nhau từ ngân hàng theo ma trận số lượng câu của từng mức độ nhận thức (`[NB]`, `[TH]`, `[VD]`, `[VDC]`).

9. **Tự động phát hiện phương án logic phụ thuộc**:
   - Tự động nhận diện các phương án như *"Cả A và B đều đúng"*, *"Tất cả các đáp án trên"*, *"Không có đáp án nào"* để **tự động bật cờ Khóa Đảo Phương Án (`[KĐ]`)**, cảnh báo ngay cho giáo viên để không làm câu hỏi bị sai logic.

10. **Tự động vẽ hình minh họa TikZ 300 DPI cho Lời giải chi tiết**:
    - Khi bóc tách ngân hàng câu hỏi, Sub-agent tự động quét và nhận diện các câu hỏi hình học phẳng, hình học không gian, lượng giác, bài toán thực tế (đo chiều cao tháp, hải đăng), hệ trục tọa độ Oxy, biểu diễn miền nghiệm bất phương trình, đồ thị hàm số và bảng biến thiên.
    - Nếu câu hỏi hoặc phần lời giải trong file gốc thiếu hình minh họa, hệ thống tự động sinh mã TikZ chuẩn Sư phạm theo quy chuẩn `ve-hinh-tikz`, gọi trình biên dịch `pdflatex` + `pymupdf` kết xuất thành ảnh PNG vector 300 DPI siêu nét và nhúng trực tiếp căn giữa vào mục Hướng dẫn giải trong các file `De_xxx_LoiGiai.docx`.

---

## 3. BỘ FILE KẾT QUẢ XUẤT RA (OUTPUT ECOSYSTEM)

Mỗi lần chạy sẽ tự động tạo thư mục `TronDe_<TênFile>_<ThờiGian>/` chứa đầy đủ:

| Tên File | Mô tả chi tiết |
| :--- | :--- |
| `De_101.docx`, `De_102.docx`... | Đề thi phát cho học sinh: có Header 2 cột Trường Cao đẳng Nghề số 1 - BQP, khung điểm sư phạm, khung mã đề, dóng lề chuẩn 1.6cm, số trang động, bảng đáp án tách riêng trang 3. |
| `De_101_LoiGiai.docx`... | Đề thi có **gạch chân đáp án đúng màu đỏ** + đính kèm **Hướng dẫn giải chi tiết** ngay dưới từng câu để giáo viên chữa bài. |
| `DapAn.xlsx` | File Excel 2 sheet: <br>• Sheet `MaTranDapAn`: Ma trận đáp án toàn bộ mã đề.<br>• Sheet `DoiChieu`: Bảng tra cứu đối chiếu câu trong mã đề tương ứng câu gốc nào trong ngân hàng. |
| `TnMaker_Import.xlsx` | File Excel cấu hình chuẩn để nạp trực tiếp vào **App TnMaker** trên điện thoại (chấm điểm bằng camera). |
| `ZipGrade_Import.csv` | File CSV cấu hình chuẩn để import vào **ZipGrade**. |
| `Azota_Import.txt` | File text chuẩn cú pháp để tải lên **Azota** thi online. |

---

## 4. HƯỚNG DẪN THỰC THI

### Cách 1: Chạy trực tiếp từ Chat với AI (Sub-agent)
Chỉ cần gửi đường dẫn file Word hoặc nói:
> "Trộn cho tôi file `C:\Users\ADMIN\Downloads\de 1.docx` thành 4 mã đề (101-104), đề 45 phút, trường CĐN số 1 BQP."

AI sẽ tự động gọi engine thực thi và trả về link mở các file thành phẩm.

### Cách 2: Chạy qua dòng lệnh Terminal / PowerShell
```powershell
python "d:\TOAN\05_CONG_CU_VA_ENGINE\tron_de_pro\cli.py" "d:\TOAN\NganHang.docx" --codes 4 --start 101 --title "ĐỀ KIỂM TRA GIỮA KỲ I" --time 45
```

Nếu muốn bật chế độ rút ngẫu nhiên bộ câu riêng cho từng mã:
```powershell
python "d:\TOAN\05_CONG_CU_VA_ENGINE\tron_de_pro\cli.py" "d:\TOAN\NganHang.docx" --separate --codes 4
```

### Cách 3: Click chuột qua File Batch
Nhấp đúp chuột vào file [CHAY_TRON_DE_PRO.bat](file:///d:/TOAN/CHAY_TRON_DE_PRO.bat) ngoài thư mục gốc `D:\TOAN\`, kéo thả file Word ngân hàng vào cửa sổ và nhấn Enter.
