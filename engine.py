# -*- coding: utf-8 -*-
"""
TRỘN ĐỀ PRO (SMART EXAM SHUFFLER & MULTI-VERSION GENERATOR)
Module: engine.py - Bộ điều phối tổng thể toàn bộ quy trình trộn đề
"""

import os
import sys
import time
from datetime import datetime
from typing import Optional, Dict, Any

from schema import ExamConfig, PartType, CognitiveLevel
from parser import ExamDocxParser
from shuffler import ExamShuffler
from builder_docx import ExamDocxBuilder
from exporter_omr import ExamOmrExporter


class TronDeProEngine:
    def __init__(self, bank_docx_path: str, config: Optional[ExamConfig] = None):
        self.bank_docx_path = os.path.abspath(bank_docx_path)
        self.config = config or ExamConfig()

    def run(self, output_dir: Optional[str] = None, seed: Optional[int] = None) -> Dict[str, Any]:
        start_time = time.time()
        print(f"\n{'='*70}")
        print(f"🚀 BẮT ĐẦU QUY TRÌNH TRỘN ĐỀ PRO (GDPT 2018 - THPT & CĐN01)")
        print(f"📄 Ngân hàng câu hỏi nguồn: {os.path.basename(self.bank_docx_path)}")
        print(f"{'='*70}\n")

        # 1. Kiểm tra tồn tại file nguồn
        if not os.path.exists(self.bank_docx_path):
            if os.path.exists(self.bank_docx_path + ".docx"):
                self.bank_docx_path = self.bank_docx_path + ".docx"
            elif os.path.exists(self.bank_docx_path + ".DOCX"):
                self.bank_docx_path = self.bank_docx_path + ".DOCX"
            else:
                # Tìm file có tên tương tự trong cùng thư mục
                parent_dir = os.path.dirname(self.bank_docx_path)
                base_query = os.path.basename(self.bank_docx_path).lower()
                matched = None
                if os.path.exists(parent_dir):
                    for f in os.listdir(parent_dir):
                        if f.lower().endswith(".docx") and base_query in f.lower():
                            matched = os.path.join(parent_dir, f)
                            break
                if matched and os.path.exists(matched):
                    print(f"ℹ️ Tự động nhận diện file khớp nhất: {os.path.basename(matched)}")
                    self.bank_docx_path = matched
                else:
                    raise FileNotFoundError(f"Không tìm thấy file nguồn: {self.bank_docx_path}")

        # Tạo thư mục xuất kết quả
        if not output_dir:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(os.path.basename(self.bank_docx_path))[0]
            output_dir = os.path.join(os.path.dirname(self.bank_docx_path), f"TronDe_{base_name}_{ts}")
        os.makedirs(output_dir, exist_ok=True)

        # 2. Phân tích bóc tách ngân hàng
        print("⏳ [1/4] Đang quét và phân tích ngân hàng câu hỏi...")
        parser = ExamDocxParser(self.bank_docx_path)
        questions = parser.parse()

        # Thống kê ngân hàng
        p1_cnt = len([q for q in questions if q.part_type == PartType.PART_1_CHOICE])
        p2_cnt = len([q for q in questions if q.part_type == PartType.PART_2_TRUE_FALSE])
        p3_cnt = len([q for q in questions if q.part_type == PartType.PART_3_SHORT_ANSWER])
        p4_cnt = len([q for q in questions if q.part_type == PartType.PART_4_ESSAY])

        nb_cnt = len([q for q in questions if q.level == CognitiveLevel.NB])
        th_cnt = len([q for q in questions if q.level == CognitiveLevel.TH])
        vd_cnt = len([q for q in questions if q.level == CognitiveLevel.VD])
        vdc_cnt = len([q for q in questions if q.level == CognitiveLevel.VDC])

        print(f"   ✓ Tổng số câu hỏi: {len(questions)} câu.")
        print(f"     - Phần I (4 lựa chọn): {p1_cnt} câu | Phần II (Đúng/Sai): {p2_cnt} câu")
        print(f"     - Phần III (Trả lời ngắn): {p3_cnt} câu | Phần IV (Tự luận): {p4_cnt} câu")
        print(f"     - Phân loại nhận thức: NB: {nb_cnt} | TH: {th_cnt} | VD: {vd_cnt} | VDC: {vdc_cnt}")

        if parser.warnings:
            print(f"   ⚠️ Cảnh báo tự động phát hiện:")
            for w in parser.warnings[:5]:
                print(f"     * {w}")

        # 3. Thực thi thuật toán trộn đề & Cân bằng Anti-Bias
        print(f"\n⏳ [2/4] Đang trộn đề (Số mã: {self.config.num_codes}, Rút riêng: {self.config.independent_subsets}, Anti-bias: {self.config.anti_bias_balancing})...")
        shuffler = ExamShuffler(questions, self.config)
        exams = shuffler.generate_all_codes(seed=seed)
        print(f"   ✓ Đã sinh thành công {len(exams)} mã đề ({', '.join([str(e.code) for e in exams])}).")

        # 4. Xuất các file Word mã đề
        print("\n⏳ [3/4] Đang dựng tài liệu Word Sư phạm (Hanging Indent 1.6cm, TabStop chuẩn)...")
        generated_docx = []
        for exam in exams:
            builder = ExamDocxBuilder(exam, self.config, template_docx_path=self.bank_docx_path)
            # File đề thi học sinh
            path_de = os.path.join(output_dir, f"De_{exam.code}.docx")
            builder.build(path_de, with_solutions=False)
            generated_docx.append(path_de)

            # File đề thi kèm lời giải
            if self.config.export_solutions_exam:
                path_sol = os.path.join(output_dir, f"De_{exam.code}_LoiGiai.docx")
                builder.build(path_sol, with_solutions=True)
                generated_docx.append(path_sol)
            print(f"   ✓ Đã xuất: De_{exam.code}.docx")

        # 5. Xuất bảng đáp án Excel & OMR
        print("\n⏳ [4/4] Đang xuất bảng đáp án Excel đối chiếu và file OMR...")
        exporter = ExamOmrExporter(exams, self.config)
        path_excel = os.path.join(output_dir, "DapAn.xlsx")
        exporter.export_excel(path_excel)
        print(f"   ✓ Đã xuất: DapAn.xlsx (gồm sheet MaTranDapAn và sheet DoiChieu)")

        if self.config.export_omr_tnmaker:
            path_tnmaker = os.path.join(output_dir, "TnMaker_Import.xlsx")
            exporter.export_tnmaker(path_tnmaker)
            print(f"   ✓ Đã xuất: TnMaker_Import.xlsx (import thẳng app TnMaker điện thoại)")

        if self.config.export_omr_zipgrade:
            path_zipgrade = os.path.join(output_dir, "ZipGrade_Import.csv")
            exporter.export_zipgrade(path_zipgrade)
            print(f"   ✓ Đã xuất: ZipGrade_Import.csv (import ZipGrade)")

        if self.config.export_azota_text:
            path_azota = os.path.join(output_dir, "Azota_Import.txt")
            exporter.export_azota_text(path_azota)
            print(f"   ✓ Đã xuất: Azota_Import.txt (text chuẩn nạp Azota)")

        elapsed = round(time.time() - start_time, 2)
        print(f"\n{'='*70}")
        print(f"🎉 HOÀN THÀNH XUẤT SẮC TRONG {elapsed} GIÂY!")
        print(f"📁 Thư mục kết quả: {output_dir}")
        print(f"{'='*70}\n")

        return {
            "success": True,
            "output_dir": output_dir,
            "codes": [e.code for e in exams],
            "docx_files": generated_docx,
            "excel_path": path_excel,
            "elapsed_seconds": elapsed
        }
