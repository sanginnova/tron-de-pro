# -*- coding: utf-8 -*-
"""
TRỘN ĐỀ PRO (SMART EXAM SHUFFLER & MULTI-VERSION GENERATOR)
Module: exporter_omr.py - Xuất bảng đáp án Excel đa năng, đối chiếu câu gốc và định dạng App TnMaker/ZipGrade/Azota
"""

import os
import csv
from typing import List, Dict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

from shuffler import ShuffledExam
from schema import ExamConfig, PartType


class ExamOmrExporter:
    def __init__(self, exams: List[ShuffledExam], config: ExamConfig):
        self.exams = exams
        self.config = config

    def export_excel(self, output_path: str):
        """Xuất file Excel gồm 2 Sheet: MaTranDapAn và DoiChieu"""
        wb = openpyxl.Workbook()

        # ----------------------------------------------------
        # Sheet 1: MaTranDapAn
        # ----------------------------------------------------
        ws1 = wb.active
        ws1.title = "MaTranDapAn"

        # Định dạng styles
        font_header = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
        fill_header = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
        fill_zebra = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
        align_center = Alignment(horizontal="center", vertical="center")
        border_thin = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC')
        )

        # Tiêu đề bảng
        ws1.merge_cells("A1:Z1")
        ws1["A1"] = f"BẢNG MA TRẬN ĐÁP ÁN TOÀN BỘ CÁC MÃ ĐỀ — {self.config.exam_title.upper()}"
        ws1["A1"].font = Font(name="Times New Roman", size=13, bold=True, color="1B365D")
        ws1["A1"].alignment = align_center

        headers1 = ["Câu"] + [f"Mã {exam.code}" for exam in self.exams]
        ws1.append([])
        ws1.append(headers1)

        for col_idx in range(1, len(headers1) + 1):
            cell = ws1.cell(row=3, column=col_idx)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center

        # Dữ liệu đáp án
        num_questions = max([len(exam.questions) for exam in self.exams] or [0])
        for q_idx in range(1, num_questions + 1):
            row_data = [f"Câu {q_idx}"]
            for exam in self.exams:
                ans = exam.answers.get(q_idx, "")
                row_data.append(ans)
            ws1.append(row_data)

            # Style cho từng dòng
            row_num = 3 + q_idx
            is_even = (q_idx % 2 == 0)
            for col_idx in range(1, len(row_data) + 1):
                cell = ws1.cell(row=row_num, column=col_idx)
                cell.font = Font(name="Times New Roman", size=11, bold=(col_idx == 1 or len(str(cell.value)) == 1))
                cell.alignment = align_center
                cell.border = border_thin
                if is_even:
                    cell.fill = fill_zebra

        # Điều chỉnh độ rộng cột Sheet 1
        ws1.column_dimensions["A"].width = 12
        for col_idx in range(2, len(headers1) + 1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws1.column_dimensions[col_letter].width = 14

        # ----------------------------------------------------
        # Sheet 2: DoiChieu (Đối chiếu câu gốc trong ngân hàng)
        # ----------------------------------------------------
        ws2 = wb.create_sheet(title="DoiChieu")

        ws2.merge_cells("A1:Z1")
        ws2["A1"] = "BẢNG ĐỐI CHIẾU NGUỒN CÂU HỎI TRONG NGÂN HÀNG CHO TỪNG MÃ ĐỀ"
        ws2["A1"].font = Font(name="Times New Roman", size=13, bold=True, color="1B365D")
        ws2["A1"].alignment = align_center

        headers2 = ["Câu trong đề"] + [f"Mã {exam.code} (Câu gốc)" for exam in self.exams]
        ws2.append([])
        ws2.append(headers2)

        for col_idx in range(1, len(headers2) + 1):
            cell = ws2.cell(row=3, column=col_idx)
            cell.font = font_header
            cell.fill = PatternFill(start_color="336699", end_color="336699", fill_type="solid")
            cell.alignment = align_center

        for q_idx in range(1, num_questions + 1):
            row_data = [f"Câu {q_idx}"]
            for exam in self.exams:
                orig_num = exam.source_mapping.get(q_idx, "---")
                row_data.append(f"Gốc: Câu {orig_num}")
            ws2.append(row_data)

            row_num = 3 + q_idx
            for col_idx in range(1, len(row_data) + 1):
                cell = ws2.cell(row=row_num, column=col_idx)
                cell.font = Font(name="Times New Roman", size=10)
                cell.alignment = align_center
                cell.border = border_thin

        ws2.column_dimensions["A"].width = 16
        for col_idx in range(2, len(headers2) + 1):
            col_letter = openpyxl.utils.get_column_letter(col_idx)
            ws2.column_dimensions[col_letter].width = 20

        wb.save(output_path)

    def export_tnmaker(self, output_path: str):
        """Xuất file cấu hình import cho App TnMaker (Excel/CSV)"""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TnMaker_Keys"

        # Định dạng chuẩn TnMaker: Cột A: Mã đề; Cột B: Số câu; Cột C trở đi: Đáp án
        ws.append(["MaDe", "SoCau"] + [f"C{i}" for i in range(1, 101)])

        for exam in self.exams:
            part1_ans = [exam.answers.get(i, "") for i in range(1, len(exam.questions) + 1) if exam.answers.get(i, "") in ["A", "B", "C", "D"]]
            row = [str(exam.code), len(part1_ans)] + part1_ans
            ws.append(row)

        wb.save(output_path)

    def export_zipgrade(self, output_path: str):
        """Xuất file CSV import key chuẩn ZipGrade"""
        with open(output_path, mode="w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            # Header ZipGrade: Key, Question 1, Question 2...
            num_q = max([len(e.questions) for e in self.exams] or [0])
            headers = ["QuizName", "KeyName"] + [f"Q{i}" for i in range(1, num_q + 1)]
            writer.writerow(headers)

            for exam in self.exams:
                row = [self.config.exam_title, str(exam.code)]
                for i in range(1, num_q + 1):
                    ans = exam.answers.get(i, "")
                    if len(ans) == 1 and ans in ["A", "B", "C", "D"]:
                        row.append(ans)
                    else:
                        row.append("")
                writer.writerow(row)

    def export_azota_text(self, output_path: str):
        """Xuất text định dạng chuẩn để dán hoặc upload thẳng lên Azota"""
        with open(output_path, mode="w", encoding="utf-8") as f:
            for exam in self.exams:
                f.write(f"=== MÃ ĐỀ: {exam.code} ===\n\n")
                for idx, q in enumerate(exam.questions, 1):
                    f.write(f"Câu {idx}: {q.raw_prompt_text}\n")
                    if q.part_type == PartType.PART_1_CHOICE:
                        for opt in q.options:
                            prefix = "*" if opt.is_correct else ""
                            f.write(f"{prefix}{opt.key}. {opt.text_preview}\n")
                    f.write("\n")
                f.write("\n" + "="*40 + "\n\n")
