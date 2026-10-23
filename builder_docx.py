# -*- coding: utf-8 -*-
"""
TRỘN ĐỀ PRO (SMART EXAM SHUFFLER & MULTI-VERSION GENERATOR)
Module: builder_docx.py - Xuất bản file Word chuẩn Sư phạm (Hanging Indent 1.6cm, TabStop, Keep-with-next)
"""

import os
import re
import copy
from typing import List
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_TAB_ALIGNMENT, WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from schema import Question, Option, PartType, CognitiveLevel, ExamConfig
from shuffler import ShuffledExam


def set_cell_borders(cell, top="single", bottom="single", left="single", right="single", sz="4", val="auto"):
    """Đặt viền mỏng sư phạm cho ô bảng"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="{top}" w:sz="{sz}" w:space="0" w:color="{val}"/>\n'
        f'  <w:left w:val="{left}" w:sz="{sz}" w:space="0" w:color="{val}"/>\n'
        f'  <w:bottom w:val="{bottom}" w:sz="{sz}" w:space="0" w:color="{val}"/>\n'
        f'  <w:right w:val="{right}" w:sz="{sz}" w:space="0" w:color="{val}"/>\n'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)


def make_table_borderless(table, left_indent_cm=1.6):
    """Thiết lập bảng hoàn toàn không viền và thụt lề dóng thẳng trục 1.6cm"""
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)
    if left_indent_cm > 0:
        twips = int(left_indent_cm * 567)
        tblInd = parse_xml(f'<w:tblInd w:w="{twips}" w:type="dxa" {nsdecls("w")}/>')
        tblPr.append(tblInd)


def strip_leading_prefix(elem_or_elems, pattern):
    """Loại bỏ tiền tố (như 'Câu 1.' hoặc 'A.') kể cả khi Word chia nhỏ ký tự thành nhiều run XML"""
    import re
    if isinstance(elem_or_elems, list):
        t_nodes = []
        for el in elem_or_elems:
            t_nodes.extend([c for c in el.iter() if c.tag.endswith('}t') and c.text])
    else:
        t_nodes = [c for c in elem_or_elems.iter() if c.tag.endswith('}t') and c.text]

    if not t_nodes:
        return

    full_text = "".join([n.text for n in t_nodes])
    m = re.match(pattern, full_text)
    if not m:
        return

    chars_to_remove = m.end()
    for n in t_nodes:
        if chars_to_remove <= 0:
            break
        if len(n.text) <= chars_to_remove:
            chars_to_remove -= len(n.text)
            n.text = ""
        else:
            n.text = n.text[chars_to_remove:]
            chars_to_remove = 0


def configure_document_default_font(doc, font_name="Times New Roman", size_pt=12):
    """Cài đặt style Normal mặc định cho toàn bộ tài liệu là Times New Roman 12pt"""
    style = doc.styles["Normal"]
    style.font.name = font_name
    style.font.size = Pt(size_pt)

    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:eastAsia="{font_name}" w:cs="{font_name}"/>')
        rPr.append(rFonts)
    else:
        rFonts.set(qn("w:ascii"), font_name)
        rFonts.set(qn("w:hAnsi"), font_name)
        rFonts.set(qn("w:eastAsia"), font_name)
        rFonts.set(qn("w:cs"), font_name)


def standardize_run_fonts_and_size(elem_or_elems, font_name="Times New Roman", size_pt=12):
    """Đồng bộ hóa font Times New Roman và cỡ chữ 12pt cho toàn bộ các run trong element"""
    sz_val = str(int(size_pt * 2))
    items = elem_or_elems if isinstance(elem_or_elems, list) else [elem_or_elems]
    for el in items:
        for r in el.iter():
            if r.tag.endswith("}r"):
                rPr = r.find(qn("w:rPr"))
                if rPr is None:
                    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')
                    r.insert(0, rPr)

                rFonts = rPr.find(qn("w:rFonts"))
                if rFonts is None:
                    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="{font_name}" w:hAnsi="{font_name}" w:cs="{font_name}"/>')
                    rPr.append(rFonts)
                else:
                    rFonts.set(qn("w:ascii"), font_name)
                    rFonts.set(qn("w:hAnsi"), font_name)
                    rFonts.set(qn("w:cs"), font_name)

                sz = rPr.find(qn("w:sz"))
                if sz is None:
                    sz = parse_xml(f'<w:sz {nsdecls("w")} w:val="{sz_val}"/>')
                    rPr.append(sz)
                else:
                    sz.set(qn("w:val"), sz_val)

                szCs = rPr.find(qn("w:szCs"))
                if szCs is None:
                    szCs = parse_xml(f'<w:szCs {nsdecls("w")} w:val="{sz_val}"/>')
                    rPr.append(szCs)
                else:
                    szCs.set(qn("w:val"), sz_val)


def strip_trailing_tabs_and_spaces(elem_or_elems):
    """Loại bỏ triệt để các thẻ <w:tab/> cũ và khoảng trắng thừa ở cuối cùng của phương án"""
    items = elem_or_elems if isinstance(elem_or_elems, list) else [elem_or_elems]
    for el in items:
        # Xóa các thẻ <w:tab/> cũ
        for tab in list(el.iter()):
            if tab.tag.endswith("}tab"):
                parent = tab.getparent()
                if parent is not None:
                    parent.remove(tab)

    # Chỉ rstrip cho text element cuối cùng có nội dung để tránh mất dấu cách giữa các run
    all_t = []
    for el in items:
        for t in el.iter():
            if t.tag.endswith("}t") and t.text:
                all_t.append(t)
    if all_t:
        all_t[-1].text = all_t[-1].text.rstrip("\t ")


class ExamDocxBuilder:
    def __init__(self, exam: ShuffledExam, config: ExamConfig, template_docx_path: str = None):
        self.exam = exam
        self.config = config
        self.template_docx_path = template_docx_path

    def build(self, output_path: str, with_solutions: bool = False):
        """Xây dựng tài liệu Word hoàn chỉnh cho mã đề (bảo toàn 100% MathType OLE & Image rels)"""
        if self.template_docx_path and os.path.exists(self.template_docx_path):
            doc = docx.Document(self.template_docx_path)
            # Xóa các element trong body ngoại trừ sectPr để giữ nguyên toàn bộ rels, embeddings, và media
            body = doc._body._body
            for child in list(body):
                if not child.tag.endswith("}sectPr"):
                    body.remove(child)
        else:
            doc = docx.Document()

        # Thiết lập Font mặc định Times New Roman 12pt
        configure_document_default_font(doc, "Times New Roman", 12)

        # 1. Page Setup A4
        section = doc.sections[0]
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(1.5)

        # Cài đặt Header & Footer số trang tự động chuẩn Bộ GD&ĐT: Mã đề thi: 101 — Trang 1 / 2
        footer = section.footer
        p_ft = footer.paragraphs[0]
        p_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_ft.paragraph_format.space_before = Pt(0)
        p_ft.paragraph_format.space_after = Pt(0)
        r_ft = p_ft.add_run(f"Mã đề thi: {self.exam.code} — Trang ")
        r_ft.font.name = "Times New Roman"
        r_ft.font.size = Pt(9.5)
        r_ft.font.italic = True

        # Field PAGE (Số trang hiện tại)
        fld_page = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="PAGE"><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i/><w:sz w:val="19"/></w:rPr><w:t>1</w:t></w:r></w:fldSimple>')
        p_ft._p.append(fld_page)

        # Dấu gạch chéo
        r_slash = p_ft.add_run(" / ")
        r_slash.font.name = "Times New Roman"
        r_slash.font.size = Pt(9.5)
        r_slash.font.italic = True

        # Field SECTIONPAGES (Tổng số trang của riêng phần đề thi, KHÔNG tính trang đáp án)
        fld_numpages = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="SECTIONPAGES"><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:i/><w:sz w:val="19"/></w:rPr><w:t>2</w:t></w:r></w:fldSimple>')
        p_ft._p.append(fld_numpages)

        # 2. Bảng Header Hành chính
        self._add_administrative_header(doc)

        # 3. Khung Điểm & Lời phê Sư phạm
        self._add_grade_frame(doc)

        # 4. Nội dung câu hỏi theo 4 Phần
        current_part = None
        for idx, q in enumerate(self.exam.questions, 1):
            if q.part_type != current_part:
                current_part = q.part_type
                self._add_part_heading(doc, current_part)

            self._add_question(doc, idx, q, with_solutions)

        # 5. Bảng đáp án cuối trang (nếu là bản có lời giải hoặc in bảng đáp án)
        if not with_solutions:
            self._add_footer_divider(doc)
            self._add_matrix_answer_table(doc)

        # Lưu file
        doc.save(output_path)

    def _add_administrative_header(self, doc: docx.Document):
        """Bảng Header 2 cột: Đơn vị và Thông tin kỳ thi + Mã đề"""
        table = doc.add_table(rows=1, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        cell_left = table.cell(0, 0)
        cell_right = table.cell(0, 1)
        cell_left.width = Cm(8.0)
        cell_right.width = Cm(9.5)

        # Ô trái: Tên trường
        p_l = cell_left.paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l.paragraph_format.line_spacing = 1.15
        p_l.paragraph_format.space_after = Pt(2)
        r1 = p_l.add_run(self.config.school_header.split("\n")[0] + "\n")
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(10)
        r1.font.bold = True

        r2 = p_l.add_run(self.config.school_header.split("\n")[-1] + "\n")
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(10)
        r2.font.bold = True

        r3 = p_l.add_run("------------------")
        r3.font.name = "Times New Roman"
        r3.font.size = Pt(9)

        # Ô phải: Tên kỳ thi, Mã đề
        p_r = cell_right.paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_r.paragraph_format.line_spacing = 1.15
        p_r.paragraph_format.space_after = Pt(2)
        r_title = p_r.add_run(f"{self.config.exam_title}\n")
        r_title.font.name = "Times New Roman"
        r_title.font.size = Pt(11)
        r_title.font.bold = True

        r_sub = p_r.add_run(f"{self.config.subject} — {self.config.academic_year}\n")
        r_sub.font.name = "Times New Roman"
        r_sub.font.size = Pt(10)

        r_time = p_r.add_run(f"Thời gian làm bài: {self.config.duration_minutes} phút (không kể phát đề)\n")
        r_time.font.name = "Times New Roman"
        r_time.font.size = Pt(9.5)
        r_time.font.italic = True

        # Hộp Mã đề
        r_code = p_r.add_run(f"MÃ ĐỀ THI: {self.exam.code}")
        r_code.font.name = "Times New Roman"
        r_code.font.size = Pt(11)
        r_code.font.bold = True

        # Khung Họ tên học sinh
        p_info = doc.add_paragraph()
        p_info.paragraph_format.space_before = Pt(4)
        p_info.paragraph_format.space_after = Pt(4)
        r_info = p_info.add_run("Họ và tên thí sinh: ................................................................ Lớp: ................. SBD: .................")
        r_info.font.name = "Times New Roman"
        r_info.font.size = Pt(10)
        r_info.font.italic = True

    def _add_grade_frame(self, doc: docx.Document):
        """Khung Điểm & Lời phê Sư phạm"""
        table = doc.add_table(rows=2, cols=2)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False

        c00 = table.cell(0, 0)
        c01 = table.cell(0, 1)
        c10 = table.cell(1, 0)
        c11 = table.cell(1, 1)

        c00.width = Cm(4.0)
        c10.width = Cm(4.0)
        c01.width = Cm(13.5)
        c11.width = Cm(13.5)

        # Tiêu đề
        p00 = c00.paragraphs[0]
        p00.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r00 = p00.add_run("ĐIỂM")
        r00.font.name = "Times New Roman"
        r00.font.size = Pt(10)
        r00.font.bold = True

        p01 = c01.paragraphs[0]
        p01.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r01 = p01.add_run("LỜI PHÊ CỦA THẦY CÔ GIÁO")
        r01.font.name = "Times New Roman"
        r01.font.size = Pt(10)
        r01.font.bold = True

        # Chiều cao ô nhập điểm
        c10.height = Cm(1.2)
        c11.height = Cm(1.2)

        for row in table.rows:
            for cell in row.cells:
                set_cell_borders(cell, sz="6")

        p_spacer = doc.add_paragraph()
        p_spacer.paragraph_format.space_after = Pt(4)

    def _add_part_heading(self, doc: docx.Document, part_type: PartType):
        """Tiêu đề từng phần theo đúng chuẩn GDPT 2018"""
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True

        if part_type == PartType.PART_1_CHOICE:
            title = "PHẦN I. CÂU TRẮC NGHIỆM NHIỀU PHƯƠNG ÁN LỰA CHỌN"
            desc = " (Thí sinh trả lời từ câu 1 đến câu 12. Mỗi câu hỏi thí sinh chỉ chọn một phương án)."
        elif part_type == PartType.PART_2_TRUE_FALSE:
            title = "PHẦN II. CÂU TRẮC NGHIỆM ĐÚNG SAI"
            desc = " (Thí sinh trả lời từ câu 1 đến câu 4. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn Đúng hoặc Sai)."
        elif part_type == PartType.PART_3_SHORT_ANSWER:
            title = "PHẦN III. CÂU TRẮC NGHIỆM TRẢ LỜI NGẮN"
            desc = " (Thí sinh trả lời từ câu 1 đến câu 6)."
        else:
            title = "PHẦN IV. TỰ LUẬN"
            desc = ""

        r1 = p.add_run(title)
        r1.font.name = "Times New Roman"
        r1.font.size = Pt(12)
        r1.font.bold = True

        r2 = p.add_run(desc)
        r2.font.name = "Times New Roman"
        r2.font.size = Pt(11)
        r2.font.italic = True

    def _add_question(self, doc: docx.Document, q_num: int, q: Question, with_solutions: bool):
        """Thêm một câu hỏi với Hanging Indent 1.6cm và TabStop chuẩn"""
        # Paragraph Đề bài
        p_q = doc.add_paragraph()
        INDENT = Cm(1.6)
        p_q.paragraph_format.left_indent = INDENT
        p_q.paragraph_format.first_line_indent = -INDENT
        p_q.paragraph_format.tab_stops.add_tab_stop(INDENT, WD_TAB_ALIGNMENT.LEFT)
        p_q.paragraph_format.space_before = Pt(4)
        p_q.paragraph_format.space_after = Pt(2)
        p_q.paragraph_format.line_spacing = 1.15
        p_q.paragraph_format.keep_with_next = True

        # Run "Câu n:"
        r_num = p_q.add_run(f"Câu {q_num}:")
        r_num.font.name = "Times New Roman"
        r_num.font.size = Pt(12)
        r_num.font.bold = True
        p_q.add_run("\t")

        # Sao chép các phần tử XML đề bài (khử prefix Câu n. cũ và chuẩn hóa Times New Roman 12pt)
        if q.prompt_xml_elements:
            copied_elements = [copy.deepcopy(elem) for elem in q.prompt_xml_elements]
            strip_leading_prefix(copied_elements, r"^\s*Câu\s+\d+(?:\s*\([^\)]+\))?[\.:\s\xa0]*")
            standardize_run_fonts_and_size(copied_elements, "Times New Roman", 12)
            for elem in copied_elements:
                for child in elem:
                    tag = child.tag.split("}")[-1]
                    if tag in ["r", "oMath", "oMathPara"]:
                        p_q._p.append(child)
        else:
            r_txt = p_q.add_run(q.raw_prompt_text)
            r_txt.font.name = "Times New Roman"
            r_txt.font.size = Pt(12)

        # Thêm các Phương án (Nếu là Phần I hoặc Phần II)
        if q.part_type == PartType.PART_1_CHOICE and q.options:
            self._render_part1_options(doc, q.options, with_solutions)
        elif q.part_type == PartType.PART_2_TRUE_FALSE and q.options:
            self._render_part2_options(doc, q.options, with_solutions)

        # Thêm Đáp án cho Phần III trong bản có Lời giải
        if with_solutions and q.part_type == PartType.PART_3_SHORT_ANSWER and q.correct_short_answer:
            p_ans = doc.add_paragraph()
            p_ans.paragraph_format.left_indent = INDENT
            p_ans.paragraph_format.space_before = Pt(2)
            p_ans.paragraph_format.space_after = Pt(2)
            r_ans_lbl = p_ans.add_run("Đáp án: ")
            r_ans_lbl.font.name = "Times New Roman"
            r_ans_lbl.font.size = Pt(10.5)
            r_ans_lbl.font.bold = True
            r_ans_lbl.font.color.rgb = RGBColor(204, 0, 0)
            r_ans_val = p_ans.add_run(q.correct_short_answer)
            r_ans_val.font.name = "Times New Roman"
            r_ans_val.font.size = Pt(10.5)
            r_ans_val.font.bold = True

        # Thêm Lời giải chi tiết (nếu là bản with_solutions)
        if with_solutions and q.solution_xml_elements:
            p_sol = doc.add_paragraph()
            p_sol.paragraph_format.left_indent = INDENT
            p_sol.paragraph_format.space_before = Pt(2)
            p_sol.paragraph_format.space_after = Pt(4)
            r_s = p_sol.add_run("Hướng dẫn giải:")
            r_s.font.name = "Times New Roman"
            r_s.font.size = Pt(10)
            r_s.font.bold = True
            r_s.font.italic = True
            r_s.font.color.rgb = RGBColor(0, 102, 204)

            for sol_elem in q.solution_xml_elements:
                p_custom = doc.add_paragraph()
                p_custom.paragraph_format.left_indent = INDENT
                for child in sol_elem:
                    p_custom._p.append(copy.deepcopy(child))

    def _calculate_option_width_cm(self, opt: Option) -> float:
        """Tính chiều rộng vật lý thực tế của phương án (cm) bao gồm chữ và công thức MathType/OMML"""
        width_cm = 0.5  # Nhãn 'A. ' (in đậm Times New Roman 12pt ~ 0.5cm)

        # Text thông thường (bỏ dấu chấm rỗng nếu là placeholder)
        t = opt.text_preview.strip()
        t_clean = re.sub(r"^\.+$", "", t)
        if t_clean:
            # 1 ký tự Times New Roman 12pt rộng trung bình ~ 0.20 cm
            width_cm += len(t_clean) * 0.20

        # Công thức MathType OLE hoặc Drawing Extent
        shape_pts = []
        if opt.xml_runs:
            for r in opt.xml_runs:
                for s in r.iter():
                    if s.tag.endswith("}shape") and "style" in s.attrib:
                        m = re.search(r"width:\s*([\d\.]+)pt", s.attrib["style"])
                        if m:
                            shape_pts.append(float(m.group(1)))
                    elif s.tag.endswith("}extent"):
                        cx = s.attrib.get("cx")
                        if cx:
                            shape_pts.append(float(cx) / 12700.0)

        # 1 pt = 0.03528 cm
        if shape_pts:
            width_cm += sum(shape_pts) * 0.03528

        return width_cm

    def _render_part1_options(self, doc: docx.Document, options: List[Option], with_solutions: bool):
        """Dàn 4 phương án A, B, C, D thông minh 1 - 2 - 4 cột sử dụng TabStop chuẩn Sư phạm GDPT 2018"""
        INDENT = Cm(1.6)

        # 1. Tính chiều rộng vật lý thực tế của từng phương án (cm)
        widths = [self._calculate_option_width_cm(opt) for opt in options]
        max_w = max(widths) if widths else 0.0

        # TH1: Đáp án ngắn (<= 3.65cm) -> Dàn 4 cột trên 1 dòng duy nhất bằng TabStop (Tiết kiệm tối đa giấy in)
        if max_w <= 3.65 and len(options) == 4:
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = INDENT
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.keep_with_next = False

            # TabStops cho 4 cột: 5.6cm (B), 9.6cm (C), 13.5cm (D)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(5.6), WD_TAB_ALIGNMENT.LEFT)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(9.6), WD_TAB_ALIGNMENT.LEFT)
            p.paragraph_format.tab_stops.add_tab_stop(Cm(13.5), WD_TAB_ALIGNMENT.LEFT)

            for i, opt in enumerate(options):
                self._add_single_option_run(p, opt, with_solutions)
                if i < 3:
                    p.add_run("\t")

        # TH2: Vừa phải (3.65cm < max_w <= 7.5cm) -> Dàn 2 cột x 2 dòng bằng TabStop
        elif max_w <= 7.5 and len(options) == 4:
            # Dòng 1: A và B (B bắt đầu tại TabStop 9.6cm)
            p1 = doc.add_paragraph()
            p1.paragraph_format.left_indent = INDENT
            p1.paragraph_format.first_line_indent = Cm(0)
            p1.paragraph_format.space_before = Pt(0)
            p1.paragraph_format.space_after = Pt(1)
            p1.paragraph_format.line_spacing = 1.15
            p1.paragraph_format.keep_with_next = True
            p1.paragraph_format.tab_stops.add_tab_stop(Cm(9.6), WD_TAB_ALIGNMENT.LEFT)

            self._add_single_option_run(p1, options[0], with_solutions)
            p1.add_run("\t")
            self._add_single_option_run(p1, options[1], with_solutions)

            # Dòng 2: C và D (D bắt đầu tại TabStop 9.6cm)
            p2 = doc.add_paragraph()
            p2.paragraph_format.left_indent = INDENT
            p2.paragraph_format.first_line_indent = Cm(0)
            p2.paragraph_format.space_before = Pt(0)
            p2.paragraph_format.space_after = Pt(2)
            p2.paragraph_format.line_spacing = 1.15
            p2.paragraph_format.keep_with_next = False
            p2.paragraph_format.tab_stops.add_tab_stop(Cm(9.6), WD_TAB_ALIGNMENT.LEFT)

            self._add_single_option_run(p2, options[2], with_solutions)
            p2.add_run("\t")
            self._add_single_option_run(p2, options[3], with_solutions)

        # TH3: Dài (> 7.5cm) hoặc số phương án != 4 -> Dàn 1 cột (mỗi phương án 1 dòng riêng)
        else:
            for idx_opt, opt in enumerate(options):
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = INDENT
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
                p.paragraph_format.keep_with_next = (idx_opt < len(options) - 1)
                self._add_single_option_run(p, opt, with_solutions)

    def _render_part2_options(self, doc: docx.Document, options: List[Option], with_solutions: bool):
        """Dàn các ý a), b), c), d) trắc nghiệm Đúng/Sai"""
        INDENT = Cm(1.6)
        for idx_opt, opt in enumerate(options):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = INDENT
            p.paragraph_format.space_after = Pt(1)
            p.paragraph_format.keep_with_next = (idx_opt < len(options) - 1)

            r_key = p.add_run(f"{opt.key}) ")
            r_key.font.name = "Times New Roman"
            r_key.font.size = Pt(12)
            r_key.font.bold = True
            if with_solutions and opt.is_correct:
                r_key.underline = True
                r_key.font.color.rgb = RGBColor(204, 0, 0)

            # Đính kèm nội dung ý (khử tiền tố a) cũ, khử tab cũ và chuẩn hóa Times New Roman 12pt)
            if opt.xml_runs:
                copied_runs = [copy.deepcopy(r_elem) for r_elem in opt.xml_runs]
                strip_leading_prefix(copied_runs, r"^\s*[a-d]\)[\s\xa0]*")
                strip_trailing_tabs_and_spaces(copied_runs)
                standardize_run_fonts_and_size(copied_runs, "Times New Roman", 12)
                for r_elem in copied_runs:
                    p._p.append(r_elem)
            else:
                r_c = p.add_run(opt.text_preview.rstrip("\t "))
                r_c.font.name = "Times New Roman"
                r_c.font.size = Pt(12)

    def _add_single_option_run(self, p: docx.text.paragraph.Paragraph, opt: Option, with_solutions: bool):
        r_key = p.add_run(f"{opt.key}. ")
        r_key.font.name = "Times New Roman"
        r_key.font.size = Pt(12)
        r_key.font.bold = True
        if with_solutions and opt.is_correct:
            r_key.underline = True
            r_key.font.color.rgb = RGBColor(204, 0, 0)

        # Đính kèm nội dung phương án (khử tiền tố A. cũ, khử tab cũ và chuẩn hóa Times New Roman 12pt)
        if opt.xml_runs:
            copied_runs = [copy.deepcopy(r_elem) for r_elem in opt.xml_runs]
            strip_leading_prefix(copied_runs, r"^\s*[A-D][\.:\)\s\xa0]*")
            strip_trailing_tabs_and_spaces(copied_runs)
            standardize_run_fonts_and_size(copied_runs, "Times New Roman", 12)
            for r_elem in copied_runs:
                p._p.append(r_elem)
        else:
            r_t = p.add_run(opt.text_preview.rstrip("\t "))
            r_t.font.name = "Times New Roman"
            r_t.font.size = Pt(12)

    def _add_footer_divider(self, doc: docx.Document):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(6)
        r = p.add_run("------------------ HẾT ------------------\n(Cán bộ coi thi không giải thích gì thêm)")
        r.font.name = "Times New Roman"
        r.font.size = Pt(9.5)
        r.font.italic = True

    def _add_matrix_answer_table(self, doc: docx.Document):
        """Bảng Ma trận Đáp án chuẩn Bộ GD&ĐT (Tách riêng sang trang mới hoàn toàn)"""
        if not self.exam.answers:
            return

        # 1. TÁCH SANG SECTION 2 ĐỘC LẬP (Không tính vào số trang đề thi)
        sec_ans = doc.add_section(WD_SECTION_START.NEW_PAGE)
        sec_ans.page_width = Cm(21.0)
        sec_ans.page_height = Cm(29.7)
        sec_ans.top_margin = Cm(1.5)
        sec_ans.bottom_margin = Cm(1.5)
        sec_ans.left_margin = Cm(2.0)
        sec_ans.right_margin = Cm(1.5)
        sec_ans.footer.is_linked_to_previous = False

        # Footer riêng cho Bảng đáp án: Không đánh số trang để không nhầm lẫn với đề thi
        p_ans_ft = sec_ans.footer.paragraphs[0]
        p_ans_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_ans_ft.paragraph_format.space_before = Pt(0)
        p_ans_ft.paragraph_format.space_after = Pt(0)
        r_ans_ft = p_ans_ft.add_run(f"Mã đề thi: {self.exam.code} — BẢNG ĐÁP ÁN THAM KHẢO")
        r_ans_ft.font.name = "Times New Roman"
        r_ans_ft.font.size = Pt(9.5)
        r_ans_ft.font.italic = True

        # Tiêu đề bảng đáp án
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_title.paragraph_format.space_before = Pt(10)
        p_title.paragraph_format.space_after = Pt(6)
        p_title.paragraph_format.keep_with_next = True
        r_t = p_title.add_run(f"BẢNG ĐÁP ÁN THAM KHẢO — MÃ ĐỀ: {self.exam.code}")
        r_t.font.name = "Times New Roman"
        r_t.font.size = Pt(13)
        r_t.font.bold = True
        r_t.font.color.rgb = RGBColor(180, 0, 0)

        # 2. PHẦN I: TRẮC NGHIỆM ABCD (Ma trận 10 cột)
        part1_answers = {k: v for k, v in self.exam.answers.items() if len(v) == 1 and v in ["A", "B", "C", "D"]}
        if part1_answers:
            p_p1 = doc.add_paragraph()
            p_p1.paragraph_format.space_before = Pt(6)
            p_p1.paragraph_format.space_after = Pt(3)
            p_p1.paragraph_format.keep_with_next = True
            r_p1 = p_p1.add_run("PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn")
            r_p1.font.name = "Times New Roman"
            r_p1.font.size = Pt(11)
            r_p1.font.bold = True

            total_q = len(part1_answers)
            num_rows = (total_q + 9) // 10

            table = doc.add_table(rows=num_rows * 2, cols=10)
            table.alignment = WD_TABLE_ALIGNMENT.CENTER
            table.autofit = True

            for block_idx in range(num_rows):
                row_q = table.rows[block_idx * 2]
                row_ans = table.rows[block_idx * 2 + 1]

                for c in range(10):
                    q_num = block_idx * 10 + c + 1
                    cell_q = row_q.cells[c]
                    cell_a = row_ans.cells[c]

                    set_cell_borders(cell_q, sz="4")
                    set_cell_borders(cell_a, sz="4")

                    if q_num <= total_q:
                        # Ghi số câu
                        pq = cell_q.paragraphs[0]
                        pq.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        rq = pq.add_run(str(q_num))
                        rq.font.name = "Times New Roman"
                        rq.font.size = Pt(9.5)
                        rq.font.bold = True

                        # Ghi đáp án
                        pa = cell_a.paragraphs[0]
                        pa.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        ra = pa.add_run(part1_answers.get(q_num, ""))
                        ra.font.name = "Times New Roman"
                        ra.font.size = Pt(10)
                        ra.font.bold = True
                        ra.font.color.rgb = RGBColor(0, 102, 204)

        # 3. PHẦN II: TRẮC NGHIỆM ĐÚNG SAI (nếu có)
        part2_items = []
        for q_idx, q in enumerate(self.exam.questions, 1):
            if q.part_type == PartType.PART_2_TRUE_FALSE:
                ans_str = self.exam.answers.get(q_idx, "")
                part2_items.append((q_idx, q, ans_str))

        if part2_items:
            p_p2 = doc.add_paragraph()
            p_p2.paragraph_format.space_before = Pt(8)
            p_p2.paragraph_format.space_after = Pt(3)
            p_p2.paragraph_format.keep_with_next = True
            r_p2 = p_p2.add_run("PHẦN II. Câu trắc nghiệm đúng sai")
            r_p2.font.name = "Times New Roman"
            r_p2.font.size = Pt(11)
            r_p2.font.bold = True

            tbl2 = doc.add_table(rows=1 + len(part2_items), cols=6)
            tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl2.autofit = True

            headers = ["Câu", "Lệnh hỏi", "a)", "b)", "c)", "d)"]
            for col_idx, h in enumerate(headers):
                cell = tbl2.cell(0, col_idx)
                set_cell_borders(cell, sz="4")
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(h)
                r.font.name = "Times New Roman"
                r.font.size = Pt(9.5)
                r.font.bold = True

            for row_idx, (q_num, q, ans_str) in enumerate(part2_items, 1):
                cell_c = tbl2.cell(row_idx, 0)
                cell_l = tbl2.cell(row_idx, 1)
                set_cell_borders(cell_c, sz="4")
                set_cell_borders(cell_l, sz="4")

                p_c = cell_c.paragraphs[0]
                p_c.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p_c.add_run(str(q_num))
                r.font.name = "Times New Roman"
                r.font.size = Pt(9.5)
                r.font.bold = True

                p_l = cell_l.paragraphs[0]
                p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p_l.add_run("Đ/S")
                r.font.name = "Times New Roman"
                r.font.size = Pt(9)

                for sub_idx in range(4):
                    cell_sub = tbl2.cell(row_idx, 2 + sub_idx)
                    set_cell_borders(cell_sub, sz="4")
                    val = ans_str[sub_idx] if sub_idx < len(ans_str) else ""
                    p_s = cell_sub.paragraphs[0]
                    p_s.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    r_s = p_s.add_run(val)
                    r_s.font.name = "Times New Roman"
                    r_s.font.size = Pt(10)
                    r_s.font.bold = True
                    r_s.font.color.rgb = RGBColor(0, 102, 204) if val == "Đ" else RGBColor(204, 0, 0)

        # 4. PHẦN III: TRẮC NGHIỆM TRẢ LỜI NGẮN (nếu có)
        part3_items = []
        for q_idx, q in enumerate(self.exam.questions, 1):
            if q.part_type == PartType.PART_3_SHORT_ANSWER:
                ans_str = self.exam.answers.get(q_idx, "")
                part3_items.append((q_idx, ans_str))

        if part3_items:
            p_p3 = doc.add_paragraph()
            p_p3.paragraph_format.space_before = Pt(8)
            p_p3.paragraph_format.space_after = Pt(3)
            p_p3.paragraph_format.keep_with_next = True
            r_p3 = p_p3.add_run("PHẦN III. Câu trắc nghiệm trả lời ngắn")
            r_p3.font.name = "Times New Roman"
            r_p3.font.size = Pt(11)
            r_p3.font.bold = True

            cols_p3 = len(part3_items)
            tbl3 = doc.add_table(rows=2, cols=cols_p3)
            tbl3.alignment = WD_TABLE_ALIGNMENT.CENTER
            tbl3.autofit = True

            for col_idx, (q_num, ans_val) in enumerate(part3_items):
                c_top = tbl3.cell(0, col_idx)
                c_bot = tbl3.cell(1, col_idx)
                set_cell_borders(c_top, sz="4")
                set_cell_borders(c_bot, sz="4")

                p_top = c_top.paragraphs[0]
                p_top.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_top = p_top.add_run(f"Câu {q_num}")
                r_top.font.name = "Times New Roman"
                r_top.font.size = Pt(9.5)
                r_top.font.bold = True

                p_bot = c_bot.paragraphs[0]
                p_bot.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r_bot = p_bot.add_run(str(ans_val))
                r_bot.font.name = "Times New Roman"
                r_bot.font.size = Pt(10)
                r_bot.font.bold = True
                r_bot.font.color.rgb = RGBColor(0, 102, 204)

