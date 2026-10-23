# -*- coding: utf-8 -*-
"""
TRỘN ĐỀ PRO (SMART EXAM SHUFFLER & MULTI-VERSION GENERATOR)
Module: schema.py - Định nghĩa cấu trúc dữ liệu cho hệ thống trộn đề
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import List, Dict, Optional, Any


class PartType(str, Enum):
    PART_1_CHOICE = "PART_1_CHOICE"             # Trắc nghiệm nhiều lựa chọn A, B, C, D
    PART_2_TRUE_FALSE = "PART_2_TRUE_FALSE"     # Trắc nghiệm Đúng / Sai a, b, c, d
    PART_3_SHORT_ANSWER = "PART_3_SHORT_ANSWER" # Trắc nghiệm trả lời ngắn
    PART_4_ESSAY = "PART_4_ESSAY"               # Tự luận / Câu hỏi mở


class CognitiveLevel(str, Enum):
    NB = "NB"       # Nhận biết
    TH = "TH"       # Thông hiểu
    VD = "VD"       # Vận dụng
    VDC = "VDC"     # Vận dụng cao
    UNKNOWN = "UNKNOWN"


@dataclass
class Option:
    key: str                          # "A", "B", "C", "D" hoặc "a", "b", "c", "d"
    is_correct: bool = False          # Đánh dấu đáp án đúng
    text_preview: str = ""            # Chuỗi text để tính toán độ dài và dàn TabStop
    xml_runs: List[Any] = field(default_factory=list) # Bản sao sâu các thẻ XML của phương án


@dataclass
class Question:
    id: str                           # Mã định danh gốc (ví dụ Q_1, Q_2...)
    original_index: int               # Thứ tự ban đầu trong ngân hàng
    part_type: PartType               # Loại phần
    level: CognitiveLevel = CognitiveLevel.UNKNOWN # Mức độ nhận thức
    topic: str = ""                   # Chuyên đề (ví dụ [CĐ1], [Hàm số]...)
    score: float = 0.25               # Điểm số của câu
    is_locked_order: bool = False     # Khóa vị trí không đảo câu
    is_locked_options: bool = False   # Khóa không đảo thứ tự phương án
    group_id: Optional[str] = None    # ID câu chùm (nếu có)
    prompt_xml_elements: List[Any] = field(default_factory=list) # XML các paragraph đề bài
    options: List[Option] = field(default_factory=list)          # Danh sách phương án
    correct_short_answer: str = ""    # Đáp án cho phần III (Trả lời ngắn)
    solution_xml_elements: List[Any] = field(default_factory=list) # XML lời giải chi tiết
    raw_prompt_text: str = ""         # Text đề bài để preview


@dataclass
class ExamConfig:
    exam_title: str = "ĐỀ KIỂM TRA CHUYÊN ĐỀ TOÁN 12"
    school_header: str = "TRƯỜNG CAO ĐẲNG NGHỀ SỐ 1 - BQP\nKHOA CƠ BẢN — TỔ TOÁN"
    subject: str = "MÔN: TOÁN"
    duration_minutes: int = 45
    academic_year: str = "Năm học 2025 - 2026"
    num_codes: int = 4                # Số lượng mã đề cần sinh (ví dụ 4 mã: 101, 102, 103, 104)
    code_start: int = 101             # Mã đề bắt đầu
    independent_subsets: bool = False # False: Dùng chung bộ câu; True: Mỗi mã rút bộ câu riêng
    matrix_counts: Dict[CognitiveLevel, int] = field(default_factory=lambda: {
        CognitiveLevel.NB: 10,
        CognitiveLevel.TH: 8,
        CognitiveLevel.VD: 4,
        CognitiveLevel.VDC: 3
    })                                # Ma trận số câu rút khi bật independent_subsets
    anti_bias_balancing: bool = True  # Cân bằng phân bố đáp án A/B/C/D
    auto_fit_pages: bool = True       # Tự tối ưu trang in vừa khít chẵn trang
    export_solutions_exam: bool = True # Xuất file đề kèm lời giải
    export_omr_tnmaker: bool = True   # Xuất file import app TnMaker
    export_omr_zipgrade: bool = True  # Xuất file import ZipGrade
    export_azota_text: bool = True    # Xuất text import Azota
