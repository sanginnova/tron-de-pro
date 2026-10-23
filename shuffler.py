# -*- coding: utf-8 -*-
"""
TRỘN ĐỀ PRO (SMART EXAM SHUFFLER & MULTI-VERSION GENERATOR)
Module: shuffler.py - Thuật toán trộn đề, rút câu ma trận & Cân bằng phân bố đáp án (Anti-Bias)
"""

import random
import copy
from typing import List, Dict, Tuple, Optional
from collections import Counter

from schema import Question, Option, PartType, CognitiveLevel, ExamConfig


class ShuffledExam:
    def __init__(self, code: int):
        self.code = code
        self.questions: List[Question] = []
        self.answers: Dict[int, str] = {}           # Thứ tự câu -> Đáp án đúng ("A", "ĐSĐS", "3.5"...)
        self.source_mapping: Dict[int, int] = {}    # Thứ tự câu trong mã đề -> Thứ tự câu gốc trong ngân hàng


class ExamShuffler:
    def __init__(self, bank_questions: List[Question], config: ExamConfig):
        self.bank_questions = bank_questions
        self.config = config

    def generate_all_codes(self, seed: Optional[int] = None) -> List[ShuffledExam]:
        """Sinh tất cả các mã đề theo cấu hình"""
        if seed is not None:
            random.seed(seed)

        exams = []
        for i in range(self.config.num_codes):
            code_num = self.config.code_start + i
            exam = self._generate_single_code(code_num)
            exams.append(exam)
        return exams

    def _generate_single_code(self, code_num: int) -> ShuffledExam:
        exam = ShuffledExam(code_num)

        # Bước 1: Chọn bộ câu hỏi (Bộ câu chung hoặc Rút câu riêng theo ma trận)
        if self.config.independent_subsets:
            selected_questions = self._extract_subset_by_matrix()
        else:
            selected_questions = copy.deepcopy(self.bank_questions)

        # Bước 2: Tách theo từng phần
        part_1 = [q for q in selected_questions if q.part_type == PartType.PART_1_CHOICE]
        part_2 = [q for q in selected_questions if q.part_type == PartType.PART_2_TRUE_FALSE]
        part_3 = [q for q in selected_questions if q.part_type == PartType.PART_3_SHORT_ANSWER]
        part_4 = [q for q in selected_questions if q.part_type == PartType.PART_4_ESSAY]

        # Bước 3: Đảo thứ tự câu trong từng phần (bảo toàn câu khóa vị trí và nhóm câu chùm)
        shuffled_p1 = self._shuffle_questions_in_part(part_1)
        shuffled_p2 = self._shuffle_questions_in_part(part_2)
        shuffled_p3 = self._shuffle_questions_in_part(part_3)
        shuffled_p4 = self._shuffle_questions_in_part(part_4)

        # Bước 4: Đảo phương án với thuật toán cân bằng phân bố Anti-Bias
        if self.config.anti_bias_balancing:
            self._balance_and_shuffle_part1_options(shuffled_p1)
        else:
            self._simple_shuffle_part1_options(shuffled_p1)

        # Gộp tất cả các phần lại theo đúng thứ tự GDPT 2018
        final_questions = shuffled_p1 + shuffled_p2 + shuffled_p3 + shuffled_p4

        # Bước 5: Đánh số lại từ 1..N và ghi nhận đáp án
        for idx, q in enumerate(final_questions, 1):
            exam.source_mapping[idx] = q.original_index
            if q.part_type == PartType.PART_1_CHOICE:
                correct_opt = next((opt.key for opt in q.options if opt.is_correct), "?")
                exam.answers[idx] = correct_opt
            elif q.part_type == PartType.PART_2_TRUE_FALSE:
                # Chuỗi Đúng/Sai ví dụ: Đ - S - Đ - S
                tf_str = "".join(["Đ" if opt.is_correct else "S" for opt in q.options])
                exam.answers[idx] = tf_str
            elif q.part_type == PartType.PART_3_SHORT_ANSWER:
                exam.answers[idx] = q.correct_short_answer or "---"
            else:
                exam.answers[idx] = "Tự luận"

        exam.questions = final_questions
        return exam

    def _extract_subset_by_matrix(self) -> List[Question]:
        """Rút ngẫu nhiên bộ câu hỏi riêng theo ma trận mức độ nhận thức"""
        extracted = []
        for level, count in self.config.matrix_counts.items():
            pool = [q for q in self.bank_questions if q.level == level]
            if len(pool) >= count:
                chosen = random.sample(pool, count)
            else:
                chosen = pool[:]
            extracted.extend(copy.deepcopy(chosen))
        return extracted

    def _shuffle_questions_in_part(self, questions: List[Question]) -> List[Question]:
        """Đảo thứ tự câu hỏi trong một phần, giữ cố định câu [KHÓA] và nhóm câu chùm"""
        if not questions:
            return []

        # Gom các câu chùm lại thành từng block
        blocks = []
        i = 0
        while i < len(questions):
            q = questions[i]
            if q.group_id:
                # Gom toàn bộ câu cùng group_id
                grp = [q]
                j = i + 1
                while j < len(questions) and questions[j].group_id == q.group_id:
                    grp.append(questions[j])
                    j += 1
                blocks.append({"type": "group", "items": grp, "is_locked": any(item.is_locked_order for item in grp)})
                i = j
            else:
                blocks.append({"type": "single", "items": [q], "is_locked": q.is_locked_order})
                i += 1

        # Tách các block tự do và block bị khóa
        unlocked_blocks = [b for b in blocks if not b["is_locked"]]
        random.shuffle(unlocked_blocks)

        # Lắp ghép lại vào vị trí ban đầu
        result_blocks = []
        unlocked_idx = 0
        for b in blocks:
            if b["is_locked"]:
                result_blocks.append(b)
            else:
                result_blocks.append(unlocked_blocks[unlocked_idx])
                unlocked_idx += 1

        # Flatten kết quả
        final_list = []
        for b in result_blocks:
            final_list.extend(b["items"])
        return final_list

    def _balance_and_shuffle_part1_options(self, questions: List[Question]):
        """
        Thuật toán Cân bằng Anti-Bias:
        1. Duy trì số lượng A, B, C, D xấp xỉ nhau trong toàn bộ đề.
        2. Ngăn chặn triệt để >2 câu liên tiếp trùng cùng một chữ cái đáp án.
        """
        keys = ["A", "B", "C", "D"]
        counts = {k: 0 for k in keys}
        recent_answers = []

        for q in questions:
            if len(q.options) != 4 or q.is_locked_options:
                # Nếu câu bị khóa đảo hoặc không đủ 4 phương án thì giữ nguyên
                for idx, opt in enumerate(q.options):
                    opt.key = keys[idx]
                    if opt.is_correct:
                        counts[opt.key] = counts.get(opt.key, 0) + 1
                        recent_answers.append(opt.key)
                continue

            # Tìm phương án đúng ban đầu
            correct_opt = next((opt for opt in q.options if opt.is_correct), None)
            other_opts = [opt for opt in q.options if not opt.is_correct]

            # Xác định các chữ cái bị cấm (nếu 2 câu liền kề trước đó đã trùng nhau)
            forbidden_key = None
            if len(recent_answers) >= 2 and recent_answers[-1] == recent_answers[-2]:
                forbidden_key = recent_answers[-1]

            # Chọn chữ cái cho phương án đúng sao cho có tần số xuất hiện thấp nhất hiện tại
            available_keys = [k for k in keys if k != forbidden_key]
            min_count = min(counts[k] for k in available_keys)
            candidates = [k for k in available_keys if counts[k] == min_count]
            chosen_key = random.choice(candidates)

            # Gán vị trí cho phương án đúng và đảo ngẫu nhiên 3 phương án còn lại
            remaining_keys = [k for k in keys if k != chosen_key]
            random.shuffle(other_opts)

            new_options = [None] * 4
            chosen_index = keys.index(chosen_key)
            correct_opt.key = chosen_key
            new_options[chosen_index] = correct_opt

            other_idx = 0
            for i in range(4):
                if new_options[i] is None:
                    opt = other_opts[other_idx]
                    opt.key = keys[i]
                    new_options[i] = opt
                    other_idx += 1

            q.options = new_options
            counts[chosen_key] += 1
            recent_answers.append(chosen_key)

    def _simple_shuffle_part1_options(self, questions: List[Question]):
        """Đảo ngẫu nhiên phương án thông thường không cân bằng"""
        keys = ["A", "B", "C", "D"]
        for q in questions:
            if len(q.options) == 4 and not q.is_locked_options:
                random.shuffle(q.options)
            for idx, opt in enumerate(q.options):
                if idx < len(keys):
                    opt.key = keys[idx]
