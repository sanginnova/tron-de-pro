# -*- coding: utf-8 -*-
"""
TRỘN ĐỀ PRO (SMART EXAM SHUFFLER & MULTI-VERSION GENERATOR)
Module: cli.py - Giao diện dòng lệnh CLI cho giáo viên
"""

import sys
import os
import argparse

# Đảm bảo mã hóa UTF-8 cho console Windows
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")

# Thêm thư mục hiện tại vào sys.path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from schema import ExamConfig
from engine import TronDeProEngine


def main():
    parser = argparse.ArgumentParser(
        description="TRỘN ĐỀ PRO — Hệ thống trộn đề trắc nghiệm chuẩn Sư phạm & GDPT 2018"
    )
    parser.add_argument(
        "input",
        nargs="?",
        help="Đường dẫn file Word (.docx) ngân hàng câu hỏi gốc",
    )
    parser.add_argument(
        "--codes",
        "-c",
        type=int,
        default=4,
        help="Số lượng mã đề cần tạo (mặc định: 4)",
    )
    parser.add_argument(
        "--start",
        "-s",
        type=int,
        default=101,
        help="Mã đề bắt đầu (mặc định: 101)",
    )
    parser.add_argument(
        "--title",
        "-t",
        default="ĐỀ KIỂM TRA CHUYÊN ĐỀ TOÁN 12",
        help="Tiêu đề kỳ thi / bài kiểm tra",
    )
    parser.add_argument(
        "--school",
        default="TRƯỜNG CAO ĐẲNG NGHỀ SỐ 1 - BQP\nKHOA CƠ BẢN — TỔ TOÁN",
        help="Tên đơn vị trường / khoa / tổ",
    )
    parser.add_argument(
        "--time",
        type=int,
        default=45,
        help="Thời gian làm bài tính bằng phút (mặc định: 45)",
    )
    parser.add_argument(
        "--separate",
        action="store_true",
        help="Bật chế độ 'Mỗi mã đề rút bộ câu riêng' (Independent Question Subsets)",
    )
    parser.add_argument(
        "--no-bias",
        action="store_true",
        help="Tắt thuật toán cân bằng phân bố đáp án A/B/C/D",
    )
    parser.add_argument(
        "--seed",
        type=int,
        default=None,
        help="Mã hạt giống ngẫu nhiên (dùng khi cần tái tạo chính xác bộ đề đã sinh)",
    )
    parser.add_argument(
        "--out",
        "-o",
        default=None,
        help="Thư mục xuất kết quả (mặc định tạo thư mục TronDe_... bên cạnh file nguồn)",
    )

    args = parser.parse_args()

    # Nếu người dùng chạy trực tiếp không truyền tham số, hỗ trợ kéo thả hoặc nhập đường dẫn
    bank_path = args.input
    if not bank_path:
        print("=" * 70)
        print("🎯 HỆ THỐNG TRỘN ĐỀ PRO (THPT & CĐN01 — GDPT 2018)")
        print("=" * 70)
        bank_path = input("👉 Nhập hoặc kéo thả đường dẫn file Word (.docx) ngân hàng vào đây: ").strip('"').strip("'")
        if not bank_path:
            print("❌ Bạn chưa nhập file nguồn. Kết thúc chương trình.")
            return

    # Khởi tạo Config
    config = ExamConfig(
        exam_title=args.title,
        school_header=args.school,
        duration_minutes=args.time,
        num_codes=args.codes,
        code_start=args.start,
        independent_subsets=args.separate,
        anti_bias_balancing=not args.no_bias,
        export_solutions_exam=True,
        export_omr_tnmaker=True,
        export_omr_zipgrade=True,
        export_azota_text=True,
    )

    engine = TronDeProEngine(bank_path, config)
    res = engine.run(output_dir=args.out, seed=args.seed)
    if res.get("success"):
        # Mở thư mục kết quả trên Windows Explorer
        try:
            os.startfile(res["output_dir"])
        except Exception:
            pass


if __name__ == "__main__":
    main()
