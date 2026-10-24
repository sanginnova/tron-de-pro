# -*- coding: utf-8 -*-
"""
TRỘN ĐỀ PRO - MODULE TIKZ HANDLER
Tự động sinh và biên dịch hình vẽ TikZ 300 DPI cho Lời giải chi tiết
"""

import os
import sys
import re
import base64
from typing import List, Optional

# Nạp compiler từ tikz_studio
TIKZ_STUDIO_DIR = r"D:\TOAN\05_CONG_CU_VA_ENGINE\tikz_studio"
if TIKZ_STUDIO_DIR not in sys.path:
    sys.path.insert(0, TIKZ_STUDIO_DIR)

try:
    from compiler import compile_tikz
except ImportError:
    compile_tikz = None

from schema import Question, PartType


def render_tikz_code_to_png(tikz_code: str, output_png_path: str, dpi: int = 300) -> bool:
    """Biên dịch mã TikZ sang ảnh PNG chất lượng cao 300 DPI"""
    if not compile_tikz:
        return False
    try:
        res = compile_tikz(tikz_code, dpi=dpi)
        if res.get("success") and "image_base64" in res:
            b64_data = res["image_base64"].split(",")[1]
            os.makedirs(os.path.dirname(os.path.abspath(output_png_path)), exist_ok=True)
            with open(output_png_path, "wb") as f:
                f.write(base64.b64decode(b64_data))
            return True
        else:
            print(f"⚠️ Lỗi biên dịch TikZ: {res.get('error')}")
            return False
    except Exception as e:
        print(f"⚠️ Ngoại lệ biên dịch TikZ: {e}")
        return False


def generate_tikz_for_bpt_x_3y_ge_6() -> str:
    """TikZ biểu diễn miền nghiệm của x + 3y >= 6"""
    return r"""\begin{tikzpicture}[>=stealth,scale=0.8,font=\footnotesize,line join=round,line cap=round]
    \draw[->,thick] (-1.5,0) -- (7.5,0) node[below]{$x$};
    \draw[->,thick] (0,-1.5) -- (0,4.5) node[left]{$y$};
    \node[below left] at (0,0) {$O$};
    
    % Miền gạch sọc: nửa mặt phẳng x + 3y < 6 (chứa O(0,0) bị gạch)
    \begin{scope}
        \clip (-1.5,-1.5) rectangle (7.5,4.5);
        \fill[pattern=north east lines, pattern color=gray!60] (-2, 2.666) -- (8, -0.666) -- (8, -2) -- (-2, -2) -- cycle;
    \end{scope}
    
    % Đường thẳng d: x + 3y = 6
    \draw[thick,blue,domain=-1:7.2,samples=2] plot (\x, {-(1/3)*\x + 2});
    \node[blue, above] at (3.5, 1.3) {$d: x+3y=6$};
    
    % Điểm đặc biệt
    \fill[red] (6,0) circle (1.5pt) node[below] {$6$};
    \fill[red] (0,2) circle (1.5pt) node[left] {$2$};
\end{tikzpicture}"""


def generate_tikz_for_bpt_2x_y_le_2() -> str:
    """TikZ biểu diễn miền nghiệm của 2x - y <= 2"""
    return r"""\begin{tikzpicture}[>=stealth,scale=0.8,font=\footnotesize,line join=round,line cap=round]
    \draw[->,thick] (-2,0) -- (4,0) node[below]{$x$};
    \draw[->,thick] (0,-3) -- (0,4) node[left]{$y$};
    \node[below left] at (0,0) {$O$};
    
    % Miền gạch sọc: 2x - y > 2 (nửa mặt phẳng không chứa O bị gạch)
    \begin{scope}
        \clip (-2,-3) rectangle (4,4);
        \fill[pattern=north east lines, pattern color=gray!60] (-2, -6) -- (4, 6) -- (4, -6) -- cycle;
    \end{scope}
    
    % Đường thẳng d: 2x - y = 2
    \draw[thick,blue,domain=-0.5:2.8,samples=2] plot (\x, {2*\x - 2});
    \node[blue, above left] at (2, 2.2) {$d: 2x-y=2$};
    
    % Điểm đặc biệt
    \fill[red] (1,0) circle (1.5pt) node[below left] {$1$};
    \fill[red] (0,-2) circle (1.5pt) node[left] {$-2$};
\end{tikzpicture}"""


def generate_tikz_for_tower_cd_24m() -> str:
    """TikZ cho bài toán đo chiều cao tháp CD (AB = 24m, góc CAD = 63, CBD = 48)"""
    return r"""\begin{tikzpicture}[>=stealth,scale=1.0,font=\footnotesize,line join=round,line cap=round]
    \coordinate (C) at (5.0, 0);
    \coordinate (D) at (5.0, 4.0);
    \coordinate (A) at (2.96, 0);
    \coordinate (B) at (1.4, 0);
    \coordinate (GroundLeft) at (0.5, 0);
    \coordinate (GroundRight) at (5.8, 0);
    
    % Mặt đất
    \draw[thick] (GroundLeft) -- (GroundRight);
    
    % Tháp CD
    \draw[very thick, brown!70!black] (4.9, 0) -- (4.95, 4.0) -- (5.05, 4.0) -- (5.1, 0) -- cycle;
    \draw[very thick, red!80!black] (C) -- (D) node[midway, right] {$h = CD$};
    
    % Tia ngắm
    \draw[thick, blue] (B) -- (D);
    \draw[thick, blue] (A) -- (D);
    
    % Góc vuông tại C
    \draw (C) rectangle ++(-0.25, 0.25);
    
    % Góc CAD = 63 độ
    \draw[teal, thick] (A) ++(0.5, 0) arc (0:63:0.5);
    \node[teal, above right] at (2.85, 0.15) {$63^\circ$};
    
    % Góc CBD = 48 độ
    \draw[magenta, thick] (B) ++(0.6, 0) arc (0:48:0.6);
    \node[magenta, above right] at (1.35, 0.15) {$48^\circ$};
    
    % Đoạn AB = 24m
    \draw[<->, thick] (1.4, -0.6) -- (2.96, -0.6) node[midway, below] {$AB = 24\text{ m}$};
    \draw[dashed, gray] (B) -- (1.4, -0.65);
    \draw[dashed, gray] (A) -- (2.96, -0.65);
    
    % Tên điểm
    \fill[black] (B) circle (1.5pt) node[below=2pt] {$B$};
    \fill[black] (A) circle (1.5pt) node[below=2pt] {$A$};
    \fill[black] (C) circle (1.5pt) node[below=2pt] {$C$};
    \fill[black] (D) circle (1.5pt) node[above] {$D$};
\end{tikzpicture}"""


def generate_tikz_for_tower_ab_500m() -> str:
    """TikZ cho bài toán đo chiều cao tháp AB (MN = 500m, góc AMB = 30, ANB = 45)"""
    return r"""\begin{tikzpicture}[>=stealth,scale=1.0,font=\footnotesize,line join=round,line cap=round]
    \coordinate (A) at (5.0, 0);
    \coordinate (B) at (5.0, 3.5);
    \coordinate (N) at (1.5, 0);
    \coordinate (M) at (-1.0, 0);
    
    \draw[thick] (-1.5, 0) -- (5.6, 0);
    \draw[very thick, red!80!black] (A) -- (B) node[midway, right] {$h = AB$};
    
    \draw[thick, blue] (M) -- (B);
    \draw[thick, blue] (N) -- (B);
    
    \draw (A) rectangle ++(-0.25, 0.25);
    \draw[magenta, thick] (M) ++(0.7, 0) arc (0:30:0.7);
    \node[magenta, above right] at (-1.0, 0.15) {$30^\circ$};
    
    \draw[teal, thick] (N) ++(0.5, 0) arc (0:45:0.5);
    \node[teal, above right] at (1.5, 0.15) {$45^\circ$};
    
    \draw[<->, thick] (-1.0, -0.5) -- (1.5, -0.5) node[midway, below] {$MN = 500\text{ m}$};
    \draw[dashed, gray] (M) -- (-1.0, -0.55);
    \draw[dashed, gray] (N) -- (1.5, -0.55);
    
    \fill[black] (M) circle (1.5pt) node[below=2pt] {$M$};
    \fill[black] (N) circle (1.5pt) node[below=2pt] {$N$};
    \fill[black] (A) circle (1.5pt) node[below=2pt] {$A$};
    \fill[black] (B) circle (1.5pt) node[above] {$B$};
\end{tikzpicture}"""


def auto_generate_tikz_for_exam(questions: List[Question], cache_dir: str):
    """Quét toàn bộ câu hỏi và tự động vẽ TikZ cho các câu hỏi cần minh họa trong lời giải"""
    os.makedirs(cache_dir, exist_ok=True)
    for q in questions:
        if q.tikz_image_path and os.path.exists(q.tikz_image_path):
            continue
            
        full_text = q.raw_prompt_text
        sol_text = ""
        if q.solution_xml_elements:
            sol_text = " ".join(["".join([t.text for t in elem.iter() if t.tag.endswith('}t') and t.text]) for elem in q.solution_xml_elements])

        combined = (full_text + " " + sol_text).lower()

        tikz_code = None
        # Heuristic 1: Bất phương trình x + 3y >= 6
        if ("miền nghiệm" in combined or "bất phương trình" in combined) and ("x+3y" in combined or "x + 3y" in combined):
            tikz_code = generate_tikz_for_bpt_x_3y_ge_6()
        # Heuristic 2: Bất phương trình 2x - y <= 2
        elif ("miền nghiệm" in combined or "bất phương trình" in combined) and ("2x-y" in combined or "2x - y" in combined):
            tikz_code = generate_tikz_for_bpt_2x_y_le_2()
        # Heuristic 3: Tháp CD với khoảng cách 24m và góc 63/48
        elif "tháp" in combined and ("24" in combined or "63" in combined or "cd" in combined):
            tikz_code = generate_tikz_for_tower_cd_24m()
        # Heuristic 4: Tháp AB với khoảng cách 500m và góc 30/45
        elif "tháp" in combined and ("500" in combined or "hải đăng" in combined or "mn" in combined):
            tikz_code = generate_tikz_for_tower_ab_500m()

        if tikz_code:
            img_path = os.path.join(cache_dir, f"tikz_q{q.original_index}.png")
            success = render_tikz_code_to_png(tikz_code, img_path, dpi=300)
            if success:
                q.tikz_image_path = img_path
                print(f"   [TikZ] Đã vẽ thành công hình minh họa TikZ 300 DPI cho Câu {q.original_index} -> {os.path.basename(img_path)}")
