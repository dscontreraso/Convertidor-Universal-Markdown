import os
import re
import io
import sys
import time
import json
import math
import hashlib
import zipfile
import threading
import subprocess
from statistics import median
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional third-party libraries
try:
    import pymupdf
except ImportError:
    pymupdf = None

try:
    import docx
except ImportError:
    docx = None

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    import pptx
except ImportError:
    pptx = None

try:
    from PIL import Image
except ImportError:
    Image = None

try:
    from rapidocr_onnxruntime import RapidOCR
except ImportError:
    RapidOCR = None


# ==============================================================================
# OCR ENGINE SINGLETON (THREAD-SAFE RAPIDOCR)
# ==============================================================================

_OCR_INSTANCE = None
_OCR_LOCK = threading.Lock()


def get_ocr_engine():
    """Lazy thread-safe initialization of RapidOCR."""
    global _OCR_INSTANCE
    if RapidOCR is None:
        return None
    if _OCR_INSTANCE is None:
        with _OCR_LOCK:
            if _OCR_INSTANCE is None:
                try:
                    _OCR_INSTANCE = RapidOCR()
                except Exception:
                    _OCR_INSTANCE = None
    return _OCR_INSTANCE


# ==============================================================================
# LATEX MATH & CLEAN TABLE FORMATTING UTILITIES
# ==============================================================================

UNICODE_TO_LATEX = {
    "α": r"\alpha ", "β": r"\beta ", "γ": r"\gamma ", "δ": r"\delta ",
    "ε": r"\varepsilon ", "ζ": r"\zeta ", "η": r"\eta ", "θ": r"\theta ",
    "ι": r"\iota ", "κ": r"\kappa ", "λ": r"\lambda ", "μ": r"\mu ",
    "ν": r"\nu ", "ξ": r"\xi ", "π": r"\pi ", "ρ": r"\rho ",
    "σ": r"\sigma ", "τ": r"\tau ", "υ": r"\upsilon ", "φ": r"\varphi ",
    "χ": r"\chi ", "ψ": r"\psi ", "ω": r"\omega ",
    "Γ": r"\Gamma ", "Δ": r"\Delta ", "Θ": r"\Theta ", "Λ": r"\Lambda ",
    "Ξ": r"\Xi ", "Π": r"\Pi ", "Σ": r"\Sigma ", "Φ": r"\Phi ",
    "Ψ": r"\Psi ", "Ω": r"\Omega ",
    "∑": r"\sum ", "∏": r"\prod ", "∫": r"\int ", "∮": r"\oint ",
    "√": r"\sqrt", "∂": r"\partial ", "∇": r"\nabla ", "∞": r"\infty ",
    "±": r"\pm ", "∓": r"\mp ", "×": r"\times ", "÷": r"\div ",
    "≤": r"\leq ", "≥": r"\geq ", "≠": r"\neq ",
    "≈": r"\approx ", "≡": r"\equiv ", "∝": r"\propto ", "∈": r"\in ",
    "∉": r"\notin ", "⊂": r"\subset ", "⊆": r"\subseteq ", "∪": r"\cup ",
    "∩": r"\cap ", "∀": r"\forall ", "∃": r"\exists ", "⇒": r"\Rightarrow ",
    "⇔": r"\Leftrightarrow ", "→": r"\rightarrow ", "←": r"\leftarrow ",
    "↔": r"\leftrightarrow ", "°": r"^{\circ}",
    "²": r"^{2}", "³": r"^{3}", "¹": r"^{1}", "⁰": r"^{0}",
    "⁴": r"^{4}", "⁵": r"^{5}", "⁶": r"^{6}", "⁷": r"^{7}", "⁸": r"^{8}", "⁹": r"^{9}",
    "₀": r"_{0}", "₁": r"_{1}", "₂": r"_{2}", "₃": r"_{3}", "₄": r"_{4}",
    "₅": r"_{5}", "₆": r"_{6}", "₇": r"_{7}", "₈": r"_{8}", "₉": r"_{9}",
}

MATH_SYMBOL_SET = set(UNICODE_TO_LATEX.keys())

# Regex patterns to detect standalone equations or inline math expressions
STANDALONE_EQ_RE = re.compile(
    r"^\s*(?:[A-Za-z\d\)\]\s_^{}\\]+\s*(?:=|\\leq|\\geq|\\neq|\\approx|≤|≥|≠|≈)\s*[^=]{2,})"
    r"|(?:\b(?:lim|sum|int|frac|sqrt)\b|\\(?:sum|int|frac|sqrt|partial|nabla|alpha|beta|gamma|sigma|Delta|Omega))"
)

INLINE_MATH_PATTERNS = [
    # Variables with subscripts/superscripts or equations like R^2 = 0.98, x_1 + x_2 = y
    re.compile(r"(?<!\$)\b([A-Za-z]\s*[\^_]\s*\{?[A-Za-z0-9+\-]+\}?(?:\s*(?:<=|>=|!=|[=<>≤≥≠≈+\-*/])\s*[A-Za-z0-9.\^_{}\\]+)*)(?!\$)"),
    # Functional expressions like f(x) = 2x + 1
    re.compile(r"(?<!\$)\b([fghFGH]\([A-Za-z,\s]+\)\s*=\s*[A-Za-z0-9+\-*/^_().\s]+)(?!\$)"),
    # Comparisons or assignments with Greek/math names or variables like alpha <= 0.05, p < 0.01
    re.compile(r"(?<!\$)\b((?:alpha|beta|gamma|delta|sigma|theta|lambda|mu|pi|[A-Za-z])\s*(?:<=|>=|!=|≤|≥|≠|≈|=|<|>)\s*-?\d+(?:\.\d+)?)(?!\$)"),
]


def normalize_math_to_latex(expr: str) -> str:
    """Convert Unicode math symbols, roots, fractions, and exponents inside an expression to LaTeX."""
    out = expr.strip()
    # Handle square root √(x), √x, sqrt(x)
    out = re.sub(r"(?:√|sqrt)\(([^)]+)\)", r"\\sqrt{\1}", out)
    out = re.sub(r"√([A-Za-z0-9]+)", r"\\sqrt{\1}", out)

    # ASCII math operators to LaTeX
    out = out.replace("<=", r" \leq ").replace(">=", r" \geq ").replace("!=", r" \neq ").replace("+-", r" \pm ")
    out = re.sub(r"\b(alpha|beta|gamma|delta|epsilon|theta|lambda|mu|pi|sigma|omega|Delta|Sigma|Omega)\b", r"\\\1", out)

    for uni_char, latex_cmd in UNICODE_TO_LATEX.items():
        if uni_char in out:
            out = out.replace(uni_char, latex_cmd)

    # Convert simple (a+b)/(c+d) fractions
    out = re.sub(r"\(([^()]+)\)\s*/\s*\(([^()]+)\)", r"\\frac{\1}{\2}", out)
    # Replace * with \cdot inside equations
    out = re.sub(r"(?<=[A-Za-z0-9\)\}])\s*\*\s*(?=[A-Za-z0-9\(\\])", r" \\cdot ", out)
    # Clean redundant spaces
    out = re.sub(r"\s{2,}", " ", out).strip()
    return out


def is_standalone_equation(line: str) -> bool:
    """Determine if a single line is predominantly a mathematical equation."""
    s = line.strip()
    if not s or len(s) > 180 or s.startswith("#") or s.startswith("|") or s.startswith("!["):
        return False

    # Already wrapped in LaTeX
    if (s.startswith("$$") and s.endswith("$$")) or (s.startswith(r"\[") and s.endswith(r"\]")):
        return False

    # Count natural language words (4+ letters, excluding common math function names)
    math_words = {"sqrt", "frac", "alfa", "alpha", "beta", "gamma", "delta", "sigma", "theta", "sin", "cos", "tan", "log", "exp", "lim", "max", "min"}
    alpha_words = [w for w in re.findall(r"[A-Za-zÁÉÍÓÚáéíóúÑñ]{4,}", s) if w.lower() not in math_words]

    # If there are 2 or more normal prose words, it is a text sentence (with possible inline math), not a standalone equation block
    if len(alpha_words) >= 2:
        return False

    math_chars = sum(1 for c in s if c in MATH_SYMBOL_SET or c in "=+-*/^_<>()[]{}")
    has_eq_sign = any(op in s for op in ("=", "<=", ">=", "!=", "≤", "≥", "≠", "≈", "∑", "∫", "√", r"\frac", r"\sum", r"\int"))
    if has_eq_sign and math_chars >= 2:
        if "http://" in s or "https://" in s or "://" in s:
            return False
        return True

    if sum(1 for c in s if c in MATH_SYMBOL_SET) >= 2 and len(alpha_words) <= 1:
        return True

    return False


def format_line_with_latex(line: str) -> str:
    """
    Convert a line of text:
    - If it's a standalone equation, wrap in display math $$ ... $$
    - If it's prose containing math symbols or formulas, wrap those segments in inline $ ... $
    """
    stripped = line.strip()
    if not stripped:
        return ""

    if is_standalone_equation(stripped):
        # Preserve optional equation numbering like (1) or (2.1) at the end
        eq_num_match = re.search(r"\s+\((\d+(?:\.\d+)?)\)\s*$", stripped)
        if eq_num_match:
            num_tag = eq_num_match.group(1)
            core_eq = stripped[: eq_num_match.start()].strip()
            latex_eq = normalize_math_to_latex(core_eq)
            return f"$$\n{latex_eq} \\tag{{{num_tag}}}\n$$"
        else:
            latex_eq = normalize_math_to_latex(stripped)
            return f"$$\n{latex_eq}\n$$"

    # Inline math replacement for prose lines
    result = stripped
    for pat in INLINE_MATH_PATTERNS:
        result = pat.sub(lambda m: f"${normalize_math_to_latex(m.group(1))}$", result)

    # Replace isolated Unicode math symbols outside $...$ with inline LaTeX
    if any(ch in result for ch in MATH_SYMBOL_SET):
        parts = re.split(r"(\$[^$]+\$)", result)
        new_parts = []
        for p in parts:
            if p.startswith("$") and p.endswith("$"):
                new_parts.append(p)
            else:
                # Wrap contiguous math symbol clusters with adjacent variables/numbers
                p_conv = re.sub(
                    r"([A-Za-z0-9]*\s*[∑∏∫∮√∂∇∞±∓×÷≤≥≠≈≡∝∈∉⊂⊆∪∩∀∃⇒⇔α-ωΓ-Ω²³¹⁰⁴-⁹₀-₉]+\s*[A-Za-z0-9^_().]*)",
                    lambda m: f" ${normalize_math_to_latex(m.group(1))}$ ",
                    p,
                )
                new_parts.append(re.sub(r"\s{2,}", " ", p_conv))
        result = "".join(new_parts).strip()

    return result


def format_clean_markdown_table(rows: list) -> str:
    """
    Convert a 2D list of rows into a clean, well-aligned Markdown table.
    - Removes completely empty rows and completely empty columns.
    - Escapes internal pipe characters (|).
    - Collapses multi-line cell text cleanly.
    """
    if not rows:
        return ""

    # Clean each cell
    cleaned_rows = []
    max_cols = 0
    for row in rows:
        if not row:
            continue
        clean_r = []
        for cell in row:
            txt = str(cell if cell is not None else "").strip()
            # Collapse internal newlines to <br> if multiple lines, or single space
            lines = [ln.strip() for ln in txt.splitlines() if ln.strip()]
            txt = " <br> ".join(lines) if len(lines) > 1 else (lines[0] if lines else "")
            txt = re.sub(r"\s{2,}", " ", txt)
            txt = txt.replace("|", r"\|")
            clean_r.append(txt)
        if any(c != "" for c in clean_r):
            cleaned_rows.append(clean_r)
            max_cols = max(max_cols, len(clean_r))

    if not cleaned_rows or max_cols == 0:
        return ""

    # Pad rows to max_cols
    for r in cleaned_rows:
        while len(r) < max_cols:
            r.append("")

    # Filter out columns that are 100% empty across all rows
    non_empty_col_indices = [
        c_idx for c_idx in range(max_cols) if any(r[c_idx] != "" for r in cleaned_rows)
    ]
    if not non_empty_col_indices:
        return ""

    final_rows = [[r[c_idx] for c_idx in non_empty_col_indices] for r in cleaned_rows]
    col_count = len(non_empty_col_indices)

    header = final_rows[0]
    # Ensure header cells have a fallback name if blank
    header_fmt = [h if h else f"Col {i+1}" for i, h in enumerate(header)]

    # Compute column widths for clean visual alignment
    col_widths = [max(len(header_fmt[i]), 3) for i in range(col_count)]
    for r in final_rows[1:]:
        for i, cell in enumerate(r):
            col_widths[i] = min(max(col_widths[i], len(cell)), 60)

    def fmt_row(r_cells):
        return "| " + " | ".join(c.ljust(col_widths[i]) for i, c in enumerate(r_cells)) + " |"

    lines = [
        fmt_row(header_fmt),
        "| " + " | ".join("-" * col_widths[i] for i in range(col_count)) + " |",
    ]
    for r in final_rows[1:]:
        lines.append(fmt_row(r))

    return "\n".join(lines)


# ==============================================================================
# OCR LAYOUT, HEADING, TABLE & LATEX RECONSTRUCTION (FOR SCANNED PDFS & IMAGES)
# ==============================================================================

def reconstruct_markdown_from_ocr(ocr_result: list) -> str:
    """
    Takes RapidOCR result: list of [box_4_pts, text, confidence]
    and reconstructs:
    - Headings (# / ## / ###) based on text box height and heading patterns
    - Clean Markdown tables when multiple boxes share the same vertical row in aligned columns
    - LaTeX equations ($$ ... $$ and $ ... $)
    - Clean paragraphs
    """
    if not ocr_result:
        return ""

    items = []
    heights = []
    for entry in ocr_result:
        if not entry or len(entry) < 2:
            continue
        box, txt = entry[0], str(entry[1]).strip()
        if not txt:
            continue
        xs = [pt[0] for pt in box]
        ys = [pt[1] for pt in box]
        x0, x1 = min(xs), max(xs)
        y0, y1 = min(ys), max(ys)
        h = max(y1 - y0, 1.0)
        heights.append(h)
        items.append({
            "x0": x0, "x1": x1, "y0": y0, "y1": y1,
            "yc": (y0 + y1) / 2.0, "h": h, "text": txt
        })

    if not items:
        return ""

    body_h = median(heights) if heights else 16.0
    # Sort top-to-bottom
    items.sort(key=lambda it: (it["yc"], it["x0"]))

    # Group items into horizontal rows based on vertical overlap
    rows = []
    current_row = [items[0]]
    row_yc = items[0]["yc"]
    row_h = items[0]["h"]

    for it in items[1:]:
        if abs(it["yc"] - row_yc) <= max(row_h, it["h"]) * 0.55:
            current_row.append(it)
            row_yc = sum(x["yc"] for x in current_row) / len(current_row)
            row_h = max(x["h"] for x in current_row)
        else:
            current_row.sort(key=lambda x: x["x0"])
            rows.append(current_row)
            current_row = [it]
            row_yc = it["yc"]
            row_h = it["h"]
    if current_row:
        current_row.sort(key=lambda x: x["x0"])
        rows.append(current_row)

    # Identify contiguous multi-column table blocks (2+ consecutive rows with >= 2 separated columns)
    md_blocks = []
    i = 0
    n_rows = len(rows)

    while i < n_rows:
        row = rows[i]
        if len(row) >= 2:
            # Check if this starts a table block (at least 2 consecutive multi-cell rows)
            table_group = [row]
            j = i + 1
            while j < n_rows and len(rows[j]) >= 2:
                table_group.append(rows[j])
                j += 1

            if len(table_group) >= 2:
                # Reconstruct column buckets by X center coordinates
                col_count = max(len(r) for r in table_group)
                if col_count >= 2:
                    raw_table = [[cell["text"] for cell in r] for r in table_group]
                    tbl_md = format_clean_markdown_table(raw_table)
                    if tbl_md:
                        md_blocks.append(tbl_md)
                        i = j
                        continue

        # Single text line
        line_text = " ".join(c["text"] for c in row).strip()
        max_h = max(c["h"] for c in row)
        h_ratio = max_h / max(body_h, 1.0)

        heading_level = detect_heading_level(line_text, h_ratio, is_bold=(h_ratio >= 1.25))
        if heading_level > 0:
            prefix = "#" * heading_level
            clean_h = re.sub(r"^#+\s*", "", line_text).strip()
            md_blocks.append(f"{prefix} {clean_h}")
        else:
            md_blocks.append(format_line_with_latex(line_text))
        i += 1

    return "\n\n".join(b for b in md_blocks if b)


def ocr_pixmap_or_image_bytes(img_bytes: bytes) -> str:
    """Run RapidOCR on image bytes and return structured Markdown (headings, tables, LaTeX)."""
    ocr_engine = get_ocr_engine()
    if ocr_engine is None:
        return ""
    try:
        result, _ = ocr_engine(img_bytes)
        if not result:
            return ""
        return reconstruct_markdown_from_ocr(result)
    except Exception:
        return ""


# ==============================================================================
# HEADING DETECTION & PDF STRUCTURAL PARSER
# ==============================================================================

HEADING_REGEX_L1 = re.compile(
    r"^(?:CAP[ÍI]TULO|T[ÍI]TULO|SECCI[ÓO]N|PARTE|ANEXO)\s+[IVXLCDM0-9]+\b", re.IGNORECASE
)
HEADING_REGEX_L2 = re.compile(
    r"^(?:ART[ÍI]CULO\s+\d+|(?:\d{1,2})\.\s+[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚáéíóúÑñ\s]{2,})"
)
HEADING_REGEX_L3 = re.compile(
    r"^(?:\d{1,2}\.\d{1,2}\.?\s+[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚáéíóúÑñ\s]{2,}|PAR[ÁA]GRAFO\b)", re.IGNORECASE
)
HEADING_REGEX_L4 = re.compile(
    r"^\d{1,2}\.\d{1,2}\.\d{1,2}\.?\s+\S+"
)


def detect_heading_level(text: str, size_ratio: float, is_bold: bool) -> int:
    """
    Determine Markdown heading level (0 = normal text, 2 = ##, 3 = ###, 4 = ####)
    based on typography (font size ratio vs body median, bold flag) and numbering patterns.
    """
    s = text.strip()
    if not s or len(s) > 160:
        return 0
    # Avoid treating equations or table lines as headings
    if is_standalone_equation(s):
        return 0

    if HEADING_REGEX_L1.match(s) or size_ratio >= 1.55:
        return 2
    if HEADING_REGEX_L2.match(s) or (size_ratio >= 1.28 and is_bold):
        return 2
    if HEADING_REGEX_L3.match(s) or size_ratio >= 1.18 or (size_ratio >= 1.10 and is_bold and len(s) < 95):
        return 3
    if HEADING_REGEX_L4.match(s) or (is_bold and len(s) < 75 and s.isupper() and len(s) > 3):
        return 4
    return 0


def extract_structured_page_elements(page, body_font_size: float, enable_ocr: bool = True) -> tuple:
    """
    Extract all elements from a PyMuPDF page in reading order (y0):
    - Headings (#, ##, ###, ####) from font sizes, bold flags, and structural patterns
    - Clean Markdown tables placed at their exact vertical position
    - Mathematical equations formatted in LaTeX ($$ ... $$ and $ ... $)
    - Superscripts/subscripts from span flags
    - Automatic RapidOCR fallback if the page is scanned or image-based
    Returns (markdown_text, used_ocr_bool)
    """
    elements = []  # list of (y0, x0, markdown_string)

    # 1. Extract tables with bounding boxes
    table_rects = []
    try:
        tabs = page.find_tables()
        if tabs and tabs.tables:
            for tab in tabs.tables:
                rect = pymupdf.Rect(tab.bbox)
                table_rects.append(rect)
                extracted = tab.extract()
                if extracted:
                    clean_tbl = format_clean_markdown_table(extracted)
                    if clean_tbl:
                        elements.append((rect.y0, rect.x0, clean_tbl))
    except Exception:
        pass

    # 2. Extract rich text spans (excluding table bounding boxes)
    text_dict = page.get_text("dict")
    digital_char_count = 0

    for block in text_dict.get("blocks", []):
        if block.get("type") != 0:
            continue
        b_rect = pymupdf.Rect(block["bbox"])
        # Skip block if its center lies inside a detected table
        b_center = pymupdf.Point((b_rect.x0 + b_rect.x1) / 2, (b_rect.y0 + b_rect.y1) / 2)
        if any( b_center in tr or (b_rect.intersects(tr) and (b_rect & tr).get_area() / max(b_rect.get_area(), 1) > 0.5) for tr in table_rects):
            continue

        block_lines_out = []
        for line in block.get("lines", []):
            spans = line.get("spans", [])
            if not spans:
                continue

            line_parts = []
            max_span_size = 0.0
            bold_chars = 0
            total_chars = 0

            for sp in spans:
                raw_txt = sp.get("text", "")
                if not raw_txt.strip():
                    line_parts.append(raw_txt)
                    continue

                sz = float(sp.get("size", body_font_size) or body_font_size)
                flags = int(sp.get("flags", 0))
                font_name = (sp.get("font") or "").lower()
                is_sup = bool(flags & 1)  # superscript bit in PyMuPDF
                is_italic = bool(flags & 2) or "italic" in font_name or "oblique" in font_name
                is_bold = bool(flags & 16) or "bold" in font_name or "black" in font_name

                max_span_size = max(max_span_size, sz)
                c_len = len(raw_txt.strip())
                total_chars += c_len
                if is_bold:
                    bold_chars += c_len

                if is_sup and len(raw_txt.strip()) <= 6:
                    line_parts.append(f"^{{{raw_txt.strip()}}}")
                else:
                    line_parts.append(raw_txt)

            full_line = "".join(line_parts).strip()
            if not full_line:
                continue

            digital_char_count += len(full_line)
            line_is_bold = (bold_chars / max(total_chars, 1)) >= 0.6
            size_ratio = max_span_size / max(body_font_size, 1.0)

            h_level = detect_heading_level(full_line, size_ratio, line_is_bold)
            if h_level > 0:
                clean_title = re.sub(r"^#+\s*", "", full_line).strip()
                block_lines_out.append(f"\n{'#' * h_level} {clean_title}\n")
            else:
                # Bullet list normalization
                if full_line.startswith(("•", "●", "▪", "‣", "– ")):
                    bullet_body = full_line.lstrip("•●▪‣– ").strip()
                    block_lines_out.append(f"- {format_line_with_latex(bullet_body)}")
                else:
                    block_lines_out.append(format_line_with_latex(full_line))

        if block_lines_out:
            # Join lines inside the same paragraph block intelligently
            merged_block = []
            for ln in block_lines_out:
                if ln.startswith("\n#") or ln.startswith("$$") or ln.startswith("- "):
                    merged_block.append(ln.strip())
                elif merged_block and not merged_block[-1].startswith("#") and not merged_block[-1].startswith("$$") and not merged_block[-1].startswith("- "):
                    merged_block[-1] = merged_block[-1] + " " + ln
                else:
                    merged_block.append(ln)
            elements.append((b_rect.y0, b_rect.x0, "\n\n".join(merged_block)))

    # 3. Check if page is scanned (very few digital characters and contains images) -> Run RapidOCR!
    used_ocr = False
    if digital_char_count < 25 and not elements:
        imgs = page.get_images()
        if enable_ocr and (imgs or digital_char_count == 0):
            try:
                pix = page.get_pixmap(dpi=175)
                png_bytes = pix.tobytes("png")
                ocr_md = ocr_pixmap_or_image_bytes(png_bytes)
                if ocr_md.strip():
                    used_ocr = True
                    return ocr_md, True
            except Exception:
                pass
        if imgs:
            return f"*[Página escaneada ({len(imgs)} imágenes detectadas — sin texto reconocible)]*", False
        return "*[Página vacía]*", False

    # Sort all extracted text blocks and tables in natural top-to-bottom reading order
    elements.sort(key=lambda el: (round(el[0], 1), round(el[1], 1)))
    page_md = "\n\n".join(el[2] for el in elements if el[2].strip())
    return page_md, used_ocr


# ==============================================================================
# CONVERSION ENGINE (ALL FILE TYPES + RECURSIVE SUBFOLDERS + NESTED ZIPS)
# ==============================================================================

CATEGORY_LABELS = {
    "pdf": "📄 PDF (Encabezados + Tablas + LaTeX + OCR)",
    "excel": "📊 Excel (.xlsx / .xls / .csv)",
    "word": "📝 Word (.docx / .doc + Ecuaciones)",
    "pptx": "📽️ PowerPoint (.pptx)",
    "image": "🖼️ Imágenes (PNG/JPG + OCR + LaTeX)",
    "other": "🧾 Texto / JSON / XML / Otros",
}

TEXT_EXTENSIONS = {
    ".txt", ".md", ".rst", ".log", ".ini", ".cfg", ".conf", ".env",
    ".json", ".xml", ".yaml", ".yml", ".toml", ".html", ".htm", ".css",
    ".py", ".js", ".ts", ".jsx", ".tsx", ".java", ".c", ".cpp", ".h",
    ".cs", ".go", ".rs", ".php", ".rb", ".sql", ".sh", ".bat", ".ps1", ".tex"
}

CODE_LANG_MAP = {
    ".py": "python", ".js": "javascript", ".ts": "typescript",
    ".json": "json", ".xml": "xml", ".html": "html", ".htm": "html",
    ".css": "css", ".sql": "sql", ".yaml": "yaml", ".yml": "yaml",
    ".sh": "bash", ".bat": "bat", ".ps1": "powershell", ".md": "markdown",
    ".tex": "latex"
}


def classify_extension(ext: str) -> str:
    ext = ext.lower()
    if ext == ".pdf":
        return "pdf"
    if ext in (".xlsx", ".xls", ".xlsm", ".csv", ".tsv"):
        return "excel"
    if ext in (".docx", ".doc", ".rtf", ".odt"):
        return "word"
    if ext in (".pptx", ".ppt", ".odp"):
        return "pptx"
    if ext in (".png", ".jpg", ".jpeg", ".bmp", ".webp", ".gif", ".tif", ".tiff", ".ico"):
        return "image"
    return "other"


def sanitize_relpath(path_str: str) -> str:
    """Sanitize path segments to be valid on Windows filesystems."""
    if not path_str or path_str == ".":
        return ""
    parts = path_str.replace("\\", "/").split("/")
    clean_parts = []
    for p in parts:
        clean_p = re.sub(r'[<>:"|?*]', "_", p).strip(". ")
        if clean_p:
            clean_parts.append(clean_p)
    return os.path.join(*clean_parts) if clean_parts else ""


def sanitize_filename(name: str) -> str:
    clean = re.sub(r'[<>:"|?*\\/]', "_", name).strip(". ")
    return clean if clean else "archivo"


def estimate_pdf_body_font_size(doc) -> float:
    """Compute the median font size of body text across the PDF document."""
    sizes = []
    sample_pages = min(len(doc), 8)
    for pno in range(sample_pages):
        try:
            tdict = doc[pno].get_text("dict")
            for b in tdict.get("blocks", []):
                if b.get("type") != 0:
                    continue
                for ln in b.get("lines", []):
                    for sp in ln.get("spans", []):
                        txt = sp.get("text", "").strip()
                        sz = float(sp.get("size", 0) or 0)
                        if len(txt) >= 3 and 5.0 <= sz <= 36.0:
                            sizes.extend([sz] * min(len(txt), 30))
        except Exception:
            pass
    return median(sizes) if sizes else 11.0


def convert_pdf_to_md(pdf_data: bytes, title: str, enable_ocr: bool = True) -> str:
    if pymupdf is None:
        return f"# {title}\n\n> PyMuPDF no está disponible para extraer este PDF.\n"
    try:
        doc = pymupdf.open(stream=pdf_data, filetype="pdf")
    except Exception as e:
        return f"# {title}\n\n> Error al abrir PDF: {e}\n"

    body_font_size = estimate_pdf_body_font_size(doc)
    md_parts = [f"# {title}\n"]
    meta = doc.metadata or {}
    info_lines = [f"- **Páginas**: {len(doc)}"]
    if meta.get("author"):
        info_lines.append(f"- **Autor**: {meta.get('author')}")
    if meta.get("title"):
        info_lines.append(f"- **Título original**: {meta.get('title')}")
    if meta.get("creationDate"):
        info_lines.append(f"- **Fecha Creación**: {meta.get('creationDate')}")

    ocr_pages_count = 0
    page_outputs = []
    for page_idx, page in enumerate(doc):
        page_md, used_ocr = extract_structured_page_elements(page, body_font_size, enable_ocr=enable_ocr)
        if used_ocr:
            ocr_pages_count += 1
        tag = " *(OCR)*" if used_ocr else ""
        page_outputs.append(f"\n## Página {page_idx + 1}{tag}\n\n{page_md}\n")

    if ocr_pages_count > 0:
        info_lines.append(f"- **Páginas procesadas con OCR**: {ocr_pages_count}")

    md_parts.append("\n".join(info_lines) + "\n\n---\n")
    md_parts.extend(page_outputs)

    doc.close()
    return "".join(md_parts)


def extract_docx_omml_math(para_element) -> list:
    """Extract Office MathML (m:oMath) text nodes from a Word paragraph XML element and convert to LaTeX."""
    equations = []
    try:
        for node in para_element.iter():
            tag = node.tag.split("}")[-1] if "}" in node.tag else node.tag
            if tag == "oMath":
                tokens = []
                for sub in node.iter():
                    sub_tag = sub.tag.split("}")[-1] if "}" in sub.tag else sub.tag
                    if sub_tag == "t" and sub.text:
                        tokens.append(sub.text)
                if tokens:
                    raw_eq = "".join(tokens).strip()
                    if raw_eq:
                        equations.append(normalize_math_to_latex(raw_eq))
    except Exception:
        pass
    return equations


def convert_docx_to_md(docx_data: bytes, title: str) -> str:
    if docx is None:
        return f"# {title}\n\n> python-docx no está disponible.\n"
    try:
        document = docx.Document(io.BytesIO(docx_data))
    except Exception as e:
        return convert_generic_binary_to_md(docx_data, title, ".docx", note=f"No se pudo abrir como DOCX estándar ({e})")

    md_parts = [f"# {title}\n"]
    props = document.core_properties
    meta_lines = [
        f"- **Párrafos**: {len(document.paragraphs)}",
        f"- **Tablas**: {len(document.tables)}",
    ]
    if props.author:
        meta_lines.append(f"- **Autor**: {props.author}")
    if props.created:
        meta_lines.append(f"- **Fecha Creación**: {props.created}")
    md_parts.append("\n".join(meta_lines) + "\n\n---\n\n")

    for para in document.paragraphs:
        txt = para.text.strip()
        omml_eqs = extract_docx_omml_math(para._element)

        if not txt and not omml_eqs:
            continue

        style_name = (para.style.name or "").lower() if para.style else ""
        if "heading 1" in style_name or "título 1" in style_name:
            md_parts.append(f"## {txt}\n\n")
        elif "heading 2" in style_name or "título 2" in style_name:
            md_parts.append(f"### {txt}\n\n")
        elif "heading 3" in style_name or "título 3" in style_name:
            md_parts.append(f"#### {txt}\n\n")
        elif "heading 4" in style_name or "título 4" in style_name:
            md_parts.append(f"##### {txt}\n\n")
        elif "list" in style_name or "viñeta" in style_name:
            md_parts.append(f"- {format_line_with_latex(txt)}\n")
        else:
            if txt:
                md_parts.append(f"{format_line_with_latex(txt)}\n\n")

        for eq in omml_eqs:
            md_parts.append(f"$$\n{eq}\n$$\n\n")

    if document.tables:
        md_parts.append("\n## Tablas del Documento\n\n")
        for idx, table in enumerate(document.tables, 1):
            rows_data = []
            for row in table.rows:
                rows_data.append([c.text.strip() for c in row.cells])
            clean_tbl = format_clean_markdown_table(rows_data)
            if clean_tbl:
                md_parts.append(f"### Tabla {idx}\n\n{clean_tbl}\n\n")

    return "".join(md_parts)


def convert_excel_or_csv_to_md(data: bytes, title: str, ext: str) -> str:
    if pd is None:
        return f"# {title}\n\n> pandas no está disponible.\n"

    ext = ext.lower()
    if ext in (".csv", ".tsv"):
        sep = "\t" if ext == ".tsv" else None
        df = None
        for enc in ("utf-8", "latin-1", "cp1252"):
            try:
                df = pd.read_csv(io.BytesIO(data), sep=sep, engine="python" if sep is None else "c", encoding=enc)
                break
            except Exception:
                continue
        if df is None:
            return convert_text_to_md(data, title, ext)
        rows, cols = df.shape
        md_parts = [
            f"# {title}\n\n",
            f"- **Formato**: `{ext.upper()}`\n",
            f"- **Filas**: {rows} · **Columnas**: {cols}\n\n---\n\n",
        ]
        if rows == 0:
            md_parts.append("*[Archivo sin filas de datos]*\n")
        else:
            raw_rows = [list(df.columns)] + df.fillna("").values.tolist()
            md_parts.append(format_clean_markdown_table(raw_rows) + "\n")
        return "".join(md_parts)

    try:
        xl = pd.read_excel(io.BytesIO(data), sheet_name=None)
    except Exception as e:
        return f"# {title}\n\n> Error al leer archivo Excel: {e}\n"

    sheet_names = list(xl.keys())
    md_parts = [
        f"# {title}\n",
        f"- **Hojas disponibles ({len(sheet_names)})**: {', '.join(sheet_names)}\n\n---\n",
    ]
    for sheet_name, df in xl.items():
        rows, cols = df.shape
        md_parts.append(f"\n## Hoja: {sheet_name}\n")
        md_parts.append(f"*(Filas: {rows}, Columnas: {cols})*\n\n")
        if rows == 0:
            md_parts.append("*[Hoja vacía]*\n\n")
            continue
        try:
            raw_rows = [list(df.columns)] + df.fillna("").values.tolist()
            tbl_md = format_clean_markdown_table(raw_rows)
            md_parts.append((tbl_md if tbl_md else "*[Hoja sin datos válidos]*") + "\n\n")
        except Exception as e:
            md_parts.append(f"> Error al formatear la hoja '{sheet_name}': {e}\n\n")

    return "".join(md_parts)


def convert_pptx_to_md(pptx_data: bytes, title: str) -> str:
    if pptx is None:
        return f"# {title}\n\n> python-pptx no está disponible.\n"
    try:
        prs = pptx.Presentation(io.BytesIO(pptx_data))
    except Exception as e:
        return f"# {title}\n\n> Error al abrir presentación PPTX: {e}\n"

    md_parts = [
        f"# {title}\n",
        f"- **Total Diapositivas**: {len(prs.slides)}\n\n---\n",
    ]
    for slide_idx, slide in enumerate(prs.slides, 1):
        md_parts.append(f"\n## Diapositiva {slide_idx}\n\n")
        slide_texts = []
        for shape in slide.shapes:
            if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    txt = para.text.strip()
                    if txt:
                        slide_texts.append(format_line_with_latex(txt))
            elif getattr(shape, "has_table", False) and shape.has_table:
                raw_rows = [[c.text.strip() for c in row.cells] for row in shape.table.rows]
                clean_tbl = format_clean_markdown_table(raw_rows)
                if clean_tbl:
                    slide_texts.append(clean_tbl)

        if slide_texts:
            for st in slide_texts:
                md_parts.append(st + "\n\n")
        else:
            md_parts.append("*[Diapositiva gráfica / sin cuadros de texto]*\n\n")

        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                notes = slide.notes_slide.notes_text_frame.text.strip()
                if notes:
                    md_parts.append(f"> **Notas del orador**: {notes}\n\n")
        except Exception:
            pass

    return "".join(md_parts)


def convert_image_to_md(img_data: bytes, title: str, ext: str, enable_ocr: bool = True) -> str:
    ext_clean = ext.lstrip(".").upper() or "IMG"
    md_parts = [
        f"# {title}\n\n",
        f"- **Tipo de archivo**: Imagen `{ext_clean}`\n",
        f"- **Tamaño**: {len(img_data) / 1024:.2f} KB\n",
    ]
    if Image is not None:
        try:
            with Image.open(io.BytesIO(img_data)) as im:
                md_parts.append(f"- **Dimensiones**: {im.width} x {im.height} píxeles\n")
                md_parts.append(f"- **Modo de color**: {im.mode}\n")
                if im.info.get("dpi"):
                    md_parts.append(f"- **DPI**: {im.info.get('dpi')}\n")
        except Exception:
            pass

    md_parts.append(f"\n![{title}]({title}{ext})\n")

    if enable_ocr:
        ocr_md = ocr_pixmap_or_image_bytes(img_data)
        if ocr_md.strip():
            md_parts.append("\n---\n\n## Texto y Ecuaciones Extraídas por OCR\n\n")
            md_parts.append(ocr_md + "\n")

    return "".join(md_parts)


def convert_text_to_md(text_data: bytes, title: str, ext: str) -> str:
    decoded = None
    for enc in ("utf-8", "utf-8-sig", "latin-1", "cp1252"):
        try:
            decoded = text_data.decode(enc)
            break
        except Exception:
            continue

    if decoded is None:
        return convert_generic_binary_to_md(text_data, title, ext)

    ext_lower = ext.lower()
    if ext_lower == ".md":
        return decoded

    if ext_lower == ".json":
        try:
            obj = json.loads(decoded)
            pretty = json.dumps(obj, indent=2, ensure_ascii=False)
            return f"# {title}\n\n```json\n{pretty}\n```\n"
        except Exception:
            pass

    lang = CODE_LANG_MAP.get(ext_lower, "")
    return (
        f"# {title}\n\n"
        f"- **Archivo original**: `{title}{ext}`\n"
        f"- **Tamaño**: {len(text_data):,} bytes\n\n---\n\n"
        f"```{lang}\n{decoded}\n```\n"
    )


def convert_generic_binary_to_md(data: bytes, title: str, ext: str, note: str = "") -> str:
    """Fallback converter for any binary or proprietary file type."""
    sha256 = hashlib.sha256(data).hexdigest()
    md_parts = [
        f"# {title}\n\n",
        f"- **Extensión original**: `{ext or '(sin extensión)'}`\n",
        f"- **Tamaño**: {len(data):,} bytes ({len(data) / 1024:.2f} KB)\n",
        f"- **SHA-256**: `{sha256}`\n",
    ]
    if note:
        md_parts.append(f"- **Nota**: {note}\n")
    md_parts.append("\n---\n\n")

    try:
        text_latin = data.decode("latin-1", errors="ignore")
        matches = re.findall(r"[A-Za-z0-9ÁÉÍÓÚáéíóúÑñ ,.;:\-_/()\[\]{}@#$%&*+=?!]{8,}", text_latin)
        if matches:
            snippets = []
            total_len = 0
            for m in matches[:200]:
                s = m.strip()
                if len(s) >= 8:
                    snippets.append(s)
                    total_len += len(s)
                    if total_len > 15000:
                        break
            if snippets:
                md_parts.append("## Fragmentos de Texto Detectados\n\n```\n")
                md_parts.append("\n".join(snippets))
                md_parts.append("\n```\n")
                return "".join(md_parts)
    except Exception:
        pass

    md_parts.append("*[Archivo binario o formato sin contenido textual directo]*\n")
    return "".join(md_parts)


def convert_bytes_to_markdown(data: bytes, original_name: str, enable_ocr: bool = True) -> str:
    base_name, ext = os.path.splitext(original_name)
    ext_lower = ext.lower()

    if ext_lower == ".pdf":
        return convert_pdf_to_md(data, title=base_name, enable_ocr=enable_ocr)
    elif ext_lower in (".xlsx", ".xls", ".xlsm", ".csv", ".tsv"):
        return convert_excel_or_csv_to_md(data, title=base_name, ext=ext_lower)
    elif ext_lower == ".docx":
        return convert_docx_to_md(data, title=base_name)
    elif ext_lower in (".doc", ".rtf", ".odt"):
        return convert_generic_binary_to_md(data, title=base_name, ext=ext_lower)
    elif ext_lower in (".pptx", ".ppt"):
        return convert_pptx_to_md(data, title=base_name)
    elif ext_lower in (".png", ".jpg", ".jpeg", ".bmp", ".webp", ".gif", ".tif", ".tiff", ".ico"):
        return convert_image_to_md(data, title=base_name, ext=ext_lower, enable_ocr=enable_ocr)
    elif ext_lower in TEXT_EXTENSIONS:
        return convert_text_to_md(data, title=base_name, ext=ext_lower)
    else:
        return convert_generic_binary_to_md(data, title=base_name, ext=ext_lower)


# ==============================================================================
# DEEP DIRECTORY + NESTED ZIP SCANNER
# ==============================================================================

def scan_zip_bytes(zip_bytes: bytes, zip_display_chain: str, target_base_dir: str,
                   scan_result: dict, parent_tree_children: list, explore_zips: bool):
    """Recursively inspect a ZIP archive (including nested ZIPs inside ZIPs)."""
    try:
        with zipfile.ZipFile(io.BytesIO(zip_bytes), "r") as zf:
            zip_subfolders_set = set()

            for info in zf.infolist():
                inner_path = info.filename.replace("\\", "/").strip("/")
                if not inner_path:
                    continue

                if info.is_dir():
                    zip_subfolders_set.add(inner_path)
                    continue

                parts = inner_path.split("/")
                if len(parts) > 1:
                    for i in range(1, len(parts)):
                        zip_subfolders_set.add("/".join(parts[:i]))

                fname = parts[-1]
                ext = os.path.splitext(fname)[1].lower()

                if ext == ".zip" and explore_zips:
                    scan_result["zip_count"] += 1
                    nested_stem = sanitize_filename(os.path.splitext(fname)[0])
                    inner_rel_dir = sanitize_relpath("/".join(parts[:-1]))
                    nested_target_dir = os.path.join(target_base_dir, inner_rel_dir, nested_stem)
                    nested_node = {
                        "name": f"🗜️ {fname}",
                        "type": "ZIP anidado",
                        "files": 0,
                        "target": nested_target_dir,
                        "children": [],
                    }
                    parent_tree_children.append(nested_node)
                    try:
                        nested_bytes = zf.read(info.filename)
                        before_files = len(scan_result["tasks"])
                        scan_zip_bytes(
                            nested_bytes,
                            f"{zip_display_chain}::{inner_path}",
                            nested_target_dir,
                            scan_result,
                            nested_node["children"],
                            explore_zips,
                        )
                        nested_node["files"] = len(scan_result["tasks"]) - before_files
                    except Exception as e:
                        scan_result["corrupt_zips"].append((f"{zip_display_chain}::{inner_path}", str(e)))
                else:
                    scan_result["zip_files"] += 1
                    cat = classify_extension(ext)
                    scan_result["categories"][cat] += 1
                    scan_result["extensions"].add(ext if ext else "(sin ext)")

                    clean_inner = sanitize_relpath(inner_path)
                    inner_base, _ = os.path.splitext(clean_inner)
                    target_md = os.path.join(target_base_dir, inner_base + ".md")

                    scan_result["tasks"].append({
                        "type": "zip_stream" if "::" in zip_display_chain else "zip_disk",
                        "zip_source": zip_display_chain,
                        "inner_entry": info.filename,
                        "raw_bytes": zf.read(info.filename) if "::" in zip_display_chain else None,
                        "target_md": target_md,
                        "original_name": fname,
                        "display_path": f"{zip_display_chain}::{inner_path}",
                    })
                    if len(parent_tree_children) < 40:
                        parent_tree_children.append({
                            "name": f"📄 {fname}",
                            "type": ext.upper() or "ARCHIVO",
                            "files": 1,
                            "target": os.path.basename(target_md),
                            "children": [],
                        })

            scan_result["zip_folders"] += len(zip_subfolders_set)
    except Exception as e:
        scan_result["corrupt_zips"].append((zip_display_chain, str(e)))


def scan_directory_recursive(source_dir: str, target_dir: str, explore_zips: bool = True,
                             preserve_structure: bool = True) -> dict:
    """
    Scan source_dir recursively across all subfolders and .zip files.
    Returns comprehensive counts, hierarchical preview tree, and task list.
    """
    scan_result = {
        "source_dir": source_dir,
        "target_dir": target_dir,
        "physical_folders": 0,
        "zip_folders": 0,
        "direct_files": 0,
        "zip_files": 0,
        "zip_count": 0,
        "categories": {k: 0 for k in CATEGORY_LABELS},
        "extensions": set(),
        "corrupt_zips": [],
        "tasks": [],
        "tree": [],
    }

    if not os.path.isdir(source_dir):
        return scan_result

    folder_nodes = {}
    root_node = {
        "name": f"📁 {os.path.basename(source_dir.rstrip('/\\')) or source_dir}/",
        "type": "Carpeta Raíz",
        "files": 0,
        "target": target_dir,
        "children": [],
    }
    scan_result["tree"].append(root_node)
    folder_nodes[""] = root_node

    for dirpath, dirnames, filenames in os.walk(source_dir):
        rel_dir = os.path.relpath(dirpath, source_dir)
        if rel_dir == ".":
            rel_dir = ""
        else:
            scan_result["physical_folders"] += 1

        clean_rel_dir = sanitize_relpath(rel_dir) if preserve_structure else ""

        if rel_dir != "":
            parent_rel = os.path.dirname(rel_dir)
            if parent_rel == ".":
                parent_rel = ""
            parent_node = folder_nodes.get(parent_rel, root_node)
            f_node = {
                "name": f"📁 {os.path.basename(dirpath)}/",
                "type": "Subcarpeta",
                "files": len(filenames),
                "target": os.path.join(target_dir, clean_rel_dir),
                "children": [],
            }
            if len(parent_node["children"]) < 100:
                parent_node["children"].append(f_node)
            folder_nodes[rel_dir] = f_node
        else:
            f_node = root_node

        for fname in filenames:
            src_file = os.path.join(dirpath, fname)
            ext = os.path.splitext(fname)[1].lower()

            if ext == ".zip" and explore_zips:
                scan_result["zip_count"] += 1
                zip_stem = sanitize_filename(os.path.splitext(fname)[0])
                target_zip_dir = (
                    os.path.join(target_dir, clean_rel_dir, zip_stem)
                    if preserve_structure
                    else os.path.join(target_dir, zip_stem)
                )
                zip_node = {
                    "name": f"🗜️ {fname}",
                    "type": "Archivo ZIP",
                    "files": 0,
                    "target": target_zip_dir,
                    "children": [],
                }
                if len(f_node["children"]) < 100:
                    f_node["children"].append(zip_node)

                try:
                    with open(src_file, "rb") as zf_raw:
                        zip_bytes = zf_raw.read()
                    before_tasks = len(scan_result["tasks"])
                    scan_zip_bytes(
                        zip_bytes,
                        src_file,
                        target_zip_dir,
                        scan_result,
                        zip_node["children"],
                        explore_zips,
                    )
                    zip_node["files"] = len(scan_result["tasks"]) - before_tasks
                except Exception as e:
                    scan_result["corrupt_zips"].append((src_file, str(e)))
            else:
                scan_result["direct_files"] += 1
                cat = classify_extension(ext)
                scan_result["categories"][cat] += 1
                scan_result["extensions"].add(ext if ext else "(sin ext)")

                stem, _ = os.path.splitext(fname)
                clean_stem = sanitize_filename(stem)
                target_md_path = (
                    os.path.join(target_dir, clean_rel_dir, clean_stem + ".md")
                    if preserve_structure
                    else os.path.join(target_dir, clean_stem + ".md")
                )

                rel_display = os.path.join(rel_dir, fname) if rel_dir else fname
                scan_result["tasks"].append({
                    "type": "direct",
                    "src_path": src_file,
                    "target_md": target_md_path,
                    "original_name": fname,
                    "display_path": rel_display,
                })

                if len(f_node["children"]) < 50:
                    f_node["children"].append({
                        "name": f"📄 {fname}",
                        "type": ext.upper() or "ARCHIVO",
                        "files": 1,
                        "target": os.path.basename(target_md_path),
                        "children": [],
                    })

    root_node["files"] = len(scan_result["tasks"])
    return scan_result


def execute_conversion_task(task: dict, enable_ocr: bool = True) -> dict:
    """Worker function to convert a single file task (direct or inside ZIP)."""
    target_md = task["target_md"]
    orig_name = task["original_name"]
    ext = os.path.splitext(orig_name)[1].lower()

    try:
        if task["type"] == "direct":
            with open(task["src_path"], "rb") as f:
                data = f.read()
        elif task["type"] == "zip_disk":
            with zipfile.ZipFile(task["zip_source"], "r") as zf:
                data = zf.read(task["inner_entry"])
        else:
            data = task["raw_bytes"]

        md_content = convert_bytes_to_markdown(data, orig_name, enable_ocr=enable_ocr)

        os.makedirs(os.path.dirname(target_md), exist_ok=True)
        with open(target_md, "w", encoding="utf-8", errors="replace") as out_f:
            out_f.write(md_content)

        return {
            "status": "OK",
            "name": orig_name,
            "display": task["display_path"],
            "target": target_md,
            "ext": ext or "(sin ext)",
        }
    except Exception as e:
        err_msg = str(e)
        try:
            os.makedirs(os.path.dirname(target_md), exist_ok=True)
            with open(target_md, "w", encoding="utf-8") as out_f:
                out_f.write(f"# {orig_name} - Error de Conversión\n\n> {err_msg}\n")
        except Exception:
            pass
        return {
            "status": "ERROR",
            "name": orig_name,
            "display": task["display_path"],
            "target": target_md,
            "ext": ext or "(sin ext)",
            "error": err_msg,
        }


# ==============================================================================
# NATIVE DESKTOP GUI APPLICATION (APPROVED DESIGN #1)
# ==============================================================================

class MarkdownConverterApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Convertidor Universal a Markdown — Daniel Contreras")
        self.geometry("1020x690")
        self.minsize(900, 620)
        self.configure(bg="#0f172a")

        self.colors = {
            "bg": "#0f172a",
            "card": "#1e293b",
            "card_border": "#334155",
            "fg": "#f8fafc",
            "muted": "#94a3b8",
            "primary": "#3b82f6",
            "primary_fg": "#ffffff",
            "secondary": "#334155",
            "success": "#10b981",
            "link": "#60a5fa",
            "console_bg": "#090d16",
        }

        default_src = r"C:\Users\danie\Downloads\Monitoreo"
        if not os.path.isdir(default_src):
            default_src = r"C:\Users\danie\Downloads\Mapeo"
        self.src_dir_var = tk.StringVar(value=default_src)
        self.dst_dir_var = tk.StringVar(value=default_src + "_MD")
        self.explore_zips_var = tk.BooleanVar(value=True)
        self.preserve_struct_var = tk.BooleanVar(value=True)
        self.enable_ocr_var = tk.BooleanVar(value=True)
        self.generate_report_var = tk.BooleanVar(value=True)
        self.workers_var = tk.IntVar(value=12)

        self.current_scan = None
        self.is_converting = False
        self.cancel_requested = False

        self._configure_ttk_styles()
        self._build_ui()

        self.after(250, self.start_scan)

    def _configure_ttk_styles(self):
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background=self.colors["card"],
            foreground=self.colors["fg"],
            fieldbackground=self.colors["card"],
            borderwidth=0,
            rowheight=26,
            font=("Consolas", 10),
        )
        style.configure(
            "Treeview.Heading",
            background=self.colors["secondary"],
            foreground=self.colors["fg"],
            font=("Segoe UI", 9, "bold"),
            borderwidth=1,
        )
        style.map("Treeview", background=[("selected", self.colors["primary"])])

        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=self.colors["secondary"],
            background=self.colors["primary"],
            bordercolor=self.colors["card_border"],
            lightcolor=self.colors["primary"],
            darkcolor=self.colors["primary"],
            thickness=14,
        )

    def _build_ui(self):
        self.main_wrap = tk.Frame(self, bg=self.colors["bg"], padx=18, pady=14)
        self.main_wrap.pack(fill="both", expand=True)

        # 1. HEADER & TAB SWITCHER
        header = tk.Frame(self.main_wrap, bg=self.colors["bg"])
        header.pack(fill="x", pady=(0, 10))

        title_left = tk.Frame(header, bg=self.colors["bg"])
        title_left.pack(side="left", fill="y")

        badge = tk.Label(
            title_left,
            text="MD",
            bg=self.colors["primary"],
            fg=self.colors["primary_fg"],
            font=("Segoe UI", 11, "bold"),
            width=3,
            pady=4,
        )
        badge.pack(side="left", padx=(0, 10))

        title_texts = tk.Frame(title_left, bg=self.colors["bg"])
        title_texts.pack(side="left")
        tk.Label(
            title_texts,
            text="Convertidor Universal a Markdown",
            bg=self.colors["bg"],
            fg=self.colors["fg"],
            font=("Segoe UI", 13, "bold"),
            anchor="w",
        ).pack(anchor="w")
        ocr_status = "RapidOCR + LaTeX Activo" if RapidOCR is not None else "Modo Estándar + LaTeX"
        tk.Label(
            title_texts,
            text=f"Escaneo Recursivo + Archivos .ZIP · Encabezados, Tablas Limpias, Ecuaciones LaTeX y OCR ({ocr_status})",
            bg=self.colors["bg"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(anchor="w")

        tab_bar = tk.Frame(
            header,
            bg=self.colors["card"],
            highlightbackground=self.colors["card_border"],
            highlightthickness=1,
            padx=3,
            pady=3,
        )
        tab_bar.pack(side="right")

        self.tab_buttons = {}
        tabs_meta = [
            ("scan", "1. Escaneo y Resumen"),
            ("tree", "2. Subcarpetas y ZIPs"),
            ("run", "3. Conversión"),
        ]
        for key, label in tabs_meta:
            btn = tk.Button(
                tab_bar,
                text=label,
                font=("Segoe UI", 9, "bold"),
                relief="flat",
                bd=0,
                padx=12,
                pady=5,
                cursor="hand2",
                command=lambda k=key: self.switch_tab(k),
            )
            btn.pack(side="left", padx=2)
            self.tab_buttons[key] = btn

        tk.Frame(self.main_wrap, bg=self.colors["card_border"], height=1).pack(fill="x", pady=(0, 12))

        # 2. FOLDER PICKERS BAR
        folders_frame = tk.Frame(self.main_wrap, bg=self.colors["bg"])
        folders_frame.pack(fill="x", pady=(0, 10))
        folders_frame.columnconfigure(0, weight=1)
        folders_frame.columnconfigure(1, weight=1)

        src_card = tk.Frame(
            folders_frame,
            bg=self.colors["card"],
            highlightbackground=self.colors["card_border"],
            highlightthickness=1,
            padx=12,
            pady=8,
        )
        src_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))

        src_info = tk.Frame(src_card, bg=self.colors["card"])
        src_info.pack(side="left", fill="both", expand=True)
        tk.Label(
            src_info,
            text="Carpeta de Origen (recursiva + ZIPs)",
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9, "bold"),
            anchor="w",
        ).pack(anchor="w")
        self.src_label = tk.Label(
            src_info,
            textvariable=self.src_dir_var,
            bg=self.colors["card"],
            fg=self.colors["fg"],
            font=("Consolas", 9),
            anchor="w",
        )
        self.src_label.pack(anchor="w", fill="x", pady=(2, 0))

        tk.Button(
            src_card,
            text="Elegir carpeta…",
            bg=self.colors["secondary"],
            fg=self.colors["fg"],
            activebackground=self.colors["primary"],
            activeforeground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=10,
            pady=5,
            cursor="hand2",
            command=self.choose_source_folder,
        ).pack(side="right", padx=(8, 0))

        dst_card = tk.Frame(
            folders_frame,
            bg=self.colors["card"],
            highlightbackground=self.colors["card_border"],
            highlightthickness=1,
            padx=12,
            pady=8,
        )
        dst_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        dst_info = tk.Frame(dst_card, bg=self.colors["card"])
        dst_info.pack(side="left", fill="both", expand=True)
        tk.Label(
            dst_info,
            text="Carpeta de Destino (.md)",
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9, "bold"),
            anchor="w",
        ).pack(anchor="w")
        self.dst_label = tk.Label(
            dst_info,
            textvariable=self.dst_dir_var,
            bg=self.colors["card"],
            fg=self.colors["fg"],
            font=("Consolas", 9),
            anchor="w",
        )
        self.dst_label.pack(anchor="w", fill="x", pady=(2, 0))

        dst_btns = tk.Frame(dst_card, bg=self.colors["card"])
        dst_btns.pack(side="right", padx=(8, 0))
        tk.Button(
            dst_btns,
            text="Cambiar…",
            bg=self.colors["secondary"],
            fg=self.colors["fg"],
            activebackground=self.colors["primary"],
            activeforeground="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=10,
            pady=5,
            cursor="hand2",
            command=self.choose_dest_folder,
        ).pack(side="left")

        # 3. 4 KPI COUNTER CARDS
        kpi_frame = tk.Frame(self.main_wrap, bg=self.colors["bg"])
        kpi_frame.pack(fill="x", pady=(0, 12))
        for i in range(4):
            kpi_frame.columnconfigure(i, weight=1)

        self.kpi_files_val, self.kpi_files_sub = self._create_kpi_card(
            kpi_frame, 0, "Total Archivos", "0", "0 directos · 0 en ZIP"
        )
        self.kpi_folders_val, self.kpi_folders_sub = self._create_kpi_card(
            kpi_frame, 1, "Subcarpetas", "0", "0 físicas · 0 en ZIP"
        )
        self.kpi_zips_val, self.kpi_zips_sub = self._create_kpi_card(
            kpi_frame, 2, "Archivos .ZIP", "0", "Incluye ZIPs anidados"
        )
        self.kpi_types_val, self.kpi_types_sub = self._create_kpi_card(
            kpi_frame, 3, "Formatos Detectados", "0 tipos", "PDF, DOCX, XLSX, PPTX…"
        )

        # 4. SWITCHABLE CONTENT VIEWS
        self.views_container = tk.Frame(self.main_wrap, bg=self.colors["bg"])
        self.views_container.pack(fill="both", expand=True)

        self.view_scan = tk.Frame(self.views_container, bg=self.colors["bg"])
        self.view_tree = tk.Frame(self.views_container, bg=self.colors["bg"])
        self.view_run = tk.Frame(self.views_container, bg=self.colors["bg"])

        self._build_view_scan()
        self._build_view_tree()
        self._build_view_run()

        # 5. FOOTER WITH STATUS AND AUTHOR SIGNATURE (BOTTOM-RIGHT CORNER)
        tk.Frame(self.main_wrap, bg=self.colors["card_border"], height=1).pack(fill="x", pady=(10, 8))
        footer = tk.Frame(self.main_wrap, bg=self.colors["bg"])
        footer.pack(fill="x", side="bottom")

        status_left = tk.Frame(footer, bg=self.colors["bg"])
        status_left.pack(side="left")

        self.status_dot = tk.Label(
            status_left,
            text="●",
            bg=self.colors["bg"],
            fg=self.colors["success"],
            font=("Segoe UI", 10, "bold"),
        )
        self.status_dot.pack(side="left", padx=(0, 6))

        self.status_text = tk.Label(
            status_left,
            text="Listo para escanear y convertir",
            bg=self.colors["bg"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
        )
        self.status_text.pack(side="left")

        author_right = tk.Frame(footer, bg=self.colors["bg"])
        author_right.pack(side="right")

        tk.Label(
            author_right,
            text="Desarrollado por ",
            bg=self.colors["bg"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
        ).pack(side="left")
        tk.Label(
            author_right,
            text="Daniel Contreras",
            bg=self.colors["bg"],
            fg=self.colors["fg"],
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left")
        tk.Label(
            author_right,
            text=" · ",
            bg=self.colors["bg"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
        ).pack(side="left")
        email_lbl = tk.Label(
            author_right,
            text="dscontreraso@unal.edu.co",
            bg=self.colors["bg"],
            fg=self.colors["link"],
            font=("Consolas", 9, "bold"),
            cursor="hand2",
        )
        email_lbl.pack(side="left")
        email_lbl.bind("<Button-1>", lambda e: self._copy_email())

        self.switch_tab("scan")

    def _create_kpi_card(self, parent, col, title, init_val, init_sub):
        padx = (0, 6) if col < 3 else (0, 0)
        if col > 0:
            padx = (6, 6) if col < 3 else (6, 0)

        card = tk.Frame(
            parent,
            bg=self.colors["card"],
            highlightbackground=self.colors["card_border"],
            highlightthickness=1,
            padx=12,
            pady=10,
        )
        card.grid(row=0, column=col, sticky="nsew", padx=padx)

        tk.Label(
            card,
            text=title,
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
            anchor="w",
        ).pack(anchor="w")

        val_lbl = tk.Label(
            card,
            text=init_val,
            bg=self.colors["card"],
            fg=self.colors["fg"],
            font=("Consolas", 18, "bold"),
            anchor="w",
        )
        val_lbl.pack(anchor="w", pady=(2, 2))

        sub_lbl = tk.Label(
            card,
            text=init_sub,
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Segoe UI", 8),
            anchor="w",
        )
        sub_lbl.pack(anchor="w")
        return val_lbl, sub_lbl

    def _build_view_scan(self):
        card = tk.Frame(
            self.view_scan,
            bg=self.colors["card"],
            highlightbackground=self.colors["card_border"],
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        card.pack(fill="both", expand=True)

        top_row = tk.Frame(card, bg=self.colors["card"])
        top_row.pack(fill="x", pady=(0, 10))

        tk.Label(
            top_row,
            text="Desglose de archivos encontrados por formato",
            bg=self.colors["card"],
            fg=self.colors["fg"],
            font=("Segoe UI", 10, "bold"),
        ).pack(side="left")

        workers_box = tk.Frame(top_row, bg=self.colors["card"])
        workers_box.pack(side="right")
        tk.Label(
            workers_box,
            text="Hilos de CPU:",
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
        ).pack(side="left", padx=(0, 6))
        spin = tk.Spinbox(
            workers_box,
            from_=1,
            to=32,
            textvariable=self.workers_var,
            width=4,
            bg=self.colors["secondary"],
            fg=self.colors["fg"],
            buttonbackground=self.colors["secondary"],
            relief="flat",
            font=("Consolas", 9, "bold"),
        )
        spin.pack(side="left")
        tk.Label(
            workers_box,
            text="workers",
            bg=self.colors["card"],
            fg=self.colors["fg"],
            font=("Segoe UI", 9, "bold"),
        ).pack(side="left", padx=(4, 0))

        grid = tk.Frame(card, bg=self.colors["card"])
        grid.pack(fill="x", pady=(4, 12))
        for c in range(3):
            grid.columnconfigure(c, weight=1)

        self.cat_count_labels = {}
        cat_keys = ["pdf", "excel", "word", "pptx", "image", "other"]
        for idx, ckey in enumerate(cat_keys):
            r = idx // 3
            c = idx % 3
            box = tk.Frame(
                grid,
                bg=self.colors["bg"],
                highlightbackground=self.colors["card_border"],
                highlightthickness=1,
                padx=12,
                pady=10,
            )
            box.grid(row=r, column=c, sticky="nsew", padx=4, pady=4)
            tk.Label(
                box,
                text=CATEGORY_LABELS[ckey],
                bg=self.colors["bg"],
                fg=self.colors["fg"],
                font=("Segoe UI", 9),
            ).pack(side="left")
            cnt_lbl = tk.Label(
                box,
                text="0",
                bg=self.colors["bg"],
                fg=self.colors["link"],
                font=("Consolas", 11, "bold"),
            )
            cnt_lbl.pack(side="right")
            self.cat_count_labels[ckey] = cnt_lbl

        controls_bar = tk.Frame(self.view_scan, bg=self.colors["bg"], pady=10)
        controls_bar.pack(fill="x")

        opts_left = tk.Frame(controls_bar, bg=self.colors["bg"])
        opts_left.pack(side="left")

        cb_style = {
            "bg": self.colors["bg"],
            "fg": self.colors["muted"],
            "activebackground": self.colors["bg"],
            "activeforeground": self.colors["fg"],
            "selectcolor": self.colors["card"],
            "font": ("Segoe UI", 9),
        }
        tk.Checkbutton(
            opts_left,
            text="Explorar .ZIP (y anidados)",
            variable=self.explore_zips_var,
            command=self.start_scan,
            **cb_style,
        ).pack(side="left", padx=(0, 10))

        tk.Checkbutton(
            opts_left,
            text="Conservar subcarpetas",
            variable=self.preserve_struct_var,
            command=self.start_scan,
            **cb_style,
        ).pack(side="left", padx=(0, 10))

        tk.Checkbutton(
            opts_left,
            text="OCR + Ecuaciones LaTeX",
            variable=self.enable_ocr_var,
            **cb_style,
        ).pack(side="left", padx=(0, 10))

        tk.Checkbutton(
            opts_left,
            text="Generar INFORME_CONVERSION.md",
            variable=self.generate_report_var,
            **cb_style,
        ).pack(side="left")

        self.start_btn = tk.Button(
            controls_bar,
            text="Iniciar Conversión →",
            bg=self.colors["primary"],
            fg=self.colors["primary_fg"],
            activebackground="#2563eb",
            activeforeground="#ffffff",
            font=("Segoe UI", 10, "bold"),
            relief="flat",
            padx=18,
            pady=7,
            cursor="hand2",
            command=self.start_conversion,
        )
        self.start_btn.pack(side="right")

        tk.Button(
            controls_bar,
            text="↻ Re-escanear",
            bg=self.colors["secondary"],
            fg=self.colors["fg"],
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=12,
            pady=7,
            cursor="hand2",
            command=self.start_scan,
        ).pack(side="right", padx=(0, 8))

    def _build_view_tree(self):
        card = tk.Frame(
            self.view_tree,
            bg=self.colors["card"],
            highlightbackground=self.colors["card_border"],
            highlightthickness=1,
            padx=10,
            pady=10,
        )
        card.pack(fill="both", expand=True)

        top_bar = tk.Frame(card, bg=self.colors["card"])
        top_bar.pack(fill="x", pady=(0, 6))
        tk.Label(
            top_bar,
            text="Estructura detectada (Subcarpetas físicas + contenido interno de archivos .ZIP)",
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
        ).pack(side="left")
        tk.Label(
            top_bar,
            text="Profundidad: ilimitada",
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Segoe UI", 9),
        ).pack(side="right")

        tree_wrap = tk.Frame(card, bg=self.colors["card"])
        tree_wrap.pack(fill="both", expand=True)

        cols = ("tipo", "archivos", "destino")
        self.tree = ttk.Treeview(tree_wrap, columns=cols, show="tree headings")
        self.tree.heading("#0", text="Carpeta / Archivo / ZIP", anchor="w")
        self.tree.heading("tipo", text="Tipo", anchor="w")
        self.tree.heading("archivos", text="Archivos", anchor="center")
        self.tree.heading("destino", text="Salida (.md)", anchor="w")

        self.tree.column("#0", width=380, anchor="w")
        self.tree.column("tipo", width=120, anchor="w")
        self.tree.column("archivos", width=90, anchor="center")
        self.tree.column("destino", width=260, anchor="w")

        vsb = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

    def _build_view_run(self):
        card = tk.Frame(
            self.view_run,
            bg=self.colors["card"],
            highlightbackground=self.colors["card_border"],
            highlightthickness=1,
            padx=14,
            pady=12,
        )
        card.pack(fill="both", expand=True)

        prog_top = tk.Frame(card, bg=self.colors["card"])
        prog_top.pack(fill="x", pady=(0, 6))

        self.prog_label = tk.Label(
            prog_top,
            text="Progreso: 0 / 0 (0.0%)",
            bg=self.colors["card"],
            fg=self.colors["fg"],
            font=("Consolas", 10, "bold"),
        )
        self.prog_label.pack(side="left")

        self.speed_label = tk.Label(
            prog_top,
            text="Velocidad: 0.0 arch/s · Restante: --",
            bg=self.colors["card"],
            fg=self.colors["muted"],
            font=("Consolas", 9),
        )
        self.speed_label.pack(side="right")

        self.progress_bar = ttk.Progressbar(
            card,
            style="Custom.Horizontal.TProgressbar",
            orient="horizontal",
            mode="determinate",
            maximum=100,
            value=0,
        )
        self.progress_bar.pack(fill="x", pady=(0, 10))

        console_wrap = tk.Frame(
            card,
            bg=self.colors["console_bg"],
            highlightbackground=self.colors["card_border"],
            highlightthickness=1,
        )
        console_wrap.pack(fill="both", expand=True)

        self.log_text = tk.Text(
            console_wrap,
            bg=self.colors["console_bg"],
            fg=self.colors["muted"],
            insertbackground=self.colors["fg"],
            font=("Consolas", 9),
            relief="flat",
            padx=10,
            pady=8,
            state="disabled",
        )
        log_scroll = ttk.Scrollbar(console_wrap, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

        run_actions = tk.Frame(self.view_run, bg=self.colors["bg"], pady=8)
        run_actions.pack(fill="x")

        self.cancel_btn = tk.Button(
            run_actions,
            text="Detener Conversión",
            bg="#dc2626",
            fg="#ffffff",
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            state="disabled",
            command=self.request_cancel,
        )
        self.cancel_btn.pack(side="left")

        tk.Button(
            run_actions,
            text="Abrir Carpeta de Destino",
            bg=self.colors["secondary"],
            fg=self.colors["fg"],
            font=("Segoe UI", 9, "bold"),
            relief="flat",
            padx=14,
            pady=6,
            cursor="hand2",
            command=self.open_destination_folder,
        ).pack(side="right")

    def switch_tab(self, tab_key: str):
        for k, btn in self.tab_buttons.items():
            if k == tab_key:
                btn.configure(bg=self.colors["primary"], fg=self.colors["primary_fg"])
            else:
                btn.configure(bg=self.colors["card"], fg=self.colors["muted"])

        self.view_scan.pack_forget()
        self.view_tree.pack_forget()
        self.view_run.pack_forget()

        if tab_key == "scan":
            self.view_scan.pack(fill="both", expand=True)
        elif tab_key == "tree":
            self.view_tree.pack(fill="both", expand=True)
        elif tab_key == "run":
            self.view_run.pack(fill="both", expand=True)

    def choose_source_folder(self):
        folder = filedialog.askdirectory(
            title="Seleccionar Carpeta de Origen para Convertir a Markdown",
            initialdir=self.src_dir_var.get() if os.path.isdir(self.src_dir_var.get()) else os.path.expanduser("~"),
        )
        if folder:
            norm_folder = os.path.normpath(folder)
            self.src_dir_var.set(norm_folder)
            self.dst_dir_var.set(norm_folder + "_MD")
            self.start_scan()

    def choose_dest_folder(self):
        folder = filedialog.askdirectory(
            title="Seleccionar Carpeta de Destino para Archivos .md",
            initialdir=os.path.dirname(self.dst_dir_var.get()) or os.path.expanduser("~"),
        )
        if folder:
            self.dst_dir_var.set(os.path.normpath(folder))
            self.start_scan()

    def open_destination_folder(self):
        dst = self.dst_dir_var.get()
        if os.path.isdir(dst):
            os.startfile(dst)
        else:
            messagebox.showinfo("Carpeta de Destino", f"La carpeta aún no existe:\n{dst}")

    def _copy_email(self):
        self.clipboard_clear()
        self.clipboard_append("dscontreraso@unal.edu.co")
        self.set_status("Correo copiado al portapapeles: dscontreraso@unal.edu.co", "#60a5fa")

    def set_status(self, message: str, color: str = None):
        self.status_text.configure(text=message)
        if color:
            self.status_dot.configure(fg=color)

    def append_log(self, line: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", line + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def start_scan(self):
        if self.is_converting:
            return
        src = self.src_dir_var.get().strip()
        dst = self.dst_dir_var.get().strip()
        if not os.path.isdir(src):
            self.set_status(f"La carpeta de origen no existe: {src}", "#f59e0b")
            return

        self.set_status("Escaneando subcarpetas y archivos .ZIP...", "#f59e0b")
        explore_zips = self.explore_zips_var.get()
        preserve = self.preserve_struct_var.get()

        def worker():
            res = scan_directory_recursive(src, dst, explore_zips=explore_zips, preserve_structure=preserve)
            self.after(0, lambda: self._on_scan_complete(res))

        threading.Thread(target=worker, daemon=True).start()

    def _on_scan_complete(self, res: dict):
        self.current_scan = res
        total_files = res["direct_files"] + res["zip_files"]
        total_folders = res["physical_folders"] + res["zip_folders"]
        total_zips = res["zip_count"]
        num_exts = len(res["extensions"])

        self.kpi_files_val.configure(text=f"{total_files:,}")
        self.kpi_files_sub.configure(text=f"{res['direct_files']:,} directos · {res['zip_files']:,} en ZIP")

        self.kpi_folders_val.configure(text=f"{total_folders:,}")
        self.kpi_folders_sub.configure(text=f"{res['physical_folders']:,} físicas · {res['zip_folders']:,} en ZIP")

        self.kpi_zips_val.configure(text=f"{total_zips:,}")
        self.kpi_zips_sub.configure(
            text="Incluye ZIPs anidados" if not res["corrupt_zips"] else f"{len(res['corrupt_zips'])} ZIP con alerta"
        )

        self.kpi_types_val.configure(text=f"{num_exts} tipos")
        sample_exts = ", ".join(sorted(e.upper().lstrip(".") for e in list(res["extensions"])[:5]))
        self.kpi_types_sub.configure(text=sample_exts if sample_exts else "Ningún archivo")

        for ckey, lbl in self.cat_count_labels.items():
            lbl.configure(text=f"{res['categories'].get(ckey, 0):,}")

        self.start_btn.configure(text=f"Iniciar Conversión ({total_files:,} archivos) →")

        for item in self.tree.get_children():
            self.tree.delete(item)

        def insert_nodes(parent_id, nodes):
            for n in nodes:
                nid = self.tree.insert(
                    parent_id,
                    "end",
                    text=n["name"],
                    values=(n["type"], f"{n['files']:,}", n["target"]),
                    open=(parent_id == ""),
                )
                if n.get("children"):
                    insert_nodes(nid, n["children"])

        insert_nodes("", res["tree"])

        self.set_status(
            f"Escaneo completado: {total_files:,} archivos y {total_folders:,} subcarpetas listas.",
            self.colors["success"],
        )

    def request_cancel(self):
        if self.is_converting:
            self.cancel_requested = True
            self.set_status("Deteniendo conversión...", "#f59e0b")

    def start_conversion(self):
        if self.is_converting:
            return
        if not self.current_scan or not self.current_scan["tasks"]:
            messagebox.showwarning("Sin archivos", "No hay archivos para convertir en la carpeta seleccionada.")
            return

        self.is_converting = True
        self.cancel_requested = False
        self.switch_tab("run")

        self.cancel_btn.configure(state="normal")
        self.start_btn.configure(state="disabled")
        self.progress_bar["value"] = 0

        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

        tasks = self.current_scan["tasks"]
        total_tasks = len(tasks)
        src_dir = self.src_dir_var.get()
        dst_dir = self.dst_dir_var.get()
        max_workers = max(1, int(self.workers_var.get()))
        enable_ocr = self.enable_ocr_var.get()

        self.append_log(f"=== INICIANDO CONVERSIÓN UNIVERSAL A MARKDOWN ===")
        self.append_log(f"Origen:  {src_dir}")
        self.append_log(f"Destino: {dst_dir}")
        self.append_log(f"OCR + LaTeX: {'ACTIVO' if enable_ocr else 'DESACTIVADO'}")
        self.append_log(f"Total archivos a convertir: {total_tasks:,} (Hilos: {max_workers})\n")
        self.set_status(f"Convirtiendo {total_tasks:,} archivos a Markdown...", self.colors["primary"])

        def run_worker():
            t_start = time.time()
            completed = 0
            counts = {"OK": 0, "ERROR": 0}
            by_ext = {}
            errors_list = []

            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(execute_conversion_task, t, enable_ocr): t for t in tasks}
                for fut in as_completed(futures):
                    if self.cancel_requested:
                        executor.shutdown(wait=False, cancel_futures=True)
                        break

                    res = fut.result()
                    completed += 1
                    st = res["status"]
                    counts[st] = counts.get(st, 0) + 1
                    ext = res["ext"]
                    by_ext[ext] = by_ext.get(ext, 0) + 1

                    if st == "ERROR":
                        errors_list.append(res)
                        log_line = f"[ERROR] {res['display']} -> {res.get('error', '')}"
                    else:
                        log_line = f"[OK] {res['display']} -> {os.path.basename(res['target'])}"

                    elapsed = max(time.time() - t_start, 0.001)
                    rate = completed / elapsed
                    rem = (total_tasks - completed) / rate if rate > 0 else 0
                    pct = (completed / total_tasks) * 100

                    if completed <= 30 or completed % max(1, total_tasks // 100) == 0 or completed == total_tasks or st == "ERROR":
                        self.after(
                            0,
                            lambda c=completed, p=pct, r=rate, rm=rem, ln=log_line: self._update_progress_ui(
                                c, total_tasks, p, r, rm, ln
                            ),
                        )

            t_total = time.time() - t_start
            if self.generate_report_var.get() and not self.cancel_requested:
                self._write_summary_report(dst_dir, src_dir, total_tasks, counts, by_ext, errors_list, t_total)

            self.after(0, lambda: self._on_conversion_finished(completed, total_tasks, counts, t_total, dst_dir))

        threading.Thread(target=run_worker, daemon=True).start()

    def _update_progress_ui(self, completed, total, pct, rate, rem_sec, log_line):
        self.progress_bar["value"] = pct
        self.prog_label.configure(text=f"Progreso: {completed:,} / {total:,} ({pct:.1f}%)")
        self.speed_label.configure(text=f"Velocidad: {rate:.1f} arch/s · Restante: ~{rem_sec:.0f}s")
        self.append_log(log_line)

    def _write_summary_report(self, dst_dir, src_dir, total_tasks, counts, by_ext, errors_list, t_total):
        try:
            os.makedirs(dst_dir, exist_ok=True)
            report_path = os.path.join(dst_dir, "INFORME_CONVERSION.md")
            with open(report_path, "w", encoding="utf-8") as rf:
                rf.write("# Informe de Conversión Universal a Markdown\n\n")
                rf.write(f"- **Fecha y Hora**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                rf.write(f"- **Carpeta Origen**: `{src_dir}`\n")
                rf.write(f"- **Carpeta Destino**: `{dst_dir}`\n")
                rf.write(f"- **Subcarpetas Físicas**: {self.current_scan['physical_folders']:,}\n")
                rf.write(f"- **Subcarpetas en ZIP**: {self.current_scan['zip_folders']:,}\n")
                rf.write(f"- **Archivos ZIP Procesados**: {self.current_scan['zip_count']:,}\n")
                rf.write(f"- **Tiempo Total**: {t_total:.2f} segundos ({t_total / 60:.2f} minutos)\n")
                rf.write(f"- **Total Archivos Procesados**: {total_tasks:,}\n")
                rf.write(f"  - **Exitosos**: {counts.get('OK', 0):,}\n")
                rf.write(f"  - **Con Error**: {counts.get('ERROR', 0):,}\n\n")

                rf.write("## Desglose por Extensión\n\n")
                rf.write("| Extensión | Cantidad | Porcentaje |\n")
                rf.write("|:---|:---:|:---:|\n")
                for ext, cnt in sorted(by_ext.items(), key=lambda x: x[1], reverse=True):
                    pct = (cnt / max(total_tasks, 1)) * 100
                    rf.write(f"| `{ext}` | {cnt:,} | {pct:.1f}% |\n")

                if errors_list:
                    rf.write("\n## Detalle de Errores\n\n")
                    for err in errors_list[:100]:
                        rf.write(f"- `{err['display']}`: {err.get('error', 'Error desconocido')}\n")

                rf.write("\n---\n*Desarrollado por Daniel Contreras · dscontreraso@unal.edu.co*\n")
        except Exception:
            pass

    def _on_conversion_finished(self, completed, total_tasks, counts, t_total, dst_dir):
        self.is_converting = False
        self.cancel_btn.configure(state="disabled")
        self.start_btn.configure(state="normal")

        if self.cancel_requested:
            self.append_log(f"\n[CANCELADO] Proceso detenido en {completed:,}/{total_tasks:,} archivos.")
            self.set_status("Conversión detenida por el usuario.", "#f59e0b")
        else:
            self.progress_bar["value"] = 100
            self.append_log(f"\n=== CONVERSIÓN FINALIZADA EN {t_total:.2f}s ===")
            self.append_log(f"Exitosos: {counts.get('OK', 0):,} | Errores: {counts.get('ERROR', 0):,}")
            self.append_log(f"Carpeta de salida: {dst_dir}")
            self.set_status(
                f"¡Conversión completada! {counts.get('OK', 0):,} archivos guardados en {dst_dir}",
                self.colors["success"],
            )
            messagebox.showinfo(
                "Conversión Completada",
                f"Se procesaron {completed:,} archivos en {t_total:.1f} segundos.\n\n"
                f"• Exitosos: {counts.get('OK', 0):,}\n"
                f"• Con alerta/error: {counts.get('ERROR', 0):,}\n\n"
                f"Destino: {dst_dir}",
            )


if __name__ == "__main__":
    app = MarkdownConverterApp()
    app.mainloop()
