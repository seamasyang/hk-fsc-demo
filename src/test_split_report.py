#!/usr/bin/env python3
"""
pypdf 


"""

import sys
import os

from split_report_pypdf import process_pdf


from spire.pdf.common import *
from spire.pdf import *

import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c

import re

def extract_by_spire(pdf_path):
    doc = PdfDocument()
    doc.LoadFromFile(pdf_path)
    extract_options = PdfTextExtractOptions()
    extract_options.IsExtractAllText = True
    for i in range(0, 3):
        page = doc.Pages.get_Item(i)
        text_extractor = PdfTextExtractor(page)
        text = text_extractor.ExtractText(extract_options)

        print(f"page {i}\text: {text}")


def extract_by_pypdfium2(pdf_path):
    pdf = pdfium.PdfDocument(pdf_path)
    for i in range(0, 3):
        page = pdf[i]

        text_page = page.get_textpage()
        text = text_page.get_text_bounded()
        print(f"page {i}\text: {text}")

def extract_by_pypdf(pdf_path):
    from pypdf import PdfReader
    reader = PdfReader(pdf_path)
    for i in range(0, 3):
        page = reader.pages[i]
        text = page.extract_text()
        print(f"page {i}\text: {text}")

        toc_page = "目 錄" in text
        print(f"toc: {toc_page}")

def extract_balance_sheet(pdf_path):
    from pypdf import PdfReader
    reader = PdfReader(pdf_path)
    page_bs = reader.pages[6]
    kws = ["合併及公司资产负债表"]
    matched = _find_header_on_page(page_bs.extract_text(), kws)
    print(f"text: {page_bs.extract_text()[:200]}")
    print(f"matched: {matched}")

def _find_header_on_page(page_text: str, keywords: list[str]) -> str | None:
    """扫描页面文本，返回第一个匹配到的关键词标头，或 None。"""
    for line in page_text.split("\n"):
        ls = line.strip()
        for kw in keywords:
            if kw in ls:
                return kw
    return None

def _score_title_quality(line: str, keyword: str) -> int:
    """给标题行质量打分，分值越高越像真正的章节标题。

    评分维度:
    - +10: 关键词在行首（无前缀内容）
    - +8: 前缀仅为年份数字（如 '2025 年度'）
    - +5: 关键词占总行长度 ≥ 50%
    - -5: 关键词后有句子功能词（為、是、的等）→ 疑似句子片段
    - -10: 前缀包含非年份内容（如 '後附'）
    """

    _SENTENCE_FUNCTION_WORDS = set("為是的由在從將把被讓令使經會可已能應須需讓向與及於")

    if keyword not in line:
        return -100
    ls = line.strip()
    idx = ls.index(keyword)
    prefix = ls[:idx].strip()
    suffix = ls[idx + len(keyword):].strip()

    score = 0

    # 前缀检查
    if not prefix:
        score += 10  # 完美：关键词在行首
    elif re.match(r'^[\d\s年月日年度]*$', prefix):
        score += 8  # 可以接受：年份前缀
    else:
        score -= 10  # 不良：有非年份前缀内容

    # 关键词占总行长度比例
    kw_ratio = len(keyword) / max(len(ls), 1)
    if kw_ratio >= 0.5:
        score += 5

    # 后缀检查句子功能词
    if suffix:
        has_func_word = any(ch in _SENTENCE_FUNCTION_WORDS for ch in suffix)
        if has_func_word and not re.match(r'^[\s（(）)]*$', suffix):
            score -= 5

    # 后缀过长（超过关键词长度的2倍）→ 可能是句子
    if len(suffix) > len(keyword) * 2 and len(suffix) > 5:
        score -= 5

    return score


def main():
    
    
    pdf_path = "857_csa.pdf"
    if not os.path.exists(pdf_path):
        print("[ERROR] 文件不存在:", pdf_path)
        sys.exit(1)
    
    # 确定页码范围
    
    process_pdf(pdf_path, output_dir="data/processed/", dry_run=False, verbose=True, debug_mode=True)

    # extract_by_pypdfium2(pdf_path)
    # print("\n\n")
    # extract_by_pypdf(pdf_path)

    # extract_balance_sheet(pdf_path)




if __name__ == "__main__":
    main()