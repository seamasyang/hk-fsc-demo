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


if __name__ == "__main__":
    main()