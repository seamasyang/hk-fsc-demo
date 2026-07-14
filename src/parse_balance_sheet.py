import os
from dotenv import load_dotenv

# Load variables from .env file
load_dotenv()


import re
import json
import pdfplumber
from openai import OpenAI
from typing import Dict, List, Tuple

# ==================== 配置 ====================
BASE_URL = os.getenv("base_url")
DEEPSEEK_API_KEY = os.getenv("api_key")
MODEL_NAME = os.getenv("model")

cur_dir = os.path.dirname(os.path.abspath(__file__))
# full_path = os.path.normpath(os.path.join(cur_dir, relative_path))
# print(cur_dir)

PDF_PATH = "../data/processed/857_csa/合併及公司资产负债表.pdf"
OUTPUT_MD_PATH = "../data/processed/857_csa/合併及公司资产负债表.md"

client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url=BASE_URL
)

# ==================== 1. 提取 PDF 文本 ====================
def extract_text_from_pdf(pdf_path: str, use_ocr=False) -> str:
    text_chunks = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            # TODO: not support OCR at this moment 
            # if use_ocr:
            #     from pdf2image import convert_from_bytes
            #     import pytesseract
            #     pil_image = page.to_image().original
            #     page_text = pytesseract.image_to_string(pil_image, lang='chi_sim+eng')
            # else:
            page_text = page.extract_text() or ""
            text_chunks.append(page_text.strip())
    return "\n\n".join(text_chunks)

# ==================== 2. 中文数字转阿拉伯数字 ====================
def chinese_to_arabic(chinese_str: str) -> int:
    """将中文数字年份（如'二零二六'）转换为阿拉伯数字"""
    mapping = {'零':'0','一':'1','二':'2','三':'3','四':'4','五':'5','六':'6','七':'7','八':'8','九':'9'}
    digits = ''.join(mapping.get(ch, '') for ch in chinese_str)
    return int(digits) if digits else 0

# ==================== 3. 表头模式检测 ====================
def detect_table_type(md_text: str) -> str:
    """
    扫描 Markdown 表头行，判断属于哪种模式：
    - 'standard' : 标准资产负债表（单实体，仅两列）
    - 'consolidated' : 合并及公司资产负债表（双实体，四列）
    """
    lines = md_text.split('\n')
    for line in lines:
        print(f"--->{line}")
        if '|' not in line:
            continue
        parts = [p.strip() for p in line.split('|') if p.strip()]
        if len(parts) < 3:
            continue
        # 合并出现“合并”和“公司”关键字
        has_consolidated = any('合并' or '合併' in p for p in parts)
        has_company = any('公司' in p for p in parts)
        if has_consolidated and has_company:
            print('consolidated balance sheet')
            return 'consolidated'
    # 默认返回标准模式
    print('standard balance sheet')
    return 'standard'

# ==================== 4. 调用 DeepSeek 清洗 ====================
def clean_and_format_balance_sheet(raw_text: str) -> str:
    # 先大致判断一下原始文本中是否包含“合并”和“公司”，以决定提示词
    has_consolidated = '合并' in raw_text
    has_company = '公司' in raw_text
    is_consolidated = has_consolidated and has_company

    if is_consolidated:
        prompt = f"""
你是一位资深财务专家，请将以下从 PDF 提取的**合并及公司资产负债表**文本转换为规范、干净的 Markdown 表格。

该表格包含**四列财务数据**：合并 2025 年、合并 2024 年、公司 2025 年、公司 2024 年。
表头格式：`项目 | 附注 | 合并_2025 | 合并_2024 | 公司_2025 | 公司_2024`。
请在表格下方添加钩稽检查，分别验证**合并**和**公司**两个实体的恒等式。

原始文本：
==========
{raw_text}
==========
"""
    else:
        prompt = f"""
你是一位资深财务专家，请将以下从 PDF 提取的**标准资产负债表**文本转换为规范、干净的 Markdown 表格。

该表格包含**两列财务数据**（如 2026 年、2025 年）。
表头格式：`项目 | 附注 | 2026 | 2025`（年份按实际内容调整）。
请在表格下方添加钩稽检查，验证总资产、负债、权益之间的恒等式。

原始文本：
==========
{raw_text}
==========
"""
    response = client.chat.completions.create(
        model=MODEL_NAME,
        messages=[{"role": "user", "content": prompt}],
        temperature=0.1,
        max_tokens=4096,
    )
    return response.choices[0].message.content


# ==================== 5. 解析 Markdown （核心修复） ====================
def parse_standard_table(md_text: str) -> Dict:
    """解析标准表（单实体，两列）"""
    lines = md_text.split('\n')
    data = {}
    years = []
    col_mapping = {}
    header_found = False

    for line in lines:
        if '|' not in line:
            continue
        parts = [p.strip() for p in line.split('|') if p.strip()]
        if not parts:
            continue

        # 检测表头
        if not header_found and '年' in ''.join(parts):
            # 提取中文或阿拉伯年份
            year_vals = []
            for p in parts:
                ch_year = re.findall(r'[零一二三四五六七八九]{4,}', p)
                ar_year = re.findall(r'\b20\d{2}\b', p)
                if ch_year:
                    year_vals.append(str(chinese_to_arabic(ch_year[0])))
                elif ar_year:
                    year_vals.append(ar_year[0])
            if len(year_vals) >= 2:
                years = year_vals[-2:]  # 取最后两个年份
                # 定位到最后两列
                col_mapping = {len(parts)-2: years[0], len(parts)-1: years[1]}
                header_found = True
                continue

        # 解析数据行
        if header_found:
            if len(parts) < 2:
                continue
            item_name = parts[0]
            vals = {}
            for idx, year in col_mapping.items():
                if idx < len(parts):
                    val_str = parts[idx].replace(',', '').replace('-', '0')
                    try:
                        vals[year] = int(val_str) if val_str.isdigit() else None
                    except:
                        vals[year] = None
            if any(v is not None for v in vals.values()):
                data[item_name] = vals
    return data

def parse_consolidated_table(md_text: str) -> Dict:
    """解析合并/公司表（双实体，四列）"""
    lines = md_text.split('\n')
    data = {}
    col_mapping = {}
    header_found = False

    for line in lines:
        if '|' not in line:
            continue
        parts = [p.strip() for p in line.split('|') if p.strip()]
        if not parts:
            continue

        # 检测表头
        if not header_found and any('合并' in p or '公司' in p for p in parts):
            patterns = []
            for p in parts:
                ch_year = re.findall(r'[零一二三四五六七八九]{4,}', p)
                ar_year = re.findall(r'\b20\d{2}\b', p)
                year = ch_year[0] if ch_year else (ar_year[0] if ar_year else None)
                if year:
                    if ch_year:
                        year = str(chinese_to_arabic(year))
                    entity = '合并' if '合并' in p else ('公司' if '公司' in p else None)
                    if entity and year:
                        patterns.append(f"{entity}_{year}")
            if len(patterns) >= 4:
                idx = 0
                for i, p in enumerate(parts):
                    if i < 2:  # 跳过“项目”和“附注”
                        continue
                    if idx < len(patterns):
                        col_mapping[i] = patterns[idx]
                        idx += 1
                header_found = True
                continue

        # 解析数据行
        if header_found:
            if len(parts) < 2:
                continue
            item_name = parts[0]
            vals = {}
            for idx, col_key in col_mapping.items():
                if idx < len(parts):
                    val_str = parts[idx].replace(',', '').replace('-', '0')
                    try:
                        vals[col_key] = int(val_str) if val_str.isdigit() else None
                    except:
                        vals[col_key] = None
            if any(v is not None for v in vals.values()):
                data[item_name] = vals
    return data

def parse_markdown_balance_sheet(md_text: str) -> Dict:
    """统一入口：根据表头模式自动选择解析函数"""
    mode = detect_table_type(md_text)
    if mode == 'consolidated':
        return parse_consolidated_table(md_text)
    else:
        return parse_standard_table(md_text)

# ==================== 6. 动态钩稽检查（修复硬编码索引） ====================
def check_balance(md_text: str) -> Tuple[bool, str]:
    mode = detect_table_type(md_text)
    data = parse_markdown_balance_sheet(md_text)
    if not data:
        return False, "无法解析表格数据，请检查 Markdown 格式。"

    all_keys = set()
    for item in data.values():
        all_keys.update(item.keys())

    checks = []
    passed = True

    if mode == 'standard':
        # ---- 标准模式：单实体校验 ----
        years = sorted(list(all_keys))
        for year in years:
            def find_value(keywords):
                for name, vals in data.items():
                    for kw in keywords:
                        if kw in name:
                            return vals.get(year)
                return None

            total_assets = find_value(['资产总计', '資產總計'])
            current_assets = find_value(['流动资产合计', '流動資產合計'])
            non_current_assets = find_value(['非流动资产合计', '非流動資產合計'])
            total_liabilities = find_value(['负债合计', '負債合計'])
            total_equity = find_value(['权益总额', '股東權益合計', '股东权益合计'])
            total_liab_equity = find_value(['权益总额及非流动负债', '負債及股東權益總計'])

            # 恒等式1：资产 = 流动资产 + 非流动资产
            if current_assets is not None and non_current_assets is not None and total_assets is not None:
                if current_assets + non_current_assets != total_assets:
                    checks.append(f"{year}年: 资产总计 ({total_assets}) ≠ 流动资产 ({current_assets}) + 非流动资产 ({non_current_assets})")
                    passed = False
                else:
                    checks.append(f"{year}年: 资产总计 = 流动资产 + 非流动资产 ✓ ({total_assets})")

            # 恒等式2：资产 = 负债 + 权益（如果有总负债和总权益）
            if total_liabilities is not None and total_equity is not None and total_assets is not None:
                if total_liabilities + total_equity != total_assets:
                    checks.append(f"{year}年: 资产总计 ({total_assets}) ≠ 负债 ({total_liabilities}) + 权益 ({total_equity})")
                    passed = False
                else:
                    checks.append(f"{year}年: 资产总计 = 负债 + 权益 ✓ ({total_assets})")

            # 恒等式3：总资产减流动负债 = 权益总额及非流动负债（常见于第一份PDF）
            if total_assets is not None and total_liab_equity is not None:
                if total_assets != total_liab_equity:
                    checks.append(f"{year}年: 资产总计 ({total_assets}) ≠ 权益总额及非流动负债 ({total_liab_equity})")
                    passed = False
                else:
                    checks.append(f"{year}年: 资产总计 = 权益总额及非流动负债 ✓ ({total_assets})")

    else:
        # ---- 合并模式：双实体分别校验 ----
        entities_years = {}
        for key in all_keys:
            match = re.match(r'^(合并|公司)_(\d{4})$', key)
            if match:
                entity, year = match.groups()
                if entity not in entities_years:
                    entities_years[entity] = set()
                entities_years[entity].add(year)

        for entity in entities_years:
            years = sorted(list(entities_years[entity]))
            for year in years:
                def find_value(keywords):
                    for name, vals in data.items():
                        for kw in keywords:
                            if kw in name:
                                return vals.get(f"{entity}_{year}")
                    return None

                total_assets = find_value(['资产总计', '資產總計'])
                current_assets = find_value(['流动资产合计', '流動資產合計'])
                non_current_assets = find_value(['非流动资产合计', '非流動資產合計'])
                total_liabilities = find_value(['负债合计', '負債合計'])
                total_equity = find_value(['股东权益合计', '股東權益合計'])
                total_liab_equity = find_value(['负债及股东权益总计', '負債及股東權益總計'])

                if current_assets is not None and non_current_assets is not None and total_assets is not None:
                    if current_assets + non_current_assets != total_assets:
                        checks.append(f"{entity} {year}年: 资产总计 ({total_assets}) ≠ 流动资产 ({current_assets}) + 非流动资产 ({non_current_assets})")
                        passed = False
                    else:
                        checks.append(f"{entity} {year}年: 资产总计 = 流动资产 + 非流动资产 ✓ ({total_assets})")

                if total_liabilities is not None and total_equity is not None and total_liab_equity is not None:
                    if total_liabilities + total_equity != total_liab_equity:
                        checks.append(f"{entity} {year}年: 负债及股东权益总计 ({total_liab_equity}) ≠ 负债 ({total_liabilities}) + 权益 ({total_equity})")
                        passed = False
                    else:
                        checks.append(f"{entity} {year}年: 负债及股东权益总计 = 负债 + 权益 ✓ ({total_liab_equity})")

                if total_assets is not None and total_liab_equity is not None:
                    if total_assets != total_liab_equity:
                        checks.append(f"{entity} {year}年: 资产总计 ({total_assets}) ≠ 负债及股东权益总计 ({total_liab_equity})")
                        passed = False
                    else:
                        checks.append(f"{entity} {year}年: 资产总计 = 负债及股东权益总计 ✓ ({total_assets})")

    if not checks:
        checks.append("未找到足够数据完成钩稽校验，请检查提取结果。")

    return passed, "\n".join(checks)

# ==================== 7. 主流程 ====================
def main():
    print("正在提取 PDF 文本...")
    raw_text = extract_text_from_pdf(PDF_PATH, use_ocr=False)
    if not raw_text.strip():
        print("警告：PDF 未能提取出文本，可能为扫描件。请改用 OCR 模式（use_ocr=True）。")
        return

    # 先检测原文本中是否包含“合并”和“公司”，以决定提示词
    mode_hint = 'consolidated' if ('合并' or '合併' in raw_text and '公司' in raw_text) else 'standard'
    print(f"检测到表头模式：{mode_hint}")

    print("正在调用 DeepSeek API 进行清洗和格式化...")
    md_result = clean_and_format_balance_sheet(raw_text)

    print("正在动态验证钩稽关系...")
    passed, summary = check_balance(md_result)
    if not passed:
        print("⚠️ 钩稽检查未通过，以下为详细检查结果：")
        print(summary)
    else:
        print("✅ 钩稽检查通过！")

    with open(OUTPUT_MD_PATH, 'w', encoding='utf-8') as f:
        f.write(md_result)
    print(f"已保存清洗后的 Markdown 文件至：{OUTPUT_MD_PATH}")

if __name__ == "__main__":
    main()