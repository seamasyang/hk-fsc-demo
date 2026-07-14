#!/usr/bin/env python3
"""
财报PDF自动拆分工具

将联交所上市公司的年度/中期财报PDF，按目录大纲和附注编号拆分为多个小PDF文件。
便于审计人员逐项检查，防止错误信息披露。

用法:
    python split_report_pypdf.py <pdf_path> [options]
    python split_report_pypdf.py --test <pdf_path>   # 执行测试：逐个阶段运行并输出诊断信息
    python split_report_pypdf.py --debug <pdf_path>   # 详细调试模式

选项:
    --output-dir <dir>    输出目录 (默认: data/processed/)
    --dry-run             仅分析，不生成文件
    --verbose             显示详细进度
    --debug               显示调试信息（每页文本预览等）
    --phase <phase>       仅运行指定阶段: toc|fs|notes|all (默认: all)
    --test <pdf_path>     测试模式：逐个阶段运行并输出诊断信息
"""

import argparse
import json
import os
import re
import sys
from pathlib import Path
from typing import Any

from pypdf import PdfReader, PdfWriter


# ──────────────────────────── 配置 ────────────────────────────

NOTES_KEYWORDS = [
    "附註",
    "附注",
    "财务报表附注",
    "財務報表附註",
    "notes to the financial statements",
    "notes to financial statements",
    "notes",
]

# 各财务表的标头关键词（按显示顺序排列）
FINANCIAL_STATEMENT_KEYWORDS = [
    "綜合收益表",
    "綜合損益及其他全面收益表",
    "綜合損益表",
    "综合损益及其他全面收益表",
    "綜合全面收益表",
    "综合全面收益表",
    "綜合資產負債表",
    "综合资产负债表",
    "綜合財務狀況表",
    "综合财务状况表",
    "綜合權益變動表",
    "综合权益变动表",
    "綜合現金流量表",
    "综合现金流量表",
]

ILLEGAL_CHARS_PATTERN = re.compile(r'[\\/:*?"<>|]')
NOTE_HEADER_PATTERN = re.compile(r"^(\d{1,3})\.\s*(.*)$")

# ──────────────────────────── 工具函数 ────────────────────────────


def sanitize_filename(name: str) -> str:
    """清洗文件名中的非法字符，并移除可能产生 illegal byte sequence 的乱码字符。

    只保留字母、数字、CJK 字符、空格、下划线、连字符和句点。
    """
    name = name.strip()
    # 移除任何非 ASCII + 非 CJK 的扩展字符，防止 macOS illegal byte sequence
    cleaned = []
    for ch in name:
        cp = ord(ch)
        # 允许: ASCII 可打印字符 (32-126)、CJK 统一表意文字 (U+4E00-U+9FFF)、
        # CJK 扩展 A (U+3400-U+4DBF)、CJK 符号 (U+3000-U+303F)、
        # 全角标点 (U+FF00-U+FFEF)
        if cp < 128:
            cleaned.append(ch)
        elif (0x4E00 <= cp <= 0x9FFF) or (0x3400 <= cp <= 0x4DBF) or (0x3000 <= cp <= 0x303F) or (0xFF00 <= cp <= 0xFFEF):
            cleaned.append(ch)
        # 跳过所有其他字符（乱码）
    cleaned_str = "".join(cleaned)
    return ILLEGAL_CHARS_PATTERN.sub("_", cleaned_str)


def is_title_garbled(title: str) -> bool:
    """检测文本提取是否产生乱码。

    如果标题中非 ASCII 字符大部分不是 CJK 统一表意文字，则判定为乱码。
    """
    if not title:
        return True
    non_ascii_count = 0
    valid_cjk_count = 0
    for ch in title:
        if ord(ch) > 127:
            non_ascii_count += 1
            cp = ord(ch)
            if (0x4E00 <= cp <= 0x9FFF) or (0x3400 <= cp <= 0x4DBF):
                valid_cjk_count += 1
    if non_ascii_count == 0:
        return False
    # 如果超过 50% 的非 ASCII 字符不是有效 CJK，判定为乱码
    garbled_ratio = 1.0 - (valid_cjk_count / non_ascii_count) if non_ascii_count > 0 else 0
    return garbled_ratio > 0.5


def info(msg: str, verbose: bool = False) -> None:
    """输出信息。"""
    if verbose:
        print(f"[INFO] {msg}")


def debug(msg: str, debug_mode: bool = False) -> None:
    """输出调试信息。"""
    if debug_mode:
        print(f"[DEBUG] {msg}")


def warn(msg: str) -> None:
    """输出警告。"""
    print(f"[WARN] {msg}", file=sys.stderr)


def err(msg: str) -> None:
    """输出错误并退出。"""
    print(f"[ERROR] {msg}", file=sys.stderr)
    sys.exit(1)


def _find_header_on_page(page_text: str, keywords: list[str]) -> str | None:
    """扫描页面文本，返回第一个匹配到的关键词标头，或 None。"""
    for line in page_text.split("\n"):
        ls = line.strip()
        for kw in keywords:
            if kw in ls:
                return kw
    return None


def _preview_text(text: str, max_len: int = 200) -> str:
    """截取文本前 max_len 个字符用于调试预览。"""
    preview = text[:max_len].replace("\n", "\\n")
    if len(text) > max_len:
        preview += "..."
    return preview


# ──────────────────────────── Phase 0: 大纲提取 ────────────────────────────


def _flatten_outline(
    reader: PdfReader, items: list[Any] | None = None, level: int = 1
) -> list[tuple[int, str, int]]:
    """将 pypdf 的分级大纲结构展开为扁平的 (level, title, page_num) 列表。

    返回:
        [(level, title, page_num), ...]
    """
    if items is None:
        items = reader.outline

    result: list[tuple[int, str, int]] = []

    for item in items:
        if isinstance(item, list):
            # 子大纲列表：保持同级
            result.extend(_flatten_outline(reader, item, level))
        else:
            try:
                page_num = reader.get_destination_page_number(item) + 1  # 转为 1-indexed
            except Exception:
                page_num = 1
            title = item.title.strip() if item.title else ""
            result.append((level, title, page_num))

            # 处理子节点：children() 返回列表
            kids = item.children()
            if kids:
                result.extend(_flatten_outline(reader, kids, level + 1))

    return result


def _resolve_page_num(label: str, label_to_physical: dict[str, int]) -> int | None:
    """将 TOC 文本中的页码标签转换为 0-based 物理页索引。

    优先精确匹配 page_labels（如 'a1' → physical_idx=0）；
    若因前导零（如"086"）匹配失败，尝试去除前导零后重试；
    最后回退为数字页码减 1（旧行为，兼容无 page labels 的 PDF）。
    """
    # 精确匹配 page labels（如 "a1", "ii", "1"）
    if label in label_to_physical:
        return label_to_physical[label]
    # 尝试去除前导零后重试（如 TOC 中 "086" → page_label 中 "86"）
    stripped = label.lstrip('0')
    if stripped and stripped in label_to_physical:
        return label_to_physical[stripped]
    # 回退：纯数字页码 → 当作物理索引（兼容无 page labels 的 PDF）
    if label.isdigit():
        return int(label) - 1
    # 无法解析
    return None


def _build_label_to_physical(reader: PdfReader) -> dict[str, int]:
    """构建页面标签 → 物理页索引（0-based）的映射。

    如果所有标签都是纯数字且连续（1,2,3...），说明 PDF 没有自定义 page labels，
    此时返回空字典，让调用方使用旧有回退逻辑，避免 TOC 中非数字页码（如"封面"）被误解析。
    """
    try:
        labels = reader.page_labels
    except Exception:
        return {}

    if not labels:
        return {}

    # 判断是否全部为纯数字且连续递增（无自定义 page labels 的默认情况）
    all_numeric = all(lbl.isdigit() for lbl in labels)
    if all_numeric:
        # 检查是否连续: labels 是 0-based，索引 i 的标签是 str(i+1)
        is_consecutive = all(int(lbl) == i + 1 for i, lbl in enumerate(labels))
        if is_consecutive:
            # 纯数字连续 = 无自定义 page labels，没必要建立映射
            return {}

    # 有自定义 page labels（如 a1, a2, 1, 2, ...），建立映射
    return {lbl: idx for idx, lbl in enumerate(labels)}


def extract_toc(reader: PdfReader, debug_mode: bool = False) -> list[dict[str, Any]]:
    """从PDF提取大纲。

    优先使用内置书签；书签不足时回退到文本解析目录页。
    文本回退时会利用 PDF page_labels 将目录中的页码精确映射到物理页索引。

    返回:
        [{level, title, page_num}, ...]
    """
    flat = _flatten_outline(reader)
    debug(f"内置书签扁平化结果: {len(flat)} 条", debug_mode)
    for item in flat:
        debug(f"  书签: L{item[0]} '{item[1]}' → 页 {item[2]}", debug_mode)

    if len(flat) >= 3:
        info(f"使用内置书签 ({len(flat)} 条)", verbose=True)
        return [
            {"level": item[0], "title": item[1], "page_num": item[2]}
            for item in flat
        ]

    # 回落：文本解析目录页
    return _extract_toc_from_text(reader, debug_mode)


def _extract_toc_from_text(reader: PdfReader, debug_mode: bool = False) -> list[dict[str, Any]]:
    """从文本目录了提取大纲。

    利用 PDF page_labels 将目录中的逻辑页码（如 "5", "a1"）精确映射到物理页索引，
    避免因前置封面/目录页导致的偏移错误。
    """
    info("内置书签不足，尝试从文本目录页提取大纲", verbose=True)

    # 构建 page_labels 映射（仅当 PDF 有自定义标签时返回非空字典）
    label_to_physical = _build_label_to_physical(reader)
    if label_to_physical:
        info(f"  使用 PDF page_labels 映射 ({len(label_to_physical)} 条)", verbose=True)
        if debug_mode:
            # 显示前几个标签的映射关系
            sample_items = list(label_to_physical.items())[:5]
            for lbl, idx in sample_items:
                debug(f"  page_label '{lbl}' → physical page {idx}", debug_mode)

    # 查找"目录"或"目錄"所在页
    toc_page_num = -1
    for i in range(min(len(reader.pages), 10)):
        text = reader.pages[i].extract_text()
        
        if (("目錄" in text) or ("目录" in text) or ("目 錄" in text)):
            toc_page_num = i
            info(f"  在第 {i+1} 页找到目录关键词", verbose=True)
            break

    if toc_page_num == -1 :
        err("无法找到目录页（目錄/目录），且内置书签不足。请手动检查PDF。")

    page = reader.pages[toc_page_num]
    text = page.extract_text()
    debug(f"目录页 ({toc_page_num+1}) 完整文本:\n{text}", debug_mode)
    lines = text.split("\n")

    entries: list[dict[str, Any]] = []
    for line in lines:
        line = line.strip()
        if not line or line in ("目錄", "目录", "目 錄"):
            continue
        
        # try title...page
        m = re.match(r"^(.+?)\s+(\d+)$", line)
        if m:
            title = m.group(1).strip()
            page_label = m.group(2).strip()
            if label_to_physical:
                physical_idx = _resolve_page_num(page_label, label_to_physical)
                if physical_idx is not None:
                    page_num = physical_idx + 1  # 转为 1-indexed
                else:
                    warn(f"  无法解析页码 '{page_label}'（条目: '{title}'），跳过此条目")
                    continue
            else:
                page_num = int(page_label)
            entries.append({"level": 1, "title": title, "page_num": page_num})
            debug(f"  解析条目: '{title}' → 页 {page_num}", debug_mode)
        
        
        # try page...title
        m = re.match(r"^(\d+)\s+(.+)$", line)
        if m:
            title = m.group(2).strip()
            page_label = m.group(1).strip()
            if label_to_physical:
                physical_idx = _resolve_page_num(page_label, label_to_physical)
                if physical_idx is not None:
                    page_num = physical_idx + 1  # 转为 1-indexed
                else:
                    warn(f"  无法解析页码 '{page_label}'（条目: '{title}'），跳过此条目")
                    continue
            else:
                page_num = int(page_label)
            entries.append({"level": 1, "title": title, "page_num": page_num})
            debug(f"  解析条目: '{title}' → 页 {page_num}", debug_mode)
       
    if not entries:
        err("无法从目录页提取条目。内置书签不足且无结构化的目录文本。")

    if len(entries) < 2:
        err(f"仅提取到 {len(entries)} 个条目，不足以作为可靠的目录。")

    info(f"从文本目录提取到 {len(entries)} 个条目", verbose=True)
    return entries


# ──────────────────────────── Phase 1: 按大纲拆分 ────────────────────────────


def compute_page_ranges(
    toc: list[dict[str, Any]], total_pages: int
) -> list[dict[str, Any]]:
    """计算每个大纲条目的页码范围。

    返回:
        [{level, title, page_num, start, end}, ...]
    """
    result: list[dict[str, Any]] = []
    n = len(toc)

    for i, entry in enumerate(toc):
        level = entry["level"]
        title = entry["title"]
        start = entry["page_num"] - 1  # 转为 0-indexed
        end = total_pages - 1  # 默认到文档末尾

        # 查找下一个同级或更高级别条目作为结束边界
        for j in range(i + 1, n):
            if toc[j]["level"] <= level:
                end = toc[j]["page_num"] - 2  # 上一个条目的最后一页
                break

        # 确保 end >= start
        end = max(end, start)

        result.append(
            {
                "level": level,
                "title": title,
                "page_num": entry["page_num"],
                "start": start,
                "end": end,
            }
        )

    return result


def build_folder_structure(
    ranges: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """构建层级结构并计算各级别的页码范围。

    返回:
        [{level, title, start, end, path, folder_path}, ...]
    """
    # 构建路径
    path_stack: list[str] = []
    level_stack: list[int] = []

    for item in ranges:
        level = item["level"]
        safe_title = sanitize_filename(item["title"])

        # 修剪栈到当前级别
        while level_stack and level_stack[-1] >= level:
            level_stack.pop()
            path_stack.pop()

        level_stack.append(level)
        path_stack.append(safe_title)

        folder_path = "/".join(path_stack[:-1]) if len(path_stack) > 1 else ""
        file_path = "/".join(path_stack) + ".pdf"

        item["path"] = file_path
        item["folder_path"] = folder_path

    return ranges


def split_by_toc(
    reader: PdfReader,
    ranges: list[dict[str, Any]],
    output_dir: str,
    dry_run: bool,
) -> dict[str, dict[str, Any]]:
    """按大纲拆分为多份PDF。

    返回:
        {章节路径: 章节信息}
    """
    sections: dict[str, dict[str, Any]] = {}
    os.makedirs(output_dir, exist_ok=True)

    for item in ranges:
        start = item["start"]
        end = item["end"]
        title = item["title"]
        path = item["path"]
        folder_path = item["folder_path"]

        # 创建文件夹
        if folder_path:
            Path(os.path.join(output_dir, folder_path)).mkdir(parents=True, exist_ok=True)

        full_path = os.path.join(output_dir, path)

        if dry_run:
            info(f"[DRY-RUN] 章节: {title} ({start+1}-{end+1}) → {path}", verbose=True)
        else:
            writer = PdfWriter()
            try:
                for i in range(start, end + 1):
                    writer.add_page(reader.pages[i])
                with open(full_path, "wb") as f:
                    writer.write(f)
                info(f"  章节: {title} ({start+1}-{end+1}) → {path}", verbose=True)
            finally:
                writer.close()

        sections[path] = {
            "title": title,
            "level": item["level"],
            "pages": [start + 1, end + 1],
            "path": path,
        }

    return sections


# ──────────────────────────── Phase 2: 财务报表内嵌拆分 ────────────────────────────


def scan_financial_statements_in_section(
    reader: PdfReader, section: dict[str, Any],
    notes_start_page: int | None = None,
    debug_mode: bool = False,
) -> list[dict[str, Any]]:
    """在章节页面范围内扫描财务报表标头。

    查找已知的财务表关键词（如"綜合收益表"、"綜合資產負債表"等），
    记录每个标头首次出现的页码。

    参数:
        notes_start_page: 如果提供，扫描将在此页之前停止（避免将附注内的引用误读为财务表头）

    返回:
        [{fs_keyword, title, start_page}, ...]  按出现顺序排列
    """
    statements: list[dict[str, Any]] = []
    seen_keywords: set[str] = set()

    start = section["start"]
    # 如果指定了附注起始页，只扫描附注开始之前的页面
    if notes_start_page is not None:
        end = min(section["end"], notes_start_page - 1)
    else:
        end = section["end"]

    for page_num in range(start, end + 1):
        text = reader.pages[page_num].extract_text()
        matched = _find_header_on_page(text, FINANCIAL_STATEMENT_KEYWORDS)
        if matched and matched not in seen_keywords:
            seen_keywords.add(matched)
            statements.append(
                {
                    "fs_keyword": matched,
                    "title": matched,
                    "start_page": page_num,
                }
            )
            debug(f"  页 {page_num+1}: 找到财务表 '{matched}'", debug_mode)

    return statements


def compute_fs_page_ranges(
    statements: list[dict[str, Any]], section_end: int,
    notes_start_page: int | None = None,
) -> list[dict[str, Any]]:
    """计算每个财务报表的页码范围。"""
    result: list[dict[str, Any]] = []
    n = len(statements)

    for i, stmt in enumerate(statements):
        start = stmt["start_page"]
        if i + 1 < n:
            end = statements[i + 1]["start_page"] - 1
        else:
            # 最后一个报表：如果有附注起始页，则截止到附注前一页
            end = section_end
            if notes_start_page is not None and notes_start_page - 1 < end:
                end = notes_start_page - 1

        end = max(end, start)

        result.append(
            {
                "title": stmt["title"],
                "start_page": start,
                "pages": [start, end],
            }
        )

    return result


def find_fs_section(ranges: list[dict[str, Any]]) -> dict[str, Any] | None:
    """查找包含财务报表的章节（"經審核財務報表"或类似名称）。"""
    fs_keywords = [
        "經審核財務報表",
        "经审核财务报表",
        "審核財務報表",
        "审核财务报表",
        "audited financial statements",
    ]

    # 优先匹配：章节标题包含"財務報表"
    for item in ranges:
        title_lower = item["title"].lower()
        for kw in fs_keywords:
            if kw.lower() in title_lower:
                return item

    # 回落：扫描内容找各财务表标头
    return None


def _safe_basename(name: str, fallback: str = "Untitled") -> str:
    """生成安全的文件名（不含扩展名）。"""
    safe = sanitize_filename(name)
    if not safe or safe.isspace():
        return fallback
    return safe


def split_financial_statements(
    reader: PdfReader,
    statements: list[dict[str, Any]],
    parent_folder: str,
    output_dir: str,
    dry_run: bool,
    verbose: bool,
) -> list[dict[str, Any]]:
    """按财务表拆分PDF。"""
    fs_results: list[dict[str, Any]] = []
    fs_folder = os.path.join(output_dir, parent_folder)
    if not dry_run:
        os.makedirs(fs_folder, exist_ok=True)

    for stmt in statements:
        start = stmt["start_page"]
        end = stmt["pages"][1]
        title = stmt["title"]
        safe_title = _safe_basename(title, "FinancialStatement")
        filename = f"{safe_title}.pdf"
        filepath = os.path.join(fs_folder, filename)

        if dry_run:
            info(
                f"[DRY-RUN] 财务表: {title} ({start+1}-{end+1}) → {parent_folder}/{filename}",
                verbose=True,
            )
        else:
            writer = PdfWriter()
            try:
                for i in range(start, end + 1):
                    writer.add_page(reader.pages[i])
                with open(filepath, "wb") as f:
                    writer.write(f)
                info(
                    f"  财务表: {title} ({start+1}-{end+1}) → {parent_folder}/{filename}",
                    verbose=verbose,
                )
            finally:
                writer.close()

        fs_results.append(
            {
                "title": title,
                "pages": [start + 1, end + 1],
                "path": os.path.join(parent_folder, filename),
            }
        )

    return fs_results


# ──────────────────────────── Phase 3: 附注识别 ────────────────────────────


def find_notes_chapter(
    ranges: list[dict[str, Any]],
) -> dict[str, Any] | None:
    """查找包含附注的章节。"""
    for item in ranges:
        title_lower = item["title"].lower()
        for kw in NOTES_KEYWORDS:
            if kw.lower() in title_lower:
                return item
    return None


def scan_notes_header_page(
    reader: PdfReader, section: dict[str, Any], debug_mode: bool = False,
) -> int | None:
    """在章节页面范围内扫描"財務報表附註"等关键标头首次出现的页码。

    返回:
        首次出现页的 0-indexed 页码，或 None
    """
    start = section["start"]
    end = section["end"]
    for page_num in range(start, end + 1):
        text = reader.pages[page_num].extract_text()
        for line in text.split("\n"):
            ls = line.strip()
            if "財務報表附註" in ls or "财务报表附注" in ls:
                debug(f"  在页 {page_num+1} 找到 '財務報表附註' 标头", debug_mode)
                return page_num
    return None


def scan_notes_in_section(
    reader: PdfReader, section: dict[str, Any],
    notes_start_page: int | None = None,
    debug_mode: bool = False,
) -> list[dict[str, Any]]:
    """在附注章节页面范围内扫描附注编号。

    处理两种格式：
    1. 同行: "1. 一般資料"
    2. 换行: "1. " 后跟 "一般資料" (下一行)

    返回:
        [{note_num, title, start_page}, ...]
    """
    notes: list[dict[str, Any]] = []
    start = section["start"]
    end = section["end"]

    for page_num in range(start, end + 1):
        page = reader.pages[page_num]
        text = page.extract_text()
        lines = text.split("\n")

        for idx, line in enumerate(lines):
            line_stripped = line.strip()
            m = NOTE_HEADER_PATTERN.match(line_stripped)
            if m:
                note_num = int(m.group(1))
                note_title = m.group(2).strip()

                # 如果本行没有标题（格式: "1." 单独一行），尝试从下一行获取标题
                if not note_title and idx + 1 < len(lines):
                    next_line = lines[idx + 1].strip()
                    # 避免取到页码或页眉
                    if next_line and not next_line.isdigit() and len(next_line) > 1:
                        note_title = next_line

                if note_num > 0:
                    # 过滤：跳过百分比误匹配（如 "96.04%" → title="04%"）
                    if "%" in note_title:
                        debug(f"  页 {page_num+1}: 跳过百分比误匹配 '{line_stripped}'", debug_mode)
                        continue
                    # 检测乱码：如果文本提取产生乱码，用空标题替代
                    if is_title_garbled(note_title):
                        debug(f"  页 {page_num+1}: 附注 {note_num} 文本 '{note_title}' 判定为乱码", debug_mode)
                        note_title = ""
                        warn(f"附注 {note_num} 文本提取为乱码，使用编号作为文件名")
                    # 去重
                    existing = [n for n in notes if n["note_num"] == note_num]
                    if not existing:
                        notes.append(
                            {
                                "note_num": note_num,
                                "title": note_title,
                                "start_page": page_num,
                            }
                        )
                        debug(f"  页 {page_num+1}: 找到附注 {note_num} '{note_title}'", debug_mode)
                    else:
                        debug(f"  页 {page_num+1}: 跳过重复附注 {note_num}", debug_mode)

    return notes


# ──────────────────────────── Phase 4: 按附注拆分 ────────────────────────────


def compute_note_page_ranges(
    notes: list[dict[str, Any]], section_end: int,
) -> list[dict[str, Any]]:
    """计算每个附注的页码范围。

    规则：
    - 正常情况下附注 N = [start, next_start - 1]（不重叠）
    - 如果下一附注与当前附注在同一页（next_start == start），
      则当前附注 = [start, start]（仅该页），下一附注也从该页开始（重叠）

    返回:
        [{note_num, title, pages: [start_page, end_page], ...}]
    """
    result: list[dict[str, Any]] = []
    n = len(notes)

    for i, note in enumerate(notes):
        start = note["start_page"]
        if i + 1 < n:
            next_start = notes[i + 1]["start_page"]
            if next_start == start:
                # 同一页有多个附注 → 当前附注仅占此页，下一附注也从同一页开始
                end = start
            else:
                # 正常情况 → 结束于下一附注前一页
                end = next_start - 1
        else:
            end = section_end

        end = max(end, start)

        result.append(
            {
                "note_num": note["note_num"],
                "title": note["title"],
                "start_page": start,
                "pages": [start, end],
            }
        )

    return result


def split_notes(
    reader: PdfReader,
    notes: list[dict[str, Any]],
    parent_folder: str,
    output_dir: str,
    dry_run: bool,
    verbose: bool,
) -> list[dict[str, Any]]:
    """按附注拆分PDF。"""
    note_results: list[dict[str, Any]] = []
    notes_folder = os.path.join(output_dir, parent_folder, "附註")
    if not dry_run:
        os.makedirs(notes_folder, exist_ok=True)

    for note in notes:
        start = note["start_page"]
        end = note["pages"][1]
        note_num = note["note_num"]
        title = note["title"]
        safe_title = sanitize_filename(title)
        if safe_title:
            filename = f"Note_{note_num}_{safe_title}.pdf"
        else:
            filename = f"Note_{note_num}.pdf"
        filepath = os.path.join(notes_folder, filename)

        if dry_run:
            info(
                f"[DRY-RUN] 附注 {note_num}: {title} ({start+1}-{end+1}) → {parent_folder}/附註/{filename}",
                verbose=True,
            )
        else:
            writer = PdfWriter()
            try:
                for i in range(start, end + 1):
                    writer.add_page(reader.pages[i])
                with open(filepath, "wb") as f:
                    writer.write(f)
                info(
                    f"  附注 {note_num}: {title} ({start+1}-{end+1}) → {parent_folder}/附註/{filename}",
                    verbose=verbose,
                )
            finally:
                writer.close()

        note_results.append(
            {
                "number": note_num,
                "title": title,
                "pages": [start + 1, end + 1],
                "path": os.path.join(parent_folder, "附註", filename),
            }
        )

    return note_results


# ──────────────────────────── 主流程 ────────────────────────────


def process_pdf(
    pdf_path: str,
    output_dir: str,
    dry_run: bool = False,
    verbose: bool = False,
    debug_mode: bool = False,
    phase: str = "all",
) -> dict[str, Any]:
    """处理单个PDF文件。

    参数:
        phase: "toc" | "fs" | "notes" | "all"  仅运行指定阶段

    返回:
        manifest 字典
    """
    if not os.path.exists(pdf_path):
        err(f"文件不存在: {pdf_path}")

    info(f"打开PDF: {pdf_path}", verbose=verbose)
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]

    info(f"总页数: {total_pages}", verbose=verbose)

    # ── Phase 0: 大纲提取 ──
    info("Phase 0: 提取大纲...", verbose=verbose)
    toc = extract_toc(reader, debug_mode)
    info(f"  提取 {len(toc)} 个大纲条目", verbose=verbose)
    if debug_mode:
        for item in toc:
            print(f"  [TOC] L{item['level']} '{item['title']}' → 页 {item['page_num']}")

    if phase == "toc":
        reader.stream.close()
        return {"source": os.path.basename(pdf_path), "total_pages": total_pages, "sections": toc, "notes": [], "financial_statements": []}

    # ── Phase 1: 按大纲拆分 ──
    info("Phase 1: 按大纲拆分...", verbose=verbose)
    ranges = compute_page_ranges(toc, total_pages)
    ranges = build_folder_structure(ranges)

    pdf_output_dir = os.path.join(output_dir, pdf_name)
    sections = split_by_toc(reader, ranges, pdf_output_dir, dry_run)

    # ── Phase 2: 财务报表内嵌拆分 ──
    info("Phase 2: 识别并拆分财务表...", verbose=verbose)
    fs_section = find_fs_section(ranges)
    fs_results: list[dict[str, Any]] = []

    if fs_section:
        # 先扫描附注标头，以限定财务报表的扫描范围
        fs_notes_header_page = scan_notes_header_page(reader, fs_section, debug_mode)

        # 在财务报表章节内扫描各财务表标头（在附注开始前停止扫描）
        fs_statements = scan_financial_statements_in_section(reader, fs_section, fs_notes_header_page, debug_mode)
        info(f"  扫描到 {len(fs_statements)} 个财务表标头", verbose=verbose)

        if len(fs_statements) >= 2:
            fs_ranges = compute_fs_page_ranges(fs_statements, fs_section["end"], fs_notes_header_page)
            parent_path = sanitize_filename(fs_section["title"])
            fs_results = split_financial_statements(
                reader, fs_ranges, parent_path, pdf_output_dir, dry_run, verbose
            )

            # 将财务表结果加入 sections（覆盖原来的合并章节）
            for fs in fs_results:
                section_key = os.path.join(parent_path, os.path.basename(fs["path"]))
                sections[section_key] = {
                    "title": fs["title"],
                    "level": 1,
                    "pages": fs["pages"],
                    "path": fs["path"],
                }
            # 标记原始合并章节已被拆分
            raw_section_key = fs_section["path"]
            if raw_section_key in sections:
                sections[raw_section_key]["sub_sections"] = len(fs_results)
        else:
            info("  财务表标头不足，跳过财务表拆分", verbose=verbose)
    else:
        info("  未找到财务报表章节，跳过财务表拆分", verbose=verbose)

    if phase == "fs":
        reader.stream.close()
        return {"source": os.path.basename(pdf_path), "total_pages": total_pages, "sections": list(sections.values()), "notes": [], "financial_statements": fs_results}

    # ── Phase 3: 附注识别 ──
    info("Phase 3: 识别附注章节...", verbose=verbose)
    notes_chapter = find_notes_chapter(ranges)

    # 如果章节标题未直接匹配，扫描各章节文本内容找"財務報表附註"
    if notes_chapter is None:
        info("  章节标题未直接匹配，尝试扫描章节文本内容...", verbose=verbose)
        for r in ranges:
            # 只扫描页数较多的章节（>= 20 页，通常是财务报告章节）
            if r["end"] - r["start"] >= 20:
                notes_header_page = scan_notes_header_page(reader, r, debug_mode)
                if notes_header_page is not None:
                    notes_chapter = r
                    info(f"  在章节内找到 '財務報表附註' 标头: {r['title']} (页 {notes_header_page+1})", verbose=verbose)
                    break

    notes_extracted: list[dict[str, Any]] = []

    if notes_chapter:
        info(f"  附注章节: {notes_chapter['title']} (页 {notes_chapter['start']+1}-{notes_chapter['end']+1})", verbose=verbose)
        notes_start_page = scan_notes_header_page(reader, notes_chapter, debug_mode)
        if notes_start_page is not None:
            info(f"  附注起始页: {notes_start_page + 1}", verbose=verbose)
        notes_extracted = scan_notes_in_section(reader, notes_chapter, notes_start_page, debug_mode)
        info(f"  扫描到 {len(notes_extracted)} 个附注", verbose=verbose)
        if debug_mode:
            for n in notes_extracted:
                print(f"  [NOTE] #{n['note_num']} '{n['title']}' → 页 {n['start_page']+1}")

        if notes_extracted:
            # ── Phase 4: 按附注拆分 ──
            info("Phase 4: 按附注拆分...", verbose=verbose)
            note_ranges = compute_note_page_ranges(notes_extracted, notes_chapter["end"])
            parent_path = sanitize_filename(notes_chapter["title"])
            note_results = split_notes(
                reader, note_ranges, parent_path, pdf_output_dir, dry_run, verbose
            )
        else:
            note_results = []
            info("  附注章节内未发现附注编号模式，跳过附注拆分", verbose=verbose)
    else:
        info("  未找到附注章节，跳过附注拆分", verbose=verbose)
        note_results = []

    # ── 移除已拆分的合并章节 ──
    # 如果财务表章节已被拆分为子文件，删除原始合并文件
    if fs_section is not None and len(fs_results) > 0:
        combined_path = os.path.join(pdf_output_dir, fs_section["path"])
        if os.path.exists(combined_path) and not dry_run:
            os.remove(combined_path)
            info(f"  移除已拆分的合并章节: {fs_section['path']}", verbose=verbose)

    reader.stream.close()

    # ── 生成清单 ──
    manifest = {
        "source": os.path.basename(pdf_path),
        "total_pages": total_pages,
        "sections": list(sections.values()),
        "notes": note_results,
        "financial_statements": fs_results,
    }

    # 写入清单
    manifest_path = os.path.join(pdf_output_dir, "_manifest.json")
    if not dry_run:
        os.makedirs(pdf_output_dir, exist_ok=True)
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        info(f"清单: {manifest_path}", verbose=verbose)

    return manifest


# ──────────────────────────── 测试模式 ────────────────────────────


def run_test(pdf_path: str, verbose: bool = True) -> None:
    """测试模式：逐个阶段运行并输出诊断信息。"""
    if not os.path.exists(pdf_path):
        err(f"文件不存在: {pdf_path}")

    debug_mode = True

    print(f"\n{'='*60}")
    print(f"测试模式 — {os.path.basename(pdf_path)}")
    print(f"{'='*60}\n")

    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    print(f"总页数: {total_pages}")
    print(f"内置书签数: {len(list(reader.outline or []))}\n")

    # ── Test Phase 0: TOC ──
    print(f"{'─'*40}")
    print(f"Phase 0: 大纲提取")
    print(f"{'─'*40}")
    flat = _flatten_outline(reader)
    print(f"扁平化书签: {len(flat)} 条")
    for item in flat:
        print(f"  L{item[0]} '{item[1]}' → 页 {item[2]}")
    toc = extract_toc(reader, debug_mode=True)
    print(f"TOC 条目: {len(toc)}")
    for t in toc:
        print(f"  L{t['level']} '{t['title']}' → 页 {t['page_num']}")

    # ── Test Phase 1: Page ranges ──
    print(f"\n{'─'*40}")
    print(f"Phase 1: 页码范围 + 目录结构")
    print(f"{'─'*40}")
    ranges = compute_page_ranges(toc, total_pages)
    ranges = build_folder_structure(ranges)
    print(f"章节数: {len(ranges)}")
    for r in ranges:
        print(f"  '{r['title']}' → 页 {r['start']+1}-{r['end']+1} → {r.get('path', 'N/A')}")

    # ── Test Phase 2: Financial statements ──
    print(f"\n{'─'*40}")
    print(f"Phase 2: 财务报表识别")
    print(f"{'─'*40}")
    fs_section = find_fs_section(ranges)
    if fs_section:
        print(f"财务章节: '{fs_section['title']}' (页 {fs_section['start']+1}-{fs_section['end']+1})")
        fs_notes_header = scan_notes_header_page(reader, fs_section, debug_mode=True)
        if fs_notes_header is not None:
            print(f"附注标头页: {fs_notes_header + 1}")
        print(f"附注标头页: {fs_notes_header}")
        print(f"扫描范围: 页 {fs_section['start']+1}", end="")
        if fs_notes_header is not None:
            print(f"-{fs_notes_header}", end="")
        print()
        fs_statements = scan_financial_statements_in_section(reader, fs_section, fs_notes_header, debug_mode=True)
        print(f"财务表数: {len(fs_statements)}")
        for s in fs_statements:
            print(f"  '{s['title']}' → 页 {s['start_page']+1}")
        if len(fs_statements) >= 2:
            fs_ranges = compute_fs_page_ranges(fs_statements, fs_section["end"], fs_notes_header)
            print(f"财务表页码范围:")
            for fr in fs_ranges:
                print(f"  '{fr['title']}' → 页 {fr['pages'][0]+1}-{fr['pages'][1]+1}")
    else:
        print("未找到财务报表章节")

    # ── Test Phase 3: Notes ──
    print(f"\n{'─'*40}")
    print(f"Phase 3: 附注识别")
    print(f"{'─'*40}")
    notes_chapter = find_notes_chapter(ranges)
    if notes_chapter is None:
        print("章节标题未直接匹配，按内容扫描...")
        for r in ranges:
            if r["end"] - r["start"] >= 20:
                nh = scan_notes_header_page(reader, r, debug_mode=True)
                if nh is not None:
                    notes_chapter = r
                    print(f"  在章节 '{r['title']}' 内找到标头 (页 {nh+1})")
                    break
    if notes_chapter:
        print(f"附注章节: '{notes_chapter['title']}' (页 {notes_chapter['start']+1}-{notes_chapter['end']+1})")
        notes_start_page = scan_notes_header_page(reader, notes_chapter, debug_mode=True)
        print(f"附注起始页: {notes_start_page + 1 if notes_start_page is not None else 'N/A'}")

        # 扫描附注前打印一些文本预览帮助调试
        print(f"\n附注章节文本预览 (前 5 页):")
        for pn in range(notes_chapter["start"], min(notes_chapter["start"] + 5, notes_chapter["end"] + 1)):
            text = reader.pages[pn].extract_text()
            print(f"  页 {pn+1}: {_preview_text(text, 300)}")

        notes_extracted = scan_notes_in_section(reader, notes_chapter, notes_start_page, debug_mode=True)
        print(f"\n附注数: {len(notes_extracted)}")
        for n in notes_extracted:
            print(f"  #{n['note_num']} '{n['title']}' → 页 {n['start_page']+1}")

        # ── Test Phase 4: Note page ranges ──
        print(f"\n{'─'*40}")
        print(f"Phase 4: 附注页码范围")
        print(f"{'─'*40}")
        if notes_extracted:
            note_ranges = compute_note_page_ranges(notes_extracted, notes_chapter["end"])
            for nr in note_ranges:
                print(f"  Note {nr['note_num']}: 页 {nr['pages'][0]+1}-{nr['pages'][1]+1} ({nr['pages'][1]-nr['pages'][0]+1} 页)")
    else:
        print("未找到附注章节")

    reader.stream.close()
    print(f"\n{'='*60}")
    print(f"测试完成")
    print(f"{'='*60}\n")


# ──────────────────────────── CLI ────────────────────────────


def main() -> None:
    parser = argparse.ArgumentParser(
        description="财报PDF自动拆分工具 - 将PDF按大纲和附注拆分为多份小PDF"
    )
    parser.add_argument("pdf_path", nargs="?", help="输入PDF文件路径")
    parser.add_argument(
        "--output-dir",
        default="data/processed/",
        help="输出目录 (默认: data/processed/)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="仅分析，不生成文件",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="显示详细进度",
    )
    parser.add_argument(
        "--debug",
        action="store_true",
        help="显示调试信息（每页文本预览等）",
    )
    parser.add_argument(
        "--phase",
        choices=["toc", "fs", "notes", "all"],
        default="all",
        help="仅运行指定阶段: toc|fs|notes|all (默认: all)",
    )
    parser.add_argument(
        "--test",
        dest="test_mode",
        action="store_true",
        help="测试模式：逐个阶段运行并输出诊断信息",
    )

    args = parser.parse_args()

    # 测试模式不需要 pdf_path 参数（可以从 data/raw/ 下读取第一个 PDF）
    if args.test_mode:
        if args.pdf_path:
            pdf_path = args.pdf_path
        else:
            # 尝试从 data/raw/ 自动查找
            raw_dir = "data/raw"
            if os.path.isdir(raw_dir):
                pdfs = [f for f in os.listdir(raw_dir) if f.lower().endswith(".pdf")]
                if pdfs:
                    pdf_path = os.path.join(raw_dir, pdfs[0])
                    print(f"未指定 PDF，自动使用: {pdf_path}")
                else:
                    err(f"{raw_dir}/ 目录下没有 PDF 文件")
            else:
                err("请指定 PDF 文件路径，或确保 data/raw/ 目录下有 PDF 文件")
        run_test(pdf_path)
        return

    if not args.pdf_path:
        parser.print_help()
        sys.exit(1)

    print(f"\n{'='*60}")
    print(f"财报PDF自动拆分工具")
    print(f"{'='*60}")

    if args.dry_run:
        print("[DRY-RUN] 开启 — 不会生成任何文件")

    manifest = process_pdf(
        pdf_path=args.pdf_path,
        output_dir=args.output_dir,
        dry_run=args.dry_run,
        verbose=args.verbose,
        debug_mode=args.debug,
        phase=args.phase,
    )

    print(f"\n{'='*60}")
    print(f"处理完成")
    print(f"  源文件:    {manifest['source']}")
    print(f"  总页数:    {manifest['total_pages']}")
    print(f"  章节数:    {len(manifest['sections'])}")
    print(f"  附注数:    {len(manifest['notes'])}")
    out_base = args.output_dir.rstrip("/")
    print(f"  输出目录:  {out_base}/{os.path.splitext(manifest['source'])[0]}/")

    if not args.dry_run:
        print(f"  清单文件:  _manifest.json")

    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()