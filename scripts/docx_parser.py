# -*- coding: utf-8 -*-
"""
docx_parser.py —— Word(.docx/.doc) 文本提取 + PDF 兜底

功能：
1. python-docx 提取 .docx 文本（按段落保留）；
2. .doc 老格式：优先用 Word COM（win32com）转换为临时 .docx 后走原解析路径；
   本机无 Word 或无 pywin32 时返回中文提示"请用 Word 另存为 .docx 后重试"；
3. pdfplumber 兜底处理误放的 .pdf（"能读则读"，不保证完整）；
4. 解析结果按文件名留档到系统临时目录备查（P3 起不写入业务目录）；
5. 从标书全文中提取「评分办法/评分细则」等评分章节文本（供 AI 逐条核对客观分项）；
6. extract_candidate_phrases()：从未匹配标书标题区（前 500 字）提取候选项目短语（建议 8）。

主路径支持 .docx/.doc/.pdf；其他格式给出中文提示，不阻塞处理。
"""

from __future__ import annotations

import os
import re
import tempfile

# 解析文本留档默认目录（P3 起中间产物写入系统临时目录，不污染业务目录）
DEFAULT_PREVIEW_DIR = os.path.join(tempfile.gettempdir(), "投标评估助手", "标书解析文本")

# 评分章节标题关键词（覆盖常见命名：评分办法/评标办法/评分细则/评分标准/评审办法/评审标准/评标标准/评分内容/评审因素/评分表/评审细则）
SCORING_KEYWORDS = (
    "评分办法", "评标办法", "评分细则", "评分标准", "评审办法", "评审标准",
    "评标标准", "评分内容", "评审因素", "评分表", "评审细则",
)
# 评分章节最大截取长度（防溢出）
SCORING_MAX_LEN = 6000


class DocxParseError(Exception):
    """标书解析过程中的中文可读异常。"""


def _safe_filename(name: str) -> str:
    """清洗文件名，去掉非法字符，避免写入出错。"""
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    return name.strip() or "未命名"


def extract_docx(path: str) -> str:
    """用 python-docx 提取 .docx 文本，按文档顺序保留段落与表格（表格行以 | 分隔）。"""
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    try:
        doc = Document(path)
    except Exception as e:
        raise DocxParseError("无法解析 Word 文件「%s」：%s" % (os.path.basename(path), e))

    parts = []
    for child in doc.element.body.iterchildren():
        tag = child.tag
        if tag.endswith("}p"):
            text = Paragraph(child, doc).text.strip()
            if text:
                parts.append(text)
        elif tag.endswith("}tbl"):
            table = Table(child, doc)
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if cells:
                    parts.append(" | ".join(cells))
    return "\n".join(parts)


def extract_pdf(path: str) -> str:
    """用 pdfplumber 兜底提取 PDF 文本。"""
    import pdfplumber

    try:
        with pdfplumber.open(path) as pdf:
            pages = []
            for page in pdf.pages:
                text = page.extract_text() or ""
                if text.strip():
                    pages.append(text)
        return "\n".join(pages)
    except Exception as e:
        raise DocxParseError("PDF 兜底解析失败「%s」：%s" % (os.path.basename(path), e))


def _convert_doc_to_docx(path: str) -> str:
    """
    用 Word COM（win32com）将 .doc 转换为临时 .docx，返回临时文件绝对路径。

    本机未安装 pywin32 或 Microsoft Word 时，抛出中文提示
    "该文件为 .doc 老格式，请用 Word 另存为 .docx 后重试"。
    """
    try:
        import win32com.client
        import pythoncom
    except ImportError:
        raise DocxParseError(
            "标书「%s」：该文件为 .doc 老格式，请用 Word 另存为 .docx 后重试"
            "（本机未安装 pywin32 组件，可执行 pip install pywin32 后支持自动转换）。"
            % os.path.basename(path)
        )

    import tempfile
    tmp_dir = tempfile.mkdtemp(prefix="bid_doc_convert_")
    tmp_path = os.path.join(
        tmp_dir, _safe_filename(os.path.splitext(os.path.basename(path))[0]) + ".docx"
    )
    word = None
    pythoncom.CoInitialize()
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        try:
            doc = word.Documents.Open(os.path.abspath(path), ReadOnly=True)
        except Exception:
            doc = word.Documents.Open(os.path.abspath(path))
        try:
            # FileFormat=16 即 wdFormatXMLDocument（.docx）
            doc.SaveAs(tmp_path, FileFormat=16)
        finally:
            doc.Close(False)
    except DocxParseError:
        raise
    except Exception:
        raise DocxParseError(
            "标书「%s」：该文件为 .doc 老格式，请用 Word 另存为 .docx 后重试"
            "（自动转换失败，可能本机未安装 Microsoft Word）。" % os.path.basename(path)
        )
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()
    return tmp_path


def _remove_tmp_dir(path: str) -> None:
    """清理 Word COM 转换产生的临时文件/目录，失败不影响主流程。"""
    try:
        import shutil
        shutil.rmtree(os.path.dirname(path), ignore_errors=True)
    except Exception:
        pass


def parse_file(path: str, save_preview: bool = True, preview_dir: str = None) -> str:
    """
    解析单个标书文件，返回纯文本。

    参数：
    - path: 文件绝对路径
    - save_preview: 是否将解析结果留档（P3 起默认写入系统临时目录 %TEMP%/投标评估助手/标书解析文本/）
    - preview_dir: 留档目录（默认系统临时目录，不写入业务目录）
    返回：解析出的文本（可为空字符串）

    .doc 老格式：优先 Word COM（win32com）转换后解析；
    本机无 Word 或未装 pywin32 时抛出中文提示"请用 Word 另存为 .docx 后重试"。
    """
    if not os.path.exists(path):
        raise DocxParseError("未找到标书文件：%s" % path)

    ext = os.path.splitext(path)[1].lower()
    basename = os.path.basename(path)

    tmp_path = None
    try:
        if ext == ".docx":
            text = extract_docx(path)
        elif ext == ".doc":
            tmp_path = _convert_doc_to_docx(path)
            text = extract_docx(tmp_path)
        elif ext == ".pdf":
            text = extract_pdf(path)
        else:
            raise DocxParseError(
                "标书文件「%s」格式为「%s」，暂不支持。请使用 .docx 格式（Word 另存为即可）。"
                % (basename, ext)
            )

        if save_preview:
            _save_preview(path, text, preview_dir)

        return text
    finally:
        if tmp_path:
            _remove_tmp_dir(tmp_path)


# 候选项目短语：明显泛化的短语不进入候选（避免把"政府采购项目"等当项目名）
GENERIC_PHRASES = (
    "政府采购项目", "采购项目", "招标采购项目", "政府采购", "招标采购",
    "公开招标", "竞争性磋商", "政府采购招标文件", "政府采购公告",
)

# 标题区的标签行前缀（采购机构/项目编号等），其后的"X采购"不是项目名
LABEL_PREFIXES = (
    "采购机构", "采购人", "采购代理", "采购单位", "委托单位", "代理机构",
    "备案单位", "项目编号", "采购编号", "项目名称", "委托方", "联系方式",
)


def extract_candidate_phrases(text: str, max_phrases: int = 3) -> list:
    """
    从未匹配标书文本前部（标题区前 500 字）提取候选项目短语（建议 8）。

    提取顺序（含"项目/采购/招标"特征）：
    1. 「项目名称：xxx」/「项目名称:xxx」的值；
    2. 含"项目"且含"采购/招标"特征的短行（≤60 字，非目录行）；
    3. 以"采购项目/采购"结尾的短句。
    返回去重后的短语列表（最多 max_phrases 条）；无候选返回 []。
    """
    if not text:
        return []
    head = text[:500]

    phrases = []

    for m in re.finditer(r"项目名称[：:]\s*([^\s，。；、\n]{4,80})", head):
        p = m.group(1).strip()
        if len(p) >= 4:
            phrases.append(p)

    for line in head.splitlines():
        s = line.strip()
        if not s or len(s) < 6 or len(s) > 60:
            continue
        if _is_toc_line(s):
            continue
        if s.startswith(LABEL_PREFIXES):
            continue
        if "项目" in s and ("采购" in s or "招标" in s):
            phrases.append(s)

    for m in re.finditer(r"([^\s，。；、\n]{6,40}(?:采购项目|采购))", head):
        p = m.group(1).strip()
        # 去掉"项目名称：""名称："等标签前缀
        p = re.sub(r"^(项目名称|项目|名称)[：:]", "", p)
        if p.startswith(LABEL_PREFIXES):
            continue
        phrases.append(p)

    # 去重、剔除泛化短语与过短干扰项，保留提取顺序
    seen = set()
    out = []
    for p in phrases:
        p = p.strip(" \t\u3000·—-_，。；、（）()")
        if not p or len(p) < 6 or p in GENERIC_PHRASES or p in seen:
            continue
        seen.add(p)
        out.append(p)
        if len(out) >= max_phrases:
            break
    return out


def _is_chapter_heading(line: str) -> bool:
    """判断是否为「第X章/第X部分/附件/附录」等明显章节标题。"""
    s = line.strip()
    if re.match(r"^第[一二三四五六七八九十百0-9]+章", s):
        return True
    if re.match(r"^第[一二三四五六七八九十百0-9]+部分", s):
        return True
    if re.match(r"^(附件|附录|附表)", s):
        return True
    return False


def _is_toc_line(line: str) -> bool:
    """判断是否为目录行（短行 + 页码），如『第六章 评标办法\t46』『- 73 -』。"""
    s = line.strip()
    if not s or len(s) > 40:
        return False
    if re.search(r"[\t\s]+\d+\s*$", s):
        return True
    if re.search(r"-\s*\d+\s*-\s*$", s):
        return True
    return False


def _is_top_heading(line: str) -> bool:
    """判断是否为「一、二、三」级顶层标题。"""
    return re.match(r"^[一二三四五六七八九十]{1,3}、", line.strip()) is not None


# 评分内容特征词（用于给候选评分选段打分）
SCORE_WORDS = (
    "得分", "分值", "评分因素", "评分项", "评分细则", "评分标准", "评分表",
    "综合评分", "满分", "评分内容", "评分办法", "评标标准", "商务技术分",
)


def _is_scoring_title(line: str) -> bool:
    """判断是否为评分章节短标题行（短、无表格分隔符、非目录行、含评分关键词）。"""
    s = line.strip()
    if not s or len(s) > 40 or "|" in s:
        return False
    if _is_toc_line(s):
        return False
    return any(kw in s for kw in SCORING_KEYWORDS)


def _score_block(lines: list) -> int:
    """对候选选段打分：得分/分值/满分行 +5，评分特征行 +2，含"分"的表格行 +1，目录行 -5。"""
    score = 0
    for ln in lines:
        s = ln.strip()
        if _is_toc_line(s):
            score -= 5
        elif any(w in s for w in ("得分", "分值", "满分")):
            score += 5
        elif any(w in s for w in SCORE_WORDS):
            score += 2
        elif "|" in s and "分" in s:
            score += 1
    return score


def extract_scoring_section(text: str, max_len: int = None) -> str:
    """
    从标书全文提取评分章节文本（供 AI 逐条核对客观分项）。

    策略：短标题锚点 + 选段打分。
    1. 候选锚点分两类：
       - 首选：短标题行（≤40 字、无表格分隔符、非目录行）含评分关键词（如"四、评分标准""第四部分 评标办法"）；
       - 兜底：任意含评分关键词的行（覆盖评分条款写在长段落里的标书）。
    2. 锚点行向后收集至下一个「第X章/第X部分/附件」标题（上限 400 行）作为候选段；
    3. 按特征行打分（得分/分值/满分 +5，评分特征词 +2，含"分"的表格行 +1，目录行 -5），
       取最高分段；得分 ≤0 视为无有效评分章节，返回空字符串；
    4. 截断 max_len 字（按行截断，不切断整行）；
    找不到评分章节时返回空字符串，由调用方标记"未提取到评分章节"（AI 从全文兜底查找）。
    """
    if not text:
        return ""

    if max_len is None:
        max_len = SCORING_MAX_LEN
    lines = text.splitlines()

    title_anchors = [i for i, ln in enumerate(lines) if _is_scoring_title(ln)]
    any_anchors = [
        i for i, ln in enumerate(lines)
        if any(kw in ln for kw in SCORING_KEYWORDS)
    ] if not title_anchors else []

    anchors = title_anchors or any_anchors
    if not anchors:
        return ""

    best = None  # (score, start, end)
    for i in anchors:
        j = i
        while j + 1 < len(lines) and (j - i) < 400 and not _is_chapter_heading(lines[j + 1]):
            j += 1
        score = _score_block(lines[i:j + 1])
        if best is None or score > best[0]:
            best = (score, i, j)

    if best is None or best[0] <= 0:
        return ""

    _, start, end = best
    collected = []
    total_len = 0
    for ln in lines[start:end + 1]:
        if total_len + len(ln) + 1 > max_len:
            break
        collected.append(ln)
        total_len += len(ln) + 1
    return "\n".join(collected)


def _save_preview(path: str, text: str, preview_dir: str = None) -> None:
    """将解析文本留档（默认系统临时目录 %TEMP%/投标评估助手/标书解析文本/）。"""
    save_dir = preview_dir or DEFAULT_PREVIEW_DIR
    os.makedirs(save_dir, exist_ok=True)
    basename = os.path.splitext(os.path.basename(path))[0]
    out_path = os.path.join(save_dir, _safe_filename(basename) + ".txt")
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(text or "（解析结果为空，请检查标书内容）")
    except Exception:
        # 留档失败不影响主流程
        pass


def list_bid_files(bid_dir: str) -> list:
    """列出标书目录下的 .docx/.doc/.pdf 文件，返回绝对路径列表。"""
    if not os.path.isdir(bid_dir):
        return []
    files = []
    for name in os.listdir(bid_dir):
        if name.startswith("."):
            continue
        if name.lower().endswith((".docx", ".doc", ".pdf")):
            files.append(os.path.join(bid_dir, name))
    return sorted(files)