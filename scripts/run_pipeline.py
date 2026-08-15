# -*- coding: utf-8 -*-
"""
run_pipeline.py —— 一键批处理入口（P3 起三个路径由参数传入，无固定目录依赖）

流程：解析 Excel（多清单合并+表头探测）→ 企业资料摘要 → 初筛（时间+类型+项目类型+相关性软标记）
      → 匹配标书（映射表→去词缀模糊→Top3 候选）→ 解析文本 → 提取评分章节
      → 输出初筛结果 Excel + 待研判清单（中间产物写入系统临时目录）

用法：
    命令行直接运行（三个路径必填，由用户在对话中告知）：
        python scripts/run_pipeline.py --list-dir "清单路径" --bid-dir "标书路径" --profile-dir "公司资料路径"
        python scripts/run_pipeline.py --list-dir ... --bid-dir ... --profile-dir ... --json   # 打印待研判清单 JSON
        python scripts/run_pipeline.py --list-dir ... --bid-dir ... --profile-dir ... --incremental   # 增量模式（融合新增）
        python scripts/run_pipeline.py --list-dir ... --bid-dir ... --profile-dir ... --reset-status # 复评（重置状态后全量）
    被 SKILL.md 编排调用：
        from scripts.run_pipeline import run_pipeline
        result = run_pipeline(list_dir, bid_dir, profile_dir)
        result = run_pipeline(list_dir, bid_dir, profile_dir, incremental=True)

增量模式（--incremental，方案 5）：
- 增量过滤位置：load_excels_from_dir 之后、filter_screening 之前；
- 按地址链接URL 读取状态（主文件列 / 清单路径/增量状态.json，URL 唯一键），
  初筛状态=已初筛 的记录自动跳过；summary 增加 new_count/skipped_screened_count/skipped_judged_count；
- 初筛结果 Excel 追加 初筛状态/研判状态 两列；收尾 status_tool 回写状态；
- 不传 --incremental 时行为与融合前 100% 一致（不读状态、不回写、不追加列）。

产物归属（P3 已确认设计）：
- 初筛结果 Excel  → 清单路径/{时间}_初筛结果.xlsx（时间格式 2026-08-05_1530，防同日覆盖）
- 研判报告 md    → 标书路径/{时间}_研判报告.md（内容由 AI 研判后写入，脚本不生成）
- 项目类型选择    → 清单路径/项目类型选择.json（步骤 2 由 AI 写入）
- 标书映射        → 标书路径/标书映射.json（随标书走，建议 10）
- 中间产物        → %TEMP%/投标评估助手/（待研判清单.json、企业资料摘要.json、标书解析文本/）

说明：
- 匹配规则（建议 7/10）：映射表（标书路径/标书映射.json）优先 → 文件名包含采购项目名称
  （去词缀：招标文件/定稿/正式稿/最终版/电子标等）→ 仍不唯一时返回候选 Top3 交 AI 请用户确认。
- 未匹配标书自动预识别（建议 8）：解析标题区前 500 字，提取候选项目短语随未匹配清单呈现。
- 无法匹配的标书单独列出，不阻塞其他项目。
"""

from __future__ import annotations

import argparse
import json
import os
import sys
import tempfile
from datetime import datetime

# 项目根目录（本文件位于 scripts/ 下）
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT_DIR)

from scripts import excel_tool, docx_parser, company_profile, bid_mapping, status_tool  # noqa: E402

# 中间产物临时目录（P3：不写入业务目录，每次运行自动覆盖）
TEMP_BASE_DIR = os.path.join(tempfile.gettempdir(), "投标评估助手")
TEMP_PREVIEW_JSON = os.path.join(TEMP_BASE_DIR, "待研判清单.json")
TEMP_COMPANY_JSON = os.path.join(TEMP_BASE_DIR, "企业资料摘要.json")
TEMP_BID_TEXT_DIR = os.path.join(TEMP_BASE_DIR, "标书解析文本")


def normalize_name(name: str) -> str:
    """规范化名称用于匹配，去掉空格与常见标点。"""
    import re
    name = re.sub(r"[\s\u3000]+", "", str(name or ""))
    name = re.sub(r"[。，、；：（）()《》【】\[\]\-\—_]", "", name)
    return name.lower()


# 标书文件名词缀表（建议 7）：去词缀后再做包含匹配
BID_AFFIXES = (
    "招标文件", "竞争性磋商文件", "竞争性谈判文件", "磋商文件", "谈判文件", "询价文件",
    "采购文件", "电子招投标文件", "线上电子招投标", "招标公告", "采购公告", "竞争性磋商公告",
    "定稿发布", "定稿", "正式稿", "最终稿", "最终版", "电子标", "发布稿", "初稿",
    "文件", "公告",
)


def strip_affixes(name: str) -> str:
    """去除名称中的常见词缀（招标文件/磋商文件/定稿/正式稿/最终版/电子标/招标公告等）。"""
    cur = str(name or "")
    for _ in range(3):
        new = cur
        for affix in BID_AFFIXES:
            new = new.replace(affix, "")
        if new == cur:
            break
        cur = new
    return cur.strip(" \t\u3000-—_（）()")


def _common_prefix_len(a: str, b: str) -> int:
    """最长公共前缀长度。"""
    n = 0
    for x, y in zip(a, b):
        if x != y:
            break
        n += 1
    return n


def _common_substring_len(a: str, b: str) -> int:
    """最长公共子串长度（DP，字符串短，开销可忽略）。"""
    if not a or not b:
        return 0
    prev = [0] * (len(b) + 1)
    best = 0
    for i in range(len(a)):
        cur = [0] * (len(b) + 1)
        for j in range(len(b)):
            if a[i] == b[j]:
                cur[j + 1] = prev[j] + 1
                if cur[j + 1] > best:
                    best = cur[j + 1]
        prev = cur
    return best


def _name_similarity(a: str, b: str) -> float:
    """基于公共前缀 + 公共子串长度的相似度（0~1）。"""
    if not a or not b:
        return 0.0
    lcp = _common_prefix_len(a, b)
    lcs = _common_substring_len(a, b)
    max_len = max(len(a), len(b), 1)
    return round((lcp + lcs) / (2.0 * max_len), 3)


def _top_candidates(project_name: str, bid_files: list, top_n: int = 3) -> list:
    """
    按公共前缀/公共子串相似度返回 TopN 候选标书（建议 7），
    返回 [{"标书文件名": str, "相似度": float}]，供 AI 在对话中请用户确认。
    """
    pn = normalize_name(project_name)
    scored = []
    for bf in bid_files:
        base = os.path.splitext(os.path.basename(bf))[0]
        bn = normalize_name(base)
        bn_stripped = normalize_name(strip_affixes(base))
        sim = max(_name_similarity(pn, bn), _name_similarity(pn, bn_stripped))
        if sim > 0:
            scored.append({"标书文件名": os.path.basename(bf), "相似度": sim})
    scored.sort(key=lambda x: (-x["相似度"], x["标书文件名"]))
    return scored[:top_n]


def match_bid_file(project_name: str, bid_files: list) -> dict:
    """
    将项目与标书文件匹配（建议 7：去词缀模糊匹配 + 候选 Top3）。

    返回：
    {
      "bid_file": 唯一匹配的标书绝对路径（无唯一匹配则为 None），
      "way": "包含匹配" / "去词缀匹配" / ""（未匹配），
      "candidates": [{"标书文件名", "相似度"}]（未唯一匹配时的 Top3 候选，无则 []）
    }

    匹配优先级：
    1. 精确包含匹配：文件名（去扩展名）包含采购项目名称，或互为包含；
    2. 去词缀后包含匹配：去掉"招标文件/竞争性磋商文件/磋商文件/定稿/正式稿/最终版/电子标/招标公告"等词缀后再包含匹配；
    3. 仍无法唯一匹配（含多文件同时命中）时，返回相似度（公共前缀+公共子串）Top3 候选，
       由 AI 在对话中请用户确认，不自动猜测。
    """
    pn = normalize_name(project_name)
    pn_stripped = normalize_name(strip_affixes(project_name))
    if not pn:
        return {"bid_file": None, "way": "", "candidates": []}

    direct, stripped = [], []
    for bf in bid_files:
        base = os.path.splitext(os.path.basename(bf))[0]
        bn = normalize_name(base)
        bn_stripped = normalize_name(strip_affixes(base))
        if bn and (pn in bn or bn in pn):
            direct.append(bf)
        elif pn_stripped and bn_stripped and (pn_stripped in bn_stripped or bn_stripped in pn_stripped):
            stripped.append(bf)

    if len(direct) == 1:
        return {"bid_file": direct[0], "way": "包含匹配", "candidates": []}
    if len(direct) > 1:
        return {"bid_file": None, "way": "", "candidates": _top_candidates(project_name, direct)}
    if len(stripped) == 1:
        return {"bid_file": stripped[0], "way": "去词缀匹配", "candidates": []}
    if len(stripped) > 1:
        return {"bid_file": None, "way": "", "candidates": _top_candidates(project_name, stripped)}

    # 无包含匹配 → 返回相似度 Top3 候选（可能为空列表 = 无候选）
    return {"bid_file": None, "way": "", "candidates": _top_candidates(project_name, bid_files)}


def find_bid_by_mapping(project_name: str, mapping: dict, bid_files: list) -> str:
    """
    按映射表查找标书（建议 10）：映射值（采购项目名称）与项目名规范化后相等，
    返回映射键对应的标书绝对路径；未命中返回 None。
    """
    pn = normalize_name(project_name)
    if not pn or not mapping:
        return None
    by_base = {os.path.splitext(os.path.basename(f))[0]: f for f in bid_files}
    for bid_name, proj in mapping.items():
        if normalize_name(proj) == pn:
            key = os.path.splitext(bid_name)[0]
            if key in by_base:
                return by_base[key]
    return None


def _match_phrase_to_projects(phrase: str, passed_projects: list) -> list:
    """
    将候选短语与初筛通过的项目做包含匹配，返回 [{"项目名称", "来源", "行号"}]。
    仅提示用户确认，不做自动匹配。
    """
    pn = normalize_name(phrase)
    if not pn:
        return []
    hits = []
    for p in passed_projects:
        proj = normalize_name(p["project_name"])
        if proj and (pn in proj or proj in pn):
            hits.append({
                "项目名称": p["project_name"],
                "来源": p.get("source", ""),
                "行号": p.get("row_index", 0),
            })
    return hits


def _pre_recognize_bid(bid_path: str, passed_projects: list) -> dict:
    """
    未匹配标书自动预识别（建议 8）：解析标题区前 500 字，提取含"项目/采购/招标"
    特征的候选短语，并尝试与清单项目关联。返回：
    {
      "标书文件名", "候选短语": [...], "匹配到清单项目": [...], "无候选": bool, "错误": str
    }
    """
    basename = os.path.basename(bid_path)
    out = {"标书文件名": basename, "候选短语": [], "匹配到清单项目": [], "无候选": True, "错误": ""}
    try:
        text = docx_parser.parse_file(bid_path)
    except docx_parser.DocxParseError as e:
        out["错误"] = str(e)
        return out
    phrases = docx_parser.extract_candidate_phrases(text or "")
    out["候选短语"] = phrases
    out["无候选"] = not phrases
    if phrases:
        for ph in phrases:
            out["匹配到清单项目"].extend(_match_phrase_to_projects(ph, passed_projects))
    return out


def run_pipeline(
    list_dir: str,
    bid_dir: str,
    profile_dir: str,
    days: int = None,
    incremental: bool = False,
    reset_status: bool = False,
) -> dict:
    """
    执行完整流水线，返回结构化结果（供 SKILL.md 与 AI 研判使用）。

    参数（P3：三个路径每次由用户在对话中提供，无固定目录依赖）：
    - list_dir: 清单路径（放置投标机会清单 .xlsx 的文件夹）
    - bid_dir: 标书路径（放置标书 .docx/.doc/.pdf 的文件夹）
    - profile_dir: 公司资料路径（放置 公司介绍.md、业绩表.xlsx、人员社保表.xlsx 的文件夹）
    - days: 时间窗口（天），默认 None。优先读取「清单路径/时间范围选择.json」中的天数；
      未配置且 days 为 None 时不限制公告时间（时间范围是每次初筛的可选项，不强制限定）。
    - incremental（融合新增，默认 False）: 增量模式。仅对新增记录（初筛状态≠已初筛）初筛，
      初筛结果 Excel 追加 初筛状态/研判状态 两列，收尾回写状态；
      不传时行为与融合前完全一致（不读状态、不回写、不追加列）。
    - reset_status（融合新增，默认 False）: 复评。先重置全部记录的初筛/研判状态
      （新增待初筛/空），再按增量路径处理（重置后无状态 → 全量视为新增）。

    流程：解析 Excel（多清单合并+表头探测）→ 增量过滤（可选）→ 企业资料摘要 → 初筛
    （时间[可选]+类型+项目类型）→ 匹配标书（映射表→去词缀模糊→候选Top3）→ 解析文本
    → 提取评分章节 → 输出初筛结果 Excel（清单路径）→ 待研判清单（临时目录）
    """
    timestamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    result = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "timestamp": timestamp,
        "paths": {"list_dir": list_dir, "bid_dir": bid_dir, "profile_dir": profile_dir},
        "summary": {},
        "company_summary": {},
        "items": [],
        "unmatched_bid_files": [],
        "errors": [],
    }

    # 步骤1：读取 Excel（支持多清单合并、表头行自动探测）
    try:
        loaded = excel_tool.load_excels_from_dir(list_dir)
        df = loaded["df"]
        sources = loaded["sources"]
        header_rows = loaded["header_rows"]
    except excel_tool.ExcelToolError as e:
        result["errors"].append(str(e))
        return result

    # 步骤1.5：复评重置（--reset-status）：全部状态重置为「新增待初筛 / 空」，再走增量路径（全量=新增）
    # 重置失败（如主文件被 Excel 打开）必须中止运行：否则后续会用旧状态把全部记录当"已初筛"
    # 静默跳过，产出误导性的空初筛结果。
    if reset_status:
        try:
            status_tool.reset_status(list_dir)
            result["errors"].append(
                "已重置全部记录的初筛/研判状态（新增待初筛/空），本次按全部新增处理。"
            )
        except status_tool.StatusToolError as e:
            result["errors"].append(
                "复评重置失败，本次运行已中止（未做任何初筛/状态回写）：%s" % e
            )
            return result
        incremental = True

    # 步骤1.6：增量过滤（融合新增，方案 5）：load_excels_from_dir 之后、filter_screening 之前；
    # 按地址链接URL 读取状态，剔除 初筛状态=已初筛 的记录；缺 URL 列时降级全量处理。
    incremental_effective = False
    if incremental:
        state, state_warnings = status_tool.read_state(list_dir)
        for w in state_warnings:
            result["errors"].append(w)
        if "地址链接URL" not in df.columns:
            result["errors"].append(
                "清单缺少地址链接URL列，无法按 URL 增量，本次全量处理。"
            )
        else:
            incremental_effective = True
            urls = df["地址链接URL"].map(
                lambda v: str(v).strip() if v is not None else ""
            )
            st_of = state.get
            screened = urls.map(
                lambda u: (st_of(u) or {}).get("初筛状态") == "已初筛" if u else False
            )
            judged = urls.map(
                lambda u: (st_of(u) or {}).get("研判状态") == "已研判" if u else False
            )
            total_loaded = int(len(df))
            new_count = int((~screened).sum())
            skipped_screened_count = int(screened.sum())
            skipped_judged_count = int((screened & judged).sum())
            result["summary"]["incremental"] = True
            result["summary"]["total_loaded"] = total_loaded
            result["summary"]["new_count"] = new_count
            result["summary"]["skipped_screened_count"] = skipped_screened_count
            result["summary"]["skipped_judged_count"] = skipped_judged_count
            df = df[~screened]

    # 步骤2：企业资料摘要（研判唯一入口）+ 占位词/完整性校验（中间产物写临时目录）
    try:
        company_summary = company_profile.build_summary(
            md_path=None, attachments_dir=profile_dir
        )
        company_profile.save_summary(company_summary, TEMP_COMPANY_JSON)
    except company_profile.CompanyProfileError as e:
        company_summary = {}
        result["errors"].append(str(e))
    result["company_summary"] = company_summary
    result["company_summary_path"] = TEMP_COMPANY_JSON
    for err in company_summary.get("errors", []):
        result["errors"].append(err)
    if company_summary.get("placeholder_hits"):
        result["errors"].append(
            "⚠️ 公司介绍仍含示例内容（\"某/示例\"字样），研判结果将基于示例数据，请先更新："
            + "；".join(company_summary["placeholder_hits"][:5])
        )
    missing = company_summary.get("completeness", {}).get("未填", [])
    if missing:
        result["errors"].append(
            "企业资料未填齐（缺：%s），研判时相关维度将标注'企业资料不足'。"
            % "、".join(missing)
        )

    # 步骤2.5：读取用户本次确认的项目类型（清单路径/项目类型选择.json，缺失时自动跳过类型过滤）
    project_types_file = os.path.join(list_dir, excel_tool.PROJECT_TYPES_SELECTION_FILE)
    project_types = excel_tool.load_project_types(project_types_file, errors=result["errors"])
    result["selected_project_types"] = [
        (t["名称"] if t.get("子类") is None else "%s（%s）" % (t["名称"], "、".join(t["子类"])))
        for t in project_types
    ]
    if not project_types:
        result["errors"].append(
            "警告：尚未选择项目类型（清单路径/%s 中无启用项），本次初筛跳过类型过滤。"
            "建议先在对话中确认项目类型后再运行。" % excel_tool.PROJECT_TYPES_SELECTION_FILE
        )

    # 步骤2.6：读取用户本次确认的时间范围（清单路径/时间范围选择.json，未配置=不限，不强制限定）
    time_file = os.path.join(list_dir, excel_tool.TIME_RANGE_FILE)
    cfg_days = excel_tool.load_time_range(time_file)
    if cfg_days is not None:
        days = cfg_days
    result["time_range_days"] = days  # None = 不限公告时间
    if days is None:
        result["errors"].append(
            "提示：本次未限定公告时间范围（不限时间）。如需限定，请在对话中确认天数并写入 清单路径/%s。" % excel_tool.TIME_RANGE_FILE
        )
    else:
        result["errors"].append("本次时间范围：公告 %d 天以内。" % days)

    # 步骤3：初筛（相关性为软标记，不参与过滤）
    company_text = company_summary.get("_raw_text", "")
    df = excel_tool.filter_screening(
        df, company_text=company_text, days=days, project_types=project_types
    )

    # 步骤3.5：初筛结果输出。
    # 多 Sheet 主文件工作流（来源名含"::"，主文件为政府采购公告.xlsx）：初筛结果直接回写主文件
    # 对应 Sheet（初筛结果/初筛原因/相关性(命中词) 三列），不新建 初筛结果.xlsx 表格；
    # 否则输出 {时间}_初筛结果.xlsx（P3：放清单路径，各清单一个 sheet）。
    screening_excel_path = ""
    try:
        back_written = False
        if len(df):
            back_written = excel_tool.write_screening_back_to_main(df, sources)
        if back_written:
            screening_excel_path = next(iter(sources.values()), "")
        else:
            extra_cols = None
            if incremental_effective and len(df):
                df["初筛状态"] = "已初筛"
                df["研判状态"] = df["初筛结果"].map(
                    lambda r: "" if r == "通过" else "已跳过"
                )
                extra_cols = ["初筛状态", "研判状态"]
            screening_excel_path = excel_tool.write_screening_excel(
                df, list_dir, sources, timestamp, extra_columns=extra_cols
            )
    except excel_tool.ExcelToolError as e:
        result["errors"].append(str(e))
    result["screening_excel"] = screening_excel_path
    if screening_excel_path:
        result["errors"].append("初筛结果已输出：%s" % screening_excel_path)

    # 步骤3.6：增量收尾——状态回写（方案 4.3）：本次批次全部记录 初筛状态=已初筛；
    # 初筛通过者 研判状态=空，未通过者 研判状态=已跳过；按 URL 定位，禁止按行号。
    if incremental_effective and len(df):
        updates = []
        for _, row in df.iterrows():
            url = excel_tool._stringify(row.get("地址链接URL"))
            if not url:
                continue
            passed = excel_tool._stringify(row.get("初筛结果")) == "通过"
            updates.append({
                "url": url,
                "初筛状态": "已初筛",
                "研判状态": "" if passed else "已跳过",
            })
        if updates:
            try:
                status_tool.write_state(list_dir, updates)
            except status_tool.StatusToolError as e:
                result["errors"].append(str(e))

    # 步骤4：匹配标书（建议 10：映射表优先 → 建议 7：去词缀模糊匹配 + 候选 Top3）
    bid_files = docx_parser.list_bid_files(bid_dir)
    if not bid_files:
        result["errors"].append(
            "未找到标书文件，请将招标文件（.docx/.doc 或 .pdf）放入标书路径后再试（%s）。" % bid_dir
        )
    # 读取手动映射（标书路径/标书映射.json，建议 10，P3 随标书走）
    mapping_file = os.path.join(bid_dir, bid_mapping.MAPPING_FILE_NAME)
    bid_mapping_dict = {}
    try:
        bid_mapping_dict = bid_mapping.load_bid_mapping(mapping_file)
        if bid_mapping_dict:
            result["bid_mapping"] = bid_mapping_dict
            result["bid_mapping_file"] = mapping_file
    except bid_mapping.BidMappingError as e:
        result["errors"].append(str(e))
    matched_set = set()

    # 解析通过的候选
    passed = excel_tool.filter_passed(df)
    passed_projects = [
        {
            "project_name": excel_tool._stringify(row.get("采购项目名称")),
            "source": excel_tool._stringify(row.get("来源文件")),
            "row_index": int(row["原始行号"]),
        }
        for _, row in passed.iterrows()
    ]
    items = []
    for _, row in passed.iterrows():
        project_name = excel_tool._stringify(row.get("采购项目名称"))
        # 1) 映射表优先（建议 10）
        bid_file = find_bid_by_mapping(project_name, bid_mapping_dict, bid_files)
        match_way = "映射表"
        match_candidates = []
        # 2) 自动模糊匹配（建议 7：去词缀 + 候选 Top3）
        if bid_file is None:
            match_result = match_bid_file(project_name, bid_files)
            bid_file = match_result["bid_file"]
            match_way = match_result["way"]
            match_candidates = match_result["candidates"]
        if bid_file:
            matched_set.add(bid_file)
        item = {
            "source": excel_tool._stringify(row.get("来源文件")),
            "row_index": int(row["原始行号"]),
            "project_type": excel_tool._stringify(row.get("项目类型")),
            "project_name": project_name,
            "announce_time": excel_tool._stringify(row.get("公告时间")),
            "announce_type": excel_tool._stringify(row.get("公告类型")),
            "district": excel_tool._stringify(row.get("所在区县")),
            "buyer": excel_tool._stringify(row.get("采购人")),
            "budget": excel_tool._stringify(row.get("预算金额(元)")),
            "url": excel_tool._stringify(row.get("地址链接URL")),
            "relevance": excel_tool._stringify(row.get("相关性初判")),
            "relevance_why": excel_tool._stringify(row.get("相关性依据")),
            "type_pending": bool(row.get("类型待确认")),
            "type_pending_why": excel_tool._stringify(row.get("初筛原因")),
            "bid_file": bid_file,
            "bid_match_way": match_way,
            "bid_match_candidates": match_candidates,
            "bid_text": "",
            "scoring_section": "",
            "scoring_missing": False,
            "matched": bool(bid_file),
        }
        if bid_file:
            try:
                item["bid_text"] = docx_parser.parse_file(
                    bid_file, save_preview=True, preview_dir=TEMP_BID_TEXT_DIR
                )
                item["scoring_section"] = docx_parser.extract_scoring_section(item["bid_text"])
                if not item["scoring_section"].strip():
                    item["scoring_missing"] = True
                    result["errors"].append(
                        "「%s」未提取到评分章节（未识别到评分标题或章节结构特殊），"
                        "研判时请从标书全文自行查找评分表逐条核对。" % os.path.basename(bid_file)
                    )
            except docx_parser.DocxParseError as e:
                item["bid_text"] = ""
                item["error"] = str(e)
                result["errors"].append(str(e))
        items.append(item)

    # 未匹配到任何项目的标书文件（映射表中的文件不列入未匹配清单，建议 10）
    unmatched_bid = []
    for f in bid_files:
        base = os.path.basename(f)
        if f in matched_set:
            continue
        if base in bid_mapping_dict or os.path.splitext(base)[0] in bid_mapping_dict:
            matched_set.add(f)
            continue
        unmatched_bid.append(f)

    # 未匹配标书自动预识别（建议 8：标题区前 500 字候选项目短语）
    bid_candidates = []
    for f in unmatched_bid:
        bid_candidates.append(_pre_recognize_bid(f, passed_projects))

    # 初筛统计
    screening = df["初筛结果"].value_counts().to_dict()
    pending = [it for it in items if it["relevance"] != "相关"]
    type_pending = [it for it in items if it["type_pending"]]
    result["summary"] = {
        "excel_files": list(sources.keys()),
        "total_rows": int(len(df)),
        "selected_project_types": [t["名称"] for t in project_types],
        "time_range_days": days,
        "screening": {str(k): int(v) for k, v in screening.items()},
        "passed_count": int(len(passed)),
        "relevance_pending_count": int(len(pending)),
        "type_pending_count": int(len(type_pending)),
        "matched_count": int(sum(1 for it in items if it["matched"])),
        "bid_file_count": int(len(bid_files)),
        "unmatched_bid_count": int(len(unmatched_bid)),
        "bid_pre_recognized_count": int(sum(1 for c in bid_candidates if not c["无候选"])),
        "bid_mapping_count": int(len(bid_mapping_dict)),
        "scoring_missing_count": int(sum(1 for it in items if it["scoring_missing"])),
    }
    # 增量模式字段在汇总区追加（步骤1.6 已写入，此处保留不被覆盖）
    if incremental_effective:
        result["summary"]["incremental"] = True
        result["summary"]["total_loaded"] = total_loaded
        result["summary"]["new_count"] = new_count
        result["summary"]["skipped_screened_count"] = skipped_screened_count
        result["summary"]["skipped_judged_count"] = skipped_judged_count
    result["items"] = items
    result["unmatched_bid_files"] = unmatched_bid
    result["bid_candidates"] = bid_candidates
    for err in loaded.get("errors", []):
        result["errors"].append(err)

    # 留档待研判清单到系统临时目录（P3：%TEMP%/投标评估助手/待研判清单.json，不污染业务目录）
    _save_preview(result)
    return result


def _save_preview(result: dict) -> None:
    """将待研判清单 JSON 留档到系统临时目录。"""
    os.makedirs(TEMP_BASE_DIR, exist_ok=True)
    out_path = TEMP_PREVIEW_JSON
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def main() -> int:
    """命令行入口（P3：三个路径由参数传入，无固定目录依赖）。"""
    # 兼容 GBK 控制台：强制 UTF-8 输出，避免中文/符号打印报错
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    parser = argparse.ArgumentParser(
        description="投标机会评估助手流水线：初筛 + 标书匹配 + 评分章节提取（P3 路径参数化）"
    )
    parser.add_argument("--list-dir", required=True, help="清单路径：放置投标机会清单 .xlsx 的文件夹")
    parser.add_argument("--bid-dir", required=True, help="标书路径：放置标书 .docx/.doc/.pdf 的文件夹")
    parser.add_argument("--profile-dir", required=True, help="公司资料路径：放置公司介绍.md、业绩表.xlsx、人员社保表.xlsx 的文件夹")
    parser.add_argument("--days", type=int, default=None, help="初筛时间窗口（天）；优先读取 清单路径/时间范围选择.json；不传且未配置=不限时间")
    parser.add_argument("--incremental", action="store_true", help="增量模式：仅对新增记录（初筛状态≠已初筛）初筛，初筛结果 Excel 追加状态两列并回写状态（不传=融合前全量行为）")
    parser.add_argument("--reset-status", action="store_true", help="复评：重置全部记录的初筛/研判状态为「新增待初筛/空」，本次按全部新增处理")
    parser.add_argument("--json", action="store_true", help="打印待研判清单 JSON（供 SKILL.md 读取）")
    args = parser.parse_args()

    try:
        result = run_pipeline(
            args.list_dir,
            args.bid_dir,
            args.profile_dir,
            days=args.days,
            incremental=args.incremental,
            reset_status=args.reset_status,
        )
    except Exception as e:
        print("【错误】%s" % e)
        return 1

    if result["errors"]:
        for err in result["errors"]:
            print("【提示】%s" % err)

    s = result["summary"]
    print("=" * 50)
    print("投标机会评估助手 · 流水线执行完成")
    print("=" * 50)
    print("清单路径：%s" % args.list_dir)
    print("标书路径：%s" % args.bid_dir)
    print("公司资料路径：%s" % args.profile_dir)
    if s.get("incremental"):
        print("增量模式：本次共读取 %s 条记录：新增 %s 条（已进入初筛）、跳过已初筛 %s 条（其中已研判 %s 条）" % (
            s.get("total_loaded", 0), s.get("new_count", 0),
            s.get("skipped_screened_count", 0), s.get("skipped_judged_count", 0)))
    print("清单文件：%s" % "、".join(s.get("excel_files", [])))
    print("总行数：%s" % s.get("total_rows", 0))
    selected = s.get("selected_project_types", [])
    print("已选项目类型：%s" % ("、".join(selected) if selected else "未配置（跳过类型过滤）"))
    tdays = s.get("time_range_days")
    print("时间范围：%s" % ("不限（未强制限定）" if tdays is None else "公告 %d 天以内" % tdays))
    print("初筛统计：%s" % json.dumps(s.get("screening", {}), ensure_ascii=False))
    print("初筛通过：%s 个（相关性待确认 %s 个，不剔除）" % (s.get("passed_count", 0), s.get("relevance_pending_count", 0)))
    if result.get("screening_excel"):
        print("初筛结果已输出：%s" % result["screening_excel"])
    print("已匹配标书：%s 个" % s.get("matched_count", 0))
    print("标书文件总数：%s 个" % s.get("bid_file_count", 0))
    print("未匹配标书：%s 个（已自动预识别候选项目 %s 个）" % (
        s.get("unmatched_bid_count", 0), s.get("bid_pre_recognized_count", 0)))
    if s.get("bid_mapping_count"):
        print("映射表生效：%s 条（标书路径/%s）" % (s.get("bid_mapping_count"), bid_mapping.MAPPING_FILE_NAME))
    print("未提取到评分章节：%s 个" % s.get("scoring_missing_count", 0))
    print("中间产物（临时目录）：%s" % TEMP_BASE_DIR)

    if result.get("bid_mapping"):
        print("\n【标书手动映射】（%s，匹配顺序：映射表优先）：" % result.get("bid_mapping_file", "标书路径/标书映射.json"))
        for k, v in result["bid_mapping"].items():
            print("  - %s → %s" % (k, v))

    if result["unmatched_bid_files"]:
        print("\n【未能匹配到项目的标书】(请在对话中确认候选或手动指定)：")
        for f in result["unmatched_bid_files"]:
            print("  - %s" % os.path.basename(f))
        print("\n【未匹配标书候选项目预识别】(标题区前 500 字)：")
        for c in result.get("bid_candidates", []):
            line = "  - %s → " % c["标书文件名"]
            if c["错误"]:
                print(line + "解析失败：%s" % c["错误"])
            elif c["无候选"]:
                print(line + "无候选")
            else:
                print(line + "候选：" + "、".join(c["候选短语"]))
                for h in c["匹配到清单项目"]:
                    print("      ↳ 与清单项目「%s」（%s 行 %s）相近，请用户确认" % (
                        h["项目名称"], h["来源"], h["行号"]))

    if args.json:
        print("\n【待研判清单 JSON】")
        print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    sys.exit(main())