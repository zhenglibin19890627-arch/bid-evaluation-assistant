# -*- coding: utf-8 -*-
"""
company_profile.py —— 企业资料摘要生成 + 完整性校验 + 占位词检测

功能：
1. 读取「公司资料路径/公司介绍.md」（路径由用户每次提供，P3 起不依赖固定 config 目录），
   抽取 9 节填写状态，检测示例/占位字样（"某""示例"中断；
   "待确认"与空值放行，仅提示缺失项——需求变更 2026-08-05）；
2. 读取可选附件「公司资料路径/业绩表.xlsx」「公司资料路径/人员社保表.xlsx」，
   **原样保留全部列与数据**（附件表头由用户实际提供、每次可能不同：
   不做列名归整、不删减字段，除非该字段对研判无影响）；
3. 生成企业资料摘要（P3 起落盘到系统临时目录，AI 研判唯一的企业资料入口，
   附件数据原样透传，字段含义由研判时按实际表头理解）。

用法：
    from scripts import company_profile
    summary = company_profile.build_summary(md_path, attachments_dir)   # dict，可打印/落盘
    company_profile.save_summary(summary, out_path)                     # 落盘临时目录
"""

from __future__ import annotations

import json
import os
import re
from datetime import datetime

import pandas as pd

# 9 节企业资料（顺序与模板一致）
SECTIONS = [
    "公司基本信息",
    "公司业务",
    "资质证书",
    "历史业绩",
    "主销产品与品牌",
    "厂家合作关系",
    "人员团队",
    "财务能力",
    "其他说明",
]

# 中断字样（需求变更 2026-08-05）：疑似假数据，命中即要求用户更新，防止基于示例数据研判
# "待确认"与空值已放行（不中断，仅提示缺失项，研判时相关维度标注"企业资料不足"）。
INTERRUPT_WORDS = ("某", "示例")

# 公司资料路径下的约定文件名（P3 起路径由用户每次提供，放在「公司资料路径」下）
COMPANY_MD_NAME = "公司介绍.md"
PERFORMANCE_FILE_NAME = "业绩表.xlsx"
PERSONNEL_FILE_NAME = "人员社保表.xlsx"


class CompanyProfileError(Exception):
    """企业资料处理过程中的中文可读异常。"""


def _read_md(path: str) -> str:
    if not os.path.exists(path):
        raise CompanyProfileError(
            "未找到企业资料文件：%s\n请确认公司资料路径下存在「%s」。"
            % (path, COMPANY_MD_NAME)
        )
    with open(path, encoding="utf-8") as f:
        return f.read()


def parse_md_sections(text: str) -> list:
    """
    解析公司介绍.md：按「## 」标题切分，返回
    [{"节名": str, "已填": bool, "占位": bool, "行": [{"标签": str, "值": str}]}]

    「占位」仅统计中断字样（某/示例，疑似假数据）；
    "待确认"与空值放行（不中断，由完整性校验提示缺失项）。
    """
    sections = []
    current = None
    for raw in text.splitlines():
        line = raw.rstrip()
        if line.startswith("## "):
            name = line[3:].strip()
            name = re.sub(r"^[一二三四五六七八九十\d]+、", "", name).strip()
            if current:
                sections.append(current)
            current = {"节名": name, "行": []}
            continue
        if current is None or not line.strip():
            continue
        if line.strip().startswith(">"):
            continue
        m = re.match(r"^[-*\s]*([^：:]+)[：:]\s*(.*)$", line)
        if m:
            current["行"].append({"标签": m.group(1).strip(), "值": m.group(2).strip()})
        else:
            current["行"].append({"标签": line.strip(), "值": ""})
    if current:
        sections.append(current)

    for s in sections:
        filled = [r for r in s["行"] if r["值"]]
        s["已填"] = bool(filled)
        s["占位"] = any(
            any(w in r["值"] for w in INTERRUPT_WORDS) for r in s["行"]
        )
    return sections


def parse_attachment(path: str) -> dict:
    """
    读取附件 xlsx，**原样保留全部列与数据**（不做列名归整、不删减字段）。

    返回 {"存在": bool, "列": [列名...], "列数": int, "行数": int, "行": [{原列名: 值}...], "错误": [str]}
    - 表头由用户实际提供，每次可能不同；字段是否对研判有影响由研判时判断；
    - 仅提供结构与原始数据，不做任何字段筛选/转换/统计。
    """
    out = {"存在": os.path.exists(path), "列": [], "列数": 0, "行数": 0, "行": [], "错误": []}
    if not out["存在"]:
        return out
    try:
        df = pd.read_excel(path, dtype=str)
    except Exception as e:
        out["错误"].append("读取失败：%s" % e)
        return out
    if df.empty:
        out["错误"].append("文件为空")
        return out

    columns = [str(c).strip() for c in df.columns]
    out["列"] = columns
    out["列数"] = len(columns)

    for _, row in df.iterrows():
        item = {}
        empty = True
        for col in columns:
            val = row.get(col)
            val = "" if val is None or str(val).lower() in ("nan", "none", "null") else str(val).strip()
            item[col] = val
            if val:
                empty = False
        if not empty:
            out["行"].append(item)
    out["行数"] = len(out["行"])
    return out


def build_summary(md_path: str = None, attachments_dir: str = None) -> dict:
    """
    生成企业资料摘要（研判唯一入口）。

    - md_path: 公司介绍.md 的完整路径；为 None 时在 attachments_dir（公司资料路径）下
      自动定位「公司介绍.md」，找不到则找该目录下其他 .md/.txt；仍找不到返回 errors
      提示（缺失文件不报错，按"待确认放行/某示例中断"口径处理，由 AI 提示用户补充）。
    - attachments_dir: 公司资料路径（存放业绩表.xlsx、人员社保表.xlsx 的文件夹）。
    """
    if md_path is None:
        md_path = _locate_md(attachments_dir)
    result = {
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "company_file": os.path.basename(md_path) if md_path else "",
        "sections": [],
        "completeness": {"已填": 0, "共": len(SECTIONS), "未填": []},
        "placeholder_hits": [],
        "performance": None,
        "personnel": None,
        "errors": [],
    }

    if not md_path or not os.path.exists(md_path):
        result["errors"].append(
            "未找到公司介绍文档（公司资料路径：%s）。\n"
            "请确认该路径下存在「%s」，或用记事本新建后重试；也可以直接在对话中口述，AI 代写。"
            % (attachments_dir or "未提供", COMPANY_MD_NAME)
        )
        return result

    try:
        text = _read_md(md_path)
    except CompanyProfileError as e:
        result["errors"].append(str(e))
        return result
    result["_raw_text"] = text  # 供初筛相关性初判使用（不入摘要正文展示）

    sections = parse_md_sections(text)
    # 按模板顺序整理，模板外的节归入"其他"
    by_name = {s["节名"]: s for s in sections}
    ordered = []
    for name in SECTIONS:
        s = by_name.pop(name, None)
        if s is None:
            s = {"节名": name, "已填": False, "占位": False, "行": []}
        ordered.append(s)
    for name, s in by_name.items():
        ordered.append(s)

    filled = [s["节名"] for s in ordered if s["已填"]]
    missing = [s["节名"] for s in ordered if not s["已填"]]
    for s in ordered:
        if s["占位"]:
            for r in s["行"]:
                if r["值"] and any(w in r["值"] for w in INTERRUPT_WORDS):
                    result["placeholder_hits"].append(
                        "%s-%s：含示例字样（%s）" % (s["节名"], r["标签"], r["值"][:30])
                    )

    result["sections"] = ordered
    result["completeness"] = {
        "已填": len(filled), "共": len(SECTIONS), "未填": missing, "已填列表": filled,
    }

    # 附件解析（原样保留全部列，不做字段归整/统计）；缺失不报错，按可选处理
    perf_path = os.path.join(attachments_dir or "", PERFORMANCE_FILE_NAME)
    pers_path = os.path.join(attachments_dir or "", PERSONNEL_FILE_NAME)
    result["performance"] = parse_attachment(perf_path)
    result["personnel"] = parse_attachment(pers_path)
    for name, attach in (("业绩表", result["performance"]), ("人员社保表", result["personnel"])):
        for err in attach.get("错误", []):
            result["errors"].append("%s：%s" % (name, err))

    return result


def _locate_md(attachments_dir: str) -> str:
    """
    在公司资料路径下定位公司介绍文档：优先「公司介绍.md」，其次任意 .md/.txt。
    目录不存在或为空时返回 None（由 build_summary 给出中文提示）。
    """
    if not attachments_dir or not os.path.isdir(attachments_dir):
        return None
    primary = os.path.join(attachments_dir, COMPANY_MD_NAME)
    if os.path.exists(primary):
        return primary
    for name in sorted(os.listdir(attachments_dir)):
        if name.lower().endswith((".md", ".txt")) and not name.startswith((".", "~$")):
            return os.path.join(attachments_dir, name)
    return None


def check_ready(summary: dict = None) -> (bool, list):
    """研判前置校验：返回 (是否就绪, 提示列表)。占位词命中或 9 项未填齐时未就绪。"""
    summary = summary or build_summary()
    msgs = []
    if summary["placeholder_hits"]:
        msgs.append(
            "检测到公司介绍仍含示例内容（%s），请先更新为贵公司真实信息后再研判"
            "（'待确认'或留空的项不会中断，可在研判时补充）。"
            % "；".join(summary["placeholder_hits"][:5])
        )
    missing = summary["completeness"]["未填"]
    if missing:
        msgs.append("企业资料未填齐（缺：%s），研判时相关维度将标注'企业资料不足'。" % "、".join(missing))
    if summary["performance"] and not summary["performance"]["存在"]:
        msgs.append("可选：业绩明细可放入公司资料路径/业绩表.xlsx（表头按你们实际格式，原样保留全部列）供逐条核对。")
    if summary["personnel"] and not summary["personnel"]["存在"]:
        msgs.append("可选：人员社保明细可放入公司资料路径/人员社保表.xlsx（表头按你们实际格式，原样保留全部列）供逐条核对。")
    ready = not summary["placeholder_hits"]
    return ready, msgs


def save_summary(summary: dict = None, out_path: str = None) -> str:
    """落盘企业资料摘要.json（P3 起写入系统临时目录），返回路径。"""
    if not out_path:
        raise CompanyProfileError(
            "无法保存企业资料摘要：未提供输出路径。\n请传入临时目录路径，如 os.path.join(temp_dir, '企业资料摘要.json')。"
        )
    summary = summary or build_summary()
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)
    return out_path


if __name__ == "__main__":
    import sys
    md = sys.argv[1] if len(sys.argv) > 1 else None
    attach = sys.argv[2] if len(sys.argv) > 2 else None
    s = build_summary(md, attach)
    print(json.dumps(s, ensure_ascii=False, indent=2))
