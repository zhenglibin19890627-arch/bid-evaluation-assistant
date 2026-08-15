# -*- coding: utf-8 -*-
"""
bid_mapping.py —— 标书手动映射落盘（建议 10，P3 起随标书走）

标书路径/标书映射.json 格式：{"标书文件名": "采购项目名称"}

- 流水线匹配顺序：映射表 → 自动模糊匹配；
- 用户在对话中手动指定的映射调用 add_bid_mapping() 自动写入，下次直接复用；
- 映射表中的标书不再列入未匹配清单；
- 文件为普通 JSON，可用记事本编辑。

用法（必须传入标书路径下的映射文件完整路径）：
    from scripts import bid_mapping
    mapping = bid_mapping.load_bid_mapping(path)            # {"标书文件名": "采购项目名称"}
    bid_mapping.add_bid_mapping("202609.docx", "项目名称", path)
    其中 path = os.path.join(标书路径, "标书映射.json")
"""

from __future__ import annotations

import json
import os

# P3 起映射文件写到标书路径下（随标书走），不再有固定的 config/标书映射.json
MAPPING_FILE_NAME = "标书映射.json"


class BidMappingError(Exception):
    """标书映射处理过程中的中文可读异常。"""


def _require_path(path) -> str:
    """映射文件路径必须显式传入（标书路径/标书映射.json）。"""
    if not path:
        raise BidMappingError(
            "未提供标书映射文件路径。\n请传入 os.path.join(标书路径, '%s') 后再调用。" % MAPPING_FILE_NAME
        )
    return str(path)


def load_bid_mapping(path: str = None) -> dict:
    """
    读取「标书路径/标书映射.json」，返回 {"标书文件名": "采购项目名称"}。

    文件不存在时返回 {}；文件损坏或格式不符时抛出中文提示。
    """
    cfg_path = _require_path(path)
    if not os.path.exists(cfg_path):
        return {}
    try:
        with open(cfg_path, encoding="utf-8-sig") as f:
            data = json.load(f)
    except Exception as e:
        raise BidMappingError(
            "读取 %s 失败：%s\n若曾用记事本编辑，请确认是合法 JSON 格式（如 {\"标书文件名\": \"采购项目名称\"}）。"
            % (cfg_path, e)
        )
    if not isinstance(data, dict):
        raise BidMappingError(
            "%s 格式应为 {\"标书文件名\": \"采购项目名称\"}，请用记事本修正后重试。" % cfg_path
        )
    return {str(k): str(v) for k, v in data.items() if str(k).strip()}


def save_bid_mapping(mapping: dict, path: str = None) -> str:
    """将映射字典落盘到「标书路径/标书映射.json」，返回文件路径。"""
    cfg_path = _require_path(path)
    os.makedirs(os.path.dirname(cfg_path) or ".", exist_ok=True)
    try:
        with open(cfg_path, "w", encoding="utf-8") as f:
            json.dump(mapping, f, ensure_ascii=False, indent=2)
    except Exception as e:
        raise BidMappingError("保存 %s 失败：%s" % (cfg_path, e))
    return cfg_path


def add_bid_mapping(bid_filename: str, project_name: str, path: str = None) -> dict:
    """
    新增/更新一条映射：{"标书文件名": "采购项目名称"}，自动落盘到标书路径。

    返回更新后的完整映射字典。对话中用户手动指定的映射应调用本函数写入
    （path 传 os.path.join(标书路径, '标书映射.json')）。
    """
    if not bid_filename or not str(bid_filename).strip():
        raise BidMappingError("标书文件名为空，无法写入映射。")
    if not project_name or not str(project_name).strip():
        raise BidMappingError("采购项目名称为空，无法写入映射。")
    mapping = load_bid_mapping(path)
    mapping[str(bid_filename).strip()] = str(project_name).strip()
    save_bid_mapping(mapping, path)
    return mapping
