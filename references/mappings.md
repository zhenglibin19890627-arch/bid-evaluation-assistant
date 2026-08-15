# Mappings Reference

## Category Codes (公告类型 → categoryCode)

Used in the POST body of the list API (`/portal/category`).

| 公告类型 (Standard) | categoryCode | 备注 |
|---|---|---|
| 采购意向 | `110-175885` | 采购意向公开 |
| 意见征询 | `110-246839` | 含单一来源公示、需求公示 |
| 采购项目公告 | `110-978863` | 招标公告、资格预审、非招标 |
| 更正公告 | `110-943756` | 修正/澄清公告 |
| 采购结果公告 | `110-900461` | 中标/成交/废标 |
| 采购合同公告 | `110-567245` | 合同/合同变更 |
| 履约验收公告 | `110-198517` | 项目验收 |
| 电子卖场公告 | `110-341071` | 电子卖场 |
| 框架协议公告 | `110-978920` | 框架协议采购 |

## pathName → Standard Type Mapping

The API returns a `pathName` field; map it to the standard type:

| pathName (API) | Standard Type |
|---|---|
| 采购意向公开 | 采购意向 |
| 采购文件需求公示 | 意见征询 |
| 单一来源采购公示 | 意见征询 |
| 资格预审公告 | 采购项目公告 |
| 招标公告 | 采购项目公告 |
| 非招标公告 | 采购项目公告 |
| 更正公告 | 更正公告 |
| 中标（成交）结果公告 | 采购结果公告 |
| 废标公告 | 采购结果公告 (→ overridden to 废标) |
| 采购合同公告 | 采购合同公告 |
| 合同变更公告 | 采购合同公告 |
| 履约验收公告 | 履约验收公告 |
| 其他电子卖场公告 | 电子卖场公告 |
| 框架协议公告 | 框架协议公告 |

## District Codes (示例：丽水市)

Format: `{districtCode: "Display Name"}`

| Code | Name |
|---|---|
| 331199 | 丽水市本级 |
| 331102 | 莲都区 |
| 331103 | 丽水开发区 |
| 331121 | 青田县 |
| 331122 | 缙云县 |
| 331123 | 遂昌县 |
| 331124 | 松阳县 |
| 331125 | 云和县 |
| 331126 | 庆元县 |
| 331127 | 景宁畲族自治县 |
| 331181 | 龙泉市 |

**To add/change regions:**
- Region configuration is **not** in `fetch_zfcg.py` — it lives in `references/districts.json`（`cities` → `{prefix, districts}`）。
- Add a city entry with its district codes and run `fetch_zfcg.py --city <地市名>`; no code change needed.
- `scripts/fetch_zfcg.py` 通过 `load_districts()` 读取该文件并动态生成目标区县集合（`TARGET_DISTRICTS`/`TARGET_DISTRICT_CODES`）。

## Supplier Suffixes (供应商合法后缀)

Organization name must end with one of these to be recognized:

```
公司|中心|事务所|集团|医院|院校|大学|研究院|书院|学院|
合作社|工程队|协会|勘查院|农资店|服务部|家庭农场|商行|经营部|养殖场
```

When adding a new suffix, update ALL regex patterns that reference the suffix group.

## Purchaser Suffixes (采购人合法后缀)

```
局|委|办|中心|院|处|部|所|站|队|园|校|厅|室|会|学
```
