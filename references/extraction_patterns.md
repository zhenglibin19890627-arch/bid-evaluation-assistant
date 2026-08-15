# Extraction Patterns Reference

Complete list of regex patterns used to extract structured data from government
procurement announcement HTML content.

## Pre-processing

```python
def strip_html(s):
    s = re.sub(r'<style[^>]*>.*?</style>', '', s, flags=re.DOTALL)
    s = re.sub(r'<[^>]+>', ' ', s)
    s = html.unescape(s)
    s = re.sub(r'&nbsp;', ' ', s)
    s = re.sub(r'\s+', ' ', s).strip()
    return s
```

## Budget Extraction (预算金额) — 16 Patterns

Ordered by specificity. `norm = re.sub(r'\s+', ' ', text)` then apply:

| # | Pattern | 万元-conversion |
|---|---------|----------------|
| 1 | `预算金额[（(]元[)）][：:\s]*([\d,.]+)` | No |
| 2 | `预算金额[：:\s]*([\d,.]+)\s*元` | No |
| 3 | `预算总价[（(]元[)）][：:\s]*([\d,.]+)` | No |
| 4 | `预算总价[：:\s]*([\d,.]+)\s*元` | No |
| 5 | `采购预算[（(]元[)）][：:\s]*([\d,.]+)` | No |
| 6 | `采购预算[：:\s]*([\d,.]+)\s*元` | No |
| 7 | `项目预算[：:\s]*([\d,.]+)\s*元` | No |
| 8 | `预算金额[：:\s]*人民币\s*([\d,.]+)` | No |
| 9 | `最高限价[（(]元[)）][：:\s]*([\d,.]+)` | No |
| 10 | `最高限价[：:\s]*([\d,.]+)\s*元` | No |
| 11 | `控制价[：:\s]*([\d,.]+)\s*元` | No |
| 12 | `单一来源采购限额[：:\s]*([\d,.]+)\s*万元` | ×10000 |
| 13 | `单一来源采购限额[：:\s]*([\d,.]+)\s*元` | No |
| 14 | `预算[（(]万元[)）][：:\s]*([\d,.]+)` | ×10000 |
| 15 | `预算[：:\s]*([\d,.]+)\s*元` | No |
| 16 | `预算[：:\s]*([\d,.]+)\s*万元` | ×10000 |
| 17 | `￥\s*([\d,]+(?:\.\d{1,2})?)` | No |

**万元 detection:** If the pattern is marked ×10000, OR `万元` appears within ±10 chars of the match:
```python
if is_wan or '万元' in norm[max(0, m.start()-10):m.end()+10]:
    v *= 10000
```

## Winning Amount Extraction (中标金额) — 25+ Patterns

**Multi-lot detection** (run first):
```python
multi_matches = re.findall(r'总价[：:,，\s]*([\d,.]+)\s*[（(]元[)）]', norm)
if len(multi_matches) >= 2:
    return sum(float(v.replace(',', '')) for v in multi_matches)
```

Single-match patterns (in order):

1. `中标[（(]成交[)）]金额[（(]元[)）][：:\s]*([\d,.]+)`
2. `中标金额[（(]元[)）][：:\s]*([\d,.]+)`
3. `成交金额[（(]元[)）][：:\s]*([\d,.]+)`
4. `中标[（(]成交[)）]金额[：:\s]*([\d,.]+)\s*元`
5. `合同金额[（(]元[)）][：:\s]*([\d,.]+)`
6. `合同金额[：:\s]*([\d,.]+)`
7. `总中标金额[：:\s]*([\d,.]+)`
8. `成交总金额[：:\s]*([\d,.]+)`
9. `中标总价[：:\s]*([\d,.]+)`
10. `总价报价[（(]元[)）][：:\s]*([\d,.]+)`
11. `总价报价[：:\s]*([\d,.]+)\s*[（(]元[)）]`
12. `总价报价[：:\s]*([\d,.]+)\s*元`
13. `总报价[（(]元[)）][：:\s]*([\d,.]+)`
14. `总报价[（(][^)）]*[)）][：:\s]*([\d,.]+)`
15. `总报价[：:\s]*([\d,.]+)\s*[（(]元[)）]`
16. `总报价[：:\s]*([\d,.]+)\s*元`
17. `总价[（(][^)）]*[)）][：:\s]*([\d,.]+)`
18. `总价[：:,，\s]*([\d,.]+)\s*[（(]元[)）]`
19. `总价[：:,，\s]*([\d,.]+)\s*元`
20. `投标报价[（(]元[)）][：:\s]*([\d,.]+)`
21. `投标报价[：:\s]*([\d,.]+)\s*[（(]元[)）]`
22. `投标报价[：:\s]*([\d,.]+)\s*元`
23. `投标价[：:\s]*([\d,.]+)\s*[（(]元[)）]`
24. `报价[：:]\s*([\d,.]+)\s*[（(]元[)）]`
25. `投标单价[（(][^)）]*[)）][：:\s]*([\d,.]+)\s*[（(]元[)）]`
26. `单价[：:\s]*([\d,.]+)\s*[（(]元[)）]`
27. `设备报价[：:\s]*([\d,.]+)\s*[（(]元[)）]`
28. `合价[（(][^)）]*[)）][：:\s]*([\d,.]+)`
29. `合价[：:\s]*([\d,.]+)\s*[（(]元[)）]`

## Supplier Extraction (供应商) — Multiple Strategies

**Strategy order:**

### 1. Multi-lot detection
If multiple `总价...元 company_name` patterns exist, return comma-joined unique names.

### 2. Joint bid (联合体) detection
```
牵头供应商：XXX公司 投标联合体：YYY公司、ZZZ公司
→ "牵头：XXX公司；联合体：YYY公司、ZZZ公司"
```

### 3. Standard patterns (6 patterns)
```
中标供应商名称：company_name
中标（成交）供应商：company_name
成交供应商：company_name
供应商（乙方）：company_name
入围供应商：company_name
供应商名称：company_name
```

### 4-13. Fallback patterns
After amount + (元) → company name extraction, covering:
总价/报价/投标报价/投标价/总价报价/总报价/中标金额/投标单价/单价/设备报价 + amount → company

### 14. Generic catch-all
`[\d,.]+\s*[（(]元[)）]\s*{company name}` — matches any amount+(元) followed by a valid company name.

### Name validation
- Must match: `[\u4e00-\u9fff（）()A-Za-z]+(合法后缀)`
- Length: 2 < len < 100
- Suffix must be one of: 公司|中心|事务所|集团|医院|院校|大学|研究院|书院|学院|合作社|工程队|协会|勘查院|农资店|服务部|家庭农场|商行|经营部|养殖场

## Agency Extraction (代理机构) — 4 Methods

| # | Method | Pattern |
|---|--------|---------|
| 1 | "XX 受 YY 委托" | `公司名\s+受\s+采购人\s*委托` |
| 2 | Structured contact section | `采购代理机构\s*信息\s*名\s*称：XX` |
| 3 | 代理机构 label | `(接收)?代理\s*机构：XX` |
| 4 | Generic contact name | `采购代理机构.*?名\s*称：XX` |

**Special cases:**
- 采购意向: always return empty (no agency)
- 采购合同公告/电子卖场公告: return empty if not found
- Others: fallback to `author` field if it contains 公司/事务所/代理

## Purchaser Extraction (采购人) — 7 Patterns

1. `采购人（甲方）：XXX局`
2. `采购人：XXX中心`
3. `采购人.*?名 称：XXX部`
4. `甲方：XXX委`
5. `受 XXX 的委托`
6. `采购人信息.*?名 称：任意名称`

## Void/废标 Detection

**Keyword matching** in the first 3000 chars of normalized text:

**Strong match** (any of these → 废标): `废标公告|招标失败公告|采购失败公告|流标公告`

**Weak match** (need confirmation): `废标|招标失败|采购失败|终止采购|取消采购|流标`
- Must also NOT have winning result indicators: `中标（成交）信息|中标（成交）结果|中标供应商名称`
- Must have failure-specific indicators: `废标理由|废标结果|流标原因`

## Cross-Referencing Rules

### Layer 1: Announcement Link Chain
- Follow `announcementLinkDtoList` to fetch linked announcements
- Extract missing budget/supplier/agency

### Layer 2: URL Extraction from HTML
- Regex: `articleId=([A-Za-z0-9+/=]+(?:%3D%3D|%3D))`
- Extract and fetch up to 10 unique linked IDs

### Layer 3: Amount Fallback
- 采购结果公告 without budget → use winning amount
- 采购合同公告 without budget → use contract amount
- **Always take the larger amount** when cross-referencing

### Critical Rules
1. **Void protection**: Never cross-reference winning info for 废标 announcements
2. **Budget priority**: Take `max(existing, new)` to avoid per-person costs
3. **Key consistency**: Always use `_winning_amount` internally
4. **Min thresholds**: Budget > 100, winning amount > 0

## Known Data Quality Patterns (30 Types)

| # | Type | Rule |
|---|------|------|
| 1 | 采购意向无代理 | 采购意向 should have no agency |
| 2 | 意见征询缺代理 | 单一来源公示 without agency is OK |
| 3 | 电子卖场 | Should have budget, no agency |
| 4 | 百分比定价 | 结果公告 with percentage pricing OK to have no amount |
| 5 | 供应商截断 | Name must have complete suffix |
| 6 | 金额异常小 | Winning amount should be reasonable |
| 7 | 废标检测 | Keyword-based void detection |
| 8+ | Various | Joint bid completeness, contract budget, correction budget, false void, special suffixes, bracket matching, percentage supplier, acceptance cross-ref, amount format, quotation format, single-source fallback, correction supplier, void protection, total quotation format, bid unit price format, total-comma-separated, farm supply suffix, unit price format, equipment quotation format, service dept suffix, family farm/shop/operation/breeding suffixes |
