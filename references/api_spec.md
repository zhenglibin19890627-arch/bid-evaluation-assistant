# API Specification — 浙江政府采购网

## Base URL

```
https://zfcg.czt.zj.gov.cn
```

## 1. List API (公告列表)

### Endpoint

```
POST /portal/category
```

### Request Headers

```
Content-Type: application/json
User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36
Accept: application/json, text/plain, */*
Origin: https://zfcg.czt.zj.gov.cn
Referer: https://zfcg.czt.zj.gov.cn/site/category?parentId=600007&childrenCode=ZcyAnnouncement
```

### Request Body (Current Format — post-2026-06-12)

```json
{
  "pageNo": 1,
  "pageSize": 100,
  "categoryCode": "110-978863",
  "isGov": true,
  "excludeDistrictPrefix": ["90", "006011", "H0", "001111"],
  "publishDateBegin": "2026-06-22",
  "publishDateEnd": "2026-06-22",
  "_t": 1719052800000
}
```

**Parameters:**
| Field | Type | Description |
|---|---|---|
| `pageNo` | int | Page number (1-based) |
| `pageSize` | int | Items per page (max 100) |
| `categoryCode` | string | Category code (see `mappings.md`) |
| `isGov` | bool | Government-only filter |
| `excludeDistrictPrefix` | array | Exclude districts by prefix |
| `publishDateBegin` | string | Start date (YYYY-MM-DD) |
| `publishDateEnd` | string | End date (YYYY-MM-DD) |
| `_t` | int | Timestamp in milliseconds |

**Note:** The old API format (`districtCode`, `isZcy`) is deprecated and returns 422.

### Response Format

```json
{
  "success": true,
  "result": {
    "data": {
      "total": 150,
      "data": [
        {
          "articleId": "base64_encoded_id",
          "projectName": "项目名称",
          "title": "公告标题",
          "publishDate": 1719052800000,
          "districtCode": "331102",
          "districtName": "浙江省丽水市莲都区",
          "pathName": "招标公告",
          "purchaseName": "采购人名称",
          "supplierName": "供应商名称",
          "author": "发布机构",
          "budgetPrice": "1000000",
          "totalContractAmount": "950000"
        }
      ]
    }
  }
}
```

**Client-side filtering required:** The API does not support `districtCode` filtering in the current version. Instead:
- Request all categories with `excludeDistrictPrefix` to exclude non-target areas
- Filter results client-side by matching `districtCode` prefix (e.g. "3311" for 丽水) or `districtName`

### Pagination

- `pageSize` max: 100
- Auto-paginate: continue while `len(items) == pageSize` and `page * pageSize < total`
- Delay 200ms between pages

## 2. Detail API (公告详情)

### Endpoint

```
GET /portal/detail?articleId={articleId}&timestamp={timestamp_ms}
```

### Parameters
| Field | Type | Description |
|---|---|---|
| `articleId` | string | URL-encoded article ID from list response |
| `timestamp` | int | Current time in milliseconds |

### Response Format

```json
{
  "success": true,
  "result": {
    "data": {
      "content": "HTML 格式的公告正文内容",
      "announcementLinkDtoList": [
        {
          "articleId": "linked_article_id",
          "title": "关联公告标题"
        }
      ],
      "author": "发布机构名称",
      "pathName": "中标（成交）结果公告"
    }
  }
}
```

### Key Fields

| Field | Description |
|---|---|
| `content` | HTML announcement body — main source for extraction |
| `announcementLinkDtoList` | Linked announcements (for cross-referencing) |
| `author` | Publishing agency — fallback when extraction fails |
| `pathName` | Announcement type path name |

## 3. Detail Page URL (Public)

```
https://zfcg.czt.zj.gov.cn/site/detail?articleId={articleId}
```

URL-encode the `articleId` parameter. Used as the output link in Excel.

## Performance & Rate Limiting

- Add 300ms delay between detail requests to avoid rate limiting
- Add 200ms delay between list page requests
- Single-day collection (<50 records): ~180s total
- Multi-day collection (>100 records): use 600s timeout or split into 2-day batches
- Never collect >200 records in one run — split into batches
- Recommended: collect at 18:00+ since afternoon announcements may be posted late
