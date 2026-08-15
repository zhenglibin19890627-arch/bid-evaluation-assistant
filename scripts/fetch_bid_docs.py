# -*- coding: utf-8 -*-
"""
fetch_bid_docs.py —— 自动下载招标文件（v3，已验证平台真实链路）

登录：先运行 browser_login.py 完成一次登录（登录态存本机浏览器档案）。

下载链路（政采云真实接口，全程通过 UI 点击触发 + 响应截获 + OSS 直链下载）：
1. 「我的采购文件」列表页按项目名称搜索（UI 点击）；
2. 命中行点击「下载采购文件」→ 弹层出现，触发 getPurFile 接口（返回文件名/fileOssId）；
3. 点击弹层「下载」按钮 → 触发 zcy/obs/file/getDownLoadUrl 接口（返回带签名的 OSS 直链）；
4. 用签名直链直接下载文件（OSS 无需登录 cookie），保存为平台文件名；
5. 未命中（尚未获取）→ 打开获取入口页自动填表（联系人三项）提交申请 → 回列表重试；
6. 每下载成功一个文件，自动写入 标书路径/标书映射.json（文件名 → 项目名）。

用法：
    python scripts/fetch_bid_docs.py --out-dir "标书路径"
    python scripts/fetch_bid_docs.py --out-dir "标书路径" --only-url "https://zfcg.czt.zj.gov.cn/site/detail?articleId=..."

依赖：pip install playwright（浏览器用系统 Edge，无需额外下载）
"""
from __future__ import annotations

import argparse
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

DETAIL_API = "https://zfcg.czt.zj.gov.cn/portal/detail"
ACQUIRE_RE = re.compile(r"https://www\.zcygov\.cn/bid-enroll/acquirepurfile/launch/(\d+)")
LIST_URL = "https://www.zcygov.cn/bid-enroll/_procurement_/acquirepurfile/list?_app_=zcy.procurement"
WORKBENCH_URL = "https://middle.zcygov.cn/guide/supplier"

UA = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}


def default_profile_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return os.path.join(base, "投标评估助手", "浏览器档案")


def default_contact_file() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return os.path.join(base, "投标评估助手", "申请信息.json")


def load_items(json_path: str) -> list:
    data = json.load(open(json_path, encoding="utf-8"))
    return data.get("items", [])


def load_contact(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    try:
        return json.load(open(path, encoding="utf-8"))
    except Exception:
        return {}


def fetch_acquire_link(announce_url: str) -> str:
    qs = urllib.parse.urlparse(announce_url).query
    article = urllib.parse.parse_qs(qs).get("articleId", [""])[0]
    if not article:
        return ""
    api = "%s?articleId=%s&timestamp=%d" % (DETAIL_API, urllib.parse.quote(article), int(time.time() * 1000))
    try:
        req = urllib.request.Request(api, headers=UA)
        j = json.loads(urllib.request.urlopen(req, timeout=30).read().decode("utf-8", "ignore"))
        content = (((j.get("result") or {}).get("data") or {}).get("content")) or ""
        m = ACQUIRE_RE.search(content)
        return m.group(0) if m else ""
    except Exception:
        return ""


def safe_filename(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', "_", name or "")
    return name.strip() or "招标文件"


def _body(page) -> str:
    try:
        return page.inner_text("body") or ""
    except Exception:
        return ""


def _find_input(page, keys):
    for inp in page.query_selector_all("input"):
        try:
            ph = inp.get_attribute("placeholder") or ""
            label = page.evaluate(
                """(el) => {
                    const wrap = el.closest('.el-form-item,.ant-form-item,.form-item,div');
                    return wrap ? (wrap.innerText || '').split('\\n')[0] : '';
                }""", inp)
            if any(k in (ph + label) for k in keys):
                return inp
        except Exception:
            continue
    return None


def _click_text(page, texts):
    for text in texts:
        try:
            loc = page.get_by_text(text, exact=False).first
            if loc.count() > 0:
                loc.click()
                return text
        except Exception:
            continue
    return None


def goto_list_page(context, page):
    page.goto(LIST_URL, wait_until="domcontentloaded", timeout=45000)
    page.wait_for_timeout(5000)
    if "无访问权限" in _body(page):
        print("  ↳ 直接访问列表被拒，改走工作台入口…")
        page.goto(WORKBENCH_URL, wait_until="domcontentloaded", timeout=45000)
        page.wait_for_timeout(5000)
        _click_text(page, ("获取采购文件",))
        page.wait_for_timeout(4000)
        for p in context.pages:
            if "acquirepurfile/list" in (p.url or ""):
                p.bring_to_front()
                p.wait_for_timeout(3000)
                return p
    return page


def find_row(page, name: str):
    """在结果表格中定位目标行：优先整名匹配，否则取最长公共前缀命中行。
    避免同名系列项目（如'丽水学院…标4'与'丽水学院…半导体'）互相误配。"""
    rows = page.locator("tr")
    n = rows.count()
    if n == 0:
        return None
    best, best_len = None, 0
    for i in range(n):
        try:
            t = (rows.nth(i).inner_text() or "").replace("\n", "").replace(" ", "")
        except Exception:
            continue
        if not t or "项目信息" in t:
            continue
        if name.replace(" ", "") in t:
            return rows.nth(i)
        # 最长公共前缀（去空格后）
        k = 0
        a = name.replace(" ", "")
        while k < min(len(a), len(t)) and a[k] == t[k]:
            k += 1
        if k > best_len:
            best, best_len = rows.nth(i), k
    return best if best_len >= 12 else None


def ui_search(page, name: str) -> bool:
    """列表页 UI 搜索；返回是否出现目标行。
    搜索词去掉括号后缀（半截括号会导致平台模糊匹配失败），并两级降级重试。"""
    inp = _find_input(page, ("项目名称",))
    if inp is None:
        return False
    base = re.sub(r"[（(].*$", "", name).strip()
    for term in (base[:20], base[:10]):
        inp.fill(term)
        # 必须点击输入框所在表单内的「搜索」按钮：
        # get_by_text(非精确) 会误中帮助说明文字；页头还有商品搜索按钮，需用表单范围限定
        clicked = False
        try:
            btn = inp.locator("xpath=ancestor::form[1]//button[contains(., '搜索')]")
            if btn.count() > 0:
                btn.first.click()
                clicked = True
        except Exception:
            pass
        if not clicked:
            page.get_by_text("搜索", exact=True).first.click()
        page.wait_for_timeout(3500)
        if find_row(page, name) is not None:
            return True
    return False


def apply_acquire(page, acquire_url: str, contact: dict) -> bool:
    """未获取项目：入口页填表 + 提交申请。"""
    page.on("dialog", lambda d: d.accept())
    page.goto(acquire_url, wait_until="domcontentloaded", timeout=45000)
    page.wait_for_timeout(3500)
    if "login" in page.url:
        print("  ✗ 入口页跳回登录（登录态失效）")
        return False
    body = _body(page)
    if any(k in body for k in ("已获取", "已完成")):
        print("  ↳ 已获取（此前已申请），跳过申请")
        return True
    if not any(k in body for k in ("提交申请", "提交")):
        print("  ↳ 入口页无提交按钮（结构变化）")
        return False
    if not contact.get("联系人姓名"):
        print("  ↳ 缺少联系人信息（%s），无法自动申请" % default_contact_file())
        return False
    for keys, val in ((("联系人姓名", "姓名"), contact.get("联系人姓名", "")),
                      (("手机号", "手机", "联系电话"), contact.get("项目经办人手机号", "")),
                      (("邮箱", "email"), contact.get("邮箱", ""))):
        inp = _find_input(page, keys)
        if inp:
            inp.fill(val)
    time.sleep(1)
    _click_text(page, ("提交申请", "提交"))
    print("  ↳ 已点击提交申请，等待平台处理…")
    page.wait_for_timeout(8000)
    return True


def main() -> int:
    parser = argparse.ArgumentParser(description="自动下载招标文件（复用浏览器登录态）")
    parser.add_argument("--out-dir", required=True, help="下载保存目录（建议=标书路径）")
    parser.add_argument("--json", default=None, help="待研判清单 JSON 路径（默认 %TEMP%/投标评估助手/待研判清单.json）")
    parser.add_argument("--only-url", default=None, help="只处理指定公告 URL（调试用）")
    parser.add_argument("--profile", default=None, help="浏览器档案目录（默认与 browser_login 一致）")
    parser.add_argument("--contact", default=None, help="申请信息 JSON 路径（默认 %%LOCALAPPDATA%%\\投标评估助手\\申请信息.json）")
    parser.add_argument("--browser", default="msedge", choices=("msedge", "chrome", "chromium"),
                        help="浏览器内核：msedge（系统自带 Edge，推荐，零下载）/ chrome / chromium")
    parser.add_argument("--headless", action="store_true", help="无界面运行（政采云 WAF 可能拦截无头浏览器，不推荐）")
    parser.add_argument("--manual-wait", type=int, default=90, help="自动流程失败后等待人工操作的秒数（默认 90）")
    args = parser.parse_args()

    json_path = args.json or os.path.join(os.environ.get("TEMP", ""), "投标评估助手", "待研判清单.json")
    if not os.path.exists(json_path):
        print("未找到待研判清单：%s" % json_path)
        return 1
    items = load_items(json_path)
    if args.only_url:
        items = [it for it in items if it.get("url") == args.only_url]
        if not items:
            print("URL 不在待研判清单中：%s" % args.only_url)
            return 1
    os.makedirs(args.out_dir, exist_ok=True)
    profile_dir = args.profile or default_profile_dir()
    if not os.path.isdir(profile_dir):
        print("浏览器档案不存在：%s\n请先运行 python scripts/browser_login.py 完成登录。" % profile_dir)
        return 1
    contact = load_contact(args.contact or default_contact_file())

    from playwright.sync_api import sync_playwright

    print("共 %d 个项目" % len(items))
    plans = []
    for it in items:
        link = fetch_acquire_link(it.get("url", ""))
        plans.append({"item": it, "acquire": link})
        print("  [%s] %s" % ("有入口" if link else "无入口", it.get("project_name", "")[:36]))
    print("有获取入口：%d 个；无入口：%d 个" % (len([p for p in plans if p["acquire"]]), len([p for p in plans if not p["acquire"]])))

    results = []
    mapping_file = os.path.join(args.out_dir, "标书映射.json")

    with sync_playwright() as pw:
        launch_kwargs = dict(
            user_data_dir=profile_dir,
            headless=args.headless,
            locale="zh-CN",
            accept_downloads=True,
            args=["--disable-blink-features=AutomationControlled"],
        )
        if args.browser in ("msedge", "chrome"):
            launch_kwargs["channel"] = args.browser
        try:
            context = pw.chromium.launch_persistent_context(**launch_kwargs)
        except Exception as e:
            print("浏览器启动失败：%s" % e)
            return 1

        state_path = os.path.join(profile_dir, "storage_state.json")
        if os.path.exists(state_path):
            try:
                st = json.load(open(state_path, encoding="utf-8"))
                saved = [c for c in st.get("cookies", []) if c.get("value")]
                if saved:
                    context.add_cookies(saved)
                    print("已从 storage_state 注入 %d 个 cookie" % len(saved))
            except Exception as e:
                print("storage_state 注入失败（忽略）：%s" % e)

        check = context.new_page()
        try:
            check.goto(WORKBENCH_URL, wait_until="domcontentloaded", timeout=30000)
            check.wait_for_timeout(2500)
            if "login" in (check.url or ""):
                print("登录态已失效。请先运行 python scripts/browser_login.py 重新登录。")
                check.close()
                context.close()
                return 1
            print("登录态预检通过")
        finally:
            try:
                check.close()
            except Exception:
                pass

        page = context.new_page()
        # 响应截获：getPurFile（文件清单 fileOssId→name）+ getDownLoadUrl（签名直链 fileId→URL）
        pur_files = {}   # fileOssId → name（弹层打开时触发，每次只对应当前行项目）
        signed_urls = {}  # fileId → 签名 URL

        def on_response(r):
            try:
                if "getPurFile?" in r.url:
                    j = json.loads(r.text())
                    for f in (j.get("result") or []):
                        key = f.get("fileOssId") or f.get("fileUrl") or ""
                        if key:
                            pur_files[key] = f.get("name", "")
                elif "getDownLoadUrl" in r.url:
                    j = json.loads(r.text())
                    signed = j.get("result") or ""
                    if signed:
                        qs = urllib.parse.urlparse(r.url).query
                        file_id = urllib.parse.parse_qs(qs).get("fileId", [""])[0]
                        signed_urls[file_id] = signed
            except Exception:
                pass
        page.on("response", on_response)

        page = goto_list_page(context, page)

        for i, plan in enumerate(plans):
            it = plan["item"]
            name = it.get("project_name", "")
            print("\n[%d/%d] %s" % (i + 1, len(plans), name[:50]))
            saved_files = []

            # 0) 无获取入口（电子卖场竞价/询价类）：平台无采购文件可下，直接跳过（不人工兜底）
            if not plan["acquire"]:
                results.append({"name": name, "status": "无入口（电子卖场类）", "file": ""})
                continue

            # 1) UI 搜索 + 弹层下载（弹层内逐个点击全部「下载」按钮，捕获所有文件签名）
            try:
                if ui_search(page, name):
                    row = find_row(page, name)
                    if row is None:
                        print("  ✗ 搜索结果中无目标行")
                    else:
                        link = row.get_by_text("下载采购文件", exact=True).first
                        if link.count() > 0:
                            link.click()  # 弹层 + getPurFile
                            page.wait_for_timeout(3500)
                            btns = page.get_by_text("下载", exact=True)
                            n_btns = btns.count()
                            print("  ↳ 弹层内下载按钮 %d 个，逐个点击…" % n_btns)
                            for b in range(n_btns):
                                try:
                                    btns.nth(b).click()
                                    page.wait_for_timeout(2500)
                                except Exception:
                                    continue
                            _click_text(page, ("关闭", "取消", "×"))
                            page.wait_for_timeout(1000)
            except Exception as e:
                print("  ✗ 弹层下载异常：%s" % e)

            # 2) 未命中 → 申请获取 → 回列表重试
            if not saved_files and plan["acquire"] and find_row(page, name) is None:
                print("  ↳ 列表未命中，尝试申请获取…")
                try:
                    if apply_acquire(page, plan["acquire"], contact):
                        page = goto_list_page(context, page)
                        if ui_search(page, name):
                            row = find_row(page, name)
                            if row is None:
                                print("  ✗ 申请后搜索结果中无目标行")
                                continue
                            link = row.get_by_text("下载采购文件", exact=True).first
                            if link.count() > 0:
                                link.click()
                                page.wait_for_timeout(3500)
                                btns = page.get_by_text("下载", exact=True)
                                print("  ↳ 弹层内下载按钮 %d 个，逐个点击…" % btns.count())
                                for b in range(btns.count()):
                                    try:
                                        btns.nth(b).click()
                                        page.wait_for_timeout(2500)
                                    except Exception:
                                        continue
                                _click_text(page, ("关闭", "取消", "×"))
                except Exception as e:
                    print("  ✗ 申请流程异常：%s" % e)

            # 3) 用截获的签名直链下载（文件名优先用 getPurFile 的平台原始名）
            for file_id, signed in signed_urls.items():
                fname = pur_files.get(file_id) or safe_filename(os.path.basename(urllib.parse.urlparse(signed).path))
                try:
                    req = urllib.request.Request(signed, headers=UA)
                    with urllib.request.urlopen(req, timeout=180) as resp:
                        data = resp.read()
                    if len(data) < 200:
                        print("  ✗ 文件过小（%d 字节），疑似失败" % len(data))
                        continue
                    dest = os.path.join(args.out_dir, fname)
                    n = 1
                    stem, ext = os.path.splitext(fname)
                    while os.path.exists(dest):
                        dest = os.path.join(args.out_dir, "%s_%d%s" % (stem, n, ext))
                        n += 1
                    with open(dest, "wb") as fp:
                        fp.write(data)
                    print("  ✓ 已下载：%s（%d 字节）" % (os.path.basename(dest), len(data)))
                    saved_files.append(dest)
                except Exception as e:
                    print("  ✗ 直链下载失败：%s" % e)
            # 清空本次已用的截获（避免串项目）
            signed_urls.clear()
            pur_files.clear()

            # 4) 人工兜底
            if not saved_files and not args.headless:
                print("     自动流程未完成，请在弹出的浏览器页面手动操作（%d 秒）…" % args.manual_wait)
                try:
                    with page.expect_download(timeout=args.manual_wait * 1000) as dl_info:
                        pass
                    dl = dl_info.value
                    fname = safe_filename(dl.suggested_filename)
                    dest = os.path.join(args.out_dir, fname)
                    dl.save_as(dest)
                    saved_files.append(dest)
                    print("  ✓ 人工完成，已保存：%s" % fname)
                except Exception:
                    pass

            if saved_files:
                # 映射优先"招标文件类"文件（公平竞争审查表/回执等附件只保存不映射，
                # 避免流水线匹配时多个文件对应同一项目造成歧义）
                TENDER_KWS = ("招标", "采购文件", "磋商", "谈判", "询价", "标书", "竞争性")
                tender_like = [f for f in saved_files
                               if any(k in os.path.basename(f) for k in TENDER_KWS)
                               or os.path.splitext(f)[1].lower() in (".doc", ".docx")]
                map_targets = tender_like or saved_files
                for dest in map_targets:
                    try:
                        from scripts import bid_mapping
                        bid_mapping.add_bid_mapping(os.path.basename(dest), name, path=mapping_file)
                        print("  ↳ 已写入标书映射：%s → %s" % (os.path.basename(dest), name[:30]))
                    except Exception as e:
                        print("  ↳ 标书映射写入失败：%s" % e)
                if len(map_targets) < len(saved_files):
                    for f in saved_files:
                        if f not in map_targets:
                            print("  ↳ 附件未映射（仅保存）：%s" % os.path.basename(f))
                results.append({"name": name, "status": "已下载", "file": "、".join(os.path.basename(f) for f in saved_files)})
                time.sleep(2)
                continue
            results.append({"name": name, "status": "需人工处理", "file": ""})
        context.close()

    ok = [r for r in results if r["status"] == "已下载"]
    print("\n" + "=" * 60)
    print("下载汇总：成功 %d | 需人工 %d" % (len(ok), len([r for r in results if r["status"] == "需人工处理"])))
    for r in results:
        print("  [%s] %s %s" % (r["status"], r["name"][:40], r["file"][:50]))
    return 0


if __name__ == "__main__":
    sys.exit(main())
