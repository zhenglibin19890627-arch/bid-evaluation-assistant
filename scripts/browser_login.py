# -*- coding: utf-8 -*-
"""
browser_login.py —— 政采云登录（Playwright 持久化浏览器档案，方案 A）

安全原则：账号密码只在弹出的浏览器窗口里由用户本人输入，脚本不接触、不保存任何凭证。
登录态由 Chromium 加密保存在本机档案目录（默认 %LOCALAPPDATA%\\投标评估助手\\浏览器档案），
后续下载脚本复用该档案，无需重复登录；登录过期时重新运行本脚本再登一次即可。

用法：
    python scripts/browser_login.py          # 打开窗口，等待用户登录完成后自动退出
    python scripts/browser_login.py --reset  # 清空旧登录态后重新登录
    python scripts/browser_login.py --profile <目录>  # 自定义档案目录

依赖：pip install playwright && python -m playwright install chromium
"""
from __future__ import annotations

import argparse
import os
import sys
import time

LOGIN_URL = "https://login.zcygov.cn/user-login/#/login"
HOME_URL = "https://www.zcygov.cn/"


def default_profile_dir() -> str:
    base = os.environ.get("LOCALAPPDATA") or os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return os.path.join(base, "投标评估助手", "浏览器档案")


def is_logged_in(page) -> bool:
    """粗略判定登录完成：离开 login 域名，或出现常见登录态 cookie。"""
    try:
        url = page.url or ""
        if "login" not in url and "zcygov" in url:
            return True
        for c in page.context.cookies():
            n = c.get("name", "").lower()
            if c.get("value") and any(k in n for k in ("token", "session", "sso", "auth", "jid")):
                return True
    except Exception:
        pass
    return False


def main() -> int:
    parser = argparse.ArgumentParser(description="政采云登录（持久化浏览器档案）")
    parser.add_argument("--profile", default=None, help="浏览器档案目录（默认 %%LOCALAPPDATA%%\\投标评估助手\\浏览器档案）")
    parser.add_argument("--browser", default="msedge", choices=("msedge", "chrome", "chromium"),
                        help="浏览器内核：msedge（系统自带 Edge，推荐，零下载）/ chrome / chromium")
    parser.add_argument("--reset", action="store_true", help="清空旧登录态后重新登录")
    parser.add_argument("--timeout", type=int, default=15 * 60, help="等待登录的超时秒数（默认 900）")
    args = parser.parse_args()

    profile_dir = args.profile or default_profile_dir()
    os.makedirs(profile_dir, exist_ok=True)

    if args.reset:
        import shutil
        shutil.rmtree(profile_dir, ignore_errors=True)
        os.makedirs(profile_dir, exist_ok=True)
        print("已清空旧登录态，重新开始。")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("未安装 playwright。请先执行：pip install playwright && python -m playwright install chromium")
        return 1

    # msedge/chrome 使用系统已装浏览器（零下载）；chromium 需要 python -m playwright install chromium
    launch_kwargs = dict(
        user_data_dir=profile_dir,
        headless=False,
        locale="zh-CN",
        args=["--disable-blink-features=AutomationControlled"],
    )
    if args.browser in ("msedge", "chrome"):
        launch_kwargs["channel"] = args.browser

    with sync_playwright() as pw:
        try:
            context = pw.chromium.launch_persistent_context(**launch_kwargs)
        except Exception as e:
            if args.browser in ("msedge", "chrome"):
                print("用系统浏览器（%s）启动失败：%s" % (args.browser, e))
                print("请确认已安装该浏览器；或改用 chromium（python -m playwright install chromium）")
            else:
                print("chromium 启动失败：%s\n请先执行 python -m playwright install chromium" % e)
            return 1

        page = context.new_page()
        print("=" * 60)
        print("已打开登录窗口：请在浏览器中完成登录（账号密码/扫码均可）。")
        print("登录成功后脚本会自动检测并退出，请勿手动关闭浏览器窗口。")
        print("档案目录：%s" % profile_dir)
        print("=" * 60)
        page.goto(LOGIN_URL, wait_until="domcontentloaded")

        deadline = time.time() + args.timeout
        done = False
        last_info = time.time()
        while time.time() < deadline:
            if is_logged_in(page):
                done = True
                break
            if time.time() - last_info > 20:
                print("[%s] 仍在等待登录完成…（当前页面：%s）" % (
                    time.strftime("%H:%M:%S"), (page.url or "")[:80]))
                last_info = time.time()
            time.sleep(2)

        if not done:
            print("等待登录超时（%d 秒）。请重新运行本脚本继续。" % args.timeout)
            context.close()
            return 1

        # 登录检测到后不立即关窗：政采云 SSO 有异步落盘会话的阶段，
        # 等待 12 秒让会话 cookie 写入档案，再二次访问验证在线状态
        print("检测到登录，等待 SSO 会话落盘（12 秒）…")
        time.sleep(12)
        print("二次验证：访问供应商工作台…")
        try:
            page.goto("https://middle.zcygov.cn/guide/supplier", wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(4000)
        except Exception as e:
            print("验证访问失败（可忽略）：%s" % e)
        if not is_logged_in(page):
            print("警告：二次验证未通过，登录态可能未持久化。请保持窗口打开继续浏览登录后页面，")
            print("      脚本再等待 60 秒后复检…")
            time.sleep(60)
        if not is_logged_in(page):
            print("仍然未检测到持久登录态。将保存当前 cookie 并退出（可能仍需重新登录一次）。")

        # 关窗前再等待：让会话级 cookie（SESSION/SSOSESSION）有机会写入档案磁盘
        # （即使档案没落盘，storage_state.json 里也已有完整会话，后续脚本可注入恢复）
        print("等待 15 秒确保会话 cookie 写入档案…")
        time.sleep(15)

        # 落盘 storage_state，便于后续用 requests 复用会话
        state_path = os.path.join(profile_dir, "storage_state.json")
        try:
            context.storage_state(path=state_path)
            print("登录态已保存：%s" % state_path)
        except Exception as e:
            print("登录态保存失败（不影响持久化档案）：%s" % e)
        names = sorted({c.get("name", "") for c in context.cookies()})
        print("档案 cookie 名称（仅名称，不含值）：%s" % names)
        print("登录完成。当前页面：%s" % page.url)
        context.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
