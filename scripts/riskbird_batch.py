# Maintained by Lu Lingyan, Deheng (Wuxi) Law Firm. / Adapted for agent-browser by WorkBuddy
#!/usr/bin/env python3
"""批量通过风鸟(riskbird.com)查询企业社保人数、电话、邮箱（agent-browser 适配版）

适配说明：原版依赖 CDP Proxy(localhost:3456) 操作已登录 Chrome，
本版改用 agent-browser CLI 操作浏览器（风鸟不拦截 agent-browser，已实测）。

用法:
    python riskbird_batch.py <企业名单.json> [结果.json]

    企业名单.json 格式: ["企业名称1", "企业名称2", ...]
    结果默认写 riskbird_results.json

前置:
    agent-browser 已安装且用 --headed 打开过风鸟(可选，脚本会自动打开)
    建议先手动登录风鸟一次(登录态会保留在 agent-browser session)
"""
import json, subprocess, time, os, re, urllib.parse, sys

DEBUG = True

# agent-browser 可执行文件绝对路径（Windows 下裸命令名在 Python subprocess 里找不到）
_AGENT_BROWSER = os.environ.get(
    "AGENT_BROWSER_EXE",
    r"C:\Users\Administrator\.workbuddy\binaries\node\versions\22.22.2\node_modules\agent-browser\bin\agent-browser-win32-x64.exe",
)


def _clean_env():
    """清空代理环境变量（否则 agent-browser 报 ERR_NO_SUPPORTED_PROXIES）"""
    env = os.environ.copy()
    for k in ("HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy",
              "ALL_PROXY", "all_proxy"):
        env.pop(k, None)
    env["NO_PROXY"] = "*"
    env["no_proxy"] = "*"
    return env


def _ab(args, timeout=60):
    """调用 agent-browser，返回 stdout 文本"""
    r = subprocess.run([_AGENT_BROWSER] + args, capture_output=True, text=True,
                       timeout=timeout, env=_clean_env())
    return r.stdout.strip()


def navigate(url):
    _ab(["open", url])


def get_page_text(max_len=10000):
    """获取页面纯文本（agent-browser eval）"""
    out = _ab(["eval", f"document.body.innerText.slice(0,{max_len})"])
    # agent-browser eval 返回 JSON 字符串（带引号）
    try:
        val = json.loads(out)
        return val if isinstance(val, str) else out
    except Exception:
        return out


def extract_from_search_results(text, company_name):
    """从搜索结果页文本中提取目标企业的信息（保留原版逻辑）"""
    lines = text.split("\n")
    info = {"company": company_name, "phone": "", "email": "",
            "website": "", "socialSecurity": "", "staffSize": "",
            "address": "", "legalPerson": "", "capital": "",
            "establishDate": "", "status": "", "creditCode": ""}

    target_idx = -1
    for i, line in enumerate(lines):
        if line.strip() == company_name:
            target_idx = i
            break

    if target_idx < 0:
        for i, line in enumerate(lines):
            if company_name in line and len(line.strip()) > 5:
                target_idx = i
                break

    if target_idx < 0:
        return None

    for j in range(target_idx, min(target_idx + 30, len(lines))):
        line = lines[j].strip()
        if not line:
            continue

        if line in ("在营", "正常", "存续", "在营/正常", "吊销", "注销"):
            info["status"] = line

        if line.startswith("电话：") or line.startswith("电话:"):
            val = line.split("：")[-1].split(":")[-1].strip()
            m = re.search(r'(\d[\d\-]{5,})', val)
            if m:
                info["phone"] = m.group(1)

        if line.startswith("邮箱：") or line.startswith("邮箱:"):
            val = line.split("：")[-1].split(":")[-1].strip()
            m = re.search(r'([A-Za-z0-9._%-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})', val)
            if m:
                info["email"] = m.group(1)

        if line.startswith("官网：") or line.startswith("官网:") or line.startswith("网址："):
            val = line.split("：")[-1].split(":")[-1].strip()
            if val and val != "-":
                info["website"] = val

        if "通信地址" in line or "注册地址" in line:
            parts = re.split(r'[：:]', line)
            if len(parts) >= 2:
                val = parts[-1].strip()
                if val and val != "-":
                    info["address"] = val

        if "法定代表人" in line and "：" in line:
            info["legalPerson"] = line.split("：")[-1].strip()
        elif "法定代表人" in line and j + 1 < len(lines):
            next_l = lines[j + 1].strip()
            if next_l and not next_l.startswith("注册资本") and not next_l.startswith("电话"):
                info["legalPerson"] = next_l.split("：")[-1].strip()

        if "注册资本" in line and "实缴" not in line and "：" in line:
            info["capital"] = line.split("：")[-1].strip()

        if "成立日期" in line and "：" in line:
            info["establishDate"] = line.split("：")[-1].strip()

        if "统一社会信用代码" in line and "：" in line:
            info["creditCode"] = line.split("：")[-1].strip()

    if not info["phone"]:
        context = "\n".join(lines[max(0, target_idx - 2):target_idx + 20])
        phone_match = re.search(r'电话[：:]\s*(\d[\d\-]{6,})', context)
        if phone_match:
            info["phone"] = phone_match.group(1).strip()

    if not info["email"]:
        context = "\n".join(lines[max(0, target_idx - 2):target_idx + 20])
        email_match = re.search(r'邮箱[：:]\s*([\w.%-]+@[\w.-]+\.[\w]{2,})', context)
        if email_match:
            info["email"] = email_match.group(1).strip()

    return info


def query_one(company):
    """查询一家企业 - 从搜索页提取"""
    encoded = urllib.parse.quote(company)
    search_url = f"https://www.riskbird.com/search/company?keyword={encoded}&_t={int(time.time() * 1000)}"

    navigate(search_url)
    time.sleep(6)

    text = get_page_text(10000)

    if "额度已用完" in text:
        print(f"  ⛔ 额度已用完")
        return {"company": company, "error": "QUOTA_EXHAUSTED"}

    info = extract_from_search_results(text, company)

    if info is None:
        print(f"  ❌ 搜索结果中未找到该企业")
        return {"company": company, "error": "NOT_FOUND", "_preview": text[:500]}

    info["query_time"] = time.strftime("%Y-%m-%d %H:%M:%S")
    return info


def save_results(results, path):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    companies_file = sys.argv[1]
    result_file = sys.argv[2] if len(sys.argv) > 2 else "riskbird_results.json"

    with open(companies_file, "r", encoding="utf-8") as f:
        companies = json.load(f)
    total = len(companies)

    print("=" * 60)
    print(f"风鸟批量查询(agent-browser适配版) | 共 {total} 家企业 | 搜索页直接提取")
    print("=" * 60)

    results = []
    if os.path.exists(result_file):
        try:
            with open(result_file, "r", encoding="utf-8") as f:
                results = json.load(f)
            print(f"📂 已有记录: {len(results)} 条")
        except Exception:
            pass

    done = set(r.get("company", "") for r in results if r.get("company") and not r.get("error"))
    todo = [c for c in companies if c not in done]
    print(f"✅ 已完成: {len(done)} | ⏳ 待查询: {len(todo)}\n")

    if not todo:
        print("全部完成！")
        return

    success = 0
    fail = 0

    for idx, company in enumerate(todo):
        try:
            info = query_one(company)
        except Exception as e:
            info = {"company": company, "error": f"异常: {e}"}

        results.append(info)
        save_results(results, result_file)

        err = info.get("error", "")
        if err:
            fail += 1
            if "QUOTA" in err:
                break
        else:
            if info.get("phone"):
                success += 1
            else:
                fail += 1

        # 精简进度：每 10 家打印一次，不逐家打印明细（省 token）
        if (idx + 1) % 10 == 0 or idx + 1 == len(todo):
            print(f"  进度 {idx+1}/{len(todo)} | 成功 {success} | 失败 {fail}")

        time.sleep(2)

    print(f"\n{'=' * 60}")
    print(f"完成！成功: {success} | 失败: {fail} | 总计: {len(results)}")
    print(f"结果: {result_file}")

    # 不打印名单明细（结果在 JSON 里，省 token）


if __name__ == "__main__":
    main()
