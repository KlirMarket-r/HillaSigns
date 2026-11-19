"""
Hilla's Signs: sync from the Kupot back office.

1. Logs into the back office (its normal login form) with KUPOT_USER / KUPOT_PASS (GitHub secrets).
2. Reads the product list from the back office API, the same JSON the מוצרים page shows.
   Read only: it never imports, saves nothing and changes nothing in the back office.
3. Rebuilds the product list inside index.html with the app's rules
   (no כפר קרע, no מבצעים, no ברקוד כללי, no קרטונים שלמים; סניף זלדמן prices, or the משאית הקסמים
   price for products sold only there) and keeps every sale on its product.
4. Writes index.html only if something actually changed.

Run locally:            KUPOT_USER=... KUPOT_PASS=... python sync.py
From a saved export:    python sync.py --file export.xlsx
From saved API JSON:    python sync.py --json products.json
"""
import datetime
import json
import os
import re
import sys
from zoneinfo import ZoneInfo

URL = os.environ.get("KUPOT_URL", "https://backoffice.kupot.co.il/")
API = os.environ.get("KUPOT_API", "https://backoffice.server.kupot.co.il/api/web/products/by-client")
APP = os.environ.get("APP_FILE", "index.html")
SHOT = "sync-failure.png"
TZ = ZoneInfo("Asia/Jerusalem")

EXCLUDED = {"כפר קרע", "מבצעים", "מבצים", "ברקוד כללי", "קרטונים שלמים"}
RENAMED = {"Unknown": "ללא קטגוריה", "": "ללא קטגוריה"}
MAIN_BRANCH = "סניף זלדמן"
TRUCK_BRANCH = "משאית הקסמים"
DATA_RE = re.compile(r"const DATA = (\[\[.*?\]\]);", re.S)
DATE_RE = re.compile(r'const PRICES_DATE = "[^"]*";')


def log(*a):
    print(*a, flush=True)


# ---------------------------------------------------------------- back office
def fetch_products():
    """Log in through the real login form, then read the product JSON with the session's token.
    The token is used for one request and never printed or saved."""
    from playwright.sync_api import sync_playwright

    user, pw = os.environ.get("KUPOT_USER"), os.environ.get("KUPOT_PASS")
    if not user or not pw:
        sys.exit("KUPOT_USER / KUPOT_PASS are not set")

    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(locale="he-IL", timezone_id="Asia/Jerusalem", viewport={"width": 1366, "height": 800})
        page = ctx.new_page()
        try:
            log("open", URL)
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)
            page.locator("input[name=email]").first.wait_for(state="visible", timeout=30000)
            page.locator("input[name=email]").first.fill(user)
            page.locator("input[name=password]").first.fill(pw)
            log("log in")
            btn = page.get_by_role("button", name=re.compile(r"^\s*כניסה\s*$"))
            if btn.count():
                btn.first.click()
            else:
                page.locator("input[name=password]").first.press("Enter")
            # Logged in = the app stored its session token
            page.wait_for_function("() => !!localStorage.getItem('token')", timeout=30000)
            token = page.evaluate("() => localStorage.getItem('token')")
            log("logged in, reading products")
            r = page.request.get(API, headers={"Authorization": "Bearer " + token, "Accept": "application/json"})
            if r.status != 200:
                raise RuntimeError(f"products API answered {r.status}")
            items = r.json()
            if not isinstance(items, list):
                raise RuntimeError("products API did not return a list")
            log(f"{len(items)} products from the back office")
            return items
        except Exception:
            try:
                page.screenshot(path=SHOT, full_page=True)
                log("screenshot saved:", SHOT)
            except Exception:
                pass
            raise
        finally:
            browser.close()


# ---------------------------------------------------------------- normalize
def norm_name(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()


def records_from_api(items):
    """API products -> [{name, barcode, cat, price}]"""
    out = []
    for p in items:
        if p.get("isDeleted"):
            continue
        prices = {}
        for d in p.get("detailsPerBranch") or []:
            try:
                prices[norm_name(d.get("branchName"))] = float(d.get("branchPrice") or 0)
            except (TypeError, ValueError):
                pass
        price = prices.get(MAIN_BRANCH) or prices.get(TRUCK_BRANCH) or 0.0
        cat = norm_name(p.get("categoryName"))
        out.append({"name": norm_name(p.get("productName")), "barcode": norm_name(p.get("barcode")),
                    "cat": RENAMED.get(cat, cat), "price": price})
    return out


def branch_price(details, branch):
    m = re.search(re.escape(branch) + r" \(([^)]*)\)", str(details))
    try:
        return float(m.group(1)) if m else 0.0
    except ValueError:
        return 0.0


def records_from_xlsx(xlsx):
    """Back office export (one sheet per category) -> [{name, barcode, cat, price}]"""
    import pandas as pd
    out = []
    for sheet, df in pd.read_excel(xlsx, sheet_name=None, dtype=str).items():
        cat = sheet.strip()
        cat = RENAMED.get(cat, cat)
        for r in df.fillna("").itertuples():
            details = getattr(r, "detailsPerBranch", "")
            out.append({"name": norm_name(r.productName), "barcode": norm_name(r.barcode), "cat": cat,
                        "price": branch_price(details, MAIN_BRANCH) or branch_price(details, TRUCK_BRANCH)})
    return out


# ---------------------------------------------------------------- build
def export_time(filename):
    """'05-11-2026_09-58_products_export.xlsx' or '19-11-2026 09-56 products export.xlsx'
    -> '05.11.2026, 09:58', else now."""
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})[_ ](\d{2})-(\d{2})", filename or "")
    if m:
        dd, mm, yy, hh, mi = m.groups()
        return f"{dd}.{mm}.{yy}, {hh}:{mi}"
    return datetime.datetime.now(TZ).strftime("%d.%m.%Y, %H:%M")


def build_rows(records, old):
    c = [r for r in records
         if r["cat"] not in EXCLUDED
         and not re.search("לא פעיל|לא בשימוש", r["name"])
         and r["price"] > 0
         and r["name"] != "כללי"]

    # Deterministic order (the API returns products in a different order on every call): categories in
    # the order the app already has (new ones last), then by name and barcode inside each category
    cat_order = {}
    for x in old:
        cat_order.setdefault(x[3], len(cat_order))
    for r in c:
        cat_order.setdefault(r["cat"], len(cat_order))
    c.sort(key=lambda r: (cat_order[r["cat"]], r["name"], r["barcode"]))

    # Sales come from the app as it is now: by barcode, or by name when the barcode changed
    # (by name only inside the same category: the truck has products with a store product's name)
    sales = {}
    for x in old:
        if len(x) > 4:
            v = (x[4], x[5], x[6] if len(x) > 6 else None, x[7] if len(x) > 7 else None)
            if x[1]:
                sales[("bc", x[1])] = v
            sales.setdefault(("nm", x[0].strip(), x[3]), v)

    rows = []
    for r in c:
        k = ("bc", r["barcode"]) if r["barcode"] and ("bc", r["barcode"]) in sales else ("nm", r["name"], r["cat"])
        price = float(r["price"])
        price = int(price) if price == int(price) else round(price, 2)
        row = [r["name"], r["barcode"], price, r["cat"]]
        if k in sales:
            q, t, g, gt = sales[k]
            row += [q, t] + ([g, gt] if g is not None else [])
        rows.append(row)
    return rows


def summary(old, new):
    ob = {x[1]: x for x in old if x[1]}
    nb = {x[1]: x for x in new if x[1]}
    on = {x[0] for x in old}
    nn = {x[0] for x in new}
    lines = [f"**{len(new)} products** (was {len(old)}), {sum(1 for x in new if len(x) > 4)} with a sale"
             f" (was {sum(1 for x in old if len(x) > 4)})", ""]
    ch = [f"- {x[0]}: ₪{ob[b][2]} → ₪{x[2]}" for b, x in nb.items() if b in ob and abs(ob[b][2] - x[2]) > 1e-9]
    if ch:
        lines += ["**Price changes (reprint these signs):**"] + ch + [""]
    added = sorted(nn - on)
    gone = sorted(on - nn)
    if added:
        lines += [f"**New or renamed ({len(added)}):** " + ", ".join(added[:40]), ""]
    if gone:
        lines += [f"**Gone or renamed ({len(gone)}):** " + ", ".join(gone[:40]), ""]
    lost = [x[0] for x in old if len(x) > 4 and x[0] not in nn and not (x[1] and x[1] in nb)]
    if lost:
        lines += ["**Sales lost (product renamed, re-attach by hand):** " + ", ".join(lost), ""]
    return "\n".join(lines)


def main():
    args = sys.argv[1:]
    if "--file" in args:
        xlsx = args[args.index("--file") + 1]
        records = records_from_xlsx(xlsx)
        stamp = export_time(os.path.basename(xlsx))
    elif "--json" in args:
        with open(args[args.index("--json") + 1], encoding="utf-8") as f:
            records = records_from_api(json.load(f))
        stamp = datetime.datetime.now(TZ).strftime("%d.%m.%Y, %H:%M")
    else:
        records = records_from_api(fetch_products())
        stamp = datetime.datetime.now(TZ).strftime("%d.%m.%Y, %H:%M")

    html = open(APP, encoding="utf-8").read()
    m = DATA_RE.search(html)
    if not m:
        sys.exit(f"No product list found in {APP}")
    old = json.loads(m.group(1))
    new = build_rows(records, old)
    if len(new) < 50:
        sys.exit(f"Only {len(new)} products came back, something is wrong; the app was not changed")

    if new == old:
        log("No changes in the back office, the app stays as it is.")
        report = "No changes since the last sync."
    else:
        data = json.dumps(new, ensure_ascii=False, separators=(",", ":"))
        html = html[:m.start()] + "const DATA = " + data + ";" + html[m.end():]
        html = DATE_RE.sub(f'const PRICES_DATE = "{stamp}";', html, count=1)
        open(APP, "w", encoding="utf-8", newline="\n").write(html)
        report = summary(old, new)
        log(report)

    step = os.environ.get("GITHUB_STEP_SUMMARY")
    if step:
        with open(step, "a", encoding="utf-8") as f:
            f.write(report + "\n")


if __name__ == "__main__":
    try:
        main()
    except SystemExit as e:
        if e.code not in (None, 0):
            print(f"::error title=sync::{e.code}", flush=True)   # shows as a GitHub annotation
        raise
    except Exception as e:
        print(f"::error title=sync::{type(e).__name__}: {str(e)[:300]}", flush=True)
        raise
