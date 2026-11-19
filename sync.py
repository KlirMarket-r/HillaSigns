"""
Hilla's Signs: sync from the Kupot back office.

1. Logs into the back office with KUPOT_USER / KUPOT_PASS (GitHub secrets).
2. Goes to מוצרים and clicks Export. It only ever exports: it never clicks import,
   saves nothing and changes nothing in the back office.
3. Rebuilds the product list inside index.html with the app's rules
   (no כפר קרע, no מבצעים, no ברקוד כללי, no קרטונים שלמים; סניף זלדמן prices, or the משאית הקסמים
   price for products sold only there) and keeps every sale on its product.
4. Writes index.html only if something actually changed.

Run locally:  KUPOT_USER=... KUPOT_PASS=... python sync.py
Use an export you already have instead of logging in:  python sync.py --file export.xlsx
"""
import datetime
import json
import os
import re
import sys
from zoneinfo import ZoneInfo

import pandas as pd

URL = os.environ.get("KUPOT_URL", "https://backoffice.kupot.co.il/")
APP = os.environ.get("APP_FILE", "index.html")
DOWNLOAD = "products_export.xlsx"
SHOT = "sync-failure.png"
TZ = ZoneInfo("Asia/Jerusalem")

EXCLUDED = {"כפר קרע", "מבצעים", "מבצים", "ברקוד כללי", "קרטונים שלמים"}
RENAMED = {"Unknown": "ללא קטגוריה"}
DATA_RE = re.compile(r"const DATA = (\[\[.*?\]\]);", re.S)
DATE_RE = re.compile(r'const PRICES_DATE = "[^"]*";')


def log(*a):
    print(*a, flush=True)


# ---------------------------------------------------------------- back office
def download_export(path):
    from playwright.sync_api import sync_playwright, TimeoutError as PwTimeout

    user, pw = os.environ.get("KUPOT_USER"), os.environ.get("KUPOT_PASS")
    if not user or not pw:
        sys.exit("KUPOT_USER / KUPOT_PASS are not set")

    def click_text(page, patterns, what, timeout=20000):
        """Click the first visible button, link or text that matches one of the patterns."""
        deadline = datetime.datetime.now() + datetime.timedelta(milliseconds=timeout)
        while datetime.datetime.now() < deadline:
            for pat in patterns:
                loc = (page.get_by_role("button", name=pat)
                       .or_(page.get_by_role("link", name=pat))
                       .or_(page.get_by_role("menuitem", name=pat))
                       .or_(page.get_by_text(pat)))
                for i in range(loc.count()):
                    el = loc.nth(i)
                    if el.is_visible():
                        txt = (el.inner_text() or "").strip()
                        if "ייבוא" in txt and "ייצוא" not in txt:
                            continue  # never touch import
                        log(f"click: {what} ({txt[:40]!r})")
                        el.click()
                        return True
            page.wait_for_timeout(500)
        return False

    with sync_playwright() as p:
        browser = p.chromium.launch()
        ctx = browser.new_context(accept_downloads=True, locale="he-IL", timezone_id="Asia/Jerusalem",
                                  viewport={"width": 1440, "height": 900})
        page = ctx.new_page()
        try:
            log("open", URL)
            page.goto(URL, wait_until="domcontentloaded", timeout=60000)

            # Log in: the username box is the first visible text field, then the password
            pw_box = page.locator("input[type=password]").first
            pw_box.wait_for(state="visible", timeout=30000)
            user_box = page.locator(
                "input:not([type=password]):not([type=hidden]):not([type=checkbox])"
                ":not([type=radio]):not([type=submit]):not([type=button])").first
            user_box.fill(user)
            pw_box.fill(pw)
            log("log in")
            def logged_in(wait):
                try:
                    pw_box.wait_for(state="hidden", timeout=wait)
                    return True
                except PwTimeout:
                    return False
            if not click_text(page, [re.compile(r"^\s*(כניסה|היכנס|התחבר(ות)?|Login|Log in|Sign in)\s*$", re.I)],
                              "login button", timeout=3000):
                pw_box.press("Enter")
            if not logged_in(10000):
                # No form to submit with Enter and no button text we know: use the page's own submit button
                btn = page.locator("button[type=submit], input[type=submit], button").filter(visible=True).first
                if btn.count():
                    log("click: login button (fallback)")
                    btn.click()
                logged_in(20000)
            page.wait_for_load_state("networkidle", timeout=60000)
            if page.locator("input[type=password]").filter(visible=True).count():
                raise RuntimeError("Still on the login page: check KUPOT_USER / KUPOT_PASS")

            # Left panel: מוצרים
            if not click_text(page, [re.compile(r"^\s*מוצרים\s*$")], "מוצרים"):
                raise RuntimeError("Couldn't find מוצרים in the menu")
            page.wait_for_load_state("networkidle", timeout=60000)

            # Export (some versions put it under ייבוא/ייצוא first)
            export = [re.compile(r"^\s*(Export|ייצוא)(\s+(לאקסל|ל-?Excel|Excel))?\s*$", re.I)]
            with page.expect_download(timeout=180000) as dl:
                if not click_text(page, export, "Export", timeout=8000):
                    if not click_text(page, [re.compile(r"ייבוא\s*.?\s*ייצוא|Import\s*.?\s*Export", re.I)], "ייבוא/ייצוא"):
                        raise RuntimeError("Couldn't find the Export button")
                    page.wait_for_timeout(1500)
                    if not click_text(page, export, "Export"):
                        raise RuntimeError("Couldn't find the Export button")
            d = dl.value
            d.save_as(path)
            log("downloaded:", d.suggested_filename)
            return d.suggested_filename
        except Exception:
            try:
                page.screenshot(path=SHOT, full_page=True)
                log("screenshot saved:", SHOT)
            except Exception:
                pass
            raise
        finally:
            browser.close()


# ---------------------------------------------------------------- build
def export_time(filename):
    """'05-11-2026_09-58_products_export.xlsx' or '19-11-2026 09-56 products export.xlsx'
    -> '05.11.2026, 09:58', else now."""
    m = re.search(r"(\d{2})-(\d{2})-(\d{4})[_ ](\d{2})-(\d{2})", filename or "")
    if m:
        dd, mm, yy, hh, mi = m.groups()
        return f"{dd}.{mm}.{yy}, {hh}:{mi}"
    return datetime.datetime.now(TZ).strftime("%d.%m.%Y, %H:%M")


def branch_price(details, branch):
    m = re.search(re.escape(branch) + r" \(([^)]*)\)", str(details))
    try:
        return float(m.group(1)) if m else 0.0
    except ValueError:
        return 0.0


def shelf_price(details):
    """The סניף זלדמן price; for products sold only on the truck, the משאית הקסמים price."""
    return branch_price(details, "סניף זלדמן") or branch_price(details, "משאית הקסמים")


def build_rows(xlsx, old):
    frames = []
    for sheet, df in pd.read_excel(xlsx, sheet_name=None, dtype=str).items():
        df["cat"] = RENAMED.get(sheet.strip(), sheet.strip())
        frames.append(df)
    d = pd.concat(frames).fillna("")
    d["price"] = d.detailsPerBranch.map(shelf_price)
    d["name"] = d.productName.str.replace(r"\s+", " ", regex=True).str.strip()
    d["barcode"] = d.barcode.astype(str).str.strip()

    c = d[~d.cat.isin(EXCLUDED)]
    c = c[~c.name.str.contains("לא פעיל|לא בשימוש")]
    c = c[c.price > 0]
    c = c[c.name != "כללי"]

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
    for r in c.itertuples():
        k = ("bc", r.barcode) if r.barcode and ("bc", r.barcode) in sales else ("nm", r.name, r.cat)
        price = float(r.price)
        price = int(price) if price == int(price) else price
        row = [r.name, r.barcode, price, r.cat]
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
    lines = [f"**{len(new)} products** (was {len(old)}), {sum(1 for x in new if len(x) > 4)} with a sale", ""]
    ch = [f"- {x[0]}: ₪{ob[b][2]} → ₪{x[2]}" for b, x in nb.items() if b in ob and abs(ob[b][2] - x[2]) > 1e-9]
    if ch:
        lines += ["**Price changes (reprint these signs):**"] + ch + [""]
    added = sorted(nn - on)
    gone = sorted(on - nn)
    if added:
        lines += [f"**New or renamed ({len(added)}):** " + ", ".join(added[:40]), ""]
    if gone:
        lines += [f"**Gone or renamed ({len(gone)}):** " + ", ".join(gone[:40]), ""]
    return "\n".join(lines)


def main():
    args = sys.argv[1:]
    if "--file" in args:
        xlsx = args[args.index("--file") + 1]
        fname = os.path.basename(xlsx)
    else:
        xlsx = DOWNLOAD
        fname = download_export(xlsx)

    html = open(APP, encoding="utf-8").read()
    m = DATA_RE.search(html)
    if not m:
        sys.exit(f"No product list found in {APP}")
    old = json.loads(m.group(1))
    new = build_rows(xlsx, old)
    if len(new) < 50:
        sys.exit(f"Only {len(new)} products in the export, something is wrong; the app was not changed")

    if new == old:
        log("No changes in the back office, the app stays as it is.")
        report = "No changes since the last sync."
    else:
        data = json.dumps(new, ensure_ascii=False, separators=(",", ":"))
        html = html[:m.start()] + "const DATA = " + data + ";" + html[m.end():]
        html = DATE_RE.sub(f'const PRICES_DATE = "{export_time(fname)}";', html, count=1)
        open(APP, "w", encoding="utf-8").write(html)
        report = summary(old, new)
        log(report)

    step = os.environ.get("GITHUB_STEP_SUMMARY")
    if step:
        with open(step, "a", encoding="utf-8") as f:
            f.write(report + "\n")


if __name__ == "__main__":
    main()
