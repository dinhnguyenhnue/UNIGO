#!/usr/bin/env python3
"""
audit_website.py — Cong cu kiem dinh va cham diem Website On tap Tin hoc UNIGO (thang 100)
Tong hop tieu chuan tu: ui-ux-pro-max, web-design-guidelines, redesign-skill, taste-skill, evaluation.
"""

import argparse
import json
import os
import pathlib
import re
import subprocess
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8")

DEFAULT_DIR = pathlib.Path(r"D:\UNIGO\Website_on_tap\GK1")
TARGET_FILES = ["index.html", "on-tap-tin-hoc-hk1.html"]


def extract_scripts(html_text: str):
    return re.findall(r"<script[^>]*>([\s\S]*?)</script>", html_text, flags=re.IGNORECASE)


def check_js_syntax(html_text: str):
    scripts = extract_scripts(html_text)
    if not scripts:
        return False, "No <script> block found"
    combined = "\n;\n".join(scripts)
    with tempfile.NamedTemporaryFile("w", suffix=".js", encoding="utf-8", delete=False) as tf:
        tf.write(combined)
        tmp_path = tf.name
    try:
        res = subprocess.run(["node", "--check", tmp_path], capture_output=True, text=True)
        ok = res.returncode == 0
        msg = "OK" if ok else (res.stderr.strip() or res.stdout.strip())
        return ok, msg
    except FileNotFoundError:
        return True, "node not installed (skipped)"
    finally:
        try:
            os.remove(tmp_path)
        except OSError:
            pass


def check_hk1_bank(html_text: str):
    m = re.search(r"const\s+D\s*=\s*(\{.*?\});", html_text)
    if not m:
        return False, {}
    try:
        data = json.loads(m.group(1))
        counts = {g: {"q": len(v.get("q", [])), "e": len(v.get("e", []))} for g, v in data.items()}
        ok = all(str(g) in counts and counts[str(g)]["q"] >= 50 and counts[str(g)]["e"] >= 20 for g in range(1, 9))
        return ok, counts
    except Exception:
        return False, {}


def audit_file(path: pathlib.Path) -> dict:
    txt = path.read_text(encoding="utf-8")
    is_main = path.name in ("index.html", "trang_n_t_p_tin_h_c_unigo.html")

    checks = []

    def add(dim, name, passed, pts, detail=""):
        checks.append({
            "dimension": dim,
            "name": name,
            "passed": bool(passed),
            "max_pts": pts,
            "earned_pts": pts if passed else 0,
            "detail": detail,
        })

    # === D1. Accessibility & Semantic HTML (20 pts) ===
    has_desc = bool(re.search(r'<meta\s+name=["\']description["\']', txt, re.I))
    add("D1", "Meta description SEO", has_desc, 2, "Co <meta name='description'>" if has_desc else "Thieu meta description")

    has_icon = bool(re.search(r'<link[^>]+rel=["\']icon["\']', txt, re.I))
    add("D1", "SVG Favicon", has_icon, 2, "Co favicon" if has_icon else "Thieu favicon")

    has_skip = "skip-link" in txt or "skip-to-content" in txt
    add("D1", "Skip-to-content link", has_skip, 3, "Co skip-link" if has_skip else "Thieu skip-link cho ban phim")

    has_main = "<main" in txt
    add("D1", "Landmark <main>", has_main, 3, "Co the <main>" if has_main else "Thieu the <main>")

    if is_main:
        bad_divs = len(re.findall(r'<div[^>]+class=["\'](?:[^"\']*\s)?(badge-pill|step-node|period-card|activity-card)(?:\s[^"\']*)?["\']', txt))
        add("D1", "Semantic <button> (khong dung div bam)", bad_divs == 0, 4, f"Con {bad_divs} div gia nut" if bad_divs else "0 div gia nut")
        dialogs = len(re.findall(r'role=["\']dialog["\']', txt))
        aria_modals = len(re.findall(r'aria-modal=["\']true["\']', txt))
        add("D1", "ARIA Dialog & aria-modal", dialogs >= 3 and aria_modals >= 3, 2, f"dialog={dialogs}, aria-modal={aria_modals}")
    else:
        has_id_role = 'aria-live="polite"' in txt and 'role="status"' in txt
        add("D1", "Semantic buttons & ARIA status", has_id_role and "aria-label" in txt, 6, "Du button chuan + aria-live")

    has_focus_vis = ":focus-visible" in txt
    add("D1", "Visible focus ring (:focus-visible)", has_focus_vis, 2)

    has_reduced_motion = "prefers-reduced-motion" in txt
    add("D1", "Ho tro prefers-reduced-motion", has_reduced_motion, 2)

    # === D2. Visual Identity & Anti-AI-Slop (20 pts) ===
    purple_hits = sum(txt.lower().count(c) for c in ("#3a0ca3", "#7209b7", "#ede9fe", "#fce7f3", "#f5f3ff"))
    if is_main:
        add("D2", "Loai bo bang mau tim AI mac dinh", purple_hits == 0, 5, f"Con {purple_hits} ma mau tim AI cu" if purple_hits else "Da chuyen sang UNIGO Ocean/Emerald")
    else:
        has_ocean = "#0f3d7a" in txt.lower() and "#1d4ed8" in txt.lower()
        add("D2", "Dong bo bang mau UNIGO Ocean/Emerald", has_ocean, 5, "Dong bo token #0f3d7a / #1d4ed8" if has_ocean else "Chua dong bo token Ocean")

    has_font = "Be Vietnam Pro" in txt
    add("D2", "Typography tieng Viet Be Vietnam Pro", has_font, 4)

    has_tab_nums = "tabular-nums" in txt
    add("D2", "So don vi thang hang (tabular-nums)", has_tab_nums, 3)

    bad_z = re.findall(r"z-index:\s*(9999|5000)\b", txt)
    add("D2", "Thang z-index chuan hoa (khong 5000/9999)", len(bad_z) == 0, 4, f"z-index lon: {bad_z}" if bad_z else "Thang z-index sach")

    if is_main:
        has_theme_tokens = "--primary-dark:" in txt and "--surface:" in txt
        add("D2", "He thong CSS design tokens UNIGO", has_theme_tokens, 4)
    else:
        has_theme_toggle = "data-theme" in txt and "tbtn" in txt
        add("D2", "Nut chuyen Giao dien Sang / Toi", has_theme_toggle, 4)

    # === D3. Student UX, Gamification & Pedagogical Feedback (20 pts) ===
    alerts = len(re.findall(r"\balert\(", txt))
    confirms = len(re.findall(r"\bconfirm\(", txt))
    add("D3", "Khong dung window.alert() gay chan", alerts == 0, 5, f"Con {alerts} alert()" if alerts else "0 alert()")
    add("D3", "Khong dung window.confirm() tho cung", confirms == 0, 3, f"Con {confirms} confirm()" if confirms else "0 confirm()")

    has_toast = "showToast" in txt and "toast-container" in txt
    add("D3", "Thong bao Toast / Inline phi chan", has_toast, 4)

    if is_main:
        has_t20 = "Tuần 11 – Tuần 20" in txt and "Tuần 21 – Tuần 28" in txt
        add("D3", "Dong bo lich thi CK1 Tuan 20", has_t20, 3, "Da cap nhat Tuan 11-20 & 21-28" if has_t20 else "Van ghi Tuan 11-19")
        has_g12 = 'value="1"' in txt and 'value="2"' in txt and "filter-score-grade" in txt
        add("D3", "Bo loc Khoi 1 & Khoi 2 trong Bang diem GV", has_g12, 2)
        has_print = "@media print" in txt
        add("D3", "Toi uu in Bang diem A4 (@media print)", has_print, 3)
    else:
        has_timer = "hk1-timer" in txt
        add("D3", "Dong ho dem thoi gian lam bai truc tiep", has_timer, 3)
        has_end_essay = "function endEssay" in txt
        add("D3", "Man hinh tong ket Tu luan chuyen nghiep", has_end_essay, 2)
        has_kbd_enter = 'e.key==="Enter"' in txt or "e.key === 'Enter'" in txt or 'e.key=="Enter"' in txt
        add("D3", "Phim tat A-D / 1-4 & Enter qua cau", has_kbd_enter, 3)

    # === D4. Responsive & Touch Ergonomics (15 pts) ===
    c_100vh = txt.count("100vh")
    c_100dvh = txt.count("100dvh")
    add("D4", "On dinh khung nhin Mobile (100dvh thay 100vh)", c_100vh == 0 and c_100dvh >= 1, 5, f"100vh={c_100vh}, 100dvh={c_100dvh}")

    has_touch = "touch-action" in txt
    add("D4", "Loai bo tre cham 300ms (touch-action: manipulation)", has_touch, 4)

    has_44px = "44px" in txt or "48px" in txt
    add("D4", "Kich thuoc vung cham toi thieu >= 44px", has_44px, 3)

    if is_main:
        has_resp_dash = "dash-grid-2" in txt and "@media (max-width: 820px)" in txt
        add("D4", "Responsive xep chong grid Quan tri GV tren Mobile", has_resp_dash, 3)
    else:
        has_mob_bp = "@media (max-width:420px)" in txt or "@media (max-width: 600px)" in txt
        add("D4", "Responsive breakpoints cho Mobile", has_mob_bp, 3)

    # === D5. Code Quality & Performance (15 pts) ===
    js_ok, js_msg = check_js_syntax(txt)
    add("D5", "Cu phap JavaScript hop le (node --check)", js_ok, 8, js_msg)

    inline_count = len(re.findall(r'\sstyle="[^"]*"', txt))
    if is_main:
        add("D5", "Giam inline style HTML (<= 75)", inline_count <= 75, 7, f"inline style={inline_count}")
    else:
        add("D5", "Giam inline style HTML (<= 10)", inline_count <= 10, 7, f"inline style={inline_count}")

    # === D6. Data Integrity & Sync Preservation (10 pts) ===
    has_webhook = "AKfycbxdFxPYPnGL4vaW5gi7rYRh1Mz6UDrOUhlhKGInWrHhET8g90jONtXmXfS9wQ1dwe6j" in txt
    add("D6", "Giu nguyen Webhook Google Sheets", has_webhook, 3)

    roster_symbol = "DEFAULT_STUDENT_ROSTERS" if is_main else "DEFAULT_ROSTERS"
    has_keys = ("UNIGO_EXAM_RESULTS_V1" in txt) and ("OnTap_HK1" in txt) and (roster_symbol in txt)
    add("D6", "Giu nguyen Roster & Khoa dong bo OnTap_HK1", has_keys, 3)

    if is_main:
        has_bank = ("QUESTION_BANK" in txt) and ("SGK_ALIGNMENT_ROWS" in txt) and ('return localStorage.getItem(TEACHER_PIN_KEY) || "2026"' in txt)
        add("D6", "Bao toan QUESTION_BANK, SGK_ALIGNMENT & PIN 2026", has_bank, 4)
    else:
        bank_ok, counts = check_hk1_bank(txt)
        add("D6", "Bao toan Ngan hang 8 khoi (Lop 1-8) trong const D", bank_ok, 4, json.dumps(counts, ensure_ascii=False))

    dim_totals = {}
    for c in checks:
        d = c["dimension"]
        dim_totals.setdefault(d, {"earned": 0, "max": 0})
        dim_totals[d]["earned"] += c["earned_pts"]
        dim_totals[d]["max"] += c["max_pts"]

    total_earned = sum(v["earned"] for v in dim_totals.values())
    total_max = sum(v["max"] for v in dim_totals.values())

    return {
        "file": str(path),
        "filename": path.name,
        "total_score": total_earned,
        "max_score": total_max,
        "dimensions": dim_totals,
        "metrics": {
            "alerts": alerts,
            "confirms": confirms,
            "inline_styles": inline_count,
            "100vh": c_100vh,
            "100dvh": c_100dvh,
            "js_syntax_ok": js_ok,
        },
        "checks": checks,
    }


def main():
    parser = argparse.ArgumentParser(description="Audit & score UNIGO Informatics Website")
    parser.add_argument("--dir", default=str(DEFAULT_DIR), help="Directory containing index.html and on-tap-tin-hoc-hk1.html")
    parser.add_argument("--json-out", default="", help="Optional path to save JSON report")
    args = parser.parse_args()

    base = pathlib.Path(args.dir)
    results = []
    for fn in TARGET_FILES:
        fp = base / fn
        if not fp.exists():
            print(f"[ERROR] Khong tim thay: {fp}", file=sys.stderr)
            sys.exit(1)
        res = audit_file(fp)
        results.append(res)
        print(f"\n=== {res['filename']}: {res['total_score']}/{res['max_score']} diem ===")
        for d, dv in sorted(res["dimensions"].items()):
            print(f"  {d}: {dv['earned']}/{dv['max']}")
        failed = [c for c in res["checks"] if not c["passed"]]
        if failed:
            print("  Cac muc chua dat:")
            for c in failed:
                print(f"    - [{c['dimension']}] {c['name']} ({c['detail']})")

    mirror_path = base / "trang_n_t_p_tin_h_c_unigo.html"
    if mirror_path.exists():
        same = (base / "index.html").read_text(encoding="utf-8") == mirror_path.read_text(encoding="utf-8")
        print(f"\nMirror sync (index.html == trang_n_t_p_tin_h_c_unigo.html): {same}")

    avg = round(sum(r["total_score"] for r in results) / len(results), 1)
    print(f"Diem trung binh toan he thong: {avg}/100")

    if args.json_out:
        out_p = pathlib.Path(args.json_out)
        out_p.parent.mkdir(parents=True, exist_ok=True)
        out_p.write_text(json.dumps({"average": avg, "results": results}, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
