#!/usr/bin/env python3
"""QSLカード レポート面 自動印刷ツール (CTESTWIN ADIF -> PDF)

事前印刷済みのQSLカード(100x148mm・縦)に、交信データだけを重ね刷りするPDFを作る。
  - 1交信 = 1ページ(1カード)
  - 印刷済みの重複防止: printed.json に記録 (QSL#は本アプリ独自の連番)
  - 依存: reportlab (pip install reportlab) / GUIは標準のtkinter

使い方:  python qsl_tool.py            -> GUI
         python qsl_tool.py log.adi    -> CLI(未印刷分をすべてPDF化)
"""
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

# ---------------------------------------------------------------- 設定 ----
APP_DIR = Path(__file__).resolve().parent
HISTORY_FILE = APP_DIR / "printed.json"   # 印刷履歴(重複チェック + QSL#採番)
CONFIG_FILE = APP_DIR / "qsl_config.json"

DEFAULT_CONFIG = {
    "rig": "",            # 例: "IC-7300"
    "output_w": "",       # 例: "50"
    "ant": "",            # 例: "Dipole"
    "qth": "",            # 例: "Kyoto"
    "op_name": "",        # 自局オペレータ名(空ならADIFのOPERATOR)
    "source_tz": "JST",   # ADIFの時刻は JST / UTC どちらか
    "print_tz": "JST",    # カードに印字する時刻 JST / UTC
    "template_image": "", # 位置合わせ確認用: カード白紙スキャン画像(png/jpg)。空なら重ね刷りのみ
    "offset_x_mm": 0.0,   # プリンタのズレ補正
    "offset_y_mm": 0.0,
    "font": "Helvetica",
    "jp_font_path": "",   # 日本語フォント(.ttf/.ttc)。空なら自動検出→無ければ内蔵CIDフォント
    "rmks_line1": "TNX FB QSO.",
    "rmks_line2": "今後ともよろしくお願いします。",
}

# 日本語フォント候補(TrueType outlineのもの。見つからなければCIDフォントにフォールバック)
JP_FONT_CANDIDATES = [
    "C:/Windows/Fonts/meiryo.ttc", "C:/Windows/Fonts/msgothic.ttc",
    "C:/Windows/Fonts/YuGothM.ttc",
    "/Library/Fonts/Arial Unicode.ttf",
    "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf",
    "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
    "/usr/share/fonts/truetype/ipafont-gothic/ipag.ttf",
]

# 白紙スキャン(896x1340px)上の座標。原点=左上。ズレたらここを調整。
SCAN_W, SCAN_H = 896, 1340
PAGE_W_MM, PAGE_H_MM = 100.0, 148.0
POS = {
    "call_box_x": [388, 470, 550, 640, 730, 820],  # TO RADIO 赤枠6つの中心x
    "call_box_y": 125,                             # 同 中心y
    "row_y": 352,                                   # 表の記入行 文字ベースライン
    "to_line_x": 80, "to_line_y": 135,             # TO RADIO下の点線欄(移動局のコール/P)
    "date_x": 62, "time_x": 305, "rst_x": 525, "mhz_x": 650,
    "twoway_x": 795,                               # 2WAY欄(中央揃え)
    "qslno_x": 160, "qslno_y": 432,
    "rig_x": 155, "rig_y": 492, "out_x": 730,
    "ant_x": 155, "ant_y": 553,
    "rmks_x": 150, "rmks_y": 612,
    "rmks_line2_y": 655,
    "free_x": 75, "free_y1": 730, "free_dy": 48,    # 下の自由記入エリア
}

BAND_MHZ = {
    "2190m": "0.136", "630m": "0.475", "160m": "1.9", "80m": "3.5", "60m": "5.3",
    "40m": "7", "30m": "10", "20m": "14", "17m": "18", "15m": "21", "12m": "24",
    "10m": "28", "6m": "50", "4m": "70", "2m": "144", "1.25m": "220",
    "70cm": "430", "33cm": "900", "23cm": "1200", "13cm": "2400",
}
BAND_ALIAS = {"430": "70cm", "430mhz": "70cm", "144": "2m", "144mhz": "2m",
              "3.5": "80m", "3.5mhz": "80m"}
HF_BANDS = {"160m", "80m", "60m", "40m", "30m", "20m", "17m", "15m", "12m", "10m"}
# Antは原則ログのバンドから自動。ここを書き換えれば変更できる。
ANT_RULES = {"80m": "1/2λDPアンテナ", "2m": "GPアンテナ", "70cm": "GPアンテナ"}
ANT_HF_DEFAULT = "1λワイヤーアンテナ"   # 上記以外のHF
MONTHS = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN",
          "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]

# ------------------------------------------------------------- ADIF -------
_TAG = re.compile(r"<(\w+):(\d+)(?::\w+)?>", re.I)


def read_text(path):
    raw = Path(path).read_bytes()
    for enc in ("utf-8", "cp932"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="replace")


def parse_adif(text):
    """ADIFを交信dictのリストにする。キーは小文字。"""
    m = re.search(r"<eoh>", text, re.I)
    body = text[m.end():] if m else text
    qsos = []
    for chunk in re.split(r"<eor>", body, flags=re.I):
        rec, pos = {}, 0
        while True:
            t = _TAG.search(chunk, pos)
            if not t:
                break
            name, n = t.group(1).lower(), int(t.group(2))
            rec[name] = chunk[t.end():t.end() + n]
            pos = t.end() + n
        if "call" in rec:
            qsos.append(rec)
    return qsos


def qso_key(q):
    """重複判定キー: コールサイン+日付+時刻+バンド+モード"""
    return "|".join(q.get(k, "").upper() for k in
                    ("call", "qso_date", "time_on", "band", "mode"))


def to_datetime(q, source_tz):
    t = (q.get("time_on", "000000") + "000000")[:6]
    dt = datetime.strptime(q["qso_date"] + t, "%Y%m%d%H%M%S")
    return dt  # source_tz の壁時計時刻


def fmt_datetime(q, source_tz, print_tz):
    dt = to_datetime(q, source_tz)
    if source_tz != print_tz:
        dt += timedelta(hours=-9 if source_tz == "JST" else 9)
    return f"{dt.day:02d} {MONTHS[dt.month - 1]} {dt.year}", dt.strftime("%H:%M"), print_tz


def mhz_of(q):
    if q.get("freq"):
        return q["freq"].rstrip("0").rstrip(".") if "." in q["freq"] else q["freq"]
    return BAND_MHZ.get(band_key(q.get("band", "")), q.get("band", ""))


def band_key(band):
    b = band.strip().lower()
    return BAND_ALIAS.get(b, b)


def ant_of(q, cfg):
    """バンドからAntを自動判定。該当しないバンドは設定のAntを使う。"""
    b = band_key(q.get("band", ""))
    if b in ANT_RULES:
        return ANT_RULES[b]
    if b in HF_BANDS:
        return ANT_HF_DEFAULT
    return cfg["ant"]


def split_call(call):
    """'JA1ABC/P' -> ('JA1ABC', '/P')。移動でなければ suffix は空。"""
    parts = [p for p in call.split("/") if p]
    if len(parts) <= 1:
        return call, ""
    base = max(parts, key=len)
    return base, "/".join(p for p in parts if p is not base)


_JP_CACHE = {}


def register_jp_font(cfg):
    """日本語フォント名を返す(失敗時None)。TrueType→無ければCIDフォント。"""
    path = cfg.get("jp_font_path", "")
    if path in _JP_CACHE:
        return _JP_CACHE[path]
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.cidfonts import UnicodeCIDFont
    from reportlab.pdfbase.ttfonts import TTFont
    name = None
    for p in ([path] if path else []) + JP_FONT_CANDIDATES:
        if p and Path(p).exists():
            try:
                pdfmetrics.registerFont(TTFont("QSL-JP", p, subfontIndex=0))
                name = "QSL-JP"
                break
            except Exception:
                continue
    if name is None:
        try:
            pdfmetrics.registerFont(UnicodeCIDFont("HeiseiKakuGo-W5"))
            name = "HeiseiKakuGo-W5"
        except Exception:
            name = None
    _JP_CACHE[path] = name
    return name


# ----------------------------------------------------------- 履歴 ---------
def load_json(path, default):
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return default


def load_history():
    return load_json(HISTORY_FILE, {"next_no": 1, "printed": {}})


def save_history(h):
    HISTORY_FILE.write_text(json.dumps(h, ensure_ascii=False, indent=1), encoding="utf-8")


def load_config():
    c = dict(DEFAULT_CONFIG)
    c.update(load_json(CONFIG_FILE, {}))
    return c


def save_config(c):
    CONFIG_FILE.write_text(json.dumps(c, ensure_ascii=False, indent=1), encoding="utf-8")


# ----------------------------------------------------------- PDF ----------
def render_pdf(qsos, out_path, cfg, history, assign_numbers=True):
    """qsos を1ページ1枚でPDF化。history に QSL# と対応を記録して返す。"""
    c = canvas.Canvas(str(out_path), pagesize=(PAGE_W_MM * mm, PAGE_H_MM * mm))
    sx, sy = PAGE_W_MM * mm / SCAN_W, PAGE_H_MM * mm / SCAN_H
    ox, oy = cfg["offset_x_mm"] * mm, -cfg["offset_y_mm"] * mm
    font = cfg["font"]
    jp = register_jp_font(cfg)

    def X(px): return px * sx + ox
    def Y(py): return (SCAN_H - py) * sy + oy

    def text(px, py, s, size=11, center=False):
        c.setFont(jp if (jp and not s.isascii()) else font, size)
        (c.drawCentredString if center else c.drawString)(X(px), Y(py), s)

    for q in qsos:
        if cfg["template_image"] and Path(cfg["template_image"]).exists():
            c.drawImage(cfg["template_image"], 0, 0, PAGE_W_MM * mm, PAGE_H_MM * mm)
        c.setFillColorRGB(0, 0, 0)

        call = q["call"].upper()
        base, suffix = split_call(call)
        if len(base) <= 6:   # 赤枠に1文字ずつ(/Pなどを除いた本体)
            for i, ch in enumerate(base):
                text(POS["call_box_x"][i], POS["call_box_y"] + 14, ch, 20, center=True)
        else:                # 7文字以上は枠の上に小さく
            text(POS["call_box_x"][0] - 25, POS["call_box_y"] + 10, base, 18)
        if suffix:           # 移動局: TO RADIO下の点線欄に「コール/P」
            text(POS["to_line_x"], POS["to_line_y"], call, 14)

        date, hhmm, tz = fmt_datetime(q, cfg["source_tz"], cfg["print_tz"])
        text(POS["date_x"], POS["row_y"], date, 11)
        text(POS["time_x"], POS["row_y"], f"{hhmm} {tz}", 11)
        text(POS["rst_x"], POS["row_y"], q.get("rst_sent", ""), 12)
        text(POS["mhz_x"], POS["row_y"], mhz_of(q), 11)
        text(POS["twoway_x"], POS["row_y"], q.get("mode", "").upper(), 11, center=True)

        key = qso_key(q)
        rec = history["printed"].get(key)
        if rec is None and assign_numbers:
            rec = {"qsl_no": history["next_no"], "call": call,
                   "printed_at": datetime.now().isoformat(timespec="seconds")}
            history["next_no"] += 1
            history["printed"][key] = rec
        if rec:
            text(POS["qslno_x"], POS["qslno_y"], f"{rec['qsl_no']:05d}", 11)

        text(POS["rig_x"], POS["rig_y"], cfg["rig"], 11)
        text(POS["out_x"], POS["rig_y"], cfg["output_w"], 11)
        text(POS["ant_x"], POS["ant_y"], ant_of(q, cfg), 11)

        text(POS["rmks_x"], POS["rmks_y"], cfg["rmks_line1"], 11)
        text(POS["rmks_x"], POS["rmks_line2_y"], cfg["rmks_line2"], 11)

        op = cfg["op_name"] or q.get("operator", "")
        text(POS["free_x"], POS["free_y1"], f"OP.  {op}", 10)
        text(POS["free_x"], POS["free_y1"] + POS["free_dy"], f"QTH. {cfg['qth']}", 9)
        c.showPage()
    c.save()
    return history


# ------------------------------------------------------------ CLI ---------
def run_cli(adif_path):
    cfg, hist = load_config(), load_history()
    qsos = [q for q in parse_adif(read_text(adif_path)) if qso_key(q) not in hist["printed"]]
    print(f"未印刷: {len(qsos)} 件")
    if not qsos:
        return
    out = APP_DIR / f"qsl_{datetime.now():%Y%m%d_%H%M%S}.pdf"
    save_history(render_pdf(qsos, out, cfg, hist))
    print("出力:", out)


# ------------------------------------------------------------ GUI ---------
def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    cfg = load_config()
    root = tk.Tk()
    root.title("QSLカード レポート面 印刷ツール")
    state = {"qsos": []}

    top = ttk.Frame(root, padding=6)
    top.pack(fill="x")
    ttk.Button(top, text="ADIFを開く", command=lambda: load()).pack(side="left")
    ttk.Button(top, text="設定", command=lambda: settings()).pack(side="left", padx=6)
    ttk.Button(top, text="PDF作成", command=lambda: make()).pack(side="right")
    info = tk.StringVar()
    ttk.Label(top, textvariable=info).pack(side="left", padx=10)

    cols = ("state", "call", "date", "time", "band", "mode", "rst", "no")
    tree = ttk.Treeview(root, columns=cols, show="headings", selectmode="extended", height=18)
    for col, w, label in zip(cols, (70, 100, 100, 70, 60, 60, 50, 140),
                             ("状態", "CALL", "日付", "時刻", "BAND", "MODE", "RST", "相手No.")):
        tree.heading(col, text=label)
        tree.column(col, width=w, anchor="w")
    tree.pack(fill="both", expand=True, padx=6, pady=(0, 6))

    def load():
        p = filedialog.askopenfilename(filetypes=[("ADIF", "*.adi *.adif *.ADI *.ADIF"), ("all", "*.*")])
        if not p:
            return
        state["qsos"] = parse_adif(read_text(p))
        refresh()

    def refresh():
        hist = load_history()
        tree.delete(*tree.get_children())
        todo = []
        for i, q in enumerate(state["qsos"]):
            done = qso_key(q) in hist["printed"]
            d, t, tz = fmt_datetime(q, cfg["source_tz"], cfg["print_tz"])
            tree.insert("", "end", iid=str(i), values=(
                "印刷済" if done else "未印刷", q["call"], d, f"{t} {tz}", q.get("band", ""),
                q.get("mode", ""), q.get("rst_sent", ""), q.get("srx", "")))
            if not done:
                todo.append(str(i))
        tree.selection_set(todo)  # 未印刷だけ選択状態にする(=重複防止)
        info.set(f"{len(state['qsos'])}件中 未印刷 {len(todo)}件")

    def make():
        sel = [state["qsos"][int(i)] for i in tree.selection()]
        if not sel:
            return messagebox.showinfo("QSL", "対象がありません")
        hist = load_history()
        again = [q for q in sel if qso_key(q) in hist["printed"]]
        if again and not messagebox.askyesno("再印刷", f"印刷済みが{len(again)}件含まれます。再印刷しますか?"):
            return
        out = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=f"qsl_{datetime.now():%Y%m%d_%H%M}.pdf")
        if not out:
            return
        save_history(render_pdf(sel, out, cfg, hist))
        refresh()
        messagebox.showinfo("QSL", f"{len(sel)}枚分のPDFを作成しました")

    def settings():
        win = tk.Toplevel(root)
        win.title("設定")
        labels = [("rig", "Rig"), ("output_w", "Output (W)"), ("ant", "Ant(自動判定外のバンド用)"), ("qth", "QTH"),
                  ("rmks_line1", "Rmks 1行目"), ("rmks_line2", "Rmks 2行目"),
                  ("jp_font_path", "日本語フォント(.ttf/.ttc 空=自動)"),
                  ("op_name", "OP名(空=ADIFのOPERATOR)"), ("source_tz", "ADIFの時刻 (JST/UTC)"),
                  ("print_tz", "印字する時刻 (JST/UTC)"), ("template_image", "位置確認用の台紙画像"),
                  ("offset_x_mm", "X補正(mm)"), ("offset_y_mm", "Y補正(mm)")]
        vars_ = {}
        for r, (k, lab) in enumerate(labels):
            ttk.Label(win, text=lab).grid(row=r, column=0, sticky="w", padx=6, pady=2)
            vars_[k] = tk.StringVar(value=str(cfg[k]))
            ttk.Entry(win, textvariable=vars_[k], width=36).grid(row=r, column=1, padx=6)

        def ok():
            for k, v in vars_.items():
                cfg[k] = float(v.get() or 0) if k.startswith("offset") else v.get().strip()
            save_config(cfg)
            win.destroy()
            refresh()
        ttk.Button(win, text="保存", command=ok).grid(row=len(labels), column=1, sticky="e", pady=6)

    root.mainloop()


if __name__ == "__main__":
    run_cli(sys.argv[1]) if len(sys.argv) > 1 else run_gui()
