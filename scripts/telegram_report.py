#!/usr/bin/env python3
"""Отчёт для Telegram: выполнение по РДТ и ШПД ниже графика (2 таблицы + Excel).

Берёт тот же data.json, что и сайт, считает по тем же правилам и отправляет в Telegram
несколько сообщений + Excel-файл:
  1. Выполнение: план / факт / % по всем показателям — по регионам (РДТ) и ДРБ (одной строкой).
  2. ШПД ниже графика: сколько сотрудников не дотягивают до плана на сегодня
     (с учётом опоздания данных в Qlik на LAG_DAYS дня и допуска TOLERANCE), по регионам и каналам,
     с разбивкой по % выполнения плана: 0%, 1–20%, 21–40%, 41–60%, 61–80%, больше 80%.

Запуск:
    python scripts/telegram_report.py                       # только собрать файлы (out/)
    TELEGRAM_BOT_TOKEN=... TELEGRAM_CHAT_ID=... python scripts/telegram_report.py --send
    TELEGRAM_BOT_TOKEN=... python scripts/telegram_report.py --chat-ids   # узнать chat_id
"""
import argparse, calendar, html, json, os, re, sys, urllib.request, uuid
from pathlib import Path

IND = [("tg", "Продажи ₸", "₸"), ("shpd", "ШПД", "ШПД"), ("fms", "ФМС", "ФМС"), ("tv", "ТВ", "ТВ")]
EXCL_REG = re.compile(r"^(-|)$|дирекция", re.I)
EXCL_CHAN = re.compile(r"online churn", re.I)
CH_ORDER = ["ПСС", "УП", "ЦАП", "ГП", "СОВО", "Online Sales"]
DRB = "ДРБ"
MONTHS = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август",
          "сентябрь", "октябрь", "ноябрь", "декабрь"]
PACE_LIMIT = 0.8
LAG_DAYS = int(os.environ.get("QLIK_LAG_DAYS", "3"))         # данные в Qlik опаздывают на 3 дня
TOLERANCE = float(os.environ.get("SHPD_TOLERANCE", "0.10"))   # допуск: до 10 % ниже нужного факта — не отстаёт
BUCKETS = [("0%", 0, 0), ("1-20", 0, 0.20), ("21-40", 0.20, 0.40), ("41-60", 0.40, 0.60),
           ("61-80", 0.60, 0.80), (">80", 0.80, 9e9)]


def prev_p(p):
    y, m = divmod(p, 100)
    return (y - 1) * 100 + 12 if m == 1 else y * 100 + m - 1


def elapsed_share(date, period):
    y, m, d = map(int, date.split("-"))
    dp = y * 100 + m
    if dp > period:
        return 1.0
    if dp < period:
        return 0.0
    return (d - 1) / calendar.monthrange(y, m)[1]


def build(data):
    tg, vol, date = data["tg"], data["vol"], data["exportDate"]
    cur = max([r[0] for r in tg] + [r[0] for r in vol])
    m3 = [prev_p(prev_p(prev_p(cur))), prev_p(prev_p(cur)), prev_p(cur)]
    want = set(m3 + [cur])
    E = {}

    def emp(L):
        L = str(L).strip().upper()
        return E.setdefault(L, {"L": L, "fio": "", "reg": None, "chan": None, "fil": "", "np": "", "m": {}})

    def slot(e, p):
        return e["m"].setdefault(p, [None] * 8)

    def add(a, i, v):
        a[i] = (a[i] or 0) + v

    def ok(reg, ch):
        return not EXCL_REG.search(str(reg or "").strip()) and not EXCL_CHAN.search(str(ch or ""))

    for p, reg, fil, L, fio, ch, pl, fa in tg:
        if p not in want or not ok(reg, ch):
            continue
        e = emp(L); a = slot(e, p); add(a, 0, pl); add(a, 1, fa)
        if fio and len(str(fio)) > len(e["fio"]):
            e["fio"] = str(fio).strip()
        if p == cur:
            e["reg"], e["chan"] = str(reg).strip(), str(ch).strip()
            if fil: e["fil"] = str(fil).strip()
    for p, reg, fil, fio, L, ch, np_, sp, sf, fp, ff, tp, tf in vol:
        if p not in want or not ok(reg, ch):
            continue
        e = emp(L); a = slot(e, p)
        for i, v in zip(range(2, 8), (sp, sf, fp, ff, tp, tf)):
            add(a, i, v)
        if fio and len(str(fio)) > len(e["fio"]):
            e["fio"] = str(fio).strip()
        if p == cur:
            e["reg"], e["chan"] = str(reg).strip(), str(ch).strip()
            if fil: e["fil"] = str(fil).strip()
            if np_: e["np"] = str(np_).strip()

    share = elapsed_share(date, cur)
    emps = [e for e in E.values() if e["reg"] and cur in e["m"]
            and any((e["m"][cur][k] or 0) > 0 for k in (0, 2, 4, 6))]
    for e in emps:
        e["ind"] = []
        for i in range(4):
            pcs = []
            for p in m3:
                a = e["m"].get(p)
                if a and a[2 * i] is not None and a[2 * i] > 0:
                    pcs.append((a[2 * i + 1] or 0) / a[2 * i])
            avg = sum(pcs) / len(pcs) if pcs else None
            c = e["m"][cur]
            plan, fact = c[2 * i], c[2 * i + 1]
            pc = (fact or 0) / plan if plan and plan > 0 else None
            pace = pc / share if pc is not None and share > 0 else None
            r = avg if avg is not None else pace
            fc = None if pc is None else pc if share >= 1 else (pc + (1 - share) * r if r is not None else None)
            e["ind"].append({"avg": avg, "plan": plan, "fact": fact, "pc": pc, "fc": fc, "pace": pace,
                             "lag": fc is not None and fc < PACE_LIMIT,
                             "chronic": fc is not None and fc < PACE_LIMIT and avg is not None and avg < 0.8})
    return {"date": date, "cur": cur, "m3": m3, "share": share, "emps": emps}


def regions(emps):
    return sorted({e["reg"] for e in emps}, key=lambda r: (r == DRB, r))


def chans(emps):
    return sorted({e["chan"] for e in emps}, key=lambda c: (CH_ORDER.index(c) if c in CH_ORDER else 99, c))


def counts(lst):
    out = []
    for i in range(4):
        withp = [e for e in lst if e["ind"][i]["fc"] is not None]
        out.append((len(withp), sum(e["ind"][i]["lag"] for e in withp), sum(e["ind"][i]["chronic"] for e in withp)))
    return out


def ddmm(date):
    y, m, d = date.split("-")
    return f"{d}.{m}.{y}"


def fmt_num(v, money):
    if money:
        return f"{v / 1e6:,.1f}".replace(",", " ")
    return f"{v:,.0f}".replace(",", " ")


def groups(emps):
    """Строки таблицы выполнения: регионы, итого регионы, ДРБ с каналами, всего."""
    regs = [r for r in regions(emps) if r != DRB]
    out = [(r, [e for e in emps if e["reg"] == r]) for r in regs]
    out.append(("Итого рег.", [e for e in emps if e["reg"] != DRB]))
    drb = [e for e in emps if e["reg"] == DRB]
    if drb:
        out.append(("ДРБ", drb))
    out.append(("Всего", emps))
    return out


def totals(lst, i, cur):
    plan = sum((e["m"][cur][2 * i] or 0) for e in lst)
    fact = sum((e["m"][cur][2 * i + 1] or 0) for e in lst)
    return plan, fact, (fact / plan if plan else None)


def message_exec(M):
    """1. Выполнение: план / факт / % по всем показателям — по РДТ и ДРБ."""
    emps, cur = M["emps"], M["cur"]
    parts = [f"<b>Выполнение плана на {ddmm(M['date'])}</b>\n"
             f"{MONTHS[cur % 100 - 1]} {cur // 100}: прошло {round(M['share'] * 100)}% дней месяца.\n"]
    W = 13
    for i, (_, name, _) in enumerate(IND):
        money = i == 0
        rows = [f"{('млн ₸' if money else ''):<{W}}{'план':>8}{'факт':>8}{'%':>5}"]
        for g, lst in groups(emps):
            plan, fact, pc = totals(lst, i, cur)
            rows.append(f"{g[:W]:<{W}}{fmt_num(plan, money):>8}{fmt_num(fact, money):>8}"
                        f"{(str(round(pc * 100)) + '%') if pc is not None else '—':>5}")
            if g.startswith("Итого"):
                rows.append("")
        parts.append(f"\n<b>{name}</b>\n<pre>" + html.escape("\n".join(rows)) + "</pre>")
    return "".join(parts)


def shpd_need_share(M):
    """Какая доля плана ШПД должна быть в Qlik на дату выгрузки (с учётом опоздания данных)."""
    y, m, d = map(int, M["date"].split("-"))
    if y * 100 + m != M["cur"]:
        return 1.0
    return max(0, d - LAG_DAYS) / calendar.monthrange(y, m)[1]


def shpd_rows(M):
    """Сотрудники с планом ШПД: план, факт, нужно на дату, отстаёт ли, корзина % выполнения."""
    need = shpd_need_share(M)
    out = []
    for e in M["emps"]:
        c = e["m"][M["cur"]]
        plan, fact = c[2] or 0, c[3] or 0
        if plan <= 0:
            continue
        must = plan * need
        lag = fact < must * (1 - TOLERANCE)
        pc = fact / plan
        b = next(n for n, lo, hi in BUCKETS if (pc == 0 if n == "0%" else lo < pc <= hi))
        out.append({"e": e, "plan": plan, "fact": fact, "must": must, "pc": pc, "lag": lag, "bucket": b})
    return out, need


def message_shpd(M):
    """2. ШПД ниже графика — по каждому региону отдельно, по каналам, с разбивкой по % выполнения."""
    rows_all, need = shpd_rows(M)
    head = (f"<b>ШПД ниже графика на {ddmm(M['date'])}</b>\n"
            f"Данные в Qlik — с опозданием на {LAG_DAYS} дня, поэтому на дату нужно ≈{round(need * 100)}% плана. "
            f"Отстающий — факт ниже нужного больше чем на {round(TOLERANCE * 100)}% "
            f"(факт &lt; {round(need * (1 - TOLERANCE) * 100)}% плана).\n"
            f"Столбцы: сотрудников с планом ШПД, из них отстают, и отстающие по % выполнения плана.\n")
    W = 9
    limit = need * (1 - TOLERANCE)
    shown = [b for b in BUCKETS if b[1] < limit or b[0] == "0%"]      # группы выше порога всегда пустые
    hdr = f"{'':<{W}}{'чел':>4}{'отст':>5}" + "".join(f"{n:>6}" for n, _, _ in shown)
    short = {"Online Sales": "Online"}

    def line(name, lst):
        lag = [r for r in lst if r["lag"]]
        name = short.get(name.strip(), name.strip())
        name = (" " + name) if lst and name not in ("Итого", "Регионы", "Всего") else name
        return f"{name[:W]:<{W}}{len(lst):>4}{len(lag):>5}" + "".join(
            f"{sum(1 for r in lag if r['bucket'] == n):>6}" for n, _, _ in shown)
    blocks = []
    for reg in regions(M["emps"]):
        R = [r for r in rows_all if r["e"]["reg"] == reg]
        if not R:
            continue
        if reg == DRB:
            lines = [hdr, line("Итого", R)]                       # ДРБ — одной строкой, без каналов
        else:
            lines = [hdr] + [line(" " + c, [r for r in R if r["e"]["chan"] == c]) for c in chans([r["e"] for r in R])]
            lines.append(line("Итого", R))
        blocks.append(f"\n<b>{html.escape(reg)}</b>\n<pre>" + html.escape("\n".join(lines)) + "</pre>")
    reg_rows = [r for r in rows_all if r["e"]["reg"] != DRB]
    lines = [hdr, line("Регионы", reg_rows), line("Всего", rows_all)]
    blocks.append("\n<b>Итого</b>\n<pre>" + html.escape("\n".join(lines)) + "</pre>")
    return head + "".join(blocks)


def message(M):
    emps = M["emps"]
    head = (f"<b>Отстающие на {ddmm(M['date'])}</b>\n"
            f"{MONTHS[M['cur'] % 100 - 1]} {M['cur'] // 100}: прошло {round(M['share'] * 100)}% дней месяца.\n"
            f"Отстающий — прогноз на конец месяца ниже 80% (с учётом выполнения за 3 прошлых месяца).\n")
    m3 = ", ".join(MONTHS[p % 100 - 1][:3] for p in M["m3"])
    W = 12

    def block(k):
        rows = [f"{'':<{W}}{'чел':>4}" + "".join(f"{s:>5}" for _, _, s in IND)]

        def line(name, lst):
            return f"{name[:W]:<{W}}{len(lst):>4}" + "".join(f"{c[k]:>5}" for c in counts(lst))
        reg_list = [e for e in emps if e["reg"] != DRB]
        for r in regions(emps):
            if r == DRB:
                rows += [line("Итого рег.", reg_list), ""]
            LR = [e for e in emps if e["reg"] == r]
            rows.append(line(r, LR))
            if r != DRB:
                for c in chans(LR):
                    rows.append(line(" " + c, [e for e in LR if e["chan"] == c]))
        if not any(e["reg"] == DRB for e in emps):
            rows.append(line("Итого рег.", reg_list))
        rows.append(line("Всего", emps))
        return "<pre>" + html.escape("\n".join(rows)) + "</pre>"
    return (head + "\n<b>1. Прогноз ниже 80%, чел.</b>\n" + block(1) +
            f"\n<b>2. Из них хронически</b> (ниже 80% и в среднем за {m3}), чел.\n" + block(2))


def build_xlsx(M, path):
    """Excel: «Выполнение», «ШПД ниже графика» (сводка по градациям) и «Сотрудники» (подробно, по ссылке)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    red = PatternFill("solid", fgColor="FBE4E4"); amber = PatternFill("solid", fgColor="FDF1D3"); green = PatternFill("solid", fgColor="E1F3E1")
    total_fill = PatternFill("solid", fgColor="A9D8A9")        # строки «Итого» — зелёные
    head_fill = PatternFill("solid", fgColor="F0EFEC")
    bold = Font(bold=True); link_font = Font(color="0563C1", underline="single")
    thin = Side(style="thin", color="D0D0D0"); box = Border(top=thin, bottom=thin, left=thin, right=thin)
    fill_for = lambda v: None if v is None else red if v < 0.8 else amber if v < 1 else green
    emps, cur = M["emps"], M["cur"]
    wb = Workbook()

    def header(ws, row):
        for c in ws[row]:
            c.font = bold; c.fill = head_fill; c.border = box
            c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

    def total_row(ws, r):
        for c in ws[r]:
            c.font = bold; c.fill = total_fill

    def link(ws, cell, text, target):
        ws[cell] = text
        ws[cell].hyperlink = f"#'{target}'!A1"
        ws[cell].font = link_font

    # ---------- 1. Выполнение ----------
    ws = wb.active; ws.title = "Выполнение"
    ws["A1"] = f"Выполнение плана на {ddmm(M['date'])} — по регионам (РДТ) и ДРБ"; ws["A1"].font = Font(bold=True, size=13)
    link(ws, "A2", "Подробно по сотрудникам →", "Сотрудники")
    ws.append([])
    ws.append(["РДТ"] + sum([[f"{n}: план", f"{n}: факт", f"{n}: %"] for _, n, _ in IND], []))
    header(ws, 4)
    for g, lst in groups(emps):
        row = [g.strip()]
        for i in range(4):
            plan, fact, pc = totals(lst, i, cur)
            row += [round(plan), round(fact), round(pc, 2) if pc is not None else None]
        ws.append(row); r = ws.max_row
        is_total = g.startswith("Итого") or g.startswith("Всего")
        for i in range(4):
            ws.cell(r, 2 + 3 * i).number_format = ws.cell(r, 3 + 3 * i).number_format = "#,##0"
            ws.cell(r, 4 + 3 * i).number_format = "0%"
            f = fill_for(row[3 + 3 * i])
            if f and not is_total:
                ws.cell(r, 4 + 3 * i).fill = f
        for c in ws[r]: c.border = box
        if is_total:
            total_row(ws, r)
    ws.column_dimensions["A"].width = 16
    for k in range(2, 14): ws.column_dimensions[get_column_letter(k)].width = 13
    ws.freeze_panes = "B5"

    # ---------- 2. ШПД ниже графика (сводка по градациям) ----------
    rows_all, need = shpd_rows(M)
    limit = need * (1 - TOLERANCE)
    ws2 = wb.create_sheet("ШПД ниже графика")
    ws2["A1"] = f"ШПД ниже графика на {ddmm(M['date'])}"; ws2["A1"].font = Font(bold=True, size=13)
    ws2["A2"] = (f"Нужно на дату ≈{round(need * 100)}% плана (данные Qlik с опозданием {LAG_DAYS} дня). "
                 f"Отстающий — факт ниже нужного больше чем на {round(TOLERANCE * 100)}% (факт < {round(limit * 100)}% плана). "
                 f"Градации — % выполнения плана ШПД у отстающих.")
    link(ws2, "A3", "Подробно по сотрудникам →", "Сотрудники")
    ws2.append([])
    bnames = {"0%": "0%", "1-20": "1–20%", "21-40": "21–40%", "41-60": "41–60%", "61-80": "61–80%", ">80": "больше 80%"}
    shown = [b for b in BUCKETS if b[1] < limit or b[0] == "0%"]      # группы выше порога всегда пустые
    ws2.append(["РДТ", "Канал", "С планом ШПД, чел", "Отстают, чел", "% отстающих"] +
               [bnames[n] for n, _, _ in shown] + ["Отставание, шт"])
    header(ws2, 5)

    def put(reg, ch, lst, is_total=False):
        lag = [x for x in lst if x["lag"]]
        row = [reg, ch, len(lst), len(lag), round(len(lag) / len(lst), 2) if lst else None]
        row += [sum(1 for x in lag if x["bucket"] == n) for n, _, _ in shown]
        row.append(round(sum(x["must"] - x["fact"] for x in lag)))
        ws2.append(row); r = ws2.max_row
        ws2.cell(r, 5).number_format = "0%"
        ws2.cell(r, len(row)).number_format = "#,##0"
        for c in ws2[r]: c.border = box
        if is_total:
            total_row(ws2, r)
    for reg in regions(emps):
        R = [x for x in rows_all if x["e"]["reg"] == reg]
        if not R or reg == DRB:
            continue
        for c in chans([x["e"] for x in R]):
            put(reg, c, [x for x in R if x["e"]["chan"] == c])
        put(f"Итого {reg}", "", R, True)
    put("Итого регионы", "", [x for x in rows_all if x["e"]["reg"] != DRB], True)
    drb = [x for x in rows_all if x["e"]["reg"] == DRB]
    if drb:
        put("ДРБ", "", drb)
    put("Всего", "", rows_all, True)
    for k, w in enumerate([20, 10, 12, 11, 11] + [10] * len(shown) + [13], 1):
        ws2.column_dimensions[get_column_letter(k)].width = w
    ws2.freeze_panes = "C6"

    # ---------- 3. Сотрудники (подробно, открывается по ссылке) ----------
    ws3 = wb.create_sheet("Сотрудники")
    ws3["A1"] = f"Выполнение по сотрудникам на {ddmm(M['date'])}"; ws3["A1"].font = Font(bold=True, size=13)
    link(ws3, "A2", "← К выполнению", "Выполнение")
    ws3.append([])
    ws3.append(["РДТ", "Канал", "Подразделение", "ФИО", "Логин"] +
               sum([[f"{n}: план", f"{n}: факт", f"{n}: %"] for _, n, _ in IND], []))
    header(ws3, 4)
    order = sorted(emps, key=lambda e: ((e["reg"] == DRB), e["reg"], CH_ORDER.index(e["chan"]) if e["chan"] in CH_ORDER else 99,
                                        e["fio"] or e["L"]))
    for e in order:
        c = e["m"][cur]
        row = [e["reg"], e["chan"], e["np"] or e["fil"], e["fio"] or e["L"], e["L"]]
        for i in range(4):
            plan, fact = c[2 * i], c[2 * i + 1]
            pc = (fact or 0) / plan if plan else None
            row += [round(plan) if plan is not None else None, round(fact) if fact is not None else None,
                    round(pc, 2) if pc is not None else None]
        ws3.append(row); r = ws3.max_row
        for i in range(4):
            ws3.cell(r, 6 + 3 * i).number_format = ws3.cell(r, 7 + 3 * i).number_format = "#,##0"
            ws3.cell(r, 8 + 3 * i).number_format = "0%"
            f = fill_for(row[7 + 3 * i])
            if f: ws3.cell(r, 8 + 3 * i).fill = f
    for k, w in enumerate([14, 12, 22, 30, 18] + [11] * 12, 1): ws3.column_dimensions[get_column_letter(k)].width = w
    ws3.freeze_panes = "F5"; ws3.auto_filter.ref = f"A4:{get_column_letter(ws3.max_column)}{ws3.max_row}"
    wb.active = 0
    wb.save(path)


def tg_call(token, method, fields, file_field=None, file_path=None):
    url = f"https://api.telegram.org/bot{token}/{method}"
    if file_path is None:
        req = urllib.request.Request(url, data=json.dumps(fields).encode(), headers={"Content-Type": "application/json"})
    else:
        b = uuid.uuid4().hex; parts = []
        for k, v in fields.items():
            parts.append(f'--{b}\r\nContent-Disposition: form-data; name="{k}"\r\n\r\n{v}\r\n'.encode())
        name = Path(file_path).name
        parts.append(f'--{b}\r\nContent-Disposition: form-data; name="{file_field}"; filename="{name}"\r\n'
                     f'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n\r\n'.encode())
        parts.append(Path(file_path).read_bytes()); parts.append(f"\r\n--{b}--\r\n".encode())
        req = urllib.request.Request(url, data=b"".join(parts), headers={"Content-Type": f"multipart/form-data; boundary={b}"})
    with urllib.request.urlopen(req, timeout=60) as r:
        res = json.load(r)
    if not res.get("ok"):
        raise RuntimeError(f"Telegram {method}: {res}")
    return res


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", default="daily/data.json", help="путь к data.json сайта")
    ap.add_argument("--out", default="out")
    ap.add_argument("--send", action="store_true", help="отправить в Telegram")
    ap.add_argument("--chat-ids", action="store_true", help="показать chat_id из последних сообщений боту")
    a = ap.parse_args()
    token = os.environ.get("TELEGRAM_BOT_TOKEN", "").strip()
    if a.chat_ids:
        res = tg_call(token, "getUpdates", {})
        for u in res["result"]:
            ch = (u.get("message") or u.get("channel_post") or u.get("my_chat_member") or {}).get("chat", {})
            if ch: print(ch.get("id"), ch.get("type"), ch.get("title") or ch.get("username") or ch.get("first_name"))
        return
    M = build(json.loads(Path(a.data).read_text(encoding="utf-8")))
    out = Path(a.out); out.mkdir(exist_ok=True)
    msgs = [message_exec(M), message_shpd(M)]
    for k, m in enumerate(msgs, 1):
        if len(m) > 4000:
            raise SystemExit(f"Сообщение {k} длиннее лимита Telegram: {len(m)} символов")
    (out / "message.html").write_text("\n\n---\n\n".join(msgs), encoding="utf-8")
    xlsx = out / f"otchet_{M['date']}.xlsx"; build_xlsx(M, xlsx)
    print("\n\n---\n\n".join(msgs)); print(f"\nФайл: {xlsx}")
    if a.send:
        chat_ids = [c.strip() for c in os.environ.get("TELEGRAM_CHAT_ID", "").split(",") if c.strip()]
        if not token or not chat_ids:
            sys.exit("Нужны переменные TELEGRAM_BOT_TOKEN и TELEGRAM_CHAT_ID")
        site = os.environ.get("SITE_URL", "").strip()
        for cid in chat_ids:
            for k, m in enumerate(msgs):
                tg_call(token, "sendMessage", {"chat_id": cid, "text": m + (f"\n\nСайт: {site}" if site and k == len(msgs) - 1 else ""),
                                               "parse_mode": "HTML", "disable_web_page_preview": True})
            tg_call(token, "sendDocument", {"chat_id": cid, "caption": "Выполнение и ШПД ниже графика"},
                    "document", xlsx)
        print("Отправлено в чаты:", ", ".join(chat_ids))


if __name__ == "__main__":
    main()
