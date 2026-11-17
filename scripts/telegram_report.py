#!/usr/bin/env python3
"""Отчёт «Отстающие по темпу» для Telegram-бота.

Берёт тот же daily/data.json, что и сайт, считает по тем же правилам и
отправляет в Telegram сводку по регионам и каналам + Excel-файл со списком.

Отстающий по темпу — сотрудник, у которого прогноз на месяц ниже 80%:
    прогноз = (факт / план) / доля прошедших дней месяца.
Для каждого отстающего указано, был ли он ниже 80% и в среднем за 3 прошлых
полных месяца (тогда это «хронически отстающий»).

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
    return (d - 1) / calendar.monthrange(y, m)[1]   # выгрузка от 17-го = 16 полных дней


def build(data):
    """Та же логика, что buildModel + enrich на сайте."""
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
    emps = [e for e in E.values() if e["reg"] and cur in e["m"]]
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
            pc = fact / plan if plan and plan > 0 else None
            fc = pc / share if pc is not None and share > 0 else None
            e["ind"].append({"avg": avg, "plan": plan, "fact": fact, "pc": pc, "fc": fc,
                             "lag": fc is not None and fc < PACE_LIMIT,
                             "chronic": fc is not None and fc < PACE_LIMIT and avg is not None and avg < 0.8})
    return {"date": date, "cur": cur, "m3": m3, "share": share, "emps": emps}


def regions(emps):
    return sorted({e["reg"] for e in emps}, key=lambda r: (r == DRB, r))


def chans(emps):
    return sorted({e["chan"] for e in emps}, key=lambda c: (CH_ORDER.index(c) if c in CH_ORDER else 99, c))


def counts(lst):
    """[(с планом, отстают, хронически)] по 4 показателям."""
    out = []
    for i in range(4):
        withp = [e for e in lst if e["ind"][i]["fc"] is not None]
        out.append((len(withp), sum(e["ind"][i]["lag"] for e in withp), sum(e["ind"][i]["chronic"] for e in withp)))
    return out


def ddmm(date):
    y, m, d = date.split("-")
    return f"{d}.{m}.{y}"


def message(M):
    emps = M["emps"]
    head = (f"<b>Отстающие по темпу на {ddmm(M['date'])}</b>\n"
            f"{MONTHS[M['cur'] % 100 - 1]} {M['cur'] // 100}: прошло {round(M['share'] * 100)}% дней месяца.\n"
            f"Отстающий — прогноз на месяц ниже 80%.\n")
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
            for c in chans(LR):
                rows.append(line(" " + c, [e for e in LR if e["chan"] == c]))
        if not any(e["reg"] == DRB for e in emps):
            rows.append(line("Итого рег.", reg_list))
        rows.append(line("Всего", emps))
        return "<pre>" + html.escape("\n".join(rows)) + "</pre>"
    return (head + "\n<b>1. Отстают по темпу, чел.</b>\n" + block(1) +
            f"\n<b>2. Из них хронически</b> (ниже 80% и в среднем за {m3}), чел.\n" + block(2))


def build_xlsx(M, path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    red = PatternFill("solid", fgColor="FBE4E4"); amber = PatternFill("solid", fgColor="FDF1D3"); green = PatternFill("solid", fgColor="E1F3E1")
    bold = Font(bold=True); grey = PatternFill("solid", fgColor="F0EFEC")
    fill_for = lambda v: None if v is None else red if v < 0.8 else amber if v < 1 else green
    emps = M["emps"]; wb = Workbook()

    ws = wb.active; ws.title = "Сводка"
    ws.append([f"Отстающие по темпу на {ddmm(M['date'])} (прогноз на месяц ниже 80%)"]); ws["A1"].font = Font(bold=True, size=13)
    ws.append([f"Прошло {round(M['share'] * 100)}% дней месяца. «Хронически» — отстают по темпу и ниже 80% в среднем за 3 прошлых месяца."])
    ws.append([])
    ws.append(["Регион", "Канал", "Сотрудников"] + sum([[f"{n}: с планом", f"{n}: отстают", f"{n}: % отстающих", f"{n}: из них хронически"] for _, n, _ in IND], []))
    for c in ws[4]: c.font = bold; c.alignment = Alignment(wrap_text=True, vertical="top")

    def put(reg, ch, lst, is_total=False):
        row = [reg, ch, len(lst)]
        for withp, lag, chron in counts(lst):
            row += [withp, lag, (lag / withp) if withp else None, chron]
        ws.append(row)
        r = ws.max_row
        for k in range(4):
            ws.cell(r, 6 + 4 * k).number_format = "0%"
        if is_total or not ch:
            for c in ws[r]: c.font = bold; c.fill = grey
    for reg in regions(emps):
        LR = [e for e in emps if e["reg"] == reg]
        put(reg + (" (не регион)" if reg == DRB else ""), "", LR)
        for c in chans(LR):
            put(reg, c, [e for e in LR if e["chan"] == c])
    put("Итого по регионам", "", [e for e in emps if e["reg"] != DRB], True)
    put("Всего с ДРБ", "", emps, True)
    ws.column_dimensions["A"].width = 22; ws.column_dimensions["B"].width = 14
    for k in range(3, 3 + 1 + 16): ws.column_dimensions[get_column_letter(k)].width = 12
    ws.freeze_panes = "C5"

    ws2 = wb.create_sheet("Сотрудники")
    ws2.append(["Регион", "Канал", "Подразделение", "ФИО", "Логин", "Отстаёт по показателям"] +
               sum([[f"{n}: план", f"{n}: факт", f"{n}: % на дату", f"{n}: прогноз", f"{n}: ср. 3 мес."] for _, n, _ in IND], []))
    for c in ws2[1]: c.font = bold; c.alignment = Alignment(wrap_text=True, vertical="top")
    lag_emps = [e for e in emps if any(x["lag"] for x in e["ind"])]
    lag_emps.sort(key=lambda e: ((e["reg"] == DRB), e["reg"], CH_ORDER.index(e["chan"]) if e["chan"] in CH_ORDER else 99, -sum(x["lag"] for x in e["ind"]), e["fio"]))
    for e in lag_emps:
        row = [e["reg"], e["chan"], e["np"] or e["fil"], e["fio"] or e["L"], e["L"], ", ".join(IND[i][2] for i, x in enumerate(e["ind"]) if x["lag"])]
        for x in e["ind"]:
            row += [x["plan"], x["fact"], x["pc"], x["fc"], x["avg"]]
        ws2.append(row)
        r = ws2.max_row
        for k, x in enumerate(e["ind"]):
            base = 7 + 5 * k
            ws2.cell(r, base).number_format = ws2.cell(r, base + 1).number_format = "#,##0"
            for j, v in ((2, x["pc"]), (3, x["fc"]), (4, x["avg"])):
                cell = ws2.cell(r, base + j); cell.number_format = "0%"
            for j, v in ((3, x["fc"]), (4, x["avg"])):
                f = fill_for(v)
                if f: ws2.cell(r, base + j).fill = f
    widths = [14, 12, 22, 30, 18, 18] + [11] * 20
    for k, w in enumerate(widths, 1): ws2.column_dimensions[get_column_letter(k)].width = w
    ws2.freeze_panes = "E2"; ws2.auto_filter.ref = ws2.dimensions
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
    msg = message(M); (out / "message.html").write_text(msg, encoding="utf-8")
    xlsx = out / f"otstayushchie_po_tempu_{M['date']}.xlsx"; build_xlsx(M, xlsx)
    print(msg); print(f"\nФайл: {xlsx}")
    if a.send:
        chat_ids = [c.strip() for c in os.environ.get("TELEGRAM_CHAT_ID", "").split(",") if c.strip()]
        if not token or not chat_ids:
            sys.exit("Нужны переменные TELEGRAM_BOT_TOKEN и TELEGRAM_CHAT_ID")
        site = os.environ.get("SITE_URL", "").strip()
        for cid in chat_ids:
            tg_call(token, "sendMessage", {"chat_id": cid, "text": msg + (f"\n\nСайт: {site}" if site else ""),
                                           "parse_mode": "HTML", "disable_web_page_preview": True})
            tg_call(token, "sendDocument", {"chat_id": cid, "caption": "Сводка и список отстающих"}, "document", xlsx)
        print("Отправлено в чаты:", ", ".join(chat_ids))


if __name__ == "__main__":
    main()
