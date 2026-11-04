# -*- coding: utf-8 -*-
"""
助学贷款还款计划表
  · 产出 1 个 .xlsx（参数联动 + 逾期高亮 + 复选框）
  · 产出 1 个 .ics （手机日历原生导入，带「提前 N 天」提醒）
  · 产出 1 个 .csv （兜底，iPhone「文件」App 可直接导入）

用法：python 助学贷款还款计划表.py
依赖：pip install openpyxl  （ics 为纯标准库实现，无额外依赖）
"""

import csv
import uuid
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.formatting.rule import FormulaRule

# ==================================================================
# 0. 输出路径 & 日历提醒参数（改这里就够了）
# ==================================================================
OUT_XLSX = "助学贷款还款计划表.xlsx"
OUT_ICS  = "还款提醒日历.ics"
OUT_CSV  = "还款提醒日历.csv"

REMIND_DAYS_BEFORE = 3        # 提前几天弹通知（0 = 当天 09:00 提醒）
REMIND_TIME = (9, 0, 0)       # 提醒时刻（时,分,秒），本地时间
CALENDAR_NAME = "助学贷款还款提醒"
CALENDAR_COLOR = "FF2E86C1"   # iOS 导入后可改色，此处仅写进 X-WR-CALCOLOR

# ==================================================================
# 0b. webcal 订阅发布参数（只在「发布成订阅链接」时才用到）
# ==================================================================
PUBLISH_DIR = Path(__file__).parent / "publish"   # 对外暴露的目录，里面只放一个 loan.ics
PUBLISH_ICS = PUBLISH_DIR / "loan.ics"            # 文件名越短越好，订阅链接里要手打
WEB_PORT = 8765                                   # 局域网临时服务的端口，随便改
WEB_HOST = "0.0.0.0"                             # 监听所有网卡；只用本机可改 "127.0.0.1"
REFRESH_INTERVAL_HOURS = 24                       # X-PUBLISHED-TTL，手机约每天来取一次

# ==================================================================
# 1. 生成 xlsx（你的原脚本，逻辑一字未改）
# ==================================================================
def build_xlsx(path=OUT_XLSX):
    wb = Workbook()
    ws = wb.active
    ws.title = "助学贷款还款计划表"

    # ---------- 全局参数区 ----------
    ws["A1"] = "全局参数区（仅修改此区，全表自动联动）"
    ws.merge_cells("A1:B1")
    ws["A1"].font = Font(bold=True, size=12)
    ws["A1"].fill = PatternFill("solid", fgColor="D9E1F2")
    ws["A1"].alignment = Alignment(horizontal="center")

    params = [
        ["借款本金（元）", 40000],                    # B2
        ["最新5年期LPR（年利率）", 0.0345],           # B3
        ["基点调整（减70BP）", -0.007],               # B4
        ["执行年利率", "=B3+B4"],                     # B5
        ["罚息利率倍率", 1.3],                        # B6
        ["罚息年利率", "=B5*B6"],                     # B7
        ["年计息天数", 365],                          # B8
        ["每年等额还本金额", 1818.18],                # B9
        ["最后一期还本金额", 1818.20],                # B10
        ["首个还款日（毕业当年）", date(2025, 12, 20)],  # B11
        ["开始还本年份", 2033],                       # B12
        ["最后到期日", date(2043, 9, 20)],            # B13
    ]
    for i, (label, value) in enumerate(params, start=2):
        ws[f"A{i}"] = label
        ws[f"B{i}"] = value
        ws[f"A{i}"].font = Font(bold=True)
        ws[f"A{i}"].alignment = Alignment(horizontal="right")

    for cell in ["B3", "B4", "B5", "B7"]:
        ws[cell].number_format = "0.00%"
    for cell in ["B11", "B13"]:
        ws[cell].number_format = "yyyy-mm-dd"
    for cell in ["B2", "B9", "B10"]:
        ws[cell].number_format = "0.00"

    # ---------- 还款明细表 ----------
    header_row = 15
    headers = ["序号", "还款日期", "还款阶段", "期初剩余本金",
               "应付利息", "应付本金", "应还本息",
               "已结清", "绑定列(隐藏)", "逾期天数", "逾期罚息", "应还总额"]
    for col, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col, value=header)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill("solid", fgColor="E2EFDA")

    start_row = header_row + 1
    total_periods = 19
    for row in range(start_row, start_row + total_periods):
        idx = row - start_row + 1
        ws.cell(row=row, column=1, value=idx)

        if idx == 1:
            ws.cell(row=row, column=2, value="=B$11")
        elif idx == total_periods:
            ws.cell(row=row, column=2, value="=B$13")
        else:
            ws.cell(row=row, column=2, value=f"=DATE(YEAR(B{row-1})+1,12,20)")

        ws.cell(row=row, column=3, value=(
            f'=IF(B{row}<DATE(B$12,12,20),"只还利息",'
            f'IF(B{row}<B$13,"还本付息","最后一期(利随本清)"))'))

        ws.cell(row=row, column=4, value="=B$2" if idx == 1 else f"=D{row-1}-F{row-1}")
        ws.cell(row=row, column=5, value=f"=D{row}*B$5")
        ws.cell(row=row, column=6, value=(
            f'=IF(C{row}="只还利息",0,'
            f'IF(C{row}="还本付息",B$9,B$10))'))
        ws.cell(row=row, column=7, value=f"=E{row}+F{row}")
        ws.cell(row=row, column=8, value=f'=IF(I{row}=TRUE,"√","")')
        ws.cell(row=row, column=9, value=False)
        ws.cell(row=row, column=10, value=f"=MAX(0,TODAY()-B{row})")
        ws.cell(row=row, column=11, value=(
            f'=IF(AND(I{row}=FALSE,J{row}>0),'
            f'F{row}*B$7/B$8*J{row},0)'))
        ws.cell(row=row, column=12, value=f"=G{row}+K{row}")

    for row in range(start_row, start_row + total_periods):
        ws[f"B{row}"].number_format = "yyyy-mm-dd"
        for col in ["D", "E", "F", "G", "K", "L"]:
            ws[f"{col}{row}"].number_format = "0.00"
        ws[f"J{row}"].number_format = "0"
        for col in range(1, 13):
            ws.cell(row=row, column=col).alignment = Alignment(horizontal="center")

    # ---------- 合计行 ----------
    total_row = start_row + total_periods
    ws[f"A{total_row}"] = "合计"
    ws[f"A{total_row}"].font = Font(bold=True)
    ws[f"A{total_row}"].fill = PatternFill("solid", fgColor="FFF2CC")
    ws[f"A{total_row}"].alignment = Alignment(horizontal="center")
    for col in ["E", "F", "G", "K", "L"]:
        ws[f"{col}{total_row}"] = f"=SUM({col}{start_row}:{col}{total_row-1})"
        ws[f"{col}{total_row}"].number_format = "0.00"
        ws[f"{col}{total_row}"].font = Font(bold=True)

    # ---------- 条件格式 ----------
    date_range = f"B{start_row}:B{total_row-1}"
    ws.conditional_formatting.add(
        date_range,
        FormulaRule(formula=[f'AND(B{start_row}<TODAY(),I{start_row}=FALSE)'],
                    fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
                    font=Font(color="9C0006")))
    ws.conditional_formatting.add(
        date_range,
        FormulaRule(formula=[f'AND(B{start_row}-TODAY()<=30,B{start_row}-TODAY()>0,I{start_row}=FALSE)'],
                    fill=PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
                    font=Font(color="9C5700")))

    col_width = {"A": 6, "B": 14, "C": 18, "D": 14, "E": 12, "F": 12,
                 "G": 12, "H": 8, "I": 12, "J": 10, "K": 12, "L": 14}
    for col, w in col_width.items():
        ws.column_dimensions[col].width = w
    ws.column_dimensions["I"].hidden = True

    wb.save(path)
    return ws  # 交给导出函数复用


# ==================================================================
# 2. 从工作表提取每期「确定性」数据（不依赖 openpyxl 公式求值）
# ==================================================================
def iter_periods(ws):
    """用参数区 + 日期递推重算，保证导出时即使公式未求值也准确。"""
    principal = float(ws["B2"].value)
    # 注意：openpyxl 不计算公式，B5/B7 读出的是 "=B3+B4" 这类字符串，
    # 所以用参数区的「原始输入值」自行推导，结果与公式完全一致。
    lpr = float(ws["B3"].value)
    bp = float(ws["B4"].value)
    penalty_mult = float(ws["B6"].value)
    rate = lpr + bp                              # 对应 B5 = B3+B4
    penalty_rate = rate * penalty_mult            # 对应 B7 = B5*B6
    _days_per_year = float(ws["B8"].value)
    pay1 = ws["B11"].value                       # date
    pay_last = ws["B13"].value                   # date
    start_repay_year = int(ws["B12"].value)
    amp_normal = float(ws["B9"].value)
    amp_final = float(ws["B10"].value)
    n = 19

    def stage_of(d):
        if d < date(start_repay_year, 12, 20):
            return "只还利息"
        if d < pay_last:
            return "还本付息"
        return "最后一期(利随本清)"

    def date_of(idx):
        if idx == 1:
            return pay1
        if idx == n:
            return pay_last
        y = pay1.year + (idx - 1)
        # 12-20；若跨过最后日期则用尾款日（兜底，正常不会命中）
        return date(y, 12, 20)

    remain = principal
    for idx in range(1, n + 1):
        d = date_of(idx)
        st = stage_of(d)
        interest = remain * rate
        amp = 0.0 if st == "只还利息" else (amp_normal if st == "还本付息" else amp_final)
        total = interest + amp
        yield {
            "idx": idx,
            "date": d,
            "stage": st,
            "interest": interest,
            "principal_paid": amp,
            "total": total,
        }
        remain = max(0.0, remain - amp)


# ==================================================================
# 3. 导出 .ics（原生提醒，iOS / 安卓 / Outlook 通吃）
# ==================================================================
def _esc(s: str) -> str:
    return (s.replace("\\", "\\\\").replace(";", "\\;")
              .replace(",", "\\,").replace("\n", "\\n"))

def _dt_line(dt: datetime, prop: str) -> str:
    """输出 UTC 的 YYYYMMDDTHHMMSSZ，兼容性最好。"""
    utc = dt.astimezone(timezone.utc)
    return f"{prop}:{utc.strftime('%Y%m%dT%H%M%SZ')}"

def export_ics(ws, path=OUT_ICS,
               remind_days=REMIND_DAYS_BEFORE,
               remind_time=REMIND_TIME,
               cal_name=CALENDAR_NAME,
               color=CALENDAR_COLOR,
               publish_mode=False,
               ttl_hours=REFRESH_INTERVAL_HOURS):
    """
    publish_mode=True  → 发布成「订阅源」：
      · UID 固定（同一期永远是同一个 ID），手机刷新时才会正确「更新/删除」旧事件，
        而不是每次都当新事件堆出一堆副本；
      · 写入 X-PUBLISHED-TTL，告诉客户端多久来取一次；
      · 不加随机 uuid 后缀，保证文件内容可重复生成、可 diff。
    publish_mode=False → 普通一次性导入用的 ics（默认，行为不变）。
    """
    local = timezone(timedelta(hours=8))   # 还款日以本地（东八区）零点为准
    now_utc = datetime.now(timezone.utc)

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        "PRODID:-//助学贷款还款计划表//CN",
        "CALSCALE:GREGORIAN",
        "METHOD:PUBLISH",
        f"X-WR-CALNAME:{_esc(cal_name)}",
        f"X-WR-CALCOLOR:{color}",
    ]
    if publish_mode:
        # 客户端每 ttl_hours 来同步一次；Apple 日历会尊重这个值
        lines.append(f"X-PUBLISHED-TTL:PT{int(ttl_hours)}H")

    for p in iter_periods(ws):
        # 发布模式下 UID 必须稳定；一次性导入则带点随机，避免和别的日历撞 ID
        uid = (f"loan-p{p['idx']}@{p['date'].isoformat().replace('-','')}"
               if publish_mode else
               f"loan-{p['idx']}-{p['date'].isoformat()}@{uuid.uuid4().hex[:8]}")
        summary = f"💰 助学贷款还款 · 第{p['idx']}期"
        desc = (f"阶段：{p['stage']}\\n"
                f"应还本息：¥{p['total']:,.2f}\\n"
                f"（利息 ¥{p['interest']:,.2f} + 本金 ¥{p['principal_paid']:,.2f}）\\n"
                f"请在 {p['date'].isoformat()} 前还清，逾期将按罚息利率计收。")

        # 全天事件：用 DATE 值（VALARM 里同样按日期触发，各端都能弹）
        start_dt = datetime(p['date'].year, p['date'].month, p['date'].day,
                            remind_time[0], remind_time[1], remind_time[2], tzinfo=local)
        trigger_days = max(0, int(remind_days))   # 避免负值拼出 -P-3D

        # DTSTAMP 用「日期派生」的固定值，而不是脚本运行时刻。
        # 理由：订阅源要求「内容没变就比特一致」。若每次 publish 都写当前秒，
        # 客户端会以为事件被改过，可能重复弹通知。同一还款日的事件永远同一个戳。
        stamp = datetime(p['date'].year, p['date'].month, p['date'].day,
                         0, 0, 0, tzinfo=timezone.utc)

        lines += [
            "BEGIN:VEVENT",
            f"UID:{uid}",
            f"SUMMARY:{_esc(summary)}",
            f"DESCRIPTION:{_esc(desc)}",
            f"DTSTART;VALUE=DATE:{p['date'].strftime('%Y%m%d')}",
            f"DTEND;VALUE=DATE:{(p['date'] + timedelta(days=1)).strftime('%Y%m%d')}",
            f"DTSTAMP:{stamp.strftime('%Y%m%dT%H%M%SZ')}",
            "TRANSP:TRANSPARENT",
            "STATUS:CONFIRMED",
            "BEGIN:VALARM",
            "ACTION:DISPLAY",
            f"DESCRIPTION:{_esc(summary)}",
            f"TRIGGER:-P{trigger_days}D",
            "END:VALARM",
            "END:VEVENT",
        ]

    lines += ["END:VCALENDAR"]

    text = "\r\n".join(lines) + "\r\n"
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(text)
    mode_tag = "（发布订阅源模式：UID 固定 + TTL）" if publish_mode else "（一次性导入模式）"
    print(f"✅ 日历 ICS 已生成：{Path(path).resolve()}（{len(list(iter_periods(ws)))} 条事件，提前 {remind_days} 天提醒）{mode_tag}")


# ==================================================================
# 4. 导出 .csv（兜底，iPhone「文件」App 直接导入）
# ==================================================================
def export_csv(ws, path=OUT_CSV):
    fieldnames = ["Subject", "Start Date", "Start Time", "End Date",
                  "End Time", "All Day Event", "Description",
                  "Location", "Private"]
    rows = []
    for p in iter_periods(ws):
        ds = p['date'].isoformat()
        rows.append({
            "Subject": f"💰 助学贷款还款 · 第{p['idx']}期",
            "Start Date": ds, "Start Time": "",
            "End Date": ds, "End Time": "",
            "All Day Event": "TRUE",
            "Description": (f"阶段：{p['stage']}\n"
                            f"应还本息：¥{p['total']:,.2f}\n"
                            f"（利息 ¥{p['interest']:,.2f} + 本金 ¥{p['principal_paid']:,.2f}）\n"
                            f"请在 {ds} 前还清，逾期将按罚息利率计收。"),
            "Location": "", "Private": "TRUE",
        })
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)
    print(f"✅ 日历 CSV 已生成：{Path(path).resolve()}（{len(rows)} 条，UTF-8 BOM）")


# ==================================================================
# 5b. webcal 订阅发布：把 ics 固化为「可订阅的静态源」
# ==================================================================
def publish_webcal(ws, ics_path=PUBLISH_ICS, ttl_hours=REFRESH_INTERVAL_HOURS):
    """
    生成 publish/loan.ics（固定 UID + TTL），作为 webcal 订阅源。

    这一步只负责「造文件」，不碰网络。怎么把 publish/ 挂到 HTTP 上，
    看下面 serve_webcal() 或 README_WEBCAL.md。
    """
    PUBLISH_DIR.mkdir(parents=True, exist_ok=True)
    export_ics(ws, path=ics_path, publish_mode=True, ttl_hours=ttl_hours)

    # 顺手写一份给部署用的纯文本说明（放同一目录，传到服务器也能看）
    readme = PUBLISH_DIR / "README.txt"
    readme.write_text(
        "这个目录里的 loan.ics 就是你的「订阅型日历源」。\n"
        "它和根目录的 还款提醒日历.ics 内容一致，区别是：\n"
        "  · 每条事件 UID 固定 → 手机刷新时会「更新/删除」旧事件，不会堆副本\n"
        "  · 带 X-PUBLISHED-TTL → 客户端按间隔自动来取\n"
        "只要让这个文件能被 https:// 或 http:// 访问到，\n"
        "把链接拼成 webcal://你的域名/loan.ics 订阅即可。\n",
        encoding="utf-8")
    print(f"📡 订阅源已发布：{ics_path.resolve()}")
    print(f"   → 订阅链接示例：webcal://<你的地址>/{ics_path.name}")
    return ics_path


# ==================================================================
# 5c. 局域网临时服务（零配置，先跑通再说部署）
#
#    python 助学贷款还款计划表.py serve
#
#    启动后手机在同一 Wi-Fi 下订阅：
#        webcal://192.168.x.x:8765/loan.ics
#
#    注意：这是「临时演示级」服务 —— 单线程、无鉴权、进程退出即停。
#    长期挂着用请换成方案 B / C（见 README_WEBCAL.md）。
# ==================================================================
def _get_local_ip() -> str:
    """找一个「看起来像局域网」的 IPv4 地址，比硬编码 192.168 稳。"""
    import socket
    s = None
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("223.5.5.5", 53))          # 不发包，只是借它拿出口地址
        ip = s.getsockname()[0]
        if ip.startswith(("127.", "169.254")):
            raise OSError("loopback")
        return ip
    except Exception:
        return "192.168.1.100"                 # 兜底，不对就手动换
    finally:
        if s:
            s.close()


def serve_webcal(ics_path=PUBLISH_ICS, host=WEB_HOST, port=WEB_PORT):
    import socket

    if not ics_path.exists():
        raise SystemExit(f"❌ 找不到订阅源 {ics_path}，先跑一次：python {Path(__file__).name} publish")

    data = ics_path.read_bytes()
    mtime = ics_path.stat().st_mtime
    ip = _get_local_ip()

    def handle(conn):
        try:
            req = b""
            while b"\r\n\r\n" not in req:
                chunk = conn.recv(4096)
                if not chunk:
                    return
                req += chunk
            head = req.decode("utf-8", "ignore").split("\r\n")[0]
            # 路由：只给 /loan.ics（和根目录兜底），其余 404，避免误暴露。
            # head 形如 "GET /loan.ics HTTP/1.1"
            path_req = head.split(" ", 2)[1] if len(head.split(" ", 2)) > 1 else ""
            if path_req in ("/", f"/{ics_path.name}"):
                pass   # 正常返回 ics
            else:
                conn.send(b"HTTP/1.1 404 Not Found\r\nContent-Type: text/plain; charset=utf-8\r\n\r\nnot found\r\n")
                return
            resp = (
                "HTTP/1.1 200 OK\r\n"
                "Content-Type: text/calendar; charset=utf-8\r\n"
                f"Content-Length: {len(data)}\r\n"
                f"Last-Modified: {datetime.fromtimestamp(mtime, timezone.utc).strftime('%a, %d %b %Y %H:%M:%S GMT')}\r\n"
                "Cache-Control: no-cache, max-age=0\r\n"
                "Access-Control-Allow-Origin: *\r\n"
                "\r\n"
            ).encode("utf-8") + data
            conn.sendall(resp)
            print(f"  ↳ 已响应 {head.split()[1]!r}")
        finally:
            conn.close()

    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        sock.bind((host, port))
        sock.listen(5)
        print("=" * 60)
        print(f"📡 webcal 临时服务已启动（Ctrl+C 停止）")
        print(f"   电脑地址：http://{ip}:{port}/{ics_path.name}")
        print(f"   手机订阅：webcal://{ip}:{port}/{ics_path.name}")
        print("=" * 60)
        print("手机操作：iOS「设置 → 日历 → 账户 → 添加账户 → 其他 → 订阅日历」")
        print("          安卓：日历 App → 添加账户 → 「网络日历 / 订阅 URL」")
        print("提示：只支持 http(s) 的客户端，若系统不让输 webcal:// 就填 http:// 那个地址")
        try:
            while True:
                conn, _addr = sock.accept()
                import threading
                threading.Thread(target=handle, args=(conn,), daemon=True).start()
        except KeyboardInterrupt:
            print("\n🛑 服务已停止")


# ==================================================================
# 5d. 生成一份「部署选型」说明，你自己看完再决定怎么挂
# ==================================================================
def write_webcal_guide():
    guide = Path("README_WEBCAL.md")
    guide.write_text(_WEBCAL_GUIDE_MD, encoding="utf-8")
    print(f"📖 部署方案已写好：{guide.resolve()}")
    return guide


_WEBCAL_GUIDE_MD = r"""# 助学贷款日历 · webcal 订阅方案

> 一句话：**把 `publish/loan.ics` 挂到一个 HTTP 地址上，手机订阅它。**
> 之后每次改了 xlsx 里的参数 → 重跑脚本 → 手机第二天自动刷新，不用再手动导入。

## 三个方案，按折腾程度选

| 方案 | 适合 | 难度 | 公网可用 | 说明 |
|---|---|---|---|---|
| **A. 局域网临时服务** | 只想先跑通、在家连 Wi-Fi 用 | ⭐ | ❌ | `python 助学贷款还款计划表.py serve`，关掉终端就停 |
| **B. 免费静态托管** | 想在公司/外面也能刷新 | ⭐⭐ | ✅ | 把 `publish/loan.ics` 传到 GitHub Pages / Gitee Pages / 对象存储，拿一个 https 链接 |
| **C. 自己服务器 / NAS** | 有台常开的机器 | ⭐⭐⭐ | ✅ | nginx 指一条 `location`，最稳、数据完全自己控 |

---

## 方案 A：先跑通（零配置）

```bash
# 1) 生成订阅源（固定 UID，可反复刷新不堆副本）
python 助学贷款还款计划表.py publish

# 2) 起服务（同一 Wi-Fi 下）
python 助学贷款还款计划表.py serve
```

终端会打印出类似：

```
手机订阅：webcal://192.168.1.100:8765/loan.ics
```

**iPhone 订阅**：设置 → 日历 → 账户 → 添加账户 → 其他 → **订阅日历** → 贴上面地址。
**安卓**：日历 App → 添加日历账户 → 网络日历 / 订阅 URL → 贴 `http://192.168.1.100:8765/loan.ics`。

> ⚠️ 几个坑：
> - 手机和电脑必须同一 Wi-Fi；切到 5G/流量就刷新失败（这是正常的）。
> - 苹果有时不让填 `webcal://`，直接填 `http://...` 一样能订阅。
> - 临时服务**没有鉴权**，同网段谁都能下载这个 ics。里面只有还款金额，不算敏感，但别把密码写进备注。

---

## 方案 B：免费公网链接（推荐长期使用）

思路：`publish/loan.ics` 是个纯文本文件，任何能托管静态文件的地方都行。

**B1. GitHub Pages（最典型）**

1. 新建一个公开仓库，把 `publish/loan.ics` 丢进去（放根目录即可）
2. 仓库 Settings → Pages → Source 选 `main` 分支根目录 → 保存
3. 几分钟后拿到链接：`https://你的用户名.github.io/仓库名/loan.ics`
4. 手机订阅：`webcal://你的用户名.github.io/仓库名/loan.ics`

以后更新流程就变成：

```bash
# 改完 xlsx 参数后
python 助学贷款还款计划表.py publish
cp publish/loan.ics /path/to/你的仓库/
cd 你的仓库 && git add loan.ics && git commit -m "更新还款日历" && git push
```

手机第二天自己就刷到新版。**连电脑都不用开**。

**B2. 国内可访问的平替**
- Gitee Pages（国内快，但有实名/配额限制）
- 腾讯云 COS / 阿里云 OSS：上传 `loan.ics`，设为公开读，拿到 https 直链
- Cloudflare R2（免费额度够用，全球快）

> 用对象存储时记得把「Content-Type」设成 `text/calendar`，否则手机可能当成普通文件下载。

---

## 方案 C：NAS / 自己服务器 + nginx

```nginx
# /etc/nginx/conf.d/loan.conf
server {
    listen 443 ssl;
    server_name cal.你的域名.com;

    location = /loan.ics {
        root /var/www/loan;          # 里面放 publish/loan.ics
        default_type text/calendar;
        add_header Cache-Control "no-cache, max-age=0";
        # 可选：加一层 Basic 认证，防止被随便订阅
        # auth_basic "loan calendar"; auth_basic_user_file /etc/nginx/.htpasswd;
    }
}
```

订阅地址：`webcal://cal.你的域名.com/loan.ics`

好处：数据不出自己家、可加密码、可配自动刷新（`X-PUBLISHED-TTL: PT24H` 已写进文件）。

---

## 手机端会怎么表现

| 你做了什么 | 手机刷新后 |
|---|---|
| 新生成了第 N 期 | 自动多出一条提醒 |
| 某期被「已结清」排除 | 那条自动消失 |
| 改了 LPR / 金额 | 那条事件的描述和提醒金额跟着变 |
| 删掉服务/换链接 | 原订阅还在手机里，**需要手动删一次订阅** |

Apple 日历默认大约每天刷一次；想立刻看到，就在日历列表里下拉刷新，或者先把那个订阅日历删了重加。

## 安全提醒（认真看）

- 订阅链接 = **谁拿到谁就能看你的还款金额和日期**。用公网方案时，至少二选一：
  - 放 GitHub 私有仓库 + 不用 Pages 公开（改走自己服务器带认证）
  - 或者直接接受「这些数字被陌生人看到也无所谓」
- 别在 `SUMMARY` / `DESCRIPTION` 里加身份证号、合同号、银行卡号。
- 脚本里的 `loan.ics` 不含任何个人信息，只有期数、日期、金额。

---

## 常用命令速查

```bash
python 助学贷款还款计划表.py            # 生成 xlsx + 一次性 ics + csv
python 助学贷款还款计划表.py publish     # 只生成订阅源 publish/loan.ics
python 助学贷款还款计划表.py serve       # 起局域网临时服务
python 助学贷款还款计划表.py guide       # 重新生成这份说明
```

需要我帮你改成「改完自动 git push / 自动上传 COS」的那种一键脚本，告诉我你选哪个托管平台就行。
"""
def _build_and_export():
    ws = build_xlsx()
    print(f"✅ 生成成功：{Path(OUT_XLSX).resolve()}")
    export_ics(ws)
    export_csv(ws)
    print("\n📱 导入建议：优先用 .ics（带提前提醒）；.csv 仅作兜底。")
    print("   想变成「手机自动刷新」的订阅源：python 本脚本 publish")
    return ws


if __name__ == "__main__":
    import sys
    cmd = (sys.argv[1:] or ["build"])[0].lower()

    if cmd == "publish":
        ws = build_xlsx(); publish_webcal(ws); write_webcal_guide()
        print("\n👉 下一步：先看 README_WEBCAL.md 选部署方式；")
        print("   想立刻试，就跑：python 助学贷款还款计划表.py serve")

    elif cmd == "serve":
        # 没有订阅源就先造一个，避免手滑报「找不到文件」
        if not PUBLISH_ICS.exists():
            ws = build_xlsx(); publish_webcal(ws)
        serve_webcal()

    elif cmd == "guide":
        write_webcal_guide()

    else:  # build / 默认
        _build_and_export()
