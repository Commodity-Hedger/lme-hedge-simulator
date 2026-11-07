"""커브 엑셀에서 6개 품목 커브를 읽어 curve.json을 만든다.

사용법: python scripts/update_curve.py "<파일명 (YYYYMMDD)>.xlsx"
저장된 curve.json의 date가 같거나 더 최신이면 아무것도 바꾸지 않는다.
6개 품목 중 하나라도 읽지 못하면 저장하지 않는다.
"""
import datetime as d, json, re, sys, time, pathlib
from openpyxl import load_workbook

SHEETS = ["CU", "AL", "ZN", "PB", "NI", "SN"]
OUT = pathlib.Path(__file__).resolve().parent.parent / "curve.json"

def to_iso(v):
    if isinstance(v, d.datetime): return v.date().isoformat()
    if isinstance(v, d.date): return v.isoformat()
    if isinstance(v, (int, float)) and 40000 < v < 60000:
        return (d.date(1899, 12, 30) + d.timedelta(int(round(v)))).isoformat()
    return None

def parse_sheet(ws):
    rows = [list(r) for r in ws.iter_rows(values_only=True)]
    cash = m3 = None; month_row = len(rows)
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            nxt = r[j + 1] if j + 1 < len(r) else None
            if v == "Cash Date": cash = to_iso(nxt)
            if v == "3M Date": m3 = to_iso(nxt)
            if isinstance(v, str) and "월물" in v and i < month_row: month_row = i
    pts = {}
    for i, r in enumerate(rows):
        for j, v in enumerate(r):
            nxt = r[j + 1] if j + 1 < len(r) else None
            if not isinstance(nxt, (int, float)) or isinstance(nxt, bool) or nxt <= 0: continue  # 0 = 빈 칸
            if i < month_row:
                if v == "Cash" and cash: pts[cash] = nxt
                elif v == "3Months" and m3: pts[m3] = nxt
                elif i < month_row - 3 and to_iso(v): pts[to_iso(v)] = nxt
            elif i > month_row + 1 and to_iso(v): pts[to_iso(v)] = nxt
    if not (cash and m3 and len(pts) > 5): return None
    return {"cash": cash, "m3": m3, "pts": [[k, round(pts[k], 4)] for k in sorted(pts)]}

def main(path):
    m = re.search(r"(20\d{6})", pathlib.Path(path).name)
    wb = load_workbook(path, data_only=True)
    data = {n: p for n in SHEETS if n in wb.sheetnames and (p := parse_sheet(wb[n]))}
    missing = [n for n in SHEETS if n not in data]
    if missing: sys.exit(f"{', '.join(missing)} 시트에서 커브를 찾지 못했습니다. 파일 형식을 확인하세요. (저장 안 함)")
    date = f"{m[1][:4]}-{m[1][4:6]}-{m[1][6:]}" if m else data["CU"]["cash"]
    if OUT.exists():
        old = json.loads(OUT.read_text(encoding="utf-8"))
        if old.get("date", "") >= date:
            print(f"curve.json이 이미 {old['date']} 기준입니다. 변경 없음."); return
    OUT.write_text(json.dumps({"date": date, "data": data, "updatedAt": int(time.time() * 1000)},
                              ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    print(f"curve.json 갱신: {date} ({', '.join(data)})")

if __name__ == "__main__":
    if len(sys.argv) != 2: sys.exit(__doc__)
    main(sys.argv[1])
