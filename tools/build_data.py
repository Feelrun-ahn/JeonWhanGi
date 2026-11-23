#!/usr/bin/env python3
"""구글 시트 「2026 목운중 전환기 일정」(xlsx 내보내기) → data.json

사용법:  python3 tools/build_data.py 시트.xlsx data.json [수정시각]
- 메인 탭(2026전환기일정)에서 날짜별 1~7교시와 병합(이어지는 교시)을 읽음
- 주간 안내 탭에서 교시별 상세 안내, 링크, 학사 일정을 읽음
- 반별 표(스포츠·농구 날)가 있으면 반마다 1~4교시를 따로 저장 (페이지에서 고른 반에 맞춰 보여줌)
"""
import json, re, sys
import openpyxl

URL_RE = re.compile(r'https?://[^\s<>"\')\]]+')

# 전환기 시작 전 날짜 (시트에 없음)
STATIC = {
    **{d: {"all": {"t": "1기 교과 프로그램", "sub": "시간표대로 운영"}} for d in
       ["2026-11-09", "2026-11-10", "2026-11-11", "2026-11-12", "2026-11-13", "2026-11-16", "2026-11-17"]},
    "2026-11-18": {"all": {"t": "동아리", "sub": "동아리 4 · 수능 감독관 예비소집일"}},
    "2026-11-19": {"all": {"t": "재량휴업일", "sub": "대학수학능력시험", "off": True}},
    "2026-11-20": {"all": {"t": "재량휴업일", "sub": "", "off": True}},
}

def norm(s):
    if s is None:
        return ""
    s = str(s).replace("\r", "").replace("\t", " ")
    s = re.sub(r"[  ]+", " ", s)
    return s.strip()

def one_line(s):
    return re.sub(r"\s*\n\s*", " ", norm(s)).strip()

def year_for(month):
    return 2027 if month <= 2 else 2026

def dkey(m, d):
    return f"{year_for(m)}-{m:02d}-{d:02d}"

FIXES = [
    (r"자기 주도학습", "자기주도학습"),
    (r"모의 고사", "모의고사"),
    (r"^교과 ?프로그램 ?(\d+)$", r"교과프로그램 \1"),
    (r"^합창$", "합창연습"),
    (r"^전환기 관련안내$", "전환기 관련 안내"),
    (r"졸업영상 ?만들기 ?(\d)", r"졸업영상 만들기 \1"),
    (r"^반별대항\((.+?)\) ?예선전$", r"반 대항 예선 (\1)"),
    (r"^협동화제작 ?", "협동화 제작 "),
    (r"^에코백제작$", "에코백 제작"),
    (r"^두뇌게임 스도쿠$", "두뇌게임 (스도쿠)"),
    (r"^두뇌게임-컬러링 숨은그림찾기$", "두뇌게임 (컬러링 · 숨은그림찾기)"),
    (r"^편지쓰기 \(메시지전달\)$", "편지 쓰기 (메시지 전달)"),
    (r"^롤링페이퍼 \(반친구돌려쓰기\)$", "롤링페이퍼 (반 친구 돌려 쓰기)"),
    (r"^학급대청소$", "학급 대청소"),
    (r"카드만들기", "카드 만들기"),
    (r"^자율\d*$", "자율"),
    (r"^크리스마스교실 꾸미기", "크리스마스 교실 꾸미기 "),
    (r"^합창대회설명 학급회의\+연습$", "합창대회 설명 · 학급회의"),
    (r"^민주주의선거교실$", "민주주의 선거교실"),
    (r"\s+\)", ")"),
    (r"\s{2,}", " "),
]

def pretty(title):
    t = title.strip()
    for a, b in FIXES:
        t = re.sub(a, b, t)
    return t.strip()

def split_cell(raw):
    """셀 원문 → (제목, 메모, 주식 반 목록)"""
    s = one_line(raw)
    note, classes = "", None
    m = re.match(r"^주식\s*((?:\d+\s*반?\s*,?\s*)+)(\D.*)$", s)
    if m:
        classes = [int(x) for x in re.findall(r"\d+", m.group(1))]
        s = m.group(2).strip()
    if s.startswith("학생회장"):
        note, s = "학생회장", re.sub(r"^학생회장,?\s*", "", s)
        s = re.sub(r"\d,\d교시$", "", s)
    m = re.match(r"^졸업영상 ?만들기 ?(\d)\s*(배구.*)$", s)
    if m:
        s, note = f"졸업영상 만들기 {m.group(1)}", "배구 대회 참여자는 8층 체육관 (연습)"
    if ":" in s:
        head, tail = s.split(":", 1)
        s, note = head.strip(), (note + " " + tail.strip()).strip()
        s = re.sub(r"\s*\d[.,]\d교시$", "", s)
    if s == "반 대항":
        s = "반 대항 결승"
        note = re.sub(r"\s*/\s*", " · ", re.sub(r"^방송송출\s*", "", note)).strip(" ·")
        note = "방송 송출 · " + note if note else "방송 송출"
    m = re.match(r"^(.+?)\s*\(방송\)$", s)
    if m:
        s, note = m.group(1), (note or "방송 송출")
    m = re.match(r"^진로콘서트\((.+)\)$", s)
    if m:
        s, note = "진로 콘서트", m.group(1)
    return pretty(s), note.replace("/ /", "/").strip(" /"), classes

ALLDAY_MAP = {
    "롯데월드": ("롯데월드", "외부활동 · 문화체험"),
    "공연- 뮤지컬": ("뮤지컬 관람", "외부활동 · 공연"),
    "국립중앙박물관": ("국립중앙박물관", "외부활동 · 역사체험"),
    "합창- 도림교회": ("합창 대회", "외부활동 · 도림교회"),
    "영화": ("영화 관람", "외부활동"),
}

def allday_of(raw):
    s = one_line(raw)
    m = re.match(r"^외부활동\s*\((.+)\)$", s)
    if m:
        inner = m.group(1).strip()
        t, sub = ALLDAY_MAP.get(inner, (inner, "외부활동"))
        return {"t": t, "sub": sub}
    if s in ("크리스마스", "신정"):
        return {"t": s, "sub": "쉬는 날", "off": True}
    if s.replace(" ", "") == "합창리허설":
        return {"t": "합창 리허설", "sub": "하루 종일"}
    if "졸업식" in s:
        return {"t": "졸업식 · 종업식", "sub": "창의적 체험활동(자율 2)"}
    return {"t": pretty(s), "sub": ""}

def merged_map(ws):
    """(row,col) → (rowspan, colspan) for top-left, and set of covered cells"""
    tops, covered = {}, set()
    for r in ws.merged_cells.ranges:
        tops[(r.min_row, r.min_col)] = (r.max_row - r.min_row + 1, r.max_col - r.min_col + 1)
        for rr in range(r.min_row, r.max_row + 1):
            for cc in range(r.min_col, r.max_col + 1):
                if (rr, cc) != (r.min_row, r.min_col):
                    covered.add((rr, cc))
    return tops, covered

def parse_main(ws, days):
    tops, covered = merged_map(ws)
    month = 11  # 첫 칸(A3)은 병합돼서 비어 있음
    for r in range(3, ws.max_row + 1):
        a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
        if a not in (None, ""):
            try:
                month = int(float(str(a).strip()))
            except ValueError:
                pass
        m = re.match(r"\s*(\d+)", str(b or ""))
        if not m:
            continue
        k = dkey(month, int(m.group(1)))
        entry = days.setdefault(k, {})
        info = []
        for col, label in ((4, "학사"), (5, "고입")):
            v = one_line(ws.cell(r, col).value)
            if v and (r, col) not in covered:
                info.append({"k": label, "t": v})
            elif (r, col) in covered:
                # 병합된 고입 일정은 위 칸 값을 이어받음
                for rr in range(r - 1, 2, -1):
                    if (rr, col) in tops:
                        rs, _ = tops[(rr, col)]
                        if rr + rs - 1 >= r:
                            v2 = one_line(ws.cell(rr, col).value)
                            if v2:
                                info.append({"k": label, "t": v2})
                        break
        note = one_line(ws.cell(r, 13).value)
        if note:
            entry["memo"] = note
        if info:
            entry["info"] = info

        first = ws.cell(r, 6)
        span_all = tops.get((r, 6), (1, 1))[1]
        if span_all >= 7:
            entry["all"] = allday_of(first.value)
            continue
        periods, p = [], 1
        while p <= 7:
            col = 5 + p
            if (r, col) in covered:
                p += 1
                continue
            span = tops.get((r, col), (1, 1))[1]
            raw = ws.cell(r, col).value
            if raw in (None, "") or not one_line(raw):
                periods.append({"t": "", "s": p, "n": span})
            else:
                t, note2, classes = split_cell(raw)
                item = {"t": t, "s": p, "n": span}
                if note2:
                    item["note"] = note2
                if classes:
                    item["stock"] = classes
                periods.append(item)
            p += span
        entry["p"] = periods

PERIOD_LABEL = re.compile(r"^\s*(\d)\s*(?:교시)?\s*(?:[~,]\s*(\d)\s*)?교시")

def label_periods(text):
    m = PERIOD_LABEL.match(text)
    if not m:
        return None
    a = int(m.group(1)); b = int(m.group(2) or a)
    return list(range(a, b + 1))

def split_summary(text):
    """'5교시 - ...\n6교시- ...' 형식의 묶음 안내를 교시별로 나눔"""
    parts = re.split(r"(?m)^\s*(?=\d(?:\s*[,~]\s*\d)?\s*교시\s*-)", text)
    out = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        m = re.match(r"^(\d)(?:\s*[,~]\s*(\d))?\s*교시\s*-\s*(.*)$", part, re.S)
        if m:
            a = int(m.group(1)); b = int(m.group(2) or a)
            out.append((list(range(a, b + 1)), m.group(3).strip()))
        else:
            out.append((None, part))
    return out

def clean_detail(t):
    t = norm(t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    t = "\n".join(line.strip() for line in t.split("\n"))
    return t.strip()

def parse_weekly(ws, days):
    tops, covered = merged_map(ws)
    # 1행에서 날짜 칸 찾기
    starts = []
    for c in range(1, ws.max_column + 1):
        v = norm(ws.cell(1, c).value)
        m = re.match(r"(\d+)/(\d+)\(", v)
        if m:
            starts.append((c, dkey(int(m.group(1)), int(m.group(2)))))
    if not starts:
        return
    bounds = []
    for i, (c, k) in enumerate(starts):
        end = starts[i + 1][0] - 1 if i + 1 < len(starts) else min(ws.max_column, c + 4)
        bounds.append((c, end, k))
    # 학사 일정 줄
    sched_row = next((r for r in range(1, 12) if norm(ws.cell(r, 1).value) == "학사 일정"), None)
    # 상세 안내 영역: '교과선생님께' 아래부터 15반 표 전까지
    table_row = next((r for r in range(1, ws.max_row + 1)
                      if re.search(r"\d+월\s*\d+일.*1~4교시", norm(ws.cell(r, 2).value))), ws.max_row + 1)
    for c0, c1, k in bounds:
        entry = days.setdefault(k, {})
        if sched_row:
            v = one_line(ws.cell(sched_row, c0).value)
            if v:
                info = entry.setdefault("info", [])
                comp = lambda t: re.sub(r"[\s()]", "", t)[:8]
                if not any(comp(x["t"]) == comp(v) for x in info):
                    info.append({"k": "학사", "t": v})
        guide = entry.setdefault("guide", [])
        for r in range(11, min(table_row, ws.max_row + 1)):
            cur = None
            for c in range(c0, c1 + 1):
                if (r, c) in covered:
                    continue
                cell = ws.cell(r, c)
                v = norm(cell.value)
                if not v or v.startswith("[매 교시]"):
                    continue
                link = cell.hyperlink.target if cell.hyperlink else None
                lp = label_periods(v) if len(v) <= 14 else None
                if lp:
                    cur = lp
                    continue
                if v.endswith("=>") or v.endswith("=&gt;"):
                    continue
                pieces = split_summary(v) if re.search(r"(?m)^\s*\d(?:\s*[,~]\s*\d)?\s*교시\s*-", v) else [(cur, v)]
                for per, txt in pieces:
                    txt = clean_detail(txt)
                    if not txt:
                        continue
                    links = []
                    if link:
                        links.append(link)
                    links += [u for u in URL_RE.findall(txt) if u not in links]
                    item = {"p": per, "t": txt}
                    if links:
                        item["links"] = links
                    if not any(g["t"] == txt and g["p"] == per for g in guide):
                        guide.append(item)
        if not guide:
            entry.pop("guide")
    # 반별 표 (스포츠·농구 경기 날 1~4교시)
    for r in range(1, ws.max_row + 1):
        title = norm(ws.cell(r, 2).value)
        m = re.search(r"(\d+)월\s*(\d+)일", title)
        if not (m and "1~4교시" in title):
            continue
        k = dkey(int(m.group(1)), int(m.group(2)))
        hdr = r + 1
        cols = {}
        for c in range(1, ws.max_column + 1):
            hv = re.sub(r"\s", "", norm(ws.cell(hdr, c).value))
            mm = re.match(r"^(\d+)(?:\.0)?반?$", hv)
            if mm:
                cols[int(mm.group(1))] = c
        table = {}
        for cls, col in cols.items():
            per = {}
            for rr in range(hdr + 1, hdr + 5):
                lp = label_periods(norm(ws.cell(rr, 1).value))
                v = norm(ws.cell(rr, col).value)
                if lp and v:
                    per[str(lp[0])] = pretty({"자기주도학": "자기주도학습", "합창": "합창연습"}.get(v, v))
            if per:
                table[str(cls)] = per
        if table:
            days.setdefault(k, {})["byClass"] = {"src": title, "c": table}

def main():
    src, out = sys.argv[1], sys.argv[2]
    updated = sys.argv[3] if len(sys.argv) > 3 else ""
    wb = openpyxl.load_workbook(src)
    days = {k: dict(v) for k, v in STATIC.items()}
    main_ws = next((ws for ws in wb.worksheets if "전환기일정" in ws.title.replace(" ", "")), wb.worksheets[0])
    parse_main(main_ws, days)
    for ws in wb.worksheets:
        if ws is main_ws:
            continue
        parse_weekly(ws, days)
    for k in list(days):
        if not days[k]:
            days.pop(k)
    # 배구 대회 안내 (체육대회 탭)
    for ws in wb.worksheets:
        if "체육" not in ws.title:
            continue
        for r in range(1, ws.max_row + 1):
            a, b, c = one_line(ws.cell(r, 1).value), norm(ws.cell(r, 2).value), one_line(ws.cell(r, 3).value)
            m = re.match(r"(\d+)/(\d+)\(.\)\s*\n?\s*(.*)", b)
            if a and m and "배구" in one_line(ws.cell(12, 1).value):
                k = dkey(int(m.group(1)), int(m.group(2)))
                if k in days:
                    txt = f"배구 대회 · {a}\n{one_line(m.group(3))}" + (f"\n{c}" if c else "")
                    days[k].setdefault("guide", []).append({"p": None, "t": txt.strip()})
    data = {"source": "2026 목운중 전환기 일정 (구글 시트)", "updated": updated,
            "days": dict(sorted(days.items()))}
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print(f"{len(days)} days → {out}")

if __name__ == "__main__":
    main()
