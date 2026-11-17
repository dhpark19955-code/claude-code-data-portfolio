# -*- coding: utf-8 -*-
"""IB 서식 DATASET 공용 엔진 — 종목별 빌더가 이걸 import 해서 쓴다.

설계 원칙 (ib-excel-format 스킬):
  파랑 = 하드코딩 원자료 · 검정 = 동일시트 수식 · 초록 = 타시트 링크
  괄호 음수 · 0은 '-' · 퍼센트는 분수 저장 · 병합 없이 centerContinuous 밴드
  A열 스페이서 / B열 라벨 / C열부터 데이터 · 격자선 off · 한글이라 맑은 고딕

★ 하지 말 것 (Excel이 파일 손상으로 보고 열기를 거부한다):
  - freeze_panes 를 걸었다가 "A1"로 되돌리기 → <pane> 없이 <selection pane=..> 만 남는다
  - cell.hyperlink = "#'시트'!A1"  → 외부 관계(rels)로 기록된다. Hyperlink(location=..) 를 쓸 것
수식을 쓴 워크북은 저장 후 recalc_dataset.py 로 Excel 재계산해야 캐시값이 생긴다.
"""
import os, sys, glob, unicodedata
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from ib_style import (IB, apply_workbook_defaults, input_cell, formula_cell,  # noqa: E402
                      link_cell, source_row, set_font_everywhere)

KFONT = "맑은 고딕"
GREY = "FF595959"


def font(bold=False, italic=False, color=IB.BLACK, size=9):
    return Font(name=KFONT, size=size, bold=bold, italic=italic, color=color)


def is_note(v):
    return isinstance(v, str) and v.strip().startswith(("★", "※", "주석", "출처", "·", "-", "  "))


def is_band(v):
    return isinstance(v, str) and v.strip().startswith(("■", "①", "②", "③", "④", "⑤"))


class Dataset:
    """한 종목의 DATASET 워크북."""

    def __init__(self, base, out_path, company, ticker, today):
        self.base, self.out = base, out_path
        self.company, self.ticker, self.today = company, ticker, today
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.reg = []          # (name, title, desc, unit, period)
        self._cache = {}
        self._idx = {}
        for root, dirs, fns in os.walk(base):
            for fn in fns:
                self._idx.setdefault(unicodedata.normalize("NFC", fn), os.path.join(root, fn))

    # ───────── 소스 읽기 (NFD 파일명 함정 회피) ─────────
    def rows_of(self, fname, sheet, skip=0, drop_notes=True, keep_bands=True):
        key = unicodedata.normalize("NFC", fname)
        if key not in self._cache:
            if key not in self._idx:
                raise FileNotFoundError(fname)
            self._cache[key] = openpyxl.load_workbook(self._idx[key], read_only=True,
                                                      data_only=True)
        wb = self._cache[key]
        tgt = unicodedata.normalize("NFC", sheet)
        real = [s for s in wb.sheetnames if unicodedata.normalize("NFC", s) == tgt]
        if not real:
            raise KeyError(f"{fname} :: {sheet}")
        out = []
        for r in list(wb[real[0]].iter_rows(values_only=True))[skip:]:
            if not any(c not in (None, "") for c in r):
                continue
            v = r[0]
            if is_band(v) and keep_bands:
                out.append(tuple(r))
                continue
            if drop_notes and is_note(v) and not any(c not in (None, "") for c in r[1:]):
                continue
            out.append(tuple(r))
        # 왼쪽에 완전히 빈 열이 있으면 떼어낸다 (원본이 A열을 스페이서로 쓴 경우).
        while out and all((len(r) == 0 or r[0] in (None, "")) for r in out):
            out = [r[1:] for r in out]
            if not any(len(r) for r in out):
                break
        # 오른쪽 완전 빈 열도 정리
        while out and all(len(r) and r[-1] in (None, "") for r in out):
            out = [r[:-1] for r in out]
        return [r for r in out if any(c not in (None, "") for c in r)]

    # ───────── 셀 프리미티브 ─────────
    def new_sheet(self, name, title, units, source, label_w=30, data_w=11):
        ws = self.wb.create_sheet(name[:31])
        apply_workbook_defaults(ws, label_width=label_w, data_width=data_w)
        c = ws.cell(row=1, column=2, value=title)
        c.font = font(bold=True, size=11)
        c = ws.cell(row=2, column=2, value=units)
        c.font = font(italic=True)
        s = source_row(ws, 3, f"출처: {source}", col=2)
        s.font = font(italic=True, color=GREY)
        return ws

    def band(self, ws, row, ncol, text, style="blue", start_col=2):
        pal = {"blue": (IB.BAND_BLUE, IB.WHITE), "black": (IB.BAND_BLACK, IB.WHITE),
               "pale": (IB.BAND_BLUE_PALE, IB.BLACK), "grey": (IB.BAND_GREY, IB.BLACK)}
        fill, fc = pal[style]
        for col in range(start_col, start_col + max(ncol, 2)):
            c = ws.cell(row=row, column=col)
            c.fill = PatternFill("solid", fgColor=fill)
            c.alignment = Alignment(horizontal="centerContinuous")
            c.font = font(bold=True, color=fc)
        ws.cell(row=row, column=start_col, value=text)
        return row + 1

    def head(self, ws, row, headers, start_col=3, label_text="항목"):
        lc = ws.cell(row=row, column=2, value=label_text)
        lc.font = font(bold=True)
        lc.fill = PatternFill("solid", fgColor=IB.BAND_GREY)
        lc.alignment = Alignment(horizontal="left")
        for i, h in enumerate(headers):
            c = ws.cell(row=row, column=start_col + i, value=h)
            c.font = font(bold=True)
            c.fill = PatternFill("solid", fgColor=IB.BAND_GREY)
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        return row + 1

    def row(self, ws, row, lab, vals, start_col=3, fmt=IB.NUM, bold=False, indent=0):
        lc = ws.cell(row=row, column=2, value=lab)
        lc.font = font(bold=bold)
        if indent:
            lc.alignment = Alignment(indent=indent)
        for i, v in enumerate(vals):
            if v is None or v == "":
                continue
            col = start_col + i
            if isinstance(v, str) and v.startswith("="):
                (link_cell if "!" in v else formula_cell)(ws, row, col, v, numfmt=fmt, bold=bold)
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                input_cell(ws, row, col, v, numfmt=fmt, bold=bold)
            else:
                c = ws.cell(row=row, column=col, value=v)
                c.font = font(bold=bold)
                c.alignment = Alignment(horizontal="right", wrap_text=True)
        return row + 1

    def notes(self, ws, row, lines, col=2):
        for n in lines:
            c = ws.cell(row=row, column=col, value=n)
            c.font = font(italic=True, color=GREY)
            row += 1
        return row

    # ───────── 선언형: 소스 시트를 그대로 표로 ─────────
    def add(self, name, title, desc, unit, period, src, units_label=None, source_label=None,
            label_w=30, data_w=11, skip=0, band_title=None, notes=(), fmt_map=None,
            bold_labels=(), default_fmt=IB.NUM, col_widths=None, drop_notes=True,
            keep_source_notes=False):
        """src = (파일명, 시트명). 첫 행을 헤더로 보고 나머지를 데이터로 찍는다."""
        f, s = src
        rows = self.rows_of(f, s, skip=skip, drop_notes=drop_notes)
        ws = self.new_sheet(name, title, units_label or f"(단위: {unit})",
                            source_label or f"{f} :: {s}", label_w=label_w, data_w=data_w)
        ncol = max(len(r) for r in rows)
        r = 4
        if band_title:
            r = self.band(ws, r, ncol, band_title)
        hdr = rows[0]
        r = self.head(ws, r, [("" if h is None else h) for h in hdr[1:]],
                      label_text=("항목" if hdr[0] in (None, "") else str(hdr[0])))
        tail_notes = []
        for src_row in rows[1:]:
            lab = src_row[0]
            body = src_row[1:]
            if is_band(lab) and not any(c not in (None, "") for c in body):
                r = self.band(ws, r, ncol, str(lab).lstrip("■ ").strip(), style="pale")
                continue
            if is_note(lab) and not any(c not in (None, "") for c in body):
                if keep_source_notes:
                    tail_notes.append(str(lab))
                continue
            sl = str(lab) if lab is not None else ""
            fmt = default_fmt
            if fmt_map:
                for k, v in fmt_map.items():
                    if k in sl:
                        fmt = v
                        break
            elif "%" in sl or "율" in sl or "비중" in sl:
                fmt = IB.NUM1
            r = self.row(ws, r, sl.strip(), list(body), fmt=fmt,
                         bold=any(b in sl for b in bold_labels) if bold_labels else False,
                         indent=1 if sl.startswith(("  ", "\u3000")) else 0)
        if notes or tail_notes:
            self.notes(ws, r + 1, list(notes) + tail_notes)
        if col_widths:
            for i, w in enumerate(col_widths):
                ws.column_dimensions[get_column_letter(i + 2)].width = w
        self.reg.append((name[:31], title, desc, unit, period))
        return ws

    def register(self, name, title, desc, unit, period):
        self.reg.append((name[:31], title, desc, unit, period))

    # ───────── README ─────────
    def readme(self, headline_lines, vintage, order=None, excluded=()):
        ws = self.wb.create_sheet("00_README")
        apply_workbook_defaults(ws, label_width=26)
        c = ws.cell(row=1, column=2, value=f"{self.company} ({self.ticker}) — DATASET")
        c.font = font(bold=True, size=13)
        c = ws.cell(row=2, column=2, value=f"수치 전용 데이터셋 · IB 서식 · 기준일 {self.today}")
        c.font = font(italic=True)
        r = 4
        base_lines = [
            "색 규칙:  파랑 = 하드코딩 원자료 입력 |  검정 = 동일시트 수식 |  초록 = 타시트 링크",
            "숫자 규칙:  음수는 괄호 · 0은 '-' 표기 · 퍼센트는 분수 저장(0.15 → 15.0%) · 배수는 0.0x",
        ]
        for line in base_lines + list(headline_lines):
            cc = ws.cell(row=r, column=2, value=line)
            cc.font = font()
            r += 1
        if excluded:
            r += 1
            r = self.band(ws, r, 6, "제외한 것")
            for a, b in excluded:
                ws.cell(row=r, column=2, value=a).font = font(bold=True)
                ws.cell(row=r, column=3, value=b).font = font()
                r += 1
        r += 1
        r = self.band(ws, r, 6, "데이터 사전")
        r = self.head(ws, r, ["제목", "내용", "단위", "기간"], label_text="시트")
        names = order or [x[0] for x in self.reg]
        by = {x[0]: x for x in self.reg}
        for nm in names:
            if nm not in by:
                continue
            _, title, desc, unit, period = by[nm]
            lc = ws.cell(row=r, column=2, value=nm)
            lc.font = Font(name=KFONT, size=9, color="FF0563C1", underline="single")
            lc.hyperlink = Hyperlink(ref=lc.coordinate, location=f"'{nm}'!A1")
            for j, v in enumerate([title, desc, unit, period]):
                cc = ws.cell(row=r, column=3 + j, value=v)
                cc.font = font()
                cc.alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        r += 1
        r = self.band(ws, r, 6, "데이터 빈티지 · 한계")
        for a, b in vintage:
            ws.cell(row=r, column=2, value=a).font = font(bold=True)
            cc = ws.cell(row=r, column=3, value=b)
            cc.font = font()
            cc.alignment = Alignment(wrap_text=True, vertical="top")
            r += 1
        for col, w in zip("BCDEF", (26, 44, 46, 16, 26)):
            ws.column_dimensions[col].width = w
        return ws

    # ───────── 저장 ─────────
    def save(self, freeze="C7"):
        for ws in self.wb.worksheets:
            set_font_everywhere(ws, KFONT)
            ws.sheet_view.showGridLines = False
            # ★ README에는 틀고정을 걸지 않는다 (걸었다 풀면 파일이 손상된다)
            if ws.title != "00_README":
                ws.freeze_panes = freeze

        def key(ws):
            n = ws.title
            if n == "00_README":
                return (0, 0)
            p = n.split("_")[0]
            return (1, int(p) if p.isdigit() else 999)

        self.wb._sheets.sort(key=key)
        self.wb.save(self.out)
        print(f"시트 {len(self.wb.sheetnames)}개 -> {os.path.basename(self.out)}")
        return self.out


def project_base():
    """프로젝트 루트 = 이 파일(scripts/)의 상위. 폴더가 이동해도 깨지지 않게 절대경로 glob 대신 __file__ 기준."""
    here = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if os.path.isdir(os.path.join(here, "마스터")):
        return here
    hits = [c for c in glob.glob(os.path.join(os.path.expanduser("~"), "Downloads", "**", "260908_*"), recursive=True)
            if os.path.isdir(c)]
    if not hits:
        raise FileNotFoundError("프로젝트 루트(260908_투자대상 스크리닝)를 찾지 못했다")
    return hits[0]


def fs_sheets(ds, prefix, fin_file, sheet_map, fs_map):
    """재무요약(연결/별도 × 연간/분기) 4시트를 공통 서식으로. fs_map에 라벨→행 기록."""
    for name, title, srcsheet, period in sheet_map:
        ws = ds.new_sheet(name, title, "(단위: 백만원, IFRS)",
                          f"DART 재무제표 API · {fin_file} :: {srcsheet}", label_w=26)
        rows = ds.rows_of(fin_file, srcsheet, drop_notes=False)
        hdr = [str(h) for h in rows[0][1:] if h]
        r = ds.band(ws, 4, len(hdr) + 1, "손익계산서")
        r = ds.head(ws, r, hdr)
        for src_row in rows[1:]:
            lab = src_row[0]
            if lab == "__1__":
                r = ds.band(ws, r, len(hdr) + 1, "재무상태표")
                continue
            if lab == "__2__":
                r = ds.band(ws, r, len(hdr) + 1, "현금흐름표")
                continue
            vals = list(src_row[1:len(hdr) + 1])
            if not any(v not in (None, "") for v in vals):
                continue
            sl = str(lab)
            bold = sl.strip() in ("매출액(영업수익)", "영업이익", "당기순이익",
                                  "자산총계", "부채총계", "자본총계")
            fmt = IB.NUM1 if "EPS" in sl else IB.NUM
            r = ds.row(ws, r, sl.strip(), vals, fmt=fmt, bold=bold,
                       indent=1 if sl.startswith("  ") else 0)
            fs_map[(name, sl.strip())] = r - 1
        ds.notes(ws, r + 1, [
            "원자료 시트 — 모든 수치는 하드코딩 입력(파랑). 가공은 재무비율·KPI 시트에서 수식으로 참조한다.",
            "빈칸은 DART 원문에 해당 계정이 없는 경우다(0이 아님).",
        ])
        ds.register(name, title, "DART 실측 재무제표", "백만원", period)


def ratio_sheet(ds, name, fs_sheet_name, fs_map, years, cols, source_note=""):
    """재무비율 시트 — 전부 타시트 참조 수식."""
    ws = ds.new_sheet(name, "Financial Ratios", "(단위: %, 배)",
                      f"{fs_sheet_name} 참조 수식" + (f" · {source_note}" if source_note else ""),
                      label_w=30)
    r = ds.band(ws, 4, len(years) + 1, "성장성 · 수익성 · 안정성")
    r = ds.head(ws, r, years)

    def ref(lab, col):
        return f"'{fs_sheet_name}'!{col}{fs_map[(fs_sheet_name, lab)]}"

    def has(lab):
        return (fs_sheet_name, lab) in fs_map

    rev = "매출액(영업수익)"
    specs = [
        ("영업수익 성장률", IB.PCT, "growth", rev),
        ("영업이익률", IB.PCT, "ratio", ("영업이익", rev)),
        ("순이익률", IB.PCT, "ratio", ("당기순이익", rev)),
        ("ROE (당기순이익/자본총계)", IB.PCT, "ratio", ("당기순이익", "자본총계")),
        ("ROA (당기순이익/자산총계)", IB.PCT, "ratio", ("당기순이익", "자산총계")),
        ("부채비율 (부채/자본)", IB.PCT, "ratio", ("부채총계", "자본총계")),
        ("현금/자산", IB.PCT, "ratio", ("현금및현금성자산", "자산총계")),
        ("영업현금흐름/영업이익", IB.MULT, "ratio", ("영업활동현금흐름", "영업이익")),
    ]
    for lab, fmt, kind, arg in specs:
        if kind == "growth":
            if not has(arg):
                continue
            vals = [None] + [f'=IFERROR({ref(arg, c1)}/{ref(arg, c0)}-1,"")'
                             for c0, c1 in zip(cols, cols[1:])]
        else:
            a, b = arg
            if not (has(a) and has(b)):
                continue
            vals = [f'=IFERROR({ref(a, c)}/{ref(b, c)},"")' for c in cols]
        r = ds.row(ws, r, lab, vals, fmt=fmt)
    ds.notes(ws, r + 1, [
        "전부 수식(초록 = 타시트 참조). 원자료 시트의 숫자를 바꾸면 자동 갱신된다.",
        "ROE·ROA는 기말 잔액 기준(평균 아님).",
    ])
    ds.register(name, "Financial Ratios", "성장성·수익성·안정성 (수식)", "%/배",
                f"{years[0]}~{years[-1]}")
