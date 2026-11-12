# -*- coding: utf-8 -*-
"""원천징수이행상황신고서 자동 작성기

급여대장 엑셀(원천1_급여대장 / 원천2_일용직 / 원천3_사업기타소득 탭 형식)을 읽어
지정한 귀속월의 원천징수이행상황신고서(소득세법 시행규칙 별지 제21호서식)를
한글(HWP) 파일로 만들고, 계산 근거와 검토사항을 담은 엑셀을 함께 만든다.

사용법:  python wht_report.py 급여대장.xlsx [--month 2025-06] [--paid 2025-06] [--yes]
"""
import sys, os, re, csv, shutil, argparse, calendar, datetime as dt
from collections import OrderedDict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

HERE = os.path.dirname(os.path.abspath(__file__))
TABLE_CSV = os.path.join(HERE, '간이세액표_2025.csv')
TEMPLATE = os.path.join(HERE, '별지21호_원천징수이행상황신고서_2025.hwp')

# ───────────────────────── 세법 기준값 (2025년 귀속) ─────────────────────────
MEAL_LIMIT = 200_000            # 식대 비과세 한도 (소득세법 제12조 제3호 러목)
CAR_LIMIT = 200_000             # 자가운전보조금 비과세 한도 (시행령 제12조 제3호)
DAILY_DEDUCTION = 150_000       # 일용근로소득공제 1일 (법 제47조 제2항)
DAILY_RATE = 0.06               # 일용근로 원천징수세율 (법 제129조)
DAILY_CREDIT = 0.55             # 근로소득세액공제율 (법 제59조 제3항)
BUSINESS_RATE = 0.03            # 사업소득 원천징수세율
OTHER_RATE = 0.20               # 기타소득 원천징수세율
OTHER_EXPENSE_HUMAN = 0.60      # 인적용역 기타소득 필요경비율 (시행령 제87조)
OTHER_MIN_INCOME = 50_000       # 기타소득 과세최저한(건별 기타소득금액) (법 제84조)
SMALL_AMOUNT = 1_000            # 소액부징수 (법 제86조)
LOCAL_RATE = 0.10               # 지방소득세 (지방세법 제103조의13)
CHILD_DEDUCT = {1: 12_500, 2: 29_160}   # 간이세액표 8~20세 자녀 공제 (별표2 제3호)

def trunc10(x):
    """10원 미만 절사"""
    x = round(x, 4)                     # 부동소수점 오차(1349.9999…) 방지
    return int(x // 10 * 10) if x > 0 else 0

# ───────────────────────── 간이세액표 ─────────────────────────
class SimpleTaxTable:
    def __init__(self, path):
        self.rows = []
        with open(path, encoding='utf-8-sig') as f:
            r = csv.reader(f); next(r)
            for row in r:
                self.rows.append([int(v) for v in row])
        self.top = self.rows[-1]          # 10,000천원 행

    def _base(self, pay, fam):
        """가족 1~11명 기준 표 금액"""
        if pay < self.rows[0][0]:
            return 0
        if pay <= 10_000_000:
            for lo, hi, *vals in self.rows:
                if lo <= pay < hi or (lo == 10_000_000 and pay == lo):
                    return vals[fam - 1]
        t10 = self.top[2 + fam - 1]
        over = pay - 10_000_000
        if pay <= 14_000_000:  return t10 + over * 0.98 * 0.35 + 25_000
        if pay <= 28_000_000:  return t10 + 1_397_000 + (pay - 14_000_000) * 0.98 * 0.38
        if pay <= 30_000_000:  return t10 + 6_610_600 + (pay - 28_000_000) * 0.98 * 0.40
        if pay <= 45_000_000:  return t10 + 7_394_600 + (pay - 30_000_000) * 0.40
        if pay <= 87_000_000:  return t10 + 13_394_600 + (pay - 45_000_000) * 0.42
        return t10 + 31_034_600 + (pay - 87_000_000) * 0.45

    def tax(self, pay, family=1, children=0, ratio=100):
        family = max(1, int(family))
        if family <= 11:
            base = self._base(pay, family)
        else:                                   # 별표2 제4호
            t11, t10 = self._base(pay, 11), self._base(pay, 10)
            base = t11 - (t10 - t11) * (family - 11)
        if children == 1:   base -= CHILD_DEDUCT[1]
        elif children == 2: base -= CHILD_DEDUCT[2]
        elif children >= 3: base -= CHILD_DEDUCT[2] + (children - 2) * 25_000
        base = max(0, base)
        return trunc10(base * ratio / 100)

# ───────────────────────── 엑셀 읽기 도우미 ─────────────────────────
def norm(s):
    return re.sub(r'[\s()\[\]·ㆍ_\-/]', '', str(s or ''))

def find_sheet(wb, *keys):
    for ws in wb.worksheets:
        if any(k in norm(ws.title) for k in keys):
            return ws
    return None

def find_header(ws, required, scan=15):
    """required: {표준이름: [동의어...]} → (머리글 행, {표준이름: 열번호})"""
    for r in range(1, min(ws.max_row, scan) + 1):
        cells = {c: norm(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)}
        found = {}
        for key, alts in required.items():
            for c, v in cells.items():
                if v and any(norm(a) == v or (len(norm(a)) >= 2 and norm(a) in v) for a in alts):
                    found.setdefault(key, c)
        if len(found) >= max(2, len(required) - 1):
            return r, found, cells
    return None, {}, {}

def to_num(v):
    if v is None or v == '': return 0
    if isinstance(v, (int, float)): return v
    s = str(v).replace(',', '').strip()
    try: return float(s)
    except ValueError: return 0

def to_month(v):
    if isinstance(v, (dt.datetime, dt.date)):
        return f'{v.year:04d}-{v.month:02d}'
    m = re.search(r'(\d{4})\D+(\d{1,2})', str(v or ''))
    return f'{int(m.group(1)):04d}-{int(m.group(2)):02d}' if m else None

def to_date(v):
    if isinstance(v, dt.datetime): return v.date()
    if isinstance(v, dt.date): return v
    m = re.search(r'(\d{4})\D+(\d{1,2})\D+(\d{1,2})', str(v or ''))
    return dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else None

def yn(v, default='N'):
    s = str(v or '').strip().upper()
    if s in ('Y', 'YES', 'O', '예', '네', '○', '있음', 'TRUE'): return 'Y'
    if s in ('N', 'NO', 'X', '아니오', '없음', 'FALSE'): return 'N'
    return default

def add_months(d, n):
    y, m = d.year + (d.month - 1 + n) // 12, (d.month - 1 + n) % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))

# ───────────────────────── 입력 읽기 ─────────────────────────
class Inputs:
    def __init__(self, path):
        self.path = path
        self.wb = openpyxl.load_workbook(path, data_only=True)
        self.warn_read = []
        self.biz = self._read_business()
        self.roster = self._read_roster()
        self.payroll = self._read_payroll()
        self.daily = self._read_daily()
        self.other = self._read_other()

    def _read_business(self):
        info = {'상호': '', '대표자': '', '사업자등록번호': '', '사업장': '', '전화번호': '', '전자우편': ''}
        ws = find_sheet(self.wb, '사업장정보', '고객개요', '기본정보')
        if not ws: return info
        labels = {'상호': ['상호', '법인명'], '대표자': ['대표자', '대표'], '사업자등록번호': ['사업자등록번호'],
                  '사업장': ['사업장', '사업장소재지', '주소'], '전화번호': ['전화번호', '연락처'], '전자우편': ['전자우편', '이메일']}
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 40)):
            for i, c in enumerate(row[:-1]):
                key = norm(c.value)
                for k, alts in labels.items():
                    if not info[k] and key in [norm(a) for a in alts] and row[i + 1].value not in (None, ''):
                        val = str(row[i + 1].value).strip()
                        if k == '대표자': val = re.sub(r'\s*\(.*\)$', '', val)
                        if k == '사업자등록번호': val = re.sub(r'\s*\(.*\)$', '', val)
                        info[k] = val
        return info

    def _read_roster(self):
        """고객개요 시트의 직원 현황 표(있으면) → 기본값 채우기용"""
        ws = find_sheet(self.wb, '고객개요', '직원현황', '사업장정보')
        out = {}
        if not ws: return out
        for r in range(1, ws.max_row + 1):
            heads = {norm(ws.cell(r, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(r, c).value}
            if '성명' in heads and ('입사일' in heads or '본인차량' in heads):
                for rr in range(r + 1, ws.max_row + 1):
                    name = ws.cell(rr, heads['성명']).value
                    if not name: break
                    g = lambda k: ws.cell(rr, heads[k]).value if k in heads else None
                    out[str(name).strip()] = {'입사일': to_date(g('입사일')), '본인차량': yn(g('본인차량')),
                                              '구분': str(g('구분') or ''), '관계': str(g('대표와의관계') or '')}
                break
        return out

    def _read_payroll(self):
        ws = find_sheet(self.wb, '급여대장')
        if not ws:
            raise SystemExit('급여대장 탭을 찾지 못했습니다. 탭 이름에 "급여대장"이 들어가야 합니다.')
        req = {'귀속월': ['귀속월', '귀속연월', '지급월'], '성명': ['성명', '이름'], '기본급': ['기본급', '급여']}
        hr, cols, cells = find_header(ws, req)
        if not hr or not {'귀속월', '성명', '기본급'} <= cols.keys():
            raise SystemExit(f'급여대장 탭의 머리글(귀속월·성명·기본급)을 찾지 못했습니다. 찾은 열: {list(cols)}')
        # 지급 항목 = 기본급 열부터 '지급합계' 직전까지 (없으면 기본급 오른쪽 숫자열 중 공제항목 전까지)
        stop_words = ['지급합계', '지급계', '합계', '총지급', '국민연금', '건강보험', '소득세']
        start = cols['기본급']; end = ws.max_column + 1
        for c in range(start + 1, ws.max_column + 1):
            if any(w in cells.get(c, '') for w in stop_words): end = c; break
        pay_cols = {c: str(ws.cell(hr, c).value).strip() for c in range(start, end) if ws.cell(hr, c).value}
        rows = []
        for r in range(hr + 1, ws.max_row + 1):
            name, mon = ws.cell(r, cols['성명']).value, to_month(ws.cell(r, cols['귀속월']).value)
            if not name or not mon: continue
            items = {h: to_num(ws.cell(r, c).value) for c, h in pay_cols.items()}
            rows.append({'월': mon, '성명': str(name).strip(), '항목': items, '행': r})
        return rows

    def _read_daily(self):
        ws = find_sheet(self.wb, '일용')
        if not ws: return []
        hr, cols, _ = find_header(ws, {'근무일': ['근무일', '일자', '날짜'], '성명': ['성명', '이름'], '일당': ['일당', '지급액', '일급']})
        if not hr or not {'근무일', '성명', '일당'} <= cols.keys():
            self.warn_read.append('일용직 탭 머리글(근무일·성명·일당)을 찾지 못해 일용직 자료를 건너뛰었습니다.'); return []
        out = []
        for r in range(hr + 1, ws.max_row + 1):
            d, name = to_date(ws.cell(r, cols['근무일']).value), ws.cell(r, cols['성명']).value
            if not d or not name: continue
            out.append({'일자': d, '성명': str(name).strip(), '일당': to_num(ws.cell(r, cols['일당']).value)})
        return out

    def _read_other(self):
        ws = find_sheet(self.wb, '사업기타', '사업소득', '기타소득')
        if not ws: return []
        hr, cols, _ = find_header(ws, {'지급일': ['지급일', '일자'], '성명': ['성명', '이름'], '지급액': ['지급액', '금액'],
                                       '소득구분': ['소득구분', '고객요청소득구분'], '내용': ['지급내용', '내용', '적요']})
        if not hr or not {'지급일', '성명', '지급액'} <= cols.keys():
            self.warn_read.append('사업·기타소득 탭 머리글(지급일·성명·지급액)을 찾지 못해 건너뛰었습니다.'); return []
        out = []
        for r in range(hr + 1, ws.max_row + 1):
            d, name = to_date(ws.cell(r, cols['지급일']).value), ws.cell(r, cols['성명']).value
            if not d or not name: continue
            g = lambda k: ws.cell(r, cols[k]).value if k in cols else ''
            out.append({'일자': d, '성명': str(name).strip(), '지급액': to_num(g('지급액')),
                        '소득구분': str(g('소득구분') or ''), '내용': str(g('내용') or '')})
        return out

    def months(self):
        return sorted({r['월'] for r in self.payroll} | {to_month(r['일자']) for r in self.daily} | {to_month(r['일자']) for r in self.other})

# ───────────────────────── 직원정보 파일 ─────────────────────────
EMP_HEAD = ['성명', '소득구분(자동/상용/일용)', '입사일', '공제대상가족수(본인포함)', '8~20세 자녀수',
            '원천징수비율(80/100/120)', '본인차량 업무사용(Y/N)', '식사 현물제공(Y/N)', '확인완료(Y)', '메모']
YEL = PatternFill('solid', fgColor='FFFFF2CC')

def load_or_make_employees(path, inp, names):
    """직원정보 파일을 읽고, 없는 사람은 기본값으로 추가해 저장한다."""
    emp, wb = OrderedDict(), None
    if os.path.exists(path):
        wb = openpyxl.load_workbook(path); ws = wb.active
        for r in range(2, ws.max_row + 1):
            n = ws.cell(r, 1).value
            if n: emp[str(n).strip()] = {h: ws.cell(r, i + 1).value for i, h in enumerate(EMP_HEAD)}
    added = []
    for n in names:
        if n in emp: continue
        ro = inp.roster.get(n, {})
        emp[n] = {'성명': n, '소득구분(자동/상용/일용)': '자동', '입사일': ro.get('입사일'),
                  '공제대상가족수(본인포함)': 1, '8~20세 자녀수': 0, '원천징수비율(80/100/120)': 100,
                  '본인차량 업무사용(Y/N)': ro.get('본인차량', 'N'), '식사 현물제공(Y/N)': 'N', '확인완료(Y)': '',
                  '메모': ('고객개요에서 가져옴. ' if ro else '') + (f"대표와의 관계: {ro['관계']}" if ro.get('관계') and ro['관계'] != '없음' else '')}
        added.append(n)
    if added or not os.path.exists(path):
        wb = openpyxl.Workbook(); ws = wb.active; ws.title = '직원정보'
        bold = Font(bold=True); thin = Side(style='thin', color='FF999999')
        for i, h in enumerate(EMP_HEAD, 1):
            c = ws.cell(1, i, h); c.font = bold; c.fill = PatternFill('solid', fgColor='FFD9E1F2')
            c.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')
        for r, (n, e) in enumerate(emp.items(), 2):
            for i, h in enumerate(EMP_HEAD, 1):
                c = ws.cell(r, i, e.get(h))
                if h == '입사일': c.number_format = 'yyyy-mm-dd'
                if str(e.get('확인완료(Y)') or '').upper() != 'Y' and i in (2, 4, 5, 6, 7, 8, 9): c.fill = YEL
        for i, w in enumerate([10, 14, 12, 14, 10, 14, 14, 14, 10, 50], 1):
            ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
        ws.row_dimensions[1].height = 45; ws.freeze_panes = 'B2'
        n = len(emp) + 3
        ws.cell(n, 1, '※ 노란 칸을 확인·수정한 뒤 "확인완료"에 Y를 적으세요. Y가 아닌 사람은 검토사항에 "기본값 사용"으로 표시됩니다.')
        ws.cell(n + 1, 1, '※ 소득구분: 자동 = 일용직 탭 근무자 중 같은 사업장에 3개월 이상 계속 근무하면 상용(A01)으로, 아니면 일용(A03)으로 처리.')
        try:
            wb.save(path)
        except PermissionError:
            print(f'  ! 직원정보 파일이 열려 있어 새 직원을 추가 저장하지 못했습니다: {os.path.basename(path)}')
    return emp, added

def eint(v, default):
    try: return int(float(v))
    except (TypeError, ValueError): return default

# ───────────────────────── 계산 ─────────────────────────
def compute(inp, emp, month, table):
    y, m = map(int, month.split('-'))
    month_end = dt.date(y, m, calendar.monthrange(y, m)[1])
    lines, warns = [], [('입력', '', w, '엑셀 머리글 확인') for w in inp.warn_read]

    def e(n): return emp.get(n, {})
    for n in sorted({r['성명'] for r in inp.payroll if r['월'] == month} | {r['성명'] for r in inp.daily if to_month(r['일자']) == month}):
        if str(e(n).get('확인완료(Y)') or '').upper() != 'Y':
            warns.append(('직원정보', n, '공제대상가족수·차량·식사 제공 여부가 확인되지 않아 기본값(가족 1명, 자녀 0명, 100%)으로 계산했습니다.', '직원정보 파일 확인 후 확인완료=Y'))

    # 1) 급여대장 → A01
    for r in [r for r in inp.payroll if r['월'] == month]:
        n, it = r['성명'], r['항목']; ee = e(n)
        meal = sum(v for h, v in it.items() if '식대' in norm(h) or '식사' in norm(h))
        car = sum(v for h, v in it.items() if '자가운전' in norm(h) or '차량' in norm(h))
        total = sum(it.values())
        meal_nt = 0 if yn(ee.get('식사 현물제공(Y/N)')) == 'Y' else min(meal, MEAL_LIMIT)
        car_nt = min(car, CAR_LIMIT) if yn(ee.get('본인차량 업무사용(Y/N)')) == 'Y' else 0
        taxable = total - meal_nt - car_nt
        fam, kids, ratio = eint(ee.get('공제대상가족수(본인포함)'), 1), eint(ee.get('8~20세 자녀수'), 0), eint(ee.get('원천징수비율(80/100/120)'), 100)
        tax = table.tax(taxable, fam, kids, ratio)
        notes = []
        if meal > meal_nt: notes.append(f'식대 {meal - meal_nt:,.0f}원 과세' + ('(식사 현물제공)' if meal_nt == 0 and meal else ''))
        if car > car_nt: notes.append(f'자가운전보조금 {car - car_nt:,.0f}원 과세' + ('(본인차량 업무사용 아님)' if car_nt == 0 else ''))
        other_items = [h for h in it if it[h] and not any(k in norm(h) for k in ('기본급', '식대', '식사', '자가운전', '차량'))]
        if other_items: notes.append('과세 처리: ' + ', '.join(other_items))
        if any('연장' in norm(h) or '야간' in norm(h) for h in other_items):
            warns.append(('비과세', n, '연장·야간근로수당을 과세로 처리했습니다. 월정액급여 210만원 이하 생산직 등이면 비과세될 수 있습니다.', '소득세법 시행령 제17조 요건 확인'))
        lines.append({'코드': 'A01', '구분': '근로-간이세액', '성명': n, '총지급액': taxable + meal_nt, '과세': taxable,
                      '제출비과세': meal_nt, '미제출비과세': car_nt, '소득세': tax,
                      '근거': f'과세급여 {taxable:,.0f}원, 가족 {fam}명' + (f', 자녀 {kids}명' if kids else '') + (f', {ratio}%' if ratio != 100 else '') + ('. ' + '; '.join(notes) if notes else '')})

    # 2) 일용직 탭 → 상용 판정 후 A01 또는 A03
    daily_m = [r for r in inp.daily if to_month(r['일자']) == month]
    for n in OrderedDict.fromkeys(r['성명'] for r in daily_m):
        recs = [r for r in daily_m if r['성명'] == n]; ee = e(n)
        days, pay = len(recs), sum(r['일당'] for r in recs)
        mode = str(ee.get('소득구분(자동/상용/일용)') or '자동').strip()
        start = to_date(ee.get('입사일')) or min(r['일자'] for r in inp.daily if r['성명'] == n)
        long_term = add_months(start, 3) <= month_end
        regular = mode == '상용' or (mode not in ('일용',) and long_term)
        if regular:
            fam, kids, ratio = eint(ee.get('공제대상가족수(본인포함)'), 1), eint(ee.get('8~20세 자녀수'), 0), eint(ee.get('원천징수비율(80/100/120)'), 100)
            tax = table.tax(pay, fam, kids, ratio)
            lines.append({'코드': 'A01', '구분': '근로-간이세액(일용직 탭)', '성명': n, '총지급액': pay, '과세': pay, '제출비과세': 0,
                          '미제출비과세': 0, '소득세': tax, '근거': f'{days}일 근무, 월 {pay:,.0f}원, 가족 {fam}명. {start:%Y-%m-%d}부터 3개월 이상 계속 근무 → 상용'})
            if mode != '상용':
                warns.append(('일용/상용', n, f'{start:%Y-%m-%d}부터 같은 사업장에 3개월 이상 계속 근무해 상용근로자(A01)로 처리했습니다. 일용으로 처리하면 세액 0원 등 결과가 달라집니다.', '소득세법 시행령 제20조 제1항 제3호. 1~전월 신고분도 같은 기준인지 확인'))
        else:
            tax = 0
            for r in recs:
                t = trunc10(max(0, r['일당'] - DAILY_DEDUCTION) * DAILY_RATE * (1 - DAILY_CREDIT))
                tax += t if t >= SMALL_AMOUNT else 0
            lines.append({'코드': 'A03', '구분': '근로-일용', '성명': n, '총지급액': pay, '과세': pay, '제출비과세': 0, '미제출비과세': 0,
                          '소득세': tax, '근거': f'{days}일 × 일당. (일당-15만원)×6%×45%, 1일 1천원 미만 소액부징수'})
            if mode == '일용' and long_term:
                warns.append(('일용/상용', n, f'직원정보에서 "일용"으로 지정했지만 {start:%Y-%m-%d}부터 3개월 이상 근무했습니다.', '소득세법 시행령 제20조 제1항 제3호'))

    # 3) 사업·기타소득
    for r in [r for r in inp.other if to_month(r['일자']) == month]:
        n, amt, kind = r['성명'], r['지급액'], norm(r['소득구분'])
        if '사업' in kind:
            tax = trunc10(amt * BUSINESS_RATE); tax = tax if tax >= SMALL_AMOUNT else 0
            lines.append({'코드': 'A25', '구분': '사업-매월징수', '성명': n, '총지급액': amt, '과세': amt, '제출비과세': 0, '미제출비과세': 0,
                          '소득세': tax, '근거': f'{r["일자"]:%m/%d} {r["내용"]} × 3%'})
            warns.append(('소득구분', n, f'"{r["내용"]}" {amt:,.0f}원을 사업소득(A25, 3%)으로 처리했습니다. 계속·반복성이 없으면 기타소득입니다.', '소득세법 제19조·제21조'))
        elif '기타' in kind:
            human = '그외' not in kind
            code = 'A59' if human else 'A42'
            exp = OTHER_EXPENSE_HUMAN if human else 0
            income = amt * (1 - exp)
            tax = 0 if income <= OTHER_MIN_INCOME else trunc10(income * OTHER_RATE)
            lines.append({'코드': code, '구분': '기타-인적용역' if human else '기타-그 외', '성명': n, '총지급액': amt, '과세': amt, '제출비과세': 0,
                          '미제출비과세': 0, '소득세': tax, '근거': f'{r["일자"]:%m/%d} {r["내용"]}. 소득금액 {income:,.0f}원 × 20%' + (' (필요경비 60%)' if human else '')})
            warns.append(('소득구분', n, f'"{r["내용"]}" {amt:,.0f}원을 기타소득 {"인적용역(A59, 필요경비 60%)" if human else "그 외(A42)"}로 처리했습니다. 계속·반복적인 용역이면 사업소득(3%)입니다.', '소득세법 제21조 제1항 제19호, 시행령 제87조'))
        else:
            warns.append(('소득구분', n, f'"{r["내용"]}" {amt:,.0f}원의 소득구분("{r["소득구분"]}")을 알 수 없어 신고서에서 뺐습니다.', '사업소득/기타소득 중 하나로 적어 주세요'))
    for l in lines: l['지방소득세'] = trunc10(l['소득세'] * LOCAL_RATE)
    return lines, warns

def summarize(lines):
    codes = OrderedDict((c, {'인원': 0, '총지급액': 0, '소득세': 0}) for c in ['A01', 'A03', 'A25', 'A59', 'A42'])
    for l in lines:
        s = codes[l['코드']]; s['인원'] += 1; s['총지급액'] += l['총지급액']; s['소득세'] += l['소득세']
    add = lambda *ks: {k: sum(codes[c][k] for c in ks) for k in ('인원', '총지급액', '소득세')}
    codes['A10'] = add('A01', 'A03'); codes['A30'] = add('A25'); codes['A40'] = add('A59', 'A42')
    codes['A99'] = {k: codes['A10'][k] + codes['A30'][k] + codes['A40'][k] for k in ('인원', '총지급액', '소득세')}
    return codes

def due_date(month):
    y, m = map(int, month.split('-'))
    d = add_months(dt.date(y, m, 10), 1)
    while d.weekday() >= 5: d += dt.timedelta(days=1)
    return d

# ───────────────────────── 출력: 검토 엑셀 ─────────────────────────
def write_review(path, inp, month, paid, lines, warns, codes, local_total, due):
    wb = openpyxl.Workbook()
    H = PatternFill('solid', fgColor='FFD9E1F2'); S = PatternFill('solid', fgColor='FFF2F2F2')
    b = Font(bold=True); t = Font(bold=True, size=14); NUM = '#,##0;(#,##0);-'
    ws = wb.active; ws.title = '요약'
    ws['A1'] = f'{inp.biz["상호"]} 원천징수이행상황신고 {month} 귀속'; ws['A1'].font = t
    info = [('귀속연월', month), ('지급연월', paid), ('신고·납부기한', f'{due:%Y-%m-%d} (공휴일이면 다음 영업일)'), ('입력 파일', os.path.basename(inp.path))]
    for i, (k, v) in enumerate(info, 3): ws.cell(i, 1, k).font = b; ws.cell(i, 2, v)
    r0 = 8
    for j, h in enumerate(['코드', '인원', '총지급액', '소득세 등', '납부세액(⑩)'], 1):
        c = ws.cell(r0, j, h); c.font = b; c.fill = H
    order = [c for c in ['A01', 'A03', 'A10', 'A25', 'A30', 'A59', 'A42', 'A40', 'A99'] if codes[c]['인원'] or c == 'A99']
    for i, c in enumerate(order, r0 + 1):
        s = codes[c]; sub = c in ('A10', 'A30', 'A40', 'A99')
        vals = [c, s['인원'], s['총지급액'], s['소득세'], s['소득세'] if sub else None]
        for j, v in enumerate(vals, 1):
            cell = ws.cell(i, j, v); cell.number_format = NUM
            if sub: cell.font = b; cell.fill = S
    r = r0 + len(order) + 2
    ws.cell(r, 1, '국세 납부(홈택스)').font = b; ws.cell(r, 2, codes['A99']['소득세']).number_format = NUM
    ws.cell(r + 1, 1, '지방소득세(위택스)').font = b; ws.cell(r + 1, 2, f"=SUM(집계근거!J2:J{len(lines) + 1})").number_format = NUM
    ws.cell(r + 2, 1, '합계').font = b; ws.cell(r + 2, 2, f'=B{r}+B{r + 1}').number_format = NUM
    ws.cell(r + 4, 1, f'검토사항 {len(warns)}건 → "검토사항" 탭을 먼저 확인하세요.').font = Font(bold=True, color='FFC00000')
    for L, w in zip('ABCDE', (22, 30, 14, 12, 14)): ws.column_dimensions[L].width = w

    d = wb.create_sheet('집계근거')
    hd = ['코드', '구분', '성명', '총지급액(⑤)', '과세 지급액', '제출비과세(식대)', '미제출비과세(자가운전)', '소득세', '', '지방소득세', '산출근거']
    for j, h in enumerate(hd, 1):
        if h: c = d.cell(1, j, h); c.font = b; c.fill = H; c.alignment = Alignment(wrap_text=True)
    for i, l in enumerate(lines, 2):
        vals = [l['코드'], l['구분'], l['성명'], l['총지급액'], l['과세'], l['제출비과세'], l['미제출비과세'], l['소득세'], None, l['지방소득세'], l['근거']]
        for j, v in enumerate(vals, 1):
            c = d.cell(i, j, v)
            if j in (4, 5, 6, 7, 8, 10): c.number_format = NUM
    tr = len(lines) + 2
    d.cell(tr, 1, '합계').font = b
    for L in 'DEFGHJ':
        c = d[f'{L}{tr}']; c.value = f'=SUM({L}2:{L}{tr - 1})'; c.number_format = NUM; c.font = b
    for L, w in zip('ABCDEFGHIJK', (6, 20, 9, 13, 13, 13, 15, 10, 2, 11, 70)): d.column_dimensions[L].width = w
    d.row_dimensions[1].height = 32; d.freeze_panes = 'D2'

    v = wb.create_sheet('검토사항')
    for j, h in enumerate(['#', '분류', '대상', '내용', '근거·확인할 것'], 1):
        c = v.cell(1, j, h); c.font = b; c.fill = H
    for i, w in enumerate(warns, 2):
        for j, x in enumerate((i - 1,) + tuple(w), 1):
            c = v.cell(i, j, x); c.alignment = Alignment(wrap_text=True, vertical='top')
    for L, w in zip('ABCDE', (4, 10, 9, 70, 40)): v.column_dimensions[L].width = w
    wb.save(path)

# ───────────────────────── 출력: HWP ─────────────────────────
def fmt(n): return f'{int(n):,}'

def write_hwp(out_path, pdf_path, inp, month, paid, codes, sign_date):
    import win32com.client as win32
    shutil.copyfile(TEMPLATE, out_path)
    hwp = win32.Dispatch('HWPFrame.HwpObject')
    log = []
    try:
        try: hwp.RegisterModule('FilePathCheckDLL', 'FilePathCheckerModule')
        except Exception: pass
        if not hwp.Open(out_path, 'HWP', 'forceopen:true'):
            raise RuntimeError('한글에서 서식 파일을 열지 못했습니다.')

        def find(s, top=True):
            if top: hwp.MovePos(2, 0, 0)
            p = hwp.HParameterSet.HFindReplace
            hwp.HAction.GetDefault('RepeatFind', p.HSet)
            p.FindString = s; p.Direction = 0; p.IgnoreMessage = 1; p.FindRegExp = 0
            ok = hwp.HAction.Execute('RepeatFind', p.HSet)
            if not ok: log.append(f'서식에서 "{s}"을(를) 찾지 못함')
            return ok
        def ins(t):
            p = hwp.HParameterSet.HInsertText
            hwp.HAction.GetDefault('InsertText', p.HSet); p.Text = t; hwp.HAction.Execute('InsertText', p.HSet)
        def replace(s, t, top=True):
            if find(s, top): ins(t)
        def right_of(label, t):
            if t and find(label):
                hwp.HAction.Run('Cancel'); hwp.HAction.Run('TableRightCell'); ins(t)
        def row(code, vals):
            if not find(code): return
            hwp.HAction.Run('Cancel')
            for v in vals:
                hwp.HAction.Run('TableRightCell')
                if v not in (None, '', 0):
                    ins(fmt(v)); hwp.HAction.Run('ParagraphShapeAlignRight')

        y, m = month.split('-'); py, pm = paid.split('-')
        replace('[ ]원천징수이행상황신고서', '[○]원천징수이행상황신고서')
        if find('매월'):
            ins('○매월')
            if find('○매월'):
                for _ in range(4): hwp.HAction.Run('CharShapeWidthDecrease')
                for _ in range(3): hwp.HAction.Run('CharShapeSpacingDecrease')
        replace('    년   월', f'{y}년 {int(m):2d}월')
        replace('    년   월', f'{py}년 {int(pm):2d}월')
        right_of('법인명(상호)', inp.biz['상호'])
        right_of('대표자(성명)', inp.biz['대표자'])
        hwp.MovePos(2, 0, 0)
        replace('여, 부', '여, 부(○)', False); replace('여, 부', '여, 부(○)', False)
        right_of('등록번호', inp.biz['사업자등록번호'])
        right_of('사업장 소재지', inp.biz['사업장'])
        right_of('전화번호', inp.biz['전화번호'])
        for code in ['A01', 'A03', 'A10', 'A25', 'A30', 'A59', 'A42', 'A40', 'A99']:
            s = codes[code]
            if not s['인원'] and code != 'A99': continue
            pay_col = s['소득세'] if code in ('A10', 'A30', 'A40', 'A99') else None
            row(code, [s['인원'], s['총지급액'], s['소득세'], None, None, None, pay_col, None])
        replace('년     월     일', f'{sign_date.year}년   {sign_date.month:2d}월   {sign_date.day:2d}일')
        if inp.biz['대표자']:
            replace('신고인               ', f'신고인     {inp.biz["대표자"]}     ')
        hwp.SaveAs(out_path, 'HWP', '')
        if pdf_path: hwp.SaveAs(pdf_path, 'PDF', '')
    finally:
        try: hwp.Clear(1)
        except Exception: pass
        hwp.Quit()
    return log

# ───────────────────────── 메인 ─────────────────────────
def ask(prompt, default, auto):
    if auto: return default
    try:
        s = input(f'{prompt} [엔터 = {default}]: ').strip()
    except EOFError:
        s = ''
    return s or default

def main():
    ap = argparse.ArgumentParser(description='급여대장 → 원천징수이행상황신고서(HWP)')
    ap.add_argument('xlsx'); ap.add_argument('--month'); ap.add_argument('--paid')
    ap.add_argument('--yes', action='store_true', help='묻지 않고 기본값으로 진행')
    ap.add_argument('--pdf', action='store_true', help='확인용 PDF도 함께 저장')
    a = ap.parse_args()
    src = os.path.abspath(a.xlsx)
    print('=' * 60); print(' 원천징수이행상황신고서 자동 작성'); print('=' * 60)
    print(f'입력: {src}')
    inp = Inputs(src)
    months = inp.months()
    if not months: raise SystemExit('급여 자료에서 귀속월을 찾지 못했습니다.')
    print(f'자료에 있는 월: {", ".join(months)}')
    month = to_month(a.month or ask('신고할 귀속월(예: 2025-06)', months[-1], a.yes))
    if month not in months: raise SystemExit(f'{month} 자료가 없습니다.')
    paid = to_month(a.paid or ask('지급연월(급여를 실제로 준 달)', month, a.yes))

    folder = os.path.dirname(src)
    biz = inp.biz['상호'] or os.path.splitext(os.path.basename(src))[0]
    emp_path = os.path.join(folder, f'{biz}_직원정보.xlsx')
    names = sorted({r['성명'] for r in inp.payroll} | {r['성명'] for r in inp.daily})
    emp, added = load_or_make_employees(emp_path, inp, names)
    if added: print(f'직원정보 파일에 {len(added)}명을 기본값으로 추가했습니다: {os.path.basename(emp_path)}')

    table = SimpleTaxTable(TABLE_CSV)
    lines, warns = compute(inp, emp, month, table)
    codes = summarize(lines)
    local_total = sum(l['지방소득세'] for l in lines)
    due = due_date(month)
    y, m = month.split('-')
    base = os.path.join(folder, f'{biz}_원천징수이행상황신고서_{y}년{int(m)}월귀속')
    write_review(base + '_검토.xlsx', inp, month, paid, lines, warns, codes, local_total, due)
    log = write_hwp(base + '.hwp', base + '.pdf' if a.pdf else None, inp, month, paid, codes, due)
    warns_hwp = [f'  - {x}' for x in log]

    print('-' * 60)
    for c in ['A01', 'A03', 'A10', 'A25', 'A30', 'A59', 'A42', 'A40', 'A99']:
        s = codes[c]
        if s['인원'] or c == 'A99':
            print(f'  {c}  인원 {s["인원"]:>3}  총지급액 {s["총지급액"]:>13,.0f}  소득세 {s["소득세"]:>10,.0f}')
    print(f'  지방소득세(위택스) {local_total:,.0f}원   납부기한 {due:%Y-%m-%d}')
    print('-' * 60)
    print(f'만든 파일:\n  {base}.hwp\n  {base}_검토.xlsx' + (f'\n  {base}.pdf' if a.pdf else ''))
    if warns: print(f'\n검토사항 {len(warns)}건 → 검토 엑셀의 "검토사항" 탭을 확인하세요.')
    if warns_hwp: print('\nHWP 기입 경고:'); print('\n'.join(warns_hwp))

if __name__ == '__main__':
    try:
        main()
    except SystemExit as e:
        if e.code not in (None, 0): print(f'\n[중단] {e.code}')
        sys.exit(1 if e.code not in (None, 0) else 0)
    except PermissionError as e:
        print(f'\n[중단] 파일이 다른 프로그램에서 열려 있어 저장하지 못했습니다. 한글/엑셀에서 닫고 다시 실행하세요.\n  {e.filename}')
        sys.exit(1)
    except Exception:
        import traceback; traceback.print_exc()
        print('\n[중단] 예기치 못한 오류입니다. 위 메시지를 캡처해 담당자에게 보내 주세요.')
        sys.exit(1)
