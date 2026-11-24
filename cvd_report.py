# -*- coding: utf-8 -*-
"""
뇌심혈관질환 위험성평가 총괄평가서 자동작성 프로그램
=====================================================
검진프로그램(누리온 등)에서 내려받은 '뇌심혈관질환 위험도 평가 조회' 개인별 엑셀(.xls/.xlsx)을
드래그&드롭(또는 파일 선택)하면, 같은 폴더에  <원본파일명>_최종보고서.docx / .xlsx  를 만들어 줍니다.

[설치]  (최초 1회, 명령 프롬프트에서)
    pip install pandas numpy matplotlib openpyxl python-docx xlrd tkinterdnd2 pdfplumber pillow pytesseract

[같은 폴더에 두어야 하는 파일]
    cvd_report.py(실행), cvd_engine.py(위험도 계산), cvd_extract.py(PDF/사진 판독),
    cvd_store.py(사업장·부서별 저장), cvd_gui_individual.py(개인 입력 창)

[실행]
    python cvd_report.py                      -> 창이 열림 (파일을 창에 끌어다 놓기)
    python cvd_report.py 파일.xls             -> 창이 열리고 해당 파일이 미리 선택됨
    python cvd_report.py 파일.xls --nogui --top 최고위험군:3,고위험군:3   -> 창 없이 바로 생성

[참고] 본 프로그램의 '관리방향'은 평가 결과(인원·위험요인 보유율)를 근거로 만든 규칙 기반 초안입니다.
       최종 확정은 직업환경의학전문의(의사인 보건관리자)의 검토를 거쳐야 합니다.
"""
import os
import re
import sys
import shutil
import argparse
import tempfile
import threading
import subprocess
import textwrap
import datetime as dt
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

# ----------------------------------------------------------------------------
# 0. 상수 / 설정
# ----------------------------------------------------------------------------
RISKS = ["저위험군", "중등도위험군", "고위험군", "최고위험군"]
RISK_COLOR = {"저위험군": "#4C9F70", "중등도위험군": "#F2C14E",
              "고위험군": "#F08A3C", "최고위험군": "#D64545"}
WORKS = ["통상근무", "조건부근무", "병가 또는 휴직", "작업전환"]
AGE_BINS = [0, 29, 39, 49, 59, 200]
AGE_LABELS = ["20대 이하", "30대", "40대", "50대", "60대 이상"]
DOC_FONT = "맑은 고딕"
SUFFIX = "_최종보고서"

# 부서명 표기 정리용: 괄호로 붙은 근무형태/구분 표기(예: (야), (일검))는 같은 부서로 합칩니다.
# 프로그램과 같은 폴더에 '부서명_매핑.csv'(열: 원본,표준) 를 두면 추가로 적용됩니다.
SHIFT_TAG = re.compile(r"\((야|일검|배|주|일|주간|야간)\)")
DEFAULT_ALIAS = {"물류": "물류팀"}

# 참고문헌 (보고서 마지막에 자동 수록). '확인' 표시가 붙은 항목은 발간 전 서지사항을 한 번 더 확인하세요.
REFERENCES = [
    "한국산업안전보건공단. KOSHA GUIDE H-200-2018 직장에서의 뇌·심혈관계질환 예방을 위한 발병위험도 평가 및 사후관리지침. 2018.12.",
    "산업안전보건기준에 관한 규칙 제669조(직무스트레스에 의한 건강장해 예방조치).",
    "대한고혈압학회. 2018 대한고혈압학회 고혈압 진료지침. (H-200-2018 인용 기준)",
    "대한의학회. 2015 이상지질혈증 진료지침. (H-200-2018 인용 기준)",
    "Williams B, et al. 2018 ESC/ESH Guidelines for the management of arterial hypertension. Eur Heart J 2018;39:3021-3104.",
    "2003 World Health Organization (WHO)/International Society of Hypertension (ISH) Statement on Management of Hypertension. J Hypertens 2003;21(11):1983-1992.",
    "Yi 등. A Comparison of KOSHA's Cardiovascular Disease Risk Assessment and the Predicted 10-year Risk of Cardiovascular Disease Developed by Jee for Male Workers at a Wallpaper and Floor Covering Manufacturer. Korean J Occup Environ Med 2009;21(2):174-.",
    "이정엽, 양선희, 정인성, 이미영. 소규모사업장 근로자의 뇌심혈관질환 발병위험도에 대한 사후관리의 효과. Keimyung Med J 2016;35(2).",
    "KOSHA와 국민건강보험공단(NHIS) 심뇌혈관질환 위험도 평가도구의 상관성·일치도 비교. Ann Occup Environ Med 2023 (PMID 37701487). [서지사항(저자·권호) 확인 필요]",
    "Comparison of risk-assessment tools for cardio-cerebrovascular diseases (CVD) in male shipyard workers: a cross-sectional study. Ann Occup Environ Med. [저자·연도·권호 확인 필요]",
    "대한고혈압학회. 2022 고혈압 진료지침. [최신 개정본 확인 필요]",
    "한국지질·동맥경화학회. 이상지질혈증 진료지침 제5판(2022). [최신 개정본 확인 필요]",
]

# 위험요인 표시순서 / 개선 가능 요인 / 요인별 관리방안
FACTORS = ["흡연", "운동부족", "음주", "비만·복부비만", "혈압 전단계 이상",
           "공복혈당 이상", "이상지질혈증(TC·LDL·TG)", "HDL 저하",
           "연령 기준(남45·여55세 이상)", "직계가족 조기발병", "기저질환 진단·치료 중"]
MODIFIABLE = {"흡연", "운동부족", "음주", "비만·복부비만", "혈압 전단계 이상",
              "공복혈당 이상", "이상지질혈증(TC·LDL·TG)", "HDL 저하"}
ACTIONS = {
    "흡연": "금연 프로그램(금연클리닉·보건소 연계, 금연 상담) 우선 운영",
    "운동부족": "주 3일 이상, 1회 30분 이상 운동 실천 프로그램(근무 중 스트레칭·걷기 등) 운영",
    "음주": "절주 교육·상담 및 고위험 음주자 선별 후 개별 상담",
    "비만·복부비만": "영양지도와 체중관리 프로그램(BMI 25 미만, 허리둘레 남 90cm·여 85cm 미만 목표)",
    "혈압 전단계 이상": "혈압 자가측정·정기 추적, 1기 이상 고혈압은 진료 연계(생활요법±약물치료)",
    "공복혈당 이상": "당뇨병 선별검사 추적(공복혈당·당화혈색소) 및 식이·운동 생활요법",
    "이상지질혈증(TC·LDL·TG)": "지질검사 추적 및 식이지도, 위험도에 따라 약물치료 고려(진료 연계)",
    "HDL 저하": "유산소 운동 증진·금연·절주 지도",
}
TIER_ACTION = {  # KOSHA GUIDE H-200-2018 <표 10>
    "최고위험군": "병가·휴직 여부를 포함한 의사 업무적합성평가, 근무상 조치 및 작업(환경) 개선 권고, 생활습관 개선과 약물치료(진료 연계)",
    "고위험군": "조건부 근무 또는 병가·휴직 검토, 의사 업무적합성평가, 근무상 조치 및 작업(환경) 개선 권고, 생활습관 개선과 약물치료",
    "중등도위험군": "조건부 근무, 의사 업무적합성평가, 보건관리자 상담, 생활습관 개선 또는 약물치료",
    "저위험군": "통상근무(또는 조건부 근무), 보건관리자 상담, 생활습관 개선",
}


def setup_font():
    names = {f.name for f in font_manager.fontManager.ttflist}
    for n in ["Malgun Gothic", "맑은 고딕", "NanumGothic", "Noto Sans CJK KR", "Noto Sans KR",
              "AppleGothic", "UnDotum", "Baekmuk Gulim", "Noto Sans CJK JP"]:
        if n in names:
            plt.rcParams["font.family"] = n
            break
    plt.rcParams["axes.unicode_minus"] = False


setup_font()


# ----------------------------------------------------------------------------
# 1. 파일 읽기 (.xls 가 xlrd 로 안 열리는 경우를 대비한 다단계 시도)
# ----------------------------------------------------------------------------
def _patch_xlrd():
    """검진프로그램이 만든 .xls 는 파일 끝에 여분 섹터가 있어 xlrd 가 'corrupt?' 오류를 내는 경우가 있음.
    해당 검사만 우회하도록 xlrd 함수를 패치한다."""
    import inspect
    from xlrd import compdoc
    src = textwrap.dedent(inspect.getsource(compdoc.CompDoc._locate_stream))
    pat = re.compile(r'raise CompDocError\(\s*"%s: size exceeds expected %d bytes; corrupt\?"\s*%\s*\(qname, found_limit \* sec_size\)\s*\)')
    new_src, n = pat.subn("break", src)
    if n == 0:
        raise RuntimeError("xlrd 패치 대상 코드를 찾지 못함")
    new_src = new_src.replace("assert s == EOCSID", "pass").replace("assert tot_found == found_limit", "pass")
    ns = dict(vars(compdoc))
    exec(new_src, ns)
    compdoc.CompDoc._locate_stream = ns["_locate_stream"]


def _convert_with_excel(path, outdir):
    import win32com.client  # pywin32
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    try:
        wb = xl.Workbooks.Open(str(Path(path).resolve()))
        out = Path(outdir) / (Path(path).stem + ".xlsx")
        wb.SaveAs(str(out), FileFormat=51)
        wb.Close(False)
    finally:
        xl.Quit()
    return out


def _convert_with_soffice(path, outdir):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        for p in [r"C:\Program Files\LibreOffice\program\soffice.exe",
                  r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"]:
            if os.path.exists(p):
                exe = p
    if not exe:
        raise RuntimeError("LibreOffice 없음")
    subprocess.run([exe, "--headless", "--convert-to", "xlsx", "--outdir", str(outdir), str(path)],
                   check=True, capture_output=True, timeout=180)
    return Path(outdir) / (Path(path).stem + ".xlsx")


def load_table(path, log=print):
    path = Path(path)
    ext = path.suffix.lower()
    if ext == ".csv":
        for enc in ("utf-8-sig", "cp949"):
            try:
                return pd.read_csv(path, encoding=enc)
            except Exception:
                pass
        raise ValueError("CSV 를 읽을 수 없습니다.")
    if ext in (".xlsx", ".xlsm"):
        return pd.read_excel(path)
    if ext == ".xls":
        errs = []
        try:
            return pd.read_excel(path)
        except Exception as e:
            errs.append(f"xlrd: {e}")
        try:
            _patch_xlrd()
            log("  (xls 호환 모드로 읽는 중)")
            return pd.read_excel(path)
        except Exception as e:
            errs.append(f"xlrd-patch: {e}")
        tmp = tempfile.mkdtemp()
        for fn in (_convert_with_excel, _convert_with_soffice):
            try:
                out = fn(path, tmp)
                log(f"  ({fn.__name__} 로 변환하여 읽는 중)")
                return pd.read_excel(out)
            except Exception as e:
                errs.append(f"{fn.__name__}: {e}")
        raise RuntimeError("엑셀 파일을 읽지 못했습니다. 엑셀에서 '다른 이름으로 저장 > .xlsx' 후 다시 시도해 주세요.\n" + "\n".join(errs))
    raise ValueError(f"지원하지 않는 파일 형식: {ext}")


# ----------------------------------------------------------------------------
# 2. 데이터 정리
# ----------------------------------------------------------------------------
def load_alias():
    alias = dict(DEFAULT_ALIAS)
    for base in (Path(sys.argv[0]).resolve().parent, Path.cwd()):
        f = base / "부서명_매핑.csv"
        if f.exists():
            try:
                m = pd.read_csv(f, encoding="utf-8-sig")
                alias.update(dict(zip(m.iloc[:, 0].astype(str), m.iloc[:, 1].astype(str))))
            except Exception:
                pass
            break
    return alias


def norm_dept(name, level="team", alias=None):
    s = "" if pd.isna(name) else str(name).strip()
    if not s:
        return "(부서 미기재)"
    s = SHIFT_TAG.sub("", s)
    s = re.sub(r"\s*-\s*\d*\s*$", "", s).strip()
    if level == "team":
        s = s.split("(")[0].strip()
    if alias and s in alias:
        s = alias[s]
    return s or "(부서 미기재)"


def norm_work(v):
    s = "" if pd.isna(v) else str(v)
    if "통상" in s:
        return "통상근무"
    if "조건" in s:
        return "조건부근무"
    if "병가" in s or "휴직" in s:
        return "병가 또는 휴직"
    if "전환" in s:
        return "작업전환"
    return None


def _s(df, col):
    if col in df.columns:
        return df[col].fillna("").astype(str)
    return pd.Series([""] * len(df), index=df.index)


def prepare(raw, level="team", log=print):
    need = ["부서명", "성별", "연령", "발병위험도평가"]
    miss = [c for c in need if c not in raw.columns]
    if miss:
        raise ValueError("필수 열이 없습니다: " + ", ".join(miss) +
                         "\n(검진프로그램의 '뇌심혈관질환 위험도 평가 조회' 개인별 엑셀인지 확인하세요)")
    df = raw.copy()
    df["위험군"] = _s(df, "발병위험도평가").str.strip()
    bad = ~df["위험군"].isin(RISKS)
    if bad.any():
        log(f"  ※ 위험군이 비어 있거나 인식되지 않는 {int(bad.sum())}명은 제외했습니다.")
    df = df[~bad].copy()
    g = _s(df, "성별").str.strip()
    df["성"] = np.where(g.str.startswith("남") | g.str.upper().str.startswith("M"), "남",
                        np.where(g.str.startswith("여") | g.str.upper().str.startswith("F"), "여", "미상"))
    df["연령"] = pd.to_numeric(df["연령"], errors="coerce")
    df["연령대"] = pd.cut(df["연령"], AGE_BINS, labels=AGE_LABELS)
    alias = load_alias()
    df["부서_분석"] = df["부서명"].map(lambda x: norm_dept(x, level, alias))
    df["근무조치"] = df["업무적합성평가"].map(norm_work) if "업무적합성평가" in df.columns else None
    if "검진일자" in df.columns:
        df["검진일자"] = pd.to_datetime(df["검진일자"], errors="coerce")
    if "공단검진위험도" in df.columns:
        v = pd.to_numeric(df["공단검진위험도"], errors="coerce")
        df["공단구간"] = pd.cut(v, [-1, 0.01 - 1e-9, 0.05 - 1e-9, 0.10 - 1e-9, 10],
                               labels=RISKS, right=True)
    # 위험요인 플래그
    rf = _s(df, "위험인자")
    bp = _s(df, "혈압수준")
    hx = _s(df, "과거및현병력") + "," + _s(df, "동반된질병상태")
    flags = pd.DataFrame(index=df.index)
    flags["흡연"] = _s(df, "흡연유무").str.strip().eq("흡연") if "흡연유무" in df.columns else rf.str.contains("흡연")
    flags["운동부족"] = _s(df, "신체활동부족").str.contains("운동부족")
    flags["음주"] = _s(df, "음주여부").str.strip().eq("음주")
    flags["비만·복부비만"] = rf.str.contains("비만")
    flags["혈압 전단계 이상"] = bp.str.contains("전단계|1기|2기")
    flags["공복혈당 이상"] = rf.str.contains("공복혈당|당화혈색소|식후")
    flags["이상지질혈증(TC·LDL·TG)"] = rf.str.contains("총콜레스테롤|LDL|중성지방")
    flags["HDL 저하"] = rf.str.contains(r"HDL콜레스테롤\(40")
    flags["연령 기준(남45·여55세 이상)"] = rf.str.contains("연령")
    flags["직계가족 조기발병"] = rf.str.contains("가족") | _s(df, "가족력").str.strip().ne("")
    flags["기저질환 진단·치료 중"] = hx.str.contains("고혈압|이상지질|당뇨|뇌혈관|심혈관|신장")
    flags["_고혈압1기이상"] = bp.str.contains("1기|2기")
    flags["_단백뇨"] = _s(df, "표적장기손상여부").str.contains("단백뇨")
    flags["_당뇨"] = hx.str.contains("당뇨")
    return df, flags


# ----------------------------------------------------------------------------
# 3. 분석
# ----------------------------------------------------------------------------
def pct(a, b):
    return (a / b * 100.0) if b else 0.0


def analyze(df, flags):
    A = {}
    A["n"] = len(df)
    A["n_m"] = int((df["성"] == "남").sum())
    A["n_f"] = int((df["성"] == "여").sum())
    A["sites"] = ", ".join(sorted(_s(df, "사업장명").str.strip().replace("", np.nan).dropna().unique())) or "(사업장명 미기재)"
    if "검진일자" in df.columns and df["검진일자"].notna().any():
        a, b = df["검진일자"].min(), df["검진일자"].max()
        A["period"] = f"{a:%Y년 %m월 %d일} ~ {b:%Y년 %m월 %d일}"
    else:
        A["period"] = "(검진일자 없음)"
    # 위험군 × 성별
    rg = pd.DataFrame(0, index=["계", "남", "여"], columns=RISKS)
    for r in RISKS:
        rg.loc["계", r] = int((df["위험군"] == r).sum())
        rg.loc["남", r] = int(((df["위험군"] == r) & (df["성"] == "남")).sum())
        rg.loc["여", r] = int(((df["위험군"] == r) & (df["성"] == "여")).sum())
    A["risk_gender"] = rg
    # 업무적합성 × 성별
    wg = pd.DataFrame(0, index=["계", "남", "여"], columns=WORKS)
    for w in WORKS:
        wg.loc["계", w] = int((df["근무조치"] == w).sum())
        wg.loc["남", w] = int(((df["근무조치"] == w) & (df["성"] == "남")).sum())
        wg.loc["여", w] = int(((df["근무조치"] == w) & (df["성"] == "여")).sum())
    A["work_gender"] = wg
    # 위험군 × 근무조치
    A["risk_work"] = pd.crosstab(df["위험군"], df["근무조치"]).reindex(index=RISKS, columns=WORKS, fill_value=0)
    # 연령대
    A["age_risk"] = pd.crosstab(df["연령대"], df["위험군"]).reindex(index=AGE_LABELS, columns=RISKS, fill_value=0)
    # 위험요인 보유율
    fl = flags[FACTORS]
    A["prev_all"] = fl.mean() * 100
    A["prev_by_risk"] = pd.DataFrame({r: fl[df["위험군"] == r].mean() * 100 for r in RISKS}).fillna(0)
    A["prev_by_sex"] = pd.DataFrame({s: fl[df["성"] == s].mean() * 100 for s in ["남", "여"]}).fillna(0)
    A["flags_n"] = fl.sum()
    # 질환·표적장기
    A["extra"] = {
        "고혈압 1·2기 해당": int(flags["_고혈압1기이상"].sum()),
        "단백뇨(표적장기손상)": int(flags["_단백뇨"].sum()),
        "당뇨병(과거·현병력/동반질환)": int(flags["_당뇨"].sum()),
    }
    # 공단 10년 위험도 구간
    if "공단구간" in df.columns:
        A["nhis"] = df["공단구간"].value_counts().reindex(RISKS).fillna(0).astype(int)
    return A


def dept_table(df, min_n):
    t = pd.crosstab(df["부서_분석"], df["위험군"]).reindex(columns=RISKS, fill_value=0)
    t["계"] = t[RISKS].sum(axis=1)
    t["고위험 이상"] = t["고위험군"] + t["최고위험군"]
    for r in RISKS:
        t[r + "(%)"] = t[r] / t["계"] * 100
    t["고위험 이상(%)"] = t["고위험 이상"] / t["계"] * 100
    sx = pd.crosstab(df["부서_분석"], df["성"])
    t["남"] = sx.get("남", 0)
    t["여"] = sx.get("여", 0)
    t["평균연령"] = df.groupby("부서_분석")["연령"].mean().round(1)
    t["분석대상"] = np.where(t["계"] >= min_n, "O", "X(인원 적음)")
    return t.sort_values(["고위험 이상(%)", "계"], ascending=[False, False])


def top_depts(t, risk, n, basis, min_n):
    s = t[(t["계"] >= min_n) & (t[risk] > 0)].copy()
    if basis == "비율":
        s = s.sort_values([risk + "(%)", risk, "계"], ascending=False)
    else:
        s = s.sort_values([risk, risk + "(%)", "계"], ascending=False)
    return s.head(n)


# ----------------------------------------------------------------------------
# 4. 관리방향 문구 생성 (규칙 기반 초안)
# ----------------------------------------------------------------------------
def summary_paragraphs(A):
    rg = A["risk_gender"]
    n = A["n"]
    hi = rg.loc["계", "고위험군"] + rg.loc["계", "최고위험군"]
    out = [
        f"{A['sites']} 근로자 {n}명(남 {A['n_m']}명, 여 {A['n_f']}명)을 대상으로 뇌심혈관질환 발병위험도평가를 실시한 결과, "
        f"저위험군 {rg.loc['계','저위험군']}명({pct(rg.loc['계','저위험군'], n):.1f}%), "
        f"중등도위험군 {rg.loc['계','중등도위험군']}명({pct(rg.loc['계','중등도위험군'], n):.1f}%), "
        f"고위험군 {rg.loc['계','고위험군']}명({pct(rg.loc['계','고위험군'], n):.1f}%), "
        f"최고위험군 {rg.loc['계','최고위험군']}명({pct(rg.loc['계','최고위험군'], n):.1f}%)으로 분류되었다. "
        f"고위험군 이상은 {hi}명({pct(hi, n):.1f}%)으로, KOSHA GUIDE H-200-2018에 따라 의사의 업무적합성평가가 필요한 대상이다."
    ]
    if A["n_m"] and A["n_f"]:
        hm = rg.loc["남", "고위험군"] + rg.loc["남", "최고위험군"]
        hf = rg.loc["여", "고위험군"] + rg.loc["여", "최고위험군"]
        out.append(f"성별로는 남성의 고위험군 이상 비율이 {pct(hm, A['n_m']):.1f}%({hm}명), "
                   f"여성이 {pct(hf, A['n_f']):.1f}%({hf}명)였다.")
    pa = A["prev_all"]
    mod = pa[[f for f in FACTORS if f in MODIFIABLE]].sort_values(ascending=False).head(3)
    out.append("개선 가능한 위험요인 중 보유율이 높은 항목은 " +
               ", ".join(f"{k} {v:.1f}%" for k, v in mod.items()) + " 순이었다.")
    return out


def overall_recommendations(A, tops):
    rg = A["risk_gender"]
    n = A["n"]
    recs = []
    b = []
    for r in RISKS[::-1]:
        c = int(rg.loc["계", r])
        b.append(f"{r} {c}명({pct(c, n):.1f}%): {TIER_ACTION[r]}")
    recs.append(("위험군별 사후관리 원칙 (KOSHA GUIDE H-200-2018 <표 10> 기준)", b))

    pa = A["prev_all"]
    mod = [(f, pa[f]) for f in FACTORS if f in MODIFIABLE]
    mod.sort(key=lambda x: -x[1])
    b = []
    for f, v in mod[:4]:
        b.append(f"{f} 보유율 {v:.1f}% → {ACTIONS[f]}")
    b.append("사업장 근로자의 보편적 생활습관 문제를 파악하여 금연·영양·운동·절주 프로그램을 사업장에서 직접 운영하거나 외부기관 지원을 받아 운영(지침 9.2).")
    recs.append(("사업장 공통 생활습관 개선 프로그램", b))

    ex = A["extra"]
    b = []
    if ex["고혈압 1·2기 해당"]:
        b.append(f"혈압 1기 이상 {ex['고혈압 1·2기 해당']}명: 진료 연계 및 치료목표(<표 8-1>) 달성 여부 추적.")
    if ex["당뇨병(과거·현병력/동반질환)"]:
        b.append(f"당뇨병 병력·동반질환 {ex['당뇨병(과거·현병력/동반질환)']}명: 당뇨병이 있으면 1도 고혈압 미만이라도 고위험군 이상으로 분류(지침 7.2(3)), 혈압·지질 동반관리.")
    if ex["단백뇨(표적장기손상)"]:
        b.append(f"단백뇨(표적장기손상) {ex['단백뇨(표적장기손상)']}명: 신장기능·요단백 추적검사 및 전문의 진료 권고.")
    b.append("고혈압·이상지질혈증·당뇨 등 기초질환이 2가지 이상 병합된 근로자는 발병위험이 가중되므로 치료방침 결정 시 함께 고려(지침 9.1(3)).")
    recs.append(("기초질환 관리", b))

    hi = int(rg.loc["계", "고위험군"] + rg.loc["계", "최고위험군"])
    recs.append(("작업관리 및 업무적합성평가", [
        f"고위험군 이상 {hi}명은 의사인 보건관리자(또는 직업환경의학전문의)의 업무적합성평가를 받아야 함(지침 10.1(2)).",
        "장시간근무·교대·야간작업 및 고열·한랭·소음·유해화학물질 노출 여부를 조사하고, <표 11>에 해당하는 경우 정밀 업무적합성평가(<표 12> 서식)를 실시.",
        "작업전환이 필요한 경우 반드시 의사 의견을 듣고 근로자에게 충분히 설명한 후 동의를 받아 시행하며, 평가결과를 이유로 고용상 불이익을 주어서는 안 됨(지침 10.3).",
    ]))
    recs.append(("평가주기 및 개인정보 보호", [
        "기본 평가주기는 2년 1회이나, 개선이 필요한 생활습관 보유자는 연 1회 이상, 기초질환자는 의사 권고에 따라 단축(지침 5.2).",
        "근무상 조치 결정 시 관계자(인사담당자·관리감독자 등)에게 제공하는 건강정보는 최소한으로 제한(지침 11.1(4)). 본 보고서의 부서별 결과는 집단 수준 자료이며 개인 식별 정보를 포함하지 않는다.",
    ]))
    return recs


def dept_bullets(dname, sub, flags_sub, A, risk, n_case, n_total_dept):
    pa = A["prev_all"]
    pd_ = flags_sub[FACTORS].mean() * 100
    all_cnt = int(A["risk_gender"].loc["계", risk])
    b = [f"{risk} {n_case}명 (부서 {n_total_dept}명 중 {pct(n_case, n_total_dept):.1f}%, 전사 {pct(all_cnt, A['n']):.1f}%) → {TIER_ACTION[risk]}."]
    ex = []
    for f in FACTORS:
        if f in MODIFIABLE:
            ex.append((f, pd_[f], pa[f], pd_[f] - pa[f]))
    ex.sort(key=lambda x: -x[3])
    picked = [x for x in ex if x[3] >= 5][:3]
    if picked:
        for f, v, a, d in picked:
            b.append(f"{f}: 부서 {v:.0f}% (전사 {a:.0f}%, +{d:.0f}%p) → {ACTIONS[f]}")
    else:
        top = sorted(ex, key=lambda x: -x[1])[:2]
        b.append("전사 평균을 뚜렷이(+5%p 이상) 상회하는 개선 가능 요인은 없음. 부서 내 보유율이 높은 "
                 + ", ".join(f"{f} {v:.0f}%" for f, v, _, _ in top) + " 중심으로 개인별 맞춤 상담 권장.")
    if risk in ("최고위험군", "고위험군"):
        b.append("해당 부서의 근무형태(장시간·교대·야간)와 작업환경(고열·한랭·소음·유해화학물질)을 현장 확인하여 <표 11>에 해당하면 정밀 업무적합성평가 실시.")
    if n_total_dept < 10:
        b.append("부서 인원이 적어 비율이 크게 변동할 수 있으므로 해석에 유의(개인 식별 방지를 위해 보고서에는 집단 수치만 기재).")
    return b


# ----------------------------------------------------------------------------
# 5. 그래프 (matplotlib → PNG)
# ----------------------------------------------------------------------------
def _save(fig, path):
    fig.tight_layout()
    fig.savefig(path, dpi=170)
    plt.close(fig)
    return path


def fig_risk_gender(A, path):
    rg = A["risk_gender"]
    fig, axs = plt.subplots(1, 2, figsize=(9, 3.6), gridspec_kw={"width_ratios": [1, 1.25]})
    vals = [rg.loc["계", r] for r in RISKS]
    axs[0].pie(vals, labels=[f"{r}\n{v}명" for r, v in zip(RISKS, vals)], colors=[RISK_COLOR[r] for r in RISKS],
               autopct=lambda p: f"{p:.1f}%" if p > 2 else "", startangle=90, counterclock=False,
               wedgeprops={"width": 0.45, "edgecolor": "white"}, textprops={"fontsize": 8}, pctdistance=0.78)
    axs[0].set_title("전체 위험군 분포", fontsize=10)
    x = np.arange(len(RISKS))
    m = [pct(rg.loc["남", r], A["n_m"]) for r in RISKS]
    f = [pct(rg.loc["여", r], A["n_f"]) for r in RISKS]
    axs[1].bar(x - 0.2, m, 0.4, label=f"남 ({A['n_m']}명)", color="#4472C4")
    axs[1].bar(x + 0.2, f, 0.4, label=f"여 ({A['n_f']}명)", color="#ED7D31")
    for i, (a, b) in enumerate(zip(m, f)):
        axs[1].text(i - 0.2, a + 0.5, f"{a:.1f}", ha="center", fontsize=7)
        axs[1].text(i + 0.2, b + 0.5, f"{b:.1f}", ha="center", fontsize=7)
    axs[1].set_xticks(x)
    axs[1].set_xticklabels([r.replace("위험군", "") for r in RISKS], fontsize=9)
    axs[1].set_ylabel("성별 내 구성비 (%)", fontsize=8)
    axs[1].set_title("성별 위험군 구성비", fontsize=10)
    axs[1].legend(fontsize=8)
    axs[1].spines[["top", "right"]].set_visible(False)
    return _save(fig, path)


def fig_age(A, path):
    t = A["age_risk"]
    tot = t.sum(axis=1).replace(0, np.nan)
    p = t.div(tot, axis=0) * 100
    fig, ax = plt.subplots(figsize=(7.5, 3.2))
    left = np.zeros(len(p))
    for r in RISKS:
        ax.barh(p.index.astype(str), p[r].fillna(0), left=left, color=RISK_COLOR[r], label=r)
        left += p[r].fillna(0).values
    for i, (lab, tt) in enumerate(zip(p.index, t.sum(axis=1))):
        ax.text(101, i, f"{tt}명", va="center", fontsize=8)
    ax.set_xlim(0, 112)
    ax.invert_yaxis()
    ax.set_xlabel("연령대 내 구성비 (%)", fontsize=8)
    ax.set_title("연령대별 위험군 구성", fontsize=10)
    ax.legend(ncol=4, fontsize=7, loc="lower center", bbox_to_anchor=(0.5, -0.38))
    ax.spines[["top", "right"]].set_visible(False)
    return _save(fig, path)


def fig_factors(A, path):
    pa = A["prev_all"].reindex(FACTORS)
    fig, ax = plt.subplots(figsize=(7.5, 4.2))
    y = np.arange(len(FACTORS))
    ax.barh(y - 0.2, pa.values, 0.4, color="#4472C4", label="전체")
    sub = A.get("prev_hi")
    if sub is not None:
        ax.barh(y + 0.2, sub.reindex(FACTORS).values, 0.4, color="#D64545", label="고위험군 이상")
        for i, v in enumerate(sub.reindex(FACTORS).values):
            ax.text(v + 0.8, i + 0.2, f"{v:.0f}", va="center", fontsize=7)
    for i, v in enumerate(pa.values):
        ax.text(v + 0.8, i - 0.2, f"{v:.0f}", va="center", fontsize=7)
    ax.set_yticks(y)
    ax.set_yticklabels(FACTORS, fontsize=8)
    ax.invert_yaxis()
    ax.set_xlim(0, 105)
    ax.set_xlabel("보유율 (%)", fontsize=8)
    ax.set_title("위험요인 보유율", fontsize=10)
    ax.legend(fontsize=8, loc="lower right")
    ax.spines[["top", "right"]].set_visible(False)
    return _save(fig, path)


def fig_dept_overview(t, min_n, path, max_rows=30):
    s = t[t["계"] >= min_n].sort_values("고위험 이상(%)", ascending=True).tail(max_rows)
    h = max(3.0, 0.27 * len(s) + 1.3)
    fig, ax = plt.subplots(figsize=(8, h))
    left = np.zeros(len(s))
    for r in RISKS:
        v = s[r + "(%)"].values
        ax.barh(s.index, v, left=left, color=RISK_COLOR[r], label=r)
        left += v
    for i, (idx, row) in enumerate(s.iterrows()):
        ax.text(101, i, f"{int(row['계'])}명 / 고위험이상 {int(row['고위험 이상'])}명", va="center", fontsize=7)
    ax.set_xlim(0, 140)
    ax.set_xticks([0, 20, 40, 60, 80, 100])
    ax.tick_params(axis="y", labelsize=7)
    ax.set_xlabel("부서 내 구성비 (%)", fontsize=8)
    ax.set_title(f"부서별 위험군 구성 (인원 {min_n}명 이상, 고위험군 이상 비율 순)", fontsize=10)
    ax.legend(ncol=4, fontsize=7, loc="lower center", bbox_to_anchor=(0.45, -0.12 - 1.2 / h))
    ax.spines[["top", "right"]].set_visible(False)
    return _save(fig, path)


def fig_top_summary(top, risk, path):
    fig, ax = plt.subplots(figsize=(7, 0.55 * len(top) + 1.4))
    ax.barh(top.index[::-1], top[risk].values[::-1], color=RISK_COLOR[risk])
    for i, (idx, row) in enumerate(list(top.iloc[::-1].iterrows())):
        ax.text(row[risk] + 0.05, i, f"{int(row[risk])}명 ({row[risk + '(%)']:.1f}%)", va="center", fontsize=8)
    ax.set_xlim(0, max(top[risk].max() * 1.35, 1))
    ax.set_xlabel("인원(명)", fontsize=8)
    ax.set_title(f"{risk} 상위 {len(top)}개 부서", fontsize=10)
    ax.xaxis.get_major_locator().set_params(integer=True)
    ax.spines[["top", "right"]].set_visible(False)
    return _save(fig, path)


def fig_dept_detail(dname, sub, flags_sub, A, path):
    fig, axs = plt.subplots(1, 2, figsize=(9, 3.3), gridspec_kw={"width_ratios": [1, 1.5]})
    n = len(sub)
    dep = [pct((sub["위험군"] == r).sum(), n) for r in RISKS]
    allp = [pct(A["risk_gender"].loc["계", r], A["n"]) for r in RISKS]
    for xi, (lab, vals) in enumerate([(f"{dname}\n({n}명)", dep), (f"전체\n({A['n']}명)", allp)]):
        bot = 0
        for r, v in zip(RISKS, vals):
            axs[0].bar(xi, v, bottom=bot, color=RISK_COLOR[r], label=r if xi == 0 else None, width=0.55)
            if v >= 6:
                axs[0].text(xi, bot + v / 2, f"{v:.0f}", ha="center", va="center", fontsize=7, color="white")
            bot += v
    axs[0].set_xticks([0, 1])
    axs[0].set_xticklabels([f"{dname}\n({n}명)", f"전체\n({A['n']}명)"], fontsize=7)
    axs[0].set_ylabel("%", fontsize=8)
    axs[0].set_title("위험군 구성", fontsize=9)
    axs[0].spines[["top", "right"]].set_visible(False)
    facs = [f for f in FACTORS if f in MODIFIABLE]
    pdv = (flags_sub[facs].mean() * 100).values
    pav = A["prev_all"][facs].values
    y = np.arange(len(facs))
    axs[1].barh(y - 0.2, pdv, 0.4, color="#D64545", label="해당 부서")
    axs[1].barh(y + 0.2, pav, 0.4, color="#9AA5B1", label="전사")
    axs[1].set_yticks(y)
    axs[1].set_yticklabels(facs, fontsize=7)
    axs[1].invert_yaxis()
    axs[1].set_xlim(0, 105)
    axs[1].set_title("개선 가능 위험요인 보유율 (%)", fontsize=9)
    axs[1].legend(fontsize=7, loc="lower right")
    axs[1].spines[["top", "right"]].set_visible(False)
    return _save(fig, path)


# ----------------------------------------------------------------------------
# 6. Word 보고서
# ----------------------------------------------------------------------------
def _font(run, size=None, bold=None, color=None):
    run.font.name = DOC_FONT
    rpr = run._element.get_or_add_rPr()
    rpr.rFonts.set(qn("w:eastAsia"), DOC_FONT)
    if size:
        run.font.size = Pt(size)
    if bold is not None:
        run.font.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def _style_font(style, size=None, bold=None, color=None):
    style.font.name = DOC_FONT
    style.element.get_or_add_rPr()
    style.element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT)
    if size:
        style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if color:
        style.font.color.rgb = RGBColor.from_string(color)


def _shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcPr.append(shd)


def new_doc():
    d = Document()
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2.0)
    sec.top_margin = sec.bottom_margin = Cm(2.0)
    _style_font(d.styles["Normal"], 10)
    d.styles["Normal"].paragraph_format.space_after = Pt(4)
    _style_font(d.styles["Heading 1"], 15, True, "1F3864")
    _style_font(d.styles["Heading 2"], 12, True, "2F5496")
    _style_font(d.styles["Heading 3"], 10.5, True, "404040")
    return d


def para(d, text, size=10, bold=False, align=None, color=None, space_after=None, italic=False):
    p = d.add_paragraph()
    r = p.add_run(text)
    _font(r, size, bold, color)
    r.italic = italic
    if align == "center":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    return p


def bullet(d, text, level=0):
    p = d.add_paragraph(style="List Bullet")
    p.paragraph_format.left_indent = Cm(0.6 + 0.6 * level)
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run(text)
    _font(r, 9.5)
    return p


def head(d, text, lvl):
    h = d.add_heading(level=lvl)
    r = h.add_run(text)
    _font(r)
    return h


def table(d, header, rows, widths=None, size=8.5, fills=None, first_bold=True):
    t = d.add_table(rows=1, cols=len(header))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]
        c.text = ""
        r = c.paragraphs[0].add_run(str(h))
        _font(r, size, True)
        c.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        _shade(c, "D9E2F3")
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            r = cells[i].paragraphs[0].add_run(str(v))
            _font(r, size, bool(first_bold and i == 0))
            cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER if i else WD_ALIGN_PARAGRAPH.LEFT
            if fills and fills.get((ri, i)):
                _shade(cells[i], fills[(ri, i)])
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    d.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def pic(d, path, width=16.0, caption=None):
    d.add_picture(str(path), width=Cm(width))
    d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    if caption:
        para(d, caption, 8.5, align="center", color="595959")


def build_docx(path, df, flags, A, dstat, selections, opts, tmp, log):
    d = new_doc()
    n = A["n"]
    title = opts.get("title") or "뇌심혈관질환 발병위험도평가 총괄평가서"
    # ---- 표지 영역
    para(d, "", 40)
    para(d, title, 22, True, "center", "1F3864", 10)
    para(d, A["sites"], 14, False, "center", "404040", 4)
    para(d, f"평가기간: {A['period']}", 11, False, "center", "404040", 4)
    para(d, f"작성일: {dt.date.today():%Y년 %m월 %d일}", 11, False, "center", "404040", 4)
    if opts.get("author"):
        para(d, f"작성: {opts['author']}", 11, False, "center", "404040", 4)
    para(d, "", 20)
    para(d, "※ 본 보고서의 '관리방향'은 평가 결과(인원·위험요인 보유율)에 근거한 자동 작성 초안이며, "
            "직업환경의학전문의(의사인 보건관리자)의 검토·확정을 전제로 합니다.", 9, False, None, "7F7F7F", italic=True)
    d.add_page_break()

    # ---- 1. 일반사항
    head(d, "1. 일반사항", 1)
    table(d, ["구분", "내용"], [
        ["사업장명", A["sites"]],
        ["평가인원", f"총 {n}명 (남자 {A['n_m']}명 / 여자 {A['n_f']}명)"],
        ["평가기간", A["period"]],
        ["평가기준", "KOSHA GUIDE H-200-2018 (종합조사표 <표 3>~<표 7>) 및 일반건강검진 10년 심뇌혈관질환 발병위험도 중 높은 군"],
        ["위험군 분류", "저위험군 / 중등도위험군 / 고위험군 / 최고위험군"],
        ["부서 구분 기준", "팀 단위(괄호 내 세부구분 합산)" if opts["level"] == "team" else "세부부서 단위(근무형태 표기 합산)"],
    ], widths=[3.5, 13.5])
    head(d, "요약", 2)
    for t in summary_paragraphs(A):
        para(d, t, 10)

    # ---- 2. 평가결과
    head(d, "2. 평가결과", 1)
    head(d, "2-1. 뇌심혈관질환 위험군별 현황", 2)
    rg = A["risk_gender"]
    rows = []
    for r in RISKS:
        rows.append([r, rg.loc["계", r], f"{pct(rg.loc['계', r], n):.1f}",
                     rg.loc["남", r], f"{pct(rg.loc['남', r], A['n_m']):.1f}",
                     rg.loc["여", r], f"{pct(rg.loc['여', r], A['n_f']):.1f}"])
    rows.append(["합계", n, "100.0", A["n_m"], "100.0", A["n_f"], "100.0"])
    fills = {(i, 0): {"저위험군": "E2F0D9", "중등도위험군": "FFF2CC", "고위험군": "FBE5D6", "최고위험군": "F4CCCC"}[r]
             for i, r in enumerate(RISKS)}
    table(d, ["구분", "계(명)", "계(%)", "남(명)", "남 중 %", "여(명)", "여 중 %"], rows,
          widths=[3.2, 2.2, 2.2, 2.2, 2.4, 2.2, 2.4], fills=fills)
    para(d, "※ 계(%)는 전체 평가인원 대비, 남·여 %는 각 성별 인원 대비 구성비.", 8.5, color="595959")
    pic(d, fig_risk_gender(A, tmp / "g_risk.png"), 16.5)

    head(d, "2-2. 업무적합성평가(근무상 조치) 현황", 2)
    wg = A["work_gender"]
    if wg.loc["계"].sum() > 0:
        rows = [[w, wg.loc["계", w], f"{pct(wg.loc['계', w], n):.1f}", wg.loc["남", w], wg.loc["여", w]] for w in WORKS]
        table(d, ["근무상 조치", "계(명)", "계(%)", "남(명)", "여(명)"], rows, widths=[4.5, 2.5, 2.5, 2.5, 2.5])
        rw = A["risk_work"]
        rows = [[r] + [int(rw.loc[r, w]) for w in WORKS] for r in RISKS]
        para(d, "위험군별 근무상 조치 (명)", 9.5, True)
        table(d, ["위험군"] + WORKS, rows, widths=[3.5, 3, 3, 3.5, 3])
    else:
        para(d, "업무적합성평가 정보가 없습니다.", 10)

    head(d, "2-3. 연령대별 위험군 분포", 2)
    pic(d, fig_age(A, tmp / "g_age.png"), 14.5)

    head(d, "2-4. 위험요인 보유 현황", 2)
    para(d, "고위험군 이상 근로자와 전체 근로자의 위험요인 보유율을 비교하였다. (한 사람이 여러 요인을 가질 수 있음)", 9.5)
    pic(d, fig_factors(A, tmp / "g_fac.png"), 14.5)
    pr = A["prev_by_risk"]
    rows = [[f, f"{A['prev_all'][f]:.1f}"] + [f"{pr.loc[f, r]:.1f}" for r in RISKS] for f in FACTORS]
    table(d, ["위험요인", "전체 %"] + [r.replace("위험군", "") + " %" for r in RISKS], rows,
          widths=[5.6, 2.1, 2.2, 2.4, 2.2, 2.4])
    if "nhis" in A:
        para(d, "참고: 일반건강검진(공단) 10년 심뇌혈관질환 발병위험도 구간별 인원", 9.5, True)
        nh = A["nhis"]
        table(d, ["구간", "<1%", "1~<5%", "5~<10%", "≥10%"],
              [["인원(명)"] + [int(nh[r]) for r in RISKS]], widths=[3.5, 3, 3, 3, 3])

    # ---- 3. 부서별 분석
    head(d, "3. 부서별 분석", 1)
    para(d, f"부서별 위험군 구성입니다. 인원 {opts['min_n']}명 미만 부서는 비율 해석의 왜곡과 개인 식별 위험을 줄이기 위해 그래프에서 제외하고, 표에는 표시하였습니다.", 9.5)
    pic(d, fig_dept_overview(dstat, opts["min_n"], tmp / "g_dept.png"), 16.5)
    rows = []
    ds = dstat.sort_values(["고위험 이상", "계"], ascending=False)
    for name, r in ds.head(25).iterrows():
        rows.append([name, int(r["계"]), int(r["저위험군"]), int(r["중등도위험군"]), int(r["고위험군"]),
                     int(r["최고위험군"]), f"{r['고위험 이상(%)']:.1f}"])
    para(d, "고위험군 이상 인원 상위 부서 (최대 25개)", 9.5, True)
    table(d, ["부서", "계", "저", "중등도", "고", "최고", "고위험 이상 %"], rows,
          widths=[6, 1.6, 1.6, 2, 1.6, 1.8, 2.8])

    # ---- 4. 위험군별 상위 부서 상세
    sec_no = 4
    if selections:
        head(d, "4. 위험군별 상위 부서 상세 분석 및 관리방향", 1)
        k = 0
        for risk, nn, basis in selections:
            k += 1
            top = top_depts(dstat, risk, nn, basis, opts["min_n"])
            head(d, f"4-{k}. {risk} 상위 {nn}개 부서 (기준: {basis}, 인원 {opts['min_n']}명 이상 부서)", 2)
            if top.empty:
                para(d, f"해당 조건({risk} 보유 부서, 인원 {opts['min_n']}명 이상)을 만족하는 부서가 없습니다.", 10)
                continue
            rows = [[i + 1, name, int(r["계"]), int(r[risk]), f"{r[risk + '(%)']:.1f}", int(r["고위험 이상"]), f"{r['고위험 이상(%)']:.1f}"]
                    for i, (name, r) in enumerate(top.iterrows())]
            table(d, ["순위", "부서", "부서 인원", f"{risk.replace('위험군','')} 인원", f"{risk.replace('위험군','')} %", "고위험 이상", "고위험 이상 %"],
                  rows, widths=[1.3, 5.5, 2, 2.2, 2, 2.2, 2.6])
            cand = dstat[(dstat["계"] >= opts["min_n"]) & (dstat[risk] > 0)]
            if basis == "인원수" and len(cand) > nn:
                cs = cand.sort_values([risk, risk + "(%)", "계"], ascending=False)
                if cs[risk].iloc[nn - 1] == cs[risk].iloc[nn]:
                    para(d, f"※ {risk} 인원이 같은 부서가 순위 경계에 있어 부서 내 비율, 부서 인원 순으로 순위를 정하였습니다. 필요 시 '상위 개수'를 늘려 확인하세요.", 8.5, color="595959")
            pic(d, fig_top_summary(top, risk, tmp / f"g_top_{k}.png"), 12.5)
            for j, (name, r) in enumerate(top.iterrows()):
                sub = df[df["부서_분석"] == name]
                fsub = flags.loc[sub.index]
                head(d, f"[{j + 1}] {name}  (인원 {len(sub)}명, 남 {int((sub['성']=='남').sum())} / 여 {int((sub['성']=='여').sum())}, 평균연령 {sub['연령'].mean():.1f}세)", 3)
                pic(d, fig_dept_detail(name, sub, fsub, A, tmp / f"g_dept_{k}_{j}.png"), 16)
                para(d, "관리방향", 9.5, True, color="1F3864", space_after=2)
                for btxt in dept_bullets(name, sub, fsub, A, risk, int(r[risk]), len(sub)):
                    bullet(d, btxt)
        sec_no = 5

    # ---- 종합 관리방향
    head(d, f"{sec_no}. 종합 관리방향 제언", 1)
    for ttl, bl in overall_recommendations(A, selections):
        head(d, ttl, 3)
        for btxt in bl:
            bullet(d, btxt)
    sec_no += 1

    # ---- 유의사항
    head(d, f"{sec_no}. 해석상 유의사항", 1)
    for t in [
        "위험군은 개인별 평가 시점의 검진 결과로 분류된 것이며, 부서별 비교는 기술통계로서 부서 간 인과관계를 의미하지 않는다.",
        "인원이 적은 부서는 1~2명의 차이로 비율이 크게 달라질 수 있다.",
        "KOSHA 평가도구는 범주형 위험도 분류로, 일반건강검진 10년 위험도 평가와 분류 결과가 다를 수 있으므로(두 도구 간 일치도 연구 참고) 최종 위험군은 지침 7.2(4)에 따라 높은 군을 채택한다.",
        "근무상 조치는 의사의 업무적합성평가 결과에 따르며, 평가결과를 이유로 고용상 불이익을 주어서는 안 된다.",
    ]:
        bullet(d, t)
    sec_no += 1

    # ---- 참고문헌
    head(d, f"{sec_no}. 참고문헌", 1)
    for i, ref in enumerate(REFERENCES, 1):
        p = d.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.8)
        p.paragraph_format.first_line_indent = Cm(-0.8)
        p.paragraph_format.space_after = Pt(2)
        _font(p.add_run(f"{i}. {ref}"), 9)
    d.save(path)
    log(f"  ✔ Word 저장: {path}")


# ----------------------------------------------------------------------------
# 7. Excel 보고서
# ----------------------------------------------------------------------------
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HFILL = PatternFill("solid", fgColor="D9E2F3")


def xl_write(ws, r0, title, header, rows, widths=None):
    ws.cell(r0, 1, title).font = Font(bold=True, size=11, name=DOC_FONT)
    r = r0 + 1
    for j, h in enumerate(header, 1):
        c = ws.cell(r, j, h)
        c.font = Font(bold=True, name=DOC_FONT)
        c.fill = HFILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    for row in rows:
        r += 1
        for j, v in enumerate(row, 1):
            if isinstance(v, (np.integer,)):
                v = int(v)
            elif isinstance(v, (np.floating,)):
                v = float(v)
            c = ws.cell(r, j, v)
            c.border = BORDER
            c.font = Font(name=DOC_FONT)
            c.alignment = Alignment(horizontal="left" if j == 1 else "center", vertical="center", wrap_text=(j == 1))
            if isinstance(v, float):
                c.number_format = "0.0"
    if widths:
        for j, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(j)].width = max(ws.column_dimensions[get_column_letter(j)].width or 0, w)
    return r + 2, r0 + 1, r  # 다음 시작행, 헤더행, 마지막 행


def stacked_chart(ws, title, hdr_row, last_row, first_col, last_col, anchor, horizontal=False, w=18, h=9):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.title = title
    ch.width, ch.height = w, h
    data = Reference(ws, min_col=first_col, max_col=last_col, min_row=hdr_row, max_row=last_row)
    cats = Reference(ws, min_col=1, min_row=hdr_row + 1, max_row=last_row)
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    colors = [RISK_COLOR[r].lstrip("#") for r in RISKS]
    for s, col in zip(ch.series, colors):
        s.graphicalProperties.solidFill = col
    ws.add_chart(ch, anchor)


def build_xlsx(path, df, flags, A, dstat, selections, opts, log):
    wb = Workbook()
    n = A["n"]
    # --- 요약
    ws = wb.active
    ws.title = "요약"
    ws.column_dimensions["A"].width = 22
    ws.cell(1, 1, opts.get("title") or "뇌심혈관질환 발병위험도평가 총괄평가서").font = Font(bold=True, size=14, name=DOC_FONT)
    info = [("사업장명", A["sites"]), ("평가인원", f"총 {n}명 (남 {A['n_m']} / 여 {A['n_f']})"),
            ("평가기간", A["period"]), ("작성일", f"{dt.date.today():%Y-%m-%d}")]
    for i, (k, v) in enumerate(info, 3):
        ws.cell(i, 1, k).font = Font(bold=True, name=DOC_FONT)
        ws.cell(i, 2, v).font = Font(name=DOC_FONT)
    r = 8
    for t in summary_paragraphs(A):
        ws.cell(r, 1, t).alignment = Alignment(wrap_text=True, vertical="top")
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
        ws.row_dimensions[r].height = 62
        r += 1
    ws.cell(r + 1, 1, "※ 관리방향은 자동 작성 초안이며 의사 검토 후 확정").font = Font(italic=True, color="7F7F7F", name=DOC_FONT)

    # --- 위험군별 현황
    ws = wb.create_sheet("위험군별현황")
    rg = A["risk_gender"]
    rows = []
    for g in ["계", "남", "여"]:
        tot = {"계": n, "남": A["n_m"], "여": A["n_f"]}[g]
        rows.append([g] + [int(rg.loc[g, r]) for r in RISKS] + [tot])
    nxt, hr, lr = xl_write(ws, 1, "1. 성별 위험군 인원(명)", ["구분"] + RISKS + ["합계"], rows, [14] + [14] * 5)
    stacked_chart(ws, "성별 위험군 인원", hr, lr, 2, 5, "I2")
    rows = []
    for g in ["계", "남", "여"]:
        tot = {"계": n, "남": A["n_m"], "여": A["n_f"]}[g]
        rows.append([g] + [round(pct(rg.loc[g, r], tot), 2) for r in RISKS])
    nxt, _, _ = xl_write(ws, nxt, "2. 성별 위험군 구성비(%) - 각 성별 인원 대비", ["구분"] + RISKS, rows)
    rows = [[g] + [int(A["work_gender"].loc[g, w]) for w in WORKS] for g in ["계", "남", "여"]]
    nxt, _, _ = xl_write(ws, nxt, "3. 업무적합성평가(근무상 조치) 인원(명)", ["구분"] + WORKS, rows)
    rows = [[r_] + [int(A["risk_work"].loc[r_, w]) for w in WORKS] for r_ in RISKS]
    nxt, _, _ = xl_write(ws, nxt, "4. 위험군별 근무상 조치(명)", ["위험군"] + WORKS, rows)
    at = A["age_risk"]
    rows = [[str(a)] + [int(at.loc[a, r_]) for r_ in RISKS] + [int(at.loc[a].sum())] for a in at.index]
    nxt, hr, lr = xl_write(ws, nxt, "5. 연령대별 위험군(명)", ["연령대"] + RISKS + ["합계"], rows)
    stacked_chart(ws, "연령대별 위험군", hr, lr, 2, 5, "I20")

    # --- 위험요인
    ws = wb.create_sheet("위험요인")
    pr = A["prev_by_risk"]
    ps = A["prev_by_sex"]
    rows = [[f, int(A["flags_n"][f]), round(A["prev_all"][f], 1)] + [round(pr.loc[f, r_], 1) for r_ in RISKS]
            + [round(ps.loc[f, "남"], 1), round(ps.loc[f, "여"], 1)] for f in FACTORS]
    xl_write(ws, 1, "위험요인 보유 현황 (%)", ["위험요인", "보유(명)", "전체%"] + [r_ + "%" for r_ in RISKS] + ["남%", "여%"],
             rows, [30, 10, 10, 12, 12, 12, 12, 10, 10])

    # --- 부서별
    ws = wb.create_sheet("부서별현황")
    cols = ["부서"] + RISKS + ["계", "고위험 이상", "고위험 이상(%)", "남", "여", "평균연령", "분석대상"]
    rows = [[nm] + [int(r_[x]) for x in RISKS] + [int(r_["계"]), int(r_["고위험 이상"]), round(r_["고위험 이상(%)"], 1),
                                                 int(r_["남"]), int(r_["여"]), r_["평균연령"], r_["분석대상"]]
            for nm, r_ in dstat.iterrows()]
    nxt, hr, lr = xl_write(ws, 1, f"부서별 위험군 인원 (고위험 이상 비율 순, 분석대상: {opts['min_n']}명 이상)", cols, rows, [30] + [12] * 11)
    ws.freeze_panes = "B3"
    top_rows = min(len(rows), 30)
    ch = BarChart()
    ch.type = "bar"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.title = "부서별 위험군 인원 (상위 30개 부서)"
    ch.width, ch.height = 24, max(10, top_rows * 0.55)
    ch.add_data(Reference(ws, min_col=2, max_col=5, min_row=hr, max_row=hr + top_rows), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=hr + 1, max_row=hr + top_rows))
    for s, rk in zip(ch.series, RISKS):
        s.graphicalProperties.solidFill = RISK_COLOR[rk].lstrip("#")
    ch.y_axis.delete = False
    ch.x_axis.delete = False
    ch.x_axis.scaling.orientation = "maxMin"
    ws.add_chart(ch, "P2")

    # --- 상위부서분석
    ws = wb.create_sheet("상위부서분석")
    r0 = 1
    if not selections:
        ws.cell(1, 1, "선택된 상위 부서 분석이 없습니다.")
    for risk, nn, basis in selections:
        top = top_depts(dstat, risk, nn, basis, opts["min_n"])
        rows = [[i + 1, nm, int(r_["계"]), int(r_[risk]), round(r_[risk + "(%)"], 1), int(r_["고위험 이상"]), round(r_["고위험 이상(%)"], 1)]
                for i, (nm, r_) in enumerate(top.iterrows())]
        r0, hr, lr = xl_write(ws, r0, f"{risk} 상위 {nn}개 부서 (기준: {basis})",
                              ["순위", "부서", "부서 인원", f"{risk} 인원", f"{risk} %", "고위험 이상", "고위험 이상 %"], rows,
                              [8, 28, 10, 12, 10, 12, 12])
        if len(rows):
            ch = BarChart()
            ch.type = "bar"
            ch.title = f"{risk} 상위 부서 (명)"
            ch.width, ch.height = 14, 6
            ch.add_data(Reference(ws, min_col=4, min_row=hr, max_row=lr), titles_from_data=True)
            ch.set_categories(Reference(ws, min_col=2, min_row=hr + 1, max_row=lr))
            ch.series[0].graphicalProperties.solidFill = RISK_COLOR[risk].lstrip("#")
            ch.legend = None
            ch.x_axis.scaling.orientation = "maxMin"
            ch.y_axis.delete = False
            ch.x_axis.delete = False
            ws.add_chart(ch, f"J{hr}")
        for j, (nm, r_) in enumerate(top.iterrows()):
            sub = df[df["부서_분석"] == nm]
            fsub = flags.loc[sub.index]
            ws.cell(r0, 1, f"[{nm}] 인원 {len(sub)}명").font = Font(bold=True, name=DOC_FONT)
            r0 += 1
            for btxt in dept_bullets(nm, sub, fsub, A, risk, int(r_[risk]), len(sub)):
                c = ws.cell(r0, 2, "• " + btxt)
                c.alignment = Alignment(wrap_text=True, vertical="top")
                ws.merge_cells(start_row=r0, start_column=2, end_row=r0, end_column=8)
                ws.row_dimensions[r0].height = 32
                r0 += 1
            r0 += 1
        r0 += 8 if len(rows) < 8 else 2

    # --- 관리방향
    ws = wb.create_sheet("관리방향")
    ws.column_dimensions["A"].width = 130
    r0 = 1
    for ttl, bl in overall_recommendations(A, selections):
        ws.cell(r0, 1, ttl).font = Font(bold=True, size=11, name=DOC_FONT)
        r0 += 1
        for btxt in bl:
            c = ws.cell(r0, 1, "• " + btxt)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            ws.row_dimensions[r0].height = 30
            r0 += 1
        r0 += 1

    # --- 부서 매핑 (검증용)
    ws = wb.create_sheet("부서명매핑")
    mp = df.groupby(["부서명", "부서_분석"]).size().reset_index(name="인원")
    xl_write(ws, 1, "원본 부서명 → 분석 부서명 (필요 시 '부서명_매핑.csv'로 수정)", ["원본 부서명", "분석 부서명", "인원"],
             mp.values.tolist(), [34, 26, 8])

    # --- 참고문헌
    ws = wb.create_sheet("참고문헌")
    ws.column_dimensions["A"].width = 150
    ws.cell(1, 1, "참고문헌").font = Font(bold=True, size=11, name=DOC_FONT)
    for i, ref in enumerate(REFERENCES, 2):
        c = ws.cell(i, 1, f"{i - 1}. {ref}")
        c.alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[i].height = 30
    wb.save(path)
    log(f"  ✔ Excel 저장: {path}")


# ----------------------------------------------------------------------------
# 8. 실행 본체
# ----------------------------------------------------------------------------
def unique_path(p):
    p = Path(p)
    if not p.exists():
        return p
    i = 2
    while True:
        q = p.with_name(f"{p.stem}({i}){p.suffix}")
        if not q.exists():
            return q
        i += 1


def generate(src, opts, log=print):
    """누리온 엑셀(.xls/.xlsx/.csv) 한 개로 보고서 생성"""
    src = Path(src)
    log(f"▶ 파일 읽는 중: {src.name}")
    raw = load_table(src, log)
    return _generate_from_raw(raw, src.with_suffix(""), opts, log)


def generate_from_store(root, site, opts, log=print, date_from=None, date_to=None, latest_only=True):
    """프로그램에 저장된 개인별 자료(사업장 폴더)로 보고서 생성. 결과는 <자료폴더>/<사업장>/ 에 저장."""
    import cvd_store
    log(f"▶ 저장 자료 읽는 중: {site}")
    raw = cvd_store.records_to_dataframe(root, site, date_from, date_to, latest_only, log)
    if raw.empty:
        raise ValueError("해당 사업장에 계산 가능한 저장 자료가 없습니다.")
    return _generate_from_raw(raw, Path(root) / site / site, opts, log)


def _generate_from_raw(raw, base, opts, log):
    base = Path(base)
    df, flags = prepare(raw, opts["level"], log)
    if df.empty:
        raise ValueError("평가 대상 자료가 없습니다.")
    log(f"  평가 인원 {len(df)}명 / 분석 부서 {df['부서_분석'].nunique()}개")
    A = analyze(df, flags)
    hi_mask = df["위험군"].isin(["고위험군", "최고위험군"])
    A["prev_hi"] = (flags.loc[hi_mask, FACTORS].mean() * 100) if hi_mask.any() else None
    dstat = dept_table(df, opts["min_n"])
    selections = [(r, int(nn), b) for r, (use, nn, b) in opts["top"].items() if use and nn > 0]
    selections.sort(key=lambda x: RISKS.index(x[0]), reverse=True)
    outs = []
    tmp = Path(tempfile.mkdtemp())
    try:
        if opts.get("word", True):
            p = unique_path(base.with_name(base.name + SUFFIX + ".docx"))
            log("▶ Word 보고서 작성 중 (그래프 생성 포함)…")
            build_docx(p, df, flags, A, dstat, selections, opts, tmp, log)
            outs.append(p)
        if opts.get("excel", True):
            p = unique_path(base.with_name(base.name + SUFFIX + ".xlsx"))
            log("▶ Excel 보고서 작성 중…")
            build_xlsx(p, df, flags, A, dstat, selections, opts, log)
            outs.append(p)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    log("완료!")
    return outs


def default_opts():
    return {"level": "team", "min_n": 5, "word": True, "excel": True, "title": "", "author": "",
            "top": {"최고위험군": (True, 3, "인원수"), "고위험군": (True, 3, "인원수"),
                    "중등도위험군": (False, 3, "인원수"), "저위험군": (False, 3, "인원수")}}


# ----------------------------------------------------------------------------
# 9. GUI (tkinter + 드래그앤드롭)
# ----------------------------------------------------------------------------
def run_gui(initial=None):
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    try:
        from tkinterdnd2 import TkinterDnD, DND_FILES
        root = TkinterDnD.Tk()
        dnd = True
    except Exception:
        root = tk.Tk()
        dnd = False
    root.title("뇌심혈관질환 위험성평가 총괄평가서 자동작성")
    root.geometry("760x640")
    state = {"path": None, "out": []}

    nb = ttk.Notebook(root)
    nb.pack(fill="both", expand=True, padx=8, pady=8)

    # ---- 탭1: 파일
    t1 = ttk.Frame(nb)
    nb.add(t1, text=" ① 파일 선택 ")
    drop = tk.Label(t1, text=("여기에 엑셀 파일(.xls/.xlsx)을\n끌어다 놓으세요" if dnd else
                              "파일 끌어놓기를 쓰려면  pip install tkinterdnd2  후 다시 실행하세요\n(아래 '파일 찾기' 버튼 사용 가능)"),
                    relief="ridge", bd=3, bg="#EEF3FA", fg="#1F3864", font=("맑은 고딕", 13), height=6)
    drop.pack(fill="x", padx=14, pady=14)
    pv = tk.StringVar(value="선택된 파일 없음")
    ttk.Label(t1, textvariable=pv, wraplength=700).pack(anchor="w", padx=14)
    info = tk.StringVar(value="")
    ttk.Label(t1, textvariable=info, foreground="#2F5496", wraplength=700).pack(anchor="w", padx=14, pady=6)

    def set_file(p):
        p = str(p).strip().strip("{}").strip('"')
        if not os.path.isfile(p):
            return
        state["path"] = p
        pv.set("선택한 파일: " + p)
        info.set("")
        out_hint.set(f"저장 위치: {Path(p).parent}  (파일명: {Path(p).stem}{SUFFIX}.docx / .xlsx)")

        def peek():
            try:
                raw = load_table(p, lambda *_: None)
                d, _ = prepare(raw, level_var.get(), lambda *_: None)
                info.set(f"✔ 읽기 성공: {len(d)}명, 부서 {d['부서_분석'].nunique()}개 (현재 '{'팀' if level_var.get()=='team' else '세부부서'}' 기준)")
            except Exception as e:
                info.set(f"⚠ {e}")
        threading.Thread(target=peek, daemon=True).start()

    if dnd:
        for w in (drop, root):
            w.drop_target_register(DND_FILES)
            w.dnd_bind("<<Drop>>", lambda e: set_file(root.tk.splitlist(e.data)[0]))

    def browse():
        p = filedialog.askopenfilename(filetypes=[("엑셀/CSV", "*.xls *.xlsx *.xlsm *.csv"), ("모든 파일", "*.*")])
        if p:
            set_file(p)
    ttk.Button(t1, text="파일 찾기…", command=browse).pack(pady=4)
    out_hint = tk.StringVar(value="")
    ttk.Label(t1, textvariable=out_hint, foreground="#595959", wraplength=700).pack(anchor="w", padx=14, pady=8)

    # ---- 탭2: 상위 부서 설정
    t2 = ttk.Frame(nb)
    nb.add(t2, text=" ② 상위 부서 분석 설정 ")
    ttk.Label(t2, text="위험군별로 '상위 몇 개 부서'를 상세 분석(그래프+관리방향)할지 선택하세요.",
              font=("맑은 고딕", 10, "bold")).grid(row=0, column=0, columnspan=4, sticky="w", padx=12, pady=10)
    for j, h in enumerate(["상세분석", "위험군", "상위 몇 개", "순위 기준"]):
        ttk.Label(t2, text=h).grid(row=1, column=j, padx=10, pady=4)
    top_vars = {}
    dflt = default_opts()["top"]
    for i, r in enumerate(RISKS[::-1]):
        use = tk.BooleanVar(value=dflt[r][0])
        nn = tk.IntVar(value=dflt[r][1])
        bs = tk.StringVar(value=dflt[r][2])
        ttk.Checkbutton(t2, variable=use).grid(row=2 + i, column=0)
        ttk.Label(t2, text=r, foreground=RISK_COLOR[r], font=("맑은 고딕", 10, "bold")).grid(row=2 + i, column=1, sticky="w", padx=10)
        ttk.Spinbox(t2, from_=1, to=30, width=5, textvariable=nn).grid(row=2 + i, column=2)
        ttk.Combobox(t2, values=["인원수", "비율"], width=8, textvariable=bs, state="readonly").grid(row=2 + i, column=3, padx=10)
        top_vars[r] = (use, nn, bs)
    ttk.Label(t2, text="· 인원수: 해당 위험군 인원이 많은 부서 순 / 비율: 부서 내 해당 위험군 비율이 높은 부서 순",
              foreground="#595959").grid(row=7, column=0, columnspan=4, sticky="w", padx=12, pady=(14, 2))
    ttk.Separator(t2).grid(row=8, column=0, columnspan=4, sticky="ew", pady=8)
    level_var = tk.StringVar(value="team")
    ttk.Label(t2, text="부서 구분 단위").grid(row=9, column=0, columnspan=2, sticky="w", padx=12)
    ttk.Radiobutton(t2, text="팀 단위 (괄호 세부구분 합산)", variable=level_var, value="team",
                    command=lambda: state["path"] and set_file(state["path"])).grid(row=9, column=2, columnspan=2, sticky="w")
    ttk.Radiobutton(t2, text="세부부서 단위 (예: 생산1팀(제피실))", variable=level_var, value="detail",
                    command=lambda: state["path"] and set_file(state["path"])).grid(row=10, column=2, columnspan=2, sticky="w")
    minn = tk.IntVar(value=5)
    ttk.Label(t2, text="분석 최소 인원(명)").grid(row=11, column=0, columnspan=2, sticky="w", padx=12, pady=8)
    ttk.Spinbox(t2, from_=1, to=50, width=5, textvariable=minn).grid(row=11, column=2, sticky="w")
    ttk.Label(t2, text="이 인원 미만 부서는 순위·그래프에서 제외 (개인 식별/비율 왜곡 방지)", foreground="#595959").grid(row=11, column=3, sticky="w")

    # ---- 탭3: 출력
    t3 = ttk.Frame(nb)
    nb.add(t3, text=" ③ 출력 옵션 ")
    wv, xv = tk.BooleanVar(value=True), tk.BooleanVar(value=True)
    ttk.Checkbutton(t3, text="Word(.docx) 보고서 생성", variable=wv).grid(row=0, column=0, sticky="w", padx=14, pady=10)
    ttk.Checkbutton(t3, text="Excel(.xlsx) 보고서 생성", variable=xv).grid(row=1, column=0, sticky="w", padx=14)
    tv, av = tk.StringVar(), tk.StringVar()
    ttk.Label(t3, text="보고서 제목 (비우면 기본값)").grid(row=2, column=0, sticky="w", padx=14, pady=(14, 2))
    ttk.Entry(t3, textvariable=tv, width=60).grid(row=3, column=0, sticky="w", padx=14)
    ttk.Label(t3, text="작성자 표기 (예: 직업환경의학과 ○○○) - 선택").grid(row=4, column=0, sticky="w", padx=14, pady=(14, 2))
    ttk.Entry(t3, textvariable=av, width=60).grid(row=5, column=0, sticky="w", padx=14)

    # ---- 하단: 실행/로그
    bottom = ttk.Frame(root)
    bottom.pack(fill="both", padx=8, pady=(0, 8))
    log_box = scrolledtext.ScrolledText(bottom, height=9, state="disabled", font=("맑은 고딕", 9))
    log_box.pack(fill="both", expand=True, pady=4)

    def log(msg):
        def _w():
            log_box.configure(state="normal")
            log_box.insert("end", str(msg) + "\n")
            log_box.see("end")
            log_box.configure(state="disabled")
        root.after(0, _w)

    btn = ttk.Button(bottom, text="▶ 보고서 생성")
    btn.pack(pady=4)

    def get_opts():
        o = {"level": level_var.get(), "min_n": int(minn.get()), "word": wv.get(), "excel": xv.get(),
             "title": tv.get().strip(), "author": av.get().strip(),
             "top": {r: (v[0].get(), int(v[1].get()), v[2].get()) for r, v in top_vars.items()}}
        if not (o["word"] or o["excel"]):
            messagebox.showwarning("출력 형식", "Word 또는 Excel 중 하나는 선택해야 합니다.")
            return None
        return o

    def go(job=None):
        """job=None: 선택한 엑셀 파일로 생성 / job=(root, site): 저장된 개인자료로 생성"""
        if job is None and not state["path"]:
            messagebox.showwarning("파일 없음", "먼저 엑셀 파일을 선택(또는 끌어놓기) 해주세요.")
            nb.select(0)
            return
        o = get_opts()
        if o is None:
            return
        btn.state(["disabled"])

        def work():
            try:
                if job is None:
                    outs = generate(state["path"], o, log)
                else:
                    outs = generate_from_store(job[0], job[1], o, log)
                state["out"] = outs
                root.after(0, lambda: done(outs))
            except Exception as e:
                log(f"✖ 오류: {e}")
                root.after(0, lambda: messagebox.showerror("오류", str(e)))
            finally:
                root.after(0, lambda: btn.state(["!disabled"]))
        threading.Thread(target=work, daemon=True).start()

    def done(outs):
        if messagebox.askyesno("완료", "보고서를 만들었습니다.\n저장 폴더를 열까요?\n\n" + "\n".join(p.name for p in outs)):
            try:
                if sys.platform.startswith("win"):
                    os.startfile(str(outs[0].parent))
                elif sys.platform == "darwin":
                    subprocess.run(["open", str(outs[0].parent)])
                else:
                    subprocess.run(["xdg-open", str(outs[0].parent)])
            except Exception:
                pass
    btn.configure(command=go)

    def open_individual():
        try:
            import cvd_gui_individual as G
        except Exception as e:
            messagebox.showerror("모듈 오류", "cvd_gui_individual.py 등 같은 폴더의 모듈을 찾지 못했습니다.\n" + str(e))
            return
        w = G.IndividualWindow(root, dnd_files=DND_FILES if dnd else None,
                               on_make_report=lambda site, rt: go((rt, site)))
        w.focus_set()
    ttk.Separator(t1).pack(fill="x", padx=14, pady=10)
    ttk.Label(t1, text="※ 개인이 제출한 검진결과지(PDF/사진)를 읽어 계산·저장하려면:").pack(anchor="w", padx=14)
    ttk.Button(t1, text="👤 개인 검진결과지 입력·관리 창 열기 (사업장/부서별 저장)", command=open_individual).pack(anchor="w", padx=14, pady=4)
    if initial:
        set_file(initial)
    root.mainloop()


def parse_top(s):
    o = default_opts()["top"]
    o = {k: (False, v[1], v[2]) for k, v in o.items()}
    for part in s.split(","):
        if not part.strip():
            continue
        bits = part.split(":")
        r = bits[0].strip()
        if r in o:
            nn = int(bits[1]) if len(bits) > 1 else 3
            bs = bits[2] if len(bits) > 2 else "인원수"
            o[r] = (True, nn, bs)
    return o


def main():
    ap = argparse.ArgumentParser(description="뇌심혈관질환 위험성평가 총괄평가서 자동작성")
    ap.add_argument("file", nargs="?", help="개인별 결과 엑셀 파일 (.xls/.xlsx)")
    ap.add_argument("--nogui", action="store_true", help="창 없이 바로 생성")
    ap.add_argument("--top", default="최고위험군:3,고위험군:3", help="예: 최고위험군:3,고위험군:3:비율")
    ap.add_argument("--level", choices=["team", "detail"], default="team")
    ap.add_argument("--min-n", type=int, default=5)
    a = ap.parse_args()
    if a.nogui:
        if not a.file:
            ap.error("--nogui 사용 시 파일 경로가 필요합니다.")
        o = default_opts()
        o.update(level=a.level, min_n=a.min_n, top=parse_top(a.top))
        generate(a.file, o)
    else:
        run_gui(a.file)


if __name__ == "__main__":
    main()
