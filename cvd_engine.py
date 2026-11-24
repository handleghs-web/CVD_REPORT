# -*- coding: utf-8 -*-
"""
뇌심혈관질환 발병위험도 계산 엔진
=================================
'인더닥터_뇌심혈관질환_위험성평가.xlsx'(문진표_개인/결과표_개인/reference 시트)의 계산 로직을 그대로 옮긴 모듈.
KOSHA GUIDE H-200-2018 의 <표 5>~<표 7> 및 일반건강검진 심뇌혈관질환 위험도 모형(복합상대위험도 × 연령별 절대위험도)을 따른다.

calc(inp, today=None)  ->  결과 dict
  - 기본값은 '지침 기준 보정' 결과 (엑셀의 오류로 보이는 부분을 바로잡음)
  - compat=True 로 부르면 엑셀과 100% 동일(오류 포함)하게 계산 (검증용)
  - 보정 결과와 엑셀 결과가 다르면 res["diffs"] 에 사유가 기록된다.
"""
import math
import datetime as dt

RISKS = ["저위험군", "중등도위험군", "고위험군", "최고위험군"]

# ---- reference 시트 ---------------------------------------------------------
# 성별·연령별 평균 복합상대위험도 (20-24 ~ 75이상, 12구간)
MEAN_RR = {
    "남": [1.83, 1.91, 2.02, 2.08, 2.11, 2.14, 2.19, 2.23, 2.25, 2.28, 2.25, 2.24],
    "여": [1.43, 1.40, 1.44, 1.47, 1.54, 1.64, 1.78, 1.92, 2.08, 2.22, 2.32, 2.40],
}
# 연령별(20~85세) 평균 절대위험도 (10만명당 10년 발생)
ABS_RISK = [99, 138, 178, 217, 256, 295, 334, 374, 449, 524, 600, 675, 751, 873, 995, 1117, 1239,
            1362, 1557, 1752, 1947, 2142, 2337, 2591, 2846, 3101, 3356, 3611, 3920, 4229, 4539,
            4848, 5158, 5573, 5988, 6403, 6818, 7234, 7928, 8622, 9316, 10010, 10705, 11462, 12219,
            12976, 13733, 14491, 15205, 15919, 16633, 17347, 18062, 18380, 18698, 19017, 19335,
            19654, 19972, 20290, 20609, 20927, 21246, 21564, 21882, 22201]
ABS_AGES = list(range(20, 86))
assert len(ABS_RISK) == len(ABS_AGES) == 66

RR = {  # 남, 여
    "bmi": {"남": [1.0, 1.04, 1.15, 1.36, 1.58], "여": [1.0, 1.02, 1.03, 1.15, 1.33]},
    "waist": {"남": [1.0, 1.44], "여": [1.0, 1.52]},
    "smoke": {"남": [1.0, 1.30, 1.60], "여": [1.0, 1.20, 1.60]},
    "bp": {"남": [1.0, 1.25, 1.78, 2.71], "여": [1.0, 1.43, 2.06, 3.20]},
    "bpmed": {"남": [1.0, 1.22], "여": [1.0, 1.22]},
    "activity": {"남": [1.0, 1.20], "여": [1.0, 1.20]},
    "fbs": {"남": [1.0, 1.04, 1.12, 1.27, 1.75], "여": [1.0, 1.03, 1.14, 1.31, 1.80]},
    "dmmed": {"남": [1.0, 1.42], "여": [1.0, 1.42]},
    "tc": {"남": [1.0, 1.15, 1.37], "여": [1.0, 1.07, 1.20]},
    "gfr": {"남": [1.0, 1.34, 2.09], "여": [1.0, 1.34, 2.09]},
    "dip": {"남": [1.0, 1.69, 2.43], "여": [1.0, 1.69, 2.43]},
}

WORK_FACTORS = [
    "주당 60시간 이상의 장시간 노동",
    "고정적인 야간작업",
    "공공의 안전을 책임지는 업무",
    "업무상 실수로 인해 타인의 생명에 영향을 주는 업무",
    "힘이 많이 드는 중노동을 연속적으로 해야하는 작업(건설업,하역업 등)",
    "운전작업,고소작업(부정맥이 있는 경우)",
    "근무일정예측이 어렵거나 시차가 큰 출장이 잦은 업무",
    "용광로작업과 같은 고열작업 또는 한랭작업",
    "과도한 소음에 노출되는 작업",
    "갱내 작업 등 산소가 부족하기 쉬운 곳에서의 작업",
    "이황화탄소,염화탄소류,니트로글리세린,메틸렌 클로라이드,일산화탄소 취급 업무",
]

FITNESS = {1: "통상적인 근무 가능", 2: "조건부 근무",
           3: "조건부 근무 또는 한시적 병가/휴직 권고", 4: "병가 또는 휴직 필요"}
FOLLOWUP = {
    1: "생활습관 개선 및 약물치료 권고합니다. 보건관리자와 상담하세요.",
    2: "생활습관 개선 및 약물치료 필요합니다. 산업보건의사의 업무적합성 평가 권고합니다. 보건관리자와 상담하세요.",
    3: "생활습관 개선 및 약물 치료 필요하며, 산업보건의사와 한시적 병가 또는 휴직에 대한 상담 및 업무적합성평가 필요합니다. 필요한 경우 작업환경 개선이 필요합니다.",
    4: "생활습관 개선 및 약물 치료 필요하며, 개선 될때까지 병가 또는 휴직이 필요합니다. 산업보건의사의 업무적합성평가  및 작업환경 개선이 필요합니다.",
}
SWITCH_TEXT = "업무적요인을 고려하였을때, 작업전환을 필요한 업무에 종사하고 있습니다. 작업전환이 필요합니다."

REQUIRED = [("sex", "성별"), ("birth", "생년월일"), ("height", "신장"), ("weight", "체중"),
            ("waist", "허리둘레"), ("sbp", "수축기혈압"), ("dbp", "이완기혈압"), ("fbs", "공복혈당")]


# ---- 입력 정리 --------------------------------------------------------------
def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip().replace(",", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def default_input():
    return {
        "site": "", "dept": "", "name": "", "emp_no": "", "sex": "", "birth": "", "exam_date": "",
        "smoke": "비흡연", "activity": "주3회이상", "alcohol": "",
        "fam_disease": "없음", "fam_who": "", "fam_father_age": "", "fam_mother_age": "", "fam_dm": "",
        "height": "", "weight": "", "waist": "", "sbp": "", "dbp": "", "fbs": "",
        "tc": "", "hdl": "", "tg": "", "ldl": "", "creat": "", "dipstick": "1+까지",
        "tx_htn": False, "tx_dm": False, "tx_lipid": False,
        "dz_cerebro": False, "dz_cardio": False, "dz_kidney": False, "dz_pad": False,
        "to_lvh": False, "to_proteinuria": False, "to_athero": False, "to_retino": False,
        "work": [False] * len(WORK_FACTORS), "note": "", "source": "",
    }


def parse_birth(s):
    s = "".join(ch for ch in str(s) if ch.isdigit())
    if len(s) != 8:
        return None
    try:
        return dt.date(int(s[:4]), int(s[4:6]), int(s[6:8]))
    except ValueError:
        return None


def parse_date(s):
    if isinstance(s, dt.date):
        return s
    d = "".join(ch for ch in str(s) if ch.isdigit())
    if len(d) == 8:
        try:
            return dt.date(int(d[:4]), int(d[4:6]), int(d[6:8]))
        except ValueError:
            return None
    return None


def _band(v, cuts):
    """cuts 오름차순 경계값: v<cuts[0] -> 0, cuts[0]<=v<cuts[1] -> 1 ..."""
    i = 0
    for c in cuts:
        if v >= c:
            i += 1
    return i


def bp_level(sbp, dbp):
    s = _band(sbp, [120, 140, 160])   # 0~3
    d = _band(dbp, [80, 90, 100])
    return max(s, d) + 1              # 1=정상, 2=전단계(120-139/80-89), 3=1기, 4=2기


def bp_label(level):
    return {1: "정상", 2: "고혈압 전단계", 3: "1기 고혈압", 4: "2기 고혈압"}[level]


def nearest_age(value):
    diffs = [abs(g - value) for g in ABS_RISK]
    return ABS_AGES[diffs.index(min(diffs))]       # 동률이면 앞(낮은 나이) = 엑셀 MATCH 와 동일


def lookup_abs(age):
    """VLOOKUP(age, 연령표, 2) 근사일치: age 이하 중 가장 큰 연령."""
    a = min(max(age, 20), 85)
    return ABS_RISK[ABS_AGES.index(a)]


# ---- 본 계산 ----------------------------------------------------------------
def calc(inp, today=None, compat=False, _inner=False):
    """inp: default_input() 형식의 dict. today: 나이 계산 기준일(기본: 검진일, 없으면 오늘)."""
    errs, warns, diffs = [], [], []
    sex = inp.get("sex") or ""
    bd = parse_birth(inp.get("birth"))
    ref = today or parse_date(inp.get("exam_date")) or dt.date.today()
    v = {k: _num(inp.get(k)) for k in ["height", "weight", "waist", "sbp", "dbp", "fbs", "tc", "hdl", "tg", "ldl", "creat"]}
    for key, label in REQUIRED:
        val = bd if key == "birth" else (sex if key == "sex" else v.get(key))
        if val in (None, ""):
            errs.append(f"{label} 미입력")
    if sex not in ("남", "여") and "성별 미입력" not in errs:
        errs.append("성별은 '남' 또는 '여'")
    if errs:
        return {"ok": False, "errors": errs, "warnings": warns, "diffs": diffs}

    age = int((ref - bd).days / 365.25)               # 엑셀: INT((TODAY()-생일)/365.25)
    if age < 20:
        warns.append(f"만 {age}세: 평가표 연령범위(20세 이상) 밖 → 20세 구간으로 계산")
    band = min(max(age, 20), 200)
    band_idx = min((max(age, 20) - 20) // 5, 11)
    s = sex
    bmi = v["weight"] / ((v["height"] / 100) ** 2)

    # --- 위험요인별 상대위험도
    bmi_i = _band(bmi, [25, 26.5, 28, 30])
    bmi_rr = RR["bmi"][s][bmi_i]
    if compat and bmi_i == 4:
        bmi_rr = 0.0                                  # 엑셀 K20 오류(30이상 → F14 참조): 해당 항목 0 처리
    waist_i = 0 if v["waist"] < (90 if s == "남" else 85) else 1
    waist_rr = RR["waist"][s][waist_i]
    obes_rr = max(bmi_rr, waist_rr)
    bpl = bp_level(v["sbp"], v["dbp"])
    bp_rr = RR["bp"][s][bpl - 1]
    bpmed_rr = RR["bpmed"][s][1 if inp.get("tx_htn") else 0]
    fbs_i = _band(v["fbs"], [100, 110, 126, 140])
    fbs_rr = RR["fbs"][s][fbs_i]
    dmmed_rr = RR["dmmed"][s][1 if inp.get("tx_dm") else 0]
    if v["tc"] is None:
        tc_rr = 1.0
        warns.append("총콜레스테롤 미입력 → 위험도 1.00 적용(과소평가 가능, 이전 검사 결과가 있으면 입력)")
    else:
        tc_rr = RR["tc"][s][_band(v["tc"], [200, 240])]
    if v["creat"] is None or v["creat"] <= 0:
        egfr = None
        gfr_rr = 1.0
        warns.append("혈중 크레아티닌 미입력 → 신기능 위험도 1.00 적용")
    else:
        egfr = 175 * (v["creat"] ** -1.154) * (age ** -0.203) * (0.742 if s == "여" else 1.0)
        gfr_rr = RR["gfr"][s][0 if egfr >= 60 else (1 if egfr >= 45 else 2)]
    dip = inp.get("dipstick") or "1+까지"
    dip_i = {"1+까지": 0, "2+": 1, "3+이상": 2}.get(dip, 0)
    dip_rr = RR["dip"][s][dip_i]
    kid_rr = max(gfr_rr, dip_rr)
    smoke_i = {"비흡연": 0, "과거흡연": 1, "현재흡연": 2}.get(inp.get("smoke") or "비흡연", 0)
    smoke_rr = RR["smoke"][s][smoke_i]
    act_rr = RR["activity"][s][1 if inp.get("activity") == "주2회이하" else 0]

    items = [obes_rr, bp_rr, bpmed_rr, fbs_rr, dmmed_rr, tc_rr, kid_rr, smoke_rr, act_rr]
    comp = sum(items) - 8
    mean_rr = MEAN_RR[s][band_idx]
    adj = comp / mean_rr
    abs_age = lookup_abs(age)
    abs100k = adj * abs_age
    abs_floor = math.floor(abs100k + 1e-7)
    abs_pct = abs100k / 1000.0
    cv_age = nearest_age(abs_floor)

    corr_comp = (bpmed_rr + dmmed_rr + kid_rr + smoke_rr) - 3
    corr_abs = math.floor((corr_comp / mean_rr) * abs_age + 1e-7)
    corr_age = nearest_age(corr_abs)

    # --- 통합형 위험도 분류 (엑셀 M31:M38)
    m31 = 1 if abs_pct < 1 else (10 if abs_pct < 5 else (100 if abs_pct < 10 else 1000))
    m32 = 1000 if any(inp.get(k) for k in ("dz_cerebro", "dz_cardio", "dz_kidney", "dz_pad")) else 0
    m22 = 1 if (v["sbp"] >= 180 and v["dbp"] >= 110) else 0
    m25 = 1 if (v["tc"] is not None and v["tc"] >= 310) else 0
    k_dip2, k_dip3 = int(dip_i == 1), int(dip_i == 2)
    k_pu = int(bool(inp.get("to_proteinuria")))
    dm_tx = bool(inp.get("tx_dm"))
    m33 = 0
    if dm_tx:
        m33 = 1000 if (m22 + m25 + k_dip2 + k_dip3 + k_pu) >= 1 else 100
    if egfr is None:
        m34 = 0
    else:
        m34 = 1000 if egfr < 30 else (100 if egfr < 60 else 0)
    m35 = 100 if (m22 + m25) >= (2 if compat else 1) else 0
    m36 = 10 if bpl == 4 else (1 if bpl == 3 else 0)
    m_org = 0
    if not compat and any(inp.get(k) for k in ("to_lvh", "to_proteinuria", "to_athero", "to_retino")):
        m_org = 1000 if bpl == 4 else 100             # 지침 <표 6> 2단계 '고혈압성 장기손상'
    total = m31 + m32 + m33 + m34 + m35 + m36 + m_org
    cls = 1 if total < 10 else (2 if total < 100 else (3 if total < 1000 else 4))

    work = list(inp.get("work") or [])
    work_any = any(work[:len(WORK_FACTORS)])
    switch_txt = SWITCH_TEXT if (cls == 4 and work_any) else ""

    # --- 결과표 표시 항목
    obese_i = {0: "정상", 1: "비만", 2: "비만", 3: "고도비만", 4: "고도비만"}[bmi_i]
    htn = bpl >= 3 or bool(inp.get("tx_htn"))
    dm = fbs_i >= 4 or dm_tx
    ldl = v["ldl"]
    dyslip = bool((v["tc"] is not None and v["tc"] >= 240) or (ldl is not None and ldl >= 160) or inp.get("tx_lipid"))
    pa, ma = _num(inp.get("fam_father_age")), _num(inp.get("fam_mother_age"))
    fam_dz = (inp.get("fam_disease") or "없음") != "없음"
    who = inp.get("fam_who") or ""
    fam_early = fam_dz and (("부" in who and pa and 0 < pa < 55) or ("모" in who and ma and 0 < ma < 65))

    res = {
        "ok": True, "errors": [], "warnings": warns, "diffs": diffs,
        "age": age, "bmi": round(bmi, 1), "egfr": None if egfr is None else round(egfr, 1),
        "bp_level": bpl, "bp_label": bp_label(bpl),
        "rr": {"비만(BMI·허리 중 높은 값)": obes_rr, "혈압": bp_rr, "혈압약 복용": bpmed_rr, "공복혈당": fbs_rr,
               "당뇨약 복용": dmmed_rr, "총콜레스테롤": tc_rr, "신기능(eGFR·요단백 중 높은 값)": kid_rr,
               "흡연": smoke_rr, "신체활동": act_rr},
        "comp_rr": round(comp, 4), "mean_rr": mean_rr, "adj_rr": round(adj, 4),
        "abs_100k": abs_floor, "abs_pct": round(abs_pct, 2), "cv_age": cv_age,
        "corr_rr": round(corr_comp, 4), "corr_abs_100k": corr_abs, "corr_age": corr_age,
        "corr_abs_pct": round(corr_abs / 1000.0, 2),
        "class_no": cls, "risk": RISKS[cls - 1], "fitness": FITNESS[cls], "followup": FOLLOWUP[cls],
        "switch": switch_txt, "score_total": total,
        "display": {"비만도": obese_i, "고혈압": "있음" if htn else "없음", "당뇨": "있음" if dm else "없음",
                    "이상지질혈증": "있음" if dyslip else "없음",
                    "신장기능": "정상" if (egfr is None or egfr >= 45) else "저하",
                    "흡연": inp.get("smoke") or "비흡연",
                    "신체활동": "부족" if inp.get("activity") == "주2회이하" else "적절",
                    "뇌심질환가족력": "있음" if fam_early else "없음"},
        "ref_date": ref.isoformat(),
    }
    if compat or _inner:
        return res
    # 엑셀(오류 포함) 결과와 비교
    ex = calc(inp, today=ref, compat=True, _inner=True)
    res["compat_class"] = ex["risk"]
    if ex["risk"] != res["risk"]:
        reasons = []
        if bmi_i == 4:
            reasons.append("BMI 30 이상 위험도(1.58) 반영(엑셀은 누락)")
        if (m22 + m25) == 1:
            reasons.append("단일 위험요인 과도 증가(180/110 이상 또는 총콜 310 이상)를 고위험으로 분류(엑셀은 2개 동시 필요)")
        if m_org:
            reasons.append("표적장기 손상(좌심실비대·단백뇨·죽상동맥경화·고혈압성 망막증) 반영(엑셀은 미반영)")
        diffs.append(f"엑셀 계산은 '{ex['risk']}' → 지침 기준 보정 결과 '{res['risk']}': " + "; ".join(reasons or ["세부 차이"]))
    if abs(ex["abs_pct"] - res["abs_pct"]) > 1e-9:
        diffs.append(f"10년 발병위험도: 엑셀 {ex['abs_pct']}% → 보정 {res['abs_pct']}% (BMI 30 이상 누락 보정)")
    return res


# ---- 누리온 내보내기 형식(보고서 프로그램 입력)으로의 변환 ----------------------------
def risk_factor_text(inp, res):
    """KOSHA <표 5> 위험인자 목록(보고서 프로그램이 쓰는 문자열)."""
    v = {k: _num(inp.get(k)) for k in ["waist", "fbs", "tc", "hdl", "tg", "ldl"]}
    out = []
    sex = inp.get("sex")
    age = res["age"]
    if (sex == "남" and age >= 45) or (sex == "여" and age >= 55):
        out.append("연령(남 45세 이상, 여 55세 이상)")
    if (inp.get("fam_disease") or "없음") != "없음" and res["display"]["뇌심질환가족력"] == "있음":
        out.append("직계가족의 심뇌혈관질환 조기발병")
    if inp.get("smoke") == "현재흡연":
        out.append("흡연")
    if res["bmi"] >= 25 or (v["waist"] is not None and v["waist"] >= (90 if sex == "남" else 85)):
        out.append("비만(BMI 25이상) 또는 복부비만")
    if v["fbs"] is not None and 100 <= v["fbs"] < 126:
        out.append("공복혈당장애(100≤ 공복혈당 <126)")
    if v["hdl"] is not None and v["hdl"] < 40:
        out.append("HDL콜레스테롤(40㎎/㎗ 미만)")
    if (v["tc"] is not None and v["tc"] >= 220) or (v["ldl"] is not None and v["ldl"] >= 150) or (v["tg"] is not None and v["tg"] >= 200):
        out.append("총콜레스테롤≥220㎎/㎗ 또는 LDL≥150 또는 중성지방≥200")
    return ", ".join(out)


def to_export_row(inp, res):
    """보고서 프로그램(cvd_report.py)이 읽는 누리온 엑셀과 같은 열 이름의 dict."""
    work_map = {1: "통상근무자", 2: "조건근무자", 3: "조건근무자(병가/휴직 권고)", 4: "병가 또는 휴직"}
    hx = []
    if inp.get("tx_htn"):
        hx.append("고혈압")
    if inp.get("tx_dm"):
        hx.append("당뇨병")
    if inp.get("tx_lipid"):
        hx.append("이상지질혈증")
    for k, t in (("dz_cerebro", "뇌혈관질환"), ("dz_cardio", "심혈관질환"), ("dz_kidney", "신장질환"), ("dz_pad", "말초혈관질환")):
        if inp.get(k):
            hx.append(t)
    org = [t for k, t in (("to_lvh", "좌심실비대"), ("to_proteinuria", "단백뇨"),
                          ("to_athero", "죽상동맥경화증"), ("to_retino", "고혈압성망막증")) if inp.get(k)]
    alc = {"1잔미만": "비음주", "1~5잔미만": "음주", "5잔이상": "음주"}.get(inp.get("alcohol") or "", "")
    return {
        "사업장명": inp.get("site", ""), "부서명": inp.get("dept", ""), "성명": inp.get("name", ""),
        "성별": "남자" if inp.get("sex") == "남" else "여자", "연령": res["age"],
        "검진일자": inp.get("exam_date", ""), "발병위험도평가": res["risk"],
        "업무적합성평가": work_map[res["class_no"]],
        "위험인자": risk_factor_text(inp, res), "혈압수준": res["bp_label"],
        "흡연유무": "흡연" if inp.get("smoke") == "현재흡연" else ("과거흡연" if inp.get("smoke") == "과거흡연" else "비흡연"),
        "음주여부": alc, "신체활동부족": "운동부족" if inp.get("activity") == "주2회이하" else "",
        "과거및현병력": ",".join(hx), "동반된질병상태": ",".join(hx),
        "표적장기손상여부": ",".join(org),
        "가족력": "있음" if res["display"]["뇌심질환가족력"] == "있음" else "",
        "공단검진위험도": res["abs_pct"] / 100.0,
        "BMI": res["bmi"], "허리둘레": inp.get("waist"), "수축기혈압": inp.get("sbp"), "이완기혈압": inp.get("dbp"),
        "공복혈당": inp.get("fbs"), "총콜레스테롤": inp.get("tc"), "HDL": inp.get("hdl"),
        "중성지방": inp.get("tg"), "LDL": inp.get("ldl"),
        "심뇌혈관나이": res["cv_age"], "교정심뇌혈관나이": res["corr_age"],
    }
