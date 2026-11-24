# -*- coding: utf-8 -*-
"""
개인 건강검진 결과지(PDF/이미지)에서 뇌심혈관질환 위험도 계산에 필요한 값을 읽어오는 모듈.

  extract_file(path) -> {"fields": {...}, "notes": {필드: 메시지}, "method": "...", "text": "..."}

* 글자가 들어 있는 PDF(병원 출력 PDF)  : pdfplumber 로 글자·표·체크박스(■)까지 읽음  (가장 정확)
* 스캔 PDF / 사진(jpg, png ...)        : Tesseract OCR 사용 (설치되어 있을 때만)  -> 숫자 위주로 읽으며 반드시 눈으로 확인
읽은 값은 '초안'이다. 프로그램 화면에서 사람이 확인·수정한 뒤 저장한다.
"""
import re
import sys
import json
from pathlib import Path

NA = "NA"   # '비해당/미실시' 로 표기된 항목

_NUM = re.compile(r"(?<![\d.])(\d+(?:\.\d+)?)")
_RANGE = re.compile(r"\d+(?:\.\d+)?\s*-\s*\d+(?:\.\d+)?")
_PAREN = re.compile(r"\([^)]*\)")


# ---------------------------------------------------------------- 줄 만들기
def _group_lines(words, tol=3.0):
    words = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lines, cur, y = [], [], None
    for w in words:
        if y is None or abs(w["top"] - y) <= tol:
            cur.append(w)
            y = w["top"] if y is None else y
        else:
            lines.append(cur)
            cur, y = [w], w["top"]
    if cur:
        lines.append(cur)
    out = []
    for l in lines:
        l = sorted(l, key=lambda w: w["x0"])
        out.append({"y": sum(w["top"] for w in l) / len(l), "text": " ".join(w["text"] for w in l)})
    return out


def pdf_pages(path):
    """-> [(page_obj, lines)] ; 글자가 없으면 lines=[]"""
    import pdfplumber
    res = []
    pdf = pdfplumber.open(str(path))
    for p in pdf.pages:
        words = p.extract_words()
        res.append((p, _group_lines(words) if words else []))
    return pdf, res


def ocr_lines_from_image(img):
    import pytesseract

    # GitHub Actions 빌드본에서는 실행파일과 함께 포함된 Tesseract를 우선 사용.
    # 개발 환경에서는 기존 표준 설치 경로를 사용하고, 마지막으로 PATH를 확인한다.
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        bundle_root = Path(sys._MEIPASS)
        bundled_tesseract = bundle_root / "Tesseract-OCR" / "tesseract.exe"
        if bundled_tesseract.exists():
            pytesseract.pytesseract.tesseract_cmd = str(bundled_tesseract)
    else:
        local_tesseract = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
        if local_tesseract.exists():
            pytesseract.pytesseract.tesseract_cmd = str(local_tesseract)

    langs = pytesseract.get_languages(config="")

    lang = "kor+eng" if "kor" in langs else "eng"

    d = pytesseract.image_to_data(
        img,
        lang=lang,
        output_type=pytesseract.Output.DICT,
        config="--psm 6"
    )

    words = []
    for i, t in enumerate(d["text"]):
        if t.strip():
            words.append({
                "text": t,
                "top": d["top"][i],
                "x0": d["left"][i]
            })

    return _group_lines(words, tol=12.0), lang
# ---------------------------------------------------------------- 값 추출 도우미
def num_after(line, label_re):
    """줄에서 라벨 뒤의 첫 수치. 비해당 등이면 NA, 못 찾으면 None."""
    m = re.search(label_re, line)
    if not m:
        return None
    rest = line[m.end():]
    rest = _PAREN.sub(" ", rest)
    rest = _RANGE.sub(" ", rest)
    rest = rest.replace("계 산 값", " ").replace("계산값", " ").strip()
    head = rest[:8]
    mn = _NUM.search(rest)
    mna = re.search(r"비\s*해당|미\s*실시|미\s*검", head)
    if mna and (not mn or mna.start() < mn.start()):
        return NA
    if mn:
        return float(mn.group(1))
    return None


def _first(lines, label_re, skip_re=None):
    for ln in lines:
        t = ln["text"]
        if skip_re and re.search(skip_re, t):
            continue
        v = num_after(t, label_re)
        if v is not None:
            return v
    return None


# ---------------------------------------------------------------- 체크박스(표) 기반: 유질환자
def checked_disease_rows(page):
    """표의 '목표질환' 열 범위 안에서 '■ 유질환자' 가 체크된 질환명 집합 (병원 양식 전용)."""
    out = set()
    try:
        tables = page.find_tables()
    except Exception:
        return out
    words = page.extract_words()
    boxes = [w for w in words if "■" in w["text"]]
    for t in tables:
        for r in t.rows:
            cells = [c for c in r.cells if c]
            if not cells:
                continue
            for c in cells:
                try:
                    txt = (page.crop(c).extract_text() or "")
                except Exception:
                    continue
                for name in ("고혈압", "당뇨병", "이상지질혈증"):
                    if name in txt.replace("\n", "") and c[0] < 170:
                        # 이 행(질환) 전체 높이 안에서 '■ 유질환자' 탐색
                        top, bot = c[1], c[3]
                        for b in boxes:
                            if top - 1 <= b["top"] <= bot + 1:
                                near = [w for w in words if abs(w["top"] - b["top"]) < 4 and 0 < w["x0"] - b["x0"] < 60]
                                joined = "".join(w["text"] for w in sorted(near, key=lambda w: w["x0"]))
                                if joined.startswith("유질환자") or "■유질환자" in (b["text"] + joined):
                                    out.add(name)
    return out


# ---------------------------------------------------------------- 본 해석
def parse(lines, pdf_pages_objs=None):
    f, notes = {}, {}
    texts = [ln["text"] for ln in lines]
    whole = "\n".join(texts)

    # 성명 / 주민번호(앞자리만 사용) / 성별
    m = re.search(r"수검자\s*성명\s*([가-힣]{2,5})", whole)
    if not m:
        m = re.search(r"([가-힣]{2,5})\s*님의\s*검진결과", whole)
    if m:
        f["name"] = m.group(1)
    m = re.search(r"(\d{6})\s*-\s*([1-8])", whole)
    if m:
        yy, mm, dd, g = m.group(1)[:2], m.group(1)[2:4], m.group(1)[4:6], int(m.group(2))
        century = 2000 if g in (3, 4, 7, 8) else 1900
        f["birth"] = f"{century + int(yy):04d}{mm}{dd}"
        f["sex"] = "남" if g in (1, 3, 5, 7) else "여"
        notes["birth"] = "주민등록번호 앞 7자리에서 생년월일·성별만 사용 (주민번호는 저장하지 않음)"
    m = re.search(r"GENDER\s*:\s*([MF])", whole)
    if m and "sex" not in f:
        f["sex"] = "남" if m.group(1) == "M" else "여"
    m = re.search(r"검\s*진\s*일(?:자)?\s*[:：]?\s*(20\d{2})\s*[.\-년]\s*(\d{1,2})\s*[.\-월]\s*(\d{1,2})", whole)
    if m:
        f["exam_date"] = f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"

    # 신장·체중
    for t in texts:
        m = re.search(r"키.*?몸무게.*?(\d{2,3}(?:\.\d)?)\s*/\s*(\d{2,3}(?:\.\d)?)", t)
        if m:
            f["height"], f["weight"] = float(m.group(1)), float(m.group(2))
            break
    if "height" not in f:
        h = _first(lines, r"(?:신\s*장)")
        w = _first(lines, r"(?:체\s*중|몸무게)")
        if isinstance(h, float) and isinstance(w, float):
            f["height"], f["weight"] = h, w
    # 허리둘레
    v = _first(lines, r"허리\s*둘레")
    if isinstance(v, float):
        f["waist"] = v
    # 혈압
    m = re.search(r"(\d{2,3})\s*/\s*(\d{2,3})\s*mmHg", whole)
    if m:
        f["sbp"], f["dbp"] = float(m.group(1)), float(m.group(2))
    else:
        m = re.search(r"혈압\s*(\d{2,3})\s*/\s*(\d{2,3})", whole)
        if m:
            f["sbp"], f["dbp"] = float(m.group(1)), float(m.group(2))
    # 혈액검사
    spec = [("fbs", r"공복\s*혈당(?!장애)", None), ("tc", r"총\s*콜레스테롤", None),
            ("hdl", r"(?:고밀도\s*콜레스테롤|HDL)", r"의심"), ("tg", r"(?<!고)중성지방(?!혈증)", None),
            ("ldl", r"(?:저밀도\s*콜레스테롤|LDL)", r"< ?100"), ("creat", r"크레아티닌", None)]
    for key, lab, skip in spec:
        v = _first(lines, lab, skip)
        if v == NA:
            notes[key] = "결과지에 '비해당/미실시' → 값 없음 (이전 검사 결과가 있으면 직접 입력)"
        elif v is not None:
            f[key] = v
    # 요단백
    for t in texts:
        if "요단백" in t:
            m = re.search(r"■\s*(정상|경계|단백뇨\s*의심)", t)
            if m:
                lab = m.group(1).replace(" ", "")
                f["dipstick"] = "1+까지" if lab in ("정상", "경계") else "2+"
                notes["dipstick"] = ("'단백뇨 의심'은 2+로 가정함 — 실제 요단백 등급(2+/3+) 확인 필요" if lab == "단백뇨의심"
                                     else "결과지 판정: " + lab)
            else:
                m = re.search(r"(3\s*\+|2\s*\+|1\s*\+|\+\+\+|\+\+|±|\+|음성|-)", t.split("요단백", 1)[1])
                if m:
                    g = m.group(1).replace(" ", "")
                    f["dipstick"] = "3+이상" if g in ("3+", "+++") else ("2+" if g in ("2+", "++") else "1+까지")
            break
    # 과거병력 / 약물치료
    for t in texts:
        m = re.search(r"과거\s*병력\s*(.*?)\s*약물\s*치료\s*(.*)$", t)
        if m:
            hist, meds = m.group(1), m.group(2)
            if re.search(r"뇌졸중|뇌경색|뇌출혈|뇌혈관", hist):
                f["dz_cerebro"] = True
            if re.search(r"심근경색|협심증|심혈관|관상동맥|심부전", hist):
                f["dz_cardio"] = True
            if re.search(r"신부전|신장질환|만성신장|콩팥", hist):
                f["dz_kidney"] = True
            if re.search(r"말초혈관", hist):
                f["dz_pad"] = True
            if re.search(r"고혈압|혈압", meds):
                f["tx_htn"] = True
            if re.search(r"당뇨", meds):
                f["tx_dm"] = True
            if re.search(r"이상지질|고지혈|콜레스테롤|스타틴", meds):
                f["tx_lipid"] = True
            if re.search(r"고혈압|당뇨|이상지질|고지혈", hist) and not re.search(r"고혈압|당뇨|이상지질|고지혈|혈압", meds):
                notes["tx_htn"] = "과거병력에 만성질환 기재 — 약물 복용 여부 확인 필요"
            break
    # 판정기관 '유질환자' 체크(표 기반)
    if pdf_pages_objs:
        sick = set()
        for p, _ in pdf_pages_objs:
            sick |= checked_disease_rows(p)
        mp = {"고혈압": "tx_htn", "당뇨병": "tx_dm", "이상지질혈증": "tx_lipid"}
        for name in sick:
            f[mp[name]] = True
            notes[mp[name]] = f"결과지 '{name} 유질환자'에 체크 → 약물치료 중으로 처리 (확인 필요)"

    # 생활습관: (1) 건강인 위험평가 페이지가 있으면 사용, (2) 없으면 '생활습관' 체크박스
    smoke_txt = next((re.split(r"흡연", t, 1)[1] for t in texts if re.match(r"\s*흡연\s", t)), None)
    if smoke_txt is not None:
        s = smoke_txt
        if re.search(r"비\s*흡연|안\s*함|없음", s):
            f["smoke"] = "비흡연"
        elif re.search(r"과거|금연\s*중|끊", s):
            f["smoke"] = "과거흡연"
        else:
            f["smoke"] = "현재흡연"
            notes["smoke"] = "위험평가 페이지의 흡연 현황 기준(전자담배 포함)"
    else:
        flag = any(re.search(r"■\s*금연\s*필요", t) for t in texts)
        hint = bool(re.search(r"금연하십시오|금연이 필요", whole))
        if flag or hint:
            f["smoke"] = "현재흡연"
            notes["smoke"] = "결과지의 '금연 필요' 표시로 추정 — 문진으로 확인 필요"
        else:
            f["smoke"] = "비흡연"
            notes["smoke"] = "결과지에 흡연 정보 없음 → 비흡연으로 가정, 과거흡연 여부 문진 확인 필요"
    act = next((t for t in texts if re.match(r"\s*신체활동\s", t)), None)
    m = re.search(r"주\s*(\d+)\s*회", act) if act else None
    if m:
        f["activity"] = "주3회이상" if int(m.group(1)) >= 3 else "주2회이하"
        notes["activity"] = f"위험평가 페이지: 주 {m.group(1)}회"
    else:
        low = any(re.search(r"■\s*신체활동\s*필요", t) for t in texts)
        f["activity"] = "주2회이하" if low else "주3회이상"
        notes["activity"] = "결과지의 '신체활동 필요' 체크 여부로 추정 — 문진으로 확인 필요"
    alc = next((t for t in texts if re.match(r"\s*음주\s", t)), "")
    if re.search(r"위험음주|폭음|5\s*잔", alc) or re.search(r"■\s*절주\s*필요", whole):
        f["alcohol"] = "5잔이상"
        notes["alcohol"] = "위험음주/절주 필요 표시 기준 (계산에는 영향 없음)"
    return f, notes


# ---------------------------------------------------------------- 진입점
def extract_file(path, log=print):
    path = Path(path)
    ext = path.suffix.lower()
    method, lines, pdf_objs, pdf = "", [], None, None
    if ext == ".pdf":
        pdf, pages = pdf_pages(path)
        allines = [ln for _, ls in pages for ln in ls]
        if len(allines) >= 5:
            lines, pdf_objs, method = allines, pages, "PDF 텍스트"
        else:
            # 스캔 PDF → OCR
            try:
                import pytesseract  # noqa
                import fitz  # PyMuPDF: 스캔 PDF 페이지를 이미지로 변환
                from PIL import Image
                scanned_pdf = fitz.open(str(path))
                try:
                    scale = 250 / 72
                    for page in scanned_pdf:
                        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
                        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        ls, lang = ocr_lines_from_image(img)
                        lines += ls
                finally:
                    scanned_pdf.close()
                method = f"스캔 PDF OCR({lang})"
            except Exception as e:
                raise RuntimeError("스캔본 PDF입니다. OCR을 쓰려면 Tesseract(한국어 포함)와 pytesseract 설치가 필요합니다.\n" + str(e))
    elif ext in (".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"):
        try:
            from PIL import Image
            img = Image.open(path)
            lines, lang = ocr_lines_from_image(img)
            method = f"이미지 OCR({lang})"
        except Exception as e:
            raise RuntimeError("이미지 판독에는 Tesseract OCR(한국어 데이터 포함)과 pytesseract 설치가 필요합니다.\n" + str(e))
    else:
        raise ValueError("지원 형식: PDF, PNG, JPG, TIF")
    fields, notes = parse(lines, pdf_objs)
    if "OCR" in method:
        notes["_ocr"] = "OCR로 읽은 값은 오인식 가능성이 있으므로 모든 수치를 원본과 대조하세요."
    if pdf is not None:
        pdf.close()
    return {"fields": fields, "notes": notes, "method": method, "text": "\n".join(ln["text"] for ln in lines)}


if __name__ == "__main__":
    r = extract_file(sys.argv[1])
    print(r["method"])
    print(json.dumps(r["fields"], ensure_ascii=False, indent=1))
    for k, v in r["notes"].items():
        print(" -", k, ":", v)
