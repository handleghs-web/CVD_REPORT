# -*- coding: utf-8 -*-
"""
개인별 평가자료 저장소  (사업장 > 부서 > 개인 JSON 파일)

    <자료폴더>/<사업장명>/<부서명>/<이름>_<생년월일8자리>_<검진일>.json

* 주민등록번호는 저장하지 않는다. (이름·생년월일·검진수치·문진 입력값만 저장)
* 저장된 파일에는 '입력값(input)'만 보관하고, 결과는 열 때마다 cvd_engine 으로 다시 계산한다.
  (계산식을 고치면 저장 자료에도 자동 반영)
* 폴더째로 백업·복사할 수 있고, 메모장으로도 열어볼 수 있다.
* 건강정보는 민감정보이므로 자료 폴더는 암호화/접근통제가 되는 위치에 두고 사용하세요.
"""
import re
import json
import datetime as dt
from pathlib import Path

import pandas as pd

import cvd_engine as E

CONFIG = Path.home() / ".cvd_report_config.json"
INVALID = re.compile(r'[\\/:*?"<>|]')


def load_config():
    try:
        return json.loads(CONFIG.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(cfg):
    try:
        CONFIG.write_text(json.dumps(cfg, ensure_ascii=False, indent=1), encoding="utf-8")
    except Exception:
        pass


def default_root():
    cfg = load_config()
    if cfg.get("data_root"):
        return Path(cfg["data_root"])
    return Path.home() / "Documents" / "뇌심혈관_평가자료"


def set_root(p):
    cfg = load_config()
    cfg["data_root"] = str(p)
    save_config(cfg)


def _clean(s, default="미지정"):
    s = INVALID.sub("_", str(s or "").strip())
    return s or default


def record_path(root, inp):
    bd = "".join(ch for ch in str(inp.get("birth", "")) if ch.isdigit()) or "00000000"
    ed = "".join(ch for ch in str(inp.get("exam_date", "")) if ch.isdigit()) or "00000000"
    fn = f"{_clean(inp.get('name'), '이름없음')}_{bd}_{ed}.json"
    return Path(root) / _clean(inp.get("site"), "사업장미지정") / _clean(inp.get("dept"), "부서미지정") / fn


def save_record(root, inp):
    """저장. 같은 사람·검진일 자료가 다른 부서 폴더에 있으면 옮긴다. -> Path"""
    root = Path(root)
    p = record_path(root, inp)
    site_dir = root / _clean(inp.get("site"), "사업장미지정")
    if site_dir.exists():
        for old in site_dir.glob(f"*/{p.name}"):
            if old != p:
                old.unlink()
    p.parent.mkdir(parents=True, exist_ok=True)
    data = dict(inp)
    data.pop("resident_no", None)                 # 혹시라도 주민번호 키가 있으면 저장 안 함
    payload = {"input": data, "saved_at": dt.datetime.now().isoformat(timespec="seconds"), "version": 1}
    p.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
    return p


def load_record(path):
    d = json.loads(Path(path).read_text(encoding="utf-8"))
    base = E.default_input()
    base.update(d.get("input", {}))
    return base


def delete_record(path):
    Path(path).unlink()


def list_sites(root):
    root = Path(root)
    return sorted(p.name for p in root.iterdir() if p.is_dir()) if root.exists() else []


def list_depts(root, site):
    p = Path(root) / _clean(site, "사업장미지정")
    return sorted(x.name for x in p.iterdir() if x.is_dir()) if p.exists() else []


def iter_records(root, site=None):
    root = Path(root)
    base = root / _clean(site) if site else root
    if not base.exists():
        return
    for p in sorted(base.rglob("*.json")):
        try:
            yield p, load_record(p)
        except Exception:
            continue


def summary_rows(root, site=None):
    """목록 화면용: (경로, 사업장, 부서, 이름, 검진일, 나이, 위험군, 경고수)"""
    rows = []
    for p, inp in iter_records(root, site):
        res = E.calc(inp)
        rows.append({
            "path": str(p), "사업장": inp.get("site", ""), "부서": inp.get("dept", ""), "이름": inp.get("name", ""),
            "검진일": inp.get("exam_date", ""),
            "나이": res.get("age", "") if res.get("ok") else "",
            "위험군": res["risk"] if res.get("ok") else "(입력 미완성)",
            "비고": ("엑셀과 상이" if res.get("diffs") else "") if res.get("ok") else "; ".join(res.get("errors", [])),
        })
    return rows


def records_to_dataframe(root, site=None, date_from=None, date_to=None, latest_only=True, log=print):
    """저장 자료 → 보고서 프로그램(cvd_report.py)이 읽는 '누리온 엑셀'과 같은 열 구성의 DataFrame."""
    items = []
    skipped = 0
    for p, inp in iter_records(root, site):
        ed = E.parse_date(inp.get("exam_date"))
        if date_from and (not ed or ed < date_from):
            continue
        if date_to and (not ed or ed > date_to):
            continue
        res = E.calc(inp)
        if not res.get("ok"):
            skipped += 1
            continue
        row = E.to_export_row(inp, res)
        row["_key"] = f"{inp.get('name')}|{inp.get('birth')}"
        row["_ed"] = ed or dt.date.min
        row["_path"] = str(p)
        items.append(row)
    if skipped:
        log(f"  ※ 입력이 미완성인 {skipped}건은 제외했습니다.")
    if not items:
        return pd.DataFrame()
    df = pd.DataFrame(items)
    if latest_only:
        df = df.sort_values("_ed").drop_duplicates("_key", keep="last")
    return df.drop(columns=["_key", "_ed"]).reset_index(drop=True)
