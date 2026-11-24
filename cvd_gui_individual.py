# -*- coding: utf-8 -*-
"""
개인 검진결과 입력·계산·저장 창 (tkinter)

  - 검진결과지(PDF/사진)를 창에 끌어놓거나 [파일 불러오기]로 읽어 값을 자동 채움
  - 값을 고치면 즉시 엑셀(문진표_개인/결과표_개인)과 같은 방식으로 위험도 재계산
  - [저장] → 자료폴더/사업장/부서/ 에 개인별 파일로 저장
  - [저장 목록] 탭에서 사업장·부서별 조회, 수정, 삭제, 보고서 생성
"""
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

import cvd_engine as E
import cvd_store as S

RISK_BG = {"저위험군": "#4C9F70", "중등도위험군": "#F2C14E", "고위험군": "#F08A3C", "최고위험군": "#D64545"}
RISK_FG = {"저위험군": "white", "중등도위험군": "black", "고위험군": "white", "최고위험군": "white"}
NUM_KEYS = ["height", "weight", "waist", "sbp", "dbp", "fbs", "tc", "hdl", "tg", "ldl", "creat",
            "fam_father_age", "fam_mother_age"]
BOOL_KEYS = ["tx_htn", "tx_dm", "tx_lipid", "dz_cerebro", "dz_cardio", "dz_kidney", "dz_pad",
             "to_lvh", "to_proteinuria", "to_athero", "to_retino"]
LABELS = {
    "height": "신장(cm)", "weight": "체중(kg)", "waist": "허리둘레(cm)", "sbp": "수축기혈압", "dbp": "이완기혈압",
    "fbs": "공복혈당(mg/dL)", "tc": "총콜레스테롤", "hdl": "HDL", "tg": "중성지방", "ldl": "LDL",
    "creat": "혈중크레아티닌", "fam_father_age": "부 발병나이", "fam_mother_age": "모 발병나이",
}


class IndividualWindow(tk.Toplevel):
    def __init__(self, master, dnd_files=None, on_make_report=None):
        super().__init__(master)
        self.title("개인 검진결과 입력·계산·저장")
        self.geometry("1240x800")
        self.on_make_report = on_make_report
        self.vars = {}
        self.entries = {}
        self.work_vars = [tk.BooleanVar(value=False) for _ in E.WORK_FACTORS]
        self._after = None
        self._loading = False
        self.current_path = None
        self.root_dir = S.default_root()

        top = ttk.Frame(self)
        top.pack(fill="x", padx=8, pady=(8, 0))
        ttk.Label(top, text="자료 폴더:").pack(side="left")
        self.root_var = tk.StringVar(value=str(self.root_dir))
        ttk.Entry(top, textvariable=self.root_var, width=70, state="readonly").pack(side="left", padx=4)
        ttk.Button(top, text="변경…", command=self.change_root).pack(side="left")
        ttk.Label(top, text="  ※ 건강정보 폴더입니다. 접근이 통제되는 위치를 사용하세요.", foreground="#9C3B3B").pack(side="left")

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=8)
        self.tab_in = ttk.Frame(self.nb)
        self.tab_list = ttk.Frame(self.nb)
        self.nb.add(self.tab_in, text="  ① 입력·계산  ")
        self.nb.add(self.tab_list, text="  ② 저장 목록 / 보고서  ")
        self._build_input()
        self._build_list()
        if dnd_files:
            try:
                self.drop_target_register(dnd_files)
                self.dnd_bind("<<Drop>>", lambda e: self.load_file(self.tk.splitlist(e.data)[0]))
                self.drop_hint.configure(text="여기(창 어디든)에 검진결과지 PDF/사진을 끌어다 놓으면 자동으로 읽습니다")
            except Exception:
                pass
        self.recalc()

    # ------------------------------------------------------------------ 입력 탭
    def _entry(self, parent, key, width=9, row=None, col=None, label=None):
        v = tk.StringVar()
        self.vars[key] = v
        v.trace_add("write", self._changed)
        if label is not None:
            ttk.Label(parent, text=label).grid(row=row, column=col, sticky="e", padx=(6, 2), pady=2)
        e = tk.Entry(parent, textvariable=v, width=width)
        e.grid(row=row, column=(col + 1) if label is not None else col, sticky="w", pady=2)
        self.entries[key] = e
        return e

    def _combo(self, parent, key, values, width=14, row=0, col=0, label=None, state="readonly"):
        v = tk.StringVar()
        self.vars[key] = v
        v.trace_add("write", self._changed)
        if label:
            ttk.Label(parent, text=label).grid(row=row, column=col, sticky="e", padx=(6, 2), pady=2)
        cb = ttk.Combobox(parent, textvariable=v, values=values, width=width, state=state)
        cb.grid(row=row, column=col + 1 if label else col, sticky="w", pady=2)
        return cb

    def _check(self, parent, key, text, row, col, span=1):
        v = tk.BooleanVar(value=False)
        self.vars[key] = v
        v.trace_add("write", self._changed)
        ttk.Checkbutton(parent, text=text, variable=v).grid(row=row, column=col, columnspan=span, sticky="w", padx=6)

    def _radio(self, parent, key, values, row, col, label):
        v = tk.StringVar()
        self.vars[key] = v
        v.trace_add("write", self._changed)
        ttk.Label(parent, text=label).grid(row=row, column=col, sticky="e", padx=(6, 2))
        f = ttk.Frame(parent)
        f.grid(row=row, column=col + 1, columnspan=3, sticky="w")
        for t in values:
            ttk.Radiobutton(f, text=t, value=t, variable=v).pack(side="left", padx=3)

    def _build_input(self):
        pw = ttk.PanedWindow(self.tab_in, orient="horizontal")
        pw.pack(fill="both", expand=True)
        left_outer = ttk.Frame(pw)
        right = ttk.Frame(pw)
        pw.add(left_outer, weight=3)
        pw.add(right, weight=2)
        self.after(150, lambda: pw.sashpos(0, 720))

        cv = tk.Canvas(left_outer, highlightthickness=0, width=700)
        sb = ttk.Scrollbar(left_outer, orient="vertical", command=cv.yview)
        cv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        cv.pack(side="left", fill="both", expand=True)
        form = ttk.Frame(cv)
        win = cv.create_window((0, 0), window=form, anchor="nw")
        form.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        cv.bind("<Configure>", lambda e: cv.itemconfigure(win, width=e.width))
        cv.bind_all("<MouseWheel>", lambda e: cv.yview_scroll(int(-e.delta / 120), "units") if str(self.focus_get()).startswith(str(self)) else None)

        self.drop_hint = ttk.Label(form, text="[파일 불러오기] 버튼으로 검진결과지 PDF/사진을 읽을 수 있습니다", foreground="#2F5496")
        self.drop_hint.grid(row=0, column=0, sticky="w", padx=6, pady=(6, 2))
        bt = ttk.Frame(form)
        bt.grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Button(bt, text="📄 검진결과지 파일 불러오기…", command=self.pick_file).pack(side="left", padx=2)
        ttk.Button(bt, text="💾 저장", command=self.save_current).pack(side="left", padx=2)
        ttk.Button(bt, text="＋ 새로 입력", command=self.new_record).pack(side="left", padx=2)

        # 1. 기본정보
        lf = ttk.LabelFrame(form, text="1. 기본정보")
        lf.grid(row=2, column=0, sticky="ew", padx=6, pady=4)
        self.cb_site = self._combo(lf, "site", [], 22, 0, 0, "사업장", state="normal")
        self.cb_dept = self._combo(lf, "dept", [], 22, 0, 2, "부서", state="normal")
        self._entry(lf, "name", 14, 1, 0, "성명")
        self._radio(lf, "sex", ["남", "여"], 1, 2, "성별")
        self._entry(lf, "birth", 12, 2, 0, "생년월일(8자리)")
        self._entry(lf, "exam_date", 12, 2, 2, "검진일(YYYY-MM-DD)")
        self.vars["site"].trace_add("write", lambda *_: self._refresh_depts())

        # 2. 문진
        lf = ttk.LabelFrame(form, text="2. 문진 (결과지에 없으면 직접 확인)")
        lf.grid(row=3, column=0, sticky="ew", padx=6, pady=4)
        self._radio(lf, "smoke", ["비흡연", "과거흡연", "현재흡연"], 0, 0, "흡연(전자담배 포함)")
        self._radio(lf, "activity", ["주3회이상", "주2회이하"], 1, 0, "신체활동(1회30분↑)")
        self._combo(lf, "alcohol", ["", "1잔미만", "1~5잔미만", "5잔이상"], 12, 2, 0, "음주량(참고)")
        self._combo(lf, "fam_disease", ["없음", "뇌졸중", "협심증", "심근경색"], 12, 3, 0, "부모 심뇌혈관질환")
        self._check(lf, "fam_f", "부", 3, 2)
        self._check(lf, "fam_m", "모", 3, 3)
        self._entry(lf, "fam_father_age", 6, 4, 0, "부 발병나이")
        self._entry(lf, "fam_mother_age", 6, 4, 2, "모 발병나이")

        # 3. 계측·검사
        lf = ttk.LabelFrame(form, text="3. 신체계측 및 검사결과")
        lf.grid(row=4, column=0, sticky="ew", padx=6, pady=4)
        keys = ["height", "weight", "waist", "sbp", "dbp", "fbs", "tc", "hdl", "tg", "ldl", "creat"]
        for i, k in enumerate(keys):
            self._entry(lf, k, 8, i // 3, (i % 3) * 2, LABELS[k])
        self._combo(lf, "dipstick", ["1+까지", "2+", "3+이상"], 8, 3, 4, "요단백(Dipstick)")

        # 4. 치료·질환
        lf = ttk.LabelFrame(form, text="4. 약물 치료 / 질환 / 표적장기 손상")
        lf.grid(row=5, column=0, sticky="ew", padx=6, pady=4)
        ttk.Label(lf, text="치료(약 복용):").grid(row=0, column=0, sticky="e")
        self._check(lf, "tx_htn", "고혈압약", 0, 1)
        self._check(lf, "tx_dm", "당뇨약", 0, 2)
        self._check(lf, "tx_lipid", "이상지질혈증약", 0, 3)
        ttk.Label(lf, text="진단 질환:").grid(row=1, column=0, sticky="e")
        for j, (k, t) in enumerate([("dz_cerebro", "뇌혈관질환"), ("dz_cardio", "심혈관질환"), ("dz_kidney", "신장질환"), ("dz_pad", "말초혈관질환")]):
            self._check(lf, k, t, 1, 1 + j)
        ttk.Label(lf, text="표적장기손상:").grid(row=2, column=0, sticky="e")
        for j, (k, t) in enumerate([("to_lvh", "좌심실비대"), ("to_proteinuria", "단백뇨"), ("to_athero", "죽상동맥경화"), ("to_retino", "고혈압성망막증")]):
            self._check(lf, k, t, 2, 1 + j)

        # 5. 업무
        lf = ttk.LabelFrame(form, text="5. 현재 종사 업무 (해당 항목 선택)")
        lf.grid(row=6, column=0, sticky="ew", padx=6, pady=4)
        for i, (t, v) in enumerate(zip(E.WORK_FACTORS, self.work_vars)):
            v.trace_add("write", self._changed)
            ttk.Checkbutton(lf, text=t, variable=v).grid(row=i, column=0, sticky="w", padx=6)
        lf = ttk.LabelFrame(form, text="메모")
        lf.grid(row=7, column=0, sticky="ew", padx=6, pady=4)
        self._entry(lf, "note", 80, 0, 0)
        form.columnconfigure(0, weight=1)

        # ---- 오른쪽: 결과
        self.risk_lbl = tk.Label(right, text="-", font=("맑은 고딕", 18, "bold"), height=2, bg="#DDDDDD")
        self.risk_lbl.pack(fill="x", padx=6, pady=(6, 2))
        self.out = tk.Text(right, wrap="word", height=26, font=("맑은 고딕", 10), state="disabled", bg="#FAFAFA")
        self.out.pack(fill="both", expand=True, padx=6, pady=2)
        self.out.tag_configure("h", font=("맑은 고딕", 10, "bold"), foreground="#1F3864", spacing1=6)
        self.out.tag_configure("warn", foreground="#8A6D00")
        self.out.tag_configure("diff", foreground="#C55A11")
        self.out.tag_configure("err", foreground="#C00000")
        ttk.Label(right, text="판독 메모 (노란 칸은 확인 필요)").pack(anchor="w", padx=6)
        self.memo = tk.Text(right, wrap="word", height=8, font=("맑은 고딕", 9), state="disabled", bg="#FFFBE6")
        self.memo.pack(fill="x", padx=6, pady=(0, 6))

    # ------------------------------------------------------------------ 목록 탭
    def _build_list(self):
        t = self.tab_list
        bar = ttk.Frame(t)
        bar.pack(fill="x", padx=6, pady=6)
        ttk.Label(bar, text="사업장").pack(side="left")
        self.list_site = tk.StringVar()
        self.cb_list_site = ttk.Combobox(bar, textvariable=self.list_site, width=26, state="readonly")
        self.cb_list_site.pack(side="left", padx=4)
        self.cb_list_site.bind("<<ComboboxSelected>>", lambda e: self.refresh_list())
        ttk.Button(bar, text="새로고침", command=self.refresh_list).pack(side="left", padx=2)
        ttk.Button(bar, text="선택 항목 열기(수정)", command=self.open_selected).pack(side="left", padx=2)
        ttk.Button(bar, text="삭제", command=self.delete_selected).pack(side="left", padx=2)
        ttk.Button(bar, text="누리온 형식 엑셀로 내보내기", command=self.export_xlsx).pack(side="left", padx=2)
        ttk.Button(bar, text="▶ 이 사업장 총괄보고서 만들기", command=self.make_report).pack(side="right", padx=2)
        cols = ["사업장", "부서", "이름", "검진일", "나이", "위험군", "비고"]
        self.tree = ttk.Treeview(t, columns=cols, show="headings", selectmode="browse")
        for c, w in zip(cols, [160, 150, 100, 100, 50, 110, 260]):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True, padx=6, pady=(0, 6))
        self.tree.bind("<Double-1>", lambda e: self.open_selected())
        self.count_lbl = ttk.Label(t, text="")
        self.count_lbl.pack(anchor="w", padx=8, pady=(0, 6))
        self._paths = {}
        self.nb.bind("<<NotebookTabChanged>>", lambda e: self.refresh_list() if self.nb.index("current") == 1 else None)

    # ------------------------------------------------------------------ 값 읽기/쓰기
    def collect(self):
        inp = E.default_input()
        for k, v in self.vars.items():
            if k in ("fam_f", "fam_m"):
                continue
            inp[k] = v.get()
        who = ("부" if self.vars["fam_f"].get() else "") + ("모" if self.vars["fam_m"].get() else "")
        inp["fam_who"] = who
        inp["work"] = [v.get() for v in self.work_vars]
        for k in BOOL_KEYS:
            inp[k] = bool(self.vars[k].get())
        if inp["fam_disease"] == "":
            inp["fam_disease"] = "없음"
        return inp

    def fill(self, inp):
        self._loading = True
        try:
            base = E.default_input()
            base.update(inp)
            for k, v in self.vars.items():
                if k == "fam_f":
                    v.set("부" in (base.get("fam_who") or ""))
                elif k == "fam_m":
                    v.set("모" in (base.get("fam_who") or ""))
                elif k in BOOL_KEYS:
                    v.set(bool(base.get(k)))
                else:
                    val = base.get(k, "")
                    if isinstance(val, float) and val == int(val) and k in ("sbp", "dbp", "fbs", "tc", "hdl", "tg", "ldl"):
                        val = int(val)
                    v.set("" if val is None else str(val))
            w = list(base.get("work") or [])
            for i, v in enumerate(self.work_vars):
                v.set(bool(w[i]) if i < len(w) else False)
        finally:
            self._loading = False
        self._clear_marks()
        self.recalc()

    def _clear_marks(self):
        for e in self.entries.values():
            e.configure(bg="white")

    def _mark(self, keys):
        for k in keys:
            if k in self.entries:
                self.entries[k].configure(bg="#FFF2A8")

    # ------------------------------------------------------------------ 계산/표시
    def _changed(self, *_):
        if self._loading:
            return
        if self._after:
            self.after_cancel(self._after)
        self._after = self.after(200, self.recalc)

    def recalc(self):
        self._after = None
        res = E.calc(self.collect())
        self.last = res
        o = self.out
        o.configure(state="normal")
        o.delete("1.0", "end")
        if not res.get("ok"):
            self.risk_lbl.configure(text="입력 중…", bg="#DDDDDD", fg="black")
            o.insert("end", "필수 항목을 입력하면 자동 계산됩니다.\n\n", "h")
            for e in res["errors"]:
                o.insert("end", f" · {e}\n", "err")
            o.configure(state="disabled")
            return
        self.risk_lbl.configure(text=f"{res['risk']}", bg=RISK_BG[res["risk"]], fg=RISK_FG[res["risk"]])
        o.insert("end", "종합\n", "h")
        o.insert("end", f"만 {res['age']}세 · BMI {res['bmi']}" + (f" · eGFR {res['egfr']}" if res["egfr"] is not None else "") + f" · 혈압 {res['bp_label']}\n")
        o.insert("end", f"10년 이내 발병위험도 {res['abs_pct']}%  (동일 연령·성별 평균 대비 {res['adj_rr']:.2f}배)\n")
        o.insert("end", f"심뇌혈관 나이 {res['cv_age']}세  →  교정 가능 시 {res['corr_age']}세 ({res['corr_abs_pct']}%)\n")
        o.insert("end", "\n업무적합성\n", "h")
        o.insert("end", res["fitness"] + "\n")
        if res["switch"]:
            o.insert("end", res["switch"] + "\n", "err")
        o.insert("end", "\n사후관리\n", "h")
        o.insert("end", res["followup"] + "\n")
        o.insert("end", "\n위험요인별 상대위험도\n", "h")
        for k, v in res["rr"].items():
            o.insert("end", f"  {k}: {v:.2f}\n")
        o.insert("end", f"  복합상대위험도 {res['comp_rr']}\n")
        d = res["display"]
        o.insert("end", "\n결과표 표시\n", "h")
        o.insert("end", "  " + " / ".join(f"{k} {v}" for k, v in d.items()) + "\n")
        if res["warnings"]:
            o.insert("end", "\n확인 필요\n", "h")
            for w in res["warnings"]:
                o.insert("end", f" · {w}\n", "warn")
        if res["diffs"]:
            o.insert("end", "\n엑셀 원본과 다른 점 (지침 기준 보정)\n", "h")
            for w in res["diffs"]:
                o.insert("end", f" · {w}\n", "diff")
        o.configure(state="disabled")

    # ------------------------------------------------------------------ 파일 불러오기
    def pick_file(self):
        p = filedialog.askopenfilename(filetypes=[("검진결과지", "*.pdf *.png *.jpg *.jpeg *.tif *.tiff"), ("모든 파일", "*.*")])
        if p:
            self.load_file(p)

    def load_file(self, path, sync=False):
        import cvd_extract as X
        self._set_memo(f"읽는 중… {Path(path).name}")

        def work():
            try:
                r = X.extract_file(path, log=lambda *_: None)
                if sync:
                    self._apply_extract(r, path)
                else:
                    self.after(0, lambda: self._apply_extract(r, path))
            except Exception as e:
                if sync:
                    raise
                self.after(0, lambda: (self._set_memo("읽기 실패: " + str(e)), messagebox.showerror("읽기 실패", str(e))))
        if sync:
            work()
        else:
            threading.Thread(target=work, daemon=True).start()

    def _apply_extract(self, r, path):
        cur = self.collect()
        keep = {k: cur[k] for k in ("site", "dept", "work", "alcohol", "fam_disease", "fam_who",
                                    "fam_father_age", "fam_mother_age", "note")}
        newinp = E.default_input()
        newinp.update(keep)
        newinp.update(r["fields"])
        newinp["source"] = Path(path).name
        self.current_path = None
        self.fill(newinp)
        lines = [f"판독 방식: {r['method']}  (파일: {Path(path).name})"]
        marks = []
        for k, v in r["notes"].items():
            lines.append(f"· {LABELS.get(k, k)}: {v}")
            if "확인" in v or "비해당" in v or k.startswith("_"):
                marks.append(k)
        missing = [LABELS[k] for k in ("height", "weight", "waist", "sbp", "dbp", "fbs") if k not in r["fields"]]
        if missing:
            lines.append("· 읽지 못한 필수 항목: " + ", ".join(missing))
            marks += [k for k in ("height", "weight", "waist", "sbp", "dbp", "fbs") if k not in r["fields"]]
        lines.append("· 부서·약물치료·가족력·업무내용은 결과지에 없으므로 직접 확인/입력하세요.")
        self._mark(marks)
        self._set_memo("\n".join(lines))
        self.nb.select(0)

    def _set_memo(self, text):
        self.memo.configure(state="normal")
        self.memo.delete("1.0", "end")
        self.memo.insert("end", text)
        self.memo.configure(state="disabled")

    # ------------------------------------------------------------------ 저장
    def _refresh_depts(self):
        try:
            self.cb_site.configure(values=S.list_sites(self.root_dir))
            self.cb_dept.configure(values=S.list_depts(self.root_dir, self.vars["site"].get()))
        except Exception:
            pass

    def change_root(self):
        p = filedialog.askdirectory(title="자료 폴더 선택", initialdir=str(self.root_dir))
        if p:
            self.root_dir = Path(p)
            S.set_root(p)
            self.root_var.set(p)
            self._refresh_depts()
            self.refresh_list()

    def save_current(self, interactive=True):
        inp = self.collect()
        res = E.calc(inp)
        miss = [n for k, n in (("site", "사업장"), ("dept", "부서"), ("name", "성명"), ("exam_date", "검진일")) if not str(inp[k]).strip()]
        if E.parse_date(inp["exam_date"]) is None and "검진일" not in miss:
            miss.append("검진일 형식(YYYY-MM-DD)")
        if not res.get("ok"):
            miss += res["errors"]
        if miss:
            if interactive:
                messagebox.showwarning("저장 불가", "다음 항목을 입력해 주세요:\n · " + "\n · ".join(miss))
            return None
        inp["exam_date"] = E.parse_date(inp["exam_date"]).isoformat()
        p = S.save_record(self.root_dir, inp)
        self.current_path = p
        self._refresh_depts()
        if interactive:
            self._set_memo(f"저장되었습니다:\n{p}")
        return p

    def new_record(self):
        keep = self.collect()
        base = E.default_input()
        base.update(site=keep["site"], dept=keep["dept"], exam_date=keep["exam_date"])
        self.current_path = None
        self.fill(base)
        self._set_memo("새 입력 (사업장·부서·검진일은 유지됨)")

    # ------------------------------------------------------------------ 목록
    def refresh_list(self):
        sites = S.list_sites(self.root_dir)
        self.cb_list_site.configure(values=sites)
        if self.list_site.get() not in sites:
            self.list_site.set(sites[0] if sites else "")
        for i in self.tree.get_children():
            self.tree.delete(i)
        self._paths.clear()
        site = self.list_site.get() or None
        rows = S.summary_rows(self.root_dir, site) if site else []
        for r in rows:
            iid = self.tree.insert("", "end", values=[r["사업장"], r["부서"], r["이름"], r["검진일"], r["나이"], r["위험군"], r["비고"]])
            self._paths[iid] = r["path"]
        self.count_lbl.configure(text=f"{len(rows)}명 저장됨" + (f"  (사업장: {site})" if site else ""))

    def _sel_path(self):
        s = self.tree.selection()
        return self._paths.get(s[0]) if s else None

    def open_selected(self):
        p = self._sel_path()
        if not p:
            return
        self.current_path = Path(p)
        self.fill(S.load_record(p))
        self._set_memo(f"저장된 자료를 열었습니다: {p}\n수정 후 [저장]하면 덮어씁니다.")
        self.nb.select(0)

    def delete_selected(self):
        p = self._sel_path()
        if p and messagebox.askyesno("삭제", f"이 자료를 삭제할까요?\n{p}"):
            S.delete_record(p)
            self.refresh_list()

    def export_xlsx(self):
        site = self.list_site.get()
        if not site:
            return
        df = S.records_to_dataframe(self.root_dir, site, latest_only=False)
        if df.empty:
            messagebox.showinfo("내보내기", "내보낼 자료가 없습니다.")
            return
        p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"{site}_개인별평가결과.xlsx")
        if p:
            df.drop(columns=["_path"], errors="ignore").to_excel(p, index=False)
            messagebox.showinfo("내보내기", "저장했습니다.\n" + p)

    def make_report(self):
        site = self.list_site.get()
        if not site:
            messagebox.showinfo("보고서", "사업장을 선택하세요.")
            return
        if self.on_make_report:
            self.on_make_report(site, self.root_dir)
