"""
PDF智能处理工具 v3.0（合并版）
- 模式1：分类拆分 - 按关键词规则将PDF页面分类并拆分到不同文件
- 模式2：按页拆分 - 自动识别标题和流水号，按页拆分并智能命名
- 支持扫描PDF的OCR识别
- 智能分类：上下文关联 + 多关键词组合匹配 + 得分制
- 自动提取标题和流水号
"""

import os
import sys
import re
import json
import threading
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from datetime import datetime

# Fix PyInstaller bundled path for rapidocr_onnxruntime
if getattr(sys, 'frozen', False):
    _base_path = sys._MEIPASS
    os.environ.setdefault('RAPIDOCR_MODEL_PATH', os.path.join(_base_path, 'rapidocr_onnxruntime'))

# Third-party imports
try:
    import fitz  # PyMuPDF
except ImportError:
    fitz = None

try:
    from rapidocr_onnxruntime import RapidOCR
except ImportError:
    RapidOCR = None


class PDFUnifiedApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF智能处理工具 v3.0")
        self.root.geometry("1050x850")
        self.root.resizable(True, True)

        # Shared state
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.is_running = False
        self.ocr_engine = None

        # Mode 1: Classification split state
        self.categories = []
        self.default_categories = [
            {
                "name": "付款申请单",
                "keywords": "付款申请单,付款申请",
                "combo_keywords": "付款帐号,请款金额,收款单位,农民工工资,累计付款额,按发票管理分类,请款总,业务摘要"
            },
            {
                "name": "支付申请表",
                "keywords": "支付申请表,支付申请,费用支付申请表",
                "combo_keywords": ""
            },
            {
                "name": "发票",
                "keywords": "发票,增值税,普通发票,专用发票,电子发票",
                "combo_keywords": ""
            },
            {
                "name": "支付情况统计表",
                "keywords": "支付情况统计表,合同款支付情况统计表,支付情况统计",
                "combo_keywords": ""
            },
            {
                "name": "合同文件",
                "keywords": "合同文件,合同协议书,招标文件,合同专用章",
                "combo_keywords": ""
            },
            {
                "name": "检测报告",
                "keywords": "检测报告,检验报告,试验报告,预防性试验报告,检验检测报告,监测技术报告",
                "combo_keywords": ""
            },
            {
                "name": "付款意见书",
                "keywords": "付款意见书,合同付款意见书,审核结果,审核情况",
                "combo_keywords": ""
            },
        ]

        # Mode 2: Page split state
        self.page_names = {}
        self.page_data = {}

        self._build_ui()
        self._load_defaults()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === 1. File Selection (shared) ===
        file_frame = ttk.LabelFrame(main_frame, text="📁 文件选择", padding=8)
        file_frame.pack(fill=tk.X, pady=(0, 8))

        row1 = ttk.Frame(file_frame)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="PDF文件:").pack(side=tk.LEFT)
        self.pdf_entry = ttk.Entry(row1, textvariable=self.pdf_path, width=75)
        self.pdf_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row1, text="浏览...", command=self._browse_pdf).pack(side=tk.LEFT)

        row2 = ttk.Frame(file_frame)
        row2.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(row2, text="输出目录:").pack(side=tk.LEFT)
        self.output_entry = ttk.Entry(row2, textvariable=self.output_dir, width=75)
        self.output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row2, text="浏览...", command=self._browse_output).pack(side=tk.LEFT)

        # === 2. Mode Selection Tabs ===
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # Tab 1: Classification Split
        self.tab1 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab1, text="📑 模式1：分类拆分")
        self._build_classification_tab()

        # Tab 2: Page Split
        self.tab2 = ttk.Frame(self.notebook)
        self.notebook.add(self.tab2, text="📄 模式2：按页拆分")
        self._build_page_split_tab()

        # === 3. Progress & Status (shared) ===
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(0, 8))

        self.progress = ttk.Progressbar(action_frame, mode="determinate", length=500)
        self.progress.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.status_label = ttk.Label(action_frame, text="就绪 - 请选择PDF并选择处理模式")
        self.status_label.pack(side=tk.LEFT, padx=5)

        # === 4. Log (shared) ===
        log_frame = ttk.LabelFrame(main_frame, text="📋 运行日志", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ==================== Tab 1: Classification Split ====================

    def _build_classification_tab(self):
        cat_frame = ttk.LabelFrame(self.tab1, text="🏷️ 分类规则（类别名称 + 主关键字 + 组合关键字，逗号分隔）", padding=8)
        cat_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        table_container = ttk.Frame(cat_frame)
        table_container.pack(fill=tk.BOTH, expand=True)

        self.cat_canvas = tk.Canvas(table_container, height=250)
        scrollbar = ttk.Scrollbar(table_container, orient="vertical", command=self.cat_canvas.yview)
        self.cat_inner_frame = ttk.Frame(self.cat_canvas)

        self.cat_inner_frame.bind(
            "<Configure>",
            lambda e: self.cat_canvas.configure(scrollregion=self.cat_canvas.bbox("all"))
        )

        self.cat_canvas.create_window((0, 0), window=self.cat_inner_frame, anchor="nw")
        self.cat_canvas.configure(yscrollcommand=scrollbar.set)

        self.cat_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        hdr = ttk.Frame(self.cat_inner_frame)
        hdr.pack(fill=tk.X, pady=(0, 3))
        ttk.Label(hdr, text="序号", width=5, anchor="center").pack(side=tk.LEFT, padx=2)
        ttk.Label(hdr, text="类别名称", width=16, anchor="center").pack(side=tk.LEFT, padx=2)
        ttk.Label(hdr, text="主关键字（逗号分隔）", width=35, anchor="center").pack(side=tk.LEFT, padx=2)
        ttk.Label(hdr, text="组合关键字（需2+命中）", width=35, anchor="center").pack(side=tk.LEFT, padx=2)
        ttk.Label(hdr, text="操作", width=8, anchor="center").pack(side=tk.LEFT, padx=2)

        self.cat_rows_frame = ttk.Frame(self.cat_inner_frame)
        self.cat_rows_frame.pack(fill=tk.X)

        btn_row = ttk.Frame(cat_frame)
        btn_row.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(btn_row, text="➕ 添加分类", command=self._add_category_row).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="🔄 重置默认", command=self._load_defaults).pack(side=tk.LEFT, padx=2)

        opt_frame = ttk.LabelFrame(self.tab1, text="⚙️ 选项", padding=8)
        opt_frame.pack(fill=tk.X, pady=(0, 8))

        self.ocr_zoom = tk.IntVar(value=2)
        self.auto_unclass = tk.BooleanVar(value=True)
        self.context_aware = tk.BooleanVar(value=True)

        opt_row = ttk.Frame(opt_frame)
        opt_row.pack(fill=tk.X)
        ttk.Label(opt_row, text="OCR缩放倍数:").pack(side=tk.LEFT)
        ttk.Spinbox(opt_row, from_=1, to=4, textvariable=self.ocr_zoom, width=5).pack(side=tk.LEFT, padx=5)
        ttk.Label(opt_row, text="(越大越清晰，但速度越慢)").pack(side=tk.LEFT)

        opt_row2 = ttk.Frame(opt_frame)
        opt_row2.pack(fill=tk.X, pady=(3, 0))
        ttk.Checkbutton(opt_row2, text="未分类页面单独输出为\"未分类.pdf\"", variable=self.auto_unclass).pack(side=tk.LEFT)
        ttk.Checkbutton(opt_row2, text="上下文关联（连续页面分类优化）", variable=self.context_aware).pack(side=tk.LEFT, padx=(20, 0))

        self.start_classify_btn = ttk.Button(self.tab1, text="🚀 开始分类拆分", command=self._start_classification_split)
        self.start_classify_btn.pack(pady=10)

    # ==================== Tab 2: Page Split ====================

    def _build_page_split_tab(self):
        opt_frame = ttk.LabelFrame(self.tab2, text="⚙️ 识别选项", padding=8)
        opt_frame.pack(fill=tk.X, pady=(0, 8))

        opt_row1 = ttk.Frame(opt_frame)
        opt_row1.pack(fill=tk.X, pady=2)

        self.max_title_len = tk.IntVar(value=30)
        ttk.Label(opt_row1, text="OCR缩放:").pack(side=tk.LEFT)
        ttk.Spinbox(opt_row1, from_=1, to=4, textvariable=self.ocr_zoom, width=5).pack(side=tk.LEFT, padx=(2, 15))

        ttk.Label(opt_row1, text="标题最大字数:").pack(side=tk.LEFT)
        ttk.Spinbox(opt_row1, from_=5, to=100, textvariable=self.max_title_len, width=5).pack(side=tk.LEFT, padx=(2, 15))

        self.prefix_mode = tk.StringVar(value="标题+编号")
        ttk.Label(opt_row1, text="命名方式:").pack(side=tk.LEFT)
        ttk.Combobox(opt_row1, textvariable=self.prefix_mode, width=15,
                     values=["标题+编号", "序号+标题+编号", "仅标题", "仅编号"],
                     state="readonly").pack(side=tk.LEFT, padx=2)

        opt_row2 = ttk.Frame(opt_frame)
        opt_row2.pack(fill=tk.X, pady=2)
        ttk.Label(opt_row2, text="提示: 程序会自动识别每页的表格标题和流水号，您可以预览并编辑后再拆分",
                  foreground="gray").pack(side=tk.LEFT)

        action_frame = ttk.Frame(self.tab2)
        action_frame.pack(fill=tk.X, pady=(0, 8))

        self.recognize_btn = ttk.Button(action_frame, text="🔍 识别标题", command=self._start_recognize)
        self.recognize_btn.pack(side=tk.LEFT, padx=5)

        self.split_btn = ttk.Button(action_frame, text="🚀 拆分PDF", command=self._start_page_split, state=tk.DISABLED)
        self.split_btn.pack(side=tk.LEFT, padx=5)

        preview_frame = ttk.LabelFrame(self.tab2, text="📋 识别结果预览（可双击编辑文件名）", padding=8)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        tree_container = ttk.Frame(preview_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)

        columns = ("page", "title", "serial_number", "filename")
        self.tree = ttk.Treeview(tree_container, columns=columns, show="headings", height=12)

        self.tree.heading("page", text="页码")
        self.tree.heading("title", text="识别标题")
        self.tree.heading("serial_number", text="流水号")
        self.tree.heading("filename", text="文件名")

        self.tree.column("page", width=50, anchor="center", minwidth=40)
        self.tree.column("title", width=220, minwidth=100)
        self.tree.column("serial_number", width=200, minwidth=80)
        self.tree.column("filename", width=350, minwidth=150)

        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._on_tree_double_click)

    # ==================== Shared Methods ====================

    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if path:
            self.pdf_path.set(path)
            if not self.output_dir.get():
                base_dir = os.path.dirname(path)
                self.output_dir.set(os.path.join(base_dir, "split_output"))

    def _browse_output(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    def _log(self, msg):
        def _append():
            timestamp = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{timestamp}] {msg}\n")
            self.log_text.see(tk.END)
        self.root.after(0, _append)

    def _update_progress(self, value, status=""):
        def _update():
            self.progress["value"] = value
            if status:
                self.status_label.config(text=status)
        self.root.after(0, _update)

    # ==================== Classification Tab Methods ====================

    def _load_defaults(self):
        for widget in self.cat_rows_frame.winfo_children():
            widget.destroy()
        self.categories = []

        for cat in self.default_categories:
            self._add_category_row(
                name=cat["name"],
                keywords=cat["keywords"],
                combo_keywords=cat.get("combo_keywords", "")
            )

    def _add_category_row(self, name="", keywords="", combo_keywords=""):
        idx = len(self.categories) + 1
        row_frame = ttk.Frame(self.cat_rows_frame)
        row_frame.pack(fill=tk.X, pady=1)

        idx_label = ttk.Label(row_frame, text=str(idx), width=5, anchor="center")
        idx_label.pack(side=tk.LEFT, padx=2)

        name_var = tk.StringVar(value=name)
        name_entry = ttk.Entry(row_frame, textvariable=name_var, width=16)
        name_entry.pack(side=tk.LEFT, padx=2)

        kw_var = tk.StringVar(value=keywords)
        kw_entry = ttk.Entry(row_frame, textvariable=kw_var, width=35)
        kw_entry.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

        combo_var = tk.StringVar(value=combo_keywords)
        combo_entry = ttk.Entry(row_frame, textvariable=combo_var, width=35)
        combo_entry.pack(side=tk.LEFT, padx=2, fill=tk.X, expand=True)

        def remove_row(rf=row_frame, nv=name_var, kv=kw_var, cv=combo_var):
            rf.destroy()
            self.categories = [c for c in self.categories if c["frame"] != rf]
            self._reindex_rows()

        ttk.Button(row_frame, text="✕", width=3, command=remove_row).pack(side=tk.LEFT, padx=2)

        self.categories.append({
            "frame": row_frame,
            "name_var": name_var,
            "kw_var": kw_var,
            "combo_var": combo_var,
            "idx_label": idx_label,
        })

    def _reindex_rows(self):
        for i, cat in enumerate(self.categories):
            cat["idx_label"].config(text=str(i + 1))

    def _get_categories(self):
        result = []
        for cat in self.categories:
            name = cat["name_var"].get().strip()
            keywords_str = cat["kw_var"].get().strip()
            combo_str = cat["combo_var"].get().strip()
            if name and keywords_str:
                keywords = [kw.strip() for kw in keywords_str.split(",") if kw.strip()]
                combo_keywords = [kw.strip() for kw in combo_str.split(",") if kw.strip()] if combo_str else []
                if keywords:
                    result.append({
                        "name": name,
                        "keywords": keywords,
                        "combo_keywords": combo_keywords,
                    })
        return result

    def _classify_page(self, text, categories):
        if not text:
            return None, 0, []

        scores = {}
        details = {}

        for cat in categories:
            score = 0
            cat_details = []

            for kw in cat["keywords"]:
                if kw in text:
                    base_score = 10
                    length_bonus = max(0, (len(kw) - 2) * 2)
                    total = base_score + length_bonus
                    score += total
                    cat_details.append(f"主关键字[{kw}] +{total}")

            if cat.get("combo_keywords"):
                combo_matches = 0
                matched_combos = []
                for ckw in cat["combo_keywords"]:
                    if ckw in text:
                        combo_matches += 1
                        matched_combos.append(ckw)

                if combo_matches >= 2:
                    combo_score = combo_matches * 3
                    score += combo_score
                    cat_details.append(f"组合关键字{matched_combos} +{combo_score}")
                elif combo_matches > 0:
                    cat_details.append(f"组合关键字部分[{matched_combos}] (不足2项，不计分)")

            scores[cat["name"]] = score
            details[cat["name"]] = cat_details

        best_cat = None
        best_score = 0
        for cat_name, score in scores.items():
            if score > best_score:
                best_score = score
                best_cat = cat_name

        return best_cat, best_score, details.get(best_cat, [])

    def _apply_context(self, page_results, total_pages, categories):
        if not self.context_aware.get():
            return page_results

        pages = sorted(page_results.keys())

        changed = True
        while changed:
            changed = False
            for i in range(len(pages)):
                pn = pages[i]
                current_cat = page_results[pn]

                if i > 0 and i < len(pages) - 1:
                    prev_cat = page_results[pages[i - 1]]
                    next_cat = page_results[pages[i + 1]]
                    if prev_cat == next_cat and prev_cat != current_cat:
                        if prev_cat != "未分类":
                            page_results[pn] = prev_cat
                            self._log(f"  🔄 上下文修正: 第{pn}页 {current_cat} → {prev_cat} (被前后同类别包围)")
                            changed = True
                            break

        all_cat_names = set(page_results.values()) - {"未分类"}
        for candidate in all_cat_names:
            candidate_pages = [pn for pn in pages if page_results[pn] == candidate]
            if len(candidate_pages) < 2:
                continue

            min_page = min(candidate_pages)
            max_page = max(candidate_pages)

            for pn in range(min_page, max_page + 1):
                if page_results[pn] != candidate:
                    old_cat = page_results[pn]
                    page_results[pn] = candidate
                    self._log(f"  🔄 连续性修正: 第{pn}页 {old_cat} → {candidate} (多页文档中间页)")

        for i in range(len(pages)):
            pn = pages[i]
            if page_results[pn] == "未分类":
                prev_cat = None
                next_cat = None
                for j in range(i - 1, -1, -1):
                    if page_results[pages[j]] != "未分类":
                        prev_cat = page_results[pages[j]]
                        break
                for j in range(i + 1, len(pages)):
                    if page_results[pages[j]] != "未分类":
                        next_cat = page_results[pages[j]]
                        break

                if prev_cat and next_cat and prev_cat == next_cat:
                    page_results[pn] = prev_cat
                    self._log(f"  🔄 未分类修正: 第{pn}页 → {prev_cat} (前后均为同类)")
                elif prev_cat:
                    page_results[pn] = prev_cat
                    self._log(f"  🔄 未分类修正: 第{pn}页 → {prev_cat} (跟随前一页)")

        return page_results

    def _start_classification_split(self):
        pdf = self.pdf_path.get().strip()
        if not pdf or not os.path.isfile(pdf):
            messagebox.showerror("错误", "请选择有效的PDF文件")
            return

        out = self.output_dir.get().strip()
        if not out:
            messagebox.showerror("错误", "请指定输出目录")
            return

        cats = self._get_categories()
        if not cats:
            messagebox.showerror("错误", "请至少定义一个分类规则")
            return

        if fitz is None:
            messagebox.showerror("错误", "PyMuPDF未安装，请运行: pip install PyMuPDF")
            return

        if RapidOCR is None:
            messagebox.showerror("错误", "RapidOCR未安装，请运行: pip install rapidocr-onnxruntime")
            return

        self.is_running = True
        self.start_classify_btn.config(state=tk.DISABLED)
        self.log_text.delete("1.0", tk.END)

        thread = threading.Thread(
            target=self._run_classification_split,
            args=(pdf, out, cats),
            daemon=True
        )
        thread.start()

    def _run_classification_split(self, pdf_path, output_dir, categories):
        try:
            self._log(f"正在打开PDF: {pdf_path}")
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            self._log(f"共 {total_pages} 页")

            self._log("正在初始化OCR引擎...")
            ocr = RapidOCR()
            self._log("OCR引擎就绪")

            page_texts = {}
            page_results = {}
            page_details = {}

            for page_num in range(total_pages):
                self._log(f"正在处理第 {page_num + 1}/{total_pages} 页...")
                self._update_progress((page_num / total_pages) * 70, f"OCR识别: {page_num+1}/{total_pages}")

                page = doc[page_num]
                zoom = self.ocr_zoom.get()
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)

                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                    tmp_path = tmp.name
                pix.save(tmp_path)

                try:
                    ocr_result, _ = ocr(tmp_path)
                    if ocr_result:
                        text = " ".join([item[1] for item in ocr_result])
                    else:
                        text = ""
                except Exception as e:
                    self._log(f"  ⚠️ OCR异常: {e}")
                    text = ""
                finally:
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass

                page_texts[page_num + 1] = text
                category, score, detail = self._classify_page(text, categories)

                if category is None or score == 0:
                    category = "未分类"

                page_results[page_num + 1] = category
                page_details[page_num + 1] = detail

                preview = text[:60].replace("\n", " ") if text else "(空)"
                detail_str = ", ".join(detail) if detail else "无匹配"
                self._log(f"  → {category} (得分:{score}) | {detail_str} | {preview}...")

            self._update_progress(75, "正在优化分类结果...")
            self._log("\n正在应用上下文关联优化...")
            page_results = self._apply_context(page_results, total_pages, categories)

            self._update_progress(85, "正在拆分PDF...")
            self._log("\n正在拆分PDF文件...")

            category_pages = {}
            for page_num, category in page_results.items():
                if category not in category_pages:
                    category_pages[category] = []
                category_pages[category].append(page_num)

            os.makedirs(output_dir, exist_ok=True)

            created_files = []
            for category, pages in sorted(category_pages.items()):
                if category == "未分类" and not self.auto_unclass.get() and len(pages) == 0:
                    continue

                idx = ""
                for cat in categories:
                    if cat["name"] == category:
                        idx = f"{categories.index(cat) + 1}_"
                        break

                filename = f"{idx}{category}.pdf"
                output_path = os.path.join(output_dir, filename)

                new_doc = fitz.open()
                for page_num in sorted(pages):
                    new_doc.insert_pdf(doc, from_page=page_num - 1, to_page=page_num - 1)
                new_doc.save(output_path)
                new_doc.close()

                created_files.append((filename, len(pages), pages))
                self._log(f"  ✅ {filename} ({len(pages)}页: 第{pages}页)")

            doc.close()

            result_path = os.path.join(output_dir, "classification_results.txt")
            with open(result_path, "w", encoding="utf-8") as f:
                f.write("PDF智能拆分结果\n")
                f.write(f"源文件: {pdf_path}\n")
                f.write(f"拆分时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"总页数: {total_pages}\n")
                f.write(f"分类方法: 得分制 + 上下文关联\n")
                f.write("=" * 60 + "\n\n")
                for filename, page_count, pages in created_files:
                    f.write(f"{filename}: {page_count}页 (原始第{pages}页)\n")
                f.write(f"\n详细页码分类：\n")
                f.write("-" * 60 + "\n")
                for page_num in range(1, total_pages + 1):
                    detail_str = ", ".join(page_details.get(page_num, []))
                    f.write(f"第{page_num:2d}页: {page_results[page_num]:10s} | {detail_str}\n")

            self._update_progress(100, "完成!")
            self._log(f"\n🎉 拆分完成！共生成 {len(created_files)} 个文件")
            self._log(f"输出目录: {output_dir}")

            self.root.after(0, lambda: messagebox.showinfo(
                "拆分完成",
                f"成功将 {total_pages} 页PDF拆分为 {len(created_files)} 个文件！\n\n输出目录: {output_dir}"
            ))

        except Exception as e:
            self._log(f"❌ 错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"拆分过程中出错:\n{e}"))

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.start_classify_btn.config(state=tk.NORMAL))

    # ==================== Page Split Tab Methods ====================

    @staticmethod
    def _extract_title(ocr_result, max_len=30):
        if not ocr_result:
            return ""

        items = []
        for item in ocr_result:
            text = item[1].strip()
            bbox = item[0]
            center_y = (bbox[0][1] + bbox[2][1]) / 2
            height = abs(bbox[2][1] - bbox[0][1])
            center_x = (bbox[0][0] + bbox[2][0]) / 2
            width = abs(bbox[2][0] - bbox[0][0])
            items.append({
                "text": text,
                "y": center_y,
                "x": center_x,
                "height": height,
                "width": width,
            })

        if not items:
            return ""

        title_suffixes = [
            "表", "单", "书", "报告", "意见书", "申请", "统计表",
            "检测表", "记录", "清单", "汇总", "概算", "预算",
            "审签表", "请款表", "付款单", "支付表",
        ]

        items_by_y = sorted(items, key=lambda i: i["y"])

        if items_by_y:
            min_y = items_by_y[0]["y"]
            max_y = max(i["y"] for i in items_by_y)
            page_height = max_y - min_y if max_y > min_y else 1
            top_threshold = min_y + page_height * 0.2
        else:
            top_threshold = 0

        top_items = [i for i in items_by_y if i["y"] <= top_threshold]

        if items:
            avg_height = sum(i["height"] for i in items) / len(items)
        else:
            avg_height = 0

        candidates = []
        for item in top_items:
            text = item["text"]
            if len(text) < 2:
                continue
            if len(text) > max_len * 2:
                continue
            if re.match(r'^[\d\s\-、]+$', text):
                continue
            is_title_like = any(text.endswith(s) or s in text for s in title_suffixes)
            font_bonus = 1.5 if item["height"] > avg_height else 1.0

            candidates.append({
                "text": text,
                "score": (font_bonus * (1 if is_title_like else 0.5)) + (0.3 / max(len(text), 1)),
                "y": item["y"],
                "is_title_like": is_title_like,
            })

        if not candidates:
            for item in top_items:
                if len(item["text"]) >= 2 and len(item["text"]) <= max_len * 2:
                    return item["text"][:max_len]
            return ""

        candidates.sort(key=lambda c: (-c["score"], c["y"]))

        best = candidates[0]["text"]
        best = re.sub(r'^[\d\s\-、.]+', '', best).strip()
        return best[:max_len] if best else ""

    @staticmethod
    def _extract_serial_number(ocr_result):
        if not ocr_result:
            return ""

        all_text = " ".join([item[1] for item in ocr_result])

        patterns = [
            r'[A-Za-z]{2,}[\d年]+第?[\d]+号',
            r'(?:编号|编号：|No\.|编号:)\s*([A-Za-z0-9\-年月第号]+)',
            r'(?:合同编号|合同号)[：:]?\s*([A-Za-z0-9\-]+)',
            r'(?:文号|发文字号)[：:]?\s*([A-Za-z0-9\-〔年]+[\d]+号?)',
            r'[A-Z]{2,}[\d]{4}年第?[\d]+号',
            r'[A-Z]{2,}-[A-Z]{2,}-[\d]{4}-[\d]+',
        ]

        for pattern in patterns:
            match = re.search(pattern, all_text)
            if match:
                result = match.group(1) if match.lastindex else match.group(0)
                return result.strip()

        return ""

    @staticmethod
    def _build_filename(title, serial, mode, page_num, max_len=30):
        clean_title = re.sub(r'[\\/:*?"<>|]', '', title)[:max_len] if title else f"第{page_num}页"
        clean_serial = re.sub(r'[\\/:*?"<>|]', '', serial) if serial else ""

        if mode == "标题+编号":
            if clean_title and clean_serial:
                return f"{clean_title}+{clean_serial}"
            elif clean_title:
                return clean_title
            elif clean_serial:
                return clean_serial
            else:
                return f"第{page_num}页"
        elif mode == "序号+标题+编号":
            parts = [f"{page_num:02d}"]
            if clean_title:
                parts.append(clean_title)
            if clean_serial:
                parts.append(clean_serial)
            return "_".join(parts) if len(parts) > 1 else f"第{page_num}页"
        elif mode == "仅标题":
            return clean_title if clean_title else f"第{page_num}页"
        elif mode == "仅编号":
            return clean_serial if clean_serial else f"第{page_num}页"
        else:
            return f"第{page_num}页"

    def _start_recognize(self):
        pdf = self.pdf_path.get().strip()
        if not pdf or not os.path.isfile(pdf):
            messagebox.showerror("错误", "请选择有效的PDF文件")
            return

        if fitz is None:
            messagebox.showerror("错误", "PyMuPDF未安装")
            return

        if RapidOCR is None:
            messagebox.showerror("错误", "RapidOCR未安装")
            return

        self.is_running = True
        self.recognize_btn.config(state=tk.DISABLED)
        self.split_btn.config(state=tk.DISABLED)
        self.log_text.delete("1.0", tk.END)

        for item in self.tree.get_children():
            self.tree.delete(item)
        self.page_names = {}
        self.page_data = {}

        thread = threading.Thread(target=self._run_recognize, args=(pdf,), daemon=True)
        thread.start()

    def _run_recognize(self, pdf_path):
        try:
            self._log(f"正在打开PDF: {pdf_path}")
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            self._log(f"共 {total_pages} 页")

            self._log("正在初始化OCR引擎...")
            ocr = RapidOCR()
            self._log("OCR引擎就绪，开始识别...")

            max_title = self.max_title_len.get()
            mode = self.prefix_mode.get()

            for page_num in range(total_pages):
                self._update_progress(
                    (page_num / total_pages) * 100,
                    f"识别中: {page_num+1}/{total_pages}"
                )
                self._log(f"正在识别第 {page_num + 1} 页...")

                page = doc[page_num]
                zoom = self.ocr_zoom.get()
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)

                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                    tmp_path = tmp.name
                pix.save(tmp_path)

                ocr_result = None
                try:
                    ocr_result, _ = ocr(tmp_path)
                except Exception as e:
                    self._log(f"  ⚠️ OCR异常: {e}")
                finally:
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass

                title = self._extract_title(ocr_result, max_len=max_title)
                serial = self._extract_serial_number(ocr_result)
                filename = self._build_filename(title, serial, mode, page_num + 1, max_title)

                self.page_names[page_num + 1] = filename
                self.page_data[page_num + 1] = {"title": title, "serial": serial, "filename": filename}

                pn = page_num + 1
                self.root.after(0, lambda p=pn, t=title, s=serial, f=filename:
                    self.tree.insert("", "end", iid=str(p), values=(p, t, s, f)))

                self._log(f"  标题: {title} | 编号: {serial}")
                self._log(f"  → 文件名: {filename}.pdf")

            doc.close()
            self._update_progress(100, f"识别完成 - 共{total_pages}页")
            self._log(f"\n✅ 识别完成！请检查并编辑文件名，然后点击拆分")

            self.root.after(0, lambda: self.split_btn.config(state=tk.NORMAL))

        except Exception as e:
            self._log(f"❌ 错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.recognize_btn.config(state=tk.NORMAL))

    def _on_tree_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)

        if not item_id:
            return

        col_idx = int(column.replace("#", "")) - 1

        if col_idx != 3:
            return

        current_values = self.tree.item(item_id, "values")
        current_text = current_values[col_idx]

        bbox = self.tree.bbox(item_id, column)
        if not bbox:
            return

        entry = tk.Entry(self.tree, font=("Microsoft YaHei", 9))
        entry.place(x=bbox[0], y=bbox[1], width=bbox[2], height=bbox[3])
        entry.insert(0, current_text)
        entry.select_range(0, tk.END)
        entry.focus_set()

        def save_edit(event=None):
            new_text = entry.get().strip()
            if new_text:
                values = list(self.tree.item(item_id, "values"))
                values[col_idx] = new_text
                self.tree.item(item_id, values=values)
                page_num = int(item_id)
                self.page_names[page_num] = new_text
            entry.destroy()

        def cancel_edit(event=None):
            entry.destroy()

        entry.bind("<Return>", save_edit)
        entry.bind("<Escape>", cancel_edit)
        entry.bind("<FocusOut>", save_edit)

    def _start_page_split(self):
        if not self.page_names:
            messagebox.showerror("错误", "请先识别PDF")
            return

        out = self.output_dir.get().strip()
        if not out:
            messagebox.showerror("错误", "请指定输出目录")
            return

        if fitz is None:
            messagebox.showerror("错误", "PyMuPDF未安装")
            return

        self.is_running = True
        self.recognize_btn.config(state=tk.DISABLED)
        self.split_btn.config(state=tk.DISABLED)

        thread = threading.Thread(
            target=self._run_page_split,
            args=(self.pdf_path.get(), out),
            daemon=True
        )
        thread.start()

    def _run_page_split(self, pdf_path, output_dir):
        try:
            self._log(f"正在打开PDF: {pdf_path}")
            doc = fitz.open(pdf_path)
            total_pages = len(doc)

            os.makedirs(output_dir, exist_ok=True)

            self._log(f"\n正在拆分 {total_pages} 页...")
            created_count = 0

            for page_num in range(total_pages):
                self._update_progress(
                    ((page_num + 1) / total_pages) * 100,
                    f"拆分中: {page_num+1}/{total_pages}"
                )

                pn = page_num + 1
                filename = self.page_names.get(pn, f"第{pn}页")
                output_path = os.path.join(output_dir, f"{filename}.pdf")

                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=page_num, to_page=page_num)
                new_doc.save(output_path)
                new_doc.close()

                created_count += 1
                self._log(f"  ✅ 第{pn}页 → {filename}.pdf")

            doc.close()

            self._update_progress(100, "完成!")
            self._log(f"\n🎉 拆分完成！共生成 {created_count} 个文件")
            self._log(f"输出目录: {output_dir}")

            self.root.after(0, lambda: messagebox.showinfo(
                "拆分完成",
                f"成功将 {total_pages} 页PDF拆分为 {created_count} 个文件！\n\n输出目录: {output_dir}"
            ))

        except Exception as e:
            self._log(f"❌ 错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"拆分过程中出错:\n{e}"))

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.split_btn.config(state=tk.NORMAL))


def main():
    missing = []
    if fitz is None:
        missing.append("PyMuPDF (pip install PyMuPDF)")
    if RapidOCR is None:
        missing.append("rapidocr-onnxruntime (pip install rapidocr-onnxruntime)")

    root = tk.Tk()

    if missing:
        messagebox.showwarning(
            "依赖缺失",
            "以下依赖未安装，程序可能无法正常工作:\n\n" +
            "\n".join(missing) +
            "\n\n请在命令行中运行上述pip install命令安装。"
        )

    app = PDFUnifiedApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
