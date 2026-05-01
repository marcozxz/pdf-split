"""
PDF智能拆分工具 v2.0
- 支持扫描PDF的OCR识别
- 自定义分类关键字
- 智能分类：上下文关联 + 多关键词组合匹配 + 得分制
- 自动拆分并导出
"""

import os
import sys
import json
import threading
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


class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF智能拆分工具 v2.0")
        self.root.geometry("850x750")
        self.root.resizable(True, True)

        # State
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.categories = []  # List of dicts: {name, keywords}
        self.is_running = False
        self.ocr_engine = None

        # Default categories with enriched keywords
        # keywords: primary keywords for direct matching
        # combo_keywords: combination keywords - need 2+ matches to trigger category
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

        self._build_ui()
        self._load_defaults()

    def _build_ui(self):
        # === Main container with padding ===
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === 1. File Selection ===
        file_frame = ttk.LabelFrame(main_frame, text="📁 文件选择", padding=8)
        file_frame.pack(fill=tk.X, pady=(0, 8))

        row1 = ttk.Frame(file_frame)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="PDF文件:").pack(side=tk.LEFT)
        self.pdf_entry = ttk.Entry(row1, textvariable=self.pdf_path, width=60)
        self.pdf_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row1, text="浏览...", command=self._browse_pdf).pack(side=tk.LEFT)

        row2 = ttk.Frame(file_frame)
        row2.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(row2, text="输出目录:").pack(side=tk.LEFT)
        self.output_entry = ttk.Entry(row2, textvariable=self.output_dir, width=60)
        self.output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row2, text="浏览...", command=self._browse_output).pack(side=tk.LEFT)

        # === 2. Category Definition ===
        cat_frame = ttk.LabelFrame(main_frame, text="🏷️ 分类规则（类别名称 + 主关键字 + 组合关键字，逗号分隔）", padding=8)
        cat_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # Category table
        table_container = ttk.Frame(cat_frame)
        table_container.pack(fill=tk.BOTH, expand=True)

        # Scrollable frame for categories
        self.cat_canvas = tk.Canvas(table_container, height=220)
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

        # Headers
        hdr = ttk.Frame(self.cat_inner_frame)
        hdr.pack(fill=tk.X, pady=(0, 3))
        ttk.Label(hdr, text="序号", width=5, anchor="center").pack(side=tk.LEFT, padx=2)
        ttk.Label(hdr, text="类别名称", width=16, anchor="center").pack(side=tk.LEFT, padx=2)
        ttk.Label(hdr, text="主关键字（逗号分隔）", width=35, anchor="center").pack(side=tk.LEFT, padx=2)
        ttk.Label(hdr, text="组合关键字（需2+命中）", width=35, anchor="center").pack(side=tk.LEFT, padx=2)
        ttk.Label(hdr, text="操作", width=8, anchor="center").pack(side=tk.LEFT, padx=2)

        # Category rows container
        self.cat_rows_frame = ttk.Frame(self.cat_inner_frame)
        self.cat_rows_frame.pack(fill=tk.X)

        # Add/Reset buttons
        btn_row = ttk.Frame(cat_frame)
        btn_row.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(btn_row, text="➕ 添加分类", command=self._add_category_row).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_row, text="🔄 重置默认", command=self._load_defaults).pack(side=tk.LEFT, padx=2)

        # === 3. Options ===
        opt_frame = ttk.LabelFrame(main_frame, text="⚙️ 选项", padding=8)
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

        # === 4. Action Buttons ===
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(0, 8))

        self.start_btn = ttk.Button(action_frame, text="🚀 开始拆分", command=self._start_split)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(action_frame, mode="determinate", length=400)
        self.progress.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.status_label = ttk.Label(action_frame, text="就绪")
        self.status_label.pack(side=tk.LEFT, padx=5)

        # === 5. Log ===
        log_frame = ttk.LabelFrame(main_frame, text="📋 运行日志", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def _load_defaults(self):
        """Load default category rows."""
        # Clear existing rows
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
        """Add a category input row."""
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
        """Re-index category rows after deletion."""
        for i, cat in enumerate(self.categories):
            cat["idx_label"].config(text=str(i + 1))

    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if path:
            self.pdf_path.set(path)
            # Auto-set output dir
            if not self.output_dir.get():
                base_dir = os.path.dirname(path)
                self.output_dir.set(os.path.join(base_dir, "split_output"))

    def _browse_output(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    def _log(self, msg):
        """Thread-safe logging."""
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

    def _get_categories(self):
        """Extract category definitions from UI."""
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

    def _start_split(self):
        """Start the splitting process in a background thread."""
        # Validate inputs
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

        # Disable UI
        self.is_running = True
        self.start_btn.config(state=tk.DISABLED)
        self.log_text.delete("1.0", tk.END)

        # Start background thread
        thread = threading.Thread(
            target=self._run_split,
            args=(pdf, out, cats),
            daemon=True
        )
        thread.start()

    def _classify_page(self, text, categories):
        """
        Classify a page based on OCR text using scoring system.

        Scoring rules:
        - Primary keyword match: +10 points per match
        - Longer primary keyword match gets bonus: +2 per extra char beyond 2
        - Combo keyword match: +3 points per match (need 2+ combo matches to contribute)
        - Category with highest total score wins

        Returns: (category_name, score, match_details)
        """
        if not text:
            return None, 0, []

        scores = {}
        details = {}

        for cat in categories:
            score = 0
            cat_details = []

            # Primary keyword matching
            for kw in cat["keywords"]:
                if kw in text:
                    base_score = 10
                    # Bonus for longer keywords (more specific)
                    length_bonus = max(0, (len(kw) - 2) * 2)
                    total = base_score + length_bonus
                    score += total
                    cat_details.append(f"主关键字[{kw}] +{total}")

            # Combo keyword matching (only counts if 2+ matches)
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

        # Find best category
        best_cat = None
        best_score = 0
        for cat_name, score in scores.items():
            if score > best_score:
                best_score = score
                best_cat = cat_name

        return best_cat, best_score, details.get(best_cat, [])

    def _apply_context(self, page_results, total_pages, categories):
        """
        Apply context-aware optimization to classification results.
        Uses page continuity to fix misclassifications.

        Rules:
        1. Single-page island: If page N is cat X, and pages N-1 and N+1
           are both cat Y (Y != X, Y != 未分类), reclassify page N to cat Y.
        2. Multi-page gap: If a category spans pages A..B, and some pages
           in between are classified differently or as 未分类, fill the gaps.
        3. Trailing 未分类: If a 未分类 page is between two classified pages,
           assign it to the preceding category.
        """
        if not self.context_aware.get():
            return page_results

        pages = sorted(page_results.keys())

        # Rule 1: Single-page island surrounded by same category
        # If page N is cat X, and pages N-1 and N+1 are cat Y (Y != X),
        # and cat Y has a strong match, reclassify page N to cat Y.
        changed = True
        while changed:
            changed = False
            for i in range(len(pages)):
                pn = pages[i]
                current_cat = page_results[pn]

                # Check if this page is an "island" (surrounded by same different category)
                if i > 0 and i < len(pages) - 1:
                    prev_cat = page_results[pages[i - 1]]
                    next_cat = page_results[pages[i + 1]]
                    if prev_cat == next_cat and prev_cat != current_cat:
                        # Island detected - but only fix if the surrounding category
                        # makes sense (not "未分类")
                        if prev_cat != "未分类":
                            page_results[pn] = prev_cat
                            self._log(f"  🔄 上下文修正: 第{pn}页 {current_cat} → {prev_cat} (被前后同类别包围)")
                            changed = True
                            break

        # Rule 2: Multi-page document gap fill
        # For each category, find its page range and fill gaps with 未分类 or
        # misclassified pages (e.g., 付款意见书 pages 4-6, where page 5 got
        # misclassified because of incidental keyword matches like "支付情况统计表"
        # being referenced as an attachment name)
        all_cat_names = set(page_results.values()) - {"未分类"}
        for candidate in all_cat_names:
            candidate_pages = [pn for pn in pages if page_results[pn] == candidate]
            if len(candidate_pages) < 2:
                continue  # Only apply to multi-page categories

            min_page = min(candidate_pages)
            max_page = max(candidate_pages)

            for pn in range(min_page, max_page + 1):
                if page_results[pn] != candidate:
                    # This page is inside the range of a multi-page document
                    # but classified differently — likely a misclassification
                    # due to incidental keyword matches (e.g., attachment references)
                    old_cat = page_results[pn]
                    page_results[pn] = candidate
                    self._log(f"  🔄 连续性修正: 第{pn}页 {old_cat} → {candidate} (多页文档中间页)")

        # Rule 3: Assign trailing 未分类 pages to adjacent classified category
        # If a 未分类 page is between two classified pages, prefer the preceding category
        for i in range(len(pages)):
            pn = pages[i]
            if page_results[pn] == "未分类":
                # Look for nearest non-未分类 neighbor
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

                # Prefer same category as neighbor if both sides agree
                if prev_cat and next_cat and prev_cat == next_cat:
                    page_results[pn] = prev_cat
                    self._log(f"  🔄 未分类修正: 第{pn}页 → {prev_cat} (前后均为同类)")
                elif prev_cat:
                    page_results[pn] = prev_cat
                    self._log(f"  🔄 未分类修正: 第{pn}页 → {prev_cat} (跟随前一页)")

        return page_results

    def _run_split(self, pdf_path, output_dir, categories):
        """Core splitting logic running in background thread."""
        try:
            self._log(f"正在打开PDF: {pdf_path}")
            doc = fitz.open(pdf_path)
            total_pages = len(doc)
            self._log(f"共 {total_pages} 页")

            # Initialize OCR engine
            self._log("正在初始化OCR引擎...")
            ocr = RapidOCR()
            self._log("OCR引擎就绪")

            # Store OCR text and scores for each page
            page_texts = {}
            page_results = {}  # page_num -> category_name
            page_details = {}  # page_num -> match details

            # Process each page
            for page_num in range(total_pages):
                self._log(f"正在处理第 {page_num + 1}/{total_pages} 页...")
                self._update_progress((page_num / total_pages) * 70, f"OCR识别: {page_num+1}/{total_pages}")

                page = doc[page_num]

                # Render to image
                zoom = self.ocr_zoom.get()
                mat = fitz.Matrix(zoom, zoom)
                pix = page.get_pixmap(matrix=mat)

                # Save temp image
                import tempfile
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                    tmp_path = tmp.name
                pix.save(tmp_path)

                # OCR
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
                    # Clean up temp file
                    try:
                        os.unlink(tmp_path)
                    except:
                        pass

                page_texts[page_num + 1] = text

                # Classify using scoring system
                category, score, detail = self._classify_page(text, categories)

                if category is None or score == 0:
                    category = "未分类"

                page_results[page_num + 1] = category
                page_details[page_num + 1] = detail

                preview = text[:60].replace("\n", " ") if text else "(空)"
                detail_str = ", ".join(detail) if detail else "无匹配"
                self._log(f"  → {category} (得分:{score}) | {detail_str} | {preview}...")

            # Apply context-aware optimization
            self._update_progress(75, "正在优化分类结果...")
            self._log("\n正在应用上下文关联优化...")
            page_results = self._apply_context(page_results, total_pages, categories)

            # Group pages by category
            self._update_progress(85, "正在拆分PDF...")
            self._log("\n正在拆分PDF文件...")

            category_pages = {}
            for page_num, category in page_results.items():
                if category not in category_pages:
                    category_pages[category] = []
                category_pages[category].append(page_num)

            # Create output directory
            os.makedirs(output_dir, exist_ok=True)

            # Split PDF
            created_files = []
            for category, pages in sorted(category_pages.items()):
                # Skip "未分类" if option disabled and no pages
                if category == "未分类" and not self.auto_unclass.get() and len(pages) == 0:
                    continue

                # Build filename with index prefix for ordering
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

            # Save classification results
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

            # Show summary
            self.root.after(0, lambda: messagebox.showinfo(
                "拆分完成",
                f"成功将 {total_pages} 页PDF拆分为 {len(created_files)} 个文件！\n\n输出目录: {output_dir}"
            ))

        except Exception as e:
            self._log(f"❌ 错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", f"拆分过程中出错:\n{e}"))

        finally:
            self.is_running = False
            self.root.after(0, lambda: self.start_btn.config(state=tk.NORMAL))


def main():
    # Check dependencies
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

    app = PDFSplitterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
