"""
PDF按页拆分命名工具
- 自动OCR识别每页的表头标题和流水号
- 以"标题+流水号"命名拆分后的PDF文件
- 支持预览和编辑识别结果后再拆分
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
# When running as a PyInstaller exe, the package data files are in sys._MEIPASS
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


class PDFPageSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF按页拆分命名工具")
        self.root.geometry("980x780")
        self.root.resizable(True, True)

        # State
        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.is_running = False
        self.ocr_engine = None
        self.page_names = {}  # page_num -> filename (editable)

        self._build_ui()

    def _build_ui(self):
        main_frame = ttk.Frame(self.root, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # === 1. File Selection ===
        file_frame = ttk.LabelFrame(main_frame, text="📁 文件选择", padding=8)
        file_frame.pack(fill=tk.X, pady=(0, 8))

        row1 = ttk.Frame(file_frame)
        row1.pack(fill=tk.X)
        ttk.Label(row1, text="PDF文件:").pack(side=tk.LEFT)
        self.pdf_entry = ttk.Entry(row1, textvariable=self.pdf_path, width=70)
        self.pdf_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row1, text="浏览...", command=self._browse_pdf).pack(side=tk.LEFT)

        row2 = ttk.Frame(file_frame)
        row2.pack(fill=tk.X, pady=(5, 0))
        ttk.Label(row2, text="输出目录:").pack(side=tk.LEFT)
        self.output_entry = ttk.Entry(row2, textvariable=self.output_dir, width=70)
        self.output_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(row2, text="浏览...", command=self._browse_output).pack(side=tk.LEFT)

        # === 2. Options ===
        opt_frame = ttk.LabelFrame(main_frame, text="⚙️ 识别选项", padding=8)
        opt_frame.pack(fill=tk.X, pady=(0, 8))

        opt_row1 = ttk.Frame(opt_frame)
        opt_row1.pack(fill=tk.X, pady=2)

        self.ocr_zoom = tk.IntVar(value=2)
        ttk.Label(opt_row1, text="OCR缩放:").pack(side=tk.LEFT)
        ttk.Spinbox(opt_row1, from_=1, to=4, textvariable=self.ocr_zoom, width=5).pack(side=tk.LEFT, padx=(2, 15))

        self.max_title_len = tk.IntVar(value=30)
        ttk.Label(opt_row1, text="标题最大字数:").pack(side=tk.LEFT)
        ttk.Spinbox(opt_row1, from_=5, to=100, textvariable=self.max_title_len, width=5).pack(side=tk.LEFT, padx=(2, 15))

        self.prefix_mode = tk.StringVar(value="标题+编号")
        ttk.Label(opt_row1, text="命名方式:").pack(side=tk.LEFT)
        ttk.Combobox(opt_row1, textvariable=self.prefix_mode, width=15,
                     values=["标题+编号", "序号+标题+编号", "仅标题", "仅编号"],
                     state="readonly").pack(side=tk.LEFT, padx=2)

        # Explanation row
        opt_row2 = ttk.Frame(opt_frame)
        opt_row2.pack(fill=tk.X, pady=2)
        ttk.Label(opt_row2, text="提示: 程序会自动识别每页的表格标题和流水号(如FPOK2025年1893号)，您可以预览并编辑后再拆分",
                  foreground="gray").pack(side=tk.LEFT)

        # === 3. Action: Recognize ===
        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(0, 8))

        self.recognize_btn = ttk.Button(action_frame, text="🔍 识别标题", command=self._start_recognize)
        self.recognize_btn.pack(side=tk.LEFT, padx=5)

        self.split_btn = ttk.Button(action_frame, text="🚀 拆分PDF", command=self._start_split, state=tk.DISABLED)
        self.split_btn.pack(side=tk.LEFT, padx=5)

        self.progress = ttk.Progressbar(action_frame, mode="determinate", length=350)
        self.progress.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        self.status_label = ttk.Label(action_frame, text="就绪 - 请选择PDF并点击识别")
        self.status_label.pack(side=tk.LEFT, padx=5)

        # === 4. Preview Table ===
        preview_frame = ttk.LabelFrame(main_frame, text="📋 识别结果预览（可双击编辑文件名）", padding=8)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # Treeview
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

        # Scrollbar
        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # Double-click to edit
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        # === 5. Log ===
        log_frame = ttk.LabelFrame(main_frame, text="📋 运行日志", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=False)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, wrap=tk.WORD, font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # ===================== File Selection =====================

    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="选择PDF文件",
            filetypes=[("PDF文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if path:
            self.pdf_path.set(path)
            if not self.output_dir.get():
                self.output_dir.set(os.path.join(os.path.dirname(path), "split_pages"))

    def _browse_output(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    # ===================== Logging =====================

    def _log(self, msg):
        def _append():
            ts = datetime.now().strftime("%H:%M:%S")
            self.log_text.insert(tk.END, f"[{ts}] {msg}\n")
            self.log_text.see(tk.END)
        self.root.after(0, _append)

    def _update_progress(self, value, status=""):
        def _update():
            self.progress["value"] = value
            if status:
                self.status_label.config(text=status)
        self.root.after(0, _update)

    # ===================== Title & Serial Number Extraction =====================

    @staticmethod
    def _extract_title(ocr_result, max_len=30):
        """Extract the table/document title from OCR results.

        Strategy:
        1. Find the topmost (smallest y) text that looks like a title
        2. Title candidates are usually short, centered, with larger font
        3. Common patterns: XXX表, XXX单, XXX书, XXX报告, etc.
        """
        if not ocr_result:
            return ""

        # Collect all text items with position info
        items = []
        for item in ocr_result:
            text = item[1].strip()
            confidence = item[2]
            bbox = item[0]  # [[x1,y1],[x2,y2],[x3,y3],[x4,y4]]
            # Center Y of the text
            center_y = (bbox[0][1] + bbox[2][1]) / 2
            # Height of the text (approximate font size)
            height = abs(bbox[2][1] - bbox[0][1])
            # Center X
            center_x = (bbox[0][0] + bbox[2][0]) / 2
            # Width
            width = abs(bbox[2][0] - bbox[0][0])
            items.append({
                "text": text,
                "y": center_y,
                "x": center_x,
                "height": height,
                "width": width,
                "confidence": confidence,
            })

        if not items:
            return ""

        # Title suffixes commonly found in Chinese documents
        title_suffixes = [
            "表", "单", "书", "报告", "意见书", "申请", "统计表",
            "检测表", "记录", "清单", "汇总", "概算", "预算",
            "审签表", "请款表", "付款单", "支付表",
        ]

        # Sort by Y position (top first)
        items_by_y = sorted(items, key=lambda i: i["y"])

        # Strategy: Look for title-like text in the top portion
        # Top 20% of the page
        if items_by_y:
            min_y = items_by_y[0]["y"]
            max_y = max(i["y"] for i in items_by_y)
            page_height = max_y - min_y if max_y > min_y else 1
            top_threshold = min_y + page_height * 0.2
        else:
            top_threshold = 0

        # Filter top items
        top_items = [i for i in items_by_y if i["y"] <= top_threshold]

        # Calculate average font height
        if items:
            avg_height = sum(i["height"] for i in items) / len(items)
        else:
            avg_height = 0

        # Find title candidates: top area, preferably larger font, short text
        candidates = []
        for item in top_items:
            text = item["text"]
            # Skip very short text (likely noise)
            if len(text) < 2:
                continue
            # Skip very long text (likely body content)
            if len(text) > max_len * 2:
                continue
            # Skip text that's clearly body (starts with numbers, common prefixes)
            if re.match(r'^[\d\s\-、]+$', text):
                continue
            # Prefer text with title suffixes or larger font
            is_title_like = any(text.endswith(s) or s in text for s in title_suffixes)
            font_bonus = 1.5 if item["height"] > avg_height else 1.0

            candidates.append({
                "text": text,
                "score": (font_bonus * (1 if is_title_like else 0.5)) + (0.3 / max(len(text), 1)),
                "y": item["y"],
                "is_title_like": is_title_like,
            })

        if not candidates:
            # Fallback: just use the first top text
            for item in top_items:
                if len(item["text"]) >= 2 and len(item["text"]) <= max_len * 2:
                    return item["text"][:max_len]
            return ""

        # Sort by score (best first)
        candidates.sort(key=lambda c: (-c["score"], c["y"]))

        # Return best candidate, truncated
        best = candidates[0]["text"]
        # Clean up common prefixes
        best = re.sub(r'^[\d\s\-、.]+', '', best).strip()
        return best[:max_len] if best else ""

    @staticmethod
    def _extract_serial_number(ocr_result):
        """Extract document serial number / reference number from OCR results.

        Common patterns:
        - FPOK2025年1893号
        - GCOK2025年第300号
        - 编号：XXXX
        - No.XXXX
        - XXX-XXXX
        """
        if not ocr_result:
            return ""

        all_text = " ".join([item[1] for item in ocr_result])

        # Pattern list (ordered by priority)
        patterns = [
            # Chinese document numbers: 字母+年+号
            r'[A-Za-z]{2,}[\d年]+第?[\d]+号',
            # 编号：XXX
            r'(?:编号|编号：|No\.|编号:)\s*([A-Za-z0-9\-年月第号]+)',
            # Contract numbers
            r'(?:合同编号|合同号)[：:]?\s*([A-Za-z0-9\-]+)',
            # Filing numbers
            r'(?:文号|发文字号)[：:]?\s*([A-Za-z0-9\-〔年]+[\d]+号?)',
            # General pattern: letters + year + number + 号
            r'[A-Z]{2,}[\d]{4}年第?[\d]+号',
            # Pattern like FXSD-XS-2025-022
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
        """Build the output filename based on mode."""
        # Clean title: remove invalid filename characters
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

    # ===================== Recognize =====================

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

        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.page_names = {}

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

                # Save temp image
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                    tmp_path = tmp.name
                pix.save(tmp_path)

                # OCR
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

                # Extract title and serial number
                title = self._extract_title(ocr_result, max_len=max_title)
                serial = self._extract_serial_number(ocr_result)
                filename = self._build_filename(title, serial, mode, page_num + 1, max_title)

                self.page_names[page_num + 1] = filename

                # Add to tree (thread-safe via root.after)
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

    # ===================== Tree Edit =====================

    def _on_tree_double_click(self, event):
        """Handle double-click to edit a cell in the treeview."""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)

        if not item_id:
            return

        col_idx = int(column.replace("#", "")) - 1
        col_names = ["page", "title", "serial_number", "filename"]

        # Only allow editing filename (col 4)
        if col_idx != 3:
            return

        # Get current value
        current_values = self.tree.item(item_id, "values")
        current_text = current_values[col_idx]

        # Create edit popup
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

    # ===================== Split =====================

    def _start_split(self):
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
            target=self._run_split,
            args=(self.pdf_path.get(), out),
            daemon=True
        )
        thread.start()

    def _run_split(self, pdf_path, output_dir):
        try:
            self._log("正在拆分PDF...")

            os.makedirs(output_dir, exist_ok=True)
            doc = fitz.open(pdf_path)

            # Check for duplicate filenames and add suffix
            filename_count = {}
            final_names = {}

            for page_num, filename in self.page_names.items():
                if filename not in filename_count:
                    filename_count[filename] = 0
                    final_names[page_num] = filename
                else:
                    filename_count[filename] += 1
                    final_names[page_num] = f"{filename}_{filename_count[filename]}"

            created_files = []
            total = len(final_names)

            for idx, (page_num, filename) in enumerate(sorted(final_names.items())):
                self._update_progress(
                    (idx / total) * 100,
                    f"拆分: {idx+1}/{total}"
                )

                # Sanitize filename
                safe_name = re.sub(r'[\\/:*?"<>|]', '', filename)
                output_path = os.path.join(output_dir, f"{safe_name}.pdf")

                # Handle existing files
                if os.path.exists(output_path):
                    base, ext = os.path.splitext(output_path)
                    counter = 1
                    while os.path.exists(f"{base}_{counter}{ext}"):
                        counter += 1
                    output_path = f"{base}_{counter}{ext}"

                new_doc = fitz.open()
                new_doc.insert_pdf(doc, from_page=page_num - 1, to_page=page_num - 1)
                new_doc.save(output_path)
                new_doc.close()

                created_files.append(os.path.basename(output_path))
                self._log(f"  ✅ 第{page_num}页 → {os.path.basename(output_path)}")

            doc.close()
            self._update_progress(100, "完成!")
            self._log(f"\n🎉 拆分完成！共生成 {len(created_files)} 个文件")
            self._log(f"输出目录: {output_dir}")

            self.root.after(0, lambda: messagebox.showinfo(
                "拆分完成",
                f"成功将PDF拆分为 {len(created_files)} 个文件！\n\n输出目录: {output_dir}"
            ))

        except Exception as e:
            self._log(f"❌ 错误: {e}")
            self.root.after(0, lambda: messagebox.showerror("错误", str(e)))

        finally:
            self.is_running = False
            self.root.after(0, lambda: (
                self.recognize_btn.config(state=tk.NORMAL),
                self.split_btn.config(state=tk.NORMAL),
            ))


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
            "以下依赖未安装:\n\n" + "\n".join(missing) +
            "\n\n请在命令行中运行上述pip install命令安装。"
        )

    app = PDFPageSplitterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
