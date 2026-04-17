import pandas as pd
import os
import glob
import json
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, simpledialog

# ------------------------- 配置管理 -------------------------
CONFIG_FILE = "filter_config.json"

class ConfigManager:
    def __init__(self):
        self.whitelist_excludes = {}
        self.global_blacklist = []
        self.whitelist = []
        self.load()

    def load(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.whitelist_excludes = data.get("whitelist_excludes", {})
                self.global_blacklist = data.get("global_blacklist", [])
                self.whitelist = data.get("whitelist", list(self.whitelist_excludes.keys()))
                for word in self.whitelist:
                    if word not in self.whitelist_excludes:
                        self.whitelist_excludes[word] = []
                for word in list(self.whitelist_excludes.keys()):
                    if word not in self.whitelist:
                        self.whitelist.append(word)
            except Exception as e:
                print(f"加载失败: {e}")
                self._init_default()
        else:
            self._init_default()
        self.whitelist.sort()

    def _init_default(self):
        self.whitelist_excludes = {
            "相机": ["线缆", "支架", "电源线", "连接线", "固定板"],
            "工控机": ["电源线", "数据线", "支架"],
            "CPU": []
        }
        self.global_blacklist = ["线缆", "支架", "电源线", "连接线", "数据线", "接头"]
        self.whitelist = list(self.whitelist_excludes.keys())
        self.whitelist.sort()

    def save(self):
        current_keys = set(self.whitelist_excludes.keys())
        whitelist_set = set(self.whitelist)
        for word in whitelist_set - current_keys:
            self.whitelist_excludes[word] = []
        for word in current_keys - whitelist_set:
            del self.whitelist_excludes[word]
        self.whitelist = sorted(list(whitelist_set))
        data = {
            "whitelist": self.whitelist,
            "whitelist_excludes": self.whitelist_excludes,
            "global_blacklist": self.global_blacklist
        }
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def add_whitelist_word(self, word):
        if word in self.whitelist:
            return False
        self.whitelist.append(word)
        self.whitelist_excludes[word] = []
        self.save()
        return True

    def remove_whitelist_word(self, word):
        if word not in self.whitelist:
            return False
        self.whitelist.remove(word)
        if word in self.whitelist_excludes:
            del self.whitelist_excludes[word]
        self.save()
        return True

    def get_whitelist(self):
        return sorted(self.whitelist)

    def get_excludes(self, word):
        return self.whitelist_excludes.get(word, [])

    def set_excludes(self, word, excludes):
        if word in self.whitelist_excludes:
            self.whitelist_excludes[word] = excludes
            self.save()

# ------------------------- 核心筛选逻辑 -------------------------
def match_keywords(text, keywords):
    if not isinstance(text, str):
        return False
    text_lower = text.lower()
    for kw in keywords:
        if kw.lower() in text_lower:
            return True
    return False

def longest_matching_whitelist(material_name, whitelist_words):
    matches = [w for w in whitelist_words if w.lower() in material_name.lower()]
    if not matches:
        return None
    return max(matches, key=len)

def should_keep(material_name, config):
    if not isinstance(material_name, str):
        return True, "其他保留"
    whitelist_words = config.get_whitelist()
    matched_word = longest_matching_whitelist(material_name, whitelist_words)
    if matched_word is None:
        if match_keywords(material_name, config.global_blacklist):
            return False, "全局黑名单删除"
        return True, "其他保留"
    excludes = config.get_excludes(matched_word)
    if match_keywords(material_name, excludes):
        return False, f"白名单排除词删除（{matched_word}的排除词）"
    return True, "白名单保留"

# ------------------------- Excel 处理 -------------------------
def process_excel(file_path, config, log_list):
    log_list.append(f"\n正在处理: {file_path}")
    try:
        xl = pd.ExcelFile(file_path)
        sheet_names = xl.sheet_names
        if not sheet_names:
            log_list.append(f"  警告：{file_path} 中没有工作表，跳过。")
            return
        source_sheet_name = sheet_names[0]
        log_list.append(f"  使用工作表 '{source_sheet_name}' 作为原始数据源。")
        df_source = pd.read_excel(file_path, sheet_name=source_sheet_name, header=None, dtype=str)
        df_source = df_source.fillna("")
        if df_source.shape[1] < 4:
            log_list.append(f"  错误：{file_path} 列数不足 4 列，无法获取 D 列数据，跳过。")
            return
        material_col = df_source.iloc[:, 3].astype(str).fillna("")
        keep_flags = []
        keep_reasons = []
        for mat in material_col:
            keep, reason = should_keep(mat, config)
            keep_flags.append(keep)
            keep_reasons.append(reason)
        keep_mask = pd.Series(keep_flags)
        df_kept = df_source[keep_mask].copy()
        df_removed = df_source[~keep_mask].copy()
        # 保留数据：添加类别列（第一列）
        if not df_kept.empty:
            kept_reasons = [keep_reasons[i] for i, flag in enumerate(keep_flags) if flag]
            df_kept = df_kept.reset_index(drop=True)
            df_kept.insert(0, "类别", kept_reasons)
            df_kept_with_cat = df_kept
        else:
            df_kept_with_cat = pd.DataFrame(columns=["类别"])
        # 删除数据：添加删除原因列（第一列）
        if not df_removed.empty:
            removed_reasons = [keep_reasons[i] for i, flag in enumerate(keep_flags) if not flag]
            df_removed = df_removed.reset_index(drop=True)
            df_removed.insert(0, "删除原因", removed_reasons)
            df_removed_with_reason = df_removed
        else:
            df_removed_with_reason = pd.DataFrame(columns=["删除原因"])
        # 写入Excel
        with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
            df_source.to_excel(writer, sheet_name="源数据", index=False, header=False)
            df_kept_with_cat.to_excel(writer, sheet_name="删减后", index=False, header=False)
            df_removed_with_reason.to_excel(writer, sheet_name="被移除", index=False, header=False)
        log_list.append(f"  处理完成：原始 {len(df_source)} 行，保留 {len(df_kept)} 行，删除 {len(df_removed)} 行。")
    except Exception as e:
        log_list.append(f"  处理 {file_path} 时出错：{e}")

def run_batch_processing(config, log_callback):
    excel_files = glob.glob(os.path.join(".", "*.xlsx"))
    # 排除汇总文件（如果存在）
    folder_name = os.path.basename(os.getcwd())
    summary_file = f"{folder_name}.xlsx"
    if summary_file in excel_files:
        excel_files.remove(summary_file)
    if not excel_files:
        log_callback("在当前目录中没有找到任何 .xlsx 文件。")
        return
    log_callback(f"找到 {len(excel_files)} 个 Excel 文件，开始批量处理...")
    for file_path in excel_files:
        temp_log = []
        process_excel(file_path, config, temp_log)
        for msg in temp_log:
            log_callback(msg)
    log_callback("\n所有文件处理完毕。")

# ------------------------- 汇总报告生成 -------------------------
def generate_summary_report(log_callback):
    """生成汇总报告：第一个sheet为文件列表，后续sheet为各文件的删减后数据"""
    folder_name = os.path.basename(os.getcwd())
    summary_file = f"{folder_name}.xlsx"
    # 获取当前目录下所有xlsx文件（排除汇总文件本身）
    all_files = glob.glob(os.path.join(".", "*.xlsx"))
    if summary_file in all_files:
        all_files.remove(summary_file)
    if not all_files:
        log_callback("没有找到任何 .xlsx 文件，无法生成汇总报告。")
        return
    # 准备第一个sheet的数据：机器名列表
    file_names = [os.path.splitext(os.path.basename(f))[0] for f in all_files]
    df_index = pd.DataFrame({
        "项目号": [""] * len(file_names),
        "机器名": file_names
    })
    # 使用ExcelWriter
    try:
        with pd.ExcelWriter(summary_file, engine="openpyxl") as writer:
            df_index.to_excel(writer, sheet_name="目录", index=False)
            # 遍历每个文件，读取第二个sheet（索引1，即“删减后”）
            for fname in all_files:
                sheet_name = os.path.splitext(os.path.basename(fname))[0]
                # Excel sheet名称长度限制31字符
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                try:
                    df_sheet = pd.read_excel(fname, sheet_name=1, header=None)  # 第二个sheet
                    df_sheet.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                    log_callback(f"已添加 {fname} 的删减后数据 -> sheet: {sheet_name}")
                except Exception as e:
                    log_callback(f"警告：读取 {fname} 的第二个sheet失败：{e}")
        log_callback(f"汇总报告已生成：{summary_file}")
    except Exception as e:
        log_callback(f"生成汇总报告失败：{e}")

# ------------------------- GUI 应用程序 -------------------------
class FilterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("SmartBOMFilter - 智能物料筛选器")
        self.root.geometry("1200x850")

        default_font = ("微软雅黑", 11)
        self.root.option_add("*Font", default_font)

        self.config = ConfigManager()

        # 分页参数
        self.whitelist_page = 0
        self.whitelist_page_size = 15
        self.blacklist_page = 0
        self.blacklist_rows_per_page = 18
        self.blacklist_cols = 3
        self.blacklist_page_size = self.blacklist_rows_per_page * self.blacklist_cols

        self._build_ui()
        self._refresh_whitelist_display()
        self._refresh_blacklist_display()

    def _build_ui(self):
        main_pane = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 左侧白名单
        left_frame = ttk.LabelFrame(main_pane, text="总白名单（快捷增删，排除列表可独立编辑）")
        main_pane.add(left_frame, weight=1)

        columns = ("白名单词", "操作")
        self.whitelist_tree = ttk.Treeview(left_frame, columns=columns, show="headings",
                                           height=self.whitelist_page_size)
        self.whitelist_tree.heading("白名单词", text="白名单词")
        self.whitelist_tree.heading("操作", text="操作")
        self.whitelist_tree.column("白名单词", width=180, anchor="center")
        self.whitelist_tree.column("操作", width=120, anchor="center")
        style = ttk.Style()
        style.configure("Treeview", font=("微软雅黑", 11), rowheight=28)
        style.configure("Treeview.Heading", font=("微软雅黑", 11, "bold"))
        self.whitelist_tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.whitelist_tree.bind("<ButtonRelease-1>", self.on_whitelist_click)

        w_panel = ttk.Frame(left_frame)
        w_panel.pack(fill=tk.X, pady=5)
        ttk.Button(w_panel, text="上一页", command=self.whitelist_prev_page).pack(side=tk.LEFT, padx=2)
        self.whitelist_page_label = ttk.Label(w_panel, text="")
        self.whitelist_page_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(w_panel, text="下一页", command=self.whitelist_next_page).pack(side=tk.LEFT, padx=2)
        ttk.Button(w_panel, text="删除当前选中词", command=self.delete_selected_whitelist).pack(side=tk.RIGHT, padx=2)

        add_w_frame = ttk.Frame(left_frame)
        add_w_frame.pack(fill=tk.X, pady=5, padx=5)
        self.whitelist_entry = ttk.Entry(add_w_frame, font=("微软雅黑", 11))
        self.whitelist_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(add_w_frame, text="添加白名单词", command=self.add_whitelist_word).pack(side=tk.RIGHT)

        # 右侧黑名单
        right_frame = ttk.LabelFrame(main_pane, text="全局黑名单（不包含任何白名单词时命中即删除）")
        main_pane.add(right_frame, weight=1)

        black_grid_frame = ttk.Frame(right_frame)
        black_grid_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.black_listboxes = []
        for i in range(self.blacklist_cols):
            col_frame = ttk.Frame(black_grid_frame)
            col_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)
            lb = tk.Listbox(col_frame, font=("微软雅黑", 11),
                            height=self.blacklist_rows_per_page,
                            selectmode=tk.SINGLE)
            lb.pack(fill=tk.BOTH, expand=True)
            self.black_listboxes.append(lb)
            lb.bind("<<ListboxSelect>>", self.on_blacklist_select)

        b_panel = ttk.Frame(right_frame)
        b_panel.pack(fill=tk.X, pady=5)
        ttk.Button(b_panel, text="上一页", command=self.blacklist_prev_page).pack(side=tk.LEFT, padx=2)
        self.blacklist_page_label = ttk.Label(b_panel, text="")
        self.blacklist_page_label.pack(side=tk.LEFT, padx=10)
        ttk.Button(b_panel, text="下一页", command=self.blacklist_next_page).pack(side=tk.LEFT, padx=2)
        ttk.Button(b_panel, text="删除选中黑名单词", command=self.delete_selected_blacklist).pack(side=tk.RIGHT, padx=2)

        add_b_frame = ttk.Frame(right_frame)
        add_b_frame.pack(fill=tk.X, pady=5, padx=5)
        self.blacklist_entry = ttk.Entry(add_b_frame, font=("微软雅黑", 11))
        self.blacklist_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(add_b_frame, text="添加黑名单词", command=self.add_blacklist_word).pack(side=tk.RIGHT)

        # 底部按钮区域
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0,10))

        btn_frame = ttk.Frame(bottom_frame)
        btn_frame.pack(pady=5)
        run_btn = ttk.Button(btn_frame, text="运行批量筛选", command=self.run_filter, width=20)
        run_btn.pack(side=tk.LEFT, padx=5)
        summary_btn = ttk.Button(btn_frame, text="生成汇总报告", command=self.generate_summary, width=20)
        summary_btn.pack(side=tk.LEFT, padx=5)

        log_frame = ttk.LabelFrame(bottom_frame, text="处理日志")
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 10), height=10)
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        self.current_selected_black_word = None

    # ------------------ 白名单操作 ------------------
    def _get_whitelist_page_items(self):
        words = self.config.get_whitelist()
        start = self.whitelist_page * self.whitelist_page_size
        end = start + self.whitelist_page_size
        return words[start:end], len(words)

    def _refresh_whitelist_display(self):
        for item in self.whitelist_tree.get_children():
            self.whitelist_tree.delete(item)
        page_words, total = self._get_whitelist_page_items()
        for word in page_words:
            self.whitelist_tree.insert("", tk.END, values=(word, "📝 排除列表"), tags=(word,))
        total_pages = max(1, (total + self.whitelist_page_size - 1) // self.whitelist_page_size)
        self.whitelist_page_label.config(text=f"第 {self.whitelist_page+1} / {total_pages} 页 (共{total}词)")

    def _highlight_whitelist_word(self, word):
        for item in self.whitelist_tree.get_children():
            if self.whitelist_tree.item(item, "values")[0] == word:
                self.whitelist_tree.selection_set(item)
                self.whitelist_tree.see(item)
                break

    def on_whitelist_click(self, event):
        region = self.whitelist_tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        column = self.whitelist_tree.identify_column(event.x)
        if column != "#2":
            return
        item = self.whitelist_tree.identify_row(event.y)
        if not item:
            return
        word = self.whitelist_tree.item(item, "values")[0]
        self.edit_exclude_list(word)

    def edit_exclude_list(self, word):
        win = tk.Toplevel(self.root)
        win.title(f"编辑“{word}”的排除列表")
        win.geometry("500x400")
        win.transient(self.root)
        win.grab_set()
        frame = ttk.Frame(win, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=f"排除词（包含以下任意词即删除）", font=("微软雅黑", 11, "bold")).pack(anchor=tk.W)
        listbox_frame = ttk.Frame(frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        listbox = tk.Listbox(listbox_frame, font=("微软雅黑", 11))
        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=listbox.yview)
        listbox.config(yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        excludes = self.config.get_excludes(word).copy()
        for ex in excludes:
            listbox.insert(tk.END, ex)
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=5)

        def add_exclude():
            new_ex = simpledialog.askstring("添加排除词", f"为“{word}”添加排除词", parent=win)
            if new_ex and new_ex.strip():
                new_ex = new_ex.strip()
                if new_ex not in excludes:
                    excludes.append(new_ex)
                    listbox.insert(tk.END, new_ex)
                    self.config.set_excludes(word, excludes)
                else:
                    messagebox.showwarning("已存在", f"排除词“{new_ex}”已存在")

        def delete_exclude():
            sel = listbox.curselection()
            if not sel:
                messagebox.showinfo("提示", "请先选中要删除的排除词", parent=win)
                return
            ex = listbox.get(sel[0])
            if messagebox.askyesno("确认删除", f"确定删除排除词“{ex}”吗？", parent=win):
                excludes.remove(ex)
                listbox.delete(sel[0])
                self.config.set_excludes(word, excludes)

        ttk.Button(btn_frame, text="添加排除词", command=add_exclude).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="删除选中排除词", command=delete_exclude).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="关闭", command=win.destroy).pack(side=tk.RIGHT, padx=2)

    def whitelist_prev_page(self):
        if self.whitelist_page > 0:
            self.whitelist_page -= 1
            self._refresh_whitelist_display()

    def whitelist_next_page(self):
        total_words = len(self.config.get_whitelist())
        max_page = max(0, (total_words - 1) // self.whitelist_page_size)
        if self.whitelist_page < max_page:
            self.whitelist_page += 1
            self._refresh_whitelist_display()

    def add_whitelist_word(self):
        new_word = self.whitelist_entry.get().strip()
        if not new_word:
            messagebox.showwarning("输入为空", "请输入白名单词")
            return
        if self.config.add_whitelist_word(new_word):
            self.whitelist_entry.delete(0, tk.END)
            words = self.config.get_whitelist()
            try:
                idx = words.index(new_word)
            except ValueError:
                return
            page = idx // self.whitelist_page_size
            self.whitelist_page = page
            self._refresh_whitelist_display()
            self._highlight_whitelist_word(new_word)
        else:
            messagebox.showwarning("已存在", f"白名单词“{new_word}”已存在")

    def delete_selected_whitelist(self):
        selected = self.whitelist_tree.selection()
        if not selected:
            messagebox.showinfo("提示", "请先点击选中要删除的白名单词（点击左侧行首）")
            return
        word = self.whitelist_tree.item(selected[0], "values")[0]
        if messagebox.askyesno("确认删除", f"确定删除白名单词“{word}”及其所有排除词吗？"):
            self.config.remove_whitelist_word(word)
            total = len(self.config.get_whitelist())
            if total == 0:
                self.whitelist_page = 0
            elif self.whitelist_page * self.whitelist_page_size >= total:
                self.whitelist_page = max(0, (total - 1) // self.whitelist_page_size)
            self._refresh_whitelist_display()

    # ------------------ 黑名单操作 ------------------
    def _get_blacklist_page_items(self):
        words = sorted(self.config.global_blacklist)
        start = self.blacklist_page * self.blacklist_page_size
        end = start + self.blacklist_page_size
        page_words = words[start:end]
        matrix = []
        for i in range(0, len(page_words), self.blacklist_cols):
            row = page_words[i:i+self.blacklist_cols]
            row += [None] * (self.blacklist_cols - len(row))
            matrix.append(row)
        while len(matrix) < self.blacklist_rows_per_page:
            matrix.append([None] * self.blacklist_cols)
        return matrix, len(words)

    def _refresh_blacklist_display(self):
        for lb in self.black_listboxes:
            lb.delete(0, tk.END)
        matrix, total = self._get_blacklist_page_items()
        for col in range(self.blacklist_cols):
            lb = self.black_listboxes[col]
            for row in range(self.blacklist_rows_per_page):
                word = matrix[row][col]
                if word:
                    lb.insert(tk.END, word)
        total_pages = max(1, (total + self.blacklist_page_size - 1) // self.blacklist_page_size)
        self.blacklist_page_label.config(text=f"第 {self.blacklist_page+1} / {total_pages} 页 (共{total}词)")

    def _highlight_blacklist_word(self, word):
        sorted_words = sorted(self.config.global_blacklist)
        start = self.blacklist_page * self.blacklist_page_size
        end = start + self.blacklist_page_size
        page_words = sorted_words[start:end]
        try:
            idx = page_words.index(word)
        except ValueError:
            return
        col = idx % self.blacklist_cols
        row = idx // self.blacklist_cols
        if row < self.blacklist_rows_per_page:
            lb = self.black_listboxes[col]
            lb.selection_clear(0, tk.END)
            lb.selection_set(row)
            lb.see(row)

    def on_blacklist_select(self, event):
        widget = event.widget
        if not isinstance(widget, tk.Listbox):
            return
        selection = widget.curselection()
        if selection:
            word = widget.get(selection[0])
            self.current_selected_black_word = word
        else:
            self.current_selected_black_word = None

    def blacklist_prev_page(self):
        if self.blacklist_page > 0:
            self.blacklist_page -= 1
            self._refresh_blacklist_display()

    def blacklist_next_page(self):
        total = len(self.config.global_blacklist)
        max_page = max(0, (total - 1) // self.blacklist_page_size)
        if self.blacklist_page < max_page:
            self.blacklist_page += 1
            self._refresh_blacklist_display()

    def add_blacklist_word(self):
        new_word = self.blacklist_entry.get().strip()
        if not new_word:
            messagebox.showwarning("输入为空", "请输入黑名单词")
            return
        if new_word in self.config.global_blacklist:
            messagebox.showwarning("已存在", f"黑名单词“{new_word}”已存在")
            return
        self.config.global_blacklist.append(new_word)
        self.config.save()
        self.blacklist_entry.delete(0, tk.END)
        sorted_words = sorted(self.config.global_blacklist)
        try:
            idx = sorted_words.index(new_word)
        except ValueError:
            return
        page = idx // self.blacklist_page_size
        self.blacklist_page = page
        self._refresh_blacklist_display()
        self._highlight_blacklist_word(new_word)

    def delete_selected_blacklist(self):
        if not self.current_selected_black_word:
            messagebox.showinfo("提示", "请先点击选中要删除的黑名单词")
            return
        word = self.current_selected_black_word
        if messagebox.askyesno("确认删除", f"确定从全局黑名单中删除“{word}”吗？"):
            self.config.global_blacklist.remove(word)
            self.config.save()
            total = len(self.config.global_blacklist)
            if total == 0:
                self.blacklist_page = 0
            elif self.blacklist_page * self.blacklist_page_size >= total:
                self.blacklist_page = max(0, (total - 1) // self.blacklist_page_size)
            self._refresh_blacklist_display()
            self.current_selected_black_word = None

    # ------------------ 运行筛选 ------------------
    def run_filter(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        def log_callback(msg):
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
            self.root.update_idletasks()

        log_callback("开始批量处理...")
        run_batch_processing(self.config, log_callback)
        log_callback("处理完成。")

    # ------------------ 生成汇总报告 ------------------
    def generate_summary(self):
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

        def log_callback(msg):
            self.log_text.config(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
            self.root.update_idletasks()

        log_callback("开始生成汇总报告...")
        generate_summary_report(log_callback)
        log_callback("汇总报告生成完毕。")

# ------------------------- 启动 -------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = FilterApp(root)
    root.mainloop()