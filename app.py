import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd

# 默认正则表达式规则
DEFAULT_MAT_REGEX = r"(?:材料|材质|Material)[：:\s]*([A-Za-z0-9\u4e00-\u9fa5\-_/\.]+)"
DEFAULT_WGT_REGEX = r"(?:重量|单重|净重|总重|Weight)[：:\s]*([0-9]+(?:\.[0-9]+)?\s*(?:kg|g|t|吨|千克|克)?)"

# 跳过扫描的系统隐藏目录
SKIP_DIRS = {"$recycle.bin", "system volume information", "windows", "program files", "program files (x86)", "appdata"}


class DrawingExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF工程图纸材料与重量批量提取工具 v1.0")
        self.root.geometry("760x560")
        self.root.resizable(True, True)

        self.is_running = False
        self.stop_flag = False

        self._build_ui()

    def _build_ui(self):
        pad_opts = {"padx": 12, "pady": 6}

        # 1. 目录与输出设置区
        frame_paths = ttk.LabelFrame(self.root, text=" 1. 路径设置 ", padding=10)
        frame_paths.pack(fill="x", **pad_opts)

        # 输入目录
        ttk.Label(frame_paths, text="图纸文件夹/盘符:").grid(row=0, column=0, sticky="w", pady=4)
        self.var_input_dir = tk.StringVar()
        ttk.Entry(frame_paths, textvariable=self.var_input_dir, width=62).grid(row=0, column=1, padx=6, pady=4, sticky="ew")
        ttk.Button(frame_paths, text="选择目录...", command=self.browse_input_dir).grid(row=0, column=2, pady=4)

        # 输出文件
        ttk.Label(frame_paths, text="输出Excel路径:").grid(row=1, column=0, sticky="w", pady=4)
        self.var_output_file = tk.StringVar()
        ttk.Entry(frame_paths, textvariable=self.var_output_file, width=62).grid(row=1, column=1, padx=6, pady=4, sticky="ew")
        ttk.Button(frame_paths, text="保存位置...", command=self.browse_output_file).grid(row=1, column=2, pady=4)

        frame_paths.columnconfigure(1, weight=1)

        # 2. 提取规则配置区（可展开微调）
        frame_rules = ttk.LabelFrame(self.root, text=" 2. 匹配规则配置（支持正则表达式，如无特殊需求保持默认即可） ", padding=10)
        frame_rules.pack(fill="x", **pad_opts)

        ttk.Label(frame_rules, text="材料匹配正则:").grid(row=0, column=0, sticky="w", pady=4)
        self.var_mat_regex = tk.StringVar(value=DEFAULT_MAT_REGEX)
        ttk.Entry(frame_rules, textvariable=self.var_mat_regex).grid(row=0, column=1, padx=6, pady=4, sticky="ew")

        ttk.Label(frame_rules, text="重量匹配正则:").grid(row=1, column=0, sticky="w", pady=4)
        self.var_wgt_regex = tk.StringVar(value=DEFAULT_WGT_REGEX)
        ttk.Entry(frame_rules, textvariable=self.var_wgt_regex).grid(row=1, column=1, padx=6, pady=4, sticky="ew")

        ttk.Button(frame_rules, text="恢复默认规则", command=self.reset_rules).grid(row=0, column=2, rowspan=2, padx=4)
        frame_rules.columnconfigure(1, weight=1)

        # 3. 控制按钮与进度条
        frame_ctrl = ttk.Frame(self.root)
        frame_ctrl.pack(fill="x", **pad_opts)

        self.btn_start = ttk.Button(frame_ctrl, text="▶ 开始扫描并导出 Excel", command=self.start_task)
        self.btn_start.pack(side="left", ipadx=12, ipady=4)

        self.btn_stop = ttk.Button(frame_ctrl, text="■ 停止", command=self.stop_task, state="disabled")
        self.btn_stop.pack(side="left", padx=10, ipadx=8, ipady=4)

        self.lbl_status = ttk.Label(frame_ctrl, text="就绪", foreground="#333333")
        self.lbl_status.pack(side="right")

        self.progress = ttk.Progressbar(self.root, mode="indeterminate")
        self.progress.pack(fill="x", padx=12, pady=2)

        # 4. 运行日志区
        frame_log = ttk.LabelFrame(self.root, text=" 3. 运行日志 ", padding=8)
        frame_log.pack(fill="both", expand=True, **pad_opts)

        self.txt_log = tk.Text(frame_log, height=14, state="disabled", font=("Consolas", 9), bg="#F8F9FA")
        scrollbar = ttk.Scrollbar(frame_log, orient="vertical", command=self.txt_log.yview)
        self.txt_log.configure(yscrollcommand=scrollbar.set)

        self.txt_log.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

    def log(self, msg):
        def _append():
            self.txt_log.configure(state="normal")
            self.txt_log.insert("end", msg + "\n")
            self.txt_log.see("end")
            self.txt_log.configure(state="disabled")
        self.root.after(0, _append)

    def browse_input_dir(self):
        folder = filedialog.askdirectory(title="选择包含PDF图纸的文件夹或盘符")
        if folder:
            self.var_input_dir.set(folder)
            if not self.var_output_file.get():
                default_out = os.path.join(folder, "图纸材料与重量汇总.xlsx")
                self.var_output_file.set(default_out)

    def browse_output_file(self):
        filepath = filedialog.asksaveasfilename(
            title="选择Excel保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="图纸材料与重量汇总.xlsx"
        )
        if filepath:
            self.var_output_file.set(filepath)

    def reset_rules(self):
        self.var_mat_regex.set(DEFAULT_MAT_REGEX)
        self.var_wgt_regex.set(DEFAULT_WGT_REGEX)

    def start_task(self):
        input_dir = self.var_input_dir.get().strip()
        output_file = self.var_output_file.get().strip()

        if not input_dir or not os.path.exists(input_dir):
            messagebox.showwarning("路径错误", "请选择有效的图纸文件夹或盘符！")
            return
        if not output_file:
            messagebox.showwarning("路径错误", "请指定输出的 Excel 文件路径！")
            return

        try:
            mat_pattern = re.compile(self.var_mat_regex.get(), re.IGNORECASE)
            wgt_pattern = re.compile(self.var_wgt_regex.get(), re.IGNORECASE)
        except re.error as e:
            messagebox.showerror("正则表达式错误", f"规则格式有误：\n{e}")
            return

        self.is_running = True
        self.stop_flag = False
        self.btn_start.configure(state="disabled")
        self.btn_stop.configure(state="normal")
        self.progress.start(12)
        self.lbl_status.configure(text="正在扫描...")

        self.txt_log.configure(state="normal")
        self.txt_log.delete("1.0", "end")
        self.txt_log.configure(state="disabled")

        thread = threading.Thread(
            target=self._worker,
            args=(input_dir, output_file, mat_pattern, wgt_pattern),
            daemon=True
        )
        thread.start()

    def stop_task(self):
        if self.is_running:
            self.stop_flag = True
            self.lbl_status.configure(text="正在停止...")
            self.log("[提示] 用户请求停止扫描，正在完成当前文件...")

    def _extract_single_pdf(self, pdf_path, mat_pattern, wgt_pattern):
        full_text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text += text + "\n"
        except Exception as e:
            return "读取失败", f"错误: {str(e)}"

        if not full_text.strip():
            return "无文本(可能是扫描件)", "需OCR识别"

        materials = mat_pattern.findall(full_text)
        weights = wgt_pattern.findall(full_text)

        mat_str = ", ".join(dict.fromkeys(m.strip() for m in materials)) if materials else "未找到"
        wgt_str = ", ".join(dict.fromkeys(w.strip() for w in weights)) if weights else "未找到"
        return mat_str, wgt_str

    def _worker(self, input_dir, output_file, mat_pattern, wgt_pattern):
        results = []
        count = 0
        self.log(f"开始扫描: {input_dir}")

        for root_dir, dirs, files in os.walk(input_dir):
            # 过滤系统目录
            dirs[:] = [d for d in dirs if d.lower() not in SKIP_DIRS]

            if self.stop_flag:
                break

            for file in files:
                if self.stop_flag:
                    break
                if file.lower().endswith(".pdf"):
                    count += 1
                    full_path = os.path.join(root_dir, file)
                    self.root.after(0, lambda c=count, f=file: self.lbl_status.configure(text=f"已处理 {c} 个: {f[:25]}"))

                    mat, wgt = self._extract_single_pdf(full_path, mat_pattern, wgt_pattern)
                    self.log(f"[{count}] {file}  -->  材料: {mat} | 重量: {wgt}")

                    results.append({
                        "序号": count,
                        "文件名": file,
                        "材料": mat,
                        "重量": wgt,
                        "所在文件夹": root_dir,
                        "完整路径": full_path
                    })

        self.root.after(0, self._finish_task, results, output_file)

    def _finish_task(self, results, output_file):
        self.progress.stop()
        self.is_running = False
        self.btn_start.configure(state="normal")
        self.btn_stop.configure(state="disabled")

        if not results:
            self.lbl_status.configure(text="未找到PDF文件")
            self.log("扫描结束，未提取到任何记录。")
            messagebox.showinfo("完成", "未在指定目录下找到 PDF 文件。")
            return

        try:
            df = pd.DataFrame(results)
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="图纸提取结果")
                ws = writer.sheets["图纸提取结果"]
                # 自动调整列宽
                col_widths = {"A": 8, "B": 32, "C": 22, "D": 18, "E": 40, "F": 55}
                for col, width in col_widths.items():
                    ws.column_dimensions[col].width = width

            self.lbl_status.configure(text=f"完成！共处理 {len(results)} 个文件")
            self.log(f"\n成功导出至: {output_file}")
            if messagebox.askyesno("导出成功", f"共处理 {len(results)} 个 PDF 图纸！\n结果已保存至：\n{output_file}\n\n是否立即打开该 Excel 文件？"):
                os.startfile(output_file)
        except Exception as e:
            self.lbl_status.configure(text="导出失败")
            self.log(f"[错误] 导出 Excel 失败: {e}")
            messagebox.showerror("导出失败", f"无法写入 Excel 文件（请检查文件是否正在被打开）：\n{e}")


if __name__ == "__main__":
    root = tk.Tk()
    app = DrawingExtractorApp(root)
    root.mainloop()
