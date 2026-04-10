#!/usr/bin/env python3
"""
GCP 用量明細 PDF 產生器 — 圖形介面
"""

import sys
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re

# 把同目錄的 gcp_pdf_generator 引入
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from gcp_pdf_generator import (
    register_chinese_font,
    load_source_data,
    aggregate_project,
    generate_pdf,
    generate_summary_pdf,
)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('GCP 用量明細 PDF 產生器')
        self.resizable(False, False)
        self._last_dir = os.path.expanduser('~')  # 記憶上次開啟目錄
        self._build_ui()
        self._center()
        # 修正 macOS 視窗焦點問題
        self.after(200, self._force_focus)

    # ── 版面 ────────────────────────────────────────────────────
    def _build_ui(self):
        PAD = dict(padx=12, pady=6)

        # ── Excel 檔案 ──
        frame_file = ttk.LabelFrame(self, text='來源 Excel 明細檔', padding=8)
        frame_file.grid(row=0, column=0, columnspan=3, sticky='ew', **PAD)

        self.var_excel = tk.StringVar()
        ttk.Entry(frame_file, textvariable=self.var_excel, width=55).grid(row=0, column=0, padx=(0,6))
        ttk.Button(frame_file, text='選擇檔案…', command=self._pick_excel).grid(row=0, column=1)

        # ── 輸出資料夾 ──
        frame_out = ttk.LabelFrame(self, text='輸出資料夾', padding=8)
        frame_out.grid(row=1, column=0, columnspan=3, sticky='ew', **PAD)

        self.var_outdir = tk.StringVar()
        ttk.Entry(frame_out, textvariable=self.var_outdir, width=55).grid(row=0, column=0, padx=(0,6))
        ttk.Button(frame_out, text='選擇資料夾…', command=self._pick_outdir).grid(row=0, column=1)

        # ── 設定 ──
        frame_cfg = ttk.LabelFrame(self, text='設定', padding=8)
        frame_cfg.grid(row=2, column=0, columnspan=3, sticky='ew', **PAD)

        # 匯率
        ttk.Label(frame_cfg, text='匯率：').grid(row=0, column=0, sticky='w')
        self.var_rate = tk.StringVar(value='')
        ttk.Entry(frame_cfg, textvariable=self.var_rate, width=12).grid(row=0, column=1, sticky='w', padx=(0,20))

        # 產生總表
        self.var_summary = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame_cfg, text='同時產生專案總表 PDF', variable=self.var_summary).grid(
            row=0, column=2, sticky='w')

        # 指定單一專案（可留空）
        ttk.Label(frame_cfg, text='只產生指定專案（留空＝全部）：').grid(row=1, column=0, columnspan=2, sticky='w', pady=(8,0))
        self.var_project = tk.StringVar(value='')
        ttk.Entry(frame_cfg, textvariable=self.var_project, width=35).grid(
            row=1, column=2, sticky='w', pady=(8,0))

        # ── 執行按鈕 ──
        self.btn_run = ttk.Button(self, text='產生 PDF', command=self._run, width=18)
        self.btn_run.grid(row=3, column=0, columnspan=3, pady=(4, 8))

        # ── 記錄視窗 ──
        frame_log = ttk.LabelFrame(self, text='執行記錄', padding=8)
        frame_log.grid(row=4, column=0, columnspan=3, sticky='nsew', padx=12, pady=(0,12))

        self.log = tk.Text(frame_log, height=14, width=72, state='disabled',
                           font=('Menlo', 11), bg='#1e1e1e', fg='#d4d4d4',
                           insertbackground='white', relief='flat')
        self.log.grid(row=0, column=0, sticky='nsew')

        sb = ttk.Scrollbar(frame_log, command=self.log.yview)
        sb.grid(row=0, column=1, sticky='ns')
        self.log.configure(yscrollcommand=sb.set)

        # 顏色標籤
        self.log.tag_config('ok',    foreground='#4ec9b0')
        self.log.tag_config('err',   foreground='#f44747')
        self.log.tag_config('info',  foreground='#9cdcfe')
        self.log.tag_config('done',  foreground='#dcdcaa')

    # ── 輔助 ────────────────────────────────────────────────────
    def _center(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f'+{(sw-w)//2}+{(sh-h)//2}')

    def _force_focus(self):
        """修正 macOS 開啟時視窗無焦點，導致按鈕需點兩次的問題"""
        self.lift()
        self.focus_force()

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title='選擇 Excel 明細檔',
            initialdir=self._last_dir,
            filetypes=[('Excel 檔案', '*.xlsx *.xls'), ('所有檔案', '*.*')],
        )
        self.after(100, self._force_focus)  # 對話框關閉後重新取回焦點
        if path:
            self._last_dir = os.path.dirname(path)  # 記憶位置
            self.var_excel.set(path)
            if not self.var_outdir.get():
                self.var_outdir.set(os.path.dirname(path))

    def _pick_outdir(self):
        init = self.var_outdir.get() or self._last_dir
        path = filedialog.askdirectory(
            title='選擇輸出資料夾',
            initialdir=init,
        )
        self.after(100, self._force_focus)  # 對話框關閉後重新取回焦點
        if path:
            self._last_dir = path  # 記憶位置
            self.var_outdir.set(path)

    def _log(self, msg, tag=''):
        self.log.configure(state='normal')
        self.log.insert('end', msg + '\n', tag)
        self.log.see('end')
        self.log.configure(state='disabled')

    def _clear_log(self):
        self.log.configure(state='normal')
        self.log.delete('1.0', 'end')
        self.log.configure(state='disabled')

    # ── 執行 ────────────────────────────────────────────────────
    def _run(self):
        excel   = self.var_excel.get().strip()
        outdir  = self.var_outdir.get().strip()
        rate_s  = self.var_rate.get().strip()
        project = self.var_project.get().strip()
        summary = self.var_summary.get()

        # 驗證
        if not excel or not os.path.exists(excel):
            messagebox.showerror('錯誤', '請選擇有效的 Excel 檔案')
            return
        if not outdir:
            messagebox.showerror('錯誤', '請選擇輸出資料夾')
            return
        if summary and not rate_s:
            messagebox.showerror('錯誤', '產生總表需要填入匯率')
            return
        if rate_s:
            try:
                rate = float(rate_s)
            except ValueError:
                messagebox.showerror('錯誤', '匯率格式不正確')
                return
        else:
            rate = None

        self.btn_run.configure(state='disabled')
        self._clear_log()

        # 在背景執行，避免 UI 凍結
        thread = threading.Thread(target=self._worker,
                                  args=(excel, outdir, rate, project, summary),
                                  daemon=True)
        thread.start()

    def _worker(self, excel, outdir, rate, project, summary):
        def log(msg, tag=''):
            self.after(0, self._log, msg, tag)

        try:
            os.makedirs(outdir, exist_ok=True)

            # 月份標籤
            basename = os.path.basename(excel)
            month_label = ''
            m = re.search(r'(\d{4}年\d{2}月)', basename)
            if m:
                month_label = m.group(1)

            log(f'來源：{basename}', 'info')
            log(f'月份：{month_label}', 'info')
            log(f'輸出：{outdir}', 'info')
            log('')

            font_name = register_chinese_font()

            log('讀取 Excel…', 'info')
            projects = load_source_data(excel)
            log(f'找到 {len(projects)} 個 Project ID\n', 'info')

            # 決定要處理哪些專案
            if project:
                if project not in projects:
                    log(f'找不到 Project ID：{project}', 'err')
                    log(f'可用：{", ".join(sorted(projects.keys()))}', 'err')
                    return
                target_pids = [project]
            else:
                target_pids = sorted(projects.keys())

            # 產生個別 PDF
            for pid in target_pids:
                agg = aggregate_project(projects[pid])
                out_path = os.path.join(outdir, f'{pid}.pdf')
                generate_pdf(pid, month_label, agg, out_path, font_name)
                total = sum(r['cost'] for r in agg)
                log(f'  ✓  {pid}.pdf   ({len(agg)} 項, USD {total:,.2f})', 'ok')

            # 產生總表
            if summary and rate is not None:
                log('')
                projects_totals = {
                    pid: sum(r['cost'] for r in aggregate_project(projects[pid]))
                    for pid in sorted(projects.keys())
                }
                m2 = re.search(r'(\d{4})年(\d{2})月', month_label)
                fname = f'尚峪{m2.group(2)}月_GCP專案總表.pdf' if m2 else '尚峪_GCP專案總表.pdf'
                summary_path = os.path.join(outdir, fname)
                generate_summary_pdf(month_label, projects_totals, rate, summary_path, font_name)
                log(f'  ✓  {fname}  （總表）', 'done')

            log('')
            log('完成！', 'done')

            # 開啟輸出資料夾
            self.after(0, lambda: os.system(f'open "{outdir}"'))

        except Exception as e:
            log(f'\n錯誤：{e}', 'err')
            import traceback
            log(traceback.format_exc(), 'err')

        finally:
            self.after(0, lambda: self.btn_run.configure(state='normal'))


if __name__ == '__main__':
    app = App()
    app.mainloop()
