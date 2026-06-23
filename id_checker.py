#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
身份证号码批量校验工具
支持导入 Excel/CSV/txt，逐条校验后导出结果。
"""

import re
import datetime
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ---------------- 校验核心逻辑 ----------------

# 省份代码表（前两位）
PROVINCE = {
    '11': '北京', '12': '天津', '13': '河北', '14': '山西', '15': '内蒙古',
    '21': '辽宁', '22': '吉林', '23': '黑龙江', '31': '上海', '32': '江苏',
    '33': '浙江', '34': '安徽', '35': '福建', '36': '江西', '37': '山东',
    '41': '河南', '42': '湖北', '43': '湖南', '44': '广东', '45': '广西',
    '46': '海南', '50': '重庆', '51': '四川', '52': '贵州', '53': '云南',
    '54': '西藏', '61': '陕西', '62': '甘肃', '63': '青海', '64': '宁夏',
    '65': '新疆', '71': '台湾', '81': '香港', '82': '澳门', '83': '台湾',
}

WEIGHTS = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
CHECK_CODES = '10X98765432'


def check_id(raw):
    """返回 (是否有效, 说明, 附加信息字典)"""
    num = str(raw).strip().upper()
    info = {'生日': '', '性别': '', '省份': ''}

    if len(num) != 18:
        return False, '长度不是18位', info
    if not re.match(r'^\d{17}[\dX]$', num):
        return False, '含非法字符', info

    # 省份
    prov = PROVINCE.get(num[:2])
    if prov is None:
        return False, '省份代码无效', info
    info['省份'] = prov

    # 出生日期
    try:
        y, m, d = int(num[6:10]), int(num[10:12]), int(num[12:14])
        bd = datetime.date(y, m, d)
        if bd > datetime.date.today():
            return False, '出生日期晚于今天', info
        if y < 1900:
            return False, '出生年份过早', info
        info['生日'] = bd.isoformat()
    except ValueError:
        return False, '出生日期非法', info

    # 性别（第17位奇男偶女）
    info['性别'] = '男' if int(num[16]) % 2 == 1 else '女'

    # 校验码
    s = sum(int(num[i]) * WEIGHTS[i] for i in range(17))
    if CHECK_CODES[s % 11] != num[17]:
        return False, '校验码不匹配', info

    return True, '有效', info


# ---------------- 文件读取 ----------------

def read_ids(path):
    """从文件读取身份证号列表，返回 list[str]。"""
    lower = path.lower()
    if lower.endswith(('.xlsx', '.xls')):
        import pandas as pd
        df = pd.read_excel(path, dtype=str, header=None)
        return _extract_from_df(df)
    elif lower.endswith('.csv'):
        import pandas as pd
        df = pd.read_csv(path, dtype=str, header=None)
        return _extract_from_df(df)
    else:  # txt 等纯文本，每行一个
        with open(path, 'r', encoding='utf-8-sig') as f:
            return [line.strip() for line in f if line.strip()]


def _extract_from_df(df):
    """找出最像身份证号的那一列；找不到就把所有非空单元格都收进来。"""
    best_col, best_hits = None, -1
    for col in df.columns:
        vals = df[col].dropna().astype(str)
        hits = sum(1 for v in vals if re.match(r'^\s*\d{17}[\dXx]\s*$', v))
        if hits > best_hits:
            best_hits, best_col = hits, col
    if best_hits > 0:
        return [str(v).strip() for v in df[best_col].dropna()]
    # 兜底：全部单元格
    out = []
    for col in df.columns:
        out += [str(v).strip() for v in df[col].dropna()]
    return out


def write_results(path, rows):
    """rows: list of dict。导出为 Excel 或 CSV。"""
    import pandas as pd
    df = pd.DataFrame(rows)
    if path.lower().endswith('.csv'):
        df.to_csv(path, index=False, encoding='utf-8-sig')
    else:
        df.to_excel(path, index=False)


# ---------------- 图形界面 ----------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('身份证批量校验工具')
        self.geometry('760x520')
        self.results = []
        self._build()

    def _build(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill='x')

        ttk.Button(top, text='导入文件', command=self.on_import).pack(side='left')
        ttk.Button(top, text='导出结果', command=self.on_export).pack(side='left', padx=6)
        self.summary = ttk.Label(top, text='请导入 Excel / CSV / txt 文件')
        self.summary.pack(side='left', padx=12)

        cols = ('号码', '结果', '说明', '生日', '性别', '省份')
        self.tree = ttk.Treeview(self, columns=cols, show='headings')
        widths = (190, 60, 130, 100, 50, 90)
        for c, w in zip(cols, widths):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor='w')
        self.tree.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        self.tree.tag_configure('ok', background='#e8f5e9')
        self.tree.tag_configure('bad', background='#ffebee')

    def on_import(self):
        path = filedialog.askopenfilename(
            title='选择文件',
            filetypes=[('表格/文本', '*.xlsx *.xls *.csv *.txt'), ('所有文件', '*.*')])
        if not path:
            return
        self.summary.config(text='读取中…')
        threading.Thread(target=self._do_check, args=(path,), daemon=True).start()

    def _do_check(self, path):
        try:
            ids = read_ids(path)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror('读取失败', str(e)))
            self.after(0, lambda: self.summary.config(text='读取失败'))
            return

        rows, ok = [], 0
        for raw in ids:
            valid, msg, info = check_id(raw)
            ok += valid
            rows.append({
                '号码': str(raw).strip(), '结果': '有效' if valid else '无效',
                '说明': msg, '生日': info['生日'],
                '性别': info['性别'], '省份': info['省份'],
            })
        self.results = rows
        self.after(0, lambda: self._render(rows, ok))

    def _render(self, rows, ok):
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            tag = 'ok' if r['结果'] == '有效' else 'bad'
            self.tree.insert('', 'end', tags=(tag,), values=(
                r['号码'], r['结果'], r['说明'], r['生日'], r['性别'], r['省份']))
        total = len(rows)
        self.summary.config(text=f'共 {total} 条，有效 {ok}，无效 {total - ok}')

    def on_export(self):
        if not self.results:
            messagebox.showinfo('提示', '没有可导出的结果')
            return
        path = filedialog.asksaveasfilename(
            title='保存结果', defaultextension='.xlsx',
            filetypes=[('Excel', '*.xlsx'), ('CSV', '*.csv')])
        if not path:
            return
        try:
            write_results(path, self.results)
            messagebox.showinfo('完成', f'已导出到\n{path}')
        except Exception as e:
            messagebox.showerror('导出失败', str(e))


if __name__ == '__main__':
    App().mainloop()
