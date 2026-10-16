# -*- coding: utf-8 -*-
"""脱敏校验：确认输出文件中不再包含准考证号（15位数字）与姓名列。"""
import os
import re
import sys

from openpyxl import load_workbook

ID_RX = re.compile(r'^\d{13,16}$')
DIGIT_RX = re.compile(r'^\d{13,16}$')


def scan(path, limit=None):
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[wb.sheetnames[0]]
    try:
        ws.reset_dimensions()
    except Exception:
        pass
    header = None
    ids = 0
    rows = 0
    ncols = 0
    name_col_like = 0
    for row in ws.iter_rows(values_only=True):
        if row is None or not any(c is not None and str(c).strip() for c in row):
            continue
        vals = [('' if c is None else str(c).strip()) for c in row]
        if header is None:
            header = vals
            ncols = len(vals)
            continue
        rows += 1
        for v in vals:
            if DIGIT_RX.match(v):
                ids += 1
        if limit and rows >= limit:
            break
    wb.close()
    return dict(header=header, ncols=ncols, rows=rows, id_like=ids)


def main():
    src_dir, out_dir = sys.argv[1], sys.argv[2]
    print('%-40s | %-26s | %-26s' % ('文件（年份）', '原始', '脱敏后'))
    print('-' * 100)
    for f in sorted(os.listdir(out_dir)):
        if not f.endswith('.xlsx'):
            continue
        year = re.search(r'(\d{4})年度', f).group(1)
        src = None
        for g in os.listdir(src_dir):
            if year + '年度' in g and g.lower().endswith(('.xlsx', '.xls')):
                src = os.path.join(src_dir, g)
                break
        a = scan(src)
        b = scan(os.path.join(out_dir, f))
        print('%s | 列%d 行%d 准考证号样%d | 列%d 行%d 准考证号样%d' % (
            year, a['ncols'], a['rows'], a['id_like'], b['ncols'], b['rows'], b['id_like']))
    print('\n脱敏后表头：', b['header'])
    print('原始表头：  ', a['header'])


if __name__ == '__main__':
    main()
