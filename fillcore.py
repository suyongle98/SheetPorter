# -*- coding: utf-8 -*-
"""自动填表 · 核心引擎（不依赖界面，可单独测试）

设计要点：
1. 源表、目标表都按「表头名称」定位列，列顺序变了也不怕；
2. 目标表数据区容量靠边框识别，数据多了自动插行并复制格式；
3. 表尾（备注/签名）保持不动，只整体下移；
4. 转换规则可扩展（取斜杠后、取斜杠前、固定值……）。
"""
import os
import re
import shutil
import datetime
import glob
import tempfile
from copy import copy

import openpyxl
from openpyxl.worksheet.cell_range import CellRange

PLACEHOLDER = re.compile(r"\{([^{}]+)\}")


# ---------------------------------------------------------------- 基础工具

def date_context(d=None):
    """生成标题/文件名里可用的日期变量。"""
    d = d or datetime.date.today()
    return {
        "年": str(d.year),
        "月": str(d.month),
        "日": str(d.day),
        "yy": "%02d" % (d.year % 100),
        "mm": "%02d" % d.month,
        "dd": "%02d" % d.day,
        "yyyy": "%04d" % d.year,
        "MM": "%02d" % d.month,
        "DD": "%02d" % d.day,
        "date": d.strftime("%Y-%m-%d"),
    }


def render(text, ctx):
    """把 {年}{月}{yy} 等占位符替换成实际值；未知占位符原样保留。"""
    if text is None:
        return ""
    text = str(text)

    def _rep(m):
        key = m.group(1)
        return str(ctx[key]) if key in ctx else m.group(0)

    return PLACEHOLDER.sub(_rep, text)


def resolve_path(base_dir, p):
    if not p:
        return p
    p = os.path.expanduser(p)
    if os.path.isabs(p):
        return os.path.normpath(p)
    return os.path.normpath(os.path.join(base_dir, p))


# ---------------------------------------------------------------- 读源表

def read_source(path, sheet=None, header_row=1, data_start=2):
    """读一张源表。

    返回 (headers, rows)：
      headers —— 有序的列名列表
      rows    —— [{列名: 值}, ...]，全空行自动跳过
    """
    if not os.path.exists(path):
        raise FileNotFoundError("源文件不存在：%s" % path)
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        if isinstance(sheet, int):
            idx = sheet if sheet >= 0 else len(wb.worksheets) + sheet
            ws = wb.worksheets[idx]
        elif sheet in (None, "", "第一个"):
            ws = wb.worksheets[0]
        else:
            ws = wb[sheet]
        return _read_sheet(ws, header_row, data_start)
    finally:
        wb.close()


def _read_sheet(ws, header_row, data_start):
    headers = []
    col_index = {}
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is None or str(v).strip() == "":
            continue
        name = str(v).strip()
        if name in col_index:
            k = 2
            while "%s#%d" % (name, k) in col_index:
                k += 1
            name = "%s#%d" % (name, k)
        col_index[name] = c
        headers.append(name)

    rows = []
    for r in range(data_start, ws.max_row + 1):
        item = {}
        empty = True
        for name, c in col_index.items():
            v = ws.cell(r, c).value
            if v is not None and str(v).strip() != "":
                empty = False
            item[name] = v
        if not empty:
            rows.append(item)
    return headers, rows


# ---------------------------------------------------------------- 目标表

def find_col(ws, header_row, name):
    """按表头名称找目标列号（1 起），找不到返回 None。"""
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v is not None and str(v).strip() == str(name).strip():
            return c
    return None


def detect_data_area(ws, data_start):
    """数据区从 data_start 开始，连续「有边框」的最后一行。

    模板里数据行都有细边框，数据区下方那行空白没有边框，可作为边界。
    若整片都没有边框（非常规模板），退化为按已有内容最后一行。
    """
    last = data_start - 1
    for r in range(data_start, ws.max_row + 1):
        c = ws.cell(r, 1)
        b = c.border
        if b and (b.bottom.style or b.top.style):
            last = r
        else:
            break
    if last < data_start:
        for r in range(ws.max_row, data_start - 1, -1):
            if any(ws.cell(r, c).value not in (None, "") for c in range(1, ws.max_column + 1)):
                last = r
                break
    return last


def insert_rows_with_style(ws, at_row, count, style_row):
    """在 at_row 前插入 count 行，复制 style_row 的样式与行高，并同步合并/筛选。"""
    if count <= 0:
        return
    ws.insert_rows(at_row, count)

    base_height = ws.row_dimensions[style_row].height
    max_col = ws.max_column
    for r in range(at_row, at_row + count):
        if base_height:
            ws.row_dimensions[r].height = base_height
        for c in range(1, max_col + 1):
            ws.cell(r, c)._style = copy(ws.cell(style_row, c)._style)

    # 合并单元格整体下移
    for m in list(ws.merged_cells.ranges):
        if m.min_row >= at_row:
            ws.merged_cells.remove(m)
            m.shift(0, count)
            ws.merged_cells.add(m)

    # 筛选范围跟着扩
    if ws.auto_filter.ref:
        cr = CellRange(ws.auto_filter.ref)
        # 插入点在筛选范围末尾或紧邻其后时，筛选范围跟着扩
        if cr.max_row >= at_row - 1:
            cr.max_row += count
            ws.auto_filter.ref = str(cr)


# ---------------------------------------------------------------- 值转换

def transform_value(value, kind="direct", arg=None, ctx=None):
    ctx = ctx or {}
    if value is None:
        raw = ""
    elif isinstance(value, float) and value == int(value):
        raw = str(int(value))
    else:
        raw = str(value)

    kind = (kind or "direct").strip()
    if kind in ("", "direct", "直接"):
        return value if value is not None else ""
    if kind in ("strip", "去空格"):
        return raw.strip()
    if kind in ("after_slash", "取斜杠后", "取/后"):
        return raw.split("/")[-1].strip() if "/" in raw else raw.strip()
    if kind in ("before_slash", "取斜杠前", "取/前"):
        return raw.split("/")[0].strip() if "/" in raw else raw.strip()
    if kind in ("fixed", "固定值"):
        return render(arg or "", ctx)
    if kind in ("number", "数字"):
        try:
            return raw
        except Exception:
            return raw
    return raw


# ---------------------------------------------------------------- 填一张表

def fill_sheet(ws, task, rows, ctx, log=None):
    """把 rows 按列映射写入目标 sheet，返回 (写入行数, 原容量, 插入行数)。"""
    log = log or (lambda s: None)
    target = task.get("target", {})
    header_row = int(target.get("header_row", 2))
    data_start = int(target.get("data_start", header_row + 1))
    columns = [c for c in task.get("columns", []) if c.get("target")]
    if header_row < 1 or data_start <= header_row:
        raise ValueError("数据起始行必须大于表头行，且行号必须为正整数")

    if not columns:
        raise ValueError("任务「%s」没有配置列映射" % task.get("name", ""))

    mapped = []
    missing = []
    for col in columns:
        idx = find_col(ws, header_row, col["target"])
        if idx is None:
            missing.append(col["target"])
        else:
            mapped.append((idx, col))
    if missing:
        raise ValueError(
            "目标 sheet「%s」第 %d 行表头里找不到这些列：%s"
            % (ws.title, header_row, "、".join(missing))
        )

    capacity = max(0, detect_data_area(ws, data_start) - data_start + 1)
    need = len(rows)
    inserted = 0
    if need > capacity:
        inserted = need - capacity
        insert_rows_with_style(ws, data_start + capacity, inserted, data_start + capacity - 1)
        log("    数据 %d 行 > 模板容量 %d 行，已自动插入 %d 行" % (need, capacity, inserted))

    # 清空映射列的旧数据，避免源表变短后混入上次的数据。
    for r in range(data_start, data_start + max(capacity, need)):
        for idx, _ in mapped:
            ws.cell(r, idx).value = None

    for i, row in enumerate(rows):
        r = data_start + i
        for idx, col in mapped:
            v = row.get(col.get("source"))
            v = transform_value(v, col.get("transform", "direct"), col.get("arg"), ctx)
            cell = ws.cell(r, idx)
            cell.value = v
            if isinstance(v, str):
                cell.data_type = "s"  # 源表文字以文字保存，不作为公式执行。
    return need, capacity, inserted


# ---------------------------------------------------------------- 跑整套规则

def run_profile(profile, base_dir, ctx=None, log=None, overwrite=False):
    """执行一套规则，返回输出文件路径。"""
    log = log or (lambda s: None)
    ctx = ctx or date_context()

    template = resolve_path(base_dir, render(profile.get("template", ""), ctx))
    if not template or not os.path.exists(template):
        raise FileNotFoundError("模板文件不存在：%s" % template)

    out_dir = resolve_path(base_dir, render(profile.get("output_dir", ""), ctx))
    out_name = render(profile.get("output_name", ""), ctx) or os.path.basename(template)
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, out_name) if out_dir else resolve_path(base_dir, out_name)

    if os.path.normcase(os.path.realpath(template)) == os.path.normcase(os.path.realpath(out_path)):
        raise ValueError("输出文件不能与模板文件相同")
    if not profile.get("tasks"):
        raise ValueError("规则至少需要一个任务")

    # 全部任务成功后才发布结果；失败时保留原输出。
    wb = openpyxl.load_workbook(template)
    temporary = None
    try:
        for task in profile.get("tasks", []):
            name = task.get("name") or task.get("target", {}).get("sheet", "?")
            src = task.get("source", {})
            src_dir = resolve_path(base_dir, render(src.get("dir", ""), ctx)) or base_dir
            pattern = render(src.get("pattern", ""), ctx)
            if pattern and any(ch in pattern for ch in "*?["):
                matches = sorted(p for p in glob.glob(os.path.join(src_dir, pattern))
                                 if os.path.isfile(p) and not os.path.basename(p).startswith("~$"))
                if not matches:
                    raise FileNotFoundError("在 %s 里找不到匹配 %s 的文件" % (src_dir, pattern))
                src_path = matches[-1]
            else:
                src_path = resolve_path(src_dir, pattern) if pattern else ""

            if os.path.normcase(os.path.realpath(src_path)) == os.path.normcase(os.path.realpath(out_path)):
                raise ValueError("输出文件不能与源文件相同")

            log("[%s] 源文件：%s" % (name, os.path.basename(src_path)))
            headers, rows = read_source(
                src_path,
                sheet=src.get("sheet", 0),
                header_row=int(src.get("header_row", 1)),
                data_start=int(src.get("data_start", 2)),
            )
            unknown = sorted({
                c.get("source") for c in task.get("columns", [])
                if c.get("source") and c.get("source") not in headers
                and c.get("transform") not in ("fixed", "固定值")
            })
            if unknown:
                raise ValueError(
                    "源表「%s」里找不到这些列：%s\n实际表头：%s"
                    % (os.path.basename(src_path), "、".join(unknown), "、".join(headers))
                )
            if not rows:
                log("    警告：源表没有数据行")
            sheet_name = task.get("target", {}).get("sheet", "")
            if sheet_name not in wb.sheetnames:
                raise ValueError("模板里没有名为「%s」的工作表" % sheet_name)
            ws = wb[sheet_name]
            for addr, val in (task.get("cells") or {}).items():
                ws[addr] = render(val, ctx)
            need, capacity, inserted = fill_sheet(ws, task, rows, ctx, log)
            log("    写入 %d 行（模板容量 %d 行%s）" % (need, capacity, "，插 %d 行" % inserted if inserted else ""))
        fd, temporary = tempfile.mkstemp(prefix=".sheetporter-", suffix=".xlsx",
                                         dir=os.path.dirname(os.path.abspath(out_path)))
        os.close(fd)
        wb.save(temporary)
        if os.path.exists(out_path) and not overwrite:
            stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S-%f")
            root, ext = os.path.splitext(out_path)
            backup = "%s.旧文件-%s%s" % (root, stamp, ext)
            shutil.copy2(out_path, backup)
            log("目标文件已存在，旧文件已备份：%s" % os.path.basename(backup))
        os.replace(temporary, out_path)
    finally:
        wb.close()
        if temporary and os.path.exists(temporary):
            os.unlink(temporary)
    log("已保存：%s" % out_path)
    return out_path
