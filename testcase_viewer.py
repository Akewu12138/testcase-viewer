#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试用例记录表 — TestCase Viewer & Recorder
============================================
将 Excel 测试用例逐条展示为清晰的 HTML 卡片页面，
执行后记录测试结果和备注，自动回写 Excel。
支持搜索筛选 + 汇总统计页。

技术栈：Python Flask + openpyxl + 浏览器前端
适配：Mac / Windows
"""

import sys
import os
import json
import webbrowser
import threading
import datetime
import re
from pathlib import Path

# ============================================================
# 0. 依赖自检 & 自动安装
# ============================================================
def _ensure_deps():
    missing = []
    for pkg in ['flask', 'openpyxl']:
        try:
            __import__(pkg if pkg != 'flask' else 'flask')
        except ImportError:
            missing.append(pkg)

    if missing:
        import subprocess
        print(f"⏳ 首次运行，正在安装必要组件: {', '.join(missing)} ...")
        print("   （仅需一次，请稍候）\n")
        try:
            subprocess.check_call(
                [sys.executable, '-m', 'pip', 'install', '--quiet'] + missing,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            print("✅ 组件安装完成！\n")
        except Exception as e:
            print(f"❌ 自动安装失败: {e}")
            print("   请手动运行: pip install flask openpyxl")
            input("\n按回车键退出...")
            sys.exit(1)

_ensure_deps()

from flask import Flask, request, jsonify, render_template_string
import openpyxl

# ============================================================
# 1. 配置常量
# ============================================================
BASE_DIR = Path(__file__).parent.resolve()
TESTCASES_DIR = BASE_DIR / 'testcases'
PORT = 8765

# 列名智能识别规则（支持中英文常见表头）
COLUMN_PATTERNS = {
    'id':           ['编号', 'ID', '用例ID', '用例编号', '序号', 'No.', '编号ID', '用例号', 'NO'],
    'title':        ['标题', '用例名称', '测试点', '用例标题', '名称', '测试项', '测试标题', '功能点', '测试内容'],
    'precondition': ['前置条件', '预置条件', '前提条件', '准备条件', '前置', '预设', '环境准备'],
    'steps':        ['步骤', '测试步骤', '操作步骤', '执行步骤', '测试过程', '操作过程', '测试操作'],
    'expected':     ['预期结果', '期望结果', '预期', '期望', '预计结果', '预期输出', '期望输出'],
    'purpose':      ['测试目的', '目的', '测试目标', '测试说明', '目标', '测试意图'],
    'priority':     ['优先级', '等级', '重要程度', '级别', '严重程度', 'P级', '重要级别'],
    'module':       ['模块', '所属模块', '功能模块', '测试模块', '需求模块', '系统模块'],
    'category':     ['分类', '用例类型', '测试类型', '类型', '测试分类'],
}

RESULT_COL = '测试结果'
REMARK_COL = '测试现象备注'
TIME_COL   = '执行时间'


# ============================================================
# 2. Excel 操作
# ============================================================
def find_excel_file():
    """在 testcases 目录中查找 Excel 文件（排除临时文件）"""
    if not TESTCASES_DIR.exists():
        TESTCASES_DIR.mkdir(parents=True)
        return None

    excels = []
    for pattern in ['*.xlsx', '*.xls']:
        for f in TESTCASES_DIR.glob(pattern):
            if not f.name.startswith('~$') and not f.name.startswith('.'):
                excels.append(f)

    return excels[0] if excels else None


def detect_columns(headers):
    """将 Excel 表头智能映射到标准字段"""
    mapping = {}
    for i, h in enumerate(headers):
        h_str = str(h).strip() if h else ''
        for field, patterns in COLUMN_PATTERNS.items():
            if h_str == '':
                continue
            matched = False
            if h_str in patterns:
                matched = True
            else:
                for pat in patterns:
                    if pat in h_str or h_str in pat:
                        matched = True
                        break
            if matched:
                mapping[field] = i
                break
    mapping['_headers'] = headers
    return mapping


def read_testcases(filepath):
    """读取 Excel，返回（用例列表, 列映射, 表头列表）"""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))
    if len(rows) < 2:
        wb.close()
        return [], {}, []

    # 解析表头（去空格）
    raw_headers = []
    for cell in rows[0]:
        v = str(cell).strip() if cell is not None else ''
        raw_headers.append(v)
    headers = raw_headers
    mapping = detect_columns(headers)

    # 找到结果列（如果已有）
    result_col_idx = None
    remark_col_idx = None
    for i, h in enumerate(headers):
        if h == RESULT_COL:
            result_col_idx = i
        if h == REMARK_COL:
            remark_col_idx = i

    # 解析数据行
    testcases = []
    for row_idx, row in enumerate(rows[1:], start=1):
        if all(cell is None or str(cell).strip() == '' for cell in row):
            continue

        tc = {'_row': row_idx + 1}  # Excel 实际行号

        for i, value in enumerate(row):
            tc[f'col_{i}'] = str(value).strip() if value is not None else ''
            tc[f'_header_{i}'] = headers[i] if i < len(headers) else f'列{i+1}'

        for field, col_idx in mapping.items():
            if field != '_headers' and col_idx < len(row):
                raw_val = row[col_idx]
                tc[field] = str(raw_val).strip() if raw_val is not None else ''

        if result_col_idx is not None and result_col_idx < len(row):
            tc['_saved_result'] = str(row[result_col_idx]).strip() if row[result_col_idx] is not None else ''
        else:
            tc['_saved_result'] = ''

        if remark_col_idx is not None and remark_col_idx < len(row):
            tc['_saved_remark'] = str(row[remark_col_idx]).strip() if row[remark_col_idx] is not None else ''
        else:
            tc['_saved_remark'] = ''

        # 构建搜索文本（所有字段拼接，用于全文搜索）
        search_parts = []
        for field in ['id', 'title', 'module', 'priority', 'category',
                       'precondition', 'steps', 'expected', 'purpose']:
            if field in tc:
                search_parts.append(tc[field])
        for i in range(len(row)):
            if tc.get(f'col_{i}', ''):
                search_parts.append(tc[f'col_{i}'])
        tc['_search_text'] = ' '.join(search_parts)

        testcases.append(tc)

    wb.close()
    return testcases, mapping, headers


def save_result(filepath, row_number, result, remark):
    """将测试结果写入 Excel 指定行"""
    wb = openpyxl.load_workbook(filepath)
    ws = wb.active

    header_cells = [ws.cell(row=1, column=c + 1) for c in range(ws.max_column)]
    headers_row = [str(c.value).strip() if c.value is not None else '' for c in header_cells]
    while len(headers_row) < ws.max_column:
        headers_row.append('')

    def find_or_create_col(col_name):
        for i, h in enumerate(headers_row):
            if h == col_name:
                return i + 1
        new_col = len(headers_row) + 1
        ws.cell(row=1, column=new_col, value=col_name)
        headers_row.append(col_name)
        return new_col

    result_col = find_or_create_col(RESULT_COL)
    remark_col = find_or_create_col(REMARK_COL)
    time_col   = find_or_create_col(TIME_COL)

    now_str = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    ws.cell(row=row_number, column=result_col, value=result)
    ws.cell(row=row_number, column=remark_col, value=remark)
    ws.cell(row=row_number, column=time_col,   value=now_str)

    wb.save(filepath)
    wb.close()


def build_display_fields(tc, mapping, headers):
    """构建前端展示用的字段列表（有序）"""
    fields = []

    ordered_specials = [
        ('id',           '\U0001f4cb 用例编号', 'field-id'),
        ('title',        '\U0001f4cc 用例标题', 'field-title'),
        ('module',       '\U0001f4c1 所属模块', 'field-meta'),
        ('priority',     '⚡ 优先级',   'field-meta'),
        ('category',     '\U0001f3f7️ 分类',     'field-meta'),
        ('precondition', '\U0001f527 前置条件', 'field-section'),
        ('steps',        '\U0001f4dd 测试步骤', 'field-section field-steps'),
        ('expected',     '✅ 预期结果', 'field-section field-expected'),
        ('purpose',      '\U0001f3af 测试目的', 'field-purpose'),
    ]

    displayed_cols = set()

    for field, label, css_class in ordered_specials:
        if field in mapping:
            col_idx = mapping[field]
            val = tc.get(field, '')
            if val:
                fields.append({
                    'label': label,
                    'value': val,
                    'css_class': css_class,
                })
                displayed_cols.add(col_idx)

    for i, h in enumerate(headers):
        if i in displayed_cols:
            continue
        if i >= len(headers):
            continue
        val = tc.get(f'col_{i}', '')
        if val and val != 'None' and h not in (RESULT_COL, REMARK_COL, TIME_COL):
            fields.append({
                'label': f'\U0001f4c4 {h}',
                'value': val,
                'css_class': 'field-section',
            })

    return fields


# ============================================================
# 3. Flask 应用 & API
# ============================================================
app = Flask(__name__)

STATE = {
    'testcases': [],
    'mapping': {},
    'headers': [],
    'filepath': None,
    'filename': '',
}


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/api/init')
def api_init():
    return jsonify({
        'loaded': len(STATE['testcases']) > 0,
        'filename': STATE['filename'],
        'total': len(STATE['testcases']),
    })


@app.route('/api/testcase/<int:index>')
def api_testcase(index):
    if index < 0 or index >= len(STATE['testcases']):
        return jsonify({'error': '索引超出范围'}), 404

    tc = STATE['testcases'][index].copy()
    tc['_index'] = index
    tc['_display_fields'] = build_display_fields(tc, STATE['mapping'], STATE['headers'])
    tc['_total'] = len(STATE['testcases'])

    return jsonify(tc)


@app.route('/api/all-status')
def api_all_status():
    """获取所有用例的执行状态"""
    if not STATE['filepath'] or not STATE['testcases']:
        return jsonify([])

    try:
        wb = openpyxl.load_workbook(STATE['filepath'], data_only=True)
        ws = wb.active

        headers_row = [
            str(ws.cell(row=1, column=c + 1).value).strip()
            if ws.cell(row=1, column=c + 1).value is not None else ''
            for c in range(ws.max_column)
        ]

        result_col = None
        for i, h in enumerate(headers_row):
            if h == RESULT_COL:
                result_col = i + 1
                break

        statuses = []
        for tc in STATE['testcases']:
            row = tc['_row']
            if result_col:
                cell_val = ws.cell(row=row, column=result_col).value
                statuses.append(str(cell_val).strip() if cell_val else '')
            else:
                statuses.append('')

        wb.close()
        return jsonify(statuses)
    except Exception:
        return jsonify([''] * len(STATE['testcases']))


@app.route('/api/search', methods=['POST'])
def api_search():
    """搜索筛选 API

    请求体示例:
    {
        "keyword": "登录",           // 全文搜索关键词（可选）
        "result_filter": "",         // 按结果筛选: ""全部 / "通过" / "失败" / "阻塞" / "跳过" / "未执行"
        "priority_filter": "",       // 按优先级筛选（可选）
        "module_filter": "",         // 按模块筛选（可选）
        "page": 0,                   // 分页页码（0-based）
        "page_size": 50              // 每页条数
    }

    返回:
    {
        "results": [ { index, id, title, module, priority, result, remark, _search_snippet } ],
        "total_matched": 150,
        "page": 0,
        "page_size": 50
    }
    """
    data = request.get_json() or {}
    keyword = (data.get('keyword') or '').strip().lower()
    result_filter = (data.get('result_filter') or '').strip()
    priority_filter = (data.get('priority_filter') or '').strip()
    module_filter = (data.get('module_filter') or '').strip()
    page = max(0, int(data.get('page', 0)))
    page_size = min(200, max(1, int(data.get('page_size', 50))))

    # 收集所有用例的状态（从内存中）
    matched = []
    for i, tc in enumerate(STATE['testcases']):
        # 全文搜索
        if keyword:
            if keyword not in tc.get('_search_text', '').lower():
                continue

        # 结果筛选
        if result_filter:
            saved = tc.get('_saved_result', '')
            if result_filter == '未执行':
                if saved != '' and saved != 'None':
                    continue
            elif saved != result_filter:
                continue

        # 优先级筛选
        if priority_filter:
            tc_priority = tc.get('priority', '')
            if priority_filter not in tc_priority:
                continue

        # 模块筛选
        if module_filter:
            tc_module = tc.get('module', '')
            if module_filter not in tc_module:
                continue

        # 构建搜索摘要（关键词高亮上下文）
        snippet = ''
        if keyword:
            search_text = tc.get('_search_text', '')
            idx = search_text.lower().find(keyword)
            if idx >= 0:
                start = max(0, idx - 30)
                end = min(len(search_text), idx + len(keyword) + 50)
                snippet = search_text[start:end]
                if start > 0:
                    snippet = '...' + snippet
                if end < len(search_text):
                    snippet = snippet + '...'

        matched.append({
            'index': i,
            'id': tc.get('id', ''),
            'title': tc.get('title', ''),
            'module': tc.get('module', ''),
            'priority': tc.get('priority', ''),
            'result': tc.get('_saved_result', ''),
            'remark': tc.get('_saved_remark', ''),
            'snippet': snippet,
        })

    total_matched = len(matched)
    start = page * page_size
    end = start + page_size
    page_results = matched[start:end]

    return jsonify({
        'results': page_results,
        'total_matched': total_matched,
        'page': page,
        'page_size': page_size,
        'total_pages': max(1, (total_matched + page_size - 1) // page_size),
    })


@app.route('/api/summary')
def api_summary():
    """汇总统计 API

    返回:
    {
        "total": 200,
        "counts": { "通过": 120, "失败": 15, "阻塞": 5, "跳过": 10, "未执行": 50 },
        "by_priority": { "P0": { "通过":5, "失败":2, "阻塞":0, "跳过":0, "未执行":3 }, ... },
        "by_module": { "登录模块": { "通过":10, "失败":1, ... }, ... },
        "execution_rate": 75.0,       // 已执行百分比
        "pass_rate": 80.0,            // 通过率（通过的/已执行）
    }
    """
    tcs = STATE['testcases']

    counts = {'通过': 0, '失败': 0, '阻塞': 0, '跳过': 0, '未执行': 0}
    by_priority = {}
    by_module = {}

    for tc in tcs:
        result = tc.get('_saved_result', '')
        if not result or result == 'None':
            result = '未执行'

        # 全局计数
        if result in counts:
            counts[result] += 1
        else:
            counts['未执行'] += 1

        # 按优先级分组
        pri = tc.get('priority', '未标注')
        if not pri or pri == 'None':
            pri = '未标注'
        if pri not in by_priority:
            by_priority[pri] = {'通过': 0, '失败': 0, '阻塞': 0, '跳过': 0, '未执行': 0}
        by_priority[pri][result] += 1

        # 按模块分组
        mod = tc.get('module', '未分类')
        if not mod or mod == 'None':
            mod = '未分类'
        if mod not in by_module:
            by_module[mod] = {'通过': 0, '失败': 0, '阻塞': 0, '跳过': 0, '未执行': 0}
        by_module[mod][result] += 1

    total = len(tcs)
    executed = total - counts['未执行']
    execution_rate = round((executed / total) * 100, 1) if total > 0 else 0
    pass_rate = round((counts['通过'] / executed) * 100, 1) if executed > 0 else 0

    return jsonify({
        'total': total,
        'counts': counts,
        'by_priority': by_priority,
        'by_module': by_module,
        'execution_rate': execution_rate,
        'pass_rate': pass_rate,
    })


@app.route('/api/filter-options')
def api_filter_options():
    """获取所有可用的筛选选项（优先级列表、模块列表）"""
    priorities = set()
    modules = set()

    for tc in STATE['testcases']:
        p = tc.get('priority', '')
        if p and p != 'None':
            priorities.add(p)
        m = tc.get('module', '')
        if m and m != 'None':
            modules.add(m)

    return jsonify({
        'priorities': sorted(priorities),
        'modules': sorted(modules),
    })


@app.route('/api/save', methods=['POST'])
def api_save():
    """保存测试结果"""
    data = request.get_json()
    index = data.get('index')
    result = data.get('result', '')
    remark = data.get('remark', '')

    if index is None or index < 0 or index >= len(STATE['testcases']):
        return jsonify({'success': False, 'error': '索引无效'}), 400
    if not STATE['filepath']:
        return jsonify({'success': False, 'error': '未找到 Excel 文件'}), 400

    try:
        row_number = STATE['testcases'][index]['_row']
        save_result(STATE['filepath'], row_number, result, remark)

        STATE['testcases'][index]['_saved_result'] = result
        STATE['testcases'][index]['_saved_remark'] = remark

        return jsonify({'success': True})
    except PermissionError:
        return jsonify({
            'success': False,
            'error': '无法写入 Excel 文件。请关闭 Excel 程序后重试。'
        }), 500
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


# ============================================================
# 4. HTML 模板（带搜索筛选 + 汇总页的完整前端）
# ============================================================
HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>测试用例记录表</title>
<style>
/* ========== CSS 变量 & 全局 ========== */
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}

:root{
    --bg:#f0f2f5;
    --card:#fff;
    --text:#1f2937;
    --text-secondary:#6b7280;
    --border:#e5e7eb;
    --primary:#4f46e5;
    --primary-light:#eef2ff;
    --pass:#16a34a;
    --pass-light:#dcfce7;
    --pass-bg:#f0fdf4;
    --fail:#dc2626;
    --fail-light:#fecaca;
    --fail-bg:#fef2f2;
    --block:#ea580c;
    --block-light:#fed7aa;
    --block-bg:#fff7ed;
    --skip:#6b7280;
    --skip-light:#e5e7eb;
    --skip-bg:#f9fafb;
    --purpose-bg:#fffbeb;
    --purpose-border:#fcd34d;
    --shadow:0 1px 3px rgba(0,0,0,0.08),0 1px 2px rgba(0,0,0,0.06);
    --shadow-lg:0 4px 6px rgba(0,0,0,0.07),0 2px 4px rgba(0,0,0,0.06);
    --radius:12px;
    --radius-sm:8px;
}

body{
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC",
               "Microsoft YaHei","Helvetica Neue",sans-serif;
    background:var(--bg);
    color:var(--text);
    line-height:1.6;
    min-height:100vh;
}

/* ========== 顶部导航栏 ========== */
.topbar{
    background:var(--card);
    border-bottom:1px solid var(--border);
    padding:0 24px;
    height:56px;
    display:flex;
    align-items:center;
    justify-content:space-between;
    position:sticky;
    top:0;
    z-index:100;
    box-shadow:var(--shadow);
    gap:16px;
}
.topbar-left{display:flex;align-items:center;gap:10px;white-space:nowrap;}
.topbar-logo{font-size:20px;}
.topbar-title{font-size:16px;font-weight:700;color:var(--text);letter-spacing:-.3px;}

/* 视图切换标签 */
.view-tabs{display:flex;gap:0;}
.view-tab{
    padding:6px 16px;
    border:2px solid var(--border);
    background:var(--card);
    font-size:13px;
    font-weight:600;
    cursor:pointer;
    transition:all .15s;
    font-family:inherit;
    color:var(--text-secondary);
}
.view-tab:first-child{border-radius:var(--radius-sm) 0 0 var(--radius-sm);}
.view-tab:last-child{border-radius:0 var(--radius-sm) var(--radius-sm) 0;}
.view-tab.active{
    background:var(--primary);
    color:#fff;
    border-color:var(--primary);
}
.view-tab:hover:not(.active){background:var(--primary-light);color:var(--primary);}

.topbar-right{display:flex;align-items:center;gap:12px;font-size:13px;color:var(--text-secondary);white-space:nowrap;}

/* ========== 搜索栏 ========== */
.search-bar-wrap{
    max-width:960px;
    margin:0 auto;
    padding:12px 24px 0;
    display:flex;
    gap:8px;
    align-items:center;
    flex-wrap:wrap;
}
.search-input{
    flex:1;
    min-width:180px;
    padding:8px 14px;
    border:2px solid var(--border);
    border-radius:var(--radius-sm);
    font-size:14px;
    font-family:inherit;
    background:var(--card);
    color:var(--text);
    transition:border-color .15s;
}
.search-input:focus{outline:none;border-color:var(--primary);box-shadow:0 0 0 3px rgba(79,70,229,.1);}
.search-input::placeholder{color:#c4c4c4;}

.filter-select{
    padding:8px 12px;
    border:2px solid var(--border);
    border-radius:var(--radius-sm);
    font-size:13px;
    font-family:inherit;
    background:var(--card);
    color:var(--text);
    cursor:pointer;
    min-width:90px;
}
.filter-select:focus{outline:none;border-color:var(--primary);}

.search-clear{
    background:none;
    border:none;
    font-size:13px;
    color:var(--primary);
    cursor:pointer;
    white-space:nowrap;
    padding:4px 8px;
    font-family:inherit;
}
.search-clear:hover{text-decoration:underline;}

.search-info{
    max-width:960px;
    margin:0 auto;
    padding:4px 24px 0;
    font-size:12px;
    color:var(--text-secondary);
}

/* ========== 进度条 ========== */
.progress-wrap{
    padding:12px 24px 0;
    max-width:960px;
    margin:0 auto;
}
.progress-info{display:flex;justify-content:space-between;margin-bottom:4px;font-size:13px;color:var(--text-secondary);}
.progress-bar{height:6px;background:var(--border);border-radius:3px;overflow:hidden;}
.progress-fill{height:100%;background:var(--primary);border-radius:3px;transition:width .4s ease;}
.progress-fill.done{background:var(--pass);}
.legend{display:flex;gap:14px;margin-top:6px;font-size:12px;color:var(--text-secondary);flex-wrap:wrap;}
.legend-dot{display:inline-block;width:10px;height:10px;border-radius:50%;margin-right:3px;vertical-align:-1px;}
.legend-dot.pass{background:var(--pass)}
.legend-dot.fail{background:var(--fail)}
.legend-dot.block{background:var(--block)}
.legend-dot.skip{background:var(--skip)}

/* ========== 主内容区 ========== */
.main-container{max-width:960px;margin:0 auto;padding:16px 24px 40px;}

/* 空状态 */
.state-message{text-align:center;padding:80px 20px;color:var(--text-secondary);}
.state-message .icon{font-size:56px;margin-bottom:16px;}
.state-message h2{font-size:20px;margin-bottom:8px;color:var(--text);}
.state-message p{font-size:14px;max-width:400px;margin:0 auto;line-height:1.7;}
.state-message .path-hint{
    display:inline-block;background:var(--card);border:1px dashed var(--border);
    border-radius:var(--radius-sm);padding:6px 14px;margin-top:12px;
    font-family:"SF Mono","Fira Code",monospace;font-size:13px;
}

/* ========== 用例卡片 ========== */
.card{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden;margin-bottom:16px;}
.card-header{
    padding:16px 24px 14px;border-bottom:1px solid var(--border);
    display:flex;align-items:flex-start;justify-content:space-between;flex-wrap:wrap;gap:8px;
}
.card-navigator{font-size:14px;color:var(--text-secondary);white-space:nowrap;}
.card-navigator strong{color:var(--primary);font-size:18px;font-weight:700;}
.card-body{padding:6px 24px 18px;}

/* 结果状态徽标 */
.result-badge{
    display:inline-block;padding:3px 12px;border-radius:20px;font-size:12px;font-weight:600;
}
.result-badge-pass{background:var(--pass-light);color:var(--pass)}
.result-badge-fail{background:var(--fail-light);color:var(--fail)}
.result-badge-block{background:var(--block-light);color:var(--block)}
.result-badge-skip{background:var(--skip-light);color:var(--skip)}

/* 字段展示 */
.field-row{padding:10px 0;border-bottom:1px solid #f3f4f6;}
.field-row:last-child{border-bottom:none;}
.field-label{font-size:13px;font-weight:600;color:var(--text-secondary);margin-bottom:3px;letter-spacing:.2px;}
.field-value{font-size:15px;color:var(--text);line-height:1.7;white-space:pre-wrap;word-break:break-word;}

.field-id .field-value{
    font-family:"SF Mono","Fira Code","Consolas",monospace;font-size:14px;
    background:var(--primary-light);display:inline-block;padding:2px 10px;border-radius:4px;color:var(--primary);
}
.field-title .field-value{font-size:20px;font-weight:700;line-height:1.3;}
.field-meta{display:inline-flex;align-items:center;gap:16px;padding:5px 0;border-bottom:none;}
.field-meta .field-label{margin-bottom:0;}
.field-meta .field-value{font-size:13px;background:#f3f4f6;padding:2px 12px;border-radius:20px;}
.field-expected .field-value{
    color:#15803d;background:#f0fdf4;padding:12px 16px;border-radius:var(--radius-sm);border-left:4px solid #22c55e;
}
.field-purpose{
    background:var(--purpose-bg);border:1px solid var(--purpose-border);
    border-radius:var(--radius-sm);padding:14px 18px;margin:6px 0 0;
}
.field-purpose .field-label{color:#92400e;font-size:14px;}
.field-purpose .field-value{color:#78350f;font-size:16px;font-weight:500;margin-top:4px;}

/* ========== 操作区 ========== */
.action-card{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);padding:20px 24px;margin-bottom:16px;}
.action-card h3{font-size:15px;margin-bottom:14px;color:var(--text);}

.result-group{display:flex;gap:10px;margin-bottom:18px;flex-wrap:wrap;}
.result-btn{
    flex:1;min-width:90px;padding:12px 6px;border:2px solid var(--border);border-radius:var(--radius-sm);
    background:var(--card);font-size:14px;font-weight:600;cursor:pointer;
    transition:all .15s ease;text-align:center;color:var(--text);user-select:none;
}
.result-btn:hover{transform:translateY(-1px);box-shadow:var(--shadow-lg);}
.result-btn:active{transform:scale(.97);}

.result-btn.selected-pass{border-color:var(--pass);background:var(--pass-bg);color:var(--pass);box-shadow:0 0 0 3px rgba(22,163,74,.15);}
.result-btn.selected-fail{border-color:var(--fail);background:var(--fail-bg);color:var(--fail);box-shadow:0 0 0 3px rgba(220,38,38,.15);}
.result-btn.selected-block{border-color:var(--block);background:var(--block-bg);color:var(--block);box-shadow:0 0 0 3px rgba(234,88,12,.15);}
.result-btn.selected-skip{border-color:var(--skip);background:var(--skip-bg);color:var(--skip);box-shadow:0 0 0 3px rgba(107,114,128,.15);}

.result-emoji{font-size:18px;display:block;margin-bottom:3px;}

.remark-area{
    width:100%;min-height:76px;padding:12px 14px;border:2px solid var(--border);
    border-radius:var(--radius-sm);font-size:14px;font-family:inherit;line-height:1.6;
    resize:vertical;transition:border-color .15s;color:var(--text);background:#fafafa;
}
.remark-area:focus{outline:none;border-color:var(--primary);background:#fff;box-shadow:0 0 0 3px rgba(79,70,229,.1);}
.remark-area::placeholder{color:#c3c3c3;}

.btn{
    display:inline-flex;align-items:center;gap:6px;padding:10px 22px;
    border:none;border-radius:var(--radius-sm);font-size:14px;font-weight:600;
    cursor:pointer;transition:all .15s ease;font-family:inherit;user-select:none;
}
.btn:active{transform:scale(.97);}

.btn-save{background:var(--primary);color:#fff;padding:12px 28px;font-size:15px;box-shadow:0 2px 4px rgba(79,70,229,.3);}
.btn-save:hover{background:#4338ca;box-shadow:0 4px 8px rgba(79,70,229,.35);}
.btn-save:disabled{background:#a5b4fc;cursor:not-allowed;box-shadow:none;}

.btn-nav{background:var(--card);border:2px solid var(--border);color:var(--text);padding:10px 20px;font-size:14px;}
.btn-nav:hover{border-color:var(--primary);color:var(--primary);background:var(--primary-light);}
.btn-nav:disabled{opacity:.4;cursor:not-allowed;}
.btn-nav:disabled:hover{border-color:var(--border);color:var(--text);background:var(--card);}

.btn-sm{padding:6px 14px;font-size:12px;}
.btn-outline{background:var(--card);border:2px solid var(--border);color:var(--text);}
.btn-outline:hover{border-color:var(--primary);color:var(--primary);background:var(--primary-light);}

/* ========== 搜索结果列表 ========== */
.search-result-list{display:flex;flex-direction:column;gap:8px;}
.search-result-item{
    background:var(--card);border:2px solid var(--border);border-radius:var(--radius-sm);
    padding:12px 16px;cursor:pointer;transition:all .15s;
    display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;
}
.search-result-item:hover{border-color:var(--primary);box-shadow:var(--shadow);}
.search-result-item .sr-left{flex:1;min-width:0;}
.search-result-item .sr-id{font-family:"SF Mono","Fira Code",monospace;font-size:12px;color:var(--primary);}
.search-result-item .sr-title{font-size:15px;font-weight:600;margin:2px 0;word-break:break-word;}
.search-result-item .sr-meta{font-size:12px;color:var(--text-secondary);}
.search-result-item .sr-snippet{font-size:12px;color:var(--text-secondary);margin-top:4px;
    white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}
.search-result-item .sr-snippet em{background:#fde68a;font-style:normal;padding:0 2px;border-radius:2px;}
.search-result-item .sr-right{flex-shrink:0;display:flex;align-items:center;gap:8px;}

/* ========== 底部分页 ========== */
.pagination{display:flex;align-items:center;justify-content:center;gap:8px;margin-top:16px;font-size:14px;}
.pagination .page-info{color:var(--text-secondary);}

/* ========== 汇总页 ========== */
.summary-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:20px;}
.summary-stat{
    background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);
    padding:20px;text-align:center;
}
.summary-stat .stat-number{font-size:36px;font-weight:800;line-height:1.1;}
.summary-stat .stat-label{font-size:13px;color:var(--text-secondary);margin-top:4px;}
.stat-pass .stat-number{color:var(--pass)}
.stat-fail .stat-number{color:var(--fail)}
.stat-block .stat-number{color:var(--block)}
.stat-skip .stat-number{color:var(--skip)}
.stat-total .stat-number{color:var(--primary)}

/* 汇总表格 */
.summary-table-wrap{background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);overflow:hidden;margin-bottom:16px;}
.summary-table-wrap h3{padding:16px 20px 12px;font-size:15px;border-bottom:1px solid var(--border);}
.summary-table{width:100%;border-collapse:collapse;font-size:13px;}
.summary-table th,.summary-table td{padding:10px 14px;text-align:center;border-bottom:1px solid #f3f4f6;}
.summary-table th{background:#f9fafb;font-weight:600;font-size:12px;color:var(--text-secondary);}
.summary-table td:first-child{text-align:left;font-weight:500;}
.summary-table .cell-pass{color:var(--pass);font-weight:600;}
.summary-table .cell-fail{color:var(--fail);font-weight:600;}
.summary-table .cell-block{color:var(--block);font-weight:600;}
.summary-table .cell-skip{color:var(--skip);}
.summary-table .bar-cell{width:120px;}
.summary-table .mini-bar{height:6px;background:var(--border);border-radius:3px;overflow:hidden;display:inline-block;width:100%;}
.summary-table .mini-bar-fill{height:100%;border-radius:3px;transition:width .4s;}
.mini-bar-fill.pass{background:var(--pass)}
.mini-bar-fill.fail{background:var(--fail)}
.mini-bar-fill.block{background:var(--block)}

/* 汇总指标行 */
.summary-metrics{display:flex;gap:20px;flex-wrap:wrap;margin-bottom:20px;}
.metric-card{
    background:var(--card);border-radius:var(--radius);box-shadow:var(--shadow);
    padding:16px 20px;display:flex;align-items:center;gap:14px;flex:1;min-width:200px;
}
.metric-ring{width:64px;height:64px;position:relative;flex-shrink:0;}
.metric-ring svg{transform:rotate(-90deg);}
.metric-ring .bg{fill:none;stroke:var(--border);stroke-width:6;}
.metric-ring .fg{fill:none;stroke-width:6;stroke-linecap:round;transition:stroke-dashoffset .6s ease;}
.metric-ring .pct{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);font-size:16px;font-weight:800;}
.metric-text h4{font-size:14px;margin-bottom:2px;color:var(--text);}
.metric-text p{font-size:12px;color:var(--text-secondary);}

/* ========== 底部导航 ========== */
.nav-bar{display:flex;align-items:center;justify-content:center;gap:16px;padding:8px 0;}
.nav-hint{font-size:12px;color:#c4c4c4;text-align:center;margin-top:6px;}

/* ========== Toast ========== */
.toast{
    position:fixed;top:20px;left:50%;transform:translateX(-50%);padding:12px 24px;
    border-radius:var(--radius-sm);font-size:14px;font-weight:600;
    z-index:999;animation:toastIn .3s ease,toastOut .3s ease 2s forwards;box-shadow:var(--shadow-lg);
}
.toast-success{background:#166534;color:#fff;}
.toast-error{background:#991b1b;color:#fff;}
.toast-warn{background:#92400e;color:#fff;}

@keyframes toastIn{from{opacity:0;transform:translateX(-50%) translateY(-20px)}to{opacity:1;transform:translateX(-50%) translateY(0)}}
@keyframes toastOut{from{opacity:1}to{opacity:0}}

/* ========== 弹窗 ========== */
.modal-overlay{
    position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.4);
    display:flex;align-items:center;justify-content:center;z-index:200;
}
.modal-box{
    background:var(--card);border-radius:var(--radius);padding:28px;max-width:400px;
    width:90%;box-shadow:0 20px 60px rgba(0,0,0,.15);text-align:center;
}
.modal-box h3{margin-bottom:10px;font-size:18px;}
.modal-box p{margin-bottom:20px;color:var(--text-secondary);font-size:14px;line-height:1.6;}
.modal-actions{display:flex;gap:10px;justify-content:center;}
.btn-modal-primary{background:var(--primary);color:#fff;padding:8px 20px;border:none;border-radius:var(--radius-sm);font-size:14px;font-weight:600;cursor:pointer;font-family:inherit;}
.btn-modal-ghost{background:var(--card);color:var(--text-secondary);padding:8px 20px;border:2px solid var(--border);border-radius:var(--radius-sm);font-size:14px;cursor:pointer;font-family:inherit;}

/* ========== 快捷键提示 ========== */
.kbd{display:inline-block;background:#f3f4f6;border:1px solid #d1d5db;border-radius:4px;padding:1px 6px;font-family:"SF Mono","Fira Code",monospace;font-size:11px;vertical-align:1px;}

.highlight{background:#fde68a;padding:0 2px;border-radius:2px;}

/* ========== 响应式 ========== */
@media(max-width:640px){
    .topbar{padding:0 12px;flex-wrap:wrap;height:auto;padding-top:8px;padding-bottom:8px;}
    .card-header{padding:12px 14px 10px}
    .card-body{padding:4px 14px 14px}
    .action-card{padding:16px}
    .result-group{gap:6px}
    .result-btn{min-width:70px;padding:10px 4px;font-size:13px}
    .result-emoji{font-size:16px}
    .field-title .field-value{font-size:17px}
    .main-container{padding:10px 12px 40px}
    .progress-wrap{padding:10px 12px 0}
    .search-bar-wrap{padding:10px 12px 0}
    .search-info{padding:2px 12px 0}
    .summary-grid{grid-template-columns:repeat(2,1fr)}
    .summary-metrics{flex-direction:column;}
}

@media print{
    body{background:#fff}
    .topbar,.action-card,.nav-bar,.nav-hint,.progress-wrap,.search-bar-wrap,.search-info{display:none}
    .card{box-shadow:none;border:1px solid #e5e7eb}
}
</style>
</head>
<body>

<!-- 顶部导航 -->
<div class="topbar" id="topbarMain">
    <div class="topbar-left">
        <span class="topbar-logo">&#128300;</span>
        <span class="topbar-title">测试用例记录表</span>
    </div>
    <div class="view-tabs">
        <button class="view-tab active" id="tabExecute" onclick="switchView('execute')">&#9654; 执行测试</button>
        <button class="view-tab" id="tabSearch" onclick="switchView('search')">&#128269; 搜索筛选</button>
        <button class="view-tab" id="tabSummary" onclick="switchView('summary')">&#128202; 汇总统计</button>
    </div>
    <div class="topbar-right">
        <span id="topbarFilename">未加载</span>
    </div>
</div>

<!-- ===== 执行测试视图 ===== -->
<div id="viewExecute">
    <!-- 进度条 -->
    <div class="progress-wrap" id="progressWrap" style="display:none">
        <div class="progress-info">
            <span id="progressText">共 0 条</span>
            <span id="progressPercent">0%</span>
        </div>
        <div class="progress-bar"><div class="progress-fill" id="progressFill" style="width:0%"></div></div>
        <div class="legend">
            <span><span class="legend-dot pass"></span>通过 <span id="countPass">0</span></span>
            <span><span class="legend-dot fail"></span>失败 <span id="countFail">0</span></span>
            <span><span class="legend-dot block"></span>阻塞 <span id="countBlock">0</span></span>
            <span><span class="legend-dot skip"></span>跳过 <span id="countSkip">0</span></span>
            <span>未执行 <span id="countNone">0</span></span>
        </div>
    </div>

    <div class="main-container" id="execContainer"></div>

    <div class="nav-bar" id="navBar" style="display:none">
        <button class="btn btn-nav" id="btnPrev" onclick="goTo(-1)" title="&#8592; 键">&#9664; 上一条</button>
        <button class="btn btn-outline btn-sm" onclick="switchView('search')" style="margin:0 8px;">&#128269; 搜索</button>
        <button class="btn btn-nav" id="btnNext" onclick="goTo(1)" title="&#8594; 键">下一条 &#9654;</button>
    </div>
    <div class="nav-hint" id="navHint" style="display:none">
        <span class="kbd">&#8592;</span> 上一条 &nbsp; <span class="kbd">&#8594;</span> 下一条 &nbsp; <span class="kbd">Ctrl+S</span> 保存
    </div>
</div>

<!-- ===== 搜索筛选视图 ===== -->
<div id="viewSearch" style="display:none;">
    <div class="search-bar-wrap" style="padding-top:16px;">
        <input type="text" class="search-input" id="searchKeyword"
               placeholder="输入关键词搜索（用例编号、标题、步骤等）..."
               oninput="doSearch()" onkeydown="if(event.key==='Enter')doSearch()">
        <select class="filter-select" id="filterResult" onchange="doSearch()">
            <option value="">全部结果</option>
            <option value="未执行">未执行</option>
            <option value="通过">通过</option>
            <option value="失败">失败</option>
            <option value="阻塞">阻塞</option>
            <option value="跳过">跳过</option>
        </select>
        <select class="filter-select" id="filterPriority" onchange="doSearch()">
            <option value="">全部优先级</option>
        </select>
        <select class="filter-select" id="filterModule" onchange="doSearch()">
            <option value="">全部模块</option>
        </select>
        <button class="search-clear" onclick="clearSearch()">&#10005; 清除</button>
    </div>
    <div class="search-info" id="searchInfo"></div>

    <div class="main-container" id="searchContainer">
        <div class="state-message"><div class="icon">&#128269;</div><h2>输入关键词开始搜索</h2></div>
    </div>

    <div class="pagination" id="searchPagination" style="display:none;"></div>
</div>

<!-- ===== 汇总统计视图 ===== -->
<div id="viewSummary" style="display:none;">
    <div class="main-container" id="summaryContainer">
        <div class="state-message"><div class="icon">&#9203;</div><h2>加载中...</h2></div>
    </div>
</div>

<!-- Toast -->
<div id="toastContainer"></div>

<script>
// ============================================================
// 全局状态
// ============================================================
let currentView = 'execute';    // execute | search | summary
let currentIndex = 0;
let totalCount = 0;
let dirty = false;
let selectedResult = '';
let allStatuses = [];

let searchPage = 0;
let searchPageSize = 30;
let searchLastParams = null;

// ============================================================
// 视图切换
// ============================================================
function switchView(view) {
    if (dirty && view !== 'execute') {
        showUnsavedDialog().then(confirmed => {
            if (confirmed) _doSwitch(view);
        });
        return;
    }
    _doSwitch(view);
}

function _doSwitch(view) {
    currentView = view;

    document.getElementById('viewExecute').style.display = view === 'execute' ? 'block' : 'none';
    document.getElementById('viewSearch').style.display = view === 'search' ? 'block' : 'none';
    document.getElementById('viewSummary').style.display = view === 'summary' ? 'block' : 'none';

    document.getElementById('tabExecute').classList.toggle('active', view === 'execute');
    document.getElementById('tabSearch').classList.toggle('active', view === 'search');
    document.getElementById('tabSummary').classList.toggle('active', view === 'summary');

    if (view === 'search') {
        loadFilterOptions();
        searchPage = 0;
        doSearch();
    } else if (view === 'summary') {
        loadSummary();
    }

    window.scrollTo({top: 0, behavior: 'smooth'});
}

// ============================================================
// 初始化
// ============================================================
async function init() {
    try {
        const resp = await fetch('/api/init');
        const data = await resp.json();

        if (!data.loaded) {
            showEmptyState();
            return;
        }

        totalCount = data.total;
        document.getElementById('topbarFilename').textContent =
            '\U0001f4c2 ' + (data.filename || '已加载');
        document.getElementById('progressWrap').style.display = 'block';
        document.getElementById('navBar').style.display = 'flex';
        document.getElementById('navHint').style.display = 'block';

        await refreshStatuses();
        await loadTestCase(0);
    } catch (err) {
        showEmptyState();
        console.error('初始化失败:', err);
    }
}

function showEmptyState() {
    const container = document.getElementById('execContainer');
    container.innerHTML = `
        <div class="state-message">
            <div class="icon">\U0001f4cb</div>
            <h2>未找到测试用例文件或未读取到用例</h2>
            <p>请将测试用例 Excel 文件（.xlsx 或 .xls）<br>放入工具目录下的 <strong>testcases</strong> 文件夹中，<br>然后重新启动程序。<br><br>第一行应为表头，从第二行开始为用例数据。</p>
            <div class="path-hint">将 Excel 文件放入<br><code>testcases/</code> 文件夹</div>
        </div>
    `;
}

// ============================================================
// 执行测试 - 加载用例
// ============================================================
async function loadTestCase(index) {
    if (dirty) {
        const confirmed = await showUnsavedDialog();
        if (!confirmed) return;
    }

    try {
        const resp = await fetch('/api/testcase/' + index);
        if (!resp.ok) throw new Error('加载失败');

        const tc = await resp.json();
        currentIndex = index;
        renderTestCase(tc);
        updateNavButtons();
        selectedResult = tc._saved_result || '';
        dirty = false;
        window.scrollTo({top: 0, behavior: 'smooth'});
    } catch (err) {
        showToast('加载用例失败: ' + err.message, 'error');
    }
}

function renderTestCase(tc) {
    const container = document.getElementById('execContainer');
    const fields = tc._display_fields || [];
    const total = tc._total || totalCount;

    let fieldsHtml = '';
    for (const f of fields) {
        let valueHtml = escapeHtml(f.value);
        if (f.css_class.includes('field-steps')) {
            valueHtml = formatSteps(f.value);
        }
        fieldsHtml += `
            <div class="field-row ${f.css_class}">
                <div class="field-label">${f.label}</div>
                <div class="field-value">${valueHtml}</div>
            </div>`;
    }

    const savedBadge = tc._saved_result
        ? `<span class="result-badge result-badge-${resultCss(tc._saved_result)}">已执行: ${escapeHtml(tc._saved_result)}</span>`
        : '';

    container.innerHTML = `
        <div class="card">
            <div class="card-header">
                <div class="card-navigator">
                    第 <strong>${tc._index + 1}</strong> 条 / 共 ${total} 条
                </div>
                ${savedBadge}
            </div>
            <div class="card-body">${fieldsHtml}</div>
        </div>

        <div class="action-card">
            <h3>\U0001f4dd 记录测试结果</h3>
            <div class="result-group">
                <button class="result-btn ${selectedResult === '通过' ? 'selected-pass' : ''}"
                        onclick="selectResult('通过', this)">
                    <span class="result-emoji">✅</span>通过
                </button>
                <button class="result-btn ${selectedResult === '失败' ? 'selected-fail' : ''}"
                        onclick="selectResult('失败', this)">
                    <span class="result-emoji">❌</span>失败
                </button>
                <button class="result-btn ${selectedResult === '阻塞' ? 'selected-block' : ''}"
                        onclick="selectResult('阻塞', this)">
                    <span class="result-emoji">\U0001f6ab</span>阻塞
                </button>
                <button class="result-btn ${selectedResult === '跳过' ? 'selected-skip' : ''}"
                        onclick="selectResult('跳过', this)">
                    <span class="result-emoji">⏭️</span>跳过
                </button>
            </div>

            <h3 style="margin-bottom:8px">\U0001f4ac 测试现象备注</h3>
            <textarea class="remark-area" id="remarkInput"
                      placeholder="请描述测试执行过程中的实际现象、发现的问题等..."
                      oninput="markDirty()">${escapeHtml(tc._saved_remark || '')}</textarea>

            <div style="margin-top:14px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;">
                <button class="btn btn-save" id="btnSave" onclick="saveResult()">
                    \U0001f4be 确认并保存
                </button>
                <span id="saveStatus" style="font-size:13px;color:var(--text-secondary);"></span>
            </div>
        </div>`;

    document.getElementById('remarkInput').addEventListener('keydown', function(e) {
        if ((e.ctrlKey || e.metaKey) && e.key === 's') {
            e.preventDefault();
            saveResult();
        }
    });
}

function formatSteps(text) {
    const lines = text.split('\n').filter(l => l.trim() !== '');
    if (lines.length <= 1) return escapeHtml(text).replace(/\n/g, '<br>');
    let html = '<ol style="padding-left:20px;margin:0;">';
    for (const line of lines) {
        html += `<li style="margin-bottom:4px;">${escapeHtml(line.trim())}</li>`;
    }
    html += '</ol>';
    return html;
}

function selectResult(value, btnEl) {
    selectedResult = value;
    markDirty();
    document.querySelectorAll('.result-btn').forEach(b => b.className = 'result-btn');
    const classMap = {'通过':'selected-pass','失败':'selected-fail','阻塞':'selected-block','跳过':'selected-skip'};
    btnEl.className = 'result-btn ' + (classMap[value] || '');
}

function resultCss(r) {
    const m = {'通过':'pass','失败':'fail','阻塞':'block','跳过':'skip'};
    return m[r] || '';
}

function markDirty() { dirty = true; }

// ============================================================
// 保存结果
// ============================================================
async function saveResult() {
    if (!selectedResult) {
        showToast('请先选择测试结果', 'warn');
        return;
    }
    const remark = document.getElementById('remarkInput').value.trim();
    const btnSave = document.getElementById('btnSave');
    const saveStatus = document.getElementById('saveStatus');
    btnSave.disabled = true;
    btnSave.textContent = '⏳ 保存中...';

    try {
        const resp = await fetch('/api/save', {
            method:'POST',
            headers:{'Content-Type':'application/json'},
            body:JSON.stringify({index:currentIndex, result:selectedResult, remark:remark}),
        });
        const data = await resp.json();
        if (data.success) {
            dirty = false;
            showToast('✅ 已保存！测试结果已写入 Excel', 'success');
            saveStatus.textContent = '✅ 已保存 ' + new Date().toLocaleTimeString('zh-CN');
            await refreshStatuses();
            allStatuses[currentIndex] = selectedResult;
        } else {
            showToast('❌ 保存失败: ' + data.error, 'error');
        }
    } catch (err) {
        showToast('❌ 网络错误: ' + err.message, 'error');
    } finally {
        btnSave.disabled = false;
        btnSave.textContent = '\U0001f4be 确认并保存';
    }
}

// ============================================================
// 导航
// ============================================================
function goTo(delta) {
    const newIndex = currentIndex + delta;
    if (newIndex < 0 || newIndex >= totalCount) return;
    loadTestCase(newIndex);
}

function updateNavButtons() {
    document.getElementById('btnPrev').disabled = currentIndex <= 0;
    document.getElementById('btnNext').disabled = currentIndex >= totalCount - 1;
}

// ============================================================
// 进度条
// ============================================================
async function refreshStatuses() {
    try {
        const resp = await fetch('/api/all-status');
        allStatuses = await resp.json();
        renderProgress();
    } catch (err) {
        console.error('获取状态失败:', err);
    }
}

function renderProgress() {
    const total = allStatuses.length;
    if (total === 0) return;
    const counts = {'通过':0,'失败':0,'阻塞':0,'跳过':0,'':0};
    for (const s of allStatuses) {
        if (counts.hasOwnProperty(s)) counts[s]++;
        else counts['']++;
    }
    const executed = total - counts[''];
    const pct = Math.round((executed / total) * 100);

    document.getElementById('progressText').textContent = `已执行 ${executed} / ${total} 条`;
    document.getElementById('progressPercent').textContent = pct + '%';
    document.getElementById('progressFill').style.width = pct + '%';
    document.getElementById('progressFill').className = 'progress-fill' + (pct === 100 ? ' done' : '');
    document.getElementById('countPass').textContent = counts['通过'];
    document.getElementById('countFail').textContent = counts['失败'];
    document.getElementById('countBlock').textContent = counts['阻塞'];
    document.getElementById('countSkip').textContent = counts['跳过'];
    document.getElementById('countNone').textContent = counts[''];
}

// ============================================================
// 搜索筛选
// ============================================================
let searchDebounceTimer = null;

function doSearch() {
    clearTimeout(searchDebounceTimer);
    searchDebounceTimer = setTimeout(_doSearch, 200);
}

async function _doSearch() {
    const keyword = document.getElementById('searchKeyword').value.trim();
    const resultFilter = document.getElementById('filterResult').value;
    const priorityFilter = document.getElementById('filterPriority').value;
    const moduleFilter = document.getElementById('filterModule').value;

    searchLastParams = {keyword, resultFilter, priorityFilter, moduleFilter, page:searchPage};

    try {
        const resp = await fetch('/api/search', {
            method:'POST',
            headers:{'Content-Type':'application/json'},
            body:JSON.stringify({
                keyword, result_filter:resultFilter,
                priority_filter:priorityFilter, module_filter:moduleFilter,
                page:searchPage, page_size:searchPageSize,
            }),
        });
        const data = await resp.json();
        renderSearchResults(data);
    } catch (err) {
        console.error('搜索失败:', err);
    }
}

function renderSearchResults(data) {
    const container = document.getElementById('searchContainer');
    const {results, total_matched, page, total_pages} = data;

    document.getElementById('searchInfo').textContent =
        total_matched > 0 ? `找到 ${total_matched} 条匹配用例` : '';

    if (results.length === 0) {
        container.innerHTML = `
            <div class="state-message">
                <div class="icon">\U0001f50d</div>
                <h2>无匹配结果</h2>
                <p>尝试更换关键词或放宽筛选条件</p>
            </div>`;
        document.getElementById('searchPagination').style.display = 'none';
        return;
    }

    // 高亮关键词
    const kw = (searchLastParams && searchLastParams.keyword) ? searchLastParams.keyword.toLowerCase() : '';

    let html = '<div class="search-result-list">';
    for (const r of results) {
        const badge = r.result
            ? `<span class="result-badge result-badge-${resultCss(r.result)}">${escapeHtml(r.result)}</span>`
            : '';
        const snippet = r.snippet ? highlightText(r.snippet, kw) : '';
        const titleDisplay = kw ? highlightText(r.title, kw) : escapeHtml(r.title);
        const idDisplay = r.id || '#' + (r.index + 1);

        html += `
            <div class="search-result-item" onclick="jumpToExecute(${r.index})" title="点击跳转到该用例">
                <div class="sr-left">
                    <span class="sr-id">${escapeHtml(idDisplay)}</span>
                    <div class="sr-title">${titleDisplay}</div>
                    <div class="sr-meta">
                        ${r.module ? escapeHtml(r.module) + ' &middot; ' : ''}
                        ${r.priority ? escapeHtml(r.priority) : ''}
                    </div>
                    ${snippet ? `<div class="sr-snippet">${snippet}</div>` : ''}
                </div>
                <div class="sr-right">
                    ${badge}
                    <span style="font-size:12px;color:var(--text-secondary);">#${r.index + 1}</span>
                </div>
            </div>`;
    }
    html += '</div>';
    container.innerHTML = html;

    // 分页
    const pagDiv = document.getElementById('searchPagination');
    if (total_pages > 1) {
        pagDiv.style.display = 'flex';
        pagDiv.innerHTML = `
            <button class="btn btn-outline btn-sm" onclick="searchGoPage(0)" ${page === 0 ? 'disabled' : ''}>首页</button>
            <button class="btn btn-outline btn-sm" onclick="searchGoPage(${page - 1})" ${page === 0 ? 'disabled' : ''}>上一页</button>
            <span class="page-info">第 ${page + 1} / ${total_pages} 页（共 ${total_matched} 条）</span>
            <button class="btn btn-outline btn-sm" onclick="searchGoPage(${page + 1})" ${page >= total_pages - 1 ? 'disabled' : ''}>下一页</button>
            <button class="btn btn-outline btn-sm" onclick="searchGoPage(${total_pages - 1})" ${page >= total_pages - 1 ? 'disabled' : ''}>末页</button>
        `;
    } else {
        pagDiv.style.display = 'none';
    }
}

function searchGoPage(p) {
    searchPage = p;
    if (searchLastParams) searchLastParams.page = p;
    _doSearch();
    window.scrollTo({top:0, behavior:'smooth'});
}

function clearSearch() {
    document.getElementById('searchKeyword').value = '';
    document.getElementById('filterResult').value = '';
    document.getElementById('filterPriority').value = '';
    document.getElementById('filterModule').value = '';
    searchPage = 0;
    doSearch();
}

function highlightText(text, keyword) {
    if (!keyword) return escapeHtml(text);
    const escaped = escapeHtml(text);
    const escapedKw = escapeHtml(keyword);
    const regex = new RegExp('(' + escapedKw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + ')', 'gi');
    return escaped.replace(regex, '<em class="highlight">$1</em>');
}

function jumpToExecute(index) {
    switchView('execute').then ? switchView('execute').then(() => loadTestCase(index)) : (() => {
        _doSwitch('execute');
        loadTestCase(index);
    })();
}

async function loadFilterOptions() {
    try {
        const resp = await fetch('/api/filter-options');
        const data = await resp.json();
        // 填充优先级下拉
        const priSel = document.getElementById('filterPriority');
        priSel.innerHTML = '<option value="">全部优先级</option>';
        for (const p of data.priorities) {
            priSel.innerHTML += `<option value="${escapeHtml(p)}">${escapeHtml(p)}</option>`;
        }
        // 填充模块下拉
        const modSel = document.getElementById('filterModule');
        modSel.innerHTML = '<option value="">全部模块</option>';
        for (const m of data.modules) {
            modSel.innerHTML += `<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`;
        }
    } catch (err) {
        console.error('加载筛选选项失败:', err);
    }
}

// ============================================================
// 汇总统计
// ============================================================
async function loadSummary() {
    const container = document.getElementById('summaryContainer');
    container.innerHTML = '<div class="state-message"><div class="icon">⏳</div><h2>加载中...</h2></div>';

    try {
        const resp = await fetch('/api/summary');
        const data = await resp.json();
        renderSummary(data);
    } catch (err) {
        container.innerHTML = '<div class="state-message"><div class="icon">❌</div><h2>加载失败</h2></div>';
    }
}

function renderSummary(data) {
    const container = document.getElementById('summaryContainer');
    const {total, counts, by_priority, by_module, execution_rate, pass_rate} = data;

    // 环形进度条 SVG
    const ringSVG = (pct, color) => {
        const r = 26, circ = 2 * Math.PI * r;
        const offset = circ * (1 - pct / 100);
        return `<svg width="64" height="64" viewBox="0 0 64 64">
            <circle class="bg" cx="32" cy="32" r="${r}"/>
            <circle class="fg" cx="32" cy="32" r="${r}"
                stroke="${color}" stroke-dasharray="${circ}" stroke-dashoffset="${offset}"/>
        </svg>`;
    };

    let html = '';

    // 统计卡片
    html += `<div class="summary-grid">
        <div class="summary-stat stat-total"><div class="stat-number">${total}</div><div class="stat-label">用例总数</div></div>
        <div class="summary-stat stat-pass"><div class="stat-number">${counts['通过']}</div><div class="stat-label">通过</div></div>
        <div class="summary-stat stat-fail"><div class="stat-number">${counts['失败']}</div><div class="stat-label">失败</div></div>
        <div class="summary-stat stat-block"><div class="stat-number">${counts['阻塞']}</div><div class="stat-label">阻塞</div></div>
        <div class="summary-stat stat-skip"><div class="stat-number">${counts['跳过']}</div><div class="stat-label">跳过</div></div>
    </div>`;

    // 指标行
    html += `<div class="summary-metrics">
        <div class="metric-card">
            <div class="metric-ring">
                ${ringSVG(execution_rate, '#4f46e5')}
                <div class="pct" style="color:#4f46e5;">${execution_rate}%</div>
            </div>
            <div class="metric-text">
                <h4>执行进度</h4>
                <p>已执行 ${total - counts['未执行']} / ${total} 条</p>
            </div>
        </div>
        <div class="metric-card">
            <div class="metric-ring">
                ${ringSVG(pass_rate, '#16a34a')}
                <div class="pct" style="color:#16a34a;">${pass_rate}%</div>
            </div>
            <div class="metric-text">
                <h4>通过率</h4>
                <p>已执行用例中的通过比例</p>
            </div>
        </div>
    </div>`;

    // 按优先级汇总
    const priKeys = Object.keys(by_priority).sort();
    if (priKeys.length > 0) {
        html += `<div class="summary-table-wrap"><h3>⚡ 按优先级分布</h3>
        <table class="summary-table">
            <tr><th>优先级</th><th>用例数</th><th style="color:var(--pass)">通过</th><th style="color:var(--fail)">失败</th><th style="color:var(--block)">阻塞</th><th>跳过</th><th>未执行</th><th>进度</th></tr>`;
        for (const p of priKeys) {
            const d = by_priority[p];
            const totalP = d['通过'] + d['失败'] + d['阻塞'] + d['跳过'] + d['未执行'];
            const execP = totalP - d['未执行'];
            const pctP = totalP > 0 ? Math.round((execP / totalP) * 100) : 0;
            html += `<tr>
                <td><strong>${escapeHtml(p)}</strong></td>
                <td>${totalP}</td>
                <td class="cell-pass">${d['通过']}</td>
                <td class="cell-fail">${d['失败']}</td>
                <td class="cell-block">${d['阻塞']}</td>
                <td>${d['跳过']}</td>
                <td>${d['未执行']}</td>
                <td class="bar-cell">
                    <div class="mini-bar"><div class="mini-bar-fill pass" style="width:${pctP}%"></div></div>
                    <span style="font-size:11px;">${pctP}%</span>
                </td>
            </tr>`;
        }
        html += `</table></div>`;
    }

    // 按模块汇总
    const modKeys = Object.keys(by_module).sort();
    if (modKeys.length > 0) {
        html += `<div class="summary-table-wrap"><h3>\U0001f4c1 按模块分布</h3>
        <table class="summary-table">
            <tr><th>模块</th><th>用例数</th><th style="color:var(--pass)">通过</th><th style="color:var(--fail)">失败</th><th style="color:var(--block)">阻塞</th><th>跳过</th><th>未执行</th><th>进度</th></tr>`;
        for (const m of modKeys) {
            const d = by_module[m];
            const totalM = d['通过'] + d['失败'] + d['阻塞'] + d['跳过'] + d['未执行'];
            const execM = totalM - d['未执行'];
            const pctM = totalM > 0 ? Math.round((execM / totalM) * 100) : 0;
            html += `<tr>
                <td><strong>${escapeHtml(m)}</strong></td>
                <td>${totalM}</td>
                <td class="cell-pass">${d['通过']}</td>
                <td class="cell-fail">${d['失败']}</td>
                <td class="cell-block">${d['阻塞']}</td>
                <td>${d['跳过']}</td>
                <td>${d['未执行']}</td>
                <td class="bar-cell">
                    <div class="mini-bar"><div class="mini-bar-fill pass" style="width:${pctM}%"></div></div>
                    <span style="font-size:11px;">${pctM}%</span>
                </td>
            </tr>`;
        }
        html += `</table></div>`;
    }

    container.innerHTML = html;
}

// ============================================================
// 未保存提示
// ============================================================
function showUnsavedDialog() {
    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.className = 'modal-overlay';
        overlay.innerHTML = `
            <div class="modal-box">
                <h3>⚠️ 有未保存的更改</h3>
                <p>当前用例的结果或备注尚未保存，<br>切换后将丢失这些更改。是否先保存？</p>
                <div class="modal-actions">
                    <button class="btn-modal-ghost" id="modalDiscard">不保存，直接切换</button>
                    <button class="btn-modal-primary" id="modalSave">先保存</button>
                </div>
            </div>`;
        document.body.appendChild(overlay);

        overlay.querySelector('#modalSave').onclick = async () => {
            document.body.removeChild(overlay);
            await saveResult();
            resolve(true);
        };
        overlay.querySelector('#modalDiscard').onclick = () => {
            document.body.removeChild(overlay);
            dirty = false;
            resolve(true);
        };
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) { document.body.removeChild(overlay); resolve(false); }
        });
    });
}

// ============================================================
// Toast
// ============================================================
function showToast(msg, type) {
    const container = document.getElementById('toastContainer');
    const toast = document.createElement('div');
    toast.className = 'toast toast-' + type;
    toast.textContent = msg;
    container.appendChild(toast);
    setTimeout(() => { if (toast.parentNode) toast.parentNode.removeChild(toast); }, 2300);
}

function escapeHtml(str) {
    const div = document.createElement('div');
    div.appendChild(document.createTextNode(str));
    return div.innerHTML;
}

// ============================================================
// 键盘快捷键
// ============================================================
document.addEventListener('keydown', function(e) {
    if (currentView !== 'execute') return;

    const tag = document.activeElement.tagName;
    const isInput = tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT';

    if (!isInput) {
        if (e.key === 'ArrowLeft') { e.preventDefault(); goTo(-1); return; }
        if (e.key === 'ArrowRight') { e.preventDefault(); goTo(1); return; }
    }

    if ((e.ctrlKey || e.metaKey) && e.key === 's') {
        e.preventDefault();
        if (currentView === 'execute') saveResult();
    }
});

// ============================================================
// 启动
// ============================================================
init();
</script>

</body>
</html>
'''


# ============================================================
# 5. 启动入口
# ============================================================
def main():
    print("=" * 50)
    print("  \U0001f52c 测试用例记录表  v1.1")
    print("=" * 50)
    print()

    filepath = find_excel_file()
    if filepath is None:
        print("⚠️  未在 testcases/ 目录中找到 Excel 文件")
        print("   请将测试用例 .xlsx 文件放入 testcases/ 文件夹后重新启动")
        print()
        print("   按回车键退出...")
        input()
    else:
        print(f"\U0001f4c2 找到文件: {filepath.name}")
        try:
            testcases, mapping, headers = read_testcases(filepath)
            STATE['testcases'] = testcases
            STATE['mapping'] = mapping
            STATE['headers'] = headers
            STATE['filepath'] = str(filepath)
            STATE['filename'] = filepath.name

            print(f"\U0001f4ca 共读取 {len(testcases)} 条测试用例")
            if mapping:
                recognized = [k for k in mapping if k != '_headers']
                print(f"\U0001f50d 识别到的字段: {', '.join(recognized)}")
            print()
        except Exception as e:
            print(f"❌ 读取 Excel 失败: {e}")
            print()
            print("   请检查:")
            print("   1. Excel 文件是否已损坏")
            print("   2. Excel 是否被其他程序占用")
            print("   按回车键退出...")
            input()

    print(f"\U0001f310 正在启动本地服务 (http://127.0.0.1:{PORT}) ...")
    print()

    def open_browser():
        import time
        time.sleep(0.8)
        webbrowser.open(f'http://127.0.0.1:{PORT}')

    threading.Thread(target=open_browser, daemon=True).start()

    import logging
    log = logging.getLogger('werkzeug')
    log.setLevel(logging.WARNING)

    print(f"✅ 服务已启动！浏览器正在打开...")
    print(f"   如果浏览器未自动打开，请手动访问: http://127.0.0.1:{PORT}")
    print()
    print("   按 Ctrl+C 可停止服务")
    print("-" * 50)

    try:
        app.run(host='127.0.0.1', port=PORT, debug=False, use_reloader=False)
    except KeyboardInterrupt:
        print("\n\U0001f44b 已停止服务，再见！")


if __name__ == '__main__':
    main()
