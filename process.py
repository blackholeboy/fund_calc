"""
账单退款对账工具
- 自动扫描当前文件夹所有 xlsx 文件（排除 Excel 临时文件 ~$ 开头）
- 解析【账单详情】工作表，匹配消费与退款，计算实际支出
- 生成 HTML 看板，包含总体概览、按日期汇总、订单明细
使用方法：
  1. 把新的账单 Excel 文件放进本文件夹（保持表头格式一致）
  2. 双击 run.command 或命令行执行：python3 process.py
  3. 自动生成/覆盖 fund_report.html，打开即可查看
"""

import pandas as pd
from collections import defaultdict
from datetime import datetime
from decimal import Decimal
from pathlib import Path
import glob
import os

# 支持的列名列表（兼容可能的小差异）
REQUIRED_COLUMNS = ['订单', '交易类型', '交易金额(个)', '状态', '资金操作人', '时间']

# 金额列的别名映射：不同导出版本可能叫「交易金额(个)」或「交易金额(元)」
COLUMN_ALIASES = {'交易金额(个)': ['交易金额(个)', '交易金额(元)']}

# 退款类型的名称集合：不同导出版本可能叫「退豆」或「退还」
REFUND_TYPES = {'退豆', '退还'}

# 充值类型名称集合
RECHARGE_TYPES = {'充值'}


def aggregate_recharge(df: pd.DataFrame) -> dict:
    """
    按日期汇总充值交易金额
    充值是独立的资金流入交易，不参与消费/退款/实际支出的净额计算，
    但需要在日期汇总表中单独列出，便于对账。

    @param df: 清洗后的账单 DataFrame
    @return: 按日期汇总的充值金额字典，key 为日期字符串，value 为 Decimal 金额
    """
    # 过滤出充值类型的交易
    recharge_df = df[df['交易类型'].isin(RECHARGE_TYPES)]
    recharge_daily = defaultdict(lambda: Decimal('0'))
    for _, row in recharge_df.iterrows():
        # 优先用订单号时间戳归属日期；充值的订单号可能是"微信豆"等占位词，
        # 此时回退用 Excel 的"时间"列取日期（该列必有值）
        order_id = str(row['订单']).strip()
        order_date = get_date_from_order(order_id)
        if order_date == '未知' and pd.notna(row['时间']):
            order_date = row['时间'].strftime('%Y-%m-%d')
        recharge_daily[order_date] += row['交易金额(个)']
    return dict(recharge_daily)


def find_excel_files(folder: Path) -> list:
    """
    扫描文件夹及 SheetData 子目录中所有可用的 Excel 文件

    @param folder: 目标文件夹路径
    @return: Excel 文件路径列表（排除 Excel 临时文件 ~$ 开头和示例文件）
    """
    files = []
    # 同时扫描根目录与 SheetData 子目录
    for directory in [folder, folder / 'SheetData']:
        if not directory.is_dir():
            continue
        for f in sorted(directory.glob('*.xlsx')):
            # 跳过 Excel 临时文件（~$ 开头）
            if f.name.startswith('~$'):
                continue
            # 跳过示例文件（仅作格式参考，不参与计算）
            if '示例' in f.name:
                continue
            files.append(f)
    return files


def load_data(filepath: Path) -> pd.DataFrame:
    """
    从单个 Excel 文件加载并清洗账单数据
    - 自动扫描所有工作表，找到包含必需列的那张（不再要求必须叫「账单详情」）
    - 兼容金额列名差异：交易金额(个) / 交易金额(元)

    @param filepath: Excel 文件路径
    @return: 清洗后的 DataFrame，额外标记来源文件；找不到可用表时返回空 DataFrame
    """
    try:
        sheets = pd.read_excel(filepath, sheet_name=None)
    except Exception as e:
        print(f'  ⚠️ 跳过 {filepath.name}：读取失败 ({e})')
        return pd.DataFrame()

    df = None
    for sheet_name, sheet_df in sheets.items():
        # 把别名列名统一成规范列名（如 交易金额(元) -> 交易金额(个)）
        rename_map = {}
        for canon, aliases in COLUMN_ALIASES.items():
            for alias in aliases:
                if alias in sheet_df.columns:
                    rename_map[alias] = canon
                    break
        sheet_df = sheet_df.rename(columns=rename_map)
        # 检查必需列是否都存在
        missing = [c for c in REQUIRED_COLUMNS if c not in sheet_df.columns]
        if not missing:
            df = sheet_df
            break

    if df is None:
        print(f'  ⚠️ 跳过 {filepath.name}：所有工作表都缺少必需列 {REQUIRED_COLUMNS}')
        return pd.DataFrame()

    # 只保留有效数据列，去除右侧汇总区域
    df = df[REQUIRED_COLUMNS].copy()
    # 去除空行
    df = df.dropna(subset=['订单', '交易类型'])
    # 转换金额为数值，再剔除无法转成数字的行
    df['交易金额(个)'] = pd.to_numeric(df['交易金额(个)'], errors='coerce')
    df = df.dropna(subset=['交易金额(个)'])
    # 金额统一转为 Decimal 精确计算，避免浮点累加误差（涉及金额必须精确）
    df['交易金额(个)'] = df['交易金额(个)'].map(lambda v: Decimal(repr(float(v))))
    # 转换时间
    df['时间'] = pd.to_datetime(df['时间'], errors='coerce')
    # 标记来源文件（用于排查）
    df['_来源文件'] = filepath.name
    return df


def fmt_money(value) -> str:
    """
    金额格式化：千分位分隔 + 固定两位小数，规避浮点显示误差

    @param value: 金额数值（Decimal / float / int）
    @return: 格式化后的金额字符串，如 296,554.28
    """
    return f'{value:,.2f}'


def load_all_data(folder: Path) -> pd.DataFrame:
    """
    加载文件夹内所有 Excel 文件的账单详情，合并为一个 DataFrame

    @param folder: 数据文件夹
    @return: 合并后的 DataFrame
    """
    excel_files = find_excel_files(folder)
    print(f'📂 发现 {len(excel_files)} 个 Excel 文件：')
    for f in excel_files:
        print(f'   - {f.name}')

    dfs = []
    for f in excel_files:
        df = load_data(f)
        if not df.empty:
            dfs.append(df)

    if not dfs:
        print('❌ 没有可读取的数据！')
        return pd.DataFrame()

    combined = pd.concat(dfs, ignore_index=True)
    # 按订单号去重（避免多文件重复）
    combined = combined.drop_duplicates(subset=['订单', '交易类型', '交易金额(个)', '时间'])
    return combined

def get_date_from_order(order_id: str) -> str:
    """
    从订单号中提取 Unix 时间戳并转本地日期（北京时间）
    订单号格式：1793116800_654400，前面 10 位是秒级 Unix 时间戳

    @param order_id: 基础订单号（不含"退款"后缀）
    @return: YYYY-MM-DD 日期字符串，解析失败返回 "未知"
    """
    try:
        ts = int(order_id.split('_')[0])
        if ts > 1e12:  # 毫秒级时间戳（13位），转秒
            ts = ts / 1000
        return datetime.fromtimestamp(ts).strftime('%Y-%m-%d')
    except Exception:
        return '未知'


def match_refund_orders(df: pd.DataFrame) -> list:
    """
    匹配所有消费订单及其退款：
    - 包含所有有消费的订单（不管有没有退款）
    - 同一基础订单号的多笔消费和退款分别累加
    - 实际支出 = 消费总额 - 退款总额
    - 日期归属：用订单号里的 Unix 时间戳转本地日期

    @param df: 清洗后的账单 DataFrame
    @return: 订单净额列表，每条包含订单号、消费总额、退款总额、实际支出、日期等
    """
    # 按基础订单号聚合
    consume_map = defaultdict(list)  # 基础订单号 -> 消费记录列表
    refund_map = defaultdict(list)   # 基础订单号 -> 退款记录列表

    for _, row in df.iterrows():
        order_id = str(row['订单']).strip()
        trade_type = row['交易类型']

        # 判断是消费还是退款（退款单号以"退款"结尾，类型可能是"退豆"或"退还"）
        if order_id.endswith('退款'):
            base_order = order_id[:-2]  # 去掉"退款"
            if trade_type in REFUND_TYPES:
                refund_map[base_order].append(row)
        else:
            if trade_type == '消费':
                consume_map[order_id].append(row)

    # 合并所有有消费的订单（有退款的订单也要包含）
    all_base_orders = set(consume_map.keys()) | set(refund_map.keys())

    matched = []
    for base_order in all_base_orders:
        consume_rows = consume_map.get(base_order, [])
        refund_rows = refund_map.get(base_order, [])

        consume_total = sum(r['交易金额(个)'] for r in consume_rows)
        refund_total = sum(r['交易金额(个)'] for r in refund_rows)
        actual_cost = consume_total - refund_total

        # 日期归属：优先用订单号时间戳（匹配 Excel 透视表口径）
        order_date = get_date_from_order(base_order)

        # 收集操作人（去重）
        operators = list(set(
            str(r['资金操作人']) for r in consume_rows + refund_rows
            if pd.notna(r['资金操作人'])
        ))

        # 明细交易记录（合并消费+退款，按时间排序）
        details = []
        for r in consume_rows:
            details.append({
                '时间': r['时间'].strftime('%Y-%m-%d %H:%M') if pd.notna(r['时间']) else '',
                '类型': '消费',
                '金额': r['交易金额(个)'],
                '操作人': str(r['资金操作人']) if pd.notna(r['资金操作人']) else '',
                '状态': str(r['状态']) if pd.notna(r['状态']) else ''
            })
        for r in refund_rows:
            details.append({
                '时间': r['时间'].strftime('%Y-%m-%d %H:%M') if pd.notna(r['时间']) else '',
                '类型': str(r['交易类型']),
                '金额': r['交易金额(个)'],
                '操作人': str(r['资金操作人']) if pd.notna(r['资金操作人']) else '',
                '状态': str(r['状态']) if pd.notna(r['状态']) else ''
            })
        # 按时间排序明细
        details.sort(key=lambda x: x['时间'])

        matched.append({
            '订单号': base_order,
            '日期': order_date,
            '消费总额': consume_total,
            '退款总额': refund_total,
            '实际支出': actual_cost,
            '操作人': '、'.join(operators),
            '消费笔数': len(consume_rows),
            '退款笔数': len(refund_rows),
            '明细': details
        })

    # 按日期倒序
    matched.sort(key=lambda x: x['日期'], reverse=True)
    return matched


def build_summary(matched: list, recharge_daily: dict) -> dict:
    """
    构建汇总看板数据：总体统计 + 按日期汇总
    口径：
    - 消费总金额 = 所有消费交易之和
    - 退款总金额 = 所有退豆交易之和
    - 实际支出总计 = 消费总金额 - 退款总金额
    - 充值总额 = 所有充值交易之和（独立维度，不影响实际支出）
    - 按日期汇总：用订单号时间戳归属日期（与 Excel 透视表一致）

    @param matched: 匹配后的订单净额列表
    @param recharge_daily: 按日期汇总的充值金额字典
    @return: 汇总数据字典
    """
    total_consume = sum(m['消费总额'] for m in matched)
    total_refund = sum(m['退款总额'] for m in matched)
    total_actual = sum(m['实际支出'] for m in matched)
    # 充值总额：所有日期的充值累加
    total_recharge = sum(recharge_daily.values(), Decimal('0'))
    total_orders = len(matched)
    total_consume_tx = sum(m['消费笔数'] for m in matched)
    total_refund_tx = sum(m['退款笔数'] for m in matched)

    # 收集所有涉及的日期（消费/退款日期 ∪ 充值日期）
    all_dates = set(m['日期'] for m in matched) | set(recharge_daily.keys())

    # 按日期汇总（日期来自订单号时间戳）
    daily = defaultdict(lambda: {'消费总额': Decimal('0'), '退款总额': Decimal('0'), '实际支出': Decimal('0'), '充值总额': Decimal('0'), '订单数': 0})
    for m in matched:
        d = m['日期']
        daily[d]['消费总额'] += m['消费总额']
        daily[d]['退款总额'] += m['退款总额']
        daily[d]['实际支出'] += m['实际支出']
        daily[d]['订单数'] += 1
    # 注入充值数据（某些日期可能只有充值没有消费/退款）
    for d, amount in recharge_daily.items():
        daily[d]['充值总额'] += amount

    daily_list = [{'日期': d, **v} for d, v in sorted(daily.items(), reverse=True)]

    # 有退款的订单数（用于区分"全额消费"和"部分退款"）
    refunded_orders = sum(1 for m in matched if m['退款笔数'] > 0)
    pure_consume_orders = sum(1 for m in matched if m['退款笔数'] == 0 and m['消费笔数'] > 0)

    return {
        '总体': {
            '订单总数': total_orders,
            '有退款订单数': refunded_orders,
            '全额消费订单数': pure_consume_orders,
            '消费总笔数': total_consume_tx,
            '退款总笔数': total_refund_tx,
            '充值总金额': total_recharge,
            '消费总金额': total_consume,
            '退款总金额': total_refund,
            '实际支出总计': total_actual,
            '退款率': f"{(total_refund / total_consume * 100):.2f}%" if total_consume > 0 else "0%"
        },
        '按日期': daily_list
    }

def generate_html(datasets: list, output_path: str):
    """
    生成 HTML 报告文件（多个数据源独立统计，页面可切换查看，数据不合并）

    @param datasets: 数据源列表，每项包含 name(显示名)/file(文件名)/matched(订单列表)/summary(汇总数据)
    @param output_path: 输出 HTML 文件路径
    """
    # 生成时间
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    # 数据源切换标签（每个 xlsx 一个标签）
    tabs_html = ''
    for i, ds in enumerate(datasets):
        active_cls = ' active' if i == 0 else ''
        tabs_html += f'<button class="source-tab{active_cls}" onclick="switchSource({i})">{ds["name"]}</button>'

    # 每个数据源独立生成一个面板，统计数据互不合并
    panels_html = ''
    for i, ds in enumerate(datasets):
        matched = ds['matched']
        summary = ds['summary']
        # 该数据源的数据来源信息
        source_info = f"数据来源：{ds['file']}　|　生成时间：{now_str}"
        # 明细表筛选下拉选项：日期（倒序）与操作人（去重排序）
        filter_dates = sorted({m['日期'] for m in matched}, reverse=True)
        filter_ops = sorted({op for m in matched for op in m['操作人'].split('、') if op})
        date_options = ''.join(f'<option value="{v}">{v}</option>' for v in filter_dates)
        op_options = ''.join(f'<option value="{v}">{v}</option>' for v in filter_ops)
        # 生成明细表 HTML
        detail_rows_html = ''
        for m in matched:
            # 生成该订单的可展开明细（5 列：时间/类型/金额/操作人/状态）
            detail_items = ''
            for d in m['明细']:
                type_class = 'type-consume' if d['类型'] == '消费' else 'type-refund'
                detail_items += f"""
                <tr class="detail-row">
                    <td>{d['时间']}</td>
                    <td class="{type_class}">{d['类型']}</td>
                    <td class="amount">{fmt_money(d['金额'])}</td>
                    <td>{d['操作人']}</td>
                    <td>{d['状态']}</td>
                </tr>"""

            refund_amount = m['退款总额']

            detail_rows_html += f"""
            <tr class="order-row" data-date="{m['日期']}" data-ops="{m['操作人']}">
                <td class="expand-cell"><span class="expand-btn" onclick="toggleDetail(this)">▼</span></td>
                <td class="order-id">{m['订单号']}</td>
                <td>{m['日期']}</td>
                <td>{m['操作人']}</td>
                <td class="amount">{fmt_money(m['消费总额'])}</td>
                <td class="amount refund-color">{fmt_money(refund_amount)}</td>
                <td class="amount actual-color">{fmt_money(m['实际支出'])}</td>
            </tr>
            <tr class="detail-container" style="display:none;">
                <td class="inner-indent"></td>
                <td class="inner-indent"></td>
                <td class="inner-indent"></td>
                <td colspan="4" style="padding:0;">
                    <table class="inner-table">
                        <colgroup>
                            <col style="width:22%;">
                            <col style="width:14%;">
                            <col style="width:16%;">
                            <col style="width:16%;">
                            <col style="width:32%;">
                        </colgroup>
                        <thead>
                            <tr>
                                <th>时间</th>
                                <th>类型</th>
                                <th class="amount">交易金额(个)</th>
                                <th>操作人</th>
                                <th>状态</th>
                            </tr>
                        </thead>
                        <tbody>{detail_items}</tbody>
                    </table>
                </td>
            </tr>"""

        # 生成日汇总表 HTML（6 列：时间 | 订单数 | 充值总额 | 消费总额 | 退款总额 | 实际支出）
        daily_rows_html = ''
        for d in summary['按日期']:
            daily_rows_html += f"""
            <tr>
                <td>{d['日期']}</td>
                <td>{d['订单数']}</td>
                <td class="amount">{fmt_money(d['充值总额'])}</td>
                <td class="amount">{fmt_money(d['消费总额'])}</td>
                <td class="amount refund-color">{fmt_money(d['退款总额'])}</td>
                <td class="amount actual-color">{fmt_money(d['实际支出'])}</td>
            </tr>"""

        # 总体数据
        t = summary['总体']

        # 组装该数据源的完整面板（看板 + 日汇总 + 明细表）
        display_style = 'block' if i == 0 else 'none'
        panels_html += f"""
<div class="source-panel" id="panel-{i}" data-source="{ds['name']}" style="display:{display_style};">
<div style="margin-bottom:16px; font-size:13px; color:#64748b;">{source_info}</div>

<!-- 总体看板 -->
<div class="card">
    <h2>📈 总体看板</h2>
    <div class="kpi-grid">
        <div class="kpi-box">
            <div class="kpi-label">订单总数</div>
            <div class="kpi-value highlight">{t['订单总数']}</div>
            <div class="kpi-value muted">{t['有退款订单数']}笔有退款 / {t['全额消费订单数']}笔全额消费</div>
        </div>
        <div class="kpi-box">
            <div class="kpi-label">充值总金额</div>
            <div class="kpi-value">{fmt_money(t['充值总金额'])}</div>
        </div>
        <div class="kpi-box">
            <div class="kpi-label">消费总金额</div>
            <div class="kpi-value highlight">{fmt_money(t['消费总金额'])}</div>
        </div>
        <div class="kpi-box">
            <div class="kpi-label">退款总金额</div>
            <div class="kpi-value refund">{fmt_money(t['退款总金额'])}</div>
        </div>
        <div class="kpi-box">
            <div class="kpi-label">实际支出总计</div>
            <div class="kpi-value actual">{fmt_money(t['实际支出总计'])}</div>
        </div>
        <div class="kpi-box" style="background:linear-gradient(135deg,#ecfdf5,#d1fae5);">
            <div class="kpi-label">📊 数据操作</div>
            <button onclick="downloadExcel({i})" class="download-btn" style="margin-top:8px;">下载汇总表</button>
            <button onclick="openChart({i})" class="download-btn" style="margin-top:8px;background:linear-gradient(135deg,#3b82f6,#1d4ed8);">查看图表</button>
        </div>
    </div>

    <!-- 按日期汇总 -->
    <h2 style="margin-top: 28px;">📅 按日期汇总（与导出 Excel 一致）</h2>
    <div class="table-wrap">
    <table class="daily-table">
        <thead>
            <tr>
                <th>时间</th>
                <th>订单数</th>
                <th class="amount">充值总额</th>
                <th class="amount">消费总额</th>
                <th class="amount">退款总额</th>
                <th class="amount">实际支出</th>
            </tr>
        </thead>
        <tbody>{daily_rows_html}
            <tr style="background:#f1f5f9;font-weight:700;border-top:2px solid #cbd5e1;">
                <td>总计</td>
                <td>{t['订单总数']}</td>
                <td class="amount">{fmt_money(t['充值总金额'])}</td>
                <td class="amount">{fmt_money(t['消费总金额'])}</td>
                <td class="amount refund-color">{fmt_money(t['退款总金额'])}</td>
                <td class="amount actual-color">{fmt_money(t['实际支出总计'])}</td>
            </tr>
        </tbody>
    </table>
    </div>
</div>

<!-- 明细表 -->
<div class="card">
    <h2>📋 订单明细（点击展开交易详情）<span class="filter-count" id="filter-count-{i}"></span></h2>
    <div class="table-wrap">
    <table class="main-table">
        <colgroup>
            <col style="width:36px;">
            <col style="width:24%;">
            <col style="width:13%;">
            <col style="width:15%;">
            <col style="width:16%;">
            <col style="width:16%;">
            <col style="width:16%;">
        </colgroup>
        <thead>
            <tr>
                <th></th>
                <th>订单号</th>
                <th>日期
                    <select class="filter-select" id="filter-date-{i}" onchange="filterOrders({i})">
                        <option value="">全部日期</option>{date_options}
                    </select>
                </th>
                <th>操作人
                    <select class="filter-select" id="filter-op-{i}" onchange="filterOrders({i})">
                        <option value="">全部操作人</option>{op_options}
                    </select>
                </th>
                <th class="amount">消费总额</th>
                <th class="amount">退款总额</th>
                <th class="amount">实际支出</th>
            </tr>
        </thead>
        <tbody id="order-tbody-{i}">{detail_rows_html}</tbody>
    </table>
    </div>
</div>
</div>"""

    html = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>账单退款对账看板</title>
<style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
        background: #f0f2f5;
        color: #1f2937;
        padding: 24px;
        line-height: 1.6;
    }}
    h1 {{
        font-size: 24px;
        font-weight: 600;
        margin-bottom: 20px;
        color: #111827;
    }}
    h2 {{
        font-size: 18px;
        font-weight: 600;
        margin-bottom: 16px;
        color: #374151;
        border-left: 4px solid #3b82f6;
        padding-left: 12px;
    }}
    .card {{
        background: #fff;
        border-radius: 12px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08), 0 1px 2px rgba(0,0,0,0.06);
        padding: 24px;
        margin-bottom: 24px;
    }}
    .kpi-grid {{
        display: grid;
        grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
        gap: 16px;
        margin-bottom: 24px;
    }}
    .kpi-box {{
        background: linear-gradient(135deg, #f8fafc, #eef2ff);
        border-radius: 10px;
        padding: 18px 20px;
        border: 1px solid #e2e8f0;
    }}
    .kpi-label {{
        font-size: 13px;
        color: #64748b;
        margin-bottom: 6px;
    }}
    .kpi-value {{
        font-size: 24px;
        font-weight: 700;
        color: #1e293b;
        font-variant-numeric: tabular-nums;
    }}
    .kpi-value.highlight {{ color: #2563eb; }}
    .kpi-value.refund {{ color: #dc2626; }}
    .kpi-value.actual {{ color: #059669; }}
    .kpi-value.muted {{ color: #64748b; font-size: 18px; }}
    .download-btn {{
        display: block;
        width: 100%;
        margin-top: 10px;
        padding: 10px 16px;
        background: linear-gradient(135deg, #059669, #047857);
        color: #fff;
        border: none;
        border-radius: 8px;
        font-size: 14px;
        font-weight: 600;
        cursor: pointer;
        transition: all 0.2s;
    }}
    .download-btn:hover {{
        background: linear-gradient(135deg, #047857, #065f46);
        transform: translateY(-1px);
        box-shadow: 0 4px 12px rgba(5,150,105,0.3);
    }}

    table {{
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
    }}
    /* 主明细表：auto 布局让 colspan 单元格能撑满整行，内部 inner-table 才会满宽 */
    .main-table {{ table-layout: auto; }}
    .daily-table {{ table-layout: fixed; }}
    .main-table th.amount, .main-table td.amount {{ text-align: right; }}
    thead th {{
        background: #f1f5f9;
        color: #475569;
        font-weight: 600;
        text-align: left;
        padding: 12px 14px;
        border-bottom: 2px solid #e2e8f0;
        white-space: nowrap;
    }}
    /* 两个主表表头吸顶（滚动时固定在视口顶部） */
    .daily-table thead th, .main-table thead th {{
        position: sticky;
        top: 0;
        z-index: 10;
        background: #f1f5f9;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.06);
    }}

    /* 数据源切换标签栏 */
    .source-tabs {{
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
        margin-bottom: 18px;
    }}
    .source-tab {{
        padding: 9px 22px;
        background: #fff;
        border: 1px solid #e2e8f0;
        border-radius: 999px;
        font-size: 14px;
        font-weight: 500;
        color: #475569;
        cursor: pointer;
        transition: all 0.2s;
    }}
    .source-tab:hover {{
        border-color: #3b82f6;
        color: #2563eb;
    }}
    .source-tab.active {{
        background: linear-gradient(135deg, #2563eb, #1d4ed8);
        border-color: #2563eb;
        color: #fff;
        box-shadow: 0 2px 8px rgba(37, 99, 235, 0.3);
    }}
    /* 桌面端不创建滚动容器，保证吸顶相对视口生效；移动端才允许横向滚动 */
    .table-wrap {{ overflow-x: visible; }}
    tbody td {{
        padding: 10px 14px;
        border-bottom: 1px solid #f1f5f9;
    }}
    tbody tr:hover {{ background: #f8fafc; }}
    .order-row {{ cursor: pointer; }}
    .order-id {{ font-family: "SF Mono", Monaco, Consolas, monospace; font-size: 13px; color: #1e40af; }}
    .amount {{ text-align: right; font-variant-numeric: tabular-nums; }}
    .refund-color {{ color: #dc2626; }}
    .actual-color {{ color: #059669; font-weight: 600; }}
    .type-consume {{ color: #2563eb; font-weight: 600; }}
    .type-refund {{ color: #dc2626; font-weight: 600; }}
    .expand-btn {{
        display: inline-block;
        width: 20px;
        height: 20px;
        line-height: 20px;
        text-align: center;
        cursor: pointer;
        font-size: 10px;
        color: #94a3b8;
        user-select: none;
        transition: transform 0.2s;
    }}
    .expand-btn.open {{ transform: rotate(180deg); }}
    .inner-indent {{ border-bottom: 1px solid #f1f5f9; background: #fafbfc; }}
    .inner-table {{ width: 100%; margin: 0; box-shadow: none; background: #fafbfc; }}
    .inner-table th {{ background: #e2e8f0; font-size: 12px; padding: 6px 10px; text-align: left; border-bottom: 1px solid #cbd5e1; }}
    .inner-table td {{ padding: 6px 10px; font-size: 12px; border-bottom: 1px dashed #e2e8f0; background: #fafbfc; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }}
    .detail-row td {{ padding: 6px 10px; font-size: 12px; color: #64748b; }}

    /* 表头筛选下拉框 */
    .filter-select {{
        display: block;
        width: 100%;
        margin-top: 4px;
        padding: 4px 6px;
        font-size: 12px;
        color: #334155;
        background: #fff;
        border: 1px solid #cbd5e1;
        border-radius: 6px;
        cursor: pointer;
        outline: none;
    }}
    .filter-select:focus {{ border-color: #3b82f6; }}
    .filter-count {{ font-size: 12px; color: #64748b; font-weight: 400; margin-left: 8px; }}

    @media (max-width: 768px) {{
        body {{ padding: 12px; }}
        .card {{ padding: 16px; }}
        table {{ font-size: 12px; }}
        thead th, tbody td {{ padding: 8px 6px; }}
        /* 移动端恢复横向滚动（会牺牲吸顶，小屏优先保证可读性） */
        .table-wrap {{ overflow-x: auto; }}
    }}

    /* 图表弹窗样式 */
    .chart-modal {{
        display: none;
        position: fixed;
        top: 0; left: 0;
        width: 100%; height: 100%;
        background: rgba(0, 0, 0, 0.5);
        z-index: 1000;
        justify-content: center;
        align-items: center;
    }}
    .chart-modal.show {{ display: flex; }}
    .chart-modal-box {{
        background: #fff;
        border-radius: 12px;
        width: min(960px, 94vw);
        height: min(600px, 86vh);
        box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        display: flex;
        flex-direction: column;
        overflow: hidden;
    }}
    .chart-modal-header {{
        padding: 16px 20px;
        border-bottom: 1px solid #e2e8f0;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }}
    .chart-modal-title {{ font-size: 16px; font-weight: 600; color: #1e293b; }}
    .chart-modal-close {{
        background: none; border: none; font-size: 20px; cursor: pointer; color: #64748b;
        padding: 4px 10px; border-radius: 6px; transition: background 0.15s;
    }}
    .chart-modal-close:hover {{ background: #f1f5f9; color: #1e293b; }}
    .chart-modal-body {{ flex: 1; padding: 16px; min-height: 0; }}
    #chart-container {{ width: 100%; height: 100%; }}
</style>
</head>
<body>
<h1>📊 账单退款对账看板</h1>

<!-- 数据源切换标签栏 -->
<div class="source-tabs">{tabs_html}</div>

<!-- 各数据源独立面板（统计不合并） -->
{panels_html}

<!-- SheetJS 本地 -->
<script src="assets/xlsx.full.min.js"></script>
<!-- ECharts 本地 -->
<script src="assets/echarts.min.js"></script>
<script>
function toggleDetail(btn) {{
    const row = btn.closest('.order-row');
    const nextRow = row.nextElementSibling;
    const isOpen = nextRow.style.display !== 'none';
    btn.classList.toggle('open');
    nextRow.style.display = isOpen ? 'none' : 'table-row';
}}

function switchSource(idx) {{
    // 切换数据源面板：显示目标面板，隐藏其他面板
    document.querySelectorAll('.source-panel').forEach((p, i) => {{
        p.style.display = i === idx ? 'block' : 'none';
    }});
    // 更新标签栏高亮状态
    document.querySelectorAll('.source-tab').forEach((b, i) => {{
        b.classList.toggle('active', i === idx);
    }});
    // 切换后重新应用该面板的筛选计数
    filterOrders(idx);
}}

function filterOrders(idx) {{
    // 读取当前面板日期/操作人两个下拉框的筛选值
    const dateVal = document.getElementById('filter-date-' + idx).value;
    const opVal = document.getElementById('filter-op-' + idx).value;
    const rows = document.querySelectorAll('#order-tbody-' + idx + ' .order-row');
    let visible = 0;
    rows.forEach(row => {{
        // 日期精确匹配；操作人为"、"分隔的多值，按包含匹配
        const matchDate = !dateVal || row.dataset.date === dateVal;
        const ops = (row.dataset.ops || '').split('、');
        const matchOp = !opVal || ops.includes(opVal);
        const show = matchDate && matchOp;
        row.style.display = show ? '' : 'none';
        // 隐藏时同步收起并隐藏其明细子行
        const detailRow = row.nextElementSibling;
        if (detailRow && detailRow.classList.contains('detail-container')) {{
            detailRow.style.display = 'none';
            const btn = row.querySelector('.expand-btn');
            if (btn) btn.classList.remove('open');
        }}
        if (show) visible++;
    }});
    // 更新表头计数，提示当前筛选结果
    const counter = document.getElementById('filter-count-' + idx);
    if (counter) counter.textContent = visible + ' / ' + rows.length + ' 单';
}}

function downloadExcel(idx) {{
    // 定位当前数据源面板：提取其每日汇总表的所有列
    const panel = document.getElementById('panel-' + idx);
    const table = panel.querySelector('.daily-table');
    const sourceName = panel.dataset.source || '';
    const rows = table.querySelectorAll('tbody tr');
    // 列索引对应：0=时间, 1=订单数, 2=充值总额, 3=消费总额, 4=退款总额, 5=实际支出
    const data = [];
    for (let i = 0; i < rows.length; i++) {{
        const cells = rows[i].querySelectorAll('td');
        const label = cells[0].textContent.trim();
        if (label === '总计') continue;
        // 提取每列数值（跳过时间列 label）
        const row = [label];
        for (let c = 1; c <= 5; c++) {{
            const str = cells[c].textContent.trim().replace(/,/g, '');
            row.push(Math.round((parseFloat(str) || 0) * 100) / 100);
        }}
        data.push(row);
    }}
    // 计算总计行（数值列逐项累加）
    const totals = ['总计', 0, 0, 0, 0, 0];
    data.forEach(r => {{
        for (let c = 1; c <= 5; c++) totals[c] = Math.round((totals[c] + r[c]) * 100) / 100;
    }});
    data.push(totals);

    // 创建 Excel（6 列：时间 | 订单数 | 充值总额 | 消费总额 | 退款总额 | 实际支出）
    const sheetData = [['时间', '订单数', '充值总额', '消费总额', '退款总额', '实际支出'], ...data];
    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    ws['!cols'] = [{{ wch: 14 }}, {{ wch: 10 }}, {{ wch: 16 }}, {{ wch: 16 }}, {{ wch: 16 }}, {{ wch: 16 }}];
    // 金额列统一两位小数显示格式（列索引 2-5：充值/消费/退款/实际支出）
    for (let r = 1; r < sheetData.length; r++) {{
        for (let c of [2, 3, 4, 5]) {{
            const addr = XLSX.utils.encode_cell({{ r: r, c: c }});
            if (ws[addr]) ws[addr].z = '#,##0.00';
        }}
    }}
    // 总计行加粗
    const totalRowIdx = sheetData.length - 1;
    for (let c = 0; c <= 5; c++) {{
        const addr = XLSX.utils.encode_cell({{ r: totalRowIdx, c: c }});
        if (ws[addr]) ws[addr].s = {{ bold: true }};
    }}

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '每日汇总');

    const now = new Date();
    // 文件名附加原表格名字，便于区分各数据源导出
    const fname = `每日汇总_${{sourceName}}_${{now.getFullYear()}}${{String(now.getMonth()+1).padStart(2,'0')}}${{String(now.getDate()).padStart(2,'0')}}.xlsx`;
    XLSX.writeFile(wb, fname);
}}

// 当前弹窗实例（避免重复创建）
let chartModal = null;
let chartInstance = null;

/**
 * 打开图表弹窗并渲染 ECharts
 * 从指定面板的日汇总表读取数据，充值/消费/退款用柱状图，实际支出用折线图叠加
 * @param {{number}} idx - 数据源面板索引
 */
function openChart(idx) {{
    const panel = document.getElementById('panel-' + idx);
    const table = panel.querySelector('.daily-table');
    const sourceName = panel.dataset.source || '';
    const rows = table.querySelectorAll('tbody tr');
    const data = [];
    // 列索引：0=时间, 1=订单数, 2=充值, 3=消费, 4=退款, 5=实际支出
    for (let i = 0; i < rows.length; i++) {{
        const cells = rows[i].querySelectorAll('td');
        const label = cells[0].textContent.trim();
        if (label === '总计') continue;
        const nums = [];
        for (let c = 1; c <= 5; c++) {{
            const str = cells[c].textContent.trim().replace(/,/g, '');
            nums.push(parseFloat(str) || 0);
        }}
        data.push({{ date: label, orders: nums[0], recharge: nums[1], consume: nums[2], refund: nums[3], actual: nums[4] }});
    }}
    if (!data.length) return;

    // 弹窗 DOM（首次创建，后续复用）
    if (!chartModal) {{
        chartModal = document.createElement('div');
        chartModal.className = 'chart-modal';
        chartModal.innerHTML = `
        <div class="chart-modal-box">
            <div class="chart-modal-header">
                <div class="chart-modal-title" id="chart-title"></div>
                <button class="chart-modal-close" onclick="closeChart()">✕</button>
            </div>
            <div class="chart-modal-body"><div id="chart-container"></div></div>
        </div>`;
        chartModal.addEventListener('click', (e) => {{ if (e.target === chartModal) closeChart(); }});
        document.body.appendChild(chartModal);
    }}
    chartModal.classList.add('show');
    document.getElementById('chart-title').textContent = '📈 每日数据图表 — ' + sourceName;

    // 初始化/复用 ECharts
    const container = document.getElementById('chart-container');
    if (chartInstance) {{ chartInstance.dispose(); }}
    chartInstance = echarts.init(container);

    const option = {{
        tooltip: {{
            trigger: 'axis',
            axisPointer: {{ type: 'cross' }},
            formatter: function(params) {{
                let tip = params[0].axisValue + '<br/>';
                params.forEach(p => {{
                    const marker = p.marker;
                    const val = p.value != null ? p.value.toLocaleString('zh-CN', {{minimumFractionDigits: 2, maximumFractionDigits: 2}}) : '-';
                    tip += marker + p.seriesName + '：' + val + '<br/>';
                }});
                return tip;
            }}
        }},
        legend: {{ data: ['充值总额', '消费总额', '退款总额', '实际支出'], top: 8 }},
        grid: {{ left: 60, right: 30, top: 50, bottom: 80 }},
        dataZoom: [
            {{ type: 'inside', start: 0, end: 100 }},
            {{ type: 'slider', start: 0, end: 100, bottom: 10, height: 20 }}
        ],
        xAxis: {{
            type: 'category',
            data: data.map(d => d.date),
            axisLabel: {{ rotate: 45, fontSize: 11, interval: 0 }},
            axisTick: {{ alignWithLabel: true }}
        }},
        yAxis: {{
            type: 'value',
            axisLabel: {{ formatter: v => v >= 10000 ? (v / 10000).toFixed(0) + 'w' : v }}
        }},
        series: [
            {{ name: '充值总额', type: 'bar', data: data.map(d => d.recharge), itemStyle: {{ color: '#60a5fa' }}, barGap: '0%' }},
            {{ name: '消费总额', type: 'bar', data: data.map(d => d.consume), itemStyle: {{ color: '#2563eb' }} }},
            {{ name: '退款总额', type: 'bar', data: data.map(d => d.refund), itemStyle: {{ color: '#f87171' }} }},
            {{ name: '实际支出', type: 'line', data: data.map(d => d.actual), smooth: true, itemStyle: {{ color: '#059669' }}, lineStyle: {{ width: 3 }}, symbol: 'circle', symbolSize: 8, z: 10 }}
        ]
    }};
    chartInstance.setOption(option);
    // 窗口尺寸变化时自适应
    setTimeout(() => chartInstance && chartInstance.resize(), 0);
}}

/** 关闭图表弹窗 */
function closeChart() {{
    if (chartModal) chartModal.classList.remove('show');
    if (chartInstance) {{ chartInstance.dispose(); chartInstance = null; }}
}}

// 页面加载后初始化各面板筛选计数显示
document.querySelectorAll('.source-panel').forEach((p, i) => filterOrders(i));
</script>
</body>
</html>"""

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)

def main():
    """主入口：扫描当前文件夹与 SheetData 下所有 Excel，按文件独立统计生成对账 HTML"""
    folder = Path(__file__).parent.resolve()
    output_path = folder / 'fund_report.html'

    print('=' * 50)
    print('📊 账单退款对账工具')
    print('=' * 50)

    print('\n📂 扫描 Excel 文件...')
    excel_files = find_excel_files(folder)
    if not excel_files:
        print('❌ 没有可读取的数据！')
        return

    # 每个 xlsx 独立统计，互不合并
    datasets = []
    for f in excel_files:
        df = load_data(f)
        if df.empty:
            continue
        # 文件名（不含扩展名）作为该数据源的显示名
        name = f.stem
        matched = match_refund_orders(df)
        # 按日期汇总充值数据（充值是独立维度，不参与消费/退款净额计算）
        recharge_daily = aggregate_recharge(df)
        summary = build_summary(matched, recharge_daily)
        datasets.append({'name': name, 'file': f.name, 'matched': matched, 'summary': summary})
        t = summary['总体']
        print(f'  ✅ {f.name}: {len(df)} 条记录 / {len(matched)} 单 / 充值 {fmt_money(t["充值总金额"])} / 实际支出 {fmt_money(t["实际支出总计"])}')

    if not datasets:
        print('❌ 没有可读取的数据！')
        return

    print(f'\n📝 生成 HTML 报告 -> {output_path.name}（{len(datasets)} 个数据源，独立统计）')
    generate_html(datasets, str(output_path))

    # 分别打印各数据源概览
    for ds in datasets:
        t = ds['summary']['总体']
        print()
        print('=' * 50)
        print(f"📊 对账概览 - {ds['name']}")
        print('=' * 50)
        print(f"  有退款订单数:   {t['有退款订单数']}")
        print(f"  消费总金额:     {fmt_money(t['消费总金额'])}")
        print(f"  退款总金额:     {fmt_money(t['退款总金额'])}")
        print(f"  退款率:         {t['退款率']}")
        print(f"  充值总金额:     {fmt_money(t['充值总金额'])}")
        print(f"  实际支出总计:   {fmt_money(t['实际支出总计'])}")
        print('=' * 50)
    print()
    print(f'✅ 完成！打开 {output_path.name} 查看报告')

    # 自动打开 HTML（可选）：Windows 用 os.startfile，Mac 用 open
    try:
        if os.name == 'nt':
            os.startfile(str(output_path))
        else:
            import subprocess
            subprocess.run(['open', str(output_path)], check=False)
    except Exception:
        pass

if __name__ == '__main__':
    main()
