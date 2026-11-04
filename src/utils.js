/**
 * 辅助函数 + 数据处理逻辑
 * 移植自 process.py 的核心算法
 * @namespace window.APP_UTILS
 */
window.APP_UTILS = {

    /* ============ 辅助函数 ============ */

    /**
     * 金额格式化：千分位分隔 + 固定两位小数
     * @param {number} value - 金额数值
     * @returns {string} 格式化后的金额，如 "296,554.28"
     */
    fmt_money(value) {
        const num = Number(value) || 0;
        return num.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    },

    /**
     * 从订单号中提取 Unix 时间戳并转本地日期
     * 订单号格式：1793116800_654400，前面 10 位是秒级 Unix 时间戳
     * @param {string} orderId - 基础订单号（不含"退款"后缀）
     * @returns {string} YYYY-MM-DD 日期字符串，解析失败返回 "未知"
     */
    get_date_from_order(orderId) {
        try {
            const ts = parseInt(String(orderId).split('_')[0], 10);
            if (isNaN(ts)) return '未知';
            // 毫秒级时间戳（13位）转秒
            const seconds = ts > 1e12 ? ts / 1000 : ts;
            const d = new Date(seconds * 1000);
            const y = d.getFullYear();
            const m = String(d.getMonth() + 1).padStart(2, '0');
            const day = String(d.getDate()).padStart(2, '0');
            return `${y}-${m}-${day}`;
        } catch (e) {
            return '未知';
        }
    },

    /**
     * 数值转整数分（避免浮点累加误差），内部计算用整数
     * @param {*} val - 任意值，尝试转数字再取整
     * @returns {number} 乘以 100 后的整数
     */
    toCents(val) {
        const num = parseFloat(String(val).replace(/,/g, ''));
        return Math.round((isNaN(num) ? 0 : num) * 100);
    },

    /**
     * 把整数"分"转回"元"（用于显示和导出）
     * @param {number} cents - 整数，单位分
     * @returns {number} 单位元的浮点数
     */
    toYuan(cents) { return Number((cents / 100).toFixed(2)); },

    /* ============ 数据处理核心 ============ */

    /**
     * 匹配所有消费订单及其退款
     * 规则与 process.py 完全一致：
     * - 退款单号以"退款"结尾，去掉后缀后匹配消费订单
     * - 同一基础订单号的多笔消费和退款分别累加
     * - 日期归属：用订单号里的 Unix 时间戳转本地日期
     * @param {Array} df - 清洗后的行数组
     * @returns {Array} 订单净额列表
     */
    match_refund_orders(df) {
        const CONST = window.APP_CONSTANTS;
        const consumeMap = {};  // baseOrder -> [rows]
        const refundMap = {};   // baseOrder -> [rows]

        for (const row of df) {
            const orderId = row['订单'];
            const tradeType = row['交易类型'];

            if (orderId.endsWith('退款')) {
                const baseOrder = orderId.slice(0, -2);
                if (CONST.REFUND_TYPES.has(tradeType)) {
                    (refundMap[baseOrder] = refundMap[baseOrder] || []).push(row);
                }
            } else {
                if (tradeType === '消费') {
                    (consumeMap[orderId] = consumeMap[orderId] || []).push(row);
                }
            }
        }

        const allBaseOrders = new Set([...Object.keys(consumeMap), ...Object.keys(refundMap)]);
        const matched = [];

        for (const baseOrder of allBaseOrders) {
            const consumeRows = consumeMap[baseOrder] || [];
            const refundRows = refundMap[baseOrder] || [];

            let consumeTotal = 0, refundTotal = 0;
            consumeRows.forEach(r => consumeTotal += r['交易金额(个)']);
            refundRows.forEach(r => refundTotal += r['交易金额(个)']);
            const actualCost = consumeTotal - refundTotal;

            const orderDate = this.get_date_from_order(baseOrder);

            // 收集操作人（去重）
            const opSet = new Set();
            [...consumeRows, ...refundRows].forEach(r => {
                if (r['资金操作人']) opSet.add(r['资金操作人']);
            });

            // 明细交易记录（按时间排序）
            const details = [];
            consumeRows.forEach(r => {
                details.push({
                    时间: r['时间'],
                    类型: '消费',
                    金额: r['交易金额(个)'],
                    操作人: r['资金操作人'],
                    状态: r['状态']
                });
            });
            refundRows.forEach(r => {
                details.push({
                    时间: r['时间'],
                    类型: r['交易类型'],
                    金额: r['交易金额(个)'],
                    操作人: r['资金操作人'],
                    状态: r['状态']
                });
            });
            details.sort((a, b) => a['时间'].localeCompare(b['时间']));

            matched.push({
                订单号: baseOrder,
                日期: orderDate,
                消费总额: consumeTotal,
                退款总额: refundTotal,
                实际支出: actualCost,
                操作人: [...opSet].join('、'),
                消费笔数: consumeRows.length,
                退款笔数: refundRows.length,
                明细: details
            });
        }

        // 按日期倒序
        matched.sort((a, b) => b['日期'].localeCompare(a['日期']));
        return matched;
    },

    /**
     * 按日期汇总充值交易金额（充值不参与消费/退款净额计算，但在日汇总表单独列出）
     * @param {Array} df - 清洗后的行数组
     * @returns {Object} 日期 -> 金额(分) 的字典
     */
    aggregate_recharge(df) {
        const CONST = window.APP_CONSTANTS;
        const daily = {};
        for (const row of df) {
            if (!CONST.RECHARGE_TYPES.has(row['交易类型'])) continue;
            const orderId = row['订单'];
            let orderDate = this.get_date_from_order(orderId);
            // 充值订单号可能是占位词，回退用时间列
            if (orderDate === '未知' && row['时间']) {
                orderDate = row['时间'].slice(0, 10);
            }
            daily[orderDate] = (daily[orderDate] || 0) + row['交易金额(个)'];
        }
        return daily;
    },

    /**
     * 构建汇总看板数据：总体统计 + 按日期汇总
     * @param {Array} matched - match_refund_orders 的返回值
     * @param {Object} rechargeDaily - aggregate_recharge 的返回值
     * @returns {Object} 汇总数据
     */
    build_summary(matched, rechargeDaily) {
        let totalConsume = 0, totalRefund = 0;
        let totalConsumeTx = 0, totalRefundTx = 0;
        let refundedOrders = 0, pureConsumeOrders = 0;

        matched.forEach(m => {
            totalConsume += m['消费总额'];
            totalRefund += m['退款总额'];
            totalConsumeTx += m['消费笔数'];
            totalRefundTx += m['退款笔数'];
            if (m['退款笔数'] > 0) refundedOrders++;
            else if (m['消费笔数'] > 0) pureConsumeOrders++;
        });

        const totalActual = totalConsume - totalRefund;
        const totalRecharge = Object.values(rechargeDaily).reduce((s, v) => s + v, 0);

        // 收集所有日期
        const allDates = new Set([
            ...matched.map(m => m['日期']),
            ...Object.keys(rechargeDaily)
        ]);

        // 按日期汇总
        const dailyMap = {};
        allDates.forEach(d => {
            dailyMap[d] = { 消费总额: 0, 退款总额: 0, 实际支出: 0, 充值总额: 0, 订单数: 0 };
        });
        matched.forEach(m => {
            const d = m['日期'];
            dailyMap[d]['消费总额'] += m['消费总额'];
            dailyMap[d]['退款总额'] += m['退款总额'];
            dailyMap[d]['实际支出'] += m['实际支出'];
            dailyMap[d]['订单数']++;
        });
        Object.entries(rechargeDaily).forEach(([d, amt]) => {
            if (!dailyMap[d]) dailyMap[d] = { 消费总额: 0, 退款总额: 0, 实际支出: 0, 充值总额: 0, 订单数: 0 };
            dailyMap[d]['充值总额'] += amt;
        });

        // 倒序排列
        const dailyList = Object.entries(dailyMap)
            .map(([d, v]) => ({ 日期: d, ...v }))
            .sort((a, b) => b['日期'].localeCompare(a['日期']));

        const refundRate = totalConsume > 0 ? ((totalRefund / totalConsume) * 100).toFixed(2) + '%' : '0%';

        return {
            '总体': {
                '订单总数': matched.length,
                '有退款订单数': refundedOrders,
                '全额消费订单数': pureConsumeOrders,
                '消费总笔数': totalConsumeTx,
                '退款总笔数': totalRefundTx,
                '充值总金额': totalRecharge,
                '消费总金额': totalConsume,
                '退款总金额': totalRefund,
                '实际支出总计': totalActual,
                '退款率': refundRate
            },
            '按日期': dailyList
        };
    }
};
