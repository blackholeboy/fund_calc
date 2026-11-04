/**
 * 全局常量定义（对应 process.py）
 * @namespace window.APP_CONSTANTS
 */
window.APP_CONSTANTS = {

    /** 必需的列名，Excel 表头必须包含 */
    REQUIRED_COLUMNS: ['订单', '交易类型', '交易金额(个)', '状态', '资金操作人', '时间'],

    /** 金额列别名：不同导出版本可能叫「交易金额(个)」或「交易金额(元)」 */
    COLUMN_ALIASES: {
        '交易金额(个)': ['交易金额(个)', '交易金额(元)']
    },

    /** 退款类型名称集合 */
    REFUND_TYPES: new Set(['退豆', '退还']),

    /** 充值类型名称集合 */
    RECHARGE_TYPES: new Set(['充值'])
};
