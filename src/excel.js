/**
 * Excel 文件解析
 * 依赖全局 XLSX（SheetJS）
 * @namespace window.APP_EXCEL
 */
window.APP_EXCEL = {

    /**
     * 读取单个 Excel 文件，找到包含必需列的工作表并清洗数据
     * @param {File} file - HTML File 对象
     * @returns {Promise<Array>} 清洗后的行数组，每行是一个对象；找不到可用表时返回 []
     */
    loadExcelData(file) {
        const CONST = window.APP_CONSTANTS;
        const U = window.APP_UTILS;

        return new Promise((resolve) => {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array', cellDates: true });

                    let df = null;
                    for (const sheetName of workbook.SheetNames) {
                        const sheet = workbook.Sheets[sheetName];
                        // 读成对象数组，第一行作为表头
                        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                        if (!rows.length) continue;

                        // 统一别名列名（如 交易金额(元) -> 交易金额(个)）
                        const renamed = rows.map(row => {
                            const newRow = {};
                            for (const [canon, aliases] of Object.entries(CONST.COLUMN_ALIASES)) {
                                for (const alias of aliases) {
                                    if (alias in row) {
                                        newRow[canon] = row[alias];
                                        break;
                                    }
                                }
                            }
                            // 复制其他非别名列
                            for (const [k, v] of Object.entries(row)) {
                                const matched = Object.values(CONST.COLUMN_ALIASES).some(a => a.includes(k));
                                if (!matched) newRow[k] = v;
                            }
                            return newRow;
                        });

                        // 检查必需列是否都存在
                        const headers = Object.keys(renamed[0]);
                        const missing = CONST.REQUIRED_COLUMNS.filter(c => !headers.includes(c));
                        if (!missing.length) {
                            df = renamed;
                            break;
                        }
                    }

                    if (!df) { resolve([]); return; }

                    // 清洗：只保留有效数据行 + 金额转整数分
                    const cleaned = [];
                    for (const row of df) {
                        const order = String(row['订单'] || '').trim();
                        const tradeType = String(row['交易类型'] || '').trim();
                        if (!order || !tradeType) continue;

                        const cents = U.toCents(row['交易金额(个)']);
                        if (cents === 0 && parseFloat(row['交易金额(个)']) === 0) {
                            // 金额为 0 跳过（可能是空行）
                            if (!row['交易金额(个)']) continue;
                        }

                        // 时间：XLSX cellDates 已返回 Date 对象，否则尝试解析
                        let timeStr = '';
                        if (row['时间']) {
                            const d = row['时间'] instanceof Date ? row['时间'] : new Date(row['时间']);
                            if (!isNaN(d.getTime())) {
                                const hh = String(d.getHours()).padStart(2, '0');
                                const mm = String(d.getMinutes()).padStart(2, '0');
                                timeStr = `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')} ${hh}:${mm}`;
                            }
                        }

                        cleaned.push({
                            订单: order,
                            交易类型: tradeType,
                            '交易金额(个)': cents,   // 单位：分
                            状态: String(row['状态'] || '').trim(),
                            资金操作人: String(row['资金操作人'] || '').trim(),
                            时间: timeStr
                        });
                    }

                    resolve(cleaned);
                } catch (err) {
                    console.error('读取 Excel 失败:', err);
                    resolve([]);
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }
};
