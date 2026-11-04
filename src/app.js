/**
 * Vue 3 主应用（CDN 模式）
 * 负责文件选择、数据解析、面板渲染、图表展示
 * 依赖：Vue 3, SheetJS(XLSX), ECharts, 以及本目录下的 constants.js / utils.js / excel.js
 */
(function() {
    const { createApp, ref, computed, reactive, nextTick, watch, onMounted, onBeforeUnmount } = Vue;
    const U = window.APP_UTILS;
    const Excel = window.APP_EXCEL;

    /* ================================================================
     *  子组件定义
     * ================================================================ */

    /**
     * 顶部标题 + 上传栏
     */
    const HeaderBar = {
        props: ['hasFiles', 'isParsing'],
        emits: ['files-selected', 'clear'],
        template: `
            <div class="header-row">
                <h1>📊 账单退款对账看板<span class="version-badge">v1.0</span></h1>
                <div class="upload-bar">
                    <span class="file-input-wrapper">
                        <button class="upload-btn" @click="triggerFile" :disabled="isParsing">选择文件</button>
                        <input type="file" ref="fileInput" multiple accept=".xlsx,.xls" @change="onFileChange" />
                    </span>
                    <button v-if="hasFiles" class="refresh-btn" @click="$emit('clear')">清除</button>
                </div>
            </div>
        `,
        methods: {
            /** 触发隐藏的 file input */
            triggerFile() { this.$refs.fileInput.click(); },
            /** 处理文件选择 change 事件 */
            onFileChange(e) {
                const files = [...e.target.files];
                const valid = files.filter(f => !f.name.startsWith('~$') && !f.name.includes('示例'));
                if (!valid.length) {
                    alert('没有有效的 Excel 文件，请检查文件是否被 Excel 打开（临时文件以 ~$ 开头会被跳过）');
                    e.target.value = '';
                    return;
                }
                this.$emit('files-selected', valid);
            }
        }
    };

    /**
     * 数据源切换标签栏
     */
    const SourceTabs = {
        props: ['datasets', 'activeIdx'],
        emits: ['switch'],
        template: `
            <div v-if="datasets.length > 1" class="source-tabs">
                <button
                    v-for="(ds, i) in datasets"
                    :key="ds.name"
                    class="source-tab"
                    :class="{ active: i === activeIdx }"
                    @click="$emit('switch', i)"
                >{{ ds.name }}</button>
            </div>
        `
    };

    /**
     * 单个数据源面板（KPI + 日汇总 + 明细 + 图表）
     */
    const DataPanel = {
        props: ['dataset'],
        data() {
            return {
                filterDate: '',       // 日期筛选值
                filterOp: '',         // 操作人筛选值
                expandedRows: {},     // 记录哪些订单行被展开  { orderId: true }
                chartInstance: null,  // ECharts 实例引用
                chartModalVisible: false
            };
        },
        computed: {
            /** 可用的日期筛选选项（去重 + 倒序） */
            dateOptions() {
                return [...new Set(this.dataset.matched.map(m => m['日期']))].sort().reverse();
            },
            /** 可用的操作人筛选选项 */
            opOptions() {
                return [...new Set(this.dataset.matched.flatMap(m =>
                    m['操作人'].split('、').filter(Boolean)
                ))].sort();
            },
            /** 根据筛选条件过滤后的订单列表 */
            filteredOrders() {
                return this.dataset.matched.filter(m => {
                    const matchDate = !this.filterDate || m['日期'] === this.filterDate;
                    const ops = m['操作人'].split('、');
                    const matchOp = !this.filterOp || ops.includes(this.filterOp);
                    return matchDate && matchOp;
                });
            },
            /** 面板的所有过滤相关数据（给模板用） */
            t() { return this.dataset.summary['总体']; },
            dailyRows() { return this.dataset.summary['按日期']; },
            sourceInfo() {
                const now = new Date().toLocaleString('zh-CN', { hour12: false });
                return `数据来源：${this.dataset.file}　|　生成时间：${now}`;
            }
        },
        template: `
            <div class="source-panel">
                <div style="margin-bottom:16px; font-size:13px; color:#64748b;">{{ sourceInfo }}</div>

                <!-- 总体看板 -->
                <div class="card">
                    <h2>📈 总体看板</h2>
                    <div class="kpi-grid">
                        <div class="kpi-box">
                            <div class="kpi-label">订单总数</div>
                            <div class="kpi-value highlight">{{ t['订单总数'] }}</div>
                            <div class="kpi-value muted">{{ t['有退款订单数'] }}笔有退款 / {{ t['全额消费订单数'] }}笔全额消费</div>
                        </div>
                        <div class="kpi-box">
                            <div class="kpi-label">充值总金额</div>
                            <div class="kpi-value">{{ fmt(t['充值总金额']) }}</div>
                        </div>
                        <div class="kpi-box">
                            <div class="kpi-label">消费总金额</div>
                            <div class="kpi-value highlight">{{ fmt(t['消费总金额']) }}</div>
                        </div>
                        <div class="kpi-box">
                            <div class="kpi-label">退款总金额</div>
                            <div class="kpi-value refund">{{ fmt(t['退款总金额']) }}</div>
                        </div>
                        <div class="kpi-box">
                            <div class="kpi-label">实际支出总计</div>
                            <div class="kpi-value actual">{{ fmt(t['实际支出总计']) }}</div>
                        </div>
                        <div class="kpi-box" style="background:linear-gradient(135deg,#ecfdf5,#d1fae5);">
                            <div class="kpi-label">📊 数据操作</div>
                            <button class="download-btn" style="margin-top:8px;" @click="downloadExcel">下载汇总表</button>
                            <button class="download-btn" style="margin-top:8px;background:linear-gradient(135deg,#3b82f6,#1d4ed8);" @click="openChart">查看图表</button>
                        </div>
                    </div>

                    <!-- 按日期汇总 -->
                    <h2 style="margin-top: 28px;">📅 按日期汇总</h2>
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
                        <tbody>
                            <tr v-for="d in dailyRows" :key="d['日期']">
                                <td>{{ d['日期'] }}</td>
                                <td>{{ d['订单数'] }}</td>
                                <td class="amount">{{ fmt(d['充值总额']) }}</td>
                                <td class="amount">{{ fmt(d['消费总额']) }}</td>
                                <td class="amount refund-color">{{ fmt(d['退款总额']) }}</td>
                                <td class="amount actual-color">{{ fmt(d['实际支出']) }}</td>
                            </tr>
                            <tr style="background:#f1f5f9;font-weight:700;border-top:2px solid #cbd5e1;">
                                <td>总计</td>
                                <td>{{ t['订单总数'] }}</td>
                                <td class="amount">{{ fmt(t['充值总金额']) }}</td>
                                <td class="amount">{{ fmt(t['消费总金额']) }}</td>
                                <td class="amount refund-color">{{ fmt(t['退款总金额']) }}</td>
                                <td class="amount actual-color">{{ fmt(t['实际支出总计']) }}</td>
                            </tr>
                        </tbody>
                    </table>
                    </div>
                </div>

                <!-- 明细表 -->
                <div class="card">
                    <h2>📋 订单明细（点击展开交易详情）<span class="filter-count">{{ filteredOrders.length }} / {{ dataset.matched.length }} 单</span></h2>
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
                                    <select class="filter-select" v-model="filterDate">
                                        <option value="">全部日期</option>
                                        <option v-for="d in dateOptions" :key="d" :value="d">{{ d }}</option>
                                    </select>
                                </th>
                                <th>操作人
                                    <select class="filter-select" v-model="filterOp">
                                        <option value="">全部操作人</option>
                                        <option v-for="op in opOptions" :key="op" :value="op">{{ op }}</option>
                                    </select>
                                </th>
                                <th class="amount">消费总额</th>
                                <th class="amount">退款总额</th>
                                <th class="amount">实际支出</th>
                            </tr>
                        </thead>
                        <tbody>
                            <template v-for="m in filteredOrders" :key="m['订单号']">
                                <tr class="order-row">
                                    <td><span class="expand-btn" :class="{ open: expandedRows[m['订单号']] }" @click="toggle(m['订单号'])">▼</span></td>
                                    <td class="order-id">{{ m['订单号'] }}</td>
                                    <td>{{ m['日期'] }}</td>
                                    <td>{{ m['操作人'] }}</td>
                                    <td class="amount">{{ fmt(m['消费总额']) }}</td>
                                    <td class="amount refund-color">{{ fmt(m['退款总额']) }}</td>
                                    <td class="amount actual-color">{{ fmt(m['实际支出']) }}</td>
                                </tr>
                                <tr v-if="expandedRows[m['订单号']]">
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
                                            <tbody>
                                                <tr v-for="(d, di) in m['明细']" :key="di" class="detail-row">
                                                    <td>{{ d['时间'] }}</td>
                                                    <td :class="d['类型'] === '消费' ? 'type-consume' : 'type-refund'">{{ d['类型'] }}</td>
                                                    <td class="amount">{{ fmt(d['金额']) }}</td>
                                                    <td>{{ d['操作人'] }}</td>
                                                    <td>{{ d['状态'] }}</td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </td>
                                </tr>
                            </template>
                        </tbody>
                    </table>
                    </div>
                </div>

                <!-- 图表弹窗 -->
                <div class="chart-modal" :class="{ show: chartModalVisible }" @click.self="closeChart">
                    <div class="chart-modal-box">
                        <div class="chart-modal-header">
                            <div class="chart-modal-title">📈 每日数据图表 — {{ dataset.name }}</div>
                            <button class="chart-modal-close" @click="closeChart">✕</button>
                        </div>
                        <div class="chart-modal-body"><div ref="chartBox" style="width:100%;height:100%;"></div></div>
                    </div>
                </div>
            </div>
        `,
        methods: {
            /** 金额格式化（分→元 + 千分位） */
            fmt(cents) { return U.fmt_money(U.toYuan(cents)); },
            /** 展开/收起明细行 */
            toggle(orderId) {
                this.expandedRows[orderId] = !this.expandedRows[orderId];
            },
            /** 导出每日汇总表为 Excel */
            downloadExcel() {
                const daily = this.dataset.summary['按日期'];
                const data = daily.map(d => [
                    d['日期'], d['订单数'],
                    U.toYuan(d['充值总额']), U.toYuan(d['消费总额']),
                    U.toYuan(d['退款总额']), U.toYuan(d['实际支出'])
                ]);
                // 总计行
                const t = this.dataset.summary['总体'];
                const totalRow = [
                    '总计', t['订单总数'],
                    U.toYuan(t['充值总金额']), U.toYuan(t['消费总金额']),
                    U.toYuan(t['退款总金额']), U.toYuan(t['实际支出总计'])
                ];
                const sheetData = [['时间', '订单数', '充值总额', '消费总额', '退款总额', '实际支出'], ...data, totalRow];
                const ws = XLSX.utils.aoa_to_sheet(sheetData);
                ws['!cols'] = [{ wch: 14 }, { wch: 10 }, { wch: 16 }, { wch: 16 }, { wch: 16 }, { wch: 16 }];
                for (let r = 1; r < sheetData.length; r++) {
                    for (const c of [2, 3, 4, 5]) {
                        const addr = XLSX.utils.encode_cell({ r, c });
                        if (ws[addr]) ws[addr].z = '#,##0.00';
                    }
                }
                const totalRowIdx = sheetData.length - 1;
                for (let c = 0; c <= 5; c++) {
                    const addr = XLSX.utils.encode_cell({ r: totalRowIdx, c });
                    if (ws[addr]) ws[addr].s = { bold: true };
                }
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, '每日汇总');
                const now = new Date();
                const fname = `每日汇总_${this.dataset.name}_${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}.xlsx`;
                XLSX.writeFile(wb, fname);
            },
            /** 打开图表弹窗并渲染 ECharts */
            openChart() {
                const daily = this.dataset.summary['按日期'];
                if (!daily.length) return;
                this.chartModalVisible = true;
                nextTick(() => {
                    const container = this.$refs.chartBox;
                    if (this.chartInstance) this.chartInstance.dispose();
                    this.chartInstance = echarts.init(container);
                    this.chartInstance.setOption({
                        tooltip: {
                            trigger: 'axis',
                            axisPointer: { type: 'cross' },
                            formatter: function(params) {
                                let tip = params[0].axisValue + '<br/>';
                                params.forEach(p => {
                                    const val = p.value != null ? p.value.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '-';
                                    tip += p.marker + p.seriesName + '：' + val + '<br/>';
                                });
                                return tip;
                            }
                        },
                        legend: { data: ['充值总额', '消费总额', '退款总额', '实际支出'], top: 8 },
                        grid: { left: 60, right: 30, top: 50, bottom: 80 },
                        dataZoom: [
                            { type: 'inside', start: 0, end: 100 },
                            { type: 'slider', start: 0, end: 100, bottom: 10, height: 20 }
                        ],
                        xAxis: {
                            type: 'category',
                            data: daily.map(d => d['日期']),
                            axisLabel: { rotate: 45, fontSize: 11, interval: 0 },
                            axisTick: { alignWithLabel: true }
                        },
                        yAxis: {
                            type: 'value',
                            axisLabel: { formatter: v => v >= 10000 ? (v / 10000).toFixed(0) + 'w' : v }
                        },
                        series: [
                            { name: '充值总额', type: 'bar', data: daily.map(d => U.toYuan(d['充值总额'])), itemStyle: { color: '#60a5fa' }, barGap: '0%' },
                            { name: '消费总额', type: 'bar', data: daily.map(d => U.toYuan(d['消费总额'])), itemStyle: { color: '#2563eb' } },
                            { name: '退款总额', type: 'bar', data: daily.map(d => U.toYuan(d['退款总额'])), itemStyle: { color: '#f87171' } },
                            { name: '实际支出', type: 'line', data: daily.map(d => U.toYuan(d['实际支出'])), smooth: true, itemStyle: { color: '#059669' }, lineStyle: { width: 3 }, symbol: 'circle', symbolSize: 8, z: 10 }
                        ]
                    });
                    setTimeout(() => this.chartInstance && this.chartInstance.resize(), 0);
                });
            },
            /** 关闭图表弹窗 */
            closeChart() {
                this.chartModalVisible = false;
                if (this.chartInstance) { this.chartInstance.dispose(); this.chartInstance = null; }
            }
        },
        onBeforeUnmount() {
            if (this.chartInstance) this.chartInstance.dispose();
        }
    };

    /* ================================================================
     *  根组件 App
     * ================================================================ */
    const App = {
        components: { HeaderBar, SourceTabs, DataPanel },
        template: `
            <HeaderBar
                :hasFiles="hasFiles"
                :isParsing="isParsing"
                @files-selected="handleFilesSelected"
                @clear="handleClear"
            />
            <SourceTabs
                v-if="hasFiles"
                :datasets="datasets"
                :activeIdx="activeIdx"
                @switch="handleSwitch"
            />
            <DataPanel
                v-if="currentDataset"
                :key="activeIdx"
                :dataset="currentDataset"
            />
        `,
        data() {
            return {
                datasets: [],          // 解析后的数据源数组
                activeIdx: 0,          // 当前显示的数据源索引
                isParsing: false       // 是否正在解析
            };
        },
        computed: {
            hasFiles() { return this.datasets.length > 0; },
            currentDataset() { return this.datasets[this.activeIdx] || null; }
        },
        methods: {
            /**
             * 处理文件选择：读取 → 解析 → 渲染
             * @param {File[]} files - 用户选择的 Excel 文件列表
             */
            async handleFilesSelected(files) {
                this.isParsing = true;
                this.datasets = [];
                this.activeIdx = 0;

                const datasets = [];
                for (const file of files) {
                    const df = await Excel.loadExcelData(file);
                    if (!df.length) {
                        console.warn(`${file.name} 没有找到可用的数据`);
                        continue;
                    }
                    const matched = U.match_refund_orders(df);
                    const rechargeDaily = U.aggregate_recharge(df);
                    const summary = U.build_summary(matched, rechargeDaily);
                    datasets.push({
                        name: file.name.replace(/\.(xlsx|xls)$/i, ''),
                        file: file.name,
                        matched,
                        summary
                    });
                }

                this.isParsing = false;

                if (!datasets.length) {
                    alert('❌ 所有文件都没有有效的数据，请检查表头是否包含：订单、交易类型、交易金额(个)、状态、资金操作人、时间');
                    return;
                }

                this.datasets = datasets;
            },

            /** 切换数据源 */
            handleSwitch(idx) { this.activeIdx = idx; },

            /** 清除所有已选文件 */
            handleClear() {
                this.datasets = [];
                this.activeIdx = 0;
            }
        }
    };

    /* ================================================================
     *  挂载应用
     * ================================================================ */
    createApp(App).mount('#app');
})();
