function loadCSS(href) {
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = href;
  document.head.appendChild(link);
}
loadCSS("assets/line-performance.css");
// ✅ 放在 Vue.component 外面
Chart.register(window['chartjs-plugin-annotation']);
Vue.component('line-performance3-view', {
 
template: `
<div class="container mt-4 line-performance-page">
  <div class="d-flex justify-content-between align-items-center mb-3">
    <h3><i class="bi bi-speedometer2 text-primary"></i> 3部-Dashboard</h3>
  </div>

  <div v-if="loading" class="text-center my-4">
    <span class="spinner-border text-primary" role="status"></span>
    <div class="mt-2">資料載入中...</div>
  </div>

  <div v-else>

    <div class="row">
      <draggable class="row w-100" v-model="topMetrics" group="metrics" :animation="200">
        <div :class="getCardSize(item)" v-for="(item, index) in topMetrics" :key="index">
          <div class="dashboard-box text-black"
               :class="[getBoxClass(item.title)]"
               @click="selectMetric(item)">

            <div class="h6">
              <i :class="getIcon(item.title)" class="me-2"></i>{{ item.title }}
            </div>

<div v-if="item.type === 'budget'" class="text-start w-100 h-100 d-flex flex-column justify-content-between">
  <div class="px-2 pt-2">
    <div class="small">
      已用 {{ item.used.toLocaleString() }} / {{ item.total.toLocaleString() }}（{{ item.percent }}%）
    </div>
  </div>
<div class="budget-chart-wrapper">
  <canvas :id="'budgetChart-' + _uid" style="width: 100%; height: 100%;"></canvas>
</div>
</div>



            <div v-else-if="item.type === 'boardEfficiency'" class="d-flex flex-column flex-grow-1 justify-content-end">
              <canvas
                :id="'miniChart'"
                style="flex-grow: 1; width: 100%; height: 100%; display: block;"
                width="400"
                height="160"
              />
            </div>

<div v-else>

<div
  class="h4 font-weight-bold"
  :style="{
    whiteSpace: 'pre-line',
    color: item.isBelowTarget ? '#d00' : '#000',
    margin: '0'
  }"
>
  {{ item.value }}
</div>


  <!-- 當月累計效率目標 -->
  <div v-if="item.title === '當月累計效率'" class="small text-end text-black mt-1" style="opacity: 0.8;">
  目標：100%  <i class="bi bi-arrow-up-right-circle-fill me-1"></i>
  </div>

  <!-- 當月累計益領目標 -->
  <div v-else-if="item.title === '當月累計溢領'" class="small text-end text-black mt-1" style="opacity: 0.8;">
  目標：0.8%   <i class="bi bi-arrow-down-right-circle-fill me-1"></i>
  </div>
    <!-- 逾期6天工單數 -->
  <div v-else-if="item.title === '逾期6天工單數'" class="small text-end text-black mt-1" style="opacity: 0.8;">
  目標： 10 張 <i class="bi bi-arrow-down-right-circle-fill me-1"></i>
  </div>
    <!-- FLOW逾期20天未結表單 -->
  <div v-else-if="item.title === 'FLOW逾期20天未結表單'" class="small text-end text-black mt-1" style="opacity: 0.8;">
  目標：0  
  </div>
    <!-- 當月累計益領目標 -->
  <div v-else-if="item.title === '當月累計溢領'" class="small text-end text-black mt-1" style="opacity: 0.8;">
  目標：0.8%   <i class="bi bi-arrow-down-right-circle-fill me-1"></i>
  </div>
</div>

          </div>
        </div>
      </draggable>
    </div>


    <table class="table table-bordered table-sm mt-4 text-center" v-if="!selectedMetric">
      <thead class="thead-dark">
        <tr>
          <th>班別</th>
          <th>標準工時(h)</th>
          <th>實際工時(h)</th>
          <th>累計效率</th>
          <th>應領金額(VND)</th>
          <th>溢領金額(VND)</th>
          <th>溢領比例</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="row in rows" :key="row.id" >
          <td>{{ row.team }}</td>
          <td>{{ row.stdHours }}</td>
          <td>{{ row.actHours }}</td>
<td :class="{ 'text-danger': row.accEfficiency < 100 }">
  {{ row.accEfficiency }}%
</td>
          <td>{{Number( row.payShould).toLocaleString() }}</td>
          <td>{{ Number(row.payActual).toLocaleString() }}</td>

<td :class="{ 'text-danger': row.accResult >0.8 }">
  {{ row.accResult }}%
</td>

        </tr>
        <tr class="font-weight-bold bg-light">
          <td>總計</td>
          <td>{{ sum(rows, 'stdHours') }}</td>
          <td>{{ sum(rows, 'actHours') }}</td>
          <td>{{ calculateEfficiencyPercent(rows) }}%</td>
          <td>{{ Number(sum(rows, 'payShould')).toLocaleString() }}</td>
          <td>{{ Number(sum(rows, 'payActual')).toLocaleString() }}</td>
          <td>{{ calculateOverClaimPercent(rows) }}%</td>

        </tr>
      </tbody>
          <span class="text-success small d-block mt-1">資料來源：ERP結案工單數據</span>
    </table>

<div v-if="selectedMetric && detailRows.length" class="mt-4">
  <div class="d-flex justify-content-between">
    <h5 class="text-primary">{{ selectedMetric }} 明細</h5>
    <el-button type="success" size="small" @click="downloadExcel">下載 Excel</el-button>
  </div>

  <!-- FLOW 表單用 Element UI -->
  <el-table v-if="selectedMetric === 'FLOW逾期20天未結表單'" :data="detailRows" class="mt-4" stripe>
    <el-table-column prop="表單名稱" label="表單名稱"></el-table-column>
    <el-table-column label="文件編號">
      <template slot-scope="scope">
        <a :href="scope.row.文件編號連結" target="_blank">{{ scope.row.文件編號 }}</a>
      </template>
    </el-table-column>
    <el-table-column prop="表單狀態" label="狀態"></el-table-column>
    <el-table-column prop="簽和關卡" label="關卡"></el-table-column>
    <el-table-column prop="申請人" label="申請人"></el-table-column>
    <el-table-column prop="建立日期" label="建立日期"></el-table-column>
  </el-table>


<!-- ✅ 即時效率專用：班別合併 -->
<table v-else-if="selectedMetric === '即時效率'" class="table table-bordered table-sm text-center mt-3">
  <thead class="thead-light">
    <tr>
      <th>班別</th>
      <th>時段</th>
      <th>預計產量</th>
      <th>實際產量</th>
      <th>NG數量</th>
      <th>即時效率</th>
      <th>不良率</th>
    </tr>
  </thead>
  <tbody>
    <tr v-for="(row, idx) in detailRows" :key="idx">
      <!-- ✅ 使用 rowspan 合併相同班別 -->
<td
  v-if="getRowspanMap(detailRows, 'department')[row.department].index === idx"
  :rowspan="getRowspanMap(detailRows, 'department')[row.department].count"
  class="text-primary"
  style="cursor: pointer; text-decoration: underline;"
  @click="gotoProdboard(row.department)"
>
  {{ row.department }}
</td>
      <td>{{ row.time }}</td>
      <td>{{ row.expected }}</td>
      <td>{{ row.actual }}</td>
      <td>{{ row.ng }}</td>
    <td :class="{ 'text-danger font-weight-bold': parseFloat(row.efficiency) < 90 }">
      {{ row.efficiency }}%
    </td>

    <!-- ✅ 不良率 > 3 紅色 -->
    <td :class="{ 'text-danger font-weight-bold': parseFloat(row.ngRate) > 3 }">
      {{ row.ngRate }}%
    </td>
    </tr>
  </tbody>
  
</table>

  <!-- 其他使用原本的 bootstrap table -->
  <table v-else class="table table-bordered table-sm text-center mt-3">
    <thead class="thead-light">
      <tr>
        <th v-for="col in detailColumns" :key="col">{{ col }}</th>
      </tr>
    </thead>
    <tbody>
      <tr v-for="(row, idx) in detailRows" :key="idx">
        <td v-for="col in detailColumns" :key="col">{{ row[col] }}</td>
      </tr>
    </tbody>
  </table>


  <div class="d-flex justify-content-between">

    <span class="text-success small d-block mt-1">{{ selectedMetricnote }}</span>
  </div>


  
</div>


  </div>
</div>
`

,

  components: {
  //  draggable: vuedraggable.default
  },

  data() {
    const now = new Date();
    const mfgMonth = now.getFullYear().toString() + String(now.getMonth() + 1).padStart(2, '0');
    return {
      query: { area: 'VN', mfgMonth },
      loading: false,
      topMetrics: [],
      rows: [],
      selectedMetric: '',
      selectedMetricnote: '',
      detailRows: [],
      detailColumns: [],
      budgetRows: [], // 👈 新增
      budgetTotal: 0, // 👈 新增
      budgetLimit: 480000000, // 固定預算
      unclosedRows: [],  // 全部資料
      unclosedCount: 0,  // 未結工單數（部=2）
      overdueCount: 0,   // 逾期工單數（逾期天數>6）
      attendanceRows: [],
      boardRows: [],     // 全部原始資料（每班每時段）
      timeSummary: [],   // 各時段統計資料（時段、總預計、總實際、總效率）\
      chartInstance: null,
      flowUnclosedRows: [],
      flowUnclosedCount: 0,
      invalidHoursRows: [],
      invalidHoursTotal: 0,
      adjustedHoursRows: [],
      adjustedHoursTotal: 0,
    };
Chart.register(window['chartjs-plugin-annotation']);
  },
  
  mounted() {
    this.fetchData();
   // this.fetchBudget(); // 👈 新增這個，載入時也抓預算
    this.fetchUnclosed(); // 👈 這個是新加的
    this.fetchAttendanceSummary();
    this.fetchBoardSummary(); // 👈加這個
    this.fetchFlowUnclosed();
    this.fetchInvalidHours();
    this.fetchAdjustedHours();
  },
  computed: {
    sortedBoardRows() {
      return this.boardRows
        .slice() // 複製原始資料
        .sort((a, b) => {
          if (a.department < b.department) return -1;
          if (a.department > b.department) return 1;
          return a.time.localeCompare(b.time);
        });
    }
  },
  methods: {

    gotoProdboard(dept) {
      this.$emit('switch-view', { target: 'production-board', dept });
    },
      getRowspanMap(data, key) {
        const map = {};
        let last = null;
        let count = 0;
    
        data.forEach((item, index) => {
          if (item[key] !== last) {
            if (count > 0) map[last].count = count;
            last = item[key];
            count = 1;
            map[item[key]] = { index, count: 1 };
          } else {
            count++;
          }
        });
    
        if (last && map[last]) {
          map[last].count = count;
        }
    
        return map;
      },
    
    
    
    sum(data, field) {
      return data.reduce((t, d) => t + parseFloat(d[field] || 0), 0).toFixed(2);
    },
    calculateEfficiencyPercent(data) {
      const totalStd = data.reduce((sum, d) => sum + parseFloat(d.stdHours || 0), 0);
      const totalAct = data.reduce((sum, d) => sum + parseFloat(d.actHours || 0), 0);
      if (totalStd === 0) return '0.00';
      return ((totalStd / totalAct) * 100).toFixed(2);
    },
    calculateOverClaimPercent(data) {
      const totalShould = data.reduce((sum, d) => sum + parseFloat(d.payShould || 0), 0);
      const totalActual = data.reduce((sum, d) => sum + parseFloat(d.payActual || 0), 0);
      if (totalShould === 0) return '0.00';
      return (((totalActual ) / totalShould) * 100).toFixed(2);
    },
    getCardSize(item) {
      switch (item.size) {
        case '1x2': return 'col-md-6 col-12 size-1x2';
        case '2x1': return 'col-md-3 col-12 size-2x1';
        case '2x2': return 'col-md-6 col-12 size-2x2';
        case '1x4': return 'col-md-6 col-12 size-1x4';   
        default:    return 'col-md-3 col-12 size-1x1';
      }
    },

    fetchAdjustedHours() {
      axios.get('https://mms.leapoptical.com:5088/api/LinePerformance/adjustedhours', {
        params: { deptPrefix: '3' }
      }).then(res => {
        this.adjustedHoursRows = res.data;
        this.adjustedHoursTotal = res.data.reduce((sum, row) => {
          return sum + parseFloat(row.調整工時總次數 || 0);
        }, 0);
        this.updateTopMetrics();
      }).catch(err => {
        console.error('調整工時查詢失敗:', err);
      });
    },
    
    fetchInvalidHours() {
      axios.get('https://mms.leapoptical.com:5088/api/LinePerformance/invalidhours', {
        params: { deptPrefix: '3' }
      }).then(res => {
        this.invalidHoursRows = res.data;
        this.invalidHoursTotal = res.data.reduce((sum, row) => {
          return sum +
            (parseFloat(row.試產工單工時 || 0) +
             parseFloat(row.樣品工單工時 || 0) +
             parseFloat(row.來料全檢工時 || 0) +
             parseFloat(row.廠內重工工時 || 0) +
             parseFloat(row.越南廠重工工時 || 0) +
             parseFloat(row.其它廠重工工時 || 0) +
             parseFloat(row.其它無效工時 || 0));
        }, 0);
        this.updateTopMetrics();
      }).catch(err => {
        console.error('無效工時查詢失敗:', err);
      });
    },
    
    fetchFlowUnclosed() {
      axios.get('https://mms.leapoptical.com:5088/api/LinePerformance/unclosedflow/list', {
        params: {
          dept: '1092710'
        }
      }).then(res => {
        this.flowUnclosedRows = res.data; // 存原始資料
        this.flowUnclosedCount = res.data.length; // 直接計算幾筆
        this.updateTopMetrics();
      }).catch(err => {
        console.error('未結FLOW表單查詢失敗:', err);
      });
    },
    fetchBoardSummary() {
      axios.get('https://mms.leapoptical.com:5088/api/LinePerformance/boardsummary', {
        params: {
          deptPrefix: '3'
        }
      }).then(res => {
        this.boardRows = res.data;  // 每班每時段原始資料
        this.timeSummary = this.aggregateByTime(res.data); // 時段統計
        this.updateTopMetrics(); // 更新小卡
      }).catch(err => {
        console.error('即時效率看板查詢失敗:', err);
      });
    },
    
    fetchUnclosed() {
      axios.get('https://mms.leapoptical.com:5088/api/WorkOrder/unclosed3')
        .then(res => {
          const data = res.data;
          this.unclosedRows = data; // 原始資料存著
           this.unclosedCount = this.unclosedRows.length;
          const validData = data.filter(x => Number(x.狀態) < 95);
    
          // ✨ 注意這邊未結工單也要部門=2
      
          this.overdueCount = validData.filter(x =>  x.逾期天數 >= 6).length;
    
          this.updateTopMetrics();
        })
        .catch(err => {
          console.error('未結工單查詢失敗:', err);
        });
    },
   
    fetchData() {
      this.loading = true;
      axios.get('https://mms.leapoptical.com:5088/api/LinePerformance/get3', {
        params: {
          AREA: this.query.area,
          MFG_MONTH: this.query.mfgMonth
        }
      }).then(res => {
        const data = res.data;
        this.rows = data.map((item, index) => ({
          id: index + 1,
          team: item.班別,
          stdHours: item.累計標準工時,
          actHours: item.累計實際工時,
          accEfficiency: item.累計效率,
          payShould: item.應領金額,
          payActual: item.實領金額,
          accResult: item.累計溢領,
          completed: item.已完成工單,
          completed1: item.未結工單數,
          completed2: item.未結逾期工單數,
          highlight: parseFloat(item.累計效率) < 90
        }));
        this.updateTopMetrics(); // 👈 別直接塞 topMetrics，改叫一個統一 function
      }).catch(err => {
        alert('查詢失敗：' + (err.response?.data?.message || err.message));
      }).finally(() => {
        this.loading = false;
      });
    },
     // 👇 新增：抓部門預算
     fetchBudget() {
      axios.get('https://mms.leapoptical.com:5088/api/OrderQuery/orderlist', {
        params: {
          department: '     1090600'
        }
      }).then(res => {
        this.budgetRows = res.data;
        this.budgetTotal = res.data.reduce((sum, item) => sum + (item.本幣金額 || 0), 0);
        this.updateTopMetrics();
      }).catch(err => {
        console.error('部門預算載入失敗:', err);
      });
    },
    
    fetchAttendanceSummary() {
      axios.get('https://mms.leapoptical.com:5088/api/LinePerformance/summary', {
        params: {
          area: 'VN',
          deptPrefix: '3'
        }
      }).then(res => {
        this.attendanceRows = res.data;
        // 這裡不直接 push 小卡，等 updateTopMetrics() 去建小卡
        this.updateTopMetrics();
      }).catch(err => {
        console.error('出席統計查詢失敗:', err);
      });
    },
    updateTopMetrics() {
      const metrics = [
        {
          title: '當月累計效率',
          value: this.calculateEfficiencyPercent(this.rows) + '%',
          size: '1x1',
          isBelowTarget:this.calculateEfficiencyPercent(this.rows)<100
        },
        {
          title: '當月累計溢領',
          value: this.calculateOverClaimPercent(this.rows) + '%',
          size: '1x1',
           isBelowTarget:this.calculateOverClaimPercent(this.rows)>0.8
        },
        {
          title: '未結工單數',
          value: this.unclosedCount  +'張',
          size: '2x1'
        },
        {
          title: '逾期6天工單數',
          value: this.overdueCount  +'張',
          size: '2x1',
           isBelowTarget:this.overdueCount>6
        }
      ];
    
      if (this.flowUnclosedCount >= 0) {
        metrics.push({
          title: 'FLOW逾期20天未結表單',
          value: this.flowUnclosedCount +'張',
          type: 'flow',
          size: '1x1',
          isBelowTarget:this.flowUnclosedCount>0
        });
      }
    
      if (this.invalidHoursTotal >= 0) {
        metrics.push({
          title: '無效工時',
          value: this.invalidHoursTotal.toFixed(1) + ' 小時',
          type: 'invalid',
          size: '2x1'
        });
      }
    
      if (this.adjustedHoursTotal >= 0) {
        metrics.push({
          title: '調整工時次數',
          value: this.adjustedHoursTotal.toFixed(0) + ' 次數',
          type: 'adjusted',
          size: '2x1'
        });
      }
    
      if (this.attendanceRows.length >= 0) {
        const totalExpected = this.attendanceRows.reduce((sum, d) => sum + d.應出席人數, 0);
        const totalExpected2 = this.attendanceRows.reduce((sum, d) => sum + d.實際出席人數, 0);
        const totalLeave = this.attendanceRows.reduce((sum, d) => sum + d.請假人數, 0);
        metrics.push({
          title: '出席統計',
          value: `應到/實到/請假\n ${totalExpected}人/${totalExpected2}人/${totalLeave}人`,
          type: 'attendance',
          size: '1x1'
        });
      }
    
      if (this.timeSummary.length >= 0) {
        metrics.unshift({
          title: '即時效率',
          type: 'boardEfficiency',
          size: '1x2'
        });
      }
    
 
      // ✅ 預先找出預算卡片資料（避免遞迴錯誤）
      const budgetMetric = metrics.find(x => x.type === 'budget');
    
      // ✅ 賦值並更新畫面
      this.topMetrics = metrics;
    
      this.$nextTick(() => {
        this.renderBudgetChart(budgetMetric);
        this.renderEfficiencyChart();
      });

    },
    renderEfficiencyChart() {
      const ctx = document.getElementById('miniChart')?.getContext('2d');
      if (!ctx || !this.timeSummary?.length) return;
    
      // ✅ 先複製資料，避免 reactive proxy 被 Chart 追蹤
      const labels = this.timeSummary.map(x => x.time.split('-')[0]);
      const values = this.timeSummary.map(x => x.efficiency);
    
      if (this.chartInstance) {
        this.chartInstance.destroy();
      }
    
      this.chartInstance = new Chart(ctx, {
        type: 'line',
        data: {
          labels,
          datasets: [{
            label: '',
            data: values,
            borderColor: '#444444',
            backgroundColor: 'rgba(255,255,255,0.2)',
            pointRadius: 3,
            borderWidth: 1,
            tension: 0.4,
    
          }]
        },
        options: {
          layout: {
            padding: {
              top: 20,
              bottom: 0,
              left:20,
            }
          },
          responsive: true,
          maintainAspectRatio: false,
          plugins: {
            legend: { display: false },
            datalabels: {
              color: 'black',
              font: { size: 10 },
              align: 'top',
              formatter: val => {
                const num = typeof val === 'object' ? val?.value ?? 0 : val;
                return `${num}%`;
              }
            }
          },
          scales: {
            x: {
              display: true,
              ticks: { font: { size: 10 }, color: 'black' },
              grid: { display: false }
            },
            y: { display: false }
          }
        },
        plugins: [ChartDataLabels]
      });
    },
    
    renderBudgetChart(budgetMetric) {
      const canvasId = `budgetChart-${this._uid}`;
      const budgetCanvas = document.getElementById(canvasId);
      const budgetCtx = budgetCanvas ? budgetCanvas.getContext('2d') : null;
  
      if (!budgetCtx || !budgetMetric) return;
  
      if (this.budgetChartInstance) {
        this.budgetChartInstance.destroy();
      }
  
      this.budgetChartInstance = new Chart(budgetCtx, {
        type: 'bar',
        data: {
          labels: budgetMetric.chartLabels,  // 1~12 月
          datasets: [{
            label: '月花費',
            data: budgetMetric.chartValues,
            backgroundColor: 'rgba(255, 193, 7, 0.8)',  // 黃色
            borderColor: 'rgba(255, 193, 7, 1)',        // 黃色邊框
            borderWidth: 1,
            barThickness: 10                           // 🎯 細長條設定
          }]
        },
        options: {
          layout: {
            padding: {
              top: 20,
              bottom: 0,
              left:0,
            }
          },
          responsive: true,
          maintainAspectRatio: false,
          scales: {
            x: {
              ticks: {
                callback: function(val, index, values) {
                  const label = this.getLabelForValue(val);
                  return label.slice(-2); // 最後兩碼
                }
              }
            },
            y: {
              beginAtZero: true,
       
              ticks: {
                callback: val => (val / 10000).toFixed(0) + '萬'
              }
            }
          },
          plugins: {
            legend: { display: false },
            datalabels: {
              anchor: 'end',
              align: 'top',
              font: { size: 10 },
              color: '#444',
              formatter: val => (val / 10000).toFixed(1) + '萬'
            },
            annotation: {
              annotations: {
                limitLine: {
                  type: 'line',
                  yMin: 40000000,
                  yMax: 40000000,
                  borderColor: 'red',
                  borderWidth: 2,
                  borderDash: [6, 4],
                  label: {
                    content: '預算上限',
                    enabled: true,
                    position: 'end',
                    backgroundColor: 'red',
                    color: 'black'
                  }
                }
              }
            }
          }
        },
        plugins: [ChartDataLabels]
      });
      
    },

    selectMetric(item) {
      this.selectedMetric = item.title;
      this.detailRows = [];
      this.detailColumns = [];
    
      if (item.title === '未結工單數') {
        // 狀態<95、部=2，並按班別、狀態排序
        const validRows = this.unclosedRows.filter(x => Number(x.狀態) < 95 );
        this.detailRows = validRows.sort((a, b) => {
          if (a.班別 < b.班別) return -1;
          if (a.班別 > b.班別) return 1;
          return Number(a.狀態) - Number(b.狀態); // 班別相同，狀態升冪排
        });
        this.detailColumns = this.detailRows.length ? Object.keys(this.detailRows[0]) : [];
          this.selectedMetricnote='資料來源:ERP工單狀態45-95'
      } else if (item.title === '逾期6天工單數') {
        // 狀態<95、部=2、逾期天數>6
        const validRows = this.unclosedRows.filter(x => Number(x.狀態) < 95  && x.逾期天數 >= 6);
        this.detailRows = validRows.sort((a, b) => {
          if (a.班別 < b.班別) return -1;
          if (a.班別 > b.班別) return 1;
          return Number(a.狀態) - Number(b.狀態);
        });
        this.detailColumns = this.detailRows.length ? Object.keys(this.detailRows[0]) : [];
              this.selectedMetricnote='資料來源:ERP工單狀態45-95,最後發料時間>6天'
      } else if (item.title === '部門預算') {
        // 部門預算明細
        this.detailRows = this.budgetRows;
        this.detailColumns = this.budgetRows.length ? Object.keys(this.budgetRows[0]) : [];
    
      } else if (item.title === '出席統計') {

        this.detailRows = this.attendanceRows;
        this.detailColumns = this.detailRows.length ? Object.keys(this.detailRows[0]) : [];
         this.selectedMetricnote='資料來源:NFC工單掛卡資料'
      } else if (item.title === '即時效率') {
        const sorted = this.boardRows
          .slice()

        this.detailRows = sorted.map(row => {
          const efficiency = row.expected > 0 ? (row.actual / row.expected * 100).toFixed(2) : '0.00';
          const ngRate = row.actual > 0 ? (row.ng / row.actual * 100).toFixed(2) : '0.00';
          return {
            department: row.department.trim(),
            time: row.time,
            expected: row.expected,
            actual: row.actual,
            ng: row.ng,
            efficiency,
            ngRate
          };
        });
      
        this.detailColumns = ['department', 'time', 'expected', 'actual', 'ng', 'efficiency', 'ngRate'];
        this.selectedMetricnote='資料來源：產線生產看板'

      }else if (item.title === 'FLOW逾期20天未結表單') {
        this.detailRows = this.flowUnclosedRows;
        this.detailColumns = this.detailRows.length ? Object.keys(this.detailRows[0]) : [];
        this.selectedMetricnote='資料來源:FLOW未結表單，部門人員(1091810)'
      }else if (item.title === '無效工時') {
        this.detailRows = this.invalidHoursRows;
        this.detailColumns = this.invalidHoursRows.length ? Object.keys(this.invalidHoursRows[0]) : [];
        this.selectedMetricnote='資料來源:FLOW共用表單，NFC資料申請單/無效工時申請單'
      }else if (item.title === '調整工時次數') {
        this.detailRows = this.adjustedHoursRows;
        this.detailColumns = this.detailRows.length ? Object.keys(this.detailRows[0]) : [];
        this.selectedMetricnote='資料來源:FLOW共用表單，NFC資料申請單/'
      }else {
        this.selectedMetric = '';
      }
    },

downloadExcel() {
  const rows = this.detailRows;
  const columns = this.detailColumns;
  if (!rows.length || !columns.length) return;

  // 建立工作表資料
  const exportData = rows.map(row => {
    const obj = {};
    columns.forEach(col => {
      obj[col] = row[col] ?? '';
    });
    return obj;
  });

  // 轉換成工作表
  const worksheet = XLSX.utils.json_to_sheet(exportData, { header: columns });

  // 建立工作簿
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

  // 寫出 Excel 並觸發下載
  const excelBuffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([excelBuffer], { type: 'application/octet-stream' });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `${this.selectedMetric}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
},
getBoxClass(title) {
  if (title.includes('FLOW')) return 'bg-white text-dark border border-primary';
  if (title.includes('出席')) return 'bg-white text-dark border border-primary';
  if (title.includes('逾期')) return 'bg-white text-dark border border-primary';
  if (title.includes('未結')) return 'bg-white text-dark border border-primary';
  return 'bg-white text-dark border-primary';
},


    
    getIcon(title) {
      if (title.includes('FLOW')) return 'bi bi-journal-text';
      if (title.includes('效率')) return 'bi bi-speedometer2';
      if (title.includes('溢領')) return 'bi bi-graph-up-arrow';
      if (title.includes('預算')) return 'bi bi-wallet2';
      if (title.includes('出席')) return 'bi bi-people-fill';
      if (title.includes('未結')) return 'bi bi-box-seam';
      if (title.includes('逾期')) return 'bi bi-exclamation-triangle';
      return 'bi bi-bar-chart';
    },
    
    aggregateByTime(rows) {
      const group = {};
      rows.forEach(row => {
        if (!group[row.time]) group[row.time] = { expected: 0, actual: 0 };
        group[row.time].expected += Number(row.expected || 0);
        group[row.time].actual += Number(row.actual || 0);
      });
    
      return Object.keys(group).map(time => {
        const expected = group[time].expected;
        const actual = group[time].actual;
        const efficiency = expected > 0 ? (actual / expected * 100).toFixed(2) : 0;
        return { time, expected, actual, efficiency };
      }).sort((a, b) => {
        const getStartMinutes = (str) => {
          const [h, m] = str.split('-')[0].split(':').map(Number);
          return h * 60 + m;
        };
        return getStartMinutes(a.time) - getStartMinutes(b.time);
      });
    }
    
    
  }
});
