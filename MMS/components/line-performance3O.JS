function loadCSS(href) {
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = href;
  document.head.appendChild(link);
}
loadCSS("assets/line-performance.css");

Vue.component('line-performance3-view', {
  template: `
    <div class="container mt-4 line-performance-page">
      <div class="d-flex justify-content-between align-items-center mb-3">
        <h3><i class="bi bi-speedometer2 text-primary"></i> 二部產線效率統計</h3>
      </div>

      <div v-if="loading" class="text-center my-4">
        <span class="spinner-border text-primary" role="status"></span>
        <div class="mt-2">資料載入中...</div>
      </div>

      <div v-else>
        <div class="row">
          <draggable class="row w-100" v-model="topMetrics" :options="{ group: 'metrics', animation: 200 }">
            <div class="col-md-3" v-for="(item, index) in topMetrics" :key="index">
              <div class="dashboard-box text-white" :class="getBoxClass(item.title)" @click="selectMetric(item)">
                <div class="h6"><i :class="getIcon(item.title)"></i> {{ item.title }}</div>
                <div class="h4 font-weight-bold">{{ item.value }}</div>
              </div>
            </div>
          </draggable>
        </div>

        <!-- 表格區塊 -->
        <table class="table table-bordered table-sm mt-4 text-center" v-if="!selectedMetric">
          <thead class="thead-dark">
            <tr>
              <th>班別</th>
              <th>標準工時</th>
              <th>實際工時</th>
              <th>累計效率</th>
              <th>應領金額</th>
              <th>實領金額</th>
              <th>溢領比例</th>
              <th>已結工單</th>
              <th>未結工單</th>
              <th>逾期 >6 天</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="row in rows" :key="row.id" :class="{ 'low-efficiency': row.highlight }">
              <td>{{ row.team }}</td>
              <td>{{ row.stdHours }}</td>
              <td>{{ row.actHours }}</td>
              <td>{{ row.accEfficiency }}%</td>
              <td>{{ row.payShould }}</td>
              <td>{{ row.payActual }}</td>
              <td>{{ row.accResult }}%</td>
              <td>{{ row.completed }}</td>
              <td>{{ row.completed1 }}</td>
              <td>{{ row.completed2 }}</td>
            </tr>
            <tr class="font-weight-bold bg-light">
              <td>總計</td>
              <td>{{ sum(rows, 'stdHours') }}</td>
              <td>{{ sum(rows, 'actHours') }}</td>
              <td>{{ calculateEfficiencyPercent(rows) }}%</td>
              <td>{{ sum(rows, 'payShould') }}</td>
              <td>{{ sum(rows, 'payActual') }}</td>
              <td>{{ calculateOverClaimPercent(rows) }}%</td>
              <td>{{ sum(rows, 'completed') }}</td>
              <td>{{ sum(rows, 'completed1') }}</td>
              <td>{{ sum(rows, 'completed2') }}</td>
            </tr>
          </tbody>
        </table>

        <!-- 明細表格切換 -->
        <div v-if="selectedMetric && detailRows.length" class="mt-4">
          <div class="d-flex justify-content-between">
            <h5 class="text-primary">{{ selectedMetric }} 明細</h5>
            <button class="btn btn-sm btn-success" @click="downloadExcel">下載 Excel</button>
          </div>
          <table class="table table-bordered table-sm text-center">
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
        </div>
      </div>
    </div>
  `,
  components: {
    draggable: vuedraggable
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
      detailRows: [],
      detailColumns: []
    };
  },
  mounted() {
    this.fetchData();
    this.fetchBudget(); // 👈 新增這個，載入時也抓預算
  },
  methods: {
    sum(data, field) {
      return data.reduce((t, d) => t + parseFloat(d[field] || 0), 0).toFixed(2);
    },
    calculateEfficiencyPercent(data) {
      const totalStd = data.reduce((sum, d) => sum + parseFloat(d.stdHours || 0), 0);
      const totalAct = data.reduce((sum, d) => sum + parseFloat(d.actHours || 0), 0);
      if (totalStd === 0) return '0.00';
      return ((totalAct / totalStd) * 100).toFixed(2);
    },
    calculateOverClaimPercent(data) {
      const totalShould = data.reduce((sum, d) => sum + parseFloat(d.payShould || 0), 0);
      const totalActual = data.reduce((sum, d) => sum + parseFloat(d.payActual || 0), 0);
      if (totalShould === 0) return '0.00';
      return (((totalActual - totalShould) / totalShould) * 100).toFixed(2);
    },
    fetchData() {
      this.loading = true;
      axios.get('http://localhost:5117/api/LinePerformance/get3', {
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
        this.topMetrics = [
          { title: '加權平均效率', value: this.calculateEfficiencyPercent(this.rows) + '%' },
          { title: '總體溢領比例', value: this.calculateOverClaimPercent(this.rows) + '%' },
          { title: '應領總額', value: this.sum(this.rows, 'payShould') },
          { title: '實領總額', value: this.sum(this.rows, 'payActual') },
          { title: '逾期工單總數', value: this.sum(this.rows, 'completed2') }
        ];
      }).catch(err => {
        alert('查詢失敗：' + (err.response?.data?.message || err.message));
      }).finally(() => {
        this.loading = false;
      });
    },
    selectMetric(item) {
      this.selectedMetric = item.title;
      this.detailRows = [];
      this.detailColumns = [];
      if (item.title === '逾期工單總數') {
        this.loading = true;
        axios.get('http://localhost:5117/api/WorkOrder/unclosed3')
          .then(res => {
            this.detailRows = res.data;
            this.detailColumns = res.data.length ? Object.keys(res.data[0]) : [];
          })
          .catch(err => {
            alert('查詢逾期工單失敗：' + (err.response?.data?.message || err.message));
          })
          .finally(() => {
            this.loading = false;
          });
      } else {
        this.selectedMetric = ''; // 顯示回原始表格
      }
    },
    downloadExcel() {
      const rows = this.detailRows;
      const columns = this.detailColumns;
      if (!rows.length) return;

      let csvContent = columns.join(',') + '\n';
      rows.forEach(row => {
        csvContent += columns.map(col => '"' + (row[col] ?? '') + '"').join(',') + '\n';
      });

      const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', `${this.selectedMetric}.csv`);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
    },
    getBoxClass(title) {
      if (title.includes('效率')) return 'bg-success';
      if (title.includes('溢領')) return 'bg-warning';
      if (title.includes('應領') || title.includes('實領')) return 'bg-info';
      return 'bg-secondary';
    },
    getIcon(title) {
      if (title.includes('效率')) return 'bi bi-speedometer2';
      if (title.includes('溢領')) return 'bi bi-graph-up-arrow';
      if (title.includes('應領')) return 'bi bi-cash-coin';
      return 'bi bi-bar-chart';
    }
  }
});
