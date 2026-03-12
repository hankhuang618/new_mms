// unclosed.js
(function loadCSS(href) {
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = href;
  document.head.appendChild(link);
})("assets/unclosed.css");


Vue.component('unclosed-workorder-view', {
  template: `
    <div class="unclosed-page container mt-4">
      <h4 class="mb-3" style="color:#1976d2; font-weight:bold;">未結工單查詢</h4>

      <canvas id="unclosedChart" height="120" class="mb-4"></canvas>

      <div v-if="filterText" class="mb-3">
        <el-alert
          :title="'篩選：' + filterText"
          type="info"
          show-icon
          :closable="true"
          @close="clearFilter"
        />
      </div>

      <el-table
        :data="summaryRows"
        border
        stripe
        size="mini"
        class="mb-4"
        :cell-class-name="cellClass"
        @cell-click="handleCellClick"
      >
        <el-table-column label="班別" prop="班別" fixed></el-table-column>
        <el-table-column v-for="s in statusList" :key="s" :label="'狀態'+s" :prop="s"></el-table-column>
        <el-table-column v-for="s in statusList" :key="'逾期'+s" :label="'逾期' + s" :prop="'逾期' + s"></el-table-column>
        <el-table-column label="總計" width="80" align="center">
          <template slot-scope="scope">
            {{ calcRowTotal(scope.row) }}
          </template>
        </el-table-column>
      </el-table>

      <el-table :data="filteredRows" border stripe>
        <el-table-column label="班別" prop="班別"></el-table-column>
        <el-table-column label="工單" prop="工單"></el-table-column>
        <el-table-column label="品名" prop="品名"></el-table-column>
        <el-table-column label="狀態" prop="狀態"></el-table-column>
        <el-table-column label="工單數量" prop="工單數量"></el-table-column>
        <el-table-column label="入庫數量" prop="入庫數量"></el-table-column>
        <el-table-column label="標準工時" prop="標準工時"></el-table-column>
        <el-table-column label="實際工時" prop="實際工時"></el-table-column>

        <el-table-column label="最後更新日" prop="最後更新日"></el-table-column>
        <el-table-column label="逾期天數" prop="逾期天數"></el-table-column>
      </el-table>
    </div>
  `,
  data() {
    return {
      rows: [],
      summary: {},
      selectedFilter: null,
      statusList: ['45', '91', '92', '95']
    };
  },
  computed: {
    filterText() {
      if (!this.selectedFilter) return '';
      const { team, status, overdue } = this.selectedFilter;
      let t = `${team}`;
      if (status) t += ` - 狀態${status}`;
      if (overdue) t += ` - 逾期`;
      return t;
    },
    summaryRows() {
      const result = [];
      const sortedTeams = Object.keys(this.summary).sort();
      for (const team of sortedTeams) {
        const item = { 班別: team };
        this.statusList.forEach(s => {
          item[s] = this.summary[team][s]?.count || 0;
          item['逾期' + s] = this.summary[team][s]?.overdue || 0;
        });
        result.push(item);
      }

      const total = { 班別: '總計' };
      result.forEach(row => {
        Object.keys(row).forEach(k => {
          if (k !== '班別') {
            total[k] = (total[k] || 0) + (row[k] || 0);
          }
        });
      });
      result.push(total);
      return result;
    },
    filteredRows() {
      if (!this.selectedFilter) return this.rows;
      return this.rows.filter(row => {
        const { team, status, overdue } = this.selectedFilter;
        const matchTeam = row.班別 === team;
        const matchStatus = !status || row.狀態 == status;
        const matchOverdue = !overdue || row.逾期天數 > 6;
        return matchTeam && matchStatus && matchOverdue;
      });
    }
  },
  methods: {
    handleCellClick(row, column) {
      if (row.班別 === '總計') return;
      const col = column.property;
      const isOverdue = col.startsWith('逾期');
      const status = isOverdue ? col.replace('逾期', '') : col;

      this.selectedFilter = {
        team: row.班別,
        status: this.statusList.includes(status) ? status : null,
        overdue: isOverdue
      };
    },
    clearFilter() {
      this.selectedFilter = null;
    },
    calcRowTotal(row) {
      return this.statusList.reduce((sum, k) => sum + (row[k] || 0), 0);
    },
    cellClass({ column, row }) {
      if (!this.selectedFilter) return '';
      const prop = column.property;
      const { team, status, overdue } = this.selectedFilter;

      if (row.班別 !== team) return '';
      if (status && !overdue && prop === status) return 'cell-selected';
      if (status && overdue && prop === '逾期' + status) return 'cell-selected';
      if (!status && !overdue && prop === '班別') return 'cell-selected';
      return '';
    },
    drawChart() {
      const ctx = document.getElementById('unclosedChart').getContext('2d');
      const labels = Object.keys(this.summary).sort().filter(t => t !== '總計');
    
      const datasets = [];
    
      this.statusList.forEach((status, i) => {
        datasets.push({
          label: '狀態' + status,
          data: labels.map(t => this.summary[t]?.[status]?.count || 0),
          backgroundColor: 'rgba(33, 150, 243, 0.3)',
          stack: status
        });
        datasets.push({
          label: '逾期' + status,
          data: labels.map(t => this.summary[t]?.[status]?.overdue || 0),
          backgroundColor: ['#ef5350', '#ffa726', '#66bb6a', '#42a5f5'][i],
          stack: status
        });
      });
    
      new Chart(ctx, {
        type: 'bar',
        data: {
          labels,
          datasets
        },
        options: {
          responsive: true,
          plugins: {
            legend: { position: 'top' },
            title: { display: true, text: '總數 vs 逾期數（各班別）' }
          }
        }
      });
    }
    
  },
  mounted() {
    axios.get('https://mms.leapoptical.com:5088/api/WorkOrder/unclosed').then(res => {
      const raw = res.data;
      this.rows = raw;

      const group = {};
      raw.forEach(row => {
        const team = row.班別 || '未填';
        const status = row.狀態;
        const overdue = row.逾期天數 > 6;

        if (!group[team]) group[team] = {};
        if (!group[team][status]) group[team][status] = { count: 0, overdue: 0 };

        group[team][status].count += 1;
        if (overdue) group[team][status].overdue += 1;
      });

      this.summary = group;
      this.$nextTick(this.drawChart);
    });
  }
});