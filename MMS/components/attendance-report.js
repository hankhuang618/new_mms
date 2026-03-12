Vue.component('attendance-report-view', {
  template: `
    <div class="attendance-report-page">
      <div class="toolbar">
        <label>廠區:</label>
        <select v-model="selectedArea" @change="fetchSummary">
          <option value="ZH">ZH</option>
          <option value="TC">TC</option>
          <option value="VN">VN</option>
          <option value="TW">TW</option>
        </select>
        <input type="date" v-model="selectedDate" @change="fetchSummary">
      </div>
      <h4 class="report-title">{{ selectedDate }} 考勤日報</h4>

      <table class="table table-bordered table-sm text-center mt-2">
        <thead class="table-dark">
          <tr>
            <th></th>
            <th>部門</th>
            <th>出勤人數</th>
            <th>實際工時</th>
            <th>考勤工時</th>
            <th>回報工時</th>
            <th>差異</th>
          </tr>
        </thead>
        <tbody>
          <template v-for="(item, index) in summaryData" :key="index">
            <tr @click="toggleDetails(index)" style="cursor: pointer;">
              <td><span :class="{'rotate': item.expanded}">▶</span></td>
              <td>{{ item.部門 }}</td>
              <td>{{ item.出勤人數 }}</td>
              <td>{{ item.實際工時 }}</td>
              <td>{{ item.考勤工時 + item.調整工時 }}</td>
              <td>{{ item.考勤工時 + item.調整工時 + item.人工調整 }}</td>
              <td>{{ item.人工調整 }}</td>
            </tr>
            <tr v-if="item.expanded">
              <td colspan="7">
                <table class="table table-bordered table-sm detail-table">
                  <thead>
                    <tr class="table-light">
                      <th>工號</th>
                      <th>姓名</th>
                      <th>班別</th>
                      <th>工時</th>
                      <th>出勤</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr v-for="(detail, dIndex) in item.details" :key="dIndex">
                      <td>{{ detail.工號 }}</td>
                      <td>{{ detail.姓名 }}</td>
                      <td>{{ detail.班別 }}</td>
                      <td>{{ detail.工時 }}</td>
                      <td>{{ detail.出勤 }}</td>
                    </tr>
                  </tbody>
                </table>
              </td>
            </tr>
          </template>
        </tbody>
      </table>
    </div>
  `,
  data() {
    return {
      selectedArea: 'ZH',
      selectedDate: new Date().toISOString().slice(0, 10),
      summaryData: []
    };
  },
  methods: {
    async fetchSummary() {
      if (!this.selectedArea || !this.selectedDate) return;

      const date = new Date(this.selectedDate);
      const formattedDate = `${date.getFullYear()}${(date.getMonth() + 1).toString().padStart(2, '0')}${date.getDate().toString().padStart(2, '0')}`;

      try {
        const response = await fetch(`http://192.168.207.17:5024/api/Report/HR_SUN_version?AREA=${this.selectedArea}&MFG_DAY=${formattedDate}`);
        const data = await response.json();

        this.summaryData = data.map(item => ({
          ...item,
          expanded: false,
          details: [] // 預設為空，展開時再抓
        }));
      } catch (error) {
        console.error('Error loading summary:', error);
      }
    },
    async toggleDetails(index) {
      const item = this.summaryData[index];
      item.expanded = !item.expanded;

      if (item.expanded && item.details.length === 0) {
        try {
          const response = await fetch(`http://192.168.207.17:5024/api/Report/HR_DETAIL?AREA=${this.selectedArea}&MFG_DAY=${this.selectedDate.replace(/-/g, '')}&DEPT=${item.部門}`);
          const detailData = await response.json();
          item.details = detailData;
        } catch (error) {
          console.error('Error loading details:', error);
        }
      }
    }
  },
  mounted() {
    this.fetchSummary();
  }
});
