Vue.component('flow-form-view', {
    template: `
    <div class="flow-form-page">
      <h4>未結 FLOW 表單</h4>
      
<div class="mt-1">
  <el-tag type="info" class="big-tag">表單總數: {{ forms.length }}</el-tag>
  <el-tag type="danger" class="big-tag">大於2天未處理表單: {{ overdueForms }}</el-tag>
</div>
  
      <el-table :data="forms" class="mt-4" stripe>
        <el-table-column prop="表單名稱" label="表單名稱"></el-table-column>
        <el-table-column label="文件編號">
          <template slot-scope="scope">
            <a :href="scope.row.文件編號連結" target="_blank">{{ scope.row.文件編號 }}</a>
          </template>
        </el-table-column>
        <el-table-column prop="簽核者" label="簽核者"></el-table-column>
        <el-table-column prop="表單狀態" label="狀態"></el-table-column>
        <el-table-column prop="簽和關卡" label="關卡"></el-table-column>
        <el-table-column prop="申請人" label="申請人"></el-table-column>
        <el-table-column prop="建立日期" label="建立日期"></el-table-column>
<el-table-column label="前一次簽核日期">
  <template slot-scope="scope">
    <div :style="highlightOverdue(scope)">
      {{ scope.row['前一次簽核日期'] }}
    </div>
  </template>
</el-table-column>


    </div>
    `,
    data() {
      return {
        forms: [],
        activeTab: 'status'
      };
    },

    mounted() {
      axios.get("https://mms.leapoptical.com:5088/api/FlowForm/unclosed").then(res => {
        this.forms = res.data;
      });
    },
    methods: {
       countByStatus(status) {
    return this.forms.filter(f => f.表單狀態 === status).length;
  },
highlightOverdue(scope) {
  const raw = scope.row['前一次簽核日期'];
  if (!raw) return {};

  const parsedDate = new Date(raw.replace(/-/g, '/').trim());
  if (isNaN(parsedDate.getTime())) return {};

  const today = new Date();
  const sevenDaysAgo = new Date(today.getFullYear(), today.getMonth(), today.getDate() - 2);

  if (scope.row['表單狀態'] !== '已完成' && parsedDate < sevenDaysAgo) {
    return {
      background: '#ffe6e6',
      color: '#b30000',
      fontWeight: 'bold'
    };
  }

  return {};
}


    },
    computed: {
  overdueForms() {
    const today = new Date();
    return this.forms.filter(f => {
      const prevDate = new Date(f['前一次簽核日期']);
      return f.表單狀態 !== '已完成' && prevDate < today.setDate(today.getDate() - 2);
    }).length;
  }
}
  });
  