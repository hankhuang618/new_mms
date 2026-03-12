Vue.component('key-staff-pending-view', {
  template: `
    <div>
      <h3>🕒 重點人員待審</h3>
      <el-table :data="pendingList" border stripe>
        <el-table-column prop="ID" label="ID" width="60"></el-table-column>
        <el-table-column prop="AREA" label="廠區" width="80"></el-table-column>
        <el-table-column prop="SHIFT" label="班級" width="120"></el-table-column>
        <el-table-column prop="JOBNUMBER" label="工號" width="120"></el-table-column>
        <el-table-column prop="NAME" label="姓名" width="120"></el-table-column>
        <el-table-column prop="task" label="工作內容"></el-table-column>
        <el-table-column prop="IMPORT_PROC" label="重點工站"></el-table-column>
        <el-table-column prop="rating" label="星級" width="80"></el-table-column>

        <!-- 狀態 -->
        <el-table-column label="狀態" width="120">
          <template slot-scope="scope">
            <el-tag v-if="scope.row.condition === '新增'" type="success">待新增</el-tag>
            <el-tag v-else-if="scope.row.condition === '刪除'" type="danger">待刪除</el-tag>
            <el-tag v-else-if="scope.row.condition === '修改'" type="warning">待修改</el-tag>
            <el-tag v-else>{{ scope.row.condition }}</el-tag>
          </template>
        </el-table-column>

        <!-- 取消 -->
        <el-table-column label="操作" width="100">
          <template slot-scope="scope">
            <el-button type="info" size="mini" @click="cancelPending(scope.row)">取消</el-button>
          </template>
        </el-table-column>
      </el-table>
    </div>
  `,
  data() {
    return { pendingList: [] }
  },
  mounted() {
    this.fetchPending();
  },
  methods: {
    fetchPending() {
      axios.get(`https://mms.leapoptical.com:5088/api/Center/SELECTPENDIMPORT?area=VN`)
        .then(res => { this.pendingList = res.data.result })
        .catch(err => console.error(err));
    },
cancelPending(row) {
  this.$confirm(`確定要取消 ${row.NAME} (${row.JOBNUMBER}) 的 ${row.condition} 嗎？`, "提示", {
    type: "warning"
  }).then(() => {
    axios.post(`https://mms.leapoptical.com:5088/api/Center/deletePENDIMPORT`, {
      id: String(row.ID)   // ⭐ 後端要求是物件 { id: "xxx" }
    }).then(() => {
      this.$message.success("已取消");
      this.fetchPending();
    }).catch(() => {
      this.$message.error("取消失敗");
    });
  }).catch(() => {});
}


  }
});
