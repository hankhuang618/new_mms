Vue.component('line-leader-view', {
  template: `
    <div>
      <h3>👨‍🏭 產線主管</h3>
      <el-table :data="leaderList" border>
        <el-table-column prop="JOBNUMBER" label="工號"></el-table-column>
        <el-table-column prop="NAME" label="姓名"></el-table-column>
        <el-table-column prop="DEPARTMENT" label="部門"></el-table-column>
      </el-table>
    </div>
  `,
  data() {
    return { leaderList: [] }
  },
  mounted() {
    axios.get(`http://192.168.207.17:5024/api/C_IMPORT/SELECTCDEP?area=VN`)
      .then(res => { this.leaderList = res.data.result })
      .catch(err => console.error(err))
  }
});