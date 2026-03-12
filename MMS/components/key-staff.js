Vue.component('key-staff-view', {
  template: `
    <div>
      <h3>📌 重點人員</h3>
<!-- ⭐ 分析區塊 -->
<div style="margin-bottom: 30px; text-align:center;">
  <h4 style="margin-bottom: 15px; font-size:18px; font-weight:bold; color:#333;">📊 重點工站人數分析</h4>
  <el-table 
    :data="analysisList" 
    style="margin: 0 auto; width: 90%; border-radius: 12px; box-shadow: 0 4px 12px rgba(0,0,0,0.1);"
    :header-cell-style="{
      background:'#409EFF', 
      color:'#fff', 
      fontWeight:'bold',
      whiteSpace:'nowrap'
    }"
    :cell-style="{
      whiteSpace:'nowrap',
      fontSize:'14px'
    }"
    highlight-current-row
    @row-click="onAnalysisRowClick"
    show-summary
    :summary-method="getSummaries"
  >
    <!-- 班級欄位加上可點選樣式 -->
    <el-table-column label="班級" width="120" align="center">
      <template slot-scope="scope">
        <span 
          style="color:#409EFF; font-weight:bold; cursor:pointer; white-space:nowrap;"
        >
          {{ scope.row.班級 }}
        </span>
      </template>
    </el-table-column>

    <el-table-column prop="成型" label="成型" width="100" align="center"
                     :cell-style="{fontWeight:'bold', color:'#606266'}"></el-table-column>
    <el-table-column prop="焊接" label="焊接" width="100" align="center"
                     :cell-style="{fontWeight:'bold', color:'#606266'}"></el-table-column>
    <el-table-column prop="做線" label="做線" width="100" align="center"
                     :cell-style="{fontWeight:'bold', color:'#606266'}"></el-table-column>

    <!-- ⭐ 星級人員 -->
    <el-table-column label="星級人員" width="120" align="center">
      <template slot-scope="scope">
        <span style="font-weight:bold; color:#67C23A;">
          {{ Number(scope.row.成型) + Number(scope.row.焊接) + Number(scope.row.做線) }}
        </span>
      </template>
    </el-table-column>

    <el-table-column prop="其它" label="其它重點人員" width="120" align="center"
                     :cell-style="{fontWeight:'bold', color:'#E6A23C'}"></el-table-column>

    <el-table-column prop="總人數" label="總人數" width="120" align="center"
                     :cell-style="{fontWeight:'bold', color:'#F56C6C'}"></el-table-column>
  </el-table>
</div>



      <!-- 篩選條件 -->
      <div style="margin-bottom: 15px;">
        <label>廠區：</label>
        <el-select v-model="selectedArea" style="width:100px;" @change="fetchStaff">
          <el-option label="VN" value="VN"></el-option>
          <el-option label="TC" value="TC"></el-option>
          <el-option label="ZH" value="ZH"></el-option>
          <el-option label="TW" value="TW"></el-option>
        </el-select>

        <label style="margin-left:10px;">部：</label>
        <el-select v-model="selectedD" style="width:80px;" @change="fetchStaff">
          <el-option v-for="d in [1,2,3]" :key="d" :label="d" :value="d"></el-option>
        </el-select>

        <label style="margin-left:10px;">課：</label>
        <el-select v-model="selectedS" style="width:80px;" @change="fetchStaff">
          <el-option v-for="s in [1,2]" :key="s" :label="s" :value="s"></el-option>
        </el-select>

        <label style="margin-left:10px;">班別：</label>
        <el-select v-model="selectedC" style="width:100px;" @change="fetchStaff">
          <el-option v-for="c in [...Array(11).keys()].map(i=>i+1)" 
                     :key="c" 
                     :label="c" 
                     :value="c"></el-option>
        </el-select>

        <el-button type="success" style="margin-left:20px;" @click="openAddDialog">➕ 新增人員</el-button>
      </div>

      <!-- 資料表 -->
      <el-table :data="staffList" border stripe style="width: 100%;">
        <el-table-column prop="ID" label="ID" width="60"></el-table-column>
        <el-table-column prop="AREA" label="廠區" width="80"></el-table-column>
        <el-table-column prop="DEPARTMENT" label="班級" width="120"></el-table-column>
        <el-table-column prop="JOBNUMBER" label="工號" width="120"></el-table-column>
        <el-table-column prop="NAME" label="姓名" width="120"></el-table-column>
        <el-table-column prop="task" label="工作內容"></el-table-column>
        <el-table-column prop="IMPORT_PROC" label="重點工站"></el-table-column>
        <el-table-column prop="rating" label="星級" width="80"></el-table-column>
        
        <el-table-column label="操作" width="180">
          <template slot-scope="scope">
            <template v-if="getPendingStatus(scope.row)">
              <el-tag v-if="getPendingStatus(scope.row) === '新增'" type="success">待新增</el-tag>
              <el-tag v-else-if="getPendingStatus(scope.row) === '刪除'" type="danger">待刪除</el-tag>
              <el-tag v-else-if="getPendingStatus(scope.row) === '修改'" type="warning">待修改</el-tag>
            </template>
            <template v-else>
              <el-button size="mini" type="primary" @click="editRow(scope.row)">修改</el-button>
              <el-button size="mini" type="danger" @click="deleteRow(scope.row)">刪除</el-button>
            </template>
          </template>
        </el-table-column>
      </el-table>

      <!-- 新增人員對話框 -->
      <el-dialog title="新增人員" :visible.sync="showAddDialog" width="400px">
        <el-form :model="newStaff" label-width="100px">
          <el-form-item label="班別">
            <el-input v-model="currentDepartment" disabled></el-input>
          </el-form-item>
          <el-form-item label="工號">
            <el-select v-model="newStaff.JOBNUMBER" placeholder="選擇工號" style="width:100%;" @change="onJobNumberChange">
              <el-option v-for="job in jobNumberOptions" 
                         :key="job.jobnumber" 
                         :label="job.jobnumber" 
                         :value="job.jobnumber"></el-option>
            </el-select>
          </el-form-item>
          <el-form-item label="姓名">
            <el-input v-model="newStaff.NAME" disabled></el-input>
          </el-form-item>
          <el-form-item label="工作內容">
            <el-input v-model="newStaff.task"></el-input>
          </el-form-item>
          <el-form-item label="重點工站">
            <el-select v-model="newStaff.IMPORT_PROC" style="width:100%;">
              <el-option v-for="opt in ['成型','焊接','做線','其它']" 
                         :key="opt" :label="opt" :value="opt"></el-option>
            </el-select>
          </el-form-item>
          <el-form-item label="星級">
            <el-select v-model="newStaff.rating" style="width:100%;">
              <el-option v-for="n in [0,1,2,3,4]" :key="n" :label="n" :value="n"></el-option>
            </el-select>
          </el-form-item>
        </el-form>
        <div slot="footer" class="dialog-footer">
          <el-button @click="showAddDialog=false">取消</el-button>
          <el-button type="primary" @click="addStaff">新增</el-button>
        </div>
      </el-dialog>
    </div>
  `,
  data() {
    return {

      analysisList: [],  // ⭐ 新增分析資料
      staffList: [],
      pendingList: [],   // ⭐ 待審清單
      selectedArea: 'VN',
      selectedD: '',
      selectedS: '',
      selectedC: '',
      showAddDialog: false,
      jobNumberOptions: [],  // 工號選項
      newStaff: { JOBNUMBER: '', NAME: '', task: '', IMPORT_PROC: '成型', rating: 0 }
    }
  },
  computed: {
    currentDepartment() {
      if (!this.selectedD || !this.selectedS || !this.selectedC) return '';
      return `${this.selectedD}D${this.selectedS}S${this.selectedC}C`;
    }
  },
  methods: {
      onAnalysisRowClick(row) {
    const match = row.班級.match(/(\d+)D(\d+)S(\d+)C/);
    if (match) {
      this.selectedD = match[1];
      this.selectedS = match[2];
      this.selectedC = match[3];
      this.fetchStaff();
      this.$message.success(`已切換到 ${row.班級}`);
    }
  },
  getSummaries(param) {
    const { columns, data } = param;
    const sums = [];

    columns.forEach((column, index) => {
      if (index === 0) {
        sums[index] = '總計';
        return;
      }
      if (['成型','焊接','做線','其它','總人數'].includes(column.property)) {
        const total = data.reduce((sum, item) => {
          const val = Number(item[column.property]);
          return !isNaN(val) ? sum + val : sum;
        }, 0);
        sums[index] = total;
      }
      if (column.label === '星級人員') {
        const total = data.reduce((sum, item) => {
          const v1 = Number(item['成型']) || 0;
          const v2 = Number(item['焊接']) || 0;
          const v3 = Number(item['做線']) || 0;
          return sum + v1 + v2 + v3;
        }, 0);
        sums[index] = total;
      }
    });

    return sums;
  },
      fetchStaff1() {
              // ⭐ 分析清單
      axios.get(`https://mms.leapoptical.com:5088/api/Center/keyStaffAnalysis?AREA=VN`, {
        params: { area: this.selectedArea }
      }).then(res => {
        this.analysisList = res.data;
      });
      },
  fetchStaff() {
    if (!this.selectedD || !this.selectedS || !this.selectedC) return;
    const department = this.currentDepartment;

    // 1. 抓主清單
    axios.get(`https://mms.leapoptical.com:5088/api/Center/SELECTIMPORT`, {
      params: { area: this.selectedArea, department: department }
    }).then(res => {
      this.staffList = res.data.result;
      
      // 2. 抓待審清單
      axios.get(`https://mms.leapoptical.com:5088/api/Center/SELECTPENDIMPORT`, {
        params: { area: this.selectedArea }
      }).then(pendRes => {
        this.pendingList = pendRes.data.result;

        // 把「待新增」補到 staffList
        const pendingAdds = this.pendingList.filter(p => 
          p.condition === "新增" && p.SHIFT === department
        ).map(p => ({
          ID: p.ID,
          AREA: p.AREA,
          DEPARTMENT: p.SHIFT,
          JOBNUMBER: p.JOBNUMBER,
          NAME: p.NAME,
          task: p.task,
          IMPORT_PROC: p.IMPORT_PROC,
          rating: p.rating,
          isPending: true,        // ⭐ 標記為待審
          condition: p.condition
        }));

        // 合併
        this.staffList = [...this.staffList, ...pendingAdds];
      });
    });

  },
    // 判斷是否在 pendingList
  getPendingStatus(row) {
    const pending = this.pendingList.find(
      p => p.JOBNUMBER === row.JOBNUMBER && p.SHIFT === row.DEPARTMENT
    );
    return pending ? pending.condition : null;
  },
    openAddDialog() {
      if (!this.currentDepartment) {
        this.$message.warning("請先選擇部 / 課 / 班別");
        return;
      }
      this.showAddDialog = true;
      this.newStaff = { JOBNUMBER: '', NAME: '', task: '', IMPORT_PROC: '成型', rating: 0 };

      // 工號清單 (過濾掉已存在 + 待審)
      axios.post('https://mms.leapoptical.com:5088/api/Center/SELECT', {
        area: this.selectedArea,
        department: this.currentDepartment
      }).then(res => {
        const allJobs = res.data;
        const existing = this.staffList.map(s => s.JOBNUMBER);
        const pending = this.pendingList.map(p => p.JOBNUMBER);
        this.jobNumberOptions = allJobs.filter(j => !existing.includes(j.jobnumber) && !pending.includes(j.jobnumber));
      }).catch(err => {
        console.error(err);
        this.$message.error("無法讀取工號清單");
      });
    },
    onJobNumberChange(val) {
      const emp = this.jobNumberOptions.find(e => e.jobnumber === val);
      if (emp) {
        this.newStaff.NAME = emp.username;
      }
    },
    addStaff() {
      if (!this.currentDepartment) {
        this.$message.warning("請先選擇部 / 課 / 班別");
        return;
      }
      const staffData = {
        area: this.selectedArea,
        shift: this.currentDepartment,
        jobNumber: this.newStaff.JOBNUMBER,
        name: this.newStaff.NAME,
        task: this.newStaff.task,
        importProc: this.newStaff.IMPORT_PROC,
        rating: String(this.newStaff.rating),
        condition: "新增"
      };
      axios.post('https://mms.leapoptical.com:5088/api/Center/insertpendingimport', staffData)
        .then(() => {
          this.$message.success("新增已送審");
          this.showAddDialog = false;
          this.fetchStaff();
        })
        .catch(() => this.$message.error("新增失敗"));
    },
    editRow(row) {
      this.$confirm(
        `
        <div style="text-align:left;">
          <label>工作內容：</label>
          <input id="editTask" value="${row.task}" style="width:95%;margin:5px 0;"/>

          <label>重點工站：</label>
          <select id="editProc" style="width:95%;margin:5px 0;">
            <option value="成型" ${row.IMPORT_PROC === '成型' ? 'selected' : ''}>成型</option>
            <option value="焊接" ${row.IMPORT_PROC === '焊接' ? 'selected' : ''}>焊接</option>
            <option value="做線" ${row.IMPORT_PROC === '做線' ? 'selected' : ''}>做線</option>
            <option value="其它" ${row.IMPORT_PROC === '其它' ? 'selected' : ''}>其它</option>
          </select>

          <label>星級：</label>
          <select id="editRating" style="width:95%;margin:5px 0;">
            ${[0,1,2,3,4].map(n => 
              `<option value="${n}" ${String(row.rating) === String(n) ? 'selected' : ''}>${n}</option>`
            ).join('')}
          </select>
        </div>
        `,
        "修改人員資料",
        {
          dangerouslyUseHTMLString: true,
          confirmButtonText: '送出',
          cancelButtonText: '取消'
        }
      ).then(() => {
        const newTask   = document.getElementById("editTask").value;
        const newProc   = document.getElementById("editProc").value;
        const newRating = document.getElementById("editRating").value;

        const staffData = {
          area: row.AREA,
          shift: row.DEPARTMENT,
          jobNumber: row.JOBNUMBER,
          name: row.NAME,
          task: newTask,
          importProc: newProc,
          rating: String(newRating),
          condition: "修改"
        };

        axios.post('https://mms.leapoptical.com:5088/api/Center/insertpendingimport', staffData)
          .then(() => {
            this.$message.success("修改已送審");
            this.fetchStaff();
          })
          .catch(() => this.$message.error("修改失敗"));
      }).catch(() => {});
    },
    deleteRow(row) {
      this.$confirm(`確定要刪除 ${row.NAME} (${row.JOBNUMBER})?`, "提示", { type: "warning" })
        .then(() => {
          const staffData = {
            area: row.AREA,
            shift: row.DEPARTMENT,
            jobNumber: row.JOBNUMBER,
            name: row.NAME,
            task: row.task,
            importProc: row.IMPORT_PROC,
            rating: String(row.rating),
            condition: '刪除'
          };
          axios.post('https://mms.leapoptical.com:5088/api/Center/insertpendingimport', staffData)
            .then(() => {
              this.$message.success("刪除已送審");
              this.fetchStaff();
            })
            .catch(() => this.$message.error("刪除失敗"));
        })
        .catch(() => {});
    },
     // 點選分析表的班級 → 自動帶入選單
    onAnalysisRowClick(row) {
      const match = row.班級.match(/(\d+)D(\d+)S(\d+)C/);
      if (match) {
        this.selectedD = match[1];
        this.selectedS = match[2];
        this.selectedC = match[3];
        this.fetchStaff();
        this.$message.success(`已切換到 ${row.班級}`);
      }
    }
  },
  mounted() {
    this.fetchStaff();
        this.fetchStaff1();
  }
});
