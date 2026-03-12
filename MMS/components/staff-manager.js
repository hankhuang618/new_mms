/**
 * MMS - 人員管理（管理課）
 * 美化版：假別彈窗、狀態純顯示、出勤(T/F)、每列編修
 * ※ 已移除頂端「修改 / 完成」批次編修功能
 * 後端 API 沿用 NFC（可用 window.API_BASE 覆寫）
 */
(function () {
  const API_BASE = window.API_BASE || 'https://mms.leapoptical.com:7238';

  // 假別設定與 UI 規則
  const LEAVE_CONFIG = {
    '年休': { needHours: true, needNote: true },
    '請假': { needHours: false, forceHours: '0', needNote: true },
    '哺乳假': { needHours: true, needNote: true },
    '工傷': { needHours: true, needNote: true },
    '工傷陪護': { needHours: true, needNote: true },
    '借出': { needHours: true, needNote: true, notePreset: ['廠內重工工時', '其它廠重工工時', '其他借出工時'] },
    '夜班': { needHours: false, forceHours: '0', needNote: false },
    '備註': { needHours: false, forceHours: '0', needNote: true },
    '新進員工': { needHours: false, forceHours: '0', needNote: false },
  };

  // 假別 → 標籤樣式
  const LEAVE_COLOR = {
    '年休': 'success',
    '請假': 'warning',
    '哺乳假': 'warning',
    '工傷': 'danger',
    '工傷陪護': 'danger',
    '借出': 'info',
    '夜班': '',
    '備註': 'info',
    '新進員工': 'success'
  };

  Vue.component('staff-manager-view', {
    template: `
      <div class="container py-3 staff-manager-page">
        <div class="d-flex align-items-center justify-content-between mb-2">
          <h4 class="mb-0">人員管理</h4>
          <div class="text-muted small">管理課 / MMS</div>
        </div>

        <!-- 篩選條：置頂吸附 -->
        <el-card shadow="never" class="mb-3 sticky-top" style="top: 64px;">
          <div class="row g-2 align-items-end">
            <div class="col-auto">
              <label class="form-label mb-1">廠區</label>
              <el-select v-model="selectedArea" placeholder="選擇" style="width:120px" @change="onFilterChange">
                <el-option label="ZH" value="ZH"></el-option>
                <el-option label="TC" value="TC"></el-option>
                <el-option label="VN" value="VN"></el-option>
                <el-option label="TW" value="TW"></el-option>
              </el-select>
            </div>
            <div class="col-auto">
              <label class="form-label mb-1">部 (D)</label>
              <el-select v-model="selectedD" style="width:90px" @change="onFilterChange">
                <el-option v-for="d in ['1','2','3','4','6']" :key="d" :label="d" :value="d"></el-option>
              </el-select>
            </div>
            <div class="col-auto">
              <label class="form-label mb-1">課 (S)</label>
              <el-select v-model="selectedS" style="width:90px" @change="onFilterChange">
                <el-option v-for="s in ['1','2','3']" :key="s" :label="s" :value="s"></el-option>
              </el-select>
            </div>
            <div class="col-auto">
              <label class="form-label mb-1">班 (C)</label>
              <el-select v-model="selectedC" style="width:120px" @change="onFilterChange">
                <el-option v-for="c in classOptions" :key="c" :label="c" :value="c"></el-option>
              </el-select>
            </div>

            <div class="col-auto">
              <div class="text-muted small">部門代碼</div>
              <div class="fw-bold">{{ selectedDepartment }}</div>
            </div>

            <div class="col-auto ms-auto">
              <el-button type="primary" size="small" @click="new_staff">新增</el-button>
            </div>
          </div>

          <!-- 新增表單 -->
          <div v-if="showForm" class="mt-3">
            <el-form :model="newStaff" :rules="formRules" ref="newForm" label-width="90px">
              <div class="row g-2">
                <div class="col-12 col-md-3">
                  <el-form-item label="部門" prop="department">
                    <el-input v-model="newStaff.department" placeholder="例如 2D1S1C"></el-input>
                  </el-form-item>
                </div>
                <div class="col-12 col-md-3">
                  <el-form-item label="工號" prop="jobnumber">
                    <el-input v-model="newStaff.jobnumber"></el-input>
                  </el-form-item>
                </div>
                <div class="col-12 col-md-3">
                  <el-form-item label="姓名" prop="username">
                    <el-input v-model="newStaff.username"></el-input>
                  </el-form-item>
                </div>
                <div class="col-12 col-md-3">
                  <el-form-item label="工作拉">
                    <el-input v-model="newStaff.work_department" placeholder="可留空"></el-input>
                  </el-form-item>
                </div>
              </div>
              <el-button type="success" @click="submitNew">完成</el-button>
              <el-button @click="showForm=false">關閉</el-button>
            </el-form>
          </div>
        </el-card>

        <el-card shadow="never">
          <el-table
            :data="staffData"
            v-loading="loading"
            height="62vh"
            size="small"
            border
          >
            <el-table-column label="#" type="index" width="56" fixed="left" align="center"></el-table-column>

            <!-- 狀態（純顯示：在線/離線，不可點） -->
            <el-table-column label="狀態" width="100" fixed="left" align="center">
              <template slot-scope="{row}">
                <el-tag :type="row.online==='01' ? 'success' : 'info'" :disable-transitions="true">
                  {{ row.online==='01' ? '在線' : '離線' }}
                </el-tag>
              </template>
            </el-table-column>

            <!-- 出勤（沿用 NFC：show=T/F，PUT /api/Attendance/UPDATE_SHOW） -->
            <el-table-column label="出勤" width="120" align="center">
              <template slot-scope="{row}">
                <el-switch
                  v-model="row.show"
                  :active-value="'T'"
                  :inactive-value="'F'"
                  active-text="出勤"
                  @change="onToggleAttend(row)"
                ></el-switch>
              </template>
            </el-table-column>

            <el-table-column label="姓名" min-width="140">
              <template slot-scope="{row}">
                <template v-if="row._editable">
                  <el-input v-model="row.username"></el-input>
                </template>
                <template v-else>{{ row.username }}</template>
              </template>
            </el-table-column>

            <el-table-column label="工號" width="140">
              <template slot-scope="{row}">
                <template v-if="row._editable">
                  <el-input v-model="row.jobnumber"></el-input>
                </template>
                <template v-else>{{ row.jobnumber }}</template>
              </template>
            </el-table-column>

            <el-table-column label="部門" width="120">
              <template slot-scope="{row}">
                <template v-if="row._editable">
                  <el-input v-model="row.department"></el-input>
                </template>
                <template v-else>{{ row.department }}</template>
              </template>
            </el-table-column>

            <el-table-column label="工作拉" width="140">
              <template slot-scope="{row}">
                <template v-if="row._editable">
                  <el-input v-model="row.work_department"></el-input>
                </template>
                <template v-else>{{ row.work_department }}</template>
              </template>
            </el-table-column>

            <!-- 假別 / 人事狀態 -->
            <el-table-column label="假別狀態" min-width="180" align="center">
              <template slot-scope="{row}">
                <template v-if="row.status">
                  <el-tag :type="leaveTagType(row.status)">{{ row.status }}</el-tag>
                  <el-button type="text" class="ms-1" @click="askCancelLeave(row)">取消</el-button>
                </template>
                <template v-else>
                  <span class="text-muted">—</span>
                </template>
              </template>
            </el-table-column>

            <el-table-column label="外派部門" min-width="140" prop="out_department"></el-table-column>

            <!-- 功能列：每列編輯 / 假別 / 刪除 -->
            <el-table-column label="功能" width="260" fixed="right" align="center">
              <template slot-scope="{row}">
                <el-button
                  v-if="row._editable"
                  type="primary" size="mini"
                  @click="updateRow(row)"
                >儲存</el-button>
                <el-button
                  v-else
                  type="default" size="mini"
                  @click="row._editable = true"
                >編輯</el-button>

                <el-button size="mini" type="success" plain @click="openLeaveDialog(row)">假別</el-button>
                <el-button size="mini" type="danger" plain @click="askDelete(row)">刪除</el-button>
              </template>
            </el-table-column>
          </el-table>

          <div class="text-muted small mt-2 d-flex align-items-center">
            <span class="me-3">提示：出勤開關 = NFC「出勤」欄位（T/F）；假別用彈窗設定。</span>
            <el-tag type="success" class="ms-2">年休</el-tag>
            <el-tag type="warning" class="ms-1">請假/哺乳假</el-tag>
            <el-tag type="danger" class="ms-1">工傷/陪護</el-tag>
            <el-tag type="info" class="ms-1">借出/備註</el-tag>
          </div>
        </el-card>

        <!-- 假別彈窗 -->
        <el-dialog
          title="設定假別"
          :visible.sync="leaveDialog.visible"
          width="520px"
          :close-on-click-modal="false"
        >
          <el-form :model="leaveDialog.form" :rules="leaveRules" ref="leaveForm" label-width="96px">
            <el-form-item label="姓名 / 工號">
              <div>{{ leaveDialog.form.username }} / {{ leaveDialog.form.jobnumber }}</div>
            </el-form-item>

            <el-form-item label="假別" prop="leaveType">
              <el-select v-model="leaveDialog.form.leaveType" placeholder="選擇假別" style="width: 260px">
                <el-option v-for="(cfg, key) in LEAVE_CONFIG" :key="key" :label="key" :value="key"></el-option>
              </el-select>
            </el-form-item>

            <el-form-item v-if="currentLeave.needHours" label="有效工時" prop="hr_time">
              <el-input v-model="leaveDialog.form.hr_time" placeholder="輸入數字（小時）" @input="numberOnly('hr_time')"></el-input>
            </el-form-item>

            <el-form-item v-if="currentLeave.notePreset" label="常用理由">
              <el-select v-model="leaveDialog.form.notePreset" placeholder="選擇一項" style="width: 260px">
                <el-option v-for="item in currentLeave.notePreset" :key="item" :label="item" :value="item"></el-option>
              </el-select>
            </el-form-item>

            <el-form-item v-if="currentLeave.needNote || leaveDialog.form.notePreset" :label="currentLeave.notePreset ? '補充說明' : '理由'" prop="note">
              <el-input
                v-model="leaveDialog.form.note"
                :placeholder="currentLeave.notePreset ? '可留空；若選擇常用理由則會一併上傳' : '請輸入理由'"
              ></el-input>
            </el-form-item>
          </el-form>

          <span slot="footer" class="dialog-footer">
            <el-button @click="leaveDialog.visible=false">取消</el-button>
            <el-button type="primary" @click="submitLeave">確定</el-button>
          </span>
        </el-dialog>

        <!-- 自製確認對話框（刪除/取消假別用） -->
        <el-dialog
          :title="confirmBox.title"
          :visible.sync="confirmBox.visible"
          width="420px"
          :close-on-click-modal="false"
        >
          <div class="mb-2">{{ confirmBox.message }}</div>
          <span slot="footer" class="dialog-footer">
            <el-button @click="onConfirmCancel">取消</el-button>
            <el-button type="primary" @click="onConfirmOk">確定</el-button>
          </span>
        </el-dialog>
      </div>
    `,

    data() {
      return {
        loading: false,
        selectedArea: 'VN',
        selectedD: '2',
        selectedS: '1',
        selectedC: '1',
        classOptions: ['1','2','3','4','5','6','7','8','9','10','S1','S2','S3','S4','S5','S6'],
        selectedDepartment: '',
        staffData: [],
        showForm: false,
        newStaff: {
          department: '',
          jobnumber: '',
          username: '',
          work_department: ''
        },
        formRules: {
          department: [{ required: true, message: '必填', trigger: 'blur' }],
          jobnumber: [{ required: true, message: '必填', trigger: 'blur' }],
          username: [{ required: true, message: '必填', trigger: 'blur' }]
        },

        // 假別彈窗資料
        leaveDialog: {
          visible: false,
          form: {
            id: null,
            username: '',
            jobnumber: '',
            leaveType: '',
            hr_time: '',
            notePreset: '',
            note: ''
          },
          row: null
        },

        // 確認盒
        confirmBox: {
          visible: false,
          title: '確認',
          message: '',
          _resolve: null,
          _reject: null,
          payload: null
        },

        LEAVE_CONFIG
      };
    },

    computed: {
      currentLeave() {
        const cfg = this.LEAVE_CONFIG[this.leaveDialog.form.leaveType] || {};
        return cfg;
      },
      leaveRules() {
        const r = {};
        if (this.leaveDialog.form.leaveType) {
          const cfg = this.currentLeave;
          if (cfg.needHours) {
            r.hr_time = [{ required: true, message: '請輸入有效工時', trigger: 'blur' }];
          }
          if (cfg.needNote) {
            r.note = [{ required: true, message: '請輸入理由', trigger: 'blur' }];
          }
        }
        r.leaveType = [{ required: true, message: '請選擇假別', trigger: 'change' }];
        return r;
      }
    },

    watch: {
      'leaveDialog.form.leaveType'(val) {
        const cfg = this.currentLeave;
        if (cfg.forceHours != null) this.leaveDialog.form.hr_time = cfg.forceHours;
        if (!cfg.needHours && !cfg.forceHours) this.leaveDialog.form.hr_time = '';
        if (!cfg.needNote) this.leaveDialog.form.note = '';
        this.leaveDialog.form.notePreset = '';
      }
    },

    created() {
      this.computeDepartment();
      this.fetchStaff();
    },

    methods: {
      computeDepartment() {
        this.selectedDepartment = `${this.selectedD}D${this.selectedS}S${this.selectedC}C`;
      },
      onFilterChange() {
        this.computeDepartment();
        this.fetchStaff();
      },

      // 取得清單
      fetchStaff() {
        if (!this.selectedArea || !this.selectedD || !this.selectedS || !this.selectedC) return;
        this.loading = true;
        axios.post(`${API_BASE}/api/Attendance`, {
          area: this.selectedArea,
          department: this.selectedDepartment
        })
        .then(res => {
          const list = Array.isArray(res.data) ? res.data : [];
          this.staffData = list.map(x => Object.assign({_editable:false}, x));
        })
        .catch(err => {
          console.error('fetchStaff error', err);
          this.staffData = [];
        })
        .finally(() => this.loading = false);
      },

      // 新增
      new_staff() {
        this.newStaff = {
          department: this.selectedDepartment,
          jobnumber: '',
          username: '',
          work_department: ''
        };
        this.showForm = true;
        this.$nextTick(() => this.$refs.newForm && this.$refs.newForm.clearValidate());
      },
      submitNew() {
        this.$refs.newForm.validate(valid => {
          if (!valid) return;
          axios.post(`${API_BASE}/api/Attendance/creat`, {
            area: this.selectedArea,
            department: this.newStaff.department,
            jobnumber: this.newStaff.jobnumber,
            username: this.newStaff.username,
            work_department: this.newStaff.work_department || ''
          })
          .then(() => {
            this.$message.success('新增完成');
            this.showForm = false;
            this.fetchStaff();
          })
          .catch(err => {
            console.error('addNewStaff error', err);
            this.$message.error('新增失敗');
          });
        });
      },

      // 單列更新
      updateRow(row) {
        const payload = {
          id: String(row.id),
          area: this.selectedArea,
          department: row.department,
          jobnumber: row.jobnumber,
          username: row.username,
          work_department: row.work_department,
          online: row.online
        };
        axios.put(`${API_BASE}/api/Attendance/UPDATE_USER`, payload, {
          headers: { 'Content-Type':'application/json' }
        })
        .then(() => {
          this.$message.success('已儲存');
          row._editable = false;
          this.fetchStaff();
        })
        .catch(err => {
          console.error('UPDATE_staffData error', err);
          this.$message.error('更新失敗');
        });
      },

      // 出勤切換（沿用 NFC：show=T/F → /api/Attendance/UPDATE_SHOW）
      onToggleAttend(row) {
        const before = row.show === 'T' ? 'F' : 'T'; // 用於失敗回滾
        axios.put(`${API_BASE}/api/Attendance/UPDATE_SHOW`, {
          id: parseInt(row.id, 10),
          area: this.selectedArea,
          show: row.show
        }, { headers: { 'Content-Type': 'application/json' }})
        .then(() => {
          this.$message.success('出勤狀態已更新');
        })
        .catch(err => {
          console.error('UPDATE_SHOW error', err);
          this.$message.error('更新失敗，已還原');
          row.show = before;
        });
      },

      // 假別：開啟彈窗
      openLeaveDialog(row) {
        this.leaveDialog.row = row;
        this.leaveDialog.form = {
          id: row.id,
          username: row.username,
          jobnumber: row.jobnumber,
          leaveType: row.status || '',
          hr_time: row.status ? (row.hr_time || '') : '',
          notePreset: '',
          note: ''
        };
        this.leaveDialog.visible = true;
        this.$nextTick(() => this.$refs.leaveForm && this.$refs.leaveForm.clearValidate());
      },

      // 假別：送出
      submitLeave() {
        this.$refs.leaveForm.validate(valid => {
          if (!valid) return;

          const cfg = this.currentLeave;
          let note = this.leaveDialog.form.note || '';
          if (cfg.notePreset && this.leaveDialog.form.notePreset) {
            note = this.leaveDialog.form.notePreset + (note ? `；${note}` : '');
          }
          const hr = cfg.forceHours != null ? cfg.forceHours : (cfg.needHours ? (this.leaveDialog.form.hr_time || '0') : '0');

          axios.post(`${API_BASE}/api/Attendance/UPDATESTATUS`, {
            area: this.selectedArea,
            jobnumber: this.leaveDialog.form.jobnumber,
            ststus: this.leaveDialog.form.leaveType,
            hr_time: hr,
            note
          })
          .then(() => {
            this.$message.success('假別已更新');
            this.leaveDialog.visible = false;
            this.fetchStaff();
          })
          .catch(err => {
            console.error('submitLeave error', err);
            this.$message.error('送出失敗');
          });
        });
      },

      // 取消假別（等同 Delete_hr）
      askCancelLeave(row) {
        this.confirm(`要取消「${row.username}」目前的【${row.status}】？`, '取消假別', { row })
          .then(({ row }) => this.doCancelLeave(row))
          .catch(() => {});
      },
      doCancelLeave(row) {
        const id = parseInt(row.id, 10);
        axios.delete(`${API_BASE}/api/Attendance/Delete_hr`, {
          data: { id, area: this.selectedArea },
          headers: { 'Content-Type': 'application/json' }
        })
        .then(() => {
          this.$message.success('已取消假別');
          this.fetchStaff();
        })
        .catch(err => {
          console.error('delete_hr error', err);
          this.$message.error('取消失敗');
        });
      },

      // 刪除資料列
      askDelete(row) {
        this.confirm(`確定刪除「${row.username} / ${row.jobnumber}」這筆人員？`, '刪除人員', { row })
          .then(({ row }) => this.doDelete(row))
          .catch(() => {});
      },
      doDelete(row) {
        const id = parseInt(row.id, 10);
        axios.delete(`${API_BASE}/api/Attendance/Delete`, {
          data: { id, area: this.selectedArea },
          headers: { 'Content-Type': 'application/json' }
        })
        .then(() => {
          this.$message.success('已刪除');
          this.fetchStaff();
        })
        .catch(err => {
          console.error('delete_staffData error', err);
          this.$message.error('刪除失敗');
        });
      },

      // ===== 自製確認框（刪除/取消假別用） =====
      confirm(message, title='確認', payload=null) {
        return new Promise((resolve, reject) => {
          this.confirmBox.title = title;
          this.confirmBox.message = message;
          this.confirmBox.payload = payload;
          this.confirmBox.visible = true;
          this.confirmBox._resolve = resolve;
          this.confirmBox._reject = reject;
        });
      },
      onConfirmOk() {
        this.confirmBox.visible = false;
        const res = this.confirmBox._resolve; const payload = this.confirmBox.payload;
        this.confirmBox._resolve = null; this.confirmBox._reject = null; this.confirmBox.payload = null;
        res && res(payload);
      },
      onConfirmCancel() {
        this.confirmBox.visible = false;
        const rej = this.confirmBox._reject;
        this.confirmBox._resolve = null; this.confirmBox._reject = null; this.confirmBox.payload = null;
        rej && rej('cancel');
      },

      // 輸入限制
      numberOnly(field) {
        const v = this.leaveDialog.form[field];
        this.leaveDialog.form[field] = String(v).replace(/[^\d.]/g, '');
      },

      // 標籤樣式
      leaveTagType(status) {
        return LEAVE_COLOR[status] || '';
      }
    }
  });
})();
