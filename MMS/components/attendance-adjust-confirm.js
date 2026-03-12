
// MMS - 作業中心 / 考勤調整確認（修正版）
// 變更：
// - adj_time 改為純顯示（不再可編輯）
// - 修正 selection：加 ref / @selection-change / 點整列可切換
// - 日期預設昨日期

(function () {
  const API_BASE = window.API_BASE || 'https://mms.leapoptical.com:5088';
  const RAW_BASE = 'https://mms.leapoptical.com:7238';

  function getYesterdayISO() {
    const d = new Date();
    d.setDate(d.getDate() - 1);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }

  Vue.component('attendance-adjust-confirm-view', {
    template: `
    <div class="p-3">
      <div class="d-flex align-items-end gap-2 flex-wrap">
        <div>
          <label class="form-label me-2">{{ $t('common.select') }} {{ $t('menu.ops') }} Area</label>
          <el-select v-model="selectedArea" size="small" style="min-width:120px" @change="fetchList">
            <el-option v-for="a in allowedAreas" :key="a" :label="a" :value="a" />
          </el-select>
        </div>
        <div>
          <label class="form-label me-2">{{ $t('menu.attendanceDaily') }}</label>
          <el-date-picker v-model="selectedDate" type="date" size="small"
                          value-format="yyyy-MM-dd" @change="fetchList"/>
        </div>
        <div class="ms-auto">
          <el-button type="primary" size="small"
                     :loading="loading"
                     :disabled="loading || selectedRows.length === 0"
                     @click="submitSelected">
            {{ $t('common.save') }}
          </el-button>
        </div>
      </div>

      <el-card class="mt-3">
        <div class="d-flex align-items-center justify-content-between mb-2">
          <div class="fw-bold">{{ $t('menu.attAdjustConfirm') }}</div>
          <div class="text-muted" v-if="list.length">
            {{ $t('common.select') }}: {{ selectedRows.length }} / {{ list.length }}
          </div>
        </div>

        <el-table ref="tbl"
                  :data="list"
                  height="60vh"
                  border
                  v-loading="loading"
                  @selection-change="onSelectionChange"
                  @row-click="onRowClick">
          <el-table-column type="selection" width="50" />
          <el-table-column prop="mfG_DAY" label="日期" width="120">
            <template slot-scope="{row}">{{ formatDate8(row.mfG_DAY) }}</template>
          </el-table-column>
          <el-table-column prop="area" label="Area" width="80" />
          <el-table-column prop="department" label="部門" width="140" />
          <el-table-column prop="jobnumber" label="工號" width="120" />
          <el-table-column prop="username" label="姓名" width="140" />
          <el-table-column prop="calc_time" label="系統計算考勤工時" width="160">
            <template slot-scope="{row}">{{ n2(row.calc_time) }}</template>
          </el-table-column>
          <el-table-column prop="adj_time" label="調整" width="100">
            <!-- 不可調整，純顯示 -->
            <template slot-scope="{row}">{{ n2(row.adj_time) }}</template>
          </el-table-column>
          <el-table-column prop="reason" label="事由" min-width="220">
            <!-- 不可調整，純顯示 -->
            <template slot-scope="{row}">{{ row.reason || '' }}</template>
          </el-table-column>
        </el-table>

        <div class="d-flex justify-content-end mt-2">
          <el-button size="small" @click="clearSelection">{{ $t('common.cancel') }}</el-button>
          <el-button size="small" type="primary" :disabled="!list.length" @click="selectAll">全選</el-button>
        </div>
      </el-card>
    </div>
    `,
    data() {
      return {
        loading: false,
        list: [],
        selectedRows: [],
        selectedArea: (window.getAllowedAreas && window.getAllowedAreas()[0]) || 'ZH',
        selectedDate: getYesterdayISO(), // 預設昨天
      };
    },
    computed: {
      allowedAreas() {
        return (window.getAllowedAreas && window.getAllowedAreas()) || ['ZH','TC','VN','TW'];
      }
    },
    mounted() {
      this.fetchList();
    },
    methods: {
      n2(v){ const n = Number(v||0); return isNaN(n)?0:n.toFixed(2); },
      yyyymmdd(dateStr) {
        if (!dateStr) return '';
        const [y,m,d] = dateStr.split('-');
        return `${y}${m}${d}`;
      },
      formatDate8(yyyymmdd) {
        const s = String(yyyymmdd||'');
        if (s.length!==8) return s;
        return `${s.slice(0,4)}-${s.slice(4,6)}-${s.slice(6,8)}`;
      },
      onSelectionChange(val){
        this.selectedRows = val || [];
      },
      onRowClick(row){
        // 點整列切換選取
        const $t = this.$refs.tbl;
        if (!$t) return;
        const selected = this.selectedRows.includes(row);
        $t.toggleRowSelection(row, !selected);
      },
      selectAll(){
        const $t = this.$refs.tbl;
        if (!$t) return;
        // 若未全選 -> 全選；若已全選 -> 清空
        if (this.selectedRows.length === this.list.length) {
          this.clearSelection();
        } else {
          this.list.forEach(r => $t.toggleRowSelection(r, true));
        }
      },
      clearSelection(){
        const $t = this.$refs.tbl;
        if ($t) $t.clearSelection();
      },
      async fetchList(){
        this.loading = true;
        this.list = [];
        this.selectedRows = [];
        try {
          const ymd = this.yyyymmdd(this.selectedDate);
          if (!this.selectedArea || !ymd) { this.loading=false; return; }
          const url = `${RAW_BASE}/api/Attendance/HR_ADJ_CONFIRM?AREA=${encodeURIComponent(this.selectedArea)}&mfg_day=${encodeURIComponent(ymd)}`;
          const { data } = await axios.get(url);

          this.list = (Array.isArray(data)? data : []).map(r => ({
            mfG_DAY: r.mfG_DAY || r.mfg_day || r.MFG_DAY,
            area: r.area || r.AREA,
            department: r.department || r.DEPARTMENT,
            jobnumber: r.jobnumber || r.JOB_NUMBER || r.job_no,
            username: r.username || r.USERNAME,
            calc_time: r.calc_time ?? r.STATUS_TIME ?? r.att_time ?? 0,
            adj_time: r.adj_time ?? 0,
            reason: r.reason || r.NOTE || ''
          }));
        } catch (err) {
          console.error('fetch HR_ADJ_CONFIRM failed', err);
          this.$message.error('讀取失敗');
        } finally {
          this.loading = false;
        }
      },
      async submitSelected(){
        if (!this.selectedRows.length) {
          this.$message.warning('請先勾選資料');
          return;
        }
        this.loading = true;
        try {
          for (const item of this.selectedRows) {
            const payload = {
              mfG_DAY: item.mfG_DAY,
              area: item.area,
              department: item.department,
              jobnumber: item.jobnumber,
              ADJTIME: item.adj_time, // 沿用原值，不再編輯
              floW_ID: 'NFC'
            };
            await axios.post(`${RAW_BASE}/api/Attendance/UPDATE_ADJS`, payload, {
              headers: { 'Content-Type': 'application/json' }
            });
          }
          this.$message.success('更新完成');
          await this.fetchList();
        } catch (err) {
          console.error('UPDATE_ADJS failed', err);
          this.$message.error('送出失敗');
        } finally {
          this.loading = false;
        }
      }
    }
  });
})();
