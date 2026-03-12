// components/attendance-daily.js — 考勤日報（統一生產日報風格）
(function () {
  // ============= 版面樣式（與生產日報一致） =============
  if (!document.getElementById('attendance-daily-style')) {
    const style = document.createElement('style');
    style.id = 'attendance-daily-style';
    style.textContent = `
      .attendance-daily-page .table th,
      .attendance-daily-page .table td { vertical-align: middle; }
      .attendance-daily-page .table-responsive { overflow: auto; }
      .attendance-daily-page .sticky-top { top: 0; z-index: 10; }
      .attendance-daily-page .muted { color: #6c757d; }
      .attendance-daily-page .nowrap { white-space: nowrap; }
      .attendance-daily-page .w-72 { width: 72px; }
      .attendance-daily-page .w-120 { width: 150px; }
      .attendance-daily-page .w-160 { width: 160px; }
    `;
    document.head.appendChild(style);
  }

  // === 視環境調整（建議正式環境用相對路徑 + 反向代理避免 CORS） ===
  const API_HOST = 'https://mms.leapoptical.com:7238';

  const ENDPOINT = {
    HR_SUM: (area, yyyymmdd) =>
      `${API_HOST}/api/Report/HR_SUN_version?AREA=${encodeURIComponent(area)}&MFG_DAY=${encodeURIComponent(yyyymmdd)}`,
    HR_DETAIL: `${API_HOST}/api/Attendance/ATT_DATA_hr`, // POST { area, department, mfG_DAY }
    ADJ:       `${API_HOST}/api/Attendance/ADJTIMES2`,    // POST { Area, MFG_DAY, JOBNUMBER, ADJTIME, RESON }
    RAWDATA:   (yyyymmdd, jobNumber) =>
      `${API_HOST}/api/Attendance/ATT_RAWDATA?MFG_DAY=${yyyymmdd}&JOB_NUMBER=${encodeURIComponent(jobNumber)}`,
    EXPORT:    `${API_HOST}/api/TableToExcel/to-excel`,  // POST blob
  };

  // ===== 工具 =====
  const toYMD = d => {
    const dt = typeof d === 'string' ? new Date(d) : d;
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const day = String(dt.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  };
  const ymdToYYYYMMDD = ymd => ymd.replace(/-/g, '');
  const getYesterdayYMD = () => {
    const d = new Date();
    d.setDate(d.getDate() - 1);
    return toYMD(d);
  };

  // Cookie helpers（記住使用者勾選的部門，依廠區區分）
  const setCookie = (name, value, days) => {
    const d = new Date();
    d.setTime(d.getTime() + days * 24 * 60 * 60 * 1000);
    document.cookie = `${name}=${encodeURIComponent(value)};expires=${d.toUTCString()};path=/`;
  };
  const getCookie = (name) => {
    const nameEQ = name + "=";
    const parts = document.cookie.split(';');
    for (let p of parts) {
      p = p.trim();
      if (p.indexOf(nameEQ) === 0) return decodeURIComponent(p.substring(nameEQ.length));
    }
    return null;
  };
  const COOKIE_DEPTS = (area) => `ATT_DEPTS_${area || 'DEFAULT'}`;

  // 建 payload：支援取消時 ADJTIME = null
  function buildAdjPayload({ area, ymd, jobnumber, adj, reason }) {
    return {
      Area: String(area || ''),
      MFG_DAY: ymdToYYYYMMDD(ymd),                 // "yyyyMMdd"
      JOBNUMBER: String(jobnumber || ''),
      ADJTIME: adj === null ? null : String(adj ?? ''), // ✅ 取消→null；調整→字串
      RESON: String(reason || '')
    };
  }
  // 產生唯一 key（避免不同部門同工號衝突）
  function rowKeyOf(row) { return `${row['部門'] || ''}|${row['工號'] || ''}`; }

  Vue.component('attendance-daily-view', {
    template: `
<div class="container-fluid attendance-daily-page py-3">
  <!-- Toolbar（統一風格） -->
  <div class="d-flex align-items-center gap-2 flex-wrap mb-3">
    <label class="me-2">廠區</label>
    <select v-model="selectedArea" @change="onAreaChange" class="form-select form-select-sm w-auto">
      <option>ZH</option><option>TC</option><option>VN</option><option>TW</option>
    </select>
    <input type="date" v-model="selectedDate" class="form-control form-control-sm w-auto"/>
    <div class="ms-2 small text-muted">{{ selectedDate }} 考勤日報</div>

    <div class="ms-auto d-flex gap-2 align-items-center">
      <button class="btn btn-outline-primary btn-sm" @click="fetchSummary" :disabled="loading">
        <span v-if="loading" class="spinner-border spinner-border-sm me-1"></span>重新載入
      </button>
      <button class="btn btn-outline-secondary btn-sm" @click="selectAll" :disabled="!ATT_time_SUM_data.length">選取全部</button>
      <button class="btn btn-outline-secondary btn-sm" @click="clearAll" :disabled="!hasAnySelected">清除選取</button>
      <button class="btn btn-outline-success btn-sm" @click="exportExcelOriginal" :disabled="exporting || ATT_time_SUM_data.length===0">
        <span v-if="exporting" class="spinner-border spinner-border-sm me-1"></span>匯出 Excel
      </button>
    </div>
  </div>

  <div v-if="error" class="alert alert-danger py-2">{{ error }}</div>
  <div v-if="notice" class="alert alert-success py-2">{{ notice }}</div>

  <!-- 摘要（與生產日報同樣高度/Sticky表頭） -->
  <div class="table-responsive mb-3" style="max-height:45vh;overflow:auto;">
    <table class="table table-bordered table-sm align-middle" id="ATTSUM">
      <thead class="table-light sticky-top">
        <tr>
          <th style="width:40px;">選</th>
          <th>部門</th>
          <th>出勤人數</th>
          <th>實際掛卡工時</th>
          <th>系統計算考勤工時</th>
          <th>回報管理部考勤工時</th>
          <th>人工調整</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(r, idx) in ATT_time_SUM_data" :key="idx" @click="toggleDept(r['部門'])" :class="{'table-primary': showInSecondTable[r['部門']]}" style="cursor:pointer;"
            v-if="Number(r['考勤工時']) + Number(r['調整工時']) !== 0">
          <td class="text-center"><input type="checkbox" :checked="showInSecondTable[r['部門']]" @click.stop @change.stop="toggleDept(r['部門'])"></td>
          <td>{{ formatDepartment(r['部門']) }}</td>
          <td>{{ toNum(r['出勤人數']) }}</td>
          <td>{{ toFixed(r['實際工時'], 2) }}</td>
          <td>{{ toFixed(sysCalc(r), 2) }}</td>
          <td>{{ toFixed(mgmtCalc(r), 2) }}</td>
          <td>{{ toFixed(r['人工調整'], 2) }}</td>
        </tr>

        <!-- 摘要合計列（勾選則為已選，否則全部） -->
        <tr class="table-secondary fw-bold" v-if="ATT_time_SUM_data.length">
          <td colspan="2">合計（{{ hasAnySelected ? '已選部門' : '全部' }}）</td>
          <td>{{ sumSelected('出勤人數') || (hasAnySelected ? sumSelected('出勤人數') : totalAll('出勤人數')) }}</td>
          <td>{{ hasAnySelected ? toFixed(sumSelected('實際工時'),2) : toFixed(totalAll('實際工時'),2) }}</td>
          <td>{{ hasAnySelected ? toFixed(sumSelected('考勤工時') + sumSelected('調整工時'),2) : toFixed(totalAll('考勤工時') + totalAll('調整工時'),2) }}</td>
          <td>{{ hasAnySelected ? toFixed(sumSelected('考勤工時') + sumSelected('調整工時') + sumSelected('人工調整'),2)
                                 : toFixed(totalAll('考勤工時') + totalAll('調整工時') + totalAll('人工調整'),2) }}</td>
          <td>{{ hasAnySelected ? toFixed(sumSelected('人工調整'),2) : toFixed(totalAll('人工調整'),2) }}</td>
        </tr>
      </tbody>
    </table>
  </div>

  <!-- 中間操作列（統一風格） -->
  <div class="d-flex justify-content-end gap-2 mb-2">
    <div class="form-check form-switch me-2">
      <input class="form-check-input" type="checkbox" id="batchSwitch" v-model="batchMode">
      <label class="form-check-label" for="batchSwitch">{{ batchMode ? '批次調整：開啟' : '批次調整：關閉' }}</label>
    </div>
    <button class="btn btn-success btn-sm" @click="submitBatch" :disabled="!batchMode || batchSelectedCount===0 || saving">
      <span v-if="saving">儲存中…</span>
      <span v-else>調整完成（{{ batchSelectedCount }}）</span>
    </button>
  </div>

  <!-- 明細（與生產日報同樣高度/Sticky表頭） -->
  <div class="d-flex align-items-center gap-2 mb-2" v-if="hasAnySelected">
    <span v-if="detailLoading" class="small text-muted"><span class="spinner-border spinner-border-sm me-1"></span>載入明細中…</span>
    <span v-else class="small text-muted">筆數：{{ ATT_time_data.length }}</span>
  </div>

  <div class="table-responsive" style="max-height:55vh;overflow:auto;">
    <table class="table table-bordered table-sm align-middle" id="ATTDATA">
      <thead class="table-light sticky-top">
        <tr>
          <th v-if="batchMode" class="w-72">批次</th>
          <th>部門</th>
          <th>工作拉</th>
          <th>工號</th>
          <th>姓名</th>
          <th>實際掛卡工時</th>
          <th>系統計算考勤工時</th>
          <template v-if="!batchMode">
            <th>調整</th>
          </template>
          <template v-else>
            <th class="w-120">調整值</th>
            <th class="w-160">原因</th>
          </template>
          <th>回報管理部考勤工時</th>
          <th>正班工時</th>
          <th>加班工時</th>
          <th>假別</th>
          <th>事由 / 差異原因</th>
        </tr>
      </thead>
      <tbody>
        <template v-for="(row, idx) in ATT_time_data" :key="idx">
          <!-- 部門小計列（配色與生產日報一致） -->
          <tr v-if="idx === 0 || row['部門'] !== ATT_time_data[idx-1]['部門']" style="background:#e8f4ff">
            <td v-if="batchMode"></td>
            <td>{{ formatDepartment(row['部門']) }}</td>
            <td colspan="3">部門小計</td>
            <td>{{ toFixed(deptTotal('實際工時', row['部門']), 2) }}</td>
            <td>{{ toFixed(deptTotal('調整工時', row['部門']) + deptTotal('考勤工時', row['部門']), 2) }}</td>
            <template v-if="!batchMode"><td>{{ toFixed(deptTotal('手動調整工時', row['部門']), 2) }}</td></template>
            <template v-else><td colspan="2">{{ toFixed(deptTotal('手動調整工時', row['部門']), 2) }}</td></template>
            <td>{{ toFixed(deptTotal('調整工時', row['部門']) + deptTotal('考勤工時', row['部門']) + deptTotal('手動調整工時', row['部門']), 2) }}</td>
            <td>{{ toFixed(excessHours(row['部門']).regular, 2) }}</td>
            <td>{{ toFixed(excessHours(row['部門']).overtime, 2) }}</td>
            <td colspan="2"></td>
          </tr>

          <!-- 資料列 -->
          <tr>
            <td v-if="batchMode">
              <input type="checkbox" v-model="getBatch(row).selected" @change="onBatchSelectChange(row)" />
            </td>
            <td>{{ formatDepartment(row['部門']) }}</td>
            <td>{{ row['排拉'] }}</td>
            <td><a href="javascript:void(0)" @click="fetchAttendanceData(row['工號'], row['姓名'])">{{ row['工號'] }}</a></td>
            <td>{{ row['姓名'] }}</td>
            <td>{{ toFixed(row['實際工時'], 2) }}</td>
            <td>{{ toFixed(row['考勤工時']+row['調整工時'], 2)   }}</td>

            <template v-if="!batchMode">
              <td class="nowrap">
                <template v-if="hasAdj(row)">
                  <a href="javascript:void(0)" class="text-danger" @click="cancelAdjust(row)">取消</a>
                  <div class="text-muted small">（目前：{{ toFixed(toNum(row['手動調整工時']),2) }}）</div>
                </template>
                <template v-else>
                  <a href="javascript:void(0)" @click="openAdjust(row)">調整</a>
                </template>
              </td>
            </template>

            <template v-else>
              <td>
                <el-input-number
                  class="w-120"
                  v-model="getBatch(row).adj"
                  :step="0.5" :min="-24" :max="24" :precision="2"
                  @change="onBatchFieldChange(row)"
                ></el-input-number>
                <div class="small muted" v-if="hasAdj(row)">目前：{{ toFixed(toNum(row['手動調整工時']),2) }}</div>
              </td>
              <td>
                <el-select filterable clearable class="w-160" v-model="getBatch(row).reason" placeholder="原因" @change="onBatchFieldChange(row)">
                  <el-option v-for="r in reasons" :key="r" :label="r" :value="r" />
                </el-select>
              </td>
            </template>

            <td>{{ toFixed(totalForRow(row), 2) }}</td>
            <td>{{ toFixed(calcRegular(row), 2) }}</td>
            <td>{{ toFixed(calcOvertime(row), 2) }}</td>
            <td>{{ row['假別'] || '' }}</td>
            <td>{{ row['事由'] || row['調整原因'] || row['調整理由'] || '' }}</td>
          </tr>
        </template>
      </tbody>
    </table>
  </div>

  <!-- 工時明細對話框 -->
  <el-dialog :title="'工時明細｜' + selectedName + '（' + selectedJobNumber + '）'" :visible.sync="timeDialog.visible" width="720px">
    <div v-if="timeDialog.loading" class="text-center my-3">
      <span class="spinner-border text-primary"></span>
    </div>
    <div v-else class="table-responsive">
      <table class="table table-bordered table-sm text-center align-middle">
        <thead class="table-light">
          <tr>
            <th>放卡時間</th>
            <th>拿卡時間</th>
            <th>工單部門</th>
            <th>工單號</th>
            <th>工時(分鐘)</th>
            <th>工作內容</th>
          </tr>
        </thead>
        <tbody>
          <tr v-for="(r, i) in attendanceData" :key="i">
            <td>{{ r.check_in_time }}</td>
            <td>{{ r.check_out_time }}</td>
            <td>{{ r.wO_DEPARTMENT }}</td>
            <td>{{ r.wO_ID }}</td>
            <td>{{ r.elapsed_time }}</td>
            <td>{{ r.work_content }}</td>
          </tr>
        </tbody>
      </table>
    </div>
    <span slot="footer" class="dialog-footer">
      <el-button @click="timeDialog.visible=false">關閉</el-button>
    </span>
  </el-dialog>

  <!-- 單筆調整對話框（保留） -->
  <el-dialog title="工時調整" :visible.sync="dialog.visible" width="420px">
    <div v-if="dialog.item">
      <div class="mb-2"><strong>工號：</strong>{{ dialog.item['工號'] }}　<strong>姓名：</strong>{{ dialog.item['姓名'] }}</div>
      <el-input-number v-model="dialog.adj" :step="0.5" :min="-24" :max="24" :precision="2" label="調整工時"></el-input-number>
      <div class="mt-3">
        <label class="form-label">原因</label>
        <el-select v-model="dialog.reason" placeholder="請選擇">
          <el-option v-for="r in reasons" :key="r" :label="r" :value="r" />
        </el-select>
      </div>
    </div>
    <span slot="footer" class="dialog-footer">
      <el-button @click="dialog.visible=false">取消</el-button>
      <el-button type="primary" :loading="dialog.saving" @click="saveAdjust">確定</el-button>
    </span>
  </el-dialog>
</div>
    `,
    data() {
      const yesterday = getYesterdayYMD();
      const lastArea = getCookie('MMS_SELECTED_AREA') || 'VN'; // ⬅ 新增：預設取 Cookie
      return {
        // 狀態
        loading: false,
        detailLoading: false,
        exporting: false,
        saving: false,
        error: '',
        notice: '',

        // 篩選
        selectedArea: lastArea,
        selectedDate: yesterday,

        // 摘要與明細
        ATT_time_SUM_data: [],
        ATT_time_data: [],
        showInSecondTable: {},

        // 工時明細彈窗
        timeDialog: { visible: false, loading: false },
        attendanceData: [],
        selectedJobNumber: '',
        selectedName: '',

        // 單筆調整
        dialog: { visible: false, item: null, adj: 0, reason: '', saving: false },
        reasons: ['忘記掛卡/取卡','掛卡異常','網路異常','斷電','自離','新員工','重複掛卡','忘記設定狀態','掛錯卡','狀態設定錯誤','其他'],

        // 批次調整
        batchMode: false,
        // key: "部門|工號" -> { selected, adj, reason }
        batchEditing: {}
      };
    },
    computed: {
      hasAnySelected() { return Object.values(this.showInSecondTable).some(v => v); },
      batchSelectedCount() {
        return Object.values(this.batchEditing).filter(x => x && x.selected).length;
      },
    },
    methods: {
      // ===== 版型一致的操作 =====
      async selectAll(){
        if (!this.ATT_time_SUM_data.length) return;
        const needLoad = [];
        this.ATT_time_SUM_data.forEach(r => {
          if (!this.showInSecondTable[r['部門']]) needLoad.push(r['部門']);
          this.$set(this.showInSecondTable, r['部門'], true);
        });
        for (const d of needLoad) await this.fetchDetails(d);
        this.persistDeptSelection();
      },
      clearAll(){
        this.ATT_time_data = [];
        Object.keys(this.showInSecondTable).forEach(k => this.$set(this.showInSecondTable, k, false));
        this.batchEditing = {};
        this.persistDeptSelection();
      },
      async toggleDept(dep){
        const nv = !this.showInSecondTable[dep];
        this.$set(this.showInSecondTable, dep, nv);
        if (nv) await this.fetchDetails(dep);
        else {
          this.ATT_time_data = this.ATT_time_data.filter(x => x['部門'] !== dep);
          Object.keys(this.batchEditing).forEach(k => { if (k.startsWith(dep + '|')) this.$delete(this.batchEditing, k); });
        }
        this.persistDeptSelection();
      },

      // Cookie（記住每個廠區的勾選部門）
      persistDeptSelection() {
        const selected = Object.keys(this.showInSecondTable).filter(k => this.showInSecondTable[k]);
        setCookie(COOKIE_DEPTS(this.selectedArea), JSON.stringify(selected), 30);
      },
      async applySavedDeptSelection() {
        const raw = getCookie(COOKIE_DEPTS(this.selectedArea));
        if (!raw) return;
        let saved = [];
        try { saved = JSON.parse(raw) || []; } catch {}
        if (!Array.isArray(saved) || !saved.length) return;

        let any = false;
        this.ATT_time_SUM_data.forEach(r => {
          const dep = r['部門'];
          if (saved.includes(dep)) {
            this.$set(this.showInSecondTable, dep, true);
            any = true;
          }
        });
        if (!any) return;

        for (const dep of Object.keys(this.showInSecondTable)) {
          if (this.showInSecondTable[dep]) {
            await this.fetchDetails(dep);
          }
        }
      },

      // ===== 批次調整 =====
      ensureBatchRow(row) {
        const key = rowKeyOf(row);
        if (!this.batchEditing[key]) {
          this.$set(this.batchEditing, key, { selected: false, adj: 0, reason: '' });
        }
        return this.batchEditing[key];
      },
      getBatch(row) { return this.ensureBatchRow(row); },
      onBatchFieldChange(row) {
        const b = this.ensureBatchRow(row);
        const adjNumber = Number(b.adj);
        // 有調整值就自動打勾（≠0 視為有調整）
        if (Number.isFinite(adjNumber) && adjNumber !== 0) {
          b.selected = true;
        }
      },
      onBatchSelectChange(row) { this.ensureBatchRow(row); },

      async submitBatch() {
        const selectedPairs = Object.entries(this.batchEditing).filter(([_, v]) => v && v.selected);
        if (selectedPairs.length === 0) {
          this.$message?.warning?.('請先勾選要調整的資料列');
          return;
        }
        // 基本檢核：批次現在只有「調整」，沒有取消
        for (const [key, v] of selectedPairs) {
          const adjNumber = Number(v.adj);
          if (!Number.isFinite(adjNumber) || adjNumber === 0) {
            this.$message?.warning?.(`選取項目調整值應為非 0 數字（${key}）`);
            return;
          }
          if (!v.reason) {
            this.$message?.warning?.(`選取項目尚未選擇原因（${key}）`);
            return;
          }
        }

        try {
          this.saving = true; this.error=''; this.notice='';
          let success = 0; let fail = 0; const msgs = [];

          for (const [key, v] of selectedPairs) {
            const [dept, job] = key.split('|');
            const row = this.ATT_time_data.find(r => (r['部門'] === dept && r['工號'] === job));
            if (!row) { fail++; msgs.push(`${key}：找不到資料列`); continue; }

            const adjString = Number(v.adj).toFixed(2);
            const reason = v.reason;

            const payload = buildAdjPayload({
              area: this.selectedArea,
              ymd: this.selectedDate,
              jobnumber: job,
              adj: adjString,
              reason
            });

            try {
              const resp = await axios.post(ENDPOINT.ADJ, payload, {
                headers: { 'Content-Type': 'application/json' },
                validateStatus: () => true
              });
              const ok = resp?.data?.success === true;
              const rc = resp?.data?.returnCode ?? resp?.data?.code ?? null;
              if (ok) success++;
              else {
                fail++;
                const serverMsg = resp?.data?.message || '';
                msgs.push(`${key}：失敗${rc != null ? `（代碼 ${rc}）` : ''}${serverMsg ? '｜' + serverMsg : ''}`);
              }
            } catch (e) {
              console.error(e);
              fail++; msgs.push(`${key}：網路或跨域問題`);
            }
          }

          if (fail === 0) { this.notice = `批次調整完成，共 ${success} 筆成功`; }
          else if (success === 0) { this.error = `批次調整全部失敗，共 ${fail} 筆`; }
          else { this.error = `成功 ${success} 筆，失敗 ${fail} 筆：` + msgs.join('；'); }

          await this.fetchSummary();
          for (const dep of Object.keys(this.showInSecondTable)) {
            if (this.showInSecondTable[dep]) await this.fetchDetails(dep);
          }
        } finally {
          this.saving = false;
        }
      },

      // ===== 主要流程 =====
      async fetchSummary() {
        try {
          this.loading = true; this.error=''; this.notice='';
          this.ATT_time_SUM_data = [];
          this.ATT_time_data = [];
          this.showInSecondTable = {};
          this.batchEditing = {};

          const yyyymmdd = ymdToYYYYMMDD(this.selectedDate);
          const url = ENDPOINT.HR_SUM(this.selectedArea, yyyymmdd);
          const { data } = await axios.get(url);

          this.ATT_time_SUM_data = (data || []).map(r => ({
            ...r,
            出勤人數: this.toNum(r['出勤人數']),
            實際工時: this.toNum(r['實際工時']),
            考勤工時: this.toNum(r['考勤工時']),
            調整工時: this.toNum(r['調整工時']),
            人工調整: this.toNum(r['人工調整']),
          }));
          this.ATT_time_SUM_data.forEach(r => { this.$set(this.showInSecondTable, r['部門'], false); });

          await this.applySavedDeptSelection();
        } catch (err) {
          console.error(err);
          this.error = '摘要讀取失敗';
        } finally {
          this.loading = false;
        }
      },

      async fetchDetails(dept) {
        try {
          this.detailLoading = true;
          const yyyymmdd = ymdToYYYYMMDD(this.selectedDate);
          const payload = { area: this.selectedArea, department: dept, mfG_DAY: yyyymmdd };
          const { data } = await axios.post(ENDPOINT.HR_DETAIL, payload);

          const merged = [...this.ATT_time_data, ...(data || [])];
          const map = new Map();
          merged.forEach(r => {
            const key = `${r['部門']}|${r['工號']}`;
            const rawAdj = r['手動調整工時'];
            const adjVal = (rawAdj === null || rawAdj === undefined || rawAdj === '') ? null : Number(rawAdj);
            map.set(key, {
              ...r,
              實際工時: this.toNum(r['實際工時']),
              考勤工時: this.toNum(r['考勤工時']),
              調整工時: this.toNum(r['調整工時']),
              手動調整工時: adjVal
            });
          });
          this.ATT_time_data = Array.from(map.values()).sort((a, b) => {
            if (a['部門'] !== b['部門']) return a['部門'] > b['部門'] ? 1 : -1;
            if ((a['排拉'] || '') !== (b['排拉'] || '')) return (a['排拉'] || '') > (b['排拉'] || '') ? 1 : -1;
            return (a['工號'] || '') > (b['工號'] || '') ? 1 : -1;
          });

          // 確保新資料也有批次物件
          this.ATT_time_data.forEach(r => this.ensureBatchRow(r));
        } catch (err) {
          console.error(err);
          this.error = `明細讀取失敗（${dept}）`;
          this.$set(this.showInSecondTable, dept, false);
        } finally {
          this.detailLoading = false;
        }
      },

      // 工號掛卡明細（點工號）
      async fetchAttendanceData(jobNumber, name) {
        try {
          this.selectedJobNumber = jobNumber;
          this.selectedName = name;
          this.timeDialog.visible = true;
          this.timeDialog.loading = true;

          const yyyymmdd = ymdToYYYYMMDD(this.selectedDate);
          const url = ENDPOINT.RAWDATA(yyyymmdd, jobNumber);
          const { data } = await axios.get(url);

          this.attendanceData = Array.isArray(data) ? data : [];
        } catch (e) {
          console.error(e);
          this.$message?.error?.('工時明細讀取失敗');
          this.attendanceData = [];
        } finally {
          this.timeDialog.loading = false;
        }
      },

      // 單筆模式（保留）
      openAdjust(row) {
        this.dialog.item = row;
        this.dialog.adj = 0;
        this.dialog.reason = '';
        this.dialog.visible = true;
      },
      async saveAdjust() {
        if (!this.dialog.item) return;
        if (!this.dialog.reason) { this.$message?.warning?.('請選擇原因'); return; }

        const adjString = Number.isFinite(this.dialog.adj) ? this.dialog.adj.toFixed(2) : '0.00';
        const dept = this.dialog.item['部門'];
        const job  = this.dialog.item['工號'];

        try {
          this.dialog.saving = true;
          const payload = buildAdjPayload({
            area: this.selectedArea, ymd: this.selectedDate, jobnumber: job,
            adj: adjString, reason: this.dialog.reason
          });
          const resp = await axios.post(ENDPOINT.ADJ, payload, { headers: { 'Content-Type': 'application/json' }, validateStatus: () => true });

          const ok = resp?.data?.success === true;
          const rc = resp?.data?.returnCode ?? resp?.data?.code ?? null;

          if (ok) {
            this.$message?.success?.(this.mapAdjMsg(rc));
            this.dialog.visible = false;
            await this.fetchSummary();
            if (this.showInSecondTable[dept]) await this.fetchDetails(dept);
          } else {
            const serverMsg = resp?.data?.message || '';
            this.$message?.error?.(`調整失敗${rc != null ? `（代碼 ${rc}）` : ''}${serverMsg ? '｜' + serverMsg : ''}`);
          }
        } catch (err) {
          console.error(err);
          this.$message?.error?.('調整失敗（網路或跨域問題）');
        } finally {
          this.dialog.saving = false;
        }
      },
      async cancelAdjust(row) {
        const dept = row['部門'];
        const job  = row['工號'];
        try {
          if (this.$confirm) {
            try { await this.$confirm('確定要取消這筆手動調整嗎？', '提示', { type: 'warning' }); }
            catch { return; }
          }
          const payload = buildAdjPayload({ area: this.selectedArea, ymd: this.selectedDate, jobnumber: job, adj: '', reason: '取消調整' });
          const resp = await axios.post(ENDPOINT.ADJ, payload, { headers: { 'Content-Type': 'application/json' }, validateStatus: () => true });
          const ok = resp?.data?.success === true;
          const rc = resp?.data?.returnCode ?? resp?.data?.code ?? null;
          if (ok) {
            this.$message?.success?.(this.mapAdjMsg(rc, true));
            await this.fetchSummary();
            if (this.showInSecondTable[dept]) await this.fetchDetails(dept);
          } else {
            const serverMsg = resp?.data?.message || '';
            this.$message?.error?.(`取消失敗${rc != null ? `（代碼 ${rc}）` : ''}${serverMsg ? '｜' + serverMsg : ''}`);
          }
        } catch (e) {
          console.error(e);
          this.$message?.error?.('取消失敗（網路或跨域問題）');
        }
      },

      // 匯出（摘要自帶加總；明細每部門插入部門加總）— 維持原 API
      async exportExcelOriginal() {
        try {
          this.exporting = true; this.error=''; this.notice='';
          const hasChecked = Object.values(this.showInSecondTable).some(Boolean);
          const rowsForExport = (this.ATT_time_SUM_data || []).filter(r => {
            const valid = this.toNum(r['考勤工時']) + this.toNum(r['調整工時']) !== 0;
            return hasChecked ? (!!this.showInSecondTable[r['部門']] && valid) : valid;
          });

          const employees = rowsForExport.map(r => {
            const sys  = this.sysCalc(r);
            const mgmt = this.mgmtCalc(r);
            return {
              部門: this.formatDepartment(r['部門']),
              出勤人數: this.toNum(r['出勤人數']),
              實際掛卡工時: Number(this.toNum(r['實際工時']).toFixed(2)),
              系統計算考勤工時: Number(sys.toFixed(2)),
              回報管理部考勤工時: Number(mgmt.toFixed(2)),
              差異: Number((sys - mgmt).toFixed(2)),
              原因: ''
            };
          });

          const sum = (acc, v) => acc + (Number.isFinite(v) ? v : 0);
          const total_出勤 = rowsForExport.reduce((a, r) => sum(a, this.toNum(r['出勤人數'])), 0);
          const total_實際 = rowsForExport.reduce((a, r) => sum(a, this.toNum(r['實際工時'])), 0);
          const total_sys  = rowsForExport.reduce((a, r) => sum(a, this.sysCalc(r)), 0);
          const total_mgmt = rowsForExport.reduce((a, r) => sum(a, this.mgmtCalc(r)), 0);
          employees.push({
            部門: '加總',
            出勤人數: total_出勤,
            實際掛卡工時: Number(total_實際.toFixed(2)),
            系統計算考勤工時: Number(total_sys.toFixed(2)),
            回報管理部考勤工時: Number(total_mgmt.toFixed(2)),
            差異: Number((total_sys - total_mgmt).toFixed(2)),
            原因: ''
          });

          // 明細：只輸出目前已載入的資料；若有勾選則只輸出勾選部門
          const detailHasChecked = Object.values(this.showInSecondTable).some(Boolean);
          const detailRows = (this.ATT_time_data || []).filter(r =>
            detailHasChecked ? !!this.showInSecondTable[r['部門']] : true
          );

          // 依部門分組，先推「部門加總」列，再推逐筆
          const groups = new Map();
          for (const r of detailRows) {
            const d = r['部門'] || '';
            if (!groups.has(d)) groups.set(d, []);
            groups.get(d).push(r);
          }
          const employmentTime = [];
          for (const [dept, list] of groups.entries()) {
            const sumActual   = list.reduce((a, r) => a + this.toNum(r['實際工時']), 0);
            const sumSystem   = list.reduce((a, r) => a + this.toNum(r['考勤工時']), 0);
            const sumManual   = list.reduce((a, r) => a + (r['手動調整工時'] == null ? 0 : this.toNum(r['手動調整工時'])), 0);
            const sumTotal    = list.reduce((a, r) => a + this.totalForRow(r), 0);
            const sumRegular  = list.reduce((a, r) => a + this.calcRegular(r), 0);
            const sumOvertime = list.reduce((a, r) => a + this.calcOvertime(r), 0);

            employmentTime.push({
              部門: this.formatDepartment(dept),
              工作拉: '部門加總',
              工號: '', 姓名: '',
              實際掛卡工時: Number(sumActual.toFixed(2)),
              系統計算考勤工時: Number(sumSystem.toFixed(2)),
              調整: Number(sumManual.toFixed(2)),
              回報管理部考勤工時: Number(sumTotal.toFixed(2)),
              正班工時: Number(sumRegular.toFixed(2)),
              加班工時: Number(sumOvertime.toFixed(2)),
              假別: '', 事由: '', 差異原因: '', 差異原因2: ''
            });
            for (const r of list) {
              const total = this.totalForRow(r);
              const regular = this.calcRegular(r);
              const overtime = this.calcOvertime(r);
              employmentTime.push({
                部門: this.formatDepartment(r['部門']),
                工作拉: r['排拉'] || '',
                工號: r['工號'] || '',
                姓名: r['姓名'] || '',
                實際掛卡工時: Number(this.toNum(r['實際工時']).toFixed(2)),
                系統計算考勤工時: Number(this.toNum(r['考勤工時']).toFixed(2)),
                調整: Number((r['手動調整工時'] == null ? 0 : this.toNum(r['手動調整工時'])).toFixed(2)),
                回報管理部考勤工時: Number(total.toFixed(2)),
                正班工時: Number(regular.toFixed(2)),
                加班工時: Number(overtime.toFixed(2)),
                假別: r['假別'] || '',
                事由: r['事由'] || r['調整原因'] || r['調整理由'] || '',
                差異原因: '',
                差異原因2: ''
              });
            }
          }

          const payload = { MFG_DAY: this.selectedDate, employees, employmentTime };
          const res = await axios.post(ENDPOINT.EXPORT, payload, { responseType: 'blob' });

          const blob = new Blob([res.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          const link = document.createElement('a');
          link.href = window.URL.createObjectURL(blob);
          link.download = `${this.selectedDate} Attendance daily report.xlsx`;
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
        } catch (e) {
          console.error(e);
          this.error = '匯出失敗';
        } finally {
          this.exporting = false;
        }
      },

      // 計算 & 顯示
      sysCalc(r) { return this.toNum(r['考勤工時']) + this.toNum(r['調整工時']); },
      mgmtCalc(r) { return this.sysCalc(r) + this.toNum(r['人工調整']); },
      hasAdj(row) { return row['手動調整工時'] !== null && row['手動調整工時'] !== undefined && this.toNum(row['手動調整工時']) !== 0; },
      totalForRow(row) {
        const manual = row['手動調整工時'] == null ? 0 : this.toNum(row['手動調整工時']);
        return this.toNum(row['考勤工時']) + this.toNum(row['調整工時']) + manual;
      },
      sumSelected(col) {
        return this.ATT_time_SUM_data.reduce((sum, r) => this.showInSecondTable[r['部門']] ? (sum + this.toNum(r[col])) : sum, 0);
      },
      totalAll(col) {
        return this.ATT_time_SUM_data.reduce((sum, r) => sum + this.toNum(r[col]), 0);
      },
      deptTotal(field, dept) {
        return this.ATT_time_data.reduce((sum, r) => r['部門'] === dept ? (sum + this.toNum(r[field])) : sum, 0);
      },
      excessHours(dept) {
        let sum = 0;
        this.ATT_time_data.forEach(r => { if (r['部門'] === dept) sum += this.totalForRow(r); });
        const regular = Math.min(sum, 8 * this.countPeople(dept));
        const overtime = Math.max(0, sum - regular);
        return { regular, overtime };
      },
      countPeople(dept) {
        const set = new Set(this.ATT_time_data.filter(r => r['部門'] === dept).map(r => r['工號']));
        return set.size;
      },
      calcRegular(row) { return Math.min(this.totalForRow(row), 8); },
      calcOvertime(row) { return Math.max(0, this.totalForRow(row) - 8); },
      formatDepartment(dept) { return String(dept || ''); },
      toNum(v) { const n = Number(v); return Number.isFinite(n) ? n : 0; },
      toFixed(v, d) { return this.toNum(v).toFixed(d ?? 0); },

      mapAdjMsg(rc, isCancel = false) {
        if (isCancel) {
          if (rc === 3) return '已取消調整';
          if (rc === 2) return '已取消（原本無資料）';
        }
        switch (rc) {
          case 0: return '調整成功（已更新）';
          case 1: return '調整成功（已新增）';
          case -10: return '失敗：日期格式錯誤';
          default: return rc != null ? `完成（代碼 ${rc}）` : '完成';
        }
      },

      onAreaChange(){
        setCookie('MMS_SELECTED_AREA', this.selectedArea, 30); // ⬅ 新增：紀錄 Cookie
        this.fetchSummary();
      }
    },
    mounted() { this.fetchSummary(); }
  });
})();
