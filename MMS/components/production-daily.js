// components/production-daily.js — 生產日報（MMS）
(function () {
  const API_HOST = 'https://mms.leapoptical.com:7238';
  const ZERO_SUM_SCOPE = 'global';   // 'global' | 'perDept'
  const COOKIE_DAYS = 30;

  const API = {
    WO_SUM: (area, ymd) => `${API_HOST}/api/Report/WO_SUM_version?AREA=${encodeURIComponent(area)}&MFG_DAY=${encodeURIComponent(ymd)}`,
    WO_VER: (area, dept, ymd) => `${API_HOST}/api/Report/WO_version?AREA=${encodeURIComponent(area)}&DEPARTMENT=${encodeURIComponent(dept)}&MFG_DAY=${encodeURIComponent(ymd)}`,
    WO_EXPORT: () => `${API_HOST}/api/WOTableToExcel/to-excel`,
    ADG_WO_TIME: () => `${API_HOST}/api/Report/ADG_WO_TIME`
  };

  // ===== Utils =====
  function toYMD(d) {
    if (typeof d === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(d)) return d.replace(/-/g, '');
    const dt = d ? new Date(d) : new Date();
    const yyyy = dt.getFullYear();
    const mm = String(dt.getMonth() + 1).padStart(2, '0');
    const dd = String(dt.getDate()).padStart(2, '0');
    return `${yyyy}${mm}${dd}`;
  }
  function yesterdayISO() {
    const dt = new Date();
    dt.setDate(dt.getDate() - 1);
    const yyyy = dt.getFullYear();
    const mm = String(dt.getMonth() + 1).padStart(2, '0');
    const dd = String(dt.getDate()).padStart(2, '0');
    return `${yyyy}-${mm}-${dd}`;
  }
function keyOf(row) {
  return [
    row.部門 || '',
    row.拉 || '',
    row.工單 || '',
    row.工序 || '',
    row.品名 || '',
    row.數量 ?? '',
    row.標準工時 ?? '',
    row.執行工時 ?? ''
  ].join('|');
}

  function setCookie(name, value, days = COOKIE_DAYS) {
    const d = new Date();
    d.setTime(d.getTime() + days * 86400000);
    document.cookie = `${name}=${encodeURIComponent(value)};expires=${d.toUTCString()};path=/`;
  }
  function getCookie(name) {
    const key = name + '=';
    const arr = document.cookie.split(';');
    for (let c of arr) {
      c = c.trim();
      if (c.indexOf(key) === 0) return decodeURIComponent(c.substring(key.length));
    }
    return '';
  }
  function cookieKey(area) { return `PD_SELECTED_${area || 'ALL'}`; }

  // ===== Vue Component =====
  Vue.component('production-daily-view', {
    template: `
<div class="container-fluid production-daily-page py-3">
  <!-- Toolbar -->
  <div class="d-flex align-items-center gap-2 flex-wrap mb-3">
<label class="me-2">區域</label>
<select
  v-model="area"
  @change="onAreaChange"
  class="form-select form-select-sm w-auto"
  :disabled="areaOptions.length === 1"
>
  <option v-for="opt in areaOptions" :key="opt" :value="opt">{{ opt }}</option>
</select>


    <input type="date" v-model="date" @change="fetchSummary" class="form-control form-control-sm w-auto"/>
    <div class="ms-2 small text-muted">{{ date }} 生產日報</div>

    <div class="ms-auto d-flex gap-2 align-items-center">
      <button class="btn btn-outline-primary btn-sm" @click="fetchSummary" :disabled="loading">
        <span v-if="loading" class="spinner-border spinner-border-sm me-1"></span>重新載入
      </button>
      <button class="btn btn-outline-secondary btn-sm" @click="selectAll" :disabled="!summary.length">選取全部</button>
      <button class="btn btn-outline-secondary btn-sm" @click="clearAll" :disabled="!anySelected">清除選取</button>
    </div>
  </div>

  <div v-if="error" class="alert alert-danger py-2">{{ error }}</div>
  <div v-if="notice" class="alert alert-success py-2">{{ notice }}</div>

  <!-- 摘要 -->
  <div class="table-responsive mb-2" style="max-height:45vh; overflow:auto;">
    <table class="table table-bordered table-sm align-middle">
      <thead class="table-light sticky-top" style="white-space:nowrap;">
        <tr>
          <th style="width: 40px;">選</th>
          <th>部門</th>
          <th>回報工時</th>
          <th>執行工時</th>
          <th>有薪假</th>
          <th>有薪工時</th>
          <th>貼紙班</th>
          <th>殘疾人</th>
          <th>無效工時</th>
          <th>未設定工單</th>
          <th>借入工時</th>
          <th>借出工時</th>
          <th>調整工時</th>
          <th>備註</th>
          <th>系統計算差異</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(item, idx) in summary" :key="'sv-'+idx"
            @click="toggleDept(item.部門)"
            :class="{'table-primary': isSelected(item.部門)}"
            style="cursor:pointer;">
          <td class="text-center">
            <input type="checkbox" :checked="isSelected(item.部門)" @click.stop @change.stop="toggleDept(item.部門)">
          </td>
          <td>{{ item.部門 }}</td>
          <td>{{ pretty(item.回報工時) }}</td>
          <td>{{ pretty(item.執行工時) }}</td>
          <td>{{ pretty(item.有薪假) }}</td>
          <td>{{ pretty(item.有薪工時) }}</td>
          <td>{{ pretty(item.貼紙班) }}</td>
          <td>{{ pretty(item.殘疾人) }}</td>
          <td>{{ pretty(item.無效工時) }}</td>
          <td>{{ pretty(item.未設定工單) }}</td>
          <td>{{ pretty(item.借入工時) }}</td>
          <td>{{ pretty(item.借出工時) }}</td>
          <td>{{ displayAdjustPretty(item.部門, item.調整工時) }}</td>
          <td class="remark-cell" style="text-align:left; white-space:normal !important; word-break:break-word;">
            <div v-for="(txt,i) in remarksLines(item)" :key="'rmk-'+i">{{ txt }}</div>
          </td>
          <td :class="{'text-danger': isNonZero(diffNumWithAdjust(item))}">{{ pretty(diffNumWithAdjust(item)) }}</td>
        </tr>

        <!-- 合計列（依已選/全部切換範圍） -->
        <tr class="table-secondary fw-bold">
          <td colspan="2">合計（{{ anySelected ? '已選部門' : '全部' }}）</td>
          <td>{{ sumPretty('回報工時') }}</td>
          <td>{{ sumPretty('執行工時') }}</td>
          <td>{{ sumPretty('有薪假') }}</td>
          <td>{{ sumPretty('有薪工時') }}</td>
          <td>{{ sumPretty('貼紙班') }}</td>
          <td>{{ sumPretty('殘疾人') }}</td>
          <td>{{ sumPretty('無效工時') }}</td>
          <td>{{ sumPretty('未設定工單') }}</td>
          <td>{{ sumPretty('借入工時') }}</td>
          <td>{{ sumPretty('借出工時') }}</td>
          <td>{{ sumAdjustedDisplay() }}</td>
          <td></td>
          <td>{{ sumDiffAdjustedDisplay() }}</td>
        </tr>
      </tbody>
    </table>
  </div>

  <!-- 中間操作列（介於 摘要 與 明細） -->
  <div class="d-flex justify-content-end gap-2 mb-2">

<!-- 新增工單按鈕（靠左） -->

  <button class="btn btn-primary btn-sm" @click="showAddWO = true">
    <i class="bi bi-plus-circle"></i> 新增工單
  </button>




  <!-- 新增工單彈窗 -->
  <el-dialog title="新增工單" :visible.sync="showAddWO" width="420px">
    <el-form :model="newWO" label-width="100px">
      <el-form-item label="班別" required>
        <el-select v-model="newWO.dept" placeholder="選擇班別">
          <el-option v-for="d in summary" :key="d.部門" :label="d.部門" :value="d.部門"></el-option>
        </el-select>
      </el-form-item>
      <el-form-item label="工單號" required>
        <el-input v-model="newWO.woid" placeholder="請輸入工單號"></el-input>
      </el-form-item>
    </el-form>
    <span slot="footer" class="dialog-footer">
      <el-button @click="showAddWO=false">取消</el-button>
      <el-button type="primary" :loading="addingWO" @click="submitAddWO">確定</el-button>
    </span>
  </el-dialog>



    <button class="btn btn-outline-warning btn-sm" @click="toggleEdit" :disabled="!anySelected">
      {{ editing ? '結束調整' : '調整' }}
    </button>
    <button class="btn btn-success btn-sm" @click="saveAdjustments" :disabled="!editing || changedCount===0 || saving || !zeroSumOK">
      <span v-if="saving">儲存中…</span>
      <span v-else>儲存調整（{{ changedCount }}）</span>
    </button>
    <span v-if="editing && changedCount>0" class="small align-self-center" :class="{'text-danger': !zeroSumOK, 'text-muted': zeroSumOK}">
      {{ zeroSumHint }}
    </span>
    <button class="btn btn-outline-success btn-sm" @click="exportExcel" :disabled="exporting">匯出 Excel</button>
  </div>

  <!-- 明細 -->
  <div class="d-flex align-items-center gap-2 mb-2" v-if="anySelected">
    <span v-if="detailLoading" class="small text-muted">
      <span class="spinner-border spinner-border-sm me-1"></span>載入明細中…
    </span>
    <span v-else class="small text-muted">總筆數：{{ details.length }}</span>
  </div>

  <div class="table-responsive" style="max-height:55vh; overflow:auto;">
    <table class="table table-bordered table-sm align-middle" id="WO2">
      <thead class="table-light sticky-top" style="white-space:nowrap;">
        <tr>
          <th>部門</th>
          <th>拉</th>
          <th>工單</th>
          <th>工序</th>
          <th>品名</th>
          <th>數量</th>
          <th>標準工時</th>
          <th>執行工時</th>
          <th>品檢工時</th>
          <th>物料員工時</th>          
          <th>調整工時</th>
          <th v-if="editing">調整(新增)</th>
          <th>備註</th>
        </tr>
      </thead>
      <tbody>
        <template v-for="dept in selectedDeptList">
          <tr v-if="grouped[dept] && grouped[dept].length" :key="'sum-'+dept" style="background:#e8f4ff">
            <td>{{ dept }}</td>
            <td></td>
            <td colspan="4">部門小計</td>
            <td>{{ pretty(deptTotal(dept,'標準工時')) }}</td>
            <td>{{ pretty(deptTotal(dept,'執行工時')) }}</td>
            <td>{{ pretty(deptTotal(dept,'品檢工時')) }}</td>
            <td>{{ pretty(deptTotal(dept,'物料員工時')) }}</td>
            <td>{{ displayAdjustPretty(dept, summaryAdjustOf(dept)) }}</td>
            <td v-if="editing">{{ pretty(deltaTotalDept(dept)) }}</td>
            <td></td>
          </tr>
          <tr v-for="(row, i) in grouped[dept]" :key="'d-'+dept+'-'+i">
            <td>{{ row.部門 }}</td>
            <td>{{ row.拉 }}</td>
            <td>{{ row.工單 || '未設定工單' }}</td>
            <td>{{ row.工序 || '-' }}</td>
            <td>{{ row.品名 }}</td>
            <td>{{ prettyInt(row.數量) }}</td>
            <td>{{ pretty(row.標準工時) }}</td>
            <td>{{ pretty(row.執行工時) }}</td>
            <td>{{ pretty(row.品檢工時) }}</td>
            <td>{{ pretty(row.物料員工時) }}</td>
            <td>{{ pretty(row.調整工時) }}</td>
            <td v-if="editing">
              <input class="form-control form-control-sm" type="number" step="0.01"
                     :value="deltaValue(row)" @input="onDeltaInput($event,row)" placeholder="0"/>
            </td>
            <td>{{ row.備註 }}</td>
          </tr>
        </template>
      </tbody>
    </table>
  </div>
</div>
    `,
    data() {
       const lastArea = getCookie('MMS_SELECTED_AREA') || 'VN'; // ⬅ 新增：預設取 Cookie
      return {
        area: lastArea,
        date: yesterdayISO(), // 預設昨天
        
        loading: false,
        detailLoading: false,
        exporting: false,
        saving: false,
        error: '',
        notice: '',
        summary: [],
        details: [],
        selected: {},          // { 部門: true/false }
        editing: false,
        deptAdjustMap: {},     // { 部門: number } 以明細「調整工時」合計覆蓋摘要
        deltaMap: {},           // { key: number }   本次「調整(新增)」
        showAddWO: false,
          addingWO: false,
        newWO: { dept: '', woid: '' },
       areaOptions: ['ZH','TC','VN','TW'],
       areaLocked: false,
      };
    },
created() {
  // 1) 取得可用廠區（由主頁提供帳號限制）
  this.areaOptions = (window.getAllowedAreas && window.getAllowedAreas()) || ['ZH','TC','VN','TW'];

  // 2) 只有一個可選就強制使用它
  if (this.areaOptions.length === 1) {
    this.area = this.areaOptions[0];
  } else {
    // 3) 多個可選時才讀 Cookie，帶回上次的廠區
    try {
      const last = (document.cookie.split(';').map(s=>s.trim()).find(s=>s.startsWith(LAST_AREA_COOKIE+'=')) || '')
                    .split('=').slice(1).join('=');
      const lastArea = decodeURIComponent(last || '');
      if (lastArea && this.areaOptions.includes(lastArea)) {
        this.area = lastArea;
      }
    } catch(e) { /* ignore */ }
  }

  // 4) 監聽登入變更：若切換帳號導致 allowedAreas 改變，調整目前廠區並避免不合法值
  window.addEventListener('mms-user-changed', () => {
    this.areaOptions = (window.getAllowedAreas && window.getAllowedAreas()) || ['ZH','TC','VN','TW'];
    if (!this.areaOptions.includes(this.area)) {
      this.area = this.areaOptions[0];
      // 帳號限制變動時順便更新 Cookie
      document.cookie = `${LAST_AREA_COOKIE}=${encodeURIComponent(this.area)};path=/;max-age=${30*86400}`;
    }
  });
},

    computed: {
      anySelected() { return Object.values(this.selected).some(Boolean); },
      selectedDeptList() { return Object.keys(this.selected).filter(k => this.selected[k]); },
      selectedCount() { return this.selectedDeptList.length; },
      summaryScope() {
        if (!this.anySelected) return this.summary;
        const set = new Set(this.selectedDeptList);
        return this.summary.filter(s => set.has(s.部門));
      },
      grouped() {
        const map = {};
        for (const r of this.details) { if (!map[r.部門]) map[r.部門] = []; map[r.部門].push(r); }
        return map;
      },
      changedCount() {
        let n = 0;
        for (const r of this.details) { const k = keyOf(r); const v = this.deltaMap[k] || 0; if (Math.abs(parseFloat(v)) > 0.0005) n++; }
        return n;
      },
      changedRows() {
        return this.details.filter(r => Math.abs(this.deltaMap[keyOf(r)] || 0) > 0.0005);
      },
      deltaSumScope() {
        return this.changedRows.reduce((t, r) => t + (this.deltaMap[keyOf(r)] || 0), 0);
      },
      zeroSumOK() {
        if (this.changedRows.length === 0) return true;
        if (ZERO_SUM_SCOPE === 'perDept') {
          const sums = {};
          for (const r of this.changedRows) { const d = r.部門 || ''; sums[d] = (sums[d] || 0) + (this.deltaMap[keyOf(r)] || 0); }
          return Object.values(sums).every(v => Math.abs(v) <= 0.005);
        } else {
          return Math.abs(this.deltaSumScope) <= 0.005;
        }
      },
      zeroSumHint() {
        if (this.changedRows.length === 0) return '';
        if (ZERO_SUM_SCOPE === 'perDept') {
          const sums = {};
          for (const r of this.changedRows) { const d = r.部門 || ''; sums[d] = (sums[d] || 0) + (this.deltaMap[keyOf(r)] || 0); }
          const bad = Object.entries(sums).filter(([_, v]) => Math.abs(v) > 0.005);
          return bad.length ? ('部門合計需為0：' + bad.map(([k, v]) => `${k}(${(+v).toFixed(2)})`).join('、')) : '';
        } else {
          return `合計：${(+this.deltaSumScope).toFixed(2)}（必須為 0）`;
        }
      }
    },
    methods: {
      async submitAddWO() {
  if (!this.newWO.dept || !this.newWO.woid) {
    this.$message.warning('請輸入班別與工單號');
    return;
  }

  try {
    this.addingWO = true;
    const payload = {
      area: this.area,
      mfgday: this.date.replace(/-/g, ''), // 例如 "20251112"
      WO_DEPARTMENT: this.newWO.dept,
      woid: this.newWO.woid
    };

    await axios.post(`${API_HOST}/api/Report/InsertH_DEPWO_DATA`, payload, {
      headers: { 'Content-Type': 'application/json' }
    });

    this.$message.success('新增工單成功');
    this.showAddWO = false;
    this.newWO = { dept: '', woid: '' };
    this.fetchSummary(); // 重新整理摘要表
  } catch (err) {
    console.error(err);
    this.$message.error('新增工單失敗');
  } finally {
    this.addingWO = false;
  }
},

      // 數值/格式（0 顯示空白）
      num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; },
      isNonZero(v) { return Math.abs(this.num(v)) > 0.0005; },
      pretty(v, d = 2) { const n = this.num(v); return this.isNonZero(n) ? n.toFixed(d) : ''; },
      prettyInt(v) { const n = this.num(v); return this.isNonZero(n) ? n.toFixed(0) : ''; },

      // 備註：轉成逐行顯示（優先 1→3→2→4；若都空，拆分「備註」）
      remarksLines(row){
        const lines = [row.備註]
          .map(v => (v == null ? '' : String(v).trim()))
          .filter(Boolean);
        if (lines.length) return lines;
        const raw = row.備註 == null ? '' : String(row.備註);
        const parts = raw.split(/\s*(?:[,，、]|\r?\n)+\s*/).filter(Boolean);
        return parts;
      },
      joinRemarks(r){ return this.remarksLines(r).join('\n'); },

      // Cookie
      persistSelected() { try { setCookie(cookieKey(this.area), JSON.stringify(this.selectedDeptList), COOKIE_DAYS); } catch (e) {} },
      restoreSelected() { try { const raw = getCookie(cookieKey(this.area)); if (!raw) return []; const arr = JSON.parse(raw); return Array.isArray(arr) ? arr : []; } catch (e) { return []; } },

      // 選取
      isSelected(dep) { return !!this.selected[dep]; },
      async toggleDept(dep) {
        if (!dep) return;
        const newVal = !this.selected[dep];
        this.$set(this.selected, dep, newVal);
        this.persistSelected();
        if (newVal) { await this.fetchDeptDetails(dep); }
        else { this.removeDeptDetails(dep); }
      },
      async selectAll() {
        if (!this.summary.length) return;
        const needFetch = []; const map = {};
        this.summary.forEach(r => { map[r.部門] = true; if (!this.selected[r.部門]) needFetch.push(r.部門); });
        this.selected = map; this.persistSelected();
        for (const dep of needFetch) { await this.fetchDeptDetails(dep); }
      },
      clearAll() { this.selected = {}; this.persistSelected(); this.details = []; this.deptAdjustMap = {}; this.deltaMap = {}; },

      // 顯示/合計
      displayAdjust(dept, fallback) {
        const v = this.deptAdjustMap[dept];
        const n = (typeof v === 'number') ? v : this.num(fallback);
        return n;
      },
      displayAdjustPretty(dept, fallback) {
        const n = this.displayAdjust(dept, fallback);
        return this.pretty(n);
      },
      summaryAdjustOf(dept) {
        const f = this.summary.find(s => s.部門 === dept);
        return f ? this.num(f.調整工時) : 0;
      },
      summaryQCOf(dept) {
        const f = this.summary.find(s => s.部門 === dept);
        return f ? this.num(f.品檢工時) : 0;
      },
      recalcDeptAdjust(dept) {
        const s = this.details.filter(x => x.部門 === dept).reduce((t, x) => t + this.num(x.調整工時), 0);
        this.$set(this.deptAdjustMap, dept, parseFloat(s.toFixed(2)));
      },
      deltaTotalDept(dept) {
        let t = 0; for (const r of (this.grouped[dept] || [])) { t += (this.deltaMap[keyOf(r)] || 0); } return t;
      },

      // 差異（不含調整工時；含有薪假/有薪工時）
      diffNumWithAdjust(r) {
        return this.num(r.回報工時) - (
          this.num(r.執行工時) +
          this.num(r.有薪假) + this.num(r.有薪工時) +
          this.num(r.未設定工單) + this.num(r.無效工時) +
          this.num(r.借出工時) - this.num(r.借入工時) +
          this.num(r.殘疾人)
        );
      },
      sumVal(field) { return this.summaryScope.reduce((tot, r) => tot + this.num(r[field]), 0); },
      sumPretty(field) { return this.pretty(this.sumVal(field)); },
      sumAdjustedDisplay() {
        const s = this.summaryScope.reduce((t, r) => t + this.displayAdjust(r.部門, r.調整工時), 0);
        return this.pretty(s);
      },
      sumDiffAdjustedDisplay() {
        const v = this.summaryScope.reduce((t, r) => t + this.diffNumWithAdjust(r), 0);
        return this.pretty(v);
      },

      toggleEdit() { this.editing = !this.editing; },

      // 明細「調整(新增)」
      deltaValue(row) { const k = keyOf(row); return this.deltaMap[k] ?? 0; },
      onDeltaInput(e, row) { const v = parseFloat(e.target.value || '0'); const k = keyOf(row); this.$set(this.deltaMap, k, isNaN(v) ? 0 : v); },

      // 資料映射（摘要/明細）
      normalizeSummary(arr) {
        return arr.map(x => ({
          部門: x.部門 ?? x.DEPARTMENT ?? x.department ?? '',
          回報工時: x.回報工時 ?? x.REPORT_HOURS ?? x.report_hours ?? 0,
          執行工時: x.執行工時 ?? x.工單工時 ?? x.ACT_HOURS ?? x.actual_hours ?? 0,
          有薪假: x.有薪假 ?? x.狀態工時 ?? x.STATUS_HOURS1 ?? 0,
          有薪工時: x.有薪工時 ?? x.狀態工時2 ?? x.STATUS_HOURS2 ?? 0,
          貼紙班: x.貼紙班 ?? x.STICKER ?? 0,
          殘疾人: x.殘疾人 ?? x.DISABLED ?? 0,
          無效工時: x.無效工時 ?? x.INVALID_HOURS ?? 0,
          未設定工單: x.未設定工單 ?? x.NO_WO ?? 0,
          借入工時: x.借入工時 ?? x.BORROW_IN ?? 0,
          借出工時: x.借出工時 ?? x.BORROW_OUT ?? 0,
          調整工時: x.調整工時 ?? x.ADJUST ?? x.adjust ?? 0,


          // 備註 1~4 + 舊欄名兼容
          備註1: x.備註1 ?? x.REMARK1 ?? '',
          備註2: x.備註2 ?? x.REMARK2 ?? '',
          備註3: x.備註3 ?? x.REMARK3 ?? '',
          備註4: x.備註4 ?? x.REMARK4 ?? '',
          備註: (x.備註 ?? x['備註考勤'] ?? x.REMARK ?? x.Remark ?? x.remark ?? x.MEMO ?? x.memo ?? x.NOTE ?? x.note ?? '')
        }));
      },
      normalizeDetails(arr, dep) {
        return arr.map(x => ({
          部門: x.部門 ?? x.DEPARTMENT ?? dep,
          拉: x.拉 ?? x.LINE ?? x.line ?? '',
          工單: x.工單 ?? x.WO_ID ?? x.wo ?? '',
          工序: x.工序 ?? x.OP ?? x.process ?? '',
          品名: x.品名 ?? x.PART_DESC ?? x.part_name ?? '',
          數量: x.數量 ?? x.QTY ?? x.qty ?? 0,
          標準工時: x.標準工時 ?? x.STD_HOURS ?? x.std_hours ?? 0,
          執行工時: x.執行工時 ?? x.ACT_HOURS ?? x.actual_hours ?? 0,
          品檢工時: x.品檢工時 ?? x.QCTIME ?? x.QC_HOURS ?? x.qc_hours ?? x.qcTime ?? 0, // ★ 新增
          物料員工時: x.物料員工時 ?? x.WLTIME ?? x.QC_HOURS ?? x.qc_hours ?? x.qcTime ?? 0, // ★ 新增
          調整工時: x.調整工時 ?? x.ADJUST ?? x.adjust ?? 0,
          備註: (x.備註 ?? x.REMARK ?? x.Remark ?? x.remark ?? x.MEMO ?? x.memo ?? x.NOTE ?? x.note ?? '')
        }));
      },

      // 讀取
      async fetchSummary() {
        try {
          this.loading = true; this.error = ''; this.notice = '';
          this.summary = []; this.details = []; this.deptAdjustMap = {}; this.deltaMap = {};
          const d = toYMD(this.date);
          const { data } = await axios.get(API.WO_SUM(this.area, d), { timeout: 60000 });
          const arr = Array.isArray(data) ? data : (data && data.result ? data.result : []);
          this.summary = this.normalizeSummary(arr);

          // 還原已選部門
          const saved = this.restoreSelected();
          if (saved.length) {
            const set = {};
            this.summary.forEach(r => { if (saved.includes(r.部門)) set[r.部門] = true; });
            this.selected = set; this.persistSelected();
            for (const dep of this.selectedDeptList) { await this.fetchDeptDetails(dep); }
          }
        } catch (err) { console.error(err); this.error = '載入總表失敗'; }
        finally { this.loading = false; }
      },
      async fetchDeptDetails(dep) {
        try {
          this.detailLoading = true; this.error = '';
          const d = toYMD(this.date);
          const { data } = await axios.get(API.WO_VER(this.area, dep, d), { timeout: 60000 });
          const arr = Array.isArray(data) ? data : (data && data.result ? data.result : []);
              console.log('WO_VER raw data sample:', arr[1]); // ⬅ 新增這行
          const rows = this.normalizeDetails(arr, dep);
          this.details = this.details.filter(r => r.部門 !== dep).concat(rows);
          this.recalcDeptAdjust(dep); // 以原『調整工時』加總覆蓋摘要
        } catch (err) { console.error(err); this.error = `載入明細失敗（${dep}）`; }
        finally { this.detailLoading = false; }
      },
      removeDeptDetails(dep) {
        this.details = this.details.filter(r => r.部門 !== dep);
        this.$delete(this.deptAdjustMap, dep);
      },

      // 檢核（新增調整合計需為 0）
      validateZeroSum() {
        const changed = this.details.filter(r => Math.abs(this.deltaMap[keyOf(r)] || 0) > 0.0005);
        if (changed.length === 0) return { ok: true };
        if (ZERO_SUM_SCOPE === 'perDept') {
          const sums = {};
          for (const row of changed) { const d = row.部門 || ''; sums[d] = (sums[d] || 0) + (this.deltaMap[keyOf(row)] || 0); }
          const notZero = Object.entries(sums).filter(([_, v]) => Math.abs(v) > 0.005);
          if (notZero.length) {
            const msg = '下列部門調整加總需為 0：' + notZero.map(([k, v]) => `${k}(${(+v).toFixed(2)})`).join('、');
            return { ok: false, msg };
          }
          return { ok: true };
        } else {
          let total = 0; for (const row of changed) { total += (this.deltaMap[keyOf(row)] || 0); }
          if (Math.abs(total) > 0.005) return { ok: false, msg: `合計：${(+total).toFixed(2)}（必須為 0）` };
          return { ok: true };
        }
      },

      // 後端 payload（數值轉字串：符合 ADJ_WO 的 string 欄位）
      buildAdjWoPayload(row) {
        const delta = this.deltaMap[keyOf(row)] || 0;
        return {
          AREA: this.area,
          MFG_DAY: toYMD(this.date),
          WO_DEPARTMENT: row.部門 || '',
          wo_id: row.工單 || '',
          part_desc: row.品名 || '',
          Qty: String(this.num(row.數量 || 0)),
          std_time: String(this.num(row.標準工時 || 0)),
          working_hours: String(delta)
        };
      },
deptTotal(dept, field){
  const rows = this.grouped[dept] || [];
  const sum = rows.reduce((t, r) => t + this.num(r[field] || 0), 0);
  return sum;
},
      async saveAdjustments() {
        try {
          this.saving = true; this.error = ''; this.notice = '';
          const rows = this.details.filter(r => Math.abs(this.deltaMap[keyOf(r)] || 0) > 0.0005);
          if (rows.length === 0) { this.notice = '沒有變更的調整需要儲存'; return; }

          const check = this.validateZeroSum();
          if (!check.ok) { this.error = check.msg || '調整欄位數字相加必須為0'; return; }

          for (const r of rows) {
            const payload = this.buildAdjWoPayload(r);
            await axios.post(API.ADG_WO_TIME(), payload, { timeout: 60000 });
          }
          this.notice = `已儲存 ${rows.length} 筆調整。`;

          for (const r of rows) { const k = keyOf(r); this.$delete(this.deltaMap, k); }

          // 自動重撈（摘要 + 目前勾選的部門明細）
          await this.fetchSummary();
          for (const dep of this.selectedDeptList) { await this.fetchDeptDetails(dep); }
        } catch (err) { console.error(err); this.error = '儲存調整失敗'; }
        finally { this.saving = false; }
      },

      async exportExcel() {
        try {
          this.exporting = true; this.error = ''; this.notice = '';
          const selectedSet = new Set(this.selectedDeptList);
          const scope = this.anySelected ? this.summary.filter(s => selectedSet.has(s.部門)) : this.summary.slice();

          // ===== 摘要（WOEmployee）=====
          const WOSUMdata = scope.map(r => ({
            部門: r.部門,
            回報管理部考勤工: this.num(r.回報工時),
            工單工時: this.num(r.執行工時),
            狀態工時: this.num(r.有薪假),      // 有薪假
            狀態工時2: this.num(r.有薪工時),    // 有薪工時
            貼紙班: this.num(r.貼紙班),
            殘疾人: this.num(r.殘疾人),
            無效工時: this.num(r.無效工時),
            未設定工單: this.num(r.未設定工單),
            借入工時: this.num(r.借入工時),
            借出工時: this.num(r.借出工時),
            調整工時: this.num(this.displayAdjust(r.部門, r.調整工時)),
            備註: this.joinRemarks(r),
            // 差異計算：不含調整工時
            系統計算差異: this.num(
              this.num(r.回報工時) - (
                this.num(r.執行工時) +
                this.num(r.有薪假) + this.num(r.有薪工時) +
                this.num(r.未設定工單) + this.num(r.無效工時) +
                this.num(r.借出工時) - this.num(r.借入工時) +
                this.num(r.殘疾人)
              )
            )
          }));

          // 摘要加總列（給後端藍底）
          const sumN = k => scope.reduce((t, r) => t + this.num(r[k]), 0);
          const sumAdj = scope.reduce((t, r) => t + this.num(this.displayAdjust(r.部門, r.調整工時)), 0);
          const sumDiff = scope.reduce((t, r) => t + (
            this.num(r.回報工時) - (
              this.num(r.執行工時) + this.num(r.有薪假) + this.num(r.有薪工時) +
              this.num(r.未設定工單) + this.num(r.無效工時) +
              this.num(r.借出工時) - this.num(r.借入工時) + this.num(r.殘疾人)
            )
          ), 0);

          WOSUMdata.push({
            部門: '加總',
            回報管理部考勤工: +sumN('回報工時').toFixed(2),
            工單工時: +sumN('執行工時').toFixed(2),
            狀態工時: +sumN('有薪假').toFixed(2),
            狀態工時2: +sumN('有薪工時').toFixed(2),
            貼紙班: +sumN('貼紙班').toFixed(2),
            殘疾人: +sumN('殘疾人').toFixed(2),
            無效工時: +sumN('無效工時').toFixed(2),
            未設定工單: +sumN('未設定工單').toFixed(2),
            借入工時: +sumN('借入工時').toFixed(2),
            借出工時: +sumN('借出工時').toFixed(2),
            調整工時: +sumAdj.toFixed(2),
            備註: '',
            系統計算差異: +sumDiff.toFixed(2)
          });

          // ===== 明細（WOEmploymentTime）依部門加總 =====
          const detailScope = this.anySelected ? this.details.filter(d => selectedSet.has(d.部門)) : this.details.slice();
          const groups = new Map();
          for (const x of detailScope) { const d = x.部門 || ''; if (!groups.has(d)) groups.set(d, []); groups.get(d).push(x); }
          const deptOrder = this.anySelected ? this.selectedDeptList : Array.from(groups.keys());

          const WOEmploymentTime = [];
          for (const dept of deptOrder) {
            const rows = groups.get(dept) || [];
            if (!rows.length) continue;

            const stdSum = rows.reduce((t, x) => t + this.num(x.標準工時 || 0), 0);
            const actSum = rows.reduce((t, x) => t + this.num(x.執行工時 || 0), 0);

            // 先插入「部門加總」列（後端會套藍底；把合計寫在 工序/品名 欄位）
            WOEmploymentTime.push({
              部門: dept, 拉: '', 工單: '部門加總',
              工序: stdSum.toFixed(2), 品名: actSum.toFixed(2),
              數量: null, 標準工時: 0, 執行工時: 0
            });

            // 再插入每筆明細
            for (const x of rows) {
              WOEmploymentTime.push({
                部門: x.部門 || '', 拉: x.拉 || '', 工單: x.工單 || '',
                工序: x.工序 || '', 品名: x.品名 || '',
                數量: (x.數量 == null ? null : this.num(x.數量)),
                標準工時: this.num(x.標準工時 || 0),
                執行工時: this.num(x.執行工時 || 0)
              });
            }
          }

          // 呼叫後端產生 Excel
          const payload = { MFG_DAY: this.date, WOEmployee: WOSUMdata, WOEmploymentTime };
          const resp = await axios.post(API.WO_EXPORT(), payload, { responseType: 'blob', timeout: 120000 });

          const blob = new Blob([resp.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
          const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `${this.date}生產日報.xlsx`; a.click();

        } catch (err) { console.error(err); this.error = '匯出失敗'; }
        finally { this.exporting = false; }
      },

      onAreaChange() { 
          document.cookie = `${LAST_AREA_COOKIE}=${encodeURIComponent(this.area)};path=/;max-age=${30*86400}`;
             setCookie('MMS_SELECTED_AREA', this.area, 30); // ⬅ 新增：紀錄 Cookie
  // 你原本的行為
        this.persistSelected(); this.fetchSummary(); 
      
      
      }
    },
    mounted() { this.fetchSummary(); }
  });
})();
