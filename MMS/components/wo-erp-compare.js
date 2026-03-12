
// ========== MMS - 作業中心：工單狀態與 ERP 對比報表 ==========
(function () {
  const API_BASE = window.API_BASE || 'https://mms.leapoptical.com:7238';

  // 工具：把 Date 轉為 YYYYMMDD
  function fmtYYYYMMDD(d) {
    const yy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, '0');
    const dd = String(d.getDate()).padStart(2, '0');
    return `${yy}${mm}${dd}`;
  }

  // 預設日期=昨天（你先前有特別要求）
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  Vue.component('wo-erp-compare-view', {
    template: `
    <div class="p-3">
      <div class="d-flex align-items-center gap-2 mb-3 flex-wrap">
        <label class="me-2">{{ t('area') || '廠區' }}</label>
        <select class="form-select form-select-sm w-auto" v-model="selectedArea" @change="loadReport">
          <option value="ZH">ZH</option>
          <option value="TC">TC</option>
          <option value="VN">VN</option>
          <option value="TW">TW</option>
        </select>

        <input type="date" class="form-control form-control-sm w-auto" v-model="selectedDate" @change="loadReport" />

        <button class="btn btn-sm btn-outline-primary ms-auto" @click="exportExcel" :disabled="loading">
          {{ t('printDetail') || '打印明細' }}
        </button>
      </div>

      <h5 class="mb-3">{{ t('woErpTitle') || '工單狀態與 ERP 對比報表' }}</h5>

      <div v-if="loading" class="text-muted">Loading…</div>

      <div v-else>
        <div class="table-responsive mb-4">
          <table class="table table-sm table-bordered align-middle">
            <caption class="text-end fst-italic">總表</caption>
            <thead class="table-light">
              <tr>
                <th>{{ t('department') || '部門' }}</th>
                <th>NFC工時與ERP工時不符</th>
                <th>NFC已完成ERP未結案</th>
                <th>NFC未完成ERP已結案</th>
                <th>NFC沒資料ERP已結案</th>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(row, idx) in woderpdata" :key="idx">
                <td>{{ row.部門 }}</td>
                <td>{{ row['nfc與erp工時不符數量'] }}</td>
                <td>{{ row['nfc已完成erp未結案數量'] }}</td>
                <td>{{ row['nfc未完成erp已結案'] }}</td>
                <td>{{ row['nfc沒資料erp已結案'] }}</td>
              </tr>
              <tr v-if="!woderpdata || !woderpdata.length">
                <td colspan="5" class="text-center text-muted">— 無資料 —</td>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- 如需顯示各類明細，可再做 1~6 類的折疊卡片；目前主要提供匯出 -->
      </div>
    </div>
    `,
    data() {
      return {
        loading: false,
        selectedArea: 'VN',
        selectedDate: fmtYYYYMMDD(yesterday), // 預設昨天（YYYYMMDD 格式先塞，下面會轉成 <input type="date"> 需要的 YYYY-MM-DD）
        // 資料集
        woderpdata: [],  // 總表
        erpdata1: [],    // 明細 1~6：與原 NFC.html 相同對應
        erpdata2: [],
        erpdata3: [],
        erpdata4: [],
        erpdata5: [],
        erpdata6: []
      };
    },
    created() {
      // 把 YYYYMMDD 轉為 input[type=date] 的 YYYY-MM-DD
      const y = this.selectedDate.slice(0,4);
      const m = this.selectedDate.slice(4,6);
      const d = this.selectedDate.slice(6,8);
      this.selectedDate = `${y}-${m}-${d}`;
      this.loadReport();
    },
    methods: {
      t(key) {
        // 輕量 i18n：沿用 index 的 I18N.menu & 其他字典
        const lang = (window.mms_lang || localStorage.getItem('mms_lang') || 'zh');
        const dict = Object.assign({
          area: { zh:'廠區', en:'Area', vi:'Khu vực' },
          department: { zh:'部門', en:'Dept.', vi:'Bộ phận' },
          woErpTitle: { zh:'工單狀態與 ERP 對比報表', en:'WO vs ERP Comparison', vi:'Đối chiếu WO/ERP' },
          printDetail: { zh:'打印明細', en:'Export Details', vi:'Xuất chi tiết' },
        }, (window.I18N || {}));
        return (dict[key] && dict[key][lang]) || null;
      },

      async loadReport() {
        this.loading = true;
        this.woderpdata = [];
        this.erpdata1 = this.erpdata2 = this.erpdata3 = this.erpdata4 = this.erpdata5 = this.erpdata6 = [];

        if (!this.selectedArea) { this.loading = false; return; }

        // 組日期（YYYYMMDD）：WO_ERP_5 需要 MFG_DAY
        const d = new Date(this.selectedDate);
        const mfg = fmtYYYYMMDD(d);

        // 完全沿用 NFC.html 中的 7 支 API（總表 + 6 組明細）
        const url1 = `${API_BASE}/api/Report/WO_ERP__SUM_version?AREA=${this.selectedArea}`;         // 總表
        const url2 = `${API_BASE}/api/Report/WO_ERP_1?AREA=${this.selectedArea}`;
        const url3 = `${API_BASE}/api/Report/WO_ERP_2?AREA=${this.selectedArea}`;
        const url4 = `${API_BASE}/api/Report/WO_ERP_3?AREA=${this.selectedArea}`;
        const url5 = `${API_BASE}/api/Report/WO_ERP_4?AREA=${this.selectedArea}`;
        const url6 = `${API_BASE}/api/Report/WO_ERP_5?AREA=${this.selectedArea}&MFG_DAY=${mfg}`;    // 這支要日期
        const url7 = `${API_BASE}/api/Report/WO_ERP_6?AREA=${this.selectedArea}`;
        // 以上對應自原 NFC.html 的 woderp() 呼叫組合 【API 列表節錄】；見檔案中的相同路由。:contentReference[oaicite:0]{index=0}

        try {
          const res = await Promise.all([url1,url2,url3,url4,url5,url6,url7].map(u => fetch(u)));
          const data = await Promise.all(res.map(r => r.json()));

          this.woderpdata = data[0];
          this.erpdata1   = data[1];
          this.erpdata2   = data[2];
          this.erpdata3   = data[3];
          this.erpdata4   = data[4];
          this.erpdata5   = data[5];
          this.erpdata6   = data[6];
        } catch (err) {
          console.error('WO/ERP compare load error:', err);
        } finally {
          this.loading = false;
        }
      },

      // 匯出 Excel（總表 + 6 張明細各一個 Sheet）
      async exportExcel() {
        if (!window.XLSX) { return alert('找不到 XLSX.js，請先在 index.html 載入'); }

        const wb = XLSX.utils.book_new();

        // 總表
        const sumRows = (this.woderpdata || []).map(x => ({
          '部門': x.部門,
          'NFC工時與ERP工時不符': x['nfc與erp工時不符數量'],
          'NFC已完成ERP未結案': x['nfc已完成erp未結案數量'],
          'NFC未完成ERP已結案': x['nfc未完成erp已結案'],
          'NFC沒資料ERP已結案': x['nfc沒資料erp已結案'],
        }));
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sumRows), '總表');

        // 明細（1~6）— 欄位依原 NFC.html 的明細表頭建立（示例以 #1）
        const buildDetail = (arr) => (arr||[]).map(r => ({
          '部門': r.部門,
          'NFC工單狀態': r.nfC工單狀態,
          '日期': r.日期,
          '工單': r.工單,
          '品名': r.品名,
          '數量': r.數量,
          '標準工時': r.標準工時,
          'NFC執行工時': r.nfC執行工時,
          'ERP輸入工時': r.erP輸入工時,
          'ERP狀態': r.erP狀態,
          'ERP已入庫數': r.erP已入庫數,
        }));

        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildDetail(this.erpdata1)), '明細1');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildDetail(this.erpdata2)), '明細2');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildDetail(this.erpdata3)), '明細3');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildDetail(this.erpdata4)), '明細4');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildDetail(this.erpdata5)), '明細5');
        XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(buildDetail(this.erpdata6)), '明細6');

        const mfg = fmtYYYYMMDD(new Date(this.selectedDate));
        XLSX.writeFile(wb, `WO_ERP_COMPARE_${this.selectedArea}_${mfg}.xlsx`);
      }
    }
  });
})();

