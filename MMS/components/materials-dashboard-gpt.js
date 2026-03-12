// 載入專屬樣式
loadCSS("assets/materials-dashboard.css");

// components/materials-dashboard.js
Vue.component('materials-dashboard-view', {
  template: `
<div class="container mt-4 materials-dashboard-page">
  <div class="d-flex justify-content-between align-items-center mb-3">
    <h3><i class="bi bi-box-seam text-success"></i> 資材部 - Dashboard</h3>
  </div>

  <!-- 載入中 -->
  <div v-if="loading" class="text-center my-4">
    <span class="spinner-border text-success" role="status"></span>
    <div class="mt-2">資料載入中...</div>
  </div>

  <div v-else>
    <!-- KPI 小卡 -->
    <div class="row mb-4">
      <div class="col-md-3 col-sm-6 mb-3" v-for="(item, idx) in topMetrics" :key="idx">
        <div class="dashboard-box"
             :class="getBoxClass(item.title)"
             @click="selectMetric(item)">
          <div class="h6"><i :class="getIcon(item.title)" class="me-2"></i>{{ item.title }}</div>
          <div class="h4">{{ item.value }}</div>
          <div v-if="item.target !== undefined" class="small text-end text-black mt-1" style="opacity: 0.8;">
            目標：{{ item.target }}
          </div>
        </div>
      </div>
    </div>

  
<!-- 未結採購單 -->
<el-card v-if="selectedMetric === '未結採購單s件'" shadow="hover" class="mt-4">
  <div class="d-flex justify-content-between align-items-center mb-2">
    <h5 class="text-primary">未結採購單 明細123</h5>
    <el-button type="success" size="small" @click="downloadExcel('unclosed-po')">下載 Excel</el-button>
  </div>

  <!-- ✅ 可左右/上下滑動，表頭固定 -->
  <div class="table-responsive" style="max-height: 600px; overflow-y: auto; overflow-x: auto;">
    <table class="table table-bordered table-sm text-center align-middle mb-0">
      <thead class="thead-light sticky-top bg-light">
        <tr>
          <th>請購日期</th>
          <th>部門名稱</th>
          <th>單據號</th>
          <th>料號</th>
          <th>供應商</th>
          <th>供應商名稱</th>
          <th>品名</th>
          <th>規格</th>
          <th>數量</th>
          <th>本幣金額</th>
          <th>外幣金額</th>
          <th>送貨單號</th>
        </tr>
      </thead>
      <tbody>
        <tr v-if="!purchaseOrders.length">
          <td colspan="12" class="text-muted">目前沒有資料</td>
        </tr>
        <tr v-for="(row, idx) in purchaseOrders" :key="idx">
          <td>{{ row.請購日期 ? row.請購日期.split(' ')[0] : '' }}</td>
          <td style="text-align: left;">{{ row.部門名稱 }}</td>
          <td>{{ row.單據號 }}</td>
          <td>{{ row.料號 }}</td>
          <td style="text-align: left;">{{ row.供應商 }}</td>
          <td style="text-align: left;">{{ row.供應商名稱 }}</td>
          <td style="text-align: left;">{{ row.品名 }}</td>
          <td style="text-align: left;">{{ row.規格 }}</td>
          <td style="text-align: right;">{{ Number(row.數量).toLocaleString() }}</td>
          <td style="text-align: right;">{{ Number(row.本幣金額).toLocaleString() }}</td>
          <td style="text-align: right;">{{ Number(row.外幣金額).toLocaleString() }}</td>
          <td>{{ row.送貨單號 }}</td>
        </tr>
      </tbody>
    </table>
  </div>
</el-card>



<!-- 待報關物料 -->
<el-card v-else-if="selectedMetric === '待報關物料'" shadow="hover" class="mt-4">
  <div class="d-flex justify-content-between align-items-center mb-2">
    <h5 class="text-primary">待報關物料 明細</h5>
    <el-button type="success" size="small" @click="downloadExcel('customs-materials')">下載 Excel</el-button>
  </div>

  <!-- ✅ 可上下滾動、表頭固定 -->
  <div class="table-responsive" style="max-height: 600px; overflow-y:auto;">
    <table class="table table-bordered table-sm text-center align-middle mb-0">
      <thead class="thead-light sticky-top bg-light">
        <tr>
          <th>收料日期</th>
          <th>訂單號</th>
          <th>行號</th>
          <th>單據類型</th>
          <th>料號</th>
          <th>品名</th>
          <th>規格</th>
          <th>單位</th>
          <th>訂單數量</th>
          <th>本幣單價</th>
          <th>外幣單價</th>
          <th>供應商名稱</th>
          <th>付款方式</th>
          <th>說明</th>
          
          <th>珠海</th>
          <th>在途</th>
          <th>檢驗中</th>
          <th>待報關</th>
          <th>送貨單號</th>
          
        </tr>
      </thead>
      <tbody>
        <tr v-if="!customsMaterials.length">
          <td colspan="21" class="text-muted">目前沒有資料</td>
        </tr>
        <tr v-for="(row, idx) in customsMaterials" :key="idx">
          <td>{{ row.收料日期 ? row.收料日期.split(' ')[0] : '' }}</td>
          <td>{{ row.訂單號 }}</td>
          <td>{{ row.行號 }}</td>
          <td>{{ row.單據類型 }}</td>
          <td>{{ row.料號 }}</td>
          <td class="text-start">{{ row.品名 }}</td>
          <td class="text-start">{{ row.規格 }}</td>
          <td>{{ row.單位 }}</td>
          <td style="text-align: right;">{{ Number(row.訂單數量).toLocaleString() }}</td>
          <td style="text-align: right;">{{ Number(row.本幣單價).toLocaleString() }}</td>
          <td style="text-align: right;">{{ Number(row.外幣單價).toLocaleString() }}</td>
          <td>{{ row.供應商名稱 }}</td>
          <td>{{ row.付款方式 }}</td>
          <td class="text-start">{{ row.說明 }}</td>
          
          <td style="text-align: right;">{{ Number(row.珠海).toLocaleString() }}</td>
          <td style="text-align: right;">{{ Number(row.在途).toLocaleString() }}</td>
          <td style="text-align: right;">{{ Number(row.檢驗中).toLocaleString() }}</td>
          <td style="text-align: right;">{{ Number(row.待報關).toLocaleString() }}</td>
          <td>{{ row.送貨單號 }}</td>
          
        </tr>
      </tbody>
    </table>
  </div>
</el-card>


    <!-- 已挑單未扣庫存 -->
    <el-card v-else-if="selectedMetric === '已挑單未扣庫存'" shadow="hover" class="mt-4">
      <div class="d-flex justify-content-between align-items-center mb-2">
        <h5 class="text-primary">已挑單未扣庫存 明細</h5>
        <el-button type="success" size="small" @click="downloadExcel('picked-not-issued')">下載 Excel</el-button>
      </div>
      <table class="table table-bordered table-sm text-center">
        <thead class="thead-light">
          <tr>
            <th>出貨單號</th><th>發運日期</th><th>訂單號</th><th>單據類型</th>
            <th>行號</th><th>料號</th><th>品名</th><th>規格</th>
            <th>數量</th><th>外幣單價</th><th>外幣金額</th><th>本幣單價</th><th>本幣金額</th>
          </tr>
        </thead>
        <tbody>
          <tr v-if="!pickedNotIssued.length">
            <td colspan="13" class="text-muted">目前沒有資料</td>
          </tr>
          <tr v-for="(row, idx) in pickedNotIssued" :key="idx">
            <td>{{ row.出貨單號 }}</td>
            <td>{{ row.發運日期.split(' ')[0] }}</td>
            <td>{{ row.訂單號 }}</td>
            <td>{{ row.單據類型 }}</td>
            <td>{{ row.行號 }}</td>
            <td style="text-align: left;">{{ row.料號 }}</td>
            <td style="text-align: left;">{{ row.品名 }}</td>
            <td>{{ row.規格 }}</td>
            <td>{{ row.數量 }}</td>
            <td>{{ row.外幣單價 }}</td>
            <td>{{ row.外幣金額 }}</td>
            <td>{{ row.本幣單價 }}</td>
            <td>{{ row.本幣金額 }}</td>
          </tr>
        </tbody>
      </table>
    </el-card>

    <!-- 已扣庫存未結案 -->
    <el-card v-else-if="selectedMetric === '已扣庫存未結案'" shadow="hover" class="mt-4">
      <div class="d-flex justify-content-between align-items-center mb-2">
        <h5 class="text-primary">已扣庫存未結案 明細</h5>
        <el-button type="success" size="small" @click="downloadExcel('issued-not-closed')">下載 Excel</el-button>
      </div>
      <table class="table table-bordered table-sm text-center">
        <thead class="thead-light">
          <tr>
            <th>出貨單號</th><th>發運日期</th><th>訂單號</th><th>單據類型</th>
            <th>行號</th><th>料號</th><th>品名</th><th>規格</th>
            <th>數量</th><th>外幣單價</th><th>外幣金額</th><th>本幣單價</th><th>本幣金額</th>
          </tr>
        </thead>
        <tbody>
          <tr v-if="!issuedNotClosed.length">
            <td colspan="13" class="text-muted">目前沒有資料</td>
          </tr>
          <tr v-for="(row, idx) in issuedNotClosed" :key="idx">
            <td>{{ row.出貨單號 }}</td>
            <td>{{ row.發運日期.split(' ')[0] }}</td>
            <td>{{ row.訂單號 }}</td>
            <td>{{ row.單據類型 }}</td>
            <td>{{ row.行號 }}</td>
            <td style="text-align: left;">{{ row.料號 }}</td>
            <td style="text-align: left;">{{ row.品名 }}</td>
            <td>{{ row.規格 }}</td>
            <td>{{ row.數量 }}</td>
            <td>{{ row.外幣單價 }}</td>
            <td>{{ row.外幣金額 }}</td>
            <td style="text-align: right;">{{ row.本幣單價 }}</td>
            <td style="text-align: right;">{{ row.本幣金額 }}</td>
          </tr>
        </tbody>
      </table>
    </el-card>

    <!-- 未結 FLOW -->
    <el-card v-else-if="selectedMetric === '未結FLOW'" shadow="hover" class="mt-4">
      <div class="d-flex justify-content-between align-items-center mb-2">
        <h5 class="text-primary">未結 FLOW 表單 明細</h5>
        <el-button type="success" size="small" @click="downloadExcel('flow-unclosed')">下載 Excel</el-button>
      </div>
      <table class="table table-bordered table-sm text-center">
        <thead class="thead-light">
          <tr>
            <th>表單名稱</th><th>文件編號</th><th>狀態</th><th>關卡</th><th>申請人</th><th>建立日期</th>
          </tr>
        </thead>
        <tbody>
          <tr v-if="!flowUnclosedRows.length">
            <td colspan="6" class="text-muted">目前沒有資料</td>
          </tr>
          <tr v-for="(row, idx) in flowUnclosedRows" :key="idx">
            <td>{{ row.表單名稱 }}</td>
            <td><a :href="row.文件編號連結" target="_blank">{{ row.文件編號 }}</a></td>
            <td>{{ row.表單狀態 }}</td>
            <td>{{ row.簽和關卡 }}</td>
            <td>{{ row.申請人 }}</td>
            <td>{{ row.建立日期 }}</td>
          </tr>
        </tbody>
      </table>
    </el-card>


<!-- 已上線工單缺料表 -->
<el-card v-else-if="selectedMetric === '已上線工單缺料表'" shadow="hover" class="mt-4">
  <div class="d-flex justify-content-between align-items-center mb-2">
    <h5 class="text-danger">已上線工單缺料表 明細</h5>
    <el-button type="success" size="small" @click="downloadExcel('workorder-shortage')">下載 Excel</el-button>
  </div>

  <!-- ✅ 可左右/上下滑動，表頭固定 -->
  <div class="table-responsive" style="max-height: 600px; overflow-y: auto; overflow-x: auto;">
    <table class="table table-bordered table-sm text-center align-middle mb-0">
      <thead class="thead-light sticky-top bg-light">
        <tr>
          <th>班別</th>
          <th>工單號</th>
          <th>狀態</th>
          <th>母件料號</th>
          <th>母件品名</th>
          <th>母件規格</th>
          <th>工單數量</th>
          <th>完工數量</th>
          <th>標準工時</th>
          <th>已入工時</th>
          <th>子件料號</th>
          <th>子件品名</th>
          <th>子件規格</th>
          <th>應發料量</th>
          <th>已發數量</th>
          <th>庫存</th>
          <th>珠海數量</th>
          <th>在途數量</th>
          <th>檢驗中數量</th>
          <th>待報關數量</th>
        </tr>
      </thead>
      <tbody>
        <tr v-if="!workorderShortage.length">
          <td colspan="21" class="text-muted">目前沒有資料</td>
        </tr>
        <tr v-for="(row, idx) in workorderShortage" :key="idx">
          <td>{{ row.班別 }}</td>
          <td>{{ row.工單號 }}</td>
          <td>{{ row.狀態 }}</td>
          <td>{{ row.母件料號 }}</td>
          <td style="text-align:left;">{{ row.母件品名 }}</td>
          <td style="text-align:left;">{{ row.母件規格 }}</td>
          <td style="text-align:right;">{{ Number(row.工單數量).toLocaleString() }}</td>
          <td style="text-align:right;">{{ Number(row.完工數量).toLocaleString() }}</td>
          <td style="text-align:right;">{{ row.標準工時 }}</td>
          <td style="text-align:right;">{{ row.已入工時 }}</td>
          <td>{{ row.子件料號 }}</td>
          <td style="text-align:left;">{{ row.子件品名 }}</td>
          <td style="text-align:left;">{{ row.子件規格 }}</td>
          <td style="text-align:right;">{{ Number(row.應發料量).toLocaleString() }}</td>
          <td style="text-align:right;">{{ Number(row.已發數量).toLocaleString() }}</td>
          <td style="text-align:right;">{{ Number(row.庫存).toLocaleString() }}</td>
          <td style="text-align:right;">{{ Number(row.珠海數量).toLocaleString() }}</td>
          <td style="text-align:right;">{{ Number(row.在途數量).toLocaleString() }}</td>
          <td style="text-align:right;">{{ Number(row.檢驗中數量).toLocaleString() }}</td>
          <td style="text-align:right;">{{ Number(row.待報關數量).toLocaleString() }}</td>
        </tr>
      </tbody>
    </table>
  </div>
</el-card>

<!-- 異常倉別追蹤-來料加工 -->
<el-card v-else-if="selectedMetric === '異常倉別追蹤-來料加工'" shadow="hover" class="mt-4">
  <h3>📦 異常倉別追蹤-來料加工</h3>
  <div class="mb-3 d-flex flex-wrap align-items-center gap-2">
    <el-button type="primary" @click="fetchStockData">重新載入庫存</el-button>
    <el-button type="success" @click="downloadStockExcel">下載 Excel</el-button>
    <el-input v-model="searchText" placeholder="搜尋料號或品名..." style="width: 200px;"></el-input>
    <el-select v-model="selectedFilter" placeholder="篩選欄位" style="width: 160px;">
      <el-option label="不篩選" value=""></el-option>
      <el-option v-for="field in numericFields" :key="field" :label="field + ' ≠ 0'" :value="field"></el-option>
    </el-select>
  </div>

  <div v-if="loading" class="text-center my-4">
    <span class="spinner-border text-primary" role="status"></span>
    <div class="mt-2">資料載入中...</div>
  </div>

  <div v-else class="table-responsive fixed-header" style="max-height: 600px; overflow-y:auto;">
    <table class="table table-bordered table-sm text-center align-middle mb-0">
      <thead class="thead-dark sticky-top bg-dark text-white">
        <tr>
          <th>料號</th>
          <th>品名</th>
          <th>良品倉(A+B)</th>
          <th>A良品倉</th>
          <th>B良品倉</th>
          <th>來料短少倉(STG)</th>
          <th>借調倉(LED)</th>
          <th>遺失倉(LOS)</th>
          <th>待判定倉(TBD)</th>
          <th>製損倉(MRB)</th>
          <th>來料不良倉(PUB)</th>
          <th v-if="canViewSensitiveColumns()">DN</th>
          <th v-if="canViewSensitiveColumns()">OFS</th>
          <th v-if="canViewSensitiveColumns()">DG</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="item in filteredStockData" :key="item.料號">
          <td>{{ (item.料號 || '').trim() }}</td>
          <td class="text-start">{{ item.品名 }}</td>
          <td>{{ item.loc }}</td>
          <td>{{ item.loc_A }}</td>
          <td>{{ item.loc_B }}</td>
          <td>{{ item.loc_STG }}</td>
          <td>{{ item.loc_LED }}</td>
          <td>{{ item.loc_LOS }}</td>
          <td>{{ item.loc_TBD }}</td>
          <td>{{ item.loc_MRB }}</td>
          <td>{{ item.loc_PUB }}</td>
          <td v-if="canViewSensitiveColumns()">{{ item.loc_DN }}</td>
          <td v-if="canViewSensitiveColumns()">{{ item.loc_OFS }}</td>
          <td v-if="canViewSensitiveColumns()" :class="{ 'text-danger': item.loc_DG > 0 }">
            {{ item.loc_DG }}
          </td>
        </tr>
      </tbody>
    </table>
  </div>
</el-card>
<!-- 異常倉別追蹤-一般物料 -->
<el-card v-else-if="selectedMetric === '異常倉別追蹤-一般物料'" shadow="hover" class="mt-4">
  <h3>📦 異常倉別追蹤-一般物料</h3>
  <div class="mb-3 d-flex flex-wrap align-items-center gap-2">
    <el-button type="primary" @click="fetchGeneralStock">重新載入庫存</el-button>
    <el-button type="success" @click="downloadGeneralStockExcel">下載 Excel</el-button>
    <el-input v-model="searchText" placeholder="搜尋料號或品名..." style="width: 200px;"></el-input>
    <el-select v-model="selectedFilter" placeholder="篩選欄位" style="width: 160px;">
      <el-option label="不篩選" value=""></el-option>
      <el-option v-for="field in numericFields" :key="field" :label="field + ' ≠ 0'" :value="field"></el-option>
    </el-select>
  </div>

  <div v-if="loading" class="text-center my-4">
    <span class="spinner-border text-primary" role="status"></span>
    <div class="mt-2">資料載入中...</div>
  </div>

  <div v-else class="table-responsive fixed-header" style="max-height: 600px; overflow-y:auto;">
    <table class="table table-bordered table-sm text-center align-middle mb-0">
      <thead class="thead-dark sticky-top bg-dark text-white">
        <tr>
          <th>料號</th>
          <th>品名</th>
          <th>良品倉(A+B)</th>
          <th>A良品倉</th>
          <th>B良品倉</th>
          <th>來料短少倉(STG)</th>
          <th>借調倉(LED)</th>
          <th>遺失倉(LOS)</th>
          <th>待判定倉(TBD)</th>
          <th>製損倉(MRB)</th>
          <th>來料不良倉(PUB)</th>
          <th>DN</th>
          <th>OFS</th>
          <th>DG</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="item in filteredStockData" :key="item.料號">
          <td>{{ item.料號 }}</td>
          <td class="text-start">{{ item.品名 }}</td>
          <td>{{ item.loc }}</td>
          <td>{{ item.loc_A }}</td>
          <td>{{ item.loc_B }}</td>
          <td>{{ item.loc_STG }}</td>
          <td>{{ item.loc_LED }}</td>
          <td>{{ item.loc_LOS }}</td>
          <td>{{ item.loc_TBD }}</td>
          <td>{{ item.loc_MRB }}</td>
          <td>{{ item.loc_PUB }}</td>
          <td>{{ item.loc_DN }}</td>
          <td>{{ item.loc_OFS }}</td>
          <td>{{ item.loc_DG }}</td>
        </tr>
      </tbody>
    </table>
  </div>
</el-card>



  </div>
</div>
`,
  data() {
    return {
      loading: true,
      selectedMetric: '',
      purchaseOrders: [],
      customsMaterials: [],
      pickedNotIssued: [],
      issuedNotClosed: [],
      flowUnclosedRows: [],
      flowUnclosedCount: 0,
      workorderShortage: [],
       stockData: [],
       generalStock: [],
    searchText: "",
    selectedFilter: "",
    numericFields: ["loc", "loc_A", "loc_B", "loc_STG", "loc_LED", "loc_LOS", "loc_TBD", "loc_MRB", "loc_PUB", "loc_DN", "loc_OFS", "loc_DG"],
      topMetrics: [
        { title: "未結採購單s件", value: "0 件", target: "0" },
        { title: "待報關物料", value: "0 件", target: "0" },
        { title: "已挑單未扣庫存", value: "0 件", target: "0" },
        { title: "已扣庫存未結案", value: "0 件", target: "0" },
        { title: "未結FLOW", value: "0 件", target: "0" },
        { title: "已上線工單缺料表", value: "0 件", target: "0" },
        { title: "異常倉別追蹤-一般物料", value: "0 件", target: "0" },
        { title: "異常倉別追蹤-來料加工", value: "0 件", target: "0" }
      ]
    };
  },
computed: {
  filteredStockData() {
    // 根據目前選到哪個模組，決定資料來源
    let data = [];

    if (this.selectedMetric === '異常倉別追蹤-來料加工') {
      data = this.stockData;
    } else if (this.selectedMetric === '異常倉別追蹤-一般物料') {
      data = this.generalStock;
    } else {
      data = [];
    }

    // 🔍 關鍵字搜尋
    if (this.searchText) {
      const keyword = this.searchText.trim().toLowerCase();
      data = data.filter(
        r =>
          (r.料號 && r.料號.toLowerCase().includes(keyword)) ||
          (r.品名 && r.品名.toLowerCase().includes(keyword))
      );
    }

    // 🔢 數值篩選 (只顯示不為 0 的)
    if (this.selectedFilter) {
      data = data.filter(r => Number(r[this.selectedFilter]) !== 0);
    }

    return data;
  }
},

  methods: {
    getBoxClass(title) { return "border-primary"; },
    getIcon(title) {
      if (title.includes("採購")) return "bi bi-cart-check";
      if (title.includes("報關")) return "bi bi-truck";
      if (title.includes("挑單")) return "bi bi-list-check";
      if (title.includes("扣庫存")) return "bi bi-box-arrow-down";
      if (title.includes("FLOW")) return "bi bi-envelope-exclamation";
      return "bi bi-circle";
    },
    selectMetric(item) { this.selectedMetric = item.title; },
    async fetchPurchaseOrders() {
      const res = await axios.get("https://mms.leapoptical.com:5088/api/Materials/unclosed-po?company=00109&docType=OM");
      this.purchaseOrders = res.data;
      this.topMetrics[0].value = this.purchaseOrders.length + " 件";
    },
    async fetchCustomsMaterials() {
      const res = await axios.get("https://mms.leapoptical.com:5088/api/Materials/customs-materials");
      this.customsMaterials = res.data;
      this.topMetrics[1].value = this.customsMaterials.length + " 件";
    },
    async fetchPickedNotIssued() {
      const res = await axios.get("https://mms.leapoptical.com:5088/api/Materials/picked-not-issued?company=00109");
      this.pickedNotIssued = res.data;
      this.topMetrics[2].value = this.pickedNotIssued.length + " 件";
    },
    async fetchIssuedNotClosed() {
      const res = await axios.get("https://mms.leapoptical.com:5088/api/Materials/issued-not-closed?company=00109");
      this.issuedNotClosed = res.data;
      this.topMetrics[3].value = this.issuedNotClosed.length + " 件";
    },
    async fetchFlowUnclosed() {
      try {
        const depts = ['1091210', '109110', '1091510', '1091910', '1092200'];
        let allData = [];
        const requests = depts.map(d =>
          axios.get('https://mms.leapoptical.com:5088/api/LinePerformance/unclosedflow/list', { params: { dept: d } })
        );
        const responses = await Promise.all(requests);
        responses.forEach(res => { allData = allData.concat(res.data); });
        this.flowUnclosedRows = allData;
        this.flowUnclosedCount = allData.length;
        this.topMetrics[4].value = this.flowUnclosedCount + " 件";
      } catch (err) {
        console.error("未結FLOW表單查詢失敗:", err);
      }
    },
    async fetchWorkorderShortage() {
  try {
    const res = await axios.get("https://mms.leapoptical.com:5088/api/Materials/shortage?factory=%20%20%20%20%201099000");
    this.workorderShortage = res.data;
    this.topMetrics[5].value = this.workorderShortage.length + " 件";
  } catch (err) {
    console.error("❌ 載入已上線工單缺料表失敗:", err);
  }
},
downloadExcel(type) {
  let rows = [], fileName = "";

  if (type === "unclosed-po") {
    rows = this.purchaseOrders;
    fileName = "未結採購單.xlsx";
  } else if (type === "customs-materials") {
    rows = this.customsMaterials;
    fileName = "待報關物料.xlsx";
  } else if (type === "picked-not-issued") {
    rows = this.pickedNotIssued;
    fileName = "已挑單未扣庫存.xlsx";
  } else if (type === "issued-not-closed") {
    rows = this.issuedNotClosed;
    fileName = "已扣庫存未結案.xlsx";
  } else if (type === "flow-unclosed") {
    rows = this.flowUnclosedRows;
    fileName = "未結FLOW表單.xlsx";
  } else if (type === "workorder-shortage") {
    rows = this.workorderShortage;
    fileName = "已上線工單缺料表.xlsx";
  }

  if (!rows || !rows.length) {
    this.$message.warning("沒有資料可匯出");
    return;
  }

  // ✅ 將 JSON 轉成 worksheet
  const worksheet = XLSX.utils.json_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");

  // ✅ 產生 Excel 檔案
  const wbout = XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });

  // ✅ 下載檔案
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
},
async fetchStockData() {
    this.loading = true;
    try {
      const res = await axios.get("https://mms.leapoptical.com:5088/api/Materials/subcontract-stock");
      this.stockData = res.data;
      this.topMetrics[7].value = this.stockData.length + " 筆";
    } catch (err) {
      this.$message.error("載入代工庫存失敗");
      console.error(err);
    } finally {
      this.loading = false;
    }
  },


 downloadStockExcel() {
    if (!this.stockData.length) return this.$message.warning("沒有資料可匯出");
    const ws = XLSX.utils.json_to_sheet(this.stockData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "代工庫存表");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "代工庫存表.xlsx";
    link.click();
  },
  canViewSensitiveColumns() {
    // 可改為依使用者權限判斷
    return true;
  },
  async fetchGeneralStock() {
  this.loading = true;
  try {
    const res = await axios.get("https://mms.leapoptical.com:5088/api/Materials/general-stock");
    this.generalStock = res.data;
    this.topMetrics[6].value = this.generalStock.length + " 筆";
  } catch (err) {
    this.$message.error("載入一般物料庫存失敗");
    console.error(err);
  } finally {
    this.loading = false;
  }
},
downloadGeneralStockExcel() {
  if (!this.generalStock.length) return this.$message.warning("沒有資料可匯出");
  const ws = XLSX.utils.json_to_sheet(this.generalStock);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "一般物料庫存表");
  const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const blob = new Blob([wbout], { type: "application/octet-stream" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "一般物料庫存表.xlsx";
  link.click();
},

  },
  mounted() {
    Promise.all([
      this.fetchPurchaseOrders(),
      this.fetchCustomsMaterials(),
      this.fetchPickedNotIssued(),
      this.fetchIssuedNotClosed(),
      this.fetchFlowUnclosed(),
      this.fetchWorkorderShortage(),
     this.fetchStockData(), // ✅ 加這行
     this.fetchGeneralStock() // ✅ 加這行
    ]).finally(() => { this.loading = false; });
  }
});
