// 載入專屬樣式 (請確保此檔案存在且內容正確)
loadCSS("assets/materials-dashboard.css");

Vue.component('materials-dashboard-view', {
  template: `
<div class="container-fluid mt-4 materials-dashboard-page">
  <div class="d-flex justify-content-between align-items-center mb-3">
    <h3><i class="bi bi-box-seam text-success"></i> 資材部 - Dashboard</h3>
  </div>

  <div class="row mb-4">
    <div class="col mb-3" v-for="item in topMetrics" :key="item.id">
      <div class="dashboard-box card card-body text-center"
           :class="{ 'active': selectedMetric === item.id }"
           @click="handleMetricClick(item)">
        <div class="h6"><i :class="getIcon(item.title)" class="me-2"></i>{{ item.title }}</div>
        <div class="h4 mb-0">{{ item.value }}</div>
        <div v-if="item.target !== undefined" class="small text-end text-muted mt-1">
          目標：{{ item.target }}
        </div>
      </div>
    </div>
  </div>

  <div v-if="appLoading" class="text-center my-5">
    <span class="spinner-border text-success" role="status" style="width: 3rem; height: 3rem;"></span>
    <h4 class="mt-3">Dashboard 初始資料載入中...</h4>
  </div>

  <el-tabs v-else v-model="selectedMetric" type="border-card">

    <el-tab-pane label="未結S件採購單" name="unclosed-po">
      <el-card shadow="hover">
         <div slot="header" class="d-flex justify-content-between align-items-center">
          <div>
  <span class="text-primary h5 d-block">未結S件採購單 明細</span>
</div>

   
            <div>
              <el-input v-model="searchText.po" placeholder="搜尋單據號、料號、品名..." style="width: 250px; margin-right: 10px;" clearable></el-input>
              <el-button type="success" @click="downloadExcel('unclosed-po')"><i class="bi bi-download"></i> 下載 Excel</el-button>
            </div>
         </div>
         <el-table :data="filteredData.po" v-loading="loading" stripe height="550" style="width: 100%">
            <el-table-column prop="請購日期" label="請購日期" width="120" sortable :formatter="formatDate"></el-table-column>
            <el-table-column prop="部門名稱" label="部門名稱" width="150" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="單據號" label="單據號" width="160" sortable></el-table-column>
            <el-table-column prop="料號" label="料號" width="160" sortable></el-table-column>
            <el-table-column prop="供應商" label="供應商" width="120" sortable></el-table-column>
            <el-table-column prop="供應商名稱" label="供應商名稱" width="250" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="品名" label="品名" width="250" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="規格" label="規格" width="250" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="數量" label="數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="本幣金額" label="本幣金額" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="外幣金額" label="外幣金額" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="送貨單號" label="送貨單號" width="150" sortable></el-table-column>
         </el-table>
      </el-card>
        <span class="text-success small d-block mt-1">資料來源：ERP未結S件採購單<980</span>
    </el-tab-pane>

    <el-tab-pane label="待報關物料" name="customs-materials">
       <el-card shadow="hover">
         <div slot="header" class="d-flex justify-content-between align-items-center">
           
                     <div>
                     <span class="text-primary h5">待報關物料 明細</span>
 
</div>
            <div>
              <el-input v-model="searchText.customs" placeholder="搜尋訂單號、料號、品名..." style="width: 250px; margin-right: 10px;" clearable></el-input>
              <el-button type="success" @click="downloadExcel('customs-materials')"><i class="bi bi-download"></i> 下載 Excel</el-button>
            </div>
         </div>
         <el-table :data="filteredData.customs" v-loading="loading" stripe height="550" style="width: 100%">
            <el-table-column prop="收料日期" label="收料日期" width="120" sortable :formatter="formatDate"></el-table-column>
            <el-table-column prop="訂單號" label="訂單號" width="150" sortable></el-table-column>
            <el-table-column prop="行號" label="行號" width="80" sortable></el-table-column>
            <el-table-column prop="單據類型" label="單據類型" width="100" sortable></el-table-column>
            <el-table-column prop="料號" label="料號" width="160" sortable></el-table-column>
            <el-table-column prop="品名" label="品名" width="250" align="left" header-align="center" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="規格" label="規格" width="250" align="left" header-align="center" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="單位" label="單位" width="80" sortable></el-table-column>
            <el-table-column prop="訂單數量" label="訂單數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="本幣單價" label="本幣單價" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="外幣單價" label="外幣單價" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="供應商名稱" label="供應商名稱" width="250" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="付款方式" label="付款方式" width="120" sortable></el-table-column>
            <el-table-column prop="說明" label="說明" width="200" align="left" header-align="center" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="珠海" label="珠海" width="100" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="在途" label="在途" width="100" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="檢驗中" label="檢驗中" width="100" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="待報關" label="待報關" width="100" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="送貨單號" label="送貨單號" width="150" sortable></el-table-column>
         </el-table>
      </el-card>
       <span class="text-success small d-block mt-1">資料來源：ERP待報關物料</span>
    </el-tab-pane>

    <el-tab-pane label="已挑單未扣庫存" name="picked-not-issued">
      <el-card shadow="hover">
         <div slot="header" class="d-flex justify-content-between align-items-center">

                      <div>
            <span class="text-primary h5">已挑單未扣庫存 明細</span>

</div>
            <div>
              <el-input v-model="searchText.picked" placeholder="搜尋出貨單號、訂單號、料號..." style="width: 250px; margin-right: 10px;" clearable></el-input>
              <el-button type="success" @click="downloadExcel('picked-not-issued')"><i class="bi bi-download"></i> 下載 Excel</el-button>
            </div>
         </div>
         <el-table :data="filteredData.picked" v-loading="loading" stripe height="550" style="width: 100%">
            <el-table-column prop="出貨單號" label="出貨單號" width="160" sortable></el-table-column>
            <el-table-column prop="發運日期" label="發運日期" width="120" sortable :formatter="formatDate"></el-table-column>
            <el-table-column prop="訂單號" label="訂單號" width="160" sortable></el-table-column>
            <el-table-column prop="單據類型" label="單據類型" width="100" sortable></el-table-column>
            <el-table-column prop="行號" label="行號" width="80" sortable></el-table-column>
            <el-table-column prop="料號" label="料號" width="160" sortable align="left" header-align="center"></el-table-column>
            <el-table-column prop="品名" label="品名" width="250" sortable align="left" header-align="center" show-overflow-tooltip></el-table-column>
            <el-table-column prop="規格" label="規格" width="250" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="數量" label="數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="外幣單價" label="外幣單價" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="外幣金額" label="外幣金額" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="本幣單價" label="本幣單價" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="本幣金額" label="本幣金額" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
         </el-table>
      </el-card>
        <span class="text-success small d-block mt-1">資料來源：ERP訂單狀態560</span>
    </el-tab-pane>

    <el-tab-pane label="已扣庫存未結案" name="issued-not-closed">
       <el-card shadow="hover">
         <div slot="header" class="d-flex justify-content-between align-items-center">

                                  <div>
            <span class="text-primary h5">已扣庫存未結案 明細</span>

</div>
            <div>
              <el-input v-model="searchText.issued" placeholder="搜尋出貨單號、訂單號、料號..." style="width: 250px; margin-right: 10px;" clearable></el-input>
              <el-button type="success" @click="downloadExcel('issued-not-closed')"><i class="bi bi-download"></i> 下載 Excel</el-button>
            </div>
         </div>
         <el-table :data="filteredData.issued" v-loading="loading" stripe height="550" style="width: 100%">
            <el-table-column prop="出貨單號" label="出貨單號" width="160" sortable></el-table-column>
            <el-table-column prop="發運日期" label="發運日期" width="120" sortable :formatter="formatDate"></el-table-column>
            <el-table-column prop="訂單號" label="訂單號" width="160" sortable></el-table-column>
            <el-table-column prop="單據類型" label="單據類型" width="100" sortable></el-table-column>
            <el-table-column prop="行號" label="行號" width="80" sortable></el-table-column>
            <el-table-column prop="料號" label="料號" width="160" sortable align="left" header-align="center"></el-table-column>
            <el-table-column prop="品名" label="品名" width="250" sortable align="left" header-align="center" show-overflow-tooltip></el-table-column>
            <el-table-column prop="規格" label="規格" width="250" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="數量" label="數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="外幣單價" label="外幣單價" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="外幣金額" label="外幣金額" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="本幣單價" label="本幣單價" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="本幣金額" label="本幣金額" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
         </el-table>
      </el-card>
        <span class="text-success small d-block mt-1">資料來源：ERP訂單狀態>560~620</span>
    </el-tab-pane>

    <el-tab-pane label="未結FLOW" name="flow-unclosed">
       <el-card shadow="hover">
         <div slot="header" class="d-flex justify-content-between align-items-center">

                                              <div>
            <span class="text-primary h5">未結 FLOW 表單 明細</span>

</div>
            <div>
              <el-input v-model="searchText.flow" placeholder="搜尋表單、文件編號、申請人..." style="width: 250px; margin-right: 10px;" clearable></el-input>
              <el-button type="success" @click="downloadExcel('flow-unclosed')"><i class="bi bi-download"></i> 下載 Excel</el-button>
            </div>
         </div>
         <el-table :data="filteredData.flow" v-loading="loading" stripe height="550" style="width: 100%">
            <el-table-column prop="表單名稱" label="表單名稱" width="250" sortable show-overflow-tooltip></el-table-column>
            <el-table-column prop="文件編號" label="文件編號" width="200" sortable>
              <template slot-scope="scope">
                <a :href="scope.row.文件編號連結" target="_blank">{{ scope.row.文件編號 }}</a>
              </template>
            </el-table-column>
            <el-table-column prop="表單狀態" label="狀態" width="120" sortable></el-table-column>
            <el-table-column prop="簽和關卡" label="關卡" width="150" sortable></el-table-column>
            <el-table-column prop="申請人" label="申請人" width="150" sortable></el-table-column>
            <el-table-column prop="建立日期" label="建立日期" sortable></el-table-column>
         </el-table>
      </el-card>
        <span class="text-success small d-block mt-1">資料來源：FLOW未簽和表單，部門人員(1091910,1091210) </span>
    </el-tab-pane>
    
    <el-tab-pane label="已上線工單缺料表" name="workorder-shortage">
      <el-card shadow="hover">
         <div slot="header" class="d-flex justify-content-between align-items-center">

                                                          <div>
            <span class="text-danger h5">已上線工單缺料表 明細</span>

</div>
            <div>
              <el-input v-model="searchText.shortage" placeholder="搜尋工單號、料號、品名..." style="width: 250px; margin-right: 10px;" clearable></el-input>
              <el-button type="success" @click="downloadExcel('workorder-shortage')"><i class="bi bi-download"></i> 下載 Excel</el-button>
            </div>
         </div>
         <el-table :data="filteredData.shortage" v-loading="loading" stripe height="550" style="width: 100%">
            <el-table-column prop="班別" label="班別" width="80" sortable></el-table-column>
            <el-table-column prop="工單號" label="工單號" width="160" sortable></el-table-column>
            <el-table-column prop="狀態" label="狀態" width="100" sortable></el-table-column>
            <el-table-column prop="母件料號" label="母件料號" width="160" sortable></el-table-column>
            <el-table-column prop="母件品名" label="母件品名" width="250" sortable align="left" header-align="center" show-overflow-tooltip></el-table-column>
            <el-table-column prop="母件規格" label="母件規格" width="250" sortable align="left" header-align="center" show-overflow-tooltip></el-table-column>
            <el-table-column prop="工單數量" label="工單數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="完工數量" label="完工數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="標準工時" label="標準工時" width="120" sortable align="right"></el-table-column>
            <el-table-column prop="已入工時" label="已入工時" width="120" sortable align="right"></el-table-column>
            <el-table-column prop="子件料號" label="子件料號" width="160" sortable></el-table-column>
            <el-table-column prop="子件品名" label="子件品名" width="250" sortable align="left" header-align="center" show-overflow-tooltip></el-table-column>
            <el-table-column prop="子件規格" label="子件規格" width="250" sortable align="left" header-align="center" show-overflow-tooltip></el-table-column>
            <el-table-column prop="應發料量" label="應發料量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="已發數量" label="已發數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="庫存" label="庫存" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="珠海數量" label="珠海數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="在途數量" label="在途數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="檢驗中數量" label="檢驗中數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
            <el-table-column prop="待報關數量" label="待報關數量" width="120" sortable align="right" :formatter="formatNumber"></el-table-column>
         </el-table>
      </el-card>
        <span class="text-success small d-block mt-1">資料來源：ERP工單狀態45-95,有入工時的工單</span>
    </el-tab-pane>

    <el-tab-pane label="異常倉別(一般物料)" name="general-stock">
      <el-card shadow="hover">
         <div slot="header" class="d-flex justify-content-between align-items-center">
                                                          <div>
            <span class="text-danger h5">已上線工單缺料表 明細</span>

</div>
            <div>
              <el-input v-model="searchText.generalStock" placeholder="搜尋料號或品名..." style="width: 200px; margin-right: 10px;" clearable></el-input>
              <el-select v-model="selectedFilter.generalStock" placeholder="篩選欄位" style="width: 180px; margin-right: 10px;" clearable>
                <el-option label="不篩選 (全部)" value=""></el-option>
                <el-option v-for="field in numericFields" :key="field.value" :label="field.label + ' ≠ 0'" :value="field.value"></el-option>
              </el-select>
              <el-button type="success" @click="downloadExcel('general-stock')"><i class="bi bi-download"></i> 下載 Excel</el-button>
            </div>
         </div>
         <el-table :data="filteredData.generalStock" v-loading="loading" stripe height="550" style="width: 100%">
            <el-table-column prop="料號" label="料號" width="150" sortable fixed></el-table-column>
            <el-table-column prop="品名" label="品名" width="250" sortable align="left" header-align="center" show-overflow-tooltip></el-table-column>
            <el-table-column v-for="field in numericFields" :key="field.value" :prop="field.value" :label="field.label" width="140" sortable align="right" header-align="center">
                <template slot-scope="scope">
                    <span :class="{'text-danger fw-bold': scope.row[field.value] > 0}">{{ formatNumber(null, null, scope.row[field.value]) }}</span>
                </template>
            </el-table-column>
         </el-table>
      </el-card>
        <span class="text-success small d-block mt-1">資料來源：ERP非代工料號庫未MRB/PBB/STG/TBD不為0料號</span>
    </el-tab-pane>

    <el-tab-pane label="異常倉別(來料加工)" name="subcontract-stock">
       <el-card shadow="hover">
         <div slot="header" class="d-flex justify-content-between align-items-center">

                                                                      <div>
            <span class="text-primary h5">📦 異常倉別追蹤-來料加工</span>

</div>
            <div>
              <el-input v-model="searchText.subcontractStock" placeholder="搜尋料號或品名..." style="width: 200px; margin-right: 10px;" clearable></el-input>
              <el-select v-model="selectedFilter.subcontractStock" placeholder="篩選欄位" style="width: 180px; margin-right: 10px;" clearable>
                <el-option label="不篩選 (全部)" value=""></el-option>
                <el-option v-for="field in numericFields" :key="field.value" :label="field.label + ' ≠ 0'" :value="field.value"></el-option>
              </el-select>
              <el-button type="success" @click="downloadExcel('subcontract-stock')"><i class="bi bi-download"></i> 下載 Excel</el-button>
            </div>
         </div>
         <el-table :data="filteredData.subcontractStock" v-loading="loading" stripe height="550" style="width: 100%">
            <el-table-column prop="料號" label="料號" width="150" sortable fixed></el-table-column>
            <el-table-column prop="品名" label="品名" width="250" sortable align="left" header-align="center" show-overflow-tooltip></el-table-column>
            <el-table-column v-for="field in numericFields.filter(f => canViewSensitiveColumns() || !['loc_DN', 'loc_OFS', 'loc_DG'].includes(f.value))" :key="field.value" :prop="field.value" :label="field.label" width="140" sortable align="right" header-align="center">
                <template slot-scope="scope">
                    <span :class="{'text-danger fw-bold': scope.row[field.value] > 0}">{{ formatNumber(null, null, scope.row[field.value]) }}</span>
                </template>
            </el-table-column>
         </el-table>
      </el-card>
        <span class="text-success small d-block mt-1">資料來源：ERP工料號庫未MRB/PBB/STG/TBD不為0料號</span>
    </el-tab-pane>

  </el-tabs>

  <!-- 成品／半成品 庫存金額明細（Element UI Dialog） -->
  <el-dialog
    :visible.sync="showStockDetailModal"
    :title="stockDetailKind === 'finished' ? '成品庫存金額明細' : '半成品庫存金額明細'"
    width="90%">
    <div class="mb-3">
      <el-tag type="info" class="mr-2">筆數：{{ stockDetailList().length }}</el-tag>
      <el-tag type="warning" class="mr-2">總數量：{{ formatVND(stockDetailSummary().總數量) }}</el-tag>
      <el-tag type="danger" class="mr-2">金額VND：{{ formatVND(stockDetailSummary().金額VND) }}</el-tag>
      <el-tag type="success">金額USD：{{ formatUSD(stockDetailSummary().金額USD) }}</el-tag>
    </div>
    <el-table :data="stockDetailList()" height="60vh" stripe>
      <el-table-column prop="料號" label="料號" width="160" sortable></el-table-column>
      <el-table-column prop="品名" label="品名" width="280" show-overflow-tooltip></el-table-column>
      <el-table-column prop="規格" label="規格" width="300" show-overflow-tooltip></el-table-column>
      <el-table-column prop="總數量" label="總數量" width="120" align="right" :formatter="formatNumber"></el-table-column>
      <el-table-column prop="成本" label="成本(VND/單位)" width="150" align="right" :formatter="formatNumber"></el-table-column>
      <el-table-column prop="金額VND" label="金額VND" width="150" align="right" :formatter="formatNumber"></el-table-column>
      <el-table-column prop="金額USD" label="金額USD" width="150" align="right">
        <template slot-scope="scope">{{ formatUSD(scope.row.金額USD) }}</template>
      </el-table-column>
    </el-table>
    <span slot="footer" class="dialog-footer">
      <el-button @click="showStockDetailModal=false">關閉</el-button>
    </span>
  </el-dialog>

</div>
`,
  data() {
    return {
      appLoading: true,
      loading: false,
      selectedMetric: 'unclosed-po',

      // 明細 Dialog 狀態
      showStockDetailModal: false,
      stockDetailKind: null, // 'finished' | 'semi'

      rawData: {
        po: [], customs: [], picked: [], issued: [],
        flow: [], shortage: [], generalStock: [], subcontractStock: []
      },

      searchText: {
        po: '', customs: '', picked: '', issued: '',
        flow: '', shortage: '', generalStock: '', subcontractStock: ''
      },

      selectedFilter: {
        generalStock: '', subcontractStock: ''
      },
      
      topMetrics: [
        { id: "unclosed-po", title: "未結S件採購單", value: "0 件", target: "0" },
        { id: "customs-materials", title: "待報關物料", value: "0 件", target: "0" },
        { id: "picked-not-issued", title: "已挑單未扣庫存", value: "0 件", target: "0" },
        { id: "issued-not-closed", title: "已扣庫存未結案", value: "0 件", target: "0" },
        { id: "flow-unclosed", title: "未結FLOW", value: "0 件", target: "0" },
        { id: "workorder-shortage", title: "已上線工單缺料表", value: "0 件", target: "0" },
        { id: "general-stock", title: "異常倉別(一般物料)", value: "0 筆", target: "0" },
        { id: "subcontract-stock", title: "異常倉別(來料加工)", value: "0 筆", target: "0" },
        { id: "finished-usd", title: "成品庫存金額(USD)", value: "$0", route: "general-stock" },
        { id: "semi-usd",     title: "半成品庫存金額(USD)", value: "$0", route: "general-stock" }
      ],

      numericFields: [
        { label: '良品倉(A+B)', value: 'loc' }, { label: 'A良品倉', value: 'loc_A' }, { label: 'B良品倉', value: 'loc_B' },
        { label: '來料短少倉(STG)', value: 'loc_STG' }, { label: '借調倉(LED)', value: 'loc_LED' }, { label: '遺失倉(LOS)', value: 'loc_LOS' },
        { label: '待判定倉(TBD)', value: 'loc_TBD' }, { label: '製損倉(MRB)', value: 'loc_MRB' }, { label: '來料不良倉(PUB)', value: 'loc_PUB' },
        { label: 'DN', value: 'loc_DN' }, { label: 'OFS', value: 'loc_OFS' }, { label: 'DG', value: 'loc_DG' },
      ],
    };
  },
  computed: {
    filteredData() {
      const createFilter = (data, searchText, searchFields) => {
        if (!searchText) return data;
        const keyword = searchText.trim().toLowerCase();
        return data.filter(row => 
          searchFields.some(field => 
            row[field] && String(row[field]).toLowerCase().includes(keyword)
          )
        );
      };
      
      const filteredGeneralStock = () => {
        let data = createFilter(this.rawData.generalStock, this.searchText.generalStock, ['料號', '品名']);
        if (this.selectedFilter.generalStock) {
          data = data.filter(r => Number(r[this.selectedFilter.generalStock]) !== 0);
        }
        return data;
      };

      const filteredSubcontractStock = () => {
        let data = createFilter(this.rawData.subcontractStock, this.searchText.subcontractStock, ['料號', '品名']);
        if (this.selectedFilter.subcontractStock) {
          data = data.filter(r => Number(r[this.selectedFilter.subcontractStock]) !== 0);
        }
        return data;
      };

      return {
        po: createFilter(this.rawData.po, this.searchText.po, ['單據號', '料號', '品名', '供應商名稱']),
        customs: createFilter(this.rawData.customs, this.searchText.customs, ['訂單號', '料號', '品名', '供應商名稱']),
        picked: createFilter(this.rawData.picked, this.searchText.picked, ['出貨單號', '訂單號', '料號']),
        issued: createFilter(this.rawData.issued, this.searchText.issued, ['出貨單號', '訂單號', '料號']),
        flow: createFilter(this.rawData.flow, this.searchText.flow, ['表單名稱', '文件編號', '申請人']),
        shortage: createFilter(this.rawData.shortage, this.searchText.shortage, ['工單號', '母件料號', '母件品名', '子件料號', '子件品名']),
        generalStock: filteredGeneralStock(),
        subcontractStock: filteredSubcontractStock(),
      };
    },
  },
  methods: {
    // ===== 小卡點擊行為：成品/半成品顯示明細，其餘維持切換頁籤 =====
    handleMetricClick(item) {
      if (item.id === 'finished-usd') {
        this.stockDetailKind = 'finished';
        this.showStockDetailModal = true;
        this.selectedMetric = (item.route || 'general-stock');
        return;
      }
      if (item.id === 'semi-usd') {
        this.stockDetailKind = 'semi';
        this.showStockDetailModal = true;
        this.selectedMetric = (item.route || 'general-stock');
        return;
      }
      this.selectedMetric = (item.route || item.id);
    },

    // 料號第一碼為 '1' 視為成品
    isFinishedRow(row) {
      const no = (row?.料號 || '').toString().trim();
      return no.startsWith('1');
    },

    // 明細列表（依目前卡片類型過濾），並映射成指定欄位
    stockDetailList() {
      if (!Array.isArray(this.rawData.generalStock)) return [];
      const mapped = this.rawData.generalStock.map(r => ({
        料號: r.料號,
        品名: r.品名,
        規格: r.規格,
        總數量: Number(r.總庫存數量 || 0),
        成本: Number(r.標準成本_VND || 0),     // 單位成本 (VND)
        金額VND: Number(r.庫存金額_VND || 0),
        金額USD: Number(r.庫存金額_USD || 0)
      }));
      if (this.stockDetailKind === 'finished') return mapped.filter(this.isFinishedRow);
      if (this.stockDetailKind === 'semi') return mapped.filter(x => !this.isFinishedRow(x));
      return mapped;
    },

    // 合計（顯示於 Dialog 上方）
    stockDetailSummary() {
      const list = this.stockDetailList();
      const sum = (k) => list.reduce((acc, x) => acc + Number(x[k] || 0), 0);
      return {
        總數量: sum('總數量'),
        金額VND: sum('金額VND'),
        金額USD: sum('金額USD')
      };
    },

    getIcon(title) {
      if (title.includes("採購")) return "bi bi-cart-check";
      if (title.includes("報關")) return "bi bi-truck";
      if (title.includes("挑單") || title.includes("扣庫存")) return "bi bi-box-arrow-down";
      if (title.includes("FLOW")) return "bi bi-envelope-exclamation";
      if (title.includes("缺料")) return "bi bi-tools text-danger";
      if (title.includes("異常倉")) return "bi bi-archive";
      if (title.includes("庫存金額")) return "bi bi-cash-coin text-success";
      return "bi bi-circle";
    },
    formatUSD(n) {
      if (n == null || isNaN(Number(n))) return "$0";
      return "$" + Number(n).toLocaleString(undefined, { maximumFractionDigits: 2 });
    },
    formatVND(n) {
      if (n == null || isNaN(Number(n))) return "0";
      return Number(n).toLocaleString(undefined, { maximumFractionDigits: 0 });
    },
    formatDate(row, column, cellValue) {
      return this.formatDateString(cellValue);
    },
    formatDateString(val) {
      if (!val) return '';
      let token = String(val).trim();

      // 去除中文上午/下午與時間
      token = token.replace(/\s*(上午|下午)\s*\d{1,2}:\d{2}:\d{2}.*/g, '');
      token = token.split(' ')[0];     // 2025/3/28 00:00:00 -> 2025/3/28
      token = token.split('T')[0];     // 2025-03-28T00:00:00 -> 2025-03-28

      if (/^\d{8}$/.test(token)) {     // 20250328
        const y = token.slice(0, 4);
        const m = Number(token.slice(4, 6));
        const d = Number(token.slice(6, 8));
        return `${y}/${m}/${d}`;
      }
      if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(token)) {
        const parts = token.includes('-') ? token.split('-') : token.split('/');
        const [y, m, d] = [parts[0], Number(parts[1]), Number(parts[2])];
        return `${y}/${m}/${d}`;
      }
      return token;
    },
    // 指定欄位強制套用「yyyy/m/d」格式（把字串轉為日期序號並設 z）
    applyDateFormat(worksheet, headerNames, fmt = 'yyyy/m/d') {
      if (!worksheet || !worksheet['!ref']) return;
      const range = XLSX.utils.decode_range(worksheet['!ref']);

      // 讀表頭：建立「欄名 -> 欄索引」對照
      const headerIndex = {};
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const addr = XLSX.utils.encode_cell({ r: range.s.r, c: C });
        const cell = worksheet[addr];
        if (cell && cell.v != null) {
          headerIndex[String(cell.v).trim()] = C;
        }
      }

      headerNames.forEach(name => {
        const col = headerIndex[String(name).trim()];
        if (col == null) return;

        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
          const addr = XLSX.utils.encode_cell({ r: R, c: col });
          const cell = worksheet[addr];
          if (!cell || cell.v == null) continue;

          // 先把任何值轉成 YYYY/M/D
          let pure = '';
          if (cell.t === 's' || cell.t === 'str') {
            pure = this.formatDateString(cell.v);
          } else if (cell.t === 'n' && typeof cell.v === 'number' && cell.z) {
            // 已是數字 + 可能有時間格式，直接改 z
            cell.z = fmt;
            continue;
          } else {
            pure = this.formatDateString(String(cell.v));
          }

          if (pure) {
            const [y, m, d] = pure.split('/').map(Number);
            if (!isNaN(y) && !isNaN(m) && !isNaN(d)) {
              const dObj = new Date(y, m - 1, d);
              dObj.setHours(0, 0, 0, 0);
              const serial = this.toExcelDateSerial(dObj);
              cell.t = 'n';          // 數字型（Excel 日期序號）
              cell.v = serial;       // 真正的日期值
              cell.z = fmt;          // 顯示成 yyyy/m/d（不會顯示時間）
            } else {
              // 兜不出有效日期時，保留純文字避免 Excel 自動轉出時間
              cell.t = 's';
              cell.v = pure;
            }
          }
        }
      });
    },

    // 轉成 Excel 的日期序號（時分秒會歸零）
    toExcelDateSerial(dateObj) {
      const utc = Date.UTC(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate());
      const excelEpoch = Date.UTC(1899, 11, 30); // Excel 起算日 1899-12-30
      return (utc - excelEpoch) / 86400000;
    },
    formatNumber(row, column, cellValue) {
      if (cellValue === null || cellValue === undefined) return '';
      const num = Number(cellValue);
      return isNaN(num) ? '' : num.toLocaleString();
    },
    canViewSensitiveColumns() {
      return true; 
    },

    async fetchData(apiCall, dataKey, metricIndex, unit = '件') {
      try {
        const res = await apiCall();
        this.rawData[dataKey] = res.data;
        this.topMetrics[metricIndex].value = `${res.data.length} ${unit}`;
      } catch (err) {
        this.$message.error(`載入 ${this.topMetrics[metricIndex].title} 失敗`);
        console.error(`Error fetching ${dataKey}:`, err);
      }
    },
    async fetchAllData() {
      this.appLoading = true;
      this.loading = true;

      const apiCalls = [
        this.fetchData(() => axios.get("https://mms.leapoptical.com:5088/api/Materials/unclosed-po?company=00109&docType=OM"), 'po', 0),
        this.fetchData(() => axios.get("https://mms.leapoptical.com:5088/api/Materials/customs-materials"), 'customs', 1),
        this.fetchData(() => axios.get("https://mms.leapoptical.com:5088/api/Materials/picked-not-issued?company=00109"), 'picked', 2),
        this.fetchData(() => axios.get("https://mms.leapoptical.com:5088/api/Materials/issued-not-closed?company=00109"), 'issued', 3),
        this.fetchData(async () => {
          const depts = ['1091210', '109110', '1091510', '1091910', '1092200'];
          const requests = depts.map(d => axios.get('https://mms.leapoptical.com:5088/api/LinePerformance/unclosedflow/list', { params: { dept: d } }));
          const responses = await Promise.all(requests);
          return { data: responses.flatMap(res => res.data) };
        }, 'flow', 4),
        this.fetchData(() => axios.get("https://mms.leapoptical.com:5088/api/Materials/shortage?factory=%20%20%20%20%201099000"), 'shortage', 5),
        this.fetchData(() => axios.get("https://mms.leapoptical.com:5088/api/Materials/general-stock"), 'generalStock', 6, '筆'),
        this.fetchData(() => axios.get("https://mms.leapoptical.com:5088/api/Materials/subcontract-stock"), 'subcontractStock', 7, '筆'),
      ];
      
      await Promise.allSettled(apiCalls);
      
      // 成品 / 半成品 庫存金額(USD) 小卡
      try {
        const { data } = await axios.get("https://mms.leapoptical.com:5088/api/Materials/general-stock-usd-split");
        const idxFinished = this.topMetrics.findIndex(x => x.id === "finished-usd");
        const idxSemi     = this.topMetrics.findIndex(x => x.id === "semi-usd");
        if (idxFinished >= 0) this.topMetrics[idxFinished].value = this.formatUSD(data?.finishedUSD ?? 0);
        if (idxSemi >= 0)     this.topMetrics[idxSemi].value     = this.formatUSD(data?.semiUSD ?? 0);
      } catch (e) {
        console.error("載入 general-stock-usd-split 失敗", e);
      }

      this.loading = false;
      this.appLoading = false;
    },

    downloadExcel(type) {
      const dataMap = {
        'unclosed-po':        { data: this.filteredData.po,              filename: '未結S件採購單.xlsx' },
        'customs-materials':  { data: this.filteredData.customs,         filename: '待報關物料.xlsx' },
        'picked-not-issued':  { data: this.filteredData.picked,          filename: '已挑單未扣庫存.xlsx' },
        'issued-not-closed':  { data: this.filteredData.issued,          filename: '已扣庫存未結案.xlsx' },
        'flow-unclosed':      { data: this.filteredData.flow,            filename: '未結FLOW表單.xlsx' },
        'workorder-shortage': { data: this.filteredData.shortage,        filename: '已上線工單缺料表.xlsx' },
        'general-stock':      { data: this.filteredData.generalStock,    filename: '異常倉別_一般物料.xlsx' },
        'subcontract-stock':  { data: this.filteredData.subcontractStock,filename: '異常倉別_來料加工.xlsx' },
      };

      // 各頁籤需要轉成「日期欄」的欄位
      const dateFieldsByType = {
        'unclosed-po':       ['請購日期'],  // ⭐ 僅日期不要時間
        'customs-materials': ['收料日期'],
        'picked-not-issued': ['發運日期'],
        'issued-not-closed': ['發運日期'],
        'flow-unclosed':     ['建立日期'],
      };

      const entry = dataMap[type];
      if (!entry) return this.$message.warning('未知的匯出類型');
      const { data, filename } = entry;

      if (!data || !data.length) {
        return this.$message.warning('沒有可匯出的資料（可能被篩選為空）');
      }

      const fields = dateFieldsByType[type] || [];

      // 1) 先把指定欄位清成 YYYY/M/D（避免帶出「上午/下午」）
      const exportRows = data.map(row => {
        const copy = { ...row };
        fields.forEach(f => {
          if (f in copy) {
            const pure = this.formatDateString(copy[f]);
            copy[f] = pure || '';
          }
        });
        return copy;
      });

      // 2) 產出工作表
      const worksheet = XLSX.utils.json_to_sheet(exportRows, { cellDates: false }); // 先用字串

      // 3) 將指定欄位改成「真正的 Excel 日期」並套用 yyyy/m/d 顯示格式
      this.applyDateFormat(worksheet, fields, 'yyyy/m/d');

      // 4) 輸出
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Data');
      XLSX.writeFile(workbook, filename);
    },

  },
  mounted() {
    this.fetchAllData();
  }
});
