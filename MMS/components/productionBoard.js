(function loadCSS(href) {
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = href;
  document.head.appendChild(link);
})("assets/production-board.css");



Vue.component('production-board-view', {
  template: `
<div class="production-board-page">
 <div class="section-content d-flex flex-wrap align-items-center">

        日期<input type="date" v-model="selectedDate" @change="handleInputChange" >
        <label >{{ translatedText.area }}:</label>
        <select id="areaSelect" v-model="selectedArea" @change="handleInputChange">
          <option value="ZH">ZH</option>
          <option value="TC">TC</option>
          <option value="VN">VN</option>
          <option value="TW">TW</option>
        </select>
    
        <label >{{ translatedText.部門 }}:</label>
        <select id="DSelect" v-model="selectedD" @change="handleInputChange">
            <option value="1">1</option>                                                     
            <option value="2">2</option>                                                                                 
            <option value="3">3</option> 
            <option value="4">4</option>
            <option value="6">6</option>
          </select>
        <label >{{ translatedText.D }}</label>
        <select id="SSelect" v-model="selectedS" @change="handleInputChange">
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
          </select>
        <label >{{ translatedText.S }}</label>
        <select id="CSelect" v-model="selectedC" @change="handleInputChange">
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
            <option value="6">6</option>
            <option value="7">7</option>
            <option value="8">8</option>
            <option value="9">9</option>
            <option value="10">10</option>
            <option value="S1">S1</option>
            <option value="S2">S2</option>
            <option value="S3">S3</option>
            <option value="S4">S4</option>
            <option value="S5">S5</option>
            <option value="S6">S6</option>
          </select>
    
        <label >{{ translatedText.C }}</label>
        <select id="LSelect" v-model="selectedL" @change="handleInputChange">
            <option value=""></option>
            <option value="1">1</option>
            <option value="2">2</option>
            <option value="3">3</option>
            <option value="4">4</option>
            <option value="5">5</option>
            <option value="6">6</option>
            <option value="7">7</option>
          </select>
        <label >{{ translatedText.L }}</label>
      </div>
 <div class="right-content">
   
  <hr></hr>

  <template >
    <div style="margin-bottom: 10px;">

  <!-- 合併欄位開關 -->
  <el-switch
    v-model="isMergeEnabled"
    :active-text="translatedText.欄位合併" /></el-switch> 
  <el-switch v-model="showline" :active-text="translatedText.做線"></el-switch>
  <el-switch v-model="showdefectForming" :active-text="translatedText.成型S"></el-switch>
  <el-switch v-model="showDefectWelding" :active-text="translatedText.焊接"></el-switch>

      <el-switch v-model="table01" :active-text="translate(selectedProcess1)"></el-switch>  
      <el-switch v-model="table02" :active-text="translate(selectedProcess2)"></el-switch>  
      <el-switch v-model="table03" :active-text="translate(selectedProcess3)"></el-switch>  
      <el-switch v-model="table04" :active-text="translate(selectedProcess4)"></el-switch>  
      <el-switch v-model="table05" :active-text="translate(selectedProcess5)"></el-switch>  
      <el-switch v-model="table06" :active-text="translate(selectedProcess6)"></el-switch>  

  </div>

<!--<el-table :data="tableData" border style="width: 100%; margin-top: 10px;" :span-method="mergeTable">-->
  <el-table :data="tableData" border  :span-method="mergeTable" :summary-method="getSummaries" show-summary  :row-class-name="tableRowClassName">



      <el-table-column v-if="editMode" label="操作"    min-width="40px"  >
          <template slot-scope="scope">
              <el-button type="danger" icon="el-icon-delete" size="mini" @click="removeRow2(scope.row,scope.$index)"></el-button>
          </template>
      </el-table-column>
      <el-table-column  prop="time" id="time"  min-width="100px"  >
        <template slot="header"  >
          <span >時段<br> thời gian</span>
      </template>
          <template slot-scope="scope">
              <el-select v-model="scope.row.time" :disabled="!editMode"  @input="resetExpected(scope.row, scope.$index)"style="width: 100%;" :class="{ 'edit-mode': editMode }" >
                  <el-option 
                      v-for="[label, value] in Object.entries(timeOptions)" 
                      :key="label" 
                      :label="label" 
                      :value="label">
                  </el-option>
              </el-select>
          </template>
      </el-table-column>

      <el-table-column   prop="people" id="people"   min-width="60px">
        <template slot="header">
          <span width="10px">人數<br><small>Số người</small></span>
      </template>
          <template slot-scope="scope">
              <el-input  v-model.number="scope.row.people"  @input="resetExpected(scope.row, scope.$index)" :disabled="!editMode || !scope.row.time" type="number" min="1"  :class="{ 'edit-mode': editMode } "   ></el-input>
          </template>
      </el-table-column>
  
  
      <el-table-column  prop="wohourse" id="wohourse"   min-width="60px" align="center">
        <template slot="header">
          <span>總時數<br><small>Tổng số giờ</small></span>
      </template>
          <template slot-scope="scope">
            {{ (calculateHours(scope.row)).toFixed(2) }}
          </template>
      </el-table-column>
  
      <el-table-column prop="product" id="product"  min-width="150px">
        <template slot="header">
          <span>品名<br>Tên sản phẩm</span>
      </template>
        <template slot-scope="scope">
            <el-select v-model="scope.row.product" :disabled="!editMode" style="width: 100%;"  @input="resetExpected(scope.row, scope.$index)":class="{ 'edit-mode': editMode }" style="width: 100%;">
                <el-option 
                    v-for="item in productOptions" 
                    :key="item.value" 
                    :label="item.label" 
                    :value="item.value">
                </el-option>
            </el-select>
        </template>
    </el-table-column>

      <el-table-column  prop="workhorse" id="workhorse" min-width="80px">
        <template slot="header">
          <span>輸入工時<br>Nhập giờ<br>làm việc</span>
      </template>
          <template slot-scope="scope">
              <el-input 
              v-model.number="scope.row.workhorse"  
              @input="resetExpected(scope.row, scope.$index)" 
              @blur="formatDecimal(scope.row, 'workhorse')" 
              :disabled="!editMode" 
              type="number" 
              step="0.1" 
              min="0.0"
              :class="{ 'edit-mode': editMode }">
          </el-input>
          </template>
      </el-table-column>

      <!-- <el-table-column label="預期產量">
          <template slot-scope="scope">
              {{ calculateExpected(scope.row, scope.$index) }}
          </template>
      </el-table-column>-->
  
      <el-table-column   prop="expected" id="expected" min-width="80px">
        <template slot="header">
          <span>標準產能<br>Số lượng <br>tiêu chuẩn</span>
      </template>
          <template slot-scope="scope">
              <!-- 使用 v-model 綁定 scope.row.expected，並觸發 @blur 檢查手動修改 -->
              <el-input 
                  v-model.number="scope.row.expected" 
                  :disabled="true" 
                  type="number" 
                  min="0" 
                  @blur="updateExpected(scope.row,scope.$index)"
                  :class="{ 'edit-mode': editMode }">
              </el-input>
          </template>
      </el-table-column>
      <el-table-column  prop="actual" id="actual" min-width="100px">
        <template slot="header">
          <span>實際產出<br>Số lượng<br>thực tế</span>
      </template>
          <template slot-scope="scope">
              <el-input v-model.number="scope.row.actual" :disabled="!editMode" type="number" min="0" :class="{ 'edit-mode': editMode }"></el-input>
          </template>
      </el-table-column>
      <el-table-column prop="efficiency" id="efficiency" min-width="70px">
        <template slot="header">
          <span class="nowrap-text">效率<br>hiệu quả</span>
        </template>
        <template slot-scope="scope">
          <span :class="{
            'low-efficiency': Number(calculateEfficiency(scope.row)) < 90 
          }">
            {{ calculateEfficiency(scope.row) }}%
          </span>
        </template>
      </el-table-column>  
      <el-table-column  prop="NG" id="NG"  min-width="70px" >
        <template slot="header">
          <span>不良品<br>Hàng NG</span>
      </template>
          <template slot-scope="scope">
              <el-input v-model.number="scope.row.NG" :disabled="!editMode" type="number" min="0" :class="{ 'edit-mode': editMode }"></el-input>
          </template>
      </el-table-column>
      <el-table-column  prop="line" id="line"  min-width="70px" class="no-space"  v-if="showline">
        <template slot="header">
          <span  class="nowrap-text">做線<br>làm dây</span>
      </template>
          <template slot-scope="scope">
              <el-input v-model="scope.row.line" :disabled="!editMode" type="number" min="0" :class="{ 'edit-mode': editMode }"></el-input>
          </template>
      </el-table-column>
         <!-- 根據開關狀態顯示/隱藏 -->
      <el-table-column  prop="defectForming" id="defectForming"  min-width="70px" v-if="showdefectForming">
          <template slot="header">
            <span>成型<br>thành hình</span>
        </template>
          <template slot-scope="scope">
              <el-input v-model.number="scope.row.defectForming" :disabled="!editMode" type="number" min="0" :class="{ 'edit-mode': editMode }"></el-input>
          </template>
      </el-table-column>
      <el-table-column prop="defectWelding" id="defectWelding"  min-width="70px" v-if="showDefectWelding">
        <template slot="header"  >
          <span>焊接<br>Hàn</span>
      </template>
          <template slot-scope="scope">
              <el-input v-model.number="scope.row.defectWelding" :disabled="!editMode" type="number" min="0" :class="{ 'edit-mode': editMode }"></el-input>
          </template>
      </el-table-column>


      <el-table-column
      :render-header="renderHeaderWithSelect1"
      v-if="table01"
      min-width="70px"
      prop="other1"
    >
      <template slot-scope="scope">
        <el-input
          v-model="scope.row.other1"
          :disabled="!editMode"
          type="number"
          min="0"
          :class="{ 'edit-mode': editMode }"
        ></el-input>
      </template>
    </el-table-column>
    
    <el-table-column
      :render-header="renderHeaderWithSelect2"
      v-if="table02"
      min-width="70px"
      prop="other2"
    >
      <template slot-scope="scope">
        <el-input
          v-model.number="scope.row.other2"
          :disabled="!editMode"
          type="number"
          min="0"
          :class="{ 'edit-mode': editMode }"
        ></el-input>
      </template>
    </el-table-column>
    
    <el-table-column
      :render-header="renderHeaderWithSelect3"
      v-if="table03"
      min-width="70px"
      prop="other3"
    >
      <template slot-scope="scope">
        <el-input
          v-model.number="scope.row.other3"
          :disabled="!editMode"
          type="number"
          min="0"
          :class="{ 'edit-mode': editMode }"
        ></el-input>
      </template>
    </el-table-column>
    <el-table-column
    :render-header="renderHeaderWithSelect4"
    v-if="table04"
    min-width="70px"
    prop="other4"
  >
    <template slot-scope="scope">
      <el-input
        v-model.number="scope.row.other4"
        :disabled="!editMode"
        type="number"
        min="0"
        :class="{ 'edit-mode': editMode }"
      ></el-input>
    </template>
  </el-table-column>
  <el-table-column
  :render-header="renderHeaderWithSelect5"
  v-if="table05"
  min-width="70px"
  prop="other5"
>
  <template slot-scope="scope">
    <el-input
      v-model.number="scope.row.other5"
      :disabled="!editMode"
      type="number"
      min="0"
      :class="{ 'edit-mode': editMode }"
    ></el-input>
  </template>
</el-table-column>
<el-table-column
:render-header="renderHeaderWithSelect6"
v-if="table06"
min-width="70px"
prop="other6"
>
<template slot-scope="scope">
  <el-input
    v-model.number="scope.row.other6"
    :disabled="!editMode"
    type="number"
    min="0"
    :class="{ 'edit-mode': editMode }"
  ></el-input>
</template>
</el-table-column>

  
  
  
  
  
  </el-table>
  
  <el-collapse v-model="activeCollapse" class="mt-3" v-if="Object.keys(productSummary).length > 0">
    <el-collapse-item :title="translatedText.依產品分類加總資料" name="1">
      <el-table :data="Object.entries(productSummary)" border size="mini">
        <el-table-column :label="translatedText.品名"  align="center">
          <template slot-scope="scope">
            {{ scope.row[0] }}
          </template>
        </el-table-column>

        <el-table-column :label="translatedText.輸入工時" align="center">
          <template slot-scope="scope">
            {{ scope.row[1].workhorse.toFixed(2) }}
          </template>
        </el-table-column>
        <el-table-column :label="translatedText.標準產能"  align="center">
          <template slot-scope="scope">
            {{ scope.row[1].expected }}
          </template>
        </el-table-column>
        <el-table-column :label="translatedText.實際產出" align="center">
          <template slot-scope="scope">
            {{ scope.row[1].actual }}
          </template>
        </el-table-column>
        <el-table-column :label="translatedText.不良品" align="center">
          <template slot-scope="scope">
            {{ scope.row[1].NG }}
          </template>
        </el-table-column>
        <el-table-column :label="translatedText.不良率" align="center">
          <template slot-scope="scope">
            <span :class="{
              'low-efficiency': Number((scope.row[1].NG / scope.row[1].actual) * 100)  >= 3 
            }">
              {{
                scope.row[1].actual > 0
                  ? ((scope.row[1].NG / scope.row[1].actual) * 100).toFixed(2) + '%'
                  : '0.00%'
              }}
            </span>
          </template>
        </el-table-column>
        
        <el-table-column :label="translatedText.效率" align="center">
          <template slot-scope="scope">
            <span :class="{
              'low-efficiency': Number(scope.row[1].efficiency) < 90 
            }">
              {{ scope.row[1].efficiency }}%
            </span>
            

          </template>
        </el-table-column>
      </el-table>
    </el-collapse-item>
  </el-collapse>
  
  
<hr>
<h5>生產工單</h5>
<el-table :data="WOVERSON" border stripe size="mini" style="margin-top: 10px;">
  <el-table-column prop="工單" label="工單" min-width="100"></el-table-column>
  <el-table-column prop="品名" label="品名" min-width="200"></el-table-column>
  <el-table-column prop="數量" label="數量" min-width="70" align="right"></el-table-column>
  <el-table-column prop="標準工時" label="標準工時" min-width="90" align="right"></el-table-column>
  <el-table-column prop="執行工時" label="執行工時" min-width="90" align="right"></el-table-column>
</el-table>

    </template>
   </div> 
       </div> 
  `,
  props: ['selectedDept'],
  data() {
    return {

      isMergeEnabled: true, // ✅ 這是開關，true=合併、false=不合併
      WOVERSON: [],
      selectedProcess1: '欄位1',
      selectedProcess2: '欄位2',
      selectedProcess3: '欄位3',
      selectedProcess4: '欄位4',
      selectedProcess5: '欄位5',
      selectedProcess6: '欄位6',
      processSwitch: {
  
          前處理:false,
          插件:false,
          组装:false,
          ATE:false,
          打扭:false,
          套套管:false,
          內模:false,
          外模: false,
          CRY:false,
          外觀: false,
          待包裝:false,
        },
      activeCollapse: ['1'],  // 折疊開關（預設展開）
      showSummaryTable: false, // 開關控制加總表格
      summaryData: [],
      tableDataRaw: [],  
      EFFOVER:[],
      summaryByProductRows:[],
       isOpen: true,
      editMode: false,
      currentSheet: 0, // 預設顯示第1個工作表
      selectedArea :'VN',
      selectedDate: new Date().toISOString().substr(0, 10),
      selectedD:'2',
      selectedS:'1',
      selectedC:'1',
      selectedL:'',
      showline:true,
      translatedText: {
        依產品分類加總資料:'依產品分類加總資料',
        依產品分類加總資料:'依產品分類加總資料',
        欄位合併:'欄位合併',
        輸入工時:'輸入工時',
        標準產能:'標準產能',
        實際產出:'實際產出',
        不良品:'不良品',
        不良率:'不良率',
        效率:'效率',
        前處理:'前處理',
          插件:'插件',
          组装:'组装',
        總數量:'總數量',
        一般:'一般',
        焊接:'焊接',
        成型:'成型',
        做線:'做線',
        焊接S:'焊接',
        成型S:'成型',
        做線S:'做線',
        在線生產看板: '在線生產看板',
        每日生產看版:'每日生產看版',
        欄位合併:'欄位合併',
        星級完成數量:'重點工站完成數量',
        工單管理: '工單管理',
        工單查詢: '工單查詢',
        工站管理: '工站管理',
        工站設定:'工站設定',
        人員管理: '人員管理',
        考勤異常: '考勤異常',
        生產看版: '生產看版',
        工單分析圖表: '工單分析圖表',
        未結工單: '未結工單',
       完工數量:'完工數量',
        生產日報:'生產日報',
        產能日報:'產能日報',
        卡片設定:'卡片設定',
        考勤日報:'考勤日報',
        登入:'登入',
        NFC_SYSTEAM:'NFC 系統',
        考勤系統:'考勤系統',
        生產數據:'生產數據',
        area:'廠區',
        部門:'部門',
        D:'部',
        S:'課',
        C:'班',
        L:'拉',
        在線人數:'在線人數',
        應出席人數:'應出席人數',
        實際出席人數:'實際出席人數',
        請假人數:'請假人數',
       流水線在線人數:'流水線在線人數',
        借入人數:'借入人數',
        借出人數:'借出人數',
        工單列表:'工單列表',
        ID:'編號',
        工單:'工單',
        品名:'品名',
        數量:'數量',
        標準工時:'標準工時',
        執行工時:'執行工時',
        當前效率:'當前效率',
        過往效率:'過往效率',
        已完成數量:'已完成數量',
        設備:'設備',
        姓名:'姓名',
        工時:'工時',
        工站:'工站',
        座位:'座位',
        工站列表:'工站列表',
        功能:'功能',
        fqcer1:'輸入值不能小於原值',
        fqcer2:'輸入值不能大於工單數量',
        完成:'完成',
        暫停:'暫停',
        復工:'復工',
        輸入工單:'輸入工單',
        查詢:'查詢',
        提交:'提交',
        workOrder1:'100:試產工單, 200:來料全檢, 400:重工,700 物料員工時',
        貼紙班:'貼紙班',
        裁剝班:'裁線班',
        料號:'料號',
        設定工單列表:'設定工單列表',
        移除:'移除',
        總和:'總和',
        choose:'選擇NFC_CODE',
        修改:'修改',
        取消:'取消',
        增加座位:'增加座位',
        刪除座位:'刪除座位',
        卡機號:'卡機號',
        工站名稱:'工站名稱',
        insert:'插入',
        delete:'刪除',
        New:'新增',
        出勤:'出勤',
        工作拉:'工作拉',
        工號:'工號',
        Status:'狀態',
        事由:'事由',
        年休:'年休',
        請假:'請假',
        哺乳假:'哺乳假',
        工傷:'工傷',
        工傷陪護:'工傷陪護',
        借出:'借出',
        新進員工:'新進員工',
        夜班:'夜班',
        考勤日報:'考勤日報',
        show:'顯示',
        出勤人數:'出勤人數',
        實際掛卡工時:'實際掛卡工時',
        總共考勤工時:'總共考勤工時',
        回報管理部考勤工時:'回報管理部考勤工時', 
        差異:'差異', 
        原因:'原因',
        輸入卡機:'輸入卡機',
        NFC卡號:'NFC卡號',
        NFC_CODE:'NFC_CODE',
        刷新:'刷新',
        在線工單:'在線工單',
        平均效率:'平均效率',
        低於90工單數:'低於90%工單數',
        已完成工單:'已完成工單',
        未結案工單:'未結案工單',
        無效工時:'無效工時(hrs)',
        本日:'本日',
        本月:'本月',
        上月:'上月',
        pd1:'備註：(本月、上月)資料來源ERP',
        應報考勤:'應報考勤',
        加總:'加總',
        導出:'導出',
        系統計算考勤工時:'系統計算考勤工時',
        增加有效考勤工時:'增加有效考勤工時',
        正班工時:'正班工時',
        加班工時:'加班工時',
        假別:'假別',
        更新完成:'更新完成',
        更新失敗:'更新失敗',
        重複輸入:'输入值重复，请重新输入！',
        查無卡號:'查無卡號',
        工單已存在:'该工單號已存在于表格中。',
        未添加工單:'未添加工單',
        查無工單:'查無工單，是否手動增加？',
        請輸暫停理由:'請輸暫停理由：',
        未生產:'未生產',
        調整:'調整',
        忘記設定:'忘記設定',
        忘記設定狀態:'忘記設定狀態',
        工作部門:'工作部門',
        備註:'備註',
        考勤調整確認:'考勤調整確認',
        日期:'日期',
        班別管理:'班別管理',
        工單狀態與ERP對比報表:'工單狀態與ERP對比報表',
        ERP入庫數量:'ERP入庫數量',
        過站管理:'過站管理',
        序號管理:'序號管理',
        序號查詢:'序號查詢',
        出貨序號表:'出貨序號表',
        不良品紀錄查詢:'不良品紀錄查詢',
        工站編號:'工站編號',
        綁定工站:'綁定工站',
        測試工站:'測試工站',
        輸入開始序號:'輸入開始序號',


    },
    query: {
      dept: '',
      // 其他查詢條件...
    },
    showline:true,
    showdefectForming: true,   // 控制「成形」顯示/隱藏
    showDefectWelding: true,   // 控制「焊接」顯示/隱藏
    table01: false,   // 控制「焊接」顯示/隱藏
    table02: false,   // 控制「焊接」顯示/隱藏
    table03: false,   // 控制「焊接」顯示/隱藏
    table04: false,   // 控制「焊接」顯示/隱藏
    table05: false,   // 控制「焊接」顯示/隱藏
    table06: false,   // 控制「焊接」顯示/隱藏
            productOptions: [], // **存放品名列表**
    tableData: [  {            
      ID: "",
      AREA: "",
      MFG_DAY: "",
      DEPARTMENT: "",
      people: "",
      time: "",
      product: "",
      expected: "",
      actual: "",
      line: "",
      defectForming: "",
      defectWelding: "",
      NG: "",
      other1: "",
      other2: "",
      other3: "",
      other4: "",
      other5: "",
      other6: ""
    }
            ],
            timeOptions: {
    },
  };
},
computed: {
  productSummary() {
    const summary = {};
  
    this.tableData.forEach(row => {
      const product = row.product || '未指定品名';
      if (!summary[product]) {
        summary[product] = {
          people: 0,
          workhorse: 0.00,
          wohourse: 0.00,
          actual: 0,
          expected: 0,
          NG: 0,
          defectForming: 0,
          defectWelding: 0,
          other1: 0,
          other2: 0,
          other3: 0,
          other4: 0,
          other5: 0,
          other6: 0,
          efficiency: 0 // 先預設為 0，最後再計算
        };
      }
  
      const data = summary[product];
  
      data.people += Number(row.people || 0);
      data.workhorse += Number(row.workhorse || 0);
      data.wohourse += Number(row.wohourse || 0);
      data.actual += Number(row.actual || 0);
      data.expected += Number(row.expected || 0);
      data.NG += Number(row.NG || 0);
      data.defectForming += Number(row.defectForming || 0);
      data.defectWelding += Number(row.defectWelding || 0);
      data.other1 += Number(row.other1 || 0);
      data.other2 += Number(row.other2 || 0);
      data.other3 += Number(row.other3 || 0);
      data.other4 += Number(row.other4 || 0);
      data.other5 += Number(row.other5 || 0);
      data.other6 += Number(row.other6 || 0);
    });
  
    // 計算效率
    for (const product in summary) {
      const item = summary[product];
      item.efficiency = item.expected > 0 ? ((item.actual / item.expected) * 100).toFixed(2) : "0.00";
    }
  
    return summary;
  },
    },
    mounted() {
      if (this.selectedDept) this.autoFillDepartment(this.selectedDept);
      this.updateTimeOptions(),
      this.fetchTableData2();

    },
    watch: {
      selectedDept(newDept) {
        if (newDept) this.autoFillDepartment(newDept);
      }
    },

  methods: {
    autoFillDepartment(deptCode) {
      const matched = deptCode.match(/^(\d)D(\d)S(\d)C$/);
      if (matched) {
        this.selectedD = matched[1];
        this.selectedS = matched[2];
        this.selectedC = matched[3];
        this.handleInputChange(); // 若需同步查詢
      } else {
        console.warn('❗班別格式不符：', deptCode);
      }
    },
    
    updateTimeOptions() {
      if (this.selectedD === "3") {
        this.timeOptions = {
            "7:00-8:00": 1.00,
            "8:00-10:00": 2.00,
            "10:00-11:40": 1.67,
            "12:00-12:40": 0.67,
            "12:40-15:00": 2.33,
            "15:00-17:00": 2,
            "17:00-20:00": 3,
            "20:00-21:00": 1,
            "21:00-22:00": 1
        };
      } else if (this.selectedD === "2") {
        this.timeOptions = {
            "7:00-8:00": 1.00,
            "8:00-10:00": 2.00,
            "10:00-11:20": 1.33,
            "11:40-12:20": 0.67,
            "12:20-15:00": 2.66,
            "15:00-17:00": 2,
            "17:00-20:00": 3,
            "20:00-21:00": 1,
            "21:00-22:00": 1
        };
      } else {
        this.timeOptions = {
            "7:00-8:00": 1.00,
            "8:00-10:00": 2.00,
            "10:00-11:10": 1.33,
            "12:20-15:00": 2.67,
            "15:00-17:00": 2,
            "17:00-20:00": 3,
            "20:00-21:00": 1,
            "21:00-22:00": 1
        };
      }
      console.log("Response:", this.selectedD);
    },
    handleInputChange() {
      this.updateTimeOptions(),
      this.fetchTableData2();

    },
    async WOVER() {
      this.WOVERSON=[];
      if (this.selectedL) {
        this.selectedDepartment = this.selectedD + 'D' + this.selectedS + 'S' + this.selectedC + '-' + this.selectedL + 'C';
        } else {
        this.selectedDepartment = this.selectedD + 'D' + this.selectedS + 'S' + this.selectedC + 'C';
        }
        
    
   
      /*const date = new Date(this.selectedDate);
      const year = date.getFullYear();
      const month = (date.getMonth() + 1).toString().padStart(2, '0');
      const day = date.getDate().toString().padStart(2, '0');
      const formattedDate = `${year}${month}${day}`;*/
        if (this.selectedDate) {
          // 把 YYYY-MM-DD 轉成 YYYYMMDD（移除 -）
           formattedDate  = this.selectedDate.replace(/-/g, '');

        }else{
        const today = new Date();
        const year = today.getFullYear();
        const month = String(today.getMonth() + 1).padStart(2, '0'); // 補零
        const day = String(today.getDate()).padStart(2, '0'); // 補零
         formattedDate = `${year}${month}${day}`; // YYYYMMDD
      }

      const apiUrl1 = `https://mms.leapoptical.com:5088/api/NFC_API/WO_version?AREA=${this.selectedArea}&DEPARTMENT=${this.selectedDepartment}&MFG_DAY=${formattedDate}`;
    
      try {
        const response = await fetch(apiUrl1);
        const data = await response.json();
        this.WOVERSON.push(...data);
        this.WOVERSON.sort((a, b) => a.部門.localeCompare(b.部門));
      } catch (error) {
        console.error("取得 WOVER 資料失敗：", error);
      } finally {
        this.loading = false;
      }
    },
    
    async fetchTableData2() {
      if (this.selectedDate) {
        // 把 YYYY-MM-DD 轉成 YYYYMMDD（移除 -）
         formattedDate  = this.selectedDate.replace(/-/g, '');
        console.log('formattedDate:',formattedDate);
      }else{
      const today = new Date();
      const year = today.getFullYear();
      const month = String(today.getMonth() + 1).padStart(2, '0'); // 補零
      const day = String(today.getDate()).padStart(2, '0'); // 補零
       formattedDate = `${year}${month}${day}`; // YYYYMMDD
    }
          if (this.selectedL) {
      this.selectedDepartment = this.selectedD + 'D' + this.selectedS + 'S' + this.selectedC + '-' + this.selectedL + 'C';
      } else {
      this.selectedDepartment = this.selectedD + 'D' + this.selectedS + 'S' + this.selectedC + 'C';
      }
      
      
      
      
          try {
            const apiUrl = `https://mms.leapoptical.com:5088/api/NFC_API/SELECPD2?area=${this.selectedArea}&department=${this.selectedDepartment}&&mfg_day=${formattedDate}`;
            const response = await axios.get(apiUrl);
            console.log("Response:", response.data);
            if (!response.data.result || !Array.isArray(response.data.result)) {
              console.error("API response is invalid:", response.data);
              return;
            }
      
            // 清理數據，把 undefined、空字串 或 "undefined" 轉為空值或 0
            this.tableData = response.data.result.map(item => ({
                ID: item.ID || "",  
              AREA: item.AREA || "",
              MFG_DAY: item.MFG_DAY, // API 回傳沒有 MFG_DAY，從參數取得
              DEPARTMENT: item.DEPARTMENT || "",
              people: item.people || "0",
              time: item.time || "",
              workhorse: item.workhorse || "",
              product: item.product || "",
              expected: item.expected || "0",
              actual: item.actual || "0",
              line: item.line && item.line !== "undefined" ? item.line : "",
              defectForming: item.defectForming && item.defectForming !== "undefined" ? item.defectForming : "0",
              defectWelding: item.defectWelding && item.defectWelding !== "undefined" ? item.defectWelding : "0",
              NG: item.NG && item.NG !== "undefined" ? item.NG : "0",
              other1: item.other1 && item.other1 !== "undefined" ? item.other1 : "0",
              other2: item.other2 && item.other2 !== "undefined" ? item.other2 : "0",
              other3: item.other3 && item.other3 !== "undefined" ? item.other3 : "0",
              other4: item.other4 && item.other4 !== "undefined" ? item.other4 : "0",
              other5: item.other5 && item.other5 !== "undefined" ? item.other5 : "0",
              other6: item.other6 && item.other6 !== "undefined" ? item.other6 : "0",
              other1_desc: item.other1_desc || "",
              other2_desc: item.other2_desc || "",
              other3_desc: item.other3_desc || "",
              other4_desc: item.other4_desc || "",
              other5_desc: item.other5_desc || "",
              other6_desc: item.other6_desc || "",
            }));
          } catch (error) {
            console.error("Error fetching data:", error);
          }
          if (this.tableData.length > 0) {
      const firstRow = this.tableData[0];
      this.selectedProcess1 = firstRow.other1_desc || "";
      this.selectedProcess2 = firstRow.other2_desc || "";
      this.selectedProcess3 = firstRow.other3_desc || "";
      this.selectedProcess4 = firstRow.other4_desc || "";
      this.selectedProcess5 = firstRow.other5_desc || "";
      this.selectedProcess6 = firstRow.other6_desc || "";
    
    
    }
    await this.WOVER();
    },
    async toggleEditMode() { // 確保這裡是 async
      this.editMode = !this.editMode;
      const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0'); // 補零
    const day = String(today.getDate()).padStart(2, '0'); // 補零
    const formattedDate = `${year}${month}${day}`; // YYYYMMDD
  
      if (!this.editMode) {
        try {
          // 確保所有數據轉成字串
          const formattedData = this.tableData.map(row => ({
         //   AREA: String(row.AREA),
       //     MFG_DAY: String(row.MFG_DAY),
         //   DEPARTMENT: String(row.DEPARTMENT),
            ID: String(row.ID)|| "0",
            AREA: this.selectedArea,
            MFG_DAY: formattedDate,
            DEPARTMENT: this.selectedDepartment,
            PEOPLE: String(row.people),
            TIME: String(row.time),
            PRODUCT: String(row.product),
            WORKHORSE:String(row.workhorse),
            EXPECTED: String(row.expected),
            ACTUAL: String(row.actual),
            LINE: String(row.line),
            DEFECTFORMING: String(row.defectForming),
            DEFECTWELDING: String(row.defectWelding),
            NG: String(row.NG),
            OTHER1: String(row.other1),
            OTHER2: String(row.other2),
            OTHER3: String(row.other3)
          }));
          console.log("送出前的資料:", JSON.stringify(formattedData, null, 2));
  
          console.log("Response:", formattedData);
  
          const response = await axios.post("http://192.168.209.18:5088/api/C_PDREPORT/InsertOrUpdate", 
    formattedData, // 直接傳遞數據
    { headers: { "Content-Type": "application/json" } }
  );
          console.log("Response:", response.data);
          this.$message.success("資料已成功提交！");
  
  
  
  
        } catch (error) {
          console.error("API Error:", error);
          this.$message.error("提交失敗，請稍後再試！");
  
        }
      } else {
        this.fetchProductOptions(); // 在編輯模式開啟時調用
      }
    },
    addRow() {
          this.tableData.push({ ID:0 ,people: 0, time: "", product: "", expected: 0, actual: 0 });
      },
      async removeRow2(row, index) {
    const deletedId = row.ID; // 取得要刪除的 ID
    console.log("🔹 刪除的 ID:", deletedId);
    if (!deletedId) {
      console.warn("⚠️ ID 為 null，直接刪除前端資料");
      this.tableData.splice(index, 1);
      return;
    }
    try {
      const response = await axios.post(
        "http://192.168.209.18:5088/api/C_PDREPORT/delete",
        { ID: String(deletedId) },
        {
          headers: {
            "Content-Type": "application/json",
          },
        }
      );
  
      console.log("🔹 API Response:", response);
  
      if (response.data && response.data.message === "Record deleted successfully.") {
        this.tableData.splice(index, 1);
      } else {
        alert("刪除失敗：" + (response.data?.message || "未知錯誤"));
      }
    } catch (error) {
      console.error("🔴 刪除時發生錯誤:", error);
      if (error.response) {
        console.error("🔴 伺服器回應:", error.response);
        alert("錯誤：" + (error.response.data?.message || "請求無效"));
      } else {
        alert("無法連線到伺服器，請稍後再試");
      }
    }
  },
  translate(text) {
    return this.translatedText[text] || text;
  },
  tableRowClassName({ row }) {
    return row._isSummary ? 'summary-row' : '';
  },

  mergeTable({ row, column, rowIndex }) {
    // 根據部門選擇不同的合併欄位清單
    /*const mergeColumns = (this.selectedDepartment?.toUpperCase() === '3D1S6C')
      ? ["time", "people", "wohourse"]
      : [
          "time", "people", "wohourse", "line",
          "defectForming", "defectWelding", "NG",
          "other1", "other2", "other3", "other4", "other5", "other6"
        ];*/
  
        const mergeColumns = this.isMergeEnabled
      ? [
          "time", "people", "wohourse", "line",
          "defectForming", "defectWelding", "NG",
          "other1", "other2", "other3", "other4", "other5", "other6"
        ]
      : ["time", "people", "wohourse"];
    // 如果該欄位不在要合併的欄位中，就回傳預設 rowspan:1
    if (!mergeColumns.includes(column.property)) {
      return { rowspan: 1, colspan: 1 };
    }
  
    const currentTime = row.time;
    const prevRow = this.tableData[rowIndex - 1];
  
    // 第一列 或時段不同 → 計算 rowspan
    if (rowIndex === 0 || currentTime !== prevRow?.time) {
      let rowspan = 1;
      for (let i = rowIndex + 1; i < this.tableData.length; i++) {
        if (this.tableData[i].time === currentTime) {
          rowspan++;
        } else {
          break;
        }
      }
      return { rowspan, colspan: 1 };
    }
  
    // 其他重複列不顯示
    return { rowspan: 0, colspan: 0 };
  },
  getSummaries(param) {
    const { columns, data } = param;
    const sums = [];
    
    // 確保變數在迴圈外部定義，避免 ReferenceError
    let totalExpected = 0; // 總標準產能
    let totalActual = 0; // 總實際產出
    
    columns.forEach((column, index) => {
      if (index === 0) {
        sums[index] = '總計';
        return;
      }
    
      const prop = column.property;
      if (prop) {
        if ([ 'expected', 'actual', 'line', 'defectForming', 'defectWelding', 'NG', 'other1', 'other2', 'other3'].includes(prop)) {
          const total = data.reduce((sum, row) => sum + (Number(row[prop]) || 0), 0);
          sums[index] =(total).toFixed(0) ;
    
          // 記錄總標準產能與總實際產出，以計算總體效率
          if (prop === 'expected') totalExpected = total;
          if (prop === 'actual') totalActual = total;
    
        } else if (prop === 'workhorse') {
          sums[index] = data.reduce((sum, row) => sum + (Number(this.calculateHours(row)) || 0), 0).toFixed(2);
    
        } else if (prop === 'wohourse') {
          sums[index] = data.reduce((sum, row) => sum + (Number(this.calculateHours(row)) || 0), 0).toFixed(2);
    
        } else if (prop === 'efficiency') {
          // 計算總體效率
          sums[index] = totalExpected > 0 ? ((totalActual / totalExpected) * 100).toFixed(2) + '%' : '0%';
    
        } else {
          sums[index] = ''; // 其他欄位不加總
        }
      }
    });
    
    return sums;
    },
    calculateHours(row) {
      return (row.people || 0.00) * (this.timeOptions[row.time] || 0.00);
  },
  calculateEfficiency(row) {
    if (row.expected === 0) return 0;
    return ((row.actual / row.expected) * 100).toFixed(2);
},
renderHeaderWithSelect1(h) {
  return this.renderHeaderSelect(h, this.selectedProcess1, val => this.selectedProcess1 = val);
},
renderHeaderWithSelect2(h) {
  return this.renderHeaderSelect(h, this.selectedProcess2, val => this.selectedProcess2 = val);
},
renderHeaderWithSelect3(h) {
  return this.renderHeaderSelect(h, this.selectedProcess3, val => this.selectedProcess3 = val);
},
renderHeaderWithSelect4(h) {
  return this.renderHeaderSelect(h, this.selectedProcess4, val => this.selectedProcess4 = val);
},
renderHeaderWithSelect5(h) {
  return this.renderHeaderSelect(h, this.selectedProcess5, val => this.selectedProcess5 = val);
},
renderHeaderWithSelect6(h) {
  return this.renderHeaderSelect(h, this.selectedProcess6, val => this.selectedProcess6 = val);
},
renderHeaderSelect(h, selectedValue, onInput) {
return h('div', [
  // 下拉選單本體
  h('el-select', {
    class: 'header-select-white',
    props: {
      value: selectedValue,
      placeholder: '選擇流程',
      size: 'mini'
    },
    on: {
      input: onInput
    }
  }, Object.keys(this.processSwitch).map(key =>
    h('el-option', {
      props: {
        label: (key),
        value: key
      }
    })
  )),

  // 條件顯示說明區塊
  selectedValue === '打扭'
    ? h('span', { style: { fontSize: '11px', display: 'block', marginTop: '2px' } }, [
        h('small', 'xoắn dây')
      ])
    : selectedValue === '套套管'
    ? h('span', { style: { fontSize: '11px', display: 'block', marginTop: '2px' } }, [
        h('small', 'lồng vỏ bọc')
      ])
    : selectedValue === '內模'
    ? h('span', { style: { fontSize: '11px', display: 'block', marginTop: '2px' } }, [
        h('small', 'khuôn trong')
      ])
    : selectedValue === '外模'
    ? h('span', { style: { fontSize: '11px', display: 'block', marginTop: '2px' } }, [
        h('small', 'khuôn ngoài')
      ])
    : selectedValue === '外觀'
    ? h('span', { style: { fontSize: '11px', display: 'block', marginTop: '2px' } }, [
        h('small', 'ngoại quan')
      ])
    :selectedValue === '包裝'
    ? h('span', { style: { fontSize: '11px', display: 'block', marginTop: '2px' } }, [
        h('small', 'đóng gói')
      ])
    : null
]);

},




  }
  
});
