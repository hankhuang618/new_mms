// components/overflowperformance.js
// 溢領考績（Over-claim Performance）— 完整版
// 風格與互動全面對齊「效率考績」，API 與 NFC.html 完全一致。
// 依序載入 5 組資料：
//  1) /overflowperformance?area=&year=&month=           （主表 班別 × 日期）
//  2) /DEPover?area=&year=&month=                        （部門彙總）
//  3) /QCDEPover?area=&year=&month=                      （QC 部門）
//  4) /CNDEPover?area=VN                                 （CN/中國 固定 area=VN）
//  5) /stff_OVER_bonus?area=&year=&month=                （員工獎金•溢領版；僅統計用）
// 匯出：將本頁 6 張表一次匯出為同一活頁簿。

(function(){
  // === 環境設定（保持與 NFC.html 一致） ===
  const API_HOST   = 'http://192.168.207.17:5024';
  const API_PREFIX = API_HOST + '/api/performance';
  const CN_FIXED_AREA = 'VN'; // NFC 的 CNDEPover 固定 area=VN

  // === 共用樣式：沿用效率考績外觀 ===
  function loadCSS(href){
    if ([...document.querySelectorAll('link[rel="stylesheet"]')].some(x=>x.href.includes(href))) return;
    const link = document.createElement('link'); link.rel='stylesheet'; link.href=href; document.head.appendChild(link);
  }
  loadCSS('assets/effperformance.css');

  const pad2 = n => (n<10? ('0'+n) : ''+n);
  const today = new Date();

  Vue.component('overflowperformance-view', {
    template: `
    <div class="eff-page">
      <!-- Header -->
      <div class="eff-header shadow-sm p-2 mb-3 bg-light rounded d-flex flex-wrap align-items-center justify-content-between">
        <div class="d-flex align-items-center flex-wrap gap-2">
          <label>🌍 區域：</label>
          <select class="form-select form-select-sm w-auto" v-model="selectedArea" @change="fetchAll">
            <option value="ZH">ZH</option>
            <option value="TC">TC</option>
            <option value="VN">VN</option>
            <option value="TW">TW</option>
          </select>

          <label>📅 年份：</label>
          <select class="form-select form-select-sm w-auto" v-model.number="selectedYear" @change="fetchAll">
            <option v-for="y in yearOptions" :key="'y'+y" :value="y">{{ y }}</option>
          </select>

          <label>月份：</label>
          <select class="form-select form-select-sm w-auto" v-model="selectedMonth" @change="fetchAll">
            <option v-for="m in 12" :key="'m'+m" :value="pad2(m)">{{ pad2(m) }}</option>
          </select>
        </div>

        <div class="d-flex flex-wrap gap-2">
          <button v-for="(sheet, index) in sheets"
                  :key="'tab'+index"
                  @click="currentSheet = index"
                  class="btn btn-outline-primary btn-sm"
                  :class="{ active: currentSheet === index }">
            {{ sheet.name }}
          </button>
          <button @click="exportExcel()" class="btn btn-success btn-sm">📤 導出 Excel</button>
          <el-button size="small" type="primary" :loading="loading" @click="fetchAll">重新整理</el-button>
        </div>
      </div>

      <el-alert v-if="error" type="error" :title="error" show-icon class="mb-2"/>

      <div class="table-zone" v-loading="loading">
        <!-- 三部（部門=3） -->
        <div v-show="currentSheet===0">
          <h5 class="mb-2 fw-bold text-primary">三部 溢領表</h5>
          <table id="WOOEVER3" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
            <thead>
              <tr>
                <th rowspan="2">班別代號</th>
                <th v-for="(date, idx) in main_dates" :key="'h3-'+date" :colspan="idx===main_dates.length-1?10:3">{{ date }}</th>
              </tr>
              <tr>
                <template v-for="(date, idx) in main_dates" :key="'s3-'+date">
                  <th>應領金額</th>
                  <th>溢領金額</th>
                  <th>溢領比例</th>
                  <template v-if="idx===main_dates.length-1">
                    <th>歸還溢領</th>
                    <th>當月溢領比例上升率</th>
                    <th>當月最終效率考核</th>
                    <th>考績</th>
                    <th>額外考績</th>
                    <th>下月歸還溢領</th>
                    <th>（保留1）</th>
                    <th>（保留2）</th>
                  </template>
                </template>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(rows, erp) in main_group" :key="'r3-'+erp" v-if="isDept(rows,'3')">
                <td class="fw-bold bg-light">{{ erp }}</td>
                <template v-for="(date,i) in main_dates" :key="'b3-'+erp+'-'+date">
                  <td>{{ fmtN(getMain(erp,date,'應領金額')) }}</td>
                  <td>{{ fmtN(getMain(erp,date,'溢領金額')) }}</td>
                  <td>{{ fmt2(getMain(erp,date,'溢領比例')) }}</td>
                  <template v-if="i===main_dates.length-1">
                    <td>{{ fmt2(getMain(erp,date,'累計')) }}</td>
                    <td>{{ fmt2(getMain(erp,date,'上升率')) }}</td>
                    <td>{{ fmt2(getMain(erp,date,'績效')) }}</td>
                    <td>{{ getMain(erp,date,'考績') }}</td>
                    <td>{{ getMain(erp,date,'額外考績') }}</td>
                    <td>{{ fmt2(getMain(erp,date,'下月累計負值')) }}</td>
                    <td></td><td></td>
                  </template>
                </template>
              </tr>
              <tr>
                <td>合計</td>
                <template v-for="(date,i) in main_dates" :key="'t3-'+date">
                  <td>{{ fmtN(sumMain(date,'應領金額','3')) }}</td>
                  <td>{{ fmtN(sumMain(date,'溢領金額','3')) }}</td>
                  <td>{{ fmt2(divZero(sumMain(date,'溢領金額','3'), sumMain(date,'應領金額','3'))) }}</td>
                  <template v-if="i===main_dates.length-1">
                    <td>{{ fmt2(sumMain(date,'累計','3')) }}</td>
                    <td>{{ fmt2(sumMain(date,'上升率','3')) }}</td>
                    <td>{{ fmt2(sumMain(date,'績效','3')) }}</td>
                    <td>-</td><td>-</td>
                    <td>{{ fmt2(sumMain(date,'下月累計負值','3')) }}</td>
                    <td></td><td></td>
                  </template>
                </template>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- 二部（部門=2） -->
        <div v-show="currentSheet===1">
          <h5 class="mb-2 fw-bold text-primary">二部 溢領表</h5>
          <table id="WOOEVER2" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
            <thead>
              <tr>
                <th rowspan="2">班別代號</th>
                <th v-for="(date, idx) in main_dates" :key="'h2-'+date" :colspan="idx===main_dates.length-1?10:3">{{ date }}</th>
              </tr>
              <tr>
                <template v-for="(date, idx) in main_dates" :key="'s2-'+date">
                  <th>應領金額</th>
                  <th>溢領金額</th>
                  <th>溢領比例</th>
                  <template v-if="idx===main_dates.length-1">
                    <th>歸還溢領</th>
                    <th>當月溢領比例上升率</th>
                    <th>當月最終效率考核</th>
                    <th>考績</th>
                    <th>額外考績</th>
                    <th>下月歸還溢領</th>
                    <th>（保留1）</th>
                    <th>（保留2）</th>
                  </template>
                </template>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(rows, erp) in main_group" :key="'r2-'+erp" v-if="isDept(rows,'2')">
                <td class="fw-bold bg-light">{{ erp }}</td>
                <template v-for="(date,i) in main_dates" :key="'b2-'+erp+'-'+date">
                  <td>{{ fmtN(getMain(erp,date,'應領金額')) }}</td>
                  <td>{{ fmtN(getMain(erp,date,'溢領金額')) }}</td>
                  <td>{{ fmt2(getMain(erp,date,'溢領比例')) }}</td>
                  <template v-if="i===main_dates.length-1">
                    <td>{{ fmt2(getMain(erp,date,'累計')) }}</td>
                    <td>{{ fmt2(getMain(erp,date,'上升率')) }}</td>
                    <td>{{ fmt2(getMain(erp,date,'績效')) }}</td>
                    <td>{{ getMain(erp,date,'考績') }}</td>
                    <td>{{ getMain(erp,date,'額外考績') }}</td>
                    <td>{{ fmt2(getMain(erp,date,'下月累計負值')) }}</td>
                    <td></td><td></td>
                  </template>
                </template>
              </tr>
              <tr>
                <td>合計</td>
                <template v-for="(date,i) in main_dates" :key="'t2-'+date">
                  <td>{{ fmtN(sumMain(date,'應領金額','2')) }}</td>
                  <td>{{ fmtN(sumMain(date,'溢領金額','2')) }}</td>
                  <td>{{ fmt2(divZero(sumMain(date,'溢領金額','2'), sumMain(date,'應領金額','2'))) }}</td>
                  <template v-if="i===main_dates.length-1">
                    <td>{{ fmt2(sumMain(date,'累計','2')) }}</td>
                    <td>{{ fmt2(sumMain(date,'上升率','2')) }}</td>
                    <td>{{ fmt2(sumMain(date,'績效','2')) }}</td>
                    <td>-</td><td>-</td>
                    <td>{{ fmt2(sumMain(date,'下月累計負值','2')) }}</td>
                    <td></td><td></td>
                  </template>
                </template>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- 一部（部門=1） -->
        <div v-show="currentSheet===2">
          <h5 class="mb-2 fw-bold text-primary">一部 溢領表</h5>
          <table id="WOOEVER1" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
            <thead>
              <tr>
                <th rowspan="2">班別代號</th>
                <th v-for="(date, idx) in main_dates" :key="'h1-'+date" :colspan="idx===main_dates.length-1?10:3">{{ date }}</th>
              </tr>
              <tr>
                <template v-for="(date, idx) in main_dates" :key="'s1-'+date">
                  <th>應領金額</th>
                  <th>溢領金額</th>
                  <th>溢領比例</th>
                  <template v-if="idx===main_dates.length-1">
                    <th>歸還溢領</th>
                    <th>當月溢領比例上升率</th>
                    <th>當月最終效率考核</th>
                    <th>考績</th>
                    <th>額外考績</th>
                    <th>下月歸還溢領</th>
                    <th>（保留1）</th>
                    <th>（保留2）</th>
                  </template>
                </template>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(rows, erp) in main_group" :key="'r1-'+erp" v-if="isDept(rows,'1')">
                <td class="fw-bold bg-light">{{ erp }}</td>
                <template v-for="(date,i) in main_dates" :key="'b1-'+erp+'-'+date">
                  <td>{{ fmtN(getMain(erp,date,'應領金額')) }}</td>
                  <td>{{ fmtN(getMain(erp,date,'溢領金額')) }}</td>
                  <td>{{ fmt2(getMain(erp,date,'溢領比例')) }}</td>
                  <template v-if="i===main_dates.length-1">
                    <td>{{ fmt2(getMain(erp,date,'累計')) }}</td>
                    <td>{{ fmt2(getMain(erp,date,'上升率')) }}</td>
                    <td>{{ fmt2(getMain(erp,date,'績效')) }}</td>
                    <td>{{ getMain(erp,date,'考績') }}</td>
                    <td>{{ getMain(erp,date,'額外考績') }}</td>
                    <td>{{ fmt2(getMain(erp,date,'下月累計負值')) }}</td>
                    <td></td><td></td>
                  </template>
                </template>
              </tr>
              <tr>
                <td>合計</td>
                <template v-for="(date,i) in main_dates" :key="'t1-'+date">
                  <td>{{ fmtN(sumMain(date,'應領金額','1')) }}</td>
                  <td>{{ fmtN(sumMain(date,'溢領金額','1')) }}</td>
                  <td>{{ fmt2(divZero(sumMain(date,'溢領金額','1'), sumMain(date,'應領金額','1'))) }}</td>
                  <template v-if="i===main_dates.length-1">
                    <td>{{ fmt2(sumMain(date,'累計','1')) }}</td>
                    <td>{{ fmt2(sumMain(date,'上升率','1')) }}</td>
                    <td>{{ fmt2(sumMain(date,'績效','1')) }}</td>
                    <td>-</td><td>-</td>
                    <td>{{ fmt2(sumMain(date,'下月累計負值','1')) }}</td>
                    <td></td><td></td>
                  </template>
                </template>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- 部門彙總（DEPover） -->
        <div v-show="currentSheet===3">
          <h5 class="mb-2 fw-bold text-primary">部門彙總</h5>
          <table id="DEPOVER" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
            <thead>
              <tr>
                <th rowspan="2">部門</th>
                <th v-for="(date, idx) in DEP_dates" :key="'hdD-'+date" :colspan="idx===DEP_dates.length-1?11:1">{{ date }}</th>
              </tr>
              <tr>
                <template v-for="(date, idx) in DEP_dates" :key="'shD-'+date">
                  <th>效率</th>
                  <template v-if="idx===DEP_dates.length-1">
                    <th>累計</th>
                    <th>上升率</th>
                    <th>績效</th>
                    <th>考績</th>
                    <th>額外考績</th>
                    <th>考績獎金</th>
                    <th>額外考績獎金</th>
                    <th>A15拉管理福利金</th>
                    <th>最終獎金</th>
                    <th>下月累計負值</th>
                  </template>
                </template>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(rows, dept) in DEP_group" :key="dept">
                <td>{{ dept }}</td>
                <template v-for="(date, i) in DEP_dates" :key="'dd-'+dept+'-'+date">
                  <td>{{ fmt2(getDEP(dept,date,'效率')) }}</td>
                  <template v-if="i===DEP_dates.length-1">
                    <td>{{ fmt2(getDEP(dept,date,'累計')) }}</td>
                    <td>{{ fmt2(getDEP(dept,date,'上升率')) }}</td>
                    <td>{{ fmt2(getDEP(dept,date,'績效')) }}</td>
                    <td>{{ dashZero(getDEP(dept,date,'考績')) }}</td>
                    <td>{{ dashZero(getDEP(dept,date,'額外考績')) }}</td>
                    <td>{{ money(getDEP(dept,date,'考績獎金')) }}</td>
                    <td>{{ money(getDEP(dept,date,'額外考績獎金')) }}</td>
                    <td>{{ money(getDEP(dept,date,'A15拉管理福利金')) }}</td>
                    <td>{{ money(getDEP(dept,date,'最終獎金')) }}</td>
                    <td>{{ fmt2(getDEP(dept,date,'下月累計負值')) }}</td>
                  </template>
                </template>
              </tr>
              <tr>
                <td colspan="16" class="text-end fw-bold">Total 最終獎金</td>
                <td class="fw-bold">{{ money(totalDEPOVERFinalBonus) }}</td>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- QC 部門（QCDEPover） -->
        <div v-show="currentSheet===4">
          <h5 class="mb-2 fw-bold text-primary">QC 部門</h5>
          <table id="QCDEPOVER" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
            <thead>
              <tr>
                <th rowspan="2">管理部門</th>
                <th rowspan="2">管理班別</th>
                <th rowspan="2">工號</th>
                <th rowspan="2">姓名</th>
                <th rowspan="2">職稱</th>
                <th v-for="(date, idx) in QC_dates" :key="'hdQ-'+date" :colspan="idx===QC_dates.length-1?11:1">{{ date }}</th>
              </tr>
              <tr>
                <template v-for="(date, idx) in QC_dates" :key="'shQ-'+date">
                  <th>效率</th>
                  <template v-if="idx===QC_dates.length-1">
                    <th>累計</th>
                    <th>上升率</th>
                    <th>績效</th>
                    <th>考績</th>
                    <th>額外考績</th>
                    <th>考績獎金</th>
                    <th>額外考績獎金</th>
                    <th>A15拉管理福利金</th>
                    <th>最終獎金</th>
                    <th>下月累計負值</th>
                  </template>
                </template>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(rows, emp) in QC_group" :key="emp">
                <td>{{ getQC(emp,lastKey(QC_dates),'部') }}</td>
                <td>{{ getQC(emp,lastKey(QC_dates),'班') }}</td>
                <td>{{ emp }}</td>
                <td>{{ getQC(emp,lastKey(QC_dates),'姓名') }}</td>
                <td>{{ getQC(emp,lastKey(QC_dates),'崗位') }}</td>
                <template v-for="(date,i) in QC_dates" :key="'dq-'+emp+'-'+date">
                  <td>{{ fmt2(getQCV(emp,date,'效率')) }}</td>
                  <template v-if="i===QC_dates.length-1">
                    <td>{{ fmt2(getQCV(emp,date,'累計')) }}</td>
                    <td>{{ fmt2(getQCV(emp,date,'上升率')) }}</td>
                    <td>{{ fmt2(getQCV(emp,date,'績效')) }}</td>
                    <td>{{ dashZero(getQCV(emp,date,'考績')) }}</td>
                    <td>{{ dashZero(getQCV(emp,date,'額外考績')) }}</td>
                    <td>{{ money(getQCV(emp,date,'考績獎金')) }}</td>
                    <td>{{ money(getQCV(emp,date,'額外考績獎金')) }}</td>
                    <td>{{ money(getQCV(emp,date,'A15拉管理福利金')) }}</td>
                    <td>{{ money(getQCV(emp,date,'最終獎金')) }}</td>
                    <td>{{ fmt2(getQCV(emp,date,'下月累計負值')) }}</td>
                  </template>
                </template>
              </tr>
              <tr>
                <td colspan="16" class="text-end fw-bold">Total 最終獎金</td>
                <td class="fw-bold">{{ money(totalQCOVERFinalBonus) }}</td>
              </tr>
            </tbody>
          </table>
        </div>

        <!-- CN（中國，CNDEPover） -->
        <div v-show="currentSheet===5">
          <h5 class="mb-2 fw-bold text-primary">CN（中國）</h5>
          <table id="CNDEPOVER" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
            <thead>
              <tr>
                <th rowspan="2">管理部門</th>
                <th rowspan="2">管理班別</th>
                <th rowspan="2">工號</th>
                <th rowspan="2">姓名</th>
                <th rowspan="2">職稱</th>
                <th v-for="(date, idx) in ZH_dates" :key="'hdZ-'+date" :colspan="idx===ZH_dates.length-1?11:1">{{ date }}</th>
              </tr>
              <tr>
                <template v-for="(date, idx) in ZH_dates" :key="'shZ-'+date">
                  <th>效率</th>
                  <template v-if="idx===ZH_dates.length-1">
                    <th>累計</th>
                    <th>上升率</th>
                    <th>績效</th>
                    <th>考績</th>
                    <th>額外考績</th>
                    <th>考績獎金</th>
                    <th>額外考績獎金</th>
                    <th>A15拉管理福利金</th>
                    <th>最終獎金</th>
                    <th>下月累計負值</th>
                  </template>
                </template>
              </tr>
            </thead>
            <tbody>
              <tr v-for="(rows, emp) in ZH_group" :key="emp">
                <td>{{ getZH(emp,lastKey(ZH_dates),'管理部門') || getZH(emp,lastKey(ZH_dates),'部門') }}</td>
                <td>{{ getZH(emp,lastKey(ZH_dates),'管理班別') || getZH(emp,lastKey(ZH_dates),'班') }}</td>
                <td>{{ emp }}</td>
                <td>{{ getZH(emp,lastKey(ZH_dates),'姓名') }}</td>
                <td>{{ getZH(emp,lastKey(ZH_dates),'職稱') }}</td>
                <template v-for="(date,i) in ZH_dates" :key="'dz-'+emp+'-'+date">
                  <td>{{ fmt2(getZHV(emp,date,'效率')) }}</td>
                  <template v-if="i===ZH_dates.length-1">
                    <td>{{ fmt2(getZHV(emp,date,'累計')) }}</td>
                    <td>{{ fmt2(getZHV(emp,date,'上升率')) }}</td>
                    <td>{{ fmt2(getZHV(emp,date,'績效')) }}</td>
                    <td>{{ dashZero(getZHV(emp,date,'考績')) }}</td>
                    <td>{{ dashZero(getZHV(emp,date,'額外考績')) }}</td>
                    <td>{{ money(getZHV(emp,date,'考績獎金')) }}</td>
                    <td>{{ money(getZHV(emp,date,'額外考績獎金')) }}</td>
                    <td>{{ money(getZHV(emp,date,'A15拉管理福利金')) }}</td>
                    <td>{{ money(getZHV(emp,date,'最終獎金')) }}</td>
                    <td>{{ fmt2(getZHV(emp,date,'下月累計負值')) }}</td>
                  </template>
                </template>
              </tr>
              <tr>
                <td colspan="16" class="text-end fw-bold">Total 最終獎金</td>
                <td class="fw-bold">{{ money(totalZHOVERFinalBonus) }}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
    `,

    data(){
      return {
        loading:false, error:'',
        selectedArea:'VN',
        selectedYear: today.getFullYear(),
        selectedMonth: pad2(today.getMonth()+1),
        yearOptions:[ today.getFullYear()-1, today.getFullYear(), today.getFullYear()+1 ],

        sheets:[ {name:'三部'}, {name:'二部'}, {name:'一部'}, {name:'部門彙總'}, {name:'QC 部門'}, {name:'中國（CN）'} ],
        currentSheet:0,

        // 主表 overflowperformance
        main_raw:[], main_dates:[], main_group:{},
        // DEPover
        DEP_raw:[], DEP_dates:[], DEP_group:{},
        // QCDEPover
        QC_raw:[], QC_dates:[], QC_group:{},
        // CNDEPover
        ZH_raw:[], ZH_dates:[], ZH_group:{},
        // stff_OVER_bonus（可作後續個人獎金表或彙總用）
        stff_over_bonus:[],
      }
    },

    computed:{
      totalDEPOVERFinalBonus(){
        const last = this.lastKey(this.DEP_dates); if(!last) return 0;
        let sum=0; for(const k in this.DEP_group){ const r=this.DEP_group[k]?.[last]; if(r) sum+=Number(r['最終獎金']||0); } return sum;
      },
      totalQCOVERFinalBonus(){
        const last = this.lastKey(this.QC_dates); if(!last) return 0;
        let sum=0; for(const k in this.QC_group){ const r=this.QC_group[k]?.[last]; if(r) sum+=Number(r['最終獎金']||0); } return sum;
      },
      totalZHOVERFinalBonus(){
        const last = this.lastKey(this.ZH_dates); if(!last) return 0;
        let sum=0; for(const k in this.ZH_group){ const r=this.ZH_group[k]?.[last]; if(r) sum+=Number(r['最終獎金']||0); } return sum;
      },
    },

    methods:{
      pad2,
      lastKey(arr){ return (arr&&arr.length)? arr[arr.length-1] : ''; },

      // 格式化
      fmtN(v){ const n=Number(v||0); return n===0? '' : n.toLocaleString('en-US',{maximumFractionDigits:0}); },
      fmt2(v){ const n=Number(v||0); return n===0? '' : (Math.round(n*100)/100).toFixed(2); },
      money(v){ const n=Number(v||0); return n===0? '-' : n.toLocaleString('en-US',{maximumFractionDigits:0}); },
      dashZero(v){ const n=Number(v||0); return n===0? '-' : n; },
      divZero(a,b){ const x=Number(a||0), y=Number(b||0); if(!y) return 0; return x/y; },

      // === 抓資料（API 與 NFC.html 一致） ===
      async fetchAll(){ this.loading=true; this.error='';
        try{
          await Promise.all([
            this.fetchOverflowMain(),
            this.fetchDEPover(),
            this.fetchQCDEPover(),
            this.fetchCNDEPover(),
            this.fetchStffOverBonus(),
          ]);
        }catch(e){ this.error = '讀取資料失敗：' + (e?.message||e); }
        finally{ this.loading=false; }
      },

      async fetchOverflowMain(){
        const url = `${API_PREFIX}/overflowperformance?area=${this.selectedArea}&year=${this.selectedYear}&month=${this.selectedMonth}`;
        const res = await axios.get(url);
        this.main_raw = res.data?.result || [];
        this.processMain();
      },
      async fetchDEPover(){
        const url = `${API_PREFIX}/DEPover?area=${this.selectedArea}&year=${this.selectedYear}&month=${this.selectedMonth}`;
        const res = await axios.get(url);
        this.DEP_raw = res.data?.result || [];
        this.processDEP();
      },
      async fetchQCDEPover(){
        const url = `${API_PREFIX}/QCDEPover?area=${this.selectedArea}&year=${this.selectedYear}&month=${this.selectedMonth}`;
        const res = await axios.get(url);
        this.QC_raw = res.data?.result || [];
        this.processQC();
      },
      async fetchCNDEPover(){
        const url = `${API_PREFIX}/CNDEPover?area=${CN_FIXED_AREA}`; // 固定 VN
        const res = await axios.get(url);
        this.ZH_raw = res.data?.result || [];
        this.processZH();
      },
      async fetchStffOverBonus(){
        const url = `${API_PREFIX}/stff_OVER_bonus?area=${this.selectedArea}&year=${this.selectedYear}&month=${this.selectedMonth}`;
        const res = await axios.get(url);
        this.stff_over_bonus = res.data?.result || [];
      },

      // === 整理分組 ===
      processMain(){
        const ds=new Set(), g={};
        (this.main_raw||[]).forEach(it=>{ ds.add(it.日期); const k=it.ERP班別代號; if(!g[k]) g[k]={}; g[k][it.日期]=it; });
        this.main_dates = Array.from(ds).sort(); this.main_group=g;
      },
      processDEP(){
        const ds=new Set(), g={};
        (this.DEP_raw||[]).forEach(it=>{ ds.add(it.日期); const k=it.ERP班別代號||it.部門||it.管理部門||it.部; if(!g[k]) g[k]={}; g[k][it.日期]=it; });
        this.DEP_dates = Array.from(ds).sort(); this.DEP_group=g;
      },
      processQC(){
        const ds=new Set(), g={};
        (this.QC_raw||[]).forEach(it=>{ ds.add(it.日期); const k=it.工號||it.員編||it.EMPID; if(!g[k]) g[k]={}; g[k][it.日期]=it; });
        this.QC_dates = Array.from(ds).sort(); this.QC_group=g;
      },
      processZH(){
        const ds=new Set(), g={};
        (this.ZH_raw||[]).forEach(it=>{ ds.add(it.日期); const k=it.工號||it.員編||it.EMPID; if(!g[k]) g[k]={}; g[k][it.日期]=it; });
        this.ZH_dates = Array.from(ds).sort(); this.ZH_group=g;
      },

      // === 取值 / 合計 ===
      getMain(erp,date,field){ return this.main_group?.[erp]?.[date]?.[field] ?? 0; },
      isDept(rows, dept){ const last=this.lastKey(this.main_dates); return (rows?.[last]?.['部門']||rows?.[last]?.['部'])==dept; },
      sumMain(date, field, dept){ let s=0; for(const erp in this.main_group){ const r=this.main_group[erp]?.[date]; if(!r) continue; if(dept && String(r['部門']||r['部'])!==String(dept)) continue; s+=Number(r[field]||0); } return s; },

      getDEP(key,date,field){ return this.DEP_group?.[key]?.[date]?.[field] ?? 0; },
      getQC(emp,date,field){ return this.QC_group?.[emp]?.[date]?.[field] ?? ''; },
      getQCV(emp,date,field){ return this.QC_group?.[emp]?.[date]?.[field] ?? 0; },
      getZH(emp,date,field){ return this.ZH_group?.[emp]?.[date]?.[field] ?? ''; },
      getZHV(emp,date,field){ return this.ZH_group?.[emp]?.[date]?.[field] ?? 0; },

      // === 匯出 Excel（多表 → 一活頁簿） ===
      exportExcel(){
        try{
          const wb = XLSX.utils.book_new();
          const tables = [
            { id:'WOOEVER3', name:'溢領_三部' },
            { id:'WOOEVER2', name:'溢領_二部' },
            { id:'WOOEVER1', name:'溢領_一部' },
            { id:'DEPOVER',  name:'部門彙總'   },
            { id:'QCDEPOVER',name:'QC部門'     },
            { id:'CNDEPOVER',name:'CN中國'     },
          ];
          tables.forEach(t=>{ const el=this.$el.querySelector('#'+t.id); if(!el) return; const ws=XLSX.utils.table_to_sheet(el,{raw:true}); XLSX.utils.book_append_sheet(wb, ws, t.name.substring(0,31)); });
          const fname = `溢領考績_${this.selectedArea}_${this.selectedYear}${this.selectedMonth}.xlsx`;
          XLSX.writeFile(wb, fname);
        }catch(err){ this.$message && this.$message.error('Excel 下載失敗：'+(err?.message||err)); }
      },
    },

    mounted(){ this.fetchAll(); }
  });
})();
