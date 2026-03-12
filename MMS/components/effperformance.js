// components/effperformance.js
// MMS 生產部 - 效率考績（完整功能 + 美化樣式）
//
// ✅ 串接 API（依既有 index.html 寫法）
//   - /api/performance/effperformance?area=..&year=..&month=..
//   - /api/performance/DEPPER?area=..&year=..&month=..
//   - /api/performance/QCDEPPER?area=..&year=..&month=..
//   - /api/performance/ZHDEPPER?area=VN   ← 原檔寫死 area=VN
//   - /api/performance/PPWO?area=..&year=..&month=..
//   - /api/performance/eff120?area=..&year=..&month=..
//   - /api/performance/stff_bonus?area=..&year=..&month=..
//
// ✅ 匯出：POST https://localhost:5001/api/PWEWOTableToExcel/to-excel
//   payload: { MFG_DAY, wopet, wopet3, PPWO, stffbonus, DEPPER, QCDEPPER, ZHDEPPER, EFF120 }
//
// ✅ UI：加載 /assets/effperformance.css 美化（表頭、按鈕、hover、高亮效率）
//

function loadCSS(href) {
  const link = document.createElement("link");
  link.rel = "stylesheet";
  link.href = href;
  document.head.appendChild(link);
}
loadCSS("assets/effperformance.css");

Vue.component('effperformance-view', {
  template: `
  <div class="eff-page" v-if="SHOW_eff_performance">
    <!-- Header -->
    <div class="eff-header shadow-sm p-2 mb-3 bg-light rounded d-flex flex-wrap align-items-center justify-content-between">
      <div class="d-flex align-items-center flex-wrap gap-2">
        <label>🌍 區域：</label>
        <select class="form-select form-select-sm w-auto" v-model="selectedArea" @change="effper">
          <option value="ZH">ZH</option>
          <option value="TC">TC</option>
          <option value="VN">VN</option>
          <option value="TW">TW</option>
        </select>

        <label>📅 年份：</label>
        <select class="form-select form-select-sm w-auto" v-model="selectedYear" @change="effper">
          <option v-for="y in [2024,2025,2026,2027]" :key="y">{{ y }}</option>
        </select>

        <label>月份：</label>
        <select class="form-select form-select-sm w-auto" v-model="selectedMonth" @change="effper">
          <option v-for="num in 12" :key="num" :value="String(num).padStart(2,'0')">
            {{ String(num).padStart(2, '0') }}
          </option>
        </select>
      </div>

      <div class="d-flex flex-wrap gap-2">
        <button v-for="(sheet, index) in sheets"
                :key="index"
                @click="currentSheet = index"
                class="btn btn-outline-primary btn-sm"
                :class="{ active: currentSheet === index }">
          {{ sheet.name }}
        </button>
        <button @click="downloadExcelWOPER" class="btn btn-success btn-sm">
          📤 導出 Excel
        </button>
      </div>
    </div>

    <!-- Tabs Content -->
    <div class="table-zone">
      <!-- WOPET2：二部（部門=2） -->
      <div v-show="currentSheet === 0">
        <h5 class="mb-2 fw-bold text-primary">二部效率表</h5>
        <table id="WOPET2" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
          <thead>
            <tr>
              <th rowspan="2">班別代號<br>Mã ca làm việc</th>
              <th v-for="(d, idx) in eff_performance_dates" :key="'header-'+d" :colspan="idx === eff_performance_dates.length - 1 ? 9 : 3">
                {{ d }}
              </th>
            </tr>
            <tr>
              <template v-for="(d, idx) in eff_performance_dates" :key="'sub-'+d">
                <th>標準工時</th>
                <th>實際工時</th>
                <th>效率</th>
                <template v-if="idx === eff_performance_dates.length - 1">
                  <th>累計負值效率</th>
                  <th>效率上升率</th>
                  <th>最終考核效率</th>
                  <th>當月考績</th>
                  <th>額外獎勵</th>
                  <th>下月累計負值</th>
                </template>
              </template>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(v, erp) in eff_performance_groupedData" :key="erp" v-if="eff_performance_getValue(erp, eff_performance_dates[2], '部門') === '2'">
              <td class="fw-bold bg-light">{{ erp }}</td>
              <template v-for="(date, i) in eff_performance_dates" :key="'d-'+erp+'-'+date">
                <td>{{ eff_performance_formatNumber(eff_performance_getValue(erp, date, '人時標準')) }}</td>
                <td>{{ eff_performance_formatNumber(eff_performance_getValue(erp, date, '人時實際')) }}</td>
                <td :class="effColorClass(eff_performance_getValue(erp, date, '效率'))">
                  {{ eff_performance_formatNumber2(eff_performance_getValue(erp, date, '效率')) }}
                </td>
                <template v-if="i === eff_performance_dates.length - 1">
                  <td>{{ eff_performance_formatNumber2(eff_performance_getValue(erp, date, '累計')) }}</td>
                  <td>{{ eff_performance_formatNumber2(eff_performance_getValue(erp, date, '上升率')) }}</td>
                  <td>{{ eff_performance_formatNumber2(eff_performance_getValue(erp, date, '績效')) }}</td>
                  <td>{{ eff_performance_getValue(erp, date, '考績') === '0' ? '-' : eff_performance_getValue(erp, date, '考績') }}</td>
                  <td>{{ eff_performance_getValue(erp, date, '考績2') === '0' ? '-' : eff_performance_getValue(erp, date, '考績2') }}</td>
                  <td>{{ eff_performance_formatNumber2(eff_performance_getValue(erp, date, '下月累計負值')) }}</td>
                </template>
              </template>
            </tr>
            <tr>
              <td>合計</td>
              <template v-for="(date, i) in eff_performance_dates" :key="'t2-'+date">
                <td>{{ eff_performance_formatNumber(eff_performance_sumColumn(eff_performance_groupedData, date, '人時標準', '2')) }}</td>
                <td>{{ eff_performance_formatNumber(eff_performance_sumColumn(eff_performance_groupedData, date, '人時實際', '2')) }}</td>
                <td>{{ eff_performance_formatNumber2(eff_performance_formatNumber(eff_performance_sumColumn(eff_performance_groupedData, date, '人時標準', '2')) / eff_performance_formatNumber(eff_performance_sumColumn(eff_performance_groupedData, date, '人時實際', '2')) * 100) }}</td>
              </template>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- WOPET3：三部（部門=3） -->
      <div v-show="currentSheet === 1">
        <h5 class="mb-2 fw-bold text-primary">三部效率表</h5>
        <table id="WOPET3" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
          <thead>
            <tr>
              <th rowspan="2">班別代號<br>Mã ca làm việc</th>
              <th v-for="(d, idx) in eff_performance_dates" :key="'header3-'+d" :colspan="idx === eff_performance_dates.length - 1 ? 9 : 3">
                {{ d }}
              </th>
            </tr>
            <tr>
              <template v-for="(d, idx) in eff_performance_dates" :key="'sub3-'+d">
                <th>標準工時</th>
                <th>實際工時</th>
                <th>效率</th>
                <template v-if="idx === eff_performance_dates.length - 1">
                  <th>累計負值效率</th>
                  <th>效率上升率</th>
                  <th>最終考核效率</th>
                  <th>當月考績</th>
                  <th>額外獎勵</th>
                  <th>下月累計負值</th>
                </template>
              </template>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(v, erp) in eff_performance_groupedData" :key="erp" v-if="eff_performance_getValue(erp, eff_performance_dates[2], '部門') === '3'">
              <td class="fw-bold bg-light">{{ erp }}</td>
              <template v-for="(date, i) in eff_performance_dates" :key="'d3-'+erp+'-'+date">
                <td>{{ eff_performance_formatNumber(eff_performance_getValue(erp, date, '人時標準')) }}</td>
                <td>{{ eff_performance_formatNumber(eff_performance_getValue(erp, date, '人時實際')) }}</td>
                <td :class="effColorClass(eff_performance_getValue(erp, date, '效率'))">
                  {{ eff_performance_formatNumber2(eff_performance_getValue(erp, date, '效率')) }}
                </td>
                <template v-if="i === eff_performance_dates.length - 1">
                  <td>{{ eff_performance_formatNumber2(eff_performance_getValue(erp, date, '累計')) }}</td>
                  <td>{{ eff_performance_formatNumber2(eff_performance_getValue(erp, date, '上升率')) }}</td>
                  <td>{{ eff_performance_formatNumber2(eff_performance_getValue(erp, date, '績效')) }}</td>
                  <td>{{ eff_performance_getValue(erp, date, '考績') === '0' ? '-' : eff_performance_getValue(erp, date, '考績') }}</td>
                  <td>{{ eff_performance_getValue(erp, date, '考績2') === '0' ? '-' : eff_performance_getValue(erp, date, '考績2') }}</td>
                  <td>{{ eff_performance_formatNumber2(eff_performance_getValue(erp, date, '下月累計負值')) }}</td>
                </template>
              </template>
            </tr>
            <tr>
              <td>合計</td>
              <template v-for="(date, i) in eff_performance_dates" :key="'t3-'+date">
                <td>{{ eff_performance_formatNumber(eff_performance_sumColumn(eff_performance_groupedData, date, '人時標準', '3')) }}</td>
                <td>{{ eff_performance_formatNumber(eff_performance_sumColumn(eff_performance_groupedData, date, '人時實際', '3')) }}</td>
                <td>{{ eff_performance_formatNumber2(eff_performance_formatNumber(eff_performance_sumColumn(eff_performance_groupedData, date, '人時標準', '3')) / eff_performance_formatNumber(eff_performance_sumColumn(eff_performance_groupedData, date, '人時實際', '3')) * 100) }}</td>
              </template>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- PPWO -->
      <div v-show="currentSheet === 2">
        <h5 class="mb-2 fw-bold text-primary">PPWO 工單明細</h5>
        <table id="PPWO" class="eff-table table table-bordered table-sm text-center">
          <thead class="table-dark">
            <tr>
              <th>工單</th>
              <th>班別代號</th>
              <th>料號</th>
              <th>數量</th>
              <th>標準工時</th>
              <th>實際工時</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="item in eff_performance_ppwodata" :key="item.ID">
              <td>{{ item.WONUM }}</td>
              <td>{{ item.SHIFT }}</td>
              <td>{{ item.ITEMNUM }}</td>
              <td>{{ item.QTY }}</td>
              <td>{{ Number(item.SWH||0).toFixed(2) }}</td>
              <td>{{ Number(item.AWH||0).toFixed(2) }}</td>
            </tr>
            <tr v-if="eff_performance_ppwodata.length">
              <td>比例</td>
              <td>{{ eff_performance_ppwodata[0].Total_PP_T }}/{{ eff_performance_ppwodata[0].Total_WONUM }}</td>
              <td colspan="2">合計</td>
              <td>{{ totalSWH.toFixed(2) }}</td>
              <td>{{ totalAWH.toFixed(2) }}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- DEPPER（主管績效） -->
      <div v-show="currentSheet === 3">
        <h5 class="mb-2 fw-bold text-primary">主管績效</h5>
        <table id="DEPPER" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
          <thead>
            <tr>
              <th rowspan="2">管理部門</th>
              <th rowspan="2">管理班別</th>
              <th rowspan="2">工號</th>
              <th rowspan="2">姓名</th>
              <th rowspan="2">上班天數</th>
              <th rowspan="2">職稱</th>
              <th v-for="(d, idx) in DEP_eff_performance_dates" :key="'header-dep-'+d" :colspan="idx === DEP_eff_performance_dates.length - 1 ? 11 : 1">
                {{ d }}
              </th>
            </tr>
            <tr>
              <template v-for="(d, idx) in DEP_eff_performance_dates" :key="'sub-dep-'+d">
                <th>效率</th>
                <template v-if="idx === DEP_eff_performance_dates.length - 1">
                  <th>累計負值效率</th>
                  <th>效率上升率</th>
                  <th>最終考核效率</th>
                  <th>當月考績</th>
                  <th>額外獎勵</th>
                  <th>考績獎金</th>
                  <th>額外考績獎金</th>
                  <th>A15拉管理福利金</th>
                  <th>最終獎金</th>
                  <th>下月累計負值效率</th>
                </template>
              </template>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(v, erp) in DEP_eff_performance_groupedData" :key="erp">
              <td>{{ DEP_eff_performance_getValue(erp, lastDEPDate, '部') }}</td>
              <td>{{ DEP_eff_performance_getValue(erp, lastDEPDate, '班') }}</td>
              <td>{{ erp }}</td>
              <td>{{ DEP_eff_performance_getValue(erp, lastDEPDate, '姓名') }}</td>
              <td>{{ DEP_eff_performance_getValue(erp, lastDEPDate, 'LAMDAY_STR') }}</td>
              <td>{{ DEP_eff_performance_getValue(erp, lastDEPDate, '崗位') }}</td>
              <template v-for="(date, i) in DEP_eff_performance_dates" :key="'dep-'+erp+'-'+date">
                <td :class="effColorClass(DEP_eff_performance_getValue(erp, date, '效率'))">
                  {{ eff_performance_formatNumber2(DEP_eff_performance_getValue(erp, date, '效率')) }}
                </td>
                <template v-if="i === DEP_eff_performance_dates.length - 1">
                  <td>{{ eff_performance_formatNumber2(DEP_eff_performance_getValue(erp, date, '累計')) }}</td>
                  <td>{{ eff_performance_formatNumber2(DEP_eff_performance_getValue(erp, date, '上升率')) }}</td>
                  <td>{{ eff_performance_formatNumber2(DEP_eff_performance_getValue(erp, date, '績效')) }}</td>
                  <td>{{ DEP_eff_performance_getValue(erp, date, '考績') === '0' ? '-' : DEP_eff_performance_getValue(erp, date, '考績') }}</td>
                  <td>{{ DEP_eff_performance_getValue(erp, date, '額外考績') === '0' ? '-' : DEP_eff_performance_getValue(erp, date, '額外考績') }}</td>
                  <td>{{ formatCurrency(DEP_eff_performance_getValue(erp, date, '考績獎金')) }}</td>
                  <td>{{ formatCurrency(DEP_eff_performance_getValue(erp, date, '額外考績獎金')) }}</td>
                  <td>{{ formatCurrency(DEP_eff_performance_getValue(erp, date, 'A15拉管理福利金')) }}</td>
                  <td>{{ formatCurrency(DEP_eff_performance_getValue(erp, date, '最終獎金')) }}</td>
                  <td>{{ eff_performance_formatNumber2(DEP_eff_performance_getValue(erp, date, '下月累計負值')) }}</td>
                </template>
              </template>
            </tr>
            <tr>
              <td colspan="17">Total</td>
              <td>{{ formatCurrency(totalFinalBonus) }}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- 員工獎金（stff_bonus） -->
      <div v-show="currentSheet === 4">
        <h5 class="mb-2 fw-bold text-primary">員工獎金</h5>
        <table id="stffbonus" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
          <thead>
            <tr>
              <th>班別代號</th>
              <th>工號</th>
              <th>姓名</th>
              <th>工作天數</th>
              <th>請假天數</th>
              <th>崗位</th>
              <th>績效等級-I</th>
              <th>獎金-I</th>
              <th>績效等級-II</th>
              <th>獎金-II</th>
              <th>最終獎金</th>
            </tr>
          </thead>
          <tbody>
            <template v-for="(group, shift) in groupedByShift">
              <tr v-for="item in group" :key="item.ID">
                <td>{{ item.SHIFT }}</td>
                <td>{{ item.JOB }}</td>
                <td>{{ item.USERNAME }}</td>
                <td>{{ item.LAMDAY }}</td>
                <td>{{ item.請假天數 }}</td>
                <td>{{ item.IMPORT }}</td>
                <td>{{ item.note }}</td>
                <td>{{ formatCurrency(item.績效) }}</td>
                <td>{{ item.note2 }}</td>
                <td>{{ formatCurrency(item.額外績效) }}</td>
                <td>{{ formatCurrency(item.實際獎金) }}</td>
              </tr>
              <tr>
                <td colspan="9">部門 {{ shift }} 統計</td>
                <td colspan="2">{{ formatCurrency(calculateShiftBonus(shift)) }}</td>
              </tr>
            </template>
            <tr>
              <td colspan="9">Total</td>
              <td colspan="2">{{ formatCurrency(totalBonus) }}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- QCDEPPER（產線外人員績效） -->
      <div v-show="currentSheet === 5">
        <h5 class="mb-2 fw-bold text-primary">產線外人員績效</h5>
        <table id="QCDEPPER" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
          <thead>
            <tr>
              <th rowspan="2">管理部門</th>
              <th rowspan="2">管理班別</th>
              <th rowspan="2">工號</th>
              <th rowspan="2">姓名</th>
              <th rowspan="2">職稱</th>
              <th v-for="(d, idx) in QCDEP_eff_performance_dates" :key="'header-qc-'+d" :colspan="idx === QCDEP_eff_performance_dates.length - 1 ? 11 : 1">
                {{ d }}
              </th>
            </tr>
            <tr>
              <template v-for="(d, idx) in QCDEP_eff_performance_dates" :key="'sub-qc-'+d">
                <th>效率</th>
                <template v-if="idx === QCDEP_eff_performance_dates.length - 1">
                  <th>累計負值效率</th>
                  <th>效率上升率</th>
                  <th>最終考核效率</th>
                  <th>當月考績</th>
                  <th>額外獎勵</th>
                  <th>考績獎金</th>
                  <th>額外考績獎金</th>
                  <th>A15拉管理福利金</th>
                  <th>最終獎金</th>
                  <th>下月累計負值效率</th>
                </template>
              </template>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(v, erp) in QCDEP_eff_performance_groupedData" :key="erp">
              <td>{{ QCDEP_eff_performance_getValue(erp, lastQCDEPDate, '部') }}</td>
              <td>{{ QCDEP_eff_performance_getValue(erp, lastQCDEPDate, '班') }}</td>
              <td>{{ erp }}</td>
              <td>{{ QCDEP_eff_performance_getValue(erp, lastQCDEPDate, '姓名') }}</td>
              <td>{{ QCDEP_eff_performance_getValue(erp, lastQCDEPDate, '崗位') }}</td>
              <template v-for="(date, i) in QCDEP_eff_performance_dates" :key="'qc-'+erp+'-'+date">
                <td :class="effColorClass(QCDEP_eff_performance_getValue(erp, date, '效率'))">
                  {{ eff_performance_formatNumber2(QCDEP_eff_performance_getValue(erp, date, '效率')) }}
                </td>
                <template v-if="i === QCDEP_eff_performance_dates.length - 1">
                  <td>{{ eff_performance_formatNumber2(QCDEP_eff_performance_getValue(erp, date, '累計')) }}</td>
                  <td>{{ eff_performance_formatNumber2(QCDEP_eff_performance_getValue(erp, date, '上升率')) }}</td>
                  <td>{{ eff_performance_formatNumber2(QCDEP_eff_performance_getValue(erp, date, '績效')) }}</td>
                  <td>{{ QCDEP_eff_performance_getValue(erp, date, '考績') === '0' ? '-' : QCDEP_eff_performance_getValue(erp, date, '考績') }}</td>
                  <td>{{ QCDEP_eff_performance_getValue(erp, date, '額外考績') === '0' ? '-' : QCDEP_eff_performance_getValue(erp, date, '額外考績') }}</td>
                  <td>{{ formatCurrency(QCDEP_eff_performance_getValue(erp, date, '考績獎金')) }}</td>
                  <td>{{ formatCurrency(QCDEP_eff_performance_getValue(erp, date, '額外考績獎金')) }}</td>
                  <td>{{ formatCurrency(QCDEP_eff_performance_getValue(erp, date, 'A15拉管理福利金')) }}</td>
                  <td>{{ formatCurrency(QCDEP_eff_performance_getValue(erp, date, '最終獎金')) }}</td>
                  <td>{{ eff_performance_formatNumber2(QCDEP_eff_performance_getValue(erp, date, '下月累計負值')) }}</td>
                </template>
              </template>
            </tr>
            <tr>
              <td colspan="16">Total</td>
              <td>{{ formatCurrency(totalQCFinalBonus) }}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- 效率<120%：eff120 -->
      <div v-show="currentSheet === 6">
        <h5 class="mb-2 fw-bold text-primary">效率  120% 明細</h5>
        <table id="eff120" class="eff-table table table-bordered table-sm text-center">
          <thead class="table-dark">
            <tr>
              <th>班別代號</th>
              <th>料號</th>
              <th>客戶料號</th>
              <th>工單數量</th>
              <th>數量</th>
              <th>標準工時</th>
              <th>實際工時</th>
              <th>效率</th>
              <th>建議調整標準工時</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="item in eff_performance_eff120data" :key="item.ID">
              <td>{{ item.SHIFT }}</td>
              <td>{{ item.ITEMNUM }}</td>
              <td>{{ item.wad01 }}</td>
              <td>{{ item.WONUM }}</td>
              <td>{{ item.QTY }}</td>
              <td>{{ Number(item.SWH || 0).toFixed(2) }}</td>
              <td>{{ Number(item.AWH || 0).toFixed(2) }}</td>
              <td>{{ Number((item.eff || 0) * 100).toFixed(2) }}%</td>
              <td>{{ item.sug }}</td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- ZHDEPPER（綜合部 ZH） -->
      <div v-show="currentSheet === 7">
        <h5 class="mb-2 fw-bold text-primary">綜合部（ZH）</h5>
        <table id="ZHDEPPER" class="eff-table table table-bordered table-hover table-sm align-middle text-center">
          <thead>
            <tr>
              <th rowspan="2">管理部門</th>
              <th rowspan="2">管理班別</th>
              <th rowspan="2">工號</th>
              <th rowspan="2">姓名</th>
              <th rowspan="2">職稱</th>
              <th v-for="(d, idx) in ZHDEP_eff_performance_dates" :key="'header-zh-'+d" :colspan="idx === ZHDEP_eff_performance_dates.length - 1 ? 11 : 1">
                {{ d }}
              </th>
            </tr>
            <tr>
              <template v-for="(d, idx) in ZHDEP_eff_performance_dates" :key="'sub-zh-'+d">
                <th>效率</th>
                <template v-if="idx === ZHDEP_eff_performance_dates.length - 1">
                  <th>累計負值效率</th>
                  <th>效率上升率</th>
                  <th>最終考核效率</th>
                  <th>當月考績</th>
                  <th>額外獎勵</th>
                  <th>考績獎金</th>
                  <th>額外考績獎金</th>
                  <th>A15拉管理福利金</th>
                  <th>最終獎金</th>
                  <th>下月累計負值效率</th>
                </template>
              </template>
            </tr>
          </thead>
          <tbody>
            <tr v-for="(v, erp) in ZHDEP_eff_performance_groupedData" :key="erp">
              <td>{{ ZHDEP_eff_performance_getValue(erp, lastZHDEPDate, '部') }}</td>
              <td>{{ ZHDEP_eff_performance_getValue(erp, lastZHDEPDate, '班') }}</td>
              <td>{{ erp }}</td>
              <td>{{ ZHDEP_eff_performance_getValue(erp, lastZHDEPDate, '姓名') }}</td>
              <td>{{ ZHDEP_eff_performance_getValue(erp, lastZHDEPDate, '崗位') }}</td>
              <template v-for="(date, i) in ZHDEP_eff_performance_dates" :key="'zh-'+erp+'-'+date">
                <td :class="effColorClass(ZHDEP_eff_performance_getValue(erp, date, '效率'))">
                  {{ eff_performance_formatNumber2(ZHDEP_eff_performance_getValue(erp, date, '效率')) }}
                </td>
                <template v-if="i === ZHDEP_eff_performance_dates.length - 1">
                  <td>{{ eff_performance_formatNumber2(ZHDEP_eff_performance_getValue(erp, date, '累計')) }}</td>
                  <td>{{ eff_performance_formatNumber2(ZHDEP_eff_performance_getValue(erp, date, '上升率')) }}</td>
                  <td>{{ eff_performance_formatNumber2(ZHDEP_eff_performance_getValue(erp, date, '績效')) }}</td>
                  <td>{{ ZHDEP_eff_performance_getValue(erp, date, '考績') === '0' ? '-' : ZHDEP_eff_performance_getValue(erp, date, '考績') }}</td>
                  <td>{{ ZHDEP_eff_performance_getValue(erp, date, '額外考績') === '0' ? '-' : ZHDEP_eff_performance_getValue(erp, date, '額外考績') }}</td>
                  <td>{{ formatCurrency(ZHDEP_eff_performance_getValue(erp, date, '考績獎金')) }}</td>
                  <td>{{ formatCurrency(ZHDEP_eff_performance_getValue(erp, date, '額外考績獎金')) }}</td>
                  <td>{{ formatCurrency(ZHDEP_eff_performance_getValue(erp, date, 'A15拉管理福利金')) }}</td>
                  <td>
                    {{
                      (total => (isNaN(total) || total === 0) ? '-' : formatCurrency(total))(
                        parseFloat(ZHDEP_eff_performance_getValue(erp, date, '考績獎金') || '0') +
                        parseFloat(ZHDEP_eff_performance_getValue(erp, date, '額外考績獎金') || '0') +
                        parseFloat(ZHDEP_eff_performance_getValue(erp, date, 'A15拉管理福利金') || '0')
                      )
                    }}
                  </td>
                  <td>{{ eff_performance_formatNumber2(ZHDEP_eff_performance_getValue(erp, date, '下月累計負值')) }}</td>
                </template>
              </template>
            </tr>
            <tr>
              <td colspan="16">Total</td>
              <td>{{ formatCurrency(totalZHFinalBonus) }}</td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  </div>
  `,

  data() {
    const today = new Date();
    return {
      // 顯示控制
      SHOW_eff_performance: true,

      // 條件
      selectedArea: 'VN',
      selectedYear: String(today.getFullYear()),
      selectedMonth: String(today.getMonth() + 1).padStart(2, '0'),

      // Tabs
      currentSheet: 0,
      sheets: [
        { name: '二部' }, { name: '三部' }, { name: 'PPWO' }, { name: '主管績效' },
        { name: '員工獎金' }, { name: '產線外人員績效' }, { name: '效率>120%料號' }, { name: '連續3月大於100%料號' }
      ],

      // 主表
      eff_performance_data: [],
      eff_performance_dates: [],
      eff_performance_groupedData: {},

      // PPWO
      eff_performance_ppwodata: [],

      // DEPPER
      DEP_eff_performance_data: [],
      DEP_eff_performance_dates: [],
      DEP_eff_performance_groupedData: {},

      // QCDEPPER
      QCDEP_eff_performance_data: [],
      QCDEP_eff_performance_dates: [],
      QCDEP_eff_performance_groupedData: {},

      // ZHDEPPER
      ZHDEP_eff_performance_data: [],
      ZHDEP_eff_performance_dates: [],
      ZHDEP_eff_performance_groupedData: {},

      // eff120
      eff_performance_eff120data: [],

      // stff_bonus
      eff_stff_bonus: []
    };
  },

  computed: {
    // 末月份（顯示人名/職稱等用）
    lastDEPDate() {
      return this.DEP_eff_performance_dates.length ? this.DEP_eff_performance_dates[this.DEP_eff_performance_dates.length - 1] : '';
    },
    lastQCDEPDate() {
      return this.QCDEP_eff_performance_dates.length ? this.QCDEP_eff_performance_dates[this.QCDEP_eff_performance_dates.length - 1] : '';
    },
    lastZHDEPDate() {
      return this.ZHDEP_eff_performance_dates.length ? this.ZHDEP_eff_performance_dates[this.ZHDEP_eff_performance_dates.length - 1] : '';
    },

    // PPWO 合計
    totalSWH() { return this.eff_performance_ppwodata.reduce((t, x) => t + parseFloat(x.SWH || 0), 0); },
    totalAWH() { return this.eff_performance_ppwodata.reduce((t, x) => t + parseFloat(x.AWH || 0), 0); },

    // 員工獎金 grouping / total
    groupedByShift() {
      const map = {};
      (this.eff_stff_bonus || []).forEach(i => {
        const k = i.SHIFT || '-';
        if (!map[k]) map[k] = [];
        map[k].push(i);
      });
      return map;
    },
    totalBonus() {
      return (this.eff_stff_bonus || []).reduce((t, x) => t + parseFloat(x.實際獎金 || 0), 0);
    },

    // DEPPER / QCDEPPER / ZHDEPPER 總最終獎金
    totalFinalBonus() {
      const last = this.lastDEPDate; if (!last) return 0;
      let sum = 0;
      Object.keys(this.DEP_eff_performance_groupedData).forEach(erp => {
        const row = this.DEP_eff_performance_groupedData[erp][last];
        if (row) sum += parseFloat(row['最終獎金'] || 0);
      });
      return sum;
    },
    totalQCFinalBonus() {
      const last = this.lastQCDEPDate; if (!last) return 0;
      let sum = 0;
      Object.keys(this.QCDEP_eff_performance_groupedData).forEach(erp => {
        const row = this.QCDEP_eff_performance_groupedData[erp][last];
        if (row) sum += parseFloat(row['最終獎金'] || 0);
      });
      return sum;
    },
    totalZHFinalBonus() {
      const last = this.lastZHDEPDate; if (!last) return 0;
      let sum = 0;
      Object.keys(this.ZHDEP_eff_performance_groupedData).forEach(erp => {
        const row = this.ZHDEP_eff_performance_groupedData[erp][last];
        if (row) sum += parseFloat(row['最終獎金'] || 0);
      });
      return sum;
    }
  },

  mounted() {
    this.effper();
  },

  methods: {
    // 效率上/下限著色
    effColorClass(val) {
      const n = parseFloat(val);
      if (isNaN(n)) return '';
      if (n >= 100) return 'bg-success text-white fw-bold';
      if (n < 80)  return 'bg-danger text-white fw-bold';
      return '';
    },

    // ---- API 依序載入 ----
    async effper() {
      await this.eff_performance_fetchData();        // 主表
      this.eff_performance_processData();

      await this.eff_performance_ppwo();             // PPWO
      await this.eff_performance_eff120();           // <120%
      await this.eff_stff_bonus_fetchData();         // 員工獎金

      await this.DEP_eff_performance_fetchData();    // 管理部 二部
      this.DEP_eff_performance_processData();

      await this.QCDEP_eff_performance_fetchData();  // 管理部 QC
      this.QCDEP_eff_performance_processData();

      await this.ZHDEP_eff_performance_fetchData();  // 綜合部 ZH（area=VN）
      this.ZHDEP_eff_performance_processData();
    },

    async eff_performance_fetchData() {
      const { selectedArea, selectedYear, selectedMonth } = this;
      const url = `http://192.168.207.17:5024/api/performance/effperformance?area=${selectedArea}&year=${selectedYear}&month=${selectedMonth}`;
      const res = await axios.get(url);
      this.eff_performance_data = res.data.result || [];
    },
    async eff_performance_ppwo() {
      const { selectedArea, selectedYear, selectedMonth } = this;
      const url = `http://192.168.207.17:5024/api/performance/PPWO?area=${selectedArea}&year=${selectedYear}&month=${selectedMonth}`;
      const res = await axios.get(url);
      this.eff_performance_ppwodata = res.data.result || [];
    },
    async eff_performance_eff120() {
      const { selectedArea, selectedYear, selectedMonth } = this;
      const url = `http://192.168.207.17:5024/api/performance/eff120?area=${selectedArea}&year=${selectedYear}&month=${selectedMonth}`;
      const res = await axios.get(url);
      this.eff_performance_eff120data = res.data.result || [];
    },
    async eff_stff_bonus_fetchData() {
      const { selectedArea, selectedYear, selectedMonth } = this;
      const url = `http://192.168.207.17:5024/api/performance/stff_bonus?area=${selectedArea}&year=${selectedYear}&month=${selectedMonth}`;
      const res = await axios.get(url);
      this.eff_stff_bonus = res.data.result || [];
    },
    async DEP_eff_performance_fetchData() {
      const { selectedArea, selectedYear, selectedMonth } = this;
      const url = `http://192.168.207.17:5024/api/performance/DEPPER?area=${selectedArea}&year=${selectedYear}&month=${selectedMonth}`;
      const res = await axios.get(url);
      this.DEP_eff_performance_data = res.data.result || [];
    },
    async QCDEP_eff_performance_fetchData() {
      const { selectedArea, selectedYear, selectedMonth } = this;
      const url = `http://192.168.207.17:5024/api/performance/QCDEPPER?area=${selectedArea}&year=${selectedYear}&month=${selectedMonth}`;
      const res = await axios.get(url);
      this.QCDEP_eff_performance_data = res.data.result || [];
    },
    async ZHDEP_eff_performance_fetchData() {
      // 注意：原始檔案固定 area=VN
      const url = `http://192.168.207.17:5024/api/performance/ZHDEPPER?area=VN`;
      const res = await axios.get(url);
      this.ZHDEP_eff_performance_data = res.data.result || [];
    },

    // ---- 資料整理/分組 ----
    eff_performance_processData() {
      const dateSet = new Set();
      const grouped = {};
      (this.eff_performance_data || []).forEach(item => {
        dateSet.add(item.日期);
        if (!grouped[item.ERP班別代號]) grouped[item.ERP班別代號] = {};
        grouped[item.ERP班別代號][item.日期] = item;
      });
      this.eff_performance_dates = Array.from(dateSet).sort();
      this.eff_performance_groupedData = grouped;
    },
    DEP_eff_performance_processData() {
      const dateSet = new Set();
      const grouped = {};
      (this.DEP_eff_performance_data || []).forEach(item => {
        dateSet.add(item.日期);
        if (!grouped[item.ERP班別代號]) grouped[item.ERP班別代號] = {};
        grouped[item.ERP班別代號][item.日期] = item;
      });
      this.DEP_eff_performance_dates = Array.from(dateSet).sort();
      this.DEP_eff_performance_groupedData = grouped;
    },
    QCDEP_eff_performance_processData() {
      const dateSet = new Set();
      const grouped = {};
      (this.QCDEP_eff_performance_data || []).forEach(item => {
        dateSet.add(item.日期);
        if (!grouped[item.ERP班別代號]) grouped[item.ERP班別代號] = {};
        grouped[item.ERP班別代號][item.日期] = item;
      });
      this.QCDEP_eff_performance_dates = Array.from(dateSet).sort();
      this.QCDEP_eff_performance_groupedData = grouped;
    },
    ZHDEP_eff_performance_processData() {
      const dateSet = new Set();
      const grouped = {};
      (this.ZHDEP_eff_performance_data || []).forEach(item => {
        dateSet.add(item.日期);
        if (!grouped[item.ERP班別代號]) grouped[item.ERP班別代號] = {};
        grouped[item.ERP班別代號][item.日期] = item;
      });
      this.ZHDEP_eff_performance_dates = Array.from(dateSet).sort();
      this.ZHDEP_eff_performance_groupedData = grouped;
    },

    // ---- 取值/加總/格式 ----
    eff_performance_getValue(erp, date, field) {
      const g = this.eff_performance_groupedData[erp];
      if (!g) return '-';
      const row = g[date];
      return row ? (row[field] ?? '-') : '-';
    },
    DEP_eff_performance_getValue(erp, date, field) {
      const g = this.DEP_eff_performance_groupedData[erp];
      if (!g) return '-';
      const row = g[date];
      return row ? (row[field] ?? '-') : '-';
    },
    QCDEP_eff_performance_getValue(erp, date, field) {
      const g = this.QCDEP_eff_performance_groupedData[erp];
      if (!g) return '-';
      const row = g[date];
      return row ? (row[field] ?? '-') : '-';
    },
    ZHDEP_eff_performance_getValue(erp, date, field) {
      const g = this.ZHDEP_eff_performance_groupedData[erp];
      if (!g) return '-';
      const row = g[date];
      return row ? (row[field] ?? '-') : '-';
    },

    eff_performance_sumColumn(grouped, date, field, deptNo) {
      let sum = 0;
      Object.keys(grouped).forEach(k => {
        const row = grouped[k] && grouped[k][date];
        if (row && String(row['部門']) === String(deptNo)) {
          sum += parseFloat(row[field] || 0);
        }
      });
      return sum;
    },

    eff_performance_formatNumber(value) {
      const n = Number(value);
      return isNaN(n) ? '-' : n.toFixed(2);
    },
    eff_performance_formatNumber2(value) {
      const n = Number(value);
      return isNaN(n) ? '-' : n.toFixed(2); // 主表/部門表顯示「值」，不加 %
    },
    formatCurrency(value) {
      const n = Number(value);
      if (isNaN(n)) return '-';
      return n.toLocaleString('en-US', { maximumFractionDigits: 0 });
    },
    calculateShiftBonus(shift) {
      const list = this.groupedByShift[shift] || [];
      return list.reduce((t, x) => t + parseFloat(x.實際獎金 || 0), 0);
    },

    // ---- 匯出 ----
    async downloadExcelWOPER() {
      const PERData = {
        MFG_DAY: `${this.selectedYear}/${this.selectedMonth}`,
        wopet:  this._tableJsonFromWOPET(2), // 二部
        wopet3: this._tableJsonFromWOPET(3), // 三部
        PPWO:   this.eff_performance_ppwodata,
        stffbonus: this.eff_stff_bonus,
        DEPPER:   this._flatGroupToArray(this.DEP_eff_performance_groupedData),
        QCDEPPER: this._flatGroupToArray(this.QCDEP_eff_performance_groupedData),
        ZHDEPPER: this._flatGroupToArray(this.ZHDEP_eff_performance_groupedData),
        EFF120:   this.eff_performance_eff120data
      };

      try {
        const resp = await axios.post('https://localhost:5001/api/PWEWOTableToExcel/to-excel', PERData, { responseType: 'blob' });
        const blob = new Blob([resp.data], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = '月效率考核表 Mẫu đánh giá hiệu quả hàng tháng .xlsx';
        link.click();
      } catch (err) {
        console.error('Excel 匯出失敗:', err);
      }
    },

    _flatGroupToArray(grouped) {
      const out = [];
      Object.keys(grouped || {}).forEach(erp => {
        Object.keys(grouped[erp] || {}).forEach(date => out.push(grouped[erp][date]));
      });
      return out;
    },
    _tableJsonFromWOPET(deptNo) {
      const out = [];
      Object.keys(this.eff_performance_groupedData || {}).forEach(erp => {
        this.eff_performance_dates.forEach(date => {
          const row = this.eff_performance_groupedData[erp][date];
          if (row && String(row['部門']) === String(deptNo)) out.push(row);
        });
      });
      return out;
    }
  }
});
