/* Overflow Performance (溢領考績) — MMS Vue2 component
 * Drop-in module to port NFC.html 的「溢領考績」到 MMS SPA
 * Dependencies: axios, (optional) Chart.js (via global `Chart`)
 */

Vue.component('overflow-performance-view', {
  template: `
<div class="container overflow-performance-page mt-3">
  <!-- Filters -->
  <div class="d-flex align-items-end gap-2 flex-wrap mb-3">
    <div>
      <label class="form-label mb-1">區域</label>
      <select v-model="selectedArea" class="form-select form-select-sm" @change="onFiltersChanged">
        <option>VN</option><option>ZH</option><option>TW</option><option>TC</option>
      </select>
    </div>
    <div>
      <label class="form-label mb-1">年份</label>
      <input type="number" v-model.number="selectedYear" class="form-control form-control-sm" @change="onFiltersChanged">
    </div>
    <div>
      <label class="form-label mb-1">月份</label>
      <input type="number" min="1" max="12" v-model.number="selectedMonth" class="form-control form-control-sm" @change="onFiltersChanged">
    </div>
    <div class="ms-auto d-flex align-items-center gap-2">
      <div class="form-check form-switch">
        <input class="form-check-input" type="checkbox" v-model="autoRefresh">
        <label class="form-check-label">自動刷新</label>
      </div>
      <button class="btn btn-primary btn-sm" @click="loadAll">查詢</button>
      <button class="btn btn-success btn-sm" @click="exportExcel">匯出Excel</button>
    </div>
  </div>

  <!-- 班別 × 日期：溢領金額矩陣 -->
  <div class="table-responsive mb-3" v-if="eff_overflow_dates.length">
    <table class="table table-sm table-bordered text-center align-middle mb-0">
      <thead class="sticky-top bg-light">
        <tr>
          <th style="white-space:nowrap">ERP班別代號</th>
          <th v-for="d in eff_overflow_dates" :key="d" style="white-space:nowrap">{{d}}</th>
          <th style="white-space:nowrap">合計</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(group, erp) in eff_overflow_grouped" :key="erp">
          <td class="text-start fw-semibold">{{erp}}</td>
          <td v-for="d in eff_overflow_dates" :key="erp+'-'+d">
            {{ fmtNumber(getOverflow(erp,d,'溢領金額')) }}
          </td>
          <td class="fw-bold">{{ fmtNumber(sumBy(group, '溢領金額')) }}</td>
        </tr>
      </tbody>
    </table>
  </div>

  <!-- 部門/品保 溢領總覽 -->
  <div class="row g-3">
    <div class="col-lg-6">
      <div class="card shadow-sm">
        <div class="card-header py-2">部門溢領 (DEPPER)</div>
        <div class="table-responsive">
          <table class="table table-sm table-hover mb-0">
            <thead><tr><th>部門</th><th>應領金額</th><th>溢領金額</th><th>溢領比例</th></tr></thead>
            <tbody>
              <tr v-for="r in dep_overflow" :key="r.部門">
                <td>{{r.部門}}</td>
                <td>{{zeroBlank(r.應領金額)}}</td>
                <td class="text-danger">{{zeroBlank(r.溢領金額)}}</td>
                <td class="text-danger">{{fmtPercent(r.溢領比例)}}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
    <div class="col-lg-6">
      <div class="card shadow-sm">
        <div class="card-header py-2">QC 溢領 (QCDEPPER)</div>
        <div class="table-responsive">
          <table class="table table-sm table-hover mb-0">
            <thead><tr><th>部門</th><th>應領金額</th><th>溢領金額</th><th>溢領比例</th></tr></thead>
            <tbody>
              <tr v-for="r in qcdep_overflow" :key="r.部門">
                <td>{{r.部門}}</td>
                <td>{{zeroBlank(r.應領金額)}}</td>
                <td class="text-danger">{{zeroBlank(r.溢領金額)}}</td>
                <td class="text-danger">{{fmtPercent(r.溢領比例)}}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <!-- 工單/品號追蹤 -->
  <div class="card shadow-sm mt-3">
    <div class="card-header py-2">工單/品號追蹤 (溢領比例)</div>
    <div class="p-2 d-flex gap-2">
      <input class="form-control form-control-sm" v-model.trim="partInput" placeholder="輸入工單或品號">
      <button class="btn btn-outline-primary btn-sm" @click="loadPartList">查詢</button>
    </div>
    <div class="table-responsive">
      <table class="table table-sm table-striped mb-0">
        <thead><tr><th>日期</th><th>工單</th><th>品號</th><th>品名</th><th>效率(%)</th><th>工單金額</th><th>應領金額</th><th>溢領金額</th><th>溢領比例</th></tr></thead>
        <tbody>
          <tr v-for="row in partRows" :key="row.WONum">
            <td>{{row.WODate}}</td><td>{{row.WONum}}</td><td>{{row.ItemNum}}</td><td class="text-start">{{row.ItemDesc}}</td>
            <td>{{row.eff}}</td><td>{{row.womoney ?? ''}}</td><td>{{row.應領金額 ?? ''}}</td><td class="text-danger">{{row.溢領金額 ?? ''}}</td><td class="text-danger">{{row.溢領比例 ?? ''}}</td>
          </tr>
        </tbody>
      </table>
    </div>
    <canvas id="over-chart" height="160" class="p-2"></canvas>
  </div>
</div>
  `,
  data(){return{
    perBase: (window.PERF_BASE || (location.origin + '/api')),
    selectedArea: 'VN',
    selectedYear: new Date().getFullYear(),
    selectedMonth: new Date().getMonth()+1,
    eff_overflow_data: [],
    eff_overflow_grouped: {},
    eff_overflow_dates: [],
    dep_overflow: [],
    qcdep_overflow: [],
    partInput:'',
    partRows: [],
    timerId:null,
    autoRefresh:true,
    chartInstance:null
  }},
  created(){ this.loadAll(); this.setupAutoRefresh(); },
  beforeDestroy(){ if(this.timerId) clearInterval(this.timerId); if(this.chartInstance){ this.chartInstance.destroy();}},
  methods:{
    setupAutoRefresh(){
      if(this.timerId) clearInterval(this.timerId);
      if(this.autoRefresh){
        this.timerId=setInterval(this.loadAll, 300000);
      }
    },
    onFiltersChanged(){
      this.loadAll();
      this.setupAutoRefresh();
    },
    async loadAll(){
      await Promise.all([this.fetchOverflow(), this.fetchDepOver(), this.fetchQCDepOver()]);
    },
    async fetchOverflow(){
      const url = `${this.perBase}/performance/overflowperformance?area=${this.selectedArea}&year=${this.selectedYear}&month=${this.selectedMonth}`;
      const {data} = await axios.get(url);
      const list = data?.result ?? data ?? [];
      this.eff_overflow_data = list;
      // group by ERP班別代號 & 日期
      const dateSet = new Set(), grouped={};
      list.forEach(it=>{
        const date = it.日期;
        const erp = it.ERP班別代號;
        if(!grouped[erp]) grouped[erp]={};
        grouped[erp][date]=it;
        dateSet.add(date);
      });
      this.eff_overflow_grouped=grouped;
      this.eff_overflow_dates=[...dateSet].sort();
    },
    async fetchDepOver(){
      const url = `${this.perBase}/performance/DEPover?area=${this.selectedArea}&year=${this.selectedYear}&month=${this.selectedMonth}`;
      const {data} = await axios.get(url);
      this.dep_overflow = data?.result ?? data ?? [];
    },
    async fetchQCDepOver(){
      const url = `${this.perBase}/performance/QCDEPover?area=${this.selectedArea}&year=${this.selectedYear}&month=${this.selectedMonth}`;
      const {data} = await axios.get(url);
      this.qcdep_overflow = data?.result ?? data ?? [];
    },
    getOverflow(erp,date,field){
      const g=this.eff_overflow_grouped?.[erp]?.[date];
      return g ? g[field] : '';
    },
    sumBy(group, field){
      return Object.values(group).reduce((sum, m)=> sum + (parseFloat(m[field]||0)||0), 0);
    },
    fmtNumber(v){ const n=Number(v); return (!n || n===0)? '' : n.toFixed(2); },
    fmtPercent(v){ const n=Number(v); if (!n || n===0) return ''; return (n).toFixed(2)+'%'; },
    zeroBlank(v){ const n=Number(v); return (!n || n===0)? '' : n; },
    async exportExcel(){
      // Minimal payload — align with NFC.html 的 PWEOVERWOTableToExcel
      const payload={
        MFG_DAY: `${this.selectedYear}/${String(this.selectedMonth).padStart(2,'0')}`,
        DEPPER: this.dep_overflow,
        QCDEPPER: this.qcdep_overflow,
        wopet: [], PPWO: [], stffbonus: [], wopet3: [], ZHDEPPER: [], EFF120: []
      };
      const url = (window.EXPORT_BASE || (location.origin + '/api')) + '/PWEOVERWOTableToExcel/to-excel';
      const blob = (await axios.post(url, payload, {responseType:'blob'})).data;
      const link=document.createElement('a');
      link.href=URL.createObjectURL(new Blob([blob], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
      link.download=`月溢領考核表_${payload.MFG_DAY}.xlsx`;
      link.click();
      URL.revokeObjectURL(link.href);
    },
    async loadPartList(){
      if(!this.partInput) return;
      const url = `${this.perBase}/READWOID/part_list?part=${encodeURIComponent(this.partInput)}`;
      const {data}= await axios.get(url);
      const rows = data?.result ?? data ?? [];
      this.partRows = rows;
      this.renderChart(rows);
    },
    renderChart(rows){
      if(this.chartInstance){ this.chartInstance.destroy(); this.chartInstance=null; }
      const labels=rows.slice().reverse().map(r=>r.WONum);
      const eff=rows.slice().reverse().map(r=>r.eff);
      const over=rows.slice().reverse().map(r=>r.溢領比例);
      const canvas=document.getElementById('over-chart');
      if(!canvas) return;
      const ctx=canvas.getContext('2d');
      if(!window.Chart || !ctx) return;
      this.chartInstance = new Chart(ctx, {
        type:'line',
        data:{ labels, datasets:[
          {label:'效率(%)', data: eff, yAxisID:'y'},
          {label:'溢領比例(%)', data: over, yAxisID:'y1'}
        ]},
        options:{ responsive:true, scales:{ y:{ beginAtZero:true }, y1:{ beginAtZero:true, position:'right' } } }
      });
    }
  }
});
