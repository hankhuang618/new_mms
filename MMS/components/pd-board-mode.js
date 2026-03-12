
// components/pd-board-mode.js
Vue.component('pd-board-mode-view', {
  template: `
  <div class="container mt-4">
    <h4 class="mb-3"><i class="el-icon-setting"></i> 產線看板調整</h4>

    <el-card shadow="hover">
      <div class="d-flex align-items-center">
        <span class="mr-3">看板模式開關：</span>
        <el-switch
          v-model="isOpen"
          active-text="優化（Model=1）"
          inactive-text="正常（Model=0）"
          @change="onToggle"
          :loading="loading">
        </el-switch>
      </div>

      <el-alert
        class="mt-3"
        v-if="message"
        :type="messageType"
        :title="message"
        show-icon>
      </el-alert>
    </el-card>
  </div>
  `,
  data() {
    return {
      isOpen: false,        // true = Model=0；false = Model=1
      loading: false,
      message: '',
      messageType: 'success'
    }
  },
  methods: {
    loadCurrent() {
      this.loading = true;
      axios.get('https://mms.leapoptical.com:5088/api/Center/getPDMode')
        .then(res => {
          // 後端回傳 { model: "0" } 或 "1"
          const m = String(res.data?.model ?? '1');
          this.isOpen = (m === '0');
        })
        .catch(err => {
          this.messageType = 'error';
          this.message = err.response?.data || '讀取目前模式失敗';
        })
        .finally(() => this.loading = false);
    },
    onToggle(val) {
      // val = true -> Model=0；false -> Model=1
      const model = val ? '0' : '1';
      this.loading = true;
      this.message = '';
      axios.post('https://mms.leapoptical.com:5088/api/Center/setPDMode', { model })
        .then(() => {
          this.messageType = 'success';
          this.message = `已更新：Model=${model}`;
        })
        .catch(err => {
          this.messageType = 'error';
          this.message = err.response?.data || '更新失敗';
          // 還原 UI 狀態
          this.isOpen = !val;
        })
        .finally(() => this.loading = false);
    }
  },
  mounted() {
    this.loadCurrent();
  }
});
