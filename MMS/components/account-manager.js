Vue.component('account-manager-view', {
  template: `
    <div class="container mt-4 account-manager-page">
      <h4><i class="bi bi-person-lines-fill text-primary"></i> 帳號管理</h4>

      <el-table :data="users" stripe class="mb-4">
        <el-table-column prop="username" label="帳號" width="150" />
        <el-table-column prop="role" label="角色" width="120" />
        <el-table-column prop="isActive" label="啟用" width="100">
          <template slot-scope="{ row }">
            <el-tag :type="row.isActive ? 'success' : 'danger'">
              {{ row.isActive ? '啟用中' : '停用' }}
            </el-tag>
          </template>
        </el-table-column>
      </el-table>

      <el-form :inline="true" :model="form" label-width="100px" class="mb-3">
        <el-form-item label="帳號">
          <el-input v-model="form.username" placeholder="請輸入帳號" />
        </el-form-item>
        <el-form-item label="密碼">
          <el-input v-model="form.password" placeholder="請輸入密碼" type="password" />
        </el-form-item>
        <el-form-item label="角色">
          <el-select v-model="form.role" placeholder="請選擇">
            <el-option label="admin" value="admin" />
            <el-option label="user" value="user" />
          </el-select>
        </el-form-item>
        <el-form-item>
          <el-button type="primary" @click="create">新增帳號</el-button>
          <el-button type="warning" @click="updatePassword">修改密碼</el-button>
        </el-form-item>
      </el-form>
    </div>
  `,
  data() {
    return {
      users: [],
      form: {
        username: '',
        password: '',
        role: 'user'
      }
    };
  },
  methods: {
    loadUsers() {
      axios.get('/api/Account/all')
        .then(res => this.users = res.data)
        .catch(() => this.$message.error('載入帳號失敗'));
    },
    create() {
      if (!this.form.username || !this.form.password) {
        this.$message.warning('請輸入帳號與密碼');
        return;
      }

      axios.post('/api/Account/create', this.form)
        .then(() => {
          this.$message.success('帳號已建立');
          this.loadUsers();
        })
        .catch(err => {
          this.$message.error(err.response?.data || '建立失敗');
        });
    },
    updatePassword() {
      if (!this.form.username || !this.form.password) {
        this.$message.warning('請輸入帳號與密碼');
        return;
      }

      axios.post('/api/Account/update-password', this.form)
        .then(() => this.$message.success('密碼已更新'))
        .catch(() => this.$message.error('修改失敗'));
    }
  },
  mounted() {
    this.loadUsers();
  }
});
