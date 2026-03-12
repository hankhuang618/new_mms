


Vue.component('permission-manager-view', {
  template: `
    <div class="container mt-4 permission-manager-page">
      <h4><i class="bi bi-shield-lock text-warning"></i> 權限管理</h4>

      <el-select v-model="selectedUser" placeholder="請選擇帳號" class="mb-3" @change="fetchPermissions" clearable filterable>
        <el-option
          v-for="user in userList"
          :key="user"
          :label="user"
          :value="user"
        />
      </el-select>

      <el-table v-if="permissions.length" :data="permissions" stripe>
        <el-table-column prop="reportName" label="報表名稱"></el-table-column>
        <el-table-column label="是否可查看" width="150">
          <template slot-scope="{ row }">
            <el-switch v-model="row.canAccess" />
          </template>
        </el-table-column>
      </el-table>

      <el-button class="mt-3" type="primary" :disabled="!selectedUser" @click="save">
        儲存權限
      </el-button>
    </div>
  `,
  data() {
    return {
      selectedUser: '',
      userList: [],
      permissions: []
    };
  },
  methods: {
    fetchUsers() {
      axios.get('https://mms.leapoptical.com:5088/api/LOGIN/all')
        .then(res => {
          this.userList = res.data;
        })
        .catch(() => {
          this.$message.error('載入帳號清單失敗');
        });
    },
    fetchPermissions() {
      if (!this.selectedUser) return;
      axios.get(`https://mms.leapoptical.com:5088/api/LOGIN/Permission/${this.selectedUser}`)
        .then(res => {
          this.permissions = res.data.map(p => ({
            reportName: p.reportName,
            canAccess: p.canAccess
          }));
        })
        .catch(() => {
          this.$message.error('載入權限失敗');
        });
    },
    save() {
      axios.post('https://mms.leapoptical.com:5088/api/LOGIN/update', {
        username: this.selectedUser,
        permissions: this.permissions
      })
      .then(() => {
        this.$message.success('權限已更新');
      })
      .catch(() => {
        this.$message.error('儲存失敗');
      });
    }
  },
  mounted() {
    this.fetchUsers();
  }
});
