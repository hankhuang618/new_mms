/* MMS - 品保部 QC 人員管理（IQC/IPQC/FQC/OQC）
 * 更新：
 * - 選擇【出勤】或【假別】後，該列功能列鎖定不可再按，直到關閉彈窗或送出完成
 * - 「出勤統計」與「假別統計」並排顯示（小螢幕自動換行）
 */
(function () {
  const API_BASE = window.API_BASE || 'https://mms.leapoptical.com:7238';

  // ===== 假別設定（沿用人員管理） =====
  const LEAVE_CONFIG = {
    '年休': { needHours: true,  needNote: true },
    '請假': { needHours: false, forceHours: '0', needNote: true },
    '哺乳假': { needHours: true,  needNote: true },
    '工傷': { needHours: true,  needNote: true },
    '工傷陪護': { needHours: true,  needNote: true },
    '借出': { needHours: true,  needNote: true, notePreset: ['廠內重工工時','其它廠重工工時','其他借出工時'] },
    '夜班': { needHours: false, forceHours: '0', needNote: false },
    '備註': { needHours: false, forceHours: '0', needNote: true },
    '新進員工': { needHours: false, forceHours: '0', needNote: false },
  };
  const LEAVE_COLOR = {
    '年休':'success','新進員工':'success',
    '請假':'warning','哺乳假':'warning',
    '工傷':'danger','工傷陪護':'danger',
    '借出':'info','備註':'info',
    '夜班':''
  };

  // ===== 本地 i18n 補字典 =====
  const localDict = {
    zh: {
      qc: {
        title: 'QC 人員管理', crumb: '品保部 / MMS',
        filter: { area: '廠區', qcType: 'QC 類別', deptCode: '部門代碼', select: '選擇' },
        options: { iqc:'IQC', ipqc:'IPQC', fqc:'FQC', oqc:'OQC', tech:'技術員', material:'物料員' },
        stats: { attendCount: '出勤人數', attendHours:'出勤工時', leaveSummary:'假別統計' },
        actions: { add:'新增' },
        form: { qcType:'QC 類別', jobNo:'工號', name:'姓名', workDept:'工作拉', submit:'完成', close:'關閉' },
        table: { name:'姓名', jobNo:'工號', qcType:'QC 類別', workDept:'工作拉', attend:'出勤', leave:'假別', effHours:'有效工時', note:'備註', ops:'功能' },
        status: { present:'出勤' },
        btn: { save:'儲存', edit:'編輯', attend:'出勤', leave:'假別', delete:'刪除', cancel:'取消', ok:'確定' },
        hint: { tip:`提示：出勤欄顯示 status=出勤 與今日（{date}）；上方統計依目前清單即時計算。` },
        modal: { title:'設定出勤', nameId:'姓名 / 工號', effHours:'有效工時', note:'備註', hhPlaceholder:'輸入數字（小時）', notePh:'輸入備註（可留空）' },
        leaveDlg: { title:'設定假別', type:'假別', hours:'有效工時', preset:'常用理由', reason:'理由', nameId:'姓名 / 工號' },
        confirm: { cancelAttend:'要取消「{name}」目前的出勤資料？', deleteUser:'確定刪除「{name} / {job}」這筆人員？', cancelLeave:'要取消「{name}」目前的【{leave}】？' },
        msg: {
          required:'必填', addOk:'新增完成', addFail:'新增失敗',
          saveOk:'已儲存', saveFail:'更新失敗',
          attendOk:'出勤已更新', attendFail:'送出失敗',
          cancelAttendOk:'已取消出勤', cancelAttendFail:'取消失敗',
          leaveOk:'假別已更新', leaveFail:'送出失敗',
          cancelLeaveOk:'已取消假別', cancelLeaveFail:'取消失敗',
          deleteOk:'已刪除', deleteFail:'刪除失敗',
          needHours:'請輸入有效工時'
        }
      }
    },
    en: {
      qc: {
        title:'QC Staff Manager', crumb:'Quality / MMS',
        filter:{ area:'Site', qcType:'QC Type', deptCode:'Dept code', select:'Select' },
        options:{ iqc:'IQC', ipqc:'IPQC', fqc:'FQC', oqc:'OQC', tech:'Technician', material:'Material clerk' },
        stats:{ attendCount:'Present', attendHours:'Attend Hours', leaveSummary:'Leave Summary' },
        actions:{ add:'Add' },
        form:{ qcType:'QC Type', jobNo:'Emp ID', name:'Name', workDept:'Work Section', submit:'Submit', close:'Close' },
        table:{ name:'Name', jobNo:'Emp ID', qcType:'QC Type', workDept:'Work Section', attend:'Attendance', leave:'Leave', effHours:'Eff. Hours', note:'Note', ops:'Actions' },
        status:{ present:'Present' },
        btn:{ save:'Save', edit:'Edit', attend:'Attend', leave:'Leave', delete:'Delete', cancel:'Cancel', ok:'OK' },
        hint:{ tip:`Hint: Attendance shows status=Present & today ({date}); stats are live.` },
        modal:{ title:'Set Attendance', nameId:'Name / ID', effHours:'Effective Hours', note:'Note', hhPlaceholder:'Enter hours', notePh:'Enter note (optional)' },
        leaveDlg:{ title:'Set Leave', type:'Leave Type', hours:'Effective Hours', preset:'Presets', reason:'Reason', nameId:'Name / ID' },
        confirm:{ cancelAttend:'Cancel current attendance for “{name}”?', deleteUser:'Delete user “{name} / {job}”?', cancelLeave:'Cancel current “{leave}” for “{name}”?' },
        msg:{
          required:'Required', addOk:'Created', addFail:'Create failed',
          saveOk:'Saved', saveFail:'Update failed',
          attendOk:'Attendance updated', attendFail:'Submit failed',
          cancelAttendOk:'Attendance canceled', cancelAttendFail:'Cancel failed',
          leaveOk:'Leave updated', leaveFail:'Submit failed',
          cancelLeaveOk:'Leave canceled', cancelLeaveFail:'Cancel failed',
          deleteOk:'Deleted', deleteFail:'Delete failed',
          needHours:'Please input valid hours'
        }
      }
    },
    vi: {
      qc:{
        title:'Quản lý nhân sự QC', crumb:'Phẩm chất / MMS',
        filter:{ area:'Khu vực', qcType:'Loại QC', deptCode:'Mã bộ phận', select:'Chọn' },
        options:{ iqc:'IQC', ipqc:'IPQC', fqc:'FQC', oqc:'OQC', tech:'Kỹ thuật viên', material:'NV vật tư' },
        stats:{ attendCount:'Số người đi làm', attendHours:'Giờ chấm công', leaveSummary:'Thống kê nghỉ phép' },
        actions:{ add:'Thêm' },
        form:{ qcType:'Loại QC', jobNo:'Mã NV', name:'Họ tên', workDept:'Chuyền', submit:'Hoàn tất', close:'Đóng' },
        table:{ name:'Họ tên', jobNo:'Mã NV', qcType:'Loại QC', workDept:'Chuyền', attend:'Đi làm', leave:'Nghỉ phép', effHours:'Giờ hiệu dụng', note:'Ghi chú', ops:'Chức năng' },
        status:{ present:'Đi làm' },
        btn:{ save:'Lưu', edit:'Sửa', attend:'Đi làm', leave:'Nghỉ phép', delete:'Xóa', cancel:'Hủy', ok:'Đồng ý' },
        hint:{ tip:`Gợi ý: Cột đi làm = trạng thái=Đi làm & hôm nay ({date}); thống kê cập nhật tức thì.` },
        modal:{ title:'Cài đặt đi làm', nameId:'Họ tên / Mã NV', effHours:'Giờ hiệu dụng', note:'Ghi chú', hhPlaceholder:'Nhập số giờ', notePh:'Nhập ghi chú (tùy chọn)' },
        leaveDlg:{ title:'Cài đặt nghỉ phép', type:'Loại nghỉ', hours:'Giờ hiệu dụng', preset:'Lý do thường dùng', reason:'Lý do', nameId:'Họ tên / Mã NV' },
        confirm:{ cancelAttend:'Hủy thông tin đi làm của “{name}”?', deleteUser:'Xóa “{name} / {job}”？', cancelLeave:'Hủy “{leave}” của “{name}”？' },
        msg:{
          required:'Bắt buộc', addOk:'Đã thêm', addFail:'Thêm thất bại',
          saveOk:'Đã lưu', saveFail:'Cập nhật thất bại',
          attendOk:'Đã cập nhật đi làm', attendFail:'Gửi thất bại',
          cancelAttendOk:'Đã hủy đi làm', cancelAttendFail:'Hủy thất bại',
          leaveOk:'Đã cập nhật nghỉ phép', leaveFail:'Gửi thất bại',
          cancelLeaveOk:'Đã hủy nghỉ phép', cancelLeaveFail:'Hủy thất bại',
          deleteOk:'Đã xóa', deleteFail:'Xóa thất bại',
          needHours:'Vui lòng nhập số giờ hợp lệ'
        }
      }
    }
  };
  function l(key, params){
    if (window.I18N && window.I18N.t) {
      const v = window.I18N.t(key, params);
      if (v && v !== key) return v;
    }
    const lang = (window.I18N && window.I18N.getLang && window.I18N.getLang()) || 'zh';
    const segs = String(key).split('.');
    let node = localDict[lang];
    for (const s of segs) node = node && node[s];
    let out = (node == null) ? key : String(node);
    if (params && typeof params === 'object') {
      for (const k in params) out = out.replace(new RegExp('\\{' + k + '\\}', 'g'), params[k]);
    }
    return out;
  }

  const QC_LIST = [
    { value:'IQC', key:'iqc' },
    { value:'IPQC', key:'ipqc' },
    { value:'FQC', key:'fqc' },
    { value:'OQC', key:'oqc' },
    { value:'技術員', key:'tech' },
    { value:'物料員', key:'material' }
  ];

  // Style（僅注入一次）
  if (!document.getElementById('qc-staff-attend-modal-style')) {
    const style = document.createElement('style');
    style.id = 'qc-staff-attend-modal-style';
    style.textContent = `
      .qs-modal-mask { position: fixed; inset: 0; background: rgba(0,0,0,0.45); display:flex; align-items:center; justify-content:center; z-index: 4000; }
      .qs-modal { width: 520px; max-width: 95vw; background: #fff; border-radius: 14px; box-shadow: 0 20px 50px rgba(0,0,0,0.25); overflow:hidden; }
      .qs-modal-header { padding: 12px 16px; font-weight: 600; border-bottom: 1px solid #eef0f3; }
      .qs-modal-body { padding: 16px; max-height: 70vh; overflow:auto; }
      .qs-modal-footer { padding: 12px 16px; border-top: 1px solid #eef0f3; text-align: right; }
      .qs-field { margin-bottom: 12px; }
      .qs-label { display:block; font-size: 13px; color:#6b7280; margin-bottom:6px; }
      .qs-input, .qs-select, .qs-textarea { width: 100%; border:1px solid #d1d5db; border-radius: 8px; padding: 8px 10px; }
      .qs-input { height: 36px; }
      .qs-textarea { min-height: 80px; resize: vertical; }
      .qs-actions { display:flex; gap:8px; justify-content:flex-end; }
      .qs-btn { display:inline-flex; align-items:center; gap:6px; padding: 6px 12px; border-radius: 8px; border:1px solid transparent; cursor:pointer; }
      .qs-btn-default { background:#fff; border-color:#d1d5db; }
      .qs-btn-primary { background:#2563eb; color:#fff; }

      .attendance-inline{ display:inline-flex; align-items:center; gap:8px; flex-wrap:wrap; }
      .attendance-inline .el-button--text{ padding:0 4px; }

      .stat-chip{ display:flex; align-items:center; gap:10px; padding:8px 12px; border:1px solid #e5e7eb; border-radius:12px; background:#fff; }
      .stat-chip .k{ font-size:12px; color:#6b7280; }
      .stat-chip .v{ font-size:18px; font-weight:700; line-height:1; }

      .el-table-scroll { overflow-x: auto; }
      .el-table-scroll > .el-table { min-width: 1120px; }
      @media (max-width: 768px){
        .qc-staff-manager-page .sticky-top { top: 56px !important; }
        .qc-staff-manager-page .el-card { border-radius: 12px; }
        .qc-staff-manager-page .el-table .cell{ font-size:12px; padding: 4px 6px; }
        .actions-wrap { display:flex; flex-wrap:wrap; gap:6px; justify-content:center; }
        .actions-wrap .el-button { padding-left: 8px; padding-right: 8px; }
      }
    `;
    document.head.appendChild(style);
  }

  Vue.component('qc-staff-manager-view', {
    template: `
      <div class="container-fluid py-3 qc-staff-manager-page">
        <div class="d-flex align-items-center justify-content-between mb-2">
          <h4 class="mb-0">{{ l('qc.title') }}</h4>
          <div class="text-muted small">{{ l('qc.crumb') }}</div>
        </div>

        <!-- 篩選 + 統計 -->
        <el-card shadow="never" class="mb-3 sticky-top" style="top:64px;">
          <div class="row g-2 align-items-end flex-wrap qc-staff-filter">
            <div class="col-6 col-md-auto">
              <label class="form-label mb-1">{{ l('qc.filter.area') }}</label>
              <el-select v-model="selectedArea" :placeholder="l('qc.filter.select')" style="width:100%" @change="onFilterChange">
                <el-option label="ZH" value="ZH"></el-option>
                <el-option label="TC" value="TC"></el-option>
                <el-option label="VN" value="VN"></el-option>
                <el-option label="TW" value="TW"></el-option>
              </el-select>
            </div>

            <div class="col-6 col-md-auto">
              <label class="form-label mb-1">{{ l('qc.filter.qcType') }}</label>
              <el-select v-model="selectedQC" :placeholder="l('qc.filter.select')" style="width:100%" @change="onFilterChange">
                <el-option v-for="q in qcOptionsLocalized" :key="q.value" :label="q.label" :value="q.value"></el-option>
              </el-select>
            </div>

            <div class="col-6 col-md-auto">
              <div class="form-label mb-1 text-muted">{{ l('qc.filter.deptCode') }}</div>
              <div class="fw-bold">{{ selectedQCLabel }}</div>
            </div>

            <div class="col-12">
              <!-- 並排：出勤統計 + 假別統計 -->
              <div class="d-flex flex-wrap align-items-center gap-2">
                <div class="stat-chip">
                  <div class="k">{{ l('qc.stats.attendCount') }}</div>
                  <div class="v">{{ attendCount }}</div>
                </div>
                <div class="stat-chip">
                  <div class="k">{{ l('qc.stats.attendHours') }}</div>
                  <div class="v">{{ attendHoursDisplay }}</div>
                </div>

                <div class="text-muted small ms-2">{{ l('qc.stats.leaveSummary') }}</div>
                <div class="d-flex flex-wrap gap-1">
                  <el-tag v-for="item in leaveStatsArray" :key="item.type" :type="item.tagType" effect="plain">
                    {{ item.type }}：{{ item.count }}
                  </el-tag>
                  <span v-if="leaveStatsArray.length===0" class="text-muted small">—</span>
                </div>

                <div class="ms-auto">
                  <el-button type="primary" size="small" @click="new_staff">{{ l('qc.actions.add') }}</el-button>
                </div>
              </div>
            </div>

            <!-- 新增表單（含工作拉） -->
            <div v-if="showForm" class="col-12 mt-2">
              <el-form :model="newStaff" :rules="formRules" ref="newForm" label-width="90px">
                <div class="row g-2">
                  <div class="col-12 col-md-3">
                    <el-form-item :label="l('qc.form.qcType')" prop="department">
                      <el-input v-model="newStaff.department" disabled></el-input>
                    </el-form-item>
                  </div>
                  <div class="col-12 col-md-3">
                    <el-form-item :label="l('qc.form.jobNo')" prop="jobnumber">
                      <el-input v-model="newStaff.jobnumber"></el-input>
                    </el-form-item>
                  </div>
                  <div class="col-12 col-md-3">
                    <el-form-item :label="l('qc.form.name')" prop="username">
                      <el-input v-model="newStaff.username"></el-input>
                    </el-form-item>
                  </div>
                  <div class="col-12 col-md-3">
                    <el-form-item :label="l('qc.form.workDept')">
                      <el-input v-model="newStaff.work_department" placeholder="2D1SS1C"></el-input>
                    </el-form-item>
                  </div>
                </div>
                <el-button type="success" @click="submitNew">{{ l('qc.form.submit') }}</el-button>
                <el-button @click="showForm=false">{{ l('qc.form.close') }}</el-button>
              </el-form>
            </div>
          </div>
        </el-card>

        <!-- 資料表 -->
        <el-card shadow="never">
          <div class="el-table-scroll">
            <el-table :data="staffData" border height="60vh" v-loading="loading">
              <el-table-column type="index" label="#" width="50"></el-table-column>

              <el-table-column :label="l('qc.table.name')" min-width="140">
                <template slot-scope="{row}">
                  <template v-if="row._editable">
                    <el-input v-model="row.username"></el-input>
                  </template>
                  <template v-else>{{ row.username }}</template>
                </template>
              </el-table-column>

              <el-table-column :label="l('qc.table.jobNo')" width="120">
                <template slot-scope="{row}">
                  <template v-if="row._editable">
                    <el-input v-model="row.jobnumber"></el-input>
                  </template>
                  <template v-else>{{ row.jobnumber }}</template>
                </template>
              </el-table-column>

              <el-table-column :label="l('qc.table.qcType')" width="110">
                <template slot-scope="{row}">
                  <template v-if="row._editable">
                    <el-select v-model="row.department" style="width:100px">
                      <el-option v-for="q in qcOptionsLocalized" :key="q.value" :label="q.label" :value="q.value"></el-option>
                    </el-select>
                  </template>
                  <template v-else>{{ mapQCLabel(row.department) }}</template>
                </template>
              </el-table-column>

              <el-table-column :label="l('qc.table.workDept')" min-width="140">
                <template slot-scope="{row}">
                  <template v-if="row._editable">
                    <el-input v-model="row.work_department" placeholder="2D1SS1C"></el-input>
                  </template>
                  <template v-else>
                    <span v-if="row.work_department && String(row.work_department).trim() !== ''">{{ row.work_department }}</span>
                    <span v-else class="text-muted">—</span>
                  </template>
                </template>
              </el-table-column>

              <!-- 出勤 -->
              <el-table-column :label="l('qc.table.attend')" min-width="200" align="center">
                <template slot-scope="{row}">
                  <template v-if="isAttend(row)">
                    <div class="attendance-inline">
                      <el-tag type="success" size="small">{{ row.status || l('qc.status.present') }}</el-tag>
                      <span class="text-muted small">{{ todayStr }}</span>
                      <el-button type="text" @click="askCancelAttend(row)">{{ l('qc.btn.cancel') }}</el-button>
                    </div>
                  </template>
                  <template v-else>
                    <span class="text-muted">—</span>
                  </template>
                </template>
              </el-table-column>

              <!-- 假別（出勤以外的狀態） -->
              <el-table-column :label="l('qc.table.leave')" min-width="180" align="center">
                <template slot-scope="{row}">
                  <template v-if="row.status && !isAttend(row)">
                    <el-tag :type="leaveTagType(row.status)">{{ row.status }}</el-tag>
                    <el-button type="text" class="ms-1" @click="askCancelLeave(row)">{{ l('qc.btn.cancel') }}</el-button>
                  </template>
                  <template v-else>
                    <span class="text-muted">—</span>
                  </template>
                </template>
              </el-table-column>

              <el-table-column :label="l('qc.table.effHours')" width="90" align="right">
                <template slot-scope="{row}">
                  <span v-if="Number(row.hr_time)">{{ Number(row.hr_time) }}</span>
                  <span v-else class="text-muted">—</span>
                </template>
              </el-table-column>

              <el-table-column :label="l('qc.table.note')" min-width="160">
                <template slot-scope="{row}">
                  <span v-if="row.note && String(row.note).trim() !== ''">{{ row.note }}</span>
                  <span v-else class="text-muted">—</span>
                </template>
              </el-table-column>

              <el-table-column :label="l('qc.table.ops')" width="320" fixed="right" align="center">
                <template slot-scope="{row}">
                  <div class="actions-wrap">
                    <el-button v-if="row._editable" type="primary" size="mini" @click="updateRow(row)" :disabled="isLocked(row)">{{ l('qc.btn.save') }}</el-button>
                    <el-button v-else type="default" size="mini" @click="row._editable = true" :disabled="isLocked(row)">{{ l('qc.btn.edit') }}</el-button>

                    <el-button size="mini" type="success" plain @click="openAttendModal(row)" :disabled="isLocked(row)">{{ l('qc.btn.attend') }}</el-button>
                    <el-button size="mini" type="warning" plain @click="openLeaveDialog(row)" :disabled="isLocked(row)">{{ l('qc.btn.leave') }}</el-button>
                    <el-button size="mini" type="danger" plain @click="askDelete(row)" :disabled="isLocked(row)">{{ l('qc.btn.delete') }}</el-button>
                  </div>
                </template>
              </el-table-column>
            </el-table>
          </div>

          <div class="text-muted small mt-2 d-flex align-items-center">
            <span class="me-3">{{ l('qc.hint.tip', { date: todayStr }) }}</span>
          </div>
        </el-card>

        <!-- 出勤彈窗（自製） -->
        <div v-if="attendModal.visible" class="qs-modal-mask" @click.self="closeAttendModal">
          <div class="qs-modal" role="dialog" aria-modal="true">
            <div class="qs-modal-header">{{ l('qc.modal.title') }}</div>
            <div class="qs-modal-body">
              <div class="qs-field">
                <div class="qs-label">{{ l('qc.modal.nameId') }}</div>
                <div>{{ attendModal.form.username }} / {{ attendModal.form.jobnumber }}</div>
              </div>

              <div class="qs-field">
                <div class="qs-label">{{ l('qc.modal.effHours') }}</div>
                <input class="qs-input" v-model="attendModal.form.hr_time" @input="onlyNumber('hr_time')" :placeholder="l('qc.modal.hhPlaceholder')" />
              </div>

              <div class="qs-field">
                <div class="qs-label">{{ l('qc.modal.note') }}</div>
                <textarea class="qs-textarea" v-model="attendModal.form.remark" :placeholder="l('qc.modal.notePh')"></textarea>
              </div>
            </div>
            <div class="qs-modal-footer">
              <div class="qs-actions">
                <button class="qs-btn qs-btn-default" @click="closeAttendModal" :disabled="attendModal.sending">{{ l('qc.btn.cancel') }}</button>
                <button class="qs-btn qs-btn-primary" @click="submitAttend" :disabled="attendModal.sending">{{ l('qc.btn.ok') }}</button>
              </div>
            </div>
          </div>
        </div>

        <!-- 假別彈窗（Element UI） -->
        <el-dialog
          :title="l('qc.leaveDlg.title')"
          :visible.sync="leaveDialog.visible"
          width="520px"
          :close-on-click-modal="false"
          :before-close="onLeaveDialogClose">
          <el-form :model="leaveDialog.form" :rules="leaveRules" ref="leaveForm" label-width="96px">
            <el-form-item :label="l('qc.leaveDlg.nameId')">
              <div>{{ leaveDialog.form.username }} / {{ leaveDialog.form.jobnumber }}</div>
            </el-form-item>

            <el-form-item :label="l('qc.leaveDlg.type')" prop="leaveType">
              <el-select v-model="leaveDialog.form.leaveType" :placeholder="l('qc.filter.select')" style="width:260px">
                <el-option v-for="(cfg, key) in LEAVE_CONFIG" :key="key" :label="key" :value="key"></el-option>
              </el-select>
            </el-form-item>

            <el-form-item v-if="currentLeave.needHours" :label="l('qc.leaveDlg.hours')" prop="hr_time">
              <el-input v-model="leaveDialog.form.hr_time" :placeholder="l('qc.modal.hhPlaceholder')" @input="numberOnly('hr_time')"></el-input>
            </el-form-item>

            <el-form-item v-if="currentLeave.notePreset" :label="l('qc.leaveDlg.preset')">
              <el-select v-model="leaveDialog.form.notePreset" :placeholder="l('qc.filter.select')" style="width:260px">
                <el-option v-for="item in currentLeave.notePreset" :key="item" :label="item" :value="item"></el-option>
              </el-select>
            </el-form-item>

            <el-form-item v-if="currentLeave.needNote || leaveDialog.form.notePreset" :label="l('qc.leaveDlg.reason')" prop="note">
              <el-input v-model="leaveDialog.form.note" :placeholder="l('qc.modal.notePh')"></el-input>
            </el-form-item>
          </el-form>

          <span slot="footer" class="dialog-footer">
            <el-button @click="onLeaveDialogClose()" :disabled="leaveDialog.sending">{{ l('qc.btn.cancel') }}</el-button>
            <el-button type="primary" @click="submitLeave" :loading="leaveDialog.sending">{{ l('qc.btn.ok') }}</el-button>
          </span>
        </el-dialog>

        <!-- 刪除/取消出勤/取消假別 -->
        <el-dialog :title="confirmBox.title" :visible.sync="confirmBox.visible" width="420px" :close-on-click-modal="false">
          <div class="mb-2">{{ confirmBox.message }}</div>
          <span slot="footer" class="dialog-footer">
            <el-button @click="onConfirmCancel">{{ l('qc.btn.cancel') }}</el-button>
            <el-button type="primary" @click="onConfirmOk">{{ l('qc.btn.ok') }}</el-button>
          </span>
        </el-dialog>
      </div>
    `,

    data() {
      return {
        loading: false,
        selectedArea: 'VN',
        selectedQC: 'IQC',

        staffData: [],
        showForm: false,
        newStaff: {
          department: '',
          jobnumber: '',
          username: '',
          work_department: ''
        },
        formRules: {
          department: [{ required: true, message: l('qc.msg.required'), trigger: 'blur' }],
          jobnumber:  [{ required: true, message: l('qc.msg.required'), trigger: 'blur' }],
          username:   [{ required: true, message: l('qc.msg.required'), trigger: 'blur' }]
        },

        attendModal: {
          visible: false,
          sending: false,
          row: null,
          form: { jobnumber:'', username:'', hr_time:'', remark:'' }
        },

        // 假別彈窗
        leaveDialog: {
          visible: false,
          sending: false,
          row: null,
          form: { id:null, username:'', jobnumber:'', leaveType:'', hr_time:'', notePreset:'', note:'' }
        },

        confirmBox: {
          visible: false,
          title: l('qc.btn.ok'),
          message: '',
          _resolve: null,
          _reject: null,
          payload: null
        },

        curLang: (window.I18N && window.I18N.getLang && window.I18N.getLang()) || 'zh',
        LEAVE_CONFIG
      };
    },

    computed: {
      qcOptionsLocalized(){ return QC_LIST.map(o => ({ value:o.value, label:l('qc.options.'+o.key) })); },
      selectedQCLabel(){ const f = QC_LIST.find(x => x.value === this.selectedQC); return f ? l('qc.options.'+f.key) : this.selectedQC; },
      todayStr() {
        const d = new Date(); const y = d.getFullYear(); const m = String(d.getMonth()+1).padStart(2,'0'); const dd = String(d.getDate()).padStart(2,'0');
        return `${y}/${m}/${dd}`;
      },
      attendCount() {
        return this.staffData.reduce((acc, r) => acc + (this.isAttend(r) ? 1 : 0), 0);
      },
      attendHours() {
        return this.staffData.reduce((sum, r) => {
          if (!this.isAttend(r)) return sum;
          const v = Number(r.hr_time);
          return sum + (isNaN(v) ? 0 : v);
        }, 0);
      },
      attendHoursDisplay() {
        const v = this.attendHours;
        return Number.isInteger(v) ? String(v) : v.toFixed(1);
      },
      // 假別統計（排除出勤）
      leaveStats(){
        const map = {};
        this.staffData.forEach(r => {
          if (r.status && !this.isAttend(r)) {
            map[r.status] = (map[r.status] || 0) + 1;
          }
        });
        return map;
      },
      leaveStatsArray(){
        return Object.keys(this.leaveStats).map(k => ({
          type: k,
          count: this.leaveStats[k],
          tagType: this.leaveTagType(k)
        }));
      },
      // 假別表單規則
      currentLeave(){
        return this.LEAVE_CONFIG[this.leaveDialog.form.leaveType] || {};
      },
      leaveRules(){
        const r = { leaveType: [{ required: true, message: l('qc.msg.required'), trigger:'change' }] };
        const cfg = this.currentLeave;
        if (cfg.needHours) r.hr_time = [{ required: true, message: l('qc.msg.needHours'), trigger:'blur' }];
        if (cfg.needNote)  r.note    = [{ required: true, message: l('qc.msg.required'),  trigger:'blur' }];
        return r;
      }
    },

    watch: {
      'leaveDialog.form.leaveType'(val){
        const cfg = this.currentLeave;
        if (cfg.forceHours != null) this.leaveDialog.form.hr_time = cfg.forceHours;
        if (!cfg.needHours && !cfg.forceHours) this.leaveDialog.form.hr_time = '';
        if (!cfg.needNote) this.leaveDialog.form.note = '';
        this.leaveDialog.form.notePreset = '';
      }
    },

    created() {
      this.newStaff.department = this.selectedQC;
      this.fetchStaff();

      window.addEventListener('mms-lang-changed', (e) => {
        this.curLang = e.detail?.lang || this.curLang;
        this.formRules = {
          department: [{ required: true, message: l('qc.msg.required'), trigger: 'blur' }],
          jobnumber:  [{ required: true, message: l('qc.msg.required'), trigger: 'blur' }],
          username:   [{ required: true, message: l('qc.msg.required'), trigger: 'blur' }]
        };
        this.confirmBox.title = l('qc.btn.ok');
        this.$forceUpdate();
      });
    },

    methods: {
      l,

      mapQCLabel(dep){ const f = QC_LIST.find(x => x.value === dep); return f ? l('qc.options.'+f.key) : dep || ''; },

      onFilterChange() {
        this.newStaff.department = this.selectedQC;
        this.fetchStaff();
      },

      isAttend(row) {
        return String(row.status || '').trim() === '出勤' || String(row.status || '').trim() === l('qc.status.present');
      },

      isLocked(row){
        return !!row._lockedAction;
      },

      // 取得清單
      fetchStaff() {
        if (!this.selectedArea || !this.selectedQC) return;
        this.loading = true;
        axios.post(`${API_BASE}/api/Attendance`, {
          area: this.selectedArea,
          department: this.selectedQC
        })
        .then(res => {
          const list = Array.isArray(res.data) ? res.data : [];
          // 期望欄位：id, area, department, jobnumber, username, status(出勤/假別), hr_time, note, work_department
          this.staffData = list.map(x => Object.assign({ _editable:false, _lockedAction:null, work_department:'' }, x, {
            work_department: x.work_department != null ? x.work_department : ''
          }));
        })
        .catch(err => {
          console.error('fetchStaff error', err);
          this.staffData = [];
        })
        .finally(() => this.loading = false);
      },

      // 新增
      new_staff() {
        this.newStaff = { department: this.selectedQC, jobnumber: '', username: '', work_department: '' };
        this.showForm = true;
        this.$nextTick(() => this.$refs.newForm && this.$refs.newForm.clearValidate());
      },
      submitNew() {
        this.$refs.newForm.validate(valid => {
          if (!valid) return;
          axios.post(`${API_BASE}/api/Attendance/creat`, {
            area: this.selectedArea,
            department: this.newStaff.department,
            jobnumber: this.newStaff.jobnumber,
            username: this.newStaff.username,
            work_department: this.newStaff.work_department || ''
          })
          .then(() => {
            this.$message.success(l('qc.msg.addOk'));
            this.showForm = false;
            this.fetchStaff();
          })
          .catch(err => {
            console.error('addNewStaff error', err);
            this.$message.error(l('qc.msg.addFail'));
          });
        });
      },

      // 單列更新
      updateRow(row) {
        const payload = {
          id: String(row.id),
          area: this.selectedArea,
          department: row.department,
          jobnumber: row.jobnumber,
          username: row.username,
          work_department: row.work_department,
          online: row.online
        };
        row._lockedAction = row._lockedAction || 'save';
        axios.put(`${API_BASE}/api/Attendance/UPDATE_USER`, payload, {
          headers: { 'Content-Type':'application/json' }
        })
        .then(() => {
          this.$message.success(l('qc.msg.saveOk'));
          row._editable = false;
          this.fetchStaff();
        })
        .catch(err => {
          console.error('UPDATE_staffData error', err);
          this.$message.error(l('qc.msg.saveFail'));
          row._lockedAction = null;
        });
      },

      // ===== 出勤：自製彈窗（鎖列） =====
      openAttendModal(row){
        if (this.isLocked(row)) return;
        row._lockedAction = 'attend';
        this.attendModal.row = row;
        this.attendModal.form = {
          jobnumber: row.jobnumber,
          username:  row.username,
          hr_time:   row.hr_time ? String(row.hr_time) : '',
          remark:    ''  // remark -> 後端 note
        };
        this.attendModal.visible = true;
      },
      closeAttendModal(){
        if (this.attendModal.row) this.attendModal.row._lockedAction = null;
        this.attendModal.visible = false;
        this.attendModal.row = null;
      },
      onlyNumber(field){
        const v = this.attendModal.form[field];
        this.attendModal.form[field] = String(v).replace(/[^\d.]/g, '');
      },
      submitAttend(){
        const f = this.attendModal.form;
        if (!f.hr_time || isNaN(Number(f.hr_time))) {
          this.$message.warning(l('qc.msg.needHours')); return;
        }
        this.attendModal.sending = true;
        axios.post(`${API_BASE}/api/Attendance/UPDATESTATUS`, {
          area: this.selectedArea,
          jobnumber: f.jobnumber,
          ststus: '出勤',
          hr_time: f.hr_time,
          note: f.remark || ''
        })
        .then(() => {
          this.$message.success(l('qc.msg.attendOk'));
          this.attendModal.visible = false;
          this.fetchStaff();
        })
        .catch(err => {
          console.error('submitAttend error', err);
          this.$message.error(l('qc.msg.attendFail'));
          if (this.attendModal.row) this.attendModal.row._lockedAction = null;
        })
        .finally(() => {
          this.attendModal.sending = false;
          this.attendModal.row = null;
        });
      },
      askCancelAttend(row){
        this.confirm(l('qc.confirm.cancelAttend', { name: row.username }), l('qc.btn.attend'), { row, type:'attend' })
          .then(({ row }) => this.doCancelAttend(row))
          .catch(() => {});
      },
      doCancelAttend(row){
        const id = parseInt(row.id, 10);
        axios.delete(`${API_BASE}/api/Attendance/Delete_hr`, {
          data: { id, area: this.selectedArea },
          headers: { 'Content-Type': 'application/json' }
        })
        .then(() => {
          this.$message.success(l('qc.msg.cancelAttendOk'));
          this.fetchStaff();
        })
        .catch(err => {
          console.error('cancel_attend error', err);
          this.$message.error(l('qc.msg.cancelAttendFail'));
        });
      },

      // ===== 假別：沿用人員管理邏輯（鎖列） =====
      openLeaveDialog(row){
        if (this.isLocked(row)) return;
        row._lockedAction = 'leave';
        this.leaveDialog.row = row;
        this.leaveDialog.form = {
          id: row.id,
          username: row.username,
          jobnumber: row.jobnumber,
          leaveType: (row.status && !this.isAttend(row)) ? row.status : '',
          hr_time: (row.status && !this.isAttend(row)) ? (row.hr_time || '') : '',
          notePreset:'',
          note:''
        };
        this.leaveDialog.visible = true;
        this.$nextTick(() => this.$refs.leaveForm && this.$refs.leaveForm.clearValidate());
      },
      numberOnly(field){
        const v = this.leaveDialog.form[field];
        this.leaveDialog.form[field] = String(v).replace(/[^\d.]/g, '');
      },
      onLeaveDialogClose(done){
        // 關閉時解除鎖定（無論按 X 或 取消）
        if (this.leaveDialog.row) this.leaveDialog.row._lockedAction = null;
        this.leaveDialog.visible = false;
        this.leaveDialog.row = null;
        if (typeof done === 'function') done();
      },
      submitLeave(){
        this.$refs.leaveForm.validate(valid => {
          if (!valid) return;
          this.leaveDialog.sending = true;
          const cfg = this.currentLeave;
          let note = this.leaveDialog.form.note || '';
          if (cfg.notePreset && this.leaveDialog.form.notePreset) {
            note = this.leaveDialog.form.notePreset + (note ? `；${note}` : '');
          }
          const hr = cfg.forceHours != null ? cfg.forceHours : (cfg.needHours ? (this.leaveDialog.form.hr_time || '0') : '0');

          axios.post(`${API_BASE}/api/Attendance/UPDATESTATUS`, {
            area: this.selectedArea,
            jobnumber: this.leaveDialog.form.jobnumber,
            ststus: this.leaveDialog.form.leaveType,
            hr_time: hr,
            note
          })
          .then(() => {
            this.$message.success(l('qc.msg.leaveOk'));
            this.leaveDialog.visible = false;
            this.fetchStaff();
          })
          .catch(err => {
            console.error('submitLeave error', err);
            this.$message.error(l('qc.msg.leaveFail'));
            if (this.leaveDialog.row) this.leaveDialog.row._lockedAction = null;
          })
          .finally(() => {
            this.leaveDialog.sending = false;
            this.leaveDialog.row = null;
          });
        });
      },
      askCancelLeave(row){
        if (!row.status || this.isAttend(row)) return;
        this.confirm(l('qc.confirm.cancelLeave', { name: row.username, leave: row.status }), l('qc.btn.leave'), { row, type:'leave' })
          .then(({ row }) => this.doCancelLeave(row))
          .catch(() => {});
      },
      doCancelLeave(row){
        const id = parseInt(row.id, 10);
        axios.delete(`${API_BASE}/api/Attendance/Delete_hr`, {
          data: { id, area: this.selectedArea },
          headers: { 'Content-Type': 'application/json' }
        })
        .then(() => {
          this.$message.success(l('qc.msg.cancelLeaveOk'));
          this.fetchStaff();
        })
        .catch(err => {
          console.error('delete_hr error', err);
          this.$message.error(l('qc.msg.cancelLeaveFail'));
        });
      },

      // 刪除資料列
      askDelete(row) {
        this.confirm(l('qc.confirm.deleteUser', { name: row.username, job: row.jobnumber }), l('qc.btn.delete'), { row })
          .then(({ row }) => this.doDelete(row))
          .catch(() => {});
      },
      doDelete(row) {
        const id = parseInt(row.id, 10);
        axios.delete(`${API_BASE}/api/Attendance/Delete`, {
          data: { id, area: this.selectedArea },
          headers: { 'Content-Type': 'application/json' }
        })
        .then(() => {
          this.$message.success(l('qc.msg.deleteOk'));
          this.fetchStaff();
        })
        .catch(err => {
          console.error('delete_staffData error', err);
          this.$message.error(l('qc.msg.deleteFail'));
        });
      },

      // ===== 自製確認框 =====
      confirm(message, title='確認', payload=null) {
        return new Promise((resolve, reject) => {
          this.confirmBox.title = title;
          this.confirmBox.message = message;
          this.confirmBox.payload = payload;
          this.confirmBox.visible = true;
          this.confirmBox._resolve = resolve;
          this.confirmBox._reject  = reject;
        });
      },
      onConfirmOk() {
        this.confirmBox.visible = false;
        const res = this.confirmBox._resolve; const payload = this.confirmBox.payload;
        this.confirmBox._resolve = null; this.confirmBox._reject = null; this.confirmBox.payload = null;
        res && res(payload);
      },
      onConfirmCancel() {
        this.confirmBox.visible = false;
        const rej = this.confirmBox._reject;
        this.confirmBox._resolve = null; this.confirmBox._reject = null; this.confirmBox.payload = null;
        rej && rej('cancel');
      },

      leaveTagType(status){ return LEAVE_COLOR[status] || ''; }
    }
  });
})();
