<script setup lang="ts">
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { fetchTestHistory, type TestRunSummary } from '@/api/test'
import TestResultDialog from '@/components/TestResultDialog.vue'

const router = useRouter()

const runs = ref<TestRunSummary[]>([])
const loading = ref(false)
const filterModule = ref('')
const resultVisible = ref(false)
const selectedRunId = ref<string | null>(null)

async function loadHistory() {
  loading.value = true
  try {
    const data = await fetchTestHistory(filterModule.value || undefined)
    runs.value = data.runs
  } catch (e: any) {
    ElMessage.error('加载历史失败: ' + e.message)
  } finally {
    loading.value = false
  }
}

function viewResult(runId: string) {
  selectedRunId.value = runId
  resultVisible.value = true
}

function formatStatus(status: string): string {
  const map: Record<string, string> = {
    pending: '等待中',
    running: '运行中',
    completed: '已完成',
    error: '异常',
    stopped: '已停止',
  }
  return map[status] || status
}

function statusType(status: string): string {
  const map: Record<string, string> = {
    completed: 'success',
    error: 'danger',
    stopped: 'warning',
    running: '',
  }
  return map[status] || 'info'
}

function formatDuration(s: number | null): string {
  if (s == null) return '-'
  const m = Math.floor(s / 60)
  const sec = Math.round(s % 60)
  return m > 0 ? `${m}分${sec}秒` : `${sec}秒`
}

onMounted(() => {
  loadHistory()
})
</script>

<template>
  <div class="page-container">
    <div class="page-header">
      <div class="header-left">
        <el-button text @click="router.push('/')">
          <el-icon><ArrowLeft /></el-icon>
          返回
        </el-button>
        <h2>测试历史</h2>
      </div>
      <div class="header-right">
        <el-select
          v-model="filterModule"
          placeholder="筛选模块"
          clearable
          size="small"
          style="width: 160px"
          @change="loadHistory"
        >
          <el-option label="Rin_FSV3004" value="Rin_FSV3004" />
          <el-option label="线宽_FSV3004" value="线宽_FSV3004" />
          <el-option label="时域" value="时域" />
          <el-option label="信噪比" value="信噪比" />
          <el-option label="单频" value="单频" />
          <el-option label="功率" value="功率" />
          <el-option label="PZT调制" value="PZT调制" />
          <el-option label="相噪" value="相噪" />
        </el-select>
        <el-button size="small" @click="loadHistory" :loading="loading">
          刷新
        </el-button>
      </div>
    </div>

    <div class="page-body">
      <el-table :data="runs" stripe size="small" v-loading="loading" height="calc(100vh - 130px)">
        <el-table-column prop="module_label" label="模块" width="130" />
        <el-table-column prop="channel" label="通道" width="80" />
        <el-table-column label="状态" width="90">
          <template #default="{ row }">
            <el-tag :type="statusType(row.status)" size="small">
              {{ formatStatus(row.status) }}
            </el-tag>
          </template>
        </el-table-column>
        <el-table-column label="耗时" width="90">
          <template #default="{ row }">
            {{ formatDuration(row.duration_seconds) }}
          </template>
        </el-table-column>
        <el-table-column prop="start_time" label="开始时间" width="170">
          <template #default="{ row }">
            {{ row.start_time?.slice(0, 19) || '-' }}
          </template>
        </el-table-column>
        <el-table-column prop="created_at" label="创建时间" width="170">
          <template #default="{ row }">
            {{ row.created_at?.slice(0, 19) || '-' }}
          </template>
        </el-table-column>
        <el-table-column label="操作" width="100" fixed="right">
          <template #default="{ row }">
            <el-button size="small" text type="primary" @click="viewResult(row.id)">
              查看结果
            </el-button>
          </template>
        </el-table-column>
      </el-table>
    </div>

    <TestResultDialog
      v-model:visible="resultVisible"
      :run-id="selectedRunId"
    />
  </div>
</template>

<style scoped>
.header-left {
  display: flex;
  align-items: center;
  gap: 12px;
}
.header-left h2 {
  font-size: 16px;
  margin: 0;
}
.header-right {
  display: flex;
  gap: 8px;
}
</style>
