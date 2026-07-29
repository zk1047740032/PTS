<script setup lang="ts">
import { ref, computed, onMounted, onUnmounted, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import { useWebSocket } from '@/composables/useWebSocket'
import LogStream from '@/components/LogStream.vue'
import ProgressIndicator from '@/components/ProgressIndicator.vue'
import LiveChart from '@/components/LiveChart.vue'

const route = useRoute()
const router = useRouter()
const moduleKey = computed(() => route.params.moduleKey as string)

// 判断是否为 RIN 测试
const isRin = computed(() => moduleKey.value === 'Rin_FSV3004')

// WebSocket 订阅
const { connected, logEntries, messages } = useWebSocket(moduleKey.value)

// State
const testStatus = ref('idle')  // idle | running | completed | error | stopped
const loading = ref(false)

// RIN 参数
const dcValue = ref(2.4)
const ipAddress = ref('192.168.7.10')

// 结果
const result = ref<any>(null)
const resultPeaks = ref<any>(null)
const resultPoints = ref<any[]>([])
const resultRmsMax = ref('')
const chartUrl = ref('')

// 实时数据（从 WS 消息中提取）
const chartX = ref<number[]>([])
const chartY = ref<number[]>([])

// 监听 WS 消息更新状态
watch(
  () => messages.value.length,
  () => {
    const latest = messages.value[messages.value.length - 1]
    if (!latest) return
    switch (latest.type) {
      case 'running':
      case 'status':
        if (latest.payload?.status === 'running') testStatus.value = 'running'
        break
      case 'completed':
        testStatus.value = 'completed'
        if (latest.payload?.result) {
          result.value = latest.payload.result
          resultPeaks.value = latest.payload.result.peak_info
          resultPoints.value = latest.payload.result.target_points || []
          resultRmsMax.value = latest.payload.result.integrated_rms_max || ''
          chartUrl.value = `/api/rin/chart?t=${Date.now()}`
        }
        break
      case 'error':
        testStatus.value = 'error'
        break
    }
    // 提取进度信息
    if (latest.payload?.data?.x && latest.payload?.data?.y) {
      chartX.value = latest.payload.data.x
      chartY.value = latest.payload.data.y
    }
  }
)

// Actions
async function startRin() {
  loading.value = true
  try {
    const res = await fetch('/api/rin/start', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        dc_value: dcValue.value,
        ip_address: ipAddress.value || undefined,
      }),
    })
    if (!res.ok) {
      const err = await res.json()
      throw new Error(err.detail || `HTTP ${res.status}`)
    }
    testStatus.value = 'running'
    ElMessage.success('RIN 测试已启动')
  } catch (e: any) {
    ElMessage.error('启动失败: ' + e.message)
  } finally {
    loading.value = false
  }
}

async function stopRin() {
  try {
    const res = await fetch('/api/rin/stop', { method: 'POST' })
    if (!res.ok) throw new Error(`HTTP ${res.status}`)
    testStatus.value = 'stopped'
    ElMessage.success('停止信号已发送')
  } catch (e: any) {
    ElMessage.error('停止失败: ' + e.message)
  }
}

async function fetchStatus() {
  try {
    const res = await fetch('/api/rin/status')
    if (!res.ok) return
    const data = await res.json()
    if (data.running) testStatus.value = 'running'
    else if (data.status !== 'idle') testStatus.value = data.status
    if (data.status === 'completed') {
      await fetchResult()
    }
  } catch {}
}

async function fetchResult() {
  try {
    const res = await fetch('/api/rin/result')
    if (!res.ok) return
    const data = await res.json()
    if (data.status === 'completed') {
      result.value = data
      resultPeaks.value = data.peak_info
      resultPoints.value = data.target_points || []
      resultRmsMax.value = data.integrated_rms_max || ''
      chartUrl.value = `/api/rin/chart?t=${Date.now()}`
    }
  } catch {}
}

onMounted(() => {
  fetchStatus()
})
</script>

<template>
  <div class="page-container">
    <!-- Header -->
    <div class="page-header">
      <div class="header-left">
        <el-button text @click="router.push('/')">
          <el-icon><ArrowLeft /></el-icon>
          返回
        </el-button>
        <h2>{{ moduleKey }} 测试详情</h2>
        <el-tag :type="connected ? 'success' : 'danger'" size="small">
          {{ connected ? '实时连接' : '连接断开' }}
        </el-tag>
      </div>
      <div class="header-right">
        <el-button
          v-if="testStatus === 'running'"
          type="danger" plain @click="stopRin"
        >
          停止测试
        </el-button>
      </div>
    </div>

    <!-- Body -->
    <div class="page-body detail-body">

      <!-- RIN 专属参数面板 -->
      <div v-if="isRin" class="rin-params">
        <div class="param-card">
          <h3>RIN 测量参数</h3>
          <div class="param-form">
            <div class="param-item">
              <label>DC 值 (V)</label>
              <el-input-number
                v-model="dcValue"
                :min="0.1" :max="100" :step="0.1" :precision="2"
                size="small"
                :disabled="testStatus === 'running'"
              />
            </div>
            <div class="param-item">
              <label>仪器 IP</label>
              <el-input
                v-model="ipAddress"
                size="small"
                :disabled="testStatus === 'running'"
                style="width: 160px"
              />
            </div>
            <div class="param-actions">
              <el-button
                v-if="testStatus !== 'running'"
                type="primary"
                @click="startRin"
                :loading="loading"
              >
                <el-icon><VideoPlay /></el-icon>
                开始测量
              </el-button>
            </div>
          </div>
        </div>
      </div>

      <!-- Progress -->
      <ProgressIndicator
        :module-key="moduleKey"
        :status="testStatus"
        :messages="messages"
        @stop="stopRin"
      />

      <!-- RIN Results -->
      <div v-if="isRin && testStatus === 'completed' && result" class="rin-results">
        <div class="result-card">
          <h3>测量结果</h3>

          <!-- Peak info -->
          <div v-if="resultPeaks" class="peak-info">
            <h4>驰豫振荡峰</h4>
            <el-descriptions :column="2" border size="small">
              <el-descriptions-item label="频率范围">
                {{ resultPeaks.start_hz }} - {{ resultPeaks.stop_hz }} Hz
              </el-descriptions-item>
              <el-descriptions-item label="峰值 RIN">
                {{ resultPeaks.rin_dbc_hz }} dBc/Hz
              </el-descriptions-item>
              <el-descriptions-item label="峰值频率">
                {{ resultPeaks.freq_hz }} Hz
              </el-descriptions-item>
              <el-descriptions-item label="RMS 最大值">
                {{ resultRmsMax }}
              </el-descriptions-item>
            </el-descriptions>
          </div>

          <!-- Target points -->
          <div v-if="resultPoints.length > 0" class="target-points">
            <h4>指定频点 RIN 值</h4>
            <el-table :data="resultPoints" size="small" max-height="200">
              <el-table-column prop="freq_hz" label="频率 (Hz)" />
              <el-table-column prop="rin_dbc_hz" label="RIN (dBc/Hz)" />
            </el-table>
          </div>

          <!-- Chart -->
          <div v-if="chartUrl" class="rin-chart">
            <h4>RIN 图表</h4>
            <img :src="chartUrl" alt="RIN Chart" style="max-width:100%; border-radius:8px; border:1px solid #E2E8F0" />
          </div>
        </div>
      </div>

      <!-- Log Stream -->
      <div class="detail-log">
        <h3>测试日志</h3>
        <LogStream :entries="logEntries" :max-height="300" />
      </div>
    </div>
  </div>
</template>

<style scoped>
.detail-body {
  display: flex;
  flex-direction: column;
  gap: 12px;
  overflow-y: auto;
}
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

/* RIN 参数面板 */
.rin-params {
  flex-shrink: 0;
}
.param-card {
  background: var(--color-card);
  border-radius: 8px;
  border: 1px solid var(--color-border);
  padding: 16px;
}
.param-card h3 {
  font-size: 14px;
  margin: 0 0 12px 0;
}
.param-form {
  display: flex;
  align-items: flex-end;
  gap: 20px;
  flex-wrap: wrap;
}
.param-item {
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.param-item label {
  font-size: 12px;
  color: var(--color-text-secondary);
}
.param-actions {
  margin-left: auto;
}

/* Results */
.rin-results {
  flex-shrink: 0;
}
.result-card {
  background: var(--color-card);
  border-radius: 8px;
  border: 1px solid var(--color-border);
  padding: 16px;
}
.result-card h3 {
  font-size: 14px;
  margin: 0 0 12px 0;
}
.peak-info {
  margin-bottom: 16px;
}
.peak-info h4, .target-points h4, .rin-chart h4 {
  font-size: 13px;
  margin: 0 0 8px 0;
  color: var(--color-text-secondary);
}
.target-points {
  margin-bottom: 16px;
}

.detail-log {
  flex-shrink: 0;
}
.detail-log h3 {
  font-size: 14px;
  margin-bottom: 8px;
}
</style>
