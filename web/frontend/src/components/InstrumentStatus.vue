<script setup lang="ts">
import { ref, onMounted, onUnmounted } from 'vue'
import { fetchInstrumentStatus, type InstrumentInfo } from '@/api/platform'

const instruments = ref<InstrumentInfo[]>([])
const loading = ref(false)
const error = ref('')
let timer: number | null = null

async function refresh() {
  loading.value = true
  error.value = ''
  try {
    instruments.value = await fetchInstrumentStatus()
  } catch (e: any) {
    error.value = e.message
  } finally {
    loading.value = false
  }
}

onMounted(() => {
  refresh()
  // 每 30 秒自动刷新
  timer = window.setInterval(refresh, 30000)
})

onUnmounted(() => {
  if (timer) clearInterval(timer)
})
</script>

<template>
  <div class="instrument-status">
    <div class="inst-header">
      <h4>仪器连接状态</h4>
      <el-button size="small" text @click="refresh" :loading="loading">
        <el-icon><Refresh /></el-icon>
        刷新
      </el-button>
    </div>
    <div v-if="error" class="inst-error">{{ error }}</div>
    <div class="inst-list">
      <div v-for="inst in instruments" :key="inst.name" class="inst-row">
        <span class="inst-name">{{ inst.name }}</span>
        <span class="inst-ip">{{ inst.ip }}</span>
        <el-tag
          :type="inst.reachable ? 'success' : 'danger'"
          size="small"
          effect="plain"
        >
          {{ inst.reachable ? '可达' : '不可达' }}
        </el-tag>
      </div>
    </div>
  </div>
</template>

<style scoped>
.instrument-status {
  background: var(--color-card);
  border-radius: 8px;
  border: 1px solid var(--color-border);
  padding: 12px;
}
.inst-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 8px;
}
.inst-header h4 {
  font-size: 14px;
  margin: 0;
}
.inst-error {
  color: var(--color-danger);
  font-size: 12px;
  margin-bottom: 6px;
}
.inst-list {
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.inst-row {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 12px;
  padding: 4px 0;
}
.inst-name {
  min-width: 120px;
  color: var(--color-text);
}
.inst-ip {
  min-width: 110px;
  color: var(--color-text-secondary);
  font-family: monospace;
}
</style>
