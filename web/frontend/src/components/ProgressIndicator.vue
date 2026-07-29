<script setup lang="ts">
import { computed } from 'vue'
import type { WSMessage } from '@/composables/useWebSocket'

const props = defineProps<{
  moduleKey: string
  status: string  // pending | running | completed | error | stopped
  messages: WSMessage[]
}>()

const emit = defineEmits<{
  stop: []
}>()

const phases = [
  { key: 'connect', label: '连接仪器' },
  { key: 'configure', label: '配置参数' },
  { key: 'measure', label: '执行测量' },
  { key: 'analyze', label: '数据分析' },
  { key: 'complete', label: '完成' },
]

const currentPhase = computed(() => {
  if (props.status === 'completed') return phases.length
  if (props.status === 'error') return -1
  const lastMsg = props.messages[props.messages.length - 1]
  if (!lastMsg) return 0
  const text = (lastMsg.payload?.message || '').toLowerCase()
  if (text.includes('分析') || text.includes('保存')) return 4
  if (text.includes('测') || text.includes('扫描') || text.includes('读取')) return 3
  if (text.includes('配置') || text.includes('设置')) return 2
  if (text.includes('连接') || text.includes('打开')) return 1
  return 1
})

const isRunning = computed(() => props.status === 'running')
const statusText = computed(() => {
  const map: Record<string, string> = {
    pending: '等待中',
    running: '运行中',
    completed: '已完成',
    error: '异常',
    stopped: '已停止',
  }
  return map[props.status] || props.status
})
const statusColor = computed(() => {
  const map: Record<string, string> = {
    pending: '#909399',
    running: '#2B6FF2',
    completed: '#38A169',
    error: '#E53E3E',
    stopped: '#F2994A',
  }
  return map[props.status] || '#909399'
})
</script>

<template>
  <div class="progress-indicator">
    <div class="progress-header">
      <span class="module-name">{{ moduleKey }}</span>
      <el-tag :color="statusColor" effect="dark" size="small">
        {{ statusText }}
      </el-tag>
      <el-button
        v-if="isRunning"
        type="danger"
        size="small"
        plain
        @click="emit('stop')"
      >
        停止
      </el-button>
    </div>

    <el-steps
      :active="currentPhase"
      align-center
      finish-status="success"
      process-status="process"
      class="progress-steps"
    >
      <el-step
        v-for="phase in phases"
        :key="phase.key"
        :title="phase.label"
      />
    </el-steps>

    <el-progress
      v-if="isRunning"
      :percentage="Math.min(currentPhase / phases.length * 100, 95)"
      :indeterminate="isRunning && currentPhase === 0"
      :stroke-width="6"
      :show-text="false"
    />
  </div>
</template>

<style scoped>
.progress-indicator {
  background: var(--color-card);
  border-radius: 8px;
  padding: 16px;
  border: 1px solid var(--color-border);
}
.progress-header {
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 12px;
}
.module-name {
  font-size: 15px;
  font-weight: 600;
}
.progress-steps {
  margin-bottom: 12px;
}
</style>
