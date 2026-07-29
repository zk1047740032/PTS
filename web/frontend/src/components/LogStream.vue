<script setup lang="ts">
import { ref, watch, nextTick, computed } from 'vue'
import type { WSMessage } from '@/composables/useWebSocket'

const props = defineProps<{
  entries: WSMessage[]
  maxHeight?: number
}>()

const logContainer = ref<HTMLElement>()

// 自动滚动到底部
watch(
  () => props.entries.length,
  async () => {
    await nextTick()
    if (logContainer.value) {
      logContainer.value.scrollTop = logContainer.value.scrollHeight
    }
  }
)

const logClass = (entry: WSMessage) => ({
  'log-line': true,
  'log-running': entry.type === 'running',
  'log-completed': entry.type === 'completed',
  'log-error': entry.type === 'error',
  'log-warning': entry.type === 'warning',
})

const typeTag = (type: string) => {
  const map: Record<string, string> = {
    running: '运行',
    completed: '完成',
    error: '错误',
    warning: '警告',
    log: '信息',
  }
  return map[type] || type
}

const tagType = (type: string) => {
  const map: Record<string, string> = {
    running: '',
    completed: 'success',
    error: 'danger',
    warning: 'warning',
  }
  return map[type] || 'info'
}

function formatTime(entry: WSMessage): string {
  if (entry.payload?.data?.timestamp) return entry.payload.data.timestamp
  const now = new Date()
  return `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}:${String(now.getSeconds()).padStart(2, '0')}`
}

const containerStyle = computed(() => ({
  maxHeight: props.maxHeight ? `${props.maxHeight}px` : '100%',
}))
</script>

<template>
  <div ref="logContainer" class="log-stream" :style="containerStyle">
    <div v-if="entries.length === 0" class="log-empty">等待日志...</div>
    <div
      v-for="(entry, idx) in entries"
      :key="idx"
      :class="logClass(entry)"
    >
      <span class="log-time">{{ formatTime(entry) }}</span>
      <el-tag :type="tagType(entry.type)" size="small" effect="plain">
        {{ tagType(entry.type) }}
      </el-tag>
      <span class="log-module">{{ entry.module }}</span>
      <span class="log-msg">{{ entry.payload?.message || JSON.stringify(entry.payload) }}</span>
    </div>
  </div>
</template>

<style scoped>
.log-stream {
  overflow-y: auto;
  font-family: 'Consolas', 'Courier New', monospace;
  font-size: 12px;
  line-height: 1.8;
  padding: 8px;
  background: #1E1E1E;
  color: #D4D4D4;
  border-radius: 6px;
}
.log-empty {
  color: #888;
  text-align: center;
  padding: 20px;
}
.log-line {
  display: flex;
  align-items: flex-start;
  gap: 6px;
  padding: 1px 0;
}
.log-time {
  color: #6A9955;
  white-space: nowrap;
  min-width: 60px;
}
.log-module {
  color: #569CD6;
  min-width: 80px;
}
.log-msg {
  color: #D4D4D4;
  white-space: pre-wrap;
  word-break: break-all;
}
.log-running { background: rgba(43, 111, 242, 0.08); }
.log-completed { background: rgba(56, 161, 105, 0.08); }
.log-error { background: rgba(229, 62, 62, 0.08); }
.log-warning { background: rgba(242, 153, 74, 0.08); }
</style>
