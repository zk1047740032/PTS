<script setup lang="ts">
import { ref, watch } from 'vue'
import { fetchTestDetail, fetchOutputFiles, type TestDetail, type OutputFile } from '@/api/test'

const props = defineProps<{
  visible: boolean
  runId: string | null
}>()

const emit = defineEmits<{
  'update:visible': [value: boolean]
}>()

const detail = ref<TestDetail | null>(null)
const files = ref<OutputFile[]>([])
const loading = ref(false)

watch(() => props.runId, async (newId) => {
  if (newId) {
    loading.value = true
    try {
      const [d, f] = await Promise.all([
        fetchTestDetail(newId),
        fetchOutputFiles(newId),
      ])
      detail.value = d
      files.value = f
    } catch (e) {
      console.error(e)
    } finally {
      loading.value = false
    }
  }
})

function close() {
  emit('update:visible', false)
  detail.value = null
  files.value = []
}

function formatDuration(s: number | null): string {
  if (s == null) return '-'
  const m = Math.floor(s / 60)
  const sec = Math.round(s % 60)
  return m > 0 ? `${m}分${sec}秒` : `${sec}秒`
}
</script>

<template>
  <el-dialog
    :model-value="visible"
    title="测试结果"
    width="560px"
    @update:model-value="(v: boolean) => !v && close()"
  >
    <div v-loading="loading">
      <template v-if="detail">
        <el-descriptions :column="2" border size="small">
          <el-descriptions-item label="模块">{{ detail.module_label }}</el-descriptions-item>
          <el-descriptions-item label="通道">{{ detail.channel }}</el-descriptions-item>
          <el-descriptions-item label="状态">
            <el-tag
              :type="detail.status === 'completed' ? 'success' : detail.status === 'error' ? 'danger' : 'info'"
              size="small"
            >
              {{ detail.status === 'completed' ? '已完成' : detail.status === 'error' ? '异常' : detail.status }}
            </el-tag>
          </el-descriptions-item>
          <el-descriptions-item label="耗时">{{ formatDuration(detail.duration_seconds) }}</el-descriptions-item>
          <el-descriptions-item label="开始时间">{{ detail.start_time || '-' }}</el-descriptions-item>
          <el-descriptions-item label="结束时间">{{ detail.end_time || '-' }}</el-descriptions-item>
        </el-descriptions>

        <div v-if="detail.result_summary" class="result-summary">
          <h4>结果摘要</h4>
          <pre>{{ JSON.stringify(detail.result_summary, null, 2) }}</pre>
        </div>

        <div v-if="files.length > 0" class="output-files">
          <h4>输出文件</h4>
          <div v-for="f in files" :key="f.name" class="file-row">
            <span>{{ f.name }}</span>
            <span class="file-size">{{ (f.size / 1024).toFixed(1) }} KB</span>
            <el-button size="small" text type="primary" @click="window.open(f.url)">
              下载
            </el-button>
          </div>
        </div>
      </template>
    </div>

    <template #footer>
      <el-button @click="close">关闭</el-button>
    </template>
  </el-dialog>
</template>

<style scoped>
.result-summary {
  margin-top: 16px;
}
.result-summary h4 {
  font-size: 13px;
  margin-bottom: 6px;
}
.result-summary pre {
  background: #f7f8fa;
  padding: 10px;
  border-radius: 6px;
  font-size: 12px;
  overflow-x: auto;
  max-height: 200px;
}
.output-files {
  margin-top: 16px;
}
.output-files h4 {
  font-size: 13px;
  margin-bottom: 6px;
}
.file-row {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 4px 0;
  font-size: 13px;
}
.file-size {
  color: var(--color-text-secondary);
  font-size: 12px;
}
</style>
