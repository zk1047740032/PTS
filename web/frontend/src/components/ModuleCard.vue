<script setup lang="ts">
import type { ModuleInfo } from '@/api/platform'

const props = defineProps<{
  module: ModuleInfo
  checked: boolean
}>()

const emit = defineEmits<{
  'update:checked': [value: boolean]
  'open': [module: ModuleInfo]
}>()

function toggle() {
  emit('update:checked', !props.checked)
}

function openModule() {
  emit('open', props.module)
}
</script>

<template>
  <div
    class="module-card"
    :class="{ 'is-running': module.is_running, 'is-checked': checked }"
    @dblclick="openModule"
  >
    <el-checkbox
      :model-value="checked"
      :disabled="module.is_running"
      @change="toggle"
    />
    <div class="module-info" @click="toggle">
      <span class="module-label">{{ module.label }}</span>
      <span class="module-channel">{{ module.channel }}</span>
    </div>
    <el-tag
      v-if="module.is_running"
      type="success"
      size="small"
      effect="dark"
    >
      运行中
    </el-tag>
    <el-button
      v-else
      size="small"
      text
      type="primary"
      @click.stop="openModule"
    >
      打开
    </el-button>
  </div>
</template>

<style scoped>
.module-card {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 8px 12px;
  border-radius: 6px;
  cursor: pointer;
  transition: background 0.2s;
}
.module-card:hover {
  background: #EBF0FA;
}
.module-card.is-checked {
  background: #EBF0FA;
}
.module-card.is-running {
  background: #F0FFF4;
}
.module-info {
  flex: 1;
  display: flex;
  flex-direction: column;
  min-width: 0;
}
.module-label {
  font-size: 13px;
  font-weight: 500;
}
.module-channel {
  font-size: 11px;
  color: var(--color-text-secondary);
}
</style>
