<script setup lang="ts">
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage, ElMessageBox } from 'element-plus'

const router = useRouter()

interface ConfigSection {
  name: string
  fields: { key: string; label: string; value: string; description: string }[]
}

const configSections = ref<ConfigSection[]>([])
const loading = ref(false)
const saving = ref(false)

async function loadConfig() {
  loading.value = true
  try {
    const res = await fetch('/api/config')
    const data = await res.json()

    const sections: ConfigSection[] = [
      {
        name: '网络仪器地址',
        fields: Object.entries(data.defaults.network || {})
          .filter(([k]) => k !== '_frozen')
          .map(([k, v]) => ({
            key: `network.${k}`,
            label: k,
            value: String(data.overrides[`network.${k}`] || v || ''),
            description: '',
          })),
      },
      {
        name: 'USB 设备',
        fields: Object.entries(data.defaults.usb || {})
          .filter(([k]) => k !== '_frozen')
          .map(([k, v]) => ({
            key: `usb.${k}`,
            label: k,
            value: String(data.overrides[`usb.${k}`] || v || ''),
            description: '',
          })),
      },
      {
        name: '数据目录',
        fields: Object.entries(data.defaults.dirs || {})
          .filter(([k]) => k !== '_frozen')
          .map(([k, v]) => ({
            key: `dirs.${k}`,
            label: k,
            value: String(data.overrides[`dirs.${k}`] || v || ''),
            description: '',
          })),
      },
    ]
    configSections.value = sections
  } catch (e: any) {
    ElMessage.error('加载配置失败: ' + e.message)
  } finally {
    loading.value = false
  }
}

async function saveConfig() {
  saving.value = true
  try {
    const overrides: Record<string, any> = {}
    for (const section of configSections.value) {
      for (const field of section.fields) {
        if (field.value) {
          overrides[field.key] = field.value
        }
      }
    }
    const res = await fetch('/api/config/user', {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(overrides),
    })
    if (!res.ok) throw new Error(`HTTP ${res.status}`)
    ElMessage.success('配置已保存')
  } catch (e: any) {
    ElMessage.error('保存失败: ' + e.message)
  } finally {
    saving.value = false
  }
}

async function resetConfig() {
  try {
    await ElMessageBox.confirm('确定要重置所有用户配置为默认值吗？', '确认重置', {
      type: 'warning',
    })
    const res = await fetch('/api/config/user/reset', { method: 'POST' })
    if (!res.ok) throw new Error(`HTTP ${res.status}`)
    ElMessage.success('配置已重置')
    await loadConfig()
  } catch (e: any) {
    if (e !== 'cancel') ElMessage.error('重置失败: ' + e.message)
  }
}

onMounted(() => {
  loadConfig()
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
        <h2>配置参数</h2>
      </div>
      <div class="header-right">
        <el-button @click="resetConfig">恢复默认</el-button>
        <el-button type="primary" @click="saveConfig" :loading="saving">保存</el-button>
      </div>
    </div>

    <div class="page-body config-body" v-loading="loading">
      <el-collapse v-model="['0', '1', '2']">
        <el-collapse-item
          v-for="(section, idx) in configSections"
          :key="idx"
          :title="section.name"
          :name="String(idx)"
        >
          <el-form label-width="180px" size="small">
            <el-form-item
              v-for="field in section.fields"
              :key="field.key"
              :label="field.label"
            >
              <el-input v-model="field.value" :placeholder="field.description" />
            </el-form-item>
          </el-form>
        </el-collapse-item>
      </el-collapse>
    </div>
  </div>
</template>

<style scoped>
.config-body {
  max-width: 800px;
  margin: 0 auto;
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
</style>
