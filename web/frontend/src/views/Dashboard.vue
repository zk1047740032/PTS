<script setup lang="ts">
import { ref, onMounted, onUnmounted, computed } from 'vue'
import { useRouter } from 'vue-router'
import { ElMessage } from 'element-plus'
import {
  fetchModules,
  startTest,
  stopTest,
  batchStartTests,
  type ModuleInfo,
  type ChannelGroups,
} from '@/api/platform'
import { useWebSocket } from '@/composables/useWebSocket'
import ModuleCard from '@/components/ModuleCard.vue'
import LogStream from '@/components/LogStream.vue'
import InstrumentStatus from '@/components/InstrumentStatus.vue'
import TestResultDialog from '@/components/TestResultDialog.vue'

const router = useRouter()

// WebSocket
const { connected, logEntries, messages } = useWebSocket()

// Agent 状态
interface AgentInfo {
  agent_id: string
  hostname: string
  modules: string[]
  online: boolean
}
const agents = ref<AgentInfo[]>([])

// 监听 Agent 上/下线消息
import { watch } from 'vue'
watch(() => messages.value.length, () => {
  const latest = messages.value[messages.value.length - 1]
  if (!latest) return
  if (latest.type === 'agent_online') {
    loadAgents()
  } else if (latest.type === 'agent_offline') {
    loadAgents()
  }
})

async function loadAgents() {
  try {
    const res = await fetch('/api/rin/agents')
    const data = await res.json()
    agents.value = data.agents
  } catch {}
}

// State
const modules = ref<ModuleInfo[]>([])
const channelGroups = ref<ChannelGroups>({})
const checkedModules = ref<Set<string>>(new Set())
const activeTab = ref('光路A')
const activeSubTab = ref('通道3')
const runningModules = ref<Set<string>>(new Set())
const loading = ref(false)

// 结果弹窗
const resultVisible = ref(false)
const resultRunId = ref<string | null>(null)

// 通道和子标签
const channelTabs = computed(() => {
  const groups = channelGroups.value[activeTab.value] || {}
  return Object.keys(groups)
})

const currentModules = computed(() => {
  if (!channelGroups.value) return []
  const groups = channelGroups.value[activeTab.value] || {}
  const moduleKeys = groups[activeSubTab.value] || []
  return modules.value.filter(m => moduleKeys.includes(m.key))
})

// Data
async function loadModules() {
  try {
    const data = await fetchModules()
    modules.value = data.modules
    channelGroups.value = data.channel_groups
    updateRunningState()
  } catch (e: any) {
    ElMessage.error('加载模块列表失败: ' + e.message)
  }
}

function updateRunningState() {
  runningModules.value = new Set(
    modules.value.filter(m => m.is_running).map(m => m.key)
  )
}

// Actions
function toggleCheck(key: string) {
  const s = new Set(checkedModules.value)
  if (s.has(key)) s.delete(key)
  else s.add(key)
  checkedModules.value = s
}

function toggleAll() {
  const allKeys = currentModules.value.map(m => m.key)
  const allChecked = allKeys.every(k => checkedModules.value.has(k))
  if (allChecked) {
    // 清空当前通道
    const s = new Set(checkedModules.value)
    allKeys.forEach(k => s.delete(k))
    checkedModules.value = s
  } else {
    checkedModules.value = new Set([...checkedModules.value, ...allKeys])
  }
}

async function openWindows() {
  const toOpen = currentModules.value.filter(
    m => checkedModules.value.has(m.key) && !runningModules.value.has(m.key)
  )
  if (toOpen.length === 0) {
    ElMessage.info('请先勾选要打开的模块')
    return
  }
  loading.value = true
  try {
    await batchStartTests(toOpen.map(m => m.key), false)
    ElMessage.success(`已打开 ${toOpen.length} 个窗口`)
    await loadModules()
  } catch (e: any) {
    ElMessage.error('启动失败: ' + e.message)
  } finally {
    loading.value = false
  }
}

async function runOneClick() {
  const toRun = currentModules.value.filter(
    m => checkedModules.value.has(m.key) && !runningModules.value.has(m.key)
  )
  if (toRun.length === 0) {
    ElMessage.info('请先勾选要测试的模块')
    return
  }
  loading.value = true
  try {
    await batchStartTests(toRun.map(m => m.key), true)
    ElMessage.success(`已启动 ${toRun.length} 个测试`)
    await loadModules()
  } catch (e: any) {
    ElMessage.error('启动失败: ' + e.message)
  } finally {
    loading.value = false
  }
}

async function openSingleModule(module: ModuleInfo) {
  if (runningModules.value.has(module.key)) {
    // 已经在运行，跳转详情页
    router.push({ name: 'TestDetail', params: { moduleKey: module.key } })
    return
  }
  try {
    await startTest(module.key, false)
    ElMessage.success(`${module.label} 窗口已打开`)
    await loadModules()
  } catch (e: any) {
    ElMessage.error('打开失败: ' + e.message)
  }
}

async function handleStop(moduleKey: string) {
  try {
    await stopTest(moduleKey)
    ElMessage.success(`${moduleKey} 已停止`)
    await loadModules()
  } catch (e: any) {
    ElMessage.error('停止失败: ' + e.message)
  }
}

function showResults() {
  // 取最近完成的 run
  const completed = modules.value.filter(m => m.is_running === false)
  if (completed.length === 0) {
    ElMessage.info('暂无完成的测试')
    return
  }
  // 打开历史页
  router.push({ name: 'TestHistory' })
}

onMounted(() => {
  loadModules()
  loadAgents()
})
</script>

<template>
  <div class="page-container">
    <!-- Header -->
    <div class="page-header">
      <div class="header-left">
        <h1 class="app-title">
          <span class="title-accent">PTS</span>-种子
        </h1>
        <span class="title-sub">单独测试或一键测试</span>
      </div>
      <div class="header-right">
        <el-tag :type="connected ? 'success' : 'danger'" size="small" effect="dark">
          {{ connected ? '已连接' : '未连接' }}
        </el-tag>
        <el-button size="small" @click="$router.push({ name: 'ConfigView' })">
          <el-icon><Setting /></el-icon>
          配置参数
        </el-button>
        <el-button size="small" @click="$router.push({ name: 'TestHistory' })">
          <el-icon><Clock /></el-icon>
          测试历史
        </el-button>
      </div>
    </div>

    <!-- Body -->
    <div class="page-body dashboard-body">
      <!-- Left Panel: Module Selection -->
      <div class="left-panel">
        <div class="panel-card">
          <div class="card-header">
            <span>控制面板</span>
            <el-button size="small" text @click="toggleAll">
              {{ currentModules.every(m => checkedModules.has(m.key)) ? '清空' : '全选' }}
            </el-button>
          </div>

          <el-tabs v-model="activeTab" class="main-tabs">
            <el-tab-pane label="光路A" name="光路A">
              <el-tabs v-model="activeSubTab" type="card" class="sub-tabs">
                <el-tab-pane
                  v-for="tab in channelTabs"
                  :key="tab"
                  :label="tab"
                  :name="tab"
                >
                  <div class="module-list">
                    <ModuleCard
                      v-for="m in currentModules"
                      :key="m.key"
                      :module="m"
                      :checked="checkedModules.has(m.key)"
                      @update:checked="toggleCheck(m.key)"
                      @open="openSingleModule"
                    />
                  </div>
                </el-tab-pane>
              </el-tabs>
            </el-tab-pane>
            <el-tab-pane label="光路B" name="光路B">
              <div class="module-list">
                <ModuleCard
                  v-for="m in currentModules"
                  :key="m.key"
                  :module="m"
                  :checked="checkedModules.has(m.key)"
                  @update:checked="toggleCheck(m.key)"
                  @open="openSingleModule"
                />
              </div>
            </el-tab-pane>
          </el-tabs>

          <div class="action-buttons">
            <el-button type="primary" @click="openWindows" :loading="loading">
              <el-icon><FolderOpened /></el-icon>
              打开
            </el-button>
            <el-button type="warning" @click="showResults">
              <el-icon><Document /></el-icon>
              测试结果
            </el-button>
            <el-button type="success" @click="runOneClick" :loading="loading">
              <el-icon><VideoPlay /></el-icon>
              一键测试
            </el-button>
          </div>
        </div>
      </div>

      <!-- Right Panel: Monitor -->
      <div class="right-panel">
        <!-- Log Stream -->
        <div class="panel-card log-card">
          <div class="card-header">
            <span>运行状态监控</span>
          </div>
          <LogStream :entries="logEntries" :max-height="400" />
        </div>

        <!-- Instrument Status -->
        <div class="bottom-row">
          <InstrumentStatus />

          <!-- Agent 状态 -->
          <div class="agent-status card">
            <div class="card-header">
              <span>仪器 Agent</span>
              <el-tag :type="agents.length > 0 ? 'success' : 'info'" size="small">
                {{ agents.length }} 在线
              </el-tag>
            </div>
            <div class="agent-list" v-if="agents.length > 0">
              <div v-for="a in agents" :key="a.agent_id" class="agent-row">
                <el-icon color="#38A169"><CircleCheckFilled /></el-icon>
                <span class="agent-name">{{ a.agent_id }}</span>
                <span class="agent-host">{{ a.hostname }}</span>
                <el-tag v-for="m in a.modules" :key="m" size="small" type="success" effect="plain">
                  {{ m }}
                </el-tag>
              </div>
            </div>
            <div v-else class="agent-empty">
              <span>无 Agent 连接 — 测试将在本机执行</span>
            </div>
          </div>
        </div>
      </div>
    </div>

    <!-- Result Dialog -->
    <TestResultDialog
      v-model:visible="resultVisible"
      :run-id="resultRunId"
    />
  </div>
</template>

<style scoped>
.dashboard-body {
  display: flex;
  gap: 12px;
  padding: 12px;
}

.left-panel {
  width: 320px;
  flex-shrink: 0;
}

.right-panel {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: 12px;
  min-width: 0;
}

.panel-card {
  background: var(--color-card);
  border-radius: 8px;
  border: 1px solid var(--color-border);
  overflow: hidden;
}

.card-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px 14px;
  border-bottom: 1px solid var(--color-border);
  font-weight: 600;
  font-size: 14px;
}

.module-list {
  max-height: 320px;
  overflow-y: auto;
}

.action-buttons {
  display: flex;
  gap: 8px;
  padding: 12px;
  border-top: 1px solid var(--color-border);
}
.action-buttons .el-button {
  flex: 1;
}

.log-card {
  flex: 1;
}

.bottom-row {
  flex-shrink: 0;
  display: flex;
  flex-direction: column;
  gap: 12px;
}

/* Agent status */
.agent-status.card {
  background: var(--color-card);
  border-radius: 8px;
  border: 1px solid var(--color-border);
  padding: 12px;
}
.agent-list {
  display: flex;
  flex-direction: column;
  gap: 6px;
  margin-top: 8px;
}
.agent-row {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 12px;
}
.agent-name {
  font-weight: 600;
  color: var(--color-text);
}
.agent-host {
  color: var(--color-text-secondary);
  font-family: monospace;
}
.agent-empty {
  margin-top: 8px;
  font-size: 12px;
  color: var(--color-text-secondary);
  text-align: center;
  padding: 8px;
}

.header-left {
  display: flex;
  align-items: baseline;
  gap: 10px;
}
.app-title {
  font-size: 18px;
  font-weight: 700;
  margin: 0;
}
.title-accent {
  color: var(--color-accent);
}
.title-sub {
  font-size: 12px;
  color: var(--color-text-secondary);
}
.header-right {
  display: flex;
  align-items: center;
  gap: 8px;
}

.main-tabs, .sub-tabs {
  --el-tabs-header-height: 34px;
}
</style>
