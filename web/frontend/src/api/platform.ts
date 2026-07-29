/** 平台 API 封装 */

const BASE = '/api'

export interface ModuleInfo {
  key: string
  label: string
  module_path: string
  gui_class: string
  channel: string
  start_method: string
  is_running: boolean
}

export interface ChannelGroups {
  [pathName: string]: { [channelName: string]: string[] }
}

export interface ModulesResponse {
  modules: ModuleInfo[]
  channel_groups: ChannelGroups
}

export async function fetchModules(): Promise<ModulesResponse> {
  const res = await fetch(`${BASE}/modules`)
  if (!res.ok) throw new Error(`HTTP ${res.status}`)
  return res.json()
}

export async function fetchRunningTests(): Promise<string[]> {
  const res = await fetch(`${BASE}/tests/running`)
  if (!res.ok) throw new Error(`HTTP ${res.status}`)
  const data = await res.json()
  return data.running
}

export async function startTest(moduleKey: string, autoStart = true): Promise<{ run_id: string }> {
  const res = await fetch(`${BASE}/tests/start`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ module: moduleKey, auto_start: autoStart }),
  })
  if (!res.ok) {
    const err = await res.json()
    throw new Error(err.detail || `HTTP ${res.status}`)
  }
  return res.json()
}

export async function stopTest(moduleKey: string): Promise<void> {
  const res = await fetch(`${BASE}/tests/stop/${moduleKey}`, { method: 'POST' })
  if (!res.ok) throw new Error(`HTTP ${res.status}`)
}

export async function batchStartTests(modules: string[], autoStart = true): Promise<any> {
  const res = await fetch(`${BASE}/tests/batch-start`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ modules, auto_start: autoStart }),
  })
  if (!res.ok) throw new Error(`HTTP ${res.status}`)
  return res.json()
}

export interface InstrumentInfo {
  name: string
  ip: string
  reachable: boolean
}

export async function fetchInstrumentStatus(): Promise<InstrumentInfo[]> {
  const res = await fetch(`${BASE}/instruments/status`)
  if (!res.ok) throw new Error(`HTTP ${res.status}`)
  const data = await res.json()
  return data.instruments
}
