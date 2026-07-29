/** 测试运行 API 封装 */

const BASE = '/api'

export interface TestRunSummary {
  id: string
  module_name: string
  module_label: string
  channel: string
  status: string
  start_time: string | null
  end_time: string | null
  duration_seconds: number | null
  result_summary: any
  output_dir: string | null
  created_at: string | null
}

export interface TestLogEntry {
  sequence: number
  timestamp: string | null
  level: string
  message: string
  data: any
}

export interface TestDetail extends TestRunSummary {
  logs: TestLogEntry[]
}

export interface OutputFile {
  name: string
  size: number
  type: string
  url: string
}

export async function fetchTestHistory(
  module?: string,
  limit = 20,
  offset = 0
): Promise<{ runs: TestRunSummary[] }> {
  const params = new URLSearchParams()
  if (module) params.set('module', module)
  params.set('limit', String(limit))
  params.set('offset', String(offset))
  const res = await fetch(`${BASE}/tests/history?${params}`)
  if (!res.ok) throw new Error(`HTTP ${res.status}`)
  return res.json()
}

export async function fetchTestDetail(runId: string): Promise<TestDetail> {
  const res = await fetch(`${BASE}/tests/${runId}`)
  if (!res.ok) throw new Error(`HTTP ${res.status}`)
  return res.json()
}

export async function fetchOutputFiles(runId: string): Promise<OutputFile[]> {
  const res = await fetch(`${BASE}/tests/${runId}/files`)
  if (!res.ok) throw new Error(`HTTP ${res.status}`)
  const data = await res.json()
  return data.files
}
