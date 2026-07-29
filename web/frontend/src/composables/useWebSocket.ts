/** WebSocket 组合式函数 */

import { ref, onUnmounted, type Ref } from 'vue'

export interface WSMessage {
  type: 'log' | 'progress' | 'status' | 'data' | 'running' | 'completed' | 'error' | 'warning'
  module: string
  payload: {
    message?: string
    seq?: number
    status?: string
    progress?: number
    data?: any
  }
}

export function useWebSocket(moduleKey?: string) {
  const connected = ref(false)
  const lastMessage: Ref<WSMessage | null> = ref(null)
  const messages: Ref<WSMessage[]> = ref([])
  const logEntries: Ref<WSMessage[]> = ref([])

  let ws: WebSocket | null = null
  let reconnectTimer: number | null = null

  function connect() {
    const protocol = location.protocol === 'https:' ? 'wss:' : 'ws:'
    const url = moduleKey
      ? `${protocol}//${location.host}/ws/${moduleKey}`
      : `${protocol}//${location.host}/ws`

    ws = new WebSocket(url)

    ws.onopen = () => {
      connected.value = true
      console.log(`WebSocket connected${moduleKey ? ` to ${moduleKey}` : ''}`)
    }

    ws.onmessage = (event) => {
      try {
        const msg: WSMessage = JSON.parse(event.data)
        lastMessage.value = msg
        messages.value.push(msg)
        // 只保留最近 500 条
        if (messages.value.length > 500) {
          messages.value = messages.value.slice(-500)
        }
        // log 类型消息单独收集
        if (['log', 'running', 'completed', 'error', 'warning'].includes(msg.type)) {
          logEntries.value.push(msg)
          if (logEntries.value.length > 1000) {
            logEntries.value = logEntries.value.slice(-1000)
          }
        }
      } catch (e) {
        console.warn('WebSocket parse error:', e)
      }
    }

    ws.onclose = () => {
      connected.value = false
      // 自动重连
      reconnectTimer = window.setTimeout(() => {
        if (ws?.readyState === WebSocket.CLOSED) {
          connect()
        }
      }, 3000)
    }

    ws.onerror = () => {
      ws?.close()
    }
  }

  function send(data: any) {
    if (ws?.readyState === WebSocket.OPEN) {
      ws.send(JSON.stringify(data))
    }
  }

  function disconnect() {
    if (reconnectTimer) {
      clearTimeout(reconnectTimer)
      reconnectTimer = null
    }
    ws?.close()
    ws = null
  }

  // 自动连接
  connect()

  onUnmounted(() => {
    disconnect()
  })

  return {
    connected,
    lastMessage,
    messages,
    logEntries,
    send,
    disconnect,
  }
}
