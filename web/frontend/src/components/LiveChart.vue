<script setup lang="ts">
import { ref, watch, onMounted, onUnmounted, computed } from 'vue'
import VChart from 'vue-echarts'
import { use } from 'echarts/core'
import { CanvasRenderer } from 'echarts/renderers'
import { LineChart, ScatterChart } from 'echarts/charts'
import {
  TitleComponent,
  TooltipComponent,
  LegendComponent,
  GridComponent,
  DataZoomComponent,
} from 'echarts/components'

use([
  CanvasRenderer,
  LineChart,
  ScatterChart,
  TitleComponent,
  TooltipComponent,
  LegendComponent,
  GridComponent,
  DataZoomComponent,
])

const props = defineProps<{
  title?: string
  xData: number[]
  yData: number[]
  xLabel?: string
  yLabel?: string
  chartType?: 'line' | 'scatter'
  height?: string
  loading?: boolean
}>()

const chartRef = ref()

const option = computed(() => ({
  title: props.title
    ? { text: props.title, left: 'center', textStyle: { fontSize: 14 } }
    : undefined,
  tooltip: { trigger: 'axis' },
  grid: { left: 50, right: 20, top: props.title ? 40 : 20, bottom: 40 },
  xAxis: {
    type: 'value',
    name: props.xLabel || '',
    nameLocation: 'center',
    nameGap: 25,
  },
  yAxis: {
    type: 'value',
    name: props.yLabel || '',
  },
  dataZoom: [{ type: 'inside' }, { type: 'slider', height: 20 }],
  series: [{
    type: props.chartType || 'line',
    data: props.xData.map((x, i) => [x, props.yData[i]]),
    showSymbol: props.chartType === 'scatter',
    symbolSize: props.chartType === 'scatter' ? 4 : undefined,
    smooth: props.chartType !== 'scatter',
  }],
}))

const chartStyle = computed(() => ({
  width: '100%',
  height: props.height || '360px',
}))

onMounted(() => {
  // handled by vue-echarts
})
</script>

<template>
  <div class="live-chart">
    <v-chart
      ref="chartRef"
      :option="option"
      :style="chartStyle"
      :loading="loading"
      autoresize
    />
  </div>
</template>

<style scoped>
.live-chart {
  background: var(--color-card);
  border-radius: 8px;
  border: 1px solid var(--color-border);
  overflow: hidden;
}
</style>
