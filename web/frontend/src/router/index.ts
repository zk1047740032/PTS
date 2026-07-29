import { createRouter, createWebHistory } from 'vue-router'

const router = createRouter({
  history: createWebHistory(),
  routes: [
    {
      path: '/',
      name: 'Dashboard',
      component: () => import('@/views/Dashboard.vue'),
    },
    {
      path: '/test/:moduleKey',
      name: 'TestDetail',
      component: () => import('@/views/TestDetail.vue'),
      props: true,
    },
    {
      path: '/history',
      name: 'TestHistory',
      component: () => import('@/views/TestHistory.vue'),
    },
    {
      path: '/config',
      name: 'ConfigView',
      component: () => import('@/views/ConfigView.vue'),
    },
  ],
})

export default router
