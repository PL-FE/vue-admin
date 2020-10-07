import Vue from 'vue'
import VueRouter from 'vue-router'
import Projet from '@/views/project'
import Bpmn from '@/views/bpmn'
import NotFind from '@/components/NotFind.vue'

Vue.use(VueRouter)

const routes = [
  {
    path: '/',
    name: 'Projet',
    component: Projet,
    children: [
      {
        path: '/Bpmn',
        name: 'Bpmn',
        component: Bpmn
      },
      {
        path: '*',
        name: '404',
        component: NotFind
      }
    ]
  }

  // {
  //   path: '/about',
  //   name: 'About',
  //   // route level code-splitting
  //   // this generates a separate chunk (about.[hash].js) for this route
  //   // which is lazy-loaded when the route is visited.
  //   component: () =>
  //     import(/* webpackChunkName: "about" */ '../views/About.vue')
  // }
]

const router = new VueRouter({
  mode: 'history',
  base: process.env.BASE_URL,
  routes
})

export default router
