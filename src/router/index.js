import Vue from 'vue'
import VueRouter from 'vue-router'
import Project from '@/views/project'
import Bpmn from '@/views/bpmn'
import ckEditor from '@/views/ck-editor'
import NotFind from '@/components/NotFind.vue'

Vue.use(VueRouter)

const routes = [
  {
    path: '/',

    name: 'Project',
    component: Project,
    children: [
      {
        path: 'Bpmn',
        name: 'Bpmn',
        components: {
          Bpmn
        }
      },
      {
        path: 'ckEditor',
        name: 'ckEditor',
        components: {
          ckEditor
        }
      },
      {
        path: '/404',
        name: '404',
        components: {
          404: NotFind
        }
      },
      {
        path: '*',
        redirect: '/Bpmn'
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
