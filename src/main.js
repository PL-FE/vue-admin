import Vue from 'vue'
import App from './App.vue'
import router from './router'
import store from './store'
import ElementUI from 'element-ui'
import components from '@/components/index.js'
import CKEditor from '@ckeditor/ckeditor5-vue'
import 'element-ui/lib/theme-chalk/index.css'
import VueCodemirror from 'vue-codemirror'
import 'codemirror/lib/codemirror.css'

Vue.use(VueCodemirror)
Vue.use(CKEditor)
Vue.use(ElementUI)
Vue.use(components)
Vue.config.productionTip = false

new Vue({
  router,
  store,
  render: h => h(App)
}).$mount('#app')
