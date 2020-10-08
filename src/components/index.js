const requireComponents = require.context('./', false, /\.vue/)
export default function install (Vue) {
  requireComponents.keys().forEach(fileName => {
    // 组件实例
    const reqCom = requireComponents(fileName)
    // 截取路径作为组件名
    const reqComName = fileName.replace(/\.\//, '').replace(/\.vue/, '')
    // 组件挂载
    Vue.component(reqComName, reqCom.default || reqCom)
  })
}
