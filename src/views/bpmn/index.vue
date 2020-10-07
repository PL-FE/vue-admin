<template>
  <div class="bpmn">
    <div class="canvas"
      ref="canvas"></div>
  </div>
</template>

<script>
// 引入相关的依赖
import BpmnModeler from 'bpmn-js/lib/Modeler'
import { xmlStr } from './xmlData' // 这里是直接引用了xml字符串
export default {
  name: '',
  components: {},
  // 生命周期 - 创建完成（可以访问当前this实例）
  created () { },
  // 生命周期 - 载入后, Vue 实例挂载到实际的 DOM 操作完成，一般在该过程进行 Ajax 交互
  mounted () {
    this.init()
  },
  data () {
    return {
      // bpmn建模器
      bpmnModeler: null,
      container: null,
      canvas: null
    }
  },
  methods: {
    init () {
      // 获取到属性ref为“canvas”的dom节点
      const canvas = this.$refs.canvas
      // 建模
      this.bpmnModeler = new BpmnModeler({
        container: canvas
      })
      this.createNewDiagram()
    },
    createNewDiagram () {
      // 将字符串转换成图显示出来
      this.bpmnModeler.importXML(xmlStr).then(res => {
        this.bpmnModeler.get('canvas').zoom('fit-viewport', 'auto')
        this.success()
      })
    },
    success () {
      // console.log('创建成功!')
    }
  }
}
</script>

<style lang="less" scoped>
.bpmn {
  width: 100%;
  height: 100%;

  /deep/.djs-container {
    background-image: linear-gradient(
        90deg,
        rgba(200, 200, 200, 0.15) 10%,
        rgba(0, 0, 0, 0) 10%
      ),
      linear-gradient(rgba(200, 200, 200, 0.15) 10%, rgba(0, 0, 0, 0) 10%);
    background-size: 10px 10px;
  }

  .canvas {
    width: 100%;
    height: 100%;
  }

  .panel {
    position: absolute;
    right: 0;
    top: 0;
    width: 300px;
  }
}
</style>
