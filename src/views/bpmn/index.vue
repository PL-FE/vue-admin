<template>
  <div class="bpmn">
    <div class="tool">
      <el-button @click="saveXML">保存 XML</el-button>
      <el-button @click="$refs.refFile.click()">导入 XML</el-button>
      <el-button @click="saveSVG">保存为 SVG</el-button>

      <input type="file"
        id="files"
        ref="refFile"
        style="display: none"
        @change="loadXML" />
    </div>
    <div class="canvas"
      ref="canvas"></div>
  </div>
</template>

<script>
// 引入相关的依赖
import BpmnModeler from 'bpmn-js/lib/Modeler'
import { xmlStr } from './xmlData' // 这里是直接引用了xml字符串
export default {
  name: 'Bpmn',
  components: {},
  data () {
    return {
      bpmnModeler: null,
      container: null,
      canvas: null,
      xml: ''
    }
  },
  mounted () {
    this.init()
  },
  methods: {
    init () {
      const canvas = this.$refs.canvas
      // 建模
      this.bpmnModeler = new BpmnModeler({
        container: canvas
      })

      // 绑定事件
      const eventBus = this.bpmnModeler.get('eventBus')
      eventBus.on('element.click', e => {
        console.log('点击了元素', e)
      })

      // 导入 xml
      this.xml = xmlStr
      this.createNewDiagram()
    },
    createNewDiagram () {
      // 将字符串转换成图显示出来
      this.bpmnModeler.importXML(this.xml).then(res => {
        this.bpmnModeler.get('canvas').zoom('fit-viewport', 'auto')
        this.success()
      })
    },
    success () {
      // console.log('创建成功!')
    },

    // 获取所有元素
    getElementAll () {
      return this.bpmnModeler.get('elementRegistry').getAll()
    },
    // 根据 id 获取元素
    getElementById (id) {
      return this.bpmnModeler.get('elementRegistry').get(id)
    },

    // 查看所有可用事件
    getEventBusAll () {
      const eventBus = this.bpmnModeler.get('eventBus')
      const eventTypes = Object.keys(eventBus._listeners)
      console.log(eventTypes) // 打印出来有242种事件
      return eventTypes
    },

    async saveXML () {
      try {
        const result = await this.bpmnModeler.saveXML({ format: true })
        const { xml } = result

        const xmlBlob = new Blob([xml], {
          type: 'application/bpmn20-xml;charset=UTF-8,'
        })

        const downloadLink = document.createElement('a')
        downloadLink.download = `bpmn-${+new Date()}.bpmn`
        downloadLink.innerHTML = 'Get BPMN SVG'
        downloadLink.href = window.URL.createObjectURL(xmlBlob)
        downloadLink.onclick = function (event) {
          document.body.removeChild(event.target)
        }
        downloadLink.style.visibility = 'hidden'
        document.body.appendChild(downloadLink)
        downloadLink.click()
      } catch (err) {
        console.log(err)
      }
    },

    async saveSVG () {
      try {
        const result = await this.bpmnModeler.saveSVG()
        const { svg } = result

        const svgBlob = new Blob([svg], {
          type: 'image/svg+xml'
        })

        const downloadLink = document.createElement('a')
        downloadLink.download = `bpmn-${+new Date()}.SVG`
        downloadLink.innerHTML = 'Get BPMN SVG'
        downloadLink.href = window.URL.createObjectURL(svgBlob)
        downloadLink.onclick = function (event) {
          document.body.removeChild(event.target)
        }
        downloadLink.style.visibility = 'hidden'
        document.body.appendChild(downloadLink)
        downloadLink.click()
      } catch (err) {
        console.log(err)
      }
    },

    async loadXML () {
      const that = this
      const file = this.$refs.refFile.files[0]

      const reader = new FileReader()
      reader.readAsText(file)
      reader.onload = function () {
        console.log('this', this)
        that.xmlStr = this.result
        that.createNewDiagram()
      }
    }

  }
}
</script>

<style lang="less" scoped>
.bpmn {
  width: 100%;
  height: 100%;
  position: relative;

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

  .tool {
    position: absolute;
    z-index: 1;
    left: 50%;
    bottom: 20px;
    transform: translateX(-50%);
  }
}
</style>
