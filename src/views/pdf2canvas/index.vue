<template>
  <div class="file-container">
    <div class="file-upload">
      <h3>PDF 转 Canvas</h3>
      <el-upload class="file2img-upload"
        drag
        action=""
        :on-change="handleChange"
        accept=".pdf"
        :http-request="httpRequest"
        :before-upload="beforeUpload">
        <i class="el-icon-upload"></i>
        <div class="el-upload__text">将文件拖到此处，或<em>点击上传</em></div>
        <div class="el-upload__tip"
          slot="tip">支持 PDF</div>
      </el-upload>
    </div>

    <div class="img-container"
      v-show="file">
      <div>
        <div class="theCanvas">
          <canvas id="theCanvas"></canvas>
        </div>
        <el-pagination layout="prev, pager, next"
          :current-page.sync="currentPage"
          :page-size="1"
          :total="total">
        </el-pagination>
      </div>
    </div>
  </div>
</template>

<script>
const pdfjsLib = require('pdfjs-dist')
// Setting worker path to worker bundle.
window.pdfjsWorker = require('pdfjs-dist/build/pdf.worker')
export default {
  data () {
    return {
      currentPage: 1,
      total: 0,
      pageRendering: false,
      file: null
    }
  },
  watch: {
    currentPage () {
      // 渲染完才允许切换页码
      if (!this.pageRendering) {
        this.loadPdf(this.file)
      }
    }
  },
  mounted () {

  },
  methods: {
    handleChange (files, fileList) {
      if (fileList.length > 1) {
        fileList.splice(0, 1)
      }
    },
    httpRequest (file) { },

    beforeUpload (file) {
      const blob = new Blob([file])
      const reader = new FileReader()
      const that = this
      reader.onload = function () {
        that.file = this.result
        that.loadPdf(this.result)
      }
      reader.readAsArrayBuffer(blob)
    },

    async loadPdf (result) {
      const currentPage = this.currentPage
      const loadingTask = pdfjsLib.getDocument(result)
      const pdfDocument = await loadingTask.promise
      this.total = pdfDocument.numPages
      if (!pdfDocument) console.error('Error: ' + pdfDocument)
      // Request a first page
      this.pageRendering = true

      const pdfPage = await pdfDocument.getPage(currentPage)
      // Display page on the existing canvas with 100% scale.
      const viewport = pdfPage.getViewport({ scale: 1.0 - 0.2 })
      const canvas = document.getElementById('theCanvas')
      canvas.width = viewport.width
      canvas.height = viewport.height
      const ctx = canvas.getContext('2d')
      const renderTask = pdfPage.render({
        canvasContext: ctx,
        viewport: viewport
      })

      // Wait for rendering to finish
      await renderTask.promise

      this.pageRendering = false
    }

  }
}
</script>

<style lang="less" scoped>
.file-container {
  overflow: auto;
  display: flex;

  .img-container {
    flex: 1;
    display: flex;
    > div {
      margin: 0 auto;
      .theCanvas {
        display: inline-block;
        border: 1px solid #ccc;
      }
    }
  }
}
</style>
