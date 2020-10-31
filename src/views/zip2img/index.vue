<template>
  <div class="file-container">
    <div class="file-upload">
      <h3>ZIP 解析</h3>
      <el-upload class="file2img-upload"
        drag
        action=""
        :on-change="handleChange"
        accept=".zip"
        :http-request="httpRequest"
        :before-upload="beforeUpload">
        <i class="el-icon-upload"></i>
        <div class="el-upload__text">将文件拖到此处，或<em>点击上传</em></div>
        <div class="el-upload__tip"
          slot="tip">支持解析ZIP(内容只能是图片)</div>
      </el-upload>
    </div>

    <div class="img-container"
      v-show="imgSrc.length">
      <img width="50%"
        :src="imgSrc[currentPage - 1]"
        alt="">
      <el-pagination layout="prev, pager, next"
        :current-page.sync="currentPage"
        :page-size="1"
        :total="imgSrc.length">
      </el-pagination>
    </div>
  </div>
</template>

<script>
import JSZip from 'jszip'

export default {
  data () {
    return {
      currentPage: 1,
      imgSrc: []
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

    async beforeUpload (file) {
      const newZip = new JSZip()
      const zip = await newZip.loadAsync(file)
      this.pushImg(zip)
    },

    async pushImg (zip) {
      const imgSrc = []
      const files = zip.files
      for (const key in files) {
        const name = files[key].name
        const res = await zip.file(name).async('blob')
        const blob = new Blob([res])
        const url = window.URL.createObjectURL(blob)
        imgSrc.push(url)
      }

      this.imgSrc = imgSrc
    }

  }
}
</script>

<style lang="less" scoped>
.file-container {
  overflow: auto;
  display: flex;

  .img-container {
    text-align: center;
    margin-top: 20px;
    flex: 1;
  }
}
</style>
