<template>
  <div class="file2img">
    <h3>解析zip中的图片</h3>
    <el-upload class="file2img-upload"
      drag
      action=""
      :limit="1"
      accept=".zip"
      :http-request="httpRequest"
      :before-upload="beforeUpload">
      <i class="el-icon-upload"></i>
      <div class="el-upload__text">将文件拖到此处，或<em>点击上传</em></div>
      <div class="el-upload__tip"
        slot="tip">只能上传Zip文件，解析其中的图片</div>
    </el-upload>

    <div class="img-container"
      v-if="imgSrc.length">
      <img width="300px"
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
      imgSrc: '',
      currentPage: 1
    }
  },
  mounted () {

  },
  methods: {
    httpRequest () {

    },
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
.img-container {
  margin-top: 20px;
}
</style>
