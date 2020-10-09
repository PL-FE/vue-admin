<template>
  <div class="SpreadJS">
    <div id="excelView"
      class="excelView"></div>
  </div>
</template>

<script>
export default {
  data () {
    return {

    }
  },
  mounted () {
    this.spread = new GC.Spread.Sheets.Workbook(document.getElementById('excelView'), { sheetCount: 1 })
    this.excelIo = new GC.Spread.Excel.IO()
  },
  methods: {

  }
}
</script>

<style lang="less" scoped>
.SpreadJS {
  height: 100%;
  .excelView {
    height: 100%;
  }
}
</style>
