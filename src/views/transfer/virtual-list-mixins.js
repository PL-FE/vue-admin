export default {
  computed: {
    contentHeight () {
      // 计算滚动条高度
      return this.data.length * this.itemHeight
    }
  },

  watch: {
    filteredData () {
      this.handleScroll()
    }
  },

  mounted () {
    this.update()
  },

  data () {
    return {
      itemHeight: 30, // 单个高度
      virtualList: [] // 渲染数据
    }
  },

  methods: {
    update (scrollTop = 0) {
      // 获取当前可展示数量
      const count = Math.ceil(this.$el.clientHeight / this.itemHeight)
      // 取得可见区域的起始数据索引
      const start = Math.floor(scrollTop / this.itemHeight)
      // 取得可见区域的结束数据索引
      const end = start + count

      // 计算出可见区域对应的数据，让 Vue.js 更新
      this.virtualList = this.filteredData.slice(start, end)

      // 把可见区域的 top 设置为起始元素在整个列表中的位置（使用 transform 是为了更好的性能）
      this.$refs.content.style.webkitTransform = `translate3d(0, ${start * this.itemHeight}px, 0)`
    },
    handleScroll (e) {
      // 获取当前滚动条滚动位置
      const scrollTop = this.$refs.container.scrollTop
      this.update(scrollTop)
    }
  }
}
