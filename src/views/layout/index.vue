<template>
  <div class="layout">
    <Header @changeCollapse="changeCollapse"></Header>
    <div class="content">
      <LeftMenu ref="LeftMenu"></LeftMenu>
    </div>
  </div>
</template>

<script>
import Header from './header'
import LeftMenu from './leftMenu'
export default {
  components: {
    Header,
    LeftMenu
  },
  data () {
    return {
    }
  },
  methods: {
    changeCollapse (v) {
      this.$refs.LeftMenu.isCollapse = !this.$refs.LeftMenu.isCollapse
    }
  }
}
</script>

<style lang="less" scoped>
.layout {
  height: 100%;
  .content {
    height: calc(100% - 65px);
    overflow: auto;
  }
}
</style>
