<template>
  <div class="layout">
    <Header @changeCollapse="changeCollapse"></Header>
    <div class="content">
      <LeftMenu ref="LeftMenu"
        class="left"></LeftMenu>
      <div class="right">
        <RightContent></RightContent>
      </div>
    </div>
  </div>
</template>

<script>
import Header from './header'
import LeftMenu from './leftMenu'
import RightContent from './rightContent'
export default {
  components: {
    Header,
    LeftMenu,
    RightContent
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
    display: flex;
    height: calc(100% - 65px);
    overflow: auto;

    .right {
      padding: 0 15px;
      width: calc(100% - 231px);
      overflow: hidden;
    }
  }
}
</style>
