<template>
  <div class="header">
    <span class="icon"
      @click="handleCollapse">
      {{ isCollapse ? '展开' : '关闭'}}
      <i :class="isCollapse ? 'el-icon-caret-left': 'el-icon-caret-right'">
      </i>
    </span>
    <span class="userInfo">Admin</span>
  </div>
</template>

<script>
export default {
  data () {
    return {
      isCollapse: false
    }
  },
  watch: {
    isCollapse (v) {
      this.$emit('changeCollapse')
    }
  },
  mounted () {

  },
  methods: {
    handleCollapse () {
      this.isCollapse = !this.isCollapse
    }
  }
}
</script>

<style lang="less" scoped>
.header {
  display: flex;
  align-items: center;
  height: 65px;
  width: 100%;
  padding: 0 10px;
  background-color: #001529;
  color: #fff;

  .icon {
    padding: 10px;
    cursor: pointer;
    user-select: none;
  }

  .userInfo {
    margin-left: auto;
    margin-right: 40px;
  }
}
</style>
