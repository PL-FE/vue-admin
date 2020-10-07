<template>
  <div class="leftMenu">
    <el-menu router
      class="el-menu-vertical-demo"
      :default-openeds="defaultOpeneds"
      :default-active="defaultActive"
      :collapse="isCollapse">
      <el-submenu :index="m.index"
        v-for="m in menuData"
        :key="m.index">
        <template slot="title">
          <i :class="m.icon"></i>
          <span slot="title">{{m.label}}</span>
        </template>
        <el-menu-item v-for="sub in m.children"
          :key="sub.index"
          :index="sub.index">{{sub.label}}</el-menu-item>
      </el-submenu>
    </el-menu>
  </div>
</template>

<script>
export default {
  data () {
    return {
      isCollapse: false,
      menuData: [
        {
          label: '插件',
          index: '1',
          icon: 'el-icon-location',
          children: [
            {
              index: 'Bpmn',
              label: 'Bpmn.js'
            },
            {
              index: 'ckEditor',
              label: 'ckEditor'
            }
          ]
        },
        {
          label: '组件',
          index: '2',
          icon: 'el-icon-location',
          children: []
        },
        {
          label: '404',
          index: '3',
          icon: 'el-icon-location',
          children: [
            {
              index: '404',
              label: '404'
            }
          ]
        }

      ]
    }
  },
  computed: {
    defaultOpeneds () {
      const { menuData } = this

      return menuData.map(({ index }) => index)
    },
    defaultActive () {
      const { menuData } = this
      const activeUrl = menuData.filter(({ children }) => children.length)[0].children[0].index
      return activeUrl
    }
  },
  methods: {
  }
}
</script>

<style lang="less" scoped>
.leftMenu {
  height: 100%;
  user-select: none;
}
.el-menu-vertical-demo:not(.el-menu--collapse) {
  width: 200px;
}

.el-menu-vertical-demo {
  overflow: auto;
  height: 100%;
}
</style>
