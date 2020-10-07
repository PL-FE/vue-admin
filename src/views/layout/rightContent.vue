<template>
  <div class="rightContent">
    <el-tabs v-model="editableTabsValue"
      type="card"
      editable
      @tab-click="tabClick"
      @edit="handleTabsEdit">
      <el-tab-pane :key="item.name"
        v-for="(item) in editableTabs"
        :label="item.title"
        :name="item.name">
        <router-view :name="item.name"></router-view>
      </el-tab-pane>
    </el-tabs>
  </div>
</template>

<script>
export default {
  data () {
    return {
      editableTabsValue: '',
      tabIndex: 1,
      editableTabs: []
    }
  },
  watch: {
    $route: {
      immediate: true,
      handler (route) {
        const name = route.name
        if (!this.isIncludesTab(name)) {
          this.editableTabs.push({
            title: name,
            name: name
          })
        }
        this.editableTabsValue = name
      }
    }
  },
  methods: {
    isIncludesTab (name) {
      return this.editableTabs.findIndex(it => it.name === name) !== -1
    },

    handleTabsEdit (targetName, action) {
      if (action === 'add') {
        const newTabName = ++this.tabIndex + ''
        this.editableTabs.push({
          title: 'New Tab',
          name: newTabName,
          content: 'New Tab content'
        })
        this.editableTabsValue = newTabName
      }
      if (action === 'remove') {
        const tabs = this.editableTabs
        let activeName = this.editableTabsValue
        if (activeName === targetName) {
          tabs.forEach((tab, index) => {
            if (tab.name === targetName) {
              const nextTab = tabs[index + 1] || tabs[index - 1]
              if (nextTab) {
                activeName = nextTab.name
              }
            }
          })
        }

        this.editableTabsValue = activeName
        this.editableTabs = tabs.filter(tab => tab.name !== targetName)
      }
    },

    tabClick (tab) {
    }
  }
}
</script>

<style lang="less" scoped>
.rightContent {
  overflow: hidden;
}
</style>
