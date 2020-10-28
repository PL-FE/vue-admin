<template>
  <div class="rightContent">
    <el-tabs v-model="DycurTab"
      type="card"
      closable
      @tab-remove="removeTab">
      <el-tab-pane :key="item.name"
        v-for="(item) in tabs"
        :label="item.title"
        :name="item.name">
        <component class="custom-component"
          :is="item.name"></component>
      </el-tab-pane>
    </el-tabs>
  </div>
</template>

<script>
import { mapGetters, mapMutations } from 'vuex'
import Bpmn from '@/views/bpmn'
import ckEditor from '@/views/ck-editor'
import SpreadJS from '@/views/spread-js'
import File2img from '@/views/file2img'

export default {
  components: {
    Bpmn,
    ckEditor,
    SpreadJS,
    File2img
  },
  data () {
    return {
      tabIndex: 1
    }
  },
  computed: {
    ...mapGetters(['tabs', 'curTab']),
    DycurTab: {
      get () {
        return this.curTab
      },
      set (newTab) {
        // TODO: 首次会传 ’0‘ 进来
        // 应该是第 0 个
        if (newTab && newTab !== '0') {
          this.SET_CURTAB(newTab)
        }
      }
    }
  },
  watch: {
  },
  methods: {
    ...mapMutations(['REMOVE_TABS', 'SET_CURTAB']),

    removeTab (targetName) {
      this.REMOVE_TABS(targetName)
    },

    tabClick (tab) {
    }
  }
}
</script>

<style lang="less" scoped>
.rightContent {
  overflow: hidden;
  height: 100%;
  /deep/.el-tabs {
    height: 100%;
    .el-tabs__content {
      height: calc(100% - 56px);
      .el-tab-pane {
        height: 100%;
      }
    }
  }

  .custom-component {
    height: 100%;
  }
}
</style>
