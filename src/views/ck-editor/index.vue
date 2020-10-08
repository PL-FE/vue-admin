<template>
  <div class="ck-editor">
    <ckeditor :editor="editor"
      v-model="editorData"
      @ready="onReady"
      class="editor"
      :config="editorConfig"></ckeditor>
  </div>
</template>

<script>
import ClassicEditor from '@ckeditor/ckeditor5-build-decoupled-document'

import data from './data'
export default {
  data () {
    return {
      editor: ClassicEditor,
      editorData: data,
      editorConfig: {
        language: 'de'
      }
    }
  },
  mounted () {
    // 可用的所有插件
    // const pluginNameAll = this.editor.builtinPlugins.map(plugin => plugin.pluginName)
    // console.log('pluginNameAll', pluginNameAll)
    // 可用工具栏
    // const componentFactoryAll = Array.from(this.editor.ui.componentFactory.names())
    // console.log('componentFactoryAll', componentFactoryAll)
  },
  beforeDestroy () {
    this.editor.destroy()
  },
  methods: {
    onReady (editor) {
      // Insert the toolbar before the editable area.
      editor.ui.getEditableElement().parentElement.insertBefore(
        editor.ui.view.toolbar.element,
        editor.ui.getEditableElement()
      )
    }
  }
}
</script>

<style lang="less" scoped>
.ck-editor {
  height: 100%;
  .editor {
    height: 100%;
  }
}
</style>
