<template>
  <div class="componentContainer" >
    <h3>快速入门指南</h3>
    <div>
      按照以下步骤快速将 SpreadJS 应用到 Vue 的应用程序中：
      <div>
        <p>1. 添加 SpreadJS 产品引用到 Vue 的应用程序中。</p>
        <p>2. 添加一个组件来提供数据和业务实现。</p>
        <p>3. 绑定数据，并设置其他 SpreadJS 属性。</p>
        <p>4. 添加一些 css 属性来定制外观。</p>
      </div>
    </div>
    <div class="spreadContainer" >
      <gc-spread-sheets
        :hostClass='"spreadHost"'>

      </gc-spread-sheets>
    </div>
  </div>
</template>
<script>
  import  '@grapecity/spread-sheets-vue'
  export default {
//    name: 'sample-header'
  }
</script>
<style scoped>
  .componentContainer {
    position: absolute;
    padding: 10px;
    left: 242px;
    top: 0;
    bottom: 20px;
    right: 0;
  }
  .spreadContainer {
    padding: 10px;
    box-shadow: 0 0 20px grey;
  }
  .spreadContainer{
    position: absolute;
    left: 0px;
    right: 30px;
    top: 260px;
    bottom: 10px;
  }
  .spreadHost{
    width: 100%;
    height: 100%;
  }


</style>
