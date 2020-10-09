<template>
  <div class="componentContainer gc-scrollbar-thin" >
    <h3>GC-Spread-Sheets</h3>
    <p>以下示例展示如何绑定 Workbook 上的属性。</p>

    <div class="spreadContainer" >
      <gc-spread-sheets
        :newTabVisible='newTabVisible'
        :tabStripVisible ='tabStripVisible'
        :tabStripRatio='tabStripRatio'
        :tabNavigationVisible='tabNavigationVisible'
        :scrollbarShowMax='scrollbarShowMax'
        :scrollbarMaxAlign='scrollbarMaxAlign'
        :showHorizontalScrollbar = 'showHorizontalScrollbar'
        :showVerticalScrollbar = 'showVerticalScrollbar'
        :allowUserZoom = 'allowUserZoom'
        :allowUserResize = 'allowUserResize'
        :grayAreaBackColor = 'grayAreaBackColor'
        :backColor = 'backColor'
        :hostClass='hostClass'

      >
        <gc-worksheet   :dataSource="dataSource"  >
          <gc-column :width="'150'" :dataField="'Name'"></gc-column>
          <gc-column :width="'150'" :dataField="'CountryRegionCode'"></gc-column>
          <gc-column :width="'150'" :dataField="'City'"></gc-column>
          <gc-column :width="'150'" :dataField="'AddressLine'"></gc-column>
          <gc-column :width="'150'" :dataField="'PostalCode'"></gc-column>
        </gc-worksheet>

      </gc-spread-sheets>
    </div>
    <div class="test-btn-list">
      <label >
        <input type="checkbox" :checked="newTabVisible" @click="newTabVisible = !newTabVisible"/>显示 new 标签页
      </label>
      <label>
        <input type="checkbox" :checked="tabStripVisible" @click="tabStripVisible = !tabStripVisible"/>显示标签页
      </label>
      <br>
      <label>
        <input type="checkbox" :checked="tabStripRatio" @click="tabStripRatio = !tabStripRatio">水平滚动条宽度比率
      </label>
      <label>
        <input type="checkbox" :checked="tabNavigationVisible" @click="tabNavigationVisible = !tabNavigationVisible">显示或隐藏表单标签导航项
      </label>
      <br>
      <label>
        <input type="checkbox" :checked="scrollbarShowMax" @click="scrollbarShowMax = !scrollbarShowMax">基于表单全部的行列总数显示滚动条
      </label>
      <label>
        <input type="checkbox" :checked="scrollbarMaxAlign" @click="scrollbarMaxAlign = !scrollbarMaxAlign">滚动条对齐视图中表单的最后一行或一列
      </label>
      <br>
      <label>
        <input type="checkbox" :checked="showHorizontalScrollbar" @click="showHorizontalScrollbar = !showHorizontalScrollbar"/>显示横向滚动条
      </label>
      <label>
        <input type="checkbox" :checked="showVerticalScrollbar" @click="showVerticalScrollbar = !showVerticalScrollbar"/>显示竖向滚动条
      </label>
      <br>
      <label>
        <input type="checkbox" :checked="allowUserZoom" @click="allowUserZoom = !allowUserZoom"/>允许用户缩放
      </label>
      <label>
        <input type="checkbox" :checked="allowUserResize" @click="allowUserResize = !allowUserResize"/>允许用户改变行列宽高
      </label>
      <br>
      <label>
        <input type="color"  v-bind:value="grayAreaBackColor" v-on:change="setAreaBackColor"/>灰色区域背景色
      </label>
      <label>
        <input type="color" v-bind:value="backColor" v-on:change="setBackColor"/>背景色
      </label>
    </div>

  </div>
</template>
<script>
  import  '@grapecity/spread-sheets-vue'
  import DataService from '../static/dataService'

  export default {
    data(){
      return {
        newTabVisible :true,
        tabStripVisible : true,
        tabStripRatio: true,
        tabNavigationVisible: true,
        scrollbarShowMax: true,
        scrollbarMaxAlign: true,
        showHorizontalScrollbar: true,
        showVerticalScrollbar:true,
        allowUserZoom:true,
        allowUserResize:true,
        grayAreaBackColor:"#E4E4E4",
        backColor:"#ffffff",
        dataSource: DataService.getPersonAddressData(),
        hostClass : 'spread-host'
      }
    },
    methods:{
      setAreaBackColor(e){
        this.grayAreaBackColor = e.target.value;
      },
      setBackColor(e){
        this.backColor = e.target.value;
      }

    }


  }
</script>
<style scoped>
  .componentContainer {
    position: absolute;
    padding: 10px;
    left: 242px;
    top: 0;
    bottom: 20px;
    right: 3px;
    overflow-y:auto ;
    overflow-x: hidden;
  }
  .spreadContainer{
    position: absolute;
    top:90px;
    padding: 10px;
    /*width: 100%;*/
    /*height: 240px;*/
    left: 10px;
    right:10px;
    bottom: 150px;
    box-shadow: 0 0 20px grey;
  }
  .test-btn-list{
    /*padding: 20px;*/
    position: absolute;
    bottom: 0px;
    height:130px;
  }
  .test-btn-list label{
    display: inline-flex;
    margin: 3px 20px;
    width: 293px;
    line-height: 23px;
  }
  .spread-host{
    width: 100%;
    height: 100%;
  }

</style>
