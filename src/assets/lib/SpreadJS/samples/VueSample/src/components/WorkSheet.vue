<template>
  <div class="componentContainer gc-scrollbar-thin" >
    <h3>GC-Worksheet</h3>
    <p>以下示例展示如何使用表单上的属性。</p>

    <div class="spreadContainer " >
      <gc-spread-sheets
        :hostClass='hostClass'

      >
        <gc-worksheet
          :dataSource="dataSource"
          :rowHeaderVisible = 'rowHeaderVisible'
          :columnHeaderVisible = 'columnHeaderVisible'
          :frozenRowCount = 'frozenRowCount'
          :frozenColumnCount = 'frozenColumnCount'
          :frozenTrailingRowCount = 'frozenTrailingRowCount'
          :frozenTrailingColumnCount = 'frozenTrailingColumnCount'
          :sheetTabColor = 'sheetTabColor'
          :frozenlineColor = 'frozenlineColor'
          :selectionBackColor = 'selectionBackColor'
          :selectionBorderColor = 'selectionBorderColor'
        >
          <gc-column :width="'150'" :dataField="'name'"></gc-column>
          <gc-column :width="'150'" :dataField="'phone'"></gc-column>
          <gc-column :width="'150'" :dataField="'country'"></gc-column>
          <gc-column :width="'150'" :dataField="'email'"></gc-column>
          <gc-column :width="'150'" :dataField="'onJob'"></gc-column>
        </gc-worksheet>

      </gc-spread-sheets>
    </div>
    <div class="test-btn-list">
        <label >
          <input type="checkbox" :checked="rowHeaderVisible" @click="rowHeaderVisible = !rowHeaderVisible"/>显示行头
        </label>
        <label>
          <input type="checkbox" :checked="columnHeaderVisible" @click="columnHeaderVisible = !columnHeaderVisible"/>显示列头
        </label>
      <br>
        <label>
          <input type="number" v-model="frozenRowCount" />冻结行数
        </label>
        <label>
          <input type="number" v-model="frozenColumnCount" />冻结列数
        </label>
      <br>
        <label>
          <input type="number" v-model="frozenTrailingRowCount" />末尾冻结行数
        </label>
        <label>
          <input type="number" v-model="frozenTrailingColumnCount" />末尾冻结列数
        </label>
      <br>
        <label>
          <input type="color"  v-bind:value="sheetTabColor" v-on:change="setSheetTabColor"/>标签色
        </label>
        <label>
          <input type="color"  v-bind:value="frozenlineColor" v-on:change="setFrozenlineColor"/>冻结线色
        </label>
      <br>
        <label>
          <input type="color"  v-bind:value="selectionBorderColor" v-on:change="setSelectionBorderColor"/>选择边框色
        </label>
        <label>
          <input type="color" v-bind:value="selectionBackColor"  v-on:change="setSelectionBackColor" rgba/>选择背景色
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
        hostClass:"spread-host",
        dataSource :DataService.getEmployeesData(),
        rowHeaderVisible :true,
        columnHeaderVisible : true,
        frozenRowCount: 2,
        frozenColumnCount:3,
        frozenTrailingRowCount:0,
        frozenTrailingColumnCount:0,
        sheetTabColor:"#61E6E6",
        frozenlineColor:"#000000",
        selectionBackColor:'#D0D0D0',
        selectionBorderColor:'#217346'
      }
    },
    methods:{
      setSheetTabColor(e){
        this.sheetTabColor = e.target.value;
      },
      setFrozenlineColor(e){
        this.frozenlineColor = e.target.value;
      },
      setSelectionBackColor(e){
        this.selectionBackColor = e.target.value;
      },
      setSelectionBorderColor(e){
        this.selectionBorderColor = e.target.value;
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
    height:150px;
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
