<template>
  <div class="componentContainer gc-scrollbar" >
    <h3>分组</h3>
    <p>以下示例展示如何使用分组。</p>

    <div class="spreadContainer" >
      <gc-spread-sheets
        :hostClass='"spread-host"'
      >
        <gc-worksheet
          :dataSource="dataSource"
          :name = "'Outline'"
          :showColumnOutline = 'showColumnOutline'
          :showRowOutline = 'showRowOutline'
          :autoGenerateColumns = 'autoGenerateColumns'
          :rowOutlineInfo = 'rowOutlineInfo'
          :columnOutlineInfo = 'columnOutlineInfo'

        >
          <gc-column :width="'120'" :dataField="'Name'"  ></gc-column>
          <gc-column :width="'120'" :dataField="'CountryRegionCode'"></gc-column>
          <gc-column :width="'120'" :dataField="'City'"  ></gc-column>
          <gc-column :width="'120'" :dataField="'AddressLine'"  ></gc-column>
          <gc-column :width="'120'" :dataField="'PostalCode'" ></gc-column>

        </gc-worksheet>

      </gc-spread-sheets>
    </div>
    <div class="test-btn-list">
      <label >
        <input type="checkbox" :checked="showRowOutline" @click="showRowOutline = !showRowOutline"/>显示行分组
      </label>
      <label >
        <input type="checkbox" :checked="showColumnOutline" @click="showColumnOutline = !showColumnOutline"/>显示列分组
      </label>
    </div>

  </div>
</template>
<script>
  import  '@grapecity/spread-sheets-vue'
  import GC from '@grapecity/spread-sheets'
  import DataService from '../static/dataService'

  export default {
    data(){
      return {
        dataSource:DataService.getPersonAddressData(),
        showRowOutline : true,
        showColumnOutline: true,
        autoGenerateColumns : false
      }
    },
    computed:{
      rowOutlineInfo(){
        return [{index: 1, count: 4}, {index: 6, count: 4}]
      },
      columnOutlineInfo(){
        return [{index: 0, count: 4}];
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
    margin: 10px 20px;
  }
  .spread-host{
    width: 100%;
    height: 100%;
  }

</style>
