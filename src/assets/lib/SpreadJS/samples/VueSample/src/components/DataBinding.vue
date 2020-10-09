<template>
  <div class="componentContainer gc-scrollbar" >
    <h3>数据绑定</h3>
    <p>以下示例展示如何绑定数据。</p>

    <div class="spreadContainer" >
      <gc-spread-sheets
        :hostClass='"spread-host"'
      >
        <gc-worksheet
          :dataSource="dataSource"
          :name = "'ALLData'"
        >
          <gc-column :dataField="'name'"></gc-column>
          <gc-column :dataField="'code'"></gc-column>
          <gc-column :dataField="'city'"></gc-column>
          <gc-column :dataField="'state'"></gc-column>
          <gc-column :dataField="'lat'"></gc-column>
          <gc-column :dataField="'lon'"></gc-column>
          <gc-column :dataField="'pop2011'"></gc-column>
          <gc-column :dataField="'vol2011'"></gc-column>
          <gc-column :dataField="'vol2010'"></gc-column>
          <gc-column :dataField="'vol2009'"></gc-column>
          <gc-column :dataField="'vol2008'"></gc-column>
          <gc-column :dataField="'vol2007'"></gc-column>
          <gc-column :dataField="'vol2006'"></gc-column>
          <gc-column :dataField="'vol2005'"></gc-column>

        </gc-worksheet>

        <gc-worksheet
          :dataSource="dataSource"
          :name = "'PartData'"
        >
          <gc-column :dataField="'name'"></gc-column>
          <gc-column :dataField="'city'"></gc-column>
          <gc-column :dataField="'state'"></gc-column>
          <gc-column :dataField="'lat'"></gc-column>
          <gc-column :dataField="'lon'"></gc-column>
          <gc-column :dataField="'vol2011'"></gc-column>
        </gc-worksheet>

      </gc-spread-sheets>
    </div>
    <div class="test-btn-list">


    </div>

  </div>
</template>
<script>
  import  '@grapecity/spread-sheets-vue'
  import DataService from '../static/dataService'

  export default {
    data(){
      return {
        dataSource:DataService.getAirpotsData(),
      };
    }
  };
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
