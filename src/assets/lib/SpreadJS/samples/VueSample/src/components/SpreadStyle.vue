<template>
  <div class="componentContainer gc-scrollbar" >
    <h3>样式</h3>
    <p>以下示例展示如何使用样式。</p>

    <div class="spreadContainer" >
      <gc-spread-sheets
        :hostClass='"spread-host"'
      >
        <gc-worksheet
          :dataSource="dataSource"
        >
          <gc-column :width="'120'" :dataField="'name'" :columnStyle="columnStyle" ></gc-column>
          <gc-column :width="'120'" :dataField="'phone'"></gc-column>
          <gc-column :width="'120'" :dataField="'country'" :cellType="comboBoxCellType" ></gc-column>
          <gc-column :width="'120'" :dataField="'email'" :cellType="hyperLinkCelleType" ></gc-column>
          <gc-column :width="'120'" :dataField="'onJob'" :cellType="checkBoxCellType"></gc-column>

        </gc-worksheet>

      </gc-spread-sheets>
    </div>
    <div class="test-btn-list">


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
        dataSource:DataService.getEmployeesData(),
        checkBoxCellType: new GC.Spread.Sheets.CellTypes.CheckBox(),
        hyperLinkCelleType : new GC.Spread.Sheets.CellTypes.HyperLink()
      }
    },
    computed:{
      columnStyle(){
        let style = new GC.Spread.Sheets.Style();
        style.backColor = 'lightgray';
        return style;
      },
      comboBoxCellType(){
        let  cellType = new GC.Spread.Sheets.CellTypes.ComboBox();
        cellType.items([
          {text: 'US', value: 'US'},
          {text: 'UK', value: 'UK'},
          {text: 'Germany', value: 'Germany'},
          {text: 'Maxico', value: 'Maxico'}]);
        return cellType;
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
