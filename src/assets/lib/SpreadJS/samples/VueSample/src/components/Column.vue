<template>
  <div class="componentContainer gc-scrollbar-thin" >
    <h3>GC-Column</h3>
    <p>以下示例展示如何绑定列上的属性。</p>

    <div class="spreadContainer" >
      <gc-spread-sheets
        :hostClass='hostClass'
      >
        <gc-worksheet
          :dataSource="dataTable"
          :autoGenerateColumns = 'autoGenerateColumns'
        >
          <gc-column
            :width="width"
            :dataField="'price'"
            :visible = 'visible'
            :formatter = 'formatter'
            :resizable = 'resizable'
          ></gc-column>

        </gc-worksheet>

      </gc-spread-sheets>
    </div>
    <div class="test-btn-list">
      <label >
        <input type="checkbox" :checked="visible" @click="visible = !visible"/>可见
      </label>
      <label >
        <input type="checkbox" :checked="resizable" @click="resizable = !resizable"/>可改变列宽
      </label>
      <br>
      <label>
        <input type="number" v-model="width" />列宽
      </label>
      <label>
        <input type="text" v-model="formatter" />数据格式
      </label>


    </div>

  </div>
</template>
<script>
  import  '@grapecity/spread-sheets-vue'


  export default {
    data(){
      return {
        hostClass:'spread-host',
        autoGenerateColumns:true,
        width:300,
        visible:true,
        resizable:true,
        formatter:"$ #.00"
      }
    },
    computed:{
      dataTable(){
        let dataTable = [];
        for (let i = 0; i < 42; i++) {
          dataTable.push({price: i + 0.56})
        }
        return dataTable
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
    height:120px;
  }
  .test-btn-list label{
    display: inline-flex;
    margin: 3px 20px;
    width: 293px;
  }
  .spread-host{
    width: 100%;
    height: 100%;
  }

</style>
