<template>
  <div class="navigationBar">
    <h2>SpreadJS</h2>
    <nav>
      <ul>
        <li><router-link exact to="/QuickStart">快速入门指南</router-link></li>
        <li><router-link exact to="/GC-Spread-Sheets">GC-Spread-Sheets</router-link></li>
        <li><router-link exact to="/WorkSheet">GC-Worksheet</router-link></li>
        <li><router-link exact to="/Column">GC-Column</router-link></li>
        <li><router-link exact to="/DataBinding">数据绑定</router-link></li>
        <li><router-link exact to="/SpreadStyle">样式</router-link></li>
        <li><router-link exact to="/OutLine">分组</router-link></li>

      </ul>
    </nav>
  </div>
</template>
<script>
  export default {
//    name: 'sample-header'
  }
</script>
<style>
  .navigationBar {
    position: absolute;
    left: 0;
    top: 0;
    bottom: 20px;
    padding: 0 10px;
    width: 200px;
    border-right: 2px solid #e3e3e3;
  }
  .navigationBar li {
    margin-bottom: 10px;
  }
  .navigationBar li a{
    margin-top: 0;
    margin-bottom: 10px;
    cursor: pointer;
    text-decoration: none;
    color: rgb(96, 125, 139);
  }
  .navigationBar li:hover{
    background-color: #CFD8DC;
  }
  .navigationBar li:hover a{
    color: rgb(3, 155, 229);
  }
  .navigationBar li a.router-link-exact-active,
  .navigationBar li a.router-link-active {
    color: rgb(3, 155, 229);
  }

</style>
