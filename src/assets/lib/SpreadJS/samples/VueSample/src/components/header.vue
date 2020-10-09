<template>
  <div class="app-header">
    <img class="logo" src="../static/logo.png" alt="logo">
    <span class="header-text">SpreadJS Vue 示例</span>
  </div>
</template>
<script>
  export default {
//    name: 'sample-header'
  }
</script>
<style>
  .app-header {
    position: absolute;
    left: 0;
    right: 0;
    top: 0;
    bottom: 0;
    height: 60px;
    background-color: black;
    color: white;
    text-align: center;
  }
  .header-text {
    position: absolute;
    left: 0;
    right: 0;
    font-size: 30px;
    font-family: 'Forza SSm A', 'Forza SSm B';
    font-weight: bold;
    height: 30px;
    margin-top: -15px;
    top: 50%;
    vertical-align: middle;
  }
  .logo {
    position: absolute;
    top: 50%;
    margin-top: -20px;
    width: 40px;
    height: 40px;
    left: 80px;
  }
</style>
