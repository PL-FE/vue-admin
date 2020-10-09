<template>
  <div id="app">
    <sample-header></sample-header>
    <div class="content">
      <div class="nav">
        <sample-nav></sample-nav>
      </div>
      <div>
        <router-view></router-view>
      </div>
    </div>
    <sample-footer></sample-footer>
  </div>
</template>

<script>
  import Header from './components/header.vue'
  import Nav from './components/nav.vue'
  import Footer from './components/footer.vue'
  import '@grapecity/spread-sheets/styles/gc.spread.sheets.excel2016colorful.css'
  import GC from '@grapecity/spread-sheets'
  import '@grapecity/spread-sheets-resources-zh'
GC.Spread.Common.CultureManager.culture("zh-cn");
  export default {
    name: 'app',
    components: {
      'sample-header':Header,
      'sample-nav':Nav,
      'sample-footer':Footer
    },
    data () {
      return {

      }
    }
  }
</script>

<style>

#app {
  font-size: 14px;
  margin-top: 60px;
  margin: 0 10%;
  position: absolute;
  top: 0;
  bottom: 0;
  left: 0;
  right: 0;
  -webkit-box-shadow: 0 0 20px #e3e3e3;
  box-shadow: 0 0 20px #e3e3e3;
}
  .content{
    position: absolute;
    top: 80px;
    bottom: 60px;
    left: 0;
    right: 0;
    display: flex;
  }
  .nav{
    position: absolute;
    left: 0;
    top: 0;
    bottom: 20px;
    padding: 0 10px;
    width: 200px;
    border-right: 2px solid #e3e3e3;
  }
.gc-scrollbar::-webkit-scrollbar-track {
  /*-webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.3);*/
  background-color: #efeff0;
  -webkit-border-radius: 10px;
  border-radius: 10px;
}
.gc-scrollbar::-webkit-scrollbar-button:decrement {
  /*-webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.3);*/
  height: 8px;
  /*background:url("../img/triangle-up.png") no-repeat center;*/
  background-size:10px 5px;
  background-color: #ffffff;
  /*background-image: url("data:image/svg+xml;charset=utf-8,<svg viewBox='0 0 10 10' xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'><defs><line id='l' x1='60' x2='60' y1='7' y2='27' stroke='%236c6c6c' stroke-width='11' stroke-linecap='round'/></defs><g><use xlink:href='%23l' opacity='.27'/><use xlink:href='%23l' opacity='.27' transform='rotate(30 60,60)'/></g></svg>");*/
}
.gc-scrollbar::-webkit-scrollbar-button:increment {
  /*-webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.3);*/
  height: 8px;
  /*background:url("../img/triangle-down.png") no-repeat center;*/
  background-size:10px 5px;
  background-color: #ffffff;
  /*background-image: url("data:image/svg+xml;charset=utf-8,<svg viewBox='0 0 10 10' xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'><defs><line id='l' x1='60' x2='60' y1='7' y2='27' stroke='%236c6c6c' stroke-width='11' stroke-linecap='round'/></defs><g><use xlink:href='%23l' opacity='.27'/><use xlink:href='%23l' opacity='.27' transform='rotate(30 60,60)'/></g></svg>");*/
}

/* Handle */
.gc-scrollbar::-webkit-scrollbar-thumb {
  -webkit-border-radius: 4px;
  border-radius: 4px;
  background-color: #a7a7a7;
  /*background:rgba(153,153,153,0.7);*/
  /*-webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.5);*/
}
.gc-scrollbar::-webkit-scrollbar-thumb:window-inactive {
  background: rgba(153,153,153,0.4);
}
.gc-scrollbar-thin::-webkit-scrollbar{
  width: 4px;
}
.gc-scrollbar-thin::-webkit-scrollbar-thumb {
  -webkit-border-radius: 1px;
  border-radius: 1px;
  background-color: #666666;
  /*background:rgba(153,153,153,0.7);*/
  /*-webkit-box-shadow: inset 0 0 6px rgba(0,0,0,0.5);*/
}
.gc-scrollbar-thin::-webkit-scrollbar-button:decrement,
.gc-scrollbar-thin::-webkit-scrollbar-button:increment
{
  height: 0;
  width: 0;
}
.logo {
  position: absolute;
  top: 50%;
  margin-top: -20px;
  width: 40px;
  height: 40px;
  left: 80px;
}

img {
  max-width: 100%;
  width: auto;
  height: auto;
  vertical-align: middle;
  border: 0;
  -ms-interpolation-mode: bicubic;
}

body {
  font-family: 'Gotham SSm A', 'Gotham SSm B';
  font-size: 14px;
  color: #333333;
  background-color: #ffffff;
}

h3 {
  font-size: 24px;
}

h2, h3, h4, h5, h6 {
  margin: 10px 0;
  font-family: 'Forza SSm A', 'Forza SSm B';
  font-weight: bold;
  line-height: 20px;
  color: inherit;
  text-rendering: optimizelegibility;
}

.test-btn-list label input[type=checkbox] {
  margin: 5px;
}

.test-btn-list label input[type=text],
.test-btn-list label input[type=color],
.test-btn-list label input[type=number]
{
  margin: 0 5px;
}

label {
  font-weight: normal;
}
</style>
