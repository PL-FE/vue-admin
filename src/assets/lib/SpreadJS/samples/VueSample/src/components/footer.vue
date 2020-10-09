<template>
  <div class="app-footer"><span class="footer-text">Copyright© GrapeCity, inc. All Rights Reserved.</span></div>
</template>

<script>
  export default {
//    name: 'footer',
  }
</script>

<style>
  .app-footer {
    position: absolute;
    text-align: center;
    background-color: black;
    color: white;
    height: 60px;
    bottom: 0;
    left: 0;
    right: 0;
  }
  .footer-text {
    position: absolute;
    left: 0;
    right: 0;
    top: 50%;
    margin-top: -8px;
    height: 16px;
    font-size: 16px;
    font-family: 'Forza SSm A', 'Forza SSm B';
  }
</style>
