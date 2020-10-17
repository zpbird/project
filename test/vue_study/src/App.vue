// tslint:disable
<template>
  <div id="app">
    <router-view></router-view>
    <main-tab-bar></main-tab-bar>
  </div>
</template>

<script lang="ts">
  // tslint:enable
  import axios from 'axios'
  import { Component, Vue } from 'vue-property-decorator';
  import MainTabBar from './components/MainTabBar/MainTabBar.vue'

  axios({
    url: 'http://127.0.0.1/data/multidata.json'
  }).then(ref=>{
    console.log(ref)
  })
  @Component({
    components:{
      MainTabBar,
    },
  })
  export default class App extends Vue {}
// tslint:disable
</script>

<style lang="scss">
  @import "./assets/css/base.scss";
</style>






