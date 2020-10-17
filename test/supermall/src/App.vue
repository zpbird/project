// tslint:disable
<template>
  <div id="app">
    <router-view></router-view>
    <main-tab-bar></main-tab-bar>

  </div>
</template>

<script lang="ts">
  // tslint:enable
  import { Component, Vue } from 'vue-property-decorator'
  import MainTabBar from './components/content/MainTabBar/MainTabBar.vue'
  
  @Component({
    components: {
      MainTabBar,
    },
  })
  export default class App extends Vue {
    
  }
  // tslint:disable
</script>

<style lang="scss">
  @import "./assets/css/reset.css";
</style>
