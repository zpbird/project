// tslint:disable
<template>
  <div id="tab-bar">
    <slot></slot>
  </div>
</template>

<script lang="ts">
// tslint:enable
  import { Component, Prop, Vue } from 'vue-property-decorator';
  @Component
  export default class TabBar extends Vue {
  }
// tslint:disable
</script>

<style scoped lang="scss">
  #tab-bar {
    display: flex;
    background-color: #e9e9e9;
    position: fixed;
    left: 0;
    right: 0;
    bottom: 0;
  }


</style>


