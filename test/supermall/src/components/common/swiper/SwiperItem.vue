// tslint:disable
<template>
  <div class="slide">
    <slot></slot>
  </div>
</template>

<script lang="ts">
  // tslint:enable
  import { Component, Prop, Vue } from 'vue-property-decorator';
  @Component
  export default class SwiperItem extends Vue {
    name: string = "Slide"
  }

  // tslint:disable
</script>

<style scoped>
  .slide {
    width: 100%;
    flex-shrink: 0;
  }

  .slide img {
    width: 100%;
  }
</style>
