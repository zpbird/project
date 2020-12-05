// tslint:disable
<template>
  <div class="tab-control">
    <div v-for="(item, index) in titles" class="tab-control-item" :class="{active: index === currentIndex}" @click="itemClick(index)">
      <span>{{item}}</span>
    </div>
  </div>
</template>

<script lang="ts">
  // tslint:enable  
  import { Component, Prop, Vue, Watch } from 'vue-property-decorator';
  @Component({
    components: {

    },
  })
  export default class TabControl extends Vue {
    @Prop({type: Array}) titles!: string[]
    currentIndex: number = 0

    itemClick(index: number): void {
      this.currentIndex = index
    }
    
    }
  // tslint:disable 
</script>

<style scoped lang="scss">
  .tab-control {
    display: flex;
    text-align: center;
    font-size: 16px;
    background-color: #fff;
  }
  .tab-control-item {
    flex: 1;
    height: 40px;
    line-height: 40px;
  }
  .active {
    color: red;
    border-bottom: solid 1px red;
  }
</style>
