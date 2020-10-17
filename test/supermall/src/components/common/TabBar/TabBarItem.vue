// tslint:disable
<template>
  <div class="tab-bar-item" @click="itemClick">
    <div v-if="!isActive"><slot name="item-icon"></slot></div>
    <div v-else><slot name="item-icon-a"></slot></div>
    <div :class="{active: isActive}"><slot name="item-text"></slot></div>
    
  </div>
</template>

<script lang="ts">
  // tslint:enable
  import { Component, Prop, Vue } from 'vue-property-decorator';
  @Component
  export default class TabBarItem extends Vue {
    @Prop(String) path!: string;
    // methods
    private itemClick () {
      this.$router.replace(this.path).catch(err => err);
    }
    // computeds
    private get isActive () : boolean {
      return this.$route.path.indexOf(this.path) !== -1;
    }
  }
  // tslint:disable  
</script>

<style scoped lang="scss">
  .tab-bar-item {
    flex: 1;
    text-align: center;
    height: 49px;
    /* line-height: 49px; */
    font-size: 14px;
  }
  .tab-bar-item img {
    width: 24px;
    height: 24px;
    margin-top: 5px;
    vertical-align: middle;
  }
  .active {
    color: red;
  }
</style>
