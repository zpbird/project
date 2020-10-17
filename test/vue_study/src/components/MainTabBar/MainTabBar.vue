// tslint:disable
<template>
  <tab-bar>
    <tab-bar-item path="/home">
      <img slot="item-icon" src="../../assets/icons/home.svg" alt="" />
      <img slot="item-icon-a" src="../../assets/icons/home-a.svg" alt="" />
      <div slot="item-text">首页</div>
    </tab-bar-item>
    <tab-bar-item path="/category">
      <img slot="item-icon" src="../../assets/icons/list.svg" alt="" />
      <img slot="item-icon-a" src="../../assets/icons/list-a.svg" alt="" />
      <div slot="item-text">分类</div>
    </tab-bar-item>
    <tab-bar-item path="/cart">
      <img slot="item-icon" src="../../assets/icons/shopping-cart.svg" alt="" />
      <img slot="item-icon-a" src="../../assets/icons/shopping-cart-a.svg" alt="" />
      <div slot="item-text">购物车</div>
    </tab-bar-item>
    <tab-bar-item path="/profile">
      <img slot="item-icon" src="../../assets/icons/user.svg" alt="" />
      <img slot="item-icon-a" src="../../assets/icons/user-a.svg" alt="" />
      <div slot="item-text">我</div>
    </tab-bar-item>
  </tab-bar>
</template>

<script lang="ts">
  // tslint:enable
  import { Component, Prop, Vue, Watch } from 'vue-property-decorator';
  import TabBar from '../../components/TabBar/TabBar.vue';
  import TabBarItem from '../../components/TabBar/TabBarItem.vue';
  @Component({
    components:{
      TabBar,
      TabBarItem,
    },
  })
  export default class MainTabBar extends Vue {
  }
  // tslint:disable
</script>

<style scoped lang="scss">
  
</style>
