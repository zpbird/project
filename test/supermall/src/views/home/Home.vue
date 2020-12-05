// tslint:disable
<template>
  <div>
      <h2>Home</h2>
      <home-swiper :banners="banners"></home-swiper>
      <recommend-view :recommends="recommends"></recommend-view>
      <feature-view></feature-view>
      <tab-control class="tab-control" :titles="['流行', '新款', '精选']"></tab-control>

      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      <div>aaa</div>
      
      
  </div>
</template>

<script lang="ts">
  // tslint:enable
  import { Component, Prop, Vue, Watch } from 'vue-property-decorator'
  import { getHomeMultidata } from '../../network/home'
  import HomeSwiper from './childComps/HomeSwiper.vue'
  import RecommendView from './childComps/RecommendView.vue'
  import FeatureView from './childComps/FeatureView.vue'
  import TabControl from '../../components/content/TabControl/TabControl.vue'
  
  @Component({
    components: {
      HomeSwiper,
      RecommendView,
      FeatureView,
      TabControl
    },
  })
  export default class Home extends Vue {
    banners: object[] = []
    recommends: object[] = []
    goods: object = {
      pop: {page: 0, list: []},
      news: {page: 0, list: []},
      sell: {page: 0, list: []}
    }
    created(): void {
      console.log("生命周期create触发")
      getHomeMultidata().then(res => {
        console.log(res)
        this.banners = res.data.banner.list
        this.recommends = res.data.recommend.list
      })
    }
  }
  // tslint:disable
</script>

<style scoped lang="scss">
  .tab-control {
    position: sticky;
    top: 0px;
  }
  
    
</style>
