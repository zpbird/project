// tslint:disable
<template>
  <div>
      <h2>Home</h2>
      <swiper>

      </swiper>
  </div>
</template>

<script lang="ts">
  // tslint:enable
  import { Component, Prop, Vue, Watch } from 'vue-property-decorator'
  import { getHomeMultidata } from '../../network/home'
  import { Swiper, SwiperItem } from '../../components/common/swiper'
  
  @Component({
    components: {
      Swiper

    },
  })
  export default class Home extends Vue {
    banners: object[] = []
    recommends: object[] = [] 
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
    
</style>
