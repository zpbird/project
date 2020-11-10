// tslint:disable
<template>
  <swiper>
    <swiper-item v-for="(item,index) in banners" :key="index">
      <a :href="item.link">
        <img alt="" :src="item.image"/>
      </a>
    </swiper-item>
  </swiper>
</template>

<script lang="ts">
  // tslint:enable  
  import { Component, Prop, Vue } from 'vue-property-decorator'
  import { Swiper, SwiperItem } from '../../../components/common/swiper'

  @Component({
    components: {
      Swiper,
      SwiperItem
    },
  })
  export default class CpnName extends Vue {
    @Prop({type: Array}) banners!: string[]
  }
  // tslint:disable 
</script>

<style scoped lang="scss">
  
</style>
