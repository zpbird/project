// tslint:disable
<template>
  <div class="recommend">
    <div v-for="item in recommends" class="recommend-item">
      <a :href="item.link">
        <img alt="" :src="item.image"/>
        <div>{{item.title}}</div>
      </a>
    </div>
  </div>
</template>

<script lang="ts">
  // tslint:enable  
  import { Component, Prop, Vue } from 'vue-property-decorator';
  @Component({
    components: {

    },
  })
  export default class RecommendView extends Vue {
    @Prop({type: Array}) recommends!: string[]
  }
  // tslint:disable 
</script>

<style scoped lang="scss">
  .recommend{
    display: flex;
    text-align: center;
    font-size: 12px;
    padding: 10px 0 20px;
    border-bottom: 10px solid #eee;
  }
  .recommend-item{
    flex: 1;
  }
  .recommend-item img {
    width: 70px;
    height: 70px;
    margin-bottom: 10px;
  }
</style>
