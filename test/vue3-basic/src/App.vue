
<template>
  <h1>{{count}}</h1>
  <h1>{{double}}</h1>
  <button @click="add">add</button>
</template>

<script lang="ts">
import { defineComponent, ref, computed } from 'vue';

export default defineComponent({
    name: 'App',
    
  setup(){
    const count = ref(0)
    
    const add = () => {
      count.value++
      
    }
    const double = computed(()=>{
      return count.value * 2
      
    })
    return {
      count,
      add,
      double
    }
  }
});
</script>

<style>
#app {
  font-family: Avenir, Helvetica, Arial, sans-serif;
  -webkit-font-smoothing: antialiased;
  -moz-osx-font-smoothing: grayscale;
  text-align: center;
  color: #2c3e50;
  margin-top: 60px;
}
</style>
