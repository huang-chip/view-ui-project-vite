<template>
  <div class="demo">
    <!-- <Button type="primary" @click="handleClick">Welcome!</Button> -->
    <createMoudleA />
  </div>
</template>

<script>
import createMoudleA from './createMoudleA.vue';
export default {
  name: 'Demo',
  components: {
    createMoudleA
  },
  setup() {
  },
  methods: {
  }
}
</script>

<style lang="less" scoped>
  .demo{
    font-size: @font-size;
    text-align: center;
  }
</style>
