<template>
  <div class="demo">
    <Button type="primary" @click="handleClick">Welcome!</Button>
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
    handleClick() {
      this.$Message.info('Welcome to View UI Plus Demo!')
    }
  }
}
</script>

<style lang="less" scoped>
  .demo{
    font-size: @font-size;
    text-align: center;
  }
</style>
