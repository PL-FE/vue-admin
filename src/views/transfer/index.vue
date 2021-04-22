<template>
  <div>
    <ElTransfer v-model="value"
      filterable
      :data="data" />
  </div>
</template>

<script>
import ElTransfer from './main.vue'
export default {
  components: { ElTransfer },
  data () {
    const generateData = _ => {
      const data = []
      for (let i = 1; i <= 10000; i++) {
        data.push({
          key: i,
          label: `备选项 ${i}`
        })
      }
      return data
    }
    return {
      data: generateData(),
      value: []
    }
  }
}
</script>
<style lang="less" scoped>
</style>
