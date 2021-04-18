<template>
  <div class="">
    <Transfer v-model="value"
      :data="data" />

  </div>
</template>

<script>
import Transfer from './main'
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
export default {
  components: { Transfer },
  data () {
    return {
      value: [],
      data: generateData()
    }
  },
  mounted () {

  },
  methods: {

  }
}
</script>

<style lang="less" scoped>
</style>
