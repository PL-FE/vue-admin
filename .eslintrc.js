module.exports = {
  root: true,
  globals: {
    GC: true
  },
  env: {
    node: true
  },
  extends: ['plugin:vue/essential', '@vue/standard'],
  parserOptions: {
    parser: 'babel-eslint'
  },
  rules: {}
}
