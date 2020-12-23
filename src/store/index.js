import Vue from 'vue'
import Vuex from 'vuex'

Vue.use(Vuex)

export default new Vuex.Store({
  state: {
    tabs: [],
    curTab: '',
    radioValue: ''
  },
  getters: {
    tabs: ({ tabs }) => tabs,
    curTab: ({ curTab }) => curTab,
    radioValue ({ radioValue }) {
      return radioValue
    }
  },
  mutations: {
    setRadioValue (state, newVAlue) {
      state.radioValue = newVAlue
    },
    SET_TABS (state, name) {
      if (state.tabs.findIndex(it => it.name === name) === -1) {
        state.tabs.push({
          title: name,
          name: name
        })
      }
      if (!state.curTab) {
        state.curTab = name
      }
    },
    REMOVE_TABS (state, name) {
      const index = state.tabs.findIndex(tab => tab.name === name)
      if (state.tabs.length > 1) {
        const o = state.tabs[index + 1] || state.tabs[index - 1]
        state.curTab = o.name
        state.tabs = state.tabs.filter(tab => tab.name !== name)
      }
    },
    SET_CURTAB (state, name) {
      state.curTab = name
    }
  },
  actions: {},
  modules: {}
})
