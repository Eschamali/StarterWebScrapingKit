import { h } from 'vue'
import type { Theme } from 'vitepress'
import DefaultTheme from 'vitepress/theme'
import PromoMovie from './PromoMovie.vue'
import HomePromo from './HomePromo.vue'

const theme: Theme = {
  extends: DefaultTheme,
  Layout: () => {
    return h(DefaultTheme.Layout, null, {
      // トップページの、`features`カードの直前にムービーを差し込む
      'home-features-before': () => h(HomePromo)
    })
  },
  enhanceApp({ app }) {
    // Markdownから `<PromoMovie src="/promo/○○.html" />` として使えるようにする
    app.component('PromoMovie', PromoMovie)
  }
}

export default theme
