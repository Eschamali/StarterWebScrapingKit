<script setup lang="ts">
import { useData } from 'vitepress'
import PromoMovie from './PromoMovie.vue'

// トップページ(`layout: home`)のフロントマターに、下記があればムービーを表示します
//   promo:
//     src: /promo/ReiwaIE_60s.html
//     title: 令和のIE 60秒
const { frontmatter } = useData()
</script>

<template>
  <PromoMovie v-if="frontmatter.promo?.src" :src="frontmatter.promo.src" :title="frontmatter.promo.title" />
</template>
