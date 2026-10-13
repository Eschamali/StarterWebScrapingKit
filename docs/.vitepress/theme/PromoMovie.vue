<script setup lang="ts">
import { withBase } from 'vitepress'

// `docs/public/` 配下の、HTML製ムービー(16:9)を埋め込む共通コンポーネント
//   src  : `public/` を起点としたパス。例) /promo/ReiwaIE_60s.html
//   title: iframe のタイトル(アクセシビリティ用)
defineProps<{
  src: string
  title?: string
}>()
</script>

<template>
  <figure class="promo-movie">
    <div class="promo-frame">
      <iframe
        :src="withBase(src)"
        :title="title ?? '紹介ムービー'"
        allow="autoplay; fullscreen"
        allowfullscreen
        loading="lazy"
      ></iframe>
    </div>
  </figure>
</template>

<style scoped>
.promo-movie {
  max-width: 960px;
  margin: 24px auto;
  padding: 0 24px;
}

.promo-frame {
  position: relative;
  width: 100%;
  aspect-ratio: 16 / 9;
  overflow: hidden;
  border-radius: 12px;
  background: #000;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.25);
}

.promo-frame iframe {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  border: 0;
}
</style>
