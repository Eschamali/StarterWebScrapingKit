import { defineConfig } from 'vitepress'
import { withMermaid } from 'vitepress-plugin-mermaid'

// GitHub Pages: https://eschamali.github.io/StarterWebScrapingKit/
const domain = 'https://eschamali.github.io'
const base = '/StarterWebScrapingKit/'
const siteUrl = `${domain}${base}` // ➔ 'https://eschamali.github.io/StarterWebScrapingKit/'

export default withMermaid(
  defineConfig({
    title: 'Starter Web Scraping Kit',
    description: 'Excel VBA で CDP / WebDriver BiDi によるブラウザ自動操作',
    lang: 'ja-JP',
    base: base,
    cleanUrls: true,
    lastUpdated: true,

    // サイトマップを自動生成！
    sitemap: {
        hostname: siteUrl
    },

    // 旧サイト #page-id リンク切れ対策（ハッシュはサーバーに届かないためクライアントで置換）
    head: [
      // ① 旧サイトのハッシュリダイレクト用スクリプト
      ['script', {}, `
        (function() {
          var hash = window.location.hash;
          var hashMap = {
            "#userform-powershell": "${base}userform/powershell",
          };
          if (hash && hashMap[hash]) {
            window.location.replace(hashMap[hash]);
          }
        })();
      `],

      // ② favicon
      ['link', { rel: 'icon', type: 'image/png', href: `${base}favicon.png` }],

      // ③ Twitter（X）埋め込み用スクリプト
      ['script', { async: '', src: 'https://platform.twitter.com/widgets.js', charset: 'utf-8' }],

      // --- ④ OGP (SNSシェア画像) 設定 ----------------------------------
      ['meta', { property: 'og:image', content: `${siteUrl}browser-control.png` }],
      ['meta', { property: 'og:url', content: siteUrl }], 
      ['meta', { property: 'og:type', content: 'website' }],
      ['meta', { name: 'twitter:card', content: 'summary_large_image' }], // 画像を大きく表示させる指定

      // --- ⑤ Google アナリティクス (GA4) 設定 --------------------------
      ['script', { async: '', src: 'https://www.googletagmanager.com/gtag/js?id=G-5CW3LKTJWH' }],
      ['script', {}, `
        window.dataLayer = window.dataLayer || [];
        function gtag(){dataLayer.push(arguments);}
        gtag('js', new Date());
        gtag('config', 'G-5CW3LKTJWH');
      `]
    ],

    themeConfig: {
      logo: {
        light: '/Logo_Light.png',
        dark: '/Logo_Dark.png',
        alt: 'Starter Web Scraping Kit'
      },
      siteTitle: 'Starter Web Scraping Kit',
      nav: [
        { text: 'はじめに', link: '/intro' },
        { text: 'ガイド', link: '/guides/navigation' },
        { text: 'API', link: '/api/cdp/CDPContext' }
      ],

      sidebar: [
        {
          text: '導入',
          items: [
            { text: '概要', link: '/intro' },
            { text: 'はじめに', link: '/getting-started' },
            { text: 'アーキテクチャ', link: '/concepts/architecture' },
            { text: '設計思想', link: '/concepts/design-philosophy' },
            { text: 'CDP と BiDi', link: '/concepts/cdp-vs-bidi' }
          ]
        },
        {
          text: 'WebSocketモードでの制御について',
          items: [
            { text: '設計思想について', link: '/websocket/design' },
            { text: 'WebSocketモードでできること', link: '/websocket/capabilities' }
          ]
        },
        {
          text: 'ガイド',
          items: [
            { text: 'ページ遷移', link: '/guides/navigation' },
            { text: '要素の取得', link: '/guides/selectors' },
            { text: '入力とクリック', link: '/guides/input' },
            { text: 'JavaScript 実行', link: '/guides/javascript' },
            { text: 'イベント購読', link: '/guides/events' },
            { text: 'マルチタブ', link: '/guides/multi-tab' },
            { text: 'タイムアウト設定方法について', link: '/guides/timeout' },
            { text: '再接続 (reattach)', link: '/guides/reattach' },
            { text: '低レイヤーBiDi/CDPコマンドについて', link: '/guides/extend-raw-protocol' }
          ]
        },
        {
          text: 'API — CDP',
          items: [
            { text: 'CDPBrowser', link: '/api/cdp/CDPBrowser' },
            { text: 'CDPContext', link: '/api/cdp/CDPContext' },
            { text: 'CDPElement', link: '/api/cdp/CDPElement' }
          ]
        },
        {
          text: 'API — WebDriver BiDi',
          items: [
            { text: 'WebDriverBiDiMode', link: '/api/bidi/WebDriverBiDiMode' },
            { text: 'WebDriverBiDiContext', link: '/api/bidi/WebDriverBiDiContext' }
          ]
        },
        {
          text: '開発秘話',
          items: [
            { text: 'BiDi 登場秘話', link: '/stories/bidi-story' }
          ]
        },
        {
          text: 'UserForm × モダンブラウザ',
          items: [
            { text: 'はじめに', link: '/userform/intro' },
            { text: 'Lv.1 Edge 埋め込み', link: '/userform/edge' },
            { text: 'Lv.10 PowerShell 経由', link: '/userform/powershell' },
            { text: 'Lv.99 Excel 単体', link: '/userform/vba-only' },
            { text: '総括・比較', link: '/userform/summary' }
          ]
        }
      ],

      search: {
        provider: 'local',
        options: {
          translations: {
            button: { buttonText: '検索', buttonAriaLabel: '検索' },
            modal: {
              noResultsText: '結果がありません',
              resetButtonTitle: 'クリア',
              footer: { selectText: '選択', navigateText: '移動', closeText: '閉じる' }
            }
          }
        }
      },

      socialLinks: [
        { icon: 'github', link: 'https://github.com/Eschamali/StarterWebScrapingKit' }
      ],

      footer: {
        message: 'Excel VBA × CDP / WebDriver BiDi',
        copyright: 'Copyright © 2026 エスカマリ'
      },

      outline: { label: 'このページ', level: [2, 3] },
      docFooter: { prev: '前へ', next: '次へ' },
      returnToTopLabel: 'トップへ戻る',
      sidebarMenuLabel: 'メニュー',
      darkModeSwitchLabel: '外観',
      lightModeSwitchTitle: 'ライトモード',
      darkModeSwitchTitle: 'ダークモード'
    },

    markdown: {
      theme: { light: 'github-light', dark: 'github-dark' },
      languages: ['vb', 'bash', 'json', 'javascript', 'powershell', 'text']
    },

    // vitepress-plugin-mermaid
    mermaid: {
      // light 用。dark はプラグイン側で自動切替
    }
  })
)
