import { defineConfig } from 'vitepress'
import { withMermaid } from 'vitepress-plugin-mermaid'

// GitHub Pages: https://eschamali.github.io/StarterWebScrapingKit/
const base = '/StarterWebScrapingKit/'

export default withMermaid(
  defineConfig({
    title: 'Starter Web Scraping Kit',
    description: 'Excel VBA で CDP / WebDriver BiDi によるブラウザ自動操作',
    lang: 'ja-JP',
    base: base,
    cleanUrls: true,
    lastUpdated: true,

    // 旧サイト #page-id リンク切れ対策（ハッシュはサーバーに届かないためクライアントで置換）
    head: [
      ['script',
        { async: '', src: 'https://platform.twitter.com/widgets.js', charset: 'utf-8' } , `
        (function() {
          var hash = window.location.hash;
          var hashMap = {
            "#userform-powershell": "/StarterWebScrapingKit/userform/powershell",
          };
          if (hash && hashMap[hash]) {
            window.location.replace(hashMap[hash]);
          }
        })();
      `]
    ],

    themeConfig: {
      logo: '/logo.svg',
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
            { text: 'CDP と BiDi', link: '/concepts/cdp-vs-bidi' }
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
            { text: '再接続 (reattach)', link: '/guides/reattach' },
            { text: 'スクリーンショット', link: '/guides/screenshots' },
            { text: '生プロトコル拡張', link: '/guides/extend-raw-protocol' }
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
