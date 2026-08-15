---
layout: home

title: Starter Web Scraping Kit - Excel VBAでブラウザを自動操作
description: WebDriver不要！CDP (Chrome DevTools Protocol) と WebDriver BiDi を Excel VBA 単体で直接ブラウザ(Edge/Chrome)制御する次世代Webスクレイピングキットの公式ドキュメントです。

hero:
  name: Starter Web Scraping Kit
  text: Excel VBA でブラウザを操る
  tagline: WebDriver 不要。CDP と WebDriver BiDi を VBA 単体で。
  image:
    light: /Top-light.png
    dark: /Top-Dark.png
    alt: Starter Web Scraping Kit
  actions:
    - theme: brand
      text: はじめに
      link: /getting-started
    - theme: alt
      text: 概要を読む
      link: /intro
    - theme: alt
      text: GitHub
      link: https://github.com/Eschamali/StarterWebScrapingKit

features:
  - title: Chrome DevTools Protocol
    details: パイプ通信で Edge / Chrome を直接制御。CDPBrowser → CDPContext → CDPElement の三層モデル。
    link: /api/cdp/CDPContext
    linkText: CDP API
  - title: Playwright 風の学び方
    details: 導入 → やりたいことガイド → クラス別 API。コードは CDP / BiDi を並べて掲載。
    link: /guides/navigation
    linkText: ガイドへ
  - title: コアロジック徹底比較
    details: Puppeteer / Playwright の実ソースと1行ずつ突き合わせ。バッファ管理・ディスパッチ・非同期処理は、どこまで並んでいるのか。
    link: /core-comparison/
    linkText: 比較レポートへ
  - title: 開発秘話
    details: Puppeteer / Playwright 並みのコアエンジンをVBAで実現するまでの道のり
    link: /stories/birth-story.md
    linkText: 登場秘話を見る
---
