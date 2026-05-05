// content-router.js (ExtensionsCDP版)
window.docsContent = window.docsContent || {};

/** ページIDとHTMLファイル名の対応表 */
const PAGE_ASSETS = {
  'WebSocketViaNamedPipe': 'js/asset/WebSocketViaNamedPipe.html',
};

/** 各ページ固有のタイトル定義 */
const PAGE_TITLES = {
  'WebSocketViaNamedPipe': 'WebSocketViaNamedPipe 拡張機能 | Starter Web Scraping Kit',
};

async function fetchPageContent(pageId) {
  if (window.docsContent[pageId]) return window.docsContent[pageId];

  const path = PAGE_ASSETS[pageId];
  if (!path) {
    console.warn(`[content-router] Unknown page: "${pageId}"`);
    return null;
  }

  try {
    const response = await fetch(path);
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const html = await response.text();
    window.docsContent[pageId] = html;
    return html;
  } catch (err) {
    console.error(`[content-router] Failed to load "${path}":`, err);
    return null;
  }
}
