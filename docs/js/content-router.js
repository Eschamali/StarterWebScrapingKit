// content-router.js
// ページIDとHTMLアセットのマッピングを一元管理します (Select Case 相当)
// HTMLコンテンツは docs/js/asset/*.html に分離されており、fetch()で取得します

window.docsContent = window.docsContent || {};

/** ページIDとHTMLファイル名の対応表 */
const PAGE_ASSETS = {
  'top':               'js/asset/top.html',
  'cdp-intro':         'js/asset/cdp-intro.html',
  'cdp-demos':         'js/asset/cdp-demos.html',
  'cdp-methods':       'js/asset/cdp-methods.html',
  'cdp-advanced':      'js/asset/cdp-advanced.html',
  'bidi-intro':        'js/asset/bidi-intro.html',
  'bidi-methods':      'js/asset/bidi-methods.html',
  'technical-journey': 'js/asset/technical-journey.html',
  'bidi-update':        'js/asset/bidi-update.html',
  'bidi-story':        'js/asset/bidi-story.html',
  'userform-edge':     'js/asset/userform-edge.html',
};

/**
 * 指定ページのHTMLをfetch()で取得し、window.docsContentにキャッシュします。
 * @param {string} pageId - ロードするページID
 * @returns {Promise<string|null>} HTMLテキスト、またはエラー時null
 */
async function fetchPageContent(pageId) {
  // キャッシュ済みならすぐ返す
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

/** 全ページを事前にプリフェッチ（任意の最適化） */
function prefetchAllPages() {
  Object.keys(PAGE_ASSETS).forEach(pageId => fetchPageContent(pageId));
}
