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
  'cdp-context-methods': 'js/asset/cdp-context-methods.html',
  'cdp-element-methods': 'js/asset/cdp-element-methods.html',
  'cdp-advanced':      'js/asset/cdp-advanced.html',
  'cdp-extension':     'js/asset/cdp-extension.html',
  'bidi-intro':        'js/asset/bidi-intro.html',
  'bidi-methods':      'js/asset/bidi-methods.html',
  'bidi-context-methods': 'js/asset/bidi-context-methods.html',
  'technical-journey': 'js/asset/technical-journey.html',
  'bidi-update':        'js/asset/bidi-update.html',
  'bidi-story':        'js/asset/bidi-story.html',
  'userform-intro':     'js/asset/userform-intro.html',
  'userform-edge':      'js/asset/userform-edge.html',
  'userform-powershell': 'js/asset/userform-powershell.html',
  'userform-vba-only':  'js/asset/userform-vba-only.html',
  'userform-summary':   'js/asset/userform-summary.html',
};

/** 各ページ固有のタイトル定義 (SEO用) */
const PAGE_TITLES = {
  'top':               'Excelでブラウザ制御 (CDP / WebDriver BiDi) | Starter Web Scraping Kit',
  'cdp-intro':         'CDP-Jsonの改良点と目的 - Excel VBA',
  'cdp-demos':         'VBA CDP実装デモコードコーナー',
  'cdp-methods':       'CDPBrowser メソッドリファレンス',
  'cdp-context-methods': 'CDPContext メソッドリファレンス',
  'cdp-element-methods': 'CDPElement メソッドリファレンス',
  'cdp-advanced':      'CDPの高度な制御手法 (インジェクション)',
  'cdp-extension':     'CDP機能拡張（テンプレート利用）',
  'bidi-story':        'WebDriver BiDi 採用の裏話 - No-EXEの真実',
  'technical-journey': 'No-EXEでBiDiを実現する仕組み',
  'bidi-intro':        'WebDriver BiDi デモコードコーナー',
  'bidi-methods':      'WebDriverBiDiMode メソッドリファレンス',
  'bidi-context-methods': 'WebDriverBiDiContext メソッドリファレンス',
  'bidi-update':       'mapperTab.js の更新手順',
  'userform-intro':    'UserFormにWebView2を導入する意義',
  'userform-edge':     'Lv.1: EdgeをUserFormに直接埋め込む方法',
  'userform-powershell': 'Lv.10: PowerShell経由でWebView2を召喚する',
  'userform-vba-only': 'Lv.99: VBAのみでWebView2を完全制御する',
  'userform-summary':  '使い分けガイド: UserFormブラウザ実装の比較',
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
