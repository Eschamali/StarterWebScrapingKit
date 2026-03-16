// main.js - HTMLコンテンツの読み込みとナビゲーション管理

const contentContainer = document.getElementById('content-container');
const navItems = document.querySelectorAll('.nav-tree .nav-item:not(.external)');

/**
 * 指定ページを読み込んで表示します。
 * fetch()で取得したHTMLをキャッシュしてから描画します。
 * @param {string} pageId
 */
async function loadContent(pageId) {
	// HTMLを取得（キャッシュ or fetch）
	const html = await fetchPageContent(pageId);
	if (!html) return;

	// ローダー表示
	contentContainer.innerHTML = '<div class="loader"></div>';

	// ナビゲーションのアクティブ状態を更新
	navItems.forEach(item => {
		item.classList.remove('active');
		const href = item.getAttribute('href');
		if (href && href.endsWith('#' + pageId)) {
			item.classList.add('active');
			// 親の <details> を展開
			const parentDetail = item.closest('details');
			if (parentDetail) parentDetail.open = true;
		}
	});

	// フェードアウト → コンテンツ差し替え → フェードイン
	contentContainer.style.opacity = '0';
	setTimeout(() => {
		contentContainer.innerHTML = html;

		// コードブロックにシンタックスハイライトを適用
		document.querySelectorAll('pre code').forEach(block => {
			hljs.highlightElement(block);
		});

		contentContainer.style.opacity = '1';
		window.scrollTo({ top: 0, behavior: 'smooth' });

		// Google Analytics にページ遷移を記録
		if (typeof gtag === 'function') {
			gtag('event', 'page_view', {
				page_title: pageId,
				page_location: location.href,
				page_path: location.pathname + location.hash
			});
		}
	}, 200);
}

/**
 * URLのハッシュに基づいてコンテンツを切り替えます。
 */
function handleHashChange() {
	const hash = window.location.hash.replace('#', '');
	loadContent(hash && PAGE_ASSETS[hash] ? hash : 'top');
}

// 初回ロード時およびハッシュ変更時にコンテンツを更新
window.addEventListener('hashchange', handleHashChange);
window.addEventListener('DOMContentLoaded', handleHashChange);
