// main.js (ExtensionsCDP版)

const contentContainer = document.getElementById('content-container');
const navItems = document.querySelectorAll('.nav-tree .nav-item:not(.external)');

async function loadContent(pageId) {
	const html = await fetchPageContent(pageId);
	if (!html) return;

	contentContainer.innerHTML = '<div class="loader"></div>';

	navItems.forEach(item => {
		item.classList.remove('active');
		const href = item.getAttribute('href');
		if (href && href.endsWith('#' + pageId)) {
			item.classList.add('active');
			const parentDetail = item.closest('details');
			if (parentDetail) parentDetail.open = true;
		}
	});

	const pageTitle = PAGE_TITLES[pageId] || 'CDP Extensions';
	document.title = pageTitle;

	contentContainer.style.opacity = '0';
	
	setTimeout(async () => {
		contentContainer.innerHTML = html;

		// Syntax Highlighting
		document.querySelectorAll('pre code').forEach(block => {
			hljs.highlightElement(block);
		});

		// Mermaid Rendering
		if (window.mermaid) {
			try {
				await mermaid.run({
					nodes: document.querySelectorAll('.mermaid')
				});
			} catch (e) {
				console.error("Mermaid error:", e);
			}
		}

		contentContainer.style.opacity = '1';
		window.scrollTo({ top: 0, behavior: 'smooth' });

		if (window.twttr && window.twttr.widgets) {
			window.twttr.widgets.load(contentContainer);
		}

		if (typeof gtag === 'function') {
			gtag('event', 'page_view', {
				page_title: pageId,
				page_location: location.href,
				page_path: location.pathname + location.hash
			});
		}
	}, 200);
}

function handleHashChange() {
	const hash = window.location.hash.replace('#', '');
	loadContent(hash && PAGE_ASSETS[hash] ? hash : 'intro');
}

window.addEventListener('hashchange', handleHashChange);
window.addEventListener('DOMContentLoaded', handleHashChange);
