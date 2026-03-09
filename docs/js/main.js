// main.js - Handles page content loading and initialization

const contentContainer = document.getElementById('content-container');
const navItems = document.querySelectorAll('.nav-tree .nav-item:not(.external)');

function loadContent(pageId) {
	if (!window.docsContent || !window.docsContent[pageId]) return;

	// Show Loader
	contentContainer.innerHTML = '<div class="loader"></div>';

	// Update Nav Activity
	navItems.forEach(item => {
		item.classList.remove('active');
		const href = item.getAttribute('href');
		if (href && href.endsWith('#' + pageId)) {
			item.classList.add('active');

			// Auto expand parent menus if needed
			let parentDetail = item.closest('details');
			if (parentDetail) {
				parentDetail.open = true;
			}
		}
	});

	// Animate content out and in
	contentContainer.style.opacity = '0';

	setTimeout(() => {
		contentContainer.innerHTML = window.docsContent[pageId];
		// Apply Syntax Highlighting
		document.querySelectorAll('pre code').forEach((block) => {
			hljs.highlightElement(block);
		});

		contentContainer.style.opacity = '1';
		window.scrollTo({ top: 0, behavior: 'smooth' });

		// Track Page View in Google Analytics
		if (typeof gtag === 'function') {
			gtag('config', 'G-5CW3LKTJWH', {
				'page_title': pageId,
				'page_path': location.pathname + location.hash
			});
		}
	}, 200);
}

// Initial Load based on URL hash
window.addEventListener('DOMContentLoaded', () => {
	const hash = window.location.hash.replace('#', '');
	if (hash && window.docsContent && window.docsContent[hash]) {
		loadContent(hash);
	} else {
		loadContent('top'); // Default page
	}
});
