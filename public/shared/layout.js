/* ==========================================================================
   Mobileum Caller ID Platform - Layout & App Shell Injector
   ========================================================================== */

(function () {
  // 1. Inject Favicon
  function injectFavicon() {
    let link = document.querySelector('link[rel="icon"]') || document.querySelector('link[rel="shortcut icon"]');
    if (!link) {
      link = document.createElement('link');
      link.rel = 'icon';
      link.type = 'image/png';
      link.href = '/shared/assets/favicon.png';
      document.head.appendChild(link);
    }
  }

  // 2. Splash Screen on first load (1.5 seconds)
  function setupSplashScreen() {
    const splashKey = 'mobileum_splash_viewed_' + (window.location.port || 'site');
    // If already shown in this tab session, skip splash for instant navigation
    if (sessionStorage.getItem(splashKey)) return;

    const overlay = document.createElement('div');
    overlay.id = 'mobileumSplash';
    overlay.className = 'splash-overlay';
    overlay.innerHTML = `
      <div class="splash-video-box">
        <video src="/shared/assets/mobileum-logo-animation.mp4" autoplay muted playsinline></video>
      </div>
      <div class="splash-brand">Mobileum</div>
      <div style="font-size:12px;color:#5B6479;margin-top:4px;font-weight:500;">Caller ID Platform &middot; Philippines</div>
    `;

    document.documentElement.appendChild(overlay);

    setTimeout(() => {
      overlay.classList.add('fade-out');
      setTimeout(() => overlay.remove(), 550);
      try { sessionStorage.setItem(splashKey, '1'); } catch (e) {}
    }, 1500);
  }

  // Run initial head setups
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      injectFavicon();
      setupSplashScreen();
    });
  } else {
    injectFavicon();
    setupSplashScreen();
  }

  // 3. Reusable Portal Shell Injector
  window.setupPortalLayout = function (config) {
    const {
      portalName = 'Mobileum Portal',
      portalRole = 'operator', // 'operator' | 'enterprise' | 'customer'
      tabs = [],
      defaultTab = tabs[0]?.id || '',
      onTabChange = null
    } = config;

    let currentTab = defaultTab;

    // SVG Icons
    const ICONS = {
      dashboard: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="7" height="9" rx="1"/><rect x="14" y="3" width="7" height="5" rx="1"/><rect x="14" y="12" width="7" height="9" rx="1"/><rect x="3" y="16" width="7" height="5" rx="1"/></svg>',
      packages: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></svg>',
      requests: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><polyline points="10 9 9 9 8 9"/></svg>',
      numbers: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/></svg>',
      sheet: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/><line x1="10" y1="9" x2="8" y2="9"/></svg>',
      register: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/><path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/></svg>',
      brand: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 2l3.09 6.26L22 9.27l-5 4.87 1.18 6.88L12 17.77l-6.18 3.25L7 14.14 2 9.27l6.91-1.01L12 2z"/></svg>',
      subscription: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></svg>',
      departments: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>',
      history: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>',
      phone: '<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="5" y="2" width="14" height="20" rx="2" ry="2"/><line x1="12" y1="18" x2="12.01" y2="18"/></svg>'
    };

    // Build sidebar navigation links HTML
    const navItemsHtml = tabs.map(t => `
      <button class="nav-item ${t.id === currentTab ? 'active' : ''}" data-tab="${t.id}" id="nav_${t.id}">
        <span class="icon">${ICONS[t.icon] || ICONS.dashboard}</span>
        <span>${t.label}</span>
        ${t.badgeId ? `<span class="badge" id="${t.badgeId}" style="display:none">0</span>` : ''}
      </button>
    `).join('');

    const isMultiPort = (window.location.port === '4001' || window.location.port === '4002' || window.location.port === '4003') && 
                        (window.location.hostname === 'localhost' || window.location.hostname === '127.0.0.1');

    const opHref = isMultiPort ? 'http://localhost:4001' : '/operator';
    const enHref = isMultiPort ? 'http://localhost:4002' : '/enterprise';
    const cuHref = isMultiPort ? 'http://localhost:4003' : '/customer';
    const sheetHref = isMultiPort ? '/sheet.html' : '/operator/sheet.html';

    const opLabel = isMultiPort ? 'Operator (4001)' : 'Operator Portal';
    const enLabel = isMultiPort ? 'Enterprise (4002)' : 'Enterprise Portal';
    const cuLabel = isMultiPort ? 'Customer Phone (4003)' : 'Customer Phone';

    // Prepare app-shell HTML structure
    const sidebarHtml = `
      <aside class="app-sidebar">
        <div class="sidebar-header">
          <div class="sidebar-logo">
            <img src="/shared/assets/mobileum-logo.png" alt="Mobileum Logo">
          </div>
          <div class="sidebar-brand">
            <span class="name">Mobileum</span>
            <span class="role">${portalName.replace('Mobileum ', '')}</span>
          </div>
        </div>

        <nav class="sidebar-nav">
          ${navItemsHtml}
        </nav>

        <div class="sidebar-footer">
          <div style="font-weight:600;color:#CBD5E0;margin-bottom:6px;display:flex;align-items:center;justify-content:space-between;">
            <span>PORTALS</span>
            <span style="color:#22D3EE;font-size:10px;">PH &middot; +63</span>
          </div>
          <div class="portal-switcher">
            <a class="portal-link ${portalRole === 'operator' ? 'active' : ''}" id="shell_link_op" href="${opHref}" target="_blank">
              <span>${opLabel}</span>
              <span>&rarr;</span>
            </a>
            <a class="portal-link ${portalRole === 'enterprise' ? 'active' : ''}" id="shell_link_en" href="${enHref}" target="_blank">
              <span>${enLabel}</span>
              <span>&rarr;</span>
            </a>
            <a class="portal-link ${portalRole === 'customer' ? 'active' : ''}" id="shell_link_cu" href="${cuHref}" target="_blank">
              <span>${cuLabel}</span>
              <span>&rarr;</span>
            </a>
            ${portalRole === 'operator' ? `
            <a class="portal-link" href="${sheetHref}" target="_blank" style="margin-top:4px;background:rgba(34,211,238,0.1);color:#22D3EE;">
              <span>Live Excel Sheet</span>
              <span>&nearr;</span>
            </a>
            <a class="portal-link" href="/api/database" style="margin-top:2px;">
              <span>Download Excel (.xlsx)</span>
              <span>&darr;</span>
            </a>
            ` : ''}
          </div>
        </div>
      </aside>
    `;

    const topbarHtml = `
      <header class="app-topbar">
        <div class="topbar-left">
          <div style="display:flex;align-items:center;gap:12px;">
            <img src="/shared/assets/mobileum-logo.png" alt="Mobileum" style="width:28px;height:28px;object-fit:contain;border-radius:4px;">
            <h1 class="topbar-title" id="shell_topbar_title">${portalName}</h1>
          </div>
          <span style="color:var(--text-secondary);font-size:14px;">/</span>
          <span id="shell_tab_title" style="font-size:14px;font-weight:600;color:var(--text-muted);">${tabs.find(t => t.id === currentTab)?.label || ''}</span>
        </div>
        <div class="topbar-right">
          <div class="live-indicator">
            <span class="pulse-dot"></span>
            <span>Live Sync</span>
          </div>
          <span class="demo-tag">Demo Mock-Up</span>
        </div>
      </header>
    `;

    // Wrap page body content in app-shell structure
    const bodyChildren = Array.from(document.body.children).filter(el => 
      el.id !== 'mobileumSplash' && 
      el.tagName !== 'SCRIPT' && 
      !el.classList.contains('toast-container')
    );

    const mainContainer = document.createElement('div');
    mainContainer.className = 'app-main';
    mainContainer.innerHTML = topbarHtml;

    const contentContainer = document.createElement('div');
    contentContainer.className = 'app-content';
    mainContainer.appendChild(contentContainer);

    // Move existing body content into contentContainer
    bodyChildren.forEach(child => contentContainer.appendChild(child));

    const shell = document.createElement('div');
    shell.className = 'app-shell';
    shell.innerHTML = sidebarHtml;
    shell.appendChild(mainContainer);

    document.body.insertBefore(shell, document.body.firstChild);

    // Tab switching handler
    function switchTab(tabId) {
      currentTab = tabId;
      document.querySelectorAll('.app-sidebar .nav-item').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
      });
      const tObj = tabs.find(t => t.id === tabId);
      const tabTitleEl = document.getElementById('shell_tab_title');
      if (tabTitleEl && tObj) tabTitleEl.textContent = tObj.label;

      // Show/hide tab panels if containers with data-tab-panel exist
      document.querySelectorAll('[data-tab-panel]').forEach(panel => {
        panel.style.display = panel.dataset.tabPanel === tabId ? '' : 'none';
      });

      if (typeof onTabChange === 'function') {
        onTabChange(tabId);
      }
    }

    // Attach click events to nav items
    document.querySelectorAll('.app-sidebar .nav-item').forEach(btn => {
      btn.addEventListener('click', () => {
        switchTab(btn.dataset.tab);
      });
    });

    // Initial tab setup
    switchTab(defaultTab);

    // Update portal links dynamically when /api/info loads
    api('/api/info').then(info => {
      const lOp = document.getElementById('shell_link_op');
      const lEn = document.getElementById('shell_link_en');
      const lCu = document.getElementById('shell_link_cu');
      if (lOp && info.operatorUrl) lOp.href = info.operatorUrl;
      if (lEn && info.enterpriseUrl) lEn.href = info.enterpriseUrl;
      if (lCu && info.customerUrl) lCu.href = info.customerUrl;
    }).catch(() => {});

    return { switchTab };
  };
})();
