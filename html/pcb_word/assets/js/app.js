/* ============================================================
   弘翔电器 HONGXIANG ELECTRONICS — 站点核心脚本
   多语言引擎 + 公共组件注入 + 页面交互（精简版）
   ============================================================ */
(function () {
  'use strict';

  /* ---------- 1. 语言配置 ---------- */
  var DEFAULT_LANG = 'zh';
  var LANGS = [
    { code: 'zh', name: '简体中文',   en: 'Chinese (CN)',   flag: '🇨🇳', dir: 'ltr' },
    { code: 'en', name: 'English',    en: 'English',        flag: '🇺🇸', dir: 'ltr' },
    { code: 'ja', name: '日本語',     en: 'Japanese',       flag: '🇯🇵', dir: 'ltr' },
    { code: 'de', name: 'Deutsch',    en: 'German',         flag: '🇩🇪', dir: 'ltr' },
    { code: 'fr', name: 'Français',   en: 'French',         flag: '🇫🇷', dir: 'ltr' },
    { code: 'es', name: 'Español',    en: 'Spanish',        flag: '🇪🇸', dir: 'ltr' },
    { code: 'pt', name: 'Português',  en: 'Portuguese',     flag: '🇵🇹', dir: 'ltr' },
    { code: 'ru', name: 'Русский',    en: 'Russian',        flag: '🇷🇺', dir: 'ltr' },
    { code: 'ko', name: '한국어',     en: 'Korean',         flag: '🇰🇷', dir: 'ltr' },
    { code: 'ar', name: 'العربية',    en: 'Arabic',         flag: '🇸🇦', dir: 'rtl' },
    { code: 'vi', name: 'Tiếng Việt', en: 'Vietnamese', flag: '🇻🇳', dir: 'ltr' },
    { code: 'th', name: 'ไทย', en: 'Thai', flag: '🇹🇭', dir: 'ltr' },
    { code: 'id', name: 'Bahasa Indonesia', en: 'Indonesian', flag: '🇮🇩', dir: 'ltr' },
    { code: 'ms', name: 'Bahasa Melayu', en: 'Malay', flag: '🇲🇾', dir: 'ltr' },
    { code: 'tl', name: 'Filipino', en: 'Filipino', flag: '🇵🇭', dir: 'ltr' },
    { code: 'my', name: 'မြန်မာစာ', en: 'Burmese', flag: '🇲🇲', dir: 'ltr' },
    { code: 'fa', name: 'فارسی', en: 'Persian', flag: '🇮🇷', dir: 'rtl' },
    { code: 'el', name: 'Ελληνικά', en: 'Greek', flag: '🇬🇷', dir: 'ltr' },
    { code: 'tr', name: 'Türkçe', en: 'Turkish', flag: '🇹🇷', dir: 'ltr' }
  ];

  // 顶栏只保留四个同层级页面入口；首页内的板块由页面本身与页脚承载
  var NAV = [
    { href: 'index.html',    key: 'nav.home',    page: 'home' },
    { href: 'services.html', key: 'nav.svc',     page: 'services' },
    { href: 'about.html',    key: 'nav.about',   page: 'about' },
    { href: 'contact.html',  key: 'nav.contact', page: 'contact' }
  ];

  /* ---------- 2. 图标 ---------- */
  function svg(d) {
    return '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.7" ' +
           'stroke-linecap="round" stroke-linejoin="round">' + d + '</svg>';
  }
  var I = {
    chip:  svg('<rect x="7" y="7" width="10" height="10" rx="2"/><path d="M10 3v2M14 3v2M10 19v2M14 19v2M3 10h2M3 14h2M19 10h2M19 14h2"/>'),
    check: svg('<path d="M20 6L9 17l-5-5"/>'),
    chev:  svg('<path d="M6 9l6 6 6-6"/>'),
    phone: svg('<path d="M22 16.9v3a2 2 0 0 1-2.2 2 19.8 19.8 0 0 1-8.6-3.1 19.5 19.5 0 0 1-6-6A19.8 19.8 0 0 1 2.1 4.2 2 2 0 0 1 4.1 2h3a2 2 0 0 1 2 1.7c.1 1 .4 1.9.7 2.8a2 2 0 0 1-.5 2.1L8.1 9.9a16 16 0 0 0 6 6l1.3-1.3a2 2 0 0 1 2.1-.4c.9.3 1.8.6 2.8.7a2 2 0 0 1 1.7 2z"/>'),
    mail:  svg('<path d="M4 4h16v16H4z"/><path d="M22 6l-10 7L2 6"/>'),
    pin:   svg('<path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/>'),
    clock: svg('<circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/>')
  };

  /* ---------- 3. 当前语言 ---------- */
  function readLang() {
    var q = new URLSearchParams(location.search).get('lang');
    if (q && window.HX_LANG[q]) return q;
    try {
      var s = localStorage.getItem('hx_lang');
      if (s && window.HX_LANG[s]) return s;
    } catch (e) {}
    var n = (navigator.language || '').toLowerCase();
    for (var i = 0; i < LANGS.length; i++) {
      if (n.indexOf(LANGS[i].code) === 0) return LANGS[i].code;
    }
    return DEFAULT_LANG;
  }
  var cur = readLang();

  function L(code) { return window.HX_LANG[code] || {}; }
  function t(key) {
    var v = L(cur)[key];
    if (v === undefined) v = L(DEFAULT_LANG)[key];
    return v === undefined ? key : v;
  }
  function meta(code) {
    for (var i = 0; i < LANGS.length; i++) if (LANGS[i].code === code) return LANGS[i];
    return LANGS[0];
  }

  /* ---------- 4. 渲染 ---------- */
  function applyI18n(root) {
    var scope = root || document;
    var groups = [
      ['[data-i18n]', 'textContent', 'data-i18n'],
      ['[data-i18n-html]', 'innerHTML', 'data-i18n-html'],
      ['[data-i18n-ph]', 'placeholder', 'data-i18n-ph'],
      ['[data-i18n-title]', 'title', 'data-i18n-title'],
      ['[data-i18n-aria]', 'aria-label', 'data-i18n-aria']
    ];
    groups.forEach(function (g) {
      var nodes = scope.querySelectorAll(g[0]);
      for (var i = 0; i < nodes.length; i++) {
        var key = nodes[i].getAttribute(g[2]);
        if (g[1] === 'placeholder' || g[1] === 'title' || g[1] === 'aria-label') {
          nodes[i].setAttribute(g[1], t(key));
        } else {
          nodes[i][g[1]] = t(key);
        }
      }
    });
    var page = document.body.getAttribute('data-page') || 'home';
    document.title = t('meta.' + page + '.t');
    var md = document.querySelector('meta[name="description"]');
    if (md) md.setAttribute('content', t('meta.' + page + '.d'));
  }

  function applyDir() {
    var m = meta(cur);
    document.documentElement.setAttribute('lang', m.code);
    document.documentElement.setAttribute('dir', m.dir);
  }

  function setLang(code) {
    if (!window.HX_LANG[code]) return;
    cur = code;
    try { localStorage.setItem('hx_lang', code); } catch (e) {}
    try {
      var u = new URL(location.href);
      u.searchParams.set('lang', code);
      history.replaceState(null, '', u.toString());
    } catch (e) {}
    applyDir();
    applyI18n();
    renderLangUI();
    document.dispatchEvent(new CustomEvent('hx:langchange', { detail: { lang: code } }));
    setLangOpen(false);
  }

  /* ---------- 5. 公共组件：Header ---------- */
  function headerHTML() {
    var page = document.body.getAttribute('data-page') || 'home';
    var links = NAV.map(function (n) {
      var on = (n.page === page) ? ' active' : '';
      return '<a href="' + n.href + '" class="' + on.trim() + '" data-i18n="' + n.key + '"></a>';
    }).join('');

    var mLang = LANGS.map(function (l) {
      return '<button type="button" data-lang="' + l.code + '">' +
        '<span class="flag">' + l.flag + '</span>' +
        '<span class="nm">' + l.name + '</span>' +
        '<span class="en">' + l.en + '</span>' +
        '<span class="ck">' + I.check + '</span></button>';
    }).join('');

    return '' +
      '<header class="header" id="header"><div class="container">' +
        '<a class="brand" href="index.html">' +
          '<span class="brand-mark">' + I.chip + '</span>' +
          '<span class="brand-txt"><b data-i18n="brand.name"></b><small data-i18n="brand.sub"></small></span>' +
        '</a>' +
        '<nav class="nav">' + links + '</nav>' +
        '<div class="header-r">' +
          '<div class="lang-wrap" id="langWrap">' +
            '<button class="lang-btn" id="langBtn" type="button" aria-haspopup="true" aria-expanded="false" data-i18n-aria="ui.lang">' +
              '<span class="flag" id="curFlag"></span><span id="curName"></span>' + I.chev.replace('<svg', '<svg class="chev"') +
            '</button>' +
            '<div class="lang-menu" id="langMenu" role="menu">' +
              '<div class="lm-head">' + LANGS.length + ' LANGUAGES / 语言</div>' + mLang +
            '</div>' +
          '</div>' +
          '<a class="btn btn--primary" href="contact.html" data-i18n="btn.quote"></a>' +
          '<button class="burger" id="burger" type="button" aria-expanded="false" aria-controls="mnav" data-i18n-aria="ui.menu"><span></span><span></span><span></span></button>' +
        '</div>' +
      '</div></header>' +
      '<div class="mnav" id="mnav">' + links +
        '<div class="m-cta"><a class="btn btn--primary" href="contact.html" data-i18n="btn.quote"></a></div>' +
      '</div>';
  }

  /* ---------- 6. 公共组件：Footer ---------- */
  function footerHTML() {
    // 服务条目指向服务页里对应的真实锚点，不再全部落到同一页顶
    var svcLinks = '';
    for (var i = 1; i <= 4; i++) {
      svcLinks += '<a href="services.html#s' + i + '" data-i18n="svc.' + i + '.t"></a>';
    }
    svcLinks += '<a href="services.html" data-i18n="svc.more"></a>';
    var year = new Date().getFullYear();
    return '' +
      '<div class="container">' +
        '<div class="foot-grid">' +
          '<div class="foot-brand">' +
            '<a class="brand" href="index.html">' +
              '<span class="brand-mark">' + I.chip + '</span>' +
              '<span class="brand-txt"><b data-i18n="brand.name"></b><small data-i18n="brand.sub"></small></span>' +
            '</a>' +
            '<p data-i18n="brand.desc"></p>' +
          '</div>' +
          '<div class="foot-col"><h4 data-i18n="ft.l1"></h4>' +
            '<a href="index.html" data-i18n="nav.home"></a>' +
            '<a href="services.html" data-i18n="nav.svc"></a>' +
            '<a href="about.html" data-i18n="nav.about"></a>' +
            '<a href="index.html#cap" data-i18n="nav.cap"></a>' +
            '<a href="index.html#ind" data-i18n="nav.ind"></a>' +
            '<a href="contact.html" data-i18n="nav.contact"></a>' +
          '</div>' +
          '<div class="foot-col"><h4 data-i18n="ft.l2"></h4>' + svcLinks + '</div>' +
          '<div class="foot-col"><h4 data-i18n="ft.l3"></h4>' +
            '<ul class="foot-contact">' +
              '<li>' + I.pin + '<span data-i18n="ft.addr"></span></li>' +
              '<li>' + I.phone + '<a href="tel:+8676088886688" data-i18n="ft.tel"></a></li>' +
              '<li>' + I.mail + '<a href="mailto:sales@hongxiang-pcba.com" data-i18n="ft.mail"></a></li>' +
              '<li>' + I.clock + '<span data-i18n="ft.hour"></span></li>' +
            '</ul>' +
          '</div>' +
        '</div>' +
        '<div class="foot-bottom">' +
          '<span>© ' + year + ' ' + t('ft.rights').replace(/^©\s*\d{4}\s*/, '') + '</span>' +
        '</div>' +
      '</div>';
  }

  function mount() {
    var h = document.getElementById('site-header');
    var f = document.getElementById('site-footer');
    if (h) h.innerHTML = headerHTML();
    if (f) f.innerHTML = footerHTML();
  }

  function renderLangUI() {
    var m = meta(cur);
    var f = document.getElementById('curFlag');
    var n = document.getElementById('curName');
    if (f) f.textContent = m.flag;
    if (n) n.textContent = m.name;
    var btns = document.querySelectorAll('#langMenu [data-lang]');
    for (var i = 0; i < btns.length; i++) {
      btns[i].classList.toggle('on', btns[i].getAttribute('data-lang') === cur);
    }
  }

  /* ---------- 7. 交互 ---------- */
  function setLangOpen(open) {
    var w = document.getElementById('langWrap');
    var b = document.getElementById('langBtn');
    if (w) w.classList.toggle('open', open);
    if (b) b.setAttribute('aria-expanded', open ? 'true' : 'false');
  }
  function closeLang() { setLangOpen(false); }

  function setMenuOpen(open) {
    var b = document.getElementById('burger');
    var m = document.getElementById('mnav');
    if (b) { b.classList.toggle('on', open); b.setAttribute('aria-expanded', open ? 'true' : 'false'); }
    if (m) m.classList.toggle('on', open);
  }

  function initFaq() {
    var items = document.querySelectorAll('.faq-item');
    for (var i = 0; i < items.length; i++) {
      var q = items[i].querySelector('.faq-q');
      var a = items[i].querySelector('.faq-a');
      if (!q || !a) continue;
      a.id = 'faq-a-' + (i + 1);
      q.setAttribute('aria-controls', a.id);
      q.setAttribute('aria-expanded', 'false');
    }
    document.addEventListener('click', function (e) {
      var q = e.target.closest ? e.target.closest('.faq-q') : null;
      if (!q) return;
      var item = q.parentNode;
      var open = item.classList.contains('open');
      var all = document.querySelectorAll('.faq-item');
      for (var j = 0; j < all.length; j++) {
        all[j].classList.remove('open');
        var aa = all[j].querySelector('.faq-a');
        var qq = all[j].querySelector('.faq-q');
        if (aa) aa.style.maxHeight = null;
        if (qq) qq.setAttribute('aria-expanded', 'false');
      }
      if (!open) {
        item.classList.add('open');
        var pa = item.querySelector('.faq-a');
        if (pa) pa.style.maxHeight = pa.scrollHeight + 'px';
        q.setAttribute('aria-expanded', 'true');
      }
    });
    // 切换语言后答案长度会变化，重新测量已展开项
    document.addEventListener('hx:langchange', function () {
      var a = document.querySelector('.faq-item.open .faq-a');
      if (a) a.style.maxHeight = a.scrollHeight + 'px';
    });
  }

  function initUI() {
    // 语言菜单：点击开关 / 选择语言 / 点击外部关闭 / Esc 关闭
    document.addEventListener('click', function (e) {
      var btn = e.target.closest ? e.target.closest('#langBtn') : null;
      var wrap = document.getElementById('langWrap');
      if (btn && wrap) { setLangOpen(!wrap.classList.contains('open')); return; }
      var pick = e.target.closest ? e.target.closest('[data-lang]') : null;
      if (pick) { setLang(pick.getAttribute('data-lang')); return; }
      if (wrap && !e.target.closest('#langWrap')) closeLang();
    });
    document.addEventListener('keydown', function (e) {
      if (e.key !== 'Escape' && e.key !== 'Esc') return;
      closeLang();
      setMenuOpen(false);
    });

    // 移动端菜单
    var burger = document.getElementById('burger');
    var mnav = document.getElementById('mnav');
    if (burger && mnav) {
      burger.addEventListener('click', function () {
        setMenuOpen(!mnav.classList.contains('on'));
      });
      mnav.addEventListener('click', function (e) {
        if (e.target.tagName === 'A') setMenuOpen(false);
      });
    }

    // 头部阴影 / 回到顶部
    var header = document.getElementById('header');
    var totop = document.getElementById('totop');
    function onScroll() {
      var y = window.scrollY || document.documentElement.scrollTop;
      if (header) header.classList.toggle('stuck', y > 8);
      if (totop) totop.classList.toggle('on', y > 520);
    }
    window.addEventListener('scroll', onScroll, { passive: true });
    onScroll();
    if (totop) totop.addEventListener('click', function () {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    initFaq();
    // 不做滚动淡入：正文始终可见，脚本失败时页面也不会变空白
  }

  /* ---------- 8. 启动 ---------- */
  function boot() {
    applyDir();
    mount();
    applyI18n();
    renderLangUI();
    initUI();

    // hreflang（多语言 SEO 提示）
    var head = document.head;
    if (head) {
      LANGS.forEach(function (l) {
        var link = document.createElement('link');
        link.rel = 'alternate';
        link.hreflang = l.code;
        link.href = (location.pathname.split('/').pop() || 'index.html') + '?lang=' + l.code;
        head.appendChild(link);
      });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
  } else {
    boot();
  }

  window.HX = { setLang: setLang, t: t, LANGS: LANGS, getLang: function () { return cur; } };
})();
