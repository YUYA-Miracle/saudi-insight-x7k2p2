/* nav.js — ナビとフッターを動的挿入 */
(function(){
  const NAV = `
<nav id="nav">
  <a class="nav-logo" href="index.html">
    SAUDI INSIGHT
    <small>by Desert Bamboo Global</small>
  </a>
  <ul class="nav-links">
    <li><a href="saudi-basics.html">サウジアラビアとは</a></li>
    <li><a href="guide.html">投資ガイド</a></li>
    <li class="has-dd">
      <a href="market.html">サウジ市場 ▾</a>
      <div class="dd">
        <a href="market.html">市場概要</a>
        <a href="neom.html">NEOM</a>
        <a href="amaala.html">AMAALA</a>
        <a href="redsea.html">Red Sea Global</a>
        <a href="diriyah.html">Diriyah</a>
        <a href="qiddiyah.html">Qiddiyah</a>
      </div>
    </li>
    <li><a href="faq.html">FAQ</a></li>
    <li><a href="about.html">Desert Bamboo</a></li>
    <li><a href="news.html">ニュース</a></li>
  </ul>
  <a class="nav-apply" href="member.html">会員申請</a>
</nav>`;

  const FOOTER = `
<footer class="site-footer">
  <div class="footer-inner">
    <div class="footer-grid">
      <div>
        <div class="footer-logo">SAUDI INSIGHT</div>
        <p class="footer-tagline">Desert Bamboo Global 日本事務局が運営する、サウジアラビア投資・建築情報のクローズドプラットフォーム。現地ネットワーク経由の一次情報を日本語でお届けします。</p>
        <a class="footer-apply" href="member.html">会員申請 →</a>
      </div>
      <div>
        <div class="footer-col-h">サウジ市場</div>
        <ul class="footer-links">
          <li><a href="market.html">市場概要</a></li>
          <li><a href="neom.html">NEOM</a></li>
          <li><a href="amaala.html">AMAALA</a></li>
          <li><a href="redsea.html">Red Sea Global</a></li>
          <li><a href="diriyah.html">Diriyah</a></li>
          <li><a href="qiddiyah.html">Qiddiyah</a></li>
        </ul>
      </div>
      <div>
        <div class="footer-col-h">情報・ガイド</div>
        <ul class="footer-links">
          <li><a href="guide.html">投資・不動産ガイド</a></li>
          <li><a href="faq.html">FAQ</a></li>
          <li><a href="news.html">ニュース・コラム</a></li>
        </ul>
      </div>
      <div>
        <div class="footer-col-h">Desert Bamboo</div>
        <ul class="footer-links">
          <li><a href="about.html">企業紹介</a></li>
          <li><a href="member.html">会員申請</a></li>
          <li><a href="tokusho.html">特定商取引法に基づく表記</a></li>
          <li><a href="terms.html">利用規約</a></li>
          <li><a href="privacy.html">プライバシーポリシー</a></li>
        </ul>
      </div>
    </div>
    <div class="footer-bottom">
      <span class="footer-copy">© 2026 Desert Bamboo Global. All rights reserved.</span>
      <div class="footer-legal">
        <a href="tokusho.html">特定商取引法</a>
        <a href="terms.html">利用規約</a>
        <a href="privacy.html">プライバシーポリシー</a>
      </div>
    </div>
  </div>
</footer>`;

  /* cursorとnavをbodyの先頭に挿入 */
  document.body.insertAdjacentHTML('afterbegin',
    '<div id="cursor"></div><div id="cursor-ring"></div>' + NAV
  );

  /* フッターをbodyの末尾に挿入 — DOMContentLoaded後に確実に末尾へ */
  function insertFooter() {
    document.body.insertAdjacentHTML('beforeend', FOOTER);
  }
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', insertFooter);
  } else {
    insertFooter();
  }

  /* nav solid mode */
  if(document.body.dataset.navSolid){
    document.getElementById('nav').classList.add('solid');
  }
})();
