/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   SAUDI INSIGHT — Global JS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */

/* ── Custom cursor ── */
(function(){
  const cur = document.getElementById('cursor');
  const ring = document.getElementById('cursor-ring');
  if(!cur||!ring) return;
  let rx=0,ry=0,mx=0,my=0;
  document.addEventListener('mousemove',e=>{
    mx=e.clientX; my=e.clientY;
    cur.style.left=mx+'px'; cur.style.top=my+'px';
  });
  (function anim(){
    rx+=(mx-rx)*.12; ry+=(my-ry)*.12;
    ring.style.left=rx+'px'; ring.style.top=ry+'px';
    requestAnimationFrame(anim);
  })();
})();

/* ── Nav scroll effect ── */
(function(){
  const nav = document.getElementById('nav');
  if(!nav) return;
  function update(){
    const solid = nav.classList.contains('nav-solid');
    if(solid) return;
    nav.classList.toggle('scrolled', window.scrollY > 40);
  }
  window.addEventListener('scroll', update, {passive:true});
  update();
})();

/* ── FAQ accordion ── */
function toggleFaq(el){
  const item = el.closest('.faq-item');
  const open = item.classList.contains('open');
  document.querySelectorAll('.faq-item.open').forEach(i=>i.classList.remove('open'));
  if(!open) item.classList.add('open');
}

function faqFilter(cat, btn){
  document.querySelectorAll('.faq-filter').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  document.querySelectorAll('.faq-item').forEach(el=>{
    el.style.display = (cat==='all'||el.dataset.c===cat) ? '' : 'none';
  });
}
