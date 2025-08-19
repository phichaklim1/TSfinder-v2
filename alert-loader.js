// alert-loader.js — แสดงข้อความเตือนจาก alert.json
(function(){
  function injectStyle(){
    if (document.getElementById('egg-alert-style')) return;
    var css = '.egg-alert{color:#dc2626;font-weight:800;background:transparent;border:0;padding:6px 10px 0;line-height:1.35;white-space:normal;}';
    var s = document.createElement('style');
    s.id = 'egg-alert-style';
    s.textContent = css;
    document.head.appendChild(s);
  }
  async function loadAlerts(){
    injectStyle();
    try{
      const res = await fetch('alert.json', { cache: 'no-store' });
      if (!res.ok) throw new Error('fetch failed');
      const data = await res.json();
      document.querySelectorAll('[data-alert-slot]').forEach(function(el){
        var key = el.getAttribute('data-alert-slot'); // "index" | "pro" | custom
        var raw = (data && (data[key] ?? data.message ?? '')).trim();
        if (raw){
          // รองรับขึ้นบรรทัดใหม่ด้วย \n
          el.innerHTML = raw.replace(/\n/g,'<br>');
          el.classList.add('egg-alert');
        }else{
          el.remove(); // ไม่มีข้อความ → ไม่แสดง
        }
      });
    }catch(e){
      // เงียบไว้ ไม่ให้รบกวนหน้า
      // console.warn('alert-loader:', e);
    }
  }
  document.addEventListener('DOMContentLoaded', loadAlerts);
})();
