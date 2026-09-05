(function(){
  function refreshMotion(){
    try {
      document.querySelectorAll('.view').forEach((el,i)=>{
        el.style.setProperty('--gw-view-index', String(i));
      });
      const btn=document.getElementById('v15NotifyBtn');
      if(btn){
        const has=btn.querySelector('.v15-notify-badge, .notify-badge, .badge');
        if(has) btn.classList.add('has-alert'); else btn.classList.remove('has-alert');
      }
    } catch(e) { console.debug('GW motion enhancement:',e); }
  }
  window.addEventListener('load', refreshMotion);
  const oldRender=window.render;
  if(typeof oldRender==='function'){
    window.render=function(){
      const result=oldRender.apply(this,arguments);
      requestAnimationFrame(refreshMotion);
      return result;
    };
  }
  document.addEventListener('click', function(e){
    const b=e.target.closest && e.target.closest('.btn,.theme-toggle,.icon-btn,.qty-btn');
    if(!b) return;
    b.classList.remove('gw-click');
    void b.offsetWidth;
    b.classList.add('gw-click');
    setTimeout(()=>b.classList.remove('gw-click'),180);
  });
})();
