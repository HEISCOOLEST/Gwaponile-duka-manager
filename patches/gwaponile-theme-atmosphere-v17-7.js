(function(){
  'use strict';
  const symbols={
    ocean:['✦','◦','💧'], emerald:['🍃','✦','❖'], royal:['✦','◆','♛'], sunset:['✦','☀','◦'], mint:['🍃','✦','◦'],
    christmas:['❄','✦','🎄'], newyear:['✦','✧','🎆'], eidfitr:['☾','✦','✨'], eidadha:['☾','✦','🕌'],
    ramadan:['☾','✦','✧'], easter:['🌸','✦','◦'], independence:['✦','🇹🇿','◦'], union:['✦','🇹🇿','◦'],
    sabasaba:['✦','🏭','◦'], naneNane:['🌾','✦','◦'], labour:['✦','🛠️','◦'], revolution:['✦','🇹🇿','◦'],
    karume:['✦','🇹🇿','◦'], nyerere:['🌿','✦','◦'], boxing:['🎁','✦','◦'], maulid:['☾','✦','◦'],
    valentine:['❤','♥','✦']
  };
  const holidayThemes=new Set(Object.keys(symbols).filter(k=>!['ocean','emerald','royal','sunset','mint'].includes(k)));
  let root=null, lastTheme='';

  function make(){
    const theme=document.documentElement.getAttribute('data-theme')||'ocean';
    if(theme===lastTheme && root) return;
    lastTheme=theme;
    if(root) root.remove();
    root=document.createElement('div');
    root.className='gw-theme-atmosphere';
    root.setAttribute('aria-hidden','true');
    const list=symbols[theme]||symbols.ocean;
    const count=holidayThemes.has(theme)?7:4;
    for(let i=0;i<count;i++){
      const el=document.createElement('span');
      el.className='gw-float';
      el.textContent=list[i%list.length];
      el.style.setProperty('--x',(8+Math.random()*84).toFixed(1)+'%');
      el.style.setProperty('--y',(18+Math.random()*70).toFixed(1)+'%');
      el.style.setProperty('--size',(12+Math.random()*10).toFixed(0)+'px');
      el.style.setProperty('--dur',(12+Math.random()*10).toFixed(1)+'s');
      el.style.setProperty('--delay',(-Math.random()*10).toFixed(1)+'s');
      root.appendChild(el);
    }
    for(let i=0;i<2;i++){
      const g=document.createElement('span');
      g.className='gw-glow';
      g.style.setProperty('--x',(i?70:5)+'%');
      g.style.setProperty('--y',(i?8:65)+'%');
      g.style.animationDelay=(-Math.random()*4).toFixed(1)+'s';
      root.appendChild(g);
    }
    document.body.appendChild(root);
  }

  function init(){
    make();
    const observer=new MutationObserver(function(mutations){
      for(const m of mutations){
        if(m.type==='attributes' && m.attributeName==='data-theme'){
          make();
          break;
        }
      }
    });
    observer.observe(document.documentElement,{attributes:true,attributeFilter:['data-theme']});
  }
  if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',init,{once:true});
  else init();
})();
