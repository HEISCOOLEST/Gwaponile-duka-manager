(function(){
  function addNav(){
    const st=window.state, role=st?.profile?.role||"";
    if(role!=="Administrator"&&role!=="Manager") return;
    if(document.getElementById("adminShiftHistoryNavBtn")) return;
    const host=[...document.querySelectorAll("nav,aside,.sidebar,.app-sidebar")][0];
    if(!host) return;
    const b=document.createElement("button");
    b.id="adminShiftHistoryNavBtn";
    b.className="nav-btn";
    b.type="button";
    b.innerHTML="🕘 <span>Shift History</span>";
    b.onclick=function(){
      if(typeof window.renderAdminShiftHistory==="function") window.renderAdminShiftHistory();
      const r=document.getElementById("adminShiftHistoryRoot");
      if(r){r.style.display="block";r.scrollIntoView({behavior:"smooth",block:"start"});}
    };
    host.appendChild(b);
  }
  setInterval(addNav,1200);
})();
