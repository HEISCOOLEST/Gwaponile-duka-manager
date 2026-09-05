(function(){
  function role(){ return window.state?.profile?.role || window.state?.user?.role || ""; }
  function canSee(){ return role()==="Administrator" || role()==="Manager"; }
  function esc(v){
    return String(v??"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[m]));
  }
  function money(v){
    try{
      if(typeof window.fmt==="function") return window.fmt(v||0);
    }catch(e){}
    return Number(v||0).toLocaleString();
  }
  window.renderAdminShiftHistory = function(){
    if(!canSee()) return;
    const root=document.getElementById("adminShiftHistoryRoot");
    if(!root) return;
    const shifts=Array.isArray(window.state?.shifts)?window.state.shifts.slice():[];
    shifts.sort((a,b)=>String(b.startTime||b.date||"").localeCompare(String(a.startTime||a.date||"")));
    if(!shifts.length){
      root.innerHTML='<div class="shift-history-empty">Hakuna historia ya shift bado.</div>';
      return;
    }
    root.innerHTML=`
      <div class="card">
        <div style="font-size:16px;font-weight:900">Historia ya Shift</div>
        <div class="shift-history-admin-note">Administrator/Manager wanaweza kuona shift zote za wafanyakazi.</div>
        <div style="overflow:auto;margin-top:12px">
          <table class="shift-history-table">
            <thead><tr>
              <th>Cashier</th><th>Shift</th><th>Kuanza</th><th>Kufunga</th><th>Start Cash</th><th>End Cash</th><th>Difference</th><th>Status</th>
            </tr></thead>
            <tbody>
              ${shifts.map(x=>`<tr>
                <td>${esc(x.cashierName||x.cashierUid||"—")}</td>
                <td>${esc(x.id||"—")}</td>
                <td>${esc(x.startTime||"—")}</td>
                <td>${esc(x.endTime||"—")}</td>
                <td>${money(x.startCash)}</td>
                <td>${x.status==="closed"?money(x.endCash):"—"}</td>
                <td>${x.status==="closed"?money(x.difference):"—"}</td>
                <td>${esc(x.status||"—")}</td>
              </tr>`).join("")}
            </tbody>
          </table>
        </div>
      </div>`;
  };

  function addHistoryPanel(){
    if(!canSee() || document.getElementById("adminShiftHistoryRoot")) return;
    const host=document.querySelector("main") || document.querySelector(".app-content") || document.body;
    if(!host) return;
    const d=document.createElement("div");
    d.id="adminShiftHistoryRoot";
    d.style.marginTop="16px";
    d.style.display="none";
    host.appendChild(d);
    window.renderAdminShiftHistory();
  }

  function ensureShiftHistoryNav(){
    if(!canSee()) return;
    addHistoryPanel();
    const candidates=[...document.querySelectorAll("button,a,[role=button]")];
    let nav=candidates.find(el=>/shift history|historia ya shift|shift history/i.test(el.textContent||""));
    if(nav && !nav.dataset.adminShiftHistoryFixed){
      nav.dataset.adminShiftHistoryFixed="1";
      nav.addEventListener("click",function(){
        const r=document.getElementById("adminShiftHistoryRoot");
        if(r){ r.style.display="block"; window.renderAdminShiftHistory(); }
      });
    }
  }
  const oldRender=window.render;
  if(typeof oldRender==="function" && !window.__adminShiftRenderWrapped){
    window.__adminShiftRenderWrapped=true;
    window.render=function(){
      const out=oldRender.apply(this,arguments);
      setTimeout(ensureShiftHistoryNav,0);
      return out;
    };
  }
  setTimeout(ensureShiftHistoryNav,800);
})();
