(function(){
  async function ensureAdminShiftData(){
    try{
      if(!window.state || !window.state.profile) return;
      const role=window.state.profile.role;
      if(role!=="Administrator" && role!=="Manager") return;
      if(window.state.shifts && window.state.shifts.length) return;
      if(!window.db || !window.collection || !window.getDocs) return;
      const snap=await window.getDocs(window.collection(window.db,"shifts"));
      window.state.shifts=snap.docs.map(d=>({id:d.id,...d.data()}));
      if(typeof window.render==="function") window.render();
    }catch(e){
      console.warn("Shift report read:", e);
    }
  }
  // The module code normally owns these symbols, so the live listener is the
  // primary path. This fallback runs only when the Admin/Manager report is empty.
  document.addEventListener("click",function(e){
    const el=e.target.closest?.("[data-tab='shiftsreport']");
    if(el) setTimeout(ensureAdminShiftData,250);
  });
  setInterval(()=>{
    if(window.state?.tab==="shiftsreport" && (!window.state.shifts || !window.state.shifts.length)){
      ensureAdminShiftData();
    }
  },2000);
})();
