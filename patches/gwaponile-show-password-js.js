(function(){
  const eye = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2 12s3.5-6 10-6 10 6 10 6-3.5 6-10 6S2 12 2 12Z"/><circle cx="12" cy="12" r="2.5"/></svg>';
  const eyeOff = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m3 3 18 18"/><path d="M10.6 10.6a2 2 0 0 0 2.8 2.8"/><path d="M9.9 5.2A10.8 10.8 0 0 1 12 5c6.5 0 10 7 10 7a18.2 18.2 0 0 1-3.1 3.8"/><path d="M6.1 6.1C3.4 8.2 2 12 2 12s3.5 7 10 7c1.1 0 2.1-.2 3-.5"/></svg>';

  function addControls(root){
    const fields = (root || document).querySelectorAll('input[type="password"]');
    fields.forEach(input => {
      if(input.dataset.gwShowPasswordReady === "1") return;
      input.dataset.gwShowPasswordReady = "1";

      const parent = input.parentElement;
      if(!parent) return;

      // Keep existing layout where possible; only wrap if not already in a suitable wrapper.
      let wrap = parent;
      if(!parent.classList.contains("gw-show-password-wrap")){
        wrap = document.createElement("div");
        wrap.className = "gw-show-password-wrap";
        parent.insertBefore(wrap, input);
        wrap.appendChild(input);
      }

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "gw-show-password-btn";
      btn.setAttribute("aria-label", "Show password");
      btn.setAttribute("title", "Show password");
      btn.innerHTML = eye;

      btn.addEventListener("click", function(){
        const showing = input.type === "text";
        input.type = showing ? "password" : "text";
        btn.innerHTML = showing ? eye : eyeOff;
        btn.setAttribute("aria-label", showing ? "Show password" : "Hide password");
        btn.setAttribute("title", showing ? "Show password" : "Hide password");
        input.focus({preventScroll:true});
      });

      wrap.appendChild(btn);
    });
  }

  // Initial fields + fields created later by the SPA.
  addControls(document);
  const observer = new MutationObserver(() => addControls(document));
  observer.observe(document.documentElement, {childList:true, subtree:true});
})();
