/* RightOwner purchase landing page.
   Microsoft sends buyers here after they buy RightOwner on AppSource / in Teams:
   https://app.rubuild.com/landing.html?token=...
   The buyer signs in with Microsoft, checks the details and clicks Activate. */
(function () {
  "use strict";
  const CFG = window.RIGHTOWNER_CONFIG || {};
  const panel = document.getElementById("panel");
  const token = new URLSearchParams(location.search).get("token") || "";
  const PLAN_NAMES = { pro: "RightOwner Pro", business: "RightOwner Business", free: "RightOwner" };
  const TEAMS_LINK = "https://teams.microsoft.com/l/app/" + (CFG.teamsAppId || "");
  let pca = null;
  let idToken = "";
  let account = null;

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  function view(html) { panel.innerHTML = html; }
  function loading(msg) { view('<div class="spinner" aria-hidden="true"></div><p>' + esc(msg) + "</p>"); }
  function errorView(msg, retry) {
    view('<h2>Something needs attention</h2><p class="err">' + esc(msg) + '</p><div class="actions">' +
      (retry ? '<button class="btn primary" id="retry" type="button">Try again</button>' : "") +
      '<a class="btn" href="mailto:' + esc(CFG.supportEmail || "support@rubuild.com") + '">Email support</a></div>');
    const b = document.getElementById("retry");
    if (b) b.addEventListener("click", retry);
  }
  function fmtDate(s) {
    if (!s) return "";
    const d = new Date(s);
    return isNaN(d) ? s : d.toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" });
  }
  function termText(t) {
    if (!t) return "";
    const unit = t.unit === "P1Y" ? "Yearly" : t.unit === "P1M" ? "Monthly" : (t.unit || "");
    return unit + (t.start ? " · " + fmtDate(t.start) + " to " + fmtDate(t.end) : "");
  }

  async function api(path, body) {
    const res = await fetch(CFG.billingUrl.replace(/\/$/, "") + path, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.message || "Something went wrong. Please try again.");
    return data;
  }

  async function initMsal() {
    pca = await msal.createStandardPublicClientApplication({
      auth: { clientId: CFG.clientId, authority: "https://login.microsoftonline.com/organizations", redirectUri: location.origin + "/redirect.html" },
      cache: { cacheLocation: "sessionStorage" }
    });
    account = pca.getActiveAccount() || pca.getAllAccounts()[0] || null;
  }

  async function signIn(interactive) {
    const req = { scopes: ["openid", "profile", "User.Read"], account: account || undefined };
    let r;
    if (!interactive && account) {
      try { r = await pca.acquireTokenSilent(req); } catch (e) { r = null; }
    }
    if (!r && interactive) r = await pca.loginPopup({ scopes: req.scopes, prompt: "select_account" });
    if (!r) return false;
    account = r.account;
    pca.setActiveAccount(account);
    idToken = r.idToken;
    return true;
  }

  function signInView() {
    view('<h2>Welcome to RightOwner</h2>' +
      '<p>Thanks for your purchase. Sign in with your Microsoft work account to confirm and activate your subscription. It takes about a minute.</p>' +
      '<div class="actions"><button class="btn primary" id="signin" type="button">Sign in with Microsoft</button></div>');
    document.getElementById("signin").addEventListener("click", async () => {
      try {
        if (await signIn(true)) await showDetails();
      } catch (e) {
        if (!/user_cancelled/i.test(String(e && (e.errorCode || e.message)))) errorView("Sign-in didn't complete. Please allow pop-ups for this page and try again.", signInView);
      }
    });
  }

  async function showDetails() {
    loading("Checking your purchase with Microsoft…");
    let d;
    try { d = await api("/api/resolve", { token, idToken }); }
    catch (e) { return errorView(e.message, showDetails); }

    if (d.status === "Subscribed") return done(d);
    if (d.status === "Unsubscribed") return errorView("This subscription has been cancelled. To buy again, open RightOwner on Microsoft AppSource.", null);
    if (d.status === "Suspended") return errorView("This subscription is suspended, usually because of a payment problem. Please check Billing in the Microsoft 365 admin center.", null);

    view('<h2>Confirm your subscription</h2>' +
      '<p>Check these details, then click Activate. Microsoft will bill your organization.</p>' +
      '<dl class="details">' +
      "<dt>Plan</dt><dd>" + esc(PLAN_NAMES[d.plan] || d.planId) + "</dd>" +
      (d.term ? "<dt>Billing</dt><dd>" + esc(termText(d.term)) + "</dd>" : "") +
      "<dt>Subscription</dt><dd>" + esc(d.subscriptionName) + "</dd>" +
      "<dt>Purchased by</dt><dd>" + esc(d.purchaser && d.purchaser.email) + "</dd>" +
      "<dt>Signed in as</dt><dd>" + esc(d.signedInName ? d.signedInName + " (" + d.signedInAs + ")" : d.signedInAs) + "</dd>" +
      "</dl>" +
      (d.beneficiary && d.signedInTenant && d.beneficiary.tenantId !== d.signedInTenant
        ? '<p class="err">Note: you are signed in to a different organization from the one this subscription is for. The plan will apply to the organization that bought it.</p>' : "") +
      '<div class="actions"><button class="btn primary" id="activate" type="button">Activate</button>' +
      '<button class="btn" id="switch" type="button">Use a different account</button></div>');

    document.getElementById("activate").addEventListener("click", activate);
    document.getElementById("switch").addEventListener("click", async () => {
      try { if (await signIn(true)) await showDetails(); } catch (e) { /* stay on page */ }
    });
  }

  async function activate() {
    loading("Activating your subscription…");
    try {
      const d = await api("/api/activate", { token, idToken });
      done(d);
    } catch (e) {
      errorView(e.message, activate);
    }
  }

  function done(d) {
    view('<div class="ok" aria-hidden="true">✓</div><h2>You\'re all set</h2>' +
      "<p>" + esc(PLAN_NAMES[d.plan] || "RightOwner") + " is active for your organization. Everyone in your company can use it straight away; there is nothing to install per person.</p>" +
      "<p><strong>Next:</strong> open RightOwner in Microsoft Teams. If your organization hasn't added it yet, a Teams admin can add it from Teams admin center → Manage apps.</p>" +
      '<div class="actions"><a class="btn primary" href="' + esc(TEAMS_LINK) + '" target="_blank" rel="noopener">Open RightOwner in Teams</a>' +
      '<a class="btn" href="https://admin.microsoft.com/adminportal/home#/subscriptions" target="_blank" rel="noopener">Manage billing</a></div>');
  }

  async function startPage() {
    if (!CFG.billingUrl) return errorView("Purchases aren't switched on yet. Please email support and we'll help you.", null);
    if (!token) {
      return view('<h2>Manage your RightOwner subscription</h2>' +
        '<p>This page is opened automatically after you buy RightOwner on Microsoft AppSource or in Teams.</p>' +
        '<p>To change or cancel your plan, use the Microsoft 365 admin center → Billing → Your products.</p>' +
        '<div class="actions"><a class="btn primary" href="https://admin.microsoft.com/adminportal/home#/subscriptions" target="_blank" rel="noopener">Open Microsoft 365 admin center</a>' +
        '<a class="btn" href="' + esc(TEAMS_LINK) + '" target="_blank" rel="noopener">Open RightOwner in Teams</a></div>');
    }
    try { await initMsal(); } catch (e) { return errorView("Microsoft sign-in couldn't start. Please reload the page.", () => location.reload()); }
    if (await signIn(false).catch(() => false)) return showDetails();
    signInView();
  }

  if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) document.body.classList.add("theme-dark");
  startPage();
})();
