/* RightOwner — Teams personal app. Runs entirely in the browser.
   Data lives in a SharePoint list in the customer's own tenant, reached via Microsoft Graph. */
(function () {
  "use strict";

  const CFG = window.RIGHTOWNER_CONFIG || {};
  const SCOPES = ["User.Read", "User.Read.All", "Sites.Manage.All"];
  const GRAPH = "https://graph.microsoft.com/v1.0";
  const FIELDS = ["Title", "Category", "Description", "OwnerEmail", "OwnerName", "BackupEmail", "BackupName"];

  const state = {
    inTeams: false,
    ctx: null,
    pca: null,
    account: null,
    siteId: null,
    listId: null,
    items: [],          // { id, fields }
    people: {},         // lowercased email -> { found, enabled, displayName, jobTitle }
    filter: "all",
    query: "",
    editing: null
  };

  const $ = (id) => document.getElementById(id);

  // ---------- helpers ----------
  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  function initials(name, email) {
    const src = (name || email || "?").trim();
    const parts = src.split(/[\s._@-]+/).filter(Boolean);
    return ((parts[0] || "?")[0] + (parts[1] ? parts[1][0] : "")).toUpperCase();
  }
  function toast(msg, ms) {
    const t = $("toast");
    t.textContent = msg;
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => { t.hidden = true; }, ms || 3200);
  }
  function show(section) {
    ["loading", "message", "directory"].forEach((id) => { $(id).hidden = id !== section; });
    $("topActions").hidden = section !== "directory";
  }
  function showMessage(title, body, actions) {
    $("msgTitle").textContent = title;
    $("msgBody").textContent = body;
    const box = $("msgActions");
    box.innerHTML = "";
    (actions || []).forEach((a) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "btn" + (a.primary ? " primary" : "");
      b.textContent = a.label;
      b.addEventListener("click", a.onClick);
      box.appendChild(b);
    });
    show("message");
  }
  function openExternal(url) {
    if (state.inTeams && window.microsoftTeams && microsoftTeams.app && microsoftTeams.app.openLink) {
      microsoftTeams.app.openLink(url).catch(() => window.open(url, "_blank", "noopener"));
    } else {
      window.open(url, "_blank", "noopener");
    }
  }
  function adminConsentUrl() {
    return "https://login.microsoftonline.com/organizations/adminconsent?client_id=" + encodeURIComponent(CFG.clientId) +
      "&redirect_uri=" + encodeURIComponent(location.origin + "/redirect.html");
  }

  // ---------- theme ----------
  function applyTheme(theme) {
    document.body.classList.remove("theme-dark", "theme-contrast");
    if (theme === "dark") document.body.classList.add("theme-dark");
    if (theme === "contrast") document.body.classList.add("theme-contrast");
  }

  // ---------- auth ----------
  class GraphError extends Error {
    constructor(status, code, message) { super(message); this.status = status; this.code = code; }
  }

  async function initAuth() {
    try {
      await microsoftTeams.app.initialize();
      state.inTeams = true;
      document.body.classList.add("in-teams"); // Teams already shows the app name and icon
      state.ctx = await microsoftTeams.app.getContext();
      applyTheme(state.ctx && state.ctx.app && state.ctx.app.theme);
      microsoftTeams.app.registerOnThemeChangeHandler(applyTheme);
    } catch (e) {
      state.inTeams = false;
      if (window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches) applyTheme("dark");
    }

    const config = {
      auth: {
        clientId: CFG.clientId,
        authority: "https://login.microsoftonline.com/organizations",
        redirectUri: location.origin + "/redirect.html"
      },
      cache: { cacheLocation: "localStorage" }
    };

    if (state.inTeams) {
      state.pca = await msal.createNestablePublicClientApplication(config);
    } else {
      state.pca = await msal.createStandardPublicClientApplication(config);
    }
    const accounts = state.pca.getAllAccounts();
    state.account = state.pca.getActiveAccount() || accounts[0] || null;
  }

  function loginHint() {
    return state.ctx && state.ctx.user ? (state.ctx.user.loginHint || state.ctx.user.userPrincipalName) : undefined;
  }

  async function getToken(interactive) {
    const req = { scopes: SCOPES };
    if (state.account) req.account = state.account;
    const hint = loginHint();
    if (hint) req.loginHint = hint;

    if (!interactive) {
      try {
        const r = await state.pca.acquireTokenSilent(req);
        state.account = r.account;
        state.pca.setActiveAccount(r.account);
        return r.accessToken;
      } catch (e) {
        // Silent sign-in failed (no session yet, consent needed, etc.) — fall back to a popup below.
      }
    }
    const r = await state.pca.acquireTokenPopup(req);
    state.account = r.account;
    state.pca.setActiveAccount(r.account);
    return r.accessToken;
  }

  async function graph(method, path, body, extraHeaders) {
    let token;
    try {
      token = await getToken(false);
    } catch (e) {
      throw markAuthError(e);
    }
    const headers = Object.assign({ Authorization: "Bearer " + token }, extraHeaders || {});
    if (body !== undefined) headers["Content-Type"] = "application/json";
    const url = path.startsWith("http") ? path : GRAPH + path;

    for (let attempt = 0; attempt < 4; attempt++) {
      const res = await fetch(url, { method, headers, body: body !== undefined ? JSON.stringify(body) : undefined });
      if (res.status === 429 || res.status === 503) {
        const wait = Math.min(parseInt(res.headers.get("Retry-After") || "2", 10), 10) * 1000;
        await new Promise((r) => setTimeout(r, wait));
        continue;
      }
      if (res.status === 204) return null;
      const text = await res.text();
      const data = text ? JSON.parse(text) : null;
      if (!res.ok) {
        const err = (data && data.error) || {};
        throw new GraphError(res.status, err.code, err.message || ("Request failed (" + res.status + ")"));
      }
      return data;
    }
    throw new GraphError(429, "throttled", "Microsoft 365 is busy right now. Please try again in a minute.");
  }

  function markAuthError(e) {
    const msg = (e && (e.errorMessage || e.message)) || "";
    const code = (e && (e.errorCode || e.code)) || "";
    const err = new Error(msg || "Sign-in failed");
    err.auth = true;
    err.consent = /consent|AADSTS65001|AADSTS90094|admin/i.test(msg + " " + code);
    err.cancelled = /user_cancelled|cancel/i.test(code + " " + msg);
    return err;
  }

  // ---------- SharePoint list ----------
  async function findList() {
    const site = await graph("GET", "/sites/root?$select=id,webUrl");
    state.siteId = site.id;
    let url = "/sites/" + state.siteId + "/lists?$select=id,displayName&$top=200";
    while (url) {
      const page = await graph("GET", url);
      const hit = (page.value || []).find((l) => l.displayName === CFG.listName);
      if (hit) { state.listId = hit.id; return true; }
      url = page["@odata.nextLink"] || null;
    }
    return false;
  }

  async function createList() {
    const body = {
      displayName: CFG.listName,
      description: "Owners and backups for company systems, processes, vendors and budgets. Managed by the RightOwner Teams app.",
      columns: [
        { name: "Category", text: {} },
        { name: "Description", text: { allowMultipleLines: true, linesForEditing: 4 } },
        { name: "OwnerEmail", text: {} },
        { name: "OwnerName", text: {} },
        { name: "BackupEmail", text: {} },
        { name: "BackupName", text: {} }
      ],
      list: { template: "genericList" }
    };
    const list = await graph("POST", "/sites/" + state.siteId + "/lists", body);
    state.listId = list.id;
  }

  async function loadItems() {
    const items = [];
    let url = "/sites/" + state.siteId + "/lists/" + state.listId +
      "/items?$expand=fields($select=" + FIELDS.join(",") + ")&$top=500";
    while (url) {
      const page = await graph("GET", url);
      (page.value || []).forEach((it) => items.push({ id: it.id, createdBy: it.createdBy, fields: it.fields || {} }));
      url = page["@odata.nextLink"] || null;
    }
    items.sort((a, b) => String(a.fields.Title || "").localeCompare(String(b.fields.Title || "")));
    state.items = items;
  }

  // Look up owners/backups in bulk to spot people who have left.
  async function checkPeople(emails) {
    const todo = Array.from(new Set(emails.filter(Boolean).map((e) => e.toLowerCase())))
      .filter((e) => !(e in state.people));
    for (let i = 0; i < todo.length; i += 20) {
      const chunk = todo.slice(i, i + 20);
      const body = {
        requests: chunk.map((email, n) => ({
          id: String(n),
          method: "GET",
          url: "/users/" + encodeURIComponent(email) + "?$select=id,displayName,jobTitle,accountEnabled,mail,userPrincipalName"
        }))
      };
      try {
        const res = await graph("POST", "/$batch", body);
        (res.responses || []).forEach((r) => {
          const email = chunk[parseInt(r.id, 10)];
          if (r.status === 200) {
            state.people[email] = { found: true, enabled: r.body.accountEnabled !== false, displayName: r.body.displayName, jobTitle: r.body.jobTitle, id: r.body.id, upn: String(r.body.userPrincipalName || "").toLowerCase() };
          } else if (r.status === 404) {
            state.people[email] = { found: false, enabled: false };
          }
        });
      } catch (e) {
        // If the lookup fails we simply don't flag anyone.
      }
    }
  }

  function needsAttention(f) {
    const p = f.OwnerEmail ? state.people[String(f.OwnerEmail).toLowerCase()] : null;
    if (!f.OwnerEmail) return "No owner assigned";
    if (p && !p.found) return "Owner not found in your organization";
    if (p && !p.enabled) return "Owner's account is disabled";
    return "";
  }

  // ---------- rendering ----------
  function personBlock(role, email, name) {
    if (!email) {
      return '<div class="person"><div class="avatar">–</div><div class="person-info"><div class="role">' + esc(role) +
        '</div><div class="pmeta">Not set</div></div></div>';
    }
    const p = state.people[String(email).toLowerCase()] || {};
    const display = p.displayName || name || email;
    const meta = [p.jobTitle, email].filter(Boolean).join(" · ");
    return '<div class="person"><div class="avatar">' + esc(initials(display, email)) + '</div><div class="person-info"><div class="role">' +
      esc(role) + '</div><div class="pname">' + esc(display) + '</div><div class="pmeta">' + esc(meta) + "</div></div></div>";
  }

  function matches(item, q) {
    if (!q) return true;
    const f = item.fields;
    const owner = state.people[String(f.OwnerEmail || "").toLowerCase()] || {};
    const backup = state.people[String(f.BackupEmail || "").toLowerCase()] || {};
    const hay = [f.Title, f.Category, f.Description, f.OwnerEmail, f.OwnerName, owner.displayName, f.BackupEmail, f.BackupName, backup.displayName]
      .join(" ").toLowerCase();
    return q.split(/\s+/).filter(Boolean).every((w) => hay.includes(w));
  }

  function render() {
    const q = state.query.trim().toLowerCase();
    const attentionCount = state.items.filter((i) => needsAttention(i.fields)).length;
    $("countAll").textContent = state.items.length;
    $("countAttention").textContent = attentionCount;

    const visible = state.items.filter((i) => matches(i, q) && (state.filter === "all" || needsAttention(i.fields)));
    const cards = $("cards");
    cards.innerHTML = visible.map((it) => {
      const f = it.fields;
      const flag = needsAttention(f);
      const actions = [];
      if (f.OwnerEmail) {
        actions.push('<button class="link-btn" data-act="chat" data-email="' + esc(f.OwnerEmail) + '">Chat with owner</button>');
        actions.push('<button class="link-btn" data-act="mail" data-email="' + esc(f.OwnerEmail) + '" data-title="' + esc(f.Title) + '">Email</button>');
      }
      actions.push('<button class="link-btn edit" data-act="edit" data-id="' + esc(it.id) + '">Edit</button>');
      return '<article class="card' + (flag ? " attention" : "") + '">' +
        '<div class="card-head"><h3>' + esc(f.Title || "(untitled)") + "</h3>" +
        (f.Category ? '<span class="cat">' + esc(f.Category) + "</span>" : "") + "</div>" +
        (f.Description ? '<p class="desc">' + esc(f.Description) + "</p>" : "") +
        (flag ? '<span class="flag">⚠ ' + esc(flag) + "</span>" : "") +
        personBlock("Owner", f.OwnerEmail, f.OwnerName) +
        personBlock("Backup", f.BackupEmail, f.BackupName) +
        '<div class="card-actions">' + actions.join("") + "</div></article>";
    }).join("");

    const empty = $("empty");
    if (visible.length === 0) {
      empty.hidden = false;
      if (state.items.length === 0) {
        $("emptyTitle").textContent = "No items yet";
        $("emptyBody").textContent = "Add your first item, or use More → Import from CSV to load a spreadsheet of owners.";
      } else if (state.filter === "attention" && !q) {
        $("emptyTitle").textContent = "All owners are active";
        $("emptyBody").textContent = "Nothing needs a new owner right now.";
      } else {
        $("emptyTitle").textContent = "No matches";
        $("emptyBody").textContent = "Try a different word, or add this item so the next person finds it.";
      }
    } else {
      empty.hidden = true;
    }

    const cats = Array.from(new Set(state.items.map((i) => i.fields.Category).filter(Boolean))).sort();
    $("categoryList").innerHTML = cats.map((c) => '<option value="' + esc(c) + '">').join("");
  }

  // ---------- people picker ----------
  async function searchPeople(term) {
    const t = term.replace(/"/g, "").trim();
    if (t.length < 2) return [];
    const path = "/users?$search=" + encodeURIComponent('"displayName:' + t + '" OR "mail:' + t + '"') +
      "&$select=displayName,mail,userPrincipalName,jobTitle,accountEnabled&$top=8";
    const res = await graph("GET", path, undefined, { ConsistencyLevel: "eventual" });
    return (res.value || []).filter((u) => u.accountEnabled !== false && (u.mail || u.userPrincipalName));
  }

  function setupPersonField(field) {
    const input = field.querySelector(".person-input");
    const list = field.querySelector(".suggest");
    let timer = null;
    let results = [];
    let active = -1;

    function choose(u) {
      const email = u.mail || u.userPrincipalName;
      field.dataset.email = email;
      field.dataset.name = u.displayName || email;
      input.value = (u.displayName || email) + " <" + email + ">";
      list.hidden = true;
      state.people[email.toLowerCase()] = { found: true, enabled: true, displayName: u.displayName, jobTitle: u.jobTitle };
    }
    function paint() {
      list.innerHTML = results.map((u, i) =>
        '<li data-i="' + i + '" class="' + (i === active ? "active" : "") + '"><div>' + esc(u.displayName) +
        '</div><div class="s-meta">' + esc([u.jobTitle, u.mail || u.userPrincipalName].filter(Boolean).join(" · ")) + "</div></li>").join("");
      list.hidden = results.length === 0;
    }
    input.addEventListener("input", () => {
      field.dataset.email = "";
      field.dataset.name = "";
      clearTimeout(timer);
      const term = input.value;
      timer = setTimeout(async () => {
        try { results = await searchPeople(term); } catch (e) { results = []; }
        active = -1;
        paint();
      }, 250);
    });
    input.addEventListener("keydown", (e) => {
      if (list.hidden) return;
      if (e.key === "ArrowDown") { active = Math.min(active + 1, results.length - 1); paint(); e.preventDefault(); }
      else if (e.key === "ArrowUp") { active = Math.max(active - 1, 0); paint(); e.preventDefault(); }
      else if (e.key === "Enter" && active >= 0) { choose(results[active]); e.preventDefault(); }
      else if (e.key === "Escape") { list.hidden = true; e.stopPropagation(); e.preventDefault(); }
    });
    list.addEventListener("mousedown", (e) => {
      const li = e.target.closest("li");
      if (li) { choose(results[parseInt(li.dataset.i, 10)]); e.preventDefault(); }
    });
    input.addEventListener("blur", () => setTimeout(() => { list.hidden = true; }, 150));
  }

  function setPerson(field, email, name) {
    field.dataset.email = email || "";
    field.dataset.name = name || "";
    field.querySelector(".person-input").value = email ? ((name || email) + " <" + email + ">") : "";
  }

  function personValue(field) {
    const email = (field.dataset.email || "").trim();
    if (email) return { email, name: field.dataset.name || email };
    // Allow typing a plain email address.
    const typed = field.querySelector(".person-input").value.trim();
    const m = typed.match(/[^\s<>]+@[^\s<>]+\.[^\s<>]+/);
    return m ? { email: m[0], name: m[0], typed: true } : { email: "", name: "", invalid: typed.length > 0 };
  }

  // ---------- dialog ----------
  function openDialog(item) {
    state.editing = item || null;
    const form = $("itemForm");
    form.reset();
    $("formError").hidden = true;
    $("dlgTitle").textContent = item ? "Edit entry" : "Add to the directory";
    $("deleteBtn").hidden = !item;
    const f = item ? item.fields : {};
    form.elements.Title.value = f.Title || "";
    form.elements.Category.value = f.Category || "";
    form.elements.Description.value = f.Description || "";
    const [ownerField, backupField] = form.querySelectorAll(".person-field");
    setPerson(ownerField, f.OwnerEmail, f.OwnerName);
    setPerson(backupField, f.BackupEmail, f.BackupName);
    $("itemDialog").showModal();
    form.elements.Title.focus();
  }

  function formError(msg) {
    const el = $("formError");
    el.textContent = msg;
    el.hidden = false;
  }

  async function saveItem(e) {
    e.preventDefault();
    const form = $("itemForm");
    const [ownerField, backupField] = form.querySelectorAll(".person-field");
    const owner = personValue(ownerField);
    const backup = personValue(backupField);
    const title = form.elements.Title.value.trim();
    if (!title) return formError("Please enter what this item is.");
    if (!owner.email) return formError(owner.invalid ? "Pick the owner from the suggestions, or type a full email address." : "Please choose an owner.");
    if (backup.invalid) return formError("Pick the backup from the suggestions, or type a full email address.");
    if (backup.email && backup.email.toLowerCase() === owner.email.toLowerCase()) return formError("The backup should be a different person from the owner.");

    const fields = {
      Title: title,
      Category: form.elements.Category.value.trim(),
      Description: form.elements.Description.value.trim(),
      OwnerEmail: owner.email,
      OwnerName: owner.name,
      BackupEmail: backup.email,
      BackupName: backup.name
    };
    const btn = $("saveBtn");
    btn.disabled = true;
    try {
      const base = "/sites/" + state.siteId + "/lists/" + state.listId + "/items";
      if (state.editing) {
        await graph("PATCH", base + "/" + state.editing.id + "/fields", fields);
        state.editing.fields = Object.assign({}, state.editing.fields, fields);
      } else {
        const created = await graph("POST", base, { fields });
        state.items.push({ id: created.id, mine: true, fields: Object.assign({}, fields) });
        state.items.sort((a, b) => String(a.fields.Title || "").localeCompare(String(b.fields.Title || "")));
      }
      await checkPeople([fields.OwnerEmail, fields.BackupEmail]);
      $("itemDialog").close();
      render();
      toast(state.editing ? "Saved" : "Added “" + title + "”");
    } catch (err) {
      formError(friendlyWriteError(err));
    } finally {
      btn.disabled = false;
    }
  }


  // Teams (desktop and mobile) blocks the browser's built-in confirm() box, so we use our own small dialog.
  function askConfirm(message, okLabel, danger) {
    return new Promise((resolve) => {
      const d = document.createElement("dialog");
      d.className = "confirm-dlg";
      d.setAttribute("aria-label", "Please confirm");
      const p = document.createElement("p");
      p.textContent = message;
      const row = document.createElement("div");
      row.className = "dlg-actions";
      const spacer = document.createElement("span");
      spacer.className = "spacer";
      const no = document.createElement("button");
      no.type = "button"; no.className = "btn"; no.textContent = "Cancel";
      const yes = document.createElement("button");
      yes.type = "button"; yes.className = "btn " + (danger ? "danger-solid" : "primary"); yes.textContent = okLabel || "OK";
      row.append(spacer, no, yes);
      d.append(p, row);
      document.body.appendChild(d);
      let done = false;
      const finish = (v) => { if (done) return; done = true; try { d.close(); } catch (e) {} d.remove(); resolve(v); };
      no.addEventListener("click", () => finish(false));
      yes.addEventListener("click", () => finish(true));
      d.addEventListener("cancel", (e) => { e.preventDefault(); finish(false); });
      try { d.showModal(); } catch (e) { d.setAttribute("open", ""); }
      yes.focus();
    });
  }


  // Who may delete an entry: its owner, its backup, or the person who added it.
  // (Hard control stays with SharePoint: only people with Edit access to the home site can change the list at all.)
  function myIds() {
    const a = state.account || {};
    const ids = new Set();
    if (a.localAccountId) ids.add(String(a.localAccountId).toLowerCase());
    if (a.idTokenClaims && a.idTokenClaims.oid) ids.add(String(a.idTokenClaims.oid).toLowerCase());
    const names = new Set();
    if (a.username) names.add(String(a.username).toLowerCase());
    const hint = loginHint && loginHint();
    if (hint) names.add(String(hint).toLowerCase());
    return { ids, names };
  }
  function canDelete(item) {
    if (!item) return false;
    if (item.mine) return true;
    const me = myIds();
    const f = item.fields || {};
    const created = item.createdBy && item.createdBy.user;
    if (created && created.id && me.ids.has(String(created.id).toLowerCase())) return true;
    if (created && created.email && me.names.has(String(created.email).toLowerCase())) return true;
    for (const e of [f.OwnerEmail, f.BackupEmail]) {
      if (!e) continue;
      const key = String(e).toLowerCase();
      if (me.names.has(key)) return true;
      const p = state.people[key];
      if (p && p.id && me.ids.has(String(p.id).toLowerCase())) return true;
      if (p && p.upn && me.names.has(p.upn)) return true;
    }
    return false;
  }
  const NO_DELETE_MSG = "Only this entry's owner, its backup or the person who added it can delete it. You can still edit it, or ask your Microsoft 365 admin.";

  async function deleteItem() {
    const item = state.editing;
    if (!item) return;
    if (!canDelete(item)) { formError(NO_DELETE_MSG); return; }
    if (!(await askConfirm("Delete “" + (item.fields.Title || "this item") + "” from the directory?\n\nEveryone in your company will stop seeing it in RightOwner.", "Delete", true))) return;
    try {
      await graph("DELETE", "/sites/" + state.siteId + "/lists/" + state.listId + "/items/" + item.id);
      state.items = state.items.filter((i) => i !== item);
      $("itemDialog").close();
      render();
      toast("Deleted");
    } catch (err) {
      formError(friendlyWriteError(err));
    }
  }

  function friendlyWriteError(err) {
    if (err && (err.status === 403 || err.status === 401)) {
      return "You can view the directory but not change it. Ask your Microsoft 365 admin to give you Edit access to your company's SharePoint home site.";
    }
    return (err && err.message) || "Something went wrong. Please try again.";
  }

  // ---------- CSV ----------
  function parseCSV(text) {
    const rows = [];
    let row = [], cell = "", inQuotes = false;
    text = text.replace(/^﻿/, "");
    for (let i = 0; i < text.length; i++) {
      const c = text[i];
      if (inQuotes) {
        if (c === '"') {
          if (text[i + 1] === '"') { cell += '"'; i++; } else inQuotes = false;
        } else cell += c;
      } else if (c === '"') inQuotes = true;
      else if (c === ",") { row.push(cell); cell = ""; }
      else if (c === "\n" || c === "\r") {
        if (c === "\r" && text[i + 1] === "\n") i++;
        row.push(cell); cell = "";
        if (row.some((v) => v.trim() !== "")) rows.push(row);
        row = [];
      } else cell += c;
    }
    row.push(cell);
    if (row.some((v) => v.trim() !== "")) rows.push(row);
    return rows;
  }

  function csvCell(v) {
    const s = String(v == null ? "" : v);
    const safe = /^[=+\-@\t\r]/.test(s) ? "'" + s : s; // stop spreadsheet formula injection
    return /[",\n\r]/.test(safe) ? '"' + safe.replace(/"/g, '""') + '"' : safe;
  }

  function downloadCSV(name, rows) {
    const text = "﻿" + rows.map((r) => r.map(csvCell).join(",")).join("\r\n");
    const blob = new Blob([text], { type: "text/csv;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }

  const HEADER = ["Item", "Category", "Description", "Owner email", "Backup email"];

  function mapHeaders(headerRow) {
    const norm = headerRow.map((h) => h.trim().toLowerCase().replace(/[^a-z]/g, ""));
    const find = (...names) => norm.findIndex((h) => names.includes(h));
    return {
      title: find("item", "title", "name", "what"),
      category: find("category", "type", "area"),
      description: find("description", "details", "notes"),
      owner: find("owneremail", "owner", "ownermail"),
      backup: find("backupemail", "backup", "backupmail")
    };
  }

  async function importCSV(file) {
    const text = await file.text();
    const rows = parseCSV(text);
    if (rows.length < 2) return toast("That file has no rows to import.", 4000);
    const map = mapHeaders(rows[0]);
    if (map.title < 0 || map.owner < 0) {
      return toast("The first row needs at least the columns “Item” and “Owner email”. Use More → Download CSV template.", 6000);
    }
    const emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    const existing = new Set(state.items.map((i) => String(i.fields.Title || "").trim().toLowerCase()));
    const toCreate = [];
    let skipped = 0;
    rows.slice(1).forEach((r) => {
      const get = (i) => (i >= 0 && r[i] != null ? r[i].trim() : "");
      const title = get(map.title);
      const owner = get(map.owner);
      const backup = get(map.backup);
      if (!title || !emailRe.test(owner) || (backup && !emailRe.test(backup)) || existing.has(title.toLowerCase())) { skipped++; return; }
      existing.add(title.toLowerCase());
      toCreate.push({ Title: title.slice(0, 255), Category: get(map.category).slice(0, 100), Description: get(map.description).slice(0, 2000), OwnerEmail: owner, BackupEmail: backup });
    });
    if (toCreate.length === 0) return toast("Nothing to import (" + skipped + " rows skipped: missing item/owner, bad email, or already listed).", 6000);
    if (!(await askConfirm("Import " + toCreate.length + " items?" + (skipped ? " (" + skipped + " rows will be skipped.)" : "")))) return;

    toast("Importing " + toCreate.length + " items…", 60000);
    await checkPeople(toCreate.flatMap((f) => [f.OwnerEmail, f.BackupEmail]));
    toCreate.forEach((f) => {
      const o = state.people[f.OwnerEmail.toLowerCase()];
      const b = f.BackupEmail ? state.people[f.BackupEmail.toLowerCase()] : null;
      f.OwnerName = (o && o.displayName) || f.OwnerEmail;
      f.BackupName = f.BackupEmail ? ((b && b.displayName) || f.BackupEmail) : "";
    });

    const base = "/sites/" + state.siteId + "/lists/" + state.listId + "/items";
    let ok = 0, failed = 0, denied = false;
    for (let i = 0; i < toCreate.length; i += 20) {
      const chunk = toCreate.slice(i, i + 20);
      const body = { requests: chunk.map((fields, n) => ({ id: String(n), method: "POST", url: base, headers: { "Content-Type": "application/json" }, body: { fields } })) };
      try {
        const res = await graph("POST", "/$batch", body);
        (res.responses || []).forEach((r) => {
          if (r.status === 201) { ok++; state.items.push({ id: r.body.id, mine: true, fields: chunk[parseInt(r.id, 10)] }); }
          else { failed++; if (r.status === 403) denied = true; }
        });
      } catch (e) {
        failed += chunk.length;
        if (e.status === 403) denied = true;
      }
    }
    state.items.sort((a, b) => String(a.fields.Title || "").localeCompare(String(b.fields.Title || "")));
    render();
    if (denied) toast("You don't have permission to add items. Ask your admin for Edit access to the SharePoint home site.", 7000);
    else toast("Imported " + ok + " items" + (failed ? ", " + failed + " failed" : "") + (skipped ? ", " + skipped + " skipped" : "") + ".", 6000);
  }

  function exportCSV() {
    const rows = [HEADER.concat(["Status"])];
    state.items.forEach((i) => {
      const f = i.fields;
      rows.push([f.Title, f.Category, f.Description, f.OwnerEmail, f.BackupEmail, needsAttention(f) || "OK"]);
    });
    downloadCSV("rightowner-directory.csv", rows);
  }

  // ---------- startup ----------
  async function loadDirectory() {
    show("loading");
    try {
      const exists = await findList();
      if (!exists) return showSetup();
      await loadItems();
      await checkPeople(state.items.flatMap((i) => [i.fields.OwnerEmail, i.fields.BackupEmail]));
      show("directory");
      render();
      $("search").focus();
    } catch (err) {
      handleStartupError(err);
    }
  }

  function showSetup() {
    showMessage(
      "Set up your company directory",
      "RightOwner keeps your directory in a SharePoint list called “" + CFG.listName + "” on your company's SharePoint home site. " +
      "It stays inside your Microsoft 365. A Microsoft 365 admin or SharePoint site owner needs to create it once.",
      [
        { label: "Create the directory", primary: true, onClick: async () => {
          show("loading");
          try { await createList(); state.items = []; show("directory"); render(); toast("Directory created. Add your first item."); }
          catch (err) {
            if (err.status === 403) {
              showMessage("Ask an admin to finish setup",
                "You don't have permission to create lists on your company's SharePoint home site. Ask a Microsoft 365 admin to open RightOwner once and click “Create the directory”.",
                [{ label: "Try again", onClick: loadDirectory }]);
            } else handleStartupError(err);
          }
        } },
        { label: "Learn more", onClick: () => openExternal("https://rubuild.com/support") }
      ]
    );
  }

  function handleStartupError(err) {
    if (err && err.auth) {
      if (err.consent) {
        return showMessage("Admin approval needed",
          "RightOwner needs a Microsoft 365 admin to approve its permissions once for your organization. If you are an admin, click Approve. Otherwise, send this page to your admin.",
          [
            { label: "Approve (admins)", primary: true, onClick: () => openExternal(adminConsentUrl()) },
            { label: "I've approved — try again", onClick: start }
          ]);
      }
      return showMessage("Sign in to RightOwner", "We need you to sign in with your work account to load your company's directory.",
        [{ label: "Sign in", primary: true, onClick: async () => {
          try { await getToken(true); loadDirectory(); } catch (e) { handleStartupError(markAuthError(e)); }
        } }]);
    }
    if (err && err.status === 403) {
      return showMessage("No access to SharePoint",
        "Your account can't read your company's SharePoint home site. Ask your Microsoft 365 admin to give you at least Read access.",
        [{ label: "Try again", onClick: loadDirectory }]);
    }
    showMessage("Something went wrong", (err && err.message) || "Please try again.",
      [{ label: "Try again", primary: true, onClick: loadDirectory }, { label: "Get help", onClick: () => openExternal("https://rubuild.com/support") }]);
  }

  function wireUI() {
    $("ver").textContent = "v" + (CFG.version || "1.0.0");
    $("search").addEventListener("input", (e) => { state.query = e.target.value; render(); });
    document.querySelectorAll(".chip").forEach((chip) => chip.addEventListener("click", () => {
      document.querySelectorAll(".chip").forEach((c) => { c.classList.remove("active"); c.setAttribute("aria-selected", "false"); });
      chip.classList.add("active");
      chip.setAttribute("aria-selected", "true");
      state.filter = chip.dataset.filter;
      render();
    }));
    $("addBtn").addEventListener("click", () => openDialog(null));
    $("cancelBtn").addEventListener("click", () => $("itemDialog").close());
    $("itemForm").addEventListener("submit", saveItem);
    $("deleteBtn").addEventListener("click", deleteItem);
    document.querySelectorAll(".person-field").forEach(setupPersonField);

    const menu = $("moreMenu");
    $("moreBtn").addEventListener("click", (e) => {
      e.stopPropagation();
      menu.hidden = !menu.hidden;
      $("moreBtn").setAttribute("aria-expanded", String(!menu.hidden));
    });
    document.addEventListener("click", () => { menu.hidden = true; $("moreBtn").setAttribute("aria-expanded", "false"); });
    $("importBtn").addEventListener("click", () => $("csvFile").click());
    $("csvFile").addEventListener("change", async (e) => {
      const file = e.target.files && e.target.files[0];
      e.target.value = "";
      if (file) { try { await importCSV(file); } catch (err) { toast(err.message || "Import failed", 5000); } }
    });
    $("templateBtn").addEventListener("click", () => downloadCSV("rightowner-template.csv", [
      HEADER,
      ["Payroll", "Finance", "Monthly payroll runs, salary changes, payslip questions", "owner@yourcompany.com", "backup@yourcompany.com"],
      ["VPN access", "IT", "New VPN accounts and connection problems", "owner@yourcompany.com", ""]
    ]));
    $("exportBtn").addEventListener("click", exportCSV);
    $("refreshBtn").addEventListener("click", () => { state.people = {}; loadDirectory(); });

    $("cards").addEventListener("click", (e) => {
      const b = e.target.closest("button[data-act]");
      if (!b) return;
      if (b.dataset.act === "edit") openDialog(state.items.find((i) => i.id === b.dataset.id));
      if (b.dataset.act === "chat") openExternal("https://teams.microsoft.com/l/chat/0/0?users=" + encodeURIComponent(b.dataset.email));
      if (b.dataset.act === "mail") openExternal("mailto:" + b.dataset.email + "?subject=" + encodeURIComponent("Question about " + (b.dataset.title || "")));
    });
  }

  async function start() {
    show("loading");
    if (!CFG.clientId || /PASTE/.test(CFG.clientId)) {
      return showMessage("Almost ready", "This copy of RightOwner hasn't been connected to Microsoft 365 yet (missing client ID in config.js).", []);
    }
    try {
      if (!state.pca) await initAuth();
    } catch (e) {
      return showMessage("Couldn't start", "RightOwner couldn't start Microsoft sign-in. Please reload the app.", [{ label: "Reload", primary: true, onClick: () => location.reload() }]);
    }
    loadDirectory();
  }

  // Exposed only for automated tests.
  window.__rightowner_test = { parseCSV, csvCell, mapHeaders, esc, initials };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", () => { wireUI(); start(); });
  else { wireUI(); start(); }
})();
