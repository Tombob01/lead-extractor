import { useState, useRef, useCallback, useEffect, useMemo } from "react";
import * as XLSX from "xlsx";

// ── Capacitor local notifications ─────────────────────────────────────────
const getLocalNotifications = async () => {
  try {
    const mod = await import('@capacitor/local-notifications');
    return mod.LocalNotifications;
  } catch { return null; }
};

const requestNotificationPermission = async () => {
  try {
    const LocalNotifications = await getLocalNotifications();
    if (!LocalNotifications) return false;
    const { display } = await LocalNotifications.requestPermissions();
    return display === 'granted';
  } catch { return false; }
};

const sendNotification = async (title, body) => {
  try {
    const LocalNotifications = await getLocalNotifications();
    if (!LocalNotifications) return;
    const { display } = await LocalNotifications.checkPermissions();
    if (display !== 'granted') await requestNotificationPermission();
    await LocalNotifications.schedule({
      notifications: [{
        title,
        body,
        id: Date.now(),
        schedule: { at: new Date(Date.now() + 500) },
        sound: null,
        actionTypeId: '',
        extra: null,
      }]
    });
  } catch (e) {
    console.warn('Notification failed:', e.message);
  }
};

// ── ScrapePlugin — native foreground service bridge ─────────────────────────
// Keeps scraping alive when app is minimized or screen is off.
// Shows persistent notification with live progress and tab-specific title.
// Falls back silently in browser preview (no Capacitor runtime).

// Register the plugin ONCE at module load time — not lazily.
// registerPlugin() must be called before any method is invoked so Capacitor
// can wire the JS proxy to the native Android implementation on startup.
const _resolveScrapePlugin = () => {
  try {
    if (!window.Capacitor) return null;
    // registerPlugin is the correct call for custom inline plugins registered
    // via registerPlugin(ScrapePlugin.class) in MainActivity.
    // Call it once here at module scope so the bridge is ready immediately.
    if (window.Capacitor.registerPlugin) {
      const p = window.Capacitor.registerPlugin('ScrapePlugin');
      console.log("[ScrapeNative] Plugin registered at module load");
      return p;
    }
  } catch (e) {
    console.warn("[ScrapeNative] Plugin registration failed:", e.message);
  }
  return null;
};
const _scrapePluginInstance = _resolveScrapePlugin();

// Request POST_NOTIFICATIONS permission at runtime (required on Android 13+ / API 33+).
// Manifest declaration alone is not enough — the user must grant it at runtime.
const _requestNotifPermission = async () => {
  try {
    if (!window.Capacitor?.isNativePlatform?.()) return;
    const { LocalNotifications } = await import('@capacitor/local-notifications').catch(() => ({}));
    if (LocalNotifications?.requestPermissions) {
      const { display } = await LocalNotifications.requestPermissions();
      console.log("[ScrapeNative] Notification permission:", display);
    }
  } catch (e) {
    console.warn("[ScrapeNative] Permission request failed:", e.message);
  }
};
// Fire permission request as soon as the module loads on a native platform.
_requestNotifPermission();

const ScrapeNative = {
  _plugin: _scrapePluginInstance,
  _get() {
    // Return the already-registered instance. If Capacitor loaded after module
    // init (rare), attempt registration once more as a safety net.
    if (this._plugin) return this._plugin;
    try {
      if (window.Capacitor?.registerPlugin) {
        this._plugin = window.Capacitor.registerPlugin('ScrapePlugin');
        return this._plugin;
      }
    } catch (e) {
      console.warn("[ScrapeNative] late _get failed:", e.message);
    }
    return null;
  },
  // type: "bulk" | "smart" | "extract" | "maps"
  async start(total, type = "bulk") {
    try {
      const p = this._get();
      if (p) {
        await p.start({ total, type });
        console.log("[ScrapeNative] Service started — type:", type, "total:", total);
      }
    } catch (e) { console.warn("[ScrapeNative] start failed:", e.message); }
  },
  async progress(done, total, emails, type = "bulk", success = 0, failed = 0) {
    try {
      const p = this._get();
      if (p) await p.progress({ done, total, emails, type, success, failed });
    } catch (e) { console.warn("[ScrapeNative] progress failed:", e.message); }
  },
  // Call when scraping finishes normally — shows "✓ Done — X emails" then auto-dismisses
  async complete(total, emails) {
    try {
      const p = this._get();
      if (p) {
        await p.complete({ total, emails });
        console.log("[ScrapeNative] complete — emails:", emails);
      }
    } catch (e) { console.warn("[ScrapeNative] complete failed:", e.message); }
  },
  // Call when user taps Stop — immediately dismisses notification
  async stop() {
    try {
      const p = this._get();
      if (p) { await p.stop(); console.log("[ScrapeNative] Service stopped"); }
    } catch (e) { console.warn("[ScrapeNative] stop failed:", e.message); }
  },
};

// ── utils ──────────────────────────────────────────────────────────────────
const readFileAsText = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsText(f); });
const readFileAsBase64 = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result.split(",")[1]); r.onerror = rej; r.readAsDataURL(f); });
const uid = () => Math.random().toString(36).slice(2, 9);
const normalizeEmail = (e) => (e || "").trim().toLowerCase();

// ── API Credit Monitor — independent, read-only, non-blocking ───────────
// Credit fetch helper — tries direct first, then two CORS proxy fallbacks
const fetchJsonWithFallback = async (endpoint, timeoutMs = 8000) => {
  // Attempt 1: direct fetch (works if API supports CORS)
  try {
    const res = await fetch(endpoint, { signal: AbortSignal.timeout(timeoutMs) });
    if (res.ok) return await res.json();
  } catch {}
  // Attempt 2: corsproxy.io — transparent proxy, preserves query params + auth
  try {
    const proxied = `https://corsproxy.io/?url=${encodeURIComponent(endpoint)}`;
    const res = await fetch(proxied, { signal: AbortSignal.timeout(timeoutMs) });
    if (res.ok) return await res.json();
  } catch {}
  // Attempt 3: allorigins — wraps response in { contents: "..." }
  try {
    const proxied = `https://api.allorigins.win/get?url=${encodeURIComponent(endpoint)}`;
    const res = await fetch(proxied, { signal: AbortSignal.timeout(timeoutMs) });
    if (!res.ok) return null;
    const wrapper = await res.json();
    if (wrapper?.contents) return JSON.parse(wrapper.contents);
  } catch {}
  return null;
};

// Credit fetch: ScraperAPI
const fetchScraperApiCredits = async (apiKey) => {
  try {
    const endpoint = `https://api.scraperapi.com/account?api_key=${encodeURIComponent(apiKey)}`;
    const data = await fetchJsonWithFallback(endpoint);
    if (!data) return "unavailable";
    return typeof data.requestLimit === "number" && typeof data.requestCount === "number"
      ? data.requestLimit - data.requestCount
      : "unavailable";
  } catch { return "unavailable"; }
};

// Credit fetch: ZenRows
// ZenRows does NOT expose a public credits/account API endpoint.
// Usage is only visible inside their dashboard at app.zenrows.com.
// We track local usage count as the only available metric.
const fetchZenRowsCredits = async (_apiKey) => "no-api";

// Local usage counters stored in localStorage (non-intrusive)
const CREDIT_KEY = "apiCreditUsage";
const loadCreditUsage = () => {
  try {
    const raw = localStorage.getItem(CREDIT_KEY);
    const defaults = { snovUsed: 0, apolloUsed: 0, mapsUsed: 0 };
    return raw ? { ...defaults, ...JSON.parse(raw) } : defaults;
  } catch { return { snovUsed: 0, apolloUsed: 0, mapsUsed: 0 }; }
};
const saveCreditUsage = (usage) => {
  try { localStorage.setItem(CREDIT_KEY, JSON.stringify(usage)); } catch {}
};
const incrementUsage = (key) => {
  const usage = loadCreditUsage();
  usage[key] = (usage[key] || 0) + 1;
  saveCreditUsage(usage);
  return usage;
};
// Exported helpers — called AFTER successful API calls only
const incrementSnovUsage   = () => incrementUsage("snovUsed");
const incrementApolloUsage = () => incrementUsage("apolloUsed");
const incrementMapsUsage   = () => incrementUsage("mapsUsed");

// ── Strip UTM and tracking params — reduces page weight and avoids heavy redirects
const stripTrackingParams = (url) => {
  try {
    const u = new URL(url.startsWith("http") ? url : `https://${url}`);
    const trackingParams = [
      "utm_source","utm_medium","utm_campaign","utm_term","utm_content","utm_id",
      "fbclid","gclid","msclkid","mc_cid","mc_eid","ref","source","campaign",
      "gbraid","wbraid","igshid","s_kwcid","ef_id","_ga",
    ];
    trackingParams.forEach(p => u.searchParams.delete(p));
    return u.toString();
  } catch { return url; }
};

// ── email validation (regex + basic MX check via dns.google) ──────────────
const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;
const validateEmailFormat = (email) => EMAIL_PATTERN.test(email);

const checkMxRecord = async (domain) => {
  try {
    const res = await fetch(`https://dns.google/resolve?name=${encodeURIComponent(domain)}&type=MX`, { signal: AbortSignal.timeout(4000) });
    const data = await res.json();
    return !!(data.Answer && data.Answer.length > 0);
  } catch { return true; } // assume valid on error
};

const verifyEmail = async (email) => {
  if (!validateEmailFormat(email)) return "invalid";
  const domain = email.split("@")[1];
  const hasMx = await checkMxRecord(domain);
  return hasMx ? "valid" : "risky";
};

// ── concurrency limiter ────────────────────────────────────────────────────
// ── JobPoller — delegates bulk scraping to Cloudflare Worker job engine ──
// Submits a URL list to POST /jobs, then polls /jobs/:id/status every 4s.
// The actual scraping runs server-side — survives app minimize, screen lock,
// or the app being fully closed. Frontend is controller/UI only.

const JobPoller = {
  _timer:         null,
  _resultsOffset: 0,
  _jobId:         null,

  async start({ urls, provider = "browser", onProgress, onResult, onComplete, onError }) {
    try {
      const base = workerUrlRef.current?.trim().replace(/\/$/, "");
      if (!base) throw new Error("Worker URL not set — add it in Settings");

      const res = await fetch(`${base}/jobs`, {
        method:  "POST",
        headers: { "Content-Type": "application/json" },
        body:    JSON.stringify({ urls, provider }),
      });
      if (!res.ok) throw new Error(`Job submit failed: ${res.status}`);

      const { jobId, total } = await res.json();
      this._jobId         = jobId;
      this._resultsOffset = 0;
      console.log(`[JobPoller] Job started — id:${jobId} | ${total} URLs`);

      this._poll(jobId, base, onProgress, onResult, onComplete, onError);
      return jobId;
    } catch (e) {
      console.error("[JobPoller] start error:", e.message);
      onError?.(e.message);
    }
  },

  _poll(jobId, base, onProgress, onResult, onComplete, onError) {
    this._timer = setInterval(async () => {
      try {
        // ── status ──────────────────────────────────────────────────────────
        const statusRes = await fetch(`${base}/jobs/${jobId}/status`);
        if (!statusRes.ok) {
          const errText = await statusRes.text().catch(() => "");
          throw new Error(`Status ${statusRes.status}: ${errText.slice(0, 80)}`);
        }
        let status;
        try { status = await statusRes.json(); }
        catch (e) { throw new Error(`Status response was not JSON: ${e.message}`); }

        console.log(`[JobPoller] tick — ${status.done}/${status.total} emails:${status.emails} state:${status.state}`);
        onProgress?.(status);

        // ── results (fetch only new results since last poll) ─────────────
        const resultsRes = await fetch(`${base}/jobs/${jobId}/results?offset=${this._resultsOffset}`);
        if (!resultsRes.ok) {
          console.warn(`[JobPoller] Results fetch ${resultsRes.status} — skipping tick`);
        } else {
          let parsed;
          try { parsed = await resultsRes.json(); }
          catch (e) {
            console.warn(`[JobPoller] Results response was not JSON: ${e.message}`);
            parsed = null;
          }
          if (parsed) {
            const { results, total: resultTotal } = parsed;
            if (results?.length > 0) {
              this._resultsOffset = resultTotal;
              onResult?.(results);
            }
          }
        }

        if (status.state === "complete" || status.state === "cancelled") {
          this.stop();
          onComplete?.(status);
        }
      } catch (e) {
        console.warn("[JobPoller] poll error:", e.message);
        onError?.(e.message);
      }
    }, 4000);
  },

  stop() {
    if (this._timer) {
      clearInterval(this._timer);
      this._timer = null;
      console.log("[JobPoller] Polling stopped");
    }
  },

  async cancel(jobId) {
    this.stop();
    try {
      const base = workerUrlRef.current?.trim().replace(/\/$/, "");
      if (base && jobId) {
        await fetch(`${base}/jobs/${jobId}/cancel`, { method: "POST" });
        console.log(`[JobPoller] Job ${jobId} cancelled`);
      }
    } catch (e) {
      console.warn("[JobPoller] cancel error:", e.message);
    }
    this._jobId = null;
  },
};

const withConcurrency = async (items, concurrency, fn, onItemDone) => {
  const results = new Array(items.length);
  let idx = 0;
  const worker = async () => {
    while (idx < items.length) {
      const i = idx++;
      results[i] = await fn(items[i], i);
      if (onItemDone) onItemDone(i, results[i]);
    }
  };
  await Promise.all(Array.from({ length: Math.min(concurrency, items.length) }, worker));
  return results;
};

// ── ETA calculator ─────────────────────────────────────────────────────────
const calcEta = (startTime, done, total) => {
  if (!done || !startTime) return null;
  const elapsed = (Date.now() - startTime) / 1000;
  const rate = done / elapsed;
  const remaining = (total - done) / rate;
  if (remaining < 60) return `~${Math.ceil(remaining)}s remaining`;
  return `~${Math.ceil(remaining / 60)}m remaining`;
};

const C = {
  bg: "#07090d", surface: "#0d1117", surface2: "#111820", border: "#1a2336",
  accent: "#00e87a", blue: "#29b6f6", purple: "#a78bfa", warn: "#ffb347",
  danger: "#ff5f72", muted: "#3a4a5e", text: "#c9d6e8", dim: "#4a5a72",
};

// ── Feature 4: Per-URL error classification ───────────────────────────────
// Maps an error message (or empty-email result) to an actionable category.
// Returns one of: "blocked" | "empty" | "timeout" | "dns" | "redirect" | "error"
const classifyBulkError = (errMsg, emails) => {
  if (emails && emails.length > 0) return null; // no error
  const msg = (errMsg || "").toLowerCase();
  if (!errMsg && emails && emails.length === 0) return "empty";
  if (msg.includes("cloudflare") || msg.includes("imperva") || msg.includes("403") || msg.includes("captcha") || msg.includes("blocked") || msg.includes("forbidden")) return "blocked";
  if (msg.includes("timeout") || msg.includes("aborted") || msg.includes("timedout") || msg.includes("etimedout")) return "timeout";
  if (msg.includes("getaddrinfo") || msg.includes("dns") || msg.includes("enotfound") || msg.includes("name not resolved") || msg.includes("502") || msg.includes("no such host")) return "dns";
  if (msg.includes("redirect") || msg.includes("too many redirect") || msg.includes("econnreset")) return "redirect";
  return "error";
};

const BULK_ERROR_META = {
  blocked:  { icon: "🔴", label: "blocked",       color: "#ff5f72", tip: "Cloudflare/Imperva — use Smart Scraper" },
  empty:    { icon: "🟡", label: "empty",          color: "#ffb347", tip: "Loaded, 0 emails — try Smart Scraper" },
  timeout:  { icon: "🟠", label: "timeout",        color: "#fb923c", tip: "Too slow — retry or increase timeout" },
  dns:      { icon: "⚫", label: "DNS / dead",     color: "#6b7280", tip: "Domain doesn't resolve — remove URL" },
  redirect: { icon: "🔵", label: "redirect loop",  color: "#29b6f6", tip: "Too many redirects" },
  error:    { icon: "🔴", label: "error",          color: "#ff5f72", tip: "Fetch failed" },
};

// ── Feature 6: Domain-match email scorer ──────────────────────────────────
// Returns "match" (business email) | "personal" (gmail/yahoo etc.) | "other" (mismatch)
const PERSONAL_DOMAINS = new Set(["gmail.com","yahoo.com","hotmail.com","outlook.com","icloud.com","aol.com","protonmail.com","me.com","live.com","msn.com","ymail.com","mail.com","zoho.com","inbox.com"]);

const scoreEmailDomain = (email, siteUrl) => {
  try {
    const emailDomain = email.split("@")[1]?.toLowerCase();
    if (!emailDomain) return "other";
    if (PERSONAL_DOMAINS.has(emailDomain)) return "personal";
    const siteDomain = new URL(siteUrl).hostname.replace(/^www\./, "").toLowerCase();
    // Match if email domain equals or is a subdomain of the site domain
    if (emailDomain === siteDomain || siteDomain.endsWith("." + emailDomain) || emailDomain.endsWith("." + siteDomain)) return "match";
    return "other";
  } catch { return "other"; }
};

const FIELDS = ["name", "email", "company", "role"];

// ── HubSpot push ───────────────────────────────────────────────────────────
const pushToHubSpot = async (apiKey, leads) => {
  const results = [];
  for (const lead of leads) {
    const props = {};
    if (lead.name) { const p = lead.name.trim().split(" "); props.firstname = p[0]; if (p.length > 1) props.lastname = p.slice(1).join(" "); }
    if (lead.email) props.email = lead.email;
    if (lead.company) props.company = lead.company;
    if (lead.role) props.jobtitle = lead.role;
    const res = await fetch("https://api.hubapi.com/crm/v3/objects/contacts", {
      method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${apiKey}` },
      body: JSON.stringify({ properties: props }),
    });
    const data = await res.json();
    results.push({ lead, ok: res.ok, id: data.id, error: data.message });
  }
  return results;
};

// ── dedup ──────────────────────────────────────────────────────────────────
const markDuplicates = (list) => {
  const seen = {};
  return list.map((l) => {
    const key = normalizeEmail(l.email) || l.name?.trim().toLowerCase();
    if (!key) return { ...l, _dup: false };
    if (seen[key]) return { ...l, _dup: true };
    seen[key] = true;
    return { ...l, _dup: false };
  });
};

// ── Gemini API helper ──────────────────────────────────────────────────────
// Only use verified, non-deprecated models.
// "gemini-2.0-flash-lite" removed — it is not a valid public model name.
const GEMINI_MODELS = ["gemini-2.0-flash", "gemini-1.5-flash"];

const callGemini = async (apiKey, prompt) => {
  // Sanitize key: trim whitespace and remove any accidental line breaks
  const cleanKey = apiKey?.replace(/[\r\n]/g, "").trim();
  if (!cleanKey) {
    console.error("Gemini API key missing or blank — aborting request");
    throw new Error("Gemini API key missing");
  }

  let lastError = null;
  for (const model of GEMINI_MODELS) {
    try {
      // API key goes in the URL query param — this is the canonical auth method
      // for the Gemini REST API and matches the working test URL format.
      // Do NOT use "Authorization: Bearer" here — that is OpenAI format.
      // The x-goog-api-key header is also valid but can be stripped by some
      // CORS proxies; the ?key= param is more reliable.
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${encodeURIComponent(cleanKey)}`;
      console.log("Gemini endpoint:", url.replace(cleanKey, "***"));
      console.log("Gemini model:", model);

      const res = await fetch(url, {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          // Note: no Authorization header — Gemini uses ?key= or x-goog-api-key, NOT Bearer tokens
        },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.1, maxOutputTokens: 2048 },
        }),
      });

      console.log("Gemini status:", res.status);

      // Always read as text first — prevents silent crashes on HTML error pages,
      // quota-exceeded responses, or malformed JSON from the API.
      const rawText = await res.text();

      if (!res.ok) {
        console.log("Gemini raw response:", rawText);
        let errMsg = `Gemini ${res.status}`;
        try {
          const errJson = JSON.parse(rawText);
          errMsg = errJson?.error?.message || errMsg;
        } catch {
          // Non-JSON response (HTML error page, etc.) — use status text
          errMsg = rawText.slice(0, 200) || errMsg;
        }
        lastError = new Error(errMsg);
        continue; // try next model
      }

      let data;
      try {
        data = JSON.parse(rawText);
      } catch (parseErr) {
        console.log("Gemini raw response (parse failed):", rawText);
        lastError = new Error(`Gemini response was not valid JSON: ${parseErr.message}`);
        continue;
      }

      if (data.error) {
        console.log("Gemini API error in body:", data.error);
        lastError = new Error(data.error.message || "Unknown Gemini error");
        continue; // try next model
      }

      return data?.candidates?.[0]?.content?.parts?.[0]?.text ?? "";
    } catch (e) {
      console.log("Gemini fetch error:", e.message);
      lastError = e;
    }
  }
  throw lastError;
};

// ── CORS proxy list (expanded — 20 proxies, tiered by reliability) ────────
const CORS_PROXIES = [
  // TIER 1 — most reliable
  { id: "allorigins-json",    make: (u) => `https://api.allorigins.win/get?url=${encodeURIComponent(u)}`,                                    mode: "json_contents" },
  { id: "corsproxy-io",       make: (u) => `https://corsproxy.io/?url=${encodeURIComponent(u)}`,                                             mode: "text" },
  { id: "cors-lol",           make: (u) => `https://api.cors.lol/?url=${encodeURIComponent(u)}`,                                             mode: "text" },
  { id: "corsfix",            make: (u) => `https://proxy.corsfix.com/?${encodeURIComponent(u)}`,                                            mode: "text" },
  { id: "corsproxy-dev",      make: (u) => `https://corsproxy.dev/?url=${encodeURIComponent(u)}`,                                            mode: "text" },
  { id: "cors-proxy-app",     make: (u) => `https://cors-proxy.org/api/?url=${encodeURIComponent(u)}`,                                       mode: "text" },
  // TIER 2 — good fallbacks
  { id: "allorigins-raw",     make: (u) => `https://api.allorigins.win/raw?url=${encodeURIComponent(u)}`,                                    mode: "text" },
  { id: "codetabs",           make: (u) => `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(u)}`,                               mode: "text" },
  { id: "thingproxy",         make: (u) => `https://thingproxy.freeboard.io/fetch/${u}`,                                                     mode: "text" },
  { id: "htmldriven",         make: (u) => `https://cors-proxy.htmldriven.com/?url=${encodeURIComponent(u)}`,                                mode: "text" },
  { id: "whateverorigin",     make: (u) => `https://api.whateverorigin.com/get?url=${encodeURIComponent(u)}`,                                mode: "json_contents" },
  { id: "nocors-io",          make: (u) => `https://nocors.io/proxy?url=${encodeURIComponent(u)}`,                                          mode: "text" },
  { id: "proxify",            make: (u) => `https://proxy.scrapeops.io/v1/?api_key=free&url=${encodeURIComponent(u)}`,                       mode: "text" },
  { id: "webscraping-ai",     make: (u) => `https://app.webscraping.ai/api/html?api_key=free&url=${encodeURIComponent(u)}`,                  mode: "text" },
  // TIER 3 — alternate formats / last resort
  { id: "cors-anywhere",      make: (u) => `https://www.cors-anywhere.com/${u}`,                                                             mode: "text" },
  { id: "crossorigin-me",     make: (u) => `https://crossorigin.me/${u}`,                                                                   mode: "text" },
  { id: "corsproxy-io-raw",   make: (u) => `https://corsproxy.io/?${encodeURIComponent(u)}`,                                                mode: "text" },
  { id: "allorigins-utf8",    make: (u) => `https://api.allorigins.win/get?url=${encodeURIComponent(u)}&charset=utf-8`,                      mode: "json_contents" },
  { id: "corsproxy-io-plain", make: (u) => `https://corsproxy.io/?${u}`,                                                                    mode: "text" },
  { id: "allorigins-json2",   make: (u) => `https://api.allorigins.win/get?url=${encodeURIComponent(u)}&disableCache=true`,                  mode: "json_contents" },
];

// Browser-like headers to pass along with proxy requests (reduces bot detection)
const BROWSER_HEADERS = {
  "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
  "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
  "Accept-Language": "en-US,en;q=0.5",
};

// ── Smart proxy rotator ────────────────────────────────────────────────────
const proxyState = (() => {
  const FAIL_COOLDOWN_MS = 60_000;
  const STORAGE_KEY = "proxyState_scores";
  let scores = {};
  try {
    const saved = sessionStorage.getItem(STORAGE_KEY);
    if (saved) { scores = JSON.parse(saved); }
  } catch { scores = {}; }

  const persist = () => { try { sessionStorage.setItem(STORAGE_KEY, JSON.stringify(scores)); } catch {} };
  const get = (id) => scores[id] ?? (scores[id] = { fails: 0, lastFail: 0, lastSuccess: 0 });

  const onSuccess = (id) => {
    const s = get(id); s.fails = 0; s.lastSuccess = Date.now(); persist();
    console.log(`[ProxyRotator] ✓ ${id} succeeded`);
  };
  const onFail = (id) => {
    const s = get(id); s.fails++; s.lastFail = Date.now(); persist();
    console.log(`[ProxyRotator] ✕ ${id} failed (total: ${s.fails})`);
  };
  const sorted = () => {
    const now = Date.now();
    const available = CORS_PROXIES.filter(p => {
      const s = get(p.id);
      return s.fails === 0 || (now - s.lastFail) > FAIL_COOLDOWN_MS;
    });
    const pool = available.length > 0 ? available : [...CORS_PROXIES];
    return pool.sort((a, b) => {
      const sa = get(a.id), sb = get(b.id);
      if (sb.lastSuccess !== sa.lastSuccess) return sb.lastSuccess - sa.lastSuccess;
      return sa.fails - sb.fails;
    });
  };
  return { onSuccess, onFail, sorted };
})();

// ── Shared HTML fetch via smart proxy rotation ─────────────────────────────
const fetchHtmlSmart = async (url) => {
  const proxies = proxyState.sorted();
  console.log(`[ProxyRotator] Fetching ${url} — trying ${proxies.length} proxies: ${proxies.slice(0,5).map(p=>p.id).join(", ")}…`);
  for (const proxy of proxies) {
    try {
      const res = await fetch(proxy.make(url), {
        signal: AbortSignal.timeout(18000),
        headers: BROWSER_HEADERS,
      });
      if (!res.ok) { proxyState.onFail(proxy.id); continue; }
      let html;
      if (proxy.mode === "json_contents") {
        const data = await res.json();
        html = data?.contents ?? data?.data ?? null;
      } else {
        html = await res.text();
      }
      // Reject proxy error pages masquerading as success
      if (html && html.length > 200 && !html.includes("Error connecting to the target") && !html.includes("The requested URL could not be retrieved")) {
        proxyState.onSuccess(proxy.id);
        return html;
      }
      proxyState.onFail(proxy.id);
    } catch (_) { proxyState.onFail(proxy.id); }
  }
  return null;
};

// ── Free fetch stack for Bulk Scraper (never touches ScraperAPI/ZenRows) ──
//
// Priority chain:
//   1. Cloudflare Browser Rendering (free, real Chromium via worker)
//   2. Google Cache (rendered snapshot, free, no key)
//   3. Wayback Machine (archive snapshot, free, no key)
//   4. CORS proxy rotation (existing fallback)
//
// ScraperAPI and ZenRows are intentionally excluded — zero paid credits used.

const fetchBulkBrowserRendering = async (url) => {
  const wurl = workerUrlRef.current?.trim();
  if (!wurl) return null;

  // ── Step 1: Direct fetch (fast, ~1-3s, no puppeteer) ─────────────────────
  // Works for small business sites with emails in static HTML.
  // Falls through to full browser only if this returns no emails or fails.
  // Pass multi=true so the worker also fetches /contact and /about server-side.
  try {
    const directEndpoint = `${wurl.replace(/\/$/, "")}/fetch?url=${encodeURIComponent(url)}&provider=direct&extract=true&multi=true`;
    console.log("[BulkFree] Trying direct fetch first:", url);
    const res = await fetch(directEndpoint, { signal: AbortSignal.timeout(12000) });
    if (res.ok) {
      const data = await res.json();
      // Feature 3: if worker signals host is in allowlist block, mark it so
      // fetchPageEmails can skip the entire seed page crawl for this domain.
      if (data?.blocked === true) {
        console.log("[BulkFree] Direct fetch blocked (allowlist) for:", url);
        return { __browserEmails: [], __screenshot: null, __blocked: true };
      }
      if (Array.isArray(data?.emails) && data.emails.length > 0) {
        console.log("[BulkFree] Direct fetch got", data.emails.length, "emails for:", url);
        return { __browserEmails: data.emails, __screenshot: null };
      }
      console.log("[BulkFree] Direct fetch returned 0 emails — escalating to browser:", url);
    }
  } catch (e) {
    console.warn("[BulkFree] Direct fetch failed:", e.message, "— escalating to browser");
  }

  // ── Step 2: Full Cloudflare Browser (JS rendering, DOM extraction) ────────
  // Only reached when direct fetch fails or finds no emails.
  // Handles JS-rendered sites, obfuscated emails, lazy-loaded content.
  // Timeout is 90s to allow for the worker's own 60s 429-retry pause.
  try {
    const browserEndpoint = `${wurl.replace(/\/$/, "")}/fetch?url=${encodeURIComponent(url)}&provider=browser&extract=true&screenshot=true`;
    console.log("[BulkFree] Escalating to browser extract mode:", url);
    const res = await fetch(browserEndpoint, { signal: AbortSignal.timeout(90000) });
    if (!res.ok) {
      // 429 came back to client — worker already retried once internally.
      // Wait an extra 30s client-side so the next URL in the batch doesn't
      // immediately hit an still-exhausted browser pool.
      if (res.status === 429) {
        console.warn("[BulkFree] Browser pool still exhausted after worker retry — pausing 30s before next URL");
        await new Promise(r => setTimeout(r, 30000));
      }
      return null;
    }
    const data = await res.json();
    if (Array.isArray(data?.emails) && data.emails.length > 0) {
      console.log("[BulkFree] Browser returned", data.emails.length, "emails for:", url);
      return { __browserEmails: data.emails, __screenshot: data.screenshot ?? null };
    }
    // Browser found 0 emails — fall through to Google Cache / Wayback
    // which may have older cached versions of the page with emails visible
    console.log("[BulkFree] Browser found 0 emails — falling through to cache layers:", url);
  } catch (e) { console.warn("[BulkFree] Browser rendering failed:", e.message); }

  return null;
};

const fetchGoogleCache = async (url) => {
  const cacheUrl = `https://webcache.googleusercontent.com/search?q=cache:${encodeURIComponent(url)}&hl=en`;
  const proxies = [
    { make: (u) => `https://api.allorigins.win/get?url=${encodeURIComponent(u)}`, mode: "json" },
    { make: (u) => `https://corsproxy.io/?url=${encodeURIComponent(u)}`,          mode: "text" },
    { make: (u) => `https://api.allorigins.win/raw?url=${encodeURIComponent(u)}`, mode: "text" },
  ];
  for (const { make, mode } of proxies) {
    try {
      const res = await fetch(make(cacheUrl), { signal: AbortSignal.timeout(15000) });
      if (!res.ok) continue;
      const html = mode === "json" ? (await res.json())?.contents : await res.text();
      if (html && html.length > 500 && !html.includes("did not match any documents")) {
        console.log("[BulkFree] Google Cache hit:", url);
        return html;
      }
    } catch { /* next proxy */ }
  }
  return null;
};

const fetchWaybackMachine = async (url) => {
  try {
    const cdxRes = await fetch(
      `https://archive.org/wayback/available?url=${encodeURIComponent(url)}`,
      { signal: AbortSignal.timeout(10000) }
    );
    if (!cdxRes.ok) return null;
    const snapshotUrl = (await cdxRes.json())?.archived_snapshots?.closest?.url;
    if (!snapshotUrl) return null;
    console.log("[BulkFree] Wayback snapshot found:", snapshotUrl);
    const proxied = `https://api.allorigins.win/get?url=${encodeURIComponent(snapshotUrl)}`;
    const snapRes = await fetch(proxied, { signal: AbortSignal.timeout(18000) });
    if (!snapRes.ok) return null;
    const html = (await snapRes.json())?.contents ?? null;
    if (html && html.length > 500) {
      console.log("[BulkFree] Wayback Machine hit:", url);
      return html;
    }
  } catch { /* fall through */ }
  return null;
};

// Main free fetch — used exclusively by Bulk Scraper
// Returns either:
//   { __browserEmails: string[], __screenshot: string|null }  ← browser extract mode (compact)
//   string (HTML)                                              ← all other fallbacks
const fetchHtmlFree = async (url) => {
  // 1. Cloudflare Browser extract mode (best — real Chromium + DOM extraction, free)
  // Returns { __browserEmails, __screenshot } ONLY when emails were found.
  // If browser loads page but finds 0 emails, falls through to cache layers below.
  const fromBrowser = await fetchBulkBrowserRendering(url);
  if (fromBrowser?.__browserEmails?.length > 0) return fromBrowser;

  // Save screenshot from browser attempt if available — passed through at the end
  const browserScreenshot = fromBrowser?.__screenshot ?? null;

  // 2. Mobile UA direct fetch via worker — free, fast, 0ms on failure.
  //    Many sites serve simpler/less-protected pages to mobile UA.
  //    Replaces Google Cache which added 15s timeout and never succeeded.
  const wurl = workerUrlRef.current?.trim();
  if (wurl) {
    try {
      const mobileEndpoint = `${wurl.replace(/\/$/, "")}/fetch?url=${encodeURIComponent(url)}&provider=direct&extract=true&mobile=true`;
      const res = await fetch(mobileEndpoint, { signal: AbortSignal.timeout(10000) });
      if (res.ok) {
        const data = await res.json();
        if (Array.isArray(data?.emails) && data.emails.length > 0) {
          console.log("[BulkFree] Mobile direct got", data.emails.length, "emails for:", url);
          return { __browserEmails: data.emails, __screenshot: null };
        }
      }
    } catch (e) { console.warn("[BulkFree] Mobile direct failed:", e.message); }
  }

  // 3. Wayback Machine (archived snapshot fallback)
  const fromWayback = await fetchWaybackMachine(url);
  if (fromWayback) return fromWayback;

  // 4. CORS proxy rotation (existing behaviour — works for simple sites)
  console.log("[BulkFree] All cache layers missed — falling back to CORS proxy:", url);
  const fromProxy = await fetchHtmlSmart(url);
  if (fromProxy) return fromProxy;

  // 5. Nothing worked — return browser screenshot if we have one so UI can show it
  if (browserScreenshot) return { __browserEmails: [], __screenshot: browserScreenshot };
  return null;
};

// ── Contact page URLs to try when homepage yields no emails ──────────────
const CONTACT_PATHS = ["/contact", "/contact-us", "/about", "/about-us", "/our-team", "/team", "/staff"];

const buildContactUrls = (baseUrl) => {
  try {
    const { origin } = new URL(baseUrl);
    return CONTACT_PATHS.map(p => `${origin}${p}`);
  } catch { return []; }
};

// Fetch + extract emails from a list of contact page URLs via worker
// Returns { emails, foundUrl } as soon as one page yields emails, or { emails: [] }
const crawlContactPages = async (baseUrl, providerHint, extractFn) => {
  const contactUrls = buildContactUrls(baseUrl);
  for (const cUrl of contactUrls) {
    try {
      console.log("[ContactCrawler] Trying", cUrl);
      const result = await fetchHtmlViaWorker(cUrl, providerHint);
      if (!result?.html) continue;
      const emails = extractFn(result.html);
      if (emails.length > 0) {
        console.log("[ContactCrawler] Found", emails.length, "emails at", cUrl);
        return { emails, foundUrl: cUrl };
      }
    } catch (e) {
      console.warn("[ContactCrawler] Failed for", cUrl, e.message);
    }
  }
  return { emails: [], foundUrl: null };
};

// ── extractEmailsSmart — decoupled, provider-aware ────────────────────────
// Each provider runs fully independently. No provider requires another.
const extractEmailsSmart = async (selectedProvider, { geminiKey, groqKey, scraperApiKey, zenRowsKey }, url) => {
  const parseEmails = (raw) => {
    const clean = raw.replace(/```json|```/g, "").trim();
    const match = clean.match(/\[[\s\S]*\]/);
    if (!match) return [];
    try { return JSON.parse(match[0]).filter(e => typeof e === "string" && e.includes("@")); }
    catch { return []; }
  };

  const stripToText = (html) => html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s{3,}/g, " ")
    .slice(0, 12000);

  const emailPrompt = (pageText) =>
    `You are an expert email extractor. Your ONLY job is to find email addresses in webpage text.

INSTRUCTIONS:
- Extract ALL email addresses you can find
- Look carefully for: standard emails, emails in footers, contact sections, attorney profiles, staff pages
- Also detect obfuscated emails like: "name [at] domain.com", "name AT domain DOT com", "name(at)domain.com"
- Decode any HTML entities like &#64; (which is @)
- Do NOT hallucinate or make up emails — only return ones actually present in the text
- Return ONLY a valid JSON array of email strings
- Example output: ["info@company.com","john@lawfirm.com"]
- If absolutely no emails found, return exactly: []
- No explanation, no markdown, no extra text — ONLY the JSON array

Webpage text:
${pageText}`;

  const groqEmailPrompt = (pageText) =>
    `You are a precise email address extractor. Find every email address in this webpage content.

SEARCH FOR:
1. Standard format: name@domain.com
2. Obfuscated: "name [at] domain [dot] com" or "name AT domain DOT com"  
3. In footer, contact, about, staff, attorney profile sections
4. In mailto: links
5. HTML encoded: &#64; means @

RULES:
- Only return emails actually present in the text — never invent them
- Include ALL emails found, even if there are many
- Return ONLY a JSON array, nothing else
- Example: ["info@firm.com","john.doe@firm.com","contact@firm.com"]
- If none found: []

Page content:
${pageText}`;

  // ── provider: gemini ───────────────────────────────────────────────────────
  if (selectedProvider === "gemini") {
    if (!geminiKey?.trim()) throw new Error("Gemini API key required");
    const { html, provider: fetchProvider } = await fetchHtmlPremium(url, scraperApiKey, zenRowsKey);
    if (!html) throw new Error("Could not fetch page (site may block scrapers)");
    const raw = await callGemini(geminiKey, emailPrompt(stripToText(html)));
    let emails = parseEmails(raw);
    console.log("[extractEmailsSmart] Gemini homepage found %d emails via fetch:%s", emails.length, fetchProvider);
    // ── Contact page fallback ──────────────────────────────────────────────
    if (emails.length === 0) {
      console.log("[extractEmailsSmart] Gemini: 0 emails on homepage — crawling contact pages");
      const contactUrls = buildContactUrls(url);
      for (const cUrl of contactUrls) {
        try {
          const { html: cHtml } = await fetchHtmlPremium(cUrl, scraperApiKey, zenRowsKey);
          if (!cHtml) continue;
          const cRaw = await callGemini(geminiKey, emailPrompt(stripToText(cHtml)));
          const cEmails = parseEmails(cRaw);
          if (cEmails.length > 0) {
            console.log("[extractEmailsSmart] Gemini found", cEmails.length, "emails at", cUrl);
            emails = cEmails;
            break;
          }
        } catch (e) { console.warn("[ContactCrawler/Gemini] Failed for", cUrl, e.message); }
      }
    }
    return { emails, provider: `${fetchProvider}+gemini` };
  }

  // ── provider: groq ─────────────────────────────────────────────────────────
  if (selectedProvider === "groq") {
    if (!groqKey?.trim()) throw new Error("Groq API key required");

    // Fetch via worker first (auto mode: ScraperAPI → ZenRows) for best HTML quality
    // Falls back to CORS proxy only if worker is not configured
    let html = null;
    let fetchUsed = "proxy";

    if (workerUrlRef.current?.trim()) {
      console.log("[Groq] Fetching via worker (auto mode)");
      const result = await fetchHtmlViaWorker(url, "auto");
      if (result?.html) {
        html = result.html;
        fetchUsed = result.provider ?? "worker";
        console.log("[Groq] Worker fetch success via", fetchUsed);
      } else {
        console.warn("[Groq] Worker failed —", result?.error, "— falling back to CORS proxy");
      }
    }

    // Fallback to CORS proxy if worker not set or failed
    if (!html) {
      console.log("[Groq] Using CORS proxy fallback");
      html = await fetchHtmlSmart(url);
      fetchUsed = "proxy";
    }

    if (!html) throw new Error("Could not fetch page — worker and proxy both failed");

    const raw = await callGroq(groqKey, groqEmailPrompt(stripToText(html)));
    let emails = parseEmails(raw);
    console.log("[extractEmailsSmart] Groq homepage found %d emails via %s", emails.length, fetchUsed);

    // ── Contact page fallback ──────────────────────────────────────────────
    if (emails.length === 0) {
      console.log("[extractEmailsSmart] Groq: 0 emails on homepage — crawling contact pages");
      const contactUrls = buildContactUrls(url);
      for (const cUrl of contactUrls) {
        try {
          let cHtml = null;
          if (workerUrlRef.current?.trim()) {
            const result = await fetchHtmlViaWorker(cUrl, "auto");
            if (result?.html) cHtml = result.html;
          }
          if (!cHtml) cHtml = await fetchHtmlSmart(cUrl);
          if (!cHtml) continue;
          const cRaw = await callGroq(groqKey, groqEmailPrompt(stripToText(cHtml)));
          const cEmails = parseEmails(cRaw);
          if (cEmails.length > 0) {
            console.log("[extractEmailsSmart] Groq found", cEmails.length, "emails at", cUrl);
            emails = cEmails;
            break;
          }
        } catch (e) { console.warn("[ContactCrawler/Groq] Failed for", cUrl, e.message); }
      }
    }
    return { emails, provider: `${fetchUsed}+groq` };
  }

  // ── provider: scraperapi ───────────────────────────────────────────────────
  if (selectedProvider === "scraperapi") {
    if (!scraperApiKey?.trim()) throw new Error("ScraperAPI key required");
    if (!workerUrlRef.current?.trim()) throw new Error("ScraperAPI requires the Cloudflare Worker — add your Worker URL above");
    console.log("[extractEmailsSmart] Using provider: scraperapi");
    const result = await fetchHtmlViaWorker(url, "scraperapi");
    if (!result?.html) throw new Error(`ScraperAPI failed: ${result?.error ?? "no response"}`);
    let emails = cleanEmails(extractEmailsFromHtml(result.html));
    console.log("[extractEmailsSmart] ScraperAPI homepage found %d emails", emails.length);
    // ── Contact page fallback ──────────────────────────────────────────────
    if (emails.length === 0) {
      console.log("[extractEmailsSmart] ScraperAPI: 0 emails on homepage — crawling contact pages");
      const { emails: contactEmails } = await crawlContactPages(url, "scraperapi", (html) => cleanEmails(extractEmailsFromHtml(html)));
      if (contactEmails.length > 0) emails = contactEmails;
    }
    return { emails, provider: "scraperapi" };
  }

  // ── provider: zenrows ──────────────────────────────────────────────────────
  if (selectedProvider === "zenrows") {
    if (!zenRowsKey?.trim()) throw new Error("ZenRows API key required");
    if (!workerUrlRef.current?.trim()) throw new Error("ZenRows requires the Cloudflare Worker — add your Worker URL above");
    console.log("[extractEmailsSmart] Using provider: zenrows");
    const result = await fetchHtmlViaWorker(url, "zenrows");
    if (!result?.html) throw new Error(`ZenRows failed: ${result?.error ?? "no response"}`);
    let emails = cleanEmails(extractEmailsFromHtml(result.html));
    console.log("[extractEmailsSmart] ZenRows homepage found %d emails", emails.length);
    // ── Contact page fallback ──────────────────────────────────────────────
    if (emails.length === 0) {
      console.log("[extractEmailsSmart] ZenRows: 0 emails on homepage — crawling contact pages");
      const { emails: contactEmails } = await crawlContactPages(url, "zenrows", (html) => cleanEmails(extractEmailsFromHtml(html)));
      if (contactEmails.length > 0) emails = contactEmails;
    }
    return { emails, provider: "zenrows" };
  }

  // ── provider: auto (full fallback chain) ───────────────────────────────────
  if (selectedProvider === "auto") {
    const { html: premiumHtml, provider: fetchProvider } = await fetchHtmlPremium(url, scraperApiKey, zenRowsKey);
    if (!premiumHtml) throw new Error("Could not fetch page (all fetch layers failed)");
    // ── Contact page fallback for auto mode ───────────────────────────────
    let autoEmails = cleanEmails(extractEmailsFromHtml(premiumHtml));
    if (autoEmails.length === 0 && workerUrlRef.current?.trim()) {
      console.log("[Auto] 0 emails on homepage — crawling contact pages");
      const { emails: contactEmails } = await crawlContactPages(url, "auto", (html) => cleanEmails(extractEmailsFromHtml(html)));
      if (contactEmails.length > 0) {
        console.log("[Auto] Contact pages found", contactEmails.length, "emails");
        return { emails: contactEmails, provider: `${fetchProvider}+contact` };
      }
    }
    const pageText = stripToText(premiumHtml);

    if (geminiKey?.trim()) {
      try {
        const raw = await callGemini(geminiKey, emailPrompt(pageText));
        const emails = parseEmails(raw);
        console.log("[Auto] Gemini found %d emails via fetch:%s", emails.length, fetchProvider);
        return { emails, provider: `${fetchProvider}+gemini` };
      } catch (err) { console.warn("[Auto] Gemini failed:", err.message, "— trying Groq"); }
    }
    if (groqKey?.trim()) {
      try {
        const raw = await callGroq(groqKey, groqEmailPrompt(pageText));
        const emails = parseEmails(raw);
        console.log("[Auto] Groq found %d emails via fetch:%s", emails.length, fetchProvider);
        return { emails, provider: `${fetchProvider}+groq` };
      } catch (err) { console.warn("[Auto] Groq failed:", err.message, "— falling back to regex"); }
    }
    const emails = cleanEmails(extractEmailsFromHtml(premiumHtml));
    console.log("[Auto] Regex found %d emails via fetch:%s", emails.length, fetchProvider);
    return { emails, provider: `${fetchProvider}+regex` };
  }

  throw new Error(`Unknown provider: ${selectedProvider}`);
};

// ── FREE bulk email extractor via proxy + regex ────────────────────────────
const EMAIL_REGEX = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
const SKIP_EMAILS = /\.(png|jpg|gif|svg|webp|css|js)$/i;
const JUNK_DOMAINS = ["sentry.io","wixpress.com","example.com","domain.com","email.com","yoursite","youremail","user@","name@","test@","noreply","no-reply","donotreply","do-not-reply","support@sentry","privacy@","legal@wix"];

const cleanEmails = (raw) => {
  const unique = [...new Set((raw || []).map(e => e.toLowerCase().trim()))];
  return unique.filter(e => {
    if (SKIP_EMAILS.test(e)) return false;
    if (JUNK_DOMAINS.some(j => e.includes(j))) return false;
    if (e.length > 80) return false;
    return true;
  });
};

// Decode HTML entities and obfuscated emails
const decodeObfuscated = (html) => {
  return html
    // HTML entities: &#64; → @, &#46; → .
    .replace(/&#64;/g, "@").replace(/&#46;/g, ".")
    .replace(/&#x40;/gi, "@").replace(/&#x2e;/gi, ".")
    // &amp; etc
    .replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")
    // [at] / (at) / " at " obfuscation
    .replace(/\s*[\[(]at[\])]\s*/gi, "@")
    .replace(/\s+at\s+/gi, "@")
    // [dot] / (dot) / " dot " obfuscation
    .replace(/\s*[\[(]dot[\])]\s*/gi, ".")
    .replace(/\s+dot\s+/gi, ".");
};

// Extract mailto: href emails directly from raw HTML
const extractMailtoEmails = (html) => {
  const results = [];
  const re = /href=["']mailto:([^"'?]+)/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    const email = m[1].trim().toLowerCase();
    if (email.includes("@")) results.push(email);
  }
  return results;
};

// Decode Cloudflare email protection (data-cfemail attribute)
// CF encodes emails as hex with XOR — this reverses it
const decodeCFEmail = (encoded) => {
  try {
    const r = parseInt(encoded.substr(0, 2), 16);
    let email = "";
    for (let n = 2; n < encoded.length; n += 2) {
      email += String.fromCharCode(parseInt(encoded.substr(n, 2), 16) ^ r);
    }
    return email.toLowerCase().trim();
  } catch { return null; }
};

// Extract all Cloudflare-protected emails from HTML
const extractCFEmails = (html) => {
  const results = [];
  // Pattern 1: data-cfemail attribute
  const re1 = /data-cfemail=["']([a-f0-9]+)["']/gi;
  let m;
  while ((m = re1.exec(html)) !== null) {
    const decoded = decodeCFEmail(m[1]);
    if (decoded && decoded.includes("@")) results.push(decoded);
  }
  // Pattern 2: CF email protection link with encoded value in URL
  const re2 = /\/cdn-cgi\/l\/email-protection#([a-f0-9]+)/gi;
  while ((m = re2.exec(html)) !== null) {
    const decoded = decodeCFEmail(m[1]);
    if (decoded && decoded.includes("@")) results.push(decoded);
  }
  return results;
};

// Decode unicode escape sequences like @ → @
const decodeUnicode = (html) => {
  return html.replace(/\u([\dA-Fa-f]{4})/g, (_, hex) =>
    String.fromCharCode(parseInt(hex, 16))
  );
};

// Fetch a single URL — delegates to smart rotating proxy engine
const fetchHtml = async (url) => fetchHtmlSmart(url);

// ── Premium fetch: ScraperAPI → ZenRows → CORS fallback ───────────────────
// ONLY used by extractEmailsSmart and single-URL extract.
// fetchHtml / fetchPageEmails / bulk scraper are NEVER touched.
// ── Cloudflare Worker fetch (ScraperAPI → ZenRows, no CORS issues) ────────
// Set WORKER_URL to your deployed worker, e.g.:
//   https://scraper-proxy.YOUR-SUBDOMAIN.workers.dev
// If not set, falls back to free CORS proxy rotation.
// workerUrlRef is set by the React component so fetchHtmlViaWorker always reads the latest value
const workerUrlRef = { current: localStorage.getItem("workerUrl") ?? "https://scraper-proxy.scraper-proxy-01.workers.dev/" };
// API key refs — kept in sync with React state so module-level fetch functions can read them
const scraperApiKeyRef = { current: localStorage.getItem("scraperApiKey") ?? "" };
const zenRowsKeyRef    = { current: localStorage.getItem("zenRowsKey") ?? "" };

const fetchHtmlViaWorker = async (targetUrl, providerHint = "auto") => {
  const _wurl = workerUrlRef.current ?? "";
  if (!_wurl.trim()) return null;
  const isPaid = providerHint === "scraperapi" || providerHint === "zenrows";
  const clientTimeoutMs = isPaid ? 90000 : 70000;
  try {
    // Pass API keys in URL params — worker reads them with fallback to env vars.
    // This lets users supply their own keys without needing Cloudflare env vars set.
    const params = new URLSearchParams({
      url: targetUrl,
      provider: providerHint,
      ...(scraperApiKeyRef.current?.trim() && { scraper_api_key: scraperApiKeyRef.current.trim() }),
      ...(zenRowsKeyRef.current?.trim()    && { zenrows_api_key: zenRowsKeyRef.current.trim() }),
    });
    const endpoint = `${_wurl.replace(/\/$/, "")}/fetch?${params}`;
    console.log("[Worker] provider:", providerHint, "→", targetUrl, `(timeout: ${clientTimeoutMs}ms)`);
    const res = await fetch(endpoint, { signal: AbortSignal.timeout(clientTimeoutMs) });
    if (!res.ok) {
      const body = await res.text().catch(() => "");
      console.warn("[Worker] HTTP", res.status, body.slice(0, 200));
      return { html: null, provider: "error", error: `HTTP ${res.status}: ${body.slice(0, 120)}` };
    }
    const html = await res.text();
    const provider = res.headers.get("X-Fetched-By") ?? providerHint;
    if (html && html.length > 200) {
      console.log("[Worker] Success via", provider, "—", html.length, "chars");
      return { html, provider };
    }
    console.warn("[Worker] Empty response body from", provider);
    return { html: null, provider: "error", error: "Empty response body" };
  } catch (err) {
    console.warn("[Worker] Request failed:", err.message);
    return { html: null, provider: "error", error: err.message };
  }
};

const fetchHtmlPremium = async (url, _scraperApiKey, _zenRowsKey) => {
  // ── Layer 1: Cloudflare Worker in auto mode (ScraperAPI → ZenRows) ────────
  if (workerUrlRef.current?.trim()) {
    console.log("[Premium] Using provider: auto →", url);
    const result = await fetchHtmlViaWorker(url, "auto");
    if (result?.html) {
      console.log("[Premium] Worker success via", result.provider);
      return result;
    }
    console.warn("[Premium] Worker failed —", result?.error, "— falling through to CORS proxy");
  }

  // ── Layer 2: Free CORS proxy rotation (fallback when no Worker or Worker failed) ──
  console.log("[Premium] Using provider: proxy →", url);
  const html = await fetchHtmlSmart(url);
  return { html, provider: "proxy" };
};

// ── Groq LLM helper (llama-3.3-70b-versatile) ─────────────────────────────
const callGroq = async (groqApiKey, prompt) => {
  console.log("[Groq] Calling Groq llama-3.3-70b-versatile");
  const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${groqApiKey.trim()}`,
    },
    body: JSON.stringify({
      model: "llama-3.3-70b-versatile",
      messages: [{ role: "user", content: prompt }],
      temperature: 0,
      max_tokens: 1024,
    }),
    signal: AbortSignal.timeout(20000),
  });
  if (!res.ok) throw new Error(`Groq HTTP ${res.status}`);
  const data = await res.json();
  const text = data?.choices?.[0]?.message?.content ?? "";
  console.log("[Groq] Response received — %d chars", text.length);
  return text;
};

// ── Hunter.io + Apollo.io ─────────────────────────────────────────────────

const SKIP_DOMAINS_FOR_APIS = [
  "facebook.com","instagram.com","twitter.com","x.com","linkedin.com",
  "yelp.com","yellowpages.com","angi.com","thumbtack.com","houzz.com",
  "tripadvisor.com","trustpilot.com","g2.com","capterra.com",
];

const isSkippedDomain = (url) => {
  try {
    const host = new URL(url).hostname.replace(/^www\./, "");
    return SKIP_DOMAINS_FOR_APIS.some(d => host === d || host.endsWith("." + d));
  } catch { return false; }
};

// ── getDomainFromUrl — strips www, subdomain-safe, path-safe ─────────────
const getDomain = (url) => {
  try { return new URL(url).hostname.replace(/^www\./, ""); } catch { return null; }
};

const getDomainFromUrl = (url) => {
  try {
    const hostname = new URL(url).hostname.replace(/^www\./, "");
    // Strip common subdomains that Hunter won't recognise (locations, blog, shop etc.)
    // Keep only root domain: e.g. locations.chaffinluhana.com → chaffinluhana.com
    const parts = hostname.split(".");
    if (parts.length > 2) {
      // check if TLD is 2-part (e.g. .co.uk) — keep last 3, else last 2
      const twoPartTlds = ["co.uk","com.au","co.nz","co.za","com.br","co.in"];
      const lastTwo = parts.slice(-2).join(".");
      return twoPartTlds.includes(lastTwo) ? parts.slice(-3).join(".") : parts.slice(-2).join(".");
    }
    return hostname;
  } catch { return null; }
};

// ── CapacitorHttp helper — native HTTP, bypasses WebView CORS on Android/iOS
// Falls back to window.fetch when not running inside Capacitor.
const capacitorHttp = async (method, url, headers = {}, data = null) => {
  try {
    const { CapacitorHttp } = await import('@capacitor/core');
    const opts = { method, url, headers };
    if (data !== null) opts.data = data;
    const res = await CapacitorHttp.request(opts);
    // CapacitorHttp returns { status, data } where data is already parsed if JSON
    const text = typeof res.data === "string" ? res.data : JSON.stringify(res.data);
    return { ok: res.status >= 200 && res.status < 300, text: () => Promise.resolve(text) };
  } catch {
    // Not in Capacitor (e.g. plain browser) — fall back to fetch
    const opts = { method, headers: { "Content-Type": "application/json", ...headers } };
    if (data !== null) opts.body = JSON.stringify(data);
    return fetch(url, { ...opts, signal: AbortSignal.timeout(12000) });
  }
};

// ── fetchHunterEmails — Snov.io domain search (replaces Hunter.io) ────────
// Snov.io API: GET /v2/domain-emails-with-info?domain=DOMAIN&apiSecret=KEY
// Uses Cloudflare Worker when set, otherwise CapacitorHttp (native — no CORS).
const fetchHunterEmails = async (url, apiKey, cache, workerUrl, apiSecret) => {
  const domain = getDomainFromUrl(url);
  if (!domain) {
    console.warn("[Snov] Could not extract domain from:", url);
    return [];
  }
  console.log("[Snov] Domain extracted:", domain);
  if (cache.hunter[domain] !== undefined) {
    console.log("[Snov] Cache hit for:", domain, "→", cache.hunter[domain]);
    return cache.hunter[domain];
  }

  // Snov.io requires getting an access token first, then using it
  try {
    // Step 1: Get access token
    const tokenEndpoint = "https://api.snov.io/v1/oauth/access_token";
    const tokenPayload = {
      grant_type: "client_credentials",
      client_id: apiKey.trim(),
      client_secret: (apiSecret ?? apiKey).trim(),
    };

    let tokenRes;
    if (workerUrl?.trim()) {
      const proxied = `${workerUrl.replace(/\/$/, "")}/fetch?url=${encodeURIComponent(tokenEndpoint)}`;
      console.log("[Snov] Getting token via worker proxy");
      tokenRes = await fetch(proxied, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(tokenPayload),
        signal: AbortSignal.timeout(15000),
      });
    } else {
      console.log("[Snov] Getting token via CapacitorHttp (native)");
      tokenRes = await capacitorHttp("POST", tokenEndpoint, { "Content-Type": "application/json" }, tokenPayload);
    }

    const tokenText = await tokenRes.text();
    console.log("[Snov] Token response:", tokenText.slice(0, 200));
    let tokenData;
    try { tokenData = JSON.parse(tokenText); } catch {
      console.error("[Snov] Failed to parse token response");
      cache.hunter[domain] = [];
      return [];
    }

    const accessToken = tokenData?.access_token;
    if (!accessToken) {
      console.error("[Snov] No access token returned:", tokenData);
      cache.hunter[domain] = [];
      return [];
    }

    // Step 2: Search emails for domain
    // Pass access_token as query param to avoid CORS header restriction
    const searchEndpoint = `https://api.snov.io/v2/domain-emails-with-info?domain=${encodeURIComponent(domain)}&type=all&limit=10&lastId=0&access_token=${encodeURIComponent(accessToken)}`;
    let searchRes;
    if (workerUrl?.trim()) {
      const proxied = `${workerUrl.replace(/\/$/, "")}/fetch?url=${encodeURIComponent(searchEndpoint)}`;
      console.log("[Snov] Searching domain via worker proxy:", domain);
      searchRes = await fetch(proxied, { signal: AbortSignal.timeout(15000) });
    } else {
      console.log("[Snov] Searching domain via CapacitorHttp (native):", domain);
      searchRes = await capacitorHttp("GET", searchEndpoint);
    }

    const searchText = await searchRes.text();
    console.log("[Snov] Search response:", searchText.slice(0, 300));

    let searchData;
    try { searchData = JSON.parse(searchText); } catch {
      console.error("[Snov] Failed to parse search response");
      cache.hunter[domain] = [];
      return [];
    }

    console.log("[Snov] Response data:", searchData);

    if (searchData?.error || searchData?.message) {
      console.error("[Snov] API error:", searchData.error ?? searchData.message);
      cache.hunter[domain] = [];
      return [];
    }

    const rawContacts = searchData?.emails || searchData?.data || [];

    // Extract full contact info — name, title, email, seniority
    const contacts = rawContacts.map(e => ({
      email: (e.email || e.value)?.toLowerCase().trim() ?? null,
      firstName: e.first_name ?? null,
      lastName: e.last_name ?? null,
      name: [e.first_name, e.last_name].filter(Boolean).join(" ") || null,
      title: e.position ?? e.title ?? null,
      seniority: e.seniority ?? null,
      confidence: e.confidence ?? null,
    })).filter(c => c.email && c.email.includes("@"));

    const emails = contacts.map(c => c.email);

    // Identify decision makers by title/seniority
    const DECISION_TITLES = ["owner","ceo","founder","co-founder","president","managing","director","partner","principal"];
    const decisionMakers = contacts.filter(c =>
      c.seniority === "executive" ||
      DECISION_TITLES.some(t => (c.title ?? "").toLowerCase().includes(t))
    );

    if (emails.length === 0) {
      console.log("[Snov] Returned no emails for domain:", domain);
    } else {
      console.log("[Snov] Found contacts:", contacts);
      if (decisionMakers.length) console.log("[Snov] Decision makers:", decisionMakers);
    }

    // Cache both emails and full contacts
    cache.hunter[domain] = emails;
    cache.hunter[`${domain}_contacts`] = contacts;
    cache.hunter[`${domain}_dm`] = decisionMakers;
    return emails;
  } catch (e) {
    console.error("[Snov] Error for", domain, ":", e.message);
    cache.hunter[domain] = [];
    return [];
  }
};

// ── fetchApolloEnrichment — native HTTP via CapacitorHttp, no CORS issues ─
const fetchApolloEnrichment = async (url, apiKey, cache, workerUrl) => {
  const domain = getDomainFromUrl(url);
  if (!domain) {
    console.warn("[Apollo] Could not extract domain from:", url);
    return null;
  }
  console.log("[Apollo] Domain extracted:", domain);
  if (cache.apollo[domain] !== undefined) {
    console.log("[Apollo] Cache hit for:", domain);
    return cache.apollo[domain];
  }

  // Apollo now requires X-Api-Key header — pass key via URL param to worker
  // Worker encodes it into the proxied request headers
  const apolloEndpoint = `https://api.apollo.io/v1/organizations/enrich?api_key=${encodeURIComponent(apiKey.trim())}`;
  const apolloPayload = { domain };

  try {
    let res;
    if (workerUrl?.trim()) {
      // Route via worker — pass api_key in URL so worker can forward it as X-Api-Key
      const proxied = `${workerUrl.replace(/\/$/, "")}/fetch?url=${encodeURIComponent(apolloEndpoint)}`;
      console.log("[Apollo] Calling via worker proxy:", proxied);
      res = await fetch(proxied, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(apolloPayload),
        signal: AbortSignal.timeout(15000),
      });
    } else {
      console.log("[Apollo] Calling via CapacitorHttp (native — no CORS)");
      res = await capacitorHttp("POST", apolloEndpoint, {
        "Content-Type": "application/json",
        "Cache-Control": "no-cache",
        "X-Api-Key": apiKey.trim(),
      }, apolloPayload);
    }

    const text = await res.text();
    console.log("[Apollo] Raw response:", text.slice(0, 300));

    let data;
    try { data = JSON.parse(text); } catch {
      console.error("[Apollo] Failed to parse JSON response");
      cache.apollo[domain] = null;
      return null;
    }

    console.log("[Apollo] Response:", data);

    if (data?.error || data?.message) {
      console.error("[Apollo] API error:", data.error ?? data.message);
      cache.apollo[domain] = null;
      return null;
    }

    const org = data?.organization;
    if (!org) {
      console.log("[Apollo] No organization data returned for:", domain);
      cache.apollo[domain] = null;
      return null;
    }

    const enrichment = {
      name: org.name ?? null,
      industry: org.industry ?? null,
      size: org.estimated_num_employees ?? null,
      city: org.city ?? null,
      country: org.country ?? null,
      linkedin: org.linkedin_url ?? null,
      decisionMakers: null, // populated separately by fetchApolloDecisionMakers
    };
    console.log("[Apollo] Enrichment found:", enrichment);
    cache.apollo[domain] = enrichment;
    return enrichment;
  } catch (e) {
    console.error("[Apollo] Error for", domain, ":", e.message);
    cache.apollo[domain] = null;
    return null;
  }
};

// ── Apollo people search — finds CEO/decision maker contacts ─────────────
const fetchApolloDecisionMakers = async (url, apiKey, cache, workerUrl) => {
  const domain = getDomainFromUrl(url);
  if (!domain) return null;
  const cacheKey = `people_${domain}`;
  if (cache.apollo[cacheKey] !== undefined) return cache.apollo[cacheKey];

  // Titles to search for — ordered by seniority
  const DECISION_MAKER_TITLES = [
    "CEO", "Chief Executive Officer", "Founder", "Co-Founder",
    "Owner", "Managing Partner", "President", "Managing Director",
    "Principal", "Director", "Partner",
  ];

  try {
    const searchEndpoint = `https://api.apollo.io/v1/mixed_people/search?api_key=${encodeURIComponent(apiKey.trim())}`;
    const searchBody = JSON.stringify({
      q_organization_domains: domain,
      person_titles: DECISION_MAKER_TITLES,
      page: 1,
      per_page: 3,
    });

    let res;
    if (workerUrl?.trim()) {
      const proxied = `${workerUrl.replace(/\/$/, "")}/fetch?url=${encodeURIComponent(searchEndpoint)}`;
      console.log("[Apollo People] Searching decision makers for:", domain);
      res = await fetch(proxied, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: searchBody,
        signal: AbortSignal.timeout(15000),
      });
    } else {
      console.log("[Apollo People] Calling via CapacitorHttp (native — no CORS)");
      res = await capacitorHttp("POST", searchEndpoint, {
        "Content-Type": "application/json",
        "X-Api-Key": apiKey.trim(),
      }, JSON.parse(searchBody));
    }

    const text = await res.text();
    console.log("[Apollo People] Raw response:", text.slice(0, 400));

    let data;
    try { data = JSON.parse(text); } catch {
      console.error("[Apollo People] Failed to parse response");
      cache.apollo[cacheKey] = null;
      return null;
    }

    const people = data?.people || data?.contacts || [];
    if (!people.length) {
      console.log("[Apollo People] No decision makers found for:", domain);
      cache.apollo[cacheKey] = null;
      return null;
    }

    const decisionMakers = people.map(p => ({
      name: [p.first_name, p.last_name].filter(Boolean).join(" ") || null,
      title: p.title ?? null,
      email: p.email ?? p.sanitized_email ?? null,
      linkedin: p.linkedin_url ?? null,
    })).filter(p => p.name || p.email);

    console.log("[Apollo People] Found decision makers:", decisionMakers);
    cache.apollo[cacheKey] = decisionMakers;
    return decisionMakers;
  } catch (e) {
    console.error("[Apollo People] Error for", domain, ":", e.message);
    cache.apollo[cacheKey] = null;
    return null;
  }
};

// Extract all emails from an html string
const extractEmailsFromHtml = (html) => {
  // 1. Cloudflare email protection — must run on RAW html before any transforms
  const fromCF = extractCFEmails(html);

  // 2. Decode unicode escapes first (e.g. \u0040 → @)
  const unicodeDecoded = decodeUnicode(html);

  // 3. Decode obfuscation ([at], (at), &#64; etc)
  const decoded = decodeObfuscated(unicodeDecoded);

  // 4. Additional entity patterns not handled by decodeObfuscated:
  //    &#x40; &#64; &commat; (all mean @)
  //    &#46; &#x2e; (both mean .)
  //    Also catch space-separated obfuscation: "user @ domain . com"
  const entityDecoded = decoded
    .replace(/&commat;/gi, "@")
    .replace(/&#0*64;/g, "@")
    .replace(/&#x0*40;/gi, "@")
    .replace(/&#0*46;/g, ".")
    .replace(/&#x0*2e;/gi, ".")
    // Space-padded @ and dot common in plain-text obfuscation
    .replace(/\s+@\s+/g, "@")
    .replace(/(\w)\s+\.\s+(\w)/g, "$1.$2");

  // 5. Strip scripts/styles/comments so regex doesn't match junk
  const stripped = entityDecoded
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ");

  // 6. mailto: links from raw html (before any transforms to avoid mutation)
  const fromMailto = extractMailtoEmails(html);

  // 7. JSON-LD structured data — many sites embed contact email here
  const fromJsonLd = (() => {
    const results = [];
    const re = /<script[^>]+type=["']application\/ld\+json["'][^>]*>([\s\S]*?)<\/script>/gi;
    let m;
    while ((m = re.exec(html)) !== null) {
      try {
        const text = m[1];
        (text.match(EMAIL_REGEX) || []).forEach(e => results.push(e.toLowerCase().trim()));
      } catch {}
    }
    return results;
  })();

  // 8. Meta tags — og:email, contact meta, schema.org meta
  const fromMeta = (() => {
    const results = [];
    const re = /<meta[^>]+content=["']([^"']+)["'][^>]*>/gi;
    let m;
    while ((m = re.exec(html)) !== null) {
      const content = m[1];
      (content.match(EMAIL_REGEX) || []).forEach(e => results.push(e.toLowerCase().trim()));
    }
    return results;
  })();

  // 9. Standard regex on cleaned/decoded text
  const fromRegex = stripped.match(EMAIL_REGEX) || [];

  // 10. Merge all sources and deduplicate
  return cleanEmails([...fromCF, ...fromMailto, ...fromJsonLd, ...fromMeta, ...fromRegex]);
};

// Extract all internal links from HTML that are likely to have contact info
const extractInternalLinks = (html, base) => {
  const links = new Set();
  const re = /href=["']([^"'#?]+)/gi;
  let m;
  while ((m = re.exec(html)) !== null) {
    let href = m[1].trim();
    try {
      // Make absolute
      const abs = href.startsWith("http") ? href : new URL(href, base).href;
      const u = new URL(abs);
      // Only same-origin, no file extensions, no query-heavy pages
      if (u.origin !== new URL(base).origin) continue;
      if (/\.(jpg|jpeg|png|gif|svg|pdf|zip|css|js|xml|json)$/i.test(u.pathname)) continue;
      links.add(u.origin + u.pathname.replace(/\/$/, "") || "/");
    } catch { continue; }
  }
  return [...links];
};

// Score a URL path — higher = more likely to have emails
const scoreUrl = (pathname) => {
  const p = pathname.toLowerCase();
  const HIGH = ["agent","team","staff","people","member","contact","about","broker","realtor","advisor","expert","associate","directory","roster","find","our-","meet-"];
  const LOW = ["blog","news","press","listing","propert","search","buy","sell","rent","gallery","media","testimonial","review","faq","privacy","terms","sitemap","login","signup","cart","checkout"];
  if (LOW.some(k => p.includes(k))) return -1;
  if (HIGH.some(k => p.includes(k))) return 2;
  return 0;
};

const fetchPageEmails = async (url, scraperApiKey = "", zenRowsKey = "", useDeepCrawl = false, maxPages = 20, onProgress = null) => {
  // Normalize — always scrape from origin so seed paths don't 404
  // e.g. "http://site.com/location/Houston" → "https://site.com"
  let base;
  let normalizedUrl;
  try {
    const u = new URL(url.startsWith("http") ? url : "https://" + url);
    base = u.origin;
    normalizedUrl = base; // always start from origin, not a sub-path
  } catch {
    base = url;
    normalizedUrl = url;
  }

  // Always-check pages — all built from origin so none will 404 due to sub-paths
  const seedPages = [
    normalizedUrl,
    `${base}/contact`,
    `${base}/contact-us`,
    `${base}/about`,
    `${base}/about-us`,
    `${base}/team`,
    `${base}/agents`,
    `${base}/staff`,
  ];

  const visited = new Set();
  const allEmails = new Set();
  const queue = [...seedPages];
  let debugScreenshot = null; // captured when browser finds no emails

  const fetchAndExtract = async (pageUrl) => {
    if (visited.has(pageUrl)) return null;
    visited.add(pageUrl);
    try {
      // Free stack: Cloudflare Browser (extract mode) → Mobile direct → Wayback → CORS proxy.
      // Never uses ScraperAPI or ZenRows — zero paid-credit cost.
      const result = await fetchHtmlFree(pageUrl);
      if (!result) return null;

      // Feature 3: if direct fetch signals "host not in allowlist", mark domain
      // as blocked so we skip all remaining seed pages (same host = same block).
      if (result.__blocked) {
        console.log("[BulkScraper] Host blocked — skipping remaining seed pages for:", pageUrl);
        return "__blocked";
      }

      // Browser extract mode returns { __browserEmails, __screenshot } sentinel
      // — emails already extracted server-side, no HTML parsing needed
      if (result.__browserEmails) {
        result.__browserEmails.forEach(e => allEmails.add(e));

        // Debug screenshot — logged to console when no emails found.
        // Open Chrome DevTools (chrome://inspect) or Android Studio Logcat to view.
        // To visualise: copy the base64 string and open in browser as:
        //   data:image/png;base64,<paste here>
        if (result.__browserEmails.length === 0 && result.__screenshot) {
          debugScreenshot = result.__screenshot; // stored for UI debug viewer
          console.warn(`[BulkDebug] No emails found at ${pageUrl} — screenshot stored for debug viewer`);
        }

        // Return a minimal truthy string so deep crawl link extraction still works
        // (link extraction only happens on homepage anyway)
        return result.__browserEmails.length > 0 ? "browser-extract" : null;
      }

      // All other fallbacks return raw HTML — parse client-side as before
      const found = extractEmailsFromHtml(result);
      found.forEach(e => allEmails.add(e));
      return result;
    } catch { return null; }
  };

  // Phase 1 — always scrape seed pages
  for (const pageUrl of seedPages) {
    const html = await fetchAndExtract(pageUrl);
    // Feature 3: if homepage (or any seed) is blocked at provider=direct level,
    // skip all remaining seed pages — they'll hit the same allowlist block.
    if (html === "__blocked") {
      console.log("[BulkScraper] Block detected — aborting seed crawl, jumping to browser fallback");
      break;
    }
    // If homepage found emails, no need for deep crawl unless explicitly enabled
    if (!useDeepCrawl && allEmails.size > 0 && pageUrl === url) break;
    // Collect links from homepage for deep crawl
    if (useDeepCrawl && html && pageUrl === url) {
      const links = extractInternalLinks(html, base);
      // Score and sort — agent/team/contact pages first
      const scored = links
        .filter(l => !visited.has(l))
        .map(l => ({ url: l, score: scoreUrl(new URL(l).pathname) }))
        .filter(l => l.score >= 0)
        .sort((a, b) => b.score - a.score)
        .map(l => l.url);
      // Add to queue (deduped), capped at maxPages
      scored.forEach(l => { if (!queue.includes(l)) queue.push(l); });
    }
  }

  // Phase 2 — deep crawl discovered links
  if (useDeepCrawl) {
    let crawled = seedPages.length;
    for (const pageUrl of queue) {
      if (visited.has(pageUrl)) continue;
      if (crawled >= maxPages) break;
      crawled++;
      if (onProgress) onProgress(`Deep crawl: ${crawled}/${maxPages} pages — ${allEmails.size} emails found`);
      const html = await fetchAndExtract(pageUrl);
      // From each agent/team page, also follow links one level deeper
      if (html) {
        const subLinks = extractInternalLinks(html, base)
          .filter(l => !visited.has(l) && scoreUrl(new URL(l).pathname) > 0);
        subLinks.forEach(l => { if (!queue.includes(l)) queue.push(l); });
      }
      await new Promise(r => setTimeout(r, 200));
    }
  }

  if (allEmails.size === 0 && visited.size <= 1) {
    throw new Error("Could not fetch page (site may block scrapers)");
  }

  return { emails: [...allEmails], screenshot: debugScreenshot };
};

// ── Maps scraper via Google Places API (paginated) ────────────────────────
const mapPlace = (p, extraFields = {}) => ({
  id: uid(),
  placeId: p.id || null,               // canonical Google Place ID — used for dedup
  name: p.displayName?.text || null,
  address: p.formattedAddress || null,
  phone: p.internationalPhoneNumber || null,
  website: p.websiteUri || null,
  rating: typeof p.rating === "number" ? p.rating : null,
  reviews: typeof p.userRatingCount === "number" ? p.userRatingCount : null,
  type: p.primaryTypeDisplayName?.text || null,
  // provenance — null when expansion is not active
  originalKeyword: extraFields.originalKeyword ?? null,
  expandedKeyword: extraFields.expandedKeyword ?? null,
  sourceKeyword:   extraFields.sourceKeyword   ?? null,
});

// ══════════════════════════════════════════════════════════════════════════
// KEYWORD EXPANSION ENGINE
// Architecture:  seed → static → Google Suggest → Groq AI → merge+dedupe
// All three layers are independent; each is additive on top of the previous.
// ══════════════════════════════════════════════════════════════════════════

// ── 1. Static expansion (pure, no side-effects) ──────────────────────────
const expandKeywordsStatic = (seed) => {
  const s = seed.trim();
  const lower = s.toLowerCase();
  const v = new Set([s]);
  const add = (...t) => t.forEach(x => v.add(x));
  const replace = (from, to) => { if (lower.includes(from)) add(s.replace(new RegExp(from, "gi"), to)); };

  if (lower.includes("lawyer") || lower.includes("attorney") || lower.includes("law firm") || lower.includes("legal")) {
    add(`${s} firm`, `${s} office`);
    replace("lawyer", "attorney"); replace("attorney", "lawyer");
    add("law office", "legal services", "litigation attorney", "legal counsel");
  }
  if (lower.includes("doctor") || lower.includes("physician") || lower.includes("medical") || lower.includes("clinic")) {
    replace("doctor", "physician"); replace("clinic", "medical center");
    add(`${s} practice`, `${s} group`, "medical office", "health clinic", "urgent care");
  }
  if (lower.includes("dentist") || lower.includes("dental")) {
    replace("dentist", "dental clinic"); replace("dental clinic", "dentist");
    add(`${s} practice`, `${s} office`, "cosmetic dentist", "family dentistry", "orthodontist");
  }
  if (lower.includes("real estate") || lower.includes("realtor") || lower.includes("property")) {
    replace("realtor", "real estate agent"); replace("real estate agent", "realtor");
    add(`${s} agency`, `${s} broker`, "property management", "real estate company", "home buying agent");
  }
  if (lower.includes("accountant") || lower.includes("cpa") || lower.includes("accounting") || lower.includes("tax")) {
    replace("accountant", "CPA"); replace("CPA", "accountant");
    add(`${s} firm`, `tax ${s}`, "bookkeeping service", "financial advisor", "tax preparation");
  }
  if (lower.includes("insurance")) {
    add(`${s} agent`, `${s} broker`, `${s} company`, `${s} provider`);
  }
  if (lower.includes("roof")) {
    add("roofer", "roof repair", "roofing company", "metal roofing", "roofing services",
        "roofing contractor", "residential roofer", "commercial roofing");
  }
  if (lower.includes("contractor") && !lower.includes("roof")) {
    add(`${s} company`, `${s} service`, `local ${s}`, `${s} near me`);
  }
  if (lower.includes("plumber") || lower.includes("plumbing")) {
    add("plumbing service", "plumbing company", "emergency plumber", "local plumber", "drain repair");
  }
  if (lower.includes("electrician") || lower.includes("electrical")) {
    add("electrical contractor", "electrical service", "licensed electrician", "electrical company");
  }
  if (lower.includes("hvac") || lower.includes("heating") || lower.includes("cooling") || lower.includes("air conditioning")) {
    add("HVAC company", "AC repair", "heating and cooling", "furnace repair", "HVAC contractor");
  }
  if (lower.includes("restaurant") || lower.includes("food") || lower.includes("cafe") || lower.includes("coffee")) {
    replace("restaurant", "eatery"); replace("cafe", "coffee shop");
    add(`${s} near me`, "dining", "bistro", "food service");
  }
  if (lower.includes("gym") || lower.includes("fitness") || lower.includes("yoga") || lower.includes("spa")) {
    add(`${s} studio`, `${s} center`, `${s} near me`, "wellness center", "health club");
  }
  if (lower.includes("auto") || lower.includes("car repair") || lower.includes("mechanic")) {
    add("auto shop", "car repair shop", "auto mechanic", "vehicle repair", "auto service center");
  }
  if (lower.includes("it service") || lower.includes("tech support") || lower.includes("software")) {
    add(`${s} company`, `${s} firm`, "managed IT services", "IT consulting", "tech company");
  }
  // generic fallback — fires when nothing above matched
  if (v.size <= 2) add(`${s} company`, `${s} services`, `${s} near me`, `local ${s}`);

  return [...v];
};

// ── 2. Keyword expansion cache (localStorage, session-scoped) ────────────
const KW_CACHE_KEY = "kwExpansionCache_v1";
const KW_CACHE_TTL = 60 * 60 * 1000; // 1 hour

const kwCache = {
  _load() {
    try { return JSON.parse(localStorage.getItem(KW_CACHE_KEY) || "{}"); } catch { return {}; }
  },
  get(seed, source) {
    const store = this._load();
    const entry = store[`${source}:${seed.toLowerCase()}`];
    if (!entry) return null;
    if (Date.now() - entry.ts > KW_CACHE_TTL) return null; // expired
    return entry.keywords;
  },
  set(seed, source, keywords) {
    try {
      const store = this._load();
      store[`${source}:${seed.toLowerCase()}`] = { keywords, ts: Date.now() };
      localStorage.setItem(KW_CACHE_KEY, JSON.stringify(store));
    } catch { /* storage quota — silently skip */ }
  },
};

// ── 3. Google Suggest expansion ──────────────────────────────────────────
// Uses Google's autocomplete endpoint via a CORS-friendly CDN proxy.
// Returns up to `limit` clean keyword phrases, deduped against `existing`.
const fetchGoogleSuggest = async (seed, limit = 8, existing = new Set()) => {
  const cached = kwCache.get(seed, "suggest");
  if (cached) return cached;

  // Google Suggest doesn't allow direct CORS — proxy through allorigins
  const url = `https://suggestqueries.google.com/complete/search?client=firefox&q=${encodeURIComponent(seed)}`;
  const proxy = `https://api.allorigins.win/get?url=${encodeURIComponent(url)}`;

  const res = await fetch(proxy, { signal: AbortSignal.timeout(8000) });
  if (!res.ok) throw new Error(`Google Suggest proxy HTTP ${res.status}`);
  const wrapper = await res.json();
  const raw = JSON.parse(wrapper.contents);
  // format: [query, [suggestion, ...], ...]
  const suggestions = (raw[1] || [])
    .map(s => s.trim().toLowerCase())
    .filter(s => s && s !== seed.toLowerCase() && !existing.has(s))
    .slice(0, limit);

  kwCache.set(seed, "suggest", suggestions);
  return suggestions;
};

// ── 4. Groq AI expansion ─────────────────────────────────────────────────
// Reuses existing callGroq() — no new provider, no new auth pattern.
const fetchGroqKeywords = async (groqApiKey, seed, limit = 10, existing = new Set()) => {
  const cached = kwCache.get(seed, "groq");
  if (cached) return cached;

  const prompt = `Generate up to ${limit} commonly used business search keywords related to:

"${seed}"

Requirements:
- use real-world business terminology
- optimized for Google Maps business discovery
- short phrases only (2-5 words max)
- avoid uncommon or overly creative wording
- no explanations, no numbering, no duplicates

Return valid JSON only:
{
  "keywords": []
}`;

  const raw = await callGroq(groqApiKey, prompt);
  let keywords = [];
  try {
    const clean = raw.replace(/```json|```/g, "").trim();
    const parsed = JSON.parse(clean);
    keywords = (parsed.keywords || [])
      .map(k => k.trim().toLowerCase())
      .filter(k => k && k.split(" ").length <= 6 && !existing.has(k))
      .slice(0, limit);
  } catch (e) {
    console.warn("[GroqKeywords] JSON parse failed:", e.message);
  }

  kwCache.set(seed, "groq", keywords);
  return keywords;
};

// ── 5. Master merge + dedupe ─────────────────────────────────────────────
// Returns a deduped, capped array of keywords.
// Order of priority: static → suggest → AI
const mergeKeywords = (staticKws, suggestKws, aiKws, maxTotal) => {
  const seen = new Set();
  const out = [];
  for (const kw of [...staticKws, ...suggestKws, ...aiKws]) {
    const norm = kw.trim().toLowerCase();
    if (!norm || seen.has(norm)) continue;
    seen.add(norm);
    out.push(kw.trim());
    if (out.length >= maxTotal) break;
  }
  return out;
};

// ── Secure Storage — dual-layer: Capacitor Preferences + localStorage ────
// Capacitor Preferences writes to Android SharedPreferences which survives
// APK updates, WebView resets, and browser cache clears.
// localStorage is kept as the fast synchronous layer and in-browser fallback.
// On load:  Preferences first → restore into localStorage → return value
// On save:  localStorage immediately + Preferences async backup
// Keys covered: all API keys, secrets, and config that must survive updates.
const SecureStorage = {
  // Keys that get dual-saved
  PROTECTED_KEYS: new Set([
    "groqKey", "placesKey", "geminiKey",
    "scraperApiKey", "zenRowsKey",
    "workerUrl",
    "extractSnovKey", "extractSnovSecret",
    "extractApolloKey",
    "hunterKey", "snovSecret", "apolloKey",
    "hunterEnabled", "apolloEnabled",
    "projectName",
  ]),

  // Internal: get Capacitor Preferences if available
  async _prefs() {
    try {
      if (!window.Capacitor?.isNativePlatform?.()) return null;
      const { Preferences } = await import("@capacitor/preferences");
      return Preferences;
    } catch { return null; }
  },

  // Save — localStorage immediately, Preferences async (non-blocking)
  async set(key, value) {
    try { localStorage.setItem(key, value); } catch {}
    if (!this.PROTECTED_KEYS.has(key)) return;
    try {
      const prefs = await this._prefs();
      if (prefs) await prefs.set({ key, value: String(value) });
    } catch (e) { console.warn("[SecureStorage] Preferences.set failed:", e.message); }
  },

  // Load — try Preferences first, fall back to localStorage
  // Also restores localStorage if Preferences had the value (recovery path)
  async get(key) {
    try {
      const prefs = await this._prefs();
      if (prefs) {
        const { value } = await prefs.get({ key });
        if (value !== null && value !== undefined && value !== "") {
          // Restore into localStorage in case it was wiped
          try { localStorage.setItem(key, value); } catch {}
          return value;
        }
      }
    } catch (e) { console.warn("[SecureStorage] Preferences.get failed:", e.message); }
    // Fallback to localStorage
    try { return localStorage.getItem(key) || ""; } catch { return ""; }
  },

  // Sync-safe get — localStorage only (for useState initialisers which are sync)
  // The async recovery happens in the useEffect below
  getSync(key) {
    try { return localStorage.getItem(key) || ""; } catch { return ""; }
  },
};

const scrapeWithPlaces = async (apiKey, query, location, onProgress, maxPages = 3, onPageResult = null, stopRef = null, extraFields = {}) => {
  const searchText = location ? `${query} in ${location}` : query;
  onProgress(`Searching for "${searchText}"…`);

  const FIELD_MASK = "places.id,places.displayName,places.formattedAddress,places.internationalPhoneNumber,places.websiteUri,places.rating,places.userRatingCount,places.primaryTypeDisplayName,nextPageToken";

  let allPlaces = [];
  let pageToken = null;
  let pageNum = 0;

  while (pageNum < maxPages) {
    if (stopRef && stopRef.current) break;
    pageNum++;

    const body = { textQuery: searchText, maxResultCount: 20 };
    if (pageToken) body.pageToken = pageToken;

    const res = await fetch("https://places.googleapis.com/v1/places:searchText", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": apiKey,
        "X-Goog-FieldMask": FIELD_MASK,
      },
      body: JSON.stringify(body),
    });

    const data = await res.json();
    if (data.error) throw new Error(data.error.message);

    const places = (data.places || []).map(p => mapPlace(p, extraFields));
    allPlaces = [...allPlaces, ...places];
    onProgress(`Page ${pageNum} — ${allPlaces.length} businesses found…`);

    // Stream results page by page so UI updates live
    if (onPageResult) onPageResult(places);

    pageToken = data.nextPageToken || null;
    if (!pageToken) break; // no more pages

    // Required delay — Google needs ~2s before next page token is valid
    await new Promise(r => setTimeout(r, 2000));
  }

  if (!allPlaces.length) throw new Error("No businesses found. Try a different query.");
  return allPlaces;
};

// ══════════════════════════════════════════════════════════════════════════
export default function LeadExtractor() {
  const [tab, setTab] = useState("extract");
  const [mode, setMode] = useState("text");
  const [textInput, setTextInput] = useState("");
  const [urlInput, setUrlInput] = useState("");
  const [file, setFile] = useState(null);
  const [leads, setLeads] = useState(() => { try { const s = localStorage.getItem("leads"); return s ? JSON.parse(s) : []; } catch { return []; } });
  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("");
  const [error, setError] = useState("");
  const [editingCell, setEditingCell] = useState(null);
  const [editVal, setEditVal] = useState("");
  const [showDupOnly, setShowDupOnly] = useState(false);
  const [exportDone, setExportDone] = useState(false);
  const [exportXlsxDone, setExportXlsxDone] = useState(false);
  const [projectName, setProjectName] = useState(() => { try { return localStorage.getItem("projectName") || ""; } catch { return ""; } });
  const [copiedIdx, setCopiedIdx] = useState(null);
  const [crmOpen, setCrmOpen] = useState(false);
  const [crmKey, setCrmKey] = useState("");
  const [crmPushing, setCrmPushing] = useState(false);
  const [crmResults, setCrmResults] = useState(null);
  const [extractSnovKey, setExtractSnovKey] = useState(() => { try { return localStorage.getItem("extractSnovKey") || ""; } catch { return ""; } });
  const [extractSnovSecret, setExtractSnovSecret] = useState(() => { try { return localStorage.getItem("extractSnovSecret") || ""; } catch { return ""; } });
  const [extractSnovSaved, setExtractSnovSaved] = useState(() => { try { return !!localStorage.getItem("extractSnovKey"); } catch { return false; } });
  const [extractApolloKey, setExtractApolloKey] = useState(() => { try { return localStorage.getItem("extractApolloKey") || ""; } catch { return ""; } });
  const [extractApolloSaved, setExtractApolloSaved] = useState(() => { try { return !!localStorage.getItem("extractApolloKey"); } catch { return false; } });
  const [extractResults, setExtractResults] = useState([]);
  const [extractRunning, setExtractRunning] = useState(false);
  const [extractProgress, setExtractProgress] = useState({ done: 0, total: 0 });

  // search/filter
  const [leadsSearch, setLeadsSearch] = useState("");
  const [leadsTagFilter, setLeadsTagFilter] = useState("all");

  // verification
  const [verifying, setVerifying] = useState(false);
  const [verifyProgress, setVerifyProgress] = useState(0);

  // copy-all
  const [copiedAll, setCopiedAll] = useState(false);

  // maps state
  const [placesKey, setPlacesKey] = useState(() => { try { return localStorage.getItem("placesKey") || ""; } catch { return ""; } });
  const [placesKeySaved, setPlacesKeySaved] = useState(() => { try { return !!localStorage.getItem("placesKey"); } catch { return false; } });
  const [mapsQuery, setMapsQuery] = useState("");
  const [mapsLocation, setMapsLocation] = useState("");
  const [mapsResults, setMapsResults] = useState([]);
  const [mapsLoading, setMapsLoading] = useState(false);
  const [mapsError, setMapsError] = useState("");
  const [mapsProgress, setMapsProgress] = useState("");
  const [selectedPlaces, setSelectedPlaces] = useState(new Set());
  const [mapsMaxPages, setMapsMaxPages] = useState(3);
  const [multiQueryMode, setMultiQueryMode] = useState(false);
  const [multiQueryProgress, setMultiQueryProgress] = useState("");
  const [customAreas, setCustomAreas] = useState("");
  const mapsStopRef = useRef(false);
  const [mapsScrapeMethod, setMapsScrapeMethod] = useState("bulk"); // "bulk" | "smart" | "extract"
  const [mapsSmartProvider, setMapsSmartProvider] = useState("auto"); // provider for smart scraper

  // ── Keyword Expansion state ──────────────────────────────────────────────
  const [keywordExpand, setKeywordExpand]   = useState(false);   // master expansion toggle
  const [kwUseSuggest, setKwUseSuggest]     = useState(true);    // Google Suggest sub-toggle
  const [kwUseAI, setKwUseAI]               = useState(false);   // Groq AI sub-toggle
  const [kwMaxExpanded, setKwMaxExpanded]   = useState(8);       // cap on # of merged keywords fed into pipeline
  const [kwMaxSearches, setKwMaxSearches]   = useState(30);      // hard cap on total scrapeWithPlaces calls per run
  // preview state (populated when expansion panel is open)
  const [kwPreview, setKwPreview] = useState({ static: [], suggest: [], ai: [], merged: [], loading: false, error: "" });

  // bulk scraper state
  const [bulkUrls, setBulkUrls] = useState("");
  const [bulkFile, setBulkFile] = useState(null);
  const [bulkResults, setBulkResults] = useState([]);
  const [bulkSummaryOpen, setBulkSummaryOpen] = useState(false); // Feature 5: end-of-run summary
  const [screenshotModal, setScreenshotModal] = useState(null); // base64 string or null
  const [bulkRunning, setBulkRunning] = useState(false);
  const [bulkJobId,   setBulkJobId]   = useState(null);  // server-side job ID when using Worker job engine
  const [bulkProgress, setBulkProgress] = useState({ done: 0, total: 0 });
  const [bulkBatch,    setBulkBatch]    = useState({ current: 0, total: 0 }); // batch tracking
  const [bulkStopped, setBulkStopped] = useState(false);
  const [bulkInputMode, setBulkInputMode] = useState("text");
  const [bulkStartTime, setBulkStartTime] = useState(null);
  const [bulkEta, setBulkEta] = useState(null);
  const [bulkConcurrency, setBulkConcurrency] = useState(1); // default 1 to avoid browser 429 rate limits
  const bulkStopRef = useRef(false);
  const bulkRetryMap = useRef({});
  const MAX_BULK_RETRIES = 3;
  // ── Feature 3: Adaptive concurrency ─────────────────────────────────────
  // Tracks effective concurrency dynamically — starts at bulkConcurrency,
  // drops to 1 on 429/pool-exhausted signals, ramps back after 60s.
  const adaptiveConcRef = useRef(1); // runtime effective concurrency (not state to avoid re-renders)
  const adaptiveRampTimer = useRef(null); // timer to ramp concurrency back up
  const BULK_BATCH_SIZE  = 50; // URLs per batch — keeps browser pool from being overloaded
  const bulkFileRef = useRef();
  const [deepCrawl, setDeepCrawl] = useState(false);
  const [bulkAutoRun, setBulkAutoRun] = useState(false);
  const [geminiAutoRun, setGeminiAutoRun] = useState(false);
  const [extractAutoRun, setExtractAutoRun] = useState(false);
  const [crawlDepth, setCrawlDepth] = useState(20);

  // scraping API keys state
  const [scraperApiKey, setScraperApiKey] = useState(() => { try { return localStorage.getItem("scraperApiKey") || ""; } catch { return ""; } });
  const [scraperApiKeySaved, setScraperApiKeySaved] = useState(() => { try { return !!localStorage.getItem("scraperApiKey"); } catch { return false; } });
  const [zenRowsKey, setZenRowsKey] = useState(() => { try { return localStorage.getItem("zenRowsKey") || ""; } catch { return ""; } });
  const [zenRowsKeySaved, setZenRowsKeySaved] = useState(() => { try { return !!localStorage.getItem("zenRowsKey"); } catch { return false; } });
  const [groqKey, setGroqKey] = useState(() => { try { return localStorage.getItem("groqKey") || ""; } catch { return ""; } });
  const [groqKeySaved, setGroqKeySaved] = useState(() => { try { return !!localStorage.getItem("groqKey"); } catch { return false; } });
  const [hunterKey, setHunterKey] = useState(() => { try { return localStorage.getItem("hunterKey") || ""; } catch { return ""; } });
  const [hunterKeySaved, setHunterKeySaved] = useState(() => { try { return !!localStorage.getItem("hunterKey"); } catch { return false; } });
  const [snovSecret, setSnovSecret] = useState(() => { try { return localStorage.getItem("snovSecret") || ""; } catch { return ""; } });
  const [snovSecretSaved, setSnovSecretSaved] = useState(() => { try { return !!localStorage.getItem("snovSecret"); } catch { return false; } });
  const [apolloKey, setApolloKey] = useState(() => { try { return localStorage.getItem("apolloKey") || ""; } catch { return ""; } });
  const [apolloKeySaved, setApolloKeySaved] = useState(() => { try { return !!localStorage.getItem("apolloKey"); } catch { return false; } });
  const [hunterEnabled, setHunterEnabled] = useState(() => { try { return localStorage.getItem("hunterEnabled") === "true"; } catch { return false; } });
  const [apolloEnabled, setApolloEnabled] = useState(() => { try { return localStorage.getItem("apolloEnabled") === "true"; } catch { return false; } });

  // gemini scraper state
  const [geminiKey, setGeminiKey] = useState(() => { try { return localStorage.getItem("geminiKey") || ""; } catch { return ""; } });
  const [geminiKeySaved, setGeminiKeySaved] = useState(() => { try { return !!localStorage.getItem("geminiKey"); } catch { return false; } });

  // ── SecureStorage recovery — runs once on mount ───────────────────────────
  // If localStorage was wiped (APK update / WebView reset), this pulls all
  // protected keys back from Android SharedPreferences and restores state.
  useEffect(() => {
    const recover = async () => {
      const load = (key) => SecureStorage.get(key); // async, tries Preferences first

      const [
        rGroqKey, rPlacesKey, rGeminiKey,
        rScraperApiKey, rZenRowsKey, rWorkerUrl,
        rExtractSnovKey, rExtractSnovSecret, rExtractApolloKey,
        rHunterKey, rSnovSecret, rApolloKey,
        rHunterEnabled, rApolloEnabled,
      ] = await Promise.all([
        load("groqKey"), load("placesKey"), load("geminiKey"),
        load("scraperApiKey"), load("zenRowsKey"), load("workerUrl"),
        load("extractSnovKey"), load("extractSnovSecret"), load("extractApolloKey"),
        load("hunterKey"), load("snovSecret"), load("apolloKey"),
        load("hunterEnabled"), load("apolloEnabled"),
      ]);

      // Only update state if Preferences had a value that localStorage didn't
      // (avoids unnecessary re-renders on normal loads)
      if (rGroqKey        && !localStorage.getItem("groqKey"))        { setGroqKey(rGroqKey);               setGroqKeySaved(true); }
      if (rPlacesKey      && !localStorage.getItem("placesKey"))      { setPlacesKey(rPlacesKey);           setPlacesKeySaved(true); }
      if (rGeminiKey      && !localStorage.getItem("geminiKey"))      { setGeminiKey(rGeminiKey);           setGeminiKeySaved(true); }
      if (rScraperApiKey  && !localStorage.getItem("scraperApiKey"))  { setScraperApiKey(rScraperApiKey);   setScraperApiKeySaved(true); }
      if (rZenRowsKey     && !localStorage.getItem("zenRowsKey"))     { setZenRowsKey(rZenRowsKey);         setZenRowsKeySaved(true); }
      if (rWorkerUrl      && !localStorage.getItem("workerUrl"))      { setWorkerUrl(rWorkerUrl);           setWorkerUrlSaved(true); }
      if (rExtractSnovKey && !localStorage.getItem("extractSnovKey")) { setExtractSnovKey(rExtractSnovKey); setExtractSnovSaved(true); }
      if (rExtractSnovSecret && !localStorage.getItem("extractSnovSecret")) setExtractSnovSecret(rExtractSnovSecret);
      if (rExtractApolloKey && !localStorage.getItem("extractApolloKey")) { setExtractApolloKey(rExtractApolloKey); setExtractApolloSaved(true); }
      if (rHunterKey      && !localStorage.getItem("hunterKey"))      { setHunterKey(rHunterKey);           setHunterKeySaved(true); }
      if (rSnovSecret     && !localStorage.getItem("snovSecret"))     { setSnovSecret(rSnovSecret);         setSnovSecretSaved(true); }
      if (rApolloKey      && !localStorage.getItem("apolloKey"))      { setApolloKey(rApolloKey);           setApolloKeySaved(true); }
      if (rHunterEnabled  !== "" && !localStorage.getItem("hunterEnabled"))  setHunterEnabled(rHunterEnabled === "true");
      if (rApolloEnabled  !== "" && !localStorage.getItem("apolloEnabled"))  setApolloEnabled(rApolloEnabled === "true");

      console.log("[SecureStorage] Recovery check complete");
    };
    recover().catch(e => console.warn("[SecureStorage] Recovery failed:", e.message));
  }, []); // runs once on mount only
  const [geminiUrls, setGeminiUrls] = useState("");
  const [geminiFile, setGeminiFile] = useState(null);
  const [geminiInputMode, setGeminiInputMode] = useState("text");
  const [geminiResults, setGeminiResults] = useState([]);
  const [geminiRunning, setGeminiRunning] = useState(false);
  const [geminiProgress, setGeminiProgress] = useState({ done: 0, total: 0 });
  const [geminiStopped, setGeminiStopped] = useState(false);
  const [geminiStartTime, setGeminiStartTime] = useState(null);
  const [geminiEta, setGeminiEta] = useState(null);
  const geminiStopRef = useRef(false);
  const apiCacheRef = useRef({ hunter: {}, apollo: {} });
  const geminiFileRef = useRef();
  const geminiTextareaRef = useRef();
  const [geminiProvider, setGeminiProvider] = useState("auto");
  const [workerUrl, setWorkerUrl] = useState(() => { try { return localStorage.getItem("workerUrl") ?? "https://scraper-proxy.scraper-proxy-01.workers.dev/"; } catch { return "https://scraper-proxy.scraper-proxy-01.workers.dev/"; } });
  const [workerUrlSaved, setWorkerUrlSaved] = useState(() => { try { return !!localStorage.getItem("workerUrl") || true; } catch { return true; } });

  // ── API Credit Monitor state ─────────────────────────────────────────────
  const [apiCredits, setApiCredits] = useState(() => {
    const usage = loadCreditUsage();
    return { scraperapi: null, zenrows: null, ...usage };
  });
  const [creditPanelOpen, setCreditPanelOpen] = useState(false);

  // Fetch live credits (ScraperAPI + ZenRows) — non-blocking, background only
  const refreshCredits = useCallback(async () => {
    const usage = loadCreditUsage();
    const updates = { ...usage };

    const [sa, zr] = await Promise.allSettled([
      scraperApiKey.trim() ? fetchScraperApiCredits(scraperApiKey) : Promise.resolve(null),
      zenRowsKey.trim()    ? fetchZenRowsCredits(zenRowsKey)        : Promise.resolve(null),
    ]);

    updates.scraperapi = sa.status === "fulfilled" ? sa.value : "unavailable";
    updates.zenrows    = zr.status === "fulfilled" ? zr.value : "unavailable";
    setApiCredits(updates);
  }, [scraperApiKey, zenRowsKey]);

  // On mount + every 60 seconds — completely non-blocking
  useEffect(() => {
    refreshCredits();
    const timer = setInterval(refreshCredits, 60_000);
    return () => clearInterval(timer);
  }, [refreshCredits]);

  // Re-read local usage counters whenever panel opens
  useEffect(() => {
    if (creditPanelOpen) {
      const usage = loadCreditUsage();
      setApiCredits(prev => ({ ...prev, ...usage }));
    }
  }, [creditPanelOpen]);

  const fileRef = useRef();
  const editRef = useRef();

  // ── localStorage persistence ──────────────────────────────────────────────
  useEffect(() => {
    try { localStorage.setItem("leads", JSON.stringify(leads)); } catch {}
  }, [leads]);

  // ── sync workerUrl state → module-level ref so fetch helpers always see latest ──
  useEffect(() => {
    workerUrlRef.current = workerUrl;
  }, [workerUrl]);

  // ── App lifecycle debug listener — logs background transitions ──────────────
  // Helps verify that scraping continues when the app is minimized.
  // Check Android Studio Logcat (tag: LeadExtractor) or browser console.
  useEffect(() => {
    if (!window.Capacitor?.isNativePlatform?.()) return;
    let cleanup = () => {};
    import(/* @vite-ignore */ '@capacitor/app').then(({ App }) => {
      const handle = App.addListener('appStateChange', ({ isActive }) => {
        if (isActive) {
          console.log('[AppLifecycle] App foregrounded');
        } else {
          const activeScrape = bulkRunning ? 'bulk' : geminiRunning ? 'smart' : extractRunning ? 'extract' : 'none';
          console.log(`[AppLifecycle] App backgrounded — active scrape: ${activeScrape}`);
          if (activeScrape !== 'none') {
            console.log('[AppLifecycle] Foreground service is live — WebView suppressed from pausing');
          }
        }
      });
      cleanup = () => handle.then(h => h.remove()).catch(() => {});
    }).catch(() => {});
    return () => cleanup();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // mount only — refs used inside are stable

  // ── sync API key state → module-level refs so fetchHtmlViaWorker can pass them ──
  useEffect(() => { scraperApiKeyRef.current = scraperApiKey; }, [scraperApiKey]);
  useEffect(() => { zenRowsKeyRef.current    = zenRowsKey;    }, [zenRowsKey]);

  // ── auto-run bulk scraper when triggered from Maps tab ───────────────────
  useEffect(() => {
    if (bulkAutoRun && tab === "bulk" && !bulkRunning) {
      setBulkAutoRun(false);
      runBulkScrape(false);
    }
  }, [bulkAutoRun, tab, bulkRunning]);

  // ── auto-run smart scraper when triggered from Maps tab ──────────────────
  useEffect(() => {
    if (geminiAutoRun && tab === "gemini" && !geminiRunning) {
      setGeminiAutoRun(false);
      runGeminiScrape();
    }
  }, [geminiAutoRun, tab, geminiRunning]);

  // ── auto-run extract leads when triggered from Maps tab ──────────────────
  useEffect(() => {
    if (extractAutoRun && tab === "extract" && !extractRunning) {
      setExtractAutoRun(false);
      runExtractLeads();
    }
  }, [extractAutoRun, tab, extractRunning]);

  // ── lead helpers ──────────────────────────────────────────────────────────
  const addLeads = (newLeads) =>
    setLeads((prev) => markDuplicates([...prev, ...newLeads.map((l) => ({ ...l, id: uid() }))]));

  const deduplicateAll = () => {
    setLeads((prev) => {
      const seen = new Set();
      return markDuplicates(prev.filter((l) => {
        const key = normalizeEmail(l.email) || l.name?.trim().toLowerCase();
        if (!key || seen.has(key)) return false;
        seen.add(key); return true;
      }));
    });
    setShowDupOnly(false);
  };

  // ── Extract Leads tab — Snov.io + Apollo.io engine ──────────────────────
  const MAX_EXTRACT_SNOV = 50;
  const MAX_EXTRACT_APOLLO = 50;

  const getSnovToken = async (clientId, clientSecret) => {
    const workerUrl = workerUrlRef.current ?? "";
    const endpoint = "https://api.snov.io/v1/oauth/access_token";
    const body = JSON.stringify({ grant_type: "client_credentials", client_id: clientId.trim(), client_secret: clientSecret.trim() });
    let res;
    if (workerUrl.trim()) {
      const proxied = `${workerUrl.replace(/\/$/, "")}/fetch?url=${encodeURIComponent(endpoint)}`;
      res = await fetch(proxied, { method: "POST", headers: { "Content-Type": "application/json" }, body, signal: AbortSignal.timeout(15000) });
    } else {
      res = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json" }, body, signal: AbortSignal.timeout(10000) });
    }
    const data = await res.json();
    return data?.access_token ?? null;
  };

  const fetchExtractSnovEmails = async (domain, token) => {
    const workerUrl = workerUrlRef.current ?? "";
    const endpoint = `https://api.snov.io/v2/domain-emails-with-info?domain=${encodeURIComponent(domain)}&type=all&limit=10&lastId=0&access_token=${encodeURIComponent(token)}`;
    let res;
    if (workerUrl.trim()) {
      const proxied = `${workerUrl.replace(/\/$/, "")}/fetch?url=${encodeURIComponent(endpoint)}`;
      res = await fetch(proxied, { signal: AbortSignal.timeout(15000) });
    } else {
      res = await fetch(endpoint, { signal: AbortSignal.timeout(10000) });
    }
    const data = await res.json();
    return (data?.emails || data?.data || []).map(e => ({
      email: (e.email || e.value)?.toLowerCase().trim() ?? null,
      name: [e.first_name, e.last_name].filter(Boolean).join(" ") || null,
      title: e.position ?? e.title ?? null,
    })).filter(c => c.email?.includes("@"));
  };

  const fetchExtractApollo = async (domain, apiKey) => {
    const workerUrl = workerUrlRef.current ?? "";
    const endpoint = `https://api.apollo.io/v1/organizations/enrich?api_key=${encodeURIComponent(apiKey.trim())}`;
    const body = JSON.stringify({ domain });
    let res;
    if (workerUrl.trim()) {
      const proxied = `${workerUrl.replace(/\/$/, "")}/fetch?url=${encodeURIComponent(endpoint)}`;
      res = await fetch(proxied, { method: "POST", headers: { "Content-Type": "application/json" }, body, signal: AbortSignal.timeout(15000) });
    } else {
      res = await fetch(endpoint, { method: "POST", headers: { "Content-Type": "application/json", "X-Api-Key": apiKey.trim() }, body, signal: AbortSignal.timeout(10000) });
    }
    const data = await res.json();
    const org = data?.organization;
    if (!org) return null;
    return { company: org.name ?? null, industry: org.industry ?? null, size: org.estimated_num_employees ?? null, city: org.city ?? null, country: org.country ?? null };
  };

  const runExtractLeads = async () => {
    // Normalise: deduplicate by domain so Snov is never called twice for the same domain
    const rawUrls = urlInput.trim().split(/\n+/).map(u => u.trim()).filter(Boolean);
    if (!rawUrls.length) return;

    // Build domain-keyed unique list, preserving first occurrence URL
    const domainMap = new Map(); // domain → original url string
    for (const url of rawUrls) {
      const domain = getDomainFromUrl(url.startsWith("http") ? url : `https://${url}`) ?? url;
      if (!domainMap.has(domain)) domainMap.set(domain, url);
    }
    const entries = Array.from(domainMap.entries()); // [[domain, url], …]

    setExtractRunning(true);
    setExtractResults([]);
    setExtractProgress({ done: 0, total: entries.length });
    ScrapeNative.start(entries.length, "extract"); // keep app alive when minimized

    let snovUsedCount = 0;
    let apolloUsedCount = 0;
    const domainCache = {}; // snov_<domain> and apollo_<domain>

    // ── Obtain Snov token once upfront ───────────────────────────────────────
    let snovToken = null;
    if (extractSnovKey.trim() && extractSnovSecret.trim()) {
      try {
        snovToken = await getSnovToken(extractSnovKey, extractSnovSecret);
        if (!snovToken) {
          setExtractResults([{ url: "", domain: "snov-auth", emails: [], contacts: [], provider: "snov", enrichment: null, status: "error", error: "Snov auth failed — check Client ID & Secret" }]);
          setExtractRunning(false);
          return;
        }
        console.log("[ExtractLeads] Snov token obtained");
      } catch (e) {
        setExtractResults([{ url: "", domain: "snov-auth", emails: [], contacts: [], provider: "snov", enrichment: null, status: "error", error: `Snov token error: ${e.message}` }]);
        setExtractRunning(false);
        return;
      }
    }

    for (let i = 0; i < entries.length; i++) {
      const [domain, url] = entries[i];

      // Step 1 — log
      console.log("Processing domain:", domain);

      setExtractResults(prev => [...prev, {
        url, domain, emails: [], contacts: [],
        provider: "snov", enrichment: null, status: "processing",
      }]);

      try {
        let contacts = [];
        let emails = [];

        // Step 2 — call Snov immediately (primary, unconditional)
        if (snovToken) {
          if (snovUsedCount >= MAX_EXTRACT_SNOV) {
            console.log("Skipping API (limit reached):", domain);
          } else {
            const cacheKey = `snov_${domain}`;
            if (domainCache[cacheKey] !== undefined) {
              contacts = domainCache[cacheKey];
            } else {
              snovUsedCount++;
              try {
                contacts = await fetchExtractSnovEmails(domain, snovToken);
                domainCache[cacheKey] = contacts;
                if (contacts.length > 0) { incrementSnovUsage(); setApiCredits(prev => ({ ...prev, ...loadCreditUsage() })); }
              } catch (e) {
                console.warn("[ExtractLeads] Snov failed for", domain, ":", e.message);
                domainCache[cacheKey] = [];
              }
            }
          }
        }

        // Step 3 — store emails from Snov
        emails = contacts.map(c => c.email).filter(Boolean);

        // Step 4 — Apollo enrichment (only if emails exist AND key present AND within limit)
        let enrichment = null;
        if (emails.length > 0 && extractApolloKey.trim()) {
          if (apolloUsedCount >= MAX_EXTRACT_APOLLO) {
            console.log("Skipping API (limit reached):", domain);
          } else {
            const cacheKey = `apollo_${domain}`;
            if (domainCache[cacheKey] !== undefined) {
              enrichment = domainCache[cacheKey];
            } else {
              console.log("[ExtractLeads] Apollo enrichment for:", domain);
              apolloUsedCount++;
              try {
                enrichment = await fetchExtractApollo(domain, extractApolloKey);
                domainCache[cacheKey] = enrichment;
                if (enrichment) { incrementApolloUsage(); setApiCredits(prev => ({ ...prev, ...loadCreditUsage() })); }
              } catch (e) {
                console.warn("[ExtractLeads] Apollo failed for", domain, ":", e.message);
                domainCache[cacheKey] = null;
              }
            }
          }
        }

        // Push to leads table — store enrichment fields directly so XLSX/CSV export works
        const enrichmentPayload = enrichment ? {
          industry: enrichment.industry ?? null,
          size: enrichment.size ?? null,
          city: enrichment.city ?? null,
          country: enrichment.country ?? null,
        } : null;
        const leadsToAdd = contacts.length > 0
          ? contacts.map(c => ({
              name: c.name ?? null,
              email: c.email,
              company: enrichment?.company ?? null,
              role: c.title ?? null,
              enrichment: enrichmentPayload,
            }))
          : emails.map(e => ({ name: null, email: e, company: enrichment?.company ?? null, role: null, enrichment: enrichmentPayload }));
        if (leadsToAdd.length > 0) addLeads(leadsToAdd);

        // Result structure per spec
        setExtractResults(prev => prev.map(r => r.domain === domain ? {
          ...r,
          emails,
          contacts,
          provider: "snov",
          enrichment: enrichment ? {
            name: enrichment.name ?? null,
            role: enrichment.role ?? null,
            company: enrichment.company ?? null,
            // keep extra fields for display
            industry: enrichment.industry ?? null,
            size: enrichment.size ?? null,
            city: enrichment.city ?? null,
            country: enrichment.country ?? null,
          } : null,
          status: emails.length > 0 ? "success" : "no-email",
        } : r));

      } catch (e) {
        console.error("[ExtractLeads] Error for", domain, ":", e.message);
        setExtractResults(prev => prev.map(r => r.domain === domain
          ? { ...r, status: "error", error: e.message } : r));
      }

      setExtractProgress({ done: i + 1, total: entries.length });
      const emailsSoFarExt = extractResults.reduce((sum, r) => sum + (r.emails?.length ?? 0), 0);
      ScrapeNative.progress(i + 1, entries.length, emailsSoFarExt, "extract");
      await new Promise(r => setTimeout(r, 600));
    }

    setExtractRunning(false);
    ScrapeNative.complete(entries.length, emailsSoFarExt); // ✓ Done — auto-dismisses
  };

  const canRun = !extractRunning && urlInput.trim().length > 0;

  const prompt = (text) => `Extract ALL emails from this text. Return only a JSON array of email strings.

${text}`;

  const extract = async () => {
    if (mode === "url") { await runExtractLeads(); return; }
    setLoading(true); setError("");
    try {
      let result = [];
      const emailRegex = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
      if (mode === "text") {
        setLoadingMsg("Scanning text…");
        const emails = (textInput.match(emailRegex) || []).map(e => e.toLowerCase().trim());
        result = emails.map(e => ({ name: null, email: e, company: null, role: null }));
      } else if (mode === "file") {
        setLoadingMsg("Reading file…");
        const text = await readFileAsText(file);
        const emails = (text.match(emailRegex) || []).map(e => e.toLowerCase().trim());
        result = emails.map(e => ({ name: null, email: e, company: null, role: null }));
      }
      if (!result.length) throw new Error("No emails found in the provided content.");
      addLeads(result);
    } catch (e) { setError(e.message || "Extraction failed."); }
    finally { setLoading(false); setLoadingMsg(""); }
  };


  // ── cell editing ──────────────────────────────────────────────────────────
  const startEdit = (id, field, val) => { setEditingCell({ id, field }); setEditVal(val ?? ""); setTimeout(() => editRef.current?.focus(), 30); };
  const commitEdit = () => {
    if (!editingCell) return;
    setLeads((prev) => markDuplicates(prev.map((l) => l.id === editingCell.id ? { ...l, [editingCell.field]: editVal || null } : l)));
    setEditingCell(null);
  };

  // ── download helper (works in Android WebView) ────────────────────────────
  const triggerDownload = async (content, filename, mime) => {
    // Convert content to base64
    let base64;
    if (typeof content === "string") {
      base64 = btoa(unescape(encodeURIComponent(content)));
    } else {
      const bytes = new Uint8Array(content);
      let bin = "";
      for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
      base64 = btoa(bin);
    }

    // Try Capacitor Filesystem first (Android/iOS)
    try {
      const mod = await import('@capacitor/filesystem');
      const Filesystem = mod.Filesystem;
      const Directory = mod.Directory;
      await Filesystem.writeFile({
        path: filename,
        data: base64,
        directory: Directory.Documents,
        recursive: true,
      });
      alert("\u2705 Saved to Documents/" + filename);
      return;
    } catch (e) {
      console.warn('[Download] Capacitor Filesystem unavailable, using browser fallback:', e.message);
    }

    // Browser fallback
    const a = document.createElement("a");
    a.href = "data:" + mime + ";base64," + base64;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  // ── export CSV ────────────────────────────────────────────────────────────
  const exportCSV = (rows = displayLeads) => {
    const esc = (v) => `"${(v ?? "").toString().replace(/"/g, '""')}"`;
    const header = ["name","email","company","role","industry","employees","city","country"].join(",");
    const csv = [header, ...rows.map((l) => [
      esc(l.name),
      esc(l.email),
      esc(l.company ?? l.enrichment?.company),
      esc(l.role),
      esc(l.enrichment?.industry),
      esc(l.enrichment?.size),
      esc(l.enrichment?.city),
      esc(l.enrichment?.country),
    ].join(","))].join("\n");
    triggerDownload(csv, "leads.csv", "text/csv");
    setExportDone(true); setTimeout(() => setExportDone(false), 2000);
  };

  // ── export XLSX ───────────────────────────────────────────────────────────
  // ── Smart filename generator ─────────────────────────────────────────────
  const buildFileName = (prefix, query = "", location = "") => {
    const parts = [prefix];
    const clean = (s) => s.trim().replace(/[^a-zA-Z0-9 _-]/g, "").replace(/\s+/g, "_").slice(0, 40);
    if (query) parts.push(clean(query));
    if (location) parts.push(clean(location));
    const date = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
    parts.push(date);
    return parts.filter(Boolean).join("_") + ".xlsx";
  };

  // ── Project-based filename for leads XLSX ─────────────────────────────────
  const buildLeadsFileName = () => {
    const date = new Date().toISOString().slice(0, 10);
    return `leads-export-${date}.xlsx`;
  };

  const promptFileName = (defaultName) => {
    const bareDefault = defaultName.replace(/\.xlsx$/i, "");
    const raw = window.prompt("Enter a filename for your download:", bareDefault);
    if (raw === null) return null;
    const sanitized = (raw.trim() || bareDefault)
      .toLowerCase()
      .replace(/\s+/g, "-")
      .replace(/[^a-z0-9._-]/g, "");
    return sanitized + ".xlsx";
  };

  const exportXLSX = (rows = displayLeads) => {
    const data = rows.map(l => ({
      Name: l.name ?? "",
      Email: l.email ?? "",
      Company: l.company ?? l.enrichment?.company ?? "",
      Role: l.role ?? "",
      Industry: l.enrichment?.industry ?? "",
      Employees: l.enrichment?.size ?? "",
      City: l.enrichment?.city ?? "",
      Country: l.enrichment?.country ?? "",
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 24 }, { wch: 32 }, { wch: 24 }, { wch: 22 }, { wch: 20 }, { wch: 12 }, { wch: 18 }, { wch: 16 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Leads");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const filename = promptFileName(buildLeadsFileName());
    if (!filename) return;
    triggerDownload(wbout, filename, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    setExportXlsxDone(true); setTimeout(() => setExportXlsxDone(false), 2000);
  };

  // ── HubSpot ───────────────────────────────────────────────────────────────
  const pushCRM = async () => {
    if (!crmKey.trim()) return;
    setCrmPushing(true); setCrmResults(null);
    try { setCrmResults(await pushToHubSpot(crmKey.trim(), displayLeads.filter((l) => !l._dup))); }
    catch (e) { setCrmResults([{ ok: false, error: e.message }]); }
    finally { setCrmPushing(false); }
  };

  // ── Gemini scraper ───────────────────────────────────────────────────────
  const geminiTotalEmails = geminiResults.reduce((s, r) => s + r.emails.length, 0);
  const geminiOk = geminiResults.filter(r => r.status === "ok").length;
  const geminiFailed = geminiResults.filter(r => r.status === "error").length;

  // MUST be above parseGeminiUrls — const is not hoisted
  const normalizeUrl = (u) => {
    u = u.trim();
    if (!u) return null;
    if (!u.startsWith("http://") && !u.startsWith("https://")) u = "https://" + u;
    try { return "https://" + new URL(u).hostname; }
    catch { return null; }
  };

  const parseGeminiUrls = (text) =>
    text.split(/[\n,]+/).map(u => normalizeUrl(u)).filter(Boolean);

  const runGeminiScrape = async () => {
    // ── Validate key for selected provider ────────────────────────────────────
    const keyMap = {
      gemini:     { key: geminiKey,     label: "Gemini API key required" },
      groq:       { key: groqKey,       label: "Groq API key required" },
      scraperapi: { key: scraperApiKey, label: "ScraperAPI key required" },
      zenrows:    { key: zenRowsKey,    label: "ZenRows API key required" },
      auto:       { key: true,          label: "" },
    };
    const { key, label } = keyMap[geminiProvider] ?? {};
    if (geminiProvider !== "auto" && !key?.trim?.()) {
      setGeminiResults([{ url: "", emails: [], status: "error", error: label, provider: "none" }]);
      return;
    }

    console.group("[SmartScraper] runGeminiScrape() called — provider:", geminiProvider);
    console.log("geminiInputMode:", geminiInputMode);
    console.log("geminiFile:", geminiFile?.name ?? "null");

    let urls = [];
    if (geminiInputMode === "file" && geminiFile) {
      const fileName = geminiFile.name.toLowerCase();
      if (fileName.endsWith(".xlsx")) {
        const arrayBuffer = await geminiFile.arrayBuffer();
        const wb = XLSX.read(arrayBuffer, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        urls = rows.flatMap(row => row.map(cell => normalizeUrl(String(cell ?? "")))).filter(Boolean);
      } else {
        const text = await readFileAsText(geminiFile);
        urls = text.split(/[\n,\r]+/).map(u => normalizeUrl(u)).filter(Boolean);
      }
    } else {
      // Fallback: read directly from DOM in case state update was delayed (Android WebView paste)
      const urlText = geminiTextareaRef.current?.value || geminiUrls;
      urls = parseGeminiUrls(urlText);
    }
    urls = [...new Set(urls)];
    console.log("Final URL list (%d):", urls.length, urls.slice(0, 10));
    if (!urls.length) {
      console.warn("[SmartScraper] No valid URLs — returning early.");
      console.groupEnd();
      return;
    }
    console.log("[SmartScraper] Starting scrape:", urls.length, "URLs");
    console.groupEnd();

    const keys = { geminiKey, groqKey, scraperApiKey, zenRowsKey };
    const startTime = Date.now();
    setGeminiStartTime(startTime);
    setGeminiEta(null);
    setGeminiRunning(true);
    setGeminiStopped(false);
    geminiStopRef.current = false;
    setGeminiResults([]);
    setGeminiProgress({ done: 0, total: urls.length });
    ScrapeNative.start(urls.length, "smart"); // keep app alive when minimized

    // Reset per-run API caches and counters
    apiCacheRef.current = { hunter: {}, apollo: {} };
    let hunterUsedCount = 0;
    let apolloUsedCount = 0;
    const MAX_HUNTER = 5;
    const MAX_APOLLO = 5;

    for (let i = 0; i < urls.length; i++) {
      if (geminiStopRef.current) { setGeminiStopped(true); break; }
      const url = stripTrackingParams(urls[i]);
      setGeminiProgress({ done: i, total: urls.length });
      setGeminiResults(prev => [...prev, { url, emails: [], status: "loading", error: null, provider: null, enrichment: null }]);
      // Scrape result — may fail, Hunter still fires after
      let emails = [];
      let provider = "unknown";
      let enrichment = null;
      let scrapeError = null;
      const skipped = isSkippedDomain(url);

      try {
        const result = await extractEmailsSmart(geminiProvider, keys, url);
        emails = result.emails;
        provider = result.provider;
      } catch (e) {
        scrapeError = e.message;
        console.warn("[SmartScraper] Scrape failed for", url, "—", e.message, "— will still try Snov fallback");
      }

      try {
        // Apollo org enrichment (company name, industry, size, location)
        if (!skipped && apolloEnabled && apolloKey.trim() && apolloUsedCount < MAX_APOLLO) {
          apolloUsedCount++;
          enrichment = await fetchApolloEnrichment(url, apolloKey, apiCacheRef.current, workerUrlRef.current);
          // Note: Apollo people search requires paid plan — decision makers come from Snov.io
        }

        // Hunter fallback — fires even when scrape errored out
        console.log("[Snov] Conditions check:", {
          emailsLength: emails.length,
          hunterEnabled,
          hunterUsedCount,
          MAX_HUNTER,
          hasKey: !!hunterKey.trim(),
          skipped,
          scrapeErrored: !!scrapeError,
        });
        if (!skipped && emails.length === 0 && hunterEnabled && hunterKey.trim() && hunterUsedCount < MAX_HUNTER) {
          console.log("[Snov] Fallback triggered for:", url);
          hunterUsedCount++;
          const hunterEmails = await fetchHunterEmails(url, hunterKey, apiCacheRef.current, workerUrlRef.current, snovSecret);
          if (hunterEmails.length > 0) {
            emails = [...new Set([...emails, ...hunterEmails])];
            provider = "hunter";
            scrapeError = null;
            console.log("[Snov] Merged emails:", emails);

            // Pull Snov contact + decision maker data from cache
            const snovDomain = getDomainFromUrl(url);
            const snovContacts = apiCacheRef.current.hunter[`${snovDomain}_contacts`] || [];
            const snovDMs = apiCacheRef.current.hunter[`${snovDomain}_dm`] || [];

            // Apollo org enrichment + merge Snov people data
            if (apolloEnabled && apolloKey.trim() && apolloUsedCount < MAX_APOLLO) {
              apolloUsedCount++;
              enrichment = await fetchApolloEnrichment(url, apolloKey, apiCacheRef.current, workerUrlRef.current);
            }

            // Merge Snov decision makers into enrichment
            if (snovDMs.length > 0) {
              const dmForDisplay = snovDMs.map(c => ({
                name: c.name,
                title: c.title,
                email: c.email,
              }));
              enrichment = enrichment
                ? { ...enrichment, decisionMakers: dmForDisplay }
                : { decisionMakers: dmForDisplay };
              console.log("[Snov] Decision makers added to enrichment:", dmForDisplay);
            } else if (snovContacts.length > 0 && !enrichment?.decisionMakers) {
              // No clear DMs — show top 3 contacts with titles
              const topContacts = snovContacts.slice(0, 3).map(c => ({
                name: c.name,
                title: c.title,
                email: c.email,
              }));
              enrichment = enrichment
                ? { ...enrichment, decisionMakers: topContacts }
                : { decisionMakers: topContacts };
            }
          } else {
            console.log("[Snov] No emails found via Snov for:", url);
          }
        }

        setGeminiResults(prev => prev.map(r => r.url === url
          ? {
              url, emails,
              status: emails.length > 0 ? "ok" : scrapeError ? "error" : "none",
              error: emails.length > 0 ? null : scrapeError,
              provider: provider ?? "unknown",
              enrichment,
            }
          : r));
      } catch (e) {
        setGeminiResults(prev => prev.map(r => r.url === url
          ? { url, emails: [], status: "error", error: e.message, provider: "unknown", enrichment: null }
          : r));
      }
      const done = i + 1;
      setGeminiProgress({ done, total: urls.length });
      setGeminiEta(calcEta(startTime, done, urls.length));
      const emailsSoFarGem = geminiResults.reduce((sum, r) => sum + (r.emails?.length ?? 0), 0);
      ScrapeNative.progress(done, urls.length, emailsSoFarGem, "smart");
      await new Promise(r => setTimeout(r, 1500));
    }
    setGeminiProgress(p => ({ ...p, done: urls.length }));
    setGeminiEta(null);
    setGeminiRunning(false);
    ScrapeNative.complete(urls.length, emailsSoFarGem); // ✓ Done — auto-dismisses
  };

  const importGeminiToLeads = () => {
    const newLeads = [];
    geminiResults.forEach(r => {
      r.emails.forEach(email => {
        const domain = (() => { try { return new URL(r.url).hostname.replace(/^www\./, ""); } catch { return ""; } })();
        newLeads.push({ name: null, email, company: domain || null, role: null });
      });
    });
    addLeads(newLeads);
    setTab("extract");
  };

  const exportGeminiCSV = () => {
    const rows = [["url", "email"].join(",")];
    geminiResults.forEach(r => r.emails.forEach(e => rows.push(`"${r.url}","${e}"`)));
    triggerDownload(rows.join("\n"), "gemini-emails.csv", "text/csv");
  };

  const exportGeminiXLSX = () => {
    const data = [];
    geminiResults.forEach(r => {
      const domain = (() => { try { return new URL(r.url).hostname.replace(/^www\./, ""); } catch { return r.url; } })();
      r.emails.forEach(e => data.push({
        URL: r.url,
        Email: e,
        Company: r.enrichment?.company ?? r.enrichment?.name ?? domain,
        Industry: r.enrichment?.industry ?? "",
        Employees: r.enrichment?.size ?? "",
        City: r.enrichment?.city ?? "",
        Country: r.enrichment?.country ?? "",
        Provider: r.provider ?? "",
      }));
    });
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 40 }, { wch: 32 }, { wch: 24 }, { wch: 22 }, { wch: 12 }, { wch: 18 }, { wch: 16 }, { wch: 14 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Gemini Emails");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const geminiDomain = geminiUrls.trim().split(/[\n,]+/)[0]?.trim() || "";
    const geminiLabel = geminiDomain ? (() => { try { return new URL(geminiDomain.startsWith("http") ? geminiDomain : "https://" + geminiDomain).hostname.replace(/^www\./, "").split(".")[0]; } catch { return "scrape"; } })() : "scrape";
    triggerDownload(wbout, (() => { const filename = promptFileName(buildFileName("smart-scraper", geminiLabel)); return filename || buildFileName("smart-scraper", geminiLabel); })(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  };

  // ── Maps scraper ──────────────────────────────────────────────────────────
  const runMapsScrape = async () => {
    if (!mapsQuery.trim() || !placesKey.trim()) return;
    setMapsLoading(true); setMapsError(""); setMapsResults([]); setSelectedPlaces(new Set());
    mapsStopRef.current = false;
    ScrapeNative.start(mapsMaxPages, "maps"); // keep app alive when minimized
    try {
      let pagesDone = 0;
      // Stream results page-by-page into state as they arrive
      const onPageResult = (newPlaces) => {
        pagesDone++;
        setMapsResults(prev => {
          const updated = [...prev, ...newPlaces];
          setSelectedPlaces(new Set(updated.map(r => r.id)));
          ScrapeNative.progress(pagesDone, mapsMaxPages, updated.length, "maps");
          return updated;
        });
      };
      await scrapeWithPlaces(
        placesKey.trim(), mapsQuery.trim(), mapsLocation.trim(),
        setMapsProgress, mapsMaxPages, onPageResult, mapsStopRef
      );
      setMapsProgress("");
      incrementMapsUsage(); setApiCredits(prev => ({ ...prev, ...loadCreditUsage() }));
    } catch (e) {
      setMapsError(e.message || "Scrape failed.");
      setMapsProgress("");
    } finally {
      setMapsLoading(false);
      ScrapeNative.stop(); // error path — just stop
    }
  };

  const stopMapsScrape = () => { mapsStopRef.current = true; ScrapeNative.stop(); };

  // ── Keyword Expansion Scrape ───────────────────────────────────────────────
  // Sits BEFORE the existing multi-query pipeline.
  // Pipeline:  seed → static → Google Suggest → Groq AI → merge+dedupe → existing system
  // This function does NOT touch generateSubQueries(), scrapeWithPlaces(), or runMultiQueryScrape().
  const runKeywordExpansionScrape = async () => {
    if (!mapsQuery.trim() || !placesKey.trim()) return;

    setMapsLoading(true);
    setMapsError("");
    setMapsResults([]);
    setSelectedPlaces(new Set());
    mapsStopRef.current = false;

    // ── Step 1: static expansion ─────────────────────────────────────────
    const staticKws = expandKeywordsStatic(mapsQuery.trim());
    console.log("Static keywords:", staticKws);

    // ── Step 2: Google Suggest expansion ─────────────────────────────────
    let suggestKws = [];
    if (kwUseSuggest) {
      try {
        setMapsProgress("Fetching Google Suggest keywords…");
        const existingSet = new Set(staticKws.map(k => k.toLowerCase()));
        suggestKws = await fetchGoogleSuggest(mapsQuery.trim(), 8, existingSet);
        console.log("Google Suggest keywords:", suggestKws);
      } catch (e) {
        console.warn("[KwExpand/Suggest] Failed:", e.message);
      }
    }

    // ── Step 3: Groq AI expansion ─────────────────────────────────────────
    let aiKws = [];
    if (kwUseAI && groqKey.trim()) {
      try {
        setMapsProgress("Generating AI keyword expansions…");
        const existingSet = new Set([...staticKws, ...suggestKws].map(k => k.toLowerCase()));
        aiKws = await fetchGroqKeywords(groqKey.trim(), mapsQuery.trim(), 10, existingSet);
        console.log("AI keywords:", aiKws);
      } catch (e) {
        console.warn("[KwExpand/Groq] Failed:", e.message);
      }
    }

    // ── Step 4: merge + dedupe ────────────────────────────────────────────
    const mergedKws = mergeKeywords(staticKws, suggestKws, aiKws, kwMaxExpanded);
    console.log("Merged keywords:", mergedKws);
    console.log("Total searches (before multi-query expansion):", mergedKws.length);

    // Dedup state shared across all keyword runs
    const seenKeys = new Set();
    let totalFound = 0;
    let searchesDone = 0;

    const isDup = (place) => {
      const domain = place.website
        ? (() => { try { return new URL(place.website).hostname.replace(/^www\./, "").toLowerCase(); } catch { return null; } })()
        : null;
      const keys = [
        place.placeId ? `pid:${place.placeId}` : null,
        place.phone   ? `ph:${place.phone.replace(/\s+/g, "")}` : null,
        domain        ? `dom:${domain}` : null,
        place.name    ? `nm:${place.name.trim().toLowerCase()}` : null,
      ].filter(Boolean);
      if (keys.some(k => seenKeys.has(k))) return true;
      keys.forEach(k => seenKeys.add(k));
      return false;
    };

    ScrapeNative.start(Math.min(mergedKws.length, kwMaxSearches), "maps");

    try {
      for (let kwIdx = 0; kwIdx < mergedKws.length; kwIdx++) {
        if (mapsStopRef.current) break;
        if (searchesDone >= kwMaxSearches) break;

        const expandedKw = mergedKws[kwIdx];
        const originalKw = mapsQuery.trim();
        const extraFields = { originalKeyword: originalKw, expandedKeyword: expandedKw, sourceKeyword: expandedKw };

        setMapsProgress(`Keyword ${kwIdx + 1}/${mergedKws.length}: "${expandedKw}"…`);

        if (multiQueryMode) {
          // Feed expanded kw into existing generateSubQueries() pipeline
          const subQueries = generateSubQueries(expandedKw, mapsLocation.trim(), customAreas);
          for (let qi = 0; qi < subQueries.length; qi++) {
            if (mapsStopRef.current) break;
            if (searchesDone >= kwMaxSearches) break;
            const q = subQueries[qi];
            searchesDone++;
            setMultiQueryProgress(`Keyword ${kwIdx + 1}/${mergedKws.length} · Query ${qi + 1}/${subQueries.length}: "${q}"…`);
            try {
              const onPageResult = (newPlaces) => {
                const fresh = newPlaces.filter(p => !isDup(p)).map(p => ({ ...p, ...extraFields }));
                totalFound += fresh.length;
                setMapsResults(prev => {
                  const updated = [...prev, ...fresh];
                  setSelectedPlaces(new Set(updated.map(r => r.id)));
                  setMapsProgress(`${totalFound} unique businesses · keyword ${kwIdx + 1}/${mergedKws.length}…`);
                  return updated;
                });
              };
              await scrapeWithPlaces(
                placesKey.trim(), q, "", () => {}, mapsMaxPages, onPageResult, mapsStopRef, extraFields
              );
            } catch (e) { console.warn(`[KwExpand] Query failed: "${q}"`, e.message); }
            ScrapeNative.progress(searchesDone, kwMaxSearches, totalFound, "maps");
            if (qi < subQueries.length - 1 && !mapsStopRef.current) await new Promise(r => setTimeout(r, 800));
          }
        } else {
          // Single-query mode: run each expanded kw directly via existing scrapeWithPlaces()
          searchesDone++;
          try {
            const onPageResult = (newPlaces) => {
              const fresh = newPlaces.filter(p => !isDup(p)).map(p => ({ ...p, ...extraFields }));
              totalFound += fresh.length;
              setMapsResults(prev => {
                const updated = [...prev, ...fresh];
                setSelectedPlaces(new Set(updated.map(r => r.id)));
                setMapsProgress(`${totalFound} unique businesses · keyword ${kwIdx + 1}/${mergedKws.length}…`);
                return updated;
              });
            };
            await scrapeWithPlaces(
              placesKey.trim(), expandedKw, mapsLocation.trim(),
              () => {}, mapsMaxPages, onPageResult, mapsStopRef, extraFields
            );
          } catch (e) { console.warn(`[KwExpand] Single query failed: "${expandedKw}"`, e.message); }
          ScrapeNative.progress(searchesDone, kwMaxSearches, totalFound, "maps");
          if (kwIdx < mergedKws.length - 1 && !mapsStopRef.current) await new Promise(r => setTimeout(r, 800));
        }
      }

      setMapsProgress("");
      setMultiQueryProgress("");
      incrementMapsUsage();
      setApiCredits(prev => ({ ...prev, ...loadCreditUsage() }));
    } catch (e) {
      setMapsError(e.message || "Keyword expansion scrape failed.");
    } finally {
      setMapsLoading(false);
      ScrapeNative.stop();
    }
  };

  // ── Multi-query scrape — splits by area to bypass 60-result limit ─────────
  // Generates sub-queries from the base query + location, runs each, deduplicates
  const CITY_AREAS = {
    // ── USA — Major Cities (deep neighborhood splits) ─────────────────────
    "new york": ["Midtown Manhattan","Lower Manhattan","Upper East Side","Upper West Side","Harlem","Washington Heights","Inwood","Financial District","Greenwich Village","East Village","West Village","SoHo","Tribeca","Chelsea Manhattan","Hell's Kitchen","Murray Hill","Gramercy","Brooklyn Heights","Williamsburg","Park Slope","Bay Ridge","Flatbush","Crown Heights","Bed-Stuy","Bushwick","Sunset Park","Coney Island","Astoria Queens","Flushing Queens","Jamaica Queens","Forest Hills Queens","Jackson Heights","Long Island City","Bayside Queens","Fordham Bronx","Riverdale Bronx","Mott Haven Bronx","Co-op City Bronx","Staten Island"],
    "nyc": ["Midtown Manhattan","Lower Manhattan","Upper East Side","Harlem","Brooklyn Heights","Williamsburg","Park Slope","Astoria Queens","Flushing Queens","Fordham Bronx","Staten Island","Financial District","Chelsea Manhattan","Bed-Stuy","Jackson Heights"],
    "los angeles": ["Downtown Los Angeles","Hollywood","West Hollywood","Santa Monica","Beverly Hills","Culver City","Inglewood","Pasadena","Long Beach","Burbank","Glendale","Torrance","Compton","East Los Angeles","Van Nuys","Woodland Hills","Sherman Oaks","Studio City","Encino","North Hollywood","Chatsworth","Northridge","Canoga Park","Westwood","Brentwood","Pacific Palisades","Koreatown","Silver Lake","Echo Park","Los Feliz","Boyle Heights","Carson","Hawthorne","Redondo Beach","Hermosa Beach","Manhattan Beach","El Monte","Baldwin Park","Pomona","Ontario California"],
    "chicago": ["Loop Chicago","River North Chicago","Lincoln Park Chicago","Wicker Park Chicago","Hyde Park Chicago","Pilsen Chicago","Bridgeport Chicago","Logan Square Chicago","Andersonville Chicago","Rogers Park Chicago","Edgewater Chicago","Uptown Chicago","Lakeview Chicago","Ravenswood Chicago","Irving Park Chicago","Portage Park Chicago","Norwood Park Chicago","Evanston Illinois","Oak Park Illinois","Naperville Illinois","Schaumburg Illinois","Bolingbrook Illinois","Joliet Illinois","Elgin Illinois","Arlington Heights Illinois","Skokie Illinois","Cicero Illinois","Berwyn Illinois"],
    "houston": ["Downtown Houston","Midtown Houston","Medical Center Houston","Galleria Houston","Sugar Land","The Woodlands","Katy Texas","Pasadena Texas","Pearland Texas","Missouri City Texas","Friendswood Texas","League City Texas","Baytown Texas","Humble Texas","Kingwood Houston","Clear Lake Houston","Spring Texas","Cypress Texas","Conroe Texas","Stafford Texas","Richmond Texas","Atascocita Texas"],
    "phoenix": ["Downtown Phoenix","Scottsdale","Tempe","Mesa","Chandler","Gilbert Arizona","Glendale Arizona","Peoria Arizona","Surprise Arizona","Avondale Arizona","Goodyear Arizona","Buckeye Arizona","Queen Creek Arizona","Apache Junction Arizona","Fountain Hills Arizona","Paradise Valley Arizona","Cave Creek Arizona","Sun City Arizona","Laveen Arizona"],
    "philadelphia": ["Center City Philadelphia","North Philadelphia","South Philadelphia","West Philadelphia","Northeast Philadelphia","Germantown","Chestnut Hill","Roxborough","Manayunk","Fishtown Philadelphia","Kensington Philadelphia","Camden New Jersey","Cherry Hill New Jersey","Voorhees New Jersey","Chester Pennsylvania","Upper Darby Pennsylvania","Norristown Pennsylvania","King of Prussia Pennsylvania","Ardmore Pennsylvania"],
    "san antonio": ["Downtown San Antonio","Alamo Heights","Stone Oak","Northside San Antonio","Southside San Antonio","Leon Valley","Converse Texas","Universal City Texas","New Braunfels Texas","Seguin Texas","Schertz Texas","Cibolo Texas","Boerne Texas","Pleasanton Texas","Floresville Texas"],
    "san diego": ["Downtown San Diego","Mission Valley","La Jolla","Chula Vista","El Cajon","Escondido","Oceanside","National City","Santee California","Spring Valley","Lemon Grove","Poway","Vista California","San Marcos California","Encinitas","Carlsbad","Fallbrook","Lakeside California"],
    "dallas": ["Downtown Dallas","Uptown Dallas","Oak Cliff","Irving Texas","Plano Texas","Garland Texas","Mesquite Texas","Richardson Texas","Carrollton Texas","Frisco Texas","Allen Texas","McKinney Texas","Lewisville Texas","Denton Texas","Flower Mound Texas","Euless Texas","Bedford Texas","Hurst Texas","Grapevine Texas","Southlake Texas","Keller Texas","North Richland Hills","Duncanville Texas","DeSoto Texas","Cedar Hill Texas","Rowlett Texas","Rockwall Texas"],
    "san jose": ["Downtown San Jose","Willow Glen","Almaden Valley","Santa Clara","Sunnyvale","Mountain View","Milpitas","Campbell California","Los Gatos","Saratoga California","Cupertino","Los Altos","Morgan Hill","Gilroy California","Fremont California","Newark California","Union City California"],
    "austin": ["Downtown Austin","South Congress Austin","East Austin","Round Rock Texas","Cedar Park Texas","Pflugerville Texas","Georgetown Texas","Kyle Texas","Buda Texas","Leander Texas","Hutto Texas","Manor Texas","Bastrop Texas","Lakeway Texas","Bee Cave Texas","Dripping Springs Texas","San Marcos Texas"],
    "jacksonville": ["Downtown Jacksonville","Southside Jacksonville","Northside Jacksonville","Arlington Jacksonville","Mandarin Jacksonville","Orange Park Florida","Fleming Island Florida","St Augustine Florida","Fernandina Beach Florida","Ponte Vedra Florida"],
    "san francisco": ["Financial District SF","Mission District SF","Castro SF","Richmond District SF","Sunset District SF","South of Market SF","Haight-Ashbury","Noe Valley","Potrero Hill","Dogpatch SF","Oakland California","Berkeley California","Daly City","San Mateo","Redwood City","Palo Alto","Menlo Park","Fremont California","Hayward California","Pleasanton California","Livermore California","Walnut Creek California","Concord California"],
    "columbus": ["Downtown Columbus","Short North Columbus","Westerville","Dublin Ohio","Grove City Ohio","Hilliard Ohio","Gahanna Ohio","Reynoldsburg Ohio","Whitehall Ohio","Bexley Ohio","Upper Arlington Ohio","Worthington Ohio","New Albany Ohio","Powell Ohio","Pickerington Ohio","Lancaster Ohio","Newark Ohio"],
    "fort worth": ["Downtown Fort Worth","Arlington Texas","Mansfield Texas","Euless Texas","Hurst Texas","Bedford Texas","North Richland Hills","Watauga Texas","Keller Texas","Southlake Texas","Colleyville Texas","Grapevine Texas","Benbrook Texas","Saginaw Texas","Azle Texas","Granbury Texas","Burleson Texas","Crowley Texas"],
    "charlotte": ["Uptown Charlotte","South End Charlotte","Ballantyne Charlotte","Concord North Carolina","Gastonia North Carolina","Matthews North Carolina","Huntersville North Carolina","Mooresville North Carolina","Kannapolis North Carolina","Monroe North Carolina","Rock Hill South Carolina","Fort Mill South Carolina","Indian Trail North Carolina","Mint Hill North Carolina","Pineville North Carolina","Waxhaw North Carolina"],
    "indianapolis": ["Downtown Indianapolis","Broad Ripple Indianapolis","Carmel Indiana","Fishers Indiana","Greenwood Indiana","Lawrence Indiana","Avon Indiana","Plainfield Indiana","Noblesville Indiana","Westfield Indiana","Zionsville Indiana","Brownsburg Indiana","Mooresville Indiana","Franklin Indiana"],
    "seattle": ["Downtown Seattle","Capitol Hill Seattle","Bellevue Washington","Redmond Washington","Kirkland Washington","Renton Washington","Tacoma Washington","Bothell Washington","Kenmore Washington","Shoreline Washington","Lynnwood Washington","Edmonds Washington","Everett Washington","Marysville Washington","Auburn Washington","Kent Washington","Federal Way Washington","Burien Washington","Issaquah Washington","Sammamish Washington"],
    "denver": ["Downtown Denver","Cherry Creek Denver","Aurora Colorado","Lakewood Colorado","Westminster Colorado","Thornton Colorado","Arvada Colorado","Englewood Colorado","Littleton Colorado","Centennial Colorado","Parker Colorado","Castle Rock Colorado","Highlands Ranch Colorado","Lone Tree Colorado","Brighton Colorado","Commerce City Colorado","Northglenn Colorado","Wheat Ridge Colorado","Greenwood Village Colorado"],
    "nashville": ["Downtown Nashville","East Nashville","Brentwood Tennessee","Franklin Tennessee","Murfreesboro Tennessee","Hendersonville Tennessee","Smyrna Tennessee","La Vergne Tennessee","Mt Juliet Tennessee","Gallatin Tennessee","Goodlettsville Tennessee","Madison Tennessee","Antioch Nashville","Hermitage Nashville","Nolensville Tennessee","Spring Hill Tennessee"],
    "oklahoma city": ["Downtown Oklahoma City","Edmond Oklahoma","Moore Oklahoma","Norman Oklahoma","Midwest City Oklahoma","Yukon Oklahoma","Mustang Oklahoma","Del City Oklahoma","Choctaw Oklahoma","Piedmont Oklahoma","Weatherford Oklahoma","El Reno Oklahoma","Guthrie Oklahoma"],
    "el paso": ["Downtown El Paso","East El Paso","West El Paso","Northeast El Paso","Socorro Texas","Horizon City Texas","Canutillo Texas","Fort Bliss","Sunland Park New Mexico"],
    "washington dc": ["Downtown DC","Georgetown DC","Capitol Hill DC","Alexandria Virginia","Arlington Virginia","Bethesda Maryland","Silver Spring Maryland","Rockville Maryland","Gaithersburg Maryland","Germantown Maryland","Frederick Maryland","Fairfax Virginia","Reston Virginia","Herndon Virginia","Sterling Virginia","Ashburn Virginia","Leesburg Virginia","Manassas Virginia","Woodbridge Virginia","Springfield Virginia","Tysons Virginia","McLean Virginia"],
    "boston": ["Downtown Boston","Back Bay Boston","South Boston","Cambridge Massachusetts","Somerville Massachusetts","Quincy Massachusetts","Brookline Massachusetts","Newton Massachusetts","Waltham Massachusetts","Watertown Massachusetts","Malden Massachusetts","Medford Massachusetts","Arlington Massachusetts","Lexington Massachusetts","Worcester Massachusetts","Framingham Massachusetts","Lowell Massachusetts","Lynn Massachusetts","Peabody Massachusetts","Salem Massachusetts"],
    "memphis": ["Downtown Memphis","Midtown Memphis","East Memphis","Bartlett Tennessee","Germantown Tennessee","Collierville Tennessee","Cordova Tennessee","Lakeland Tennessee","Arlington Tennessee","Millington Tennessee","Southaven Mississippi","Olive Branch Mississippi","Horn Lake Mississippi"],
    "louisville": ["Downtown Louisville","East Louisville","South Louisville","St Matthews Louisville","Jeffersontown Kentucky","Jeffersonville Indiana","New Albany Indiana","Clarksville Indiana","Shelbyville Kentucky","Elizabethtown Kentucky"],
    "portland": ["Downtown Portland","Pearl District Portland","Beaverton Oregon","Hillsboro Oregon","Gresham Oregon","Lake Oswego Oregon","Tigard Oregon","Tualatin Oregon","Wilsonville Oregon","Oregon City","Milwaukie Oregon","Happy Valley Oregon","Clackamas Oregon","Camas Washington","Vancouver Washington"],
    "las vegas": ["Downtown Las Vegas","Henderson Nevada","Summerlin Las Vegas","North Las Vegas","Boulder City Nevada","Mesquite Nevada","Paradise Nevada","Spring Valley Nevada","Enterprise Nevada","Whitney Nevada","Winchester Nevada","Sunrise Manor Nevada"],
    "milwaukee": ["Downtown Milwaukee","Bay View Milwaukee","Wauwatosa Wisconsin","West Allis Wisconsin","Brookfield Wisconsin","Waukesha Wisconsin","Menomonee Falls Wisconsin","Germantown Wisconsin","Mequon Wisconsin","Racine Wisconsin","Kenosha Wisconsin","Oak Creek Wisconsin","Franklin Wisconsin","Greenfield Wisconsin","Greendale Wisconsin"],
    "albuquerque": ["Downtown Albuquerque","Rio Rancho New Mexico","South Valley Albuquerque","Northeast Heights Albuquerque","Northwest Albuquerque","Bernalillo New Mexico","Corrales New Mexico","Los Lunas New Mexico","Belen New Mexico","Edgewood New Mexico","Tijeras New Mexico"],
    "tucson": ["Downtown Tucson","Marana Arizona","Oro Valley Arizona","South Tucson","Sahuarita Arizona","Green Valley Arizona","Sierra Vista Arizona","Nogales Arizona","Vail Arizona","Catalina Arizona"],
    "fresno": ["Downtown Fresno","Clovis California","North Fresno","South Fresno","West Fresno","Madera California","Selma California","Sanger California","Reedley California","Kingsburg California","Fowler California","Kerman California"],
    "sacramento": ["Downtown Sacramento","Elk Grove California","Roseville California","Folsom California","Citrus Heights California","Rancho Cordova","West Sacramento","Davis California","Woodland California","Rocklin California","Lincoln California","Auburn California","Placerville California","Yuba City California"],
    "kansas city": ["Downtown Kansas City","Overland Park Kansas","Olathe Kansas","Lenexa Kansas","Lee Summit Missouri","Independence Missouri","Blue Springs Missouri","Raymore Missouri","Belton Missouri","Grandview Missouri","Liberty Missouri","Shawnee Kansas","Prairie Village Kansas","Leawood Kansas","Gardner Kansas","Merriam Kansas"],
    "atlanta": ["Downtown Atlanta","Buckhead Atlanta","Midtown Atlanta","Decatur Georgia","Marietta Georgia","Sandy Springs Georgia","Alpharetta Georgia","Roswell Georgia","Johns Creek Georgia","Dunwoody Georgia","Smyrna Georgia","Kennesaw Georgia","Acworth Georgia","Woodstock Georgia","Canton Georgia","Cumming Georgia","Gainesville Georgia","Lawrenceville Georgia","Duluth Georgia","Norcross Georgia","Peachtree City Georgia","Newnan Georgia","Fayetteville Georgia","Stockbridge Georgia","McDonough Georgia","Conyers Georgia"],
    "miami": ["Downtown Miami","Coral Gables","Hialeah Florida","Miami Beach","Doral Florida","Aventura Florida","Homestead Florida","North Miami","Kendall Florida","Pembroke Pines","Hollywood Florida","Miramar Florida","Davie Florida","Sunrise Florida","Plantation Florida","Fort Lauderdale","Pompano Beach","Boca Raton","Delray Beach","West Palm Beach","Boynton Beach","Lake Worth Florida","Wellington Florida","Cutler Bay Florida","Palmetto Bay Florida","Pinecrest Florida"],
    "minneapolis": ["Downtown Minneapolis","Saint Paul Minnesota","Bloomington Minnesota","Plymouth Minnesota","Brooklyn Park Minnesota","Edina Minnesota","Minnetonka Minnesota","Eden Prairie Minnesota","Burnsville Minnesota","Apple Valley Minnesota","Eagan Minnesota","Woodbury Minnesota","Maple Grove Minnesota","Coon Rapids Minnesota","Blaine Minnesota","Fridley Minnesota","Roseville Minnesota","Maplewood Minnesota","Shakopee Minnesota","Prior Lake Minnesota","Savage Minnesota"],
    "new orleans": ["Downtown New Orleans","French Quarter","Garden District New Orleans","Metairie Louisiana","Kenner Louisiana","Chalmette Louisiana","Gretna Louisiana","Harvey Louisiana","Marrero Louisiana","Westwego Louisiana","Slidell Louisiana","Covington Louisiana","Mandeville Louisiana","Hammond Louisiana"],
    "cleveland": ["Downtown Cleveland","Parma Ohio","Lakewood Ohio","Euclid Ohio","Strongsville Ohio","Mentor Ohio","Lorain Ohio","Elyria Ohio","Medina Ohio","Brunswick Ohio","Solon Ohio","Beachwood Ohio","Westlake Ohio","North Olmsted Ohio","Middleburg Heights Ohio","Garfield Heights Ohio","Shaker Heights Ohio","Cleveland Heights Ohio"],
    "tampa": ["Downtown Tampa","St Petersburg Florida","Clearwater Florida","Brandon Florida","Wesley Chapel Florida","Riverview Florida","Lakeland Florida","Plant City Florida","Bradenton Florida","Sarasota Florida","New Port Richey Florida","Spring Hill Florida","Land O Lakes Florida","Lutz Florida","Odessa Florida","Apollo Beach Florida","Sun City Center Florida","Ruskin Florida"],
    "baltimore": ["Downtown Baltimore","Towson Maryland","Dundalk Maryland","Essex Maryland","Columbia Maryland","Catonsville Maryland","Pikesville Maryland","Owings Mills Maryland","Timonium Maryland","Ellicott City Maryland","Bowie Maryland","Laurel Maryland","College Park Maryland","Greenbelt Maryland","Annapolis Maryland","Glen Burnie Maryland","Severn Maryland","Pasadena Maryland"],
    "raleigh": ["Downtown Raleigh","Cary North Carolina","Durham North Carolina","Chapel Hill North Carolina","Wake Forest North Carolina","Apex North Carolina","Holly Springs North Carolina","Fuquay-Varina North Carolina","Garner North Carolina","Clayton North Carolina","Morrisville North Carolina","Knightdale North Carolina"],
    "cincinnati": ["Downtown Cincinnati","Covington Kentucky","Florence Kentucky","Mason Ohio","West Chester Ohio","Fairfield Ohio","Blue Ash Ohio","Sharonville Ohio","Norwood Ohio","Anderson Township Ohio","Milford Ohio","Loveland Ohio","Lebanon Ohio","Middletown Ohio","Hamilton Ohio","Dayton Ohio","Kettering Ohio","Beavercreek Ohio","Centerville Ohio"],
    "pittsburgh": ["Downtown Pittsburgh","Squirrel Hill Pittsburgh","Mt Lebanon Pennsylvania","Bethel Park Pennsylvania","Monroeville Pennsylvania","Cranberry Township Pennsylvania","Ross Township Pennsylvania","North Hills Pittsburgh","South Hills Pittsburgh","McKeesport Pennsylvania","Greensburg Pennsylvania","New Kensington Pennsylvania","Butler Pennsylvania","Wexford Pennsylvania","Allison Park Pennsylvania"],
    "richmond": ["Downtown Richmond","Henrico Virginia","Chesterfield Virginia","Midlothian Virginia","Short Pump Virginia","Glen Allen Virginia","Mechanicsville Virginia","Colonial Heights Virginia","Petersburg Virginia","Hopewell Virginia","Chester Virginia","Ashland Virginia"],
    "orlando": ["Downtown Orlando","Kissimmee Florida","Sanford Florida","Altamonte Springs Florida","Winter Park Florida","Lake Mary Florida","Apopka Florida","Ocoee Florida","Winter Garden Florida","Clermont Florida","Leesburg Florida","Deltona Florida","Daytona Beach Florida","Palm Bay Florida","Melbourne Florida","Titusville Florida"],
    "st louis": ["Downtown St Louis","Clayton Missouri","Chesterfield Missouri","Florissant Missouri","Belleville Illinois","O Fallon Missouri","St Charles Missouri","Wentzville Missouri","St Peters Missouri","Ballwin Missouri","Kirkwood Missouri","Webster Groves Missouri","Fenton Missouri","Arnold Missouri","Festus Missouri","Edwardsville Illinois"],
    "salt lake city": ["Downtown Salt Lake City","Sandy Utah","West Valley City","Provo Utah","Orem Utah","Ogden Utah","Layton Utah","Taylorsville Utah","Murray Utah","Millcreek Utah","Cottonwood Heights Utah","Holladay Utah","Draper Utah","South Jordan Utah","Riverton Utah","Herriman Utah","Eagle Mountain Utah","Lehi Utah","American Fork Utah","Spanish Fork Utah","Springville Utah","Tooele Utah"],
    "virginia beach": ["Virginia Beach Oceanfront","Chesapeake Virginia","Norfolk Virginia","Portsmouth Virginia","Suffolk Virginia","Hampton Virginia","Newport News Virginia","Williamsburg Virginia","Yorktown Virginia"],
    "colorado springs": ["Downtown Colorado Springs","Pueblo Colorado","Fountain Colorado","Security Colorado","Widefield Colorado","Falcon Colorado","Monument Colorado","Manitou Springs Colorado","Canon City Colorado"],
    "omaha": ["Downtown Omaha","Bellevue Nebraska","Papillion Nebraska","La Vista Nebraska","Ralston Nebraska","Gretna Nebraska","Millard Omaha","Elkhorn Omaha","Council Bluffs Iowa","Fremont Nebraska","Lincoln Nebraska"],
    "long island": ["Nassau County Long Island","Suffolk County Long Island","Hempstead New York","Valley Stream","Freeport New York","Rockville Centre","Garden City New York","Mineola New York","Great Neck New York","Hicksville New York","Levittown New York","Farmingdale New York","Babylon New York","Bay Shore New York","Islip New York","Brentwood New York","Huntington New York","Patchogue New York"],
    "new jersey": ["Newark New Jersey","Jersey City","Paterson New Jersey","Elizabeth New Jersey","Edison New Jersey","Woodbridge New Jersey","Lakewood New Jersey","Toms River New Jersey","Hamilton New Jersey","Trenton New Jersey","Clifton New Jersey","Camden New Jersey","Brick New Jersey","Cherry Hill New Jersey","Passaic New Jersey","Union City New Jersey","Bayonne New Jersey","East Orange New Jersey","Vineland New Jersey"],
    "atlanta suburbs": ["Alpharetta Georgia","Roswell Georgia","Johns Creek Georgia","Dunwoody Georgia","Smyrna Georgia","Kennesaw Georgia","Acworth Georgia","Woodstock Georgia","Canton Georgia","Cumming Georgia","Gainesville Georgia","Lawrenceville Georgia","Duluth Georgia","Norcross Georgia","Peachtree City Georgia","Newnan Georgia","Fayetteville Georgia","Stockbridge Georgia","McDonough Georgia","Conyers Georgia"],
    "birmingham al": ["Downtown Birmingham Alabama","Hoover Alabama","Tuscaloosa Alabama","Decatur Alabama","Bessemer Alabama","Vestavia Hills Alabama","Pelham Alabama","Alabaster Alabama","Trussville Alabama","Gardendale Alabama","Hueytown Alabama","Moody Alabama","Calera Alabama","Helena Alabama"],
    "buffalo": ["Downtown Buffalo New York","Amherst New York","Cheektowaga New York","Tonawanda New York","Lockport New York","Niagara Falls New York","West Seneca New York","Lackawanna New York","Depew New York","Lancaster New York","Hamburg New York","Orchard Park New York","Williamsville New York"],
    "greenville sc": ["Downtown Greenville South Carolina","Spartanburg South Carolina","Anderson South Carolina","Simpsonville South Carolina","Mauldin South Carolina","Easley South Carolina","Greer South Carolina","Taylors South Carolina","Travelers Rest South Carolina","Fountain Inn South Carolina","Laurens South Carolina","Gaffney South Carolina"],
    "knoxville": ["Downtown Knoxville Tennessee","Maryville Tennessee","Morristown Tennessee","Oak Ridge Tennessee","Alcoa Tennessee","Sevierville Tennessee","Lenoir City Tennessee","Farragut Tennessee","Powell Tennessee","Halls Tennessee"],
    "tulsa": ["Downtown Tulsa Oklahoma","Broken Arrow Oklahoma","Sapulpa Oklahoma","Owasso Oklahoma","Bixby Oklahoma","Jenks Oklahoma","Sand Springs Oklahoma","Claremore Oklahoma","Bartlesville Oklahoma"],
    "bakersfield": ["Downtown Bakersfield California","Delano California","Wasco California","Shafter California","Arvin California","Tehachapi California","McFarland California"],
    "baton rouge": ["Downtown Baton Rouge","Denham Springs Louisiana","Zachary Louisiana","Central Louisiana","Baker Louisiana","Port Allen Louisiana","Gonzales Louisiana","Prairieville Louisiana"],
    "lexington ky": ["Downtown Lexington Kentucky","Nicholasville Kentucky","Georgetown Kentucky","Winchester Kentucky","Richmond Kentucky","Frankfort Kentucky","Paris Kentucky"],
    "stockton": ["Downtown Stockton California","Lodi California","Manteca California","Tracy California","Modesto California","Turlock California","Merced California"],
    "corpus christi": ["Downtown Corpus Christi","Ingleside Texas","Portland Texas","Rockport Texas","Calallen Texas","Flour Bluff Texas","Kingsville Texas"],
    "riverside": ["Downtown Riverside California","San Bernardino California","Ontario California","Rancho Cucamonga","Fontana California","Rialto California","Colton California","Redlands California","Yucaipa California","Beaumont California","Palm Springs California","Cathedral City California","Palm Desert California","Indio California","Murrieta California","Temecula California","Menifee California","Hemet California","Perris California"],
    "anaheim": ["Downtown Anaheim","Garden Grove California","Santa Ana California","Irvine California","Orange California","Fullerton California","Buena Park California","Westminster California","Huntington Beach","Newport Beach","Costa Mesa","Laguna Niguel","Mission Viejo","Lake Forest California","Aliso Viejo","San Clemente California","Yorba Linda California","Brea California","La Habra California","Whittier California","Norwalk California"],
    "wichita": ["Downtown Wichita Kansas","Derby Kansas","Haysville Kansas","Valley Center Kansas","Andover Kansas","Augusta Kansas","El Dorado Kansas","Hutchinson Kansas","Salina Kansas"],
    "anchorage": ["Downtown Anchorage","Eagle River Alaska","Wasilla Alaska","Palmer Alaska","Chugiak Alaska","Birchwood Alaska","Girdwood Alaska","Kenai Alaska","Soldotna Alaska"],
    "providence": ["Downtown Providence Rhode Island","Cranston Rhode Island","Pawtucket Rhode Island","Woonsocket Rhode Island","North Providence Rhode Island","Warwick Rhode Island","East Providence Rhode Island"],
    "hartford": ["Downtown Hartford Connecticut","New Haven Connecticut","Bridgeport Connecticut","Stamford Connecticut","Waterbury Connecticut","Norwalk Connecticut","Danbury Connecticut","Meriden Connecticut","Bristol Connecticut","New Britain Connecticut","Middletown Connecticut","Norwich Connecticut"],

    // ── UK — Deep City Splits ─────────────────────────────────────────────
    "london": ["City of London","Westminster","Canary Wharf","Shoreditch","Camden","Islington","Hackney","Southwark","Lambeth","Wandsworth","Hammersmith","Kensington","Croydon","Bromley","Ealing","Brent","Harrow","Barnet","Enfield","Haringey","Waltham Forest","Redbridge","Havering","Barking","Newham","Tower Hamlets","Lewisham","Greenwich","Bexley","Sutton","Merton","Kingston upon Thames","Richmond upon Thames","Hounslow"],
    "birmingham uk": ["Birmingham City Centre","Edgbaston","Handsworth","Sutton Coldfield","Solihull","Wolverhampton","Coventry","Walsall","West Bromwich","Dudley","Stourbridge","Halesowen","Redditch","Tamworth","Lichfield","Burton upon Trent"],
    "manchester": ["Manchester City Centre","Salford","Trafford","Stockport","Bolton","Oldham","Rochdale","Bury","Wigan","Ashton-under-Lyne","Altrincham","Sale Manchester","Urmston","Stretford","Eccles","Swinton Manchester","Leigh Greater Manchester"],
    "leeds": ["Leeds City Centre","Headingley","Chapeltown","Morley","Pudsey","Horsforth","Bradford","Wakefield","Harrogate","Huddersfield","Halifax","Dewsbury","Castleford","Keighley","Shipley"],
    "glasgow": ["Glasgow City Centre","East End Glasgow","West End Glasgow","Paisley","Rutherglen","Hamilton","Motherwell","Clydebank","Dumbarton","Airdrie","Coatbridge","Kilmarnock","Greenock","Livingston","Falkirk"],
    "liverpool": ["Liverpool City Centre","Birkenhead","Bootle","Huyton","St Helens","Warrington","Widnes","Runcorn","Ellesmere Port","Chester","Southport","Formby","Crosby Liverpool","Kirkby Liverpool"],
    "bristol": ["Bristol City Centre","Clifton Bristol","Bedminster","Southmead","Bath Somerset","Weston-super-Mare","Taunton","Swindon","Chippenham","Gloucester","Cheltenham","Clevedon","Portishead","Nailsea","Yate"],
    "sheffield": ["Sheffield City Centre","Hillsborough Sheffield","Rotherham","Doncaster","Barnsley","Chesterfield","Mansfield","Worksop","Stocksbridge","Chapeltown Sheffield","Mosborough"],
    "edinburgh": ["Edinburgh City Centre","Leith","Portobello","Musselburgh","Livingston","Kirkcaldy","Dunfermline","Perth Scotland","Dundee","Stirling"],
    "nottingham": ["Nottingham City Centre","Beeston","Arnold","Derby","Leicester","Loughborough","Hucknall","Ilkeston","Long Eaton","Eastwood Nottinghamshire","Stapleford","Kirkby-in-Ashfield"],
    "newcastle uk": ["Newcastle City Centre","Gateshead","Sunderland","Durham","Middlesbrough","Stockton-on-Tees","Darlington","Hartlepool","South Shields","North Shields","Tynemouth","Cramlington"],
    "cardiff": ["Cardiff City Centre","Barry Wales","Penarth","Pontypridd","Merthyr Tydfil","Bridgend","Swansea","Neath","Port Talbot","Llanelli","Newport Wales","Cwmbran","Caerphilly"],
    "southampton": ["Southampton City Centre","Portsmouth","Eastleigh","Fareham","Gosport","Havant","Winchester","Romsey","Totton","Hedge End","Chandlers Ford","Bishopstoke","Waterlooville"],
    "leicester": ["Leicester City Centre","Loughborough","Hinckley","Melton Mowbray","Market Harborough","Wigston","Oadby","Blaby","Coalville","Nuneaton","Rugby","Lutterworth","Syston","Birstall Leicestershire"],
    "oxford": ["Oxford City Centre","Abingdon","Didcot","Banbury","Witney","Bicester","Kidlington","Carterton","Chipping Norton","Faringdon","Wallingford","Thame","Henley-on-Thames"],
    "cambridge uk": ["Cambridge City Centre","Ely Cambridgeshire","Huntingdon","St Ives Cambridgeshire","March Cambridgeshire","Wisbech","Peterborough","Royston","Saffron Walden","Newmarket","Haverhill","Bury St Edmunds"],
    "plymouth uk": ["Plymouth City Centre","Exeter","Torquay","Paignton","Newton Abbot","Barnstaple","Saltash Cornwall","Liskeard","Tavistock","Ivybridge","Plympton","Plymstock"],
    "stoke": ["Stoke-on-Trent","Stafford","Cannock","Lichfield","Newcastle-under-Lyme","Crewe","Nantwich","Congleton","Macclesfield"],
    "wolverhampton": ["Wolverhampton City Centre","Walsall","West Bromwich","Dudley","Stourbridge","Halesowen","Tipton","Wednesbury","Oldbury","Sandwell"],
    "brighton": ["Brighton City Centre","Hove","Worthing","Eastbourne","Hastings","Lewes","Crawley","Horsham","Chichester","Bognor Regis"],

    // ── Canada — Deep City Splits ─────────────────────────────────────────
    "toronto": ["Downtown Toronto","North York","Scarborough","Etobicoke","Mississauga","Brampton","Markham","Vaughan","Richmond Hill","Pickering","Ajax Ontario","Whitby Ontario","Oshawa","Burlington Ontario","Oakville Ontario","Milton Ontario","Hamilton Ontario","St Catharines","Niagara Falls Ontario","Barrie Ontario","Newmarket Ontario","Aurora Ontario"],
    "montreal": ["Downtown Montreal","Plateau Mont Royal","Laval Quebec","Longueuil Quebec","Brossard Quebec","Saint Laurent Montreal","Verdun Montreal","Rosemont Montreal","Westmount","NDG Montreal","Pointe-Claire","Dollard-des-Ormeaux","Kirkland Quebec","Terrebonne Quebec","Repentigny Quebec","Saint-Jerome Quebec","Blainville Quebec"],
    "vancouver": ["Downtown Vancouver","Burnaby","Richmond BC","Surrey BC","Coquitlam","Langley BC","Abbotsford BC","Maple Ridge BC","Port Coquitlam","Port Moody","New Westminster","Delta BC","White Rock BC","North Vancouver","West Vancouver","Chilliwack BC","Mission BC"],
    "calgary": ["Downtown Calgary","Northwest Calgary","Northeast Calgary","Southeast Calgary","Southwest Calgary","Airdrie Alberta","Cochrane Alberta","Chestermere Alberta","Okotoks Alberta","High River Alberta","Strathmore Alberta","Canmore Alberta"],
    "edmonton": ["Downtown Edmonton","West Edmonton","North Edmonton","South Edmonton","St Albert Alberta","Sherwood Park Alberta","Spruce Grove Alberta","Leduc Alberta","Fort Saskatchewan Alberta","Beaumont Alberta","Stony Plain Alberta"],
    "ottawa": ["Downtown Ottawa","Gatineau Quebec","Kanata Ontario","Orleans Ontario","Barrhaven Ottawa","Nepean Ontario","Gloucester Ontario","Rockland Ontario","Carleton Place Ontario","Kemptville Ontario"],
    "winnipeg": ["Downtown Winnipeg","St Boniface Winnipeg","St Vital Winnipeg","Fort Garry Winnipeg","Transcona Winnipeg","Selkirk Manitoba","Steinbach Manitoba","Portage la Prairie","Brandon Manitoba"],
    "quebec city": ["Old Quebec City","Sainte-Foy","Charlesbourg","Beauport Quebec","Levis Quebec","Saint-Romuald Quebec","Breakeyville Quebec","Pintendre Quebec","Lac-Beauport Quebec","Boischatel Quebec","LAncienne-Lorette Quebec"],
    "hamilton ontario": ["Downtown Hamilton","Ancaster Ontario","Dundas Ontario","Stoney Creek","Grimsby Ontario","Burlington Ontario","Niagara Falls Ontario","St Catharines","Welland Ontario"],
    "kitchener": ["Downtown Kitchener","Waterloo Ontario","Cambridge Ontario","Guelph Ontario","Brantford Ontario","Woodstock Ontario","Stratford Ontario"],
    "london ontario": ["Downtown London Ontario","Sarnia Ontario","Windsor Ontario","Chatham Ontario","Brantford Ontario","Woodstock Ontario","St Thomas Ontario"],
    "halifax": ["Downtown Halifax","Dartmouth Nova Scotia","Bedford Nova Scotia","Sackville Nova Scotia","Truro Nova Scotia","Bridgewater Nova Scotia","New Glasgow Nova Scotia"],
    "victoria bc": ["Downtown Victoria BC","Saanich BC","Oak Bay BC","Langford BC","Colwood BC","View Royal BC","Esquimalt BC","Sidney BC","Duncan BC","Nanaimo BC","Parksville BC","Courtenay BC"],
    "saskatoon": ["Downtown Saskatoon","Warman Saskatchewan","Martensville Saskatchewan","Osler Saskatchewan","Dalmeny Saskatchewan","Langham Saskatchewan","Delisle Saskatchewan","Outlook Saskatchewan","Humboldt Saskatchewan"],

    // ── UAE — Deep Splits ─────────────────────────────────────────────────
    "dubai": ["Downtown Dubai","Deira","Bur Dubai","Jumeirah","Business Bay","Dubai Marina","JLT Jumeirah Lake Towers","Al Quoz","Mirdif","International City","Discovery Gardens","Dubai Silicon Oasis","Al Barsha","Jumeirah Village Circle","Dubai Sports City","Motor City Dubai","Palm Jumeirah","Dubai Creek","Oud Metha","Karama Dubai","Satwa Dubai","Mankhool Dubai","Al Nahda Dubai","Al Qusais","Al Twar","Muhaisnah","Al Mamzar","Rashidiya Dubai","Al Warqa","Al Mizhar","Nad Al Hammar","Al Khawaneej","Hatta Dubai"],
    "abu dhabi": ["Abu Dhabi City Centre","Khalifa City","Al Reem Island","Mussafah","Al Ain","Khalidiyah","Corniche Abu Dhabi","Hamdan Street","Tourist Club Area","Al Mushrif","Mohammed Bin Zayed City","Shakhbout City","Yas Island","Saadiyat Island","Al Falah","Baniyas Abu Dhabi","Madinat Zayed","Al Shahama","Al Wathba","Al Shamkha"],
    "sharjah": ["Sharjah City Centre","Al Nabba Sharjah","Industrial Area Sharjah","Al Majaz Sharjah","Al Taawun Sharjah","Muwaileh Sharjah","University City Sharjah","Al Khan Sharjah","Al Qasimia","Rolla Sharjah"],
    "ajman": ["Ajman City Centre","Al Nuaimiya Ajman","Al Rashidiya Ajman","Al Jurf Ajman","Al Rumaila Ajman","Emirates City Ajman"],
    "ras al khaimah": ["RAK City Centre","Al Nakheel RAK","Mina Al Arab RAK","Al Marjan Island","Al Hamra RAK","Khuzam RAK"],
    "al ain": ["Al Ain City Centre","Jimi Al Ain","Al Hili","Al Muwaiji","Zakher Al Ain","Al Khabisi","Al Jahili","Al Qattara"],
    "fujairah": ["Fujairah City","Dibba Fujairah","Khor Fakkan","Kalba Sharjah","Masafi Fujairah"],

    // ── Generic fallback ──────────────────────────────────────────────────
    "_generic": ["city centre","downtown","north","south","east","west","central","uptown","midtown","suburbs","old town","business district"],
  };

  const generateSubQueries = (query, location, customAreasText = "") => {
    const loc = location.toLowerCase().trim();

    // Use custom areas if provided
    let areas = [];
    if (customAreasText.trim()) {
      areas = customAreasText.split(/[\n,]+/).map(a => a.trim()).filter(Boolean);
    } else {
      // Match city to areas database
      const cityKey = Object.keys(CITY_AREAS).find(key =>
        key !== "_generic" && (loc.includes(key) || key.includes(loc))
      );
      areas = cityKey ? CITY_AREAS[cityKey] : CITY_AREAS["_generic"].map(a =>
        location ? `${a} ${location}` : a
      );
    }

    // Keyword variations based on query type
    const q = query.toLowerCase();
    const keywords = [query];
    if (q.includes("lawyer") || q.includes("attorney") || q.includes("legal")) {
      keywords.push(query.replace(/lawyer/gi, "attorney"));
      keywords.push(query.replace(/attorney/gi, "lawyer"));
      keywords.push(`${query} firm`);
      keywords.push(`${query} office`);
    } else if (q.includes("doctor") || q.includes("physician") || q.includes("clinic")) {
      keywords.push(query.replace(/doctor/gi, "physician"));
      keywords.push(query.replace(/clinic/gi, "medical center"));
      keywords.push(`${query} practice`);
    } else if (q.includes("dentist") || q.includes("dental")) {
      keywords.push(query.replace(/dentist/gi, "dental clinic"));
      keywords.push(`${query} practice`);
    } else if (q.includes("real estate") || q.includes("realtor")) {
      keywords.push(query.replace(/realtor/gi, "real estate agent"));
      keywords.push(`${query} agency`);
      keywords.push(`${query} broker`);
    } else if (q.includes("accountant") || q.includes("cpa") || q.includes("accounting")) {
      keywords.push(query.replace(/accountant/gi, "CPA"));
      keywords.push(`${query} firm`);
      keywords.push(`tax ${query}`);
    } else if (q.includes("insurance")) {
      keywords.push(`${query} agent`);
      keywords.push(`${query} broker`);
      keywords.push(`${query} company`);
    } else if (q.includes("restaurant") || q.includes("food")) {
      keywords.push(query.replace(/restaurant/gi, "eatery"));
      keywords.push(`${query} near me`);
    } else if (q.includes("contractor") || q.includes("plumber") || q.includes("electrician")) {
      keywords.push(`${query} company`);
      keywords.push(`${query} service`);
      keywords.push(`local ${query}`);
    } else {
      // Generic variations for any business type
      keywords.push(`${query} company`);
      keywords.push(`${query} near me`);
    }

    // Build all combinations: keyword × area — deduplicated
    const queries = new Set();
    for (const kw of [...new Set(keywords)]) {
      for (const area of areas) {
        queries.add(location && !area.toLowerCase().includes(location.toLowerCase())
          ? `${kw} in ${area}, ${location}`
          : `${kw} in ${area}`
        );
      }
    }
    return [...queries];
  };

  const runMultiQueryScrape = async () => {
    if (!mapsQuery.trim() || !placesKey.trim()) return;
    setMapsLoading(true);
    setMapsError("");
    setMapsResults([]);
    setSelectedPlaces(new Set());
    mapsStopRef.current = false;

    const subQueries = generateSubQueries(mapsQuery.trim(), mapsLocation.trim(), customAreas);
    ScrapeNative.start(subQueries.length, "maps"); // keep app alive — multi-query can run a long time
    const seenNames = new Set(); // deduplicate by business name
    let totalFound = 0;

    try {
      for (let i = 0; i < subQueries.length; i++) {
        if (mapsStopRef.current) break;
        const q = subQueries[i];
        setMultiQueryProgress(`Query ${i + 1}/${subQueries.length}: "${q}"…`);

        try {
          const onPageResult = (newPlaces) => {
            setMapsResults(prev => {
              // Deduplicate by business name + address
              const fresh = newPlaces.filter(p => {
                const key = `${p.name}|${p.address}`;
                if (seenNames.has(key)) return false;
                seenNames.add(key);
                return true;
              });
              totalFound += fresh.length;
              const updated = [...prev, ...fresh];
              setSelectedPlaces(new Set(updated.map(r => r.id)));
              setMapsProgress(`${totalFound} unique businesses found across ${i + 1} searches…`);
              return updated;
            });
          };

          await scrapeWithPlaces(
            placesKey.trim(), q, "",
            () => {}, mapsMaxPages, onPageResult, mapsStopRef
          );
        } catch (e) {
          console.warn(`[MultiQuery] Query failed: "${q}"`, e.message);
        }

        ScrapeNative.progress(i + 1, subQueries.length, totalFound, "maps");

        // Small delay between queries to avoid rate limiting
        if (i < subQueries.length - 1 && !mapsStopRef.current) {
          await new Promise(r => setTimeout(r, 1000));
        }
      }

      setMapsProgress("");
      setMultiQueryProgress("");
    } catch (e) {
      setMapsError(e.message || "Multi-query scrape failed.");
    } finally {
      setMapsLoading(false);
      ScrapeNative.stop(); // error path — just stop
    }
  };

  const importSelectedToLeads = () => {
    const toImport = mapsResults
      .filter(r => selectedPlaces.has(r.id))
      .map(r => ({ name: null, email: null, company: r.name, role: r.type, _website: r.website, _phone: r.phone, _address: r.address }));
    addLeads(toImport);
    setTab("extract");
  };

  // ── Scrape emails for selected Maps results — routes to chosen method ──
  const scrapeEmailsForSelected = () => {
    const websites = mapsResults
      .filter(r => selectedPlaces.has(r.id) && r.website)
      .map(r => r.website.trim())
      .filter(Boolean);

    if (!websites.length) {
      alert("No websites found for selected businesses. Make sure the selected businesses have a website listed.");
      return;
    }

    const joined = websites.join("\n");

    if (mapsScrapeMethod === "bulk") {
      // ── Bulk Scraper ────────────────────────────────────────────────────
      setBulkUrls(joined);
      setBulkInputMode("text");
      setTab("bulk");
      setTimeout(() => setBulkAutoRun(true), 300);

    } else if (mapsScrapeMethod === "smart") {
      // ── Smart Scraper ───────────────────────────────────────────────────
      setGeminiUrls(joined);
      setGeminiInputMode("text");
      setGeminiProvider(mapsSmartProvider); // apply chosen provider
      setTab("gemini");
      setTimeout(() => setGeminiAutoRun(true), 300);

    } else if (mapsScrapeMethod === "extract") {
      // ── Extract Leads (Snov/Apollo) ─────────────────────────────────────
      setUrlInput(joined);
      setMode("url");
      setTab("extract");
      setTimeout(() => setExtractAutoRun(true), 300);
    }
  };

  const exportMapsCSV = () => {
    const cols = ["name", "address", "phone", "website", "rating", "reviews", "type", "originalKeyword", "expandedKeyword", "sourceKeyword"];
    const rows = [cols.join(","), ...mapsResults.filter(r => selectedPlaces.has(r.id)).map(r => cols.map(c => `"${(r[c] ?? "").toString().replace(/"/g, '""')}"`).join(","))];
    triggerDownload(rows.join("\n"), "places.csv", "text/csv");
  };

  const exportMapsXLSX = () => {
    const data = mapsResults.filter(r => selectedPlaces.has(r.id)).map(r => ({
      Name: r.name ?? "",
      Address: r.address ?? "",
      Phone: r.phone ?? "",
      Website: r.website ?? "",
      Rating: r.rating ?? "",
      Reviews: r.reviews ?? "",
      Type: r.type ?? "",
      "Original Keyword":  r.originalKeyword  ?? "",
      "Expanded Keyword":  r.expandedKeyword  ?? "",
      "Source Keyword":    r.sourceKeyword    ?? "",
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 28 }, { wch: 36 }, { wch: 18 }, { wch: 32 }, { wch: 8 }, { wch: 10 }, { wch: 22 }, { wch: 22 }, { wch: 22 }, { wch: 22 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Places");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    triggerDownload(wbout, (() => { const filename = promptFileName(buildFileName("places", mapsQuery, mapsLocation)); return filename || buildFileName("places", mapsQuery, mapsLocation); })(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  };

  const togglePlace = (id) => setSelectedPlaces(prev => { const s = new Set(prev); s.has(id) ? s.delete(id) : s.add(id); return s; });
  const toggleAll = () => setSelectedPlaces(prev => prev.size === mapsResults.length ? new Set() : new Set(mapsResults.map(r => r.id)));

  // ── Bulk scraper ──────────────────────────────────────────────────────────
  const parseBulkUrls = (text) =>
    text.split(/[\n,]+/).map(u => normalizeUrl(u)).filter(Boolean);

  const runBulkScrape = async (retryFailed = false) => {
    console.group("[BulkScraper] runBulkScrape() called");
    console.log("retryFailed:", retryFailed);
    console.log("bulkInputMode:", bulkInputMode);
    console.log("bulkUrls (raw, first 200):", JSON.stringify(bulkUrls.slice(0, 200)));
    console.log("bulkFile:", bulkFile?.name ?? "null");

    let urls = [];
    if (retryFailed) {
      // Feature 2: only retry true errors — "empty" means page loaded fine, skip
      urls = bulkResults.filter(r => r.status === "error").map(r => r.url);
      console.log("Retry mode — failed URLs (excl. empty):", urls);
    } else if (bulkInputMode === "file" && bulkFile) {
      const fileName = bulkFile.name.toLowerCase();
      console.log("File mode — extension:", fileName.split(".").pop());
      if (fileName.endsWith(".xlsx")) {
        const arrayBuffer = await bulkFile.arrayBuffer();
        const wb = XLSX.read(arrayBuffer, { type: "array" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
        console.log("XLSX rows (first 5):", rows.slice(0, 5));
        urls = rows.flatMap(row => row.map(cell => normalizeUrl(String(cell ?? "")))).filter(Boolean);
      } else {
        const text = await readFileAsText(bulkFile);
        console.log("File text (first 300 chars):", text.slice(0, 300));
        urls = text.split(/[\n,\r]+/).map(u => normalizeUrl(u)).filter(Boolean);
      }
    } else {
      console.log("Text mode — raw lines (first 10):", bulkUrls.split(/[\n,\r]+/).slice(0, 10));
      urls = parseBulkUrls(bulkUrls);
    }

    // deduplicate URLs
    urls = [...new Set(urls)];
    console.log("Final URL list (%d):", urls.length, urls.slice(0, 10));
    if (!urls.length) {
      console.warn("[BulkScraper] No valid URLs — returning early. Check input format.");
      console.groupEnd();
      return;
    }
    console.log("[BulkScraper] Starting scrape:", urls.length, "URLs | concurrency:", bulkConcurrency, "| deepCrawl:", deepCrawl);
    console.groupEnd();

    const startTime = Date.now();
    setBulkStartTime(startTime);
    setBulkRunning(true);
    setBulkStopped(false);
    bulkStopRef.current = false;
    // ── Feature 3: Adaptive concurrency — reset to user-chosen value ────────
    adaptiveConcRef.current = bulkConcurrency;
    if (adaptiveRampTimer.current) clearTimeout(adaptiveRampTimer.current);
    setBulkSummaryOpen(false); // hide previous summary
    ScrapeNative.start(urls.length, "bulk"); // start foreground service — keeps app alive when minimized

    // Reset retry counts: fresh run clears all; retry-failed run only resets the failing ones
    if (!retryFailed) {
      bulkRetryMap.current = {};
      setBulkResults([]);
    } else {
      urls.forEach(url => { bulkRetryMap.current[url] = 0; });
      setBulkResults(prev =>
        prev.map(r => urls.includes(r.url) ? { ...r, status: "loading", error: null, emails: [], retries: 0 } : r)
      );
    }

    setBulkProgress({ done: 0, total: urls.length });
    setBulkEta(null);

    // init placeholder rows for fresh run
    if (!retryFailed) {
      setBulkResults(urls.map(url => ({ url, emails: [], status: "loading", error: null, retries: 0 })));
    }

    // ── Delegate to Cloudflare Worker job engine ──────────────────────────────
    // Scraping runs server-side — survives app minimize, screen lock, full close.
    // Frontend polls /jobs/:id/status every 4s and streams results back live.
    ScrapeNative.start(urls.length, "bulk"); // start foreground notification

    const jobId = await JobPoller.start({
      urls,
      provider: scraperApiKey ? "scraperapi" : zenRowsKey ? "zenrows" : "browser",

      onProgress: (status) => {
        setBulkProgress({ done: status.done, total: status.total });
        setBulkEta(calcEta(startTime, status.done, status.total));
        ScrapeNative.progress(
          status.done, status.total, status.emails,
          "bulk", status.success, status.failed
        );
        console.log(`[BulkScraper] Scraping loop tick — done:${status.done}/${status.total} emails:${status.emails} success:${status.success} failed:${status.failed}`);
      },

      onResult: (newResults) => {
        // Merge server results into bulkResults — same shape as on-device results
        setBulkResults(prev => {
          const map = new Map(prev.map(r => [r.url, r]));
          newResults.forEach(r => map.set(r.url, {
            url:    r.url,
            emails: r.emails  ?? [],
            status: r.status  ?? "ok",
            error:  r.error   ?? null,
            retries: 0,
          }));
          return [...map.values()];
        });
      },

      onComplete: (status) => {
        setBulkProgress(p => ({ ...p, done: status.total }));
        setBulkEta(null);
        setBulkRunning(false);
        setBulkJobId(null);
        setBulkSummaryOpen(true);
        ScrapeNative.complete(status.total, status.emails);
        console.log(`[BulkScraper] Complete — ${status.emails} emails from ${status.total} URLs`);
      },

      onError: (msg) => {
        console.error("[BulkScraper] Job error:", msg);
        setBulkRunning(false);
        setBulkJobId(null);
        ScrapeNative.stop();
      },
    });

    if (jobId) setBulkJobId(jobId);
    // ── END Worker job delegation ─────────────────────────────────────────────
    // onComplete / onError callbacks handle all teardown — nothing more to do here
  };

  const stopBulk = () => {
    bulkStopRef.current = true;
    if (bulkJobId) JobPoller.cancel(bulkJobId);
    setBulkJobId(null);
    ScrapeNative.stop();
  };

  const importBulkToLeads = () => {
    const newLeads = [];
    bulkResults.forEach(r => {
      r.emails.forEach(email => {
        const domain = (() => { try { return new URL(r.url).hostname.replace(/^www\./, ""); } catch { return ""; } })();
        newLeads.push({ name: null, email, company: domain || null, role: null });
      });
    });
    addLeads(newLeads);
    setTab("extract");
  };

  const exportBulkCSV = () => {
    const rows = [["url", "email"].join(",")];
    bulkResults.forEach(r => r.emails.forEach(e => rows.push(`"${r.url}","${e}"`)));
    triggerDownload(rows.join("\n"), "bulk-emails.csv", "text/csv");
  };

  const exportBulkXLSX = () => {
    const data = [];
    bulkResults.forEach(r => r.emails.forEach(e => data.push({ URL: r.url, Email: e })));
    const ws = XLSX.utils.json_to_sheet(data);
    ws["!cols"] = [{ wch: 48 }, { wch: 36 }];
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Bulk Emails");
    const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const bulkDomain = bulkUrls.trim().split(/[\n,]+/)[0]?.trim() || "";
    const bulkLabel = bulkDomain ? (() => { try { return new URL(bulkDomain.startsWith("http") ? bulkDomain : "https://" + bulkDomain).hostname.replace(/^www\./, "").split(".")[0]; } catch { return "bulk"; } })() : "bulk";
    triggerDownload(wbout, (() => { const filename = promptFileName(buildFileName("bulk-scraper", bulkLabel)); return filename || buildFileName("bulk-scraper", bulkLabel); })(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
  };

  const totalBulkEmails = bulkResults.reduce((s, r) => s + r.emails.length, 0);
  const bulkOk = bulkResults.filter(r => r.status === "ok").length;
  const bulkFailed = bulkResults.filter(r => r.status === "error").length;
  const bulkEmpty = bulkResults.filter(r => r.status === "empty").length;   // Feature 2/5
  const bulkBlocked = bulkResults.filter(r => r.errorClass === "blocked").length;  // Feature 4/5
  const bulkDns = bulkResults.filter(r => r.errorClass === "dns").length;          // Feature 4/5
  const bulkCrawling = bulkResults.filter(r => r.status === "crawling").length;

  const dupCount = leads.filter((l) => l._dup).length;

  const displayLeads = useMemo(() => {
    let list = showDupOnly ? leads.filter(l => l._dup) : leads;
    if (leadsTagFilter !== "all") list = list.filter(l => (l._tag || "none") === leadsTagFilter);
    if (leadsSearch.trim()) {
      const q = leadsSearch.trim().toLowerCase();
      list = list.filter(l =>
        (l.name || "").toLowerCase().includes(q) ||
        (l.email || "").toLowerCase().includes(q) ||
        (l.company || "").toLowerCase().includes(q) ||
        (l.role || "").toLowerCase().includes(q)
      );
    }
    return list;
  }, [leads, showDupOnly, leadsSearch, leadsTagFilter]);

  // ── verify all emails ─────────────────────────────────────────────────────
  const verifyAllEmails = async () => {
    const withEmail = leads.filter(l => l.email);
    if (!withEmail.length) return;
    setVerifying(true);
    setVerifyProgress(0);
    let done = 0;
    for (const lead of withEmail) {
      const status = await verifyEmail(lead.email);
      setLeads(prev => prev.map(l => l.id === lead.id ? { ...l, _verify: status } : l));
      done++;
      setVerifyProgress(Math.round((done / withEmail.length) * 100));
    }
    setVerifying(false);
  };

  // ── copy all emails ───────────────────────────────────────────────────────
  const copyAllEmails = () => {
    const emails = displayLeads.map(l => l.email).filter(Boolean).join("\n");
    navigator.clipboard.writeText(emails);
    setCopiedAll(true);
    setTimeout(() => setCopiedAll(false), 2000);
  };

  // ── tag a lead ────────────────────────────────────────────────────────────
  const setLeadTag = (id, tag) => {
    setLeads(prev => prev.map(l => l.id === id ? { ...l, _tag: l._tag === tag ? null : tag } : l));
  };

  const TAG_OPTIONS = [
    { value: "hot", label: "🔥 Hot", color: "#ff5f72" },
    { value: "contacted", label: "✉ Contacted", color: "#29b6f6" },
    { value: "ignore", label: "✕ Ignore", color: "#3a4a5e" },
  ];

  // ══════════════════════════════════════════════════════════════════════════
  return (
    <div style={{ fontFamily: "'DM Mono','Fira Mono',monospace", background: C.bg, minHeight: "100vh", color: C.text }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@300;400;500&display=swap');
        *{box-sizing:border-box;margin:0;padding:0;}
        textarea,input{outline:none;}
        textarea:focus,input:focus{border-color:${C.accent}55!important;}
        .hover-row:hover{background:#0f1820!important;}
        .cell-edit:hover{background:#111c28!important;cursor:text;}
        .icon-btn:hover{border-color:#2a3a55!important;color:${C.text}!important;}
        .run-btn:hover:not(:disabled){filter:brightness(1.1);transform:translateY(-1px);}
        .place-row:hover{background:#0f1622!important;}
        .bulk-row:hover{background:#0d1520!important;}
        @keyframes spin{to{transform:rotate(360deg);}}
        @keyframes fadeUp{from{opacity:0;transform:translateY(5px);}to{opacity:1;transform:translateY(0);}}
        @keyframes pulse{0%,100%{opacity:1;}50%{opacity:.4;}}
        .fade-row{animation:fadeUp 0.18s ease both;}
        .dup-badge{display:inline-flex;align-items:center;padding:1px 6px;background:${C.warn}18;border:1px solid ${C.warn}44;border-radius:3px;font-size:9px;color:${C.warn};letter-spacing:.08em;}
        .loading-dot{animation:pulse 1.2s ease-in-out infinite;}
        ::-webkit-scrollbar{width:5px;height:5px;}
        ::-webkit-scrollbar-track{background:#0a0e14;}
        ::-webkit-scrollbar-thumb{background:#1a2a3a;border-radius:3px;}
        a{color:${C.blue};text-decoration:none;}
        a:hover{text-decoration:underline;}
        @media(max-width:600px){
          .leads-table-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch;}
          .leads-table-wrap > div[style*="minWidth"]{min-width:600px;}
        }
      `}</style>

      {/* ── top bar ── */}
      <div style={{ borderBottom: `1px solid ${C.border}`, padding: "13px 22px", display: "flex", alignItems: "center", justifyContent: "space-between", background: C.surface }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ width: 8, height: 8, borderRadius: "50%", background: C.accent, boxShadow: `0 0 8px ${C.accent}`, display: "inline-block" }} />
          <span style={{ fontSize: 12, fontWeight: 500, letterSpacing: ".18em", color: "#fff", textTransform: "uppercase" }}>Lead Extractor</span>

        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          <button onClick={() => setCreditPanelOpen(v => !v)} title="API Credit Monitor" style={{
            padding: "4px 10px", background: creditPanelOpen ? `${C.blue}18` : "transparent",
            color: creditPanelOpen ? C.blue : C.dim,
            border: `1px solid ${creditPanelOpen ? C.blue + "55" : C.border}`,
            borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontSize: 10,
            letterSpacing: ".07em", transition: "all .15s",
          }}>⚡ credits</button>
          {dupCount > 0 && <SmBtn color={C.warn} onClick={() => setShowDupOnly(v => !v)}>{showDupOnly ? "show all" : `${dupCount} dup${dupCount !== 1 ? "s" : ""}`}</SmBtn>}
          {dupCount > 0 && <SmBtn color={C.accent} onClick={deduplicateAll}>deduplicate</SmBtn>}
          {leads.length > 0 && <SmBtn color={verifying ? C.warn : C.purple} onClick={verifyAllEmails} disabled={verifying}>{verifying ? `verifying ${verifyProgress}%` : "✓ verify"}</SmBtn>}
          {leads.length > 0 && <SmBtn color={copiedAll ? C.accent : C.dim} onClick={copyAllEmails}>{copiedAll ? "✓ copied" : "⧉ emails"}</SmBtn>}
          {leads.length > 0 && <SmBtn color={C.blue} onClick={() => setCrmOpen(v => !v)}>↑ HubSpot</SmBtn>}
          {leads.length > 0 && <SmBtn color={exportDone ? C.accent : C.dim} onClick={() => exportCSV()}>{exportDone ? "✓ saved" : "↓ CSV"}</SmBtn>}
          {leads.length > 0 && (
            <SmBtn color={exportXlsxDone ? C.accent : "#1ed98a"} onClick={() => exportXLSX()}>
              {exportXlsxDone ? "✓ saved" : "↓ XLSX"}
            </SmBtn>
          )}
          {leads.length > 0 && <SmBtn color={C.danger} onClick={() => { setLeads([]); setCrmResults(null); }}>clear</SmBtn>}
        </div>
      </div>

      {/* ── API Credit Monitor panel ── */}
      {creditPanelOpen && (
        <div style={{
          position: "fixed", top: 52, right: 16, zIndex: 9999,
          background: "#0d1117", border: `1px solid ${C.border}`,
          borderRadius: 8, padding: "14px 18px", minWidth: 230,
          boxShadow: "0 8px 32px rgba(0,0,0,.7)",
        }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
            <span style={{ fontSize: 10, letterSpacing: ".16em", color: C.muted, textTransform: "uppercase" }}>API Credit Monitor</span>
            <button onClick={() => { refreshCredits(); }} title="Refresh"
              style={{ background: "none", border: "none", color: C.dim, cursor: "pointer", fontSize: 13, padding: "0 2px" }}>↻</button>
          </div>
          {[
            {
              label: "ScraperAPI",
              value: apiCredits.scraperapi,
              type: "live",
              hasKey: !!scraperApiKey.trim(),
              suffix: "left",
              color: C.blue,
            },
            {
              label: "ZenRows",
              value: apiCredits.zenrows,
              type: "dashboard",
              hasKey: !!zenRowsKey.trim(),
              suffix: "used (est.)",
              color: "#eab308",
            },
            {
              label: "Snov.io",
              value: apiCredits.snovUsed,
              type: "local",
              hasKey: !!(extractSnovKey.trim() || hunterKey.trim()),
              suffix: "used (est.)",
              color: "#f97316",
            },
            {
              label: "Apollo.io",
              value: apiCredits.apolloUsed,
              type: "local",
              hasKey: !!(extractApolloKey.trim() || apolloKey.trim()),
              suffix: "used (est.)",
              color: C.purple,
            },
            {
              label: "Google Maps",
              value: apiCredits.mapsUsed,
              type: "local",
              hasKey: !!placesKey.trim(),
              suffix: "requests",
              color: C.accent,
            },
          ].map(row => (
            <div key={row.label} style={{
              display: "flex", alignItems: "center", justifyContent: "space-between",
              padding: "7px 0", borderBottom: `1px solid ${C.border}`,
            }}>
              <span style={{ fontSize: 11, color: C.dim }}>{row.label}</span>
              <span style={{ fontSize: 11, fontWeight: 500 }}>
                {!row.hasKey
                  ? <span style={{ color: "#1e2840" }}>no key</span>
                  : row.value === "no-api"
                  ? <span style={{ fontSize: 9, color: "#eab308" }}>see dashboard ↗</span>
                  : row.value === null
                  ? <span style={{ color: C.dim }}>loading…</span>
                  : row.value === "unavailable"
                  ? <span style={{ color: C.danger }}>unavailable</span>
                  : <span style={{ color: row.color }}>
                      {typeof row.value === "number" ? row.value.toLocaleString() : row.value}
                      <span style={{ fontSize: 9, color: C.dim, marginLeft: 4 }}>{row.suffix}</span>
                    </span>
                }
              </span>
            </div>
          ))}
          <div style={{ marginTop: 10, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <span style={{ fontSize: 9, color: "#1e2840" }}>live = API · est. = local count</span>
            <button onClick={() => {
              saveCreditUsage({ snovUsed: 0, apolloUsed: 0, mapsUsed: 0 });
              setApiCredits(prev => ({ ...prev, snovUsed: 0, apolloUsed: 0, mapsUsed: 0 }));
            }} style={{ fontSize: 9, color: C.danger, background: "none", border: `1px solid ${C.danger}33`, borderRadius: 3, padding: "2px 7px", cursor: "pointer", fontFamily: "inherit" }}>
              reset counts
            </button>
          </div>
        </div>
      )}

      {/* ── main tabs ── */}
      <div style={{ borderBottom: `1px solid ${C.border}`, padding: "0 22px", background: C.surface, display: "flex", gap: 0 }}>
        {[
          { id: "extract", label: "🔗  Extract Leads" },
          { id: "bulk", label: "⚡  Bulk URL Scraper" },
          { id: "gemini", label: "✦  Smart Scraper" },
          { id: "maps", label: "🗺  Maps Scraper" },
        ].map(t => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{
            padding: "11px 18px", background: "transparent", border: "none", borderBottom: `2px solid ${tab === t.id ? C.accent : "transparent"}`,
            color: tab === t.id ? C.accent : C.dim, cursor: "pointer", fontFamily: "inherit", fontSize: 11,
            fontWeight: tab === t.id ? 500 : 400, letterSpacing: ".07em", transition: "all .15s",
          }}>{t.label}</button>
        ))}
        {leads.length > 0 && (
          <span style={{ marginLeft: "auto", fontSize: 11, color: C.dim, alignSelf: "center" }}>
            {leads.length} lead{leads.length !== 1 ? "s" : ""}
            {dupCount > 0 && <span style={{ color: C.warn }}> · {dupCount} dup{dupCount !== 1 ? "s" : ""}</span>}
          </span>
        )}
      </div>

      <div style={{ maxWidth: 960, margin: "0 auto", padding: "22px 18px" }}>

        {/* ════════════════ EXTRACT TAB ════════════════ */}
        {tab === "extract" && (<>

          {/* Input mode tabs */}
          <div style={{ display: "flex", gap: 3, marginBottom: 16 }}>
            {[{ id: "url", label: "🌐  URL" }, { id: "text", label: "✏  Text" }, { id: "file", label: "📄  File" }].map(t => (
              <button key={t.id} onClick={() => { setMode(t.id); setError(""); setFile(null); }} style={{
                padding: "6px 16px", background: mode === t.id ? C.accent : "transparent",
                color: mode === t.id ? "#000" : C.muted, border: `1px solid ${mode === t.id ? C.accent : C.border}`,
                borderRadius: 4, cursor: "pointer", fontSize: 11, fontFamily: "inherit",
                fontWeight: mode === t.id ? 500 : 400, letterSpacing: ".06em", transition: "all .15s",
              }}>{t.label}</button>
            ))}
          </div>

          {/* Input area */}
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            <Label>{mode === "url" ? "URLs to scan (one per line)" : mode === "text" ? "Paste text containing emails" : "Upload file (CSV, TXT)"}</Label>

            {mode === "url" && (
              <>
                <textarea value={urlInput} onChange={e => setUrlInput(e.target.value)}
                  placeholder={"https://company.com\nhttps://lawfirm.com/team\nanotherdomain.org"}
                  rows={5}
                  style={{ width: "100%", background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "10px 14px", resize: "vertical", boxSizing: "border-box", lineHeight: 1.7 }} />
                <div style={{ marginTop: 6, fontSize: 10, color: C.dim }}>
                  {urlInput.trim().split(/[\n,]+/).filter(u => u.trim()).length} URL{urlInput.trim().split(/[\n,]+/).filter(u => u.trim()).length !== 1 ? "s" : ""} detected
                </div>
              </>
            )}

            {mode === "text" && (
              <textarea value={textInput} onChange={e => setTextInput(e.target.value)}
                placeholder={"john@acme.com\ninfo@lawfirm.com\ncontact@company.org"}
                style={{ width: "100%", height: 110, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "10px 14px", resize: "none", lineHeight: 1.75 }} />
            )}

            {mode === "file" && (
              <label htmlFor="main-file-input" style={{ display: "block", border: `2px dashed ${file ? C.accent : C.border}`, borderRadius: 6, padding: "26px 20px", textAlign: "center", cursor: "pointer", background: file ? `${C.accent}07` : "transparent", transition: "all .2s" }}>
                <div style={{ fontSize: 26, marginBottom: 7 }}>📂</div>
                <div style={{ fontSize: 12, color: file ? C.accent : C.muted }}>
                  {file ? `✓  ${file.name}  (${(file.size / 1024).toFixed(1)} KB)` : "Tap to choose — CSV, TXT"}
                </div>
                {file && <div onClick={e => { e.preventDefault(); e.stopPropagation(); setFile(null); }} style={{ marginTop: 8, fontSize: 10, color: C.danger, cursor: "pointer" }}>✕ remove</div>}
              </label>
            )}
            <input id="main-file-input" key={mode} ref={fileRef} type="file" accept=".csv,.txt"
              style={{ position: "absolute", opacity: 0, width: 1, height: 1, pointerEvents: "none" }}
              onClick={e => { e.target.value = ""; }} onChange={e => { setFile(e.target.files[0] || null); setError(""); }} />
          </div>

          {/* API Keys panel — only shown in URL mode */}
          {mode === "url" && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                <Label>API Keys</Label>
                <span style={{ fontSize: 9, color: C.dim, letterSpacing: ".1em" }}>SNOV PRIMARY · APOLLO ENRICHMENT</span>
              </div>

              {/* Snov.io */}
              <div style={{ marginBottom: 14, paddingBottom: 14, borderBottom: `1px solid ${C.border}` }}>
                <div style={{ fontSize: 9, color: "#f97316", letterSpacing: ".1em", marginBottom: 6 }}>SNOV.IO — PRIMARY EMAIL SOURCE (up to 50/run)</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                  <input type="password" value={extractSnovKey} onChange={e => { setExtractSnovKey(e.target.value); setExtractSnovSaved(false); }}
                    placeholder="Snov.io Client ID…"
                    style={{ background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "8px 12px" }} />
                  <div style={{ display: "flex", gap: 8 }}>
                    <input type="password" value={extractSnovSecret} onChange={e => { setExtractSnovSecret(e.target.value); setExtractSnovSaved(false); }}
                      placeholder="Snov.io Client Secret…"
                      style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "8px 12px" }} />
                    <button onClick={() => { SecureStorage.set("extractSnovKey", extractSnovKey); SecureStorage.set("extractSnovSecret", extractSnovSecret); setExtractSnovSaved(true); }}
                      disabled={!extractSnovKey.trim() || !extractSnovSecret.trim()}
                      style={{ padding: "8px 16px", background: (extractSnovKey.trim() && extractSnovSecret.trim()) ? C.accent : C.surface2, color: (extractSnovKey.trim() && extractSnovSecret.trim()) ? "#000" : C.muted, border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontSize: 11, fontWeight: 500, whiteSpace: "nowrap" }}>
                      {extractSnovSaved ? "✓ Saved" : "Save"}
                    </button>
                  </div>
                </div>
                <div style={{ fontSize: 10, color: C.dim, marginTop: 5 }}>Free at <a href="https://snov.io" target="_blank" rel="noreferrer" style={{ color: C.blue }}>snov.io</a> — 50 credits/month. Called for every domain as the primary email source.</div>
              </div>

              {/* Apollo.io */}
              <div>
                <div style={{ fontSize: 9, color: C.dim, letterSpacing: ".1em", marginBottom: 6 }}>APOLLO.IO — COMPANY ENRICHMENT (up to 50/run)</div>
                <div style={{ display: "flex", gap: 8 }}>
                  <input type="password" value={extractApolloKey} onChange={e => { setExtractApolloKey(e.target.value); setExtractApolloSaved(false); }}
                    placeholder="Apollo.io API key…"
                    style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "8px 12px" }} />
                  <button onClick={() => { SecureStorage.set("extractApolloKey", extractApolloKey); setExtractApolloSaved(true); }}
                    disabled={!extractApolloKey.trim()}
                    style={{ padding: "8px 16px", background: extractApolloKey.trim() ? C.purple : C.surface2, color: extractApolloKey.trim() ? "#000" : C.muted, border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontSize: 11, fontWeight: 500, whiteSpace: "nowrap" }}>
                    {extractApolloSaved ? "✓ Saved" : "Save"}
                  </button>
                </div>
                <div style={{ fontSize: 10, color: C.dim, marginTop: 5 }}>Free at <a href="https://app.apollo.io" target="_blank" rel="noreferrer" style={{ color: C.blue }}>apollo.io</a> — 50 credits/month. Adds company name, industry, size.</div>
              </div>
            </div>
          )}

          {error && <div style={{ marginBottom: 10, fontSize: 11, color: "#ff7070", padding: "7px 12px", background: "#ff606012", borderRadius: 4, border: "1px solid #ff606033" }}>⚠  {error}</div>}

          {/* Run button + progress */}
          <div style={{ display: "flex", alignItems: "center", gap: 12, marginBottom: 16 }}>
            <button className="run-btn" onClick={extract} disabled={!canRun} style={{
              padding: "9px 26px", background: canRun ? C.accent : C.surface2, color: canRun ? "#000" : C.muted,
              border: canRun ? "none" : `1px solid ${C.border}`, borderRadius: 4, cursor: canRun ? "pointer" : "not-allowed",
              fontFamily: "inherit", fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase",
              transition: "all .15s", display: "flex", alignItems: "center", gap: 8,
            }}>
              {extractRunning ? <><Spin />Processing…</> : loading ? <><Spin />{loadingMsg || "Working…"}</> : "▶  Extract Leads"}
            </button>
            {extractRunning && extractProgress.total > 0 && (
              <span style={{ fontSize: 11, color: C.dim }}>
                Processing <span style={{ color: C.accent }}>{extractProgress.done}</span> / {extractProgress.total}
              </span>
            )}
          </div>

          {/* Per-URL results — table layout matching Smart Scraper */}
          {extractResults.length > 0 && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, overflow: "hidden", marginBottom: 16 }}>
              {/* Header row */}
              <div style={{
                display: "grid", gridTemplateColumns: "1.4fr 2.2fr 1.6fr 80px 90px",
                padding: "8px 14px", borderBottom: `1px solid ${C.border}`,
                fontSize: 9, letterSpacing: ".13em", color: C.muted, textTransform: "uppercase",
              }}>
                <span>URL</span>
                <span>Emails Found</span>
                <span>Enrichment</span>
                <span>Status</span>
                <span>Provider</span>
              </div>
              {extractResults.map((row, i) => (
                <div key={i} className="bulk-row fade-row" style={{
                  display: "grid", gridTemplateColumns: "1.4fr 2.2fr 1.6fr 80px 90px",
                  padding: "9px 14px", borderBottom: i < extractResults.length - 1 ? `1px solid ${C.border}` : "none",
                  alignItems: "center", fontSize: 11, background: i % 2 ? "#0a0f16" : "transparent",
                }}>
                  {/* URL */}
                  <span style={{ fontSize: 10, color: C.dim, wordBreak: "break-all", paddingRight: 8 }}>
                    {row.domain}
                  </span>
                  {/* Emails Found */}
                  <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                    {row.status === "processing" && <span className="loading-dot" style={{ fontSize: 10, color: C.dim }}>scanning…</span>}
                    {row.status === "error" && <span style={{ fontSize: 10, color: C.danger }}>⚠ {row.error?.slice(0, 40)}</span>}
                    {row.status === "no-email" && <span style={{ fontSize: 10, color: "#1e2840" }}>no emails found</span>}
                    {row.emails?.map((e, j) => (
                      <span key={j} style={{ fontSize: 10, color: C.accent, background: `${C.accent}12`, padding: "1px 7px", borderRadius: 3, border: `1px solid ${C.accent}22` }}>{e}</span>
                    ))}
                  </div>
                  {/* Enrichment */}
                  <div style={{ fontSize: 9, color: C.dim, lineHeight: 1.7 }}>
                    {row.enrichment ? (<>
                      {row.enrichment.company && <div style={{ color: C.text, fontWeight: 500 }}>{row.enrichment.company}</div>}
                      {row.enrichment.industry && <div>{row.enrichment.industry}</div>}
                      {row.enrichment.size && <div>{row.enrichment.size.toLocaleString()} employees</div>}
                      {row.enrichment.city && <div>{row.enrichment.city}{row.enrichment.country ? `, ${row.enrichment.country}` : ""}</div>}
                      {row.contacts?.some(c => c.name || c.title) && (
                        <div style={{ marginTop: 4, paddingTop: 4, borderTop: `1px solid #1e2840` }}>
                          {row.contacts.filter(c => c.name || c.title).slice(0, 3).map((c, ci) => (
                            <div key={ci} style={{ marginBottom: 2 }}>
                              {c.name && <span style={{ color: "#f97316", fontWeight: 500 }}>{c.name}</span>}
                              {c.title && <span style={{ color: C.dim }}> · {c.title}</span>}
                            </div>
                          ))}
                        </div>
                      )}
                    </>) : <span style={{ color: "#1e2840" }}>—</span>}
                  </div>
                  {/* Status badge */}
                  <span style={{ fontSize: 10, color:
                    row.status === "success" ? C.accent :
                    row.status === "error" ? C.danger :
                    row.status === "no-email" ? "#1e2840" : C.dim
                  }}>
                    {row.status === "processing" ? "…" :
                     row.status === "success" ? `✓ ${row.emails.length}` :
                     row.status === "error" ? "✕ fail" : "— none"}
                  </span>
                  {/* Provider */}
                  <span style={{ fontSize: 9, color: "#f97316", display: "flex", alignItems: "center", gap: 3, letterSpacing: ".04em" }}>
                    {row.status === "processing" ? <span style={{ color: C.dim }}>…</span> : <>🟠 snov.io</>}
                  </span>
                </div>
              ))}
            </div>
          )}

        </>)}

                {tab === "bulk" && (<>

          <div style={{ background: `${C.accent}0a`, border: `1px solid ${C.accent}33`, borderRadius: 6, padding: "12px 18px", marginBottom: 16, display: "flex", alignItems: "flex-start", gap: 10 }}>
            <span style={{ fontSize: 18, marginTop: 1 }}>⚡</span>
            <div>
              <div style={{ fontSize: 11, color: C.accent, fontWeight: 500, letterSpacing: ".06em", marginBottom: 4 }}>100% Free — No API key, No limits</div>
              <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.7 }}>
                Fetches each website via a free CORS proxy and extracts all email addresses using pattern matching. Works on 300+ URLs. No AI used — instant, unlimited, free.<br />
                <span style={{ color: C.muted }}>Extracts: emails only. For names/roles/companies use the Extract Leads tab.</span>
              </div>
            </div>
          </div>

          <div style={{ display: "flex", gap: 3, marginBottom: 12 }}>
            {[{ id: "text", label: "✏  Paste URLs" }, { id: "file", label: "📄  Upload File" }].map(t => (
              <button key={t.id} onClick={() => setBulkInputMode(t.id)} style={{
                padding: "6px 16px", background: bulkInputMode === t.id ? C.accent : "transparent",
                color: bulkInputMode === t.id ? "#000" : C.muted, border: `1px solid ${bulkInputMode === t.id ? C.accent : C.border}`,
                borderRadius: 4, cursor: "pointer", fontSize: 11, fontFamily: "inherit",
                fontWeight: bulkInputMode === t.id ? 500 : 400, letterSpacing: ".06em", transition: "all .15s",
              }}>{t.label}</button>
            ))}
          </div>

          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            {bulkInputMode === "text" ? (<>
              <Label>Paste URLs (one per line, or comma-separated)</Label>
              <textarea value={bulkUrls} onChange={e => setBulkUrls(e.target.value)}
                placeholder={"https://company1.com/contact\nhttps://company2.com/about\nhttps://company3.com/team"}
                style={{ width: "100%", height: 130, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 11, padding: "10px 14px", resize: "vertical", lineHeight: 1.8, transition: "border-color .15s" }} />
              <div style={{ fontSize: 10, color: C.dim, marginTop: 6 }}>
                {parseBulkUrls(bulkUrls).length > 0 && <span style={{ color: C.accent }}>{parseBulkUrls(bulkUrls).length} URLs detected</span>}
              </div>
            </>) : (<>
              <Label>Upload CSV, XLSX or TXT with one URL per line</Label>
              <label htmlFor="bulk-file-input" style={{ display: "block", border: `2px dashed ${bulkFile ? C.accent : C.border}`, borderRadius: 6, padding: "26px 20px", textAlign: "center", cursor: "pointer", background: bulkFile ? `${C.accent}07` : "transparent", transition: "all .2s" }}>
                <div style={{ fontSize: 26, marginBottom: 7 }}>📂</div>
                <div style={{ fontSize: 12, color: bulkFile ? C.accent : C.muted }}>{bulkFile ? `✓  ${bulkFile.name}  (${(bulkFile.size / 1024).toFixed(1)} KB)` : "Tap to choose — CSV, XLSX or TXT"}</div>
              </label>
              <input
                id="bulk-file-input"
                ref={bulkFileRef}
                type="file"
                accept=".csv,.xlsx,.txt"
                style={{ position: "absolute", opacity: 0, width: 1, height: 1, pointerEvents: "none" }}
                onClick={e => { e.target.value = ""; }}
                onChange={e => setBulkFile(e.target.files[0] || null)}
              />
            </>)}

            {/* Deep Crawl toggle */}
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 14, marginBottom: 6, padding: "10px 14px", background: deepCrawl ? `${C.purple}12` : C.surface2, border: `1px solid ${deepCrawl ? C.purple + "44" : C.border}`, borderRadius: 6 }}>
              <input type="checkbox" id="deep-crawl-toggle" checked={deepCrawl} onChange={e => setDeepCrawl(e.target.checked)}
                style={{ accentColor: C.purple, cursor: "pointer", width: 15, height: 15 }} />
              <label htmlFor="deep-crawl-toggle" style={{ cursor: "pointer", flex: 1 }}>
                <div style={{ fontSize: 11, color: deepCrawl ? C.purple : C.dim, fontWeight: 500 }}>🕷 Deep Crawl Mode</div>
                <div style={{ fontSize: 10, color: C.muted, marginTop: 2 }}>Follows internal links (agent profiles, team pages) to find hidden emails. Slower but finds much more.</div>
              </label>
              {deepCrawl && (
                <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 2 }}>
                  <span style={{ fontSize: 9, color: C.muted, letterSpacing: ".06em" }}>MAX PAGES</span>
                  <select value={crawlDepth} onChange={e => setCrawlDepth(Number(e.target.value))}
                    style={{ background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>
                    <option value={10}>10</option>
                    <option value={20}>20</option>
                    <option value={50}>50</option>
                    <option value={100}>100</option>
                  </select>
                </div>
              )}
            </div>

            {/* Concurrency control */}
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8, padding: "8px 14px", background: C.surface2, border: `1px solid ${C.border}`, borderRadius: 6 }}>
              <span style={{ fontSize: 10, color: C.dim, flex: 1 }}>⚡ <strong style={{ color: C.text }}>Parallel workers</strong> — scrape multiple URLs at once</span>
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 2 }}>
                <span style={{ fontSize: 9, color: C.muted, letterSpacing: ".06em" }}>CONCURRENCY</span>
                <select value={bulkConcurrency} onChange={e => setBulkConcurrency(Number(e.target.value))}
                  style={{ background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>
                  <option value={1}>1 (recommended)</option>
                  <option value={2}>2 (moderate)</option>
                  <option value={3}>3 (fast)</option>
                  <option value={5}>5 (may hit limits)</option>
                </select>
              </div>
            </div>

            <div style={{ display: "flex", gap: 8, marginTop: 8, alignItems: "center" }}>
              {!bulkRunning ? (
                <button
                  onClick={() => {
                    const ready = bulkInputMode === "text" ? parseBulkUrls(bulkUrls).length > 0 : !!bulkFile;
                    console.log("[BulkScraper] Button clicked | ready:", ready, "| mode:", bulkInputMode, "| urls:", bulkUrls.trim().slice(0, 80));
                    if (!ready) return;
                    runBulkScrape();
                  }}
                  style={{
                    padding: "9px 26px",
                    background: (bulkInputMode === "text" ? parseBulkUrls(bulkUrls).length > 0 : !!bulkFile) ? C.accent : C.surface2,
                    color: (bulkInputMode === "text" ? parseBulkUrls(bulkUrls).length > 0 : !!bulkFile) ? "#000" : C.muted,
                    border: "none", borderRadius: 4,
                    cursor: (bulkInputMode === "text" ? parseBulkUrls(bulkUrls).length > 0 : !!bulkFile) ? "pointer" : "not-allowed",
                    fontFamily: "inherit", fontWeight: 500, fontSize: 11,
                    letterSpacing: ".1em", textTransform: "uppercase",
                    transition: "all .15s", display: "flex", alignItems: "center", gap: 8,
                  }}>▶  Start Bulk Scrape</button>
              ) : (
                <button onClick={stopBulk} style={{ padding: "9px 22px", background: C.danger, color: "#fff", border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase" }}>
                  ◼  Stop
                </button>
              )}

              {bulkRunning && (
                <div style={{ fontSize: 11, color: C.dim, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                  <Spin />
                  {bulkBatch.total > 1 && (
                    <span style={{ color: C.purple, fontWeight: 500 }}>
                      Batch {bulkBatch.current}/{bulkBatch.total}
                    </span>
                  )}
                  <span>{bulkProgress.done} / {bulkProgress.total} URLs</span>
                  <span style={{ color: C.accent }}>{totalBulkEmails} emails found</span>
                  {/* Feature 3: adaptive concurrency indicator */}
                  {adaptiveConcRef.current < bulkConcurrency && (
                    <span style={{ color: C.warn, fontSize: 10 }}>⚡ throttled to 1</span>
                  )}
                </div>
              )}
            </div>
          </div>

          {(bulkRunning || bulkResults.length > 0) && bulkProgress.total > 0 && (
            <div style={{ marginBottom: 14 }}>
              <div style={{ height: 3, background: C.surface2, borderRadius: 2, overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${(bulkProgress.done / bulkProgress.total) * 100}%`, background: C.accent, transition: "width .3s", borderRadius: 2 }} />
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6, fontSize: 10, color: C.dim }}>
                <span>{bulkStopped ? "⏹ Stopped" : bulkRunning ? "⚡ Running…" : "✓ Complete"}</span>
                {bulkCrawling > 0 && <span style={{ color: C.purple }}>🕷 {bulkCrawling} crawling</span>}
                <span style={{ color: C.accent }}>{bulkOk} ok</span>
                <span style={{ color: C.warn }}>{bulkEmpty} empty</span>
                {bulkBlocked > 0 && <span style={{ color: "#ff5f72" }}>🔴 {bulkBlocked} blocked</span>}
                <span style={{ color: C.danger }}>{bulkFailed} failed</span>
                {bulkEta && <span style={{ color: C.dim }}>{bulkEta}</span>}
                <span style={{ color: C.text, fontWeight: 500 }}>{totalBulkEmails} total emails</span>
              </div>
            </div>
          )}

          {/* ── Feature 5: End-of-run summary panel ───────────────────── */}
          {bulkResults.length > 0 && !bulkRunning && bulkSummaryOpen && (() => {
            const copyUrls = (filterFn) => {
              const urls = bulkResults.filter(filterFn).map(r => r.url).join("\n");
              navigator.clipboard.writeText(urls).catch(() => {});
            };
            return (
              <div style={{ marginBottom: 14, background: C.surface2, border: `1px solid ${C.border}`, borderRadius: 6, overflow: "hidden" }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 14px", borderBottom: `1px solid ${C.border}` }}>
                  <span style={{ fontSize: 10, letterSpacing: ".14em", color: C.muted, textTransform: "uppercase" }}>Run Summary</span>
                  <button onClick={() => setBulkSummaryOpen(false)} style={{ background: "none", border: "none", color: C.muted, cursor: "pointer", fontSize: 13, lineHeight: 1 }}>✕</button>
                </div>
                <div style={{ padding: "12px 14px", display: "flex", flexWrap: "wrap", gap: 10 }}>
                  <div style={{ display: "flex", gap: 8, flexWrap: "wrap", flex: 1 }}>
                    <span style={{ fontSize: 11, color: C.accent }}>✓ {totalBulkEmails} emails found</span>
                    <span style={{ fontSize: 11, color: C.warn }}>⚠ {bulkEmpty} loaded, 0 emails</span>
                    {bulkBlocked > 0 && <span style={{ fontSize: 11, color: "#ff5f72" }}>🔴 {bulkBlocked} blocked</span>}
                    {bulkDns > 0 && <span style={{ fontSize: 11, color: "#6b7280" }}>⚫ {bulkDns} DNS/dead</span>}
                    {bulkFailed > 0 && <span style={{ fontSize: 11, color: C.danger }}>✕ {bulkFailed} errors</span>}
                  </div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {bulkEmpty > 0 && (
                      <button onClick={() => copyUrls(r => r.status === "empty")} style={{ fontSize: 9, padding: "3px 9px", background: "transparent", color: C.warn, border: `1px solid ${C.warn}33`, borderRadius: 3, cursor: "pointer", fontFamily: "inherit", letterSpacing: ".04em" }}>
                        ⧉ Copy empty URLs
                      </button>
                    )}
                    {bulkBlocked > 0 && (
                      <button onClick={() => copyUrls(r => r.errorClass === "blocked")} style={{ fontSize: 9, padding: "3px 9px", background: "transparent", color: "#ff5f72", border: "1px solid #ff5f7233", borderRadius: 3, cursor: "pointer", fontFamily: "inherit", letterSpacing: ".04em" }}>
                        ⧉ Copy blocked URLs
                      </button>
                    )}
                    {bulkDns > 0 && (
                      <button onClick={() => copyUrls(r => r.errorClass === "dns")} style={{ fontSize: 9, padding: "3px 9px", background: "transparent", color: "#6b7280", border: "1px solid #6b728033", borderRadius: 3, cursor: "pointer", fontFamily: "inherit", letterSpacing: ".04em" }}>
                        ⧉ Copy dead URLs
                      </button>
                    )}
                  </div>
                </div>
                <div style={{ padding: "6px 14px 10px", fontSize: 10, color: C.muted }}>
                  💡 Copy <strong style={{ color: C.warn }}>empty</strong> or <strong style={{ color: "#ff5f72" }}>blocked</strong> URLs → paste into <strong style={{ color: C.text }}>Smart Scraper</strong> for deeper extraction.
                  Copy <strong style={{ color: "#6b7280" }}>dead URLs</strong> → remove from your list.
                </div>
              </div>
            );
          })()}

          {bulkResults.length > 0 && !bulkRunning && (
            <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
              <button onClick={() => importBulkToLeads()} disabled={!totalBulkEmails} style={{ padding: "7px 18px", background: totalBulkEmails ? C.accent : C.surface2, color: totalBulkEmails ? "#000" : C.muted, border: "none", borderRadius: 4, cursor: totalBulkEmails ? "pointer" : "not-allowed", fontFamily: "inherit", fontWeight: 500, fontSize: 11, letterSpacing: ".07em" }}>
                → Import {totalBulkEmails} emails to Leads
              </button>
              <button onClick={exportBulkCSV} disabled={!totalBulkEmails} style={{ padding: "7px 18px", background: "transparent", color: totalBulkEmails ? C.blue : C.muted, border: `1px solid ${totalBulkEmails ? C.blue + "44" : C.border}`, borderRadius: 4, cursor: totalBulkEmails ? "pointer" : "not-allowed", fontFamily: "inherit", fontSize: 11, letterSpacing: ".07em" }}>
                ↓ Export CSV
              </button>
              <button onClick={exportBulkXLSX} disabled={!totalBulkEmails} style={{ padding: "7px 18px", background: "transparent", color: totalBulkEmails ? "#1ed98a" : C.muted, border: `1px solid ${totalBulkEmails ? "#1ed98a44" : C.border}`, borderRadius: 4, cursor: totalBulkEmails ? "pointer" : "not-allowed", fontFamily: "inherit", fontSize: 11, letterSpacing: ".07em" }}>
                ↓ Export XLSX
              </button>
              {bulkFailed > 0 && (
                <button onClick={() => runBulkScrape(true)} style={{ padding: "7px 18px", background: "transparent", color: C.warn, border: `1px solid ${C.warn}44`, borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontSize: 11, letterSpacing: ".07em" }}>
                  ↺ Retry {bulkFailed} errors
                </button>
              )}
              {!bulkSummaryOpen && bulkResults.length > 0 && (
                <button onClick={() => setBulkSummaryOpen(true)} style={{ padding: "7px 14px", background: "transparent", color: C.blue, border: `1px solid ${C.blue}33`, borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontSize: 11, letterSpacing: ".07em" }}>
                  ≡ Summary
                </button>
              )}
              <button onClick={() => { setBulkResults([]); setBulkProgress({ done: 0, total: 0 }); }} style={{ padding: "7px 14px", background: "transparent", color: C.danger, border: `1px solid ${C.danger}33`, borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontSize: 11 }}>
                clear
              </button>
            </div>
          )}

          {/* ── Screenshot debug modal ─────────────────────────────────── */}
          {screenshotModal && (
            <div
              onClick={() => setScreenshotModal(null)}
              style={{
                position: "fixed", inset: 0, zIndex: 9999,
                background: "rgba(0,0,0,0.85)",
                display: "flex", alignItems: "center", justifyContent: "center",
                padding: 16,
              }}
            >
              <div
                onClick={e => e.stopPropagation()}
                style={{
                  position: "relative", background: "#0d1117", borderRadius: 8,
                  border: "1px solid #1e2840", maxWidth: "92vw", maxHeight: "88vh",
                  display: "flex", flexDirection: "column", overflow: "hidden",
                }}
              >
                {/* Header */}
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "10px 14px", borderBottom: "1px solid #1e2840" }}>
                  <span style={{ fontSize: 11, color: "#7a8aa0", letterSpacing: ".1em", textTransform: "uppercase" }}>Debug Screenshot</span>
                  <button
                    onClick={() => setScreenshotModal(null)}
                    style={{ background: "transparent", border: "none", color: "#7a8aa0", fontSize: 18, cursor: "pointer", lineHeight: 1, padding: "0 4px" }}
                  >✕</button>
                </div>
                {/* Image */}
                <div style={{ overflowY: "auto", padding: 12 }}>
                  <img
                    src={`data:image/png;base64,${screenshotModal}`}
                    alt="Page screenshot"
                    style={{ width: "100%", borderRadius: 4, display: "block" }}
                  />
                </div>
                {/* Footer hint */}
                <div style={{ padding: "8px 14px", borderTop: "1px solid #1e2840", fontSize: 10, color: "#3a4a60", textAlign: "center" }}>
                  Tap outside or ✕ to close
                </div>
              </div>
            </div>
          )}

          {bulkResults.length > 0 && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, overflow: "hidden" }}>
              <div style={{ display: "grid", gridTemplateColumns: "2fr 3fr 60px 80px", padding: "8px 14px", borderBottom: `1px solid ${C.border}`, fontSize: 9, letterSpacing: ".18em", color: C.muted, textTransform: "uppercase" }}>
                <span>URL</span><span>Emails Found</span><span title="✓ business domain  ⚠ personal/mismatch">Match</span><span>Status</span>
              </div>
              <div style={{ maxHeight: 460, overflowY: "auto" }}>
                {bulkResults.map((row, i) => {
                  const [expanded, setExpanded] = [row._expanded, (v) => setBulkResults(prev => prev.map((r, j) => j === i ? { ...r, _expanded: v } : r))];
                  const showAll = expanded || row.emails.length <= 4;
                  const visibleEmails = showAll ? row.emails : row.emails.slice(0, 3);
                  return (
                    <div key={i} className="bulk-row fade-row" style={{
                      display: "grid", gridTemplateColumns: "2fr 3fr 60px 80px",
                      padding: "9px 14px", borderBottom: i < bulkResults.length - 1 ? `1px solid ${C.border}` : "none",
                      alignItems: "start", fontSize: 11,
                      background: i % 2 ? "#0a0f16" : "transparent",
                    }}>
                      <span style={{ color: C.dim, fontSize: 10, wordBreak: "break-all", paddingRight: 10, paddingTop: 2 }}>
                        {(() => { try { return new URL(row.url).hostname; } catch { return row.url; } })()}
                      </span>
                      <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                        {row.status === "loading" && <span className="loading-dot" style={{ fontSize: 10, color: C.dim }}>scanning…</span>}
                        {row.status === "crawling" && <span className="loading-dot" style={{ fontSize: 10, color: C.purple }}>🕷 {row.crawlMsg || "deep crawling…"}</span>}
                        {/* Feature 4: classified error badges */}
                        {row.status === "error" && (() => {
                          const meta = BULK_ERROR_META[row.errorClass] ?? BULK_ERROR_META.error;
                          return (
                            <span title={meta.tip} style={{ fontSize: 10, color: meta.color, display: "flex", alignItems: "center", gap: 4 }}>
                              {meta.icon} {meta.label}
                              {row.error && <span style={{ color: C.muted, fontSize: 9 }}> — {row.error.slice(0, 48)}</span>}
                            </span>
                          );
                        })()}
                        {(row.status === "none" || row.status === "empty") && (
                          <span style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            <span style={{ fontSize: 10, color: "#1e2840" }}>no emails found</span>
                            {row.screenshot && (
                              <button
                                onClick={() => setScreenshotModal(row.screenshot)}
                                style={{
                                  fontSize: 9, color: "#4a90d9", background: "transparent",
                                  border: "1px solid #4a90d922", borderRadius: 3,
                                  padding: "1px 7px", cursor: "pointer", fontFamily: "inherit",
                                  letterSpacing: ".05em",
                                }}
                              >📷 View Screenshot</button>
                            )}
                          </span>
                        )}
                        {/* Feature 6: domain-match scored emails */}
                        {visibleEmails.map((e, j) => {
                          const matchScore = scoreEmailDomain(e, row.url);
                          const emailColor = matchScore === "match" ? C.accent : matchScore === "personal" ? C.warn : C.blue;
                          const emailBg = matchScore === "match" ? `${C.accent}12` : matchScore === "personal" ? `${C.warn}12` : `${C.blue}12`;
                          const emailBorder = matchScore === "match" ? `${C.accent}22` : matchScore === "personal" ? `${C.warn}22` : `${C.blue}22`;
                          return (
                            <span key={j} title={matchScore === "match" ? "✓ Business email" : matchScore === "personal" ? "⚠ Personal email" : "Domain mismatch"} style={{ fontSize: 10, color: emailColor, background: emailBg, padding: "1px 7px", borderRadius: 3, border: `1px solid ${emailBorder}` }}>{e}</span>
                          );
                        })}
                        {!showAll && (
                          <button onClick={() => setExpanded(true)} style={{ fontSize: 9, color: C.dim, background: "transparent", border: `1px solid ${C.border}`, borderRadius: 3, padding: "1px 6px", cursor: "pointer", fontFamily: "inherit" }}>
                            +{row.emails.length - 3} more
                          </button>
                        )}
                      </div>
                      {/* Feature 6: match summary column */}
                      <span style={{ fontSize: 9, paddingTop: 3 }}>
                        {row.status === "ok" && (() => {
                          const matches = row.emails.filter(e => scoreEmailDomain(e, row.url) === "match").length;
                          const personal = row.emails.filter(e => scoreEmailDomain(e, row.url) === "personal").length;
                          return (
                            <span style={{ display: "flex", flexDirection: "column", gap: 2 }}>
                              {matches > 0 && <span style={{ color: C.accent }}>✓ {matches}</span>}
                              {personal > 0 && <span style={{ color: C.warn }}>⚠ {personal}</span>}
                            </span>
                          );
                        })()}
                      </span>
                      {/* Feature 4: classified status badge */}
                      <span style={{ fontSize: 10, paddingTop: 2 }}>
                        {row.status === "loading" ? <span style={{ color: C.dim }}>…</span>
                          : row.status === "crawling" ? <span style={{ color: C.purple }}>🕷</span>
                          : row.status === "ok" ? <span style={{ color: C.accent }}>✓ {row.emails.length}</span>
                          : row.status === "error" ? <span style={{ color: (BULK_ERROR_META[row.errorClass] ?? BULK_ERROR_META.error).color }}>{(BULK_ERROR_META[row.errorClass] ?? BULK_ERROR_META.error).icon}</span>
                          : <span style={{ color: "#1e2840" }}>— none</span>}
                      </span>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {!bulkResults.length && !bulkRunning && (
            <div style={{ textAlign: "center", padding: "52px 0", color: "#1a2435" }}>
              <div style={{ fontSize: 36, marginBottom: 10 }}>⚡</div>
              <div style={{ fontSize: 10, letterSpacing: ".22em", textTransform: "uppercase", marginBottom: 8 }}>Bulk email scraper — 100% free</div>
              <div style={{ fontSize: 11, color: C.muted }}>Paste URLs or upload a file · No API key needed</div>
            </div>
          )}
        </>)}

        {/* ════════════════ GEMINI TAB ════════════════ */}
        {tab === "gemini" && (<>

          {/* ── Cloudflare Worker URL ────────────────────────────────────── */}
          <div style={{ background: C.surface, border: `1px solid ${workerUrl.trim() ? C.blue : C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            <Label>
              Cloudflare Worker URL
              <span style={{ color: workerUrl.trim() ? C.blue : C.dim, fontWeight: 400, letterSpacing: 0, marginLeft: 8 }}>
                {workerUrl.trim() ? "✓ Connected — ScraperAPI & ZenRows run natively" : "(optional — enables native ScraperAPI & ZenRows without CORS proxies)"}
              </span>
            </Label>
            <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
              <input
                type="text"
                value={workerUrl}
                onChange={e => { setWorkerUrl(e.target.value); setWorkerUrlSaved(false); }}
                placeholder="https://scraper-proxy.YOUR-SUBDOMAIN.workers.dev"
                style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 12px" }}
              />
              <button onClick={() => { SecureStorage.set("workerUrl", workerUrl); setWorkerUrlSaved(true); }} disabled={!workerUrl.trim()} style={{
                padding: "9px 18px", background: workerUrl.trim() ? C.blue : C.surface2,
                color: workerUrl.trim() ? "#000" : C.muted, border: "none", borderRadius: 4,
                cursor: workerUrl.trim() ? "pointer" : "default", fontFamily: "inherit", fontSize: 11, fontWeight: 500, letterSpacing: ".06em",
              }}>{workerUrlSaved ? "✓ Saved" : "Save"}</button>
            </div>
            <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.6 }}>
              Deploy the included <strong style={{ color: C.text }}>worker.js</strong> to Cloudflare (free tier). Without it, ScraperAPI &amp; ZenRows fall back to CORS proxy rotation.
            </div>
          </div>

          {/* ── Provider selector ────────────────────────────────────────── */}
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            <Label>Extraction Provider</Label>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 10 }}>
              {[
                { id: "auto",       label: "⚡ Auto",      tip: "Fallback chain: Gemini → Groq → Regex" },
                { id: "gemini",     label: "✦ Gemini",     tip: "Requires Gemini key" },
                { id: "groq",       label: "🦙 Groq",       tip: "Requires Groq key — no Gemini needed" },
                { id: "scraperapi", label: "🕷 ScraperAPI", tip: "Fetch + regex — no AI required" },
                { id: "zenrows",    label: "🌐 ZenRows",    tip: "Fetch + regex — no AI required" },
              ].map(p => (
                <button key={p.id} title={p.tip} onClick={() => setGeminiProvider(p.id)} style={{
                  padding: "6px 14px", borderRadius: 4, fontFamily: "inherit", fontSize: 11,
                  fontWeight: geminiProvider === p.id ? 600 : 400, letterSpacing: ".05em", cursor: "pointer",
                  transition: "all .15s",
                  background: geminiProvider === p.id ? C.purple : "transparent",
                  color: geminiProvider === p.id ? "#000" : C.muted,
                  border: `1px solid ${geminiProvider === p.id ? C.purple : C.border}`,
                }}>{p.label}</button>
              ))}
            </div>
            <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.7 }}>
              {geminiProvider === "auto"       && "Tries all configured providers in order. No key required — falls back to CORS proxy + regex at minimum."}
              {geminiProvider === "gemini"     && <><a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noreferrer" style={{ color: C.blue }}>Google AI Studio</a> key required. Free tier: 15 req/min, 1,500/day.</>}
              {geminiProvider === "groq"       && <><a href="https://console.groq.com/keys" target="_blank" rel="noreferrer" style={{ color: C.blue }}>console.groq.com</a> key required. Uses llama-3.3-70b-versatile — <strong style={{ color: C.text }}>runs without Gemini</strong>.</>}
              {geminiProvider === "scraperapi" && <>ScraperAPI key required. Fetches JS-rendered pages then extracts emails via regex — <strong style={{ color: C.text }}>no AI needed</strong>.</>}
              {geminiProvider === "zenrows"    && <>ZenRows key required. Fetches JS-rendered pages then extracts emails via regex — <strong style={{ color: C.text }}>no AI needed</strong>.</>}
            </div>
          </div>

          {/* ── Gemini key — shown for gemini + auto modes ───────────────── */}
          {(geminiProvider === "gemini" || geminiProvider === "auto") && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
              <Label>Gemini API Key {geminiProvider === "auto" && <span style={{ color: C.dim, fontWeight: 400, letterSpacing: 0 }}>(optional in auto mode)</span>}</Label>
              <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
                <input type="password" value={geminiKey}
                  onChange={e => { setGeminiKey(e.target.value); setGeminiKeySaved(false); }}
                  placeholder="Paste your Gemini API key…"
                  style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 12px" }}
                />
                <button onClick={() => { SecureStorage.set("geminiKey", geminiKey); setGeminiKeySaved(true); }} disabled={!geminiKey.trim()} style={{
                  padding: "9px 18px", background: geminiKey.trim() ? C.purple : C.surface2,
                  color: geminiKey.trim() ? "#000" : C.muted, border: "none", borderRadius: 4,
                  cursor: geminiKey.trim() ? "pointer" : "default", fontFamily: "inherit", fontSize: 11, fontWeight: 500, letterSpacing: ".06em",
                }}>{geminiKeySaved ? "✓ Saved" : "Save"}</button>
              </div>
              <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.6 }}>
                Free at <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noreferrer">Google AI Studio</a>. Free tier: 15 requests/min, 1,500/day.
              </div>
            </div>
          )}

          {/* ── Groq key — shown for groq + auto modes ───────────────────── */}
          {(geminiProvider === "groq" || geminiProvider === "auto") && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
              <Label>Groq API Key {geminiProvider === "auto" && <span style={{ color: C.dim, fontWeight: 400, letterSpacing: 0 }}>(optional in auto mode)</span>}</Label>
              <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
                <input type="password" value={groqKey}
                  onChange={e => { setGroqKey(e.target.value); setGroqKeySaved(false); }}
                  placeholder="Paste your Groq API key…"
                  style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 12px" }}
                />
                <button onClick={() => { SecureStorage.set("groqKey", groqKey); setGroqKeySaved(true); }} disabled={!groqKey.trim()} style={{
                  padding: "9px 18px", background: groqKey.trim() ? C.purple : C.surface2,
                  color: groqKey.trim() ? "#000" : C.muted, border: "none", borderRadius: 4,
                  cursor: groqKey.trim() ? "pointer" : "default", fontFamily: "inherit", fontSize: 11, fontWeight: 500, letterSpacing: ".06em",
                }}>{groqKeySaved ? "✓ Saved" : "Save"}</button>
              </div>
              <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.6 }}>
                Free at <a href="https://console.groq.com/keys" target="_blank" rel="noreferrer">console.groq.com</a>. <strong style={{ color: C.text }}>Runs independently — no Gemini key needed.</strong> Uses llama-3.3-70b-versatile.
              </div>
            </div>
          )}

          {/* ── ScraperAPI key — shown for scraperapi + auto modes ────────── */}
          {(geminiProvider === "scraperapi" || geminiProvider === "auto") && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
              <Label>ScraperAPI Key {geminiProvider === "auto" && <span style={{ color: C.dim, fontWeight: 400, letterSpacing: 0 }}>(optional in auto mode)</span>}</Label>
              <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
                <input type="password" value={scraperApiKey}
                  onChange={e => { setScraperApiKey(e.target.value); setScraperApiKeySaved(false); }}
                  placeholder="Paste your ScraperAPI key…"
                  style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 12px" }}
                />
                <button onClick={() => { SecureStorage.set("scraperApiKey", scraperApiKey); setScraperApiKeySaved(true); }} disabled={!scraperApiKey.trim()} style={{
                  padding: "9px 18px", background: scraperApiKey.trim() ? C.blue : C.surface2,
                  color: scraperApiKey.trim() ? "#000" : C.muted, border: "none", borderRadius: 4,
                  cursor: scraperApiKey.trim() ? "pointer" : "default", fontFamily: "inherit", fontSize: 11, fontWeight: 500, letterSpacing: ".06em",
                }}>{scraperApiKeySaved ? "✓ Saved" : "Save"}</button>
              </div>
              <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.6 }}>
                Get a key at <a href="https://www.scraperapi.com" target="_blank" rel="noreferrer">scraperapi.com</a>. <strong style={{ color: C.text }}>Runs independently — no AI key needed.</strong> Uses regex extraction.
              </div>
            </div>
          )}

          {/* ── ZenRows key — shown for zenrows + auto modes ──────────────── */}
          {(geminiProvider === "zenrows" || geminiProvider === "auto") && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
              <Label>ZenRows API Key {geminiProvider === "auto" && <span style={{ color: C.dim, fontWeight: 400, letterSpacing: 0 }}>(optional in auto mode)</span>}</Label>
              <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
                <input type="password" value={zenRowsKey}
                  onChange={e => { setZenRowsKey(e.target.value); setZenRowsKeySaved(false); }}
                  placeholder="Paste your ZenRows API key…"
                  style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 12px" }}
                />
                <button onClick={() => { SecureStorage.set("zenRowsKey", zenRowsKey); setZenRowsKeySaved(true); }} disabled={!zenRowsKey.trim()} style={{
                  padding: "9px 18px", background: zenRowsKey.trim() ? C.blue : C.surface2,
                  color: zenRowsKey.trim() ? "#000" : C.muted, border: "none", borderRadius: 4,
                  cursor: zenRowsKey.trim() ? "pointer" : "default", fontFamily: "inherit", fontSize: 11, fontWeight: 500, letterSpacing: ".06em",
                }}>{zenRowsKeySaved ? "✓ Saved" : "Save"}</button>
              </div>
              <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.6 }}>
                Get a key at <a href="https://www.zenrows.com" target="_blank" rel="noreferrer">zenrows.com</a>. <strong style={{ color: C.text }}>Runs independently — no AI key needed.</strong> Uses regex extraction.
              </div>
            </div>
          )}

          {/* ── Hunter + Apollo enrichment panel ─────────────────────────────── */}
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
              <Label>Lead Enrichment (optional)</Label>
              <span style={{ fontSize: 9, color: C.dim, letterSpacing: ".1em" }}>HUNTER FALLBACK · APOLLO ENRICHMENT</span>
            </div>

            {/* Hunter */}
            <div style={{ marginBottom: 14, paddingBottom: 14, borderBottom: `1px solid ${C.border}` }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 11, color: hunterEnabled ? C.accent : C.muted }}>
                  <input type="checkbox" checked={hunterEnabled}
                    onChange={e => { setHunterEnabled(e.target.checked); try { localStorage.setItem("hunterEnabled", e.target.checked); } catch {} }}
                    style={{ accentColor: C.accent, cursor: "pointer" }} />
                  Enable Snov.io fallback
                </label>
                <span style={{ fontSize: 9, color: C.dim }}>(up to 5 calls/run · fires only when 0 emails found · Snov.io)</span>
              </div>
              {hunterEnabled && (
                <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 6 }}>
                  <div style={{ display: "flex", gap: 8 }}>
                    <input type="password" value={hunterKey}
                      onChange={e => { setHunterKey(e.target.value); setHunterKeySaved(false); setSnovSecretSaved(false); }}
                      placeholder="Snov.io Client ID…"
                      style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "8px 12px" }} />
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <input type="password" value={snovSecret}
                      onChange={e => { setSnovSecret(e.target.value); setSnovSecretSaved(false); setHunterKeySaved(false); }}
                      placeholder="Snov.io Client Secret…"
                      style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "8px 12px" }} />
                    <button onClick={() => {
                      SecureStorage.set("hunterKey", hunterKey); SecureStorage.set("snovSecret", snovSecret);
                      setHunterKeySaved(true); setSnovSecretSaved(true);
                    }} disabled={!hunterKey.trim() || !snovSecret.trim()} style={{
                      padding: "8px 16px", background: (hunterKey.trim() && snovSecret.trim()) ? C.accent : C.surface2,
                      color: (hunterKey.trim() && snovSecret.trim()) ? "#000" : C.muted, border: "none", borderRadius: 4,
                      cursor: (hunterKey.trim() && snovSecret.trim()) ? "pointer" : "default", fontFamily: "inherit", fontSize: 11, fontWeight: 500,
                    }}>{hunterKeySaved && snovSecretSaved ? "✓ Saved" : "Save"}</button>
                  </div>
                </div>
              )}
              {hunterEnabled && <div style={{ fontSize: 10, color: C.dim }}>Free at <a href="https://snov.io" target="_blank" rel="noreferrer" style={{ color: C.blue }}>snov.io</a> — 50 credits/month. Go to Profile → API to get your Client ID and Secret.</div>}
            </div>

            {/* Apollo */}
            <div>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 8 }}>
                <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer", fontSize: 11, color: apolloEnabled ? C.purple : C.muted }}>
                  <input type="checkbox" checked={apolloEnabled}
                    onChange={e => { setApolloEnabled(e.target.checked); try { localStorage.setItem("apolloEnabled", e.target.checked); } catch {} }}
                    style={{ accentColor: C.purple, cursor: "pointer" }} />
                  Enable Apollo.io enrichment
                </label>
                <span style={{ fontSize: 9, color: C.dim }}>(up to 5 calls/run · adds name, industry, size, location)</span>
              </div>
              {apolloEnabled && (
                <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
                  <input type="password" value={apolloKey}
                    onChange={e => { setApolloKey(e.target.value); setApolloKeySaved(false); }}
                    placeholder="Paste your Apollo.io API key…"
                    style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "8px 12px" }} />
                  <button onClick={() => { SecureStorage.set("apolloKey", apolloKey); setApolloKeySaved(true); }}
                    disabled={!apolloKey.trim()} style={{
                      padding: "8px 16px", background: apolloKey.trim() ? C.purple : C.surface2,
                      color: apolloKey.trim() ? "#000" : C.muted, border: "none", borderRadius: 4,
                      cursor: apolloKey.trim() ? "pointer" : "default", fontFamily: "inherit", fontSize: 11, fontWeight: 500,
                    }}>{apolloKeySaved ? "✓ Saved" : "Save"}</button>
                </div>
              )}
              {apolloEnabled && <div style={{ fontSize: 10, color: C.dim }}>Free at <a href="https://app.apollo.io" target="_blank" rel="noreferrer" style={{ color: C.blue }}>apollo.io</a> — 50 credits/month. Returns company name, industry, size, location.</div>}
            </div>
          </div>

          <div style={{ display: "flex", gap: 3, marginBottom: 12 }}>
            {[{ id: "text", label: "✏  Paste URLs" }, { id: "file", label: "📄  Upload File" }].map(t => (
              <button key={t.id} onClick={() => setGeminiInputMode(t.id)} style={{
                padding: "6px 16px", background: geminiInputMode === t.id ? C.purple : "transparent",
                color: geminiInputMode === t.id ? "#000" : C.muted, border: `1px solid ${geminiInputMode === t.id ? C.purple : C.border}`,
                borderRadius: 4, cursor: "pointer", fontSize: 11, fontFamily: "inherit",
                fontWeight: geminiInputMode === t.id ? 500 : 400, letterSpacing: ".06em", transition: "all .15s",
              }}>{t.label}</button>
            ))}
          </div>

          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            {geminiInputMode === "text" ? (<>
              <Label>Paste URLs (one per line)</Label>
              <textarea
                ref={geminiTextareaRef}
                defaultValue=""
                placeholder={"https://company1.com/contact\nhttps://company2.com/about"}
                onInput={e => setGeminiUrls(e.target.value)}
                onChange={e => setGeminiUrls(e.target.value)}
                onPaste={e => {
                  // On Android WebView, paste does not reliably fire onChange/onInput.
                  // Read clipboard directly and sync to state after a short delay
                  // to let the WebView finish writing to the DOM first.
                  setTimeout(() => {
                    const val = geminiTextareaRef.current?.value ?? "";
                    if (val) setGeminiUrls(val);
                  }, 100);
                }}
                style={{ width: "100%", height: 130, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 11, padding: "10px 14px", resize: "vertical", lineHeight: 1.8, transition: "border-color .15s" }} />
              <div style={{ fontSize: 10, color: C.dim, marginTop: 6 }}>
                {parseGeminiUrls(geminiUrls).length > 0 && <span style={{ color: C.purple }}>{parseGeminiUrls(geminiUrls).length} URLs detected</span>}
              </div>
            </>) : (<>
              <Label>Upload CSV, XLSX or TXT with one URL per line</Label>
              <label htmlFor="gemini-file-input" style={{ display: "block", border: `2px dashed ${geminiFile ? C.purple : C.border}`, borderRadius: 6, padding: "26px 20px", textAlign: "center", cursor: "pointer", background: geminiFile ? `${C.purple}07` : "transparent", transition: "all .2s" }}>
                <div style={{ fontSize: 26, marginBottom: 7 }}>📂</div>
                <div style={{ fontSize: 12, color: geminiFile ? C.purple : C.muted }}>{geminiFile ? `✓  ${geminiFile.name}  (${(geminiFile.size / 1024).toFixed(1)} KB)` : "Tap to choose — CSV, XLSX or TXT"}</div>
              </label>
              <input
                id="gemini-file-input"
                ref={geminiFileRef}
                type="file"
                accept=".csv,.xlsx,.txt"
                style={{ position: "absolute", opacity: 0, width: 1, height: 1, pointerEvents: "none" }}
                onClick={e => { e.target.value = ""; }}
                onChange={e => setGeminiFile(e.target.files[0] || null)}
              />
            </>)}

            <div style={{ display: "flex", gap: 8, marginTop: 14, alignItems: "center" }}>
              {!geminiRunning ? (
                <button className="run-btn" onClick={runGeminiScrape}
                  disabled={
                    (geminiInputMode === "text" ? !parseGeminiUrls(geminiUrls).length : !geminiFile) ||
                    (geminiProvider === "gemini"     && !geminiKey.trim()) ||
                    (geminiProvider === "groq"       && !groqKey.trim()) ||
                    (geminiProvider === "scraperapi" && !scraperApiKey.trim()) ||
                    (geminiProvider === "zenrows"    && !zenRowsKey.trim())
                  }
                  style={{
                    padding: "9px 26px",
                    background: (() => {
                      const hasUrls = geminiInputMode === "text" ? parseGeminiUrls(geminiUrls).length : geminiFile;
                      const hasKey = geminiProvider === "auto" ? true
                        : geminiProvider === "gemini"     ? geminiKey.trim()
                        : geminiProvider === "groq"       ? groqKey.trim()
                        : geminiProvider === "scraperapi" ? scraperApiKey.trim()
                        : geminiProvider === "zenrows"    ? zenRowsKey.trim() : false;
                      return hasUrls && hasKey ? C.purple : C.surface2;
                    })(),
                    color: (() => {
                      const hasUrls = geminiInputMode === "text" ? parseGeminiUrls(geminiUrls).length : geminiFile;
                      const hasKey = geminiProvider === "auto" ? true
                        : geminiProvider === "gemini"     ? geminiKey.trim()
                        : geminiProvider === "groq"       ? groqKey.trim()
                        : geminiProvider === "scraperapi" ? scraperApiKey.trim()
                        : geminiProvider === "zenrows"    ? zenRowsKey.trim() : false;
                      return hasUrls && hasKey ? "#000" : C.muted;
                    })(),
                    border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit",
                    fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase",
                    transition: "all .15s", display: "flex", alignItems: "center", gap: 8,
                  }}>
                  ✦  Start {geminiProvider === "auto" ? "Smart" : geminiProvider.charAt(0).toUpperCase() + geminiProvider.slice(1)} Scrape
                </button>
              ) : (
                <button onClick={() => { geminiStopRef.current = true; ScrapeNative.stop(); }} style={{ padding: "9px 22px", background: C.danger, color: "#fff", border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase" }}>
                  ◼  Stop
                </button>
              )}
              {geminiRunning && (
                <div style={{ fontSize: 11, color: C.dim, display: "flex", alignItems: "center", gap: 6 }}>
                  <Spin /><span>{geminiProgress.done} / {geminiProgress.total} URLs</span>
                  <span style={{ color: C.purple }}>{geminiTotalEmails} emails found</span>
                </div>
              )}
            </div>
          </div>

          {(geminiRunning || geminiResults.length > 0) && geminiProgress.total > 0 && (
            <div style={{ marginBottom: 14 }}>
              <div style={{ height: 3, background: C.surface2, borderRadius: 2, overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${(geminiProgress.done / geminiProgress.total) * 100}%`, background: C.purple, transition: "width .3s", borderRadius: 2 }} />
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6, fontSize: 10, color: C.dim }}>
                <span>{geminiStopped ? "⏹ Stopped" : geminiRunning ? "✦ Running…" : "✓ Complete"}</span>
                <span style={{ color: C.purple }}>{geminiOk} ok</span>
                <span style={{ color: C.warn }}>{geminiResults.filter(r => r.status === "none").length} no emails</span>
                <span style={{ color: C.danger }}>{geminiFailed} failed</span>
                {geminiEta && <span style={{ color: C.dim }}>{geminiEta}</span>}
                <span style={{ color: C.text, fontWeight: 500 }}>{geminiTotalEmails} total emails</span>
              </div>
            </div>
          )}

          {geminiResults.length > 0 && !geminiRunning && (
            <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
              <button onClick={importGeminiToLeads} disabled={!geminiTotalEmails} style={{ padding: "7px 18px", background: geminiTotalEmails ? C.purple : C.surface2, color: geminiTotalEmails ? "#000" : C.muted, border: "none", borderRadius: 4, cursor: geminiTotalEmails ? "pointer" : "not-allowed", fontFamily: "inherit", fontWeight: 500, fontSize: 11, letterSpacing: ".07em" }}>
                → Import {geminiTotalEmails} emails to Leads
              </button>
              <button onClick={exportGeminiCSV} disabled={!geminiTotalEmails} style={{ padding: "7px 18px", background: "transparent", color: geminiTotalEmails ? C.blue : C.muted, border: `1px solid ${geminiTotalEmails ? C.blue + "44" : C.border}`, borderRadius: 4, cursor: geminiTotalEmails ? "pointer" : "not-allowed", fontFamily: "inherit", fontSize: 11, letterSpacing: ".07em" }}>
                ↓ Export CSV
              </button>
              <button onClick={exportGeminiXLSX} disabled={!geminiTotalEmails} style={{ padding: "7px 18px", background: "transparent", color: geminiTotalEmails ? "#1ed98a" : C.muted, border: `1px solid ${geminiTotalEmails ? "#1ed98a44" : C.border}`, borderRadius: 4, cursor: geminiTotalEmails ? "pointer" : "not-allowed", fontFamily: "inherit", fontSize: 11, letterSpacing: ".07em" }}>
                ↓ Export XLSX
              </button>
              <button onClick={() => { setGeminiResults([]); setGeminiProgress({ done: 0, total: 0 }); }} style={{ padding: "7px 14px", background: "transparent", color: C.danger, border: `1px solid ${C.danger}33`, borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontSize: 11 }}>
                clear
              </button>
            </div>
          )}

          {geminiResults.length > 0 && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, overflow: "hidden" }}>
              <div style={{ display: "grid", gridTemplateColumns: "1.6fr 2.4fr 1.4fr 80px 90px", padding: "8px 14px", borderBottom: `1px solid ${C.border}`, fontSize: 9, letterSpacing: ".18em", color: C.muted, textTransform: "uppercase" }}>
                <span>URL</span><span>Emails Found</span><span>Enrichment</span><span>Status</span><span>Provider</span>
              </div>
              <div style={{ maxHeight: 460, overflowY: "auto" }}>
                {geminiResults.map((row, i) => {
                  // ── provider badge config ──────────────────────────────────
                  const providerBadge = (() => {
                    const p = row.provider ?? "";
                    if (p.includes("scraperapi")) return { dot: "🔵", label: "scraperapi", color: "#3b82f6" };
                    if (p.includes("zenrows"))    return { dot: "🟡", label: "zenrows",    color: "#eab308" };
                    if (p.includes("groq"))       return { dot: "🟣", label: "groq",       color: "#a855f7" };
                    if (p.includes("gemini"))     return { dot: "🟢", label: "gemini",     color: "#22c55e" };
                    if (p.includes("proxy"))      return { dot: "⚪", label: "proxy",      color: "#6b7280" };
                    if (p.includes("hunter"))     return { dot: "🟠", label: "hunter",     color: "#f97316" };
                    if (p.includes("apollo"))     return { dot: "🔴", label: "apollo",     color: "#ef4444" };
                    return { dot: "⬜", label: p || "—", color: C.dim };
                  })();
                  return (
                  <div key={i} className="bulk-row fade-row" style={{
                    display: "grid", gridTemplateColumns: "1.6fr 2.4fr 1.4fr 80px 90px",
                    padding: "9px 14px", borderBottom: i < geminiResults.length - 1 ? `1px solid ${C.border}` : "none",
                    alignItems: "center", fontSize: 11, background: i % 2 ? "#0a0f16" : "transparent",
                  }}>
                    <span style={{ color: C.dim, fontSize: 10, wordBreak: "break-all", paddingRight: 10 }}>
                      {(() => { try { return new URL(row.url).hostname; } catch { return row.url; } })()}
                    </span>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                      {row.status === "loading" && <span className="loading-dot" style={{ fontSize: 10, color: C.dim }}>scanning…</span>}
                      {row.status === "crawling" && <span className="loading-dot" style={{ fontSize: 10, color: C.purple }}>🕷 {row.crawlMsg || "deep crawling…"}</span>}
                      {row.status === "error" && <span style={{ fontSize: 10, color: C.danger }}>⚠ {row.error?.slice(0, 40)}</span>}
                      {row.status === "none" && <span style={{ fontSize: 10, color: "#1e2840" }}>no emails found</span>}
                      {row.emails.map((e, j) => (
                        <span key={j} style={{ fontSize: 10, color: C.purple, background: `${C.purple}12`, padding: "1px 7px", borderRadius: 3, border: `1px solid ${C.purple}22` }}>{e}</span>
                      ))}
                    </div>
                    <div style={{ fontSize: 9, color: C.dim, lineHeight: 1.7 }}>
                      {row.enrichment ? (<>
                        {row.enrichment.name && <div style={{ color: C.text, fontWeight: 500 }}>{row.enrichment.name}</div>}
                        {row.enrichment.industry && <div>{row.enrichment.industry}</div>}
                        {row.enrichment.size && <div>{row.enrichment.size.toLocaleString()} employees</div>}
                        {row.enrichment.city && <div style={{ color: C.dim }}>{row.enrichment.city}{row.enrichment.country ? `, ${row.enrichment.country}` : ""}</div>}
                        {row.enrichment.decisionMakers?.length > 0 && (
                          <div style={{ marginTop: 4, paddingTop: 4, borderTop: `1px solid #1e2840` }}>
                            {row.enrichment.decisionMakers.map((dm, di) => (
                              <div key={di} style={{ marginBottom: 3 }}>
                                {dm.name && <span style={{ color: "#f97316", fontWeight: 500, fontSize: 9 }}>{dm.name}</span>}
                                {dm.title && <span style={{ color: C.dim, fontSize: 9 }}> · {dm.title}</span>}
                                {dm.email && <div style={{ color: C.accent, fontSize: 9 }}>{dm.email}</div>}
                              </div>
                            ))}
                          </div>
                        )}
                      </>) : <span style={{ color: "#1e2840" }}>—</span>}
                    </div>
                    <span style={{ fontSize: 10, color: row.status === "ok" ? C.purple : row.status === "error" ? C.danger : row.status === "none" ? "#1e2840" : C.dim }}>
                      {row.status === "loading" ? "…" : row.status === "crawling" ? <span style={{ color: C.purple }}>🕷</span> : row.status === "ok" ? `✓ ${row.emails.length}` : row.status === "error" ? "✕ fail" : "— none"}
                    </span>
                    <span style={{ fontSize: 9, color: providerBadge.color, display: "flex", alignItems: "center", gap: 3, letterSpacing: ".04em" }}>
                      {row.status === "loading" ? <span style={{ color: C.dim }}>…</span> : <>{providerBadge.dot} {providerBadge.label}</>}
                    </span>
                  </div>
                  );
                })}
              </div>
            </div>
          )}

          {!geminiResults.length && !geminiRunning && (
            <div style={{ textAlign: "center", padding: "52px 0", color: "#1a2435" }}>
              <div style={{ fontSize: 36, marginBottom: 10 }}>✦</div>
              <div style={{ fontSize: 10, letterSpacing: ".22em", textTransform: "uppercase", marginBottom: 8 }}>Smart email extraction — any provider</div>
              <div style={{ fontSize: 11, color: C.muted }}>Select a provider above · Gemini, Groq, ScraperAPI, ZenRows or Auto</div>
            </div>
          )}
        </>)}

        {/* ════════════════ MAPS TAB ════════════════ */}
        {tab === "maps" && (<>
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            <Label>Google Places API Key</Label>
            <div style={{ display: "flex", gap: 8, marginBottom: 6 }}>
              <input
                type="password"
                value={placesKey}
                onChange={e => { setPlacesKey(e.target.value); setPlacesKeySaved(false); }}
                placeholder="Paste your Google Places API key…"
                style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 12px" }}
              />
              <button onClick={() => { SecureStorage.set("placesKey", placesKey); setPlacesKeySaved(true); }} disabled={!placesKey.trim()} style={{
                padding: "9px 18px", background: placesKey.trim() ? C.accent : C.surface2,
                color: placesKey.trim() ? "#000" : C.muted, border: "none", borderRadius: 4,
                cursor: placesKey.trim() ? "pointer" : "default", fontFamily: "inherit", fontSize: 11, fontWeight: 500, letterSpacing: ".06em",
              }}>{placesKeySaved ? "✓ Saved" : "Save"}</button>
            </div>
            <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.6 }}>
              Get a free key at <a href="https://console.cloud.google.com/apis/library/places-backend.googleapis.com" target="_blank" rel="noreferrer">Google Cloud Console</a> → Enable "Places API (New)" → Create API Key. Free tier: $200/month credit (~5,000 searches free).
            </div>
          </div>

          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            <Label>Search Query</Label>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 14 }}>
              <div>
                <div style={{ fontSize: 10, color: C.dim, marginBottom: 5, letterSpacing: ".1em" }}>WHAT (business type / keyword)</div>
                <input value={mapsQuery} onChange={e => setMapsQuery(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && !mapsLoading && mapsQuery && placesKey && runMapsScrape()}
                  placeholder="e.g. dental clinic, law firm, coffee shop"
                  style={{ width: "100%", background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 12px" }} />
              </div>
              <div>
                <div style={{ fontSize: 10, color: C.dim, marginBottom: 5, letterSpacing: ".1em" }}>WHERE (city, area, or leave blank)</div>
                <input value={mapsLocation} onChange={e => setMapsLocation(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && !mapsLoading && mapsQuery && placesKey && runMapsScrape()}
                  placeholder="e.g. Lagos, Lekki, New York"
                  style={{ width: "100%", background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 12px" }} />
              </div>
            </div>

            {mapsError && <div style={{ marginBottom: 12, fontSize: 11, color: "#ff7070", padding: "7px 12px", background: "#ff606012", borderRadius: 4, border: "1px solid #ff606033" }}>⚠  {mapsError}</div>}

            {/* Max pages control */}
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12, padding: "10px 14px", background: C.surface2, border: `1px solid ${C.border}`, borderRadius: 6 }}>
              <span style={{ fontSize: 11, color: C.dim, flex: 1 }}>
                🗂 <strong style={{ color: C.text }}>Auto-paginate</strong> — fetches multiple pages of results automatically
              </span>
              <div style={{ display: "flex", flexDirection: "column", alignItems: "center", gap: 2 }}>
                <span style={{ fontSize: 9, color: C.muted, letterSpacing: ".06em" }}>MAX PAGES</span>
                <select value={mapsMaxPages} onChange={e => setMapsMaxPages(Number(e.target.value))}
                  style={{ background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>
                  <option value={1}>1 (~20)</option>
                  <option value={3}>3 (~60)</option>
                  <option value={5}>5 (~100)</option>
                  <option value={10}>10 (~200)</option>
                </select>
              </div>
            </div>

            {/* Multi-query mode toggle */}
            <div style={{ marginBottom: 14, padding: "12px 14px", background: multiQueryMode ? "#0d1f0d" : C.surface2, border: `1px solid ${multiQueryMode ? "#2d6a2d" : C.border}`, borderRadius: 6 }}>
              <label style={{ display: "flex", alignItems: "flex-start", gap: 10, cursor: "pointer", marginBottom: multiQueryMode ? 12 : 0 }}>
                <input type="checkbox" checked={multiQueryMode} onChange={e => setMultiQueryMode(e.target.checked)}
                  style={{ accentColor: "#4ade80", marginTop: 2, cursor: "pointer" }} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 11, color: multiQueryMode ? "#4ade80" : C.text, fontWeight: 500, marginBottom: 3 }}>
                    🚀 Multi-Query Mode — bypass 60-result limit
                  </div>
                  <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.6 }}>
                    Auto-splits search by neighborhoods/districts + keyword variations.
                    Covers <strong style={{ color: C.text }}>100+ cities worldwide</strong> including US, Nigeria, UK, Canada, Australia, India, UAE and more.
                    Returns <strong style={{ color: C.text }}>500–2000+ unique results</strong> per search.
                  </div>
                  {multiQueryMode && multiQueryProgress && (
                    <div style={{ fontSize: 10, color: "#4ade80", marginTop: 6 }}>{multiQueryProgress}</div>
                  )}
                </div>
              </label>

              {multiQueryMode && (
                <div style={{ marginTop: 4 }}>
                  <div style={{ fontSize: 9, color: C.dim, letterSpacing: ".1em", marginBottom: 5 }}>
                    CUSTOM AREAS (optional — one per line or comma separated. Overrides auto-split)
                  </div>
                  <textarea
                    value={customAreas}
                    onChange={e => setCustomAreas(e.target.value)}
                    placeholder={"e.g.\nMidtown Manhattan\nBrooklyn Heights\nQueens Village\n\nLeave blank to use automatic city splits"}
                    rows={4}
                    style={{
                      width: "100%", background: "#070b0f", border: `1px solid ${C.border}`,
                      borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 11,
                      padding: "8px 10px", resize: "vertical", boxSizing: "border-box",
                    }}
                  />
                  <div style={{ fontSize: 9, color: C.dim, marginTop: 4 }}>
                    {customAreas.trim()
                      ? `✓ Using ${customAreas.split(/[\n,]+/).filter(a => a.trim()).length} custom areas`
                      : "Auto-detecting city — leave blank for automatic neighborhood splits"}
                  </div>
                </div>
              )}
            </div>

            {/* ── Keyword Expansion panel ────────────────────────────────────── */}
            <div style={{ marginBottom: 14, padding: "12px 14px", background: keywordExpand ? "#0d0d1f" : C.surface2, border: `1px solid ${keywordExpand ? C.purple : C.border}`, borderRadius: 6 }}>
              {/* Master toggle */}
              <label style={{ display: "flex", alignItems: "flex-start", gap: 10, cursor: "pointer", marginBottom: keywordExpand ? 14 : 0 }}>
                <input type="checkbox" checked={keywordExpand} onChange={e => {
                  setKeywordExpand(e.target.checked);
                  if (!e.target.checked) setKwPreview({ static: [], suggest: [], ai: [], merged: [], loading: false, error: "" });
                }} style={{ accentColor: C.purple, marginTop: 2, cursor: "pointer" }} />
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 11, color: keywordExpand ? C.purple : C.text, fontWeight: 500, marginBottom: 3 }}>
                    ✦ Keyword Expansion — turn one keyword into a lead-discovery engine
                  </div>
                  <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.6 }}>
                    Expands your keyword using <strong style={{ color: C.text }}>static rules</strong>, <strong style={{ color: C.text }}>Google Suggest</strong>, and optional <strong style={{ color: C.text }}>Groq AI</strong>.
                    All expanded keywords feed into the existing scraper and are merged + deduplicated.
                  </div>
                </div>
              </label>

              {keywordExpand && (<>
                {/* Sub-toggles row */}
                <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
                  {[
                    { key: "static",  label: "📋 Static Rules",    active: true,        color: C.accent,  locked: true,  tip: "Always active — built-in category mappings" },
                    { key: "suggest", label: "🔍 Google Suggest",  active: kwUseSuggest, color: C.blue,   locked: false, tip: "Real Google autocomplete phrases" },
                    { key: "ai",      label: "✦ Groq AI",          active: kwUseAI,      color: C.purple, locked: false, tip: groqKey.trim() ? "AI-generated niche variations" : "Add Groq API key to enable" },
                  ].map(({ key, label, active, color, locked, tip }) => (
                    <button key={key} title={tip} onClick={() => {
                      if (locked) return;
                      if (key === "suggest") setKwUseSuggest(v => !v);
                      if (key === "ai") setKwUseAI(v => !v);
                    }} style={{
                      padding: "5px 11px", borderRadius: 4, cursor: locked ? "default" : "pointer",
                      fontFamily: "inherit", fontSize: 10, letterSpacing: ".04em",
                      background: active ? `${color}18` : "transparent",
                      border: `1px solid ${active ? color : C.border}`,
                      color: active ? color : C.muted,
                      opacity: key === "ai" && !groqKey.trim() ? 0.45 : 1,
                    }}>
                      {label}{locked ? " ✓" : active ? " ✓" : ""}
                    </button>
                  ))}
                </div>

                {/* Limits row */}
                <div style={{ display: "flex", gap: 14, marginBottom: 12, flexWrap: "wrap", alignItems: "flex-end" }}>
                  <div>
                    <div style={{ fontSize: 9, color: C.muted, letterSpacing: ".1em", marginBottom: 4 }}>MAX KEYWORDS</div>
                    <select value={kwMaxExpanded} onChange={e => setKwMaxExpanded(Number(e.target.value))}
                      style={{ background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>
                      {[4, 6, 8, 12, 16, 20].map(n => <option key={n} value={n}>{n}</option>)}
                    </select>
                  </div>
                  <div>
                    <div style={{ fontSize: 9, color: C.muted, letterSpacing: ".1em", marginBottom: 4 }}>MAX SEARCHES / RUN</div>
                    <select value={kwMaxSearches} onChange={e => setKwMaxSearches(Number(e.target.value))}
                      style={{ background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 11, padding: "4px 8px", cursor: "pointer" }}>
                      {[10, 20, 30, 50, 100, 200].map(n => <option key={n} value={n}>{n}</option>)}
                    </select>
                  </div>
                  {/* Preview button */}
                  <button disabled={!mapsQuery.trim() || kwPreview.loading} onClick={async () => {
                    if (!mapsQuery.trim()) return;
                    setKwPreview(p => ({ ...p, loading: true, error: "" }));
                    const staticKws = expandKeywordsStatic(mapsQuery.trim());
                    let suggestKws = [], aiKws = [];
                    if (kwUseSuggest) {
                      try { suggestKws = await fetchGoogleSuggest(mapsQuery.trim(), 8, new Set(staticKws.map(k => k.toLowerCase()))); }
                      catch (e) { console.warn("Suggest preview failed:", e.message); }
                    }
                    if (kwUseAI && groqKey.trim()) {
                      try {
                        const existing = new Set([...staticKws, ...suggestKws].map(k => k.toLowerCase()));
                        aiKws = await fetchGroqKeywords(groqKey.trim(), mapsQuery.trim(), 10, existing);
                      } catch (e) { console.warn("AI preview failed:", e.message); }
                    }
                    const merged = mergeKeywords(staticKws, suggestKws, aiKws, kwMaxExpanded);
                    setKwPreview({ static: staticKws, suggest: suggestKws, ai: aiKws, merged, loading: false, error: "" });
                  }} style={{
                    padding: "5px 14px", background: "transparent", border: `1px solid ${C.purple}55`,
                    borderRadius: 4, color: C.purple, cursor: mapsQuery.trim() ? "pointer" : "default",
                    fontFamily: "inherit", fontSize: 10, opacity: mapsQuery.trim() ? 1 : 0.4,
                  }}>
                    {kwPreview.loading ? "Loading…" : "↺ Preview Keywords"}
                  </button>
                </div>

                {/* Keyword preview chips — grouped by source */}
                {kwPreview.merged.length > 0 && (() => {
                  const groups = [
                    { label: "📋 Static",         kws: kwPreview.static,   color: C.accent  },
                    { label: "🔍 Google Suggest", kws: kwPreview.suggest,  color: C.blue    },
                    { label: "✦ Groq AI",         kws: kwPreview.ai,       color: C.purple  },
                  ].filter(g => g.kws.length > 0);
                  return (
                    <div style={{ marginBottom: 4 }}>
                      {groups.map(g => (
                        <div key={g.label} style={{ marginBottom: 8 }}>
                          <div style={{ fontSize: 9, color: C.muted, letterSpacing: ".1em", marginBottom: 4 }}>{g.label} ({g.kws.length})</div>
                          <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                            {g.kws.map((kw, i) => {
                              const inMerged = kwPreview.merged.includes(kw);
                              return (
                                <span key={i} style={{
                                  fontSize: 10, padding: "2px 9px", borderRadius: 4,
                                  background: inMerged ? `${g.color}20` : "transparent",
                                  border: `1px solid ${inMerged ? g.color : g.color + "33"}`,
                                  color: inMerged ? g.color : C.muted,
                                  textDecoration: inMerged ? "none" : "line-through",
                                  opacity: inMerged ? 1 : 0.5,
                                }}>{kw}</span>
                              );
                            })}
                          </div>
                        </div>
                      ))}
                      <div style={{ fontSize: 9, color: C.dim, marginTop: 4 }}>
                        <span style={{ color: C.purple }}>{kwPreview.merged.length}</span> keywords → up to{" "}
                        <span style={{ color: C.purple }}>
                          {multiQueryMode
                            ? `${Math.min(kwPreview.merged.length * 3, kwMaxSearches)}+ searches (multi-query)`
                            : `${Math.min(kwPreview.merged.length, kwMaxSearches)} searches`}
                        </span>
                      </div>
                    </div>
                  );
                })()}

                {!mapsQuery.trim() && (
                  <div style={{ fontSize: 10, color: C.warn }}>⚠ Enter a keyword above to preview expansions</div>
                )}
              </>)}
            </div>

            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              {!mapsLoading ? (
                <button className="run-btn"
                  onClick={keywordExpand ? runKeywordExpansionScrape : multiQueryMode ? runMultiQueryScrape : runMapsScrape}
                  disabled={!mapsQuery.trim() || !placesKey.trim()} style={{
                  padding: "9px 26px", background: mapsQuery && placesKey ? C.purple : C.surface2,
                  color: mapsQuery && placesKey ? "#000" : C.muted,
                  border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit",
                  fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase",
                  transition: "all .15s", display: "flex", alignItems: "center", gap: 8,
                }}>🔍  {keywordExpand ? "Expand & Search" : multiQueryMode ? "Multi-Search Businesses" : "Search Businesses"}</button>
              ) : (
                <button onClick={stopMapsScrape} style={{
                  padding: "9px 22px", background: C.danger, color: "#fff",
                  border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit",
                  fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase",
                }}>◼  Stop</button>
              )}
              {mapsLoading && (
                <div style={{ fontSize: 11, color: C.dim, display: "flex", alignItems: "center", gap: 6 }}>
                  <Spin /><span style={{ color: C.purple }}>{mapsProgress || "Searching…"}</span>
                </div>
              )}
              {mapsResults.length > 0 && !mapsLoading && (
                <span style={{ fontSize: 11, color: C.dim }}>{mapsResults.length} unique businesses found</span>
              )}
            </div>
          </div>

          {mapsResults.length > 0 && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, overflow: "hidden", marginBottom: 16 }}>
              {/* ── Scrape method selector ── */}
              <div style={{ padding: "10px 16px", borderBottom: `1px solid ${C.border}`, background: C.surface2 }}>
                <div style={{ fontSize: 9, color: C.muted, letterSpacing: ".14em", textTransform: "uppercase", marginBottom: 8 }}>
                  ⚡ Extract Emails Using
                </div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {[
                    { id: "bulk",    label: "⚡ Bulk Scraper",    desc: "Free · fast · best for most sites",           color: C.accent  },
                    { id: "smart",   label: "✦ Smart Scraper",   desc: "AI-powered · handles complex sites",          color: C.purple  },
                    { id: "extract", label: "🔗 Extract Leads",  desc: "Snov/Apollo · database lookup · verified",    color: C.blue    },
                  ].map(m => (
                    <button key={m.id} onClick={() => setMapsScrapeMethod(m.id)} style={{
                      padding: "7px 12px", borderRadius: 5, cursor: "pointer", fontFamily: "inherit",
                      fontSize: 10, letterSpacing: ".05em", transition: "all .15s",
                      background: mapsScrapeMethod === m.id ? `${m.color}18` : "transparent",
                      border: `1px solid ${mapsScrapeMethod === m.id ? m.color : C.border}`,
                      color: mapsScrapeMethod === m.id ? m.color : C.dim,
                    }}>
                      <div style={{ fontWeight: mapsScrapeMethod === m.id ? 500 : 400 }}>{m.label}</div>
                      <div style={{ fontSize: 9, color: mapsScrapeMethod === m.id ? m.color + "aa" : "#2a3a55", marginTop: 2 }}>{m.desc}</div>
                    </button>
                  ))}
                </div>
                {mapsScrapeMethod === "smart" && (
                  <div style={{ marginTop: 10, paddingTop: 10, borderTop: `1px solid ${C.border}` }}>
                    <div style={{ fontSize: 9, color: C.muted, letterSpacing: ".12em", textTransform: "uppercase", marginBottom: 6 }}>
                      Provider
                    </div>
                    <div style={{ display: "flex", gap: 5, flexWrap: "wrap" }}>
                      {[
                        { id: "auto",       label: "Auto",       desc: "Try all"         },
                        { id: "gemini",     label: "Gemini",     desc: "Google AI"       },
                        { id: "groq",       label: "Groq",       desc: "Llama 3.3"       },
                        { id: "scraperapi", label: "ScraperAPI", desc: "JS render"       },
                        { id: "zenrows",    label: "ZenRows",    desc: "Anti-bot"        },
                      ].map(p => {
                        const missingKey =
                          (p.id === "gemini"     && !geminiKey.trim())     ||
                          (p.id === "groq"       && !groqKey.trim())       ||
                          (p.id === "scraperapi" && !scraperApiKey.trim()) ||
                          (p.id === "zenrows"    && !zenRowsKey.trim());
                        return (
                          <button key={p.id} onClick={() => setMapsSmartProvider(p.id)} style={{
                            padding: "5px 10px", borderRadius: 4, cursor: "pointer",
                            fontFamily: "inherit", fontSize: 10, transition: "all .15s",
                            background: mapsSmartProvider === p.id ? `${C.purple}22` : "transparent",
                            border: `1px solid ${mapsSmartProvider === p.id ? C.purple : C.border}`,
                            color: mapsSmartProvider === p.id ? C.purple : missingKey ? C.muted : C.dim,
                            opacity: missingKey && p.id !== "auto" ? 0.5 : 1,
                          }}>
                            <span style={{ fontWeight: mapsSmartProvider === p.id ? 600 : 400 }}>{p.label}</span>
                            <span style={{ fontSize: 8, color: mapsSmartProvider === p.id ? C.purple + "99" : "#2a3a55", marginLeft: 4 }}>{p.desc}</span>
                            {missingKey && p.id !== "auto" && <span style={{ fontSize: 8, color: C.danger, marginLeft: 3 }}>no key</span>}
                          </button>
                        );
                      })}
                    </div>
                    {!geminiKey.trim() && !groqKey.trim() && mapsSmartProvider === "auto" && (
                      <div style={{ marginTop: 6, fontSize: 10, color: C.warn }}>
                        ⚠ Auto mode works best with a Gemini or Groq key — add one in the Smart Scraper tab.
                      </div>
                    )}
                    {mapsSmartProvider !== "auto" && (
                      (mapsSmartProvider === "gemini"     && !geminiKey.trim())     ||
                      (mapsSmartProvider === "groq"       && !groqKey.trim())       ||
                      (mapsSmartProvider === "scraperapi" && !scraperApiKey.trim()) ||
                      (mapsSmartProvider === "zenrows"    && !zenRowsKey.trim())
                    ) && (
                      <div style={{ marginTop: 6, fontSize: 10, color: C.danger }}>
                        ⚠ No API key saved for {mapsSmartProvider} — add it in the Smart Scraper tab first.
                      </div>
                    )}
                  </div>
                )}
                {mapsScrapeMethod === "extract" && !extractSnovKey.trim() && !extractApolloKey.trim() && (
                  <div style={{ marginTop: 8, fontSize: 10, color: C.warn }}>
                    ⚠ Extract Leads needs a Snov.io or Apollo API key — add one in the Extract Leads tab settings.
                  </div>
                )}
              </div>
              {/* ── Toolbar row ── */}
              <div style={{ padding: "10px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10 }}>
                <input type="checkbox" checked={selectedPlaces.size === mapsResults.length} onChange={toggleAll} style={{ accentColor: C.purple, cursor: "pointer" }} />
                <span style={{ fontSize: 11, color: C.dim }}>{selectedPlaces.size} of {mapsResults.length} selected</span>
                <div style={{ marginLeft: "auto", display: "flex", gap: 8, flexWrap: "wrap" }}>
                  <SmBtn color={C.accent} onClick={scrapeEmailsForSelected}>⚡ Scrape Emails</SmBtn>
                  <SmBtn color={C.purple} onClick={importSelectedToLeads}>→ Import to Leads</SmBtn>
                  <SmBtn color={C.dim} onClick={exportMapsCSV}>↓ CSV</SmBtn>
                  <SmBtn color="#1ed98a" onClick={exportMapsXLSX}>↓ XLSX</SmBtn>
                </div>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "32px 1.8fr 1.2fr 1fr 0.8fr 0.6fr", padding: "7px 14px", borderBottom: `1px solid ${C.border}`, fontSize: 9, letterSpacing: ".16em", color: C.muted, textTransform: "uppercase" }}>
                <span /><span>Business Name</span><span>Address</span><span>Website</span><span>Phone</span><span>Rating</span>
              </div>
              <div style={{ maxHeight: 440, overflowY: "auto" }}>
                {mapsResults.map((place, i) => (
                  <div key={place.id} className="place-row fade-row" onClick={() => togglePlace(place.id)} style={{
                    display: "grid", gridTemplateColumns: "32px 1.8fr 1.2fr 1fr 0.8fr 0.6fr",
                    padding: "9px 14px", borderBottom: i < mapsResults.length - 1 ? `1px solid ${C.border}` : "none",
                    alignItems: "center", cursor: "pointer", fontSize: 11,
                    background: selectedPlaces.has(place.id) ? `${C.purple}0a` : i % 2 ? "#0a0f16" : "transparent",
                  }}>
                    <input type="checkbox" checked={selectedPlaces.has(place.id)} onChange={() => togglePlace(place.id)} onClick={e => e.stopPropagation()} style={{ accentColor: C.purple, cursor: "pointer" }} />
                    <div>
                      <span style={{ color: C.accent, fontWeight: 500 }}>{place.name}</span>
                      {place.sourceKeyword && (
                        <div style={{ fontSize: 8, marginTop: 2, color: C.purple + "99", letterSpacing: ".04em" }}>
                          ✦ {place.sourceKeyword}
                          {place.expandedKeyword && place.expandedKeyword !== place.originalKeyword && place.expandedKeyword !== place.sourceKeyword &&
                            <span style={{ color: C.muted }}> · via {place.expandedKeyword}</span>}
                        </div>
                      )}
                    </div>
                    <span style={{ color: C.dim, fontSize: 10 }}>{place.address || "—"}</span>
                    <span>{place.website ? <a href={place.website} target="_blank" rel="noreferrer" onClick={e => e.stopPropagation()} style={{ color: C.blue, fontSize: 10 }}>{place.website.replace(/^https?:\/\/(www\.)?/, "").slice(0, 28)}</a> : <span style={{ color: "#1e2840" }}>—</span>}</span>
                    <span style={{ color: C.text, fontSize: 11 }}>{place.phone || <span style={{ color: "#1e2840" }}>—</span>}</span>
                    <span style={{ color: place.rating >= 4 ? C.accent : place.rating >= 3 ? C.warn : C.danger }}>
                      {place.rating ? `★ ${place.rating}` : "—"}
                      {place.reviews ? <span style={{ color: C.muted, fontSize: 9 }}> ({place.reviews})</span> : ""}
                    </span>
                  </div>
                ))}
              </div>
              <div style={{ padding: "10px 16px", borderTop: `1px solid ${C.border}`, fontSize: 10, color: C.muted }}>
                💡 <strong style={{ color: C.text }}>Tip:</strong> Select businesses with websites → click <strong style={{ color: "#00e87a" }}>⚡ Scrape Emails</strong> to extract all emails automatically — no copy-pasting needed.
              </div>
            </div>
          )}

          {!mapsResults.length && !mapsLoading && (
            <div style={{ textAlign: "center", padding: "52px 0", color: "#1a2435" }}>
              <div style={{ fontSize: 36, marginBottom: 10 }}>🗺</div>
              <div style={{ fontSize: 10, letterSpacing: ".22em", textTransform: "uppercase", marginBottom: 8 }}>Search any business type in any city</div>
              <div style={{ fontSize: 11, color: C.muted }}>e.g. "dental clinics" in "Lagos" · "law firms" in "Nairobi" · "hotels" in "Dubai"</div>
            </div>
          )}
        </>)}

      </div>
    </div>
  );
}

// ── micro components ────────────────────────────────────────────────────────
function Label({ children }) {
  return <div style={{ fontSize: 10, color: C.muted, letterSpacing: ".16em", textTransform: "uppercase", marginBottom: 10 }}>{children}</div>;
}
function SmBtn({ children, onClick, color }) {
  return <button onClick={onClick} style={{ padding: "5px 12px", background: "transparent", color, border: `1px solid ${color}44`, borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontSize: 10, letterSpacing: ".05em", transition: "all .15s" }}>{children}</button>;
}
function IBtn({ children, onClick, title, red }) {
  return <button className="icon-btn" onClick={onClick} title={title} style={{ width: 22, height: 22, padding: 0, display: "flex", alignItems: "center", justifyContent: "center", background: "transparent", border: `1px solid ${C.border}`, borderRadius: 3, color: red ? C.danger : C.muted, cursor: "pointer", fontSize: 10, fontFamily: "inherit", transition: "all .12s" }}>{children}</button>;
}
function Spin() {
  return <span style={{ display: "inline-block", width: 10, height: 10, border: "2px solid #0004", borderTop: "2px solid #000", borderRadius: "50%", animation: "spin .7s linear infinite" }} />;
}
