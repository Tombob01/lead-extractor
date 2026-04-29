import { useState, useRef, useCallback } from "react";

// ── utils ──────────────────────────────────────────────────────────────────
const readFileAsText = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsText(f); });
const readFileAsBase64 = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result.split(",")[1]); r.onerror = rej; r.readAsDataURL(f); });
const uid = () => Math.random().toString(36).slice(2, 9);
const normalizeEmail = (e) => (e || "").trim().toLowerCase();

const C = {
  bg: "#07090d", surface: "#0d1117", surface2: "#111820", border: "#1a2336",
  accent: "#00e87a", blue: "#29b6f6", purple: "#a78bfa", warn: "#ffb347",
  danger: "#ff5f72", muted: "#3a4a5e", text: "#c9d6e8", dim: "#4a5a72",
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

// ── Claude API helper ──────────────────────────────────────────────────────
const callClaude = async (messages, webSearch = false, system = null) => {
  const body = { model: "claude-sonnet-4-20250514", max_tokens: 1000, messages };
  if (system) body.system = system;
  if (webSearch) body.tools = [{ type: "web_search_20250305", name: "web_search" }];
  const res = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
  });
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data;
};

const extractTextFromResponse = (data) =>
  (data.content || []).map(b => b.text || "").join("").trim();

// ── Gemini API helper ──────────────────────────────────────────────────────
const callGemini = async (apiKey, prompt) => {
  // FIX 1: Use gemini-2.0-flash — gemini-1.5-flash is deprecated
  const res = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`,
    {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 2048 },
      }),
    }
  );
  const data = await res.json();
  if (data.error) throw new Error(data.error.message);
  return data?.candidates?.[0]?.content?.parts?.[0]?.text ?? "";
};

// FIX 2: Gemini cannot visit URLs — fetch page HTML first via CORS proxy,
// then send the page text to Gemini to extract emails from real content.
const geminiExtractEmails = async (apiKey, url) => {
  // Step 1: fetch the actual page HTML via CORS proxy (same proxies as free scraper)
  let pageText = null;
  for (const proxy of CORS_PROXIES) {
    try {
      const proxyUrl = proxy.make(url);
      const res = await fetch(proxyUrl, { signal: AbortSignal.timeout(15000) });
      if (!res.ok) continue;
      let html;
      if (proxy.mode === "json_contents") {
        const d = await res.json();
        html = d?.contents ?? null;
      } else {
        html = await res.text();
      }
      if (html && html.length > 200) {
        // Strip tags and scripts, keep visible text only
        pageText = html
          .replace(/<script[\s\S]*?<\/script>/gi, " ")
          .replace(/<style[\s\S]*?<\/style>/gi, " ")
          .replace(/<[^>]+>/g, " ")
          .replace(/\s{3,}/g, " ")
          .slice(0, 12000); // cap to avoid Gemini token limits
        break;
      }
    } catch (_) { continue; }
  }

  if (!pageText) throw new Error("Could not fetch page (site may block scrapers)");

  // Step 2: send extracted text to Gemini to find emails
  const prompt = `From the following webpage text, extract ALL email addresses.

Return ONLY a JSON array of email strings. Example: ["info@company.com","sales@company.com"]
If no emails found return: []
No explanation, no markdown — just the raw JSON array.

Webpage text:
${pageText}`;

  const raw = await callGemini(apiKey, prompt);
  const clean = raw.replace(/```json|```/g, "").trim();
  const match = clean.match(/\[[\s\S]*\]/);
  if (!match) return [];
  try { return JSON.parse(match[0]).filter(e => typeof e === "string" && e.includes("@")); }
  catch { return []; }
};

// ── FREE bulk email extractor via proxy + regex ────────────────────────────
// Uses a CORS proxy to fetch pages, then regex to extract all emails.
// Entirely free — no AI API needed.
const CORS_PROXIES = [
  { make: (u) => `https://api.allorigins.win/get?url=${encodeURIComponent(u)}`, mode: "json_contents" },
  { make: (u) => `https://corsproxy.io/?url=${encodeURIComponent(u)}`,          mode: "text" },
  { make: (u) => `https://thingproxy.freeboard.io/fetch/${encodeURIComponent(u)}`, mode: "text" },
  { make: (u) => `https://api.codetabs.com/v1/proxy?quest=${encodeURIComponent(u)}`, mode: "text" },
];

const EMAIL_REGEX = /[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g;
const SKIP_EMAILS = /\.(png|jpg|gif|svg|webp|css|js)$/i;
const JUNK_DOMAINS = ["sentry.io","wixpress.com","example.com","domain.com","email.com","yoursite","youremail","user@","name@","test@","noreply","no-reply","donotreply","do-not-reply","support@sentry","privacy@","legal@wix"];

const cleanEmails = (raw, domain) => {
  const unique = [...new Set((raw || []).map(e => e.toLowerCase().trim()))];
  return unique.filter(e => {
    if (SKIP_EMAILS.test(e)) return false;
    if (JUNK_DOMAINS.some(j => e.includes(j))) return false;
    if (e.length > 80) return false;
    return true;
  });
};

const fetchPageEmails = async (url) => {
  let html = null;
  let lastErr = null;

  for (const proxy of CORS_PROXIES) {
    try {
      const proxyUrl = proxy.make(url);
      const res = await fetch(proxyUrl, { signal: AbortSignal.timeout(15000) });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      if (proxy.mode === "json_contents") {
        const data = await res.json();
        html = data?.contents ?? null;
      } else {
        html = await res.text();
      }
      if (html && html.length > 200) break;
      html = null; // too short, try next proxy
    } catch (e) {
      lastErr = e;
      html = null;
    }
  }

  if (!html) throw new Error("Could not fetch page (site may block scrapers)");

  // strip scripts/styles to reduce false positives
  const stripped = html
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<!--[\s\S]*?-->/g, " ");

  const raw = stripped.match(EMAIL_REGEX) || [];
  const domain = (() => { try { return new URL(url).hostname.replace(/^www\./, ""); } catch { return ""; } })();
  return cleanEmails(raw, domain);
};

// ── Maps scraper via Claude web_search ────────────────────────────────────
const scrapeWithClaude = async (query, location, onProgress) => {
  const locationPart = location ? ` in ${location}` : "";
  const searchQuery = `${query}${locationPart}`;
  onProgress(`Searching for "${searchQuery}"…`);
  const systemPrompt = "You are a business data extractor. Use web_search to find real businesses. CRITICAL: return ONLY a raw JSON array, no markdown, no backticks, no explanation. Escape any double-quote characters inside string values as \\\\\". Use null for missing fields. Format: [{\"name\":\"Biz\",\"address\":\"123 St\",\"phone\":\"+1 (555) 000-0000\",\"website\":\"https://example.com\",\"rating\":4.5,\"reviews\":128,\"type\":\"type\"}]";
  const userMessage = `Search for: ${query}${locationPart}\nFind real businesses and return their details as JSON.`;
  const data = await callClaude([{ role: "user", content: userMessage }], true, systemPrompt);
  onProgress("Parsing business data…");
  let finalData = data;
  if (data.stop_reason === "tool_use") {
    const toolUses = data.content.filter(b => b.type === "tool_use");
    const toolResults = toolUses.map(tu => ({ type: "tool_result", tool_use_id: tu.id, content: "Search completed. Please compile results into the JSON array now." }));
    const messages = [{ role: "user", content: userMessage }, { role: "assistant", content: data.content }, { role: "user", content: toolResults }];
    finalData = await callClaude(messages, true, systemPrompt);
  }
  const raw = extractTextFromResponse(finalData).replace(/```json|```/g, "").trim();
  const match = raw.match(/\[[\s\S]*\]/);
  if (!match) throw new Error("No businesses found. Try a different query.");
  let parsed = null;
  try { parsed = JSON.parse(match[0]); } catch (_) {}
  if (!parsed) {
    try {
      const sanitized = match[0].replace(/[\u201C\u201D]/g, '"').replace(/[\u2018\u2019]/g, "'");
      parsed = JSON.parse(sanitized);
    } catch (_) {}
  }
  if (!Array.isArray(parsed) || parsed.length === 0) throw new Error("No businesses found for that search.");
  return parsed.map(p => ({ id: uid(), name: p.name || null, address: p.address || null, phone: p.phone || null, website: p.website || null, rating: typeof p.rating === "number" ? p.rating : null, reviews: typeof p.reviews === "number" ? p.reviews : null, type: p.type || null }));
};

// ══════════════════════════════════════════════════════════════════════════
export default function LeadExtractor() {
  const [tab, setTab] = useState("extract");
  const [mode, setMode] = useState("text");
  const [textInput, setTextInput] = useState("");
  const [urlInput, setUrlInput] = useState("");
  const [file, setFile] = useState(null);
  const [leads, setLeads] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadingMsg, setLoadingMsg] = useState("");
  const [error, setError] = useState("");
  const [editingCell, setEditingCell] = useState(null);
  const [editVal, setEditVal] = useState("");
  const [showDupOnly, setShowDupOnly] = useState(false);
  const [exportDone, setExportDone] = useState(false);
  const [copiedIdx, setCopiedIdx] = useState(null);
  const [crmOpen, setCrmOpen] = useState(false);
  const [crmKey, setCrmKey] = useState("");
  const [crmPushing, setCrmPushing] = useState(false);
  const [crmResults, setCrmResults] = useState(null);

  // maps state
  const [mapsQuery, setMapsQuery] = useState("");
  const [mapsLocation, setMapsLocation] = useState("");
  const [mapsResults, setMapsResults] = useState([]);
  const [mapsLoading, setMapsLoading] = useState(false);
  const [mapsError, setMapsError] = useState("");
  const [mapsProgress, setMapsProgress] = useState("");
  const [selectedPlaces, setSelectedPlaces] = useState(new Set());

  // bulk scraper state
  const [bulkUrls, setBulkUrls] = useState("");
  const [bulkFile, setBulkFile] = useState(null);
  const [bulkResults, setBulkResults] = useState([]); // [{url, emails, status, error}]
  const [bulkRunning, setBulkRunning] = useState(false);
  const [bulkProgress, setBulkProgress] = useState({ done: 0, total: 0 });
  const [bulkStopped, setBulkStopped] = useState(false);
  const [bulkInputMode, setBulkInputMode] = useState("text"); // text | file
  const bulkStopRef = useRef(false);
  const bulkFileRef = useRef();

  // gemini scraper state
  const [geminiKey, setGeminiKey] = useState("");
  const [geminiKeySaved, setGeminiKeySaved] = useState(false);
  const [geminiUrls, setGeminiUrls] = useState("");
  const [geminiFile, setGeminiFile] = useState(null);
  const [geminiInputMode, setGeminiInputMode] = useState("text");
  const [geminiResults, setGeminiResults] = useState([]);
  const [geminiRunning, setGeminiRunning] = useState(false);
  const [geminiProgress, setGeminiProgress] = useState({ done: 0, total: 0 });
  const [geminiStopped, setGeminiStopped] = useState(false);
  const geminiStopRef = useRef(false);
  const geminiFileRef = useRef();

  const fileRef = useRef();
  const editRef = useRef();

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

  // ── Claude extract ────────────────────────────────────────────────────────
  const callClaudeExtract = useCallback(async (messages, webSearch = false) => {
    const data = await callClaude(messages, webSearch);
    const raw = extractTextFromResponse(data).replace(/```json|```/g, "").trim();
    const match = raw.match(/\[[\s\S]*\]/);
    if (!match) throw new Error("No contacts found");
    return JSON.parse(match[0]);
  }, []);

  const prompt = (text) =>
    `Extract ALL people/contacts. Return ONLY a valid JSON array, no markdown. Each object: {name,email,company,role}, null for missing.\n\nText:\n${text}`;

  const extract = async () => {
    setLoading(true); setError("");
    try {
      let result;
      if (mode === "text") { setLoadingMsg("Analyzing text…"); result = await callClaudeExtract([{ role: "user", content: prompt(textInput) }]); }
      else if (mode === "file") {
        const isPDF = file.type === "application/pdf";
        setLoadingMsg(isPDF ? "Reading PDF…" : "Reading file…");
        const messages = isPDF
          ? [{ role: "user", content: [{ type: "document", source: { type: "base64", media_type: "application/pdf", data: await readFileAsBase64(file) } }, { type: "text", text: prompt("(see document)") }] }]
          : [{ role: "user", content: prompt(await readFileAsText(file)) }];
        setLoadingMsg("Extracting…"); result = await callClaudeExtract(messages);
      } else {
        setLoadingMsg("Fetching page…");
        result = await callClaudeExtract([{ role: "user", content: `Fetch ${urlInput.trim()} and extract ALL contacts. Return ONLY JSON array: [{name,email,company,role}], null for missing.` }], true);
      }
      addLeads(result);
    } catch (e) { setError(e.message || "Extraction failed."); }
    finally { setLoading(false); setLoadingMsg(""); }
  };

  const canRun = !loading && (mode === "text" ? textInput.trim() : mode === "file" ? !!file : urlInput.trim());

  // ── cell editing ──────────────────────────────────────────────────────────
  const startEdit = (id, field, val) => { setEditingCell({ id, field }); setEditVal(val ?? ""); setTimeout(() => editRef.current?.focus(), 30); };
  const commitEdit = () => {
    if (!editingCell) return;
    setLeads((prev) => markDuplicates(prev.map((l) => l.id === editingCell.id ? { ...l, [editingCell.field]: editVal || null } : l)));
    setEditingCell(null);
  };

  // ── export ────────────────────────────────────────────────────────────────
  const exportCSV = (rows = displayLeads) => {
    const header = FIELDS.join(",");
    const csv = [header, ...rows.map((l) => FIELDS.map((f) => `"${(l[f] ?? "").toString().replace(/"/g, '""')}"`).join(","))];
    const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([csv.join("\n")], { type: "text/csv" })); a.download = "leads.csv"; a.click();
    setExportDone(true); setTimeout(() => setExportDone(false), 2000);
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

  const parseGeminiUrls = (text) =>
    text.split(/[\n,]+/).map(u => normalizeUrl(u)).filter(Boolean);

  const runGeminiScrape = async () => {
    if (!geminiKey.trim()) return;
    let urls = [];
    if (geminiInputMode === "file" && geminiFile) {
      const text = await readFileAsText(geminiFile);
      urls = text.split(/[\n,]+/).map(u => normalizeUrl(u)).filter(Boolean);
    } else {
      urls = parseGeminiUrls(geminiUrls);
    }
    if (!urls.length) return;

    setGeminiRunning(true);
    setGeminiStopped(false);
    geminiStopRef.current = false;
    setGeminiResults([]);
    setGeminiProgress({ done: 0, total: urls.length });

    for (let i = 0; i < urls.length; i++) {
      if (geminiStopRef.current) { setGeminiStopped(true); break; }
      const url = urls[i];
      setGeminiProgress({ done: i, total: urls.length });
      setGeminiResults(prev => [...prev, { url, emails: [], status: "loading", error: null }]);
      try {
        const emails = await geminiExtractEmails(geminiKey.trim(), url);
        setGeminiResults(prev => prev.map(r => r.url === url
          ? { url, emails, status: emails.length > 0 ? "ok" : "none", error: null }
          : r));
      } catch (e) {
        setGeminiResults(prev => prev.map(r => r.url === url
          ? { url, emails: [], status: "error", error: e.message }
          : r));
      }
      // 1.5s delay to respect Gemini free tier rate limits (15 req/min)
      await new Promise(r => setTimeout(r, 1500));
    }
    setGeminiProgress(p => ({ ...p, done: urls.length }));
    setGeminiRunning(false);
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
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([rows.join("\n")], { type: "text/csv" }));
    a.download = "gemini-emails.csv"; a.click();
  };

  // ── Maps scraper ──────────────────────────────────────────────────────────
  const runMapsScrape = async () => {
    if (!mapsQuery.trim()) return;
    setMapsLoading(true); setMapsError(""); setMapsResults([]); setSelectedPlaces(new Set());
    try {
      const results = await scrapeWithClaude(mapsQuery.trim(), mapsLocation.trim(), setMapsProgress);
      setMapsResults(results);
      setSelectedPlaces(new Set(results.map(r => r.id)));
      setMapsProgress("");
    } catch (e) {
      setMapsError(e.message || "Scrape failed.");
      setMapsProgress("");
    } finally { setMapsLoading(false); }
  };

  const importSelectedToLeads = () => {
    const toImport = mapsResults
      .filter(r => selectedPlaces.has(r.id))
      .map(r => ({ name: null, email: null, company: r.name, role: r.type, _website: r.website, _phone: r.phone, _address: r.address }));
    addLeads(toImport);
    setTab("extract");
  };

  const exportMapsCSV = () => {
    const cols = ["name", "address", "phone", "website", "rating", "reviews", "type"];
    const rows = [cols.join(","), ...mapsResults.filter(r => selectedPlaces.has(r.id)).map(r => cols.map(c => `"${(r[c] ?? "").toString().replace(/"/g, '""')}"`).join(","))];
    const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([rows.join("\n")], { type: "text/csv" })); a.download = "places.csv"; a.click();
  };

  const togglePlace = (id) => setSelectedPlaces(prev => { const s = new Set(prev); s.has(id) ? s.delete(id) : s.add(id); return s; });
  const toggleAll = () => setSelectedPlaces(prev => prev.size === mapsResults.length ? new Set() : new Set(mapsResults.map(r => r.id)));

  // ── Bulk scraper ──────────────────────────────────────────────────────────
  const normalizeUrl = (u) => {
    u = u.trim();
    if (!u) return null;
    if (!u.startsWith("http://") && !u.startsWith("https://")) u = "https://" + u;
    try { new URL(u); return u; } catch { return null; }
  };

  const parseBulkUrls = (text) =>
    text.split(/[\n,]+/).map(u => normalizeUrl(u)).filter(Boolean);

  const runBulkScrape = async () => {
    let urls = [];
    if (bulkInputMode === "file" && bulkFile) {
      const text = await readFileAsText(bulkFile);
      urls = text.split(/[\n,]+/).map(u => normalizeUrl(u)).filter(Boolean);
    } else {
      urls = parseBulkUrls(bulkUrls);
    }
    if (!urls.length) return;

    setBulkRunning(true);
    setBulkStopped(false);
    bulkStopRef.current = false;
    setBulkResults([]);
    setBulkProgress({ done: 0, total: urls.length });

    const results = [];
    for (let i = 0; i < urls.length; i++) {
      if (bulkStopRef.current) { setBulkStopped(true); break; }
      const url = urls[i];
      setBulkProgress({ done: i, total: urls.length });

      // update UI with pending row
      setBulkResults(prev => [...prev, { url, emails: [], status: "loading", error: null }]);

      try {
        const emails = await fetchPageEmails(url);
        results.push({ url, emails, status: emails.length > 0 ? "ok" : "none", error: null });
        setBulkResults(prev => prev.map(r => r.url === url ? { url, emails, status: emails.length > 0 ? "ok" : "none", error: null } : r));
      } catch (e) {
        results.push({ url, emails: [], status: "error", error: e.message });
        setBulkResults(prev => prev.map(r => r.url === url ? { url, emails: [], status: "error", error: e.message } : r));
      }

      // small delay to avoid hammering proxies
      await new Promise(r => setTimeout(r, 300));
    }

    setBulkProgress({ done: urls.length, total: urls.length });
    setBulkRunning(false);
  };

  const stopBulk = () => { bulkStopRef.current = true; };

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
    const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([rows.join("\n")], { type: "text/csv" })); a.download = "bulk-emails.csv"; a.click();
  };

  const totalBulkEmails = bulkResults.reduce((s, r) => s + r.emails.length, 0);
  const bulkOk = bulkResults.filter(r => r.status === "ok").length;
  const bulkFailed = bulkResults.filter(r => r.status === "error").length;

  const dupCount = leads.filter((l) => l._dup).length;
  const displayLeads = showDupOnly ? leads.filter((l) => l._dup) : leads;

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
      `}</style>

      {/* ── top bar ── */}
      <div style={{ borderBottom: `1px solid ${C.border}`, padding: "13px 22px", display: "flex", alignItems: "center", justifyContent: "space-between", background: C.surface }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <span style={{ width: 8, height: 8, borderRadius: "50%", background: C.accent, boxShadow: `0 0 8px ${C.accent}`, display: "inline-block" }} />
          <span style={{ fontSize: 12, fontWeight: 500, letterSpacing: ".18em", color: "#fff", textTransform: "uppercase" }}>Lead Extractor</span>
        </div>
        <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
          {dupCount > 0 && <SmBtn color={C.warn} onClick={() => setShowDupOnly(v => !v)}>{showDupOnly ? "show all" : `${dupCount} dup${dupCount !== 1 ? "s" : ""}`}</SmBtn>}
          {dupCount > 0 && <SmBtn color={C.accent} onClick={deduplicateAll}>deduplicate</SmBtn>}
          {leads.length > 0 && <SmBtn color={C.blue} onClick={() => setCrmOpen(v => !v)}>↑ HubSpot</SmBtn>}
          {leads.length > 0 && <SmBtn color={exportDone ? C.accent : C.dim} onClick={() => exportCSV()}>{exportDone ? "✓ saved" : "↓ CSV"}</SmBtn>}
          {leads.length > 0 && <SmBtn color={C.danger} onClick={() => { setLeads([]); setCrmResults(null); }}>clear</SmBtn>}
        </div>
      </div>

      {/* ── main tabs ── */}
      <div style={{ borderBottom: `1px solid ${C.border}`, padding: "0 22px", background: C.surface, display: "flex", gap: 0 }}>
        {[
          { id: "extract", label: "🔗  Extract Leads" },
          { id: "bulk", label: "⚡  Bulk URL Scraper" },
          { id: "gemini", label: "✦  Gemini Scraper" },
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
          <div style={{ display: "flex", gap: 3, marginBottom: 16 }}>
            {[{ id: "text", label: "✏  Text" }, { id: "file", label: "📄  File" }, { id: "url", label: "🌐  URL" }].map(t => (
              <button key={t.id} onClick={() => { setMode(t.id); setError(""); }} style={{
                padding: "6px 16px", background: mode === t.id ? C.accent : "transparent",
                color: mode === t.id ? "#000" : C.muted, border: `1px solid ${mode === t.id ? C.accent : C.border}`,
                borderRadius: 4, cursor: "pointer", fontSize: 11, fontFamily: "inherit",
                fontWeight: mode === t.id ? 500 : 400, letterSpacing: ".06em", transition: "all .15s",
              }}>{t.label}</button>
            ))}
          </div>

          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            <Label>{mode === "text" ? "Paste text containing contacts" : mode === "file" ? "Upload CSV, TXT, or PDF" : "URL to scan"}</Label>
            {mode === "text" && (
              <textarea value={textInput} onChange={e => setTextInput(e.target.value)}
                placeholder={"John Smith, VP Engineering @ Acme — jsmith@acme.com\nSarah Chen | Product Manager, TechFlow | s.chen@techflow.io"}
                style={{ width: "100%", height: 110, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "10px 14px", resize: "none", lineHeight: 1.75, transition: "border-color .15s" }} />
            )}
            {mode === "file" && (
              <div onClick={() => fileRef.current?.click()} style={{ border: `2px dashed ${file ? C.accent : C.border}`, borderRadius: 6, padding: "26px 20px", textAlign: "center", cursor: "pointer", background: file ? `${C.accent}07` : "transparent", transition: "all .2s" }}>
                <div style={{ fontSize: 26, marginBottom: 7 }}>📂</div>
                <div style={{ fontSize: 12, color: file ? C.accent : C.muted }}>{file ? `✓  ${file.name}  (${(file.size / 1024).toFixed(1)} KB)` : "Click to choose — CSV, TXT, PDF"}</div>
              </div>
            )}
            <input ref={fileRef} type="file" accept=".csv,.txt,.pdf" style={{ display: "none" }} onChange={e => { setFile(e.target.files[0] || null); setError(""); }} />
            {mode === "url" && (
              <input value={urlInput} onChange={e => setUrlInput(e.target.value)} onKeyDown={e => e.key === "Enter" && canRun && extract()}
                placeholder="https://company.com/team"
                style={{ width: "100%", background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 13, padding: "10px 14px", transition: "border-color .15s" }} />
            )}
            {error && <div style={{ marginTop: 10, fontSize: 11, color: "#ff7070", padding: "7px 12px", background: "#ff606012", borderRadius: 4, border: "1px solid #ff606033" }}>⚠  {error}</div>}
            <button className="run-btn" onClick={extract} disabled={!canRun} style={{
              marginTop: 14, padding: "9px 26px", background: canRun ? C.accent : C.surface2, color: canRun ? "#000" : C.muted,
              border: canRun ? "none" : `1px solid ${C.border}`, borderRadius: 4, cursor: canRun ? "pointer" : "not-allowed",
              fontFamily: "inherit", fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase",
              transition: "all .15s", display: "flex", alignItems: "center", gap: 8,
            }}>{loading ? <><Spin />{loadingMsg || "Working…"}</> : "▶  Extract Leads"}</button>
          </div>

          {crmOpen && (
            <div style={{ background: C.surface, border: `1px solid ${C.blue}44`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
                <Label>HubSpot Private App Token</Label>
                <button onClick={() => { setCrmOpen(false); setCrmResults(null); }} style={{ background: "none", border: "none", color: C.dim, cursor: "pointer", fontSize: 14 }}>✕</button>
              </div>
              <div style={{ display: "flex", gap: 8 }}>
                <input type="password" value={crmKey} onChange={e => setCrmKey(e.target.value)} placeholder="pat-na1-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
                  style={{ flex: 1, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "8px 12px" }} />
                <button onClick={pushCRM} disabled={crmPushing || !crmKey.trim()} style={{ padding: "8px 18px", background: C.blue, color: "#000", border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontWeight: 500, fontSize: 11, whiteSpace: "nowrap" }}>
                  {crmPushing ? "Pushing…" : `Push ${leads.filter(l => !l._dup).length} contacts`}
                </button>
              </div>
              <div style={{ fontSize: 10, color: C.muted, marginTop: 7 }}>Settings → Integrations → Private Apps → create app → add <em>crm.objects.contacts.write</em> scope.</div>
              {crmResults && (
                <div style={{ marginTop: 10, maxHeight: 140, overflowY: "auto" }}>
                  {crmResults.map((r, i) => (
                    <div key={i} style={{ fontSize: 11, padding: "4px 0", borderBottom: `1px solid ${C.border}`, display: "flex", gap: 8 }}>
                      <span style={{ color: r.ok ? C.accent : C.danger }}>{r.ok ? "✓" : "✕"}</span>
                      <span style={{ color: C.text }}>{r.lead?.name || r.lead?.email || "—"}</span>
                      {!r.ok && <span style={{ color: C.danger }}>{r.error}</span>}
                    </div>
                  ))}
                  <div style={{ marginTop: 6, fontSize: 11, color: C.dim }}>{crmResults.filter(r => r.ok).length} pushed · {crmResults.filter(r => !r.ok).length} failed</div>
                </div>
              )}
            </div>
          )}

          {leads.length > 0 && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, overflow: "hidden" }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1.5fr 1.2fr 1fr 52px", padding: "8px 14px", borderBottom: `1px solid ${C.border}`, fontSize: 9, letterSpacing: ".18em", color: C.muted, textTransform: "uppercase" }}>
                <span>Name</span><span>Email</span><span>Company</span><span>Role</span><span />
              </div>
              {displayLeads.map((lead, i) => (
                <div key={lead.id} className="hover-row fade-row" style={{ display: "grid", gridTemplateColumns: "1fr 1.5fr 1.2fr 1fr 52px", borderBottom: i < displayLeads.length - 1 ? `1px solid ${C.border}` : "none", alignItems: "stretch", background: lead._dup ? `${C.warn}06` : i % 2 ? "#0a0f16" : "transparent" }}>
                  {FIELDS.map(field => {
                    const isEditing = editingCell?.id === lead.id && editingCell?.field === field;
                    const val = lead[field];
                    return (
                      <div key={field} className="cell-edit" onClick={() => !isEditing && startEdit(lead.id, field, val)}
                        style={{ padding: "9px 14px", fontSize: 12, display: "flex", alignItems: "center", borderRight: `1px solid ${C.border}`, transition: "background .1s" }}>
                        {isEditing ? (
                          <input ref={editRef} value={editVal} onChange={e => setEditVal(e.target.value)}
                            onBlur={commitEdit} onKeyDown={e => { if (e.key === "Enter") commitEdit(); if (e.key === "Escape") setEditingCell(null); }}
                            style={{ width: "100%", background: "#0a1520", border: `1px solid ${C.accent}66`, borderRadius: 3, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "2px 6px" }} />
                        ) : (
                          <span style={{ color: field === "name" ? C.accent : field === "email" ? C.blue : field === "role" ? C.dim : C.text, wordBreak: "break-all" }}>
                            {val && val !== "null" ? val : <span style={{ color: "#1e2840" }}>—</span>}
                          </span>
                        )}
                        {field === "name" && lead._dup && <span className="dup-badge" style={{ marginLeft: 6, flexShrink: 0 }}>dup</span>}
                      </div>
                    );
                  })}
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 3, padding: "0 5px" }}>
                    <IBtn title="Copy" onClick={() => { navigator.clipboard.writeText(FIELDS.map(f => lead[f] ?? "").join(" | ")); setCopiedIdx(lead.id); setTimeout(() => setCopiedIdx(null), 1500); }}>{copiedIdx === lead.id ? "✓" : "⧉"}</IBtn>
                    <IBtn title="Remove" red onClick={() => setLeads(p => markDuplicates(p.filter(l => l.id !== lead.id)))}>✕</IBtn>
                  </div>
                </div>
              ))}
            </div>
          )}

          {leads.length === 0 && !loading && (
            <div style={{ textAlign: "center", padding: "52px 0", color: "#1a2435" }}>
              <div style={{ fontSize: 36, marginBottom: 10 }}>◈</div>
              <div style={{ fontSize: 10, letterSpacing: ".22em", textTransform: "uppercase" }}>No leads yet</div>
            </div>
          )}
        </>)}

        {/* ════════════════ BULK SCRAPER TAB ════════════════ */}
        {tab === "bulk" && (<>

          {/* free badge */}
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

          {/* input mode toggle */}
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

          {/* URL input */}
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
              <Label>Upload CSV or TXT with one URL per line</Label>
              <div onClick={() => bulkFileRef.current?.click()} style={{ border: `2px dashed ${bulkFile ? C.accent : C.border}`, borderRadius: 6, padding: "26px 20px", textAlign: "center", cursor: "pointer", background: bulkFile ? `${C.accent}07` : "transparent", transition: "all .2s" }}>
                <div style={{ fontSize: 26, marginBottom: 7 }}>📂</div>
                <div style={{ fontSize: 12, color: bulkFile ? C.accent : C.muted }}>{bulkFile ? `✓  ${bulkFile.name}  (${(bulkFile.size / 1024).toFixed(1)} KB)` : "Click to choose — CSV or TXT"}</div>
              </div>
              <input ref={bulkFileRef} type="file" accept=".csv,.txt" style={{ display: "none" }} onChange={e => setBulkFile(e.target.files[0] || null)} />
            </>)}

            <div style={{ display: "flex", gap: 8, marginTop: 14, alignItems: "center" }}>
              {!bulkRunning ? (
                <button className="run-btn" onClick={runBulkScrape}
                  disabled={bulkInputMode === "text" ? !parseBulkUrls(bulkUrls).length : !bulkFile}
                  style={{
                    padding: "9px 26px",
                    background: (bulkInputMode === "text" ? parseBulkUrls(bulkUrls).length : bulkFile) ? C.accent : C.surface2,
                    color: (bulkInputMode === "text" ? parseBulkUrls(bulkUrls).length : bulkFile) ? "#000" : C.muted,
                    border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit",
                    fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase",
                    transition: "all .15s", display: "flex", alignItems: "center", gap: 8,
                  }}>▶  Start Bulk Scrape</button>
              ) : (
                <button onClick={stopBulk} style={{ padding: "9px 22px", background: C.danger, color: "#fff", border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase" }}>
                  ◼  Stop
                </button>
              )}

              {bulkRunning && (
                <div style={{ fontSize: 11, color: C.dim, display: "flex", alignItems: "center", gap: 6 }}>
                  <Spin />
                  <span>{bulkProgress.done} / {bulkProgress.total} URLs</span>
                  <span style={{ color: C.accent }}>{totalBulkEmails} emails found</span>
                </div>
              )}
            </div>
          </div>

          {/* progress bar */}
          {(bulkRunning || bulkResults.length > 0) && bulkProgress.total > 0 && (
            <div style={{ marginBottom: 14 }}>
              <div style={{ height: 3, background: C.surface2, borderRadius: 2, overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${(bulkProgress.done / bulkProgress.total) * 100}%`, background: C.accent, transition: "width .3s", borderRadius: 2 }} />
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6, fontSize: 10, color: C.dim }}>
                <span>{bulkStopped ? "⏹ Stopped" : bulkRunning ? "⚡ Running…" : "✓ Complete"}</span>
                <span style={{ color: C.accent }}>{bulkOk} ok</span>
                <span style={{ color: C.warn }}>{bulkResults.filter(r => r.status === "none").length} no emails</span>
                <span style={{ color: C.danger }}>{bulkFailed} failed</span>
                <span style={{ color: C.text, fontWeight: 500 }}>{totalBulkEmails} total emails</span>
              </div>
            </div>
          )}

          {/* action bar */}
          {bulkResults.length > 0 && !bulkRunning && (
            <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
              <button onClick={importBulkToLeads} disabled={!totalBulkEmails} style={{ padding: "7px 18px", background: totalBulkEmails ? C.accent : C.surface2, color: totalBulkEmails ? "#000" : C.muted, border: "none", borderRadius: 4, cursor: totalBulkEmails ? "pointer" : "not-allowed", fontFamily: "inherit", fontWeight: 500, fontSize: 11, letterSpacing: ".07em" }}>
                → Import {totalBulkEmails} emails to Leads
              </button>
              <button onClick={exportBulkCSV} disabled={!totalBulkEmails} style={{ padding: "7px 18px", background: "transparent", color: totalBulkEmails ? C.blue : C.muted, border: `1px solid ${totalBulkEmails ? C.blue + "44" : C.border}`, borderRadius: 4, cursor: totalBulkEmails ? "pointer" : "not-allowed", fontFamily: "inherit", fontSize: 11, letterSpacing: ".07em" }}>
                ↓ Export CSV
              </button>
              <button onClick={() => { setBulkResults([]); setBulkProgress({ done: 0, total: 0 }); }} style={{ padding: "7px 14px", background: "transparent", color: C.danger, border: `1px solid ${C.danger}33`, borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontSize: 11 }}>
                clear
              </button>
            </div>
          )}

          {/* results table */}
          {bulkResults.length > 0 && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, overflow: "hidden" }}>
              <div style={{ display: "grid", gridTemplateColumns: "2fr 3fr 80px", padding: "8px 14px", borderBottom: `1px solid ${C.border}`, fontSize: 9, letterSpacing: ".18em", color: C.muted, textTransform: "uppercase" }}>
                <span>URL</span><span>Emails Found</span><span>Status</span>
              </div>
              <div style={{ maxHeight: 460, overflowY: "auto" }}>
                {bulkResults.map((row, i) => (
                  <div key={i} className="bulk-row fade-row" style={{
                    display: "grid", gridTemplateColumns: "2fr 3fr 80px",
                    padding: "9px 14px", borderBottom: i < bulkResults.length - 1 ? `1px solid ${C.border}` : "none",
                    alignItems: "center", fontSize: 11,
                    background: i % 2 ? "#0a0f16" : "transparent",
                  }}>
                    <span style={{ color: C.dim, fontSize: 10, wordBreak: "break-all", paddingRight: 10 }}>
                      {(() => { try { return new URL(row.url).hostname; } catch { return row.url; } })()}
                    </span>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                      {row.status === "loading" && <span className="loading-dot" style={{ fontSize: 10, color: C.dim }}>scanning…</span>}
                      {row.status === "error" && <span style={{ fontSize: 10, color: C.danger }}>⚠ {row.error?.slice(0, 40)}</span>}
                      {row.status === "none" && <span style={{ fontSize: 10, color: "#1e2840" }}>no emails found</span>}
                      {row.emails.map((e, j) => (
                        <span key={j} style={{ fontSize: 10, color: C.blue, background: `${C.blue}12`, padding: "1px 7px", borderRadius: 3, border: `1px solid ${C.blue}22` }}>{e}</span>
                      ))}
                    </div>
                    <span style={{
                      fontSize: 9, letterSpacing: ".07em", textTransform: "uppercase", textAlign: "center",
                      color: row.status === "ok" ? C.accent : row.status === "loading" ? C.dim : row.status === "none" ? "#1e3050" : C.danger,
                    }}>
                      {row.status === "ok" ? `✓ ${row.emails.length}` : row.status === "loading" ? "…" : row.status === "none" ? "empty" : "error"}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {!bulkResults.length && !bulkRunning && (
            <div style={{ textAlign: "center", padding: "52px 0", color: "#1a2435" }}>
              <div style={{ fontSize: 36, marginBottom: 10 }}>⚡</div>
              <div style={{ fontSize: 10, letterSpacing: ".22em", textTransform: "uppercase", marginBottom: 8 }}>Bulk email extraction</div>
              <div style={{ fontSize: 11, color: C.muted }}>Paste 300+ URLs and extract all emails — free, fast, no limits</div>
            </div>
          )}
        </>)}

        {/* ════════════════ GEMINI SCRAPER TAB ════════════════ */}
        {tab === "gemini" && (<>

          {/* info banner */}
          <div style={{ background: "#4285f410", border: "1px solid #4285f433", borderRadius: 6, padding: "12px 18px", marginBottom: 16, display: "flex", alignItems: "flex-start", gap: 10 }}>
            <span style={{ fontSize: 18, marginTop: 1 }}>✦</span>
            <div>
              <div style={{ fontSize: 11, color: "#4fc3f7", fontWeight: 500, letterSpacing: ".06em", marginBottom: 4 }}>Gemini 1.5 Flash — Free API (1,500 requests/day)</div>
              <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.7 }}>
                Gemini reads each website directly and extracts all emails — bypasses CORS entirely.<br />
                Get a free API key (no credit card) at <a href="https://aistudio.google.com/app/apikey" target="_blank" rel="noreferrer" style={{ color: "#4fc3f7" }}>aistudio.google.com</a> · Resets every 24 hours.
              </div>
            </div>
          </div>

          {/* API key input */}
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            <Label>Gemini API Key</Label>
            <div style={{ display: "flex", gap: 8 }}>
              <input
                type={geminiKeySaved ? "password" : "text"}
                value={geminiKey}
                onChange={e => { setGeminiKey(e.target.value); setGeminiKeySaved(false); }}
                placeholder="AIza..."
                style={{ flex: 1, background: "#070b0f", border: `1px solid ${geminiKeySaved ? C.accent : C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 14px", transition: "border-color .2s" }}
              />
              <button onClick={() => setGeminiKeySaved(true)} disabled={!geminiKey.trim()} style={{
                padding: "9px 18px", background: geminiKey.trim() ? C.accent : C.surface2,
                color: geminiKey.trim() ? "#000" : C.muted, border: "none", borderRadius: 4,
                cursor: geminiKey.trim() ? "pointer" : "not-allowed", fontFamily: "inherit", fontWeight: 500, fontSize: 11,
              }}>{geminiKeySaved ? "✓ Saved" : "Save Key"}</button>
            </div>
            <div style={{ fontSize: 10, color: C.dim, marginTop: 7 }}>Your key is stored in memory only — never sent anywhere except Google's API.</div>
          </div>

          {/* input mode */}
          <div style={{ display: "flex", gap: 3, marginBottom: 12 }}>
            {[{ id: "text", label: "✏  Paste URLs" }, { id: "file", label: "📄  Upload File" }].map(t => (
              <button key={t.id} onClick={() => setGeminiInputMode(t.id)} style={{
                padding: "6px 16px", background: geminiInputMode === t.id ? "#4fc3f7" : "transparent",
                color: geminiInputMode === t.id ? "#000" : C.muted, border: `1px solid ${geminiInputMode === t.id ? "#4fc3f7" : C.border}`,
                borderRadius: 4, cursor: "pointer", fontSize: 11, fontFamily: "inherit",
                fontWeight: geminiInputMode === t.id ? 500 : 400, letterSpacing: ".06em", transition: "all .15s",
              }}>{t.label}</button>
            ))}
          </div>

          {/* URL input */}
          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            {geminiInputMode === "text" ? (<>
              <Label>Paste URLs (one per line, or comma-separated)</Label>
              <textarea value={geminiUrls} onChange={e => setGeminiUrls(e.target.value)}
                placeholder={"https://www.rectorhayden.com\nhttps://www.lexingtonkyhomesearch.com\nhttps://www.bluegrasshomegroup.com"}
                style={{ width: "100%", height: 130, background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 11, padding: "10px 14px", resize: "vertical", lineHeight: 1.8, transition: "border-color .15s" }} />
              <div style={{ fontSize: 10, color: C.dim, marginTop: 6 }}>
                {parseGeminiUrls(geminiUrls).length > 0 && <span style={{ color: "#4fc3f7" }}>{parseGeminiUrls(geminiUrls).length} URLs detected</span>}
              </div>
            </>) : (<>
              <Label>Upload CSV or TXT with one URL per line</Label>
              <div onClick={() => geminiFileRef.current?.click()} style={{ border: `2px dashed ${geminiFile ? "#4fc3f7" : C.border}`, borderRadius: 6, padding: "26px 20px", textAlign: "center", cursor: "pointer", background: geminiFile ? "#4fc3f710" : "transparent", transition: "all .2s" }}>
                <div style={{ fontSize: 26, marginBottom: 7 }}>📂</div>
                <div style={{ fontSize: 12, color: geminiFile ? "#4fc3f7" : C.muted }}>{geminiFile ? `✓  ${geminiFile.name}  (${(geminiFile.size / 1024).toFixed(1)} KB)` : "Click to choose — CSV or TXT"}</div>
              </div>
              <input ref={geminiFileRef} type="file" accept=".csv,.txt" style={{ display: "none" }} onChange={e => setGeminiFile(e.target.files[0] || null)} />
            </>)}

            <div style={{ display: "flex", gap: 8, marginTop: 14, alignItems: "center" }}>
              {!geminiRunning ? (
                <button className="run-btn" onClick={runGeminiScrape}
                  disabled={!geminiKey.trim() || (geminiInputMode === "text" ? !parseGeminiUrls(geminiUrls).length : !geminiFile)}
                  style={{
                    padding: "9px 26px",
                    background: geminiKey.trim() && (geminiInputMode === "text" ? parseGeminiUrls(geminiUrls).length : geminiFile) ? "#4fc3f7" : C.surface2,
                    color: geminiKey.trim() && (geminiInputMode === "text" ? parseGeminiUrls(geminiUrls).length : geminiFile) ? "#000" : C.muted,
                    border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit",
                    fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase",
                    transition: "all .15s", display: "flex", alignItems: "center", gap: 8,
                  }}>{geminiRunning ? <><Spin />Running…</> : "✦  Start Gemini Scrape"}</button>
              ) : (
                <button onClick={() => { geminiStopRef.current = true; }} style={{ padding: "9px 22px", background: C.danger, color: "#fff", border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase" }}>◼  Stop</button>
              )}
              {!geminiKey.trim() && <span style={{ fontSize: 10, color: C.warn }}>⚠ Enter your Gemini API key first</span>}
              {geminiRunning && (
                <div style={{ fontSize: 11, color: C.dim, display: "flex", alignItems: "center", gap: 6 }}>
                  <Spin /><span>{geminiProgress.done} / {geminiProgress.total} URLs</span>
                  <span style={{ color: "#4fc3f7" }}>{geminiTotalEmails} emails found</span>
                </div>
              )}
            </div>
          </div>

          {/* progress bar */}
          {(geminiRunning || geminiResults.length > 0) && geminiProgress.total > 0 && (
            <div style={{ marginBottom: 14 }}>
              <div style={{ height: 3, background: C.surface2, borderRadius: 2, overflow: "hidden" }}>
                <div style={{ height: "100%", width: `${(geminiProgress.done / geminiProgress.total) * 100}%`, background: "#4fc3f7", transition: "width .3s", borderRadius: 2 }} />
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6, fontSize: 10, color: C.dim }}>
                <span>{geminiStopped ? "⏹ Stopped" : geminiRunning ? "✦ Running…" : "✓ Complete"}</span>
                <span style={{ color: C.accent }}>{geminiOk} ok</span>
                <span style={{ color: C.warn }}>{geminiResults.filter(r => r.status === "none").length} no emails</span>
                <span style={{ color: C.danger }}>{geminiFailed} failed</span>
                <span style={{ color: C.text, fontWeight: 500 }}>{geminiTotalEmails} total emails</span>
              </div>
            </div>
          )}

          {/* action bar */}
          {geminiResults.length > 0 && !geminiRunning && (
            <div style={{ display: "flex", gap: 8, marginBottom: 14 }}>
              <button onClick={importGeminiToLeads} disabled={!geminiTotalEmails} style={{ padding: "7px 18px", background: geminiTotalEmails ? "#4fc3f7" : C.surface2, color: geminiTotalEmails ? "#000" : C.muted, border: "none", borderRadius: 4, cursor: geminiTotalEmails ? "pointer" : "not-allowed", fontFamily: "inherit", fontWeight: 500, fontSize: 11, letterSpacing: ".07em" }}>
                → Import {geminiTotalEmails} emails to Leads
              </button>
              <button onClick={exportGeminiCSV} disabled={!geminiTotalEmails} style={{ padding: "7px 18px", background: "transparent", color: geminiTotalEmails ? "#4fc3f7" : C.muted, border: `1px solid ${geminiTotalEmails ? "#4fc3f744" : C.border}`, borderRadius: 4, cursor: geminiTotalEmails ? "pointer" : "not-allowed", fontFamily: "inherit", fontSize: 11, letterSpacing: ".07em" }}>
                ↓ Export CSV
              </button>
              <button onClick={() => { setGeminiResults([]); setGeminiProgress({ done: 0, total: 0 }); }} style={{ padding: "7px 14px", background: "transparent", color: C.danger, border: `1px solid ${C.danger}33`, borderRadius: 4, cursor: "pointer", fontFamily: "inherit", fontSize: 11 }}>clear</button>
            </div>
          )}

          {/* results table */}
          {geminiResults.length > 0 && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, overflow: "hidden" }}>
              <div style={{ display: "grid", gridTemplateColumns: "2fr 3fr 80px", padding: "8px 14px", borderBottom: `1px solid ${C.border}`, fontSize: 9, letterSpacing: ".18em", color: C.muted, textTransform: "uppercase" }}>
                <span>URL</span><span>Emails Found</span><span>Status</span>
              </div>
              <div style={{ maxHeight: 460, overflowY: "auto" }}>
                {geminiResults.map((row, i) => (
                  <div key={i} className="bulk-row fade-row" style={{
                    display: "grid", gridTemplateColumns: "2fr 3fr 80px",
                    padding: "9px 14px", borderBottom: i < geminiResults.length - 1 ? `1px solid ${C.border}` : "none",
                    alignItems: "center", fontSize: 11,
                    background: i % 2 ? "#0a0f16" : "transparent",
                  }}>
                    <span style={{ color: C.dim, fontSize: 10, wordBreak: "break-all", paddingRight: 10 }}>
                      {(() => { try { return new URL(row.url).hostname; } catch { return row.url; } })()}
                    </span>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 4 }}>
                      {row.status === "loading" && <span className="loading-dot" style={{ fontSize: 10, color: C.dim }}>Gemini reading…</span>}
                      {row.status === "error" && <span style={{ fontSize: 10, color: C.danger }}>⚠ {row.error?.slice(0, 50)}</span>}
                      {row.status === "none" && <span style={{ fontSize: 10, color: "#1e2840" }}>no emails found</span>}
                      {row.emails.map((e, j) => (
                        <span key={j} style={{ fontSize: 10, color: "#4fc3f7", background: "#4fc3f712", padding: "1px 7px", borderRadius: 3, border: "1px solid #4fc3f722" }}>{e}</span>
                      ))}
                    </div>
                    <span style={{
                      fontSize: 9, letterSpacing: ".07em", textTransform: "uppercase", textAlign: "center",
                      color: row.status === "ok" ? C.accent : row.status === "loading" ? C.dim : row.status === "none" ? "#1e3050" : C.danger,
                    }}>
                      {row.status === "ok" ? `✓ ${row.emails.length}` : row.status === "loading" ? "…" : row.status === "none" ? "empty" : "error"}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          )}

          {!geminiResults.length && !geminiRunning && (
            <div style={{ textAlign: "center", padding: "52px 0", color: "#1a2435" }}>
              <div style={{ fontSize: 36, marginBottom: 10 }}>✦</div>
              <div style={{ fontSize: 10, letterSpacing: ".22em", textTransform: "uppercase", marginBottom: 8 }}>Gemini-powered email extraction</div>
              <div style={{ fontSize: 11, color: C.muted }}>AI reads each website and finds all emails — 1,500 URLs free per day</div>
            </div>
          )}
        </>)}

        {/* ════════════════ MAPS TAB ════════════════ */}
        {tab === "maps" && (<>
          <div style={{ background: `${C.accent}0a`, border: `1px solid ${C.accent}33`, borderRadius: 6, padding: "12px 18px", marginBottom: 16, display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 18 }}>⚡</span>
            <div>
              <div style={{ fontSize: 11, color: C.accent, fontWeight: 500, letterSpacing: ".06em", marginBottom: 2 }}>No API key required</div>
              <div style={{ fontSize: 10, color: C.dim, lineHeight: 1.6 }}>Uses Claude's built-in web search to find real business listings.</div>
            </div>
          </div>

          <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, padding: 18, marginBottom: 16 }}>
            <Label>Search Query</Label>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 14 }}>
              <div>
                <div style={{ fontSize: 10, color: C.dim, marginBottom: 5, letterSpacing: ".1em" }}>WHAT (business type / keyword)</div>
                <input value={mapsQuery} onChange={e => setMapsQuery(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && !mapsLoading && mapsQuery && runMapsScrape()}
                  placeholder="e.g. dental clinic, law firm, coffee shop"
                  style={{ width: "100%", background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 12px" }} />
              </div>
              <div>
                <div style={{ fontSize: 10, color: C.dim, marginBottom: 5, letterSpacing: ".1em" }}>WHERE (city, area, or leave blank)</div>
                <input value={mapsLocation} onChange={e => setMapsLocation(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && !mapsLoading && mapsQuery && runMapsScrape()}
                  placeholder="e.g. Lagos, Lekki, New York"
                  style={{ width: "100%", background: "#070b0f", border: `1px solid ${C.border}`, borderRadius: 4, color: C.text, fontFamily: "inherit", fontSize: 12, padding: "9px 12px" }} />
              </div>
            </div>

            {mapsError && <div style={{ marginBottom: 12, fontSize: 11, color: "#ff7070", padding: "7px 12px", background: "#ff606012", borderRadius: 4, border: "1px solid #ff606033" }}>⚠  {mapsError}</div>}

            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <button className="run-btn" onClick={runMapsScrape} disabled={mapsLoading || !mapsQuery.trim()} style={{
                padding: "9px 26px", background: !mapsLoading && mapsQuery ? C.purple : C.surface2,
                color: !mapsLoading && mapsQuery ? "#000" : C.muted,
                border: "none", borderRadius: 4, cursor: "pointer", fontFamily: "inherit",
                fontWeight: 500, fontSize: 11, letterSpacing: ".1em", textTransform: "uppercase",
                transition: "all .15s", display: "flex", alignItems: "center", gap: 8,
              }}>{mapsLoading ? <><Spin />{mapsProgress || "Searching…"}</> : "🔍  Search Businesses"}</button>
              {mapsResults.length > 0 && !mapsLoading && (
                <span style={{ fontSize: 11, color: C.dim }}>{mapsResults.length} businesses found</span>
              )}
            </div>
          </div>

          {mapsResults.length > 0 && (
            <div style={{ background: C.surface, border: `1px solid ${C.border}`, borderRadius: 6, overflow: "hidden", marginBottom: 16 }}>
              <div style={{ padding: "10px 16px", borderBottom: `1px solid ${C.border}`, display: "flex", alignItems: "center", gap: 10 }}>
                <input type="checkbox" checked={selectedPlaces.size === mapsResults.length} onChange={toggleAll} style={{ accentColor: C.purple, cursor: "pointer" }} />
                <span style={{ fontSize: 11, color: C.dim }}>{selectedPlaces.size} of {mapsResults.length} selected</span>
                <div style={{ marginLeft: "auto", display: "flex", gap: 8 }}>
                  <SmBtn color={C.purple} onClick={importSelectedToLeads}>→ Import to Leads</SmBtn>
                  <SmBtn color={C.accent} onClick={exportMapsCSV}>↓ Export CSV</SmBtn>
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
                    <span style={{ color: C.accent, fontWeight: 500 }}>{place.name}</span>
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
                💡 <strong style={{ color: C.text }}>Tip:</strong> Import businesses → copy their websites into the Bulk Scraper tab to extract all emails for free.
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
