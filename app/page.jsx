"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const DEFAULT_WA_MESSAGE = `Dear Hiring Manager,

I hope you are doing well.

I wanted to check if there are any openings for a Java Developer role in your organization. I have 3+ years of experience in Java, Spring Boot, Microservices, REST APIs, SQL, AWS, Docker, and RabbitMQ.

I am available to join immediately. Please find my resume below:
https://drive.google.com/file/d/1XlYCB6zAIqecGxJm3CLlxmpZIckvdjOW/view?usp=sharing

Looking forward to hearing from you.

Best regards,
Pranav Singh
+91-8269338614
pranavsinghruhela@gmail.com`;

const DEFAULT_EMAIL_SUBJECT = "Application for Java Developer Role — Pranav Singh (3+ Years Experience)";

const DEFAULT_EMAIL_BODY = `Dear Hiring Manager,

I hope this email finds you well.

I am writing to express my interest in any open Java Developer positions in your organization. I am a backend developer with 3+ years of hands-on experience in Java, Spring Boot, Microservices, REST APIs, SQL, AWS, Docker, and RabbitMQ. I also bring frontend experience with HTML, CSS, and Vaadin, enabling me to contribute across the full stack.

Highlights:
- 3+ years building production Java / Spring Boot services
- Cloud (AWS), containerization (Docker), and async messaging (RabbitMQ)
- Available to join immediately

Please find my resume here:
https://drive.google.com/file/d/1XlYCB6zAIqecGxJm3CLlxmpZIckvdjOW/view?usp=sharing

I would welcome the opportunity to discuss how my skills align with your team's needs. Thank you for your time and consideration.

Best regards,
Pranav Singh
+91-8269338614
pranavsinghruhela@gmail.com`;

const STORAGE = {
  channel: "app.channel",
  mode: "app.mode",
  waContacts: "wa.contacts",
  waSent: "wa.sent",
  waMessage: "wa.message",
  emailContacts: "email.contacts",
  emailSent: "email.sent",
  emailSubject: "email.subject",
  emailBody: "email.body",
};

const COUNTRIES = [
  { code: "91", label: "+91 IN" },
  { code: "1", label: "+1 US" },
  { code: "44", label: "+44 UK" },
  { code: "61", label: "+61 AU" },
  { code: "971", label: "+971 AE" },
  { code: "65", label: "+65 SG" },
  { code: "49", label: "+49 DE" },
  { code: "33", label: "+33 FR" },
];

const PHONE_HEADER_KEYWORDS = ["phone", "mobile", "number", "contact", "whatsapp", "wa", "cell"];
const EMAIL_HEADER_KEYWORDS = ["email", "e-mail", "mail", "id", "address"];
const NAME_HEADER_KEYWORDS = ["name", "hr", "recruiter", "person", "contact name"];

const EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function normalizePhone(raw) {
  if (raw == null) return null;
  const digits = String(raw).replace(/\D/g, "");
  if (digits.length < 10 || digits.length > 15) return null;
  if (digits.length === 10) return "91" + digits;
  return digits;
}

function normalizeEmail(raw) {
  if (raw == null) return null;
  const v = String(raw).trim().toLowerCase();
  return EMAIL_RE.test(v) ? v : null;
}

function detectColumn(headerRow, keywords) {
  if (!Array.isArray(headerRow)) return -1;
  for (let i = 0; i < headerRow.length; i++) {
    const cell = String(headerRow[i] ?? "").toLowerCase().trim();
    if (!cell) continue;
    if (keywords.some((k) => cell.includes(k))) return i;
  }
  return -1;
}

function extractItems(rows, valueKeywords, normalize, key) {
  if (!rows || rows.length === 0) return [];
  const header = rows[0] || [];
  const valIdx = detectColumn(header, valueKeywords);
  const nameIdx = detectColumn(header, NAME_HEADER_KEYWORDS);
  const seen = new Set();
  const out = [];

  if (valIdx >= 0) {
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i] || [];
      const val = normalize(row[valIdx]);
      if (!val || seen.has(val)) continue;
      seen.add(val);
      const name = nameIdx >= 0 ? String(row[nameIdx] ?? "").trim() : "";
      out.push({ [key]: val, name });
    }
    return out;
  }

  // fallback: scan every cell
  for (const row of rows) {
    if (!Array.isArray(row)) continue;
    for (let c = 0; c < row.length; c++) {
      const val = normalize(row[c]);
      if (!val || seen.has(val)) continue;
      seen.add(val);
      let name = "";
      for (let n = 0; n < row.length; n++) {
        if (n === c) continue;
        const v = String(row[n] ?? "").trim();
        if (!v) continue;
        if (key === "phone" && /^\+?\d[\d\s\-]+$/.test(v)) continue;
        if (key === "email" && EMAIL_RE.test(v.toLowerCase())) continue;
        name = v;
        break;
      }
      out.push({ [key]: val, name });
    }
  }
  return out;
}

export default function Page() {
  const [channel, setChannel] = useState("whatsapp");
  const [mode, setMode] = useState("sheet");

  // WhatsApp state
  const [waContacts, setWaContacts] = useState([]);
  const [waSent, setWaSent] = useState({});
  const [waMessage, setWaMessage] = useState(DEFAULT_WA_MESSAGE);

  // Email state
  const [emailContacts, setEmailContacts] = useState([]);
  const [emailSent, setEmailSent] = useState({});
  const [emailSubject, setEmailSubject] = useState(DEFAULT_EMAIL_SUBJECT);
  const [emailBody, setEmailBody] = useState(DEFAULT_EMAIL_BODY);

  // Single inputs
  const [singleCountry, setSingleCountry] = useState("91");
  const [singlePhone, setSinglePhone] = useState("");
  const [singleEmail, setSingleEmail] = useState("");
  const [singleError, setSingleError] = useState("");

  const [search, setSearch] = useState("");
  const [hydrated, setHydrated] = useState(false);

  useEffect(() => {
    try {
      const ch = localStorage.getItem(STORAGE.channel);
      const md = localStorage.getItem(STORAGE.mode);
      const wc = localStorage.getItem(STORAGE.waContacts);
      const ws = localStorage.getItem(STORAGE.waSent);
      const wm = localStorage.getItem(STORAGE.waMessage);
      const ec = localStorage.getItem(STORAGE.emailContacts);
      const es = localStorage.getItem(STORAGE.emailSent);
      const esub = localStorage.getItem(STORAGE.emailSubject);
      const eb = localStorage.getItem(STORAGE.emailBody);
      if (ch === "whatsapp" || ch === "email") setChannel(ch);
      if (md === "sheet" || md === "single") setMode(md);
      if (wc) setWaContacts(JSON.parse(wc));
      if (ws) setWaSent(JSON.parse(ws));
      if (wm) setWaMessage(wm);
      if (ec) setEmailContacts(JSON.parse(ec));
      if (es) setEmailSent(JSON.parse(es));
      if (esub) setEmailSubject(esub);
      if (eb) setEmailBody(eb);
    } catch {}
    setHydrated(true);
  }, []);

  useEffect(() => { if (hydrated) localStorage.setItem(STORAGE.channel, channel); }, [channel, hydrated]);
  useEffect(() => { if (hydrated) localStorage.setItem(STORAGE.mode, mode); }, [mode, hydrated]);
  useEffect(() => { if (hydrated) localStorage.setItem(STORAGE.waContacts, JSON.stringify(waContacts)); }, [waContacts, hydrated]);
  useEffect(() => { if (hydrated) localStorage.setItem(STORAGE.waSent, JSON.stringify(waSent)); }, [waSent, hydrated]);
  useEffect(() => { if (hydrated) localStorage.setItem(STORAGE.waMessage, waMessage); }, [waMessage, hydrated]);
  useEffect(() => { if (hydrated) localStorage.setItem(STORAGE.emailContacts, JSON.stringify(emailContacts)); }, [emailContacts, hydrated]);
  useEffect(() => { if (hydrated) localStorage.setItem(STORAGE.emailSent, JSON.stringify(emailSent)); }, [emailSent, hydrated]);
  useEffect(() => { if (hydrated) localStorage.setItem(STORAGE.emailSubject, emailSubject); }, [emailSubject, hydrated]);
  useEffect(() => { if (hydrated) localStorage.setItem(STORAGE.emailBody, emailBody); }, [emailBody, hydrated]);

  useEffect(() => { setSearch(""); }, [channel, mode]);

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const all = [];
    for (const sheetName of wb.SheetNames) {
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      if (channel === "whatsapp") {
        all.push(...extractItems(rows, PHONE_HEADER_KEYWORDS, normalizePhone, "phone"));
      } else {
        all.push(...extractItems(rows, EMAIL_HEADER_KEYWORDS, normalizeEmail, "email"));
      }
    }
    const map = new Map();
    const key = channel === "whatsapp" ? "phone" : "email";
    for (const item of all) map.set(item[key], item);
    const merged = Array.from(map.values());
    if (channel === "whatsapp") setWaContacts(merged);
    else setEmailContacts(merged);
    e.target.value = "";
  }

  function openContact(contact) {
    if (channel === "whatsapp") {
      const url = `https://wa.me/${contact.phone}?text=${encodeURIComponent(waMessage)}`;
      window.open(url, "_blank", "noopener,noreferrer");
      setWaSent((prev) => ({ ...prev, [contact.phone]: Date.now() }));
    } else {
      const url = `mailto:${contact.email}?subject=${encodeURIComponent(emailSubject)}&body=${encodeURIComponent(emailBody)}`;
      window.location.href = url;
      setEmailSent((prev) => ({ ...prev, [contact.email]: Date.now() }));
    }
  }

  function sendSingle() {
    setSingleError("");
    if (channel === "whatsapp") {
      const digits = singlePhone.replace(/\D/g, "");
      if (digits.length < 6) return setSingleError("Please enter a valid phone number.");
      if (!waMessage.trim()) return setSingleError("Message is empty — fill in the template above.");
      const phone = digits.length === 10 ? singleCountry + digits : digits;
      const url = `https://wa.me/${phone}?text=${encodeURIComponent(waMessage)}`;
      window.open(url, "_blank", "noopener,noreferrer");
    } else {
      const email = normalizeEmail(singleEmail);
      if (!email) return setSingleError("Please enter a valid email address.");
      if (!emailSubject.trim()) return setSingleError("Subject is empty.");
      if (!emailBody.trim()) return setSingleError("Body is empty.");
      const url = `mailto:${email}?subject=${encodeURIComponent(emailSubject)}&body=${encodeURIComponent(emailBody)}`;
      window.location.href = url;
    }
  }

  function clearAll() {
    if (!confirm("Remove all contacts in this channel? This cannot be undone.")) return;
    if (channel === "whatsapp") { setWaContacts([]); setWaSent({}); }
    else { setEmailContacts([]); setEmailSent({}); }
  }

  function resetSent() {
    if (!confirm("Reset 'sent' marks in this channel? Items become clickable again.")) return;
    if (channel === "whatsapp") setWaSent({});
    else setEmailSent({});
  }

  const contacts = channel === "whatsapp" ? waContacts : emailContacts;
  const sentMap = channel === "whatsapp" ? waSent : emailSent;
  const itemKey = channel === "whatsapp" ? "phone" : "email";

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return contacts;
    return contacts.filter(
      (c) => String(c[itemKey]).toLowerCase().includes(q) || (c.name || "").toLowerCase().includes(q)
    );
  }, [contacts, search, itemKey]);

  const sentCount = useMemo(
    () => contacts.filter((c) => sentMap[c[itemKey]]).length,
    [contacts, sentMap, itemKey]
  );

  return (
    <main className="shell">
      <div className="topbar">
        <div className="brand">{channel === "whatsapp" ? "W" : "@"}</div>
        <h1>Hiring Outreach Sender</h1>
      </div>
      <p className="subtitle">
        Send your hiring pitch on WhatsApp or via email — upload a sheet of contacts, or paste one at a time.
      </p>

      <div className="tabs" role="tablist">
        <button
          role="tab"
          aria-selected={channel === "whatsapp"}
          className={`tab ${channel === "whatsapp" ? "active" : ""}`}
          onClick={() => setChannel("whatsapp")}
        >
          WhatsApp
        </button>
        <button
          role="tab"
          aria-selected={channel === "email"}
          className={`tab tab-email ${channel === "email" ? "active" : ""}`}
          onClick={() => setChannel("email")}
        >
          Email
        </button>
      </div>

      <div className="panel">
        <h2>{channel === "whatsapp" ? "WhatsApp message" : "Email subject & body"}</h2>
        {channel === "whatsapp" ? (
          <textarea
            value={waMessage}
            onChange={(e) => setWaMessage(e.target.value)}
            placeholder="Type your WhatsApp message..."
          />
        ) : (
          <>
            <label className="field-label">Subject</label>
            <input
              className="text-input"
              style={{ width: "100%", marginBottom: 12 }}
              value={emailSubject}
              onChange={(e) => setEmailSubject(e.target.value)}
              placeholder="Email subject..."
            />
            <label className="field-label">Body</label>
            <textarea
              value={emailBody}
              onChange={(e) => setEmailBody(e.target.value)}
              placeholder="Email body..."
            />
          </>
        )}
        <p className="muted">Edit freely — saved in your browser.</p>
      </div>

      <div className="tabs sub-tabs" role="tablist">
        <button
          role="tab"
          aria-selected={mode === "sheet"}
          className={`tab ${mode === "sheet" ? "active" : ""}`}
          onClick={() => setMode("sheet")}
        >
          From Sheet
        </button>
        <button
          role="tab"
          aria-selected={mode === "single"}
          className={`tab ${mode === "single" ? "active" : ""}`}
          onClick={() => setMode("single")}
        >
          {channel === "whatsapp" ? "Single Number" : "Single Email"}
        </button>
      </div>

      {mode === "sheet" && (
        <>
          <div className="panel">
            <h2>Upload sheet</h2>
            <div className="row">
              <label className="upload-label" htmlFor="file">Choose Excel / CSV</label>
              <input id="file" type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} />
              <button className="btn btn-ghost" onClick={resetSent} disabled={!sentCount}>
                Reset sent marks
              </button>
              <button className="btn btn-danger" onClick={clearAll} disabled={!contacts.length}>
                Clear all
              </button>
            </div>
            <p className="muted">
              {channel === "whatsapp"
                ? <>Auto-detects headers like <code>name</code>, <code>phone</code>, <code>mobile</code>. 10-digit numbers default to +91.</>
                : <>Auto-detects headers like <code>name</code>, <code>email</code>. Falls back to scanning every cell for an email pattern.</>}
            </p>
          </div>

          <div className="panel">
            <div className="panel-head">
              <h2 style={{ margin: 0 }}>Contacts</h2>
              <div className="stats">
                <div>Total: <span className="stat-num">{contacts.length}</span></div>
                <div>Sent: <span className="stat-num">{sentCount}</span></div>
                <div>Remaining: <span className="stat-num">{contacts.length - sentCount}</span></div>
              </div>
            </div>

            {contacts.length > 0 && (
              <input
                className="search"
                placeholder={`Search by name or ${channel === "whatsapp" ? "number" : "email"}...`}
                value={search}
                onChange={(e) => setSearch(e.target.value)}
              />
            )}

            {contacts.length === 0 ? (
              <div className="empty">No contacts yet. Upload an Excel or CSV file above.</div>
            ) : filtered.length === 0 ? (
              <div className="empty">No matches for "{search}".</div>
            ) : (
              <div className="grid">
                {filtered.map((c) => {
                  const id = c[itemKey];
                  const isSent = !!sentMap[id];
                  return (
                    <button
                      key={id}
                      className={`card ${isSent ? "sent" : ""}`}
                      onClick={() => openContact(c)}
                      disabled={isSent}
                      type="button"
                    >
                      <div className="name">{c.name || "(no name)"}</div>
                      <div className="number">{channel === "whatsapp" ? `+${id}` : id}</div>
                      {isSent && <span className="badge">SENT</span>}
                    </button>
                  );
                })}
              </div>
            )}
          </div>
        </>
      )}

      {mode === "single" && (
        <div className="panel">
          <h2>{channel === "whatsapp" ? "Send to a single number" : "Send to a single email"}</h2>
          <p className="muted" style={{ marginBottom: 12 }}>
            {channel === "whatsapp"
              ? "Paste any phone number — uses the WhatsApp message above."
              : "Paste any email address — opens your default mail app with the subject & body above."}
          </p>

          {channel === "whatsapp" ? (
            <>
              <label className="field-label">Phone number</label>
              <div className="phone-row">
                <select
                  className="select"
                  value={singleCountry}
                  onChange={(e) => setSingleCountry(e.target.value)}
                >
                  {COUNTRIES.map((c) => (
                    <option key={c.code} value={c.code}>{c.label}</option>
                  ))}
                </select>
                <input
                  type="tel"
                  className="text-input"
                  placeholder="9876543210"
                  inputMode="numeric"
                  value={singlePhone}
                  onChange={(e) => setSinglePhone(e.target.value)}
                  onKeyDown={(e) => { if (e.key === "Enter") sendSingle(); }}
                />
              </div>
              <p className="muted">Country code is added only if you enter exactly 10 digits.</p>
            </>
          ) : (
            <>
              <label className="field-label">Email address</label>
              <input
                type="email"
                className="text-input"
                style={{ width: "100%" }}
                placeholder="hr@company.com"
                value={singleEmail}
                onChange={(e) => setSingleEmail(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter") sendSingle(); }}
              />
            </>
          )}

          <button className="btn btn-primary btn-block" onClick={sendSingle}>
            {channel === "whatsapp" ? "Open in WhatsApp" : "Open in Mail App"}
          </button>
          {singleError && <p className="error">{singleError}</p>}
        </div>
      )}
    </main>
  );
}
