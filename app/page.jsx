"use client";

import { useEffect, useMemo, useState } from "react";
import * as XLSX from "xlsx";

const DEFAULT_MESSAGE = `Dear Hiring Manager,

I hope this message finds you well. I came across the Java Developer position at your organization and I am excited to express my interest.

With over 3 years of hands-on experience in Java, Spring Boot, Microservices, REST APIs, SQL, AWS, Docker, and RabbitMQ, I bring a strong backend development skill set. Additionally, I have frontend experience with HTML, CSS, and Vaadin, enabling me to contribute across the full stack.

I am available to join immediately and would welcome the opportunity to discuss how my skills align with your team's needs.

Please find my resume here:
https://drive.google.com/file/d/1XlYCB6zAIqecGxJm3CLlxmpZIckvdjOW/view?usp=sharing

Looking forward to hearing from you.

Best regards,
Pranav Singh
+91-8269338614
pranavsinghruhela@gmail.com`;

const STORAGE_CONTACTS = "wa.contacts";
const STORAGE_SENT = "wa.sent";
const STORAGE_MESSAGE = "wa.message";
const STORAGE_TAB = "wa.tab";

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
const NAME_HEADER_KEYWORDS = ["name", "hr", "recruiter", "person", "contact name"];

function normalizePhone(raw) {
  if (raw == null) return null;
  const digits = String(raw).replace(/\D/g, "");
  if (digits.length < 10 || digits.length > 15) return null;
  if (digits.length === 10) return "91" + digits;
  return digits;
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

function extractContacts(rows) {
  if (!rows || rows.length === 0) return [];
  const header = rows[0] || [];
  const phoneIdx = detectColumn(header, PHONE_HEADER_KEYWORDS);
  const nameIdx = detectColumn(header, NAME_HEADER_KEYWORDS);
  const seen = new Set();
  const out = [];

  if (phoneIdx >= 0) {
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i] || [];
      const phone = normalizePhone(row[phoneIdx]);
      if (!phone || seen.has(phone)) continue;
      seen.add(phone);
      const name = nameIdx >= 0 ? String(row[nameIdx] ?? "").trim() : "";
      out.push({ phone, name });
    }
    return out;
  }

  for (const row of rows) {
    if (!Array.isArray(row)) continue;
    for (let c = 0; c < row.length; c++) {
      const phone = normalizePhone(row[c]);
      if (!phone || seen.has(phone)) continue;
      seen.add(phone);
      let name = "";
      for (let n = 0; n < row.length; n++) {
        if (n === c) continue;
        const v = String(row[n] ?? "").trim();
        if (v && !/^\+?\d[\d\s\-]+$/.test(v)) {
          name = v;
          break;
        }
      }
      out.push({ phone, name });
    }
  }
  return out;
}

export default function Page() {
  const [tab, setTab] = useState("bulk");
  const [contacts, setContacts] = useState([]);
  const [sent, setSent] = useState({});
  const [message, setMessage] = useState(DEFAULT_MESSAGE);
  const [search, setSearch] = useState("");
  const [hydrated, setHydrated] = useState(false);

  // single-tab state
  const [singleCountry, setSingleCountry] = useState("91");
  const [singlePhone, setSinglePhone] = useState("");
  const [singleError, setSingleError] = useState("");

  useEffect(() => {
    try {
      const c = localStorage.getItem(STORAGE_CONTACTS);
      const s = localStorage.getItem(STORAGE_SENT);
      const m = localStorage.getItem(STORAGE_MESSAGE);
      const t = localStorage.getItem(STORAGE_TAB);
      if (c) setContacts(JSON.parse(c));
      if (s) setSent(JSON.parse(s));
      if (m) setMessage(m);
      if (t === "single" || t === "bulk") setTab(t);
    } catch {}
    setHydrated(true);
  }, []);

  useEffect(() => {
    if (hydrated) localStorage.setItem(STORAGE_CONTACTS, JSON.stringify(contacts));
  }, [contacts, hydrated]);

  useEffect(() => {
    if (hydrated) localStorage.setItem(STORAGE_SENT, JSON.stringify(sent));
  }, [sent, hydrated]);

  useEffect(() => {
    if (hydrated) localStorage.setItem(STORAGE_MESSAGE, message);
  }, [message, hydrated]);

  useEffect(() => {
    if (hydrated) localStorage.setItem(STORAGE_TAB, tab);
  }, [tab, hydrated]);

  async function handleFile(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const all = [];
    for (const sheetName of wb.SheetNames) {
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
      all.push(...extractContacts(rows));
    }
    const map = new Map();
    for (const item of all) map.set(item.phone, item);
    setContacts(Array.from(map.values()));
    e.target.value = "";
  }

  function openWhatsApp(phone, markSent = true) {
    const url = `https://wa.me/${phone}?text=${encodeURIComponent(message)}`;
    window.open(url, "_blank", "noopener,noreferrer");
    if (markSent) setSent((prev) => ({ ...prev, [phone]: Date.now() }));
  }

  function sendSingle() {
    setSingleError("");
    const digits = singlePhone.replace(/\D/g, "");
    if (digits.length < 6) {
      setSingleError("Please enter a valid phone number.");
      return;
    }
    if (!message.trim()) {
      setSingleError("Message is empty — fill in the template above.");
      return;
    }
    const phone = digits.length === 10 && singleCountry ? singleCountry + digits : digits;
    openWhatsApp(phone, false);
  }

  function clearAll() {
    if (!confirm("Remove all contacts? This cannot be undone.")) return;
    setContacts([]);
    setSent({});
  }

  function resetSent() {
    if (!confirm("Reset all 'sent' marks? Numbers will become clickable again.")) return;
    setSent({});
  }

  const filtered = useMemo(() => {
    const q = search.trim().toLowerCase();
    if (!q) return contacts;
    return contacts.filter(
      (c) => c.phone.includes(q) || (c.name || "").toLowerCase().includes(q)
    );
  }, [contacts, search]);

  const sentCount = useMemo(
    () => contacts.filter((c) => sent[c.phone]).length,
    [contacts, sent]
  );

  return (
    <main className="shell">
      <div className="topbar">
        <div className="brand">W</div>
        <h1>WhatsApp Hiring Sender</h1>
      </div>
      <p className="subtitle">
        Send your hiring message on WhatsApp — pick from an uploaded sheet, or paste a single number.
      </p>

      <div className="panel">
        <h2>Message template</h2>
        <textarea
          value={message}
          onChange={(e) => setMessage(e.target.value)}
          placeholder="Type your hiring message..."
        />
        <p className="muted">Edit freely — saved in your browser. Used by both tabs below.</p>
      </div>

      <div className="tabs" role="tablist">
        <button
          role="tab"
          aria-selected={tab === "bulk"}
          className={`tab ${tab === "bulk" ? "active" : ""}`}
          onClick={() => setTab("bulk")}
        >
          From Sheet
        </button>
        <button
          role="tab"
          aria-selected={tab === "single"}
          className={`tab ${tab === "single" ? "active" : ""}`}
          onClick={() => setTab("single")}
        >
          Single Number
        </button>
      </div>

      {tab === "bulk" && (
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
              Auto-detects headers like <code>name</code>, <code>phone</code>, <code>mobile</code>, <code>whatsapp</code>.
              10-digit numbers are treated as Indian (+91).
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
                placeholder="Search by name or number..."
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
                  const isSent = !!sent[c.phone];
                  return (
                    <button
                      key={c.phone}
                      className={`card ${isSent ? "sent" : ""}`}
                      onClick={() => openWhatsApp(c.phone)}
                      disabled={isSent}
                      type="button"
                    >
                      <div className="name">{c.name || "(no name)"}</div>
                      <div className="number">+{c.phone}</div>
                      {isSent && <span className="badge">SENT</span>}
                    </button>
                  );
                })}
              </div>
            )}
          </div>
        </>
      )}

      {tab === "single" && (
        <div className="panel">
          <h2>Send to a single number</h2>
          <p className="muted" style={{ marginBottom: 12 }}>
            Paste any phone number — uses the message template above.
          </p>

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

          <button className="btn btn-primary btn-block" onClick={sendSingle}>
            Open in WhatsApp
          </button>
          {singleError && <p className="error">{singleError}</p>}
        </div>
      )}
    </main>
  );
}
