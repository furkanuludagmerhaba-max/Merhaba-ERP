import { useState, useEffect, useMemo, useCallback, useRef, Fragment } from "react";
import * as XLSX from "xlsx";
import { supabase } from "./supabase";
import { CATALOG_SEED, PROD, RECIPES } from "./data";

/* ============================ veri ============================ */


const DEPTS_SEED = [
  { id: "pasta", name: "Pasta" }, { id: "sutlu-tatli", name: "Sütlü Tatlı" },
  { id: "ekler", name: "Ekler" }, { id: "ozel-pasta", name: "Özel Pasta" },
  { id: "baklava", name: "Baklava" }, { id: "hamur", name: "Hamur" },
  { id: "kuru-pasta", name: "Kuru Pasta" }, { id: "tek-pasta", name: "Tek Pasta" },
  { id: "sandvic", name: "Sandviç" }, { id: "depo", name: "Depo" },
];
const URETIM_DEPTS = DEPTS_SEED.filter((d) => d.id !== "depo").map((d) => d.id);

const BRANCHES_SEED = [
  { id: "cevizlibag", name: "Cevizlibağ" }, { id: "sirkeci", name: "Sirkeci" },
  { id: "musiad", name: "Musiad" }, { id: "merter", name: "Merter" },
  { id: "besiktas", name: "Beşiktaş" }, { id: "basaksehir", name: "Başakşehir" },
  { id: "mavera", name: "Mavera" },
];

const CFG_V = 4;
/* ilk kurulumdan sonra eklenen yetkiler — kayıtlı hesaplara bir kez uygulanır */
const LATER_PERMS = ["stock.view", "stock.count", "purchase.view", "purchase.request", "purchase.buy", "supply.request"];

function migrateCfg(c) {
  if ((c.v || 1) >= CFG_V) return false;
  const have = new Set((c.users || []).map((u) => lower(u.username)));
  USERS_SEED.forEach((u) => { if (!have.has(lower(u.username))) c.users.push({ ...u }); });
  c.users = c.users.map((u) => {
    const def = ROLES[u.role]?.perms || [];
    const add = LATER_PERMS.filter((p) => def.includes(p) && !(u.perms || []).includes(p));
    return add.length ? { ...u, perms: [...(u.perms || []), ...add] } : u;
  });
  c.v = CFG_V;
  return true;
}

const SETTINGS_SEED = { leadProd: 1, leadShip: 1, openHour: 6, closeHour: 24 };

const FIRE_REASONS = ["Kırık / bozuk", "Tarihi geçmiş", "Şubeden iade", "Sevkiyatta döküldü", "Üretim hatası", "Personel / ikram", "Diğer"];
const REQ_STATUS = {
  talep: { label: "Talep açıldı", cls: "bg-stone-200 text-stone-700" },
  siparis: { label: "Sipariş verildi", cls: "bg-sky-200 text-sky-900" },
  kismi: { label: "Kısmen geldi", cls: "bg-orange-200 text-orange-900" },
  teslim: { label: "Teslim alındı", cls: "bg-emerald-200 text-emerald-900" },
  iptal: { label: "İptal", cls: "bg-stone-200 text-stone-500" },
};
const URGENCY = ["Normal", "Acil", "Çok acil"];

const SHORT_REASONS = ["Üretim yetişmedi", "Fire çıktı", "Hammadde bitti", "Sevkiyatta kaldı", "Şube iptal etti", "Diğer"];

const PERMS = [
  ["order.create", "Talep girer (yalnızca şube hesapları)"],
  ["order.view.own", "Kendi şubesinin taleplerini görür"],
  ["dept.queue", "Bölüm taleplerini görür, hazırlanan adedi girer"],
  ["production.entry", "Günlük üretim adedi girer"],
  ["fire.entry", "Fire girer"],
  ["shipping", "Sevk ekranını kullanır, gönderilen adedi girer"],
  ["stock.view", "Depo stok listesini görür"],
  ["stock.count", "Depo sayımı yapar, stok düzeltir"],
  ["purchase.view", "Satınalma ekranını görür"],
  ["purchase.request", "Satınalmaya talep açar, gelen malı teslim alır"],
  ["purchase.buy", "Talepleri sipariş eder, tedarikçi ve termin girer"],
  ["supply.request", "Bölümü adına depodan malzeme ister"],
  ["list.view", "Şube taleplerini yalnızca görüntüler"],
  ["report.dept", "Kendi bölümünün raporlarını görür"],
  ["report.all", "Bütün raporları görür"],
  ["admin", "Yönetim panelini kullanır"],
];

const ROLES = {
  admin: { name: "Admin", perms: ["dept.queue", "shipping", "list.view", "stock.view", "stock.count", "purchase.view", "purchase.request", "purchase.buy", "supply.request", "report.all", "admin"] },
  sube: { name: "Şube", perms: ["order.create", "order.view.own"] },
  sef: { name: "Bölüm Şefi", perms: ["dept.queue", "production.entry", "fire.entry", "supply.request"] },
  depo: { name: "Depocu", perms: ["dept.queue", "shipping", "stock.view", "stock.count", "purchase.view", "purchase.request"] },
  uretim_muduru: { name: "Üretim Müdürü", perms: ["dept.queue", "production.entry", "fire.entry", "supply.request", "stock.view", "report.all"] },
  genel_mudur: { name: "Genel Müdür", perms: ["dept.queue", "list.view", "purchase.view", "report.all"] },
  satinalma: { name: "Satınalma", perms: ["purchase.view", "purchase.buy", "stock.view"] },
  operasyon: { name: "Operasyon", perms: ["dept.queue", "list.view", "report.all"] },
  sevkiyat: { name: "Sevkiyat", perms: ["shipping", "fire.entry", "report.all"] },
  sofor: { name: "Şoför", perms: ["list.view"] },
};

const STATUS = {
  bekliyor: { label: "Bekliyor", cls: "bg-stone-200 text-stone-700" },
  hazirlaniyor: { label: "Hazırlanıyor", cls: "bg-amber-200 text-amber-900" },
  hazir: { label: "Hazır", cls: "bg-sky-200 text-sky-900" },
  kismen: { label: "Kısmen gönderildi", cls: "bg-orange-200 text-orange-900" },
  gonderildi: { label: "Gönderildi", cls: "bg-emerald-200 text-emerald-900" },
};

const USERS_SEED = [
  { id: "u1", name: "Sistem Yöneticisi", username: "admin", pin: "1234", role: "admin", branchId: null, depts: [] },
  ...BRANCHES_SEED.map((b, i) => ({ id: "b" + i, name: b.name + " Şube", username: "sube." + b.id, pin: "1111", role: "sube", branchId: b.id, depts: [] })),
  { id: "u4", name: "Pasta Şefi", username: "sef.pasta", pin: "2222", role: "sef", branchId: null, depts: ["pasta"] },
  { id: "u5", name: "Sütlü Tatlı Şefi", username: "sef.sutlu", pin: "2222", role: "sef", branchId: null, depts: ["sutlu-tatli"] },
  { id: "u6", name: "Ekler Şefi", username: "sef.ekler", pin: "2222", role: "sef", branchId: null, depts: ["ekler"] },
  { id: "u7", name: "Özel Pasta Şefi", username: "sef.ozel", pin: "2222", role: "sef", branchId: null, depts: ["ozel-pasta"] },
  { id: "u8", name: "Baklava Şefi", username: "sef.baklava", pin: "2222", role: "sef", branchId: null, depts: ["baklava"] },
  { id: "u9", name: "Hamur Şefi", username: "sef.hamur", pin: "2222", role: "sef", branchId: null, depts: ["hamur"] },
  { id: "u10", name: "Kuru Pasta Şefi", username: "sef.kuru", pin: "2222", role: "sef", branchId: null, depts: ["kuru-pasta"] },
  { id: "u11", name: "Tek Pasta Şefi", username: "sef.tek", pin: "2222", role: "sef", branchId: null, depts: ["tek-pasta"] },
  { id: "u18", name: "Sandviç Şefi", username: "sef.sandvic", pin: "2222", role: "sef", branchId: null, depts: ["sandvic"] },
  { id: "u12", name: "Depo", username: "depo", pin: "3333", role: "depo", branchId: null, depts: ["depo"] },
  { id: "u13", name: "Üretim Müdürü", username: "uretim", pin: "4444", role: "uretim_muduru", branchId: null, depts: URETIM_DEPTS },
  { id: "u14", name: "Genel Müdür", username: "genelmudur", pin: "4444", role: "genel_mudur", branchId: null, depts: [] },
  { id: "u15", name: "Operasyon", username: "operasyon", pin: "4444", role: "operasyon", branchId: null, depts: [] },
  { id: "u16", name: "Sevkiyat", username: "sevkiyat", pin: "5555", role: "sevkiyat", branchId: null, depts: [] },
  { id: "u17", name: "Şoför", username: "sofor", pin: "5555", role: "sofor", branchId: null, depts: [] },
  { id: "u19", name: "Satınalma", username: "satinalma", pin: "6666", role: "satinalma", branchId: null, depts: [] },
].map((u) => ({ ...u, perms: ROLES[u.role].perms.slice() }));

/* geçmiş üretim/fire kayıtları yalnızca raporlar açıldığında kurulur */
let _hist = null;
function histData() {
  if (_hist) return _hist;
  const prod = [], fire = [];
  PROD.hist.forEach(([ii, di, q], k) => {
    const it = PROD.items[ii];
    const rec = { id: "h" + k, date: PROD.dates[di], deptId: it.d, itemId: it.id, name: it.n, unit: it.u, kind: it.k, qty: q, by: "Geçmiş kayıt", seed: true };
    if (it.k === "mamul") prod.push(rec); else fire.push({ ...rec, reason: "Geçmiş fire kaydı" });
  });
  _hist = { prod, fire };
  return _hist;
}

/* ============================ tarih ============================ */

const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
const today = () => iso(new Date());
const addDays = (s, n) => { const [y, m, d] = s.split("-").map(Number); const x = new Date(y, m - 1, d + n); return iso(x); };
const daysAgo = (n) => addDays(today(), -n);
const monthStart = () => today().slice(0, 8) + "01";
const monthKey = (s) => s.slice(0, 7);
const dateTR = (s) => (s ? s.split("-").reverse().join(".") : "");
const inRange = (d, from, to) => (!from || d >= from) && (!to || d <= to);
function rangeMonths(from, to) {
  const out = [];
  let c = from.slice(0, 7);
  const end = to.slice(0, 7);
  let guard = 0;
  while (c <= end && guard++ < 120) {
    out.push(c);
    const [y, m] = c.split("-").map(Number);
    c = m === 12 ? `${y + 1}-01` : `${y}-${String(m + 1).padStart(2, "0")}`;
  }
  return out;
}
function rangeDates(from, to) {
  const out = [];
  let c = from, guard = 0;
  while (c <= to && guard++ < 800) { out.push(c); c = addDays(c, 1); }
  return out;
}

/* ============================ depolama ============================ */

const NS = "erp7";
const kCfg = `${NS}:cfg`;
const kSes = `${NS}:session`;
const kOrd = (d) => `${NS}:ord:${d}`;
const kPrd = (m) => `${NS}:prd:${m}`;
const kFire = (m) => `${NS}:fire:${m}`;
const kReq = (m) => `${NS}:req:${m}`;      /* satınalma talepleri, ay bazlı */
const kLevel = `${NS}:stklevel`;           /* güncel stok seviyeleri */
const kMove = (m) => `${NS}:stkmv:${m}`;   /* stok hareket günlüğü */
const KEY_RE = new RegExp("^" + NS + ":(ord|prd|fire|req|stkmv):(.+)$");

const STORE_TIMEOUT = 10000;
let storageDown = false;

/* Kalıcılık katmanı:
   - Supabase (ortak veritabanı) — ayarlıysa bütün cihazlar aynı veriyi görür.
   - Supabase ayarlı değilse localStorage — yalnız bu cihazda kalıcı.
   - O da yoksa bellek — sayfa yenilenince silinir.
   Bellek önbelleği her zaman doğruluk kaynağıdır: arka uç yazamasa bile
   girilen veri ekranda kaybolmaz, kullanıcı uyarı bandını görür. */
const cache = new Map();
const revs = new Map();            /* her anahtarın sunucudaki sürüm numarası */
let backend = null;
let sb = null;
let channel = null;
const listeners = new Set();

function localOk() {
  try {
    const k = "__erp_probe__";
    window.localStorage.setItem(k, "1");
    window.localStorage.removeItem(k);
    return true;
  } catch (e) { return false; }
}

async function initStore() {
  if (backend) return backend;
  if (supabase) { sb = supabase; backend = "supabase"; startRealtime(); return backend; }
  backend = localOk() ? "local" : "memory";
  if (backend === "memory") storageDown = true;
  return backend;
}
const storeMode = () => backend || "hazırlanıyor";
const storeOk = () => !storageDown;

/* başka bir cihaz yazdığında haber ver */
function subscribeKeys(fn) {
  listeners.add(fn);
  return () => listeners.delete(fn);
}
function startRealtime() {
  if (channel || !sb) return;
  try {
    channel = sb
      .channel("erp-kv")
      .on("postgres_changes", { event: "*", schema: "public", table: "kv" }, (p) => {
        const row = p.new && p.new.key ? p.new : p.old;
        const key = row && row.key;
        if (!key) return;
        if (p.new && p.new.rev !== undefined && revs.get(key) === p.new.rev) return; /* kendi yazdığımız */
        listeners.forEach((fn) => { try { fn(key); } catch (e) { /* yok sayılır */ } });
      })
      .subscribe();
  } catch (e) { /* canlı yenileme açılamadıysa program yine çalışır */ }
}

/* hiçbir depolama çağrısı sonsuza kadar askıda kalmaz */
function withTimeout(run, ms, fallback) {
  return new Promise((resolve) => {
    let done = false;
    const finish = (v) => { if (!done) { done = true; clearTimeout(t); resolve(v); } };
    const t = setTimeout(() => { finish(fallback); }, ms);
    let p;
    try { p = Promise.resolve(run()); } catch (e) { return finish(fallback); }
    p.then(finish, () => finish(fallback));
  });
}

/* Supabase'den tek anahtarı okur. { ok, value, rev } döndürür. */
async function readRow(key) {
  const res = await withTimeout(
    () => sb.from("kv").select("value,rev").eq("key", key).maybeSingle(),
    STORE_TIMEOUT, null
  );
  if (!res || res.error) return { ok: false };
  if (!res.data) return { ok: true, value: undefined, rev: 0 };
  return { ok: true, value: res.data.value, rev: res.data.rev };
}

async function loadKey(key, fallback) {
  await initStore();
  try {
    if (backend === "supabase") {
      const r = await readRow(key);
      if (r.ok) {
        storageDown = false;
        if (r.value !== undefined) { cache.set(key, r.value); revs.set(key, r.rev); }
        else { revs.set(key, 0); if (!cache.has(key)) cache.set(key, fallback); }
      } else {
        storageDown = true;
        if (!cache.has(key)) cache.set(key, fallback);
      }
    } else if (backend === "local") {
      const raw = window.localStorage.getItem(key);
      if (raw !== null) cache.set(key, JSON.parse(raw));
      else if (!cache.has(key)) cache.set(key, fallback);
    } else if (!cache.has(key)) cache.set(key, fallback);
  } catch (e) {
    if (!cache.has(key)) cache.set(key, fallback);
  }
  return cache.has(key) ? cache.get(key) : fallback;
}

/* Belgenin tamamını yazar (ayarlar gibi tek yazarlı kayıtlar için). */
async function saveKey(key, value) {
  cache.set(key, value);              /* önce bellek: ekrandaki veri asla kaybolmaz */
  await initStore();
  try {
    if (backend === "supabase") {
      const rev = (revs.get(key) || 0) + 1;
      const res = await withTimeout(
        () => sb.from("kv").upsert({ key, value, rev, updated_at: new Date().toISOString() }, { onConflict: "key" }).select("rev").maybeSingle(),
        STORE_TIMEOUT, null
      );
      if (res && !res.error) { revs.set(key, res.data ? res.data.rev : rev); storageDown = false; return true; }
      storageDown = true; return false;
    }
    if (backend === "local") { window.localStorage.setItem(key, JSON.stringify(value)); return true; }
  } catch (e) { storageDown = true; return false; }
  return false;
}

async function removeKey(key) {
  cache.delete(key);
  revs.delete(key);
  await initStore();
  try {
    if (backend === "supabase") { await withTimeout(() => sb.from("kv").delete().eq("key", key), STORE_TIMEOUT, null); return; }
    if (backend === "local") window.localStorage.removeItem(key);
  } catch (e) { /* yok sayılır */ }
}

/* oturum bilgisi kişiye özeldir: ortak veritabanına yazılmaz */
function loadLocal(key, fallback) {
  try { const r = window.localStorage.getItem(key); return r === null ? fallback : JSON.parse(r); } catch (e) { return cache.has("L" + key) ? cache.get("L" + key) : fallback; }
}
function saveLocal(key, value) {
  cache.set("L" + key, value);
  try { window.localStorage.setItem(key, JSON.stringify(value)); } catch (e) { /* bellek yeter */ }
}
function removeLocal(key) {
  cache.delete("L" + key);
  try { window.localStorage.removeItem(key); } catch (e) { /* yok sayılır */ }
}

/* aynı anda en fazla `limit` istek */
async function mapLimit(items, limit, fn) {
  const out = new Array(items.length);
  let i = 0;
  await Promise.all(new Array(Math.min(limit, items.length)).fill(0).map(async () => {
    while (i < items.length) { const k = i++; out[k] = await fn(items[k]); }
  }));
  return out;
}

/* oku → birleştir → yaz. İki şube aynı anda aynı güne yazarsa biri diğerinin
   üstüne yazmasın diye sürüm numarası karşılaştırılır; çakışırsa yeniden okunup
   tekrar denenir. Yazdığımız değeri döndürür. */
async function mutate(key, fn) {
  await initStore();
  if (backend !== "supabase") {
    const cur = await loadKey(key, []);
    const next = fn(cur === undefined || cur === null ? [] : cur);
    await saveKey(key, next);
    return next;
  }
  let next = null;
  for (let i = 0; i < 6; i++) {
    const r = await readRow(key);
    if (!r.ok) { await wait(200 + i * 250); continue; }        /* ağ hatası: tekrar dene */
    const cur = r.value === undefined || r.value === null ? [] : r.value;
    cache.set(key, cur); revs.set(key, r.rev);
    next = fn(cur);
    if (r.rev === 0) {
      const res = await withTimeout(
        () => sb.from("kv").insert({ key, value: next, rev: 1 }).select("rev").maybeSingle(),
        STORE_TIMEOUT, null
      );
      if (res && !res.error) { cache.set(key, next); revs.set(key, 1); storageDown = false; return next; }
    } else {
      const res = await withTimeout(
        () => sb.from("kv").update({ value: next, rev: r.rev + 1, updated_at: new Date().toISOString() }).eq("key", key).eq("rev", r.rev).select("rev"),
        STORE_TIMEOUT, null
      );
      if (res && !res.error && Array.isArray(res.data) && res.data.length) {
        cache.set(key, next); revs.set(key, r.rev + 1); storageDown = false; return next;
      }
    }
    await wait(150 + i * 200);                                  /* çakışma: baştan oku */
  }
  storageDown = true;
  if (next !== null) cache.set(key, next);
  return next !== null ? next : await loadKey(key, []);
}
const wait = (ms) => new Promise((r) => setTimeout(r, ms));

/* ============================ yardımcılar ============================ */

const uid = () => Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-4);
const fmt = (n) => { const v = Number(n) || 0; return Math.abs(v % 1) > 0.001 ? v.toLocaleString("tr-TR", { maximumFractionDigits: 2 }) : Math.round(v).toLocaleString("tr-TR"); };
const money = (n) => (Number(n) || 0).toLocaleString("tr-TR", { maximumFractionDigits: 0 }) + " ₺";
const lineStatus = (l) => l.status || "bekliyor";
const prep = (l) => Number(l.prepared) || 0;
const sentQ = (l) => Number(l.sent) || 0;
const openQty = (l) => Math.max(0, (l.qty || 0) - sentQ(l));
const lower = (s) => String(s).toLocaleLowerCase("tr");

/* hızlı arama indeksi — her satırda dizi taraması yapılmaz */
/* an: şubenin not yazabilmesi · ag: görsel ekleyebilmesi
   vi: bölümlerin depodan isteyebilmesi · dd: hangi bölümler isteyebilir, boşsa hepsi */
const defAn = (c) => (c.d === "ozel-pasta" || /butter\s*cream/i.test(c.n || "") ? 1 : 0);
const defAg = (c) => (c.d === "ozel-pasta" ? 1 : 0);
function withDefaults(c, ov) {
  return { vi: c.vd ? 1 : 0, an: defAn(c), ag: defAg(c), dd: [], ...c, ...(ov || {}) };
}
function composeCatalog(cfg) {
  const del = new Set(cfg.catDel || []);
  const ov = cfg.catOv || {};
  const out = [];
  for (const c of CATALOG_SEED) { if (!del.has(c.id)) out.push(withDefaults(c, ov[c.id])); }
  for (const c of cfg.catAdd || []) { if (!del.has(c.id)) out.push(withDefaults(c, ov[c.id])); }
  return out;
}
function buildIndex(cfg) {
  const catalog = composeCatalog(cfg);
  return {
    ...cfg,
    catalog,
    _cat: new Map(catalog.map((c) => [c.id, c])),
    _dept: new Map(cfg.depts.map((d) => [d.id, d.name])),
    _branch: new Map(cfg.branches.map((b) => [b.id, b.name])),
  };
}
const deptName = (cfg, id) => cfg._dept.get(id) || "—";
const DEPT_REQ = "dept:";
const isDeptReq = (id) => typeof id === "string" && id.startsWith(DEPT_REQ);
const branchName = (cfg, id) => {
  if (isDeptReq(id)) return (cfg._dept.get(id.slice(DEPT_REQ.length)) || id.slice(DEPT_REQ.length)) + " bölümü";
  return cfg._branch.get(id) || "—";
};
const catFind = (cfg, id) => cfg._cat.get(id);

function sumUnits(records) {
  let adet = 0, kg = 0;
  records.forEach((r) => { if (r.unit === "Adet") adet += r.qty; else if (r.unit === "Gr") kg += r.qty / 1000; else kg += r.qty; });
  return { adet, kg };
}
const unitCell = (u) => (u.adet ? fmt(u.adet) : "–") + (u.kg ? ` + ${fmt(u.kg)} kg` : "");

/* telefonla çekilen fotoğrafı küçültüp veri olarak saklar */
function readImage(file, maxPx = 900, quality = 0.7) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      try {
        const scale = Math.min(1, maxPx / Math.max(img.width, img.height));
        const w = Math.max(1, Math.round(img.width * scale)), h = Math.max(1, Math.round(img.height * scale));
        const cv = document.createElement("canvas");
        cv.width = w; cv.height = h;
        cv.getContext("2d").drawImage(img, 0, 0, w, h);
        URL.revokeObjectURL(url);
        resolve(cv.toDataURL("image/jpeg", quality));
      } catch (e) { URL.revokeObjectURL(url); reject(e); }
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error("Görsel okunamadı")); };
    img.src = url;
  });
}

function ImageView({ src, onClose }) {
  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-4" onClick={onClose}>
      <img src={src} alt="Talep görseli" className="max-h-[85vh] max-w-full rounded-lg" onClick={(e) => e.stopPropagation()} />
    </div>
  );
}

function exportXlsx(filename, sheets) {
  try {
    const wb = XLSX.utils.book_new();
    sheets.forEach((s, i) => {
      const ws = XLSX.utils.aoa_to_sheet([s.head, ...s.rows]);
      XLSX.utils.book_append_sheet(wb, ws, (s.name || "Sayfa" + (i + 1)).replace(/[\\/?*[\]:]/g, "").slice(0, 31));
    });
    XLSX.writeFile(wb, filename);
  } catch (e) {
    const s = sheets[0];
    const text = [s.head, ...s.rows].map((r) => r.join(";")).join("\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob(["\ufeff" + text], { type: "text/csv;charset=utf-8" }));
    a.download = filename.replace(/\.xlsx$/, ".csv");
    a.click();
  }
}

/* talep penceresi */
function windowOpen(s) {
  const h = new Date().getHours();
  const o = s.openHour ?? 6, c = s.closeHour ?? 24;
  return c > o ? h >= o && h < c : h >= o || h < c;
}

/* ============================ ortak parçalar ============================ */

function Badge({ status }) {
  const s = STATUS[status] || STATUS.bekliyor;
  return <span className={`inline-block rounded px-2 py-0.5 text-xs font-medium whitespace-nowrap ${s.cls}`}>{s.label}</span>;
}
function Card({ children, className = "" }) { return <div className={`rounded-lg border border-stone-200 bg-white ${className}`}>{children}</div>; }
function Field({ label, children, hint }) {
  return (<label className="block"><span className="mb-1 block text-sm font-medium text-stone-700">{label}</span>{children}{hint && <span className="mt-1 block text-xs text-stone-500">{hint}</span>}</label>);
}
const inputCls = "w-full rounded-md border border-stone-300 bg-white px-3 py-2 text-base text-stone-900 focus:border-emerald-600 focus:outline-none focus:ring-1 focus:ring-emerald-600";
const numCls = "w-[4.5rem] shrink-0 rounded-md border border-stone-300 px-1 py-2 text-center text-base tabular-nums focus:border-emerald-600 focus:outline-none";
const numFireCls = "w-[4.5rem] shrink-0 rounded-md border border-red-200 bg-red-50 px-1 py-2 text-center text-base tabular-nums focus:border-red-500 focus:outline-none";

function Btn({ children, onClick, variant = "primary", className = "", disabled, type = "button" }) {
  const v = {
    primary: "bg-emerald-800 text-white hover:bg-emerald-900 disabled:bg-stone-300",
    ghost: "bg-white text-stone-700 border border-stone-300 hover:bg-stone-50 disabled:text-stone-400",
    danger: "bg-red-700 text-white hover:bg-red-800 disabled:bg-stone-300",
    amber: "bg-amber-600 text-white hover:bg-amber-700",
  }[variant];
  return <button type={type} onClick={onClick} disabled={disabled} className={`rounded-md px-3 py-2 text-sm font-medium transition-colors disabled:cursor-not-allowed ${v} ${className}`}>{children}</button>;
}
function Empty({ title, hint }) {
  return (<div className="rounded-lg border border-dashed border-stone-300 p-8 text-center"><p className="font-medium text-stone-700">{title}</p>{hint && <p className="mt-1 text-sm text-stone-500">{hint}</p>}</div>);
}
function Stat({ label, value, sub, tone = "stone" }) {
  const tones = { stone: "text-stone-900", emerald: "text-emerald-800", orange: "text-orange-700", sky: "text-sky-800", red: "text-red-700" };
  return (<div className="rounded-lg border border-stone-200 bg-white p-3 sm:p-4"><div className="text-xs text-stone-500">{label}</div>
    <div className={`mt-1 text-xl font-semibold tabular-nums sm:text-2xl ${tones[tone]}`}>{value}</div>{sub && <div className="mt-0.5 text-xs text-stone-500">{sub}</div>}</div>);
}
function Table({ head, rows, foot, max }) {
  if (!rows.length) return <Empty title="Seçilen filtrede kayıt yok" hint="Filtreyi değiştirip Ara'ya basın." />;
  const shown = max ? rows.slice(0, max) : rows;
  return (
    <Card className="overflow-x-auto">
      <table className="w-full min-w-max text-sm">
        <thead className="bg-stone-50 text-left text-xs text-stone-600"><tr>{head.map((h, i) => <th key={h + i} className={`px-3 py-2 font-medium ${i > 0 ? "text-right" : ""}`}>{h}</th>)}</tr></thead>
        <tbody className="divide-y divide-stone-100">{shown.map((r, i) => <tr key={i}>{r.map((c, j) => <td key={j} className={`px-3 py-2 ${j > 0 ? "text-right tabular-nums" : "text-stone-800"}`}>{c}</td>)}</tr>)}</tbody>
        {foot && <tfoot className="border-t-2 border-stone-300 bg-stone-50 font-medium"><tr>{foot.map((c, j) => <td key={j} className={`px-3 py-2 ${j > 0 ? "text-right tabular-nums" : "text-stone-900"}`}>{c}</td>)}</tr></tfoot>}
      </table>
      {max && rows.length > max && <p className="border-t border-stone-100 px-3 py-2 text-xs text-stone-500">{fmt(rows.length)} satırın ilk {max} tanesi. Tamamı için Excel'e aktarın.</p>}
    </Card>
  );
}
function Tabs({ value, onChange, items }) {
  return (<div className="flex flex-wrap items-end gap-1 border-b border-stone-200">
    {items.map(([id, label]) => <button key={id} onClick={() => onChange(id)}
      className={`-mb-px border-b-2 px-3 py-2 text-sm font-medium ${value === id ? "border-emerald-800 text-emerald-900" : "border-transparent text-stone-600 hover:text-stone-900"}`}>{label}</button>)}
  </div>);
}
function Modal({ title, children, onClose, footer, wide }) {
  return (
    <div className="fixed inset-0 z-50 flex items-end justify-center bg-black/40 sm:items-center sm:p-4" onClick={onClose}>
      <div className={`max-h-[90vh] w-full overflow-y-auto rounded-t-xl bg-white sm:rounded-xl ${wide ? "sm:max-w-3xl" : "sm:max-w-lg"}`} onClick={(e) => e.stopPropagation()}>
        <div className="sticky top-0 flex items-center justify-between border-b border-stone-200 bg-white px-4 py-3">
          <h3 className="font-medium text-stone-900">{title}</h3>
          <button onClick={onClose} className="rounded px-2 py-1 text-sm text-stone-500 hover:bg-stone-100">Kapat</button>
        </div>
        <div className="p-4">{children}</div>
        {footer && <div className="sticky bottom-0 flex justify-end gap-2 border-t border-stone-200 bg-white px-4 py-3">{footer}</div>}
      </div>
    </div>
  );
}

/* filtre değişikliği anında uygulanır; ayrıca "Ara" tuşuna basmak gerekmez */
function useFilters(initial) {
  const [draft, setDraft] = useState(initial);
  return {
    draft, applied: draft,
    set: (k, v) => setDraft((d) => ({ ...d, [k]: v })),
    setMany: (o) => setDraft((d) => ({ ...d, ...o })),
  };
}
function FilterBar({ f, onApply, withDates = true, children, label = "Tarih" }) {
  return (
    <Card className="p-4">
      <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
        {withDates && (<>
          <Field label={label + " başlangıç"}><input type="date" className={inputCls} value={f.draft.from || ""} onChange={(e) => f.set("from", e.target.value)} /></Field>
          <Field label={label + " bitiş"}><input type="date" className={inputCls} value={f.draft.to || ""} onChange={(e) => f.set("to", e.target.value)} /></Field>
        </>)}
        {children}
      </div>
      {withDates && (
        <div className="mt-3 flex flex-wrap items-center gap-2">
          <Btn variant="ghost" onClick={() => f.setMany({ from: today(), to: today() })}>Bugün</Btn>
          <Btn variant="ghost" onClick={() => f.setMany({ from: addDays(today(), 1), to: addDays(today(), 1) })}>Yarın</Btn>
          <Btn variant="ghost" onClick={() => f.setMany({ from: daysAgo(6), to: today() })}>Son 7 gün</Btn>
          <Btn variant="ghost" onClick={() => f.setMany({ from: monthStart(), to: today() })}>Bu ay</Btn>
        </div>
      )}
    </Card>
  );
}
function DeptSelect({ f, cfg, allowed }) {
  return (<Field label="Bölüm"><select className={inputCls} value={f.draft.dept} onChange={(e) => f.set("dept", e.target.value)}>
    {allowed.length > 1 && <option value="all">Bütün bölümler</option>}
    {allowed.map((id) => <option key={id} value={id}>{deptName(cfg, id)}</option>)}
  </select></Field>);
}
function BranchSelect({ f, cfg }) {
  return (<Field label="Şube"><select className={inputCls} value={f.draft.branch} onChange={(e) => f.set("branch", e.target.value)}>
    <option value="all">Bütün şubeler</option>{cfg.branches.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}
  </select></Field>);
}

/* ============================ giriş ============================ */

function Login({ users, onLogin }) {
  const [u, setU] = useState(""); const [pin, setPin] = useState(""); const [err, setErr] = useState("");
  const submit = () => {
    const f = users.find((x) => lower(x.username) === lower(u.trim()) && x.pin === pin.trim());
    if (!f) return setErr("Kullanıcı adı veya şifre hatalı.");
    onLogin(f);
  };
  return (
    <div className="flex min-h-screen items-center justify-center bg-stone-100 p-4">
      <div className="w-full max-w-sm">
        <div className="mb-6"><h1 className="text-2xl font-semibold text-stone-900">Üretim ve Sevkiyat Takibi</h1>
          <p className="mt-1 text-sm text-stone-600">Talep bugün, üretim yarın, sevk ertesi gün.</p></div>
        <Card className="p-5"><div className="space-y-4">
          <Field label="Kullanıcı adı"><input className={inputCls} value={u} autoCapitalize="none" onChange={(e) => setU(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} placeholder="admin" /></Field>
          <Field label="Şifre"><input className={inputCls} type="password" inputMode="numeric" value={pin} onChange={(e) => setPin(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} placeholder="••••" /></Field>
          {err && <p className="text-sm text-red-700">{err}</p>}
          <Btn onClick={submit} className="w-full">Giriş yap</Btn>
          <p className="text-xs text-stone-500">Çıkış yapana kadar oturum açık kalır.</p>
        </div></Card>
        <details className="mt-4 text-sm text-stone-600"><summary className="cursor-pointer">Deneme hesapları</summary>
          <ul className="mt-2 space-y-1 text-xs">
            <li>admin / 1234 · sube.merter / 1111 · sef.sutlu / 2222</li>
            <li>depo / 3333 · uretim / 4444 · sevkiyat / 5555 · sofor / 5555</li>
            <li>genelmudur / 4444 · operasyon / 4444</li>
          </ul></details>
      </div>
    </div>
  );
}

/* ============================ şube: talep girişi ============================ */

function NewOrder({ me, cfg, orders, onSubmit, ensure }) {
  const s = cfg.settings;
  useEffect(() => { ensure(daysAgo(21), addDays(today(), 10)); }, [ensure]);
  const open = windowOpen(s);
  const [prodDate, setProdDate] = useState(addDays(today(), s.leadProd));
  const [shipDate, setShipDate] = useState(addDays(today(), s.leadProd + s.leadShip));
  const [q, setQ] = useState("");
  const [deptFilter, setDeptFilter] = useState("all");
  const [cart, setCart] = useState({});
  const [msg, setMsg] = useState("");
  const [zoom, setZoom] = useState(null);

  const list = useMemo(() => {
    const t = lower(q.trim());
    return cfg.catalog.filter((p) => p.vt).filter((p) => (deptFilter === "all" ? true : p.d === deptFilter))
      .filter((p) => (t ? lower(p.n).includes(t) || String(p.id).includes(t) : true)).slice(0, 500);
  }, [q, deptFilter, cfg.catalog]);

  const existing = orders.find((o) => o.branchId === me.branchId && o.prodDate === prodDate);
  const cartLines = Object.entries(cart).filter(([, v]) => Number(v && v.qty) > 0);
  const setItem = (id, patch) => setCart((c) => ({ ...c, [id]: { qty: "", note: "", img: "", ...(c[id] || {}), ...patch } }));
  if (!me.branchId) return <Empty title="Bu hesap bir şubeye bağlı değil" hint="Talep yalnızca şube hesaplarından girilir." />;

  const setProd = (d) => { setProdDate(d); setShipDate(addDays(d, s.leadShip)); };

  const copyPrev = () => {
    const prev = orders.filter((o) => o.branchId === me.branchId && o.prodDate < prodDate).sort((a, b) => (a.prodDate < b.prodDate ? 1 : -1))[0];
    if (!prev) return setMsg("Kopyalanacak önceki talep bulunamadı");
    const c = {}; prev.lines.forEach((l) => { c[l.productId] = { qty: String(l.qty), note: "", img: "" }; });
    setCart(c); setMsg(`${dateTR(prev.prodDate)} talebi getirildi`);
    setTimeout(() => setMsg(""), 4000);
  };

  const submit = async () => {
    if (!cartLines.length || !open) return;
    await onSubmit({
      branchId: me.branchId, prodDate, shipDate, note: "", createdBy: me.id,
      lines: cartLines.map(([pid, v]) => {
        const p = catFind(cfg, pid);
        return { id: uid(), productId: pid, name: p.n, deptId: p.d, price: p.p || 0, qty: Number(v.qty),
          note: p.an ? (v.note || "").trim() : "", img: p.ag ? v.img || "" : "",
          prepared: 0, sent: 0, status: "bekliyor" };
      }),
    });
    setCart({});
    setMsg(existing ? "Mevcut talebe eklendi" : "Talep gönderildi");
    setTimeout(() => setMsg(""), 4000);
  };

  return (
    <div className="space-y-4 pb-20">
      {!open && (
        <Card className="border-red-300 bg-red-50 p-4">
          <p className="text-sm font-medium text-red-900">Talep girişi şu anda kapalı.</p>
          <p className="mt-1 text-sm text-red-800">Talepler her gün saat {String(s.openHour).padStart(2, "0")}:00 ile {s.closeHour === 24 ? "24" : String(s.closeHour).padStart(2, "0")}:00 arasında girilir. Sabah {String(s.openHour).padStart(2, "0")}:00'da tekrar açılır.</p>
        </Card>
      )}
      <Card className="p-4">
        <div className="grid gap-4 sm:grid-cols-3">
          <Field label="Şube"><input className={inputCls + " bg-stone-100"} value={branchName(cfg, me.branchId)} readOnly /></Field>
          <Field label="Üretim tarihi" hint={`Talepten ${s.leadProd} gün sonra`}>
            <input type="date" className={inputCls} value={prodDate} onChange={(e) => setProd(e.target.value)} disabled={!open} />
          </Field>
          <Field label="Planlanan sevk tarihi" hint="Sevkiyat gerçek tarihi kendi girer">
            <input type="date" className={inputCls} value={shipDate} onChange={(e) => setShipDate(e.target.value)} disabled={!open} />
          </Field>
        </div>
        <div className="mt-3 flex flex-wrap items-center gap-2">
          <Btn variant="ghost" disabled={!open} onClick={() => setProd(addDays(today(), s.leadProd))}>Varsayılan gün</Btn>
          <Btn variant="ghost" disabled={!open} onClick={copyPrev}>Önceki talebi kopyala</Btn>
        </div>
        {existing && open && (
          <p className="mt-3 rounded bg-amber-50 px-3 py-2 text-sm text-amber-900">
            {dateTR(prodDate)} üretimi için bu şubenin {existing.lines.length} kalemlik talebi var. Gireceğiniz kalemler ona eklenir, aynı üründe adetler toplanır. Düzeltme için Talepler sekmesini kullanın.
          </p>
        )}
      </Card>

      <Card className="p-4">
        <div className="flex flex-col gap-2 sm:flex-row">
          <input className={inputCls + " sm:flex-1"} placeholder="Ürün ara" value={q} onChange={(e) => setQ(e.target.value)} />
          <select className={inputCls + " sm:w-52"} value={deptFilter} onChange={(e) => setDeptFilter(e.target.value)}>
            <option value="all">Bütün bölümler</option>{cfg.depts.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}
          </select>
        </div>
        <div className="mt-2 divide-y divide-stone-100">
          {list.map((p) => {
            const v = cart[p.id] || { qty: "", note: "", img: "" };
            const extra = (p.an || p.ag) && Number(v.qty) > 0;
            return (
              <div key={p.id} className="py-2">
                <div className="flex items-center gap-3">
                  <div className="min-w-0 flex-1">
                    <div className="text-sm font-medium leading-tight text-stone-800">{p.n}</div>
                    <div className="text-xs text-stone-500">
                      {deptName(cfg, p.d)}{p.p ? ` · ${money(p.p)}` : ""}
                      {(p.an || p.ag) && <span className="ml-1 text-emerald-800">· özel sipariş</span>}
                    </div>
                  </div>
                  <input type="number" min="0" inputMode="numeric" className={numCls} disabled={!open}
                    value={v.qty} onChange={(e) => setItem(p.id, { qty: e.target.value })} placeholder="0" />
                </div>
                {extra && (
                  <div className="mt-2 rounded-md border border-emerald-200 bg-emerald-50 p-2">
                    {p.an ? (
                      <input className="w-full rounded-md border border-emerald-300 bg-white px-2 py-1.5 text-sm" disabled={!open}
                        placeholder="Not: kaç kişilik, üzerine yazı, renk, teslim saati…"
                        value={v.note} onChange={(e) => setItem(p.id, { note: e.target.value })} />
                    ) : null}
                    {p.ag ? (
                      <div className="mt-2 flex flex-wrap items-center gap-2">
                        <label className="cursor-pointer rounded-md border border-emerald-300 bg-white px-3 py-1.5 text-sm text-emerald-900">
                          {v.img ? "Görseli değiştir" : "Görsel ekle"}
                          <input type="file" accept="image/*" className="hidden" disabled={!open}
                            onChange={async (e) => {
                              const fl = e.target.files && e.target.files[0];
                              e.target.value = "";
                              if (!fl) return;
                              try { setItem(p.id, { img: await readImage(fl) }); }
                              catch (err) { setMsg("Görsel eklenemedi"); setTimeout(() => setMsg(""), 3000); }
                            }} />
                        </label>
                        {v.img ? (
                          <>
                            <button type="button" onClick={() => setZoom(v.img)}>
                              <img src={v.img} alt="" className="h-12 w-12 rounded border border-emerald-300 object-cover" />
                            </button>
                            <button type="button" className="text-xs text-red-700 hover:underline" onClick={() => setItem(p.id, { img: "" })}>kaldır</button>
                          </>
                        ) : null}
                      </div>
                    ) : null}
                  </div>
                )}
              </div>
            );
          })}
          {!list.length && <p className="py-6 text-center text-sm text-stone-500">Eşleşen ürün yok.</p>}
        </div>
      </Card>

      <div className="fixed inset-x-0 bottom-0 z-20 border-t border-stone-200 bg-white px-4 py-3">
        <div className="mx-auto flex max-w-6xl items-center justify-between gap-3">
          <div className="text-sm text-stone-600">{cartLines.length} kalem · <span className="font-medium text-stone-900">{fmt(cartLines.reduce((a, [, v]) => a + Number(v.qty), 0))} adet</span>
            <span className="ml-2 text-stone-500">{money(cartLines.reduce((a, [id, v]) => a + Number(v.qty) * (catFind(cfg, id)?.p || 0), 0))}</span></div>
          <div className="flex items-center gap-2">{msg && <span className="text-sm text-emerald-700">{msg}</span>}
            <Btn onClick={submit} disabled={!cartLines.length || !open}>Talebi gönder</Btn></div>
        </div>
      </div>
      {zoom && <ImageView src={zoom} onClose={() => setZoom(null)} />}
    </div>
  );
}

/* ============================ şube: talepler ve düzeltme ============================ */

function OrderList({ orders, cfg, canSeeAll, me, onEditLine, onDeleteLine, ensure }) {
  const [open, setOpen] = useState(null);
  const [mode, setMode] = useState("bugun");
  const [editOrder, setEditOrder] = useState(null);
  const [zoom, setZoom] = useState(null);
  const f = useFilters({ from: daysAgo(7), to: today() });
  const s = cfg.settings;
  useEffect(() => { ensure(daysAgo(3), addDays(today(), 10)); }, [ensure]);
  useEffect(() => {
    if (mode === "aralik") ensure(f.applied.from, addDays(f.applied.to, (s.leadProd || 0) + (s.leadShip || 0) + 2));
  }, [mode, f.applied.from, f.applied.to, ensure, s.leadProd, s.leadShip]);

  const mine = useMemo(() => {
    const base = orders.filter((o) => (canSeeAll ? true : o.branchId === me.branchId));
    const sel = mode === "bugun"
      ? base.filter((o) => (o.orderDate || o.prodDate) === today())
      : base.filter((o) => inRange(o.orderDate || o.prodDate, f.applied.from, f.applied.to));
    return sel.slice().sort((a, b) => (a.prodDate < b.prodDate ? 1 : -1));
  }, [orders, canSeeAll, me.branchId, mode, f.applied.from, f.applied.to]);

  const canEditOrder = (o) => canSeeAll === false && mode === "bugun" && windowOpen(s) && o.prodDate > today();
  const editable = (o, l) => canEditOrder(o) && editOrder === o.id && lineStatus(l) === "bekliyor";

  return (
    <div className="space-y-3">
      {!windowOpen(s) && !canSeeAll && <Card className="border-amber-300 bg-amber-50 p-3 text-sm text-amber-900">Talep penceresi kapalı. Düzeltme sabah {String(s.openHour).padStart(2, "0")}:00'da tekrar açılır.</Card>}
      <Card className="p-4">
        <div className="flex flex-wrap items-center gap-2">
          <Btn variant={mode === "bugun" ? "primary" : "ghost"} onClick={() => setMode("bugun")}>Bugün açılanlar</Btn>
          <Btn variant={mode === "aralik" ? "primary" : "ghost"} onClick={() => setMode("aralik")}>Geçmiş talepler</Btn>
          {mode === "bugun" && <span className="text-xs text-stone-500">{dateTR(today())} tarihinde oluşturulan talepler</span>}
        </div>
        {mode === "aralik" && (
          <div className="mt-3 grid gap-3 sm:grid-cols-2 lg:grid-cols-4">
            <Field label="Başlangıç"><input type="date" className={inputCls} value={f.draft.from} onChange={(e) => f.set("from", e.target.value)} /></Field>
            <Field label="Bitiş"><input type="date" className={inputCls} value={f.draft.to} onChange={(e) => f.set("to", e.target.value)} /></Field>
            <div className="flex items-end gap-2 sm:col-span-2">
              <Btn variant="ghost" onClick={() => f.setMany({ from: daysAgo(6), to: today() })}>Son 7 gün</Btn>
              <Btn variant="ghost" onClick={() => f.setMany({ from: daysAgo(29), to: today() })}>Son 30 gün</Btn>
              <Btn variant="ghost" onClick={() => f.setMany({ from: monthStart(), to: today() })}>Bu ay</Btn>
            </div>
          </div>
        )}
      </Card>
      {!mine.length && <Empty title={mode === "bugun" ? "Bugün açılmış talep yok" : "Seçilen aralıkta talep yok"}
        hint={mode === "bugun" ? "Geçmiş talepler için üstteki düğmeyi kullanın." : "Tarih aralığını genişletin."} />}
      {mine.map((o) => {
        const total = o.lines.reduce((a, l) => a + l.qty, 0);
        const sent = o.lines.reduce((a, l) => a + sentQ(l), 0);
        const done = o.lines.every((l) => lineStatus(l) === "gonderildi");
        return (
          <Card key={o.id}>
            <button className="flex w-full items-center gap-3 p-4 text-left" onClick={() => setOpen(open === o.id ? null : o.id)}>
              <div className="min-w-0 flex-1">
                <div className="font-medium text-stone-900">{branchName(cfg, o.branchId)} · üretim {dateTR(o.prodDate)}</div>
                <div className="mt-0.5 text-sm text-stone-600">sevk {dateTR(o.shipDate)} · {o.lines.length} kalem · {fmt(sent)}/{fmt(total)} gönderildi{o.note ? ` · ${o.note}` : ""}</div>
              </div>
              <Badge status={done ? "gonderildi" : sent > 0 ? "kismen" : "bekliyor"} />
              <span className="text-stone-400">{open === o.id ? "▲" : "▼"}</span>
            </button>
            {open === o.id && (
              <div className="overflow-x-auto border-t border-stone-100 px-4 pb-4">
                <div className="flex flex-wrap items-center gap-2 py-3">
                  {canEditOrder(o) ? (
                    editOrder === o.id ? (
                      <>
                        <Btn variant="amber" onClick={() => setEditOrder(null)}>Düzenlemeyi bitir</Btn>
                        <span className="text-xs text-amber-800">Adetleri değiştirebilir, satır silebilirsiniz.</span>
                      </>
                    ) : (
                      <>
                        <Btn variant="ghost" onClick={() => setEditOrder(o.id)}>Düzenle</Btn>
                        <span className="text-xs text-stone-500">Adetler kilitli. Değiştirmek için Düzenle'ye basın.</span>
                      </>
                    )
                  ) : (
                    <span className="text-xs text-stone-500">
                      {mode === "aralik" ? "Geçmiş talepler değiştirilemez." : o.prodDate <= today() ? "Üretim günü geldiği için bu talep kilitlendi." : "Bu talep değiştirilemez."}
                    </span>
                  )}
                </div>
                <table className="w-full min-w-max text-sm">
                  <thead><tr className="text-left text-xs text-stone-500">
                    <th className="py-2">Ürün</th><th className="py-2 text-right">İstenen</th><th className="py-2 text-right">Hazırlanan</th>
                    <th className="py-2 text-right">Gönderilen</th><th className="py-2 text-right">Sevk</th><th className="py-2 text-right">Sebep</th><th className="py-2 text-right">Durum</th><th /></tr></thead>
                  <tbody className="divide-y divide-stone-100">
                    {o.lines.map((l) => (
                      <tr key={l.id}>
                        <td className="py-2 pr-2"><div className="text-stone-800">{l.name}</div>
                          <div className="text-xs text-stone-500">{deptName(cfg, l.deptId)}</div>
                          {l.note && <div className="mt-0.5 text-xs text-emerald-900">Not: {l.note}</div>}
                          {l.img && <button type="button" className="mt-1" onClick={() => setZoom(l.img)}>
                            <img src={l.img} alt="" className="h-10 w-10 rounded border border-stone-300 object-cover" /></button>}
                        </td>
                        <td className="py-2 text-right tabular-nums">
                          {editable(o, l)
                            ? <input type="number" min="1" inputMode="numeric" defaultValue={l.qty} className="w-20 rounded border border-stone-300 px-1 py-1 text-right tabular-nums"
                                onBlur={(e) => { const v = Number(e.target.value); if (v > 0 && v !== l.qty) onEditLine(o, l.id, v); }} />
                            : fmt(l.qty)}
                        </td>
                        <td className="py-2 text-right tabular-nums">{fmt(prep(l))}</td>
                        <td className="py-2 text-right tabular-nums">{fmt(sentQ(l))}</td>
                        <td className="py-2 pl-2 text-right text-xs text-stone-600">{l.sentDate ? dateTR(l.sentDate) : "—"}</td>
                        <td className="py-2 pl-2 text-right text-xs text-stone-600">{l.shortReason || "—"}</td>
                        <td className="py-2 pl-2 text-right"><Badge status={lineStatus(l)} /></td>
                        <td className="py-2 pl-2 text-right">{editable(o, l) && <button className="text-xs text-red-700 hover:underline" onClick={() => onDeleteLine(o, l.id)}>sil</button>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </Card>
        );
      })}
      {zoom && <ImageView src={zoom} onClose={() => setZoom(null)} />}
    </div>
  );
}

/* ============================ bölüm talepleri ============================ */

function DeptRequests({ me, cfg, orders, onUpdateLine, ensure }) {
  const allowed = me.depts && me.depts.length ? me.depts : cfg.depts.map((d) => d.id);
  const f = useFilters({ from: today(), to: today(), dept: allowed.length === 1 ? allowed[0] : "all", branch: "all" });
  const fd = useFilters({ from: daysAgo(29), to: today(), branch: "all" });
  const [tab, setTab] = useState("ozet");
  const [hideDone, setHideDone] = useState(true);
  const [open, setOpen] = useState({});
  const A = f.applied;
  const activeDepts = A.dept === "all" ? allowed : [A.dept];
  useEffect(() => { ensure(A.from, A.to); }, [A.from, A.to, ensure]);
  useEffect(() => { ensure(fd.applied.from, fd.applied.to); }, [fd.applied.from, fd.applied.to, ensure]);

  const rows = useMemo(() => {
    const out = [];
    orders.forEach((o) => {
      if (!inRange(o.prodDate, A.from, A.to)) return;
      if (A.branch !== "all" && o.branchId !== A.branch) return;
      o.lines.forEach((l) => { if (activeDepts.includes(l.deptId)) out.push({ order: o, line: l }); });
    });
    return out.sort((a, b) => (a.order.prodDate < b.order.prodDate ? -1 : 1));
  }, [orders, A.from, A.to, A.branch, activeDepts.join()]);

  const totals = rows.reduce((a, r) => { a.qty += r.line.qty; a.prep += prep(r.line); a.sent += sentQ(r.line); if (!["hazir", "kismen", "gonderildi"].includes(lineStatus(r.line))) a.openLines += 1; return a; }, { qty: 0, prep: 0, sent: 0, openLines: 0 });
  const singleDept = activeDepts.length === 1;

  const byProduct = useMemo(() => {
    const m = new Map();
    rows.forEach(({ order, line }) => {
      let x = m.get(line.productId);
      if (!x) { x = { key: line.productId, name: line.name, qty: 0, prep: 0, sent: 0, branches: {} }; m.set(line.productId, x); }
      x.qty += line.qty; x.prep += prep(line); x.sent += sentQ(line);
      x.branches[order.branchId] = (x.branches[order.branchId] || 0) + line.qty;
    });
    return [...m.values()].sort((a, b) => b.qty - a.qty);
  }, [rows]);

  const shownDepts = cfg.depts.filter((d) => activeDepts.includes(d.id));
  const matrix = useMemo(() => {
    const out = [];
    cfg.branches.forEach((b) => {
      const mine = rows.filter((r) => r.order.branchId === b.id);
      if (!mine.length) return;
      const cells = shownDepts.map((d) => mine.filter((r) => r.line.deptId === d.id).reduce((a, r) => a + r.line.qty, 0));
      const detail = {};
      mine.forEach(({ line }) => { detail[line.deptId] = detail[line.deptId] || {}; detail[line.deptId][line.name] = (detail[line.deptId][line.name] || 0) + line.qty; });
      out.push({ b, cells, detail, tot: cells.reduce((a, c) => a + c, 0) });
    });
    return out.filter((r) => r.tot > 0);
  }, [rows, shownDepts.map((d) => d.id).join()]);
  const deptTotals = shownDepts.map((d) => rows.filter((r) => r.line.deptId === d.id).reduce((a, r) => a + r.line.qty, 0));
  const pending = rows.filter((r) => !(hideDone && ["hazir", "kismen", "gonderildi"].includes(lineStatus(r.line))));

  const doneRows = useMemo(() => {
    const out = [];
    orders.forEach((o) => {
      if (!inRange(o.prodDate, fd.applied.from, fd.applied.to)) return;
      if (fd.applied.branch !== "all" && o.branchId !== fd.applied.branch) return;
      o.lines.forEach((l) => { if (activeDepts.includes(l.deptId) && ["hazir", "kismen", "gonderildi"].includes(lineStatus(l))) out.push({ o, l }); });
    });
    return out.sort((a, b) => (a.o.prodDate < b.o.prodDate ? 1 : -1));
  }, [orders, fd.applied.from, fd.applied.to, fd.applied.branch, activeDepts.join()]);

  const toggle = (k) => setOpen((s) => ({ ...s, [k]: !s[k] }));

  return (
    <div className="space-y-4">
      <FilterBar f={f} onApply={(d) => ensure(d.from, d.to)} label="Üretim tarihi">
        <DeptSelect f={f} cfg={cfg} allowed={allowed} /><BranchSelect f={f} cfg={cfg} />
      </FilterBar>

      <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
        <Stat label="Gelen talep" value={fmt(totals.qty)} sub={`${rows.length} kalem`} />
        <Stat label="Hazırlanan" value={fmt(totals.prep)} sub={`%${totals.qty ? Math.round((totals.prep / totals.qty) * 100) : 0}`} tone="sky" />
        <Stat label="Gönderilen" value={fmt(totals.sent)} sub={`%${totals.qty ? Math.round((totals.sent / totals.qty) * 100) : 0}`} tone="emerald" />
        <Stat label="Hazırlanacak" value={fmt(totals.openLines)} sub="henüz hazır değil" tone={totals.openLines ? "orange" : "emerald"} />
      </div>

      <Tabs value={tab} onChange={setTab} items={[["ozet", "Özet"], ["kalem", "Talep kalemleri"], ["tamam", "Tamamlananlar"]]} />

      {tab === "ozet" && (singleDept ? (
        <Card className="overflow-x-auto">
          <table className="w-full min-w-max text-sm">
            <thead className="bg-stone-50 text-left text-xs text-stone-600"><tr>
              <th className="px-3 py-2 font-medium">Ürün</th><th className="px-3 py-2 text-right font-medium">Talep</th>
              <th className="px-3 py-2 text-right font-medium">Hazırlanan</th><th className="px-3 py-2 text-right font-medium">Gönderilen</th><th /></tr></thead>
            <tbody className="divide-y divide-stone-100">
              {byProduct.map((p) => (
                <Fragment key={p.key}>
                  <tr className="cursor-pointer hover:bg-stone-50" onClick={() => toggle("p" + p.key)}>
                    <td className="px-3 py-2 text-stone-800">{p.name}</td>
                    <td className="px-3 py-2 text-right font-medium tabular-nums">{fmt(p.qty)}</td>
                    <td className="px-3 py-2 text-right tabular-nums">{fmt(p.prep)}</td>
                    <td className="px-3 py-2 text-right tabular-nums">{fmt(p.sent)}</td>
                    <td className="px-3 py-2 text-right text-stone-400">{open["p" + p.key] ? "▲" : "▼"}</td>
                  </tr>
                  {open["p" + p.key] && (
                    <tr className="bg-stone-50"><td colSpan={5} className="px-6 py-2">
                      <div className="text-xs text-stone-500">Hangi şube istedi</div>
                      <ul className="mt-1 space-y-0.5">{Object.entries(p.branches).map(([bid, q]) => (
                        <li key={bid} className="flex justify-between text-sm text-stone-700"><span>{branchName(cfg, bid)}</span><span className="tabular-nums">{fmt(q)}</span></li>))}</ul>
                    </td></tr>
                  )}
                </Fragment>
              ))}
            </tbody>
            {byProduct.length > 0 && <tfoot className="border-t-2 border-stone-300 bg-stone-50 font-medium"><tr>
              <td className="px-3 py-2">Toplam</td><td className="px-3 py-2 text-right tabular-nums">{fmt(totals.qty)}</td>
              <td className="px-3 py-2 text-right tabular-nums">{fmt(totals.prep)}</td><td className="px-3 py-2 text-right tabular-nums">{fmt(totals.sent)}</td><td /></tr></tfoot>}
          </table>
          {!byProduct.length && <p className="p-6 text-center text-sm text-stone-500">Seçilen filtrede talep yok.</p>}
        </Card>
      ) : (
        <Card className="overflow-x-auto">
          <table className="w-full min-w-max text-sm">
            <thead className="bg-stone-50 text-left text-xs text-stone-600"><tr><th className="px-3 py-2 font-medium">Şube</th>
              {shownDepts.map((d) => <th key={d.id} className="px-3 py-2 text-right font-medium">{d.name}</th>)}
              <th className="px-3 py-2 text-right font-medium">Toplam</th><th /></tr></thead>
            <tbody className="divide-y divide-stone-100">
              {matrix.map((r) => (
                <Fragment key={r.b.id}>
                  <tr className="cursor-pointer hover:bg-stone-50" onClick={() => toggle("b" + r.b.id)}>
                    <td className="px-3 py-2 text-stone-800">{r.b.name}</td>
                    {r.cells.map((c, i) => <td key={i} className="px-3 py-2 text-right tabular-nums">{c ? fmt(c) : "–"}</td>)}
                    <td className="px-3 py-2 text-right font-medium tabular-nums">{fmt(r.tot)}</td>
                    <td className="px-3 py-2 text-right text-stone-400">{open["b" + r.b.id] ? "▲" : "▼"}</td>
                  </tr>
                  {open["b" + r.b.id] && (
                    <tr className="bg-stone-50"><td colSpan={shownDepts.length + 3} className="px-6 py-3">
                      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
                        {Object.entries(r.detail).map(([did, prods]) => (
                          <div key={did}><div className="mb-1 text-xs font-medium text-emerald-900">{deptName(cfg, did)}</div>
                            <ul className="space-y-0.5">{Object.entries(prods).sort((a, b) => b[1] - a[1]).map(([n, q]) => (
                              <li key={n} className="flex justify-between gap-3 text-sm text-stone-700"><span className="truncate">{n}</span><span className="shrink-0 tabular-nums">{fmt(q)}</span></li>))}</ul></div>
                        ))}
                      </div></td></tr>
                  )}
                </Fragment>
              ))}
            </tbody>
            {matrix.length > 0 && <tfoot className="border-t-2 border-stone-300 bg-stone-50 font-medium"><tr><td className="px-3 py-2">Toplam</td>
              {deptTotals.map((n, i) => <td key={i} className="px-3 py-2 text-right tabular-nums">{fmt(n)}</td>)}
              <td className="px-3 py-2 text-right tabular-nums">{fmt(totals.qty)}</td><td /></tr></tfoot>}
          </table>
          {!matrix.length && <p className="p-6 text-center text-sm text-stone-500">Seçilen filtrede talep yok.</p>}
        </Card>
      ))}

      {tab === "kalem" && (
        <div className="space-y-3">
          <label className="flex items-center gap-2 text-sm text-stone-700">
            <input type="checkbox" checked={hideDone} onChange={(e) => setHideDone(e.target.checked)} className="h-4 w-4 rounded border-stone-300" />Tamamlananları gizle</label>
          {!pending.length ? <Empty title="Bu filtrede bekleyen kalem yok" /> :
            pending.map(({ order, line }) => <PrepRow key={line.id} order={order} line={line} cfg={cfg} onUpdate={(p) => onUpdateLine(order, line.id, p)} />)}
        </div>
      )}

      {tab === "tamam" && (
        <div className="space-y-4">
          <FilterBar f={fd} onApply={(d) => ensure(d.from, d.to)} label="Üretim tarihi"><BranchSelect f={fd} cfg={cfg} /></FilterBar>
          <Table max={400} head={["Üretim", "Şube", "Bölüm", "Ürün", "İstenen", "Hazırlanan", "Gönderilen", "Eksik", "Sebep", "Durum"]}
            rows={doneRows.map(({ o, l }) => [dateTR(o.prodDate), branchName(cfg, o.branchId), deptName(cfg, l.deptId), l.name,
              fmt(l.qty), fmt(prep(l)), fmt(sentQ(l)), openQty(l) ? fmt(openQty(l)) : "–", l.shortReason || "—", STATUS[lineStatus(l)].label])} />
        </div>
      )}
    </div>
  );
}

function PrepRow({ order, line, cfg, onUpdate }) {
  const [val, setVal] = useState(String(prep(line) || ""));
  const [zoom, setZoom] = useState(null);
  const typedDiff = val !== "" && Number(val) !== line.qty;
  const apply = (status) => {
    let p = Number(val) || 0;
    if (status === "hazir") p = line.qty;
    if (status === "hazirlaniyor" && p === 0) p = prep(line);
    setVal(String(p || "")); onUpdate({ status, prepared: p });
  };
  const saveQty = () => { const p = Number(val) || 0; if (p <= 0) return; onUpdate({ status: p >= line.qty ? "hazir" : "hazirlaniyor", prepared: p }); };
  return (
    <Card className="p-4">
      <div className="flex flex-wrap items-start justify-between gap-2">
        <div className="min-w-0"><div className="font-medium text-stone-900">{line.name}</div>
          <div className="mt-0.5 text-sm text-stone-600">{branchName(cfg, order.branchId)} · üretim {dateTR(order.prodDate)} · istenen <span className="font-medium tabular-nums text-stone-900">{fmt(line.qty)}</span>
            {prep(line) > 0 && <> · hazırlanan <span className="tabular-nums">{fmt(prep(line))}</span></>}</div>
          {line.note && <div className="mt-1 rounded bg-emerald-50 px-2 py-1 text-sm text-emerald-900">Not: {line.note}</div>}
          {line.img && <button type="button" className="mt-2 block" onClick={() => setZoom(line.img)}>
            <img src={line.img} alt="" className="h-20 w-20 rounded border border-stone-300 object-cover" /></button>}</div>
        <Badge status={lineStatus(line)} />
      </div>
      <div className="mt-3 flex flex-wrap items-center gap-2">
        <Btn variant={lineStatus(line) === "hazirlaniyor" ? "amber" : "ghost"} onClick={() => apply("hazirlaniyor")}>Hazırlanıyor</Btn>
        <Btn onClick={() => apply("hazir")} disabled={typedDiff}>Tamamı hazır ({fmt(line.qty)})</Btn>
        <div className="flex items-center gap-1">
          <input type="number" min="0" max={line.qty} inputMode="numeric" className={numCls} value={val} onChange={(e) => setVal(e.target.value)} placeholder="adet" />
          <Btn variant={typedDiff ? "primary" : "ghost"} onClick={saveQty}>Hazırlananı kaydet</Btn>
        </div>
        {typedDiff && <span className="text-xs text-amber-800">{fmt(Number(val))} yazdınız; bu yüzden "Tamamı hazır" kapalı. Yazdığınız adedi kaydetmek için yandaki düğmeyi kullanın.</span>}
      </div>
      {zoom && <ImageView src={zoom} onClose={() => setZoom(null)} />}
    </Card>
  );
}


/* --- bölümlerin depodan günlük talebi --- */
function DeptSupply({ me, cfg, orders, onSubmit, ensure }) {
  const own = (me.depts && me.depts.length ? me.depts : cfg.depts.map((d) => d.id)).filter((d) => d !== "depo");
  const allowed = own.length ? own : cfg.depts.filter((d) => d.id !== "depo").map((d) => d.id);
  const [dept, setDept] = useState(allowed[0]);
  const [date, setDate] = useState(today());
  const [q, setQ] = useState("");
  const [cart, setCart] = useState({});
  const [tab, setTab] = useState("yeni");
  const [msg, setMsg] = useState("");
  const f = useFilters({ from: daysAgo(13), to: addDays(today(), 3) });

  useEffect(() => { ensure(daysAgo(14), addDays(today(), 7)); }, [ensure]);
  useEffect(() => { ensure(f.applied.from, f.applied.to); }, [f.applied.from, f.applied.to, ensure]);

  const list = useMemo(() => {
    const t = lower(q.trim());
    return cfg.catalog
      .filter((c) => c.vi)
      .filter((c) => !c.dd || !c.dd.length || c.dd.includes(dept))
      .filter((c) => (t ? lower(c.n).includes(t) || lower(String(c.c || c.id)).includes(t) : true))
      .slice(0, 500);
  }, [cfg.catalog, q, dept]);

  const reqBranch = DEPT_REQ + dept;
  const lines = Object.entries(cart).filter(([, v]) => Number(v) > 0);
  const existing = orders.find((o) => o.branchId === reqBranch && o.prodDate === date);

  const mine = useMemo(() => orders
    .filter((o) => isDeptReq(o.branchId) && allowed.includes(o.branchId.slice(DEPT_REQ.length)))
    .filter((o) => inRange(o.prodDate, f.applied.from, f.applied.to))
    .sort((a, b) => (a.prodDate < b.prodDate ? 1 : -1)), [orders, allowed.join(), f.applied.from, f.applied.to]);

  const submit = async () => {
    if (!lines.length) return;
    await onSubmit({
      branchId: reqBranch, prodDate: date, shipDate: date, note: "", createdBy: me.id,
      lines: lines.map(([pid, qty]) => {
        const p = catFind(cfg, pid);
        return { id: uid(), productId: pid, name: p.n, deptId: "depo", price: p.p || 0, qty: Number(qty),
          note: "", img: "", prepared: 0, sent: 0, status: "bekliyor" };
      }),
    });
    setCart({});
    setMsg(existing ? "Mevcut talebe eklendi" : "Talep depoya iletildi");
    setTimeout(() => setMsg(""), 4000);
  };

  return (
    <div className={tab === "yeni" ? "space-y-4 pb-20" : "space-y-4"}>
      <Card className="p-4">
        <p className="mb-3 text-sm text-stone-600">
          Bölümünüzün depodan isteyeceği malzemeleri buradan gönderin. Talep, depocunun <b>Depo talepleri</b> ekranına düşer.
        </p>
        <div className="grid gap-3 sm:grid-cols-2">
          <Field label="Bölüm">
            <select className={inputCls} value={dept} onChange={(e) => setDept(e.target.value)}>
              {allowed.map((id) => <option key={id} value={id}>{deptName(cfg, id)}</option>)}
            </select>
          </Field>
          <Field label="İhtiyaç tarihi" hint="Varsayılan bugün">
            <input type="date" className={inputCls} value={date} onChange={(e) => setDate(e.target.value)} />
          </Field>
        </div>
        {existing && tab === "yeni" && (
          <p className="mt-3 rounded bg-amber-50 px-3 py-2 text-sm text-amber-900">
            {dateTR(date)} için bu bölümün {existing.lines.length} kalemlik talebi zaten var; gireceğiniz kalemler ona eklenir.
          </p>
        )}
      </Card>

      <Tabs value={tab} onChange={setTab} items={[["yeni", "Yeni talep"], ["gecmis", "Gönderdiklerim"]]} />

      {tab === "yeni" && (
        <Card className="p-4">
          <input className={inputCls} placeholder="Malzeme ara" value={q} onChange={(e) => setQ(e.target.value)} />
          <div className="mt-2 divide-y divide-stone-100">
            {list.map((c) => (
              <div key={c.id} className="flex items-center gap-3 py-2">
                <div className="min-w-0 flex-1">
                  <div className="text-sm font-medium leading-tight text-stone-800">{c.n}</div>
                  <div className="text-xs text-stone-500">{c.c || c.id} · {c.u}</div>
                </div>
                <input type="number" min="0" step="any" inputMode="decimal" className={numCls} placeholder="0"
                  value={cart[c.id] || ""} onChange={(e) => setCart((s) => ({ ...s, [c.id]: e.target.value }))} />
              </div>
            ))}
            {!list.length && <p className="py-6 text-center text-sm text-stone-500">
              Bu bölüm için depo talep listesinde kalem yok. Yönetim &gt; Ürünler ekranından kalem açabilirsiniz.
            </p>}
          </div>
        </Card>
      )}

      {tab === "gecmis" && (
        <div className="space-y-4">
          <FilterBar f={f} label="İhtiyaç tarihi" />
          {!mine.length ? <Empty title="Seçilen aralıkta gönderilmiş talep yok" /> : mine.map((o) => (
            <Card key={o.id} className="overflow-x-auto">
              <div className="border-b border-stone-100 px-3 py-2 font-medium text-stone-900">
                {branchName(cfg, o.branchId)} · {dateTR(o.prodDate)}
              </div>
              <table className="w-full min-w-max text-sm">
                <thead className="bg-stone-50 text-left text-xs text-stone-600"><tr>
                  <th className="px-3 py-2 font-medium">Malzeme</th><th className="px-3 py-2 text-right font-medium">İstenen</th>
                  <th className="px-3 py-2 text-right font-medium">Hazırlanan</th><th className="px-3 py-2 text-right font-medium">Gönderilen</th>
                  <th className="px-3 py-2 text-right font-medium">Durum</th></tr></thead>
                <tbody className="divide-y divide-stone-100">
                  {o.lines.map((l) => (
                    <tr key={l.id}>
                      <td className="px-3 py-2 text-stone-800">{l.name}</td>
                      <td className="px-3 py-2 text-right tabular-nums">{fmt(l.qty)}</td>
                      <td className="px-3 py-2 text-right tabular-nums">{fmt(prep(l))}</td>
                      <td className="px-3 py-2 text-right tabular-nums">{fmt(sentQ(l))}</td>
                      <td className="px-3 py-2 text-right"><Badge status={lineStatus(l)} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Card>
          ))}
        </div>
      )}

      {tab === "yeni" && (
        <div className="fixed inset-x-0 bottom-0 z-20 border-t border-stone-200 bg-white px-4 py-3">
          <div className="mx-auto flex max-w-6xl items-center justify-between gap-3">
            <div className="text-sm text-stone-600">{lines.length} kalem seçildi</div>
            <div className="flex items-center gap-2">
              {msg && <span className="text-sm text-emerald-700">{msg}</span>}
              <Btn onClick={submit} disabled={!lines.length}>Depoya gönder</Btn>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ============================ sevk ============================ */

function ShipScreen({ me, cfg, orders, onUpdateLine, onAddFire, ensure }) {
  const allowed = me.depts && me.depts.length ? me.depts : cfg.depts.map((d) => d.id);
  const f = useFilters({ from: today(), to: addDays(today(), 1), dept: allowed.length === 1 ? allowed[0] : "all", branch: "all" });
  const fh = useFilters({ from: daysAgo(29), to: today(), branch: "all" });
  const [tab, setTab] = useState("bekleyen");
  const [draft, setDraft] = useState({});
  const [sentDate, setSentDate] = useState(today());
  const [msg, setMsg] = useState("");
  const [ask, setAsk] = useState(null);
  const [fireFor, setFireFor] = useState(null);
  const A = f.applied;
  const activeDepts = A.dept === "all" ? allowed : [A.dept];
  useEffect(() => { ensure(addDays(A.from, -3), A.to); }, [A.from, A.to, ensure]);
  useEffect(() => { ensure(addDays(fh.applied.from, -3), fh.applied.to); }, [fh.applied.from, fh.applied.to, ensure]);

  const pending = useMemo(() => {
    const out = [];
    orders.forEach((o) => {
      if (!inRange(o.shipDate || o.prodDate, A.from, A.to)) return;
      if (A.branch !== "all" && o.branchId !== A.branch) return;
      o.lines.forEach((l) => { if (activeDepts.includes(l.deptId) && lineStatus(l) !== "gonderildi") out.push({ o, l }); });
    });
    return out.sort((a, b) => ((a.o.shipDate || a.o.prodDate) < (b.o.shipDate || b.o.prodDate) ? -1 : 1));
  }, [orders, A.from, A.to, A.branch, activeDepts.join()]);

  const shipped = useMemo(() => {
    const out = [];
    orders.forEach((o) => {
      o.lines.forEach((l) => {
        if (!activeDepts.includes(l.deptId)) return;
        if (!["gonderildi", "kismen"].includes(lineStatus(l))) return;
        const d = l.sentDate || o.shipDate || o.prodDate;
        if (!inRange(d, fh.applied.from, fh.applied.to)) return;
        if (fh.applied.branch !== "all" && o.branchId !== fh.applied.branch) return;
        out.push({ o, l, d });
      });
    });
    return out.sort((a, b) => (a.d < b.d ? 1 : -1));
  }, [orders, fh.applied.from, fh.applied.to, fh.applied.branch, activeDepts.join()]);

  const byBranch = {};
  pending.forEach((r) => { (byBranch[r.o.branchId] = byBranch[r.o.branchId] || []).push(r); });

  const commit = async (list, reasons, note) => {
    for (const { row, s } of list) {
      const patch = { sent: s, status: s >= row.l.qty ? "gonderildi" : "kismen", sentDate };
      if (s < row.l.qty) { patch.shortReason = reasons?.[row.l.id] || "Diğer"; patch.shortNote = note || ""; }
      await onUpdateLine(row.o, row.l.id, patch);
    }
    setDraft({}); setMsg(`${list.length} kalem ${dateTR(sentDate)} tarihiyle kaydedildi`);
    setTimeout(() => setMsg(""), 4000);
  };

  const startSave = () => {
    if (!sentDate) { setMsg("Sevk tarihi zorunlu"); return; }
    const parsed = [];
    Object.entries(draft).forEach(([lid, v]) => {
      if (v === "" || v === undefined) return;
      const row = pending.find((r) => r.l.id === lid);
      if (!row) return;
      parsed.push({ row, s: Math.max(0, Math.min(Number(v), row.l.qty)) });
    });
    if (!parsed.length) return;
    const short = parsed.filter((x) => x.s < x.row.l.qty);
    if (!short.length) return commit(parsed, {}, "");
    const reasons = {}; short.forEach((x) => { reasons[x.row.l.id] = x.row.l.shortReason || SHORT_REASONS[0]; });
    setAsk({ all: parsed, short, reasons, note: "" });
  };

  const nDraft = Object.values(draft).filter((v) => v !== "" && v !== undefined).length;

  return (
    <div className="space-y-4 pb-24">
      <FilterBar f={f} onApply={(d) => ensure(d.from, d.to)} label="Sevk tarihi">
        {allowed.length > 1 && <DeptSelect f={f} cfg={cfg} allowed={allowed} />}<BranchSelect f={f} cfg={cfg} />
      </FilterBar>

      <Tabs value={tab} onChange={setTab} items={[["bekleyen", "Gönderilecekler"], ["gonderilen", "Gönderilenler"]]} />

      {tab === "bekleyen" && (!pending.length ? <Empty title="Bu filtrede gönderilecek kalem yok" /> :
        Object.entries(byBranch).map(([bid, list]) => (
          <Card key={bid} className="overflow-x-auto">
            <div className="border-b border-stone-100 px-3 py-2 font-medium text-stone-900">{branchName(cfg, bid)} · {list.length} kalem</div>
            <table className="w-full min-w-max text-sm">
              <thead className="bg-stone-50 text-left text-xs text-stone-600"><tr>
                <th className="px-3 py-2 font-medium">Ürün</th><th className="px-3 py-2 text-right font-medium">Talep</th>
                <th className="px-3 py-2 text-right font-medium">Hazırlanan</th><th className="px-3 py-2 text-right font-medium">Daha önce</th>
                <th className="px-3 py-2 text-right font-medium">Gönderilen</th><th className="px-3 py-2 text-right font-medium">Fire</th><th className="px-3 py-2 text-right font-medium">Durum</th></tr></thead>
              <tbody className="divide-y divide-stone-100">
                {list.map(({ o, l }) => (
                  <tr key={l.id}>
                    <td className="px-3 py-2"><div className="text-stone-800">{l.name}</div>
                      <div className="text-xs text-stone-500">{deptName(cfg, l.deptId)} · üretim {dateTR(o.prodDate)}</div>
                      {l.note && <div className="text-xs text-emerald-900">Not: {l.note}</div>}
                      {l.img && <img src={l.img} alt="" className="mt-1 h-10 w-10 rounded border border-stone-300 object-cover" />}</td>
                    <td className="px-3 py-2 text-right font-medium tabular-nums">{fmt(l.qty)}</td>
                    <td className="px-3 py-2 text-right tabular-nums text-sky-800">{prep(l) ? fmt(prep(l)) : "–"}</td>
                    <td className="px-3 py-2 text-right tabular-nums text-stone-500">{sentQ(l) ? fmt(sentQ(l)) : "–"}</td>
                    <td className="px-3 py-2 text-right"><input type="number" min="0" max={l.qty} inputMode="numeric" className={numCls}
                      placeholder={String(prep(l) || l.qty)} value={draft[l.id] ?? ""} onChange={(e) => setDraft((d) => ({ ...d, [l.id]: e.target.value }))} /></td>
                    <td className="px-3 py-2 text-right"><button className="rounded border border-red-300 px-2 py-1 text-xs text-red-700 hover:bg-red-50" onClick={() => setFireFor({ o, l, qty: "", reason: FIRE_REASONS[3] })}>fire</button></td>
                    <td className="px-3 py-2 text-right"><Badge status={lineStatus(l)} /></td>
                  </tr>
                ))}
              </tbody>
            </table>
            <div className="flex flex-wrap justify-end gap-2 border-t border-stone-100 p-3">
              <Btn variant="ghost" onClick={() => setDraft((d) => { const n = { ...d }; list.forEach(({ l }) => { n[l.id] = String(prep(l) || l.qty); }); return n; })}>Hazırlananı doldur</Btn>
              <Btn variant="ghost" onClick={() => setDraft((d) => { const n = { ...d }; list.forEach(({ l }) => { n[l.id] = String(l.qty); }); return n; })}>Talebin tamamı</Btn>
            </div>
          </Card>
        )))}

      {tab === "gonderilen" && (
        <div className="space-y-4">
          <FilterBar f={fh} onApply={(d) => ensure(d.from, d.to)} label="Sevk tarihi"><BranchSelect f={fh} cfg={cfg} /></FilterBar>
          <Table max={400} head={["Sevk", "Şube", "Bölüm", "Ürün", "Talep", "Hazırlanan", "Gönderilen", "Eksik", "Sebep", "Durum"]}
            rows={shipped.map(({ o, l, d }) => [dateTR(d), branchName(cfg, o.branchId), deptName(cfg, l.deptId), l.name,
              fmt(l.qty), fmt(prep(l)), fmt(sentQ(l)), openQty(l) ? fmt(openQty(l)) : "–", l.shortReason || "—", STATUS[lineStatus(l)].label])} />
        </div>
      )}

      {tab === "bekleyen" && (
        <div className="fixed inset-x-0 bottom-0 z-20 border-t border-stone-200 bg-white px-4 py-3">
          <div className="mx-auto flex max-w-6xl flex-wrap items-center justify-between gap-3">
            <div className="flex items-center gap-2 text-sm text-stone-600">
              <span className="font-medium text-stone-800">Sevk tarihi</span>
              <input type="date" className="rounded-md border border-stone-300 px-2 py-1.5 text-sm" value={sentDate} onChange={(e) => setSentDate(e.target.value)} />
              <span className="hidden sm:inline">· {nDraft} kalem</span>
            </div>
            <div className="flex items-center gap-2">{msg && <span className="text-sm text-emerald-700">{msg}</span>}
              <Btn onClick={startSave} disabled={!nDraft || !sentDate}>Gönderildi olarak kaydet</Btn></div>
          </div>
        </div>
      )}

      {ask && (
        <Modal wide title="Eksik gönderim sebebi" onClose={() => setAsk(null)}
          footer={<><Btn variant="ghost" onClick={() => setAsk(null)}>Vazgeç</Btn><Btn onClick={() => { commit(ask.all, ask.reasons, ask.note); setAsk(null); }}>Kaydet</Btn></>}>
          <p className="mb-3 text-sm text-stone-600">{ask.short.length} kalem talebin altında gönderiliyor. Hepsine birden uygulamak için üstteki düğmeleri kullanın.</p>
          <div className="mb-4 flex flex-wrap gap-2">
            {SHORT_REASONS.map((r) => <button key={r} onClick={() => setAsk((s) => { const n = { ...s.reasons }; s.short.forEach((x) => { n[x.row.l.id] = r; }); return { ...s, reasons: n }; })}
              className="rounded-full border border-emerald-700 bg-emerald-50 px-3 py-1.5 text-sm text-emerald-900 hover:bg-emerald-100">{r}</button>)}
          </div>
          <div className="divide-y divide-stone-100 rounded-md border border-stone-200">
            {ask.short.map(({ row, s }) => (
              <div key={row.l.id} className="flex flex-wrap items-center gap-2 p-3">
                <div className="min-w-0 flex-1"><div className="text-sm font-medium text-stone-800">{row.l.name}</div>
                  <div className="text-xs text-stone-500">{branchName(cfg, row.o.branchId)} · talep {fmt(row.l.qty)} · gönderilen {fmt(s)} · <span className="text-orange-700">eksik {fmt(row.l.qty - s)}</span></div></div>
                <select className="rounded-md border border-stone-300 px-2 py-1.5 text-sm" value={ask.reasons[row.l.id]}
                  onChange={(e) => setAsk((st) => ({ ...st, reasons: { ...st.reasons, [row.l.id]: e.target.value } }))}>
                  {SHORT_REASONS.map((r) => <option key={r} value={r}>{r}</option>)}</select>
              </div>
            ))}
          </div>
          <div className="mt-3"><Field label="Açıklama (isteğe bağlı)"><input className={inputCls} value={ask.note} onChange={(e) => setAsk((s) => ({ ...s, note: e.target.value }))} placeholder="örn. araçta kaldı" /></Field></div>
        </Modal>
      )}

      {fireFor && (
        <Modal title="Sevkiyat firesi" onClose={() => setFireFor(null)}
          footer={<><Btn variant="ghost" onClick={() => setFireFor(null)}>Vazgeç</Btn>
            <Btn variant="danger" disabled={!(Number(fireFor.qty) > 0)} onClick={async () => {
              const c = catFind(cfg, fireFor.l.productId);
              await onAddFire([{ id: uid(), date: sentDate, deptId: fireFor.l.deptId, itemId: fireFor.l.productId, name: fireFor.l.name,
                unit: c?.u || "Adet", kind: "mamul", qty: Number(fireFor.qty), reason: fireFor.reason, by: me.name, at: new Date().toISOString() }]);
              setFireFor(null); setMsg("Fire kaydedildi"); setTimeout(() => setMsg(""), 3000);
            }}>Fireyi kaydet</Btn></>}>
          <p className="mb-3 text-sm text-stone-600">{fireFor.l.name} · {branchName(cfg, fireFor.o.branchId)} · {dateTR(sentDate)}</p>
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="Fire adedi"><input type="number" min="1" inputMode="numeric" className={inputCls} value={fireFor.qty} onChange={(e) => setFireFor((s) => ({ ...s, qty: e.target.value }))} /></Field>
            <Field label="Sebep"><select className={inputCls} value={fireFor.reason} onChange={(e) => setFireFor((s) => ({ ...s, reason: e.target.value }))}>
              {FIRE_REASONS.map((r) => <option key={r} value={r}>{r}</option>)}</select></Field>
          </div>
          <p className="mt-3 text-xs text-stone-500">Fire kaydı, gönderilen adedi kendiliğinden değiştirmez. Eksik gidecekse gönderilen kutusuna gerçek adedi yazın.</p>
        </Modal>
      )}
    </div>
  );
}

/* ============================ talep listesi (şoför / patron) ============================ */

function DriverList({ cfg, orders, ensure }) {
  const f = useFilters({ from: today(), to: addDays(today(), 1), branch: "all" });
  const [dept, setDept] = useState("all");
  const A = f.applied;
  useEffect(() => { ensure(addDays(A.from, -3), A.to); }, [A.from, A.to, ensure]);
  const rows = useMemo(() => {
    const out = [];
    orders.forEach((o) => {
      if (!inRange(o.shipDate || o.prodDate, A.from, A.to)) return;
      if (A.branch !== "all" && o.branchId !== A.branch) return;
      o.lines.forEach((l) => { if (dept === "all" || l.deptId === dept) out.push({ o, l }); });
    });
    return out;
  }, [orders, A.from, A.to, A.branch, dept]);
  const byBranch = {};
  rows.forEach((r) => { (byBranch[r.o.branchId] = byBranch[r.o.branchId] || []).push(r); });
  return (
    <div className="space-y-4">
      <FilterBar f={f} onApply={(d) => ensure(d.from, d.to)} label="Sevk tarihi"><BranchSelect f={f} cfg={cfg} /></FilterBar>
      <div className="-mx-4 overflow-x-auto px-4"><div className="flex gap-2 pb-1">
        <button onClick={() => setDept("all")} className={`whitespace-nowrap rounded-full border px-3 py-1.5 text-sm ${dept === "all" ? "border-emerald-800 bg-emerald-800 text-white" : "border-stone-300 bg-white text-stone-700"}`}>Hepsi</button>
        {cfg.depts.map((d) => <button key={d.id} onClick={() => setDept(d.id)}
          className={`whitespace-nowrap rounded-full border px-3 py-1.5 text-sm ${dept === d.id ? "border-emerald-800 bg-emerald-800 text-white" : "border-stone-300 bg-white text-stone-700"}`}>{d.name}</button>)}
      </div></div>
      {!rows.length ? <Empty title="Bu filtrede talep yok" hint="Tarihi değiştirip Ara'ya basın." /> :
        Object.entries(byBranch).map(([bid, list]) => (
          <Card key={bid}>
            <div className="border-b border-stone-100 px-4 py-2 font-medium text-stone-900">{branchName(cfg, bid)}</div>
            <ul className="divide-y divide-stone-100">
              {list.map(({ o, l }) => (
                <li key={l.id} className="flex items-center gap-3 px-4 py-2">
                  <div className="min-w-0 flex-1"><div className="text-sm text-stone-800">{l.name}</div>
                    <div className="text-xs text-stone-500">{deptName(cfg, l.deptId)} · sevk {dateTR(o.shipDate || o.prodDate)}</div></div>
                  <span className="shrink-0 rounded bg-stone-100 px-2 py-1 text-sm font-medium tabular-nums text-stone-800">{fmt(l.qty)}</span>
                  <Badge status={lineStatus(l)} />
                </li>
              ))}
            </ul>
          </Card>
        ))}
    </div>
  );
}

/* ============================ üretim ve fire girişi ============================ */

function EntryScreen({ me, cfg, orders, production, fires, onAdd, onAddFire, onDelete, onDeleteFire, canProduce, canFire, ensure }) {
  const allowed = me.depts && me.depts.length ? me.depts : cfg.depts.map((d) => d.id);
  const [dept, setDept] = useState(allowed[0]);
  const [date, setDate] = useState(today());
  const [q, setQ] = useState("");
  const [draft, setDraft] = useState({});
  const [fireDraft, setFireDraft] = useState({});
  const [fireOpen, setFireOpen] = useState({});
  const [recipe, setRecipe] = useState(null);
  const [msg, setMsg] = useState("");
  useEffect(() => { ensure(date, date); }, [date, ensure]);

  const all = useMemo(() => cfg.catalog.filter((p) => p.vu && p.d === dept).sort((a, b) => a.o - b.o), [cfg.catalog, dept]);
  const filtered = useMemo(() => { const t = lower(q.trim()); return t ? all.filter((p) => lower(p.n).includes(t)) : all; }, [all, q]);
  const mamul = filtered.filter((p) => p.k === "mamul");
  const yari = filtered.filter((p) => p.k !== "mamul");

  /* o günün talebi — katalog kaleminin kendisi veya bağlı olduğu talep ürünü üzerinden */
  const demand = useMemo(() => {
    const m = new Map();
    orders.forEach((o) => { if (o.prodDate === date) o.lines.forEach((l) => { if (l.deptId === dept) m.set(l.productId, (m.get(l.productId) || 0) + l.qty); }); });
    return m;
  }, [orders, date, dept]);
  const demandOf = (p) => (demand.get(p.id) || 0) + (p.l ? demand.get(p.l) || 0 : 0);

  const dayProd = production.filter((r) => r.date === date && r.deptId === dept && !r.seed);
  const dayFire = fires.filter((r) => r.date === date && r.deptId === dept && !r.seed);
  const madeOf = (p) => dayProd.filter((r) => r.itemId === p.id).reduce((a, r) => a + r.qty, 0);
  const nP = Object.values(draft).filter((v) => Number(v) > 0).length;
  const nF = Object.values(fireDraft).filter((v) => Number(v && v.q) > 0).length;

  const submit = async () => {
    const pRows = Object.entries(draft).filter(([, v]) => Number(v) > 0);
    const fRows = Object.entries(fireDraft).filter(([, v]) => Number(v && v.q) > 0);
    if (!pRows.length && !fRows.length) return;
    const base = (iid) => {
      const it = catFind(cfg, iid);
      return { id: uid(), date, deptId: dept, itemId: iid, name: it?.n || iid, unit: it?.u || "Adet", kind: it?.k || "mamul", by: me.name, at: new Date().toISOString() };
    };
    if (pRows.length && canProduce) await onAdd(date, pRows.map(([iid, v]) => ({ ...base(iid), qty: Number(v) })));
    if (fRows.length && canFire) await onAddFire(fRows.map(([iid, v]) => ({ ...base(iid), qty: Number(v.q), reason: (v.note || "").trim() || "Sebep yazılmadı" })), date);
    setDraft({}); setFireDraft({}); setFireOpen({});
    setMsg(`${canProduce ? pRows.length + " üretim, " : ""}${fRows.length} fire kaydedildi`);
    setTimeout(() => setMsg(""), 3000);
  };

  const renderRow = (p, withProd) => {
    const d = withProd ? demandOf(p) : 0;
    const made = withProd ? madeOf(p) : 0;
    const diff = made + (Number(draft[p.id]) || 0) - d;
    const fd = fireDraft[p.id] || { q: "", note: "" };
    const logged = dayFire.filter((r) => r.itemId === p.id);
    const loggedQty = logged.reduce((a, r) => a + r.qty, 0);
    const open = !!fireOpen[p.id];
    const marked = open || Number(fd.q) > 0 || loggedQty > 0;
    return (
      <div key={p.id} className="py-2">
        <div className="flex items-center gap-2">
          <div className="min-w-0 flex-1">
            <div className="text-sm font-medium leading-tight text-stone-800">{p.n}</div>
            <div className="mt-0.5 flex flex-wrap items-center gap-x-2 gap-y-1 text-xs text-stone-500">
              <span>{p.u}</span>
              {withProd && d > 0 && (
                <span>· talep <span className="font-medium text-stone-700">{fmt(d)}</span>
                  {made > 0 && <> · girilen {fmt(made)}</>}
                  <span className={diff < 0 ? "text-orange-700" : diff > 0 ? "text-sky-700" : "text-emerald-700"}> · {diff === 0 ? "tam" : diff > 0 ? `+${fmt(diff)} fazla` : `${fmt(-diff)} eksik`}</span>
                </span>
              )}
              {RECIPES[p.id] && <button type="button" onClick={() => setRecipe(p)} className="text-emerald-800 hover:underline">reçete</button>}
              {canFire && (
                <button type="button" onClick={() => setFireOpen((x) => ({ ...x, [p.id]: !x[p.id] }))}
                  className={`rounded border px-1.5 py-0.5 text-[11px] font-medium ${marked ? "border-red-400 bg-red-50 text-red-800" : "border-stone-300 text-stone-600 hover:bg-stone-50"}`}>
                  {loggedQty > 0 ? `fire ${fmt(loggedQty)}` : "+ fire"}
                </button>
              )}
            </div>
          </div>
          {withProd && canProduce && (
            <input type="number" min="0" step="any" inputMode="decimal" className={numCls} placeholder="0"
              value={draft[p.id] || ""} onChange={(e) => setDraft((x) => ({ ...x, [p.id]: e.target.value }))} />
          )}
        </div>

        {canFire && open && (
          <div className="mt-2 rounded-md border border-red-200 bg-red-50 p-2">
            <div className="flex flex-wrap items-center gap-2">
              <input type="number" min="0" step="any" inputMode="decimal" placeholder="adet"
                className="w-20 shrink-0 rounded-md border border-red-300 bg-white px-2 py-1.5 text-center text-base tabular-nums focus:border-red-500 focus:outline-none"
                value={fd.q} onChange={(e) => setFireDraft((x) => ({ ...x, [p.id]: { ...fd, q: e.target.value } }))} />
              <input placeholder="Sorun neydi? (örn. fırında yandı)"
                className="min-w-0 flex-1 rounded-md border border-red-300 bg-white px-2 py-1.5 text-sm focus:border-red-500 focus:outline-none"
                value={fd.note} onChange={(e) => setFireDraft((x) => ({ ...x, [p.id]: { ...fd, note: e.target.value } }))} />
            </div>
            {logged.length > 0 && (
              <ul className="mt-2 space-y-1 border-t border-red-200 pt-2">
                {logged.map((r) => (
                  <li key={r.id} className="flex items-center justify-between gap-2 text-xs text-red-900">
                    <span className="min-w-0 truncate">{fmt(r.qty)} {r.unit} — {r.reason}</span>
                    <button className="shrink-0 hover:underline" onClick={() => onDeleteFire(r.id, r.date)}>sil</button>
                  </li>
                ))}
              </ul>
            )}
          </div>
        )}
      </div>
    );
  };

  const unmatched = demand.size > 0 && mamul.every((p) => demandOf(p) === 0);

  return (
    <div className="space-y-4 pb-20">
      <Card className="p-4">
        <div className="grid gap-3 sm:grid-cols-2">
          <Field label="Bölüm"><select className={inputCls} value={dept} onChange={(e) => setDept(e.target.value)}>
            {allowed.map((id) => <option key={id} value={id}>{deptName(cfg, id)}</option>)}</select></Field>
          <Field label="Üretim tarihi" hint="Bu tarihe düşen şube talepleri satırlarda görünür"><input type="date" className={inputCls} value={date} onChange={(e) => setDate(e.target.value)} /></Field>
        </div>
        {canFire && <p className="mt-3 text-xs text-stone-500">Fire girmek için ürünün altındaki <span className="rounded border border-stone-300 px-1 py-0.5">+ fire</span> düğmesine basın, adedi ve sebebini oraya yazın.</p>}
        {unmatched && <p className="mt-3 rounded bg-amber-50 px-3 py-2 text-xs text-amber-900">
          Bu tarihte bölüme talep var ama üretim kalemleriyle eşleşmiyor. Yönetim &gt; Ürün kataloğu'ndan üretim kalemini bir talep ürününe bağlayın.
        </p>}
      </Card>

      <Card className="p-4">
        <input className={inputCls} placeholder="Listede ara" value={q} onChange={(e) => setQ(e.target.value)} />
        <div className="mt-3 flex items-center gap-2 border-b border-stone-200 pb-1 text-xs text-stone-500">
          <span className="flex-1">Ürün</span>
          {canProduce && <span className="w-[4.5rem] shrink-0 text-center">Üretim</span>}
        </div>
        <div className="divide-y divide-stone-100">
          {mamul.map((p) => renderRow(p, true))}
          {!mamul.length && <p className="py-6 text-center text-sm text-stone-500">Bu bölümde üretim listesine açık kalem yok.</p>}
        </div>
        {canFire && yari.length > 0 && (<>
          <div className="mt-4 rounded bg-red-50 px-3 py-2 text-sm font-medium text-red-900">Yarı mamul ve diğer fire kalemleri</div>
          <div className="divide-y divide-stone-100">{yari.map((p) => renderRow(p, false))}</div>
        </>)}
      </Card>

      <Card className="p-4">
        <h3 className="mb-2 font-medium text-stone-900">{dateTR(date)} · bu ekrandan girilen kayıtlar</h3>
        {!dayProd.length && !dayFire.length ? <p className="text-sm text-stone-500">Bu tarihte kayıt yok.</p> : (
          <table className="w-full text-sm"><tbody className="divide-y divide-stone-100">
            {dayProd.map((r) => (<tr key={r.id}><td className="py-2 text-stone-800">{r.name}</td><td className="py-2 text-right text-xs text-stone-500">üretim · {r.by}</td>
              <td className="py-2 pl-3 text-right font-medium tabular-nums">{fmt(r.qty)} {r.unit}</td>
              <td className="py-2 pl-2 text-right"><button className="text-xs text-red-700 hover:underline" onClick={() => onDelete(r.id, r.date)}>sil</button></td></tr>))}
            {dayFire.map((r) => (<tr key={r.id} className="bg-red-50"><td className="py-2 text-stone-800">{r.name}</td>
              <td className="py-2 text-right text-xs text-red-700">fire · {r.reason}</td>
              <td className="py-2 pl-3 text-right font-medium tabular-nums text-red-700">{fmt(r.qty)} {r.unit}</td>
              <td className="py-2 pl-2 text-right"><button className="text-xs text-red-700 hover:underline" onClick={() => onDeleteFire(r.id, r.date)}>sil</button></td></tr>))}
          </tbody></table>
        )}
      </Card>

      {recipe && <RecipeModal cfg={cfg} item={recipe} onClose={() => setRecipe(null)} />}

      <div className="fixed inset-x-0 bottom-0 z-20 border-t border-stone-200 bg-white px-4 py-3">
        <div className="mx-auto flex max-w-6xl items-center justify-between gap-3">
          <div className="text-sm text-stone-600">{canProduce && <>{nP} üretim · </>}<span className="text-red-700">{nF} fire</span></div>
          <div className="flex items-center gap-2">{msg && <span className="text-sm text-emerald-700">{msg}</span>}
            <Btn onClick={submit} disabled={!nP && !nF}>Kaydet</Btn></div>
        </div>
      </div>
    </div>
  );
}

/* ============================ depo stoğu ve satınalma ============================ */

const stockOf = (cfg, levels, id) =>
  levels && levels[id] !== undefined ? Number(levels[id]) : Number(catFind(cfg, id)?.ss) || 0;

function ReqBadge({ status }) {
  const s = REQ_STATUS[status] || REQ_STATUS.talep;
  return <span className={`inline-block rounded px-2 py-0.5 text-xs font-medium whitespace-nowrap ${s.cls}`}>{s.label}</span>;
}

/* --- yeni satınalma talebi --- */
function NewRequestModal({ cfg, levels, preset, onClose, onSave }) {
  const [q, setQ] = useState("");
  const [cart, setCart] = useState(() => (preset ? { [preset.id]: { qty: "", urgency: URGENCY[0], note: "" } } : {}));
  const list = useMemo(() => {
    const t = lower(q.trim());
    return cfg.catalog.filter((c) => c.vd)
      .filter((c) => (t ? lower(c.n).includes(t) || lower(String(c.c || c.id)).includes(t) : true))
      .slice(0, 300);
  }, [cfg.catalog, q]);
  const chosen = Object.entries(cart).filter(([, v]) => Number(v.qty) > 0);
  const set = (id, patch) => setCart((s) => ({ ...s, [id]: { qty: "", urgency: URGENCY[0], note: "", ...(s[id] || {}), ...patch } }));

  return (
    <Modal wide title="Satınalmaya talep aç" onClose={onClose}
      footer={<><Btn variant="ghost" onClick={onClose}>Vazgeç</Btn>
        <Btn disabled={!chosen.length} onClick={() => { onSave(chosen.map(([id, v]) => ({ id, ...v }))); onClose(); }}>
          {chosen.length ? `${chosen.length} kalemi gönder` : "Gönder"}</Btn></>}>
      <input className={inputCls} placeholder="Malzeme ara" value={q} onChange={(e) => setQ(e.target.value)} />
      <div className="mt-3 max-h-80 divide-y divide-stone-100 overflow-y-auto">
        {list.map((c) => {
          const v = cart[c.id] || { qty: "", urgency: URGENCY[0], note: "" };
          const stok = stockOf(cfg, levels, c.id);
          const min = Number(c.ms) || 0;
          return (
            <div key={c.id} className="py-2">
              <div className="flex items-center gap-2">
                <div className="min-w-0 flex-1">
                  <div className="text-sm font-medium text-stone-800">{c.n}</div>
                  <div className="text-xs text-stone-500">
                    {c.c || c.id} · stok {fmt(stok)} {c.u}{min ? ` · min ${fmt(min)}` : ""}
                    {min > 0 && stok < min && <span className="ml-1 font-medium text-red-700">kritik</span>}
                  </div>
                </div>
                <input type="number" min="0" step="any" inputMode="decimal" placeholder="miktar" className={numCls}
                  value={v.qty} onChange={(e) => set(c.id, { qty: e.target.value })} />
              </div>
              {Number(v.qty) > 0 && (
                <div className="mt-2 flex flex-wrap items-center gap-2">
                  <select className="rounded-md border border-stone-300 px-2 py-1.5 text-sm" value={v.urgency} onChange={(e) => set(c.id, { urgency: e.target.value })}>
                    {URGENCY.map((u) => <option key={u} value={u}>{u}</option>)}
                  </select>
                  <input className="min-w-0 flex-1 rounded-md border border-stone-300 px-2 py-1.5 text-sm" placeholder="Not (marka, tedarikçi, açıklama)"
                    value={v.note} onChange={(e) => set(c.id, { note: e.target.value })} />
                </div>
              )}
            </div>
          );
        })}
        {!list.length && <p className="py-6 text-center text-sm text-stone-500">Eşleşen malzeme yok.</p>}
      </div>
    </Modal>
  );
}

/* --- depo stoğu --- */
function DepotStock({ cfg, levels, moves, requests, ensureStock, ensureMoves, ensureReq, onCount, onRequest, canCount, canRequest }) {
  const [tab, setTab] = useState("liste");
  const [q, setQ] = useState("");
  const [only, setOnly] = useState("all");
  const [recipe, setRecipe] = useState(null);
  const [reqFor, setReqFor] = useState(null);
  const [counts, setCounts] = useState({});
  const [msg, setMsg] = useState("");
  const fm = useFilters({ from: daysAgo(13), to: today() });

  useEffect(() => { ensureStock(); ensureReq(monthStart(), today()); }, [ensureStock, ensureReq]);
  useEffect(() => { if (tab === "hareket") ensureMoves(fm.applied.from, fm.applied.to); }, [tab, fm.applied.from, fm.applied.to, ensureMoves]);

  const openReq = useMemo(() => {
    const m = new Map();
    requests.forEach((r) => {
      if (r.status === "teslim" || r.status === "iptal") return;
      m.set(r.itemId, (m.get(r.itemId) || 0) + (Number(r.qty) || 0) - (Number(r.recvQty) || 0));
    });
    return m;
  }, [requests]);

  const rows = useMemo(() => {
    const t = lower(q.trim());
    return cfg.catalog.filter((c) => c.vd)
      .filter((c) => (t ? lower(c.n).includes(t) || lower(String(c.c || c.id)).includes(t) : true))
      .map((c) => {
        const stok = stockOf(cfg, levels, c.id), min = Number(c.ms) || 0;
        return { ...c, stok, min, kritik: min > 0 && stok < min, acik: openReq.get(c.id) || 0 };
      })
      .filter((c) => (only === "kritik" ? c.kritik : only === "talepli" ? c.acik > 0 : true))
      .sort((a, b) => (a.kritik === b.kritik ? lower(a.n).localeCompare(lower(b.n), "tr") : a.kritik ? -1 : 1));
  }, [cfg.catalog, levels, q, only, openReq]);

  const toplam = cfg.catalog.filter((c) => c.vd).length;
  const kritik = cfg.catalog.filter((c) => c.vd).filter((c) => { const s = stockOf(cfg, levels, c.id), m = Number(c.ms) || 0; return m > 0 && s < m; }).length;
  const nCount = Object.entries(counts).filter(([, v]) => v !== "" && v !== undefined).length;

  const saveCounts = async () => {
    const rowsOut = Object.entries(counts).filter(([, v]) => v !== "" && v !== undefined).map(([id, v]) => {
      const c = catFind(cfg, id);
      return { itemId: id, name: c?.n || id, unit: c?.u || "", type: "sayim", counted: Number(v), note: "Depo sayımı", date: today() };
    });
    if (!rowsOut.length) return;
    await onCount(rowsOut);
    setCounts({});
    setMsg(`${rowsOut.length} kalemde sayım kaydedildi`);
    setTimeout(() => setMsg(""), 4000);
  };

  const mvRows = useMemo(() => moves
    .filter((r) => inRange(r.date, fm.applied.from, fm.applied.to))
    .slice().sort((a, b) => (a.at < b.at ? 1 : -1)), [moves, fm.applied.from, fm.applied.to]);

  const TYPE = { giris: "Giriş", cikis: "Çıkış", sayim: "Sayım düzeltmesi" };

  return (
    <div className={tab === "sayim" ? "space-y-4 pb-20" : "space-y-4"}>
      <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
        <Stat label="Depo kalemi" value={fmt(toplam)} sub="katalogda Depo işaretli" />
        <Stat label="Kritik seviyede" value={fmt(kritik)} sub="min. stoğun altında" tone={kritik ? "red" : "emerald"} />
        <Stat label="Açık talep" value={fmt([...openReq.values()].filter((v) => v > 0).length)} sub="satınalmada bekleyen kalem" tone="sky" />
        <Stat label="Listelenen" value={fmt(rows.length)} sub="filtreye uyan" />
      </div>

      <Tabs value={tab} onChange={setTab} items={[["liste", "Stok listesi"], ...(canCount ? [["sayim", "Depo sayımı"]] : []), ["hareket", "Stok hareketleri"]]} />

      {tab !== "hareket" && (
        <Card className="p-4">
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="Ara"><input className={inputCls} placeholder="Malzeme adı veya kodu" value={q} onChange={(e) => setQ(e.target.value)} /></Field>
            {tab === "liste" && <Field label="Görünüm"><select className={inputCls} value={only} onChange={(e) => setOnly(e.target.value)}>
              <option value="all">Bütün kalemler</option><option value="kritik">Yalnızca kritik seviyedekiler</option><option value="talepli">Açık talebi olanlar</option></select></Field>}
          </div>
          <div className="mt-3 flex flex-wrap gap-2">
            {canRequest && tab === "liste" && <Btn onClick={() => setReqFor({})}>Satınalmaya talep aç</Btn>}
            <Btn variant="ghost" disabled={!rows.length} onClick={() => exportXlsx(`depo-stok-${today()}.xlsx`, [{
              name: "Depo stok", head: ["Kod", "Malzeme", "Birim", "Güncel stok", "Min. seviye", "Açık talep", "Durum"],
              rows: rows.map((r) => [r.c || r.id, r.n, r.u, r.stok, r.min, r.acik, r.kritik ? "KRİTİK" : "Yeterli"]),
            }])}>Excel indir</Btn>
          </div>
          {tab === "sayim" && <p className="mt-3 text-xs text-stone-500">Saydığınız miktarı yazın; kaydedince stok o değere çekilir ve fark hareket olarak kaydedilir. Boş bıraktığınız satırlara dokunulmaz.</p>}
        </Card>
      )}

      {tab === "liste" && (
        <Card className="overflow-x-auto">
          <table className="w-full min-w-max text-sm">
            <thead className="bg-stone-50 text-left text-xs text-stone-600"><tr>
              <th className="px-3 py-2 font-medium">Malzeme</th><th className="px-3 py-2 font-medium">Birim</th>
              <th className="px-3 py-2 text-right font-medium">Stok</th><th className="px-3 py-2 text-right font-medium">Min.</th>
              <th className="px-3 py-2 text-right font-medium">Açık talep</th><th className="px-3 py-2 text-right font-medium">Durum</th><th /></tr></thead>
            <tbody className="divide-y divide-stone-100">
              {rows.slice(0, 250).map((r) => (
                <tr key={r.id} className={r.kritik ? "bg-red-50" : ""}>
                  <td className="px-3 py-2"><div className="font-medium text-stone-800">{r.n}</div>
                    <div className="text-xs text-stone-500">{r.c || r.id} · {r.g}
                      {RECIPES[r.id] && <button className="ml-2 text-emerald-800 hover:underline" onClick={() => setRecipe(r)}>reçete</button>}</div></td>
                  <td className="px-3 py-2 text-xs text-stone-500">{r.u}</td>
                  <td className="px-3 py-2 text-right font-medium tabular-nums">{fmt(r.stok)}</td>
                  <td className="px-3 py-2 text-right tabular-nums text-stone-500">{r.min ? fmt(r.min) : "–"}</td>
                  <td className="px-3 py-2 text-right tabular-nums text-sky-800">{r.acik > 0 ? fmt(r.acik) : "–"}</td>
                  <td className="px-3 py-2 text-right">
                    <span className={`rounded px-2 py-0.5 text-xs font-medium ${r.kritik ? "bg-red-200 text-red-900" : "bg-emerald-200 text-emerald-900"}`}>{r.kritik ? "Kritik" : "Yeterli"}</span></td>
                  <td className="px-3 py-2 text-right">{canRequest && <button className="text-xs text-emerald-800 hover:underline" onClick={() => setReqFor(r)}>talep aç</button>}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {!rows.length && <p className="p-6 text-center text-sm text-stone-500">Eşleşen kalem yok.</p>}
          {rows.length > 250 && <p className="border-t border-stone-100 px-3 py-2 text-xs text-stone-500">{fmt(rows.length)} kalemin ilk 250'si gösteriliyor.</p>}
        </Card>
      )}

      {tab === "sayim" && (
        <Card className="divide-y divide-stone-100">
          {rows.slice(0, 250).map((r) => (
            <div key={r.id} className="flex items-center gap-3 p-3">
              <div className="min-w-0 flex-1">
                <div className="text-sm font-medium text-stone-800">{r.n}</div>
                <div className="text-xs text-stone-500">{r.c || r.id} · sistemde {fmt(r.stok)} {r.u}
                  {counts[r.id] !== "" && counts[r.id] !== undefined && Number(counts[r.id]) - r.stok !== 0 && (
                    <span className={Number(counts[r.id]) - r.stok < 0 ? " text-red-700" : " text-sky-700"}>
                      {" "}· fark {Number(counts[r.id]) - r.stok > 0 ? "+" : ""}{fmt(Number(counts[r.id]) - r.stok)}
                    </span>)}
                </div>
              </div>
              <input type="number" min="0" step="any" inputMode="decimal" placeholder="sayılan" className={numCls} disabled={!canCount}
                value={counts[r.id] ?? ""} onChange={(e) => setCounts((s) => ({ ...s, [r.id]: e.target.value }))} />
            </div>
          ))}
          {!rows.length && <p className="p-6 text-center text-sm text-stone-500">Eşleşen kalem yok.</p>}
        </Card>
      )}

      {tab === "hareket" && (
        <div className="space-y-4">
          <FilterBar f={fm} onApply={(d) => ensureMoves(d.from, d.to)} label="Hareket tarihi" />
          <Table max={300} head={["Tarih", "Malzeme", "Hareket", "Miktar", "Sonra", "Not", "Kaydeden"]}
            rows={mvRows.map((r) => [dateTR(r.date), r.name, TYPE[r.type] || r.type,
              `${r.qty > 0 ? "+" : ""}${fmt(r.qty)} ${r.unit || ""}`, fmt(r.after), r.note || "—", r.by || "—"])} />
        </div>
      )}

      {tab === "sayim" && canCount && (
        <div className="fixed inset-x-0 bottom-0 z-20 border-t border-stone-200 bg-white px-4 py-3">
          <div className="mx-auto flex max-w-6xl items-center justify-between gap-3">
            <div className="text-sm text-stone-600">{nCount} kalemde sayım girildi</div>
            <div className="flex items-center gap-2">{msg && <span className="text-sm text-emerald-700">{msg}</span>}
              <Btn onClick={saveCounts} disabled={!nCount}>Sayımı kaydet</Btn></div>
          </div>
        </div>
      )}

      {recipe && <RecipeModal cfg={cfg} item={recipe} onClose={() => setRecipe(null)} />}
      {reqFor && <NewRequestModal cfg={cfg} levels={levels} preset={reqFor.id ? reqFor : null}
        onClose={() => setReqFor(null)} onSave={(rows) => onRequest(rows)} />}
    </div>
  );
}

/* --- satınalma koridoru --- */
function Purchasing({ cfg, me, requests, levels, ensureReq, ensureStock, onAdd, onPatch, onReceive, canBuy, canRequest }) {
  const [tab, setTab] = useState("gun");
  const [newReq, setNewReq] = useState(false);
  const [order, setOrder] = useState(null);
  const [recv, setRecv] = useState(null);
  const [msg, setMsg] = useState("");
  const fg = useFilters({ from: today(), to: today() });
  const fh = useFilters({ from: daysAgo(29), to: today() });

  useEffect(() => { ensureStock(); ensureReq(daysAgo(90), addDays(today(), 1)); }, [ensureStock, ensureReq]);
  useEffect(() => { ensureReq(fh.applied.from, fh.applied.to); }, [fh.applied.from, fh.applied.to, ensureReq]);
  useEffect(() => { ensureReq(daysAgo(90), fg.applied.to); }, [fg.applied.to, ensureReq]);

  const open = useMemo(() => requests.filter((r) => r.status === "talep" || r.status === "siparis" || r.status === "kismi")
    .sort((a, b) => ((a.eta || a.reqDate) < (b.eta || b.reqDate) ? -1 : 1)), [requests]);

  const G = fg.applied.from;
  const bekleyenTalep = open.filter((r) => r.status === "talep");
  const bugunBeklenen = open.filter((r) => r.status !== "talep" && r.eta === G);
  const geciken = open.filter((r) => r.status !== "talep" && r.eta && r.eta < today());
  const bugunGelen = requests.filter((r) => (r.recvDate || "") === G);

  const gecmis = useMemo(() => requests
    .filter((r) => r.status === "teslim" || r.status === "iptal")
    .filter((r) => inRange(r.recvDate || r.reqDate, fh.applied.from, fh.applied.to))
    .sort((a, b) => ((a.recvDate || a.reqDate) < (b.recvDate || b.reqDate) ? 1 : -1)), [requests, fh.applied.from, fh.applied.to]);

  const flash = (t) => { setMsg(t); setTimeout(() => setMsg(""), 4000); };

  const remind = async (r) => {
    await onPatch(r, { rem: [...(r.rem || []), { at: new Date().toISOString(), by: me.name }] });
    flash("Satınalmaya hatırlatma bırakıldı");
  };
  const cancel = async (r) => { await onPatch(r, { status: "iptal", cancelBy: me.name, cancelAt: new Date().toISOString() }); flash("Talep iptal edildi"); };

  const kalan = (r) => Math.max(0, (Number(r.qty) || 0) - (Number(r.recvQty) || 0));

  const Row = (r) => {
    const late = r.status !== "talep" && r.eta && r.eta < today();
    return (
      <Card key={r.id} className={`p-4 ${late ? "border-red-300 bg-red-50" : ""}`}>
        <div className="flex flex-wrap items-start justify-between gap-2">
          <div className="min-w-0">
            <div className="font-medium text-stone-900">{r.name}</div>
            <div className="mt-0.5 text-sm text-stone-600">
              {fmt(r.qty)} {r.unit} · talep {dateTR(r.reqDate)} · {r.by}
              {r.urgency && r.urgency !== URGENCY[0] && <span className="ml-1 font-medium text-red-700">{r.urgency}</span>}
            </div>
            {r.note && <div className="mt-1 text-sm text-stone-600">Not: {r.note}</div>}
            {r.status !== "talep" && (
              <div className="mt-1 text-sm text-stone-600">
                {r.supplier ? `Tedarikçi: ${r.supplier}` : "Tedarikçi girilmedi"}
                {r.eta && <> · termin <span className={late ? "font-medium text-red-700" : ""}>{dateTR(r.eta)}</span></>}
                {Number(r.recvQty) > 0 && <> · gelen {fmt(r.recvQty)} / kalan {fmt(kalan(r))}</>}
              </div>
            )}
            {r.rem && r.rem.length > 0 && (
              <div className="mt-1 text-xs font-medium text-amber-800">{r.rem.length} hatırlatma · son: {new Date(r.rem[r.rem.length - 1].at).toLocaleString("tr-TR")}</div>
            )}
          </div>
          <ReqBadge status={r.status} />
        </div>
        <div className="mt-3 flex flex-wrap items-center gap-2">
          {canBuy && r.status === "talep" && <Btn onClick={() => setOrder({ r, supplier: r.supplier || "", eta: addDays(today(), 2), note: "" })}>Sipariş ver</Btn>}
          {canBuy && r.status !== "talep" && <Btn variant="ghost" onClick={() => setOrder({ r, supplier: r.supplier || "", eta: r.eta || addDays(today(), 1), note: r.orderNote || "" })}>Sipariş bilgisini güncelle</Btn>}
          {canRequest && r.status !== "talep" && <Btn onClick={() => setRecv({ r, qty: String(kalan(r)), date: today(), note: "" })}>Teslim al</Btn>}
          {canRequest && r.status !== "talep" && <Btn variant="ghost" onClick={() => remind(r)}>Hatırlat</Btn>}
          {(canBuy || canRequest) && <Btn variant="ghost" onClick={() => cancel(r)}>İptal</Btn>}
        </div>
      </Card>
    );
  };

  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 gap-3 lg:grid-cols-4">
        <Stat label="Bekleyen talep" value={fmt(bekleyenTalep.length)} sub="sipariş verilmedi" tone={bekleyenTalep.length ? "orange" : "stone"} />
        <Stat label="Yolda" value={fmt(open.filter((r) => r.status !== "talep").length)} sub="sipariş verildi" tone="sky" />
        <Stat label="Geciken" value={fmt(geciken.length)} sub="termini geçti" tone={geciken.length ? "red" : "emerald"} />
        <Stat label="Bugün gelen" value={fmt(bugunGelen.length)} sub="teslim alınan kalem" tone="emerald" />
      </div>

      <Tabs value={tab} onChange={setTab} items={[["gun", "Günlük kontrol"], ["acik", "Açık talepler"], ["gecmis", "Geçmiş"]]} />

      {canRequest && (
        <div className="flex flex-wrap items-center gap-2">
          <Btn onClick={() => setNewReq(true)}>Yeni talep aç</Btn>
          {msg && <span className="text-sm text-emerald-700">{msg}</span>}
        </div>
      )}
      {!canRequest && msg && <p className="text-sm text-emerald-700">{msg}</p>}

      {tab === "gun" && (
        <div className="space-y-4">
          <FilterBar f={fg} onApply={(d) => ensureReq(daysAgo(90), d.to)} label="Kontrol tarihi" />
          <h3 className="text-sm font-medium text-stone-900">Geciken siparişler ({geciken.length})</h3>
          {!geciken.length ? <Empty title="Geciken sipariş yok" /> : geciken.map(Row)}
          <h3 className="pt-2 text-sm font-medium text-stone-900">{dateTR(G)} için beklenenler ({bugunBeklenen.length})</h3>
          {!bugunBeklenen.length ? <Empty title="Bu tarihe termin verilmiş sipariş yok" /> : bugunBeklenen.map(Row)}
          <h3 className="pt-2 text-sm font-medium text-stone-900">Sipariş bekleyen talepler ({bekleyenTalep.length})</h3>
          {!bekleyenTalep.length ? <Empty title="Sipariş bekleyen talep yok" /> : bekleyenTalep.map(Row)}
          <h3 className="pt-2 text-sm font-medium text-stone-900">{dateTR(G)} teslim alınanlar ({bugunGelen.length})</h3>
          {!bugunGelen.length ? <Empty title="Bu tarihte teslim alınan kalem yok" /> : (
            <Table head={["Malzeme", "Gelen", "Birim", "Tedarikçi", "Teslim alan"]}
              rows={bugunGelen.map((r) => [r.name, fmt(r.recvQty), r.unit, r.supplier || "—", r.recvBy || "—"])} />
          )}
        </div>
      )}

      {tab === "acik" && (!open.length ? <Empty title="Açık talep yok" hint="Depo stok ekranından yeni talep açabilirsiniz." /> : open.map(Row))}

      {tab === "gecmis" && (
        <div className="space-y-4">
          <FilterBar f={fh} onApply={(d) => ensureReq(d.from, d.to)} label="Tarih" />
          <div className="flex justify-end">
            <Btn variant="ghost" disabled={!gecmis.length} onClick={() => exportXlsx(`satinalma-${fh.applied.from}_${fh.applied.to}.xlsx`, [{
              name: "Satınalma", head: ["Talep tarihi", "Malzeme", "İstenen", "Gelen", "Birim", "Tedarikçi", "Termin", "Teslim tarihi", "Durum", "Talep eden", "Not"],
              rows: gecmis.map((r) => [r.reqDate, r.name, r.qty, r.recvQty || 0, r.unit, r.supplier || "", r.eta || "", r.recvDate || "", REQ_STATUS[r.status]?.label || r.status, r.by || "", r.note || ""]),
            }])}>Excel indir</Btn>
          </div>
          <Table max={300} head={["Talep", "Malzeme", "İstenen", "Gelen", "Tedarikçi", "Termin", "Teslim", "Durum"]}
            rows={gecmis.map((r) => [dateTR(r.reqDate), r.name, `${fmt(r.qty)} ${r.unit}`, r.recvQty ? fmt(r.recvQty) : "–",
              r.supplier || "—", r.eta ? dateTR(r.eta) : "—", r.recvDate ? dateTR(r.recvDate) : "—", REQ_STATUS[r.status]?.label || r.status])} />
        </div>
      )}

      {newReq && <NewRequestModal cfg={cfg} levels={levels} preset={null} onClose={() => setNewReq(false)}
        onSave={async (rows) => { await onAdd(rows); flash(`${rows.length} kalem satınalmaya iletildi`); }} />}

      {order && (
        <Modal title="Sipariş bilgisi" onClose={() => setOrder(null)}
          footer={<><Btn variant="ghost" onClick={() => setOrder(null)}>Vazgeç</Btn>
            <Btn disabled={!order.eta} onClick={async () => {
              await onPatch(order.r, { status: order.r.status === "talep" ? "siparis" : order.r.status, supplier: order.supplier.trim(),
                eta: order.eta, orderNote: order.note.trim(), orderedBy: me.name, orderedAt: new Date().toISOString() });
              setOrder(null); flash("Sipariş bilgisi kaydedildi");
            }}>Kaydet</Btn></>}>
          <p className="mb-3 text-sm text-stone-600">{order.r.name} · {fmt(order.r.qty)} {order.r.unit}{order.r.note ? ` · ${order.r.note}` : ""}</p>
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="Tedarikçi"><input className={inputCls} value={order.supplier} onChange={(e) => setOrder((s) => ({ ...s, supplier: e.target.value }))} placeholder="örn. Anadolu Gıda" /></Field>
            <Field label="Termin tarihi" hint="Günlük kontrol ekranı bu tarihe bakar"><input type="date" className={inputCls} value={order.eta} onChange={(e) => setOrder((s) => ({ ...s, eta: e.target.value }))} /></Field>
          </div>
          <div className="mt-3"><Field label="Not (isteğe bağlı)"><input className={inputCls} value={order.note} onChange={(e) => setOrder((s) => ({ ...s, note: e.target.value }))} placeholder="sipariş no, fiyat, açıklama" /></Field></div>
        </Modal>
      )}

      {recv && (
        <Modal title="Teslim al" onClose={() => setRecv(null)}
          footer={<><Btn variant="ghost" onClick={() => setRecv(null)}>Vazgeç</Btn>
            <Btn disabled={!(Number(recv.qty) > 0)} onClick={async () => {
              await onReceive(recv.r, { qty: Number(recv.qty), date: recv.date, note: recv.note.trim() });
              setRecv(null); flash("Teslim alındı, stoğa işlendi");
            }}>Teslim al ve stoğa ekle</Btn></>}>
          <p className="mb-3 text-sm text-stone-600">
            {recv.r.name} · istenen {fmt(recv.r.qty)} {recv.r.unit}
            {Number(recv.r.recvQty) > 0 && <> · daha önce gelen {fmt(recv.r.recvQty)}</>}
          </p>
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label={`Gelen miktar (${recv.r.unit})`}><input type="number" min="0" step="any" inputMode="decimal" className={inputCls}
              value={recv.qty} onChange={(e) => setRecv((s) => ({ ...s, qty: e.target.value }))} /></Field>
            <Field label="Teslim tarihi"><input type="date" className={inputCls} value={recv.date} onChange={(e) => setRecv((s) => ({ ...s, date: e.target.value }))} /></Field>
          </div>
          <div className="mt-3"><Field label="Not (isteğe bağlı)"><input className={inputCls} value={recv.note} onChange={(e) => setRecv((s) => ({ ...s, note: e.target.value }))} placeholder="irsaliye no, eksik/fazla açıklaması" /></Field></div>
          <p className="mt-3 text-xs text-stone-500">Kaydedince bu miktar depo stoğuna eklenir. İstenenden az girerseniz talep açık kalır.</p>
        </Modal>
      )}
    </div>
  );
}

/* ============================ raporlar ============================ */

function Reports({ me, cfg, orders, production, fires, ensure }) {
  const allDepts = cfg.depts.map((d) => d.id);
  const myDepts = me.depts && me.depts.length ? me.depts : allDepts;
  const f = useFilters({ from: monthStart(), to: today(), dept: myDepts.length === 1 ? myDepts[0] : "all", branch: "all", kind: "all", q: "" });
  const [tab, setTab] = useState("ozet");
  const A = f.applied;
  const activeDepts = A.dept === "all" ? myDepts : [A.dept];
  const term = lower(A.q.trim());
  const dset = new Set(activeDepts);
  useEffect(() => { ensure(A.from, A.to); }, [A.from, A.to, ensure]);
  const allProd = useMemo(() => production.concat(histData().prod), [production]);
  const allFire = useMemo(() => fires.concat(histData().fire), [fires]);

  const lines = useMemo(() => {
    const out = [];
    orders.forEach((o) => {
      if (!inRange(o.prodDate, A.from, A.to)) return;
      if (A.branch !== "all" && o.branchId !== A.branch) return;
      o.lines.forEach((l) => { if (dset.has(l.deptId) && (!term || lower(l.name).includes(term))) out.push({ o, l }); });
    });
    return out;
  }, [orders, A.from, A.to, A.branch, activeDepts.join(), term]);

  const prodRows = useMemo(() => allProd.filter((r) => inRange(r.date, A.from, A.to) && dset.has(r.deptId) && (A.kind === "all" || r.kind === A.kind) && (!term || lower(r.name).includes(term))), [allProd, A.from, A.to, activeDepts.join(), A.kind, term]);
  const fireRows = useMemo(() => allFire.filter((r) => inRange(r.date, A.from, A.to) && dset.has(r.deptId) && (A.kind === "all" || r.kind === A.kind) && (!term || lower(r.name).includes(term))), [allFire, A.from, A.to, activeDepts.join(), A.kind, term]);

  const lp = (l) => Number(l.price) || catFind(cfg, l.productId)?.p || 0;
  const fp = (r) => catFind(cfg, r.itemId)?.p || 0;

  const t = lines.reduce((a, { l }) => {
    a.qty += l.qty; a.prep += prep(l); a.sent += sentQ(l); a.open += openQty(l);
    a.tl += l.qty * lp(l); a.sentTl += sentQ(l) * lp(l); a.openTl += openQty(l) * lp(l);
    if (lineStatus(l) !== "gonderildi") a.openLines += 1;
    return a;
  }, { qty: 0, prep: 0, sent: 0, open: 0, openLines: 0, tl: 0, sentTl: 0, openTl: 0 });
  const rate = t.qty ? Math.round((t.sent / t.qty) * 100) : 0;
  const pU = sumUnits(prodRows), fU = sumUnits(fireRows);
  const fireTl = fireRows.reduce((a, r) => a + r.qty * fp(r), 0);

  const byDept = cfg.depts.filter((d) => dset.has(d.id)).map((d) => {
    const ls = lines.filter(({ l }) => l.deptId === d.id);
    const qty = ls.reduce((a, { l }) => a + l.qty, 0), sent = ls.reduce((a, { l }) => a + sentQ(l), 0);
    const tl = ls.reduce((a, { l }) => a + l.qty * lp(l), 0), openTl = ls.reduce((a, { l }) => a + openQty(l) * lp(l), 0);
    const p = sumUnits(prodRows.filter((r) => r.deptId === d.id)), fr = sumUnits(fireRows.filter((r) => r.deptId === d.id));
    const frTl = fireRows.filter((r) => r.deptId === d.id).reduce((a, r) => a + r.qty * fp(r), 0);
    return { ...d, qty, sent, open: qty - sent, tl, openTl, p, f: fr, frTl, rate: qty ? Math.round((sent / qty) * 100) : 0 };
  }).filter((d) => d.qty || d.p.adet || d.p.kg || d.f.adet || d.f.kg);

  const prodByItem = useMemo(() => {
    const m = new Map();
    const touch = (r) => { let x = m.get(r.itemId); if (!x) { x = { id: r.itemId, name: r.name, deptId: r.deptId, unit: r.unit, kind: r.kind, qty: 0, fire: 0, gun: new Set() }; m.set(r.itemId, x); } return x; };
    prodRows.forEach((r) => { const x = touch(r); x.qty += r.qty; x.gun.add(r.date); });
    fireRows.forEach((r) => { const x = touch(r); x.fire += r.qty; });
    return [...m.values()].map((x) => ({ ...x, gun: x.gun.size })).sort((a, b) => b.qty - a.qty);
  }, [prodRows, fireRows]);

  const missing = lines.filter(({ l }) => lineStatus(l) !== "gonderildi" && openQty(l) > 0).map(({ o, l }) => ({
    date: o.prodDate, branch: branchName(cfg, o.branchId), dept: deptName(cfg, l.deptId), name: l.name,
    qty: l.qty, prep: prep(l), sent: sentQ(l), open: openQty(l), tl: openQty(l) * lp(l),
    reason: l.shortReason || "—", note: l.shortNote || "", status: STATUS[lineStatus(l)].label,
  })).sort((a, b) => (a.date === b.date ? b.open - a.open : a.date < b.date ? 1 : -1));

  const byReason = useMemo(() => {
    const m = new Map();
    missing.forEach((r) => { let x = m.get(r.reason); if (!x) { x = { reason: r.reason, kalem: 0, adet: 0, tl: 0 }; m.set(r.reason, x); } x.kalem++; x.adet += r.open; x.tl += r.tl; });
    return [...m.values()].sort((a, b) => b.adet - a.adet);
  }, [missing]);

  const daily = useMemo(() => {
    const m = new Map();
    const touch = (d) => { let x = m.get(d); if (!x) { x = { d, talep: 0, sevk: 0, uretim: 0, uretimKg: 0, fire: 0, fireKg: 0, tl: 0 }; m.set(d, x); } return x; };
    lines.forEach(({ o, l }) => { const x = touch(o.prodDate); x.talep += l.qty; x.sevk += sentQ(l); x.tl += sentQ(l) * lp(l); });
    prodRows.forEach((r) => { const x = touch(r.date); if (r.unit === "Adet") x.uretim += r.qty; else x.uretimKg += r.unit === "Gr" ? r.qty / 1000 : r.qty; });
    fireRows.forEach((r) => { const x = touch(r.date); if (r.unit === "Adet") x.fire += r.qty; else x.fireKg += r.unit === "Gr" ? r.qty / 1000 : r.qty; });
    return [...m.values()].sort((a, b) => (a.d < b.d ? 1 : -1));
  }, [lines, prodRows, fireRows]);

  const byBranch = cfg.branches.map((b) => {
    const ls = lines.filter(({ o }) => o.branchId === b.id);
    const qty = ls.reduce((a, { l }) => a + l.qty, 0), sent = ls.reduce((a, { l }) => a + sentQ(l), 0);
    const tl = ls.reduce((a, { l }) => a + l.qty * lp(l), 0), sentTl = ls.reduce((a, { l }) => a + sentQ(l) * lp(l), 0);
    return { ...b, qty, sent, open: qty - sent, tl, sentTl, rate: qty ? Math.round((sent / qty) * 100) : 0 };
  }).filter((b) => b.qty);

  const sheets = () => [
    { name: "Bölüm özeti", head: ["Bölüm", "Talep", "Talep ₺", "Gönderilen", "Açık", "Eksik ₺", "Karşılama %", "Üretim adet", "Üretim kg", "Fire adet", "Fire kg", "Fire ₺"],
      rows: byDept.map((d) => [d.name, d.qty, Math.round(d.tl), d.sent, d.open, Math.round(d.openTl), d.rate, Math.round(d.p.adet), +d.p.kg.toFixed(2), Math.round(d.f.adet), +d.f.kg.toFixed(2), Math.round(d.frTl)]) },
    { name: "Üretim dökümü", head: ["Ürün", "Bölüm", "Tür", "Birim", "Üretim", "Fire", "Fire ₺", "Kayıt günü"],
      rows: prodByItem.map((p) => [p.name, deptName(cfg, p.deptId), p.kind === "mamul" ? "Mamul" : "Yarı mamul", p.unit, p.qty, p.fire, Math.round(p.fire * (catFind(cfg, p.id)?.p || 0)), p.gun]) },
    { name: "Fire kayıtları", head: ["Tarih", "Bölüm", "Ürün", "Tür", "Miktar", "Birim", "Tutar ₺", "Sebep", "Kaydeden"],
      rows: fireRows.slice().sort((a, b) => (a.date < b.date ? 1 : -1)).map((r) => [r.date, deptName(cfg, r.deptId), r.name, r.kind === "mamul" ? "Mamul" : "Yarı mamul", r.qty, r.unit, Math.round(r.qty * fp(r)), r.reason || "", r.by]) },
    { name: "Eksik gönderim", head: ["Üretim tarihi", "Şube", "Bölüm", "Ürün", "İstenen", "Hazırlanan", "Gönderilen", "Eksik", "Eksik ₺", "Sebep", "Açıklama", "Durum"],
      rows: missing.map((r) => [r.date, r.branch, r.dept, r.name, r.qty, r.prep, r.sent, r.open, Math.round(r.tl), r.reason, r.note, r.status]) },
    { name: "Eksik sebepleri", head: ["Sebep", "Kalem", "Eksik adet", "Tutar ₺"], rows: byReason.map((r) => [r.reason, r.kalem, r.adet, Math.round(r.tl)]) },
    { name: "Şube talepleri", head: ["Şube", "Talep", "Talep ₺", "Gönderilen", "Gönderilen ₺", "Açık", "Karşılama %"],
      rows: byBranch.map((b) => [b.name, b.qty, Math.round(b.tl), b.sent, Math.round(b.sentTl), b.open, b.rate]) },
    { name: "Günlük seyir", head: ["Tarih", "Talep", "Sevk", "Sevk ₺", "Üretim adet", "Üretim kg", "Fire adet", "Fire kg"],
      rows: daily.map((d) => [d.d, d.talep, d.sevk, Math.round(d.tl), Math.round(d.uretim), +d.uretimKg.toFixed(2), Math.round(d.fire), +d.fireKg.toFixed(2)]) },
  ];
  const exportOne = () => {
    const map = { ozet: 0, uretim: 1, fire: 2, eksik: 3, sebep: 4, sube: 5, gunluk: 6 };
    const s = sheets()[map[tab] ?? 0];
    exportXlsx(`${s.name.toLowerCase().replace(/ /g, "-")}-${A.from}_${A.to}.xlsx`, [s]);
  };

  return (
    <div className="space-y-4">
      <FilterBar f={f} onApply={(d) => ensure(d.from, d.to)} label="Üretim tarihi">
        <DeptSelect f={f} cfg={cfg} allowed={myDepts} />
        <Field label="Ürün türü"><select className={inputCls} value={f.draft.kind} onChange={(e) => f.set("kind", e.target.value)}>
          <option value="all">Hepsi</option><option value="mamul">Yalnızca mamul</option><option value="yarimamul">Yalnızca yarı mamul</option><option value="hammadde">Yalnızca hammadde</option></select></Field>
        <BranchSelect f={f} cfg={cfg} />
        <Field label="Ürün ara"><input className={inputCls} value={f.draft.q} onChange={(e) => f.set("q", e.target.value)} placeholder="örn. magnolia" /></Field>
      </FilterBar>

      <div className="flex flex-wrap gap-2">
        <Btn onClick={() => exportXlsx(`rapor-${A.from}_${A.to}.xlsx`, sheets())}>Tüm raporları Excel'e aktar</Btn>
        <Btn variant="ghost" onClick={exportOne}>Bu sekmeyi indir</Btn>
      </div>

      <div className="grid grid-cols-2 gap-3 lg:grid-cols-6">
        <Stat label="Talep" value={fmt(t.qty)} sub={money(t.tl)} />
        <Stat label="Gönderilen" value={fmt(t.sent)} sub={money(t.sentTl)} tone="emerald" />
        <Stat label="Eksik kalan" value={fmt(t.open)} sub={money(t.openTl)} tone="orange" />
        <Stat label="Üretim" value={fmt(pU.adet)} sub={pU.kg ? `+ ${fmt(pU.kg)} kg` : "adet"} tone="sky" />
        <Stat label="Fire" value={fmt(fU.adet)} sub={fU.kg ? `+ ${fmt(fU.kg)} kg` : "adet"} tone="red" />
        <Stat label="Fire tutarı" value={money(fireTl)} sub="katalog fiyatıyla" tone="red" />
      </div>

      <Tabs value={tab} onChange={setTab} items={[["ozet", "Bölüm özeti"], ["uretim", "Üretim dökümü"], ["fire", "Fire"], ["eksik", "Eksik gönderim"], ["sebep", "Eksik sebepleri"], ["sube", "Şube"], ["gunluk", "Günlük seyir"]]} />

      {tab === "ozet" && (<>
        <Card className="p-4"><p className="text-sm text-stone-700">
          {dateTR(A.from)} – {dateTR(A.to)} · {A.dept === "all" ? "bütün bölümler" : deptName(cfg, A.dept)} · {fmt(t.qty)} adet ({money(t.tl)}) talep, {fmt(t.sent)} adet gönderim (karşılama %{rate}).
          {t.openLines > 0 && <> {t.openLines} kalemde {fmt(t.open)} adet, {money(t.openTl)} değerinde eksik kaldı.</>}
          {" "}Üretim {fmt(pU.adet)} adet{pU.kg ? ` ve ${fmt(pU.kg)} kg` : ""}, fire {fmt(fU.adet)} adet{fU.kg ? ` ve ${fmt(fU.kg)} kg` : ""}.
        </p></Card>
        <Table head={["Bölüm", "Talep", "Talep ₺", "Gönderilen", "Açık", "Eksik ₺", "Karşılama", "Üretim", "Fire", "Fire ₺"]}
          rows={byDept.map((d) => [d.name, fmt(d.qty), money(d.tl), fmt(d.sent), fmt(d.open), money(d.openTl), `%${d.rate}`, unitCell(d.p), unitCell(d.f), money(d.frTl)])}
          foot={["Toplam", fmt(t.qty), money(t.tl), fmt(t.sent), fmt(t.open), money(t.openTl), `%${rate}`, unitCell(pU), unitCell(fU), money(fireTl)]} />
      </>)}
      {tab === "uretim" && <Table max={300} head={["Ürün", "Bölüm", "Tür", "Birim", "Üretim", "Fire", "Fire ₺", "Kayıt günü"]}
        rows={prodByItem.map((p) => [p.name, deptName(cfg, p.deptId), p.kind === "mamul" ? "Mamul" : "Yarı mamul", p.unit, fmt(p.qty), fmt(p.fire), money(p.fire * (catFind(cfg, p.id)?.p || 0)), p.gun])}
        foot={["Toplam", "", "", "", unitCell(pU), unitCell(fU), money(fireTl), ""]} />}
      {tab === "fire" && <Table max={300} head={["Tarih", "Bölüm", "Ürün", "Tür", "Miktar", "Birim", "Tutar", "Sebep"]}
        rows={fireRows.slice().sort((a, b) => (a.date < b.date ? 1 : -1)).map((r) => [dateTR(r.date), deptName(cfg, r.deptId), r.name, r.kind === "mamul" ? "Mamul" : "Yarı mamul", fmt(r.qty), r.unit, money(r.qty * fp(r)), r.reason || "—"])}
        foot={["Toplam", "", "", "", unitCell(fU), "", money(fireTl), ""]} />}
      {tab === "eksik" && <Table max={300} head={["Üretim", "Şube", "Bölüm", "Ürün", "İstenen", "Gönderilen", "Eksik", "Tutar", "Sebep", "Açıklama"]}
        rows={missing.map((r) => [dateTR(r.date), r.branch, r.dept, r.name, fmt(r.qty), fmt(r.sent), fmt(r.open), money(r.tl), r.reason, r.note || "—"])}
        foot={["Toplam", "", "", "", fmt(missing.reduce((a, r) => a + r.qty, 0)), fmt(missing.reduce((a, r) => a + r.sent, 0)), fmt(missing.reduce((a, r) => a + r.open, 0)), money(missing.reduce((a, r) => a + r.tl, 0)), "", ""]} />}
      {tab === "sebep" && <Table head={["Sebep", "Kalem", "Eksik adet", "Tutar"]} rows={byReason.map((r) => [r.reason, fmt(r.kalem), fmt(r.adet), money(r.tl)])}
        foot={["Toplam", fmt(byReason.reduce((a, r) => a + r.kalem, 0)), fmt(byReason.reduce((a, r) => a + r.adet, 0)), money(byReason.reduce((a, r) => a + r.tl, 0))]} />}
      {tab === "sube" && <Table head={["Şube", "Talep", "Talep ₺", "Gönderilen", "Gönderilen ₺", "Açık", "Karşılama"]}
        rows={byBranch.map((b) => [b.name, fmt(b.qty), money(b.tl), fmt(b.sent), money(b.sentTl), fmt(b.open), `%${b.rate}`])}
        foot={["Toplam", fmt(t.qty), money(t.tl), fmt(t.sent), money(t.sentTl), fmt(t.open), `%${rate}`]} />}
      {tab === "gunluk" && <Table max={200} head={["Tarih", "Talep", "Sevk", "Sevk ₺", "Üretim (adet)", "Üretim (kg)", "Fire (adet)", "Fire (kg)"]}
        rows={daily.map((d) => [dateTR(d.d), fmt(d.talep), fmt(d.sevk), money(d.tl), fmt(d.uretim), d.uretimKg ? fmt(d.uretimKg) : "–", fmt(d.fire), d.fireKg ? fmt(d.fireKg) : "–"])}
        foot={["Toplam", fmt(t.qty), fmt(t.sent), money(t.sentTl), fmt(pU.adet), fmt(pU.kg), fmt(fU.adet), fmt(fU.kg)]} />}
    </div>
  );
}

/* ============================ yönetim ============================ */

function Admin({ cfg, setCfg, orders, production, fires }) {
  const [tab, setTab] = useState("catalog");
  return (
    <div className="space-y-4">
      <Tabs value={tab} onChange={setTab} items={[["catalog", "Ürünler"], ["users", "Kullanıcılar"], ["settings", "Ayarlar"], ["depts", "Bölümler"], ["branches", "Şubeler"]]} />
      {tab === "catalog" && <AdminCatalog cfg={cfg} setCfg={setCfg} />}
      {tab === "users" && <AdminUsers cfg={cfg} setCfg={setCfg} />}
      {tab === "settings" && <AdminSettings cfg={cfg} setCfg={setCfg} orders={orders} production={production} fires={fires} />}
      {tab === "depts" && <AdminSimple cfg={cfg} setCfg={setCfg} field="depts" label="Bölüm" />}
      {tab === "branches" && <AdminSimple cfg={cfg} setCfg={setCfg} field="branches" label="Şube" />}
    </div>
  );
}

function AdminSettings({ cfg, setCfg, orders, production, fires }) {
  const s = cfg.settings;
  const set = (k, v) => setCfg((c) => ({ ...c, settings: { ...c.settings, [k]: Number(v) } }));
  const backup = () => {
    const blob = new Blob([JSON.stringify({ cfg: { ...cfg, _cat: undefined, _dept: undefined, _branch: undefined }, orders, production: production.filter((r) => !r.seed), fires: fires.filter((r) => !r.seed) }, null, 1)], { type: "application/json" });
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `yedek-${today()}.json`; a.click();
  };
  return (
    <div className="space-y-4">
      <Card className="p-4">
        <h3 className="mb-3 font-medium text-stone-900">Gün akışı</h3>
        <div className="grid gap-3 sm:grid-cols-2">
          <Field label="Talepten üretime kaç gün" hint="1 = bugünkü talep yarın üretilir">
            <input type="number" min="0" max="7" className={inputCls} value={s.leadProd} onChange={(e) => set("leadProd", e.target.value)} /></Field>
          <Field label="Üretimden sevke kaç gün" hint="1 = yarın üretilen, ertesi gün sevk edilir">
            <input type="number" min="0" max="7" className={inputCls} value={s.leadShip} onChange={(e) => set("leadShip", e.target.value)} /></Field>
        </div>
      </Card>
      <Card className="p-4">
        <h3 className="mb-3 font-medium text-stone-900">Talep penceresi</h3>
        <div className="grid gap-3 sm:grid-cols-2">
          <Field label="Açılış saati" hint="Bu saatten önce talep girilemez ve düzeltilemez">
            <input type="number" min="0" max="23" className={inputCls} value={s.openHour} onChange={(e) => set("openHour", e.target.value)} /></Field>
          <Field label="Kapanış saati" hint="24 = gece yarısı">
            <input type="number" min="1" max="24" className={inputCls} value={s.closeHour} onChange={(e) => set("closeHour", e.target.value)} /></Field>
        </div>
        <p className="mt-3 text-sm text-stone-600">Şu an pencere {windowOpen(s) ? "açık" : "kapalı"}.</p>
      </Card>
      <Card className="p-4">
        <h3 className="mb-2 font-medium text-stone-900">Veri saklama</h3>
        <p className="text-sm text-stone-700">
          Şu anki kayıt yeri:{" "}
          <b>{storeMode() === "supabase" ? "Ortak veritabanı (Supabase)" : storeMode() === "local" ? "Bu tarayıcı" : "Yalnızca bellek"}</b>.
        </p>
        <p className="mt-1 text-sm text-stone-600">
          {storeMode() === "supabase"
            ? "Girilen veriler bütün cihazlarda ortak görünür, değişiklikler anında yansır."
            : storeMode() === "local"
            ? "Veriler yalnızca bu bilgisayarda/tarayıcıda saklanır. Başka bir cihazdan girilenler burada görünmez; çok kullanıcılı çalışma için sunucu gerekir."
            : "Kalıcı depolama açılamadı, veriler sayfa yenilenince silinir."}
        </p>
      </Card>
      <Card className="p-4">
        <h3 className="mb-2 font-medium text-stone-900">Yedek</h3>
        <p className="mb-3 text-sm text-stone-600">Katalog, kullanıcılar, talepler, üretim ve fire kayıtlarını tek dosyada indirir.</p>
        <Btn variant="ghost" onClick={backup}>Yedeği indir</Btn>
      </Card>
    </div>
  );
}

function RecipeModal({ cfg, item, onClose }) {
  const lines = RECIPES[item.id] || [];
  return (
    <Modal wide title={`Reçete · ${item.n}`} onClose={onClose}>
      <p className="mb-3 text-sm text-stone-600">
        Bir parti için gereken bileşenler{item.pb ? ` · parti çıktısı ${fmt(item.pb)} ${item.u}` : ""}.
      </p>
      {!lines.length ? <Empty title="Bu kalem için reçete tanımlı değil" /> : (
        <table className="w-full text-sm">
          <thead className="bg-stone-50 text-left text-xs text-stone-600">
            <tr><th className="px-2 py-2 font-medium">Bileşen</th><th className="px-2 py-2 font-medium">Kod</th>
              <th className="px-2 py-2 text-right font-medium">Miktar</th><th className="px-2 py-2 text-right font-medium">Birim</th></tr>
          </thead>
          <tbody className="divide-y divide-stone-100">
            {lines.map(([cid, q], i) => {
              const b = catFind(cfg, cid);
              return (<tr key={i}><td className="px-2 py-2 text-stone-800">{b ? b.n : cid}</td>
                <td className="px-2 py-2 text-xs text-stone-500">{b ? b.c || b.id : cid}</td>
                <td className="px-2 py-2 text-right tabular-nums">{fmt(q)}</td>
                <td className="px-2 py-2 text-right text-xs text-stone-500">{b ? b.u : ""}</td></tr>);
            })}
          </tbody>
        </table>
      )}
    </Modal>
  );
}

function AdminCatalog({ cfg, setCfg }) {
  const [q, setQ] = useState(""); const [dept, setDept] = useState("all"); const [vis, setVis] = useState("all");
  const [src, setSrc] = useState("all"); const [kind, setKind] = useState("all");
  const [adding, setAdding] = useState(false);
  const [recipe, setRecipe] = useState(null);
  const [opt, setOpt] = useState(null);
  const [ni, setNi] = useState({ n: "", d: cfg.depts[0]?.id, u: "Adet", k: "mamul", p: "", g: "Elle eklendi", vt: true, vu: false, vd: false, vi: false });

  const list = useMemo(() => {
    const t = lower(q.trim());
    return cfg.catalog
      .filter((c) => (dept === "all" ? true : c.d === dept))
      .filter((c) => (src === "all" ? true : (c.src || "liste") === src))
      .filter((c) => (kind === "all" ? true : (c.k || "mamul") === kind))
      .filter((c) => vis === "all" ? true : vis === "talep" ? c.vt : vis === "uretim" ? c.vu : vis === "depo" ? c.vd : vis === "bolum" ? c.vi : !c.vt && !c.vu && !c.vd && !c.vi)
      .filter((c) => (t ? lower(c.n).includes(t) || lower(String(c.c || c.id)).includes(t) : true));
  }, [cfg.catalog, q, dept, vis, src, kind]);

  const patch = (id, o) => setCfg((s) => ({ ...s, catOv: { ...(s.catOv || {}), [id]: { ...((s.catOv || {})[id] || {}), ...o } } }));
  const bulk = (field, val) => setCfg((s) => {
    const ov = { ...(s.catOv || {}) };
    list.forEach((c) => { ov[c.id] = { ...(ov[c.id] || {}), [field]: val }; });
    return { ...s, catOv: ov };
  });
  const remove = (id) => setCfg((s) => ({ ...s, catDel: [...(s.catDel || []), id], catAdd: (s.catAdd || []).filter((x) => x.id !== id) }));

  const byDeptTalep = useMemo(() => {
    const m = new Map();
    cfg.catalog.forEach((c) => { if (c.vt) { if (!m.has(c.d)) m.set(c.d, []); m.get(c.d).push(c); } });
    return m;
  }, [cfg.catalog]);
  const talepOfDept = (d) => byDeptTalep.get(d) || [];

  const counts = {
    talep: cfg.catalog.filter((c) => c.vt).length,
    uretim: cfg.catalog.filter((c) => c.vu).length,
    depo: cfg.catalog.filter((c) => c.vd).length,
    bolum: cfg.catalog.filter((c) => c.vi).length,
    pasif: cfg.catalog.filter((c) => !c.vt && !c.vu && !c.vd && !c.vi).length,
  };

  return (
    <div className="space-y-4">
      <Card className="p-4">
        <p className="mb-3 text-sm text-stone-600">
          Bir kalemin hangi ekrana düşeceğini buradan belirlersiniz. <b>Şube</b>: şubelerin talep ekranında çıkar ve
          talep, kalemin <b>bölümüne</b> düşer — bölümü Depo ise depocuya, Pasta ise pasta şefine gider.
          <b> Üretim</b>: ustaların üretim/fire ekranında çıkar. <b>Depo</b>: depo stok listesinde yer alır.
          <b> Bölüm</b>: üretim bölümlerinin depodan isteyebileceği listede çıkar. Bir kalem birden fazla yerde işaretlenebilir.
          Satır sonundaki <b>seçenekler</b> bağlantısından not/görsel iznini ve hangi bölümlerin isteyebileceğini ayarlayın.
        </p>
        <div className="mb-3 flex flex-wrap gap-3 text-sm text-stone-600">
          <span>Toplam <b className="text-stone-900">{cfg.catalog.length}</b></span>
          <span>Şubede <b className="text-emerald-800">{counts.talep}</b></span>
          <span>Üretimde <b className="text-sky-800">{counts.uretim}</b></span>
          <span>Depoda <b className="text-amber-800">{counts.depo}</b></span>
          <span>Bölüm talebinde <b className="text-purple-800">{counts.bolum}</b></span>
          <span>Pasif <b className="text-stone-500">{counts.pasif}</b></span>
        </div>
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-5">
          <Field label="Ara"><input className={inputCls} placeholder="Ürün adı veya kodu" value={q} onChange={(e) => setQ(e.target.value)} /></Field>
          <Field label="Bölüm"><select className={inputCls} value={dept} onChange={(e) => setDept(e.target.value)}>
            <option value="all">Bütün bölümler</option>{cfg.depts.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}</select></Field>
          <Field label="Liste"><select className={inputCls} value={vis} onChange={(e) => setVis(e.target.value)}>
            <option value="all">Hepsi</option><option value="talep">Şube talep listesi</option><option value="uretim">Üretim listesi</option>
            <option value="depo">Depo stok listesi</option><option value="bolum">Bölüm talep listesi</option><option value="hicbiri">Pasif</option></select></Field>
          <Field label="Tür"><select className={inputCls} value={kind} onChange={(e) => setKind(e.target.value)}>
            <option value="all">Hepsi</option><option value="mamul">Mamul</option><option value="yarimamul">Yarı mamul</option><option value="hammadde">Hammadde</option></select></Field>
          <Field label="Kaynak"><select className={inputCls} value={src} onChange={(e) => setSrc(e.target.value)}>
            <option value="all">Hepsi</option><option value="liste">Fiyat + üretim listesi</option><option value="stok">Stok/reçete dosyası</option><option value="elle">Elle eklenen</option></select></Field>
        </div>
        <div className="mt-3 flex flex-wrap items-center gap-2">
          <Btn onClick={() => setAdding(!adding)}>Ürün ekle</Btn>
          <span className="text-xs text-stone-500">Listedeki {list.length} kaleme topluca:</span>
          <Btn variant="ghost" onClick={() => bulk("vt", 1)}>Şubeye aç</Btn><Btn variant="ghost" onClick={() => bulk("vt", 0)}>Şubeden kaldır</Btn>
          <Btn variant="ghost" onClick={() => bulk("vu", 1)}>Üretime aç</Btn><Btn variant="ghost" onClick={() => bulk("vu", 0)}>Üretimden kaldır</Btn>
          <Btn variant="ghost" onClick={() => bulk("vd", 1)}>Depoya aç</Btn><Btn variant="ghost" onClick={() => bulk("vd", 0)}>Depodan kaldır</Btn>
          <Btn variant="ghost" onClick={() => bulk("vi", 1)}>Bölüm talebine aç</Btn><Btn variant="ghost" onClick={() => bulk("vi", 0)}>Bölüm talebinden kaldır</Btn>
        </div>
        {adding && (
          <div className="mt-3 grid gap-3 border-t border-stone-100 pt-3 sm:grid-cols-3 lg:grid-cols-6">
            <Field label="Ürün adı"><input className={inputCls} value={ni.n} onChange={(e) => setNi({ ...ni, n: e.target.value })} /></Field>
            <Field label="Bölüm"><select className={inputCls} value={ni.d} onChange={(e) => setNi({ ...ni, d: e.target.value })}>
              {cfg.depts.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}</select></Field>
            <Field label="Birim"><select className={inputCls} value={ni.u} onChange={(e) => setNi({ ...ni, u: e.target.value })}>{["Adet", "Kg", "Gr"].map((u) => <option key={u} value={u}>{u}</option>)}</select></Field>
            <Field label="Tür"><select className={inputCls} value={ni.k} onChange={(e) => setNi({ ...ni, k: e.target.value })}>
              <option value="mamul">Mamul</option><option value="yarimamul">Yarı mamul</option><option value="hammadde">Hammadde</option></select></Field>
            <Field label="Fiyat (₺)"><input type="number" className={inputCls} value={ni.p} onChange={(e) => setNi({ ...ni, p: e.target.value })} placeholder="0" /></Field>
            <div className="flex items-end gap-3 text-sm">
              <label className="flex items-center gap-1"><input type="checkbox" className="h-4 w-4" checked={ni.vt} onChange={(e) => setNi({ ...ni, vt: e.target.checked })} />Şube</label>
              <label className="flex items-center gap-1"><input type="checkbox" className="h-4 w-4" checked={ni.vu} onChange={(e) => setNi({ ...ni, vu: e.target.checked })} />Üretim</label>
              <label className="flex items-center gap-1"><input type="checkbox" className="h-4 w-4" checked={ni.vd} onChange={(e) => setNi({ ...ni, vd: e.target.checked })} />Depo</label>
              <label className="flex items-center gap-1"><input type="checkbox" className="h-4 w-4" checked={ni.vi} onChange={(e) => setNi({ ...ni, vi: e.target.checked })} />Bölüm</label>
            </div>
            <div className="sm:col-span-3 lg:col-span-6"><Btn onClick={() => {
              if (!ni.n.trim()) return;
              const id = "X" + uid().toUpperCase();
              const o = Math.max(0, ...cfg.catalog.filter((c) => c.d === ni.d).map((c) => c.o || 0)) + 1;
              setCfg((s) => ({ ...s, catAdd: [...(s.catAdd || []), { id, c: id, n: ni.n.trim(), g: ni.g, d: ni.d, u: ni.u, k: ni.k, p: Number(ni.p) || 0, o, vt: ni.vt ? 1 : 0, vu: ni.vu ? 1 : 0, vd: ni.vd ? 1 : 0, vi: ni.vi ? 1 : 0, an: 0, ag: 0, dd: [], src: "elle" }] }));
              setNi({ ...ni, n: "", p: "" }); setAdding(false);
            }}>Kaydet</Btn></div>
          </div>
        )}
      </Card>

      <Card className="overflow-x-auto">
        <table className="w-full min-w-max text-sm">
          <thead className="bg-stone-50 text-left text-xs text-stone-600"><tr>
            <th className="px-3 py-2 font-medium">Ürün</th><th className="px-3 py-2 font-medium">Bölüm</th><th className="px-3 py-2 font-medium">Birim</th>
            <th className="px-3 py-2 font-medium">Tür</th><th className="px-3 py-2 text-right font-medium">Fiyat</th>
            <th className="px-3 py-2 text-center font-medium">Şube</th><th className="px-3 py-2 text-center font-medium">Üretim</th>
            <th className="px-3 py-2 text-center font-medium">Depo</th><th className="px-3 py-2 text-center font-medium">Bölüm</th>
            <th className="px-3 py-2 font-medium">Talep eşi</th><th /></tr></thead>
          <tbody className="divide-y divide-stone-100">
            {list.slice(0, 500).map((c) => (
              <tr key={c.id} className={!c.vt && !c.vu && !c.vd ? "bg-stone-50 text-stone-400" : ""}>
                <td className="px-3 py-2"><div className="font-medium text-stone-800">{c.n}</div>
                  <div className="text-xs text-stone-500">{c.c || c.id} · {c.g}
                    {RECIPES[c.id] && <button className="ml-2 text-emerald-800 hover:underline" onClick={() => setRecipe(c)}>reçete</button>}</div></td>
                <td className="px-3 py-2"><select className="rounded-md border border-stone-300 px-2 py-1 text-sm" value={c.d} onChange={(e) => patch(c.id, { d: e.target.value, l: "" })}>
                  {cfg.depts.map((d) => <option key={d.id} value={d.id}>{d.name}</option>)}</select></td>
                <td className="px-3 py-2"><select className="rounded-md border border-stone-300 px-2 py-1 text-sm" value={c.u} onChange={(e) => patch(c.id, { u: e.target.value })}>
                  {["Adet", "Kg", "Gr"].map((u) => <option key={u} value={u}>{u}</option>)}</select></td>
                <td className="px-3 py-2"><select className="rounded-md border border-stone-300 px-2 py-1 text-sm" value={c.k} onChange={(e) => patch(c.id, { k: e.target.value })}>
                  <option value="mamul">Mamul</option><option value="yarimamul">Yarı mamul</option><option value="hammadde">Hammadde</option></select></td>
                <td className="px-3 py-2 text-right"><input type="number" className="w-24 rounded-md border border-stone-300 px-2 py-1 text-right text-sm tabular-nums" value={c.p || ""} placeholder="0" onChange={(e) => patch(c.id, { p: Number(e.target.value) || 0 })} /></td>
                <td className="px-3 py-2 text-center"><input type="checkbox" className="h-5 w-5 rounded border-stone-300" checked={!!c.vt} onChange={(e) => patch(c.id, { vt: e.target.checked ? 1 : 0 })} /></td>
                <td className="px-3 py-2 text-center"><input type="checkbox" className="h-5 w-5 rounded border-stone-300" checked={!!c.vu} onChange={(e) => patch(c.id, { vu: e.target.checked ? 1 : 0 })} /></td>
                <td className="px-3 py-2 text-center"><input type="checkbox" className="h-5 w-5 rounded border-stone-300" checked={!!c.vd} onChange={(e) => patch(c.id, { vd: e.target.checked ? 1 : 0 })} /></td>
                <td className="px-3 py-2 text-center"><input type="checkbox" className="h-5 w-5 rounded border-stone-300" checked={!!c.vi} onChange={(e) => patch(c.id, { vi: e.target.checked ? 1 : 0 })} /></td>
                <td className="px-3 py-2">{c.vu && !c.vt ? (
                  <select className="max-w-[16rem] rounded-md border border-stone-300 px-2 py-1 text-sm" value={c.l || ""} onChange={(e) => patch(c.id, { l: e.target.value })}>
                    <option value="">— bağlı değil —</option>
                    {talepOfDept(c.d).map((x) => <option key={x.id} value={x.id}>{x.n}</option>)}
                  </select>) : <span className="text-xs text-stone-400">—</span>}</td>
                <td className="whitespace-nowrap px-3 py-2 text-right">
                  <button className="text-xs text-emerald-800 hover:underline" onClick={() => setOpt(c)}>seçenekler</button>
                  <button className="ml-3 text-xs text-red-700 hover:underline" onClick={() => remove(c.id)}>sil</button></td>
              </tr>
            ))}
          </tbody>
        </table>
        {!list.length && <p className="p-6 text-center text-sm text-stone-500">Eşleşen ürün yok.</p>}
        {list.length > 500 && <p className="border-t border-stone-100 px-3 py-2 text-xs text-stone-500">{list.length} kalemin ilk 500'ü gösteriliyor, aramayı daraltın.</p>}
      </Card>
      {recipe && <RecipeModal cfg={cfg} item={recipe} onClose={() => setRecipe(null)} />}
      {opt && (
        <Modal title={`Seçenekler · ${opt.n}`} onClose={() => setOpt(null)}>
          <p className="mb-3 text-sm text-stone-600">Şubenin bu ürünü isterken ekleyebilecekleri ve hangi bölümlerin depodan isteyebileceği.</p>
          <div className="space-y-2">
            <label className="flex items-start gap-2 text-sm text-stone-700">
              <input type="checkbox" className="mt-0.5 h-4 w-4 rounded border-stone-300" checked={!!opt.an}
                onChange={(e) => { patch(opt.id, { an: e.target.checked ? 1 : 0 }); setOpt({ ...opt, an: e.target.checked ? 1 : 0 }); }} />
              Şube bu ürüne not yazabilsin (kaç kişilik, üzerine yazı, teslim saati)
            </label>
            <label className="flex items-start gap-2 text-sm text-stone-700">
              <input type="checkbox" className="mt-0.5 h-4 w-4 rounded border-stone-300" checked={!!opt.ag}
                onChange={(e) => { patch(opt.id, { ag: e.target.checked ? 1 : 0 }); setOpt({ ...opt, ag: e.target.checked ? 1 : 0 }); }} />
              Şube bu ürüne görsel ekleyebilsin (özel pasta siparişleri için)
            </label>
          </div>
          {!!opt.vd && (
            <div className="mt-4">
              <p className="mb-1 text-sm font-medium text-stone-700">Kritik stok seviyesi</p>
              <p className="mb-2 text-xs text-stone-500">Depo stoğu bu miktarın altına düşerse "kritik" olarak işaretlenir. 0 bırakılırsa uyarı verilmez.</p>
              <div className="flex items-center gap-2">
                <input type="number" min="0" className="w-32 rounded-md border border-stone-300 px-2 py-1 text-right text-sm tabular-nums"
                  value={opt.ms || ""} placeholder="0"
                  onChange={(e) => { const ms = Number(e.target.value) || 0; patch(opt.id, { ms }); setOpt({ ...opt, ms }); }} />
                <span className="text-sm text-stone-600">{opt.u}</span>
              </div>
            </div>
          )}
          <div className="mt-4">
            <p className="mb-1 text-sm font-medium text-stone-700">Depodan isteyebilecek bölümler</p>
            <p className="mb-2 text-xs text-stone-500">Hiçbiri seçilmezse bütün bölümler isteyebilir. Kalemin ayrıca <b>Bölüm</b> kutusu işaretli olmalı.</p>
            <div className="flex flex-wrap gap-2">
              {cfg.depts.filter((d) => d.id !== "depo").map((d) => {
                const on = (opt.dd || []).includes(d.id);
                return (
                  <button key={d.id} onClick={() => {
                    const dd = on ? (opt.dd || []).filter((x) => x !== d.id) : [...(opt.dd || []), d.id];
                    patch(opt.id, { dd }); setOpt({ ...opt, dd });
                  }} className={`rounded-full border px-3 py-1 text-sm ${on ? "border-emerald-800 bg-emerald-800 text-white" : "border-stone-300 bg-white text-stone-700"}`}>{d.name}</button>
                );
              })}
            </div>
          </div>
        </Modal>
      )}
    </div>
  );
}

function AdminUsers({ cfg, setCfg }) {
  const [edit, setEdit] = useState(null);
  const [sync, setSync] = useState("");
  const blank = () => ({ id: uid(), name: "", username: "", pin: "", role: "sef", branchId: null, depts: [], perms: ROLES.sef.perms.slice(), isNew: true });
  const save = () => {
    if (!edit.name.trim() || !edit.username.trim() || !edit.pin.trim()) return;
    const c = { ...edit }; delete c.isNew;
    if (c.role !== "sube") c.perms = c.perms.filter((p) => p !== "order.create");
    setCfg((s) => ({ ...s, users: edit.isNew ? [...s.users, c] : s.users.map((u) => (u.id === c.id ? c : u)) }));
    setEdit(null);
  };
  return (
    <div className="space-y-4">
      <Card className="p-4">
        <p className="mb-3 text-sm text-stone-600">
          Programa sonradan yeni ekranlar eklendiğinde kayıtlı hesaplar eski yetkileriyle kalır ve yeni sekmeleri göremez.
          Aşağıdaki düğme eksik hesapları ekler ve herkesin yetkilerini rolünün varsayılanına çeker.
        </p>
        <div className="flex flex-wrap items-center gap-2">
          <Btn variant="ghost" onClick={() => {
            setCfg((c) => {
              const have = new Set((c.users || []).map((u) => lower(u.username)));
              const users = [...c.users];
              USERS_SEED.forEach((u) => { if (!have.has(lower(u.username))) users.push({ ...u }); });
              return { ...c, v: CFG_V, users: users.map((u) => ({ ...u, perms: (ROLES[u.role]?.perms || u.perms || []).slice() })) };
            });
            setSync("Hesaplar ve yetkiler rol varsayılanına getirildi");
            setTimeout(() => setSync(""), 4000);
          }}>Eksik hesapları ekle, yetkileri yenile</Btn>
          {sync && <span className="text-sm text-emerald-700">{sync}</span>}
          <span className="flex-1" />
          <Btn onClick={() => setEdit(blank())}>Kullanıcı ekle</Btn>
        </div>
      </Card>
      {edit && (
        <Card className="p-4">
          <h3 className="mb-3 font-medium text-stone-900">{edit.isNew ? "Yeni kullanıcı" : "Kullanıcıyı düzenle"}</h3>
          <div className="grid gap-3 sm:grid-cols-2">
            <Field label="Ad soyad"><input className={inputCls} value={edit.name} onChange={(e) => setEdit({ ...edit, name: e.target.value })} /></Field>
            <Field label="Kullanıcı adı"><input className={inputCls} value={edit.username} autoCapitalize="none" onChange={(e) => setEdit({ ...edit, username: e.target.value })} /></Field>
            <Field label="Şifre"><input className={inputCls} value={edit.pin} onChange={(e) => setEdit({ ...edit, pin: e.target.value })} /></Field>
            <Field label="Rol"><select className={inputCls} value={edit.role} onChange={(e) => setEdit({ ...edit, role: e.target.value, perms: ROLES[e.target.value].perms.slice() })}>
              {Object.entries(ROLES).map(([k, v]) => <option key={k} value={k}>{v.name}</option>)}</select></Field>
            <Field label="Bağlı şube"><select className={inputCls} value={edit.branchId || ""} onChange={(e) => setEdit({ ...edit, branchId: e.target.value || null })}>
              <option value="">—</option>{cfg.branches.map((b) => <option key={b.id} value={b.id}>{b.name}</option>)}</select></Field>
          </div>
          <div className="mt-4"><p className="mb-1 text-sm font-medium text-stone-700">Görebileceği bölümler</p>
            <p className="mb-2 text-xs text-stone-500">Hiçbiri seçilmezse bütün bölümlere erişir.</p>
            <div className="flex flex-wrap gap-2">{cfg.depts.map((d) => { const on = edit.depts.includes(d.id);
              return <button key={d.id} onClick={() => setEdit({ ...edit, depts: on ? edit.depts.filter((x) => x !== d.id) : [...edit.depts, d.id] })}
                className={`rounded-full border px-3 py-1 text-sm ${on ? "border-emerald-800 bg-emerald-800 text-white" : "border-stone-300 bg-white text-stone-700"}`}>{d.name}</button>; })}</div>
          </div>
          <div className="mt-4"><p className="mb-2 text-sm font-medium text-stone-700">Yetkiler</p>
            <div className="grid gap-1 sm:grid-cols-2">{PERMS.map(([id, label]) => { const locked = id === "order.create" && edit.role !== "sube";
              return (<label key={id} className={`flex items-start gap-2 text-sm ${locked ? "text-stone-400" : "text-stone-700"}`}>
                <input type="checkbox" disabled={locked} className="mt-0.5 h-4 w-4 rounded border-stone-300" checked={!locked && edit.perms.includes(id)}
                  onChange={(e) => setEdit({ ...edit, perms: e.target.checked ? [...edit.perms, id] : edit.perms.filter((p) => p !== id) })} />{label}</label>); })}</div>
          </div>
          <div className="mt-4 flex gap-2"><Btn onClick={save}>Kaydet</Btn><Btn variant="ghost" onClick={() => setEdit(null)}>Vazgeç</Btn></div>
        </Card>
      )}
      <Card className="overflow-x-auto">
        <table className="w-full min-w-max text-sm">
          <thead className="bg-stone-50 text-left text-xs text-stone-600"><tr><th className="px-3 py-2 font-medium">Ad</th>
            <th className="px-3 py-2 font-medium">Kullanıcı adı</th><th className="px-3 py-2 font-medium">Rol</th><th className="px-3 py-2 font-medium">Bölüm / şube</th><th /></tr></thead>
          <tbody className="divide-y divide-stone-100">
            {cfg.users.map((u) => (<tr key={u.id}>
              <td className="px-3 py-2 text-stone-800">{u.name}</td><td className="px-3 py-2 text-stone-600">{u.username}</td>
              <td className="px-3 py-2 text-stone-600">{ROLES[u.role]?.name || u.role}</td>
              <td className="px-3 py-2 text-stone-600">{u.branchId ? branchName(cfg, u.branchId) : u.depts.length ? u.depts.map((d) => deptName(cfg, d)).join(", ") : "hepsi"}</td>
              <td className="whitespace-nowrap px-3 py-2 text-right">
                <button className="text-xs text-emerald-800 hover:underline" onClick={() => setEdit({ ...u, isNew: false })}>düzenle</button>
                <button className="ml-3 text-xs text-stone-600 hover:underline" title="Yetkileri rolün varsayılanına çeker"
                  onClick={() => setCfg((s) => ({ ...s, users: s.users.map((x) => (x.id === u.id ? { ...x, perms: (ROLES[x.role]?.perms || []).slice() } : x)) }))}>yetkileri yenile</button>
                <button className="ml-3 text-xs text-red-700 hover:underline" onClick={() => setCfg((s) => ({ ...s, users: s.users.filter((x) => x.id !== u.id) }))}>sil</button>
              </td></tr>))}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

function AdminSimple({ cfg, setCfg, field, label }) {
  const [name, setName] = useState("");
  const items = cfg[field];
  const add = () => {
    if (!name.trim()) return;
    const id = lower(name.trim()).replace(/[^a-z0-9ğüşıöç]+/g, "-").replace(/^-|-$/g, "");
    if (!id || items.some((i) => i.id === id)) return;
    setCfg((c) => ({ ...c, [field]: [...c[field], { id, name: name.trim() }] }));
    setName("");
  };
  return (
    <div className="space-y-4">
      <Card className="p-4"><div className="flex gap-2">
        <input className={inputCls} placeholder={label + " adı"} value={name} onChange={(e) => setName(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()} />
        <Btn onClick={add}>Ekle</Btn></div></Card>
      <Card className="divide-y divide-stone-100">
        {items.map((i) => (<div key={i.id} className="flex items-center gap-3 p-3">
          <input className="min-w-0 flex-1 rounded border border-transparent px-2 py-1 text-sm text-stone-800 hover:border-stone-300 focus:border-emerald-600 focus:outline-none"
            value={i.name} onChange={(e) => setCfg((c) => ({ ...c, [field]: c[field].map((x) => (x.id === i.id ? { ...x, name: e.target.value } : x)) }))} />
          <span className="text-xs text-stone-400">{i.id}</span>
          <button className="text-xs text-red-700 hover:underline" onClick={() => setCfg((c) => ({ ...c, [field]: c[field].filter((x) => x.id !== i.id) }))}>sil</button>
        </div>))}
      </Card>
    </div>
  );
}

/* ============================ kök ============================ */

const seedCfg = () => ({ v: CFG_V, users: USERS_SEED, depts: DEPTS_SEED, branches: BRANCHES_SEED, settings: SETTINGS_SEED, catOv: {}, catAdd: [], catDel: [] });

export default function App() {
  const [ready, setReady] = useState(false);
  const [cfg, setCfgState] = useState(() => buildIndex(seedCfg()));
  const [ordersByDate, setOrdersByDate] = useState({});
  const [loaded, setLoaded] = useState({});
  const [prodByMonth, setProdByMonth] = useState({});
  const [fireByMonth, setFireByMonth] = useState({});
  const [reqByMonth, setReqByMonth] = useState({});
  const [levels, setLevels] = useState(null);
  const [movesByMonth, setMovesByMonth] = useState({});
  const reqRef = useRef({});
  const mvRef = useRef({});
  const [me, setMe] = useState(null);
  const [tab, setTab] = useState("");
  const [busy, setBusy] = useState(false);
  const [warn, setWarn] = useState("");
  const [fetching, setFetching] = useState(0);
  const dayRef = useRef({});
  const monRef = useRef({});

  const ensure = useCallback(async (from, to) => {
    if (!from || !to || from > to) return;
    const days = rangeDates(from, to).filter((d) => !dayRef.current[d]);
    const mons = rangeMonths(from, to).filter((m) => !monRef.current[m]);
    if (!days.length && !mons.length) return;
    days.forEach((d) => { dayRef.current[d] = 1; });
    mons.forEach((m) => { monRef.current[m] = 1; });
    setFetching((x) => x + 1);
    try {
      if (days.length) {
        const got = await mapLimit(days, 4, (d) => loadKey(kOrd(d), []));
        setOrdersByDate((p) => { const n = { ...p }; days.forEach((d, i) => { n[d] = got[i]; }); return n; });
        setLoaded((p) => { const n = { ...p }; days.forEach((d) => { n[d] = 1; }); return n; });
      }
      if (mons.length) {
        const ps = await mapLimit(mons, 3, (m) => loadKey(kPrd(m), []));
        const fs = await mapLimit(mons, 3, (m) => loadKey(kFire(m), []));
        setProdByMonth((p) => { const n = { ...p }; mons.forEach((m, i) => { n[m] = ps[i]; }); return n; });
        setFireByMonth((p) => { const n = { ...p }; mons.forEach((m, i) => { n[m] = fs[i]; }); return n; });
      }
    } finally { setFetching((x) => x - 1); }
  }, []);

  /* satınalma talepleri: istenen aylar */
  const ensureReq = useCallback(async (from, to) => {
    if (!from || !to || from > to) return;
    const mons = rangeMonths(from, to).filter((m) => !reqRef.current[m]);
    if (!mons.length) return;
    mons.forEach((m) => { reqRef.current[m] = 1; });
    setFetching((x) => x + 1);
    try {
      const got = await mapLimit(mons, 3, (m) => loadKey(kReq(m), []));
      setReqByMonth((p) => { const n = { ...p }; mons.forEach((m, i) => { n[m] = got[i]; }); return n; });
    } finally { setFetching((x) => x - 1); }
  }, []);

  /* stok seviyeleri tek belgede tutulur; hareket günlüğü ay bazlıdır */
  const ensureStock = useCallback(async () => {
    if (levels !== null) return;
    setFetching((x) => x + 1);
    try { setLevels(await loadKey(kLevel, {})); } finally { setFetching((x) => x - 1); }
  }, [levels]);

  const ensureMoves = useCallback(async (from, to) => {
    if (!from || !to || from > to) return;
    const mons = rangeMonths(from, to).filter((m) => !mvRef.current[m]);
    if (!mons.length) return;
    mons.forEach((m) => { mvRef.current[m] = 1; });
    setFetching((x) => x + 1);
    try {
      const got = await mapLimit(mons, 3, (m) => loadKey(kMove(m), []));
      setMovesByMonth((p) => { const n = { ...p }; mons.forEach((m, i) => { n[m] = got[i]; }); return n; });
    } finally { setFetching((x) => x - 1); }
  }, []);

  /* başka bir cihaz veri yazdığında açık ekranı tazele */
  useEffect(() => {
    const off = subscribeKeys((key) => {
      if (key === kCfg) { loadKey(kCfg, null).then((c) => { if (c) setCfgState(buildIndex(c)); }); return; }
      if (key === kLevel) { loadKey(kLevel, {}).then((v) => setLevels(v && typeof v === "object" ? v : {})); return; }
      const m = KEY_RE.exec(key);
      if (!m) return;
      const kind = m[1], id = m[2];
      loadKey(key, []).then((v) => {
        const arr = Array.isArray(v) ? v : [];
        if (kind === "ord" && dayRef.current[id]) setOrdersByDate((p) => ({ ...p, [id]: arr }));
        else if (kind === "prd" && monRef.current[id]) setProdByMonth((p) => ({ ...p, [id]: arr }));
        else if (kind === "fire" && monRef.current[id]) setFireByMonth((p) => ({ ...p, [id]: arr }));
        else if (kind === "req" && reqRef.current[id]) setReqByMonth((p) => ({ ...p, [id]: arr }));
        else if (kind === "stkmv" && mvRef.current[id]) setMovesByMonth((p) => ({ ...p, [id]: arr }));
      });
    });
    return off;
  }, []);

  useEffect(() => {
    let alive = true;
    const seed = seedCfg;
    const guard = setTimeout(() => {
      if (!alive) return;
      setWarn("Kayıtlı ayarlar zamanında okunamadı, varsayılanlarla açıldı.");
      setReady(true);
    }, 4000);
    (async () => {
      try {
        let c = await loadKey(kCfg, null);
        if (!c) {
          c = seed();
          const ok = await saveKey(kCfg, c);
          if (!ok && alive) setWarn("Veriler kaydedilemiyor. Girdikleriniz kalıcı olmayabilir.");
        }
        delete c.catalog;
        if (!c.catOv) c.catOv = {};
        if (!c.catAdd) c.catAdd = [];
        if (!c.catDel) c.catDel = [];
        if (!c.settings) c.settings = SETTINGS_SEED;
        if (!c.users || !c.users.length) c.users = USERS_SEED;
        if (!c.depts || !c.depts.length) c.depts = DEPTS_SEED;
        if (!c.branches || !c.branches.length) c.branches = BRANCHES_SEED;
        if (migrateCfg(c)) await saveKey(kCfg, c);
        if (!alive) return;
        setCfgState(buildIndex(c));
        const ses = loadLocal(kSes, null);
        if (alive && ses) { const u = c.users.find((x) => x.id === ses); if (u) setMe(u); }
      } catch (e) {
        if (alive) setWarn("Kayıtlı ayarlar okunamadı, varsayılanlarla açıldı.");
      } finally {
        clearTimeout(guard);
        if (alive) setReady(true);
      }
    })();
    return () => { alive = false; clearTimeout(guard); };
  }, []);

  /* giriş yapıldıktan sonra yakın günleri arka planda getir */
  useEffect(() => { if (me) ensure(daysAgo(3), addDays(today(), 3)); }, [me, ensure]);

  const orders = useMemo(() => Object.values(ordersByDate).flat(), [ordersByDate]);
  const production = useMemo(() => Object.values(prodByMonth).flat(), [prodByMonth]);
  const fires = useMemo(() => Object.values(fireByMonth).flat(), [fireByMonth]);
  const requests = useMemo(() => Object.values(reqByMonth).flat(), [reqByMonth]);
  const moves = useMemo(() => Object.values(movesByMonth).flat(), [movesByMonth]);

  const setCfg = (u) => setCfgState((prev) => {
    const raw = { ...prev }; delete raw._cat; delete raw._dept; delete raw._branch; delete raw.catalog;
    const n = typeof u === "function" ? u(raw) : u;
    saveKey(kCfg, n);
    return buildIndex(n);
  });

  const putOrders = (date, arr) => setOrdersByDate((p) => ({ ...p, [date]: arr }));

  const submitOrder = async (draft) => {
    const date = draft.prodDate;
    const arr = await mutate(kOrd(date), (cur) => {
      const i = cur.findIndex((o) => o.branchId === draft.branchId && o.prodDate === date);
      if (i === -1) return [...cur, { ...draft, id: uid(), orderDate: today(), createdAt: new Date().toISOString() }];
      const o = cur[i], lines = o.lines.slice();
      draft.lines.forEach((nl) => {
        const j = lines.findIndex((l) => l.productId === nl.productId);
        if (j === -1) lines.push(nl);
        else { const l = lines[j], qty = l.qty + nl.qty;
          lines[j] = { ...l, qty, status: lineStatus(l) === "gonderildi" && sentQ(l) < qty ? "kismen" : lineStatus(l) }; }
      });
      const next = cur.slice();
      next[i] = { ...o, lines, shipDate: draft.shipDate || o.shipDate, note: [o.note, draft.note].filter(Boolean).join(" · "), updatedAt: new Date().toISOString() };
      return next;
    });
    putOrders(date, arr);
  };

  const updateLine = async (order, lineId, patch) => {
    const date = order.prodDate;
    const arr = await mutate(kOrd(date), (cur) => cur.map((o) => o.id !== order.id ? o : {
      ...o, lines: o.lines.map((l) => l.id !== lineId ? l : { ...l, ...patch, at: new Date().toISOString(), by: me?.name }),
    }));
    putOrders(date, arr);
  };
  const editLineQty = async (order, lineId, qty) => updateLine(order, lineId, { qty });
  const deleteLine = async (order, lineId) => {
    const date = order.prodDate;
    const arr = await mutate(kOrd(date), (cur) => cur.map((o) => o.id !== order.id ? o : { ...o, lines: o.lines.filter((l) => l.id !== lineId) }).filter((o) => o.lines.length));
    putOrders(date, arr);
  };

  const addProduction = async (date, rows) => {
    const m = monthKey(date);
    const arr = await mutate(kPrd(m), (cur) => { const ids = new Set(cur.map((r) => r.id)); return [...cur, ...rows.filter((r) => !ids.has(r.id))]; });
    setProdByMonth((p) => ({ ...p, [m]: arr }));
  };
  const delProduction = async (id, date) => {
    const m = monthKey(date);
    const arr = await mutate(kPrd(m), (cur) => cur.filter((r) => r.id !== id));
    setProdByMonth((p) => ({ ...p, [m]: arr }));
  };
  const addFire = async (rows, date) => {
    const m = monthKey(date || rows[0]?.date || today());
    const arr = await mutate(kFire(m), (cur) => { const ids = new Set(cur.map((r) => r.id)); return [...cur, ...rows.filter((r) => !ids.has(r.id))]; });
    setFireByMonth((p) => ({ ...p, [m]: arr }));
  };
  const delFire = async (id, date) => {
    const m = monthKey(date);
    const arr = await mutate(kFire(m), (cur) => cur.filter((r) => r.id !== id));
    setFireByMonth((p) => ({ ...p, [m]: arr }));
  };

  /* ---- satınalma ---- */
  const addRequests = async (rows) => {
    const m = monthKey(rows[0]?.reqDate || today());
    const arr = await mutate(kReq(m), (cur) => { const ids = new Set(cur.map((r) => r.id)); return [...cur, ...rows.filter((r) => !ids.has(r.id))]; });
    reqRef.current[m] = 1;
    setReqByMonth((p) => ({ ...p, [m]: arr }));
  };
  const patchRequest = async (req, patch) => {
    const m = monthKey(req.reqDate);
    const arr = await mutate(kReq(m), (cur) => cur.map((r) => (r.id === req.id ? { ...r, ...patch } : r)));
    setReqByMonth((p) => ({ ...p, [m]: arr }));
  };

  /* ---- stok ---- */
  const applyStock = async (rows) => {
    /* rows: [{itemId, name, unit, type:'giris'|'cikis'|'sayim', qty, counted, note, ref, date}] */
    if (!rows.length) return;
    let log = [];
    const saved = await mutate(kLevel, (cur) => {
      const start = cur && typeof cur === "object" && !Array.isArray(cur) ? cur : {};
      const next = { ...start };
      log = [];
      rows.forEach((r) => {
        const base = next[r.itemId] !== undefined ? Number(next[r.itemId]) : Number(catFind(cfg, r.itemId)?.ss) || 0;
        let after = base, delta = 0;
        if (r.type === "sayim") { after = Number(r.counted) || 0; delta = after - base; }
        else if (r.type === "cikis") { delta = -Math.abs(Number(r.qty) || 0); after = base + delta; }
        else { delta = Math.abs(Number(r.qty) || 0); after = base + delta; }
        next[r.itemId] = Math.round(after * 100) / 100;
        log.push({ id: uid(), date: r.date || today(), itemId: r.itemId, name: r.name, unit: r.unit,
          type: r.type, qty: Math.round(delta * 100) / 100, after: next[r.itemId], note: r.note || "",
          ref: r.ref || "", by: me?.name || "", at: new Date().toISOString() });
      });
      return next;
    });
    setLevels(saved && typeof saved === "object" && !Array.isArray(saved) ? saved : {});
    if (!log.length) return;
    const m = monthKey(log[0].date);
    const arr = await mutate(kMove(m), (cur) => [...(Array.isArray(cur) ? cur : []), ...log]);
    mvRef.current[m] = 1;
    setMovesByMonth((p) => ({ ...p, [m]: arr }));
  };

  const createRequests = async (rows) => {
    const out = rows.map((r) => {
      const c = catFind(cfg, r.id);
      return { id: uid(), reqDate: today(), itemId: r.id, code: c?.c || r.id, name: c?.n || r.id, unit: c?.u || "",
        qty: Number(r.qty), urgency: r.urgency || URGENCY[0], note: (r.note || "").trim(),
        by: me?.name || "", byId: me?.id || "", at: new Date().toISOString(), status: "talep" };
    });
    if (out.length) await addRequests(out);
  };

  const receiveRequest = async (req, { qty, date, note }) => {
    const gelen = Number(qty) || 0;
    if (gelen <= 0) return;
    await applyStock([{ itemId: req.itemId, name: req.name, unit: req.unit, type: "giris", qty: gelen,
      note: note || `Satınalma teslimi${req.supplier ? " · " + req.supplier : ""}`, ref: req.id, date }]);
    const total = Math.round(((Number(req.recvQty) || 0) + gelen) * 100) / 100;
    await patchRequest(req, { recvQty: total, recvDate: date, recvBy: me?.name || "", recvNote: note || "",
      status: total >= (Number(req.qty) || 0) ? "teslim" : "kismi" });
  };

  const refresh = async () => {
    setBusy(true);
    try {
      const dates = Object.keys(loaded);
      if (dates.length) {
        const got = await mapLimit(dates, 4, (d) => loadKey(kOrd(d), []));
        const obd = {}; dates.forEach((d, i) => { obd[d] = got[i]; });
        setOrdersByDate(obd);
      }
      const ms = Object.keys(prodByMonth);
      if (ms.length) {
        const ps = await mapLimit(ms, 3, (m) => loadKey(kPrd(m), []));
        const fs = await mapLimit(ms, 3, (m) => loadKey(kFire(m), []));
        const pm = {}, fm = {}; ms.forEach((m, i) => { pm[m] = ps[i]; fm[m] = fs[i]; });
        setProdByMonth(pm); setFireByMonth(fm);
      }
      const rm = Object.keys(reqRef.current);
      if (rm.length) {
        const rs = await mapLimit(rm, 3, (m) => loadKey(kReq(m), []));
        const o = {}; rm.forEach((m, i) => { o[m] = rs[i]; }); setReqByMonth(o);
      }
      const vm = Object.keys(mvRef.current);
      if (vm.length) {
        const vs = await mapLimit(vm, 3, (m) => loadKey(kMove(m), []));
        const o = {}; vm.forEach((m, i) => { o[m] = vs[i]; }); setMovesByMonth(o);
      }
      if (levels !== null) setLevels(await loadKey(kLevel, {}));
    } finally { setBusy(false); }
  };

  const login = async (u) => { setMe(u); setTab(""); saveLocal(kSes, u.id); };
  const logout = async () => { setMe(null); removeLocal(kSes); };

  if (!ready) return <div className="flex min-h-screen items-center justify-center bg-stone-100 text-stone-600">Yükleniyor…</div>;
  if (!me) return <Login users={cfg.users} onLogin={login} />;

  const can = (p) => me.perms.includes(p);
  const canOrder = can("order.create") && me.role === "sube" && !!me.branchId;
  const canProduce = can("production.entry");
  const canFire = can("fire.entry");
  const isDepo = me.depts.length === 1 && me.depts[0] === "depo";

  const navItems = [
    canOrder && ["yeni", "Talep gir"],
    can("order.view.own") && ["talepler", "Talepler"],
    can("dept.queue") && ["bolum", isDepo ? "Depo talepleri" : "Şube talepleri"],
    can("shipping") && ["sevk", isDepo ? "Depo sevk" : "Sevkiyat"],
    (canProduce || canFire) && ["uretim", canProduce ? "Üretim girişi" : "Fire girişi"],
    can("supply.request") && ["depoistek", "Depo talep listesi"],
    can("stock.view") && ["stok", "Depo stoğu"],
    can("purchase.view") && ["satinalma", "Satınalma"],
    can("list.view") && ["liste", "Talep listesi"],
    (can("report.all") || can("report.dept")) && ["rapor", "Raporlar"],
    can("admin") && ["admin", "Yönetim"],
  ].filter(Boolean);
  const active = navItems.some((n) => n[0] === tab) ? tab : navItems[0]?.[0] || "";

  return (
    <div className="min-h-screen bg-stone-100 font-sans">
      <header className="sticky top-0 z-30 border-b border-emerald-900 bg-emerald-950 text-white">
        <div className="mx-auto flex max-w-6xl items-center gap-3 px-4 py-3">
          <div className="min-w-0 flex-1">
            <div className="truncate text-sm font-semibold">{me.name}</div>
            <div className="truncate text-xs text-emerald-200">{ROLES[me.role]?.name || me.role}
              {me.branchId ? " · " + branchName(cfg, me.branchId) : ""}
              {!me.branchId && me.depts.length && me.depts.length < cfg.depts.length ? " · " + me.depts.map((d) => deptName(cfg, d)).join(", ") : ""}</div>
          </div>
          {fetching > 0 && <span className="text-xs text-emerald-300">veriler yükleniyor…</span>}
          <span className="hidden text-xs text-emerald-300 sm:block">{dateTR(today())}</span>
          <button onClick={refresh} className="rounded px-2 py-1 text-xs text-emerald-200 hover:bg-emerald-900">{busy ? "…" : "Yenile"}</button>
          <button onClick={logout} className="rounded px-2 py-1 text-xs text-emerald-200 hover:bg-emerald-900">Çıkış</button>
        </div>
        <nav className="mx-auto max-w-6xl overflow-x-auto px-2"><div className="flex gap-1 pb-1">
          {navItems.map(([id, label]) => (<button key={id} onClick={() => setTab(id)}
            className={`whitespace-nowrap rounded-t px-3 py-2 text-sm font-medium ${active === id ? "bg-stone-100 text-emerald-950" : "text-emerald-100 hover:bg-emerald-900"}`}>{label}</button>))}
        </div></nav>
      </header>

      <main className="mx-auto max-w-6xl p-4 pb-10">
        {(warn || storeMode() === "memory") && (
          <div className="mb-4 rounded-lg border border-amber-300 bg-amber-50 px-4 py-3 text-sm text-amber-900">
            {warn || "Bu ortamda kalıcı depolama açılamadı. Girdikleriniz yalnızca bu sekme açık kaldığı sürece tutulur; sayfayı yenilerseniz silinir."}
          </div>
        )}
        {active === "yeni" && <NewOrder me={me} cfg={cfg} orders={orders} onSubmit={submitOrder} ensure={ensure} />}
        {active === "talepler" && <OrderList orders={orders} cfg={cfg} me={me} canSeeAll={!me.branchId} onEditLine={editLineQty} onDeleteLine={deleteLine} ensure={ensure} />}
        {active === "bolum" && <DeptRequests me={me} cfg={cfg} orders={orders} onUpdateLine={updateLine} ensure={ensure} />}
        {active === "sevk" && <ShipScreen me={me} cfg={cfg} orders={orders} onUpdateLine={updateLine} onAddFire={addFire} ensure={ensure} />}
        {active === "uretim" && <EntryScreen me={me} cfg={cfg} orders={orders} production={production} fires={fires}
          onAdd={addProduction} onAddFire={addFire} onDelete={delProduction} onDeleteFire={delFire} canProduce={canProduce} canFire={canFire} ensure={ensure} />}
        {active === "stok" && <DepotStock cfg={cfg} levels={levels} moves={moves} requests={requests}
          ensureStock={ensureStock} ensureMoves={ensureMoves} ensureReq={ensureReq}
          onCount={applyStock} onRequest={createRequests}
          canCount={can("stock.count")} canRequest={can("purchase.request")} />}
        {active === "satinalma" && <Purchasing cfg={cfg} me={me} requests={requests} levels={levels}
          ensureReq={ensureReq} ensureStock={ensureStock} onAdd={createRequests} onPatch={patchRequest}
          onReceive={receiveRequest} canBuy={can("purchase.buy")} canRequest={can("purchase.request")} />}
        {active === "depoistek" && <DeptSupply me={me} cfg={cfg} orders={orders} onSubmit={submitOrder} ensure={ensure} />}
        {active === "liste" && <DriverList cfg={cfg} orders={orders} ensure={ensure} />}
        {active === "rapor" && <Reports me={me} cfg={cfg} orders={orders} production={production} fires={fires} ensure={ensure} />}
        {active === "admin" && <Admin cfg={cfg} setCfg={setCfg} orders={orders} production={production} fires={fires} />}
        {!navItems.length && <Empty title="Bu hesaba tanımlı yetki yok" hint="Yöneticinizden yetki isteyin." />}
      </main>
    </div>
  );
}
