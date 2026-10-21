import { createClient } from "@supabase/supabase-js";

/* ─────────────────────────────────────────────────────────────
   SUPABASE BAĞLANTISI

   Aşağıdaki iki satırı Supabase panelinden alırsın:
     Settings → Data API  →  "Project URL"
     Settings → API Keys  →  "Publishable key"

   Bu iki değer tarayıcıya gider, gizli değildir. Gizli olan
   "Secret key"i buraya ASLA yazma.
   ───────────────────────────────────────────────────────────── */

const SUPABASE_URL = "https://arydcvqveumwguorvwiw.supabase.co";
const SUPABASE_KEY = "sb_publishable_Gsswjp3RdalIRGwyoPZVfA_vq6pN2WS";

/* İstersen bu değerleri koda yazmak yerine Cloudflare Pages'te
   "Environment variables" bölümüne VITE_SUPABASE_URL ve
   VITE_SUPABASE_KEY olarak da girebilirsin; oradakiler önceliklidir. */
const url = import.meta.env.VITE_SUPABASE_URL || SUPABASE_URL;
const key = import.meta.env.VITE_SUPABASE_KEY || SUPABASE_KEY;

const configured = /^https:\/\/[a-z0-9-]+\.supabase\.co\/?$/i.test(url) && key.length > 20;

/* Ayar yapılmamışsa null döner; program o zaman tek cihaz modunda
   (tarayıcı hafızası) çalışmaya devam eder, veri kaybolmaz. */
export const supabase = configured
  ? createClient(url, key, {
      auth: { persistSession: false },
      realtime: { params: { eventsPerSecond: 5 } },
    })
  : null;

if (!configured && typeof console !== "undefined") {
  console.warn("[ERP] Supabase ayarlanmadı; veriler yalnızca bu tarayıcıda tutulacak.");
}
