# Merhaba Pastaneleri · Üretim ve Sevkiyat — kurulum

Bu klasör, programı internete almak için gereken her şeyi içerir.
Sırayla üç adım var: **Supabase → GitHub → Cloudflare**.

---

## 1) Supabase (veritabanı)

1. Supabase panelinde projeni aç → sol menüden **SQL Editor** → **New query**.
2. Bu klasördeki **`supabase-kurulum.sql`** dosyasının içeriğini kopyalayıp yapıştır → **Run**.
3. Altta tek satırlık bir sonuç çıkmalı: `kayit_sayisi = 0`, `politika_sayisi = 1`, `canli_yenileme = 1`.
   Üçü de böyleyse veritabanı hazır.
Proje adresi ve anahtar `src/supabase.js` dosyasına zaten yazıldı, orada
yapacak bir şey yok.

---

## 2) GitHub (dosyaların durduğu yer)

1. github.com → sağ üstten **+** → **New repository**
   - Repository name: `merhaba-erp`
   - **Private** seç
   - "Add a README file" kutusunu **işaretleme**
   - **Create repository**
2. Açılan sayfada **"uploading an existing file"** bağlantısına tıkla.
3. Bu klasörün **içindeki her şeyi** (klasörler dahil) sürükleyip bırak:
   `src`, `public`, `index.html`, `package.json`, `vite.config.js`,
   `tailwind.config.js`, `postcss.config.js`, `.nvmrc`, `.gitignore`
4. Aşağıdaki **Commit changes** düğmesine bas.

> `node_modules` veya `dist` klasörü yükleme — zaten yok, olmamalı.

---

## 3) Cloudflare Pages (yayın)

1. dash.cloudflare.com → **Workers & Pages** → **Create** → **Pages** →
   **Connect to Git** → GitHub hesabını bağla → `merhaba-erp` deposunu seç.
2. Ayarlar:
   - **Project name:** `merhaba-erp` (küçük harf, boşluksuz)
   - **Production branch:** `main`  ← depoya dosya yükledikten sonra listede çıkar
   - **Framework preset:** None
   - **Build command:** `npm run build`
   - **Build output directory:** `dist`
3. **Environment variables (advanced)** bölümüne tek değişken ekle:
   - `NODE_VERSION` = `20`
4. **Save and Deploy**.

2-3 dakika sonra `https://merhaba-erp.pages.dev` adresi hazır olur.

---

## Sonrasında

- **Giriş:** `admin` / `1234`. İlk iş: Yönetim → Kullanıcılar'dan bütün PIN'leri değiştir.
- **Kontrol:** Yönetim → Ayarlar → "Veri saklama" kutusunda
  **"Ortak veritabanı (Supabase)"** yazmalı. "Bu tarayıcı" yazıyorsa
  `src/supabase.js` içindeki adres yanlış demektir.
- **İki cihazdan dene:** telefondan bir talep gir, bilgisayarda ekranın
  kendiliğinden tazelenmesi gerekir.
- **Tablete kurmak:** adresi Chrome/Safari'de aç → menüden
  "Ana ekrana ekle" → uygulama gibi açılır.
- **Değişiklik yapmak:** GitHub'da dosyayı düzenleyip kaydettiğin anda
  Cloudflare yeniden yayınlar, 2 dakikada canlıya geçer.

## Bilinmesi gerekenler

- Supabase ücretsiz planda proje **1 hafta hiç kullanılmazsa duraklatılır**;
  panelden tek tıkla geri açılır, veri kaybolmaz.
- Ücretsiz planda otomatik yedek yok. Ayda bir Raporlar ekranından
  Excel çıktısı alıp saklamanı öneririm.
- Güvenlik seviyesi "iç kullanım" düzeyindedir: adresi ve bir PIN'i bilen girer.
  Daha sıkısı gerekirse Cloudflare Zero Trust (ücretsiz, 50 kullanıcıya kadar)
  ile sitenin önüne e-posta doğrulamalı giriş konulabilir.
