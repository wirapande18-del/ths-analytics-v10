UPDATE V8 LEAD COLOR

Yang ditambahkan:
1. Data Lama / Pembanding sekarang mode GABUNG/TAMBAH, tidak menghapus data lama.
2. Data duplikat otomatis dilewati berdasarkan tanggal, plat/no rangka, repair type, SA, dan KM.
3. Warna status lead di tabel:
   - Hijau: Sudah Datang
   - Biru: Datang Lebih Awal
   - Kuning: Due bulan ini
   - Orange: Telat 1-30 hari
   - Merah: Telat >30 hari
   - Merah tua: Lost Customer
   - Ungu: Customer Baru
4. Grafik dashboard:
   - Status Lead / Warna
   - Lead Time
   - Repair Type
   - Estimasi per Bulan
5. Klik grafik Status Lead untuk membuka detail customer.
6. Export Excel lebih lengkap:
   - Sheet Dashboard
   - Sheet Detail Filter
   - Sheet Semua Data
   - Kolom Status Warna dan Lead Bucket untuk filter.

Cara deploy:
1. Upload folder ini ke GitHub.
2. Vercel akan menjalankan npm install dan npm run build.
3. Pastikan environment variable VITE_SUPABASE_URL dan VITE_SUPABASE_ANON_KEY tetap diisi.
