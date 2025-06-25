
<h1 align="center">🚀 Crypto Excel Price & Reward Updater 📈</h1>
<p align="center">
  <b>Otomatisasi update harga & reward kripto ke file Excel menggunakan API CoinGecko dan CoinMarketCap</b>
</p>

---

## 🚀 Deskripsi Singkat

**Crypto Excel Price & Reward Updater** adalah kumpulan skrip Python yang secara otomatis:
- Mengambil harga terbaru cryptocurrency dari dua sumber berbeda (CoinGecko & CoinMarketCap) dan memperbaruinya ke file Excel Anda.
- Menyinkronkan dan memperbarui data reward di Excel berdasarkan harga terbaru.
- Dengan dua sumber harga, Anda bisa membandingkan harga dan mendapatkan hasil yang lebih akurat, serta memastikan reward selalu terupdate.

---

## 🏆 Studi Kasus: Monitoring Reward & Airdrop Crypto (Bybit Splash & Lainnya)

Proyek ini dibuat sebagai **alat bantu monitoring usaha pada BYBIT SPLASH dan airdrop crypto lainnya**.  
Pada program airdrop seperti ini, reward biasanya diberikan dalam bentuk koin kripto dengan harga yang selalu berubah-ubah.  
Syarat mendapatkan reward biasanya berdasarkan volume transaksi tertentu, dan reward yang didapat bisa berbeda-beda setiap waktu.

Dengan alat ini, Anda cukup **sekali input jenis koin dan jumlah koin yang didapat** pada file Excel, maka:
- Estimasi nilai reward (dalam USD atau mata uang lain) akan otomatis terupdate secara real-time.
- Data harga diambil dari dua sumber (CoinGecko & CoinMarketCap) untuk akurasi dan backup.
- Semua data reward, volume, dan estimasi value akan otomatis tersinkronisasi dan tercatat di Excel, baik di file lokal maupun cloud (OneDrive) sebagai backup otomatis.

### Manfaat untuk Pengguna BYBIT SPLASH & Airdrop Crypto:
- **Tidak perlu cek harga manual:** Harga dan estimasi reward selalu update otomatis.
- **Backup otomatis:** Data Anda aman di dua lokasi (lokal & cloud) jika diaktifkan.
- **Mudah monitoring:** Semua histori reward, volume, dan estimasi value tercatat rapi di Excel.
- **Bisa digunakan untuk berbagai event airdrop dan reward exchange lain.**

> **Catatan:**  
> Anda bisa menyesuaikan path file Excel, sheet, dan kolom sesuai kebutuhan.  
> Tool ini sangat cocok untuk trader, hunter airdrop, dan siapapun yang ingin memantau reward kripto.

---

## 🛠️ Fitur

- 🔄 Update harga kripto otomatis dari CoinGecko dan CoinMarketCap API.
- 📝 Menulis harga dan waktu update ke kolom khusus di sheet Excel.
- 🌏 Konversi waktu update ke zona waktu WIB (Asia/Jakarta).
- 📊 Mendukung multi-koin sekaligus.
- 🎨 Tampilan terminal berwarna (khusus CoinMarketCap).
- ⏳ Otomatis jeda jika request API sudah mencapai limit.
- 🔗 Sinkronisasi otomatis data reward dengan harga terbaru di Excel.
- 📁 **Mendukung multi-file Excel untuk backup** (bisa custom, contoh di README hanya 1 file, bisa diperbanyak sesuai kebutuhan).

---

## 📦 Instalasi

1. **Clone repository ini:**
    ```bash
    git clone https://github.com/username/price_update.git
    ```
2. **Install dependencies:**
    ```bash
    pip install -r requirements.txt
    ```

**Dependencies:**
- requests
- pandas
- openpyxl
- pytz
- colorama

---

## 📂 Struktur Folder Project

```
price_update/
├── asset/
│   ├── coingecko.png
│   ├── coinmarketcap.png
│   └── hasil.png
├── doc/
│   ├── coba.xlsx
│   └── tes.xlsx
├── apikey.json
├── fungsi_coingecko.py
├── fungsi_coinmarketcap.py
├── fungsi_reward.py
├── update.py
├── requirements.txt
├── README.md
└── .gitignore
```

---

## ⚙️ Cara Penggunaan

### 1. Update Harga Kripto

- **CoinGecko:**
    ```python
    from fungsi_coingecko import run
    run(r"PATH/TO/YOUR/EXCEL.xlsx")
    ```

- **CoinMarketCap:**
    ```python
    from fungsi_coinmarketcap import main
    main(r"PATH/TO/YOUR/EXCEL.xlsx")
    ```

### 2. Sinkronisasi & Update Reward Otomatis

- **Reward Updater:**
    ```python
    from fungsi_reward import main
    main()
    ```

---

## ⚡ Penjelasan Singkat Kode

| File Python              | Fungsi                                                                 |
|--------------------------|------------------------------------------------------------------------|
| `fungsi_coingecko.py`    | Menulis ke kolom "CG PRICE" (F) dan "LAST UPDATE (CG)" (I)             |
| `fungsi_coinmarketcap.py`| Menulis ke kolom "CMC PRICE" (E) dan "LAST UPDATE (CMC)" (H)           |
| `fungsi_reward.py`       | Menyinkronkan reward berdasarkan harga terbaru di sheet `REWARD`       |

---

### ▶️ Menjalankan Script Otomatis Lewat Batch (Windows)

```bat
@echo off
python "update.py"
pause
```

**Langkah:**
1. Simpan sebagai `update.bat`
2. Double klik untuk menjalankan otomatis update harga dan reward

---

## 🖼️ Contoh Tampilan

### 1. CoinGecko (Excel)
![CoinGecko](asset/coingecko.png)

### 2. CoinMarketCap (Terminal)
![CMC](asset/coinmarketcap.png)

### 3. Sinkronisasi Sheet Reward
![Hasil](asset/hasil.png)


---

## ⚠️ Catatan

- Pastikan file Excel tidak sedang dibuka saat menjalankan skrip.
- Jangan pernah membagikan file `apikey.json` ke publik.
- Tambahkan `apikey.json` ke `.gitignore` agar tidak ikut terupload ke GitHub.

---

### 🔒 Sistem Backup Otomatis

Anda dapat mengatur update otomatis ke lebih dari satu file Excel (lokal dan cloud) untuk keperluan backup.

---

Dibuat dengan ❤️ oleh AHMAD NUR IKHSAN
