# 🚀 Crypto to Excel Automation 📈

<p align="center"><b>Otomatisasi update harga & reward kripto ke file Excel menggunakan API CoinGecko dan CoinMarketCap</b></p>

---

## 🧩 Ringkasan

**Crypto Excel Price & Reward Updater** adalah kumpulan skrip Python untuk:

* Mengambil harga kripto terbaru dari CoinGecko & CoinMarketCap
* Menyinkronkan estimasi nilai reward ke Excel secara otomatis
* Monitoring reward kripto seperti BYBIT SPLASH dan airdrop lainnya

---

## 💡 Studi Kasus: Airdrop & Reward BYBIT SPLASH

Alat ini dirancang untuk:

* Memantau estimasi nilai reward kripto dari volume transaksi
* Merekam histori harga & nilai reward ke Excel
* Menyimpan data di lokal dan cloud (OneDrive) jika diaktifkan

### Keunggulan:

* 🔄 Update harga otomatis dari 2 sumber
* 📊 Estimasi nilai reward real-time
* ☁️ Backup lokal & cloud
* 📋 Cocok untuk semua program reward dan airdrop exchange

---

## 🛠️ Fitur Utama

* Update harga otomatis dari CoinGecko dan CoinMarketCap
* Menulis data harga dan waktu update ke Excel
* Dukungan multi-koin dan multi-file
* Konversi waktu update ke WIB (Asia/Jakarta)
* Tampilan terminal berwarna (CMC)
* Otomatis jeda jika terkena limit API
* Sinkronisasi reward dengan harga terbaru

---

## 📦 Instalasi

```bash
git clone https://github.com/username/price_update.git
cd price_update
pip install -r requirements.txt
```

### Dependencies:

* `requests`, `pandas`, `openpyxl`, `pytz`, `colorama`

---

## 📁 Struktur Project

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

```python
from fungsi_coingecko import run
run("PATH/TO/EXCEL.xlsx")
```

```python
from fungsi_coinmarketcap import main
main("PATH/TO/EXCEL.xlsx")
```

### 2. Sinkronisasi Reward

```python
from fungsi_reward import main
main()
```

---

## 🧠 Penjelasan File Kode

| File Python               | Deskripsi                                      |
| ------------------------- | ---------------------------------------------- |
| `fungsi_coingecko.py`     | Update kolom "CG PRICE" & "LAST UPDATE (CG)"   |
| `fungsi_coinmarketcap.py` | Update kolom "CMC PRICE" & "LAST UPDATE (CMC)" |
| `fungsi_reward.py`        | Sinkronisasi nilai reward di sheet `REWARD`    |

---

## 🖥️ Otomasi Lewat Batch (Windows)

```bat
@echo off
python "update.py"
pause
```

Simpan sebagai `update.bat` dan jalankan untuk update otomatis.

---

## 🖼️ Contoh Tampilan

* CoinGecko (Excel): ![CoinGecko](asset/coingecko.png)
* CoinMarketCap (Terminal): ![CMC](asset/coinmarketcap.png)
* Sinkronisasi Reward: ![Hasil](asset/hasil.png)

---

## ⚠️ Tips Keamanan

* Jangan buka file Excel saat script dijalankan
* Jangan bagikan `apikey.json`
* Tambahkan `apikey.json` ke `.gitignore`

---

## 🛡️ Sistem Backup

Tool ini bisa digunakan untuk menyimpan data ke beberapa file Excel sekaligus (lokal & cloud).

---

*Dibuat dengan ❤️ oleh Ahmad Nur Ikhsan*
