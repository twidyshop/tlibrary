require('dotenv').config();
const express = require('express');
const crypto = require('crypto');
const axios = require('axios');
const midtransClient = require('midtrans-client');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx'); 

const app = express();
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));
app.use(express.static('public'));

const dbFile = path.join(__dirname, 'transactions.json');
const digitalDbFile = path.join(__dirname, 'digital_products.json');

const readDB = () => {
    try {
        if (!fs.existsSync(dbFile)) return [];
        return JSON.parse(fs.readFileSync(dbFile, 'utf8'));
    } catch (e) {
        return [];
    }
};

const saveDB = (data) => {
    fs.writeFileSync(dbFile, JSON.stringify(data.slice(-100), null, 2));
};

const readDigitalDB = () => {
    try {
        if (!fs.existsSync(digitalDbFile)) return [];
        return JSON.parse(fs.readFileSync(digitalDbFile, 'utf8'));
    } catch (e) { return []; }
};

const saveDigitalDB = (data) => {
    fs.writeFileSync(digitalDbFile, JSON.stringify(data, null, 2));
};

let snap = new midtransClient.Snap({
  isProduction: true,
  serverKey: process.env.MIDTRANS_SERVER_KEY,
  clientKey: process.env.MIDTRANS_CLIENT_KEY
});

// --- FUNGSI DETEKSI SUB-KATEGORI OTOMATIS KHUSUS EBOOK ---
function detectEbookSubCategory(name, description) {
    const text = (name + " " + description).toLowerCase();
    
    const kamusKategori = {
        'Novel & Fiksi': ['novel', 'fiksi', 'cerpen', 'romance', 'fantasi', 'thriller', 'misteri'],
        'Buku Anak': ['anak', 'dongeng', 'balita', 'kids', 'cerita anak', 'buku mewarnai'],
        'Remaja': ['remaja', 'teen', 'young adult', 'sekolah menengah', 'masa muda'],
        'Agama & Spiritual': ['agama', 'spiritual', 'islam', 'kristen', 'doa', 'ibadah', 'tuhan', 'iman'],
        'Sejarah & Budaya': ['sejarah', 'budaya', 'history', 'kerajaan', 'kuno', 'culture', 'tradisi'],
        'Pengembangan Diri': ['pengembangan diri', 'self improvement', 'motivasi', 'produktivitas', 'self-help', 'sukses'],
        'Psikologi': ['psikologi', 'mental', 'jiwa', 'psychology', 'mindset', 'trauma', 'pikiran'],
        'Bisnis & Keuangan': ['bisnis', 'keuangan', 'marketing', 'saham', 'investasi', 'uang', 'ekonomi', 'cuan', 'startup'],
        'Pendidikan': ['pendidikan', 'buku sekolah', 'pelajaran', 'ujian', 'cpns', 'soal', 'kampus', 'guru'],
        'Teknologi': ['teknologi', 'coding', 'javascript', 'programmer', 'aplikasi', 'web', 'komputer', 'tutorial it', 'python', 'ai'],
        'Kesehatan & Kebugaran': ['kesehatan', 'kebugaran', 'medis', 'diet', 'dokter', 'penyakit', 'olahraga', 'fitness', 'sehat'],
        'Resep & Masakan': ['resep', 'masakan', 'masak', 'kuliner', 'makanan', 'food', 'kue', 'dapur'],
        'Seni & Desain': ['seni', 'desain', 'art', 'design', 'menggambar', 'melukis', 'grafis', 'ilustrasi'],
        'Hobi & Keterampilan': ['hobi', 'keterampilan', 'craft', 'kerajinan', 'berkebun', 'jahit', 'fotografi', 'otomotif'],
        'Biografi & Memoar': ['biografi', 'memoar', 'tokoh', 'kisah hidup', 'biography', 'perjalanan hidup'],
        'Politik & Sosial': ['politik', 'sosial', 'hukum', 'pemerintahan', 'sosiologi', 'negara', 'masyarakat'],
        'Sastra & Puisi': ['sastra', 'puisi', 'sajak', 'syair', 'literature', 'pantun'],
        'Komik': ['komik', 'manga', 'manhwa', 'comic', 'webtoon', 'grafis novel'],
        'Travel & Wisata': ['travel', 'wisata', 'jalan-jalan', 'liburan', 'panduan wisata', 'guidebook', 'destinasi'],
        'Referensi': ['referensi', 'kamus', 'ensiklopedia', 'jurnal', 'pedoman', 'panduan resmi', 'direktori']
    };

    for (const [subCat, keywords] of Object.entries(kamusKategori)) {
        if (keywords.some(keyword => text.includes(keyword))) {
            return subCat;
        }
    }
    
    return 'Lainnya'; 
}
// --- END FUNGSI DETEKSI ---

let cachedProducts = null;
let cacheTimestamp = 0;
const CACHE_DURATION = 5 * 60 * 1000;

app.get('/api/config', (req, res) => {
  res.json({ clientKey: process.env.MIDTRANS_CLIENT_KEY });
});

// --- FUNGSI HELPER: KIRIM EMAIL VIA RESEND ---
async function sendEmailReceipt(trx, targetEmail) {
    const resendKey = process.env.RESEND_API_KEY;
    if (!resendKey) {
        console.log('[Resend] API Key tidak ditemukan. Melewati pengiriman email.');
        return;
    }

    let itemsHtml = '';
    if (trx.cart_items && Array.isArray(trx.cart_items)) {
        itemsHtml = trx.cart_items.map(item => `
            <div style="margin-bottom: 15px; padding: 15px; border: 1px solid #e5e7eb; border-radius: 8px; background-color: #ffffff;">
                <h4 style="margin: 0 0 10px 0; color: #1f2937; font-size: 16px;">${item.name}</h4>
                <a href="${item.downloadUrl}" style="background-color: #3b82f6; color: #ffffff; padding: 10px 15px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold; font-size: 14px;">Unduh Produk</a>
            </div>
        `).join('');
    } else {
        itemsHtml = `
            <div style="margin-bottom: 15px; padding: 15px; border: 1px solid #e5e7eb; border-radius: 8px; background-color: #ffffff;">
                <h4 style="margin: 0 0 10px 0; color: #1f2937; font-size: 16px;">${trx.product_name}</h4>
                <a href="${trx.download_url}" style="background-color: #3b82f6; color: #ffffff; padding: 10px 15px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: bold; font-size: 14px;">Unduh Produk</a>
            </div>
        `;
    }

    const emailHtml = `
        <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 600px; margin: 0 auto; color: #374151; background-color: #f9fafb; padding: 20px; border-radius: 12px; border: 1px solid #e5e7eb;">
            <div style="text-align: center; margin-bottom: 20px;">
                <h1 style="color: #2563eb; margin: 0; font-size: 24px; font-weight: 800; letter-spacing: 1px;">TLIBRARY</h1>
                <p style="margin: 5px 0 0 0; font-size: 12px; color: #6b7280;">tlibrary.my.id</p>
            </div>
            
            <div style="background-color: #ffffff; padding: 20px; border-radius: 8px; border: 1px solid #e5e7eb;">
                <h2 style="color: #111827; margin-top: 0; font-size: 20px;">Terima Kasih atas Pembelian Anda! 🎉</h2>
                <p style="line-height: 1.6;">Pembayaran untuk pesanan digital Anda telah berhasil dikonfirmasi. Berikut adalah detail pesanan dan tautan akses produk Anda:</p>
                
                <div style="background-color: #f3f4f6; padding: 15px; border-radius: 8px; margin: 20px 0; font-size: 14px;">
                    <p style="margin: 5px 0;"><strong>Order ID:</strong> <span style="font-family: monospace;">${trx.order_id}</span></p>
                    <p style="margin: 5px 0;"><strong>Total Bayar:</strong> Rp ${trx.amount.toLocaleString('id-ID')}</p>
                    <p style="margin: 5px 0;"><strong>Tanggal:</strong> ${new Date().toLocaleString('id-ID')}</p>
                </div>

                <h3 style="color: #111827; margin-bottom: 15px; font-size: 16px; border-bottom: 2px solid #e5e7eb; padding-bottom: 8px;">Daftar Produk & Link Unduh:</h3>
                ${itemsHtml}
            </div>

            <p style="margin-top: 20px; font-size: 12px; color: #9ca3af; text-align: center; line-height: 1.5;">
                Harap simpan email ini sebagai bukti pembelian yang sah.<br>
                Jika Anda memiliki pertanyaan, silakan hubungi Customer Service kami via WhatsApp.<br><br>
                &copy; ${new Date().getFullYear()} TLIBRARY. All rights reserved.
            </p>
        </div>
    `;

    try {
        const response = await axios.post('https://api.resend.com/emails', {
            from: 'TLIBRARY <noreply@tlibrary.my.id>',
            to: targetEmail,
            subject: `✅ Akses Produk: Pesanan Anda Berhasil! (${trx.order_id})`,
            html: emailHtml
        }, {
            headers: {
                'Authorization': `Bearer ${resendKey}`,
                'Content-Type': 'application/json'
            }
        });
        console.log(`[Resend] Sukses kirim email nota ke ${targetEmail} (ID: ${response.data.id})`);
    } catch (error) {
        console.error(`[Resend Error] Gagal kirim email:`, error.response ? error.response.data : error.message);
    }
}
// --- END FUNGSI HELPER ---

// --- FITUR PRODUK DIGITAL & ADMIN ---
app.get('/api/digital-products', (req, res) => {
    res.json({ data: readDigitalDB() });
});

app.post('/api/admin/login', (req, res) => {
    const { username, password } = req.body;
    const adminUser = process.env.ADMIN_USER || 'admin';
    const adminPass = process.env.ADMIN_PASS || 'twidy2026';

    if (username === adminUser && password === adminPass) {
        res.json({ success: true, token: 'twidy-admin-secure-token' });
    } else {
        res.status(401).json({ success: false, message: 'Username atau Password salah!' });
    }
});

app.post('/api/admin/products', (req, res) => {
    const { category, name, price, description, downloadUrl, image } = req.body;
    if (!category || !name || !price || !downloadUrl) {
        return res.status(400).json({ success: false, message: 'Data produk kurang lengkap!' });
    }

    const digitalProducts = readDigitalDB();
    const cat = category.toLowerCase();
    const subCat = (cat === 'ebook') ? detectEbookSubCategory(name, description || '') : '';

    const newProduct = {
        id: `DIGI-${Date.now()}-${Math.floor(Math.random()*1000)}`,
        category: cat,
        subCategory: subCat,
        name,
        price: parseInt(price),
        description: description || 'Produk digital siap download',
        downloadUrl,
        image: image && image.trim() !== '' ? image : 'https://via.placeholder.com/150?text=TLIBRARY',
        created_at: new Date().toISOString()
    };

    digitalProducts.push(newProduct);
    saveDigitalDB(digitalProducts);
    res.json({ success: true, message: 'Produk digital berhasil ditambahkan!' });
});

app.post('/api/admin/products/bulk', (req, res) => {
    let products = req.body.products;
    
    if (req.body.fileData) {
        try {
            const buffer = Buffer.from(req.body.fileData, 'base64');
            const workbook = XLSX.read(buffer, { type: 'buffer' });
            const sheetName = workbook.SheetNames[0];
            const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
            
            products = rows.map(r => ({
                category: String(r.category || r.Kategori || 'ebook').toLowerCase(),
                name: String(r.name || r.Nama_Produk || 'Produk Tanpa Nama'),
                price: parseInt(r.price || r.Harga || 0),
                description: String(r.description || r.Deskripsi || ''),
                downloadUrl: String(r.downloadUrl || r.Link_Download || '#'),
                image: String(r.image || r.URL_Gambar || '')
            }));
        } catch (e) {
            return res.status(400).json({ success: false, message: 'Gagal membaca format Excel!' });
        }
    }

    if (!Array.isArray(products) || products.length === 0) {
        return res.status(400).json({ success: false, message: 'Data produk kosong atau format tidak valid!' });
    }

    let digitalProducts = readDigitalDB();
    products.forEach(p => {
        const cat = (p.category || 'ebook').toLowerCase();
        const namaProduk = p.name || 'Produk Tanpa Nama';
        const descProduk = p.description || '';
        const subCat = (cat === 'ebook') ? detectEbookSubCategory(namaProduk, descProduk) : '';

        digitalProducts.push({
            id: `DIGI-${Date.now()}-${Math.floor(Math.random()*1000)}`,
            category: cat,
            subCategory: subCat,
            name: namaProduk,
            price: parseInt(p.price || 0),
            description: descProduk,
            downloadUrl: p.downloadUrl || '#',
            image: p.image && p.image.trim() !== '' ? p.image : 'https://via.placeholder.com/150?text=TLIBRARY',
            created_at: new Date().toISOString()
        });
    });

    saveDigitalDB(digitalProducts);
    res.json({ success: true, message: `Berhasil mengimpor ${products.length} produk secara masal!` });
});

app.put('/api/admin/products/:id', (req, res) => {
    const { id } = req.params;
    const { category, name, price, description, downloadUrl, image } = req.body;
    
    let digitalProducts = readDigitalDB();
    const index = digitalProducts.findIndex(p => p.id === id);
    if (index > -1) {
        const cat = category.toLowerCase();
        const subCat = (cat === 'ebook') ? detectEbookSubCategory(name, description || '') : '';

        digitalProducts[index] = {
            ...digitalProducts[index],
            category: cat,
            subCategory: subCat,
            name,
            price: parseInt(price),
            description: description || '',
            downloadUrl,
            image: image && image.trim() !== '' ? image : 'https://via.placeholder.com/150?text=TLIBRARY'
        };
        saveDigitalDB(digitalProducts);
        res.json({ success: true, message: 'Produk berhasil diperbarui!' });
    } else {
        res.status(404).json({ success: false, message: 'Produk tidak ditemukan!' });
    }
});

app.delete('/api/admin/products/:id', (req, res) => {
    const { id } = req.params;
    let digitalProducts = readDigitalDB();
    digitalProducts = digitalProducts.filter(p => p.id !== id);
    saveDigitalDB(digitalProducts);
    res.json({ success: true, message: 'Produk berhasil dihapus!' });
});
// --- END FITUR DIGITAL ---


// --- FUNGSI HELPER MARGIN (1%, Min Rp 200, Max Rp 500) ---
function calculateMargin(hargaAsli) {
    const persentase = Math.round(hargaAsli * 0.01);
    return Math.max(200, Math.min(500, persentase)); // Batas bawah 200, batas atas 500
}

// --- INTEGRASI H2H HYBRID (DIGIFLAZZ + HAYBI) ---
app.get('/api/products', async (req, res) => {
  try {
    const now = Date.now();
    if (!cachedProducts || (now - cacheTimestamp > CACHE_DURATION)) {
      let combinedProducts = [];
      const dfUser = process.env.DIGIFLAZZ_USERNAME;
      const dfKey = process.env.DIGIFLAZZ_API_KEY;
      const hbUser = process.env.HAYBI_USERNAME;
      const hbKey = process.env.HAYBI_API_KEY;

      // 1. FETCH DARI DIGIFLAZZ (Semua kecuali PLN & E-Money)
      if (dfUser && dfKey) {
          try {
              const sign = crypto.createHash('md5').update(dfUser + dfKey + 'depo').digest('hex');
              const dfRes = await axios.post('https://api.digiflazz.com/v1/price-list', { 
                  cmd: 'prepaid', 
                  username: dfUser, 
                  sign: sign 
              }, { headers: { 'Content-Type': 'application/json' } });
              
              if (dfRes.data && Array.isArray(dfRes.data.data)) {
                  const dfMapped = dfRes.data.data.map(p => ({
                      buyer_sku_code: p.buyer_sku_code,
                      product_name: p.product_name,
                      category: p.category,
                      price: parseInt(p.price) + calculateMargin(parseInt(p.price)),
                      buyer_product_status: p.buyer_product_status && p.seller_product_status,
                      brand: p.brand,
                      note: 'Tersedia', // <-- DIPERBAIKI: Hapus p.desc agar tulisan ijo bersih
                      isPasca: false,
                      provider: 'digiflazz' // Label penanda eksekusi webhook nanti
                  })).filter(p => !['E-Money', 'PLN'].includes(p.category) && !p.category.toLowerCase().includes('token')); 
                  // Filter out E-Money & PLN dari Digiflazz biar Haybi yang ambil alih
                  
                  combinedProducts.push(...dfMapped);
              }
          } catch (e) {
              console.error("Digiflazz Fetch Error:", e.message);
          }
      }

      // 2. FETCH DARI HAYBI (Khusus PLN & E-Money saja)
      if (hbUser && hbKey) {
          try {
              const refId = `SYNC_${Date.now()}`;
              const sign = crypto.createHash('md5').update(hbUser + hbKey + refId).digest('hex');
              const hRes = await axios.post('https://haybi.id/api/h2h/produk', { 
                  username: hbUser, 
                  ref_id: refId, 
                  sign: sign, 
                  kategori: "" 
              }, { headers: { 'Content-Type': 'application/json' } });
              
              const rawH = hRes.data?.data || hRes.data;
              if (Array.isArray(rawH)) {
                  const hbMapped = rawH.map(produk => {
                      let hargaDasar = parseInt(produk.harga || produk.price || produk.hargadasar || 0);
                      if (hargaDasar === 0) {
                          for (const k in produk) {
                              const val = parseInt(produk[k]);
                              if (!isNaN(val) && val > 500) { hargaDasar = val; break; }
                          }
                      }

                      const skuCode = produk.kode_produk || produk.kode || produk.buyer_sku_code;
                      const prodName = produk.nama_produk || produk.nama || produk.product_name;
                      let textCheck = (skuCode + " " + prodName).toUpperCase();

                      // Deteksi Kategori Haybi
                      let detectedCategory = 'Umum';
                      let detectedBrand = produk.brand || produk.provider || 'Umum';

                      if (textCheck.includes('DANA') || textCheck.includes('OVO') || textCheck.includes('GOPAY') || textCheck.includes('SHOPEE') || textCheck.includes('LINKAJA') || textCheck.includes('E-MONEY')) {
                          detectedCategory = 'E-Money';
                          if (textCheck.includes('DANA')) detectedBrand = 'DANA';
                          else if (textCheck.includes('OVO')) detectedBrand = 'OVO';
                          else if (textCheck.includes('GOPAY') || textCheck.includes('GO PAY')) detectedBrand = 'GO PAY';
                          else if (textCheck.includes('SHOPEE')) detectedBrand = 'SHOPEE PAY';
                          else if (textCheck.includes('LINKAJA')) detectedBrand = 'LINKAJA';
                      } else if (textCheck.includes('PLN') || textCheck.includes('TOKEN')) {
                          detectedCategory = 'PLN';
                          detectedBrand = 'Token PLN';
                      }

                      return {
                          buyer_sku_code: skuCode,
                          product_name: prodName,
                          category: detectedCategory,
                          price: hargaDasar + calculateMargin(hargaDasar),
                          buyer_product_status: true,
                          brand: detectedBrand,
                          note: 'Tersedia', // <-- DIPERBAIKI: Bersih dan seragam dengan Digiflazz
                          isPasca: false,
                          provider: 'haybi' // Label penanda
                      };
                  }).filter(p => ['E-Money', 'PLN'].includes(p.category)); // AMBIL HANYA E-MONEY & PLN DARI HAYBI

                  combinedProducts.push(...hbMapped);
              }
          } catch (e) {
              console.error("Haybi Fetch Error:", e.message);
          }
      }

      cachedProducts = combinedProducts;
      cacheTimestamp = now;
    }
    res.json({ data: cachedProducts });
  } catch (err) {
    res.status(500).json({ message: err.message });
  }
});

// --- API INQUIRY PASCA BAYAR & TAGIHAN (KEMBALI KE DIGIFLAZZ) ---
app.post('/api/inquiry-pasca', async (req, res) => {
    const { sku, targetId } = req.body;
    const user = process.env.DIGIFLAZZ_USERNAME;
    const key = process.env.DIGIFLAZZ_API_KEY;
    
    if (!user || !key) return res.status(500).json({ success: false, message: 'API Key Digiflazz belum diatur' });
    
    const refId = `INQ-${Date.now()}`; 
    const sign = crypto.createHash('md5').update(user + key + refId).digest('hex');

    try {
        const dfRes = await axios.post('https://api.digiflazz.com/v1/transaction', {
            commands: "inq-pasca",
            username: user,
            buyer_sku_code: sku,
            customer_no: targetId,
            ref_id: refId,
            sign: sign
        });

        const result = dfRes.data?.data;
        if (result && result.status === 'Sukses') {
            const mappedData = {
                status: 'Sukses',
                selling_price: parseInt(result.selling_price) || parseInt(result.price),
                customer_name: result.customer_name || targetId,
                desc: result.desc || {}
            };
            res.json({ success: true, data: mappedData });
        } else {
            res.json({ success: false, message: result?.message || 'Tagihan tidak ditemukan / Gagal cek' });
        }
    } catch (err) {
        res.status(500).json({ success: false, message: err.message });
    }
});
// --- END INTEGRASI API PASCA ---

// --- ENDPOINT TRANSAKSI (DENGAN AUTO-POLLING SINKRONISASI REALTIME KE 2 SERVER) ---
app.get('/api/transactions', async (req, res) => {
    try {
        let db = readDB();
        let needsSave = false;
        
        // AUTO SYNC STATUS: Cek max 5 transaksi terakhir yang masih 'DIPROSES'
        const pendingTrx = db.filter(t => t.status === 'DIPROSES' && !t.is_digital).slice(-5); 
        if (pendingTrx.length > 0) {
            await Promise.all(pendingTrx.map(async (trx) => {
                try {
                    if (trx.provider === 'haybi') {
                        const user = process.env.HAYBI_USERNAME;
                        const key = process.env.HAYBI_API_KEY;
                        if (user && key) {
                            const sign = crypto.createHash('md5').update(user + key + trx.order_id).digest('hex');
                            const checkRes = await axios.post('https://haybi.id/api/h2h/cek-status', {
                                username: user,
                                ref_id: trx.order_id,
                                sign: sign
                            });
                            const result = checkRes.data;
                            if (result && result.status) {
                                const hStatus = result.status.toLowerCase();
                                if (hStatus === 'sukses' || hStatus === 'success') {
                                    trx.status = 'SUKSES';
                                    trx.sn = result.sn || result.pesan || trx.sn;
                                    needsSave = true;
                                } else if (hStatus === 'gagal' || hStatus === 'error') {
                                    trx.status = 'GAGAL';
                                    trx.sn = result.pesan || 'Transaksi Gagal';
                                    needsSave = true;
                                }
                            }
                        }
                    } else {
                        // Default ke Digiflazz
                        const user = process.env.DIGIFLAZZ_USERNAME;
                        const key = process.env.DIGIFLAZZ_API_KEY;
                        if (user && key) {
                            const sign = crypto.createHash('md5').update(user + key + trx.order_id).digest('hex');
                            const checkRes = await axios.post('https://api.digiflazz.com/v1/transaction', {
                                username: user,
                                buyer_sku_code: trx.product_code,
                                customer_no: trx.target_id,
                                ref_id: trx.order_id,
                                sign: sign
                            });
                            const result = checkRes.data?.data;
                            if (result && result.status) {
                                const dStatus = result.status.toLowerCase();
                                if (dStatus === 'sukses') {
                                    trx.status = 'SUKSES';
                                    trx.sn = result.sn || trx.sn;
                                    needsSave = true;
                                } else if (dStatus === 'gagal') {
                                    trx.status = 'GAGAL';
                                    trx.sn = result.message || 'Transaksi Gagal';
                                    needsSave = true;
                                }
                            }
                        }
                    }
                } catch (err) {
                    // Abaikan error jaringan
                }
            }));
            if (needsSave) saveDB(db);
        }
        return res.status(200).json(db.reverse());
    } catch (e) {
        return res.status(500).json({ message: 'Error Database' });
    }
});
// --- END TRANSAKSI ---

app.post('/api/checkout', async (req, res) => {
  try {
    const { targetId, serverId, price, productName, productCode, isDigital, isPasca, downloadUrl, cartItems } = req.body;
    if (!targetId) return res.status(400).json({ message: 'Data kurang lengkap' });

    const orderId = `TLIB-${Date.now()}`;
    const fullTarget = serverId ? `${targetId}${serverId}` : targetId;
    
    let amount = 0;
    let originalProductName = '';
    let itemDetails = [];
    let finalProductCode = productCode || 'DIGI-MULTI';

    if (isDigital && cartItems && Array.isArray(cartItems)) {
        amount = cartItems.reduce((sum, item) => sum + parseInt(item.price), 0);
        originalProductName = `Pembelian ${cartItems.length} Produk Digital`;
        
        itemDetails = cartItems.map((item, index) => ({
            id: `DIGI-${index}`,
            price: parseInt(item.price),
            quantity: 1,
            name: item.name.replace(/[\[\]]/g, '').substring(0, 50),
            merchant_data: fullTarget
        }));
    } else {
        if (!productCode) return res.status(400).json({ message: 'Data kurang lengkap' });
        amount = parseInt(price || 0);
        originalProductName = productName || 'Produk Digital TLIBRARY';
        
        itemDetails = [{
            id: productCode.substring(0, 50),
            price: amount,
            quantity: 1,
            name: originalProductName.replace(/[\[\]]/g, '').substring(0, 50),
            merchant_data: fullTarget
        }];
    }

    // Deteksi provider dari cache yang sudah difilter
    let selectedProvider = 'digiflazz'; 
    if (cachedProducts && !isDigital) {
        const found = cachedProducts.find(p => p.buyer_sku_code === finalProductCode);
        if (found && found.provider) {
            selectedProvider = found.provider;
        }
    }

    const db = readDB();
    db.push({
        order_id: orderId,
        target_id: fullTarget, 
        product_code: finalProductCode,
        product_name: originalProductName,
        amount: amount,
        status: 'UNPAID',
        sn: isDigital ? 'Menunggu Pembayaran (Link akan muncul otomatis setelah lunas)...' : '-',
        is_digital: !!isDigital,
        is_pasca: !!isPasca, 
        download_url: downloadUrl || '',
        cart_items: cartItems || null,
        provider: selectedProvider,
        created_at: new Date().toISOString()
    });
    saveDB(db);

    let parameter = {
      transaction_details: { order_id: orderId, gross_amount: amount },
      item_details: itemDetails,
      customer_details: { first_name: "Pelanggan", last_name: "TLIBRARY", email: isDigital ? targetId : "customer@tlibrary.my.id" }
    };

    let transaction = await snap.createTransaction(parameter);
    res.json({ token: transaction.token, orderId });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

app.post('/api/webhook', async (req, res) => {
  try {
    const notif = req.body;
    if (!notif || !notif.transaction_status) return res.status(200).send("OK");

    const { transaction_status, order_id } = notif;
    let db = readDB();
    let trx = db.find(t => t.order_id === order_id);
    if (!trx) return res.status(200).send("OK");

    if (transaction_status === 'settlement' || transaction_status === 'capture') {
        if (['SUKSES', 'DIPROSES', 'GAGAL'].includes(trx.status)) return res.status(200).send("OK");
        
        if (trx.is_digital) {
            trx.status = 'SUKSES';
            if (trx.cart_items && Array.isArray(trx.cart_items)) {
                trx.sn = trx.cart_items.map(item => `[${item.name}]: ${item.downloadUrl}`).join(' \n ');
            } else {
                trx.sn = `DOWNLOAD LINK: ${trx.download_url}`;
            }
            saveDB(db);

            sendEmailReceipt(trx, trx.target_id);
            return res.status(200).send("OK");
        }

        trx.status = 'DIPROSES';
        saveDB(db);

        // EKSEKUSI TRANSAKSI HYBRID BERDASARKAN PROVIDER YANG TERSIMPAN
        if (trx.provider === 'haybi') {
            const user = process.env.HAYBI_USERNAME;
            const key = process.env.HAYBI_API_KEY;
            if (user && key) {
                const sign = crypto.createHash('md5').update(user + key + order_id).digest('hex');
                try {
                    const haybiRes = await axios.post('https://haybi.id/api/h2h/transaksi', {
                        username: user,
                        ref_id: order_id,
                        sign: sign,
                        produk: trx.product_code,
                        no_tujuan: trx.target_id
                    });
                    const result = haybiRes.data || {};
                    const responseStatus = (result.status || '').toLowerCase();
                    
                    if (responseStatus === 'sukses' || responseStatus === 'success') {
                        trx.status = 'SUKSES';
                    } else if (responseStatus === 'gagal' || responseStatus === 'error') {
                        trx.status = 'GAGAL';
                    } else {
                        trx.status = 'DIPROSES';
                    }
                    trx.sn = result.sn || result.pesan || result.message || 'Diproses (Menunggu Pembaruan)';
                    saveDB(db);
                } catch (err) {
                    console.error("Haybi Execution Error:", err.message);
                }
            }
        } else {
            // Default Eksekusi ke Digiflazz (Game, Pulsa, Stream, dsb)
            const user = process.env.DIGIFLAZZ_USERNAME;
            const key = process.env.DIGIFLAZZ_API_KEY;
            if (user && key) {
                const sign = crypto.createHash('md5').update(user + key + order_id).digest('hex');
                try {
                    const dfRes = await axios.post('https://api.digiflazz.com/v1/transaction', {
                        username: user,
                        buyer_sku_code: trx.product_code,
                        customer_no: trx.target_id,
                        ref_id: order_id,
                        sign: sign
                    });
                    const result = dfRes.data?.data || {};
                    const responseStatus = (result.status || '').toLowerCase();
                    
                    if (responseStatus === 'sukses') {
                        trx.status = 'SUKSES';
                    } else if (responseStatus === 'gagal') {
                        trx.status = 'GAGAL';
                    } else {
                        trx.status = 'DIPROSES';
                    }
                    trx.sn = result.sn || result.message || 'Diproses (Menunggu Pembaruan)';
                    saveDB(db);
                } catch (err) {
                    console.error("Digiflazz Execution Error:", err.message);
                }
            }
        }
    } else if (['expire', 'cancel', 'deny'].includes(transaction_status)) {
        trx.status = 'GAGAL';
        saveDB(db);
    }
    return res.status(200).send("OK");
  } catch (e) {
    return res.status(500).send("Error");
  }
});

app.post('/api/haybi-webhook', (req, res) => {
  try {
    const payload = req.body;
    if (!payload || !payload.ref_id) {
        return res.status(200).send("OK");
    }

    const { ref_id, status, sn, pesan } = payload;
    let db = readDB();
    let trx = db.find(t => t.order_id === ref_id);
    if (!trx) return res.status(200).send("OK");

    const currentStatus = (status || '').toLowerCase();

    if (currentStatus === 'sukses' || currentStatus === 'success') {
        trx.status = 'SUKSES';
    } else if (currentStatus === 'gagal' || currentStatus === 'error') {
        trx.status = 'GAGAL';
    } else {
        trx.status = 'DIPROSES';
    }
    
    const currentSn = sn || pesan;
    if (currentSn && currentSn.trim() !== '') {
        trx.sn = currentSn;
    } else if (trx.status === 'SUKSES') {
        trx.sn = 'Transaksi Berhasil';
    }
    
    saveDB(db);
    return res.status(200).send("OK");
  } catch (error) {
    console.error("Haybi Webhook Error:", error.message);
    return res.status(500).send("Error");
  }
});

app.post('/api/digiflazz-webhook', (req, res) => {
    try {
      const payload = req.body?.data;
      if (!payload || !payload.ref_id) return res.status(200).send("OK");
  
      let db = readDB();
      let trx = db.find(t => t.order_id === payload.ref_id);
      if (!trx) return res.status(200).send("OK");
  
      const currentStatus = (payload.status || '').toLowerCase();
      if (currentStatus === 'sukses') {
          trx.status = 'SUKSES';
      } else if (currentStatus === 'gagal') {
          trx.status = 'GAGAL';
      }
      
      if (payload.sn) trx.sn = payload.sn;
      if (payload.message && trx.status === 'GAGAL') trx.sn = payload.message;
      
      saveDB(db);
      return res.status(200).send("OK");
    } catch (error) {
      console.error("Digiflazz Webhook Error:", error.message);
      return res.status(500).send("Error");
    }
});

// --- FITUR SEO: ROBOTS.TXT & SITEMAP.XML ---
app.get('/robots.txt', (req, res) => {
    res.type('text/plain');
    res.send(`User-agent: *
Allow: /
Sitemap: https://tlibrary.my.id/sitemap.xml`);
});

app.get('/sitemap.xml', (req, res) => {
    const digitalProducts = readDigitalDB();
    const baseUrl = 'https://tlibrary.my.id';
    const generateSlug = (text) => text.toLowerCase().replace(/[^\w ]+/g, '').replace(/ +/g, '-');

    let xml = `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
    <url>
        <loc>${baseUrl}/</loc>
        <changefreq>daily</changefreq>
        <priority>1.0</priority>
    </url>
    <url>
        <loc>${baseUrl}/topup</loc>
        <changefreq>daily</changefreq>
        <priority>0.9</priority>
    </url>
    <url>
        <loc>${baseUrl}/riwayat</loc>
        <changefreq>weekly</changefreq>
        <priority>0.8</priority>
    </url>
    <url>
        <loc>${baseUrl}/promo</loc>
        <changefreq>weekly</changefreq>
        <priority>0.8</priority>
    </url>
    <url>
        <loc>${baseUrl}/panduan</loc>
        <changefreq>monthly</changefreq>
        <priority>0.7</priority>
    </url>`;

    if (digitalProducts && digitalProducts.length > 0) {
        digitalProducts.forEach(p => {
            if (p.name) {
                const slug = generateSlug(p.name);
                xml += `
    <url>
        <loc>${baseUrl}/${slug}</loc>
        <changefreq>monthly</changefreq>
        <priority>0.6</priority>
    </url>`;
            }
        });
    }

    xml += `\n</urlset>`;
    res.header('Content-Type', 'application/xml');
    res.send(xml);
});
// --- END FITUR SEO ---

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server TLIBRARY berjalan di port ${PORT}`));
