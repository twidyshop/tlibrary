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


// --- INTEGRASI API HAYBI (DENGAN LENGKAP GAME BRAND DETECTOR & 1% MARGIN) ---
app.get('/api/products', async (req, res) => {
  const user = process.env.HAYBI_USERNAME;
  const key = process.env.HAYBI_API_KEY;
  if (!user || !key) return res.status(500).json({ message: 'API Key Haybi belum diset' });

  try {
    const now = Date.now();
    if (!cachedProducts || (now - cacheTimestamp > CACHE_DURATION)) {
      
      const refId = `SYNC_${Date.now()}`;
      const sign = crypto.createHash('md5').update(user + key + refId).digest('hex');
      
      const response = await axios.post('https://haybi.id/api/h2h/produk', {
        username: user,
        ref_id: refId,
        sign: sign,
        kategori: ""
      }, {
        headers: { 'Content-Type': 'application/json' }
      });

      const raw = response.data;
      let targetData = [];
      
      if (raw.data && Array.isArray(raw.data)) {
        targetData = raw.data;
      } else if (Array.isArray(raw)) {
        targetData = raw;
      } else {
        return res.status(400).json({ message: 'Gagal ambil data', error: raw });
      }

      // MAPPER DENGAN SMART BRAND & GAME DETECTOR
      cachedProducts = targetData.map(produk => {
          let hargaDasar = 0;
          const possiblePriceKeys = ['hargareseller', 'hargadasar', 'hargamember', 'hargajual', 'harga', 'price', 'harga_dasar', 'harga_jual', 'base_price', 'amount', 'nominal', 'harian'];
          
          for (const key of possiblePriceKeys) {
              if (produk[key] !== undefined && !isNaN(parseInt(produk[key])) && parseInt(produk[key]) > 0) {
                  hargaDasar = parseInt(produk[key]);
                  break;
              }
          }

          if (hargaDasar === 0) {
              for (const k in produk) {
                  const val = parseInt(produk[k]);
                  if (!isNaN(val) && val > 500) {
                      hargaDasar = val;
                      break;
                  }
              }
          }

          // Margin profit 1% dengan batas minimum Rp200
          const margin = Math.max(200, Math.round(hargaDasar * 0.01));

          const skuCode = produk.kode_produk || produk.kode || produk.buyer_sku_code || produk.sku || produk.product_code || 'UNKNOWN';
          const prodName = produk.nama_produk || produk.nama || produk.product_name || produk.title || produk.name || 'Produk Haybi';
          
          // DETEKSI OTOMATIS BRAND & GAME
          let detectedBrand = produk.brand || produk.kategori || produk.provider || 'Umum';
          const textCheck = (skuCode + " " + prodName + " " + (produk.kategori || '')).toUpperCase();

          // Pulsa & Operator
          if (textCheck.startsWith('TS') || textCheck.includes('TELKOMSEL')) detectedBrand = 'Telkomsel';
          else if (textCheck.startsWith('IS') || textCheck.includes('INDOSAT') || textCheck.includes('IM3')) detectedBrand = 'Indosat';
          else if (textCheck.startsWith('AX') || textCheck.includes('AXIS')) detectedBrand = 'Axis';
          else if (textCheck.startsWith('SM') || textCheck.includes('SF') || textCheck.includes('SMART')) detectedBrand = 'Smartfren';
          else if (textCheck.startsWith('TR') || textCheck.includes('TRI')) detectedBrand = 'Tri';
          else if (textCheck.startsWith('XL') || textCheck.includes('XL')) detectedBrand = 'XL';
          else if (textCheck.startsWith('BY') || textCheck.includes('BY.U')) detectedBrand = 'by.U';
          // E-Money
          else if (textCheck.includes('DANA')) detectedBrand = 'DANA';
          else if (textCheck.includes('OVO')) detectedBrand = 'OVO';
          else if (textCheck.includes('GOPAY')) detectedBrand = 'GO PAY';
          else if (textCheck.includes('SHOPEE') || textCheck.includes('SPAY')) detectedBrand = 'SHOPEE PAY';
          else if (textCheck.includes('LINK') || textCheck.includes('LINKAJA')) detectedBrand = 'LINKAJA';
          // PLN
          else if (textCheck.includes('PLN') || textCheck.includes('TOKEN')) detectedBrand = 'Token PLN';
          // Games
          else if (textCheck.includes('MOBILE LEGENDS') || textCheck.includes('MLBB') || textCheck.startsWith('ML')) detectedBrand = 'Mobile Legends';
          else if (textCheck.includes('FREE FIRE') || textCheck.includes('FF')) detectedBrand = 'Free Fire';
          else if (textCheck.includes('PUBG')) detectedBrand = 'PUBG Mobile';
          else if (textCheck.includes('DOMINO') || textCheck.includes('HIGGS')) detectedBrand = 'Higgs Domino';
          else if (textCheck.includes('GENSHIN')) detectedBrand = 'Genshin Impact';
          else if (textCheck.includes('VALORANT')) detectedBrand = 'Valorant';
          else if (textCheck.includes('POINT BLANK') || textCheck.includes('PB')) detectedBrand = 'Point Blank';
          else if (textCheck.includes('STEAM') || textCheck.includes('HAGO') || textCheck.includes('ROBLOX')) detectedBrand = 'Voucher Game';

          return {
              buyer_sku_code: skuCode,
              product_name: prodName,
              price: hargaDasar > 0 ? (hargaDasar + margin) : 1000,
              buyer_product_status: true,
              brand: detectedBrand, // Brand sudah disesuaikan dengan frontend
              note: produk.keterangan || produk.desc || 'Tersedia',
              isPasca: false
          };
      });
      
      cacheTimestamp = now;
    }
    res.json({ data: cachedProducts });
  } catch (err) {
    console.error("Haybi Error:", err.response?.data || err.message);
    res.status(500).json({ message: err.message, detail: err.response?.data });
  }
});

app.post('/api/inquiry-pasca', async (req, res) => {
    const { sku, targetId } = req.body;
    const user = process.env.HAYBI_USERNAME;
    const key = process.env.HAYBI_API_KEY;
    
    if (!user || !key) return res.status(500).json({ success: false, message: 'API Key Haybi belum diatur' });
    
    const refId = `INQ-${Date.now()}`; 
    const sign = crypto.createHash('md5').update(user + key + refId).digest('hex');

    try {
        const haybiRes = await axios.post('https://haybi.id/api/h2h/transaksi', {
            username: user,
            ref_id: refId,
            sign: sign,
            produk: sku,
            no_tujuan: targetId
        });

        const result = haybiRes.data;
        if (result && (result.status === 'sukses' || result.status === 'pending') && result.tagihan) {
            const mappedData = {
                status: 'Sukses',
                selling_price: parseInt(result.tagihan) || parseInt(result.harga),
                customer_name: result.nama_pelanggan || result.nama || targetId,
                desc: { detail: [{ periode: result.periode || 'Bulan Berjalan' }] }
            };
            res.json({ success: true, data: mappedData });
        } else {
            res.json({ success: false, message: result?.pesan || result?.message || 'Tagihan tidak ditemukan / Gagal cek' });
        }
    } catch (err) {
        res.status(500).json({ success: false, message: err.response?.data?.pesan || err.message });
    }
});
// --- END INTEGRASI API HAYBI ---

app.get('/api/transactions', (req, res) => {
    try {
        return res.status(200).json(readDB().reverse());
    } catch (e) {
        return res.status(500).json({ message: 'Error Database' });
    }
});

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

        const user = process.env.HAYBI_USERNAME;
        const key = process.env.HAYBI_API_KEY;
        if (user && key) {
            const sign = crypto.createHash('md5').update(user + key + order_id).digest('hex');
            
            let payloadHaybi = {
                username: user,
                ref_id: order_id,
                sign: sign,
                produk: trx.product_code,
                no_tujuan: trx.target_id
            };

            try {
                const haybiRes = await axios.post('https://haybi.id/api/h2h/transaksi', payloadHaybi);
                const result = haybiRes.data || {};
                
                const responseStatus = (result.status || '').toLowerCase();
                if (responseStatus === 'sukses' || responseStatus === 'success') {
                    trx.status = 'SUKSES';
                } else if (responseStatus === 'gagal' || responseStatus === 'error') {
                    trx.status = 'GAGAL';
                } else {
                    trx.status = 'DIPROSES';
                }
                
                const resultSn = result.sn || result.pesan || result.message;
                trx.sn = (resultSn && resultSn.trim() !== '') ? resultSn : 'Diproses (Menunggu Pembaruan)';
                saveDB(db);
            } catch (err) {
                console.error("Haybi Execution Error:", err.message);
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

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server TLIBRARY berjalan di port ${PORT}`));
