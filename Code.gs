// ================================================
// TOKO BAGIR v2 - Code.gs (Google Apps Script)
// ✅ CORS diurus Cloudflare Worker — tidak perlu setHeader di sini
// ================================================
const PRODUCTS_SHEET  = 'Produk';
const ORDERS_SHEET    = 'Pesanan';
const NOTIF_SHEET     = 'Notifikasi';
const SETTINGS_SHEET  = 'Settings';
const SHIPPING_SHEET  = 'Ongkir';
const ADMINS_SHEET    = 'Admins';
const SESSIONS_SHEET  = 'Sessions';

const SESSION_TTL_HOURS = 12;
const DRIVE_FOLDER_ID   = '1KLZTqiV4Dcye9UELpNsVNb7GU_dMlBKK';

// ── SATU-SATUNYA fungsi response — tidak ada setHeader ────────
function json(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET ──────────────────────────────────────────────────────
function doGet(e) {
  const a = e.parameter.action;
  try {
    if (a === 'getProducts')      return json(getProducts());
    if (a === 'getOrders')        return json(getOrders());
    if (a === 'getNotifications') return json(getNotifications());
    if (a === 'getSettings')      return json(getSettings());
    if (a === 'getShippingRates') return json(getShippingRates());
    if (a === 'getSalesData')     return json(getSalesData());
    if (a === 'verifySession')    return json(verifySession(e.parameter.token));
    return json({ success: false, message: 'Unknown action' });
  } catch (err) {
    return json({ success: false, message: err.toString() });
  }
}

// ── POST ─────────────────────────────────────────────────────
function doPost(e) {
  let d;
  try {
    d = JSON.parse(e.postData.contents);
  } catch (parseErr) {
    return json({ success: false, message: 'Payload tidak valid: ' + parseErr.toString() });
  }
  const a = d.action;
  try {
    const publicActions = ['addOrder','markReadNotif','uploadImage','adminLogin','adminLogout'];
    if (!publicActions.includes(a)) {
      const check = verifySession(d.token);
      if (!check.valid) return json({ success: false, message: 'Sesi admin tidak valid, silakan login kembali', authError: true });
    }
    if (a === 'addProduct')          return json(addProduct(d));
    if (a === 'updateProduct')       return json(updateProduct(d));
    if (a === 'deleteProduct')       return json(deleteProduct(d));
    if (a === 'addOrder')            return json(addOrder(d));
    if (a === 'updateOrderStatus')   return json(updateOrderStatus(d));
    if (a === 'updateResi')          return json(updateResi(d));
    if (a === 'deleteOrder')         return json(deleteOrder(d));
    if (a === 'addNotification')     return json(addNotification(d));
    if (a === 'markReadNotif')       return json(markReadNotif(d));
    if (a === 'saveSettings')        return json(saveSettings(d));
    if (a === 'saveShippingRate')    return json(saveShippingRate(d));
    if (a === 'deleteShippingRate')  return json(deleteShippingRate(d));
    if (a === 'uploadImage')         return json(uploadImage(d.base64, d.fileName, d.mimeType));
    if (a === 'adminLogin')          return json(adminLogin(d));
    if (a === 'adminLogout')         return json(adminLogout(d));
    if (a === 'adminChangePassword') return json(adminChangePassword(d));
    return json({ success: false, message: 'Unknown action' });
  } catch (err) {
    return json({ success: false, message: err.toString() });
  }
}

// ── INIT SHEETS ──────────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  createSheet(ss, PRODUCTS_SHEET,
    ['id','name','cat','emoji','price','weight','stock','desc','image','price500','price1kg','price1bal']);
  createSheet(ss, ORDERS_SHEET,
    ['id','orderDate','deliveryDate','name','phone','address','items','total',
     'delivery','courier','service','shippingCost','city','province',
     'payment','paymentProof','dpAmount','note','status','resi','createdAt','updatedAt']);
  createSheet(ss, NOTIF_SHEET,
    ['id','orderId','customerName','message','type','isRead','createdAt']);
  createSheet(ss, SETTINGS_SHEET, ['key','value']);
  createSheet(ss, SHIPPING_SHEET, ['id','province','city','courier','service','cost','etd']);
  createSheet(ss, ADMINS_SHEET,   ['username','passwordHash','salt','createdAt']);
  createSheet(ss, SESSIONS_SHEET, ['token','username','createdAt','expiresAt']);

  const as = ss.getSheetByName(ADMINS_SHEET);
  if (as.getLastRow() <= 1) {
    const salt = Utilities.getUuid();
    as.appendRow(['admin', hashPassword('bagir2025', salt), salt, new Date().toISOString()]);
  }

  const ps = ss.getSheetByName(PRODUCTS_SHEET);
  if (ps.getLastRow() <= 1) {
    const sample = [
      [1,'Nastar Nanas','classic','🍍',65000,'250g',15,'Nastar lembut isi selai nanas pilihan.','',130000,250000,600000],
      [2,'Putri Salju','classic','❄️',60000,'250g',12,'Lumer di mulut, taburan gula halus.','',120000,230000,560000],
      [3,'Kastengel','savory','🧀',70000,'250g',8,'Gurih keju edam asli pilihan.','',140000,270000,650000],
      [4,'Lidah Kucing','classic','🍪',55000,'250g',20,'Renyah tipis dan manis sempurna.','',110000,210000,510000]
    ];
    ps.getRange(2,1,sample.length,12).setValues(sample);
  }

  const sh = ss.getSheetByName(SHIPPING_SHEET);
  if (sh.getLastRow() <= 1) {
    sh.getRange(2,1,3,7).setValues([
      [1,'DKI Jakarta','Jakarta','JNE','REG',15000,'1-2 hari'],
      [2,'Jawa Barat','Bandung','JNE','REG',20000,'1-2 hari'],
      [3,'Jawa Barat','Bogor','JNE','REG',18000,'1-2 hari']
    ]);
  }
  return { success: true, message: 'Semua sheet berhasil dibuat!' };
}

function createSheet(ss, name, headers) {
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(headers);
    s.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1A6CC8').setFontColor('#FFFFFF');
  }
  return s;
}

// ── UPLOAD GAMBAR ─────────────────────────────────────────────
function uploadImage(base64, fileName, mimeType) {
  if (!base64) return { success: false, message: 'Data gambar kosong' };
  try {
    const folder  = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const decoded = Utilities.base64Decode(base64);
    const blob    = Utilities.newBlob(decoded, mimeType || 'image/jpeg', fileName || ('img_'+Date.now()+'.jpg'));
    const file    = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: 'https://drive.google.com/uc?export=view&id='+file.getId(), fileId: file.getId() };
  } catch (err) {
    return { success: false, message: 'Gagal upload: '+err.toString() };
  }
}

// ── PRODUK ───────────────────────────────────────────────────
function getProducts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS_SHEET);
  if (!sheet) return { success: false, data: [] };
  const rows = sheet.getDataRange().getValues();
  const h = rows[0];
  return { success: true, data: rows.slice(1).map(r => {
    const o = {}; h.forEach((k,i) => o[k]=r[i]);
    o.price    = Number(o.price)    || 0;
    o.stock    = Number(o.stock)    || 0;
    o.id       = Number(o.id)       || 0;
    o.price500 = Number(o.price500) || 0;
    o.price1kg = Number(o.price1kg) || 0;
    o.price1bal= Number(o.price1bal)|| 0;
    return o;
  }).filter(p => p.id && p.name) };
}

function addProduct(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  const ids   = rows.slice(1).map(r=>Number(r[0])).filter(Boolean);
  const newId = ids.length ? Math.max(...ids)+1 : 1;
  sheet.appendRow([newId,d.name,d.cat,d.emoji||'',Number(d.price),d.weight||'250g',
    Number(d.stock),d.desc||'',d.image||'',
    Number(d.price500)||0,Number(d.price1kg)||0,Number(d.price1bal)||0]);
  return { success: true, id: newId };
}

function updateProduct(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (Number(rows[i][0])===Number(d.id)) {
      sheet.getRange(i+1,1,1,12).setValues([[
        Number(d.id),d.name,d.cat,d.emoji||'',Number(d.price),d.weight||'250g',
        Number(d.stock),d.desc||'',d.image||'',
        Number(d.price500)||0,Number(d.price1kg)||0,Number(d.price1bal)||0
      ]]);
      return { success: true };
    }
  }
  return { success: false, message: 'Produk tidak ditemukan' };
}

function deleteProduct(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (Number(rows[i][0])===Number(d.id)) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: false };
}

function reduceStock(items) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PRODUCTS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  items.forEach(item => {
    const name = item.name.replace(/\s*\(.*?\)\s*/g,'').trim(); // hapus "(250g)" dll dari nama
    for (let i=1;i<rows.length;i++) {
      if (String(rows[i][1]).trim()===name) {
        const newStock = Math.max(0, Number(rows[i][6])-item.qty);
        sheet.getRange(i+1,7).setValue(newStock);
        rows[i][6] = newStock;
        if (newStock<=5) addNotification({
          orderId:'', customerName:'Sistem',
          message: `Stok ${rows[i][1]} ${newStock===0?'HABIS ❗':'hampir habis (sisa '+newStock+') ⚠️'}`,
          type:'low_stock'
        });
        break;
      }
    }
  });
}

// ── PESANAN ──────────────────────────────────────────────────
function getOrders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET);
  if (!sheet||sheet.getLastRow()<=1) return { success: true, data: [] };
  const rows = sheet.getDataRange().getValues();
  const h = rows[0];
  return { success: true, data: rows.slice(1).map(r=>{
    const o={}; h.forEach((k,i)=>o[k]=r[i]);
    o.total=Number(o.total); o.shippingCost=Number(o.shippingCost)||0;
    try{o.items=JSON.parse(o.items);}catch(e){o.items=[];}
    return o;
  }).filter(o=>o.id) };
}

function addOrder(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET);
  const now   = new Date().toISOString();
  const newId = 'BG'+String(sheet.getLastRow()).padStart(3,'0');
  sheet.appendRow([
    newId, d.orderDate||now.slice(0,10), d.deliveryDate||'',
    d.name, d.phone, d.address||'',
    JSON.stringify(d.items||[]), Number(d.total),
    d.delivery||'pickup', d.courier||'', d.service||'',
    Number(d.shippingCost)||0, d.city||'', d.province||'',
    d.payment||'transfer', d.paymentProof||'', Number(d.dpAmount)||0,
    d.note||'', 'baru', '', now, now
  ]);
  reduceStock(d.items||[]);
  addNotification({
    orderId: newId, customerName: d.name,
    message: `🛒 Pesanan baru dari ${d.name} senilai Rp ${Number(d.total).toLocaleString('id-ID')}`,
    type: 'order'
  });
  if (d.paymentProof) addNotification({
    orderId: newId, customerName: d.name,
    message: `💰 Bukti bayar dari ${d.name} untuk pesanan ${newId} masuk`,
    type: 'payment'
  });
  return { success: true, id: newId };
}

function updateOrderStatus(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  const h     = rows[0];
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][0])===String(d.id)) {
      sheet.getRange(i+1,h.indexOf('status')+1).setValue(d.status);
      sheet.getRange(i+1,h.indexOf('updatedAt')+1).setValue(new Date().toISOString());
      return { success: true };
    }
  }
  return { success: false };
}

function updateResi(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  const h     = rows[0];
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][0])===String(d.id)) {
      sheet.getRange(i+1,h.indexOf('resi')+1).setValue(d.resi||'');
      sheet.getRange(i+1,h.indexOf('updatedAt')+1).setValue(new Date().toISOString());
      return { success: true };
    }
  }
  return { success: false };
}

function deleteOrder(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][0])===String(d.id)) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: false };
}

// ── NOTIFIKASI ───────────────────────────────────────────────
function getNotifications() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOTIF_SHEET);
  if (!sheet||sheet.getLastRow()<=1) return { success: true, data: [] };
  const rows = sheet.getDataRange().getValues();
  const h    = rows[0];
  return { success: true, data: rows.slice(1).map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);return o;}).filter(n=>n.id).reverse() };
}
function addNotification(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOTIF_SHEET);
  const rows  = sheet.getDataRange().getValues();
  const ids   = rows.slice(1).map(r=>Number(r[0])).filter(Boolean);
  const newId = ids.length ? Math.max(...ids)+1 : 1;
  sheet.appendRow([newId,d.orderId,d.customerName,d.message,d.type||'info',false,new Date().toISOString()]);
  return { success: true };
}
function markReadNotif(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(NOTIF_SHEET);
  const rows  = sheet.getDataRange().getValues();
  if (d.markAll) { for(let i=1;i<rows.length;i++) sheet.getRange(i+1,6).setValue(true); return {success:true}; }
  for (let i=1;i<rows.length;i++) {
    if (Number(rows[i][0])===Number(d.id)) { sheet.getRange(i+1,6).setValue(true); return {success:true}; }
  }
  return { success: false };
}

// ── SETTINGS ─────────────────────────────────────────────────
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  if (!sheet||sheet.getLastRow()<=1) return { success: true, data: {} };
  const rows = sheet.getDataRange().getValues();
  const data = {};
  rows.slice(1).forEach(r=>{ if(r[0]) data[r[0]]=r[1]; });
  return { success: true, data };
}
function saveSettings(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  const keys  = ['logoImage','qrisImage','tokoName','tokoTagline','tokoWa','tokoInstagram','tokoAlamat','headerBg','heroBgImage'];
  keys.forEach(key => {
    if (d[key]===undefined) return;
    let found = false;
    for (let i=1;i<rows.length;i++) {
      if (rows[i][0]===key) { sheet.getRange(i+1,2).setValue(d[key]); found=true; break; }
    }
    if (!found) sheet.appendRow([key, d[key]]);
  });
  return { success: true };
}

// ── ONGKIR ───────────────────────────────────────────────────
function getShippingRates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHIPPING_SHEET);
  if (!sheet||sheet.getLastRow()<=1) return { success: true, data: [] };
  const rows = sheet.getDataRange().getValues();
  const h    = rows[0];
  return { success: true, data: rows.slice(1).map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);o.cost=Number(o.cost);o.id=Number(o.id);return o;}).filter(r=>r.id) };
}
function saveShippingRate(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHIPPING_SHEET);
  const rows  = sheet.getDataRange().getValues();
  if (d.id) {
    for (let i=1;i<rows.length;i++) {
      if (Number(rows[i][0])===Number(d.id)) {
        sheet.getRange(i+1,1,1,7).setValues([[Number(d.id),d.province,d.city,d.courier,d.service,Number(d.cost),d.etd]]);
        return { success: true };
      }
    }
  }
  const ids   = rows.slice(1).map(r=>Number(r[0])).filter(Boolean);
  const newId = ids.length ? Math.max(...ids)+1 : 1;
  sheet.appendRow([newId,d.province,d.city,d.courier,d.service,Number(d.cost),d.etd]);
  return { success: true, id: newId };
}
function deleteShippingRate(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHIPPING_SHEET);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (Number(rows[i][0])===Number(d.id)) { sheet.deleteRow(i+1); return { success: true }; }
  }
  return { success: false };
}

// ── SALES DATA ────────────────────────────────────────────────
function getSalesData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET);
  if (!sheet||sheet.getLastRow()<=1) return { success: true, data:{ totalRevenue:0, totalOrders:0, monthly:[], topProducts:[], byStatus:{}, activeOrders:0 } };
  const rows   = sheet.getDataRange().getValues();
  const h      = rows[0];
  const orders = rows.slice(1).map(r=>{const o={};h.forEach((k,i)=>o[k]=r[i]);o.total=Number(o.total);try{o.items=JSON.parse(o.items);}catch(e){o.items=[];}return o;}).filter(o=>o.id);
  const active = orders.filter(o=>o.status!=='batal');
  const totalRevenue = active.reduce((s,o)=>s+o.total,0);
  const monthlyMap={};
  active.forEach(o=>{const d=o.orderDate?o.orderDate.toString().slice(0,7):'';if(!d)return;monthlyMap[d]=(monthlyMap[d]||0)+o.total;});
  const monthly     = Object.entries(monthlyMap).sort().slice(-12).map(([m,v])=>({month:m,revenue:v}));
  const prodMap={};
  active.forEach(o=>{if(!Array.isArray(o.items))return;o.items.forEach(item=>{prodMap[item.name]=(prodMap[item.name]||0)+item.qty;});});
  const topProducts = Object.entries(prodMap).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([name,qty])=>({name,qty}));
  const byStatus    = {baru:0,proses:0,selesai:0,batal:0};
  orders.forEach(o=>{if(byStatus[o.status]!==undefined)byStatus[o.status]++;});
  return { success: true, data:{ totalRevenue, totalOrders:orders.length, monthly, topProducts, byStatus, activeOrders:active.length } };
}

// ── AUTENTIKASI ADMIN ─────────────────────────────────────────
function hashPassword(password, salt) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password+':'+salt);
  return raw.map(b=>(b<0?b+256:b).toString(16).padStart(2,'0')).join('');
}

function adminLogin(d) {
  const username = (d.username||'').trim().toLowerCase();
  const password = d.password||'';
  if (!username||!password) return { success:false, message:'Username dan password wajib diisi' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMINS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][0]).toLowerCase()===username) {
      if (hashPassword(password,rows[i][2])===rows[i][1]) {
        const token   = Utilities.getUuid();
        const now     = new Date();
        const expires = new Date(now.getTime()+SESSION_TTL_HOURS*3600*1000);
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET)
          .appendRow([token,username,now.toISOString(),expires.toISOString()]);
        return { success:true, token, username, expiresAt:expires.toISOString() };
      }
      return { success:false, message:'Username atau password salah' };
    }
  }
  return { success:false, message:'Username atau password salah' };
}

function adminLogout(d) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][0]===d.token) { sheet.deleteRow(i+1); return { success:true }; }
  }
  return { success:true };
}

function verifySession(token) {
  if (!token) return { success:false, valid:false };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SESSIONS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  const now   = new Date();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][0]===token) {
      const expires = new Date(rows[i][3]);
      if (expires>now) return { success:true, valid:true, username:rows[i][1] };
      sheet.deleteRow(i+1);
      return { success:true, valid:false, message:'Sesi sudah berakhir, silakan login kembali' };
    }
  }
  return { success:true, valid:false };
}

function adminChangePassword(d) {
  const check = verifySession(d.token);
  if (!check.valid) return { success:false, message:'Sesi tidak valid, silakan login kembali' };
  if (!d.newPassword||d.newPassword.length<6) return { success:false, message:'Password baru minimal 6 karakter' };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMINS_SHEET);
  const rows  = sheet.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (String(rows[i][0]).toLowerCase()===check.username) {
      const salt = Utilities.getUuid();
      sheet.getRange(i+1,2).setValue(hashPassword(d.newPassword,salt));
      sheet.getRange(i+1,3).setValue(salt);
      return { success:true };
    }
  }
  return { success:false, message:'Akun tidak ditemukan' };
}