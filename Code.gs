const PETUGAS = {
  ANGGA: { role: "USER" },
  HENDRA: { role: "USER" },
  AGUS: { role: "USER" },
  SUTIKNO: { role: "USER" },
  PURCHASING: {
    role: "ADMIN",
    password: "xxxxxxx" // ganti sesuai keinginan
  }
};

function doGet(e) {

  const template = HtmlService.createTemplateFromFile('Index');
  template.prefillCode = e.parameter.code || '';

  return template
    .evaluate()
    .setTitle('Stok Maintenance')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME); // 🔥 TAMBAHKAN INI

}

/* ============================= */
/* AMBIL DATA BARANG */
/* ============================= */
function getItemByCode(code) {

  const clean = (code || '')
    .toString()
    .trim()
    .toUpperCase();

  if (!clean) return { ok: false, message: 'Kode kosong' };

  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('STOK BARANG');

  if (!sh) return { ok: false, message: 'Sheet tidak ditemukan' };

  const finder = sh
    .getRange("A2:A")
    .createTextFinder(clean)
    .matchEntireCell(true)
    .findNext();

  if (!finder) return { ok: false, message: 'Barang tidak ditemukan' };

  const row = finder.getRow();
  const data = sh.getRange(row, 1, 1, 7).getValues()[0];

  return {
    ok: true,
    code: data[0],
    name: data[1],
    unit: data[2],
    stock: Number(data[6]) || 0
  };
}

/* ============================= */
/* SIMPAN TRANSAKSI */
/* ============================= */
function submitTransaction(payload) {

  const lock = LockService.getScriptLock();
  lock.waitLock(5000);

  try {

    const petugas = (payload.petugas || '')
      .toString()
      .trim()
      .toUpperCase();

    const code = (payload.code || '')
      .toString()
      .trim()
      .toUpperCase();

    let type = (payload.type || '')
      .toString()
      .trim()
      .toUpperCase();

    const qty = Number(payload.qty);
    const password = (payload.password || '').toString();

    if (!PETUGAS[petugas]) {
      return { ok: false, message: 'Petugas tidak dikenal' };
    }

    const role = PETUGAS[petugas].role;

    if (role === "ADMIN") {

      if (!password) {
        return { ok: false, message: 'Password wajib diisi' };
      }

      if (password !== PETUGAS[petugas].password) {
        return { ok: false, message: 'Password Purchasing salah' };
      }

    } else {
      type = "KELUAR";
    }

    if (!code) return { ok: false, message: 'Kode kosong' };

    if (isNaN(qty)) {
      return { ok: false, message: 'Qty harus angka' };
    }

    if (qty === 0) {
      return { ok: false, message: 'Qty tidak boleh 0' };
    }

    const item = getItemByCode(code);
    if (!item.ok) return item;

    if (type === 'KELUAR') {

      const stokBaru = item.stock - qty;

      if (stokBaru < 0) {
        return {
          ok: false,
          message: `STOK TIDAK CUKUP. STOK: ${item.stock}`
        };
      }
    }

    let newStock = item.stock;

    if (type === "MASUK") {
      newStock += qty;
    } else {
      newStock -= qty;
    }

    const shT = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName('TRANSAKSI');

shT.appendRow([
  new Date(),
  item.code,
  item.name,
  type,
  qty,
  petugas,
  payload.catatan || '',
  payload.pos || '' // 🔥 TAMBAH INI
]);


    return {
      ok: true,
      name: item.name,
      type: type,
      qty: qty,
      stock: newStock
    };

  } finally {
    lock.releaseLock(); // 🔥 SELALU dilepas walaupun error
  }
}

function getMasterList() {

  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('STOK BARANG');

  if (!sh) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const values = sh
    .getRange(2, 1, lastRow - 1, 8)
    .getValues();

  return values.map(r => {

    const kode = (r[0] || '').toString().trim().toUpperCase();
    const nama = (r[1] || '').toString().trim();
    const satuan = (r[2] || '').toString().trim();
    const stokAkhir = Number(r[6]) || 0;
    const barcode = (r[7] || '').toString().trim().toUpperCase();

    return {
      code: kode,
      name: nama,
      unit: satuan,
      stock: stokAkhir, // 🔥 selalu dari kolom stok akhir
      barcode: barcode
    };

  });

}

function validatePurchasingPassword(password) {

  const admin = PETUGAS["PURCHASING"];

  if (!admin) return { ok: false };

  if (password === admin.password) {
    return { ok: true };
  }

  return { ok: false };
}

function getDashboardData() {

  const sh = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName('TRANSAKSI');

  if (!sh) return { summary: {}, daily: [] };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { summary: {}, daily: [] };

  const data = sh.getRange(2, 1, lastRow - 1, 6).getValues();

  let totalMasuk = 0;
  let totalKeluar = 0;

  const dailyMap = {};

  data.forEach(r => {

    const date = Utilities.formatDate(
      new Date(r[0]),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    const type = (r[3] || '').toString().toUpperCase();
    const qty = Number(r[4]) || 0;

    if (!dailyMap[date]) {
      dailyMap[date] = { masuk: 0, keluar: 0 };
    }

    if (type === "MASUK") {
      totalMasuk += qty;
      dailyMap[date].masuk += qty;
    }

    if (type === "KELUAR") {
      totalKeluar += qty;
      dailyMap[date].keluar += qty;
    }

  });

  const daily = Object.keys(dailyMap).sort().map(date => ({
    date,
    masuk: dailyMap[date].masuk,
    keluar: dailyMap[date].keluar
  }));

  return {
    summary: {
      totalMasuk,
      totalKeluar,
      net: totalMasuk - totalKeluar
    },
    daily
  };
}



