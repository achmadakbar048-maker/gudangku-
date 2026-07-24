import React, { useState, useMemo, useEffect, useRef } from "react";
import { createClient } from "@supabase/supabase-js";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer, LineChart, Line, BarChart, Bar, PieChart, Pie, Cell,
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, AreaChart, Area
} from "recharts";
import {
  LayoutDashboard, Boxes, Receipt, BarChart3, Search, Plus, Trash2, Pencil,
  X, AlertTriangle, TrendingUp, Wallet, ShoppingCart, History,
  LogOut, Bell, ArrowUp, ArrowDown, Download, FileSpreadsheet, Printer,
  Building2, Mail, MessageCircle, Lock, User, Settings, ArrowRightLeft,
  PackagePlus, PackageMinus, CheckCircle2, XCircle, Info, FileText, Tags,
  Sparkles, ChevronLeft, ChevronRight, Database
} from "lucide-react";
import LoginAnimation from "./LoginAnimation";

class ErrorBoundary extends React.Component {
  constructor(props) { super(props); this.state = { error: null, info: null }; }
  componentDidCatch(error, info) { console.error(error, info); this.setState({ error, info }); }
  render() {
    if (this.state.error) {
      return (
        <div style={{ padding: 24, color: '#E2574C', background: '#111', minHeight: '100vh' }}>
          <h2>Terjadi kesalahan pada aplikasi</h2>
          <div style={{ whiteSpace: 'pre-wrap', fontFamily: 'monospace', marginTop: 12 }}>{String(this.state.error)}</div>
          <div style={{ whiteSpace: 'pre-wrap', fontSize: 12, marginTop: 12 }}>{this.state.info?.componentStack}</div>
        </div>
      );
    }
    return this.props.children;
  }
}

// ---------- Aset Logo ----------
// Ikon rumah + gembok, digambar sebagai SVG (bukan raster) supaya tajam di semua ukuran.
const LOGO_ICON_SVG = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#EDEFF2" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M3 9l9-7 9 7v11a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"></path>
  <polyline points="9 22 9 12 15 12 15 22"></polyline>
  <circle cx="12" cy="14" r="2" fill="#1DB9A0" stroke="none"></circle>
  <path d="M11 15h2v3h-2z" fill="#1DB9A0" stroke="none"></path>
</svg>`;
const LOGO_ICON = `data:image/svg+xml,${encodeURIComponent(LOGO_ICON_SVG)}`;

// Wordmark "Gudang" + "Ku" (teal), dipakai di sidebar, halaman depan, dan halaman login.
function Logo({ ukuranIkon = 26, ukuranTeks = 18 }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
      <img src={LOGO_ICON} alt="GudangKu" style={{ width: ukuranIkon, height: "auto" }} />
      <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: ukuranTeks, color: "#EDEFF2" }}>
        Gudang<span style={{ color: "#1DB9A0" }}>Ku</span>
      </span>
    </div>
  );
}

// ---------- Utilitas ----------
const rupiah = (n) =>
  new Intl.NumberFormat("id-ID", { style: "currency", currency: "IDR", maximumFractionDigits: 0 }).format(n || 0);

const tanggalID = (iso) =>
  new Date(iso).toLocaleDateString("id-ID", { day: "2-digit", month: "short", year: "numeric" });

const tanggalWaktuID = (iso) =>
  new Date(iso).toLocaleString("id-ID", { day: "2-digit", month: "short", year: "numeric", hour: "2-digit", minute: "2-digit" });

const uid = () => Math.random().toString(36).slice(2, 10);

const pad = (n, l) => String(n).padStart(l, "0");

// ---------- Penyimpanan lokal (localStorage) ----------
// Semua data transaksional (produk, mutasi, penjualan, dll) disimpan di browser
// supaya tidak hilang saat halaman di-refresh. Dibungkus try/catch karena
// localStorage bisa gagal (mode incognito penuh, browser lama, dsb).
const LS_PREFIX = "gudangku_v1_";

function loadLS(key, fallback) {
  try {
    const raw = window.localStorage.getItem(LS_PREFIX + key);
    if (raw === null || raw === undefined) return fallback;
    return JSON.parse(raw);
  } catch (e) {
    console.error(`Gagal memuat data tersimpan "${key}":`, e);
    return fallback;
  }
}

function saveLS(key, value) {
  try {
    window.localStorage.setItem(LS_PREFIX + key, JSON.stringify(value));
    return true;
  } catch (e) {
    console.error(`Gagal menyimpan data "${key}":`, e);
    return false;
  }
}

function hapusSemuaDataLS() {
  try {
    Object.keys(window.localStorage)
      .filter(k => k.startsWith(LS_PREFIX))
      .forEach(k => window.localStorage.removeItem(k));
    return true;
  } catch (e) {
    console.error("Gagal menghapus data tersimpan:", e);
    return false;
  }
}

// ---------- Koneksi Supabase ----------
// URL & anon/publishable key AMAN ditulis langsung di kode sisi browser --
// itu memang fungsinya (mirip API key publik), akses ditentukan oleh
// Row Level Security di sisi database, bukan dengan menyembunyikan kunci ini.
const SUPABASE_URL = "https://cqkznabymijarsoaqnmg.supabase.co";
const SUPABASE_KEY = "sb_publishable_kQ3IAKjuN2PWeT3HG_ZJdw_0C3E0Wta";
const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

// Pemetaan nama kolom camelCase (dipakai di kode) <-> snake_case (dipakai di tabel Supabase).
// Daftar di sini jadi "whitelist" kolom yang sah -- field lokal yang tidak
// terdaftar (mis. field sementara seperti mutasiId di sales) otomatis
// tidak ikut terkirim ke Supabase.
const KOLOM = {
  products: { id: "id", kodeBarang: "kode_barang", nama: "nama", kategori: "kategori", gudang: "gudang", satuan: "satuan", stok: "stok", stokMin: "stok_min", hargaBeli: "harga_beli", hargaJual: "harga_jual", supplier: "supplier", createdAt: "created_at" },
  mutasi: { id: "id", kode: "kode", jenis: "jenis", produkId: "produk_id", namaProduk: "nama_produk", kodeBarang: "kode_barang", kategori: "kategori", gudang: "gudang", jumlah: "jumlah", satuan: "satuan", keterangan: "keterangan", oleh: "oleh", stokSebelum: "stok_sebelum", stokSesudah: "stok_sesudah", tanggal: "tanggal", transferId: "transfer_id", saleId: "sale_id", hargaSatuan: "harga_satuan" },
  sales: { id: "id", produkId: "produk_id", namaProduk: "nama_produk", kodeBarang: "kode_barang", kategori: "kategori", gudang: "gudang", jumlah: "jumlah", hargaJualSaat: "harga_jual_saat", hargaBeliSaat: "harga_beli_saat", total: "total", profit: "profit", tanggal: "tanggal", keterangan: "keterangan" },
  riwayat_harga: { id: "id", produkId: "produk_id", namaProduk: "nama_produk", kodeBarang: "kode_barang", field: "field", hargaLama: "harga_lama", hargaBaru: "harga_baru", tanggal: "tanggal" },
  categories: { id: "id", nama: "nama", prefix: "prefix", warna: "warna", counter: "counter" },
};

function keSnake(obj, table) {
  const peta = KOLOM[table];
  const out = {};
  for (const camel in peta) {
    if (obj[camel] !== undefined) out[peta[camel]] = obj[camel];
  }
  return out;
}

function keCamel(row, table) {
  const peta = KOLOM[table];
  const out = {};
  for (const camel in peta) {
    const kolom = peta[camel];
    if (row[kolom] !== undefined) out[camel] = row[kolom];
  }
  return out;
}

// Ambil bentuk singkat satuan, mis. "Piece (pcs)" -> "pcs"
const satuanSingkat = (s) => {
  if (!s) return "";
  const m = s.match(/\(([^)]+)\)/);
  return m ? m[1] : s;
};

// ---------- Satuan Kuantitas / Kemasan / Berat / Volume / Panjang ----------
const SATUAN_GROUPS = [
  { grup: "Kuantitas / Jumlah", opsi: ["Piece (pcs)", "Unit", "Lusin (12 pcs)", "Kodi (20 helai/buah)", "Gross (144 pcs)", "Rim (500 lembar)"] },
  { grup: "Satuan Kemasan", opsi: ["Box (kotak)", "Carton (karton/kardus)", "Pack/Paket", "Sachet (bungkus kecil)", "Bks (bungkus sedang)", "Bal (ikatan besar)", "Botol"] },
  { grup: "Satuan Berat", opsi: ["Kilogram (kg)", "Gram (gr)", "Ons", "Ton", "Pound (lbs)"] },
  { grup: "Satuan Volume / Isi", opsi: ["Liter (L)", "Mililiter (ml)", "Galon"] },
  { grup: "Satuan Panjang / Luas", opsi: ["Meter (m)", "Sentimeter (cm)", "Inci (in)"] },
];

const GUDANG = ["Gudang Mayabon", "Gudang Panancangan", "Gudang Panosogan", "Gudang Sentral"];
const GUDANG_WARNA = { "Gudang Mayabon": "#3FA796", "Gudang Panancangan": "#F2C14E", "Gudang Panosogan": "#6C8EBF", "Gudang Sentral": "#A16207" };

// ---------- Akun demo (login front-end, tanpa server) ----------
const AKUN = [
  { username: "admin", password: "akbarganteng", nama: "Admin Utama", role: "admin", gudang: "Semua" },
  { username: "mayabon1", password: "staff123", nama: "Staf Gudang Mayabon", role: "staff", gudang: "Gudang Mayabon" },
  { username: "panancangan1", password: "staff123", nama: "Staf Gudang Panancangan", role: "staff", gudang: "Gudang Panancangan" },
  { username: "panosogan1", password: "staff123", nama: "Staf Gudang Panosogan", role: "staff", gudang: "Gudang Panosogan" },
  { username: "sentral1", password: "staff123", nama: "Staf Gudang Sentral", role: "staff", gudang: "Gudang Sentral" },
];

// ---------- Kategori barang (sekarang dinamis, admin bisa tambah/edit/hapus) ----------
const CATEGORY_SEED = [
  { id: uid(), nama: "Elektronik", prefix: "ELK", warna: "#3FA796", counter: 3 },
  { id: uid(), nama: "Sparepart", prefix: "SPR", warna: "#F2C14E", counter: 2 },
  { id: uid(), nama: "Bahan Baku", prefix: "BHN", warna: "#E8A33D", counter: 2 },
  { id: uid(), nama: "Kemasan", prefix: "KMS", warna: "#6C8EBF", counter: 2 },
  { id: uid(), nama: "Alat Tulis", prefix: "ATK", warna: "#B48EAD", counter: 1 },
  { id: uid(), nama: "Peralatan", prefix: "PRL", warna: "#E2574C", counter: 2 },
];

// ---------- Data awal (seed) ----------
const SEED_PRODUCTS = [
  { id: uid(), kodeBarang: "ELK-001", nama: "Kabel HDMI 2m", kategori: "Elektronik", gudang: "Gudang Mayabon", satuan: "Piece (pcs)", stok: 84, stokMin: 30, hargaBeli: 22000, hargaJual: 38000, supplier: "CV Sumber Elektrik", createdAt: new Date().toISOString() },
  { id: uid(), kodeBarang: "ELK-002", nama: "Lampu LED 12W", kategori: "Elektronik", gudang: "Gudang Mayabon", satuan: "Piece (pcs)", stok: 18, stokMin: 25, hargaBeli: 15000, hargaJual: 27000, supplier: "PT Cahaya Terang", createdAt: new Date().toISOString() },
  { id: uid(), kodeBarang: "ELK-003", nama: "Adaptor 12V 2A", kategori: "Elektronik", gudang: "Gudang Mayabon", satuan: "Unit", stok: 52, stokMin: 20, hargaBeli: 31000, hargaJual: 52000, supplier: "CV Sumber Elektrik", createdAt: new Date().toISOString() },
  { id: uid(), kodeBarang: "SPR-001", nama: "Bearing 6203ZZ", kategori: "Sparepart", gudang: "Gudang Panancangan", satuan: "Piece (pcs)", stok: 9, stokMin: 15, hargaBeli: 18000, hargaJual: 32000, supplier: "UD Mesin Jaya" },
  { id: uid(), kodeBarang: "SPR-002", nama: "V-Belt A38", kategori: "Sparepart", gudang: "Gudang Panancangan", satuan: "Piece (pcs)", stok: 40, stokMin: 12, hargaBeli: 25000, hargaJual: 45000, supplier: "UD Mesin Jaya" },
  { id: uid(), kodeBarang: "BHN-002", nama: "Resin Epoxy 1kg", kategori: "Bahan Baku", gudang: "Gudang Panosogan", satuan: "Kilogram (kg)", stok: 22, stokMin: 20, hargaBeli: 48000, hargaJual: 75000, supplier: "PT Kimia Abadi" },
  { id: uid(), kodeBarang: "KMS-001", nama: "Kardus Box Sedang", kategori: "Kemasan", gudang: "Gudang Panosogan", satuan: "Box (kotak)", stok: 310, stokMin: 100, hargaBeli: 3200, hargaJual: 6000, supplier: "CV Karton Sejahtera" },
  { id: uid(), kodeBarang: "KMS-002", nama: "Bubble Wrap 1 Roll", kategori: "Kemasan", gudang: "Gudang Mayabon", satuan: "Unit", stok: 14, stokMin: 15, hargaBeli: 45000, hargaJual: 72000, supplier: "CV Karton Sejahtera" },
  { id: uid(), kodeBarang: "ATK-001", nama: "Spidol Permanen Hitam", kategori: "Alat Tulis", gudang: "Gudang Mayabon", satuan: "Piece (pcs)", stok: 76, stokMin: 30, hargaBeli: 4500, hargaJual: 8000, supplier: "Toko Sinar ATK" },
  { id: uid(), kodeBarang: "PRL-001", nama: "Sarung Tangan Safety", kategori: "Peralatan", gudang: "Gudang Panancangan", satuan: "Pack/Paket", stok: 60, stokMin: 25, hargaBeli: 12000, hargaJual: 22000, supplier: "PT Aman Kerja", createdAt: new Date().toISOString() },
  { id: uid(), kodeBarang: "PRL-002", nama: "Helm Proyek", kategori: "Peralatan", gudang: "Gudang Panancangan", satuan: "Unit", stok: 7, stokMin: 10, hargaBeli: 38000, hargaJual: 65000, supplier: "PT Aman Kerja", createdAt: new Date().toISOString() },
  { id: uid(), kodeBarang: "SNT-001", nama: "Tape Packing 48mm", kategori: "Kemasan", gudang: "Gudang Sentral", satuan: "Unit", stok: 150, stokMin: 50, hargaBeli: 1200, hargaJual: 2200, supplier: "CV Gudang Sentral", createdAt: new Date().toISOString() },
];

function buatSeedPenjualan(produkList) {
  const hasil = [];
  const hariIni = new Date();
  for (let i = 29; i >= 0; i--) {
    const tgl = new Date(hariIni);
    tgl.setDate(tgl.getDate() - i);
    const jumlahTransaksi = 1 + Math.floor(Math.random() * 4);
    for (let t = 0; t < jumlahTransaksi; t++) {
      const p = produkList[Math.floor(Math.random() * produkList.length)];
      const jml = 1 + Math.floor(Math.random() * 6);
      hasil.push({
        id: uid(), produkId: p.id, namaProduk: p.nama, kodeBarang: p.kodeBarang,
        kategori: p.kategori, gudang: p.gudang, jumlah: jml,
        hargaJualSaat: p.hargaJual, hargaBeliSaat: p.hargaBeli,
        total: jml * p.hargaJual, profit: jml * (p.hargaJual - p.hargaBeli),
        tanggal: tgl.toISOString(),
      });
    }
  }
  return hasil;
}

function buatSeedRiwayatHarga(produkList) {
  const contoh = [produkList[0], produkList[3], produkList[6], produkList[10]];
  const hasil = [];
  contoh.forEach((p, idx) => {
    const hariLalu = 50 - idx * 8;
    const tgl = new Date();
    tgl.setDate(tgl.getDate() - hariLalu);
    const hargaBeliLama = Math.round(p.hargaBeli * 0.88);
    const hargaJualLama = Math.round(p.hargaJual * 0.9);
    hasil.push({ id: uid(), produkId: p.id, namaProduk: p.nama, kodeBarang: p.kodeBarang, field: "Harga Beli", hargaLama: hargaBeliLama, hargaBaru: p.hargaBeli, tanggal: tgl.toISOString() });
    hasil.push({ id: uid(), produkId: p.id, namaProduk: p.nama, kodeBarang: p.kodeBarang, field: "Harga Jual", hargaLama: hargaJualLama, hargaBaru: p.hargaJual, tanggal: tgl.toISOString() });
  });
  return hasil;
}

function buatSeedMutasi(produkList) {
  const hasil = [];
  const jenisList = ["Masuk", "Masuk", "Keluar"];
  const keteranganMasuk = ["Pembelian dari supplier", "Retur dari pelanggan", "Hasil produksi/rakitan"];
  const keteranganKeluar = ["Transfer ke gudang lain", "Penyesuaian stok opname", "Pemakaian internal"];
  let counter = 1;
  for (let i = 11; i >= 1; i--) {
    const tgl = new Date();
    tgl.setDate(tgl.getDate() - i * 2);
    const p = produkList[Math.floor(Math.random() * produkList.length)];
    const jenis = jenisList[Math.floor(Math.random() * jenisList.length)];
    const jumlah = 2 + Math.floor(Math.random() * 15);
    const ymd = `${tgl.getFullYear()}${pad(tgl.getMonth() + 1, 2)}${pad(tgl.getDate(), 2)}`;
    hasil.push({
      id: uid(), kode: `MUT-${ymd}-${pad(counter, 4)}`, jenis, produkId: p.id, namaProduk: p.nama,
      kodeBarang: p.kodeBarang, kategori: p.kategori, gudang: p.gudang, jumlah, satuan: p.satuan,
      keterangan: jenis === "Masuk" ? keteranganMasuk[i % keteranganMasuk.length] : keteranganKeluar[i % keteranganKeluar.length],
      oleh: "Admin Utama", stokSebelum: p.stok, stokSesudah: p.stok, tanggal: tgl.toISOString(),
    });
    counter++;
  }
  return hasil;
}

// ---------- Komponen angka beranimasi (bikin dashboard terasa hidup) ----------
function AnimatedNumber({ value, formatFn }) {
  const [tampil, setTampil] = useState(0);
  const dariRef = useRef(0);

  useEffect(() => {
    const dari = dariRef.current;
    const tujuan = value || 0;
    const durasi = 650;
    const mulai = performance.now();
    let frame;
    function step(now) {
      const t = Math.min(1, (now - mulai) / durasi);
      const ease = 1 - Math.pow(1 - t, 3);
      setTampil(dari + (tujuan - dari) * ease);
      if (t < 1) frame = requestAnimationFrame(step);
      else dariRef.current = tujuan;
    }
    frame = requestAnimationFrame(step);
    return () => cancelAnimationFrame(frame);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [value]);

  const bulat = Math.round(tampil);
  return <span>{formatFn ? formatFn(bulat) : bulat}</span>;
}

// ---------- Komponen indikator stok bergaya "rak gudang" ----------
function IndikatorStok({ stok, stokMin, satuan }) {
  const rasio = stokMin > 0 ? stok / (stokMin * 2) : 1;
  const segmenAktif = Math.max(0, Math.min(6, Math.round(rasio * 6)));
  let warna = "#3FA796", label = "Aman";
  if (stok <= stokMin) { warna = "#E2574C"; label = "Kritis"; }
  else if (stok <= stokMin * 1.5) { warna = "#E8A33D"; label = "Menipis"; }

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 4, minWidth: 110 }}>
      <div style={{ display: "flex", gap: 2 }}>
        {Array.from({ length: 6 }).map((_, i) => (
          <div key={i} className="segmen-stok" style={{ height: 10, flex: 1, borderRadius: 1, background: i < segmenAktif ? warna : "#2A3138" }} />
        ))}
      </div>
      <span style={{ fontSize: 11, color: warna, fontFamily: "'IBM Plex Mono', monospace", letterSpacing: 0.3 }}>
        {label} · {stok} {satuanSingkat(satuan) || "unit"}
      </span>
    </div>
  );
}

function BarcodeDivider() {
  const bars = useMemo(() => Array.from({ length: 40 }).map(() => 1 + Math.random() * 3), []);
  return (
    <div style={{ display: "flex", alignItems: "flex-end", gap: 2, height: 14, opacity: 0.5, margin: "4px 0 18px" }}>
      {bars.map((w, i) => <div key={i} style={{ width: w, height: 8 + (i % 3) * 2, background: "#3FA796", opacity: 0.6 }} />)}
    </div>
  );
}

function KartuKPI({ ikon: Ikon, label, nilai, formatFn, sub, warna }) {
  return (
    <div className="kartu-kpi" style={{ background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, padding: "18px 20px", flex: 1, minWidth: 200 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <span style={{ fontSize: 12, color: "#8B95A1", letterSpacing: 0.4, textTransform: "uppercase" }}>{label}</span>
        <Ikon size={18} color={warna} />
      </div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 24, fontWeight: 600, color: "#EDEFF2", marginTop: 10 }}>
        <AnimatedNumber value={nilai} formatFn={formatFn} />
      </div>
      {sub && <div style={{ fontSize: 12, color: "#8B95A1", marginTop: 4 }}>{sub}</div>}
    </div>
  );
}

function Badge({ children, warna }) {
  return <span style={{ fontSize: 11, padding: "3px 8px", borderRadius: 20, background: "#232B32", color: warna }}>{children}</span>;
}

// ---------- Notifikasi Toast ----------
function ToastContainer({ toasts, onTutup }) {
  const ikon = { success: CheckCircle2, error: XCircle, info: Info };
  const warna = { success: "#3FA796", error: "#E2574C", info: "#6C8EBF" };
  return (
    <div className="no-print" style={{ position: "fixed", bottom: 20, right: 20, zIndex: 100, display: "flex", flexDirection: "column", gap: 10 }}>
      {toasts.map(t => {
        const Ikon = ikon[t.type] || Info;
        return (
          <div key={t.id} className="toast-masuk" onClick={() => onTutup(t.id)} style={{
            display: "flex", alignItems: "center", gap: 10, background: "#1D2329", border: `1px solid ${warna[t.type]}55`,
            borderLeft: `3px solid ${warna[t.type]}`, borderRadius: 8, padding: "12px 16px", minWidth: 260, maxWidth: 340,
            boxShadow: "0 8px 24px rgba(0,0,0,0.35)", cursor: "pointer"
          }}>
            <Ikon size={17} color={warna[t.type]} style={{ flexShrink: 0 }} />
            <span style={{ fontSize: 13, color: "#EDEFF2" }}>{t.message}</span>
          </div>
        );
      })}
    </div>
  );
}

// ---------- Ekspor data ----------
function eksporExcel(sheets, namaFile) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ nama, data }) => {
    const ws = XLSX.utils.json_to_sheet(data);
    XLSX.utils.book_append_sheet(wb, ws, nama.slice(0, 31));
  });
  XLSX.writeFile(wb, `${namaFile}.xlsx`);
}

function eksporCSV(data, namaFile) {
  if (!data.length) return;
  const kolom = Object.keys(data[0]);
  const baris = data.map(row => kolom.map(k => `"${String(row[k] ?? "").replace(/"/g, '""')}"`).join(","));
  const csv = [kolom.join(","), ...baris].join("\n");
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = `${namaFile}.csv`; a.click();
  URL.revokeObjectURL(url);
}

function cetakInvoiceMutasi(m) {
  const w = window.open("", "_blank", "width=420,height=650");
  if (!w) return;
  const warnaJenis = m.jenis === "Masuk" ? "#3FA796" : "#E2574C";
  w.document.write(`
    <html>
      <head>
        <title>${m.kode}</title>
        <style>
          body { font-family: 'Courier New', monospace; padding: 26px; color: #14181D; }
          .kop { display:flex; align-items:center; gap:10px; margin-bottom: 4px; }
          .kop img { height: 34px; }
          .kop b { font-size: 18px; letter-spacing: 0.5px; }
          .sub { font-size: 11px; color: #666; margin-bottom: 18px; }
          hr { border: none; border-top: 1px dashed #999; margin: 14px 0; }
          .baris { display:flex; justify-content:space-between; font-size: 13px; margin: 6px 0; }
          .label { color: #666; }
          .jenis { display:inline-block; padding: 3px 10px; border-radius: 20px; color:#fff; background:${warnaJenis}; font-size:12px; font-weight:bold; }
          .kode { text-align:center; font-size: 15px; letter-spacing: 1px; font-weight:bold; margin: 14px 0 4px; }
          .catatan { font-size: 12px; color:#555; margin-top: 10px; }
          .footer { text-align:center; font-size: 11px; color:#999; margin-top: 24px; }
        </style>
      </head>
      <body onload="window.print()">
        <div class="kop"><img src="${LOGO_ICON}" /><b>GudangKu</b></div>
        <div class="sub">Sistem Gudang Terintegrasi &mdash; Invoice Mutasi Barang</div>
        <div class="kode">${m.kode}</div>
        <div style="text-align:center;"><span class="jenis">${m.jenis === "Masuk" ? "BARANG MASUK" : "BARANG KELUAR"}</span></div>
        <hr/>
        <div class="baris"><span class="label">Tanggal</span><span>${tanggalWaktuID(m.tanggal)}</span></div>
        <div class="baris"><span class="label">Produk</span><span>${m.namaProduk}</span></div>
        <div class="baris"><span class="label">Kode Barang</span><span>${m.kodeBarang}</span></div>
        <div class="baris"><span class="label">Kategori</span><span>${m.kategori}</span></div>
        <div class="baris"><span class="label">Gudang</span><span>${m.gudang}</span></div>
        <div class="baris"><span class="label">Jumlah</span><span>${m.jumlah} ${satuanSingkat(m.satuan)}</span></div>
        <div class="baris"><span class="label">Stok Sebelum</span><span>${m.stokSebelum}</span></div>
        <div class="baris"><span class="label">Stok Sesudah</span><span>${m.stokSesudah}</span></div>
        <div class="baris"><span class="label">Dicatat oleh</span><span>${m.oleh}</span></div>
        <hr/>
        <div class="catatan"><b>Keterangan:</b> ${m.keterangan || "-"}</div>
        <div class="footer">Dicetak otomatis oleh GudangKu &middot; ${tanggalWaktuID(new Date().toISOString())}</div>
      </body>
    </html>
  `);
  w.document.close();
}

function TombolEkspor({ onExcel, onCSV, onPDF }) {
  const [buka, setBuka] = useState(false);
  return (
    <div style={{ position: "relative" }}>
      <button onClick={() => setBuka(v => !v)} style={{ ...btnSecondary, display: "flex", alignItems: "center", gap: 6 }}>
        <Download size={15} /> Ekspor
      </button>
      {buka && (
        <div className="dropdown-masuk" style={{ position: "absolute", right: 0, top: "110%", background: "#1D2329", border: "1px solid #2A3138", borderRadius: 8, zIndex: 20, minWidth: 170, overflow: "hidden" }}>
          <button onClick={() => { onExcel(); setBuka(false); }} style={itemDropdown}><FileSpreadsheet size={14} /> Excel (.xlsx)</button>
          <button onClick={() => { onCSV(); setBuka(false); }} style={itemDropdown}><FileSpreadsheet size={14} /> Spreadsheet (.csv)</button>
          <button onClick={() => { onPDF(); setBuka(false); }} style={itemDropdown}><Printer size={14} /> PDF (cetak)</button>
        </div>
      )}
    </div>
  );
}

const itemDropdown = {
  display: "flex", alignItems: "center", gap: 8, width: "100%", background: "none", border: "none",
  color: "#EDEFF2", padding: "10px 14px", fontSize: 13, textAlign: "left"
};

// ============================================================
export default function InventoryApp() {
  const [currentUser, setCurrentUser] = useState(null);
  const [tampilLanding, setTampilLanding] = useState(true);
  const [categories, setCategories] = useState(() => loadLS("categories", CATEGORY_SEED));
  const [products, setProducts] = useState(() => loadLS("products", SEED_PRODUCTS.map(p => ({ ...p, createdAt: p.createdAt || new Date().toISOString() }))));
  const [sales, setSales] = useState(() => loadLS("sales", buatSeedPenjualan(SEED_PRODUCTS)));
  const [riwayatHarga, setRiwayatHarga] = useState(() => loadLS("riwayatHarga", buatSeedRiwayatHarga(SEED_PRODUCTS)));
  const [mutasi, setMutasi] = useState(() => loadLS("mutasi", buatSeedMutasi(SEED_PRODUCTS)));
  const [nextMutasiNo, setNextMutasiNo] = useState(() => loadLS("nextMutasiNo", 12));

  const [tab, setTab] = useState("dashboard");
  const [sidebarTerbuka, setSidebarTerbuka] = useState(() => loadLS("sidebarTerbuka", true));
  useEffect(() => { saveLS("sidebarTerbuka", sidebarTerbuka); }, [sidebarTerbuka]);
  const [search, setSearch] = useState("");
  const [filterKategori, setFilterKategori] = useState("Semua");
  const [filterGudang, setFilterGudang] = useState("Semua");

  const [modalProduk, setModalProduk] = useState(null);
  const [modalJual, setModalJual] = useState(false);
  const [modalEditJual, setModalEditJual] = useState(null);
  const [modalRiwayat, setModalRiwayat] = useState(null);
  const [modalEditRiwayat, setModalEditRiwayat] = useState(null);
  const [modalNotifikasi, setModalNotifikasi] = useState(false);
  const [modalKategori, setModalKategori] = useState(false);
  const [modalMutasi, setModalMutasi] = useState(false);
  const [modalSetting, setModalSetting] = useState(false);
  const [invoiceTampil, setInvoiceTampil] = useState(null);
  const [errorForm, setErrorForm] = useState("");
  const [toasts, setToasts] = useState([]);
  const [settings, setSettings] = useState(() => loadLS("settings", { theme: "dark", waNumber: "6281234567890", notifSchedule: "08:30", autoNotify: false, hargaVisibility: {} }));
  const lastNotifyTimeRef = useRef(null);

  // Simpan otomatis ke localStorage setiap kali data berubah, supaya tidak hilang saat refresh.
  useEffect(() => { saveLS("categories", categories); }, [categories]);
  useEffect(() => { saveLS("products", products); }, [products]);
  useEffect(() => { saveLS("sales", sales); }, [sales]);
  useEffect(() => { saveLS("riwayatHarga", riwayatHarga); }, [riwayatHarga]);
  useEffect(() => { saveLS("mutasi", mutasi); }, [mutasi]);
  useEffect(() => { saveLS("nextMutasiNo", nextMutasiNo); }, [nextMutasiNo]);
  useEffect(() => { saveLS("settings", settings); }, [settings]);

  function pushToast(type, message) {
    const id = uid();
    setToasts(prev => [...prev, { id, type, message }]);
    setTimeout(() => setToasts(prev => prev.filter(t => t.id !== id)), 3800);
  }

  // ---------- Sinkronisasi ke Supabase ----------
  // Pola yang dipakai: state lokal diperbarui langsung (biar UI responsif),
  // lalu perubahan yang sama dikirim ke Supabase di latar belakang. Jika
  // gagal (mis. tidak ada internet), muncul toast error tapi perubahan
  // tetap tersimpan secara lokal.
  function simpanKeSupabase(table, camelObj) {
    supabase.from(table).upsert(keSnake(camelObj, table)).then(({ error }) => {
      if (error) { console.error(`Supabase upsert ${table}:`, error); pushToast("error", `Gagal sinkron ke server (${table}). Perubahan tersimpan di perangkat ini.`); }
    });
  }

  function simpanBanyakKeSupabase(table, camelArr) {
    if (!camelArr || !camelArr.length) return;
    supabase.from(table).upsert(camelArr.map(o => keSnake(o, table))).then(({ error }) => {
      if (error) { console.error(`Supabase upsert banyak ${table}:`, error); pushToast("error", `Gagal sinkron ke server (${table}).`); }
    });
  }

  function hapusDariSupabase(table, id) {
    supabase.from(table).delete().eq("id", id).then(({ error }) => {
      if (error) { console.error(`Supabase delete ${table}:`, error); pushToast("error", `Gagal menghapus di server (${table}).`); }
    });
  }

  function updateMassalKeSupabase(table, filterKolom, filterNilai, perubahan) {
    supabase.from(table).update(perubahan).eq(filterKolom, filterNilai).then(({ error }) => {
      if (error) { console.error(`Supabase update massal ${table}:`, error); pushToast("error", `Gagal sinkron perubahan massal ke server (${table}).`); }
    });
  }

  function simpanSettingsKeSupabase(s) {
    supabase.from("app_settings").upsert({ id: 1, theme: s.theme, wa_number: s.waNumber, notif_schedule: s.notifSchedule, auto_notify: s.autoNotify, harga_visibility: s.hargaVisibility || {} }).then(({ error }) => {
      if (error) { console.error("Supabase settings:", error); pushToast("error", "Gagal menyimpan pengaturan ke server."); }
    });
  }

  function simpanCounterKeSupabase(nilai) {
    supabase.from("counters").upsert({ key: "next_mutasi_no", value: nilai }).then(({ error }) => {
      if (error) console.error("Supabase counter:", error);
    });
  }

  function simpanSettings(baru) {
    setSettings(baru);
    simpanSettingsKeSupabase(baru);
  }

  // Pemuatan data awal dari Supabase saat aplikasi dibuka -- ini yang membuat
  // data terlihat sama di semua perangkat/akun, bukan cuma di localStorage
  // perangkat masing-masing seperti sebelumnya.
  useEffect(() => {
    let terpasang = true;
    async function muatAwal() {
      try {
        const [p, k, m, s, r, st, c] = await Promise.all([
          supabase.from("products").select("*"),
          supabase.from("categories").select("*"),
          supabase.from("mutasi").select("*").order("tanggal", { ascending: false }),
          supabase.from("sales").select("*").order("tanggal", { ascending: false }),
          supabase.from("riwayat_harga").select("*").order("tanggal", { ascending: false }),
          supabase.from("app_settings").select("*").eq("id", 1).maybeSingle(),
          supabase.from("counters").select("*").eq("key", "next_mutasi_no").maybeSingle(),
        ]);
        if (!terpasang) return;
        if (p.error || k.error || m.error || s.error || r.error) {
          throw p.error || k.error || m.error || s.error || r.error;
        }

        const produkKosong = p.data && p.data.length === 0;
        const kategoriKosong = k.data && k.data.length === 0;

        if (produkKosong && kategoriKosong) {
          // Database Supabase masih benar-benar kosong (baru pertama kali disiapkan).
          // Isi dengan data contoh supaya tampilan awal tidak kosong, sekaligus jadi
          // titik awal yang sama untuk semua akun/gudang. Catatan: kalau dua perangkat
          // sama-sama membuka aplikasi untuk PERTAMA KALINYA di saat bersamaan sebelum
          // proses ini selesai, keduanya bisa sama-sama melakukan seeding (jarang terjadi,
          // dan tidak merusak data -- hanya berpotensi duplikat data contoh).
          const produkAwal = SEED_PRODUCTS.map(pr => ({ ...pr, createdAt: pr.createdAt || new Date().toISOString() }));
          const penjualanAwal = buatSeedPenjualan(produkAwal);
          const riwayatAwal = buatSeedRiwayatHarga(produkAwal);
          const mutasiAwal = buatSeedMutasi(produkAwal);

          await supabase.from("categories").insert(CATEGORY_SEED.map(x => keSnake(x, "categories")));
          await supabase.from("products").insert(produkAwal.map(x => keSnake(x, "products")));
          if (penjualanAwal.length) await supabase.from("sales").insert(penjualanAwal.map(x => keSnake(x, "sales")));
          if (riwayatAwal.length) await supabase.from("riwayat_harga").insert(riwayatAwal.map(x => keSnake(x, "riwayat_harga")));
          if (mutasiAwal.length) await supabase.from("mutasi").insert(mutasiAwal.map(x => keSnake(x, "mutasi")));

          if (!terpasang) return;
          setCategories(CATEGORY_SEED);
          setProducts(produkAwal);
          setSales(penjualanAwal);
          setRiwayatHarga(riwayatAwal);
          setMutasi(mutasiAwal);
        } else {
          if (p.data) setProducts(p.data.map(row => keCamel(row, "products")));
          if (k.data) setCategories(k.data.map(row => keCamel(row, "categories")));
          if (m.data) setMutasi(m.data.map(row => keCamel(row, "mutasi")));
          if (s.data) setSales(s.data.map(row => keCamel(row, "sales")));
          if (r.data) setRiwayatHarga(r.data.map(row => keCamel(row, "riwayat_harga")));
        }

        if (st.data) setSettings(prev => ({ ...prev, theme: st.data.theme, waNumber: st.data.wa_number, notifSchedule: st.data.notif_schedule, autoNotify: st.data.auto_notify, hargaVisibility: st.data.harga_visibility || {} }));
        if (c.data) setNextMutasiNo(c.data.value);
      } catch (e) {
        console.error("Gagal memuat data dari server:", e);
        pushToast("error", "Gagal terhubung ke server Supabase. Menampilkan data tersimpan di perangkat ini sementara.");
      }
    }
    muatAwal();
    return () => { terpasang = false; };
  }, []);

  // Langganan realtime -- setiap perubahan dari perangkat/akun lain langsung
  // masuk ke state lokal tanpa perlu refresh halaman.
  useEffect(() => {
    function terapkanPerubahan(setter, table) {
      return (payload) => {
        if (payload.eventType === "DELETE") {
          setter(prev => prev.filter(x => x.id !== payload.old.id));
        } else {
          const baris = keCamel(payload.new, table);
          setter(prev => prev.some(x => x.id === baris.id) ? prev.map(x => x.id === baris.id ? baris : x) : [baris, ...prev]);
        }
      };
    }

    const channel = supabase.channel("gudangku-realtime")
      .on("postgres_changes", { event: "*", schema: "public", table: "products" }, terapkanPerubahan(setProducts, "products"))
      .on("postgres_changes", { event: "*", schema: "public", table: "mutasi" }, terapkanPerubahan(setMutasi, "mutasi"))
      .on("postgres_changes", { event: "*", schema: "public", table: "sales" }, terapkanPerubahan(setSales, "sales"))
      .on("postgres_changes", { event: "*", schema: "public", table: "riwayat_harga" }, terapkanPerubahan(setRiwayatHarga, "riwayat_harga"))
      .on("postgres_changes", { event: "*", schema: "public", table: "categories" }, terapkanPerubahan(setCategories, "categories"))
      .subscribe();

    return () => { supabase.removeChannel(channel); };
  }, []);

  function resetSemuaData() {
    hapusSemuaDataLS();
    pushToast("info", "Cache lokal dihapus. Memuat ulang data dari server...");
    setTimeout(() => window.location.reload(), 900);
  }

  const isAdmin = currentUser?.role === "admin";
  const gudangAktif = currentUser ? (isAdmin ? filterGudang : currentUser.gudang) : "Semua";
  const namaKategori = useMemo(() => categories.map(c => c.nama), [categories]);
  const warnaKategoriMap = useMemo(() => Object.fromEntries(categories.map(c => [c.nama, c.warna])), [categories]);
  const theme = useMemo(() => settings.theme === "dark" ? {
    bg: "#14181D", text: "#EDEFF2", surface: "#1D2329", surfaceAlt: "#171B20", border: "#2A3138", muted: "#8B95A1", accent: "#3FA796", danger: "#E2574C"
  } : {
    bg: "#F4F5F8", text: "#111827", surface: "#FFFFFF", surfaceAlt: "#F8FAFC", border: "#D1D5DB", muted: "#6B7280", accent: "#047857", danger: "#b91c1c"
  }, [settings.theme]);

  // ---------- Filter berdasarkan gudang ----------
  const productsGudang = useMemo(() =>
    gudangAktif === "Semua" ? products : products.filter(p => p.gudang === gudangAktif),
    [products, gudangAktif]);

  const salesGudang = useMemo(() =>
    gudangAktif === "Semua" ? sales : sales.filter(s => s.gudang === gudangAktif),
    [sales, gudangAktif]);

  const mutasiGudang = useMemo(() =>
    gudangAktif === "Semua" ? mutasi : mutasi.filter(m => m.gudang === gudangAktif),
    [mutasi, gudangAktif]);

  // ---------- Perhitungan turunan ----------
  const nilaiTotalStok = useMemo(() => productsGudang.reduce((s, p) => s + p.stok * p.hargaBeli, 0), [productsGudang]);
  const stokMenipis = useMemo(() => productsGudang.filter(p => p.stok <= p.stokMin), [productsGudang]);
  const totalPendapatan = useMemo(() => salesGudang.reduce((s, x) => s + x.total, 0), [salesGudang]);
  const totalProfit = useMemo(() => salesGudang.reduce((s, x) => s + x.profit, 0), [salesGudang]);
  const totalTransaksi = salesGudang.length;
  const mutasiHariIni = useMemo(() => {
    const hariIni = new Date().toISOString().slice(0, 10);
    return mutasiGudang.filter(m => m.tanggal.slice(0, 10) === hariIni).length;
  }, [mutasiGudang]);

  const trenHarian = useMemo(() => {
    const peta = {};
    salesGudang.forEach(s => {
      const k = s.tanggal.slice(0, 10);
      if (!peta[k]) peta[k] = { tanggal: k, pendapatan: 0, profit: 0 };
      peta[k].pendapatan += s.total; peta[k].profit += s.profit;
    });
    return Object.values(peta).sort((a, b) => a.tanggal.localeCompare(b.tanggal))
      .map(d => ({ ...d, label: tanggalID(d.tanggal).replace(/ \d{4}/, "") }));
  }, [salesGudang]);

  const produkTerlaris = useMemo(() => {
    const peta = {};
    salesGudang.forEach(s => { peta[s.namaProduk] = (peta[s.namaProduk] || 0) + s.jumlah; });
    return Object.entries(peta).map(([nama, jumlah]) => ({ nama, jumlah })).sort((a, b) => b.jumlah - a.jumlah).slice(0, 6);
  }, [salesGudang]);

  const distribusiKategori = useMemo(() => {
    const peta = {};
    productsGudang.forEach(p => { peta[p.kategori] = (peta[p.kategori] || 0) + p.stok * p.hargaBeli; });
    return Object.entries(peta).map(([kategori, nilai]) => ({ kategori, nilai }));
  }, [productsGudang]);

  useEffect(() => {
    if (!settings.autoNotify || !settings.waNumber || stokMenipis.length === 0) return;
    const sekarang = new Date();
    const waktu = `${pad(sekarang.getHours(), 2)}:${pad(sekarang.getMinutes(), 2)}`;
    if (waktu !== settings.notifSchedule) return;
    if (lastNotifyTimeRef.current === waktu) return;
    lastNotifyTimeRef.current = waktu;

    const tujuan = settings.waNumber.replace(/\D/g, "");
    if (!tujuan) {
      pushToast("error", "Nomor WhatsApp tidak valid untuk notifikasi otomatis.");
      return;
    }
    const baris = stokMenipis.map(p => `• ${p.nama} (${p.kodeBarang}) - sisa ${p.stok} ${satuanSingkat(p.satuan)}`).join("\n");
    const teks = `Peringatan Stok Menipis - GudangKu\n\n${baris}`;
    window.open(`https://wa.me/${tujuan}?text=${encodeURIComponent(teks)}`, "_blank");
    pushToast("info", "Mencoba mengirim notifikasi WhatsApp stok minim secara otomatis.");
  }, [settings, stokMenipis]);

  const produkFilter = useMemo(() => {
    return productsGudang.filter(p => {
      const cocokCari = (p.nama + p.kodeBarang).toLowerCase().includes(search.toLowerCase());
      const cocokKategori = filterKategori === "Semua" || p.kategori === filterKategori;
      return cocokCari && cocokKategori;
    }).sort((a, b) => {
      const ta = a.createdAt ? new Date(a.createdAt).getTime() : 0;
      const tb = b.createdAt ? new Date(b.createdAt).getTime() : 0;
      if (ta !== tb) return tb - ta;
      return a.stok / (a.stokMin || 1) - b.stok / (b.stokMin || 1);
    });
  }, [productsGudang, search, filterKategori]);

  useEffect(() => {
    console.log('currentUser', currentUser);
  }, [currentUser]);

  if (!currentUser) {
    if (tampilLanding) return <LandingPage onMulai={() => setTampilLanding(false)} />;
    return <LoginView onLogin={setCurrentUser} onKembali={() => setTampilLanding(true)} />;
  }

  // Admin selalu bisa lihat harga. Untuk akun gudang/staf, tergantung pengaturan
  // yang dipilih admin per gudang (default: tampil, kecuali admin sengaja menyembunyikannya).
  const bisaLihatHarga = isAdmin || settings.hargaVisibility?.[currentUser.gudang] !== false;

  function setHargaVisibility(gudang, tampil) {
    const baru = { ...settings, hargaVisibility: { ...(settings.hargaVisibility || {}), [gudang]: tampil } };
    simpanSettings(baru);
  }

  // ---------- Logic CRUD Kategori ----------
  function tambahKategori(nama, prefix, warna) {
    if (!nama.trim()) return "Nama kategori wajib diisi.";
    if (categories.some(c => c.nama.toLowerCase() === nama.trim().toLowerCase())) return "Kategori sudah ada.";
    const pfx = (prefix || nama.slice(0, 3)).toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 4) || "GEN";
    const baru = { id: uid(), nama: nama.trim(), prefix: pfx, warna: warna || "#3FA796", counter: 0 };
    setCategories(prev => [...prev, baru]);
    simpanKeSupabase("categories", baru);
    pushToast("success", `Kategori "${nama.trim()}" berhasil ditambahkan.`);
    return null;
  }

  function editKategori(id, nama, prefix, warna) {
    if (!nama.trim()) return "Nama kategori wajib diisi.";
    const lama = categories.find(c => c.id === id);
    if (!lama) return "Kategori tidak ditemukan.";
    const namaBaru = nama.trim();
    const pfx = (prefix || namaBaru.slice(0, 3)).toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 4) || "GEN";
    const diperbarui = { ...lama, nama: namaBaru, prefix: pfx, warna };
    setCategories(prev => prev.map(c => c.id === id ? diperbarui : c));
    simpanKeSupabase("categories", diperbarui);
    if (lama.nama !== namaBaru) {
      setProducts(prev => prev.map(p => p.kategori === lama.nama ? { ...p, kategori: namaBaru } : p));
      setSales(prev => prev.map(s => s.kategori === lama.nama ? { ...s, kategori: namaBaru } : s));
      setMutasi(prev => prev.map(m => m.kategori === lama.nama ? { ...m, kategori: namaBaru } : m));
      updateMassalKeSupabase("products", "kategori", lama.nama, { kategori: namaBaru });
      updateMassalKeSupabase("sales", "kategori", lama.nama, { kategori: namaBaru });
      updateMassalKeSupabase("mutasi", "kategori", lama.nama, { kategori: namaBaru });
    }
    pushToast("success", `Kategori "${namaBaru}" berhasil diperbarui.`);
    return null;
  }

  function hapusKategori(id) {
    const kat = categories.find(c => c.id === id);
    if (!kat) return;
    const dipakai = products.some(p => p.kategori === kat.nama);
    if (dipakai) { pushToast("error", `Kategori "${kat.nama}" masih dipakai produk, tidak bisa dihapus.`); return; }
    setCategories(prev => prev.filter(c => c.id !== id));
    hapusDariSupabase("categories", id);
    pushToast("success", `Kategori "${kat.nama}" dihapus.`);
  }

  // ---------- Logic CRUD Produk ----------
  function simpanProduk(data) {
    if (!data.nama.trim() || !data.kategori) { setErrorForm("Nama dan kategori wajib diisi."); return; }
    if (!data.satuan) { setErrorForm("Satuan wajib dipilih."); return; }
    if (bisaLihatHarga && data.hargaJual <= data.hargaBeli) { setErrorForm("Harga jual harus lebih besar dari harga beli."); return; }
    if (data.stok < 0 || data.stokMin < 0) { setErrorForm("Stok tidak boleh negatif."); return; }

    if (data.id) {
      const lama = products.find(p => p.id === data.id);
      // Pengaman: stok hanya boleh berubah lewat Mutasi Barang, bukan lewat form edit produk,
      // supaya histori stok (Mutasi) selalu 100% sinkron dengan angka stok di Inventori.
      if (lama) data = { ...data, stok: lama.stok };
      const perubahanHarga = [];
      if (lama && lama.hargaBeli !== data.hargaBeli) {
        perubahanHarga.push({ id: uid(), produkId: data.id, namaProduk: data.nama, kodeBarang: data.kodeBarang, field: "Harga Beli", hargaLama: lama.hargaBeli, hargaBaru: data.hargaBeli, tanggal: new Date().toISOString() });
      }
      if (lama && lama.hargaJual !== data.hargaJual) {
        perubahanHarga.push({ id: uid(), produkId: data.id, namaProduk: data.nama, kodeBarang: data.kodeBarang, field: "Harga Jual", hargaLama: lama.hargaJual, hargaBaru: data.hargaJual, tanggal: new Date().toISOString() });
      }
      if (perubahanHarga.length) { setRiwayatHarga(prev => [...perubahanHarga, ...prev]); simpanBanyakKeSupabase("riwayat_harga", perubahanHarga); }
      setProducts(prev => prev.map(p => p.id === data.id ? data : p));
      simpanKeSupabase("products", data);
      pushToast("success", `Produk "${data.nama}" berhasil diperbarui.`);
    } else {
      // Kalau nama produk ini sudah pernah diinput sebelumnya (di gudang manapun),
      // pakai ulang kode barang yang sama -- jangan buat kode baru / menambah nomor urut kategori.
      const produkSerupa = products.find(p => p.nama.trim().toLowerCase() === data.nama.trim().toLowerCase());
      let kodeBarang;
      if (produkSerupa) {
        kodeBarang = produkSerupa.kodeBarang;
      } else {
        const kat = categories.find(c => c.nama === data.kategori);
        const nomorUrut = (kat?.counter || 0) + 1;
        kodeBarang = `${kat?.prefix || "GEN"}-${pad(nomorUrut, 3)}`;
        if (kat) {
          const katDiperbarui = { ...kat, counter: nomorUrut };
          setCategories(prev => prev.map(c => c.id === kat.id ? katDiperbarui : c));
          simpanKeSupabase("categories", katDiperbarui);
        }
      }
      const idBaru = uid();
      const createdAt = new Date().toISOString();
      const produkBaru = { ...data, id: idBaru, kodeBarang, createdAt };
      setProducts(prev => [...prev, produkBaru]);
      simpanKeSupabase("products", produkBaru);

      // Sinkronkan dengan Riwayat Harga: harga awal produk baru dicatat supaya langsung
      // muncul di halaman Riwayat Harga (bukan hanya perubahan harga berikutnya).
      const riwayatAwal = [];
      if (data.hargaBeli > 0) {
        riwayatAwal.push({ id: uid(), produkId: idBaru, namaProduk: data.nama, kodeBarang, field: "Harga Beli", hargaLama: 0, hargaBaru: data.hargaBeli, tanggal: createdAt });
      }
      if (data.hargaJual > 0) {
        riwayatAwal.push({ id: uid(), produkId: idBaru, namaProduk: data.nama, kodeBarang, field: "Harga Jual", hargaLama: 0, hargaBaru: data.hargaJual, tanggal: createdAt });
      }
      if (riwayatAwal.length) {
        setRiwayatHarga(prev => [...riwayatAwal, ...prev]);
        simpanBanyakKeSupabase("riwayat_harga", riwayatAwal);
      }

      // Sinkronkan dengan Mutasi Barang: stok awal produk baru dicatat sebagai transaksi "Masuk"
      // supaya produk baru langsung muncul di menu Mutasi, KPI, dan grafik terkait.
      if (data.stok > 0) {
        const tgl = new Date(createdAt);
        const ymd = `${tgl.getFullYear()}${pad(tgl.getMonth() + 1, 2)}${pad(tgl.getDate(), 2)}`;
        const kodeMutasi = `MUT-${ymd}-${pad(nextMutasiNo, 4)}`;
        const mutasiBaru = {
          id: uid(), kode: kodeMutasi, jenis: "Masuk", produkId: idBaru, namaProduk: data.nama, kodeBarang,
          kategori: data.kategori, gudang: data.gudang, jumlah: data.stok, satuan: data.satuan,
          keterangan: "Stok awal produk baru", oleh: currentUser.nama,
          stokSebelum: 0, stokSesudah: data.stok, tanggal: createdAt, hargaSatuan: data.hargaBeli,
        };
        setMutasi(prev => [mutasiBaru, ...prev]);
        simpanKeSupabase("mutasi", mutasiBaru);
        const nomorMutasiBaru = nextMutasiNo + 1;
        setNextMutasiNo(nomorMutasiBaru);
        simpanCounterKeSupabase(nomorMutasiBaru);
      }
      pushToast("success", `Produk "${data.nama}" ditambahkan dengan kode ${kodeBarang}.`);
    }
    setErrorForm("");
    setModalProduk(null);
  }

  function hapusProduk(id) {
    const p = products.find(x => x.id === id);
    setProducts(prev => prev.filter(p => p.id !== id));
    hapusDariSupabase("products", id);
    if (p) pushToast("info", `Produk "${p.nama}" dihapus dari inventori.`);
  }

  // ---------- Logic Penjualan ----------
  function catatPenjualan(produkId, jumlah, tanggal, hargaJualOverride, keterangan, transactionType = "Penjualan", tujuanGudang = "") {
    const p = products.find(x => x.id === produkId);
    if (!p) return "Produk tidak ditemukan.";
    if (jumlah <= 0) return "Jumlah harus lebih dari 0.";

    const satuanFinal = p.satuan;
    const hargaJualFinal = hargaJualOverride > 0 ? hargaJualOverride : p.hargaJual;
    const defaultKeterangan = keterangan || (transactionType === "Transfer" ? `Transfer ke ${tujuanGudang}` : transactionType === "Keluar" ? "Barang Keluar" : "Penjualan");

    if (transactionType === "Keluar" && jumlah > p.stok) return `Stok tidak mencukupi. Sisa stok: ${p.stok} ${satuanSingkat(p.satuan)}.`;
    if (transactionType === "Transfer" && (!tujuanGudang || tujuanGudang === p.gudang)) return "Pilih gudang tujuan yang berbeda untuk transfer.";

    if (transactionType === "Keluar") {
      return catatMutasi("Keluar", produkId, jumlah, tanggal, keterangan || defaultKeterangan, satuanFinal, 0, "");
    }

    if (transactionType === "Transfer") {
      return catatMutasi("Transfer", produkId, jumlah, tanggal, keterangan || defaultKeterangan, satuanFinal, 0, tujuanGudang);
    }

    const total = jumlah * hargaJualFinal;
    const saleId = uid();
    const tanggalIso = new Date(tanggal).toISOString();
    const mutasiId = uid();
    const kode = `MUT-${tanggalIso.slice(0,4)}${tanggalIso.slice(5,7)}${tanggalIso.slice(8,10)}-${pad(nextMutasiNo, 4)}`;

    const perubahanHarga = [];
    if (hargaJualFinal !== p.hargaJual) {
      perubahanHarga.push({ id: uid(), produkId: p.id, namaProduk: p.nama, kodeBarang: p.kodeBarang, field: "Harga Jual", hargaLama: p.hargaJual, hargaBaru: hargaJualFinal, tanggal: tanggalIso });
    }
    if (perubahanHarga.length) { setRiwayatHarga(prev => [...perubahanHarga, ...prev]); simpanBanyakKeSupabase("riwayat_harga", perubahanHarga); }

    const produkDiperbarui = { ...p, stok: p.stok - jumlah, hargaJual: hargaJualFinal };
    setProducts(prev => prev.map(x => x.id === produkId ? produkDiperbarui : x));
    simpanKeSupabase("products", produkDiperbarui);

    const saleBaru = {
      id: saleId, produkId, namaProduk: p.nama, kodeBarang: p.kodeBarang, kategori: p.kategori, gudang: p.gudang,
      jumlah, hargaJualSaat: hargaJualFinal, hargaBeliSaat: p.hargaBeli,
      total, profit: jumlah * (hargaJualFinal - p.hargaBeli), tanggal: tanggalIso, keterangan: defaultKeterangan, mutasiId,
    };
    setSales(prev => [saleBaru, ...prev]);
    simpanKeSupabase("sales", saleBaru);

    const mutasiBaru = {
      id: mutasiId, kode, jenis: "Keluar", produkId, namaProduk: p.nama, kodeBarang: p.kodeBarang,
      kategori: p.kategori, gudang: p.gudang, jumlah, satuan: p.satuan, keterangan: keterangan || "Penjualan",
      oleh: currentUser.nama, stokSebelum: p.stok, stokSesudah: p.stok - jumlah, tanggal: tanggalIso, saleId,
    };
    setMutasi(prev => [mutasiBaru, ...prev]);
    simpanKeSupabase("mutasi", mutasiBaru);

    const nomorMutasiBaru = nextMutasiNo + 1;
    setNextMutasiNo(nomorMutasiBaru);
    simpanCounterKeSupabase(nomorMutasiBaru);
    pushToast("success", `Penjualan ${p.nama} sebanyak ${jumlah} ${satuanSingkat(p.satuan)} dicatat dan disinkronkan dengan mutasi.`);
    return null;
  }

  function perbaikiPenjualan(saleId, produkIdBaru, jumlahBaru, tanggalBaru, hargaJualOverride, keterangan) {
    const saleLama = sales.find(s => s.id === saleId);
    if (!saleLama) return "Transaksi tidak ditemukan.";
    const produkBaru = products.find(p => p.id === produkIdBaru);
    if (!produkBaru) return "Produk tidak ditemukan.";
    if (jumlahBaru <= 0) return "Jumlah harus lebih dari 0.";

    const stokTersedia = produkIdBaru === saleLama.produkId ? produkBaru.stok + saleLama.jumlah : produkBaru.stok;
    if (jumlahBaru > stokTersedia) return `Stok tidak mencukupi. Sisa stok: ${stokTersedia} unit.`;

    const hargaJualFinal = hargaJualOverride > 0 ? hargaJualOverride : produkBaru.hargaJual;
    const tanggalIso = new Date(tanggalBaru).toISOString();

    let daftarProdukBaru;
    setProducts(prev => {
      daftarProdukBaru = prev.map(p => {
        if (p.id === saleLama.produkId) p = { ...p, stok: p.stok + saleLama.jumlah };
        if (p.id === produkIdBaru) p = { ...p, stok: p.stok - jumlahBaru };
        return p;
      });
      return daftarProdukBaru;
    });
    const produkTerpengaruh = daftarProdukBaru.filter(p => p.id === saleLama.produkId || p.id === produkIdBaru);
    simpanBanyakKeSupabase("products", produkTerpengaruh);

    const perubahanHarga = [];
    if (hargaJualFinal !== produkBaru.hargaJual) {
      perubahanHarga.push({ id: uid(), produkId: produkBaru.id, namaProduk: produkBaru.nama, kodeBarang: produkBaru.kodeBarang, field: "Harga Jual", hargaLama: produkBaru.hargaJual, hargaBaru: hargaJualFinal, tanggal: tanggalIso });
    }
    if (perubahanHarga.length) { setRiwayatHarga(prev => [...perubahanHarga, ...prev]); simpanBanyakKeSupabase("riwayat_harga", perubahanHarga); }

    const saleDiperbarui = {
      ...saleLama, produkId: produkIdBaru, namaProduk: produkBaru.nama, kodeBarang: produkBaru.kodeBarang,
      kategori: produkBaru.kategori, gudang: produkBaru.gudang, jumlah: jumlahBaru, hargaJualSaat: hargaJualFinal,
      hargaBeliSaat: produkBaru.hargaBeli, total: jumlahBaru * hargaJualFinal,
      profit: jumlahBaru * (hargaJualFinal - produkBaru.hargaBeli), tanggal: tanggalIso,
      keterangan: keterangan || saleLama.keterangan,
    };
    setSales(prev => prev.map(s => s.id === saleId ? saleDiperbarui : s));
    simpanKeSupabase("sales", saleDiperbarui);

    let mutasiDiperbarui = null;
    setMutasi(prev => prev.map(m => {
      if (m.saleId !== saleId) return m;
      mutasiDiperbarui = {
        ...m, produkId: produkIdBaru, namaProduk: produkBaru.nama, kodeBarang: produkBaru.kodeBarang,
        kategori: produkBaru.kategori, gudang: produkBaru.gudang, jumlah: jumlahBaru,
        satuan: produkBaru.satuan, keterangan: keterangan || m.keterangan,
        stokSesudah: m.stokSebelum - jumlahBaru, tanggal: tanggalIso,
      };
      return mutasiDiperbarui;
    }));
    if (mutasiDiperbarui) simpanKeSupabase("mutasi", mutasiDiperbarui);

    pushToast("success", "Transaksi penjualan berhasil diperbarui dan mutasi disinkronkan.");
    return null;
  }

  function hapusPenjualan(saleId) {
    const s = sales.find(x => x.id === saleId);
    if (!s) return;
    let produkDiperbarui = null;
    setProducts(prev => prev.map(p => {
      if (p.id !== s.produkId) return p;
      produkDiperbarui = { ...p, stok: p.stok + s.jumlah };
      return produkDiperbarui;
    }));
    if (produkDiperbarui) simpanKeSupabase("products", produkDiperbarui);
    setSales(prev => prev.filter(x => x.id !== saleId));
    hapusDariSupabase("sales", saleId);
    pushToast("info", "Transaksi penjualan dihapus & stok dikembalikan.");
  }

  function editRiwayatHarga(entryId, hargaLama, hargaBaru, keterangan) {
    const existing = riwayatHarga.find(r => r.id === entryId);
    if (!existing) return "Riwayat tidak ditemukan.";
    if (hargaLama <= 0 || hargaBaru <= 0) return "Harga harus lebih dari 0.";
    const diperbarui = { ...existing, hargaLama, hargaBaru, keterangan };
    setRiwayatHarga(prev => prev.map(r => r.id === entryId ? diperbarui : r));
    simpanKeSupabase("riwayat_harga", diperbarui);
    pushToast("success", "Riwayat harga berhasil diperbarui.");
    return null;
  }

  // ---------- Logic Mutasi Stok (Barang Masuk / Keluar) ----------
  function catatMutasi(jenis, produkId, jumlah, tanggal, keterangan, satuanOverride, hargaSatuan, tujuanGudang) {
    const p = products.find(x => x.id === produkId);
    if (!p) return "Produk tidak ditemukan.";
    if (jumlah <= 0) return "Jumlah harus lebih dari 0.";
    const satuanFinal = satuanOverride || p.satuan;
    const hargaFinal = hargaSatuan > 0 ? hargaSatuan : p.hargaBeli;

    if (jenis === "Keluar" && jumlah > p.stok) return `Stok tidak mencukupi. Sisa stok: ${p.stok} ${satuanSingkat(p.satuan)}.`;
    if (jenis === "Transfer" && (!tujuanGudang || tujuanGudang === p.gudang)) return "Pilih gudang tujuan yang berbeda untuk transfer.";

    const tgl = new Date(tanggal);
    const ymd = `${tgl.getFullYear()}${pad(tgl.getMonth() + 1, 2)}${pad(tgl.getDate(), 2)}`;
    const kode = `MUT-${ymd}-${pad(nextMutasiNo, 4)}`;
    const transferId = jenis === "Transfer" ? uid() : null;

    if (jenis === "Transfer") {
      const q = products.find(x => x.kodeBarang === p.kodeBarang && x.gudang === tujuanGudang);
      const stokSebelumA = p.stok;
      const stokSesudahA = p.stok - jumlah;
      const targetProduct = q ? { ...q, stok: q.stok + jumlah } : { ...p, id: uid(), gudang: tujuanGudang, stok: jumlah, kodeBarang: p.kodeBarang, satuan: satuanFinal, createdAt: tgl.toISOString() };
      const stokSebelumB = q ? q.stok : 0;
      const stokSesudahB = stokSebelumB + jumlah;

      const recordA = {
        id: uid(), kode, jenis: "Keluar", produkId: p.id, namaProduk: p.nama, kodeBarang: p.kodeBarang, kategori: p.kategori,
        gudang: p.gudang, jumlah, satuan: satuanFinal, keterangan: keterangan || `Transfer ke ${tujuanGudang}`, oleh: currentUser.nama,
        stokSebelum: stokSebelumA, stokSesudah: stokSesudahA, tanggal: tgl.toISOString(), transferId,
      };

      const recordB = {
        id: uid(), kode, jenis: "Masuk", produkId: targetProduct.id, namaProduk: targetProduct.nama, kodeBarang: targetProduct.kodeBarang, kategori: targetProduct.kategori,
        gudang: tujuanGudang, jumlah, satuan: satuanFinal, keterangan: keterangan || `Transfer dari ${p.gudang}`, oleh: currentUser.nama,
        stokSebelum: stokSebelumB, stokSesudah: stokSesudahB, tanggal: tgl.toISOString(), transferId,
      };

      const produkAsalDiperbarui = { ...p, stok: stokSesudahA };
      setProducts(prev => {
        const updated = prev.map(x => x.id === p.id ? produkAsalDiperbarui : x);
        if (q) return updated.map(x => x.id === targetProduct.id ? targetProduct : x);
        return [...updated, targetProduct];
      });
      simpanKeSupabase("products", produkAsalDiperbarui);
      simpanKeSupabase("products", targetProduct);
      setMutasi(prev => [recordB, recordA, ...prev]);
      simpanKeSupabase("mutasi", recordA);
      simpanKeSupabase("mutasi", recordB);
      const nomorMutasiBaru = nextMutasiNo + 1;
      setNextMutasiNo(nomorMutasiBaru);
      simpanCounterKeSupabase(nomorMutasiBaru);
      setInvoiceTampil(recordA);
      pushToast("success", `Barang transfer dicatat — ${kode}`);
      return null;
    }

    const stokSebelum = p.stok;
    const stokSesudah = jenis === "Masuk" ? p.stok + jumlah : p.stok - jumlah;
    const record = {
      id: uid(), kode, jenis, produkId: p.id, namaProduk: p.nama, kodeBarang: p.kodeBarang, kategori: p.kategori,
      gudang: p.gudang, jumlah, satuan: satuanFinal, keterangan, oleh: currentUser.nama,
      stokSebelum, stokSesudah, tanggal: tgl.toISOString(), hargaSatuan: hargaFinal,
    };
    const produkDiperbarui = { ...p, stok: stokSesudah, hargaBeli: jenis === "Masuk" ? hargaFinal : p.hargaBeli, satuan: satuanFinal };

    setProducts(prev => prev.map(x => x.id === p.id ? produkDiperbarui : x));
    simpanKeSupabase("products", produkDiperbarui);
    setMutasi(prev => [record, ...prev]);
    simpanKeSupabase("mutasi", record);
    const nomorMutasiBaru = nextMutasiNo + 1;
    setNextMutasiNo(nomorMutasiBaru);
    simpanCounterKeSupabase(nomorMutasiBaru);
    setInvoiceTampil(record);
    pushToast("success", `Barang ${jenis.toLowerCase()} dicatat — ${kode}`);
    return null;
  }

  const NAV = [
    { id: "dashboard", label: "Dashboard", ikon: LayoutDashboard },
    { id: "inventory", label: "Inventori", ikon: Boxes },
    { id: "masterdata", label: "Master Data", ikon: Database },
    { id: "mutasi", label: "Mutasi Barang", ikon: ArrowRightLeft },
    { id: "sales", label: "Penjualan", ikon: Receipt },
    { id: "reports", label: "Laporan", ikon: BarChart3 },
    { id: "riwayat", label: "Riwayat Harga", ikon: History },
  ];

  return (
    <ErrorBoundary>
      <div style={{ display: "flex", minHeight: "100vh", background: theme.bg, color: theme.text, fontFamily: "'Inter', sans-serif", fontSize: 14 }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=Inter:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        button { cursor: pointer; font-family: inherit; transition: all 0.15s ease; }
        button:active { transform: scale(0.97); }
        input, select { font-family: inherit; transition: border-color 0.15s ease, box-shadow 0.15s ease; }
        input:focus, select:focus { border-color: ${theme.accent} !important; box-shadow: 0 0 0 3px rgba(63,167,150,0.15); }
        table { border-collapse: collapse; width: 100%; }
        th, td { text-align: left; padding: 10px 12px; }
        tbody tr { transition: background 0.12s ease; }
        tbody tr:hover { background: rgba(63,167,150,0.06); }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: #2A3138; border-radius: 4px; }
        .kartu-kpi { transition: transform 0.18s ease, box-shadow 0.18s ease, border-color 0.18s ease; }
        .kartu-kpi:hover { transform: translateY(-3px); box-shadow: 0 10px 26px rgba(0,0,0,0.28); border-color: #3FA796; }
        .nav-btn { transition: background 0.15s ease, color 0.15s ease, transform 0.15s ease; position: relative; }
        .nav-btn:hover { background: #1E252B !important; transform: translateX(2px); }
        .nav-btn.aktif::before { content: ""; position: absolute; left: -16px; top: 8px; bottom: 8px; width: 3px; background: #3FA796; border-radius: 3px; }
        .icon-btn-hover:hover { border-color: #3FA796 !important; color: #3FA796 !important; }
        .tab-fade { animation: fadeInUp 0.32s ease; }
        @keyframes fadeInUp { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
        @keyframes overlayIn { from { opacity: 0; } to { opacity: 1; } }
        @keyframes modalIn { from { opacity: 0; transform: scale(0.96) translateY(6px); } to { opacity: 1; transform: scale(1) translateY(0); } }
        .overlay-anim { animation: overlayIn 0.18s ease; }
        .modal-anim { animation: modalIn 0.22s cubic-bezier(.2,.9,.3,1.2); }
        @keyframes toastIn { from { opacity: 0; transform: translateX(30px); } to { opacity: 1; transform: translateX(0); } }
        .toast-masuk { animation: toastIn 0.25s ease; }
        @keyframes dropdownIn { from { opacity: 0; transform: translateY(-6px); } to { opacity: 1; transform: translateY(0); } }
        .dropdown-masuk { animation: dropdownIn 0.15s ease; }
        .segmen-stok { transition: background 0.3s ease; }
        .jenis-pill { transition: all 0.15s ease; }
        .btn-primary-glow:hover { box-shadow: 0 4px 16px rgba(63,167,150,0.35); transform: translateY(-1px); }
        .btn-danger-glow:hover { box-shadow: 0 4px 16px rgba(226,87,76,0.3); }
        @media print {
          .no-print { display: none !important; }
          .area-cetak { color: #000 !important; }
        }
      `}</style>

      {/* Sidebar */}
      <div className="no-print" style={{
        width: sidebarTerbuka ? 226 : 72, background: "#171B20", borderRight: "1px solid #2A3138",
        padding: sidebarTerbuka ? "22px 16px" : "22px 12px", display: "flex", flexDirection: "column", gap: 4,
        flexShrink: 0, position: "relative", transition: "width 0.22s ease, padding 0.22s ease", overflow: "hidden",
      }}>
        <button
          onClick={() => setSidebarTerbuka(v => !v)}
          title={sidebarTerbuka ? "Ciutkan menu" : "Buka menu"}
          className="icon-btn-hover"
          style={{
            position: "absolute", top: 24, right: -12, width: 24, height: 24, borderRadius: "50%",
            background: "#1D2329", border: "1px solid #2A3138", color: "#8B95A1", display: "flex",
            alignItems: "center", justifyContent: "center", padding: 0, zIndex: 5,
          }}>
          {sidebarTerbuka ? <ChevronLeft size={14} /> : <ChevronRight size={14} />}
        </button>

        <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 26, padding: "0 6px", justifyContent: sidebarTerbuka ? "flex-start" : "center" }}>
          <div style={{ width: 34, height: 34, borderRadius: 9, background: "rgba(255,255,255,0.06)", border: "1px solid #2A3138", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
            <img src={LOGO_ICON} alt="GudangKu" style={{ width: 22, height: "auto" }} />
          </div>
          {sidebarTerbuka && (
            <div style={{ whiteSpace: "nowrap" }}>
              <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16, lineHeight: 1.1 }}>GudangKu</div>
              <div style={{ fontSize: 9.5, color: "#5C6570", letterSpacing: 0.4 }}>SISTEM GUDANG TERINTEGRASI</div>
            </div>
          )}
        </div>
        {NAV.map(n => (
          <button key={n.id} onClick={() => setTab(n.id)} title={sidebarTerbuka ? undefined : n.label} className={`nav-btn${tab === n.id ? " aktif" : ""}`} style={{
            display: "flex", alignItems: "center", gap: 10, padding: sidebarTerbuka ? "10px 12px" : "10px 0", borderRadius: 8, border: "none", textAlign: "left",
            background: tab === n.id ? "#232B32" : "transparent", color: tab === n.id ? "#EDEFF2" : "#8B95A1", fontWeight: tab === n.id ? 600 : 500,
            justifyContent: sidebarTerbuka ? "flex-start" : "center", whiteSpace: "nowrap",
          }}>
            <n.ikon size={17} style={{ flexShrink: 0 }} /> {sidebarTerbuka && n.label}
          </button>
        ))}

        <button onClick={() => setModalNotifikasi(true)} title={sidebarTerbuka ? undefined : `Notifikasi${stokMenipis.length ? ` (${stokMenipis.length})` : ""}`} className="nav-btn" style={{
          display: "flex", alignItems: "center", gap: 10, padding: sidebarTerbuka ? "10px 12px" : "10px 0", borderRadius: 8, border: "none", textAlign: "left",
          background: "transparent", color: stokMenipis.length ? "#E8A33D" : "#8B95A1", fontWeight: 500, marginTop: 4,
          justifyContent: sidebarTerbuka ? "flex-start" : "center", whiteSpace: "nowrap",
        }}>
          <Bell size={17} style={{ flexShrink: 0 }} /> {sidebarTerbuka && `Notifikasi ${stokMenipis.length > 0 ? `(${stokMenipis.length})` : ""}`}
        </button>
        <button onClick={() => setModalSetting(true)} title={sidebarTerbuka ? undefined : "Pengaturan"} className="nav-btn" style={{
          display: "flex", alignItems: "center", gap: 10, padding: sidebarTerbuka ? "10px 12px" : "10px 0", borderRadius: 8, border: "none", textAlign: "left",
          background: "transparent", color: "#8B95A1", fontWeight: 500, marginTop: 4,
          justifyContent: sidebarTerbuka ? "flex-start" : "center", whiteSpace: "nowrap",
        }}>
          <Settings size={17} style={{ flexShrink: 0 }} /> {sidebarTerbuka && "Pengaturan"}
        </button>

        <div style={{ marginTop: "auto", borderTop: "1px solid #2A3138", paddingTop: 14 }}>
          {sidebarTerbuka ? (
            <>
              <div style={{ fontSize: 12, fontWeight: 600, whiteSpace: "nowrap" }}>{currentUser.nama}</div>
              <div style={{ fontSize: 11, color: "#5C6570", marginBottom: 10, whiteSpace: "nowrap" }}>{currentUser.role === "admin" ? "Administrator" : "Staf"} · {currentUser.gudang}</div>
              <button onClick={() => setCurrentUser(null)} style={{ display: "flex", alignItems: "center", gap: 8, background: "none", border: "1px solid #2A3138", borderRadius: 8, padding: "8px 10px", color: "#8B95A1", width: "100%" }}>
                <LogOut size={14} /> Keluar
              </button>
            </>
          ) : (
            <button onClick={() => setCurrentUser(null)} title={`Keluar (${currentUser.nama})`} style={{ display: "flex", alignItems: "center", justifyContent: "center", background: "none", border: "1px solid #2A3138", borderRadius: 8, padding: "8px 0", color: "#8B95A1", width: "100%" }}>
              <LogOut size={16} />
            </button>
          )}
        </div>
      </div>

      {/* Konten utama */}
      <div className="area-cetak" style={{ flex: 1, padding: "26px 32px", overflowY: "auto" }}>
        {/* Top bar */}
        <div className="no-print" style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 10, marginBottom: 18 }}>
          <Building2 size={15} color="#8B95A1" />
          {isAdmin ? (
            <select value={filterGudang} onChange={e => setFilterGudang(e.target.value)} style={{ ...inputStyle, padding: "7px 10px" }}>
              <option>Semua</option>
              {GUDANG.map(g => <option key={g}>{g}</option>)}
            </select>
          ) : (
            <span style={{ fontSize: 13, color: "#8B95A1" }}>{currentUser.gudang}</span>
          )}
        </div>

        {tab === "dashboard" && (
          <div className="tab-fade" key="dashboard">
            <DashboardView nilaiTotalStok={nilaiTotalStok} totalItems={productsGudang.length} stokMenipis={stokMenipis}
              totalPendapatan={totalPendapatan} totalProfit={totalProfit} totalTransaksi={totalTransaksi}
              mutasiHariIni={mutasiHariIni} trenHarian={trenHarian} produkTerlaris={produkTerlaris}
              distribusiKategori={distribusiKategori} warnaKategoriMap={warnaKategoriMap} bisaLihatHarga={bisaLihatHarga} />
          </div>
        )}

        {tab === "inventory" && (
          <div className="tab-fade" key="inventory">
            <InventoryView produk={produkFilter} search={search} setSearch={setSearch}
              namaKategori={namaKategori} warnaKategoriMap={warnaKategoriMap}
              filterKategori={filterKategori} setFilterKategori={setFilterKategori} isAdmin={isAdmin} bisaLihatHarga={bisaLihatHarga}
              onTambah={() => { setErrorForm(""); setModalProduk("baru"); }}
              onEdit={(p) => { setErrorForm(""); setModalProduk(p); }}
              onHapus={hapusProduk}
              onRiwayat={(id) => setModalRiwayat(id)}
              onKelolaKategori={() => setModalKategori(true)} />
          </div>
        )}

        {tab === "masterdata" && (
          <div className="tab-fade" key="masterdata">
            <MasterDataView produk={products} warnaKategoriMap={warnaKategoriMap} bisaLihatHarga={bisaLihatHarga} />
          </div>
        )}

        {tab === "mutasi" && (
          <div className="tab-fade" key="mutasi">
            <MutasiView mutasi={mutasiGudang} isAdmin={isAdmin} onCatat={() => setModalMutasi(true)} onCetak={cetakInvoiceMutasi} bisaLihatHarga={bisaLihatHarga} />
          </div>
        )}

        {tab === "sales" && (
          <div className="tab-fade" key="sales">
            <SalesView sales={salesGudang} isAdmin={isAdmin} bisaLihatHarga={bisaLihatHarga}
              onJualBaru={() => setModalJual(true)}
              onEdit={(s) => setModalEditJual(s)}
              onHapus={hapusPenjualan} />
          </div>
        )}

        {tab === "reports" && (
          <div className="tab-fade" key="reports">
            <ReportsView trenHarian={trenHarian} produkTerlaris={produkTerlaris} distribusiKategori={distribusiKategori} products={productsGudang} warnaKategoriMap={warnaKategoriMap} mutasi={mutasiGudang} sales={salesGudang} isAdmin={isAdmin} bisaLihatHarga={bisaLihatHarga} />
          </div>
        )}

        {tab === "riwayat" && (
          <div className="tab-fade" key="riwayat">
            <RiwayatHargaView riwayat={riwayatHarga} products={products} produkTerpilih={modalRiwayat} isAdmin={isAdmin} onEditEntry={setModalEditRiwayat} bisaLihatHarga={bisaLihatHarga} />
          </div>
        )}
      </div>

      {modalProduk && (
        <ModalProduk data={modalProduk === "baru" ? null : modalProduk} error={errorForm} categories={categories} products={products} bisaLihatHarga={bisaLihatHarga}
          onBatal={() => setModalProduk(null)} onSimpan={simpanProduk} defaultGudang={isAdmin ? GUDANG[0] : currentUser.gudang} />
      )}

      {modalKategori && (
        <ModalKategori categories={categories} products={products} onBatal={() => setModalKategori(false)}
          onTambah={tambahKategori} onEdit={editKategori} onHapus={hapusKategori} />
      )}

      {modalJual && (
        <ModalJual products={productsGudang.length ? productsGudang : products} onBatal={() => setModalJual(false)} bisaLihatHarga={bisaLihatHarga}
          onSimpan={(produkId, jumlah, tanggal, hargaJualOverride, keterangan, transactionType, tujuanGudang) => { const err = catatPenjualan(produkId, jumlah, tanggal, hargaJualOverride, keterangan, transactionType, tujuanGudang); if (!err) setModalJual(false); return err; }} />
      )}

      {modalEditJual && (
        <ModalEditJual sale={modalEditJual} products={products} onBatal={() => setModalEditJual(null)}
          onSimpan={(produkId, jumlah, tanggal, hargaJualOverride, keterangan) => { const err = perbaikiPenjualan(modalEditJual.id, produkId, jumlah, tanggal, hargaJualOverride, keterangan); if (!err) setModalEditJual(null); return err; }} />
      )}

      {modalEditRiwayat && (
        <ModalEditRiwayat entry={modalEditRiwayat} onBatal={() => setModalEditRiwayat(null)} onSimpan={(entryId, hargaLama, hargaBaru, keterangan) => { const err = editRiwayatHarga(entryId, hargaLama, hargaBaru, keterangan); if (!err) setModalEditRiwayat(null); return err; }} />
      )}

      {modalMutasi && (
        <ModalMutasi products={products} isAdmin={isAdmin} currentUser={currentUser} onBatal={() => setModalMutasi(false)} bisaLihatHarga={bisaLihatHarga}
          onSimpan={(jenis, produkId, jumlah, tanggal, keterangan, satuanOverride, hargaSatuan, tujuanGudang) => { const err = catatMutasi(jenis, produkId, jumlah, tanggal, keterangan, satuanOverride, hargaSatuan, tujuanGudang); if (!err) setModalMutasi(false); return err; }} />
      )}

      {invoiceTampil && (
        <ModalInvoiceMutasi record={invoiceTampil} onTutup={() => setInvoiceTampil(null)} onCetak={cetakInvoiceMutasi} />
      )}

      {modalNotifikasi && (
        <ModalNotifikasi stokMenipis={stokMenipis} settings={settings} onTutup={() => setModalNotifikasi(false)} />
      )}
      {modalSetting && (
        <ModalSetting settings={settings} onUpdate={simpanSettings} onTutup={() => setModalSetting(false)} onResetData={resetSemuaData} isAdmin={isAdmin} />
      )}

        <ToastContainer toasts={toasts} onTutup={(id) => setToasts(prev => prev.filter(t => t.id !== id))} />
      </div>
    </ErrorBoundary>
  );
}

function ModalSetting({ settings, onUpdate, onTutup, onResetData, isAdmin }) {
  const [local, setLocal] = useState(settings);
  const [konfirmasiReset, setKonfirmasiReset] = useState(false);
  return (
    <Overlay onBatal={onTutup}>
      <h3 style={{ marginTop: 0, fontFamily: "'Space Grotesk', sans-serif", display: "flex", alignItems: "center", gap: 8 }}>
        <Settings size={18} color="#3FA796" /> Pengaturan Aplikasi
      </h3>
      <p style={{ color: "#8B95A1", fontSize: 12, marginTop: -8 }}>Pilih tema, nomor WhatsApp, dan pengingat stok minim otomatis.</p>

      <Grid>
        <Field label="Tema Aplikasi">
          <select value={local.theme} onChange={e => setLocal(s => ({ ...s, theme: e.target.value }))} style={inputStyle}>
            <option value="dark">Gelap</option>
            <option value="light">Terang</option>
          </select>
        </Field>
        <Field label="Nomor WhatsApp tujuan">
          <input value={local.waNumber} onChange={e => setLocal(s => ({ ...s, waNumber: e.target.value }))} placeholder="62812xxxxxxx" style={inputStyle} />
        </Field>
        <Field label="Jadwal Notifikasi Stok Minim">
          <input type="time" value={local.notifSchedule} onChange={e => setLocal(s => ({ ...s, notifSchedule: e.target.value }))} style={inputStyle} />
        </Field>
        <Field label="Kirim Notifikasi Otomatis">
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <input type="checkbox" checked={local.autoNotify} onChange={e => setLocal(s => ({ ...s, autoNotify: e.target.checked }))} />
            <span style={{ color: "#EDEFF2", fontSize: 13 }}>Aktifkan pengiriman otomatis WA</span>
          </div>
        </Field>
      </Grid>

      {isAdmin && (
        <div style={{ marginTop: 20, paddingTop: 16, borderTop: "1px solid #2A3138" }}>
          <h4 style={{ margin: "0 0 6px", fontSize: 13 }}>Tampilkan Harga per Akun Gudang</h4>
          <p style={{ fontSize: 12, color: "#8B95A1", marginTop: 0, marginBottom: 12, lineHeight: 1.5 }}>
            Pilih akun gudang mana saja yang boleh melihat & mengubah harga beli/jual. Akun yang tidak dicentang tidak akan melihat kolom harga sama sekali (di Inventori, Riwayat Harga, maupun Laporan) -- tapi data harga tetap tersimpan lengkap di server, tidak hilang. Akun admin selalu bisa melihat harga.
          </p>
          <div style={{ display: "flex", gap: 8, marginBottom: 10 }}>
            <button type="button" onClick={() => setLocal(s => ({ ...s, hargaVisibility: Object.fromEntries(GUDANG.map(g => [g, true])) }))} style={{ ...btnSecondary, fontSize: 12, padding: "6px 12px" }}>
              Tampilkan semua
            </button>
            <button type="button" onClick={() => setLocal(s => ({ ...s, hargaVisibility: Object.fromEntries(GUDANG.map(g => [g, false])) }))} style={{ ...btnSecondary, fontSize: 12, padding: "6px 12px" }}>
              Sembunyikan semua
            </button>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {GUDANG.map(g => {
              const tampil = (local.hargaVisibility || {})[g] !== false;
              return (
                <label key={g} style={{ display: "flex", alignItems: "center", gap: 10, cursor: "pointer" }}>
                  <input type="checkbox" checked={tampil} onChange={e => setLocal(s => ({ ...s, hargaVisibility: { ...(s.hargaVisibility || {}), [g]: e.target.checked } }))} />
                  <span style={{ color: "#EDEFF2", fontSize: 13 }}>{g}</span>
                  <span style={{ color: tampil ? "#3FA796" : "#8B95A1", fontSize: 11, marginLeft: "auto" }}>{tampil ? "Harga terlihat" : "Harga disembunyikan"}</span>
                </label>
              );
            })}
          </div>
        </div>
      )}

      <div style={{ fontSize: 12, color: "#8B95A1", marginTop: 16, lineHeight: 1.5 }}>
        Pengaturan ini akan menyimpan tema dan nomor WA untuk notifikasi tanpa konfirmasi manual. Jika browser memblokir pembukaan WhatsApp otomatis, gunakan tombol di halaman Notifikasi sebagai cadangan.
      </div>

      <div style={{ marginTop: 20, paddingTop: 16, borderTop: "1px solid #2A3138" }}>
        <h4 style={{ margin: "0 0 6px", fontSize: 13 }}>Muat Ulang dari Server</h4>
        <p style={{ fontSize: 12, color: "#8B95A1", marginTop: 0, lineHeight: 1.5 }}>
          Data utama (produk, mutasi, penjualan, riwayat harga) sekarang tersimpan di server Supabase dan sama untuk semua akun/gudang. Tombol ini hanya membersihkan salinan cache di perangkat ini lalu memuat ulang data terbaru dari server -- data bersama tidak akan terhapus.
        </p>
        {!konfirmasiReset ? (
          <button onClick={() => setKonfirmasiReset(true)} style={btnSecondary}>
            Muat Ulang Cache & Data
          </button>
        ) : (
          <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
            <span style={{ fontSize: 12, color: "#8B95A1" }}>Cache lokal akan dibersihkan dan halaman dimuat ulang.</span>
            <button onClick={() => setKonfirmasiReset(false)} style={btnSecondary}>Batal</button>
            <button onClick={onResetData} className="btn-primary-glow" style={btnPrimary}>Ya, Muat Ulang</button>
          </div>
        )}
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onTutup} style={btnSecondary}>Batal</button>
        <button onClick={() => { onUpdate(local); onTutup(); }} className="btn-primary-glow" style={btnPrimary}>Simpan Pengaturan</button>
      </div>
    </Overlay>
  );
}

// ============================================================
function LandingPage({ onMulai }) {
  const fitur = [
    { ikon: Boxes, judul: "Inventori Terpusat", teks: "Kelola stok semua gudang dari satu tempat, lengkap dengan kode barang otomatis dan peringatan stok menipis." },
    { ikon: ArrowRightLeft, judul: "Mutasi & Transfer Barang", teks: "Catat barang masuk, keluar, dan transfer antar gudang — semua histori tersimpan rapi dan bisa ditelusuri kapan saja." },
    { ikon: ShoppingCart, judul: "Penjualan Terintegrasi", teks: "Catat penjualan langsung terhubung ke stok dan laporan keuntungan, tanpa perlu input dua kali." },
    { ikon: BarChart3, judul: "Laporan & Analitik", teks: "Pantau tren pendapatan, profit, dan produk terlaris lewat dasbor yang selalu diperbarui." },
  ];

  return (
    <div style={{ minHeight: "100vh", background: "#14181D", color: "#EDEFF2", fontFamily: "'Inter', sans-serif" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@600;700&family=Inter:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        button { transition: all 0.15s ease; }
      `}</style>

      {/* Navigasi atas */}
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "18px 32px", borderBottom: "1px solid #2A3138" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <img src={LOGO_ICON} alt="" style={{ width: 26, height: "auto" }} />
          <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 18 }}>GudangKu</span>
        </div>
        <button onClick={onMulai} className="btn-primary-glow" style={{ ...btnPrimary, padding: "9px 20px" }}>Masuk ke Aplikasi</button>
      </div>

      {/* Hero */}
      <div style={{ textAlign: "center", padding: "72px 24px 56px", maxWidth: 720, margin: "0 auto" }}>
        <div style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "rgba(63,167,150,0.12)", color: "#3FA796", fontSize: 12, fontWeight: 600, padding: "6px 14px", borderRadius: 999, marginBottom: 20 }}>
          <Sparkles size={13} /> Sinkron real-time di semua gudang
        </div>
        <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 40, lineHeight: 1.25, margin: 0 }}>
          Satu Dasbor untuk Semua Gudang Anda
        </h1>
        <p style={{ color: "#8B95A1", fontSize: 15, lineHeight: 1.7, marginTop: 16 }}>
          Kelola stok, mutasi barang, dan penjualan lintas gudang dalam satu sistem terintegrasi.
          Setiap perubahan dari gudang manapun langsung terlihat oleh semua akun, kapan saja.
        </p>
        <button onClick={onMulai} className="btn-primary-glow" style={{ ...btnPrimary, marginTop: 28, padding: "12px 28px", fontSize: 15 }}>
          Mulai Sekarang
        </button>
      </div>

      {/* Fitur */}
      <div style={{ display: "flex", flexWrap: "wrap", gap: 16, justifyContent: "center", padding: "0 24px 72px", maxWidth: 1040, margin: "0 auto" }}>
        {fitur.map((f, i) => (
          <div key={i} style={{ flex: "1 1 220px", maxWidth: 240, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 12, padding: 22 }}>
            <div style={{ width: 38, height: 38, borderRadius: 9, background: "rgba(63,167,150,0.12)", display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 14 }}>
              <f.ikon size={19} color="#3FA796" />
            </div>
            <h3 style={{ fontSize: 14.5, margin: "0 0 8px", fontFamily: "'Space Grotesk', sans-serif" }}>{f.judul}</h3>
            <p style={{ color: "#8B95A1", fontSize: 12.5, lineHeight: 1.6, margin: 0 }}>{f.teks}</p>
          </div>
        ))}
      </div>

      <div style={{ textAlign: "center", color: "#5C6570", fontSize: 12, padding: "20px 24px 32px", borderTop: "1px solid #2A3138" }}>
        © {new Date().getFullYear()} GudangKu — Sistem Gudang Terintegrasi
      </div>
    </div>
  );
}

function LoginView({ onLogin, onKembali }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [animasiAktif, setAnimasiAktif] = useState(true);

  useEffect(() => {
    if (typeof window === "undefined") return;
    const prefersReduced = window.matchMedia?.('(prefers-reduced-motion: reduce)').matches;
    if (prefersReduced) setAnimasiAktif(false);
  }, []);

  function submit(e) {
    e.preventDefault();
    const akun = AKUN.find(a => a.username === username && a.password === password);
    if (!akun) { setError("Masa gak inget password? parah!"); return; }
    onLogin(akun);
  }

  return (
    <div style={{
      minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", position: "relative",
      overflow: "hidden", color: "#ffffff", fontFamily: "'Inter', sans-serif", background: "#0d1e20",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@600;700&family=Inter:wght@400;500;600&display=swap');
        * { box-sizing: border-box; }
        .kartu-login { animation: fadeInKartu 0.55s ease; }
        @keyframes fadeInKartu { from { opacity: 0; transform: translateY(14px) scale(0.98); } to { opacity: 1; transform: translateY(0) scale(1); } }

        .input-group { position: relative; margin-bottom: 20px; border-radius: 8px; overflow: hidden; }
        .input-group svg.ikon-input { position: absolute; top: 50%; left: 15px; transform: translateY(-50%); width: 18px; height: 18px; color: #8c9b9d; z-index: 2; }
        .input-group input {
          width: 100%; padding: 15px 15px 15px 45px; background: transparent; border: 1px solid #1DB9A0;
          border-radius: 8px; color: #ffffff; font-size: 14px; outline: none; transition: box-shadow 0.3s ease;
          position: relative; z-index: 2;
        }
        .input-group input::placeholder { color: #8c9b9d; }
        .input-group input:focus { box-shadow: 0 0 10px rgba(29,185,160,0.3); }
        .input-group::before {
          content: ''; position: absolute; top: 0; left: -150%; width: 50%; height: 100%;
          background: linear-gradient(90deg, transparent, rgba(29,185,160,0.4), transparent);
          transform: skewX(-20deg); z-index: 1; animation: lightSweep 3s infinite linear; pointer-events: none;
        }
        .input-group.kedua::before { animation-delay: 1.5s; }
        @keyframes lightSweep { 0%, 90% { left: -150%; } 100% { left: 150%; } }

        .btn-login { width: 100%; padding: 15px; background-color: #1DB9A0; color: #fff; border: none; border-radius: 8px;
          font-size: 16px; font-weight: 600; cursor: pointer; margin-top: 10px; transition: background-color 0.2s ease, transform 0.1s ease;
          box-shadow: 0 4px 15px rgba(29,185,160,0.3); }
        .btn-login:hover { background-color: #179b85; }
        .btn-login:active { transform: scale(0.98); }
      `}</style>

      <LoginAnimation enabled={animasiAktif} />

      <form onSubmit={submit} className="kartu-login" style={{
        position: "relative", zIndex: 3, width: 360, background: "rgba(255,255,255,0.03)", backdropFilter: "blur(16px)",
        WebkitBackdropFilter: "blur(16px)", border: "1px solid rgba(255,255,255,0.1)", borderRadius: 20,
        padding: "40px 30px", boxShadow: "0 25px 45px rgba(0,0,0,0.2)", textAlign: "center",
      }}>
        {onKembali && (
          <button type="button" onClick={onKembali} style={{ background: "none", border: "none", color: "#8c9b9d", fontSize: 12, padding: 0, marginBottom: 18, cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}>
            ← Kembali ke beranda
          </button>
        )}

        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 10, marginBottom: 20 }}>
          <img src={LOGO_ICON} alt="" style={{ width: 32, height: 32 }} />
          <span style={{ fontSize: 24, fontWeight: 600 }}>Gudang<span style={{ color: "#1DB9A0" }}>Ku</span></span>
        </div>
        <h2 style={{ fontSize: 18, fontWeight: 400, margin: "0 0 30px", letterSpacing: 1 }}>Masuk</h2>

        <div className="input-group">
          <svg className="ikon-input" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <path d="M20 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"></path>
            <circle cx="12" cy="7" r="4"></circle>
          </svg>
          <input type="text" value={username} onChange={e => setUsername(e.target.value)} placeholder="Username" autoComplete="off" />
        </div>

        <div className="input-group kedua">
          <svg className="ikon-input" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
            <rect x="3" y="11" width="18" height="11" rx="2" ry="2"></rect>
            <path d="M7 11V7a5 5 0 0 1 10 0v4"></path>
          </svg>
          <input type="password" value={password} onChange={e => setPassword(e.target.value)} placeholder="Password" />
        </div>

        {error && <div style={{ color: "#FF8A80", fontSize: 12, marginTop: -8, marginBottom: 12, textAlign: "left" }}>{error}</div>}

        <button type="submit" className="btn-login">Masuk</button>

        <p style={{ marginTop: 25, color: "#8c9b9d", fontSize: 12, letterSpacing: 0.5 }}>Sistem Gudang Terintegrasi</p>

        <div style={{ display: "flex", justifyContent: "center", marginTop: 10 }}>
          <label style={{ display: "flex", alignItems: "center", gap: 6, cursor: "pointer" }}>
            <input type="checkbox" checked={animasiAktif} onChange={() => setAnimasiAktif(v => !v)} style={{ width: 13, height: 13 }} />
            <span style={{ fontSize: 11, color: "#8c9b9d" }}>Animasi latar</span>
          </label>
        </div>
      </form>
    </div>
  );
}

// ============================================================
function DashboardView({ nilaiTotalStok, totalItems, stokMenipis, totalPendapatan, totalProfit, totalTransaksi, mutasiHariIni, trenHarian, produkTerlaris, distribusiKategori, warnaKategoriMap, bisaLihatHarga = true }) {
  return (
    <div>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, margin: 0, display: "flex", alignItems: "center", gap: 8 }}>
        Dashboard Operasional <Sparkles size={16} color="#3FA796" />
      </h1>
      <p style={{ color: "#8B95A1", marginTop: 4 }}>Ringkasan performa gudang dan penjualan 30 hari terakhir.</p>
      <BarcodeDivider />

      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 24 }}>
        {bisaLihatHarga && <KartuKPI ikon={Wallet} label="Nilai Stok" nilai={nilaiTotalStok} formatFn={rupiah} sub={`${totalItems} jenis produk`} warna="#3FA796" />}
        <KartuKPI ikon={AlertTriangle} label="Stok Menipis" nilai={stokMenipis.length} sub="perlu restock segera" warna="#E8A33D" />
        {bisaLihatHarga && <KartuKPI ikon={TrendingUp} label="Pendapatan" nilai={totalPendapatan} formatFn={rupiah} sub="30 hari terakhir" warna="#3FA796" />}
        {bisaLihatHarga && <KartuKPI ikon={ShoppingCart} label="Profit Kotor" nilai={totalProfit} formatFn={rupiah} sub={`${totalTransaksi} transaksi`} warna="#F2C14E" />}
        <KartuKPI ikon={ArrowRightLeft} label="Mutasi Hari Ini" nilai={mutasiHariIni} sub="barang masuk & keluar" warna="#6C8EBF" />
      </div>

      {bisaLihatHarga && (
        <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
          <div style={{ flex: 2, minWidth: 420, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, padding: 18 }}>
            <h3 style={judulKartu}>Tren Pendapatan & Profit Harian</h3>
            <ResponsiveContainer width="100%" height={260}>
              <AreaChart data={trenHarian}>
                <defs>
                  <linearGradient id="gPendapatan" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#3FA796" stopOpacity={0.4} />
                    <stop offset="100%" stopColor="#3FA796" stopOpacity={0} />
                  </linearGradient>
                </defs>
                <CartesianGrid strokeDasharray="3 3" stroke="#2A3138" />
                <XAxis dataKey="label" stroke="#5C6570" fontSize={11} />
                <YAxis stroke="#5C6570" fontSize={11} tickFormatter={(v) => `${(v / 1000).toFixed(0)}rb`} />
                <Tooltip contentStyle={tooltipStyle} formatter={(v) => rupiah(v)} labelStyle={{ color: "#EDEFF2" }} />
                <Area type="monotone" dataKey="pendapatan" stroke="#3FA796" fill="url(#gPendapatan)" strokeWidth={2} name="Pendapatan" />
                <Line type="monotone" dataKey="profit" stroke="#F2C14E" strokeWidth={2} dot={false} name="Profit" />
              </AreaChart>
            </ResponsiveContainer>
          </div>

          <div style={{ flex: 1, minWidth: 260, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, padding: 18 }}>
            <h3 style={judulKartu}>Nilai Stok per Kategori</h3>
            <ResponsiveContainer width="100%" height={260}>
              <PieChart>
                <Pie data={distribusiKategori} dataKey="nilai" nameKey="kategori" innerRadius={55} outerRadius={85} paddingAngle={2}>
                  {distribusiKategori.map((d, i) => <Cell key={i} fill={warnaKategoriMap[d.kategori] || "#3FA796"} />)}
                </Pie>
                <Tooltip contentStyle={tooltipStyle} formatter={(v) => rupiah(v)} />
                <Legend wrapperStyle={{ fontSize: 11, color: "#8B95A1" }} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      <div style={{ display: "flex", gap: 16, marginTop: 16, flexWrap: "wrap" }}>
        <div style={{ flex: 1, minWidth: 320, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, padding: 18 }}>
          <h3 style={judulKartu}>Produk Terlaris</h3>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={produkTerlaris} layout="vertical" margin={{ left: 10 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#2A3138" horizontal={false} />
              <XAxis type="number" stroke="#5C6570" fontSize={11} />
              <YAxis type="category" dataKey="nama" stroke="#5C6570" fontSize={11} width={140} />
              <Tooltip contentStyle={tooltipStyle} />
              <Bar dataKey="jumlah" fill="#3FA796" radius={[0, 4, 4, 0]} name="Unit terjual" />
            </BarChart>
          </ResponsiveContainer>
        </div>

        <div style={{ flex: 1, minWidth: 320, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, padding: 18 }}>
          <h3 style={judulKartu}>Peringatan Stok Menipis</h3>
          {stokMenipis.length === 0 ? (
            <p style={{ color: "#8B95A1", fontSize: 13 }}>Tidak ada produk dengan stok kritis saat ini.</p>
          ) : (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              {stokMenipis.map(p => (
                <div key={p.id} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", fontSize: 13 }}>
                  <div>
                    <div style={{ fontWeight: 500 }}>{p.nama}</div>
                    <div style={{ color: "#5C6570", fontSize: 11, fontFamily: "'IBM Plex Mono', monospace" }}>{p.kodeBarang}</div>
                  </div>
                  <IndikatorStok stok={p.stok} stokMin={p.stokMin} satuan={p.satuan} />
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

const judulKartu = { margin: "0 0 12px", fontSize: 13, color: "#8B95A1", fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.4 };
const tooltipStyle = { background: "#1D2329", border: "1px solid #2A3138", borderRadius: 8 };

// ============================================================
function InventoryView({ produk, search, setSearch, namaKategori, warnaKategoriMap, filterKategori, setFilterKategori, isAdmin, bisaLihatHarga, onTambah, onEdit, onHapus, onRiwayat, onKelolaKategori }) {
  const [exportFrom, setExportFrom] = useState("");
  const [exportTo, setExportTo] = useState("");
  const [exportKategori, setExportKategori] = useState("Semua");
  const [exportGudang, setExportGudang] = useState("Semua");

  const dataUntukEkspor = produk.filter(p => {
    if (exportKategori !== "Semua" && p.kategori !== exportKategori) return false;
    if (isAdmin && exportGudang !== "Semua" && p.gudang !== exportGudang) return false;
    if (exportFrom && p.createdAt && p.createdAt < `${exportFrom}T00:00:00.000Z`) return false;
    if (exportTo && p.createdAt && p.createdAt > `${exportTo}T23:59:59.999Z`) return false;
    return true;
  });

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
        <div>
          <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, margin: 0 }}>Inventori Gudang</h1>
          <p style={{ color: "#8B95A1", marginTop: 4 }}>Kelola stok, harga, satuan, dan status ketersediaan produk.</p>
        </div>
        <div className="no-print" style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <TombolEkspor
            onExcel={() => eksporExcel([{ nama: "Inventori", data: dataUntukEkspor.map(p => ({ "Tanggal": tanggalID(p.createdAt), "Kode Barang": p.kodeBarang, Nama: p.nama, Kategori: p.kategori, Gudang: p.gudang, Satuan: p.satuan, Stok: p.stok, "Stok Minimum": p.stokMin, ...(bisaLihatHarga ? { "Harga Beli": p.hargaBeli, "Harga Jual": p.hargaJual, "Nilai Stok": p.stok * p.hargaBeli } : {}) })) }], "inventori")}
            onCSV={() => eksporCSV(dataUntukEkspor.map(p => ({ Tanggal: tanggalID(p.createdAt), KodeBarang: p.kodeBarang, Nama: p.nama, Kategori: p.kategori, Gudang: p.gudang, Satuan: p.satuan, Stok: p.stok, StokMinimum: p.stokMin, ...(bisaLihatHarga ? { HargaBeli: p.hargaBeli, HargaJual: p.hargaJual, NilaiStok: p.stok * p.hargaBeli } : {}) })), "inventori")}
            onPDF={() => window.print()}
          />
          {isAdmin && (
            <button onClick={onKelolaKategori} style={{ ...btnSecondary, display: "flex", alignItems: "center", gap: 6 }}>
              <Tags size={15} /> Kelola Kategori
            </button>
          )}
          {isAdmin && (
            <button onClick={onTambah} className="btn-primary-glow" style={{ ...btnPrimary, display: "flex", alignItems: "center", gap: 6 }}>
              <Plus size={16} /> Tambah Produk
            </button>
          )}
        </div>
      </div>
      <div className="no-print" style={{ display: "flex", flexWrap: "wrap", gap: 10, marginTop: 14, marginBottom: 16 }}>
        <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
          <label style={{ fontSize: 12, color: "#8B95A1", display: "flex", flexDirection: "column", gap: 6 }}>
            Dari tanggal
            <input type="date" value={exportFrom} onChange={e => setExportFrom(e.target.value)} style={inputStyle} />
          </label>
          <label style={{ fontSize: 12, color: "#8B95A1", display: "flex", flexDirection: "column", gap: 6 }}>
            Sampai tanggal
            <input type="date" value={exportTo} onChange={e => setExportTo(e.target.value)} style={inputStyle} />
          </label>
          <label style={{ fontSize: 12, color: "#8B95A1", display: "flex", flexDirection: "column", gap: 6 }}>
            Kategori ekspor
            <select value={exportKategori} onChange={e => setExportKategori(e.target.value)} style={inputStyle}>
              <option>Semua</option>
              {namaKategori.map(k => <option key={k}>{k}</option>)}
            </select>
          </label>
          {isAdmin && (
            <label style={{ fontSize: 12, color: "#8B95A1", display: "flex", flexDirection: "column", gap: 6 }}>
              Gudang ekspor
              <select value={exportGudang} onChange={e => setExportGudang(e.target.value)} style={inputStyle}>
                <option>Semua</option>
                {GUDANG.map(g => <option key={g}>{g}</option>)}
              </select>
            </label>
          )}
        </div>
      </div>
      <BarcodeDivider />

      <div className="no-print" style={{ display: "flex", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 8, padding: "8px 12px", flex: 1, maxWidth: 320 }}>
          <Search size={15} color="#5C6570" />
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Cari nama atau kode barang..." style={{ background: "transparent", border: "none", outline: "none", color: "#EDEFF2", width: "100%" }} />
        </div>
        <select value={filterKategori} onChange={e => setFilterKategori(e.target.value)} style={inputStyle}>
          <option>Semua</option>
          {namaKategori.map(k => <option key={k}>{k}</option>)}
        </select>
      </div>

      <div style={{ background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, overflow: "hidden", overflowX: "auto" }}>
        <table>
          <thead>
            <tr style={{ background: "#171B20", color: "#8B95A1", fontSize: 12, textTransform: "uppercase" }}>
              <th>Tanggal</th><th>Kode Barang</th><th>Produk</th><th>Kategori</th><th>Satuan</th><th>Gudang</th><th>Status Stok</th>
              {bisaLihatHarga && <><th>Harga Beli</th><th>Harga Jual</th><th>Nilai Stok</th></>}
              <th className="no-print"></th>
            </tr>
          </thead>
          <tbody>
            {produk.map(p => (
              <tr key={p.id} style={{ borderTop: "1px solid #2A3138", fontSize: 13 }}>
                <td style={{ color: "#8B95A1" }}>{tanggalID(p.createdAt)}</td>
                <td style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#8B95A1" }}>{p.kodeBarang}</td>
                <td><div style={{ fontWeight: 500 }}>{p.nama}</div></td>
                <td><Badge warna={warnaKategoriMap[p.kategori]}>{p.kategori}</Badge></td>
                <td style={{ fontSize: 12, color: "#8B95A1" }}>{p.satuan}</td>
                <td><Badge warna={GUDANG_WARNA[p.gudang]}>{p.gudang}</Badge></td>
                <td><IndikatorStok stok={p.stok} stokMin={p.stokMin} satuan={p.satuan} /></td>
                {bisaLihatHarga && <>
                  <td style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{rupiah(p.hargaBeli)}</td>
                  <td style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{rupiah(p.hargaJual)}</td>
                  <td style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{rupiah(p.stok * p.hargaBeli)}</td>
                </>}
                <td className="no-print">
                  <div style={{ display: "flex", gap: 6 }}>
                    {bisaLihatHarga && <button onClick={() => onRiwayat(p.id)} title="Riwayat harga" className="icon-btn-hover" style={iconBtn}><History size={13} /></button>}
                    {isAdmin && <button onClick={() => onEdit(p)} title="Edit" className="icon-btn-hover" style={iconBtn}><Pencil size={13} /></button>}
                    {isAdmin && <button onClick={() => onHapus(p.id)} title="Hapus" className="icon-btn-hover" style={{ ...iconBtn, color: "#E2574C" }}><Trash2 size={13} /></button>}
                  </div>
                </td>
              </tr>
            ))}
            {produk.length === 0 && <tr><td colSpan={bisaLihatHarga ? 11 : 8} style={{ textAlign: "center", padding: 30, color: "#5C6570" }}>Tidak ada produk yang cocok dengan pencarian.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

const iconBtn = { background: "none", border: "1px solid #2A3138", borderRadius: 6, padding: 6, color: "#8B95A1" };

// ============================================================
// Master Data: daftar seluruh barang yang sudah masuk dan belum keluar (stok saat ini)
// dari SEMUA gudang, terlihat sama oleh semua akun -- bukan cuma gudang milik akun yang login.
function MasterDataView({ produk, warnaKategoriMap, bisaLihatHarga }) {
  const [search, setSearch] = useState("");
  const [filterKategori, setFilterKategori] = useState("Semua");
  const [filterGudang, setFilterGudang] = useState("Semua");

  const namaKategori = useMemo(() => [...new Set(produk.map(p => p.kategori))], [produk]);

  const daftar = useMemo(() => {
    return produk.filter(p => {
      if (filterKategori !== "Semua" && p.kategori !== filterKategori) return false;
      if (filterGudang !== "Semua" && p.gudang !== filterGudang) return false;
      if (search) {
        const q = search.toLowerCase();
        if (![p.nama, p.kodeBarang, p.supplier].some(v => String(v || "").toLowerCase().includes(q))) return false;
      }
      return true;
    }).sort((a, b) => a.nama.localeCompare(b.nama));
  }, [produk, filterKategori, filterGudang, search]);

  const totalStokUnit = daftar.reduce((s, p) => s + p.stok, 0);
  const jumlahGudangTerpakai = new Set(daftar.map(p => p.gudang)).size;
  const stokMenipisCount = daftar.filter(p => p.stok <= p.stokMin).length;

  return (
    <div>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, margin: 0 }}>Master Data Barang</h1>
      <p style={{ color: "#8B95A1", marginTop: 4 }}>Daftar seluruh barang yang sudah masuk dan belum keluar (stok saat ini) di semua gudang -- terlihat sama untuk semua akun.</p>
      <BarcodeDivider />

      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 18 }}>
        <KartuKPI ikon={Database} label="Jenis Barang" nilai={daftar.length} sub="tercatat di master data" warna="#3FA796" />
        <KartuKPI ikon={Boxes} label="Total Unit Stok" nilai={totalStokUnit} sub="seluruh satuan digabung" warna="#6C8EBF" />
        <KartuKPI ikon={Building2} label="Gudang Terpakai" nilai={jumlahGudangTerpakai} sub={`dari ${GUDANG.length} gudang`} warna="#F2C14E" />
        <KartuKPI ikon={AlertTriangle} label="Stok Menipis" nilai={stokMenipisCount} sub="perlu restock segera" warna="#E8A33D" />
      </div>

      <div className="no-print" style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 8, padding: "8px 12px", flex: 1, maxWidth: 320 }}>
          <Search size={15} color="#5C6570" />
          <input value={search} onChange={e => setSearch(e.target.value)} placeholder="Cari nama, kode, atau supplier..." style={{ background: "transparent", border: "none", outline: "none", color: "#EDEFF2", width: "100%" }} />
        </div>
        <select value={filterKategori} onChange={e => setFilterKategori(e.target.value)} style={inputStyle}>
          <option>Semua</option>
          {namaKategori.map(k => <option key={k}>{k}</option>)}
        </select>
        <select value={filterGudang} onChange={e => setFilterGudang(e.target.value)} style={inputStyle}>
          <option>Semua</option>
          {GUDANG.map(g => <option key={g}>{g}</option>)}
        </select>
      </div>

      <div style={{ background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, overflow: "hidden", overflowX: "auto" }}>
        <table>
          <thead>
            <tr style={{ background: "#171B20", color: "#8B95A1", fontSize: 12, textTransform: "uppercase" }}>
              <th>Kode Barang</th><th>Produk</th><th>Kategori</th><th>Gudang</th><th>Stok Saat Ini</th><th>Status</th><th>Supplier</th>
              {bisaLihatHarga && <th>Nilai Stok</th>}
            </tr>
          </thead>
          <tbody>
            {daftar.map(p => (
              <tr key={p.id} style={{ borderTop: "1px solid #2A3138", fontSize: 13 }}>
                <td style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#8B95A1" }}>{p.kodeBarang}</td>
                <td><div style={{ fontWeight: 500 }}>{p.nama}</div></td>
                <td><Badge warna={warnaKategoriMap[p.kategori]}>{p.kategori}</Badge></td>
                <td><Badge warna={GUDANG_WARNA[p.gudang]}>{p.gudang}</Badge></td>
                <td style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{p.stok} {satuanSingkat(p.satuan)}</td>
                <td><IndikatorStok stok={p.stok} stokMin={p.stokMin} satuan={p.satuan} /></td>
                <td style={{ color: "#8B95A1", fontSize: 12 }}>{p.supplier || "-"}</td>
                {bisaLihatHarga && <td style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{rupiah(p.stok * p.hargaBeli)}</td>}
              </tr>
            ))}
            {daftar.length === 0 && <tr><td colSpan={bisaLihatHarga ? 8 : 7} style={{ textAlign: "center", padding: 30, color: "#5C6570" }}>Tidak ada barang yang cocok dengan pencarian.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================
function MutasiView({ mutasi, onCatat, onCetak, isAdmin }) {
  const [filterJenis, setFilterJenis] = useState("Semua");
  const [filterKategori, setFilterKategori] = useState("Semua");
  const [filterGudang, setFilterGudang] = useState("Semua");
  const [searchQuery, setSearchQuery] = useState("");
  const [tanggalMulai, setTanggalMulai] = useState("");
  const [tanggalSelesai, setTanggalSelesai] = useState("");
  const daftar = useMemo(() => {
    return mutasi.filter(m => {
      if (filterJenis !== "Semua" && m.jenis !== filterJenis) return false;
      if (filterKategori !== "Semua" && (m.kategori || "Lainnya") !== filterKategori) return false;
      if (filterGudang !== "Semua" && m.gudang !== filterGudang) return false;
      if (tanggalMulai && m.tanggal < `${tanggalMulai}T00:00:00.000Z`) return false;
      if (tanggalSelesai && m.tanggal > `${tanggalSelesai}T23:59:59.999Z`) return false;
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        if (![m.kode, m.namaProduk, m.kodeBarang, m.gudang, m.keterangan].some(v => String(v || "").toLowerCase().includes(q))) return false;
      }
      return true;
    });
  }, [mutasi, filterJenis, filterKategori, filterGudang, searchQuery, tanggalMulai, tanggalSelesai]);
  const totalMasuk = mutasi.filter(m => m.jenis === "Masuk").reduce((s, m) => s + m.jumlah, 0);
  const totalKeluar = mutasi.filter(m => m.jenis === "Keluar").reduce((s, m) => s + m.jumlah, 0);
  const totalTransfer = mutasi.filter(m => m.jenis === "Transfer").reduce((s, m) => s + m.jumlah, 0);

  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 12 }}>
        <div>
          <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, margin: 0 }}>Mutasi Barang Masuk & Keluar</h1>
          <p style={{ color: "#8B95A1", marginTop: 4 }}>Catat pergerakan stok untuk gudang manapun. Setiap transaksi otomatis mendapat kode invoice unik.</p>
        </div>
        <div className="no-print" style={{ display: "flex", gap: 10 }}>
          <TombolEkspor
            onExcel={() => eksporExcel([{ nama: "Mutasi", data: daftar.map(m => ({ Kode: m.kode, Tanggal: tanggalWaktuID(m.tanggal), Jenis: m.jenis, "Kode Barang": m.kodeBarang, Produk: m.namaProduk, Gudang: m.gudang, Jumlah: m.jumlah, Satuan: m.satuan, "Stok Sebelum": m.stokSebelum, "Stok Sesudah": m.stokSesudah, Keterangan: m.keterangan, Oleh: m.oleh })) }], "mutasi-barang")}
            onCSV={() => eksporCSV(daftar.map(m => ({ Kode: m.kode, Tanggal: tanggalID(m.tanggal), Jenis: m.jenis, Produk: m.namaProduk, Gudang: m.gudang, Jumlah: m.jumlah, Keterangan: m.keterangan })), "mutasi-barang")}
            onPDF={() => window.print()}
          />
          <button onClick={onCatat} className="btn-primary-glow" style={{ ...btnPrimary, display: "flex", alignItems: "center", gap: 6 }}>
            <ArrowRightLeft size={16} /> Catat Mutasi
          </button>
        </div>
      </div>
      <BarcodeDivider />

      <div style={{ display: "flex", gap: 16, flexWrap: "wrap", marginBottom: 18 }}>
        <KartuKPI ikon={PackagePlus} label="Total Barang Masuk" nilai={totalMasuk} sub="seluruh histori" warna="#3FA796" />
        <KartuKPI ikon={PackageMinus} label="Total Barang Keluar" nilai={totalKeluar} sub="seluruh histori" warna="#E2574C" />
        <KartuKPI ikon={ArrowRightLeft} label="Total Transfer" nilai={totalTransfer} sub="stok antar gudang" warna="#6C8EBF" />
        <KartuKPI ikon={FileText} label="Jumlah Transaksi" nilai={mutasi.length} sub="tercatat dengan kode invoice" warna="#6C8EBF" />
      </div>

      <div className="no-print" style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 14 }}>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {["Semua", "Masuk", "Keluar", "Transfer"].map(j => (
            <button key={j} onClick={() => setFilterJenis(j)} className="jenis-pill" style={{
              border: "1px solid #2A3138", borderRadius: 20, padding: "6px 16px", fontSize: 12,
              background: filterJenis === j ? "#232B32" : "transparent", color: filterJenis === j ? "#EDEFF2" : "#8B95A1", fontWeight: 600
            }}>{j}</button>
          ))}
        </div>
        <input value={searchQuery} onChange={e => setSearchQuery(e.target.value)} placeholder="Cari kode, produk, gudang..." style={{ ...inputStyle, minWidth: 240, width: 260 }} />
        <select value={filterKategori} onChange={e => setFilterKategori(e.target.value)} style={{ ...inputStyle, width: 180 }}>
          <option value="Semua">Semua Kategori</option>
          {[...new Set(mutasi.map(m => m.kategori || "Lainnya"))].map(k => <option key={k} value={k}>{k}</option>)}
        </select>
        {isAdmin && (
          <select value={filterGudang} onChange={e => setFilterGudang(e.target.value)} style={{ ...inputStyle, width: 180 }}>
            <option value="Semua">Semua Gudang</option>
            {GUDANG.map(g => <option key={g} value={g}>{g}</option>)}
          </select>
        )}
        <input type="date" value={tanggalMulai} onChange={e => setTanggalMulai(e.target.value)} style={{ ...inputStyle, minWidth: 150 }} />
        <input type="date" value={tanggalSelesai} onChange={e => setTanggalSelesai(e.target.value)} style={{ ...inputStyle, minWidth: 150 }} />
      </div>

      <div style={{ background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, overflow: "hidden", overflowX: "auto", maxHeight: 560, overflowY: "auto" }}>
        <table>
          <thead>
            <tr style={{ background: "#171B20", color: "#8B95A1", fontSize: 12, textTransform: "uppercase", position: "sticky", top: 0 }}>
              <th>Kode Invoice</th><th>Tanggal</th><th>Jenis</th><th>Produk</th><th>Gudang</th><th>Jumlah</th><th>Keterangan</th><th>Oleh</th><th className="no-print"></th>
            </tr>
          </thead>
          <tbody>
            {daftar.map(m => (
              <tr key={m.id} style={{ borderTop: "1px solid #2A3138", fontSize: 13 }}>
                <td style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#8B95A1" }}>{m.kode}</td>
                <td style={{ color: "#8B95A1" }}>{tanggalID(m.tanggal)}</td>
                <td>
                  <span style={{ display: "inline-flex", alignItems: "center", gap: 4, fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 20, color: m.jenis === "Masuk" ? "#3FA796" : "#E2574C", background: m.jenis === "Masuk" ? "rgba(63,167,150,0.12)" : "rgba(226,87,76,0.12)" }}>
                    {m.jenis === "Masuk" ? <PackagePlus size={12} /> : <PackageMinus size={12} />} {m.jenis}
                  </span>
                </td>
                <td><div style={{ fontWeight: 500 }}>{m.namaProduk}</div><div style={{ color: "#5C6570", fontSize: 11, fontFamily: "'IBM Plex Mono', monospace" }}>{m.kodeBarang}</div></td>
                <td><Badge warna={GUDANG_WARNA[m.gudang]}>{m.gudang}</Badge></td>
                <td style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{m.jumlah} {satuanSingkat(m.satuan)}</td>
                <td style={{ color: "#8B95A1", fontSize: 12, maxWidth: 200 }}>{m.keterangan || "-"}</td>
                <td style={{ color: "#8B95A1", fontSize: 12 }}>{m.oleh}</td>
                <td className="no-print">
                  <button onClick={() => onCetak(m)} title="Cetak invoice" className="icon-btn-hover" style={iconBtn}><Printer size={13} /></button>
                </td>
              </tr>
            ))}
            {daftar.length === 0 && <tr><td colSpan={9} style={{ textAlign: "center", padding: 30, color: "#5C6570" }}>Belum ada mutasi barang yang tercatat.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================
function SalesView({ sales, isAdmin, onJualBaru, onEdit, onHapus, bisaLihatHarga = true }) {
  return (
    <div>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div>
          <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, margin: 0 }}>Riwayat Penjualan</h1>
          <p style={{ color: "#8B95A1", marginTop: 4 }}>Setiap transaksi otomatis mengurangi stok gudang. {isAdmin && "Admin dapat mengedit atau menghapus transaksi yang salah input."}</p>
        </div>
        <div className="no-print" style={{ display: "flex", gap: 10 }}>
          <TombolEkspor
            onExcel={() => eksporExcel([{ nama: "Penjualan", data: sales.map(s => ({ Tanggal: tanggalID(s.tanggal), "Kode Barang": s.kodeBarang, Produk: s.namaProduk, Gudang: s.gudang, Qty: s.jumlah, ...(bisaLihatHarga ? { "Harga Jual": s.hargaJualSaat, Total: s.total, Profit: s.profit } : {}) })) }], "penjualan")}
            onCSV={() => eksporCSV(sales.map(s => ({ Tanggal: tanggalID(s.tanggal), KodeBarang: s.kodeBarang, Produk: s.namaProduk, Gudang: s.gudang, Qty: s.jumlah, ...(bisaLihatHarga ? { HargaJual: s.hargaJualSaat, Total: s.total, Profit: s.profit } : {}) })), "penjualan")}
            onPDF={() => window.print()}
          />
          <button onClick={onJualBaru} className="btn-primary-glow" style={{ ...btnPrimary, display: "flex", alignItems: "center", gap: 6 }}><Plus size={16} /> Catat Penjualan</button>
        </div>
      </div>
      <BarcodeDivider />

      <div style={{ background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, overflow: "hidden", maxHeight: 560, overflowY: "auto" }}>
        <table>
          <thead>
            <tr style={{ background: "#171B20", color: "#8B95A1", fontSize: 12, textTransform: "uppercase", position: "sticky", top: 0 }}>
              <th>Tanggal</th><th>Produk</th><th>Kode Barang</th><th>Gudang</th><th>Qty</th>
              {bisaLihatHarga && <><th>Harga Jual</th><th>Total</th><th>Profit</th></>}
              <th className="no-print"></th>
            </tr>
          </thead>
          <tbody>
            {sales.map(s => (
              <tr key={s.id} style={{ borderTop: "1px solid #2A3138", fontSize: 13 }}>
                <td style={{ color: "#8B95A1" }}>{tanggalID(s.tanggal)}</td>
                <td>{s.namaProduk}</td>
                <td style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#8B95A1" }}>{s.kodeBarang}</td>
                <td><Badge warna={GUDANG_WARNA[s.gudang]}>{s.gudang}</Badge></td>
                <td style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{s.jumlah}</td>
                {bisaLihatHarga && <>
                  <td style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{rupiah(s.hargaJualSaat)}</td>
                  <td style={{ fontFamily: "'IBM Plex Mono', monospace", fontWeight: 600 }}>{rupiah(s.total)}</td>
                  <td style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#3FA796" }}>{rupiah(s.profit)}</td>
                </>}
                <td className="no-print">
                  {isAdmin && (
                    <div style={{ display: "flex", gap: 6 }}>
                      <button onClick={() => onEdit(s)} title="Edit transaksi" className="icon-btn-hover" style={iconBtn}><Pencil size={13} /></button>
                      <button onClick={() => onHapus(s.id)} title="Hapus transaksi" className="icon-btn-hover" style={{ ...iconBtn, color: "#E2574C" }}><Trash2 size={13} /></button>
                    </div>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================
function ReportsView({ trenHarian, distribusiKategori, products, warnaKategoriMap, mutasi, sales, isAdmin, bisaLihatHarga = true }) {
  const [filterKategori, setFilterKategori] = useState("Semua");
  const [filterDateFrom, setFilterDateFrom] = useState("");
  const [filterDateTo, setFilterDateTo] = useState("");
  const [grafikMode, setGrafikMode] = useState(bisaLihatHarga ? "Pendapatan & Profit" : "Barang Masuk");

  const produkUntukMargin = useMemo(() => {
    return products.filter(p => filterKategori === "Semua" || p.kategori === filterKategori);
  }, [products, filterKategori]);

  const margin = useMemo(() => produkUntukMargin
    .map(p => ({ nama: p.nama, margin: Math.round(((p.hargaJual - p.hargaBeli) / p.hargaJual) * 100) }))
    .sort((a, b) => b.margin - a.margin).slice(0, 8), [produkUntukMargin]);

  const mutasiTersaring = useMemo(() => {
    return mutasi.filter(m => {
      if (filterKategori !== "Semua" && m.kategori !== filterKategori) return false;
      if (filterDateFrom && m.tanggal < `${filterDateFrom}T00:00:00.000Z`) return false;
      if (filterDateTo && m.tanggal > `${filterDateTo}T23:59:59.999Z`) return false;
      return true;
    });
  }, [mutasi, filterKategori, filterDateFrom, filterDateTo]);

  const chartData = useMemo(() => {
    if (grafikMode === "Pendapatan & Profit") {
      return trenHarian.filter(item => {
        const labelDate = new Date(item.label);
        if (filterDateFrom && labelDate < new Date(`${filterDateFrom}T00:00:00`)) return false;
        if (filterDateTo && labelDate > new Date(`${filterDateTo}T23:59:59`)) return false;
        return true;
      });
    }

    if (grafikMode === "Barang Masuk" || grafikMode === "Barang Keluar") {
      const jenis = grafikMode === "Barang Masuk" ? "Masuk" : "Keluar";
      const peta = {};
      mutasiTersaring.filter(m => m.jenis === jenis).forEach(m => {
        const tanggal = tanggalID(m.tanggal);
        peta[tanggal] = (peta[tanggal] || 0) + m.jumlah;
      });
      return Object.entries(peta).map(([tanggal, jumlah]) => ({ tanggal, jumlah })).sort((a, b) => new Date(a.tanggal) - new Date(b.tanggal));
    }

    if (grafikMode === "Aset Mengendap") {
      const nilaiProduk = products.filter(p => filterKategori === "Semua" || p.kategori === filterKategori)
        .map(p => ({ kategori: p.kategori, nilai: p.stok * p.hargaBeli }));
      const peta = {};
      nilaiProduk.forEach(item => { peta[item.kategori] = (peta[item.kategori] || 0) + item.nilai; });
      return Object.entries(peta).map(([kategori, nilai]) => ({ kategori, nilai }));
    }

    return [];
  }, [grafikMode, trenHarian, mutasiTersaring, products, filterKategori, filterDateFrom, filterDateTo]);

  const filteredProducts = useMemo(() => products.filter(p => filterKategori === "Semua" || p.kategori === filterKategori), [products, filterKategori]);

  return (
    <div>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, margin: 0 }}>Laporan & Analitik</h1>
      <p style={{ color: "#8B95A1", marginTop: 4 }}>Analisis mendalam performa produk, mutasi, dan aset stock.</p>
      <BarcodeDivider />

      <div className="no-print" style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
        <select value={grafikMode} onChange={e => setGrafikMode(e.target.value)} style={{ ...inputStyle, minWidth: 200 }}>
          {bisaLihatHarga && <option>Pendapatan & Profit</option>}
          <option>Barang Masuk</option>
          <option>Barang Keluar</option>
          {bisaLihatHarga && <option>Aset Mengendap</option>}
        </select>
        <select value={filterKategori} onChange={e => setFilterKategori(e.target.value)} style={{ ...inputStyle, minWidth: 180 }}>
          <option>Semua</option>
          {[...new Set(products.map(p => p.kategori))].map(k => <option key={k}>{k}</option>)}
        </select>
        <input type="date" value={filterDateFrom} onChange={e => setFilterDateFrom(e.target.value)} style={{ ...inputStyle, minWidth: 150 }} />
        <input type="date" value={filterDateTo} onChange={e => setFilterDateTo(e.target.value)} style={{ ...inputStyle, minWidth: 150 }} />
      </div>

      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
        {bisaLihatHarga && (
          <div style={{ flex: 1, minWidth: 420, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, padding: 18 }}>
            <h3 style={judulKartu}>Margin Keuntungan Tertinggi (%)</h3>
            <ResponsiveContainer width="100%" height={280}>
              <BarChart data={margin}>
                <CartesianGrid strokeDasharray="3 3" stroke="#2A3138" />
                <XAxis dataKey="nama" stroke="#5C6570" fontSize={10} angle={-20} textAnchor="end" height={70} />
                <YAxis stroke="#5C6570" fontSize={11} tickFormatter={(v) => `${v}%`} />
                <Tooltip contentStyle={tooltipStyle} formatter={(v) => `${v}%`} />
                <Bar dataKey="margin" fill="#F2C14E" radius={[4, 4, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        )}

        <div style={{ flex: 1, minWidth: 420, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, padding: 18 }}>
          <h3 style={judulKartu}>{grafikMode}</h3>
          <ResponsiveContainer width="100%" height={280}>
            {grafikMode === "Pendapatan & Profit" ? (
              <LineChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#2A3138" />
                <XAxis dataKey="label" stroke="#5C6570" fontSize={10} />
                <YAxis stroke="#5C6570" fontSize={11} tickFormatter={(v) => `${(v / 1000).toFixed(0)}rb`} />
                <Tooltip contentStyle={tooltipStyle} formatter={(v) => rupiah(v)} />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Line type="monotone" dataKey="pendapatan" stroke="#3FA796" strokeWidth={2} dot={false} name="Pendapatan" />
                <Line type="monotone" dataKey="profit" stroke="#F2C14E" strokeWidth={2} dot={false} name="Profit" />
              </LineChart>
            ) : grafikMode === "Aset Mengendap" ? (
              <BarChart data={chartData} layout="vertical" margin={{ left: 20 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#2A3138" horizontal={false} />
                <XAxis type="number" stroke="#5C6570" fontSize={11} tickFormatter={(v) => `${(v / 1000000).toFixed(1)}jt`} />
                <YAxis type="category" dataKey="kategori" stroke="#5C6570" fontSize={12} width={110} />
                <Tooltip contentStyle={tooltipStyle} formatter={(v) => rupiah(v)} />
                <Bar dataKey="nilai" radius={[0, 4, 4, 0]}>
                  {chartData.map((d, i) => <Cell key={i} fill={warnaKategoriMap[d.kategori] || "#3FA796"} />)}
                </Bar>
              </BarChart>
            ) : (
              <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#2A3138" />
                <XAxis dataKey="tanggal" stroke="#5C6570" fontSize={10} />
                <YAxis stroke="#5C6570" fontSize={11} />
                <Tooltip contentStyle={tooltipStyle} />
                <Bar dataKey="jumlah" fill={grafikMode === "Barang Masuk" ? "#3FA796" : "#E2574C"} radius={[4, 4, 0, 0]} name={grafikMode === "Barang Masuk" ? "Unit masuk" : "Unit keluar"} />
              </BarChart>
            )}
          </ResponsiveContainer>
          {chartData.length === 0 && (
            <p style={{ color: "#5C6570", fontSize: 12, textAlign: "center", marginTop: 8 }}>Tidak ada data untuk kombinasi filter ini.</p>
          )}
        </div>
      </div>

      {bisaLihatHarga && (
        <div style={{ marginTop: 16, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, padding: 18 }}>
          <h3 style={judulKartu}>Kontribusi Nilai Stok per Kategori</h3>
          <ResponsiveContainer width="100%" height={240}>
            <BarChart data={distribusiKategori} layout="vertical" margin={{ left: 20 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#2A3138" horizontal={false} />
              <XAxis type="number" stroke="#5C6570" fontSize={11} tickFormatter={(v) => `${(v / 1000000).toFixed(1)}jt`} />
              <YAxis type="category" dataKey="kategori" stroke="#5C6570" fontSize={12} width={110} />
              <Tooltip contentStyle={tooltipStyle} formatter={(v) => rupiah(v)} />
              <Bar dataKey="nilai" radius={[0, 4, 4, 0]}>
                {distribusiKategori.map((d, i) => <Cell key={i} fill={warnaKategoriMap[d.kategori] || "#3FA796"} />)}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      )}
    </div>
  );
}

// ============================================================
function RiwayatHargaView({ riwayat, products, produkTerpilih, isAdmin, onEditEntry, bisaLihatHarga = true }) {
  const [filterProduk, setFilterProduk] = useState(produkTerpilih && produkTerpilih !== "semua" ? produkTerpilih : "semua");
  const [filterField, setFilterField] = useState("Semua");
  const [searchQuery, setSearchQuery] = useState("");
  const [tanggalMulai, setTanggalMulai] = useState("");
  const [tanggalSelesai, setTanggalSelesai] = useState("");

  if (!bisaLihatHarga) {
    return (
      <div>
        <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, margin: 0 }}>Riwayat Perubahan Harga</h1>
        <div style={{ marginTop: 20, background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, padding: 30, textAlign: "center", color: "#8B95A1" }}>
          <Lock size={22} style={{ marginBottom: 10, opacity: 0.6 }} />
          <p style={{ margin: 0 }}>Akun ini tidak diberi akses untuk melihat harga. Hubungi admin kalau ini keliru.</p>
        </div>
      </div>
    );
  }

  const daftar = useMemo(() => {
    return riwayat.filter(r => {
      if (filterProduk !== "semua" && r.produkId !== filterProduk) return false;
      if (filterField !== "Semua" && r.field !== filterField) return false;
      if (tanggalMulai && r.tanggal < `${tanggalMulai}T00:00:00.000Z`) return false;
      if (tanggalSelesai && r.tanggal > `${tanggalSelesai}T23:59:59.999Z`) return false;
      if (searchQuery) {
        const q = searchQuery.toLowerCase();
        if (![r.kodeBarang, r.namaProduk, r.field, r.keterangan].some(v => String(v || "").toLowerCase().includes(q))) return false;
      }
      return true;
    });
  }, [riwayat, filterProduk, filterField, searchQuery, tanggalMulai, tanggalSelesai]);

  const dataGrafik = useMemo(() => {
    if (filterProduk === "semua") return [];
    return [...daftar].reverse().map(r => ({ tanggal: tanggalID(r.tanggal), [r.field]: r.hargaBaru }));
  }, [daftar, filterProduk]);

  return (
    <div>
      <h1 style={{ fontFamily: "'Space Grotesk', sans-serif", fontSize: 22, margin: 0 }}>Riwayat Perubahan Harga</h1>
      <p style={{ color: "#8B95A1", marginTop: 4 }}>Pantau kenaikan dan penurunan harga beli maupun harga jual dari waktu ke waktu.</p>
      <BarcodeDivider />

      <div className="no-print" style={{ display: "flex", flexWrap: "wrap", gap: 10, marginBottom: 16 }}>
        <select value={filterProduk} onChange={e => setFilterProduk(e.target.value)} style={{ ...inputStyle, minWidth: 220, flex: 1 }}>
          <option value="semua">Semua Produk</option>
          {products.map(p => <option key={p.id} value={p.id}>{p.nama} ({p.kodeBarang})</option>)}
        </select>
        <select value={filterField} onChange={e => setFilterField(e.target.value)} style={{ ...inputStyle, width: 180 }}>
          <option value="Semua">Semua Jenis Harga</option>
          <option value="Harga Beli">Harga Beli</option>
          <option value="Harga Jual">Harga Jual</option>
        </select>
        <input value={searchQuery} onChange={e => setSearchQuery(e.target.value)} placeholder="Cari produk, kode, keterangan..." style={{ ...inputStyle, minWidth: 220, flex: 1 }} />
        <input type="date" value={tanggalMulai} onChange={e => setTanggalMulai(e.target.value)} style={{ ...inputStyle, minWidth: 150 }} />
        <input type="date" value={tanggalSelesai} onChange={e => setTanggalSelesai(e.target.value)} style={{ ...inputStyle, minWidth: 150 }} />
      </div>

      {filterProduk !== "semua" && dataGrafik.length > 0 && (
        <div style={{ background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, padding: 18, marginBottom: 16 }}>
          <h3 style={judulKartu}>Grafik Perubahan Harga</h3>
          <ResponsiveContainer width="100%" height={240}>
            <LineChart data={dataGrafik}>
              <CartesianGrid strokeDasharray="3 3" stroke="#2A3138" />
              <XAxis dataKey="tanggal" stroke="#5C6570" fontSize={11} />
              <YAxis stroke="#5C6570" fontSize={11} tickFormatter={(v) => rupiah(v)} />
              <Tooltip contentStyle={tooltipStyle} formatter={(v) => rupiah(v)} />
              <Legend wrapperStyle={{ fontSize: 11 }} />
              <Line type="monotone" dataKey="Harga Beli" stroke="#F2C14E" strokeWidth={2} connectNulls />
              <Line type="monotone" dataKey="Harga Jual" stroke="#3FA796" strokeWidth={2} connectNulls />
            </LineChart>
          </ResponsiveContainer>
        </div>
      )}

      <div style={{ background: "#1D2329", border: "1px solid #2A3138", borderRadius: 10, overflow: "hidden" }}>
        <table>
          <thead>
            <tr style={{ background: "#171B20", color: "#8B95A1", fontSize: 12, textTransform: "uppercase" }}>
              <th>Tanggal</th><th>Produk</th><th>Kode Barang</th><th>Jenis Harga</th><th>Harga Lama</th><th>Harga Baru</th><th>Perubahan</th><th className="no-print"></th>
            </tr>
          </thead>
          <tbody>
            {daftar.map(r => {
              const naik = r.hargaBaru > r.hargaLama;
              const persen = r.hargaLama ? Math.abs(((r.hargaBaru - r.hargaLama) / r.hargaLama) * 100).toFixed(1) : 0;
              return (
                <tr key={r.id} style={{ borderTop: "1px solid #2A3138", fontSize: 13 }}>
                  <td style={{ color: "#8B95A1" }}>{tanggalWaktuID(r.tanggal)}</td>
                  <td>{r.namaProduk}</td>
                  <td style={{ fontFamily: "'IBM Plex Mono', monospace", color: "#8B95A1" }}>{r.kodeBarang}</td>
                  <td>{r.field}</td>
                  <td style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{rupiah(r.hargaLama)}</td>
                  <td style={{ fontFamily: "'IBM Plex Mono', monospace" }}>{rupiah(r.hargaBaru)}</td>
                  <td>
                    <span style={{ display: "flex", alignItems: "center", gap: 4, color: naik ? "#E2574C" : "#3FA796", fontFamily: "'IBM Plex Mono', monospace" }}>
                      {naik ? <ArrowUp size={13} /> : <ArrowDown size={13} />} {persen}%
                    </span>
                  </td>
                  <td className="no-print">
                    {isAdmin && onEditEntry && (
                      <button onClick={() => onEditEntry(r)} title="Edit riwayat harga" className="icon-btn-hover" style={iconBtn}><Pencil size={13} /></button>
                    )}
                  </td>
                </tr>
              );
            })}
            {daftar.length === 0 && <tr><td colSpan={8} style={{ textAlign: "center", padding: 30, color: "#5C6570" }}>Belum ada riwayat perubahan harga yang cocok.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================
function ModalProduk({ data, error, onBatal, onSimpan, defaultGudang, categories, products = [], bisaLihatHarga = true }) {
  const [form, setForm] = useState(data || {
    id: null, kodeBarang: "", nama: "", kategori: categories[0]?.nama || "", gudang: defaultGudang,
    satuan: SATUAN_GROUPS[0].opsi[0], stok: 0, stokMin: 10, hargaBeli: 0, hargaJual: 0, supplier: ""
  });
  const [saranTerbuka, setSaranTerbuka] = useState(false);

  const set = (field) => (e) => {
    const v = e.target.value;
    const numFields = ["stok", "stokMin", "hargaBeli", "hargaJual"];
    setForm(prev => ({ ...prev, [field]: numFields.includes(field) ? Number(v) : v }));
  };

  const kodePreview = useMemo(() => {
    if (data) return data.kodeBarang;
    const serupa = products.find(p => p.nama.trim().toLowerCase() === form.nama.trim().toLowerCase());
    if (serupa) return serupa.kodeBarang;
    const kat = categories.find(c => c.nama === form.kategori);
    return kat ? `${kat.prefix}-${pad((kat.counter || 0) + 1, 3)}` : "-";
  }, [data, form.kategori, form.nama, categories, products]);

  // Saat menambah produk baru: cari nama produk yang pernah diinput sebelumnya (di gudang manapun)
  // supaya kategori/satuan/harga bisa otomatis terisi dari riwayat, tanpa perlu ketik ulang dari nol.
  const namaLower = form.nama.trim().toLowerCase();
  const saranProduk = useMemo(() => {
    if (data || namaLower.length < 2) return [];
    const dilihat = new Set();
    return products
      .filter(p => p.nama.toLowerCase().includes(namaLower))
      .filter(p => (dilihat.has(p.nama.toLowerCase()) ? false : (dilihat.add(p.nama.toLowerCase()), true)))
      .slice(0, 6);
  }, [products, namaLower, data]);

  const sudahAdaDiGudangIni = !data && products.some(p => p.nama.toLowerCase() === namaLower && p.gudang === form.gudang);

  function pakaiSaran(p) {
    setForm(prev => ({
      ...prev, nama: p.nama, kategori: p.kategori, satuan: p.satuan,
      stokMin: p.stokMin, hargaBeli: p.hargaBeli, hargaJual: p.hargaJual, supplier: p.supplier,
    }));
    setSaranTerbuka(false);
  }

  return (
    <Overlay onBatal={onBatal}>
      <h3 style={{ marginTop: 0, fontFamily: "'Space Grotesk', sans-serif" }}>{data ? "Edit Produk" : "Tambah Produk Baru"}</h3>
      {data ? (
        <p style={{ color: "#8B95A1", fontSize: 12, marginTop: -8 }}>Perubahan harga beli/jual akan otomatis tercatat di Riwayat Harga. Kode barang bersifat permanen.</p>
      ) : (
        <p style={{ color: "#8B95A1", fontSize: 12, marginTop: -8 }}>Kode barang akan otomatis terisi berurutan sesuai kategori yang dipilih. Ketik nama produk yang pernah diinput untuk mengisi otomatis dari riwayat.</p>
      )}

      <Grid>
        <Field label="Nama Produk">
          <div style={{ position: "relative" }}>
            <input
              value={form.nama}
              onChange={set("nama")}
              onFocus={() => setSaranTerbuka(true)}
              onBlur={() => setTimeout(() => setSaranTerbuka(false), 150)}
              autoComplete="off"
              style={inputStyle}
            />
            {saranTerbuka && saranProduk.length > 0 && (
              <div style={{
                position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, zIndex: 20,
                background: "#171B20", border: "1px solid #2A3138", borderRadius: 8, overflow: "hidden",
                boxShadow: "0 8px 20px rgba(0,0,0,0.35)",
              }}>
                {saranProduk.map(p => (
                  <button type="button" key={p.id} onMouseDown={() => pakaiSaran(p)}
                    style={{
                      display: "block", width: "100%", textAlign: "left", padding: "9px 12px",
                      background: "none", border: "none", borderBottom: "1px solid #2A3138",
                      color: "#EDEFF2", fontSize: 12.5, cursor: "pointer",
                    }}>
                    <div style={{ fontWeight: 600 }}>{p.nama}</div>
                    <div style={{ color: "#8B95A1", fontSize: 11 }}>
                      {p.kategori} • {p.gudang} • Stok: {p.stok} {satuanSingkat(p.satuan)} • {rupiah(p.hargaJual)}
                    </div>
                  </button>
                ))}
              </div>
            )}
          </div>
          {sudahAdaDiGudangIni && (
            <span style={{ fontSize: 11, color: "#E8A33D" }}>
              Produk dengan nama ini sudah ada di {form.gudang}. Untuk menambah stoknya, gunakan menu <b>Mutasi Barang</b> agar tercatat, bukan membuat produk baru.
            </span>
          )}
        </Field>
        <Field label="Kode Barang (otomatis, permanen)">
          <div style={{ ...inputStyle, display: "flex", alignItems: "center", fontFamily: "'IBM Plex Mono', monospace", color: "#8B95A1", background: "#171B20", cursor: "not-allowed" }}>{kodePreview}</div>
        </Field>
        <Field label="Kategori">
          <select value={form.kategori} onChange={set("kategori")} style={inputStyle}>{categories.map(k => <option key={k.id}>{k.nama}</option>)}</select>
        </Field>
        <Field label="Satuan">
          <select value={form.satuan} onChange={set("satuan")} style={inputStyle}>
            {SATUAN_GROUPS.map(g => (
              <optgroup key={g.grup} label={g.grup}>
                {g.opsi.map(o => <option key={o} value={o}>{o}</option>)}
              </optgroup>
            ))}
          </select>
        </Field>
        <Field label="Gudang">
          <select value={form.gudang} onChange={set("gudang")} style={inputStyle}>{GUDANG.map(g => <option key={g}>{g}</option>)}</select>
        </Field>
        <Field label="Supplier"><input value={form.supplier} onChange={set("supplier")} style={inputStyle} /></Field>
        <Field label={`Stok Saat Ini (${satuanSingkat(form.satuan)})`}>
          {data ? (
            <>
              <div style={{ ...inputStyle, display: "flex", alignItems: "center", fontFamily: "'IBM Plex Mono', monospace", color: "#8B95A1", background: "#171B20", cursor: "not-allowed" }}>
                {form.stok} {satuanSingkat(form.satuan)}
              </div>
              <span style={{ fontSize: 11, color: "#5C6570" }}>Untuk mengubah jumlah stok, gunakan menu <b>Mutasi Barang</b> agar tercatat.</span>
            </>
          ) : (
            <input type="number" value={form.stok} onChange={set("stok")} style={inputStyle} />
          )}
        </Field>
        <Field label={`Stok Minimum (${satuanSingkat(form.satuan)})`}><input type="number" value={form.stokMin} onChange={set("stokMin")} style={inputStyle} /></Field>
        {bisaLihatHarga ? (
          <>
            <Field label="Harga Beli (Rp)"><input type="number" value={form.hargaBeli} onChange={set("hargaBeli")} style={inputStyle} /></Field>
            <Field label="Harga Jual (Rp)"><input type="number" value={form.hargaJual} onChange={set("hargaJual")} style={inputStyle} /></Field>
          </>
        ) : (
          <Field label="Harga">
            <div style={{ ...inputStyle, display: "flex", alignItems: "center", color: "#5C6570", background: "#171B20", cursor: "not-allowed" }}>Tidak ditampilkan untuk akun ini</div>
          </Field>
        )}
      </Grid>

      {error && <div style={{ color: "#E2574C", fontSize: 13, marginTop: 10 }}>{error}</div>}

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onBatal} style={btnSecondary}>Batal</button>
        <button onClick={() => onSimpan(form)} className="btn-primary-glow" style={btnPrimary}>Simpan</button>
      </div>
    </Overlay>
  );
}

// ============================================================
function ModalKategori({ categories, products, onBatal, onTambah, onEdit, onHapus }) {
  const kosong = { nama: "", prefix: "", warna: "#3FA796" };
  const [form, setForm] = useState(kosong);
  const [editId, setEditId] = useState(null);
  const [error, setError] = useState("");

  function mulaiEdit(k) {
    setEditId(k.id);
    setForm({ nama: k.nama, prefix: k.prefix, warna: k.warna });
    setError("");
  }

  function batalEdit() {
    setEditId(null);
    setForm(kosong);
    setError("");
  }

  function submit() {
    const err = editId ? onEdit(editId, form.nama, form.prefix, form.warna) : onTambah(form.nama, form.prefix, form.warna);
    if (err) { setError(err); return; }
    setError("");
    setForm(kosong);
    setEditId(null);
  }

  return (
    <Overlay onBatal={onBatal}>
      <h3 style={{ marginTop: 0, fontFamily: "'Space Grotesk', sans-serif", display: "flex", alignItems: "center", gap: 8 }}>
        <Tags size={18} color="#3FA796" /> Kelola Kategori Barang
      </h3>
      <p style={{ color: "#8B95A1", fontSize: 12, marginTop: -8 }}>Tambah, ubah, atau hapus kategori. Prefix menentukan awalan kode barang otomatis (contoh: ELK-001).</p>

      <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 14, maxHeight: 220, overflowY: "auto" }}>
        {categories.map(k => {
          const jumlahProduk = products.filter(p => p.kategori === k.nama).length;
          return (
            <div key={k.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "#171B20", border: "1px solid #2A3138", borderRadius: 8, padding: "8px 12px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 12, height: 12, borderRadius: 4, background: k.warna }} />
                <div>
                  <div style={{ fontWeight: 500, fontSize: 13 }}>{k.nama}</div>
                  <div style={{ fontSize: 11, color: "#5C6570", fontFamily: "'IBM Plex Mono', monospace" }}>{k.prefix}-XXX · {jumlahProduk} produk</div>
                </div>
              </div>
              <div style={{ display: "flex", gap: 6 }}>
                <button onClick={() => mulaiEdit(k)} className="icon-btn-hover" style={iconBtn}><Pencil size={13} /></button>
                <button onClick={() => onHapus(k.id)} className="icon-btn-hover" style={{ ...iconBtn, color: "#E2574C" }}><Trash2 size={13} /></button>
              </div>
            </div>
          );
        })}
      </div>

      <div style={{ borderTop: "1px solid #2A3138", marginTop: 16, paddingTop: 16 }}>
        <div style={{ fontSize: 12, color: "#8B95A1", marginBottom: 10, fontWeight: 600, textTransform: "uppercase" }}>{editId ? "Edit Kategori" : "Tambah Kategori Baru"}</div>
        <Grid>
          <Field label="Nama Kategori"><input value={form.nama} onChange={e => setForm(f => ({ ...f, nama: e.target.value }))} style={inputStyle} /></Field>
          <Field label="Prefix Kode (mis. ELK)"><input value={form.prefix} onChange={e => setForm(f => ({ ...f, prefix: e.target.value.toUpperCase() }))} maxLength={4} style={inputStyle} placeholder="Otomatis dari nama" /></Field>
          <Field label="Warna Label">
            <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
              <input type="color" value={form.warna} onChange={e => setForm(f => ({ ...f, warna: e.target.value }))} style={{ width: 40, height: 34, border: "1px solid #2A3138", borderRadius: 6, background: "none", padding: 2 }} />
              <span style={{ fontSize: 12, color: "#8B95A1", fontFamily: "'IBM Plex Mono', monospace" }}>{form.warna}</span>
            </div>
          </Field>
        </Grid>
        {error && <div style={{ color: "#E2574C", fontSize: 13, marginTop: 10 }}>{error}</div>}
        <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 16 }}>
          {editId && <button onClick={batalEdit} style={btnSecondary}>Batal Edit</button>}
          <button onClick={submit} className="btn-primary-glow" style={btnPrimary}>{editId ? "Simpan Perubahan" : "Tambah Kategori"}</button>
        </div>
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", marginTop: 16 }}>
        <button onClick={onBatal} style={btnSecondary}>Tutup</button>
      </div>
    </Overlay>
  );
}

// ============================================================
function ModalMutasi({ products, isAdmin, currentUser, onBatal, onSimpan, bisaLihatHarga = true }) {
  const [gudangPilihan, setGudangPilihan] = useState(isAdmin ? GUDANG[0] : currentUser.gudang);
  const produkGudang = useMemo(() => products.filter(p => p.gudang === gudangPilihan), [products, gudangPilihan]);
  const [jenis, setJenis] = useState("Masuk");
  const [produkId, setProdukId] = useState(produkGudang[0]?.id || "");
  const [jumlah, setJumlah] = useState(1);
  const [tanggal, setTanggal] = useState(new Date().toISOString().slice(0, 10));
  const [keterangan, setKeterangan] = useState("");
  const [hargaSatuan, setHargaSatuan] = useState("");
  const [tujuanGudang, setTujuanGudang] = useState(() => GUDANG.find(g => g !== (isAdmin ? GUDANG[0] : currentUser.gudang)) || GUDANG[0]);
  const [error, setError] = useState("");

  useEffect(() => {
    if (!produkGudang.find(p => p.id === produkId)) setProdukId(produkGudang[0]?.id || "");
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [gudangPilihan]);

  const produkDipilih = produkGudang.find(p => p.id === produkId);

  return (
    <Overlay onBatal={onBatal}>
      <h3 style={{ marginTop: 0, fontFamily: "'Space Grotesk', sans-serif", display: "flex", alignItems: "center", gap: 8 }}>
        <ArrowRightLeft size={18} color="#3FA796" /> Catat Mutasi Barang
      </h3>
      <p style={{ color: "#8B95A1", fontSize: 12, marginTop: -8 }}>Berlaku untuk semua gudang. Kode invoice akan dibuat otomatis setelah transaksi disimpan.</p>

      <div className="no-print" style={{ display: "flex", gap: 8, marginTop: 14 }}>
        {["Masuk", "Keluar", "Transfer"].map(j => {
          const Ikon = j === "Masuk" ? PackagePlus : j === "Keluar" ? PackageMinus : ArrowRightLeft;
          const aktif = jenis === j;
          const warna = j === "Masuk" ? "#3FA796" : j === "Keluar" ? "#E2574C" : "#6C8EBF";
          return (
            <button key={j} onClick={() => setJenis(j)} className="jenis-pill" style={{
              flex: 1, display: "flex", alignItems: "center", justifyContent: "center", gap: 6, padding: "10px 12px",
              borderRadius: 8, border: `1px solid ${aktif ? warna : "#2A3138"}`, background: aktif ? `${warna}1A` : "transparent",
              color: aktif ? warna : "#8B95A1", fontWeight: 600
            }}>
              <Ikon size={15} /> Barang {j}
            </button>
          );
        })}
      </div>

      <Grid>
        {isAdmin ? (
          <Field label="Gudang">
            <select value={gudangPilihan} onChange={e => setGudangPilihan(e.target.value)} style={inputStyle}>
              {GUDANG.map(g => <option key={g}>{g}</option>)}
            </select>
          </Field>
        ) : (
          <Field label="Gudang"><div style={{ ...inputStyle, display: "flex", alignItems: "center", color: "#8B95A1" }}>{currentUser.gudang}</div></Field>
        )}
        <Field label="Tanggal"><input type="date" value={tanggal} onChange={e => setTanggal(e.target.value)} style={inputStyle} /></Field>
        <Field label="Produk">
          <select value={produkId} onChange={e => setProdukId(e.target.value)} style={inputStyle}>
            {produkGudang.map(p => <option key={p.id} value={p.id}>{p.nama} ({p.stok} {satuanSingkat(p.satuan)} tersedia)</option>)}
          </select>
        </Field>
        <Field label={`Jumlah ${produkDipilih ? `(${satuanSingkat(produkDipilih.satuan)})` : ""}`}>
          <input type="number" min={1} value={jumlah} onChange={e => setJumlah(Number(e.target.value))} style={inputStyle} />
        </Field>
        {bisaLihatHarga && (
          <Field label="Harga Satuan (Rp)">
            <input type="text" inputMode="numeric" value={hargaSatuan} onChange={e => {
              const raw = e.target.value.replace(/[^0-9]/g, '');
              setHargaSatuan(raw.replace(/^0+/, ''));
            }} placeholder={jenis === "Masuk" ? "Masukkan harga beli per unit" : "Opsional untuk keluar/transfer"} style={inputStyle} />
          </Field>
        )}
        {jenis === "Transfer" && (
          <Field label="Gudang Tujuan">
            <select value={tujuanGudang} onChange={e => setTujuanGudang(e.target.value)} style={inputStyle}>
              {GUDANG.filter(g => g !== (isAdmin ? gudangPilihan : currentUser.gudang)).map(g => <option key={g} value={g}>{g}</option>)}
            </select>
          </Field>
        )}
      </Grid>

      <div style={{ marginTop: 14 }}>
        <Field label="Keterangan (opsional)">
          <input value={keterangan} onChange={e => setKeterangan(e.target.value)} placeholder="mis. Pembelian dari supplier, Transfer gudang, Penyesuaian stok..." style={inputStyle} />
        </Field>
      </div>

      {error && <div style={{ color: "#E2574C", fontSize: 13, marginTop: 10 }}>{error}</div>}

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onBatal} style={btnSecondary}>Batal</button>
        <button onClick={() => {
          if (!produkId) { setError("Tidak ada produk di gudang ini."); return; }
          const err = onSimpan(jenis, produkId, jumlah, tanggal, keterangan, null, Number(hargaSatuan || 0), tujuanGudang);
          if (err) setError(err);
        }} className="btn-primary-glow" style={btnPrimary}>Simpan & Buat Invoice</button>
      </div>
    </Overlay>
  );
}

// ============================================================
function ModalInvoiceMutasi({ record, onTutup, onCetak }) {
  const warna = record.jenis === "Masuk" ? "#3FA796" : "#E2574C";
  return (
    <Overlay onBatal={onTutup}>
      <div style={{ textAlign: "center" }}>
        <CheckCircle2 size={40} color={warna} style={{ marginBottom: 8 }} />
        <h3 style={{ margin: "0 0 2px", fontFamily: "'Space Grotesk', sans-serif" }}>Mutasi Berhasil Dicatat</h3>
        <p style={{ color: "#8B95A1", fontSize: 12, marginTop: 2 }}>Kode invoice otomatis dibuat untuk transaksi ini.</p>
      </div>

      <div style={{ background: "#171B20", border: "1px solid #2A3138", borderRadius: 10, padding: 16, marginTop: 16 }}>
        <div style={{ textAlign: "center", fontFamily: "'IBM Plex Mono', monospace", fontSize: 16, fontWeight: 700, letterSpacing: 1, marginBottom: 10 }}>{record.kode}</div>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 12 }}>
          <span style={{ fontSize: 11, fontWeight: 700, padding: "4px 12px", borderRadius: 20, color: warna, background: `${warna}1A` }}>
            {record.jenis === "Masuk" ? "BARANG MASUK" : "BARANG KELUAR"}
          </span>
        </div>
        {[
          ["Produk", record.namaProduk], ["Kode Barang", record.kodeBarang], ["Gudang", record.gudang],
          ["Jumlah", `${record.jumlah} ${satuanSingkat(record.satuan)}`],
          ["Stok Sebelum → Sesudah", `${record.stokSebelum} → ${record.stokSesudah}`],
          ["Dicatat oleh", record.oleh],
        ].map(([label, val]) => (
          <div key={label} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, padding: "5px 0", borderTop: "1px solid #232B32" }}>
            <span style={{ color: "#8B95A1" }}>{label}</span><span>{val}</span>
          </div>
        ))}
      </div>

      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onTutup} style={btnSecondary}>Tutup</button>
        <button onClick={() => onCetak(record)} className="btn-primary-glow" style={{ ...btnPrimary, display: "flex", alignItems: "center", gap: 6 }}>
          <Printer size={15} /> Cetak Invoice
        </button>
      </div>
    </Overlay>
  );
}

function ModalJual({ products, onBatal, onSimpan, bisaLihatHarga = true }) {
  const [produkId, setProdukId] = useState(products[0]?.id || "");
  const [productSearch, setProductSearch] = useState(`${products[0]?.nama} (${products[0]?.kodeBarang})`);
  const [transactionType, setTransactionType] = useState("Penjualan");
  const [tujuanGudang, setTujuanGudang] = useState(GUDANG.find(g => g !== products[0]?.gudang) || GUDANG[0]);
  const [jumlah, setJumlah] = useState("");
  const [tanggal, setTanggal] = useState(new Date().toISOString().slice(0, 10));
  const [hargaJualOverride, setHargaJualOverride] = useState("");
  const [keterangan, setKeterangan] = useState("");
  const [error, setError] = useState("");
  const filteredProducts = products.filter(p => p.nama.toLowerCase().includes(productSearch.toLowerCase()) || p.kodeBarang.toLowerCase().includes(productSearch.toLowerCase()));
  const produkDipilih = products.find(p => p.id === produkId);

  return (
    <Overlay onBatal={onBatal}>
      <h3 style={{ marginTop: 0, fontFamily: "'Space Grotesk', sans-serif" }}>Catat Penjualan</h3>
      <Grid>
        <Field label="Cari Produk">
          <div style={{ position: 'relative' }}>
            <input type="text" value={productSearch} onChange={e => {
              const next = e.target.value;
              setProductSearch(next);
              const match = products.find(p => p.nama.toLowerCase() === next.toLowerCase() || p.kodeBarang.toLowerCase() === next.toLowerCase() || `${p.nama} (${p.kodeBarang})`.toLowerCase() === next.toLowerCase());
              if (match) {
                setProdukId(match.id);
              } else {
                setProdukId("");
              }
            }} style={inputStyle} placeholder="Ketik nama atau kode" autoComplete="off" />
            {filteredProducts.length > 0 && productSearch && (
              <div style={{ position: 'absolute', zIndex: 10, background: '#12161B', border: '1px solid #2A3138', borderRadius: 10, marginTop: 6, maxHeight: 180, overflowY: 'auto', width: '100%' }}>
                {filteredProducts.slice(0, 6).map(p => (
                  <button key={p.id} type="button" onClick={() => {
                    setProdukId(p.id);
                    setProductSearch(`${p.nama} (${p.kodeBarang})`);
                  }} style={{ width: '100%', textAlign: 'left', padding: '10px 12px', background: 'none', border: 'none', color: '#EDEFF2', cursor: 'pointer' }}>
                    <div style={{ fontSize: 13, fontWeight: 600 }}>{p.nama}</div>
                    <div style={{ fontSize: 11, color: '#8B95A1' }}>{p.kodeBarang}</div>
                  </button>
                ))}
              </div>
            )}
          </div>
        </Field>
        <Field label="Nama Produk"><input type="text" value={produkDipilih?.nama || ''} readOnly style={{ ...inputStyle, background: '#161C21' }} /></Field>
        <Field label="Kode Barang"><input type="text" value={produkDipilih?.kodeBarang || ''} readOnly style={{ ...inputStyle, background: '#161C21' }} /></Field>
        <Field label="Jenis Transaksi">
          <select value={transactionType} onChange={e => setTransactionType(e.target.value)} style={inputStyle}>
            <option value="Penjualan">Penjualan</option>
            <option value="Keluar">Keluar</option>
            <option value="Transfer">Transfer</option>
          </select>
        </Field>
        <Field label="Tanggal"><input type="date" value={tanggal} onChange={e => setTanggal(e.target.value)} style={inputStyle} /></Field>
        <Field label="Jumlah"><input type="text" inputMode="numeric" value={jumlah} onChange={e => {
              const raw = e.target.value.replace(/[^0-9]/g, '');
              setJumlah(raw.replace(/^0+/, ''));
            }} style={inputStyle} placeholder="Kosongkan jika belum tahu" /></Field>
        {bisaLihatHarga && (
          <Field label="Harga Jual per Unit (Rp)"><input type="text" inputMode="numeric" value={hargaJualOverride} onChange={e => {
              const raw = e.target.value.replace(/[^0-9]/g, '');
              setHargaJualOverride(raw.replace(/^0+/, ''));
            }} style={inputStyle} placeholder="Kosongkan untuk gunakan harga default" /></Field>
        )}
      </Grid>
      {transactionType === "Transfer" && (
        <Field label="Gudang Tujuan">
          <select value={tujuanGudang} onChange={e => setTujuanGudang(e.target.value)} style={inputStyle}>
            {GUDANG.filter(g => g !== (produkDipilih?.gudang || "")).map(g => <option key={g} value={g}>{g}</option>)}
          </select>
        </Field>
      )}
      <Field label="Keterangan (opsional)"><input value={keterangan} onChange={e => setKeterangan(e.target.value)} style={inputStyle} placeholder="Contoh: Penjualan online, Pesanan toko" /></Field>
      {bisaLihatHarga && (
        <Field label="Estimasi Total">
          <div style={{ ...inputStyle, display: "flex", alignItems: "center", fontFamily: "'IBM Plex Mono', monospace" }}>
            {rupiah((Number(hargaJualOverride) > 0 ? Number(hargaJualOverride) : (produkDipilih?.hargaJual || 0)) * Number(jumlah || 0))}
          </div>
        </Field>
      )}
      {error && <div style={{ color: "#E2574C", fontSize: 13, marginTop: 10 }}>{error}</div>}
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onBatal} style={btnSecondary}>Batal</button>
        <button onClick={() => { const err = onSimpan(produkId, Number(jumlah || 0), tanggal, Number(hargaJualOverride || 0), keterangan, transactionType, tujuanGudang); if (err) setError(err); }} className="btn-primary-glow" style={btnPrimary}>Simpan Transaksi</button>
      </div>
    </Overlay>
  );
}

function ModalEditRiwayat({ entry, onBatal, onSimpan }) {
  const [hargaLama, setHargaLama] = useState(entry?.hargaLama || 0);
  const [hargaBaru, setHargaBaru] = useState(entry?.hargaBaru || 0);
  const [keterangan, setKeterangan] = useState(entry?.keterangan || "");
  const [error, setError] = useState("");

  useEffect(() => {
    if (entry) {
      setHargaLama(entry.hargaLama || 0);
      setHargaBaru(entry.hargaBaru || 0);
      setKeterangan(entry.keterangan || "");
      setError("");
    }
  }, [entry]);

  return (
    <Overlay onBatal={onBatal}>
      <h3 style={{ marginTop: 0, fontFamily: "'Space Grotesk', sans-serif" }}>Edit Riwayat Harga</h3>
      <p style={{ color: "#8B95A1", fontSize: 12, marginTop: -8 }}>Perbarui detail riwayat perubahan harga. Ini tidak mengubah harga produk saat ini secara otomatis.</p>
      <Grid>
        <Field label="Produk">
          <input value={entry?.namaProduk || ""} readOnly style={{ ...inputStyle, background: "#161C21" }} />
        </Field>
        <Field label="Kode Barang">
          <input value={entry?.kodeBarang || ""} readOnly style={{ ...inputStyle, background: "#161C21" }} />
        </Field>
        <Field label="Jenis Harga">
          <input value={entry?.field || ""} readOnly style={{ ...inputStyle, background: "#161C21" }} />
        </Field>
        <Field label="Tanggal">
          <input value={tanggalWaktuID(entry?.tanggal || new Date().toISOString())} readOnly style={{ ...inputStyle, background: "#161C21" }} />
        </Field>
        <Field label="Harga Lama (Rp)">
          <input type="text" inputMode="numeric" value={hargaLama} onChange={e => {
            const raw = e.target.value.replace(/[^0-9]/g, '');
            setHargaLama(raw ? Number(raw) : 0);
          }} style={inputStyle} />
        </Field>
        <Field label="Harga Baru (Rp)">
          <input type="text" inputMode="numeric" value={hargaBaru} onChange={e => {
            const raw = e.target.value.replace(/[^0-9]/g, '');
            setHargaBaru(raw ? Number(raw) : 0);
          }} style={inputStyle} />
        </Field>
      </Grid>
      <div style={{ marginTop: 14 }}>
        <Field label="Keterangan (opsional)">
          <input value={keterangan} onChange={e => setKeterangan(e.target.value)} style={inputStyle} />
        </Field>
      </div>
      {error && <div style={{ color: "#E2574C", fontSize: 13, marginTop: 10 }}>{error}</div>}
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onBatal} style={btnSecondary}>Batal</button>
        <button onClick={() => {
          if (hargaLama <= 0 || hargaBaru <= 0) { setError("Harga harus lebih dari 0."); return; }
          const err = onSimpan(entry.id, hargaLama, hargaBaru, keterangan);
          if (err) setError(err);
          else onBatal();
        }} className="btn-primary-glow" style={btnPrimary}>Simpan Perubahan</button>
      </div>
    </Overlay>
  );
}

function ModalEditJual({ sale, products, onBatal, onSimpan }) {
  const [produkId, setProdukId] = useState(sale.produkId);
  const [productInput, setProductInput] = useState(`${sale.namaProduk} (${sale.kodeBarang})`);
  const [jumlah, setJumlah] = useState(sale.jumlah);
  const [tanggal, setTanggal] = useState(sale.tanggal.slice(0, 10));
  const [hargaJualOverride, setHargaJualOverride] = useState(sale.hargaJualSaat);
  const [keterangan, setKeterangan] = useState(sale.keterangan || "");
  const [error, setError] = useState("");
  const produkDipilih = products.find(p => p.id === produkId);

  return (
    <Overlay onBatal={onBatal}>
      <h3 style={{ marginTop: 0, fontFamily: "'Space Grotesk', sans-serif" }}>Edit Transaksi Penjualan</h3>
      <p style={{ color: "#8B95A1", fontSize: 12, marginTop: -8 }}>Gunakan ini untuk memperbaiki kesalahan input seperti jumlah, harga, atau produk yang keliru dipilih.</p>
      <Grid>
        <Field label="Nama / Kode Barang">
          <input list="produk-edit-list" value={productInput} onChange={e => {
            const next = e.target.value;
            setProductInput(next);
            const match = products.find(p => `${p.nama} (${p.kodeBarang})` === next);
            if (match) setProdukId(match.id);
          }} style={inputStyle} placeholder="Ketik nama atau kode" />
          <datalist id="produk-edit-list">
            {products.map(p => <option key={p.id} value={`${p.nama} (${p.kodeBarang})`} />)}
          </datalist>
        </Field>
        <Field label="Tanggal"><input type="date" value={tanggal} onChange={e => setTanggal(e.target.value)} style={inputStyle} /></Field>
        <Field label="Jumlah"><input type="number" min={1} value={jumlah} onChange={e => setJumlah(Math.max(1, Number(e.target.value) || 1))} style={inputStyle} /></Field>
        <Field label="Harga Jual per Unit (Rp)"><input type="number" min={0} value={hargaJualOverride} onChange={e => setHargaJualOverride(Number(e.target.value))} style={inputStyle} /></Field>
      </Grid>
      <Field label="Keterangan (opsional)"><input value={keterangan} onChange={e => setKeterangan(e.target.value)} style={inputStyle} placeholder="Contoh: Penjualan online, Pesanan toko" /></Field>
      <Field label="Estimasi Total">
        <div style={{ ...inputStyle, display: "flex", alignItems: "center", fontFamily: "'IBM Plex Mono', monospace" }}>
          {rupiah((hargaJualOverride > 0 ? hargaJualOverride : (produkDipilih?.hargaJual || 0)) * jumlah)}
        </div>
      </Field>
      {error && <div style={{ color: "#E2574C", fontSize: 13, marginTop: 10 }}>{error}</div>}
      <div style={{ display: "flex", justifyContent: "flex-end", gap: 10, marginTop: 20 }}>
        <button onClick={onBatal} style={btnSecondary}>Batal</button>
        <button onClick={() => { const err = onSimpan(produkId, jumlah, tanggal, hargaJualOverride, keterangan); if (err) setError(err); }} className="btn-primary-glow" style={btnPrimary}>Simpan Perubahan</button>
      </div>
    </Overlay>
  );
}

function ModalNotifikasi({ stokMenipis, onTutup }) {
  const [nomorWA, setNomorWA] = useState("");
  const [email, setEmail] = useState("");
  const [sudahUnduh, setSudahUnduh] = useState(false);

  function ringkasanTeks() {
    const baris = stokMenipis.map(p => `• ${p.nama} (${p.kodeBarang}) — sisa ${p.stok} ${satuanSingkat(p.satuan)}, minimum ${p.stokMin}`).join("\n");
    return `Peringatan Stok Menipis - GudangKu\nTanggal: ${tanggalID(new Date().toISOString())}\n\n${baris}\n\nRincian lengkap ada pada file Excel terlampir.`;
  }

  function unduhExcel() {
    eksporExcel([{ nama: "Stok Menipis", data: stokMenipis.map(p => ({ "Kode Barang": p.kodeBarang, Produk: p.nama, Kategori: p.kategori, Gudang: p.gudang, Satuan: p.satuan, "Stok Saat Ini": p.stok, "Stok Minimum": p.stokMin, Supplier: p.supplier })) }], "notifikasi-stok-menipis");
    setSudahUnduh(true);
  }

  function kirimWhatsApp() {
    const teks = encodeURIComponent(ringkasanTeks());
    const tujuan = nomorWA.replace(/\D/g, "");
    window.open(`https://wa.me/${tujuan}?text=${teks}`, "_blank");
  }

  function kirimEmail() {
    const subjek = encodeURIComponent("Peringatan Stok Menipis - GudangKu");
    const badan = encodeURIComponent(ringkasanTeks());
    window.open(`mailto:${email}?subject=${subjek}&body=${badan}`, "_blank");
  }

  return (
    <Overlay onBatal={onTutup}>
      <h3 style={{ marginTop: 0, fontFamily: "'Space Grotesk', sans-serif", display: "flex", alignItems: "center", gap: 8 }}>
        <Bell size={18} color="#E8A33D" /> Notifikasi Stok Menipis
      </h3>

      {stokMenipis.length === 0 ? (
        <p style={{ color: "#8B95A1" }}>Semua stok dalam kondisi aman, tidak ada yang perlu dilaporkan.</p>
      ) : (
        <>
          <div style={{ display: "flex", flexDirection: "column", gap: 8, marginTop: 12, maxHeight: 180, overflowY: "auto" }}>
            {stokMenipis.map(p => (
              <div key={p.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 13, borderBottom: "1px solid #2A3138", paddingBottom: 6 }}>
                <span>{p.nama} <span style={{ color: "#5C6570" }}>({p.kodeBarang})</span></span>
                <span style={{ color: "#E8A33D", fontFamily: "'IBM Plex Mono', monospace" }}>{p.stok}/{p.stokMin} {satuanSingkat(p.satuan)}</span>
              </div>
            ))}
          </div>

          <div style={{ marginTop: 16 }}>
            <button onClick={unduhExcel} style={{ ...btnSecondary, display: "flex", alignItems: "center", gap: 6, width: "100%", justifyContent: "center" }}>
              <FileSpreadsheet size={15} /> 1. Unduh Laporan Excel
            </button>
          </div>

          <Grid>
            <Field label="Nomor WhatsApp tujuan (format 62xxx)">
              <input value={nomorWA} onChange={e => setNomorWA(e.target.value)} placeholder="62812xxxxxxx" style={inputStyle} />
            </Field>
            <Field label="Alamat Email tujuan">
              <input value={email} onChange={e => setEmail(e.target.value)} placeholder="nama@perusahaan.com" style={inputStyle} />
            </Field>
          </Grid>

          <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
            <button onClick={kirimWhatsApp} style={{ ...btnSecondary, flex: 1, display: "flex", alignItems: "center", gap: 6, justifyContent: "center", color: "#3FA796" }}>
              <MessageCircle size={15} /> Kirim ke WhatsApp
            </button>
            <button onClick={kirimEmail} style={{ ...btnSecondary, flex: 1, display: "flex", alignItems: "center", gap: 6, justifyContent: "center", color: "#6C8EBF" }}>
              <Mail size={15} /> Kirim ke Email
            </button>
          </div>

          <p style={{ fontSize: 11, color: "#5C6570", marginTop: 12, lineHeight: 1.5 }}>
            Catatan: browser tidak mengizinkan pengiriman lampiran file secara otomatis melalui WhatsApp maupun email.
            Unduh dulu file Excel di atas ({sudahUnduh ? "sudah diunduh ✓" : "belum diunduh"}), lalu lampirkan secara manual
            pada jendela WhatsApp/Email yang terbuka — pesan ringkasannya sudah otomatis terisi.
          </p>
        </>
      )}
    </Overlay>
  );
}

function Overlay({ children, onBatal }) {
  return (
    <div className="no-print overlay-anim" onClick={onBatal} style={{ position: "fixed", inset: 0, background: "rgba(10,13,16,0.7)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}>
      <div className="modal-anim" onClick={e => e.stopPropagation()} style={{ background: "#1D2329", border: "1px solid #2A3138", borderRadius: 12, padding: 24, width: 540, maxWidth: "92vw", maxHeight: "88vh", overflowY: "auto" }}>
        <button onClick={onBatal} style={{ float: "right", background: "none", border: "none", color: "#8B95A1" }}><X size={18} /></button>
        {children}
      </div>
    </div>
  );
}

function Grid({ children }) { return <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginTop: 16 }}>{children}</div>; }
function Field({ label, children }) { return <label style={{ display: "flex", flexDirection: "column", gap: 6, fontSize: 12, color: "#8B95A1" }}>{label}{children}</label>; }

const inputStyle = { background: "#14181D", border: "1px solid #2A3138", borderRadius: 6, padding: "9px 10px", color: "#EDEFF2", fontSize: 13, outline: "none" };
const btnPrimary = { background: "#3FA796", color: "#0E1519", border: "none", padding: "10px 18px", borderRadius: 8, fontWeight: 600 };
const btnSecondary = { background: "none", color: "#8B95A1", border: "1px solid #2A3138", padding: "10px 18px", borderRadius: 8, fontWeight: 500 };
