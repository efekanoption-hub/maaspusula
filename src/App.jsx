import React, { useState, useEffect, useRef, useMemo } from 'react';
import {
  Calculator, Wallet, Percent, ShieldCheck, TrendingDown, Info, Save,
  History, Download, Trash2, Loader2, CheckCircle2, Settings, Share2,
  TrendingUp, AlertCircle, ChevronDown, ChevronUp, Image as ImageIcon,
  UserCheck, Heart, Briefcase, Calendar, Clock, BarChart3, PieChart as PieChartIcon
} from 'lucide-react';
import { supabase } from './lib/supabase';
import * as XLSX from 'xlsx';
import { toPng } from 'html-to-image';
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell,
  AreaChart, Area
} from 'recharts';

// 2026 Verileri (Güncel Mevzuat)
const ASGARI_UCRET_BRUT = 33030.00;
const ASGARI_UCRET_GV_ISTISNA = 4211.33;
const ASGARI_UCRET_DV_ISTISNA = 250.70;
const SGK_TAVAN = 297270.00;
const KIDEM_TAVANI = 64948.77;

const SGK_ISCI_ORAN = 0.14;
const ISS_ISCI_ORAN = 0.01;
const SGK_ISVEREN_ORAN = 0.155;
const ISS_ISVEREN_ORAN = 0.02;

// İstisnalar
const GUNLUK_YEMEK_ISTISNA_GV = 300.00;
const GUNLUK_YEMEK_ISTISNA_SGK = 198.00;
const GUNLUK_YOL_ISTISNA_GV = 158.00;

// 2026 Gelir Vergisi Dilimleri
const VERGI_DILIMLERI = [
  { limit: 190000, oran: 0.15 },
  { limit: 400000, oran: 0.20 },
  { limit: 1500000, oran: 0.27 },
  { limit: 5300000, oran: 0.35 },
  { limit: Infinity, oran: 0.40 }
];

const ENGELLILIK_INDIRIMI = {
  0: 0,
  1: 12000,
  2: 7000,
  3: 3000
};

const AYLAR = [
  "Ocak", "Şubat", "Mart", "Nisan", "Mayıs", "Haziran",
  "Temmuz", "Ağustos", "Eylül", "Ekim", "Kasım", "Aralık"
];

const Logo = ({ size = 32, className = "" }) => (
  <div className={`logo-container flex items-center gap-2 ${className}`}>
    <div className="relative group">
      <div className="absolute -inset-1 bg-emerald-500/20 rounded-full blur opacity-0 group-hover:opacity-100 transition duration-500"></div>
      <img src="/logo.png" alt="MaasPusula Logo" style={{ width: size, height: size }} className="relative rounded-lg object-contain logo-icon" />
    </div>
    <span className="text-xl font-black tracking-tight bg-gradient-to-r from-white to-navy-300 bg-clip-text text-transparent">
      Maas<span className="text-emerald-500">Pusula</span>
    </span>
  </div>
);

const SplashScreen = ({ isVisible }) => (
  <div className={`splash-screen ${!isVisible ? 'hide' : ''}`}>
    <div className="splash-logo flex flex-col items-center justify-center">
      <div className="relative flex items-center justify-center">
        <div className="absolute -inset-10 bg-emerald-500/20 rounded-full blur-3xl animate-pulse"></div>
        <img src="/logo.png" alt="MaasPusula Splash Logo" className="w-32 h-32 md:w-48 md:h-48 relative object-contain drop-shadow-[0_0_30px_rgba(16,185,129,0.4)]" />
      </div>
      <h1 className="mt-8 text-3xl md:text-5xl font-black tracking-tighter text-white">Maas<span className="text-emerald-500">Pusula</span></h1>
    </div>
    <div className="loader-container"><div className="loader-bar" /></div>
    <p className="mt-6 text-navy-400 text-[10px] md:text-sm font-bold uppercase tracking-[0.4em] animate-pulse">Veriler Hazırlanıyor</p>
  </div>
);

function App() {
  const [showSplash, setShowSplash] = useState(true);
  const [activeTab, setActiveTab] = useState('isci');
  const [brutMaas, setBrutMaas] = useState(50000);
  const [netMaasInput, setNetMaasInput] = useState(0);
  const [lastChanged, setLastChanged] = useState('brut');
  const [savedRecords, setSavedRecords] = useState([]);
  const [isSaving, setIsSaving] = useState(false);
  const [isLoadingRecords, setIsLoadingRecords] = useState(false);
  const [label, setLabel] = useState('');

  // Settings
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth());
  const [showSettings, setShowSettings] = useState(false);
  const [besActive, setBesActive] = useState(false);
  const [sendikaAidati, setSendikaAidati] = useState(0);
  const [ozelSaglikSigortasi, setOzelSaglikSigortasi] = useState(0);
  const [disabilityLevel, setDisabilityLevel] = useState(0);
  const [isEmployerDiscount, setIsEmployerDiscount] = useState(true);
  const [bonusMonths, setBonusMonths] = useState([]);
  const [includeSeveranceReserve, setIncludeSeveranceReserve] = useState(false);

  // Yemek & Yol
  const [gunlukYemek, setGunlukYemek] = useState(0);
  const [gunlukYol, setGunlukYol] = useState(0);
  const [calisilanGun, setCalisilanGun] = useState(22);

  // Tazminat State
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');

  const [isSharing, setIsSharing] = useState(false);
  const resultCardRef = useRef(null);
  const [viewMode, setViewMode] = useState('monthly');

  const getTaxForMatrah = (matrah) => {
    let tax = 0;
    let previousLimit = 0;
    for (const dilim of VERGI_DILIMLERI) {
      if (matrah > previousLimit) {
        const taxableAmount = Math.min(matrah, dilim.limit) - previousLimit;
        tax += taxableAmount * dilim.oran;
        previousLimit = dilim.limit;
      } else { break; }
    }
    return tax;
  };

  const calculateDetailedMonth = (monthlyBrut, startCumulativeMatrah) => {
    const totalBrut = Number(monthlyBrut);
    const yemekOdemesi = Number(gunlukYemek) * Number(calisilanGun);
    const yolOdemesi = Number(gunlukYol) * Number(calisilanGun);

    const yemekTableSGK = Math.max(0, Number(gunlukYemek) - GUNLUK_YEMEK_ISTISNA_SGK) * Number(calisilanGun);
    const yemekVergiyeTabi = Math.max(0, Number(gunlukYemek) - GUNLUK_YEMEK_ISTISNA_GV) * Number(calisilanGun);
    const yolVergiyeTabi = Math.max(0, Number(gunlukYol) - GUNLUK_YOL_ISTISNA_GV) * Number(calisilanGun);

    const sgkMatrah = Math.min(totalBrut + yemekTableSGK, SGK_TAVAN);
    const sgkIsci = sgkMatrah * SGK_ISCI_ORAN;
    const issIsci = sgkMatrah * ISS_ISCI_ORAN;

    let monthlyMatrah = (totalBrut - (sgkIsci + issIsci)) + yemekVergiyeTabi + yolVergiyeTabi;
    monthlyMatrah = Math.max(0, monthlyMatrah - ENGELLILIK_INDIRIMI[disabilityLevel]);

    const endCumulativeMatrah = startCumulativeMatrah + monthlyMatrah;
    const taxPrev = getTaxForMatrah(startCumulativeMatrah);
    const taxCurrent = getTaxForMatrah(endCumulativeMatrah);
    const hamGelirVergisi = taxCurrent - taxPrev;
    const finalGelirVergisi = Math.max(0, hamGelirVergisi - ASGARI_UCRET_GV_ISTISNA);

    const hamDamgaVergisi = (totalBrut + yemekOdemesi + yolOdemesi) * 0.00759;
    const finalDamgaVergisi = Math.max(0, hamDamgaVergisi - ASGARI_UCRET_DV_ISTISNA);

    const besKesintisi = besActive ? (totalBrut * 0.03) : 0;
    const digerKesintiler = Number(sendikaAidati) + Number(ozelSaglikSigortasi);

    const sgkIsverenOranFinal = isEmployerDiscount ? 0.155 : 0.205;
    const sgkIsveren = sgkMatrah * sgkIsverenOranFinal;
    const issIsveren = sgkMatrah * ISS_ISVEREN_ORAN;

    const toplamKesinti = sgkIsci + issIsci + finalGelirVergisi + finalDamgaVergisi + besKesintisi + digerKesintiler;
    const netMaas = totalBrut + yemekOdemesi + yolOdemesi - toplamKesinti;
    const baseMaliyet = totalBrut + yemekOdemesi + yolOdemesi + sgkIsveren + issIsveren;
    const tazminatKarsiligi = includeSeveranceReserve ? (totalBrut / 12) : 0;
    const toplamMaliyet = baseMaliyet + tazminatKarsiligi;

    let currentBracket = 15;
    let nextBracketLimit = VERGI_DILIMLERI[0].limit;
    for (const dilim of VERGI_DILIMLERI) {
      if (endCumulativeMatrah > dilim.limit) continue;
      currentBracket = Math.round(dilim.oran * 100);
      nextBracketLimit = dilim.limit;
      break;
    }

    return {
      sgkIsci, issIsci, gelirVergisi: finalGelirVergisi, damgaVergisi: finalDamgaVergisi,
      toplamKesinti, netMaas, sgkIsveren, issIsveren, toplamMaliyet, baseMaliyet, tazminatKarsiligi,
      brutMaas: totalBrut, monthlyMatrah, cumulativeMatrah: endCumulativeMatrah,
      currentBracket, nextBracketLimit, besKesintisi, yemekOdemesi, yolOdemesi
    };
  };

  const yearlyProjection = useMemo(() => {
    let results = [];
    let currentCumulative = 0;
    for (let m = 0; m < 12; m++) {
      const isBonusMonth = bonusMonths.includes(m);
      const monthlyBrut = Number(brutMaas) + (isBonusMonth ? Number(brutMaas) : 0);
      const mRes = calculateDetailedMonth(monthlyBrut, currentCumulative);
      results.push({
        id: m,
        name: AYLAR[m],
        net: mRes.netMaas,
        maliyet: mRes.toplamMaliyet,
        bracket: mRes.currentBracket,
        fullRes: mRes
      });
      currentCumulative = mRes.cumulativeMatrah;
    }
    return results;
  }, [brutMaas, gunlukYemek, gunlukYol, calisilanGun, besActive, sendikaAidati, ozelSaglikSigortasi, disabilityLevel, isEmployerDiscount, bonusMonths, includeSeveranceReserve]);

  const currentMonthResults = useMemo(() => {
    return yearlyProjection[selectedMonth].fullRes;
  }, [yearlyProjection, selectedMonth]);

  const yearlyStats = useMemo(() => {
    const totalNet = yearlyProjection.reduce((sum, m) => sum + m.net, 0);
    const totalMaliyet = yearlyProjection.reduce((sum, m) => sum + m.maliyet, 0);
    const avgNet = totalNet / 12;
    return { totalNet, totalMaliyet, avgNet };
  }, [yearlyProjection]);

  const calculateFromNet = (net) => {
    let low = net / 2; let high = net * 5; let guess = (low + high) / 2;
    for (let i = 0; i < 30; i++) {
      const result = calculateDetailedMonth(guess, 0);
      if (result.netMaas < net) low = guess; else high = guess;
      guess = (low + high) / 2;
    }
    setBrutMaas(guess.toFixed(2));
  };

  const calculateTazminat = () => {
    if (!startDate || !endDate) return null;
    const start = new Date(startDate);
    const end = new Date(endDate);
    if (end < start) return null;

    const diffTime = Math.abs(end - start);
    const totalDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
    const years = Math.floor(totalDays / 365);
    const months = Math.floor((totalDays % 365) / 30);
    const days = totalDays % 30;

    const kidemMatrah = Math.min(brutMaas, KIDEM_TAVANI);
    const brutKidem = kidemMatrah * (years + (months / 12) + (days / 365));
    const damgaKidem = brutKidem * 0.00759;
    const netKidem = brutKidem - damgaKidem;

    let weekCount = 2;
    if (totalDays >= 180 && totalDays < 540) weekCount = 4;
    else if (totalDays >= 540 && totalDays < 1080) weekCount = 6;
    else if (totalDays >= 1080) weekCount = 8;

    const ihbarGun = weekCount * 7;
    const brutIhbar = (brutMaas / 30) * ihbarGun;
    const ihbarGV = brutIhbar * 0.15;
    const ihbarDV = brutIhbar * 0.00759;
    const netIhbar = brutIhbar - ihbarGV - ihbarDV;

    return { years, months, days, brutKidem, netKidem, weekCount, brutIhbar, netIhbar, totalNet: netKidem + netIhbar };
  };

  useEffect(() => {
    if (activeTab === 'tazminat') return;
    if (lastChanged === 'net') calculateFromNet(Number(netMaasInput) || 0);
    else setNetMaasInput(currentMonthResults.netMaas.toFixed(2));
  }, [brutMaas, netMaasInput, lastChanged]);

  useEffect(() => {
    const timer = setTimeout(() => setShowSplash(false), 2500);
    return () => clearTimeout(timer);
  }, []);

  const handleSave = async () => {
    setIsSaving(true);
    const record = { label: label || `Hesap - ${AYLAR[selectedMonth]} 2026`, brut_maas: currentMonthResults.brutMaas, net_maas: currentMonthResults.netMaas, isveren_maliyeti: currentMonthResults.toplamMaliyet, created_at: new Date().toISOString() };
    const { error } = await supabase.from('maas_kayitlari').insert([record]);
    if (!error) { setLabel(''); alert('Kaydedildi!'); fetchRecords(); }
    setIsSaving(false);
  };

  const fetchRecords = async () => {
    setIsLoadingRecords(true);
    const { data, error } = await supabase.from('maas_kayitlari').select('*').order('created_at', { ascending: false });
    if (!error) setSavedRecords(data);
    setIsLoadingRecords(false);
  };

  const shareResult = async () => {
    if (!resultCardRef.current) return;
    setIsSharing(true);
    try {
      const dataUrl = await toPng(resultCardRef.current, { backgroundColor: '#0f172a' });
      const blob = await (await fetch(dataUrl)).blob();
      const file = new File([blob], 'MaasPusula.png', { type: 'image/png' });
      if (navigator.share) await navigator.share({ files: [file], title: 'Maaş Pusulası' });
      else { const link = document.createElement('a'); link.download = 'MaasPusula.png'; link.href = dataUrl; link.click(); }
    } catch { }
    setIsSharing(false);
  };

  const formatCurrency = (v) => new Intl.NumberFormat('tr-TR', { style: 'currency', currency: 'TRY' }).format(v);

  const toggleBonusMonth = (idx) => {
    setBonusMonths(prev => prev.includes(idx) ? prev.filter(i => i !== idx) : [...prev, idx]);
  };

  const CustomTooltip = ({ active, payload }) => {
    if (active && payload && payload.length) {
      const data = payload[0].payload;
      return (
        <div className="bg-navy-950/90 backdrop-blur-xl p-3 rounded-2xl border border-white/10 shadow-2xl">
          <p className="text-xs font-black text-white/50 uppercase mb-1">{data.name}</p>
          <p className="text-sm font-black text-white">{formatCurrency(data.net)} <span className="text-[10px] text-emerald-400">Net</span></p>
          <p className="text-[10px] font-bold text-orange-400 mt-1">Vergi Dilimi: %{data.bracket}</p>
        </div>
      );
    }
    return null;
  };

  const tazminatResults = calculateTazminat();

  return (
    <>
      <SplashScreen isVisible={showSplash} />
      <div className={`min-h-screen bg-navy-900 text-white font-sans flex flex-col app-container ${!showSplash ? 'visible' : ''}`}>
        <nav className="p-3 md:p-6 flex flex-wrap items-center justify-between max-w-2xl mx-auto w-full gap-3">
          <Logo size={28} />
          <div className="flex items-center gap-2 md:gap-3">
            <button onClick={() => setViewMode(v => v === 'monthly' ? 'yearly' : 'monthly')} className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-full bg-navy-800 border border-white/5 text-[9px] md:text-[10px] font-black uppercase hover:bg-navy-700 transition-all shrink-0">
              {viewMode === 'monthly' ? <BarChart3 size={12} className="text-emerald-500" /> : <Calculator size={12} className="text-sky-500" />}
              <span className="hidden xs:inline">{viewMode === 'monthly' ? 'Yıllık Projeksiyon' : 'Aylık Detay'}</span>
              <span className="xs:hidden">{viewMode === 'monthly' ? 'Yıllık' : 'Aylık'}</span>
            </button>
            <button onClick={() => setShowSettings(!showSettings)} className={`p-2 rounded-full transition-all ${showSettings ? 'bg-emerald-500' : 'bg-navy-800 text-navy-400'}`}>
              <Settings size={18} />
            </button>
          </div>
        </nav>

        <main className="max-w-2xl mx-auto px-3 md:px-6 pb-16 space-y-3 md:space-y-6 flex-1 w-full">
          {showSettings && (
            <div className="glass p-4 rounded-2xl border border-emerald-500/20 bg-navy-800/40 space-y-4 animate-in slide-in-from-top-4 duration-300">
              <div className="flex items-center justify-between">
                <h3 className="text-xs font-black text-emerald-400 uppercase tracking-widest flex items-center gap-2"><Settings size={16} /> Gelişmiş Ayarlar</h3>
                <div className="flex items-center gap-2">
                  <span className="text-[9px] font-black text-navy-400">İşveren Tazminat Karşılığı (%8.33)</span>
                  <button onClick={() => setIncludeSeveranceReserve(!includeSeveranceReserve)} className={`w-10 h-5 rounded-full relative transition-all ${includeSeveranceReserve ? 'bg-emerald-500' : 'bg-navy-700'}`}>
                    <div className={`absolute top-0.5 w-4 h-4 bg-white rounded-full transition-all ${includeSeveranceReserve ? 'left-5.5' : 'left-0.5'}`} />
                  </button>
                </div>
              </div>

              <div className="space-y-3">
                <label className="text-[10px] text-navy-400 font-bold uppercase">13. Maaş / İkramiye Ayları</label>
                <div className="flex flex-wrap gap-2">
                  {AYLAR.map((ay, i) => (
                    <button key={ay} onClick={() => toggleBonusMonth(i)} className={`px-3 py-2 rounded-xl text-[10px] font-black uppercase transition-all border ${bonusMonths.includes(i) ? 'bg-emerald-500 border-emerald-400 text-white' : 'bg-navy-900/50 border-white/5 text-navy-500'}`}>{ay.slice(0, 3)}</button>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-3">
                <div className="space-y-1">
                  <label className="text-[10px] text-navy-400 font-bold uppercase">Günlük Yemek (₺)</label>
                  <input type="number" value={gunlukYemek} onChange={(e) => setGunlukYemek(e.target.value)} className="w-full bg-navy-900 border border-white/5 rounded-lg px-3 py-2 text-xs font-bold" />
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] text-navy-400 font-bold uppercase">Günlük Yol (₺)</label>
                  <input type="number" value={gunlukYol} onChange={(e) => setGunlukYol(e.target.value)} className="w-full bg-navy-900 border border-white/5 rounded-lg px-3 py-2 text-xs font-bold" />
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] text-navy-400 font-bold uppercase">Çalışılan Gün</label>
                  <input type="number" value={calisilanGun} onChange={(e) => setCalisilanGun(e.target.value)} className="w-full bg-navy-900 border border-white/5 rounded-lg px-3 py-2 text-xs font-bold text-center" />
                </div>
              </div>
            </div>
          )}

          <div className="flex p-0.5 bg-navy-800/50 rounded-xl border border-white/5 overflow-x-auto no-scrollbar scroll-smooth">
            {['isci', 'isveren', 'tazminat', 'kayitlar'].map(tab => (
              <button key={tab} onClick={() => setActiveTab(tab)} className={`flex-1 py-2 px-1.5 rounded-lg font-black text-[8px] md:text-xs uppercase transition-all whitespace-nowrap min-w-[60px] ${activeTab === tab ? 'bg-emerald-500 text-white shadow-lg' : 'text-navy-400 hover:text-white'}`}>
                {tab === 'isci' ? 'Çalışan' : tab === 'isveren' ? 'İşveren' : tab === 'tazminat' ? 'Tazminat' : 'Kayıtlar'}
              </button>
            ))}
          </div>

          {activeTab === 'kayitlar' ? (
            <div className="space-y-3">
              {isLoadingRecords ? <Loader2 className="mx-auto animate-spin" /> : savedRecords.length === 0 ? <p className="text-center text-navy-500 py-10">Kayıt bulunamadı.</p> : savedRecords.map(r => (
                <div key={r.id} className="glass p-4 rounded-2xl flex justify-between items-center group">
                  <div><h4 className="font-black text-xs">{r.label}</h4><p className="text-[9px] text-navy-500 font-bold uppercase">Net: {formatCurrency(r.net_maas)} • {new Date(r.created_at).toLocaleDateString()}</p></div>
                  <button onClick={async () => { await supabase.from('maas_kayitlari').delete().eq('id', r.id); fetchRecords(); }} className="text-navy-500 hover:text-rose-500 transition-colors"><Trash2 size={18} /></button>
                </div>
              ))}
            </div>
          ) : activeTab === 'tazminat' ? (
            <div className="space-y-4 animate-in fade-in slide-in-from-bottom-2 duration-300">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-navy-400 uppercase flex items-center gap-2"><Calendar size={12} /> İşe Giriş</label>
                  <input type="date" value={startDate} onChange={(e) => setStartDate(e.target.value)} className="w-full bg-navy-800 border-2 border-white/5 rounded-2xl p-4 font-bold outline-none focus:border-emerald-500" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-navy-400 uppercase flex items-center gap-2"><Calendar size={12} /> İşten Çıkış</label>
                  <input type="date" value={endDate} onChange={(e) => setEndDate(e.target.value)} className="w-full bg-navy-800 border-2 border-white/5 rounded-2xl p-4 font-bold outline-none focus:border-emerald-500" />
                </div>
              </div>

              {tazminatResults ? (
                <div className="space-y-4">
                  <div className="glass-emerald p-8 rounded-[1.5rem] md:rounded-[2rem] text-center relative overflow-hidden">
                    <h2 className="text-3xl md:text-6xl font-black tracking-tighter mb-2">{formatCurrency(tazminatResults.totalNet)}</h2>
                    <p className="text-emerald-400 text-[10px] md:text-xs font-black uppercase tracking-widest">Tahmini Toplam Net Alacak</p>
                    <div className="mt-6 flex justify-center gap-4 text-[9px] md:text-[10px] font-bold text-navy-300">
                      <span>{tazminatResults.years} Yıl {tazminatResults.months} Ay {tazminatResults.days} Gün Çalışma</span>
                    </div>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <InfoCard title="Kıdem Tazminatı" value={tazminatResults.netKidem} icon={<Briefcase className="text-blue-400" />} desc={`Tavan: ${formatCurrency(KIDEM_TAVANI)}`} />
                    <InfoCard title="İhbar Tazminatı" value={tazminatResults.netIhbar} icon={<Clock className="text-orange-400" />} desc={`${tazminatResults.weekCount} Haftalık Ücret`} />
                  </div>
                </div>
              ) : (
                <div className="text-center py-16 glass rounded-[2rem] border-dashed border-2 border-white/5 space-y-4">
                  <Calculator className="mx-auto text-navy-600" size={48} />
                  <p className="text-navy-400 text-sm font-medium px-4">Tarihleri girerek kıdem ve ihbar tazminatınızı hesaplayın.</p>
                </div>
              )}
            </div>
          ) : viewMode === 'yearly' ? (
            <div className="space-y-6 animate-in fade-in zoom-in-95 duration-500">
              <div className="glass p-4 md:p-6 rounded-[1.5rem] md:rounded-[2.5rem] border-white/5 bg-navy-800/30 overflow-hidden">
                <div className="flex items-center justify-between mb-4 md:mb-6">
                  <div>
                    <h3 className="text-base md:text-lg font-black tracking-tight">2026 Gelir Projeksiyonu</h3>
                    <p className="text-[9px] md:text-[10px] text-navy-500 font-bold uppercase">12 Aylık Tahmini Net Maaş Değişimi</p>
                  </div>
                  <BarChart3 className="text-emerald-500 hidden xs:block" size={24} />
                </div>

                <div className="h-48 md:h-64 w-full min-h-[192px] md:min-h-[256px]">
                  <ResponsiveContainer width="100%" height="100%" minHeight={192}>
                    <BarChart data={yearlyProjection} margin={{ top: 10, right: 0, left: -20, bottom: 0 }}>
                      <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="rgba(255,255,255,0.05)" />
                      <XAxis dataKey="name" axisLine={false} tickLine={false} tick={{ fontSize: 10, fontWeight: 900, fill: '#64748b' }} dy={10} />
                      <YAxis axisLine={false} tickLine={false} tick={{ fontSize: 10, fontWeight: 900, fill: '#64748b' }} />
                      <Tooltip content={<CustomTooltip />} cursor={{ fill: 'rgba(255,255,255,0.03)' }} />
                      <Bar dataKey="net" radius={[6, 6, 0, 0]} barSize={30}>
                        {yearlyProjection.map((entry, index) => (
                          <Cell key={`cell-${index}`} fill={entry.bracket >= 27 ? '#f97316' : entry.bracket >= 20 ? '#fbbf24' : '#10b981'} />
                        ))}
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                </div>

                <div className="flex flex-wrap justify-center gap-4 mt-6">
                  <div className="flex items-center gap-1.5"><div className="w-2.5 h-2.5 rounded-full bg-emerald-500" /><span className="text-[9px] font-black uppercase text-navy-400">%15 Dilim</span></div>
                  <div className="flex items-center gap-1.5"><div className="w-2.5 h-2.5 rounded-full bg-yellow-400" /><span className="text-[9px] font-black uppercase text-navy-400">%20 Dilim</span></div>
                  <div className="flex items-center gap-1.5"><div className="w-2.5 h-2.5 rounded-full bg-orange-500" /><span className="text-[9px] font-black uppercase text-navy-400">%27+ Dilim</span></div>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                <div className="glass p-4 rounded-[1.5rem] border-white/5 bg-emerald-500/5 text-center space-y-1">
                  <p className="text-[8px] font-black text-emerald-400 uppercase tracking-widest">Yıllık Toplam Net</p>
                  <h4 className="text-lg font-black">{formatCurrency(yearlyStats.totalNet)}</h4>
                </div>
                <div className="glass p-4 rounded-[1.5rem] border-white/5 bg-sky-500/5 text-center space-y-1">
                  <p className="text-[8px] font-black text-sky-400 uppercase tracking-widest">Ortalama Aylık Net</p>
                  <h4 className="text-lg font-black">{formatCurrency(yearlyStats.avgNet)}</h4>
                </div>
                <div className="glass p-4 rounded-[1.5rem] border-white/5 bg-orange-500/5 text-center space-y-1">
                  <p className="text-[8px] font-black text-orange-400 uppercase tracking-widest">Yıllık Toplam Maliyet</p>
                  <h4 className="text-lg font-black">{formatCurrency(yearlyStats.totalMaliyet)}</h4>
                </div>
              </div>

              <div className="bg-navy-800/50 p-4 rounded-2xl border border-white/5 flex items-start gap-3">
                <AlertCircle className="text-sky-500 shrink-0" size={18} />
                <p className="text-[9px] font-medium text-navy-300 leading-relaxed italic">
                  Not: Yıllık projeksiyon hesaplamasında {bonusMonths.length > 0 ? `${bonusMonths.length} aylık ikramiye` : 'ikramiye bulunmayan'} düz maaş ve kümülatif vergi geçişleri baz alınmıştır. Gerçek maaşınızda ay içindeki ek mesai veya rapor gibi durumlar fark yaratabilir.
                </p>
              </div>
            </div>
          ) : (
            <div className="space-y-4 animate-in fade-in slide-in-from-bottom-2 duration-300">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-navy-400 uppercase flex items-center gap-2"><Wallet size={12} className="text-emerald-500" /> Brüt Maaş</label>
                  <input type="number" value={brutMaas} onChange={(e) => { setLastChanged('brut'); setBrutMaas(e.target.value); }} className="w-full bg-navy-800/70 border-2 border-white/10 rounded-2xl p-4 text-2xl sm:text-3xl font-black outline-none focus:border-emerald-500" />
                </div>
                <div className="space-y-2">
                  <label className="text-[10px] font-black text-navy-400 uppercase flex items-center gap-2"><Wallet size={12} className="text-sky-500" /> Net Maaş</label>
                  <input type="number" value={netMaasInput} onChange={(e) => { setLastChanged('net'); setNetMaasInput(e.target.value); }} className="w-full bg-navy-800/70 border-2 border-white/10 rounded-2xl p-4 text-2xl sm:text-3xl font-black outline-none focus:border-emerald-500" />
                </div>
              </div>

              <div className="glass p-4 rounded-[1.5rem] border-orange-500/20">
                <div className="flex justify-between items-center mb-4">
                  <div className="flex items-center gap-2">
                    <select value={selectedMonth} onChange={(e) => setSelectedMonth(Number(e.target.value))} className="bg-navy-900 text-white text-[10px] font-black uppercase px-2 py-1 rounded-lg border border-white/10">{AYLAR.map((ay, i) => <option key={ay} value={i}>{ay}</option>)}</select>
                    <span className="text-[11px] font-black uppercase text-navy-300">Vergi Dilimi: %{currentMonthResults.currentBracket}</span>
                  </div>
                  {bonusMonths.includes(selectedMonth) && <span className="text-[8px] font-black bg-emerald-500 text-white px-2 py-0.5 rounded-full uppercase">İKRAMİYE AYI</span>}
                </div>
                <div className="h-2 w-full bg-navy-900 rounded-full overflow-hidden mb-2">
                  <div className="h-full bg-gradient-to-r from-orange-400 to-orange-600 transition-all duration-1000" style={{ width: `${Math.min(100, (currentMonthResults.cumulativeMatrah / currentMonthResults.nextBracketLimit) * 100)}%` }}></div>
                </div>
                <p className="text-[9px] text-navy-500 font-bold uppercase">Kümülatif Matrah: {formatCurrency(currentMonthResults.cumulativeMatrah)} / Üst Dilim: {formatCurrency(currentMonthResults.nextBracketLimit)}</p>
              </div>

              <div ref={resultCardRef} className="bg-navy-900 rounded-[1.5rem] md:rounded-[2.5rem] p-1 shadow-2xl overflow-hidden relative">
                <section className={`glass-emerald rounded-[1.3rem] md:rounded-[2.3rem] p-6 md:p-12 text-center space-y-2 relative`}>
                  <p className="text-emerald-400 font-black tracking-[0.2em] md:tracking-[0.3em] text-[9px] md:text-[10px] uppercase">{activeTab === 'isci' ? 'Aylık Ele Geçen Toplam' : 'Aylık İşveren Maliyeti'}</p>
                  <h2 className="text-3xl sm:text-4xl md:text-7xl font-black tracking-tighter break-all">{formatCurrency(activeTab === 'isci' ? currentMonthResults.netMaas : currentMonthResults.toplamMaliyet)}</h2>
                  <div className="flex justify-center flex-wrap gap-1.5 md:gap-2 pt-3 md:pt-4">
                    {currentMonthResults.yemekOdemesi > 0 && <span className="bg-navy-900/50 px-2 md:px-3 py-1 rounded-full text-[8px] md:text-[10px] font-bold text-navy-300">Yemek: {formatCurrency(currentMonthResults.yemekOdemesi)}</span>}
                    {currentMonthResults.yolOdemesi > 0 && <span className="bg-navy-900/50 px-2 md:px-3 py-1 rounded-full text-[8px] md:text-[10px] font-bold text-navy-300">Yol: {formatCurrency(currentMonthResults.yolOdemesi)}</span>}
                    {includeSeveranceReserve && activeTab === 'isveren' && <span className="bg-navy-900/50 px-2 md:px-3 py-1 rounded-full text-[8px] md:text-[10px] font-bold text-orange-400">Kıdem: {formatCurrency(currentMonthResults.tazminatKarsiligi)}</span>}
                  </div>
                </section>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 md:gap-3">
                <input type="text" value={label} onChange={(e) => setLabel(e.target.value)} placeholder="Kaydet..." className="sm:col-span-1 bg-navy-800/70 border border-white/10 rounded-xl md:rounded-2xl px-4 md:px-5 py-3 md:py-4 text-xs md:text-sm font-bold outline-none focus:border-emerald-500" />
                <div className="flex gap-2 sm:col-span-2">
                  <button onClick={handleSave} className="flex-1 bg-navy-800 p-3 md:p-4 rounded-xl md:rounded-2xl border border-white/10 text-emerald-400 flex items-center justify-center transition-all active:scale-95"><Save size={20} /></button>
                  <button onClick={shareResult} className="flex-[3] bg-emerald-500 px-4 md:px-6 py-3 md:py-4 rounded-xl md:rounded-2xl font-black text-xs md:text-sm shadow-xl shadow-emerald-500/30 flex items-center justify-center gap-2 active:scale-95"><Share2 size={18} /> Paylaş</button>
                </div>
              </div>

              <div className="grid grid-cols-2 md:grid-cols-3 gap-3 md:gap-4">
                <InfoCard title="Gelir Vergisi" value={currentMonthResults.gelirVergisi} icon={<TrendingDown className="text-orange-400" />} />
                <InfoCard title="SGK İşçi" value={currentMonthResults.sgkIsci} icon={<ShieldCheck className="text-blue-400" />} />
                <InfoCard title="Damga Vergisi" value={currentMonthResults.damgaVergisi} icon={<TrendingDown className="text-rose-400" />} />
                {activeTab === 'isveren' && (
                  <React.Fragment>
                    <InfoCard title="SGK İşveren" value={currentMonthResults.sgkIsveren} icon={<ShieldCheck className="text-emerald-400" />} />
                    <InfoCard title="İşsizlik İşv." value={currentMonthResults.issIsveren} icon={<Percent className="text-sky-400" />} />
                  </React.Fragment>
                )}
              </div>
            </div>
          )}

          <section className="glass rounded-[1.2rem] md:rounded-[1.5rem] p-4 md:p-5 flex gap-3 md:gap-4 items-center border-white/5 bg-navy-800/30">
            <Info className="text-emerald-500 shrink-0" size={20} />
            <p className="text-[9px] md:text-xs text-navy-300 font-medium italic leading-relaxed">Asgari ücret istisnaları ve kümülatif vergi geçişleri {AYLAR[selectedMonth]} ayı için {currentMonthResults.currentBracket === 20 ? '(%20 diliminde)' : '(%15 diliminde)'} hesaplanmıştır. Yıllık görünümde vergi dilimi geçişlerini inceleyebilirsiniz.</p>
          </section>
        </main>
        <footer className="p-8 text-center text-[10px] text-navy-600 font-black uppercase tracking-widest">Maaş Pusula • {new Date().getFullYear()}</footer>
      </div>
    </>
  );
}

function InfoCard({ title, value, icon, desc }) {
  return (
    <div className="glass p-5 rounded-[2rem] hover:bg-navy-800/80 transition-all border-white/5 space-y-3">
      <div className="p-2 bg-navy-900 rounded-xl w-fit">{icon}</div>
      <div>
        <h3 className="text-[9px] font-black uppercase text-navy-400 tracking-widest">{title}</h3>
        <p className="text-lg font-black text-white">{new Intl.NumberFormat('tr-TR', { style: 'currency', currency: 'TRY' }).format(value)}</p>
        {desc && <p className="text-[8px] text-navy-500 font-bold uppercase">{desc}</p>}
      </div>
    </div>
  );
}

export default App;
