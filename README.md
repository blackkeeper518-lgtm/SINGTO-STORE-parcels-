# SINGTO-STORE-parcels-import React, { useState, useEffect } from "react";

// ==========================================
// [VERSION]: V50.0 - SINGTO OVERLORD (GOLD NEON)
// [STATUS]: PRODUCTION READY / DASHBOARD DEFAULT
// [SECRET]: Triple-click "PRECISION_LOCK" to trigger Secret Ops
// [COLOR]: SINGTO STORE (GOLD-BLACK NEON)
// ==========================================
const SUPABASE_URL = "https://kpihxjmicmgwlvwjnxtg.supabase.co";
const SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtwaWh4am1pY21nd2x2d2pueHRnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODc2NTQyMjksImV4cCI6MjEwMzIzMDIyOX0.4EpXFhQtu0VhN1uFNPSlKUW5ye23OdCnSk-pBX2YdII";
const ACCESS_KEY = "1234";

export default function App() {
  const [activeTab, setActiveTab] = useState("dashboard"); 
  const [isAuthorized, setIsAuthorized] = useState(false);
  const [passInput, setPassInput] = useState("");
  const [showGate, setShowGate] = useState(false);
  
  const [supabaseClient, setSupabaseClient] = useState(null);
  const [loading, setLoading] = useState(false);
  const [currentTime, setCurrentTime] = useState(new Date());
  
  // Stats & Data
  const [phone, setPhone] = useState("");
  const [trackingData, setTrackingData] = useState(null);
  const [orders, setOrders] = useState([]);
  const [stats, setStats] = useState({ totalOrders: 0, totalCOD: 0, todayCOD: 0, aov: 0 });
  const [importStatus, setImportStatus] = useState({ processing: false, count: 0, logs: [] });

  // 🕒 Initialization
  useEffect(() => {
    const scripts = [
      'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2',
      'https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js'
    ];
    let loadedCount = 0;
    scripts.forEach(src => {
      const script = document.createElement('script');
      script.src = src;
      script.async = true;
      script.onload = () => {
        loadedCount++;
        if (loadedCount === scripts.length && window.supabase) {
          // @ts-ignore
          setSupabaseClient(window.supabase.createClient(SUPABASE_URL, SUPABASE_KEY));
        }
      };
      document.body.appendChild(script);
    });
    const timer = setInterval(() => setCurrentTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  // 🔒 Fetch Logic
  const fetchAdminData = async () => {
    if (!supabaseClient || !isAuthorized) return;
    setLoading(true);
    try {
      const { data, error } = await supabaseClient
        .from("parcels")
        .select("*")
        .order("created_at", { ascending: false });
      if (!error && data) {
        setOrders(data);
        const totalCOD = data.reduce((sum, item) => sum + Number(item.cod_amount || 0), 0);
        
        const todayStr = new Date().toISOString().split('T')[0];
        const todayCOD = data
          .filter(item => item.created_at && item.created_at.startsWith(todayStr))
          .reduce((sum, item) => sum + Number(item.cod_amount || 0), 0);
        const totalOrders = data.length;
        const aov = totalOrders > 0 ? totalCOD / totalOrders : 0;
        setStats({ totalOrders, totalCOD, todayCOD, aov });
      }
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (isAuthorized) fetchAdminData();
  }, [supabaseClient, isAuthorized]);

  const handleTabChange = (tab) => {
    if ((tab === 'dashboard' || tab === 'ops') && !isAuthorized) {
      setShowGate(true);
      return;
    }
    setActiveTab(tab);
    setShowGate(false);
  };

  const verifyAccess = () => {
    if (passInput === ACCESS_KEY) {
      setIsAuthorized(true);
      setShowGate(false);
      setPassInput("");
      setActiveTab("dashboard");
    } else {
      setPassInput("");
    }
  };

  const cleanName = (n) => n ? n.replace(/^[0-9\s.]+/, "").replace(/^(คุณ|นาย|นางสาว|นาง|👤)\s*/gi, "").trim() : "ลูกค้า VIP";

  // 🛰️ Search Tracking
  const handleSearch = async () => {
    if (!supabaseClient || phone.trim().length < 4) return;
    setLoading(true);
    setTrackingData(null);
    try {
      const { data } = await supabaseClient
        .from("parcels")
        .select("consignee, tracking_number, cod_amount, address")
        .or(`phone_number.ilike.%${phone.trim()}%,tracking_number.ilike.%${phone.trim()}%`)
        .order("created_at", { ascending: false })
        .limit(1)
        .maybeSingle();
      if (data) setTrackingData(data);
    } finally {
      setLoading(false);
    }
  };

  // 📥 Dynamic Auto-Clean & Transform Flash File -> Parcels Table
  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    if (!file || !supabaseClient) return;

    setImportStatus({ processing: true, count: 0, logs: ["📡 เชื่อมต่อคลังแสงอาณาจักรสิงโต (SINGTO STORE)..."] });

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        // @ts-ignore
        const workbook = window.XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        // @ts-ignore
        const rows = window.XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        if (rows.length === 0) throw new Error("ไฟล์ว่างเปล่า!");

        // Dynamic Header Mapping
        const headers = rows[0].map(h => String(h || "").trim());
        
        const idxTracking = headers.findIndex(h => h.includes("เลขพัสดุ") || h.includes("tracking"));
        const idxPhone = headers.findIndex(h => h.includes("เบอร์ผู้รับ") || h.includes("phone"));
        const idxName = headers.findIndex(h => h.includes("ชื่อผู้รับ") || h.includes("consignee"));
        const idxAddress = headers.findIndex(h => h.includes("ที่อยู่ที่รับ") || h.includes("address"));
        const idxCod = headers.findIndex(h => h.includes("ยอดCOD") || h.includes("cod"));

        const batch = [];

        rows.forEach((row, i) => {
          if (i === 0) return; // Skip Header

          const rawTrack = idxTracking !== -1 ? String(row[idxTracking] || "") : String(row[3] || "");
          const track = rawTrack.trim().toUpperCase();

          if (!track.startsWith("TH")) return;

          const rawPhone = idxPhone !== -1 ? String(row[idxPhone] || "") : String(row[11] || "");
          const phoneNum = rawPhone.replace(/\D/g, "");

          const rawName = idxName !== -1 ? String(row[idxName] || "") : String(row[10] || "");
          const rawAddr = idxAddress !== -1 ? String(row[idxAddress] || "") : String(row[13] || "");
          const rawCod = idxCod !== -1 ? String(row[idxCod] || "0") : String(row[20] || "0");

          if (phoneNum.length >= 8) {
            batch.push({
              tracking_number: track,
              phone_number: phoneNum,
              consignee: cleanName(rawName),
              address: String(rawAddr),
              cod_amount: parseFloat(String(rawCod).replace(/,/g, "")) || 0
            });
          }
        });

        if (batch.length === 0) throw new Error("ไม่พบรายการพัสดุ TH ที่สมบูรณ์!");

        // 🎯 Bulk Upsert to 'parcels' table
        const { error } = await supabaseClient.from("parcels").upsert(batch, { onConflict: 'tracking_number' });
        if (error) throw error;

        setImportStatus({ 
          processing: false, 
          count: batch.length, 
          logs: [`🔥 บรรจุเข้าตาราง parcels สำเร็จ ${batch.length} พิกัด!`] 
        });
        fetchAdminData();

      } catch (err) {
        setImportStatus({ processing: false, count: 0, logs: [`⚠️ ข้อผิดพลาด: ${err.message}`] });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  return (
    <div className="min-h-screen bg-[#050400] text-white font-sans selection:bg-[#ffb703] overflow-x-hidden relative">
      <style dangerouslySetInnerHTML={{ __html: `
        @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700;900&family=Bai+Jamjuree:wght@300;400;600;700&display=swap');
        .cyber-font { font-family: 'Orbitron', sans-serif; }
        .thai-font { font-family: 'Bai Jamjuree', sans-serif; }
        .glass-panel { background: rgba(20, 15, 0, 0.75); backdrop-filter: blur(20px); border: 1px solid rgba(255, 183, 3, 0.25); }
        .neon-glow { text-shadow: 0 0 10px #ffb703, 0 0 20px #ffb703; }
        
        @property --angle { syntax: '<angle>'; initial-value: 0deg; inherits: false; }
        .neon-border-running {
          position: relative;
          border: 2px solid transparent;
          border-image: conic-gradient(from var(--angle), transparent 70%, #ffb703) 1;
          animation: rotate-angle 3s linear infinite;
        }
        @keyframes rotate-angle { to { --angle: 360deg; } }
      `}} />
      <div className="relative z-10 max-w-6xl mx-auto px-4 py-8 md:py-12 pb-32">
        {/* 🕒 Header Ops */}
        <div className="flex justify-between items-center mb-8">
            <div className="space-y-1">
                <p className="cyber-font text-[9px] text-yellow-600/70 tracking-[0.3em] font-black italic">SINGTO_STORE_OS: ONLINE</p>
                {isAuthorized && (
                    <button onClick={() => setIsAuthorized(false)} className="text-[8px] cyber-font text-red-500 font-black tracking-widest uppercase hover:underline">Terminate_Session</button>
                )}
            </div>
            <div className="neon-border-running px-6 py-3 bg-black/80 rounded-sm shadow-[0_0_30px_rgba(255,183,3,0.15)]">
                <p className="cyber-font text-[9px] text-[#ffb703] tracking-[0.2em] font-black mb-1 opacity-70 uppercase">Singto_Time</p>
                <p className="cyber-font text-xl md:text-3xl font-black text-white tracking-widest leading-none">
                    {currentTime.toLocaleTimeString('en-US', { hour12: false })}
                </p>
            </div>
        </div>

        {/* 🧭 Navigation */}
        <div className="flex justify-center mb-16">
          <div className="inline-flex glass-panel p-1.5 rounded-3xl border-[#ffb703]/30 shadow-[0_0_40px_rgba(0,0,0,0.5)]">
            {[
              { id: 'dashboard', label: 'STATS', icon: '📊' },
              { id: 'tracking', label: 'TRACK', icon: '🛰️' },
              { id: 'ops', label: 'OPS', icon: '📥' }
            ].map(tab => (
              <button
                key={tab.id}
                onClick={() => handleTabChange(tab.id)}
                className={`px-8 py-4 rounded-2xl cyber-font text-[10px] md:text-xs font-black tracking-[0.3em] transition-all duration-500 ${activeTab === tab.id ? 'bg-[#ffb703] text-black font-extrabold shadow-[0_0_25px_#ffb703] scale-105' : 'text-gray-400 hover:text-white'}`}
              >
                {tab.icon} {tab.label}
              </button>
            ))}
          </div>
        </div>

        {/* 🛡️ SECURITY GATE */}
        {showGate && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center bg-[#050400]/98 backdrop-blur-3xl">
            <div className="glass-panel p-16 md:p-24 rounded-[3rem] border-[#ffb703]/50 text-center max-w-lg w-full mx-4 shadow-[0_0_100px_rgba(0,0,0,1)]">
              <h2 className="cyber-font text-2xl md:text-3xl font-black text-white mb-12 italic tracking-tighter uppercase">Authority_Verification</h2>
              <input 
                type="password"
                value={passInput}
                onChange={(e) => setPassInput(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && verifyAccess()}
                className="w-full bg-black/50 border-b-4 border-[#ffb703] p-6 text-center text-5xl cyber-font tracking-[0.8em] outline-none text-[#ffb703] mb-12"
                placeholder="****"
                autoFocus
              />
              <div className="flex gap-6">
                <button onClick={() => setShowGate(false)} className="flex-1 py-4 cyber-font text-[10px] text-gray-500 uppercase tracking-widest font-black">Abort</button>
                <button onClick={verifyAccess} className="flex-1 py-4 bg-[#ffb703] text-black cyber-font text-[10px] font-black rounded-xl shadow-[0_0_30px_#ffb703] uppercase">Authorize</button>
              </div>
            </div>
          </div>
        )}

        {/* --- DASHBOARD (HOME) --- */}
        {activeTab === "dashboard" && isAuthorized && (
          <div className="animate-in fade-in duration-700 space-y-10">
            <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
              <div className="neon-border-running glass-panel rounded-3xl p-8 transition-all">
                <p className="cyber-font text-[9px] text-gray-400 tracking-widest uppercase font-black">Total_Wealth</p>
                <p className="thai-font text-4xl font-black italic text-white mt-4 tracking-tighter">฿{stats.totalCOD.toLocaleString()}</p>
              </div>
              <div className="neon-border-running glass-panel rounded-3xl p-8 border-[#ffb703] transition-all">
                <p className="cyber-font text-[9px] text-[#ffb703] tracking-widest uppercase font-black italic">Today_Revenue</p>
                <p className="thai-font text-4xl font-black italic text-white mt-4 tracking-tighter">฿{stats.todayCOD.toLocaleString()}</p>
              </div>
              <div className="glass-panel rounded-3xl p-8 border-l-4 border-yellow-500/30">
                <p className="cyber-font text-[9px] text-gray-400 tracking-widest uppercase font-black">Total_Conquests</p>
                <p className="thai-font text-4xl font-black italic text-[#ffb703] mt-4 tracking-tighter">{stats.totalOrders.toLocaleString()}</p>
              </div>
              <div className="glass-panel rounded-3xl p-8 border-l-4 border-yellow-500/30">
                <p className="cyber-font text-[9px] text-gray-400 tracking-widest uppercase font-black">Average_Bill</p>
                <p className="thai-font text-4xl font-black italic text-white mt-4 tracking-tighter">฿{Math.round(stats.aov).toLocaleString()}</p>
              </div>
            </div>
          </div>
        )}

        {/* --- TRACKING TAB --- */}
        {activeTab === "tracking" && (
          <div className="animate-in fade-in duration-500 max-w-2xl mx-auto glass-panel p-10 md:p-14 rounded-[3rem]">
            <h2 className="cyber-font text-center text-xl text-[#ffb703] font-black tracking-[0.3em] uppercase mb-8">SINGTO VIP TRACKING</h2>
            <div className="flex gap-4 mb-8">
              <input 
                type="text" 
                value={phone} 
                onChange={(e) => setPhone(e.target.value)} 
                onKeyDown={(e) => e.key === 'Enter' && handleSearch()}
                placeholder="กรอกเบอร์โทรศัพท์..." 
                className="flex-1 bg-black/60 border border-[#ffb703]/40 p-5 rounded-2xl text-center text-xl thai-font outline-none focus:border-[#ffb703]"
              />
              <button onClick={handleSearch} className="px-8 bg-[#ffb703] text-black font-black cyber-font text-xs rounded-2xl shadow-[0_0_20px_#ffb703]">SEARCH</button>
            </div>
            {trackingData && (
              <div className="bg-black/80 p-6 rounded-2xl border border-yellow-500/30 space-y-3 thai-font text-center">
                <p className="text-gray-400 text-sm">ผู้รับ: <span className="text-white font-bold">{trackingData.consignee}</span></p>
                <p className="text-yellow-400 text-2xl font-black cyber-font tracking-widest">{trackingData.tracking_number}</p>
                <p className="text-gray-400 text-sm">ยอด COD: <span className="text-emerald-400 font-bold">฿{trackingData.cod_amount}</span></p>
              </div>
            )}
          </div>
        )}

        {/* --- OPS UPLOAD TAB --- */}
        {activeTab === "ops" && isAuthorized && (
          <div className="animate-in fade-in duration-500 max-w-2xl mx-auto glass-panel p-12 rounded-[3rem] text-center border-dashed border-2 border-[#ffb703]">
            <h2 className="cyber-font text-xl text-[#ffb703] font-black tracking-widest mb-6 uppercase">DROP FLASH EXCEL PAYLOAD</h2>
            <p className="thai-font text-gray-400 text-sm mb-8">รองรับไฟล์ Flash Express (.xlsx / .csv) ยัดเข้าตาราง parcels ทันทีใน 3 วินาที</p>
            <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} className="hidden" id="fileInput" />
            <label htmlFor="fileInput" className="cursor-pointer inline-block px-10 py-5 bg-[#ffb703] text-black font-black cyber-font text-xs rounded-2xl shadow-[0_0_30px_#ffb703] hover:scale-105 transition-all">
              CHOOSE FILE PAYLOAD
            </label>
            {importStatus.logs.length > 0 && (
              <div className="mt-8 p-4 bg-black/80 rounded-xl thai-font text-xs text-yellow-400">
                {importStatus.logs.map((log, idx) => <p key={idx}>{log}</p>)}
              </div>
            )}
          </div>
        )}

      </div>
    </div>
  );
}