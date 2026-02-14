import React, { useState, useEffect, useMemo, useRef } from 'react';
import DataSheet from './components/DataSheet';
import DashboardAnalysis from './components/DashboardAnalysis';
import RSBTCellReport from './components/RSBTCellReport';
import BookingReport from './components/BookingReport';
import { parseExcelFile, fetchMasterDataFromUrl, fetchBookingDataFromUrl } from './utils';
import { BookingDataRow, RawBookingRow, MasterDataRow } from './types';
import { MASTER_DATA as DEFAULT_MASTER_DATA } from './constants';
import { LayoutDashboard, Table, UploadCloud, AlertCircle, RefreshCw, Link as LinkIcon, CalendarDays, FileSpreadsheet, FileText, ClipboardList, ArrowLeft } from 'lucide-react';

const MASTER_DATA_URL = 'https://docs.google.com/spreadsheets/d/1sqgOjtJ5uaiI6qIG_LZMZ-D0yaUhJG_FVxZLDeQORFA/export?format=csv';

const ANNUAL_REPORT = {
  name: 'FY 2025-26 (Complete Year)',
  url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=1746788119'
};

const MONTHLY_REPORTS = [
    { name: 'April 2025', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=1286396342' },
    { name: 'May 2025', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=1250449318' },
    { name: 'June 2025', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=1901550934' },
    { name: 'July 2025', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=727421896' },
    { name: 'Aug 2025', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=1386667955' },
    { name: 'Sept 2025', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=1678804497' },
    { name: 'Oct 2025', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=1866207457' },
    { name: 'Nov 2025', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=0' },
    { name: 'Dec 2025', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=1401133633' },
    { name: 'Jan 2026', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=1660855752' },
    { name: 'Feb 2026', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=303541532' },
    { name: 'March 2026', url: 'https://docs.google.com/spreadsheets/d/1D_d3iwih0aqEBLD1JQVZr1GUtqtsCQryPT-WoxtCfrc/edit?gid=1940660600' },
];

const convertToExportUrl = (inputUrl: string): string | null => {
  const url = inputUrl.trim();
  if (url.includes('docs.google.com/spreadsheets') && !url.includes('export?format=csv')) {
      const idMatch = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (idMatch && idMatch[1]) {
          const id = idMatch[1];
          const gidMatch = url.match(/[#&?]gid=([0-9]+)/);
          const gid = gidMatch ? gidMatch[1] : null;
          return `https://docs.google.com/spreadsheets/d/${id}/export?format=csv${gid ? `&gid=${gid}` : ''}`;
      }
  }
  return url;
};

type AppTab = 'datasheet' | 'dashboard' | 'rsbt' | 'booking';

function App() {
  const [activeTab, setActiveTab] = useState<AppTab>('datasheet');
  const [rawBookingData, setRawBookingData] = useState<RawBookingRow[]>([]);
  const [masterData, setMasterData] = useState<MasterDataRow[]>(DEFAULT_MASTER_DATA);
  const [isMasterDataLoading, setIsMasterDataLoading] = useState(false);
  const [masterDataSource, setMasterDataSource] = useState<'Default' | 'Synced'>('Default');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [customUrl, setCustomUrl] = useState('');

  useEffect(() => {
    handleSyncMasterData(true);
  }, []);

  const handleSyncMasterData = async (silent = false) => {
    if (!silent) setIsMasterDataLoading(true);
    try {
      const data = await fetchMasterDataFromUrl(MASTER_DATA_URL);
      if (data && data.length > 0) {
        setMasterData(data);
        setMasterDataSource('Synced');
        if(!silent) alert(`Successfully synced ${data.length} master data records.`);
      }
    } catch (e) {
      console.error("Sync failed", e);
    } finally {
      if (!silent) setIsMasterDataLoading(false);
    }
  };

  const enrichedBookingData: BookingDataRow[] = useMemo(() => {
    return rawBookingData.map(row => {
      const master = masterData.find(m => m.officeId === row.officeId);
      return {
        ...row,
        subDivision: master?.subDivisionName || 'Unknown Sub-Division',
        headOffice: master?.headOfficeName || 'Unknown',
        officeType: master?.officeType || 'Unknown',
        officeJurisdiction: master?.officeJurisdiction || 'Unknown Jurisdiction',
        areaType: master?.areaType || 'Rural'
      };
    });
  }, [rawBookingData, masterData]);

  const handleFileUpload = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setLoading(true);
    setError(null);
    try {
      const rawData = await parseExcelFile(file);
      if (rawData.length === 0) throw new Error("No valid booking data found.");
      setRawBookingData(rawData);
      setActiveTab('dashboard'); 
    } catch (err: any) {
      setError(err.message || "Failed to parse file");
      alert(err.message || "Failed to parse file");
    } finally {
      setLoading(false);
      event.target.value = ''; 
    }
  };

  const loadFromUrl = async (inputUrl: string) => {
    setLoading(true);
    setError(null);
    const fetchUrl = convertToExportUrl(inputUrl);
    if (!fetchUrl) {
       setLoading(false);
       setError("Invalid URL format.");
       return;
    }
    try {
      const rawData = await fetchBookingDataFromUrl(fetchUrl);
      if (rawData.length === 0) throw new Error("No data found in provided sheet.");
      setRawBookingData(rawData);
      setActiveTab('dashboard');
    } catch (err: any) {
      setError("Failed to fetch data. Ensure sheet is 'Published to web'.");
    } finally {
      setLoading(false);
    }
  };

  // Resets the app to the main screen so the user can re-upload or select new data
  const triggerReupload = () => {
    setRawBookingData([]);
    setError(null);
    setCustomUrl('');
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const hardReload = () => {
    if (confirm("This will clear all uploaded data and reload the app. Continue?")) {
      window.location.reload();
    }
  };

  return (
    <div className="min-h-screen bg-slate-100 font-sans">
      <header className="bg-[#CE2029] shadow-md sticky top-0 z-50 transition-all"> 
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center gap-3">
             <button onClick={hardReload} title="Hard Reload App" className="w-10 h-10 bg-white rounded-lg flex items-center justify-center text-[#CE2029] font-bold shadow-sm hover:scale-110 transition-transform active:scale-95">IP</button>
             <h1 className="text-xl font-bold text-white tracking-wide hidden lg:block">Mail Booking Monitoring</h1>
          </div>
          
          <div className="flex items-center gap-2 sm:gap-4 overflow-x-auto no-scrollbar">
             {rawBookingData.length > 0 && (
               <button 
                onClick={triggerReupload} 
                className="flex items-center gap-1.5 px-3 py-1.5 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-all text-[10px] font-black uppercase shadow-lg border border-blue-400"
                title="Go to main screen to re-upload"
               >
                 <ArrowLeft size={14}/> Re-upload
               </button>
             )}

             <div className="hidden xl:flex items-center gap-2 text-xs text-white/90 bg-white/10 px-3 py-1.5 rounded-full backdrop-blur-sm border border-white/20">
               <span className={`w-2 h-2 rounded-full ${masterDataSource === 'Synced' ? 'bg-green-400' : 'bg-amber-400'}`}></span>
               <span>Master Data: {masterDataSource}</span>
               <button onClick={() => handleSyncMasterData()} className="ml-2 p-1 hover:bg-white/20 rounded-full transition-colors" title="Sync Master Data">
                 <RefreshCw size={12} className={isMasterDataLoading ? 'animate-spin' : ''}/>
               </button>
             </div>

             <div className="flex gap-1 bg-black/10 p-1 rounded-lg">
                <NavTab active={activeTab === 'datasheet'} onClick={() => setActiveTab('datasheet')} label="Data" icon={<Table size={16}/>}/>
                <NavTab active={activeTab === 'dashboard'} onClick={() => setActiveTab('dashboard')} label="Analysis" icon={<LayoutDashboard size={16}/>}/>
                <NavTab active={activeTab === 'rsbt'} onClick={() => setActiveTab('rsbt'} label="RSBT Cell" icon={<ClipboardList size={16}/>}/>
                <NavTab active={activeTab === 'booking'} onClick={() => setActiveTab('booking')} label="Booking Report" icon={<FileText size={16}/>}/>
             </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {loading && (
          <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-[100] flex items-center justify-center">
            <div className="bg-white p-6 rounded-2xl shadow-2xl flex flex-col items-center gap-4">
              <RefreshCw className="text-[#CE2029] animate-spin" size={40}/>
              <span className="font-bold text-slate-800 tracking-wider">Processing Data...</span>
            </div>
          </div>
        )}

        {rawBookingData.length === 0 && (
          <div className="flex flex-col items-center justify-center min-h-[80vh]">
            <div className="bg-white p-0 rounded-2xl shadow-xl border border-slate-200 max-w-3xl w-full overflow-hidden">
              <div className="h-2 bg-[#CE2029]"></div>
              <div className="p-8 md:p-10">
                <div className="flex flex-col items-center text-center mb-8">
                    <div className="w-16 h-16 bg-[#CE2029] text-white rounded-2xl flex items-center justify-center mb-4 shadow-lg"><UploadCloud size={32} /></div>
                    <h2 className="text-3xl font-extrabold text-slate-800 mb-2">Operations Dashboard</h2>
                    <p className="text-slate-500">Select a report or upload a booking data file.</p>
                </div>
                <button onClick={() => loadFromUrl(ANNUAL_REPORT.url)} disabled={loading} className="w-full relative overflow-hidden group rounded-xl shadow-md mb-8">
                    <div className="absolute inset-0 bg-gradient-to-r from-[#CE2029] to-[#B31B24]"></div>
                    <div className="relative p-6 flex items-center justify-between">
                         <div className="flex items-center gap-4">
                            <FileSpreadsheet size={28} className="text-white/80" />
                            <div className="text-left"><h3 className="text-white font-bold">Annual Report FY 2025-26</h3></div>
                         </div>
                         <div className="bg-white text-[#CE2029] px-4 py-2 rounded-lg font-bold text-sm">View Report</div>
                    </div>
                </button>
                <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-4 mb-8">
                  {MONTHLY_REPORTS.map((month) => (
                    <button key={month.name} onClick={() => loadFromUrl(month.url)} disabled={loading} className="p-3 rounded-xl border border-slate-200 bg-slate-50 hover:border-[#CE2029] hover:bg-white transition-all text-center">
                      <span className="font-semibold text-slate-600 text-xs">{month.name}</span>
                    </button>
                  ))}
                </div>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    <label className="cursor-pointer block">
                      <div className="flex items-center gap-3 p-4 border border-dashed border-slate-300 rounded-xl bg-slate-50 hover:bg-red-50 hover:border-red-300 transition-all h-full">
                          <UploadCloud size={20} className="text-slate-400"/>
                          <span className="text-sm font-semibold text-slate-700">Upload Excel/CSV</span>
                          <input type="file" accept=".csv, .xlsx, .xls" onChange={handleFileUpload} className="hidden" />
                      </div>
                    </label>
                    <div className="flex items-center gap-2 p-1 border border-slate-300 rounded-xl bg-white shadow-sm">
                        <input type="text" value={customUrl} onChange={(e) => setCustomUrl(e.target.value)} placeholder="Google Sheet Link..." className="flex-1 min-w-0 pl-3 py-2 text-sm bg-transparent border-none focus:ring-0 text-slate-800"/>
                        <button onClick={() => loadFromUrl(customUrl)} disabled={loading} className="bg-[#CE2029] text-white p-2 rounded-lg transition-colors"><LinkIcon size={18}/></button>
                    </div>
                </div>
                {error && <div className="mt-4 p-3 bg-red-50 text-[#CE2029] rounded-lg text-xs border border-red-100">{error}</div>}
              </div>
            </div>
          </div>
        )}

        {rawBookingData.length > 0 && (
          <div className="space-y-6">
            {activeTab === 'datasheet' && <DataSheet data={enrichedBookingData} onReupload={triggerReupload} />}
            {activeTab === 'dashboard' && <DashboardAnalysis data={enrichedBookingData} masterData={masterData} onReupload={triggerReupload} />}
            {activeTab === 'rsbt' && <RSBTCellReport data={enrichedBookingData} masterData={masterData} onReupload={triggerReupload} />}
            {activeTab === 'booking' && <BookingReport data={enrichedBookingData} onReupload={triggerReupload} />}
          </div>
        )}
      </main>
    </div>
  );
}

const NavTab = ({ active, onClick, label, icon }: any) => (
  <button onClick={onClick} className={`flex items-center gap-2 px-3 py-1.5 rounded-md text-xs font-bold transition-all ${
    active ? 'bg-white text-[#CE2029] shadow-sm' : 'text-white/80 hover:text-white hover:bg-white/10'
  }`}>
    {icon} <span className="hidden sm:inline">{label}</span>
  </button>
);

export default App;