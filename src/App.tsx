/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect, useMemo, useRef } from 'react';
import { motion, AnimatePresence } from 'motion/react';
import { 
  Users, 
  UserPlus, 
  Trash2, 
  Shuffle, 
  Settings2, 
  Copy, 
  Download, 
  RotateCcw,
  Plus,
  CheckCircle2,
  AlertCircle,
  FileUp,
  FileText,
  History,
  Clock,
  X
} from 'lucide-react';
import confetti from 'canvas-confetti';
import { read, utils, writeFile } from 'xlsx';
import { 
  Cloud,
  ExternalLink,
  LogOut
} from 'lucide-react';

// --- Types ---
enum GroupMode {
  COUNT = 'count', // Fixed number of groups
  SIZE = 'size'    // Fixed number of people per group
}

interface GroupResult {
  id: number;
  members: string[];
}

interface HistoryRecord {
  id: string;
  timestamp: string;
  mode: GroupMode;
  targetValue: number;
  namesList: string[];
  results: GroupResult[];
  identity: string;
}

// --- Constants ---
const SAMPLE_NAMES = [
  '趙敏', '周星馳', '梁朝偉', '張國榮', '梅艷芳', '舒淇'
].join('\n');

const IDENTITY_OPTIONS = ['管理者', '老師', '主持人', '活動組長', '自訂'];

export default function App() {
  // --- States ---
  const [namesInput, setNamesInput] = useState('');
  const [namesList, setNamesList] = useState<string[]>([]);
  const [mode, setMode] = useState<GroupMode>(GroupMode.COUNT);
  const [targetValue, setTargetValue] = useState<number>(4);
  const [results, setResults] = useState<GroupResult[]>([]);
  const [isShuffling, setIsShuffling] = useState(false);
  const [message, setMessage] = useState<{ text: string, type: 'info' | 'error' | 'success' } | null>(null);
  const [isGoogleLinked, setIsGoogleLinked] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [history, setHistory] = useState<HistoryRecord[]>([]);
  const [showHistory, setShowHistory] = useState(false);
  const [userIdentity, setUserIdentity] = useState('管理者');
  const [customIdentity, setCustomIdentity] = useState('');
  const fileInputRef = useRef<HTMLInputElement>(null);

  // --- Persistence ---
  useEffect(() => {
    const savedNames = localStorage.getItem('grouping-tool-names');
    if (savedNames) {
      setNamesInput(savedNames);
      processNames(savedNames);
    }
    
    const savedHistory = localStorage.getItem('grouping-tool-history');
    if (savedHistory) {
      setHistory(JSON.parse(savedHistory));
    }

    checkGoogleAuth();
  }, []);

  useEffect(() => {
    localStorage.setItem('grouping-tool-history', JSON.stringify(history));
  }, [history]);

  useEffect(() => {
    const handleMessage = (event: MessageEvent) => {
      if (event.data?.type === 'OAUTH_AUTH_SUCCESS') {
        setIsGoogleLinked(true);
        showMessage('Google 帳號連動成功！', 'success');
      }
    };
    window.addEventListener('message', handleMessage);
    return () => window.removeEventListener('message', handleMessage);
  }, []);

  const checkGoogleAuth = async () => {
    try {
      const resp = await fetch('/api/auth/status');
      const data = await resp.json();
      setIsGoogleLinked(data.isAuthenticated);
    } catch (e) {
      console.error('Failed to check auth status');
    }
  };

  const handleGoogleLink = async () => {
    try {
      const resp = await fetch('/api/auth/google/url');
      const { url } = await resp.json();
      window.open(url, 'google_auth', 'width=600,height=700');
    } catch (e) {
      showMessage('無法取得授權連結', 'error');
    }
  };

  const handleGoogleLogout = async () => {
    try {
      await fetch('/api/auth/logout', { method: 'POST' });
      setIsGoogleLinked(false);
      showMessage('已解除連動', 'info');
    } catch (e) {
      showMessage('登出失敗', 'error');
    }
  };

  const exportToGoogleSheets = async () => {
    if (!isGoogleLinked) {
      handleGoogleLink();
      return;
    }

    setIsExporting(true);
    try {
      const resp = await fetch('/api/export/google-sheets', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ results })
      });
      const data = await resp.json();
      if (data.success) {
        showMessage('已成功建立 Google 試算表！', 'success');
        window.open(data.spreadsheetUrl, '_blank');
      } else {
        throw new Error(data.error);
      }
    } catch (e) {
      showMessage('匯出至 Google 試算表失敗', 'error');
    } finally {
      setIsExporting(false);
    }
  };

  useEffect(() => {
    localStorage.setItem('grouping-tool-names', namesInput);
  }, [namesInput]);

  // --- Handlers ---
  const processNames = (input: string) => {
    const list = input
      .split(/[\n,，\s\t]+/)
      .map(n => n.trim())
      .filter(n => n.length > 0);
    setNamesList(list);
  };

  const handleInputChange = (e: React.ChangeEvent<HTMLTextAreaElement>) => {
    setNamesInput(e.target.value);
    processNames(e.target.value);
  };

  const importSample = () => {
    setNamesInput(SAMPLE_NAMES);
    processNames(SAMPLE_NAMES);
    showMessage('已匯入範本名單', 'success');
  };

  const showMessage = (text: string, type: 'info' | 'error' | 'success' = 'info') => {
    setMessage({ text, type });
    setTimeout(() => setMessage(null), 3000);
  };

  const handleFileSelect = () => {
    fileInputRef.current?.click();
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const fileName = file.name.toLowerCase();
      
      if (fileName.endsWith('.txt') || fileName.endsWith('.csv')) {
        const text = await file.text();
        // Simple heuristic to split CSV/TXT
        const names = text.split(/[\n\r,，\s\t]+/).map(n => n.trim()).filter(n => n.length > 0);
        const newInput = names.join('\n');
        setNamesInput(newInput);
        processNames(newInput);
        showMessage(`已成功匯入 ${names.length} 位成員`, 'success');
      } 
      else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        const data = await file.arrayBuffer();
        const workbook = read(data);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        
        // Flatten and extract non-empty strings
        const names = json.flat().map(v => String(v).trim()).filter(v => v.length > 0 && v !== 'undefined' && v !== 'null');
        const newInput = names.join('\n');
        setNamesInput(newInput);
        processNames(newInput);
        showMessage(`已成功匯入 Excel 成員 ${names.length} 位`, 'success');
      } else {
        showMessage('不支援此檔案格式', 'error');
      }
    } catch (err) {
      console.error(err);
      showMessage('解析檔案失敗', 'error');
    } finally {
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleShuffle = () => {
    if (namesList.length === 0) {
      showMessage('請先輸入名單或點擊範本匯入', 'error');
      return;
    }

    setIsShuffling(true);
    setResults([]);

    // Simulate thinking/shuffling animation
    setTimeout(() => {
      const shuffled = [...namesList].sort(() => Math.random() - 0.5);
      const newResults: GroupResult[] = [];
      
      let groupCount = 0;
      if (mode === GroupMode.COUNT) {
        groupCount = targetValue;
      } else {
        groupCount = Math.ceil(shuffled.length / targetValue);
      }

      if (groupCount <= 0) groupCount = 1;
      if (groupCount > shuffled.length) groupCount = shuffled.length;

      // Initialize groups
      for (let i = 0; i < groupCount; i++) {
        newResults.push({ id: i + 1, members: [] });
      }

      // Distribute
      shuffled.forEach((name, index) => {
        if (newResults.length > 0) {
          newResults[index % groupCount].members.push(name);
        }
      });

      setResults(newResults);
      setIsShuffling(false);
      
      // Add to history
      const finalIdentity = userIdentity === '自訂' ? (customIdentity || '神秘訪客') : userIdentity;
      const newRecord: HistoryRecord = {
        id: crypto.randomUUID(),
        timestamp: new Date().toLocaleString('zh-TW'),
        mode,
        targetValue,
        namesList: [...namesList],
        results: newResults,
        identity: finalIdentity
      };
      setHistory(prev => [newRecord, ...prev].slice(0, 50)); 

      confetti({
        particleCount: 150,
        spread: 70,
        origin: { y: 0.6 },
        colors: ['#6366f1', '#8b5cf6', '#a855f7']
      });
      showMessage('分組成功！', 'success');
    }, 800);
  };

  const copyToClipboard = () => {
    const text = results.map(g => `第 ${g.id} 組 (${g.members.length}人)：${g.members.join(', ')}`).join('\n');
    navigator.clipboard.writeText(text).then(() => {
      showMessage('已複製到剪貼簿', 'success');
    });
  };

  const copyForSpreadsheet = () => {
    // Tab-separated format (TSV) is the best for pasting into Excel/Google Sheets
    const header = "組別\t人數\t成員名單\n";
    const rows = results.map(g => `第 ${g.id} 組\t${g.members.length}\t${g.members.join(', ')}`).join('\n');
    navigator.clipboard.writeText(header + rows).then(() => {
      showMessage('已複製為表格格式，請至 Google 試算表直接貼上', 'success');
    });
  };

  const exportToExcel = () => {
    try {
      const data = results.map(g => ({
        '組別': `第 ${g.id} 組`,
        '人數': g.members.length,
        '成員名單': g.members.join(', ')
      }));
      const ws = utils.json_to_sheet(data);
      const wb = utils.book_new();
      utils.book_append_sheet(wb, ws, "分組結果");
      writeFile(wb, `隨機分組結果_${new Date().toISOString().split('T')[0]}.xlsx`);
      showMessage('Excel 檔案已下載', 'success');
    } catch (err) {
      showMessage('匯出失敗', 'error');
    }
  };

  const clearAll = () => {
    if (window.confirm('確定要清空所有名單嗎？')) {
      setNamesInput('');
      setNamesList([]);
      setResults([]);
    }
  };

  const loadHistory = (record: HistoryRecord) => {
    setNamesInput(record.namesList.join('\n'));
    setNamesList(record.namesList);
    setResults(record.results);
    setMode(record.mode);
    setTargetValue(record.targetValue);
    setShowHistory(false);
    showMessage('已載入歷程資料', 'success');
  };

  const deleteHistoryItem = (id: string, e: React.MouseEvent) => {
    e.stopPropagation();
    setHistory(prev => prev.filter(h => h.id !== id));
  };

  const clearHistory = () => {
    if (window.confirm('確定要永久刪除所有分組紀錄嗎？')) {
      setHistory([]);
    }
  };

  // --- Render ---
  return (
    <div className="min-h-screen p-6 md:p-12 lg:p-16 max-w-7xl mx-auto">
      {/* Hidden File Input */}
      <input 
        type="file" 
        ref={fileInputRef} 
        onChange={handleFileUpload} 
        accept=".txt,.csv,.xlsx,.xls" 
        className="hidden" 
      />

      {/* Header */}
      <header className="mb-14 flex flex-col md:flex-row md:items-center justify-between gap-8">
        <div>
          <motion.div 
            initial={{ opacity: 0, x: -20 }}
            animate={{ opacity: 1, x: 0 }}
            className="flex items-center gap-4 mb-3"
          >
            <div className="p-4 bg-[#9d8189] text-white rounded-[1.5rem] shadow-xl flex items-center justify-center">
              <Users size={32} strokeWidth={2.5} />
            </div>
            <h1 className="text-4xl font-extrabold tracking-tight text-[#9d8189]">
              分組<span className="text-[#f4acb7]">隨機</span>神器
            </h1>
          </motion.div>
          <p className="text-[#9d8189]/60 font-semibold text-lg ml-1">Professional Grouping Infrastructure.</p>
        </div>
        
        <div className="flex flex-wrap gap-3 self-center md:self-auto">
          <motion.button 
            whileHover={{ scale: 1.02 }}
            whileTap={{ scale: 0.98 }}
            onClick={() => setShowHistory(true)}
            className="btn-sample bg-white border-slate-200 text-[#9d8189]"
          >
            <History size={18} className="text-[#f4acb7]" />
            <span className="pr-1 text-base">歷史紀錄 ({history.length})</span>
          </motion.button>

          <motion.button 
            whileHover={{ scale: 1.02 }}
            whileTap={{ scale: 0.98 }}
            onClick={handleFileSelect}
            className="btn-sample"
          >
            <FileUp size={18} className="text-[#f4acb7]" />
            <span className="pr-1 text-base">匯入檔案</span>
          </motion.button>

          <motion.button 
            whileHover={{ scale: 1.02 }}
            whileTap={{ scale: 0.98 }}
            onClick={importSample}
            className="btn-sample group"
          >
            <div className="w-8 h-8 bg-[#ffe5d9] rounded-lg flex items-center justify-center text-[#f4acb7] group-hover:rotate-90 transition-transform duration-300">
              <Plus size={20} strokeWidth={3} />
            </div>
            <span className="pr-1 text-base">24人範本</span>
          </motion.button>
        </div>
      </header>

      <div className="grid grid-cols-1 lg:grid-cols-12 gap-10">
        {/* Left Column: Configuration */}
        <div className="lg:col-span-5 space-y-8">
          {/* Input Section */}
          <section className="glass-card p-8">
            <div className="flex items-center justify-between mb-6">
              <div className="flex items-center gap-3 font-bold text-[#9d8189] text-lg">
                <UserPlus size={20} className="text-[#f4acb7]" />
                <span>成員清單</span>
              </div>
              <button 
                onClick={clearAll}
                className="w-10 h-10 flex items-center justify-center rounded-xl text-[#9d8189]/40 hover:text-rose-400 hover:bg-rose-50 transition-all"
                title="清空名單"
              >
                <Trash2 size={20} />
              </button>
            </div>
            
            <textarea
              className="input-field min-h-[350px] font-mono text-base leading-relaxed resize-none"
              placeholder="輸入名字，每行一位成員..."
              value={namesInput}
              onChange={handleInputChange}
            />
            
            <div className="mt-4 flex items-center justify-between">
              <div className="flex gap-2 items-center text-[#9d8189]/50">
                <FileText size={14} />
                <span className="text-xs font-bold uppercase tracking-wider">TXT, CSV, XLSX</span>
              </div>
              <div className="text-sm font-bold">
                <span className="text-[#9d8189]/40">Total: </span>
                <span className="text-[#f4acb7] text-lg">{namesList.length}</span>
              </div>
            </div>
          </section>

          {/* Settings Section */}
          <section className="glass-card p-8">
            <div className="flex items-center gap-3 font-bold text-[#9d8189] text-lg mb-8">
              <Settings2 size={20} className="text-[#f4acb7]" />
              <span>分組邏輯</span>
            </div>
            
            <div className="space-y-8">
              <div className="flex bg-[#ffe5d9]/50 p-1.5 rounded-2xl">
                <button
                  onClick={() => setMode(GroupMode.COUNT)}
                  className={`flex-1 py-3 text-sm font-extrabold rounded-xl transition-all ${mode === GroupMode.COUNT ? 'bg-white text-[#f4acb7] shadow-sm' : 'text-[#9d8189]/60 hover:text-[#9d8189]'}`}
                >
                  固定組數
                </button>
                <button
                  onClick={() => setMode(GroupMode.SIZE)}
                  className={`flex-1 py-3 text-sm font-extrabold rounded-xl transition-all ${mode === GroupMode.SIZE ? 'bg-white text-[#f4acb7] shadow-sm' : 'text-[#9d8189]/60 hover:text-[#9d8189]'}`}
                >
                  每組上限
                </button>
              </div>

              <div className="flex items-center justify-between px-2">
                <label className="font-bold text-[#9d8189]/80">
                  {mode === GroupMode.COUNT ? '需要分成幾組？' : '每組理想人數？'}
                </label>
                <div className="flex items-center gap-2">
                  <button 
                    onClick={() => setTargetValue(Math.max(1, targetValue - 1))}
                    className="w-12 h-12 flex items-center justify-center rounded-2xl bg-white/60 hover:bg-white text-[#9d8189] transition-all active:scale-90"
                  >
                    -
                  </button>
                  <input 
                    type="number" 
                    className="w-16 text-center font-black text-2xl bg-transparent border-none focus:ring-0 text-[#9d8189]" 
                    value={targetValue}
                    onChange={(e) => setTargetValue(parseInt(e.target.value) || 1)}
                  />
                  <button 
                    onClick={() => setTargetValue(targetValue + 1)}
                    className="w-12 h-12 flex items-center justify-center rounded-2xl bg-white/60 hover:bg-white text-[#9d8189] transition-all active:scale-90"
                  >
                    +
                  </button>
                </div>
              </div>

              {/* Identity Selection */}
              <div className="space-y-3">
                <label className="font-bold text-[#9d8189]/80 text-sm block">操作者身份：</label>
                <div className="flex flex-wrap gap-2">
                  {IDENTITY_OPTIONS.map(option => (
                    <button
                      key={option}
                      onClick={() => setUserIdentity(option)}
                      className={`px-3 py-1.5 rounded-lg text-xs font-bold transition-all ${
                        userIdentity === option 
                        ? 'bg-[#9d8189] text-white shadow-md' 
                        : 'bg-white border border-slate-100 text-[#9d8189]'
                      }`}
                    >
                      {option}
                    </button>
                  ))}
                </div>
                {userIdentity === '自訂' && (
                  <motion.input
                    initial={{ opacity: 0, y: -10 }}
                    animate={{ opacity: 1, y: 0 }}
                    type="text"
                    placeholder="請輸入您的自訂身份..."
                    className="input-field py-2 text-sm"
                    value={customIdentity}
                    onChange={(e) => setCustomIdentity(e.target.value)}
                  />
                )}
              </div>

              <button 
                onClick={handleShuffle}
                disabled={isShuffling || namesList.length === 0}
                className="btn-primary w-full text-xl py-5"
              >
                {isShuffling ? (
                  <RotateCcw className="animate-spin" size={24} />
                ) : (
                  <Shuffle size={24} strokeWidth={2.5} />
                )}
                {isShuffling ? '運算中...' : '生成隨機分組'}
              </button>
            </div>
          </section>
        </div>

        {/* Right Column: Display */}
        <div className="lg:col-span-7">
          <AnimatePresence mode="wait">
            {results.length > 0 ? (
              <motion.div
                key="results"
                initial={{ opacity: 0, scale: 0.98 }}
                animate={{ opacity: 1, scale: 1 }}
                className="space-y-8"
              >
                <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-6">
                  <h2 className="text-3xl font-black text-[#9d8189] flex items-center gap-3">
                    <span className="w-10 h-10 rounded-2xl bg-[#ffcad4] text-white flex items-center justify-center shadow-lg shadow-[#ffcad4]/20">
                      <CheckCircle2 size={24} />
                    </span>
                    分組完成
                  </h2>
                  <div className="flex flex-wrap items-center gap-3">
                    {isGoogleLinked ? (
                      <button 
                        onClick={exportToGoogleSheets} 
                        disabled={isExporting}
                        className="btn-secondary py-2.5 px-4 text-sm bg-emerald-50 border-emerald-100 text-emerald-700 hover:bg-emerald-100"
                      >
                        {isExporting ? (
                          <RotateCcw className="animate-spin" size={16} />
                        ) : (
                          <Cloud size={16} />
                        )}
                        存至 Google 試算表
                      </button>
                    ) : (
                      <button 
                        onClick={handleGoogleLink}
                        className="btn-secondary py-2.5 px-4 text-sm bg-indigo-50 border-indigo-100 text-indigo-700 hover:bg-indigo-100"
                      >
                        <ExternalLink size={16} />
                        連動 Google 帳號
                      </button>
                    )}
                    
                    <button onClick={exportToExcel} className="btn-secondary py-2.5 px-4 text-sm">
                      <Download size={16} />
                      下載副檔名
                    </button>
                    
                    <button onClick={handleShuffle} className="btn-secondary py-2.5 px-4 text-sm bg-[#9d8189] text-white border-none hover:bg-[#8a6f77]">
                      <RotateCcw size={16} className="text-white" />
                      重新分組
                    </button>

                    {isGoogleLinked && (
                      <button 
                        onClick={handleGoogleLogout}
                        className="p-2 text-[#9d8189]/40 hover:text-rose-400 transition-colors"
                        title="解除 Google 連動"
                      >
                        <LogOut size={16} />
                      </button>
                    )}
                  </div>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  {results.map((group, idx) => (
                    <motion.div
                      key={group.id}
                      initial={{ opacity: 0, y: 20 }}
                      animate={{ opacity: 1, y: 0 }}
                      transition={{ delay: idx * 0.05, type: 'spring', stiffness: 100 }}
                      className="glass-card overflow-hidden group hover:shadow-xl hover:shadow-[#9d8189]/5 transition-all border-none ring-1 ring-white/50"
                    >
                      <div className="bg-[#9d8189] px-6 py-4 flex justify-between items-center">
                        <span className="font-extrabold text-white tracking-wide uppercase text-sm">Group {group.id}</span>
                        <div className="px-2 py-1 bg-white/20 rounded-lg backdrop-blur-sm">
                          <span className="text-xs font-black text-white">{group.members.length} Members</span>
                        </div>
                      </div>
                      <div className="p-6 flex flex-wrap gap-2.5 min-h-[100px] content-start bg-white/40">
                        {group.members.map((member, mIdx) => (
                          <motion.span 
                            key={mIdx} 
                            whileHover={{ scale: 1.05, y: -2 }}
                            className="inline-flex px-4 py-2 bg-white/80 border border-white rounded-xl text-sm font-bold text-[#9d8189] shadow-sm transition-colors hover:border-[#f4acb7] hover:text-[#f4acb7]"
                          >
                            {member}
                          </motion.span>
                        ))}
                      </div>
                    </motion.div>
                  ))}
                </div>
              </motion.div>
            ) : (
              <motion.div
                key="empty"
                initial={{ opacity: 0 }}
                animate={{ opacity: 1 }}
                className="h-full min-h-[600px] flex flex-col items-center justify-center text-center p-12 glass-card border-none ring-1 ring-white/40 bg-white/20"
              >
                <motion.div 
                  animate={{ 
                    rotate: [0, 10, -10, 0],
                    y: [0, -10, 0]
                  }}
                  transition={{ repeat: Infinity, duration: 4 }}
                  className="w-28 h-28 bg-[#ffe5d9] rounded-[2.5rem] flex items-center justify-center text-[#f4acb7] mb-8 shadow-inner"
                >
                  <Shuffle size={52} strokeWidth={1.5} />
                </motion.div>
                <h3 className="text-2xl font-black text-[#9d8189] mb-4">即刻創造公平分組</h3>
                <p className="text-[#9d8189]/60 text-lg max-w-sm leading-relaxed mb-10 font-medium">
                  只需在左側貼上成員名單，我們強大的隨機引擎將為您處理繁瑣的計算。
                </p>
                <div className="flex gap-4">
                  <div className="w-3 h-3 rounded-full bg-[#f4acb7] animate-bounce"></div>
                  <div className="w-3 h-3 rounded-full bg-[#ffcad4] animate-bounce delay-100"></div>
                  <div className="w-3 h-3 rounded-full bg-[#ffe5d9] animate-bounce delay-200"></div>
                </div>
              </motion.div>
            )}
          </AnimatePresence>
        </div>
      </div>

      {/* History Drawer */}
      <AnimatePresence>
        {showHistory && (
          <>
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowHistory(false)}
              className="fixed inset-0 bg-slate-900/40 backdrop-blur-sm z-[60]"
            />
            <motion.div 
              initial={{ x: '100%' }}
              animate={{ x: 0 }}
              exit={{ x: '100%' }}
              transition={{ type: 'spring', damping: 25, stiffness: 200 }}
              className="fixed top-0 right-0 h-full w-full max-w-md bg-white shadow-2xl z-[70] flex flex-col"
            >
              <div className="p-6 border-b border-slate-100 flex items-center justify-between bg-[#9d8189] text-white">
                <div className="flex items-center gap-3">
                  <Clock size={22} />
                  <h2 className="text-xl font-bold">分組歷程紀錄</h2>
                </div>
                <button onClick={() => setShowHistory(false)} className="p-2 hover:bg-white/20 rounded-full transition-colors">
                  <X size={24} />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto p-4 space-y-4 bg-slate-50">
                {history.length > 0 ? (
                  history.map((record) => (
                    <div 
                      key={record.id}
                      onClick={() => loadHistory(record)}
                      className="group glass-card p-5 cursor-pointer hover:border-[#f4acb7] hover:shadow-lg transition-all relative overflow-hidden"
                    >
                      <div className="flex justify-between items-start mb-3">
                        <span className="text-xs font-mono font-bold text-[#9d8189]/60">{record.timestamp}</span>
                        <button 
                          onClick={(e) => deleteHistoryItem(record.id, e)}
                          className="text-slate-300 hover:text-rose-500 transition-colors p-1"
                        >
                          <Trash2 size={16} />
                        </button>
                      </div>
                      <div className="font-bold text-slate-800 mb-1 truncate">
                        {record.namesList.length} 人 · 分成 {record.results.length} 組
                      </div>
                      <div className="text-[11px] font-bold text-[#f4acb7] mb-3 flex items-center gap-1">
                        <Users size={10} /> 執行身份：{record.identity || '未設定'}
                      </div>
                      <div className="flex flex-wrap gap-1">
                        {record.results.slice(0, 3).map(g => (
                          <span key={g.id} className="text-[10px] px-2 py-0.5 bg-slate-100 rounded text-slate-500">
                            G{g.id}: {g.members.length}人
                          </span>
                        ))}
                        {record.results.length > 3 && <span className="text-[10px] text-slate-400">...</span>}
                      </div>
                      {/* Hover Indicator */}
                      <div className="absolute right-0 top-0 bottom-0 w-1 bg-[#f4acb7] translate-x-1 group-hover:translate-x-0 transition-transform" />
                    </div>
                  ))
                ) : (
                  <div className="h-full flex flex-col items-center justify-center text-center p-12">
                    <History size={48} className="text-slate-200 mb-4" />
                    <p className="text-slate-400 font-medium">尚無歷程紀錄</p>
                  </div>
                )}
              </div>

              {history.length > 0 && (
                <div className="p-4 border-t border-slate-100">
                  <button 
                    onClick={clearHistory}
                    className="w-full py-3 text-sm font-bold text-rose-500 hover:bg-rose-50 rounded-xl transition-colors flex items-center justify-center gap-2"
                  >
                    <Trash2 size={16} />
                    清空所有紀錄
                  </button>
                </div>
              )}
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* Toast Notification */}
      <AnimatePresence>
        {message && (
          <motion.div
            initial={{ opacity: 0, y: 50, scale: 0.95 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            exit={{ opacity: 0, y: 20, scale: 0.95 }}
            className={`fixed bottom-10 left-1/2 -translate-x-1/2 px-8 py-4 rounded-2xl shadow-2xl flex items-center gap-3 z-50 border ${
              message.type === 'success' ? 'bg-[#9d8189] border-[#9d8189]/50 text-white' :
              message.type === 'error' ? 'bg-rose-500 border-rose-400 text-white' :
              'bg-slate-900 border-slate-800 text-white'
            }`}
          >
            {message.type === 'success' && <CheckCircle2 size={20} />}
            {message.type === 'error' && <AlertCircle size={20} />}
            <span className="text-base font-bold">{message.text}</span>
          </motion.div>
        )}
      </AnimatePresence>

      {/* Footer Info */}
      <footer className="mt-20 py-8 border-t border-white/20 text-center">
        <p className="text-[#9d8189]/50 text-sm font-medium tracking-wide">
          © 2026 分組神器 · 打造極致的分組體驗
        </p>
      </footer>
    </div>
  );
}
