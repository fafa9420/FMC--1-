import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Upload, DollarSign, Package, Target, TrendingUp, Search, ChevronDown, ChevronUp, CheckCircle, AlertTriangle, AlertCircle, Info, LogOut } from 'lucide-react';
import { ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer, ScatterChart, Scatter, Cell } from 'recharts';
import * as XLSX from 'xlsx';
import { initializeApp } from "firebase/app";
import { getFirestore, collection, addDoc, getDocs, query, orderBy, limit } from "firebase/firestore";
import { getAuth, signInWithEmailAndPassword, onAuthStateChanged, signOut } from "firebase/auth";

// 替換為您的專屬 Firebase 金鑰
const firebaseConfig = {
  apiKey: "AIzaSyAihDw2kc6XPwS8Udn7FeTp41GkinDZMVs",
  authDomain: "fmc-1-7f868.firebaseapp.com",
  projectId: "fmc-1-7f868",
  storageBucket: "fmc-1-7f868.firebasestorage.app",
  messagingSenderId: "276341569449",
  appId: "1:276341569449:web:38986f02ef079e67b0f6b8",
  measurementId: "G-K24K5ZTF46"
};

// 初始化 Firebase
const app = initializeApp(firebaseConfig);
const db = getFirestore(app);
const auth = getAuth(app);

// 工具函式
const formatCurrency = (val) => new Intl.NumberFormat('zh-TW').format(Math.round(val));
const formatPercent = (val) => `${(val * 100).toFixed(1)}%`;
const parseNumber = (val) => {
  if (!val) return 0;
  if (typeof val === 'number') return val;
  const cleaned = val.toString().replace(/[\r\n\s\$,]/g, '').trim();
  if (cleaned === '-' || cleaned === '') return 0;
  const num = parseFloat(cleaned);
  return isNaN(num) ? 0 : num;
};

export default function FamilyMartDashboard() {
  const [user, setUser] = useState(null);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  
  const [rawData, setRawData] = useState([]);
  const [historyList, setHistoryList] = useState([]);
  const [currentHistoryId, setCurrentHistoryId] = useState('');
  const [uploadStatus, setUploadStatus] = useState({ type: 'idle', msg: '' });
  const [filters, setFilters] = useState({ keyword: '', dateRange: 'all', category: 'all', deliveryType: 'all' });
  const [sortConfig, setSortConfig] = useState({ key: 'fulfillRate', direction: 'desc' });
  const [highlightedItem, setHighlightedItem] = useState(null);
  
  const detailRef = useRef(null);

  // 監聽登入狀態與載入歷史紀錄
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
      if (currentUser) {
        loadHistoryList();
      }
    });
    return () => unsubscribe();
  }, []);

  const loadHistoryList = async () => {
    try {
      const q = query(collection(db, "uploadHistory"), orderBy("timestamp", "desc"), limit(5));
      const querySnapshot = await getDocs(q);
      const list = [];
      querySnapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() });
      });
      setHistoryList(list);
      if (list.length > 0) {
        setCurrentHistoryId(list[0].id);
        setRawData(JSON.parse(list[0].dataStr));
      }
    } catch (error) {
      console.error("載入歷史紀錄失敗", error);
    }
  };

  const handleLogin = async (e) => {
    e.preventDefault();
    setLoginError('');
    try {
      await signInWithEmailAndPassword(auth, email, password);
    } catch (error) {
      setLoginError('登入失敗，請檢查帳號密碼。');
    }
  };

  const handleLogout = async () => {
    await signOut(auth);
    setRawData([]);
    setHistoryList([]);
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setUploadStatus({ type: 'loading', msg: '解析與上傳中...' });

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        let allData = [];

        wb.SheetNames.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
          
          let headerRowIdx = -1;
          for (let i = 0; i < Math.min(15, jsonData.length); i++) {
            const rowStr = jsonData[i].join('').toLowerCase();
            if (rowStr.includes('代號') || rowStr.includes('商品')) {
              headerRowIdx = i;
              break;
            }
          }
          if (headerRowIdx === -1) return;

          const headers = jsonData[headerRowIdx].map(h => h.toString().replace(/[\r\n\s]/g, ''));
          const findCol = (keywords, excludes = []) => headers.findIndex(h => {
             const matchKeyword = keywords.some(k => h.includes(k));
             const matchExclude = excludes.some(e => h.includes(e));
             return matchKeyword && !matchExclude;
          });

          const cols = {
            date: findCol(['檔期']),
            itemId: findCol(['代號', '品號']),
            itemName: findCol(['名稱', '品名']),
            category: findCol(['採購別', '分類']),
            deliveryType: findCol(['交期']),
            unitPrice: findCol(['單價', '銷售單價']),
            orderQty: findCol(['下單數', '預購數'], ['基本', '金額']),
            orderAmt: findCol(['下單金額', '預購金額']),
            stockQty: findCol(['備貨數'], ['金額', '率']),
          };

          for (let i = headerRowIdx + 1; i < jsonData.length; i++) {
            const row = jsonData[i];
            const itemNameStr = (row[cols.itemName] || '').toString();
            const categoryStr = (row[cols.category] || '').toString();
            
            // 排除無效商品
            if (itemNameStr.includes('未知商品') || itemNameStr.includes('無資料') || 
                categoryStr.includes('未分類') || categoryStr.includes('無資料') || !row[cols.itemId]) {
              continue;
            }

            const orderQty = parseNumber(row[cols.orderQty]);
            const stockQty = parseNumber(row[cols.stockQty]);
            const isNoLimit = (row[cols.deliveryType] || '').toString().toLowerCase().replace(/\s/g, '').includes('d+');
            
            let fulfillRate = 0;
            if (!isNoLimit && stockQty > 0) {
              fulfillRate = Math.min(orderQty / stockQty, 1); // 100% 防爆
            }

            allData.push({
              date: row[cols.date] || sheetName,
              itemId: row[cols.itemId],
              itemName: itemNameStr,
              category: categoryStr,
              deliveryType: row[cols.deliveryType] || '一般',
              unitPrice: parseNumber(row[cols.unitPrice]),
              orderQty: orderQty,
              orderAmt: parseNumber(row[cols.orderAmt]),
              stockQty: stockQty,
              fulfillRate: fulfillRate,
              isNoLimit: isNoLimit
            });
          }
        });

        const dataStr = JSON.stringify(allData);
        await addDoc(collection(db, "uploadHistory"), {
          timestamp: new Date().getTime(),
          fileName: file.name,
          dataStr: dataStr
        });

        await loadHistoryList();
        setUploadStatus({ type: 'success', msg: '上傳並同步至雲端成功！' });
        setTimeout(() => setUploadStatus({ type: 'idle', msg: '' }), 3000);
      } catch (err) {
        setUploadStatus({ type: 'error', msg: '解析失敗，請確認檔案格式。' });
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleHistoryChange = (e) => {
    const id = e.target.value;
    setCurrentHistoryId(id);
    const selected = historyList.find(h => h.id === id);
    if (selected) {
      setRawData(JSON.parse(selected.dataStr));
    }
  };

  const handleSort = (key) => {
    setSortConfig((prev) => ({
      key,
      direction: prev.key === key && prev.direction === 'desc' ? 'asc' : 'desc'
    }));
  };

  const filteredData = useMemo(() => {
    return rawData.filter(item => {
      const matchKeyword = item.itemName.includes(filters.keyword) || item.itemId.toString().includes(filters.keyword);
      const matchDate = filters.dateRange === 'all' || item.date === filters.dateRange;
      const matchCategory = filters.category === 'all' || item.category === filters.category;
      const matchDelivery = filters.deliveryType === 'all' || 
                           (filters.deliveryType === 'no_limit' ? item.isNoLimit : !item.isNoLimit);
      return matchKeyword && matchDate && matchCategory && matchDelivery;
    });
  }, [rawData, filters]);

  const sortedData = useMemo(() => {
    const sortableItems = [...filteredData];
    sortableItems.sort((a, b) => {
      // 永遠將無備貨限制 (d+2/d+3) 沉到底部
      if (a.isNoLimit && !b.isNoLimit) return 1;
      if (!a.isNoLimit && b.isNoLimit) return -1;
      
      const aVal = a[sortConfig.key];
      const bVal = b[sortConfig.key];
      
      if (aVal < bVal) return sortConfig.direction === 'asc' ? -1 : 1;
      if (aVal > bVal) return sortConfig.direction === 'asc' ? 1 : -1;
      return 0;
    });
    return sortableItems;
  }, [filteredData, sortConfig]);

  const stats = useMemo(() => {
    let totalOrder = 0, totalStockAmt = 0;
    let limitItems = 0, sumFulfill = 0, hotItems = 0;
    
    filteredData.forEach(item => {
      totalOrder += item.orderAmt;
      if (!item.isNoLimit) {
        totalStockAmt += (item.stockQty * item.unitPrice);
        limitItems++;
        sumFulfill += item.fulfillRate;
        if (item.fulfillRate >= 0.8) hotItems++;
      }
    });

    return {
      totalOrder,
      totalStockAmt,
      avgFulfill: limitItems ? (sumFulfill / limitItems) : 0,
      hotItems
    };
  }, [filteredData]);

  const trendData = useMemo(() => {
    const map = new Map();
    filteredData.forEach(item => {
      if (!map.has(item.date)) map.set(item.date, { date: item.date, orderAmt: 0, sumFulfill: 0, count: 0 });
      const d = map.get(item.date);
      d.orderAmt += item.orderAmt;
      if (!item.isNoLimit) {
        d.sumFulfill += item.fulfillRate;
        d.count++;
      }
    });
    return Array.from(map.values()).map(d => ({
      date: d.date,
      orderAmt: d.orderAmt,
      fulfillRate: d.count ? (d.sumFulfill / d.count) : 0
    })).sort((a, b) => a.date.localeCompare(b.date));
  }, [filteredData]);

  const scatterData = useMemo(() => filteredData.filter(d => !d.isNoLimit && d.orderAmt > 0), [filteredData]);

  const filterOptions = useMemo(() => ({
    dates: [...new Set(rawData.map(d => d.date))].sort(),
    categories: [...new Set(rawData.map(d => d.category))].filter(Boolean)
  }), [rawData]);

  if (!user) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4">
        <div className="bg-white p-8 rounded-xl shadow-lg max-w-md w-full">
          <div className="text-center mb-8">
            <Package className="h-12 w-12 text-blue-600 mx-auto mb-4" />
            <h1 className="text-2xl font-bold text-gray-800">全家團購銷售分析系統</h1>
            <p className="text-gray-500 mt-2">請登入以檢視企業機密數據</p>
          </div>
          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Email</label>
              <input type="email" value={email} onChange={e => setEmail(e.target.value)} className="w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" required />
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">密碼</label>
              <input type="password" value={password} onChange={e => setPassword(e.target.value)} className="w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" required />
            </div>
            {loginError && <p className="text-red-500 text-sm">{loginError}</p>}
            <button type="submit" className="w-full bg-blue-600 text-white py-2 rounded-lg hover:bg-blue-700 transition font-medium">登入系統</button>
          </form>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 p-6 font-sans">
      <div className="max-w-7xl mx-auto space-y-6">
        
        {/* Header 區塊 */}
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white p-6 rounded-xl shadow-sm border border-gray-100">
          <div>
            <h1 className="text-2xl font-bold text-gray-800 flex items-center gap-2">
              <Package className="h-6 w-6 text-blue-600" />
              全家團購商品銷售儀表板
            </h1>
            <p className="text-sm text-gray-500 mt-1">雲端安全連線中 | 帳號: {user.email}</p>
          </div>
          
          <div className="flex items-center gap-4 mt-4 md:mt-0">
            {historyList.length > 0 && (
              <select value={currentHistoryId} onChange={handleHistoryChange} className="px-4 py-2 bg-gray-50 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                {historyList.map(h => (
                  <option key={h.id} value={h.id}>{h.fileName} ({new Date(h.timestamp).toLocaleDateString()})</option>
                ))}
              </select>
            )}
            
            <label className="cursor-pointer bg-green-50 hover:bg-green-100 text-green-700 px-4 py-2 rounded-lg font-medium transition flex items-center gap-2 border border-green-200">
              <Upload className="h-4 w-4" />
              上傳新報表
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFileUpload} />
            </label>
            
            <button onClick={handleLogout} className="p-2 text-gray-400 hover:text-red-600 transition" title="登出">
              <LogOut className="h-5 w-5" />
            </button>
          </div>
        </div>

        {/* Filters */}
        <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="relative">
            <Search className="absolute left-3 top-2.5 h-4 w-4 text-gray-400" />
            <input type="text" placeholder="搜尋商品代號或名稱..." value={filters.keyword} onChange={(e) => setFilters(f => ({ ...f, keyword: e.target.value }))} className="w-full pl-9 pr-4 py-2 bg-gray-50 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <select value={filters.dateRange} onChange={(e) => setFilters(f => ({ ...f, dateRange: e.target.value }))} className="px-4 py-2 bg-gray-50 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
            <option value="all">所有檔期</option>
            {filterOptions.dates.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
          <select value={filters.category} onChange={(e) => setFilters(f => ({ ...f, category: e.target.value }))} className="px-4 py-2 bg-gray-50 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
            <option value="all">所有採購別</option>
            {filterOptions.categories.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
          <select value={filters.deliveryType} onChange={(e) => setFilters(f => ({ ...f, deliveryType: e.target.value }))} className="px-4 py-2 bg-gray-50 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
            <option value="all">所有交期</option>
            <option value="limit">有限量商品 (可算完銷率)</option>
            <option value="no_limit">無限制商品 (d+2/d+3)</option>
          </select>
        </div>

        {/* KPI Cards */}
        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="group relative bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <div className="absolute opacity-0 group-hover:opacity-100 transition bottom-full left-1/2 -translate-x-1/2 mb-2 w-48 bg-gray-800 text-white text-xs p-2 rounded pointer-events-none z-10 text-center">
              計算邏輯：當前篩選條件下所有商品的下單金額加總。
            </div>
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-blue-50 text-blue-600 rounded-lg"><DollarSign className="h-5 w-5" /></div>
              <h3 className="text-gray-500 font-medium">總下單金額</h3>
            </div>
            <p className="text-2xl font-bold text-gray-800">{formatCurrency(stats.totalOrder)}</p>
          </div>
          <div className="group relative bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <div className="absolute opacity-0 group-hover:opacity-100 transition bottom-full left-1/2 -translate-x-1/2 mb-2 w-48 bg-gray-800 text-white text-xs p-2 rounded pointer-events-none z-10 text-center">
              計算邏輯：排除 d+2/d+3，備貨數量 × 單價。
            </div>
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-purple-50 text-purple-600 rounded-lg"><Package className="h-5 w-5" /></div>
              <h3 className="text-gray-500 font-medium">總備貨金額</h3>
            </div>
            <p className="text-2xl font-bold text-gray-800">{formatCurrency(stats.totalStockAmt)}</p>
          </div>
          <div className="group relative bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <div className="absolute opacity-0 group-hover:opacity-100 transition bottom-full left-1/2 -translate-x-1/2 mb-2 w-48 bg-gray-800 text-white text-xs p-2 rounded pointer-events-none z-10 text-center">
              計算邏輯：排除 d+2/d+3，各項商品完銷率的平均值(上限100%)。
            </div>
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-orange-50 text-orange-600 rounded-lg"><Target className="h-5 w-5" /></div>
              <h3 className="text-gray-500 font-medium">平均備貨完銷率</h3>
            </div>
            <p className="text-2xl font-bold text-gray-800">{formatPercent(stats.avgFulfill)}</p>
          </div>
          <div className="group relative bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <div className="absolute opacity-0 group-hover:opacity-100 transition bottom-full left-1/2 -translate-x-1/2 mb-2 w-48 bg-gray-800 text-white text-xs p-2 rounded pointer-events-none z-10 text-center">
              計算邏輯：完銷率 &ge; 80% 的商品數量。
            </div>
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-green-50 text-green-600 rounded-lg"><TrendingUp className="h-5 w-5" /></div>
              <h3 className="text-gray-500 font-medium">熱門商品數 (完銷&ge;80%)</h3>
            </div>
            <p className="text-2xl font-bold text-gray-800">{stats.hotItems} <span className="text-sm font-normal text-gray-500">項</span></p>
          </div>
        </div>

        {/* Charts Container - 上下直式排列 */}
        <div className="space-y-6">
          {/* Chart 1: 檔期趨勢 */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <h3 className="text-lg font-bold text-gray-800 mb-6 flex items-center gap-2">各檔期銷售趨勢</h3>
            <div className="h-[300px] w-full">
              <ResponsiveContainer>
                <ComposedChart data={trendData}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#E5E7EB" />
                  <XAxis dataKey="date" tick={{ fill: '#6B7280', fontSize: 12 }} />
                  <YAxis yAxisId="left" tickFormatter={val => `$${val/1000}k`} tick={{ fill: '#6B7280', fontSize: 12 }} />
                  <YAxis yAxisId="right" orientation="right" tickFormatter={formatPercent} tick={{ fill: '#6B7280', fontSize: 12 }} />
                  <RechartsTooltip 
                    formatter={(value, name) => {
                      if (name === '總下單金額') return [formatCurrency(value), name];
                      if (name === '平均完銷率') return [formatPercent(value), name];
                      return [value, name];
                    }}
                  />
                  <Legend />
                  <Bar yAxisId="left" dataKey="orderAmt" name="總下單金額" fill="#3B82F6" radius={[4, 4, 0, 0]} maxBarSize={50} />
                  <Line yAxisId="right" type="monotone" dataKey="fulfillRate" name="平均完銷率" stroke="#F97316" strokeWidth={3} dot={{ r: 4 }} activeDot={{ r: 6 }} />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>

          {/* Chart 2: 商品落點分析 */}
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-lg font-bold text-gray-800 flex items-center gap-2">
                商品落點分析
                <span className="text-sm font-normal text-gray-400 bg-gray-50 px-2 py-1 rounded">排除 d+2/d+3</span>
              </h3>
            </div>
            <div className="h-[350px] w-full">
              <ResponsiveContainer>
                <ScatterChart margin={{ top: 20, right: 20, bottom: 20, left: 20 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E5E7EB" />
                  <XAxis type="number" dataKey="fulfillRate" name="完銷率" tickFormatter={formatPercent} domain={[0, 1]} label={{ value: '完銷率 (%)', position: 'bottom', offset: 0 }} />
                  <YAxis type="number" dataKey="orderAmt" name="下單金額" tickFormatter={val => `$${val/1000}k`} label={{ value: '下單金額 (NT$)', angle: -90, position: 'insideLeft' }} />
                  <RechartsTooltip 
                    cursor={{ strokeDasharray: '3 3' }}
                    content={({ payload }) => {
                      if (payload && payload.length) {
                        const d = payload[0].payload;
                        return (
                          <div className="bg-white p-3 border border-gray-200 shadow-lg rounded-lg text-sm">
                            <p className="font-bold text-gray-800 border-b pb-1 mb-1">{d.itemId} - {d.itemName}</p>
                            <p className="text-blue-600">下單金額: {formatCurrency(d.orderAmt)}</p>
                            <p className="text-orange-600">完銷率: {formatPercent(d.fulfillRate)}</p>
                            <p className="text-gray-500 text-xs mt-2 italic">點擊查看明細</p>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  <Scatter name="商品" data={scatterData} 
                    onClick={(d) => {
                      setHighlightedItem(d.itemId);
                      detailRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
                    }}>
                    {scatterData.map((entry, index) => (
                      <Cell key={`cell-${index}`} fill={entry.fulfillRate >= 0.8 ? '#F97316' : '#3B82F6'} className="cursor-pointer hover:opacity-80" />
                    ))}
                  </Scatter>
                </ScatterChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        {/* Detail Table */}
        <div ref={detailRef} className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
          <div className="p-6 border-b border-gray-100 flex justify-between items-center">
            <h3 className="text-lg font-bold text-gray-800">商品銷售明細清單</h3>
            <span className="text-sm text-gray-500">共 {sortedData.length} 筆資料</span>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-sm text-left">
              <thead className="text-xs text-gray-600 bg-gray-50 uppercase sticky top-0">
                <tr>
                  {[
                    { key: 'date', label: '檔期' },
                    { key: 'itemId', label: '商品代號' },
                    { key: 'itemName', label: '商品名稱' },
                    { key: 'category', label: '採購別' },
                    { key: 'deliveryType', label: '交期類型' },
                    { key: 'unitPrice', label: '銷售單價' },
                    { key: 'orderQty', label: '下單數量' },
                    { key: 'orderAmt', label: '下單金額' },
                    { key: 'stockQty', label: '備貨數量' },
                    { key: 'fulfillRate', label: '完銷率' }
                  ].map(col => (
                    <th key={col.key} 
                        className="px-4 py-3 cursor-pointer hover:bg-gray-100 group whitespace-nowrap"
                        onClick={() => handleSort(col.key)}>
                      <div className="flex items-center gap-1">
                        {col.label}
                        <span className="text-gray-400 group-hover:text-gray-600">
                          {sortConfig.key === col.key ? 
                            (sortConfig.direction === 'asc' ? <ChevronUp className="h-4 w-4 text-green-600" /> : <ChevronDown className="h-4 w-4 text-green-600" />) 
                            : <ChevronDown className="h-4 w-4 opacity-0 group-hover:opacity-50" />}
                        </span>
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sortedData.map((item, idx) => (
                  <tr key={`${item.itemId}-${idx}`} 
                      className={`border-b hover:bg-blue-50 transition ${highlightedItem === item.itemId ? 'bg-green-50 border-2 border-green-500' : ''}`}>
                    <td className="px-4 py-3 whitespace-nowrap">{item.date}</td>
                    <td className="px-4 py-3 font-medium text-gray-900">{item.itemId}</td>
                    <td className="px-4 py-3 max-w-[200px] truncate" title={item.itemName}>{item.itemName}</td>
                    <td className="px-4 py-3">
                      <span className="px-2 py-1 bg-gray-100 text-gray-600 rounded-full text-xs">{item.category}</span>
                    </td>
                    <td className="px-4 py-3 whitespace-nowrap">{item.deliveryType}</td>
                    <td className="px-4 py-3 text-right">{formatCurrency(item.unitPrice)}</td>
                    <td className="px-4 py-3 text-right">{formatCurrency(item.orderQty)}</td>
                    <td className="px-4 py-3 text-right font-medium text-blue-600">{formatCurrency(item.orderAmt)}</td>
                    <td className="px-4 py-3 text-right text-gray-500">
                      {item.isNoLimit ? '-' : formatCurrency(item.stockQty)}
                    </td>
                    <td className="px-4 py-3 text-right">
                      {item.isNoLimit ? (
                        <span className="text-gray-400 text-xs">無限制</span>
                      ) : (
                        <div className="flex items-center justify-end gap-2">
                          <span className={item.fulfillRate >= 0.8 ? 'text-green-600 font-bold' : item.fulfillRate < 0.3 ? 'text-red-500' : 'text-gray-700'}>
                            {formatPercent(item.fulfillRate)}
                          </span>
                          {item.fulfillRate >= 0.8 && <CheckCircle className="h-4 w-4 text-green-500" />}
                          {item.fulfillRate < 0.3 && <AlertTriangle className="h-4 w-4 text-red-400" />}
                        </div>
                      )}
                    </td>
                  </tr>
                ))}
                {sortedData.length === 0 && (
                  <tr><td colSpan="10" className="px-4 py-8 text-center text-gray-500">目前沒有符合條件的資料</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

      </div>
    </div>
  );
}
