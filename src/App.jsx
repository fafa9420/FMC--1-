import React, { useState, useEffect, useMemo, useRef } from 'react';
import { Upload, DollarSign, Package, Target, TrendingUp, Search, ChevronDown, ChevronUp, CheckCircle, AlertTriangle, LogOut } from 'lucide-react';
import { ComposedChart, Bar, Line, XAxis, YAxis, CartesianGrid, Tooltip as RechartsTooltip, Legend, ResponsiveContainer, ScatterChart, Scatter, Cell } from 'recharts';
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

export default function App() {
  const [user, setUser] = useState(null);
  const [email, setEmail] = useState('');
  const [password, setPassword] = useState('');
  const [loginError, setLoginError] = useState('');
  
  const [rawData, setRawData] = useState([]);
  const [historyList, setHistoryList] = useState([]);
  const [currentHistoryId, setCurrentHistoryId] = useState('');
  const [uploadStatus, setUploadStatus] = useState({ type: 'idle', msg: '' });
  const [filters, setFilters] = useState({ keyword: '', dateRange: 'all', category: 'all', deliveryType: 'all', tabType: 'all' });
  const [sortConfig, setSortConfig] = useState({ key: 'fulfillRate', direction: 'desc' });
  const [highlightedItem, setHighlightedItem] = useState(null);
  const [xlsxLoaded, setXlsxLoaded] = useState(false);
  
  const detailRef = useRef(null);

  useEffect(() => {
    const script = document.createElement('script');
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
    script.async = true;
    script.onload = () => setXlsxLoaded(true);
    document.body.appendChild(script);

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
      // 🚀 將 limit(5) 改為 limit(3)，只抓取最近 3 筆歷史紀錄
      const q = query(collection(db, "uploadHistory"), orderBy("timestamp", "desc"), limit(3));
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
      const loginAccount = email.includes('@') ? email : `${email}@fmc.com`;
      await signInWithEmailAndPassword(auth, loginAccount, password);
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

    // 🚀 關鍵修復：選取檔案後立刻清空 input 記憶，讓您可以重複上傳相同檔名的報表
    e.target.value = '';

    if (!xlsxLoaded || !window.XLSX) {
      setUploadStatus({ type: 'error', msg: '系統模組載入中，請稍後再試。' });
      return;
    }

    setUploadStatus({ type: 'loading', msg: '解析與上傳中...' });

    const reader = new FileReader();
    reader.onload = async (evt) => {
      try {
        const bstr = evt.target.result;
        const wb = window.XLSX.read(bstr, { type: 'binary' });
        let allData = [];

        wb.SheetNames.forEach(sheetName => {
          const ws = wb.Sheets[sheetName];
          const jsonData = window.XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

          for (let i = 0; i < jsonData.length; i++) {
            const row = jsonData[i];
            if (!row || row.length === 0) continue; 
            
            const itemId = (row[0] || '').toString().trim(); 
            
            let dateStr = (row[1] || '').toString().trim();
            const dateNums = dateStr.match(/\d+/g);
            if (dateNums) {
                const numStr = dateNums.join('');
                dateStr = numStr.length >= 4 ? numStr.slice(-4) : numStr;
            } else {
                const sheetNums = sheetName.match(/\d+/g);
                if (sheetNums) {
                     const sNumStr = sheetNums.join('');
                     dateStr = sNumStr.length >= 4 ? sNumStr.slice(-4) : sNumStr;
                } else {
                     dateStr = '未標示';
                }
            }

            const tabTypeRaw = (row[2] || '').toString().trim(); 
            const deliveryType = (row[3] || '').toString().trim(); 
            const categoryStr = (row[4] || '').toString().trim(); 
            const itemNameStr = (row[5] || '').toString().trim(); 
            const unitPrice = parseNumber(row[6]); 
            const orderAmt = parseNumber(row[8]); 
            const orderQty = parseNumber(row[10]); 
            const stockQty = parseNumber(row[11]); 
            
            const cleanId = itemId.replace(/\s+/g, '');
            const cleanName = itemNameStr.replace(/\s+/g, '');
            const cleanCat = categoryStr.replace(/\s+/g, '');
            const cleanTab = tabTypeRaw.replace(/\s+/g, '');
            
            if (!itemId || 
                cleanId.includes('合計') || cleanTab.includes('合計') || 
                cleanId.includes('代號') || cleanName.includes('名稱') || cleanCat.includes('採購別') || cleanTab.includes('頁籤別') ||
                cleanName.includes('未知商品') || cleanName.includes('無資料') || 
                cleanCat.includes('未分類') || cleanCat.includes('無資料')) {
              continue;
            }

            if (orderQty === 0 && orderAmt === 0) {
              continue;
            }

            const isNoLimit = deliveryType.toLowerCase().replace(/\s/g, '').includes('d+');
            
            let fulfillRate = 0;
            if (!isNoLimit && stockQty > 0) {
              fulfillRate = Math.min(orderQty / stockQty, 1);
            }

            allData.push({
              date: dateStr, 
              itemId: itemId,
              itemName: itemNameStr,
              tabType: tabTypeRaw,
              category: categoryStr,
              deliveryType: deliveryType || '一般',
              unitPrice: unitPrice,
              orderQty: orderQty,
              orderAmt: orderAmt,
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
      const cleanId = String(item.itemId || '').replace(/\s+/g, '');
      const cleanName = String(item.itemName || '').replace(/\s+/g, '');
      const cleanCat = String(item.category || '').replace(/\s+/g, '');
      const cleanTab = String(item.tabType || '').replace(/\s+/g, '');

      if (cleanId.includes('代號') || cleanName.includes('名稱') || cleanCat.includes('採購別') || cleanTab.includes('頁籤別') || cleanId.includes('合計')) {
         return false; 
      }

      if (item.orderQty === 0 && item.orderAmt === 0) {
        return false;
      }

      const matchKeyword = item.itemName.includes(filters.keyword) || item.itemId.toString().includes(filters.keyword);
      const matchDate = filters.dateRange === 'all' || item.date === filters.dateRange;
      const matchCategory = filters.category === 'all' || item.category === filters.category;
      const matchTabType = filters.tabType === 'all' || item.tabType === filters.tabType;
      
      const matchDelivery = filters.deliveryType === 'all' || 
                            String(item.deliveryType || '').toLowerCase().includes(filters.deliveryType.toLowerCase());
                            
      return matchKeyword && matchDate && matchCategory && matchTabType && matchDelivery;
    });
  }, [rawData, filters]);

  const sortedData = useMemo(() => {
    const sortableItems = [...filteredData];
    sortableItems.sort((a, b) => {
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

  const filterOptions = useMemo(() => {
    const cleanData = rawData.filter(item => {
      const cleanId = String(item.itemId || '').replace(/\s+/g, '');
      const cleanName = String(item.itemName || '').replace(/\s+/g, '');
      const cleanCat = String(item.category || '').replace(/\s+/g, '');
      const cleanTab = String(item.tabType || '').replace(/\s+/g, '');
      
      const isHeaderOrJunk = cleanId.includes('代號') || cleanName.includes('名稱') || cleanCat.includes('採購別') || cleanTab.includes('頁籤別') || cleanId.includes('合計');
      const isZeroSales = item.orderQty === 0 && item.orderAmt === 0; 
      
      return !isHeaderOrJunk && !isZeroSales;
    });
    
    return {
      dates: [...new Set(cleanData.map(d => d.date))].filter(Boolean).sort(),
      categories: [...new Set(cleanData.map(d => d.category))].filter(Boolean),
      tabTypes: [...new Set(cleanData.map(d => d.tabType))].filter(Boolean).sort(),
    };
  }, [rawData]);

  if (!user) {
    return (
      <div className="min-h-screen bg-gray-50 flex items-center justify-center p-4">
        <div className="bg-white p-8 rounded-xl shadow-lg max-w-md w-full">
          <div className="text-center mb-8">
            <Package className="h-12 w-12 text-blue-600 mx-auto mb-4"/>
            <h1 className="text-2xl font-bold text-gray-800">全家團購銷售分析系統</h1>
            <p className="text-gray-500 mt-2">請登入以檢視企業機密數據</p>
          </div>
          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">帳號</label>
              <input type="text" value={email} onChange={e => setEmail(e.target.value)} placeholder="請輸入自訂帳號" className="w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-blue-500 outline-none" required />
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
        
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center bg-white p-6 rounded-xl shadow-sm border border-gray-100">
          <div>
            <h1 className="text-2xl font-bold text-gray-800 flex items-center gap-2">
              <Package className="h-6 w-6 text-blue-600"/>
              全家團購商品銷售儀表板
            </h1>
            <p className="text-sm text-gray-500 mt-1">雲端安全連線中 | 帳號: {user.email.replace('@fmc.com', '')}</p>
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
              <Upload className="h-4 w-4"/>
              上傳新報表
              <input type="file" accept=".xlsx,.xls" className="hidden" onChange={handleFileUpload} />
            </label>
            
            <button onClick={handleLogout} className="p-2 text-gray-400 hover:text-red-600 transition" title="登出">
              <LogOut className="h-5 w-5"/>
            </button>
          </div>
        </div>

        <div className="bg-white p-4 rounded-xl shadow-sm border border-gray-100 grid grid-cols-1 md:grid-cols-5 gap-4">
          <div className="relative">
            <Search className="absolute left-3 top-2.5 h-4 w-4 text-gray-400"/>
            <input type="text" placeholder="搜尋代號或名稱..." value={filters.keyword} onChange={(e) => setFilters(f => ({ ...f, keyword: e.target.value }))} className="w-full pl-9 pr-4 py-2 bg-gray-50 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500" />
          </div>
          <select value={filters.dateRange} onChange={(e) => setFilters(f => ({ ...f, dateRange: e.target.value }))} className="px-4 py-2 bg-gray-50 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
            <option value="all">所有檔期</option>
            {filterOptions.dates.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
          <select value={filters.tabType} onChange={(e) => setFilters(f => ({ ...f, tabType: e.target.value }))} className="px-4 py-2 bg-gray-50 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
            <option value="all">所有頁籤別</option>
            {filterOptions.tabTypes.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
          <select value={filters.category} onChange={(e) => setFilters(f => ({ ...f, category: e.target.value }))} className="px-4 py-2 bg-gray-50 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
            <option value="all">所有採購別</option>
            {filterOptions.categories.map(c => <option key={c} value={c}>{c}</option>)}
          </select>
          <select value={filters.deliveryType} onChange={(e) => setFilters(f => ({ ...f, deliveryType: e.target.value }))} className="px-4 py-2 bg-gray-50 border rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
            <option value="all">所有交期</option>
            <option value="d+2">d+2</option>
            <option value="d+3">d+3</option>
            <option value="指取">指取</option>
            <option value="團宅">團宅</option>
          </select>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
          <div className="group relative bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-blue-50 text-blue-600 rounded-lg"><DollarSign className="h-5 w-5"/></div>
              <h3 className="text-gray-500 font-medium">總下單金額</h3>
            </div>
            <p className="text-2xl font-bold text-gray-800">{formatCurrency(stats.totalOrder)}</p>
          </div>
          <div className="group relative bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-purple-50 text-purple-600 rounded-lg"><Package className="h-5 w-5"/></div>
              <h3 className="text-gray-500 font-medium">總備貨金額</h3>
            </div>
            <p className="text-2xl font-bold text-gray-800">{formatCurrency(stats.totalStockAmt)}</p>
          </div>
          <div className="group relative bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-orange-50 text-orange-600 rounded-lg"><Target className="h-5 w-5"/></div>
              <h3 className="text-gray-500 font-medium">平均備貨完銷率</h3>
            </div>
            <p className="text-2xl font-bold text-gray-800">{formatPercent(stats.avgFulfill)}</p>
          </div>
          <div className="group relative bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <div className="flex items-center gap-3 mb-2">
              <div className="p-2 bg-green-50 text-green-600 rounded-lg"><TrendingUp className="h-5 w-5"/></div>
              <h3 className="text-gray-500 font-medium">熱門商品數 (完銷&ge;80%)</h3>
            </div>
            <p className="text-2xl font-bold text-gray-800">{stats.hotItems} <span className="text-sm font-normal text-gray-500">項</span></p>
          </div>
        </div>

        <div className="space-y-6">
          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <h3 className="text-lg font-bold text-gray-800 mb-6 flex items-center gap-2">各檔期銷售趨勢</h3>
            <div className="h-[300px] w-full">
              <ResponsiveContainer>
                <ComposedChart data={trendData}>
                  <CartesianGrid stroke="#E5E7EB" strokeDasharray="3 3" vertical={false}/>
                  <XAxis dataKey="date" tick={{ fill: '#6B7280', fontSize: 12 }}/>
                  <YAxis tickFormatter={val => `$${val/1000}k`} yAxisId="left" tick={{ fill: '#6B7280', fontSize: 12 }} />
                  <YAxis orientation="right" tickFormatter={formatPercent} yAxisId="right" tick={{ fill: '#6B7280', fontSize: 12 }}/>
                  <RechartsTooltip formatter={(value, name) => {
                      if (name === '總下單金額') return [formatCurrency(value), name];
                      if (name === '平均完銷率') return [formatPercent(value), name];
                      return [value, name];
                    }}
                  />
                  <Legend/>
                  <Bar dataKey="orderAmt" fill="#3B82F6" maxBarSize={50} name="總下單金額" radius={[4, 4, 0, 0]} yAxisId="left"/>
                  <Line activeDot={{ r: 6 }} dataKey="fulfillRate" dot={{ r: 4 }} name="平均完銷率" stroke="#F97316" strokeWidth={3} type="monotone" yAxisId="right" />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>

          <div className="bg-white p-6 rounded-xl shadow-sm border border-gray-100">
            <div className="flex justify-between items-center mb-6">
              <h3 className="text-lg font-bold text-gray-800 flex items-center gap-2">
                商品落點分析
                <span className="text-sm font-normal text-gray-400 bg-gray-50 px-2 py-1 rounded">排除 d+2/d+3</span>
              </h3>
            </div>
            <div className="h-[350px] w-full">
              <ResponsiveContainer>
                <ScatterChart margin={{ top: 20, right: 20, bottom: 20, left: 60 }}>
                  <CartesianGrid strokeDasharray="3 3" stroke="#E5E7EB" />
                  <XAxis type="number" dataKey="fulfillRate" name="完銷率" tickFormatter={formatPercent} domain={[0, 1]} label={{ value: '完銷率 (%)', position: 'bottom', offset: 0 }} />
                  <YAxis type="number" dataKey="orderAmt" name="下單金額" tickFormatter={val => `$${val/1000}k`} label={{ value: '下單金額', angle: -90, position: 'insideLeft', dx: -40 }} />
                  <RechartsTooltip cursor={{ strokeDasharray: '3 3' }} content={({ payload }) => {
                      if (payload && payload.length) {
                        const d = payload[0].payload;
                        return (
                          <div className="bg-white p-3 border border-gray-200 shadow-lg rounded-lg text-sm">
                            <p className="font-bold text-gray-800 border-b pb-1 mb-1">{d.itemId} - {d.itemName}</p>
                            <p className="text-blue-600">下單金額: {formatCurrency(d.orderAmt)}</p>
                            <p className="text-orange-600">完銷率: {formatPercent(d.fulfillRate)}</p>
                          </div>
                        );
                      }
                      return null;
                    }}
                  />
                  <Scatter name="商品" data={scatterData} 
                    onClick={(d) => {
                      setHighlightedItem(d.itemId);
                      setTimeout(() => {
                        const targetRow = document.getElementById(`row-${d.itemId}`);
                        if (targetRow) {
                          targetRow.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        } else {
                          detailRef.current?.scrollIntoView({ behavior: 'smooth', block: 'center' });
                        }
                      }, 100);
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
                    { key: 'tabType', label: '頁籤別' }, 
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
                            (sortConfig.direction === 'asc' ? <ChevronUp className="h-4 w-4 text-green-600"/> : <ChevronDown className="h-4 w-4 text-green-600"/>) 
                            : <ChevronDown className="h-4 w-4 opacity-0 group-hover:opacity-50"/>}
                        </span>
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sortedData.map((item, idx) => (
                  <tr key={`${item.itemId}-${idx}`} 
                      id={`row-${item.itemId}`}
                      className={`border-b hover:bg-blue-50 transition ${highlightedItem === item.itemId ? 'bg-green-50 border-2 border-green-500' : ''}`}>
                    <td className="px-4 py-3 whitespace-nowrap">{item.date}</td>
                    <td className="px-4 py-3 whitespace-nowrap">
                      {item.tabType ? <span className="px-2 py-1 bg-indigo-50 text-indigo-700 border border-indigo-100 rounded text-xs">{item.tabType}</span> : <span className="text-gray-400">-</span>}
                    </td>
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
                          {item.fulfillRate >= 0.8 && <CheckCircle className="h-4 w-4 text-green-500"/>}
                          {item.fulfillRate < 0.3 && <AlertTriangle className="h-4 w-4 text-red-400"/>}
                        </div>
                      )}
                    </td>
                  </tr>
                ))}
                {sortedData.length === 0 && (
                  <tr><td colSpan="11" className="px-4 py-8 text-center text-gray-500">目前沒有符合條件的資料</td></tr>
                )}
              </tbody>
            </table>
          </div>
        </div>

      </div>
    </div>
  );
}
