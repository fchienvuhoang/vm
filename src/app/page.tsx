"use client";

import { useState, useMemo, useEffect } from "react";
import { parseExcelBankStatement } from "@/lib/excelParser";
import { Category, ParsedRecord, UploadLog, UploadLogRecord } from "@/lib/types";
import {
  UploadCloud,
  FileSpreadsheet,
  Plus,
  Trash2,
  Tag,
  Search,
  Filter,
  CheckCircle,
  Pencil,
  Copy,
  Check,
  Loader2,
  Save,
  Sparkles,
  BarChart2,
  X,
  ArrowLeft,
  ChevronRight
} from "lucide-react";
import { 
  checkAuthStatus, 
  checkPasswordAction, 
  getCategoriesAction, 
  saveCategoriesAction, 
  saveAppliedCategoriesAction, 
  logoutAction,
  logFileUploadAction,
  getUploadStatsAction
} from "./actions";

const cleanNumber = (val: any) => {
  if (val === "" || val === null || val === undefined) return NaN;
  if (typeof val === 'number') return val;
  const str = String(val).trim();
  const usFormat = str.replace(/,/g, '');
  if (!isNaN(Number(usFormat)) && usFormat !== '') return Number(usFormat);
  const vnFormat = str.replace(/\./g, '').replace(/,/g, '.');
  if (!isNaN(Number(vnFormat)) && vnFormat !== '') return Number(vnFormat);
  return NaN;
};

const formatDateStr = (dateStr: string) => {
  if (!dateStr) return "";
  const parts = dateStr.match(/\d+/g);
  if (!parts || parts.length < 3) return dateStr;
  
  let day = parseInt(parts[0], 10);
  let month = parseInt(parts[1], 10);
  let year = parseInt(parts[2], 10);
  
  if (parts[0].length === 4) {
     year = parseInt(parts[0], 10);
     month = parseInt(parts[1], 10);
     day = parseInt(parts[2], 10);
  } else if (year < 100) {
     year += 2000;
  }
  
  let hours = parts[3] ? parseInt(parts[3], 10) : 0;
  let minutes = parts[4] ? parseInt(parts[4], 10) : 0;
  
  const mm = month.toString().padStart(2, '0');
  const dd = day.toString().padStart(2, '0');
  const yyyy = year.toString();
  const hh = hours.toString().padStart(2, '0');
  const min = minutes.toString().padStart(2, '0');
  
  return `${yyyy}-${mm}-${dd} ${hh}:${min}`;
};

const TABLE_COLUMNS = [
  { label: "Ngày giờ", key: "ngayGioGiaoDich" as keyof ParsedRecord, width: 180 },
  { label: "Tên chủ tài khoản", key: "tenChuTaiKhoan" as keyof ParsedRecord, width: 220 },
  { label: "Chi tiết giao dịch", key: "chiTietGiaoDich" as keyof ParsedRecord, left: 50, width: 350 },
  { label: "Tiền ra", key: "tienRa" as keyof ParsedRecord, width: 130 },
  { label: "Tiền vào", key: "tienVao" as keyof ParsedRecord, width: 130 },
];

const classifyRecords = (records: ParsedRecord[], categories: Category[]) => {
  return records.map((record) => {
    let matchedCategoryId: string | undefined;
    const searchNorm = record.searchText.toLowerCase().replace(/[\s\-_.+,/#!$%^&*;:{}=\\`~()]/g, '');

    for (const cat of categories) {
      const matched = cat.keywords.some((keyword) => {
        const keywordNorm = keyword.toLowerCase().replace(/[\s\-_.+,/#!$%^&*;:{}=\\`~()]/g, '');
        return keywordNorm.length > 0 && searchNorm.includes(keywordNorm);
      });
      if (matched) {
        matchedCategoryId = cat.id;
        break;
      }
    }

    return { ...record, matchedCategoryId };
  });
};

export default function Home() {
  const [categories, setCategories] = useState<Category[]>([]);
  const [appliedCategories, setAppliedCategories] = useState<Category[]>([]);
  const [records, setRecords] = useState<ParsedRecord[]>([]);
  const [isLoaded, setIsLoaded] = useState(false);
  
  const [isAuthenticated, setIsAuthenticated] = useState<boolean | null>(null);
  const [password, setPassword] = useState("");
  const [loginError, setLoginError] = useState("");

  useEffect(() => {
    async function init() {
      const auth = await checkAuthStatus();
      setIsAuthenticated(auth);
      
      if (auth) {
        const data = await getCategoriesAction();
        setCategories(data.categories);
        setAppliedCategories(data.appliedCategories);
        setIsLoaded(true);
      }
    }
    init();
  }, []);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoginError("");
    const success = await checkPasswordAction(password);
    if (success) {
      setIsAuthenticated(true);
      const data = await getCategoriesAction();
      setCategories(data.categories);
      setAppliedCategories(data.appliedCategories);
      setIsLoaded(true);
    } else {
      setLoginError("Mật khẩu không đúng");
    }
  };

  // Form states for category
  const [newCatKeywords, setNewCatKeywords] = useState("");
  const [editingId, setEditingId] = useState<string | null>(null);

  // Tab & UI status states
  const [activeTab, setActiveTab] = useState<string>("all");
  const [isApplying, setIsApplying] = useState(false);
  const [applyMessage, setApplyMessage] = useState("");
  const [copiedTab, setCopiedTab] = useState(false);
  const [copiedAll, setCopiedAll] = useState(false);

  // Stats Modal
  const [showStats, setShowStats] = useState(false);
  const [statsData, setStatsData] = useState<{ total: number; logs: UploadLog[] }>({ total: 0, logs: [] });
  const [isLoadingStats, setIsLoadingStats] = useState(false);
  const [selectedLog, setSelectedLog] = useState<UploadLog | null>(null);

  const handleOpenStats = async () => {
    setSelectedLog(null);
    setShowStats(true);
    setIsLoadingStats(true);
    const data = await getUploadStatsAction();
    setStatsData(data);
    setIsLoadingStats(false);
  };

  const handleAddOrEditCategory = () => {
    // Split by comma, trim whitespace, and remove empty
    const keywords = newCatKeywords
      .split(",")
      .map((k) => k.trim())
      .filter((k) => k.length > 0);

    if (keywords.length === 0) return;

    // Use the first keyword as the category name, capitalized
    const generatedName = keywords[0].toUpperCase();

    let updated: Category[];
    if (editingId) {
      updated = categories.map(c => c.id === editingId ? { ...c, name: generatedName, keywords } : c);
      setCategories(updated);
      setEditingId(null);
    } else {
      const newCategory: Category = {
        id: crypto.randomUUID(),
        name: generatedName,
        keywords,
      };
      updated = [...categories, newCategory];
      setCategories(updated);
    }

    // Save to database
    saveCategoriesAction(updated);

    // Reset form
    setNewCatKeywords("");
  };

  const handleEditCategory = (cat: Category) => {
    setNewCatKeywords(cat.keywords.join(", "));
    setEditingId(cat.id);
  };

  const handleRemoveCategory = (id: string) => {
    const updated = categories.filter((c) => c.id !== id);
    setCategories(updated);
    saveCategoriesAction(updated);

    if (editingId === id) {
      setEditingId(null);
      setNewCatKeywords("");
    }
  };

  const handleApplyCategories = () => {
    setIsApplying(true);
    setApplyMessage("");
    setTimeout(async () => {
      // Save to the applied state which drives the matching logic
      setAppliedCategories([...categories]);
      await saveAppliedCategoriesAction(categories);

      // If active tab doesn't exist anymore, reset to "all"
      if (!categories.find(c => c.id === activeTab) && activeTab !== "all" && activeTab !== "uncategorized") {
        setActiveTab("all");
      }
      setIsApplying(false);
      setApplyMessage("Đã phân loại xong dữ liệu!");
      setTimeout(() => setApplyMessage(""), 4000);
    }, 600); // 600ms loading simulation
  };

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const parsed = await parseExcelBankStatement(file);
      setRecords(parsed);

      const classified = classifyRecords(parsed, appliedCategories);
      const logRecords: UploadLogRecord[] = classified
        .filter((record) => !record.isFooter)
        .map((record) => ({
          soThamChieu: record.soThamChieu,
          ngayGioGiaoDich: record.ngayGioGiaoDich,
          tenChuTaiKhoan: record.tenChuTaiKhoan,
          chiTietGiaoDich: record.chiTietGiaoDich,
          tienRa: record.tienRa,
          tienVao: record.tienVao,
          categoryName: appliedCategories.find((cat) => cat.id === record.matchedCategoryId)?.name || "Chưa phân loại",
        }));

      logFileUploadAction(parsed.length, logRecords).catch(console.error);
    } catch (error) {
      console.error("Error parsing file", error);
      alert("Đã xảy ra lỗi khi đọc file Excel.");
    }
  };

  const handleCopyTable = () => {
    if (filteredRecords.length === 0) return;

    // Headers
    const headers = ["STT", ...TABLE_COLUMNS.map(c => c.label)].join("\t");
    // Rows
    const rows = filteredRecords.map((record, index) => {
      const rowStr = TABLE_COLUMNS.map(col => {
        let val = record[col.key];
        if (col.key === "ngayGioGiaoDich") val = formatDateStr(String(val || ""));
        else if (col.key === "tienRa" || col.key === "tienVao") {
          const num = cleanNumber(val);
          if (!isNaN(num)) val = num;
        }
        // Replacing newlines so it doesn't break Excel TSV rows
        return String(val || "").replace(/\n/g, " ");
      }).join("\t");
      const stt = record.isFooter ? "" : (index + 1);
      return `${stt}\t${rowStr}`;
    }).join("\n");

    navigator.clipboard.writeText(rows);
    setCopiedTab(true);
    setTimeout(() => setCopiedTab(false), 2000);
  };

  const handleCopyAllData = () => {
    if (computedRecords.length === 0) return;

    // Headers with Nhóm Phân Loại
    const headers = ["Nhóm Phân Loại", "STT", ...TABLE_COLUMNS.map(c => c.label)].join("\t");
    
    // Rows
    const rows = computedRecords.filter(r => !r.isFooter).map((record, index) => {
      let groupName = "Chưa phân loại";
      if (record.matchedCategoryId) {
        groupName = appliedCategories.find(c => c.id === record.matchedCategoryId)?.name || "Chưa phân loại";
      }

      const rowStr = TABLE_COLUMNS.map(col => {
        let val = record[col.key];
        if (col.key === "ngayGioGiaoDich") val = formatDateStr(String(val || ""));
        else if (col.key === "tienRa" || col.key === "tienVao") {
           const num = cleanNumber(val);
           if (!isNaN(num)) val = num;
        }
        return String(val || "").replace(/\n/g, " ");
      }).join("\t");
      
      return `${groupName}\t${index + 1}\t${rowStr}`;
    }).join("\n");

    navigator.clipboard.writeText(rows);
    setCopiedAll(true);
    setTimeout(() => setCopiedAll(false), 2000);
  };

  // Compute matches
  const computedRecords = useMemo(() => {
    return classifyRecords(records, appliedCategories);
  }, [records, appliedCategories]);

  // Filter records based on active tab
  const filteredRecords = useMemo(() => {
    if (activeTab === "all") return computedRecords;
    if (activeTab === "uncategorized")
      return computedRecords.filter((r) => !r.matchedCategoryId);

    return computedRecords.filter((r) => r.matchedCategoryId === activeTab);
  }, [computedRecords, activeTab]);

  // Auto-generate suggestions from chiTietGiaoDich
  const suggestedKeywords = useMemo(() => {
    if (records.length === 0) return [];

    const STOP_WORDS = new Set([
      "chuyển", "chuyen", "tiền", "tien", "ck", "cho", "thanh", "toán", "toan", "từ", "tu",
      "tk", "tài", "khoản", "tai", "khoan", "ngân", "hàng", "ngan", "hang", "nh", "đến", "den",
      "vào", "vao", "ra", "phí", "phi", "gd", "giao", "dịch", "dich", "số", "so", "thẻ", "the",
      "qua", "mb", "vcb", "bidv", "tcb", "vietcombank", "techcombank", "mbbank", "ngày", "ngay",
      "tháng", "thang", "năm", "nam", "ibft", "nạp", "nap", "rút", "rut", "phát", "sinh", "phat",
      "chi", "thu", "của", "cua", "nhận", "nhan", "gửi", "gui", "qrvn", "vnpay", "mã", "ma", "vietinbank", "agribank"
    ]);

    const counts: Record<string, number> = {};

    records.forEach(r => {
      const text = r.chiTietGiaoDich.toLowerCase().trim();
      if (!text) return;

      // Extract words without stripping out numbers too early so bigrams are built correctly
      const words = text
        .replace(/[.,/#!$%^&*;:{}=\-_`~()+\\|\n]/g, " ")
        .split(/\s+/)
        .filter(w => w.length >= 2 && !STOP_WORDS.has(w));

      // 1-grams (ignore standalone numbers)
      words.forEach(w => {
        if (!/^\d+$/.test(w)) {
          counts[w] = (counts[w] || 0) + 1;
        }
      });

      // 2-grams (ignore pure numbers combined like '123 456')
      for (let i = 0; i < words.length - 1; i++) {
        const bigram = words[i] + " " + words[i + 1];
        if (!/^\d+\s\d+$/.test(bigram)) {
          counts[bigram] = (counts[bigram] || 0) + 1;
        }
      }
    });

    return Object.entries(counts)
      .filter(([word, count]) => count > 1) // appear more than once
      .sort((a, b) => {
        const aHasTp = a[0].startsWith('tp');
        const bHasTp = b[0].startsWith('tp');
        if (aHasTp && !bHasTp) return -1;
        if (!aHasTp && bHasTp) return 1;
        // Secondary sort by frequency
        return b[1] - a[1];
      })
      .slice(0, 15) // top 15 suggestions
      .map(entry => entry[0]);
  }, [records]);

  if (isAuthenticated === null) {
    return <div className="min-h-screen flex items-center justify-center bg-gray-50"><Loader2 className="w-8 h-8 animate-spin text-indigo-600" /></div>;
  }

  if (!isAuthenticated) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50 flex-col p-4 font-sans">
        <div className="w-full max-w-sm bg-white p-8 rounded-2xl shadow-sm border border-gray-100">
           <h1 className="text-2xl font-bold text-gray-900 mb-6 text-center">Đăng Nhập Vimutti</h1>
           <form onSubmit={handleLogin} className="flex flex-col gap-4">
              <div>
                 <label className="block text-sm font-medium text-gray-700 mb-1">Mật khẩu</label>
                 <input 
                   type="password" 
                   value={password} 
                   onChange={e => setPassword(e.target.value)} 
                   className="w-full border border-gray-300 rounded-lg px-4 py-2 focus:ring-2 focus:ring-indigo-500 outline-none" 
                   placeholder="Nhập mật khẩu..." 
                   autoFocus
                 />
              </div>
              {loginError && <p className="text-red-500 text-sm">{loginError}</p>}
              <button type="submit" className="w-full bg-indigo-600 text-white font-semibold py-2.5 rounded-lg hover:bg-indigo-700 transition">
                Vào Trang
              </button>
           </form>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 text-gray-900 p-6 md:p-10 font-sans">
      <header className="mb-10 max-w-6xl mx-auto flex justify-between items-center">
        <div>
          <h1 className="text-3xl font-bold tracking-tight text-gray-900">
            Vimutti - Phân loại sao kê
          </h1>
          <p className="text-gray-500 mt-2">
            Bóc tách thông tin thiện pháp dựa trên từ khóa khớp với nội dung giao dịch.
          </p>
        </div>
        <div className="flex items-center gap-3">
          <button 
            onClick={handleOpenStats}
            className="flex items-center gap-2 text-sm cursor-pointer font-medium text-gray-700 bg-white px-4 py-2 rounded-lg border border-gray-200 hover:bg-gray-50 transition shadow-sm"
          >
            <BarChart2 className="w-4 h-4" />
            Thống kê
          </button>
          <button 
            onClick={async () => { await logoutAction(); setIsAuthenticated(false); }}
            className="text-sm cursor-pointer font-medium text-gray-600 bg-white px-4 py-2 rounded-lg border border-gray-200 hover:bg-gray-50 transition shadow-sm"
          >
            Đăng xuất
          </button>
        </div>
      </header>

      <main className="max-w-6xl mx-auto flex flex-col gap-8">

        {/* TOP SECTION */}
        <div className="flex flex-col gap-6">
          {/* File Upload Panel */}
          <section className="bg-white rounded-2xl p-4 shadow-sm border border-gray-100 flex flex-col sm:flex-row items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <UploadCloud className="w-5 h-5 text-indigo-600" />
              <h2 className="text-lg font-semibold m-0">Tải Lên Sao Kê</h2>
            </div>
            
            <div className="flex items-center gap-4 w-full sm:w-auto justify-between sm:justify-end">
              {records.length > 0 && (
                <div className="text-sm text-green-700 flex items-center gap-1.5 font-medium bg-green-50 px-3 py-1.5 rounded-lg border border-green-100">
                  <Search className="w-4 h-4" />
                  Đã tải {records.length} dòng
                </div>
              )}
              <label className="cursor-pointer bg-indigo-600 text-white px-4 py-2.5 rounded-lg text-sm font-medium hover:bg-indigo-700 transition flex items-center gap-2 shadow-sm">
                <FileSpreadsheet className="w-4 h-4" />
                Chọn file Excel
                <input
                  type="file"
                  accept=".xlsx, .xls"
                  className="hidden"
                  onChange={handleFileUpload}
                />
              </label>
            </div>
          </section>

          {/* Keyword Management Panel */}
          <section className="bg-white rounded-2xl p-6 shadow-sm border border-gray-100 flex flex-col relative overflow-hidden">
            <h2 className="text-xl font-semibold mb-4 flex items-center gap-2">
              <Filter className="w-5 h-5 text-indigo-600" />
              Cài Đặt Thiện Pháp
            </h2>

            {/* Add form */}
            <div className={`bg-gray-50 rounded-xl p-4 border mb-4 transition-colors ${editingId ? 'border-yellow-300 bg-yellow-50' : 'border-gray-100'}`}>
              <div className="space-y-3">
                <div>
                  <label className="block text-xs font-medium text-gray-600 mb-1">
                    Danh Sách Từ Khóa (cách nhau bởi dấu phẩy)
                  </label>
                  <input
                    type="text"
                    value={newCatKeywords}
                    onChange={(e) => setNewCatKeywords(e.target.value)}
                    placeholder="tp44, hoa sen, ..."
                    className="w-full border border-gray-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none"
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') handleAddOrEditCategory();
                    }}
                  />
                  {suggestedKeywords.length > 0 && (
                    <div className="mt-3 bg-white p-2 rounded border border-gray-100">
                      <span className="flex items-center gap-1 text-[11px] font-medium text-amber-600 mb-2">
                        <Sparkles className="w-3 h-3" />
                        Gợi ý từ khóa thông minh (Nhấn để thêm):
                      </span>
                      <div className="flex flex-wrap gap-1.5">
                        {suggestedKeywords.map((kw, i) => (
                          <button
                            key={i}
                            onClick={() => {
                              const current = newCatKeywords.trim();
                              setNewCatKeywords(current ? current + ", " + kw : kw);
                            }}
                            className="text-[11px] bg-amber-50 text-amber-700 px-2 py-1 rounded hover:bg-amber-100 transition cursor-pointer"
                          >
                            {kw}
                          </button>
                        ))}
                      </div>
                    </div>
                  )}
                </div>
                <div className="flex gap-2 mt-2">
                  <button
                    onClick={handleAddOrEditCategory}
                    className="w-full flex items-center justify-center cursor-pointer gap-2 bg-gray-900 text-white px-4 py-2 rounded-lg text-sm font-medium hover:bg-gray-800 transition"
                  >
                    {editingId ? <Save className="w-4 h-4" /> : <Plus className="w-4 h-4" />}
                    {editingId ? "Lưu Chỉnh Sửa" : "Thêm Nhóm Từ Khóa"}
                  </button>
                  {editingId && (
                    <button
                      onClick={() => {
                        setEditingId(null);
                        setNewCatKeywords("");
                      }}
                      className="whitespace-nowrap px-4 py-2 cursor-pointer bg-gray-200 text-gray-700 rounded-lg text-sm font-medium hover:bg-gray-300 transition"
                    >
                      Hủy Sửa
                    </button>
                  )}
                </div>
              </div>
            </div>

            {/* List of draft categories */}
            <div className="flex-1 overflow-y-auto space-y-3 max-h-[500px] mb-4 pr-1">
              {categories.map((cat) => (
                <div key={cat.id} className="border border-gray-200 rounded-xl p-3 hover:shadow-sm transition bg-white">
                  <div className="flex justify-between items-start mb-2">
                    <h3 className="font-medium text-sm text-gray-900">{cat.name}</h3>
                    <div className="flex gap-1 -mt-1">
                      <button
                        onClick={() => handleEditCategory(cat)}
                        className="text-gray-400 hover:text-blue-500 transition p-1 cursor-pointer"
                        title="Sửa"
                      >
                        <Pencil className="w-4 h-4" />
                      </button>
                      <button
                        onClick={() => handleRemoveCategory(cat.id)}
                        className="text-gray-400 hover:text-red-500 transition p-1 cursor-pointer"
                        title="Xóa"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                  <div className="flex flex-wrap gap-2">
                    {cat.keywords.map((kw, i) => (
                      <span key={i} className="inline-flex items-center gap-1 bg-indigo-50 text-indigo-700 text-[10px] px-2 py-1 rounded-md font-medium">
                        <Tag className="w-3 h-3" />
                        {kw}
                      </span>
                    ))}
                  </div>
                </div>
              ))}
              {categories.length === 0 && (
                <div className="text-sm text-center text-gray-500 py-4">
                  Chưa có nhóm từ khóa nào được thiết lập.
                </div>
              )}
            </div>

            {/* Apply Button */}
            <button
              onClick={handleApplyCategories}
              disabled={isApplying}
              className={`w-full relative mt-auto flex items-center justify-center cursor-pointer gap-2 bg-indigo-600 text-white px-4 py-3 rounded-lg text-base font-semibold transition shadow-sm ${isApplying ? 'bg-indigo-500 cursor-wait' : 'hover:bg-indigo-700'
                }`}
            >
              {isApplying ? (
                <>
                  <Loader2 className="w-5 h-5 animate-spin" />
                  Đang phân loại...
                </>
              ) : (
                <>
                  <CheckCircle className="w-5 h-5" />
                  Áp Dụng Phân Loại
                </>
              )}
            </button>

            {/* Success Banner */}
            <div
              className={`mt-4 overflow-hidden transition-all duration-300 ease-in-out ${applyMessage ? 'max-h-20 opacity-100' : 'max-h-0 opacity-0'
                }`}
            >
              <div className="bg-emerald-50 text-emerald-700 text-sm font-medium border border-emerald-200 rounded-lg p-3 flex items-center gap-2">
                <Check className="w-4 h-4 text-emerald-600 flex-shrink-0" />
                {applyMessage}
              </div>
            </div>

          </section>
        </div>

        {/* BOTTOM SECTION: Results & Display */}
        <div className="w-full">
          <section className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden min-h-[500px] flex flex-col sm:flex-row relative">

            {/* Left Sidebar for Tabs */}
            <div className="w-full sm:w-56 flex-shrink-0 flex flex-col border-b sm:border-b-0 sm:border-r border-gray-200 bg-gray-50">
              {/* Tabs */}
              <div className="flex-1 p-3 flex flex-col gap-1.5 overflow-y-auto max-h-[250px] sm:max-h-none">
                <button
                  onClick={() => setActiveTab("all")}
                  className={`w-full text-left flex-shrink-0 cursor-pointer px-3 py-2 text-sm font-medium rounded-lg transition-colors ${activeTab === "all" ? "bg-white text-indigo-600 shadow-sm border border-gray-200/60" : "text-gray-600 hover:bg-gray-200/50"
                    }`}
                >
                  Tất cả ({computedRecords.filter(r => !r.isFooter).length})
                </button>
                {appliedCategories.map(cat => (
                  <button
                    key={cat.id}
                    onClick={() => setActiveTab(cat.id)}
                    className={`w-full text-left flex-shrink-0 cursor-pointer px-3 py-2 text-sm font-medium rounded-lg transition-colors ${activeTab === cat.id ? "bg-white text-indigo-600 shadow-sm border border-gray-200/60" : "text-gray-600 hover:bg-gray-200/50"
                      }`}
                  >
                    {cat.name} ({computedRecords.filter(r => r.matchedCategoryId === cat.id && !r.isFooter).length})
                  </button>
                ))}
                <button
                  onClick={() => setActiveTab("uncategorized")}
                  className={`w-full text-left flex-shrink-0 cursor-pointer px-3 py-2 text-sm font-medium rounded-lg transition-colors ${activeTab === "uncategorized" ? "bg-white text-indigo-600 shadow-sm border border-gray-200/60" : "text-gray-600 hover:bg-gray-200/50"
                    }`}
                >
                  Chưa phân loại ({computedRecords.filter(r => !r.matchedCategoryId && !r.isFooter).length})
                </button>
              </div>

            </div>

            {/* Table Area */}
            <div className="flex-1 flex flex-col min-w-0 bg-white">
              {/* Header Copy Button */}
              {filteredRecords.length > 0 && (
                <div className="p-2 border-b border-gray-100 flex justify-end gap-2 bg-white z-10 sticky top-0">
                  <button
                    onClick={handleCopyAllData}
                    className="flex justify-center items-center cursor-pointer gap-2 px-3 py-1.5 text-sm font-medium rounded-lg bg-indigo-50 text-indigo-700 hover:bg-indigo-100 transition shadow-sm"
                  >
                    {copiedAll ? (
                      <>
                        <Check className="w-4 h-4 text-indigo-500" />
                        Đã Copy Tất Cả!
                      </>
                    ) : (
                      <>
                        <Copy className="w-4 h-4" />
                        Copy tất cả tab
                      </>
                    )}
                  </button>
                  <button
                    onClick={handleCopyTable}
                    className="flex justify-center items-center cursor-pointer gap-2 px-3 py-1.5 text-sm font-medium rounded-lg bg-gray-900 text-white hover:bg-gray-800 transition shadow-sm"
                  >
                    {copiedTab ? (
                      <>
                        <Check className="w-4 h-4 text-green-400" />
                        Đã Copy!
                      </>
                    ) : (
                      <>
                        <Copy className="w-4 h-4" />
                        Copy bảng {activeTab === "all" ? "Tất cả" : activeTab === "uncategorized" ? "Chưa phân loại" : appliedCategories.find(c => c.id === activeTab)?.name || ""}
                      </>
                    )}
                  </button>
                </div>
              )}

              {/* Table Wrapper */}
              <div className="flex-1 overflow-auto p-0">
              {filteredRecords.length === 0 ? (
                <div className="flex flex-col items-center justify-center h-full text-gray-400 py-20 min-h-[300px]">
                  <FileSpreadsheet className="w-12 h-12 mb-3 opacity-20" />
                  <p>Không có dữ liệu hiển thị.</p>
                </div>
              ) : (
                <table className="w-full text-left border-collapse text-sm">
                  <thead className="bg-gray-50 sticky top-0 bg-white border-b border-gray-200 z-30">
                    <tr>
                      <th
                        className="px-4 py-3 font-medium text-gray-500 whitespace-nowrap text-center sticky left-0 z-30 bg-gray-50 border-r border-gray-200 shadow-[1px_0_0_0_#e5e7eb]"
                        style={{ width: "50px", minWidth: "50px", maxWidth: "50px" }}
                      >
                        STT
                      </th>
                      {TABLE_COLUMNS.map((col, i) => (
                        <th
                          key={i}
                          className={`px-4 py-3 font-medium text-gray-500 whitespace-nowrap bg-gray-50 ${col.left !== undefined ? 'sticky z-30 border-r border-gray-200 shadow-[1px_0_0_0_#e5e7eb]' : 'z-20 relative'
                            }`}
                          style={{
                            left: col.left !== undefined ? `${col.left}px` : undefined,
                            width: col.width ? `${col.width}px` : undefined,
                            minWidth: col.width ? `${col.width}px` : undefined,
                            maxWidth: col.width ? `${col.width}px` : undefined,
                          }}
                        >
                          {col.label}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-100">
                    {filteredRecords.map((record, rowIndex) => {
                      let rowBgClass = 'hover:bg-gray-50';
                      let stickyBgClass = 'bg-white group-hover:bg-gray-50';

                      if (record.isFooter) {
                        rowBgClass = 'bg-amber-50 hover:bg-amber-100/50';
                        stickyBgClass = 'bg-amber-50 group-hover:bg-amber-100/50';
                      } else if (record.tienRa) {
                        rowBgClass = 'bg-slate-100/60 hover:bg-slate-200/50';
                        stickyBgClass = 'bg-slate-100/60 group-hover:bg-slate-200/50';
                      }

                      return (
                        <tr key={rowIndex} className={`transition-colors group ${rowBgClass}`}>
                          <td
                            className={`px-4 py-3 whitespace-nowrap text-center font-medium sticky left-0 z-10 border-r border-gray-100 shadow-[1px_0_0_0_#f3f4f6] ${stickyBgClass} ${record.tienRa && !record.isFooter ? 'text-slate-500' : 'text-gray-700'}`}
                            style={{ width: "50px", minWidth: "50px", maxWidth: "50px" }}
                          >
                            {record.isFooter ? "" : (rowIndex + 1)}
                          </td>
                          {TABLE_COLUMNS.map((col, colIndex) => {
                            const val = record[col.key];
                            let displayVal = String(val || "");
                            let baseClass = `px-4 py-3 ${record.tienRa && !record.isFooter ? 'text-slate-500' : 'text-gray-700'}`;

                            if (col.key === "chiTietGiaoDich") {
                              baseClass += " min-w-[250px] whitespace-pre-wrap break-words leading-relaxed";
                            } else if (col.key === "tienRa" || col.key === "tienVao") {
                              if (val !== "" && val !== null && val !== undefined) {
                                const numVal = cleanNumber(val);
                                if (!isNaN(numVal)) {
                                  displayVal = new Intl.NumberFormat('en-US').format(numVal);
                                }
                              }
                              baseClass += " font-medium whitespace-nowrap text-left text-slate-700";
                            } else if (col.key === "ngayGioGiaoDich") {
                              displayVal = formatDateStr(displayVal);
                              baseClass += " whitespace-nowrap text-sm text-slate-600";
                            } else if (col.key === "tenChuTaiKhoan") {
                              baseClass += " whitespace-normal break-words text-sm";
                            } else {
                              baseClass += " whitespace-nowrap";
                            }

                            const styleObj: any = {
                              left: col.left !== undefined ? `${col.left}px` : undefined,
                              width: col.width ? `${col.width}px` : undefined,
                              minWidth: col.width ? `${col.width}px` : undefined,
                              maxWidth: col.width ? `${col.width}px` : undefined,
                            };

                            return (
                              <td
                                key={colIndex}
                                className={`${baseClass} ${col.left !== undefined ? `sticky z-10 border-r border-gray-100 shadow-[1px_0_0_0_#f3f4f6] ${stickyBgClass}` : ''
                                  }`}
                                style={styleObj}
                              >
                                {displayVal}
                              </td>
                            );
                          })}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>

              {/* Footer Copy Button */}
              {filteredRecords.length > 0 && (
                <div className="p-2 border-t border-gray-100 flex justify-end gap-2 bg-white">
                  <button
                    onClick={handleCopyAllData}
                    className="flex justify-center items-center cursor-pointer gap-2 px-3 py-1.5 text-sm font-medium rounded-lg bg-indigo-50 text-indigo-700 hover:bg-indigo-100 transition shadow-sm"
                  >
                    {copiedAll ? (
                      <>
                        <Check className="w-4 h-4 text-indigo-500" />
                        Đã Copy Tất Cả!
                      </>
                    ) : (
                      <>
                        <Copy className="w-4 h-4" />
                        Copy tất cả tab
                      </>
                    )}
                  </button>
                  <button
                    onClick={handleCopyTable}
                    className="flex justify-center items-center cursor-pointer gap-2 px-3 py-1.5 text-sm font-medium rounded-lg bg-gray-900 text-white hover:bg-gray-800 transition shadow-sm"
                  >
                    {copiedTab ? (
                      <>
                        <Check className="w-4 h-4 text-green-400" />
                        Đã Copy!
                      </>
                    ) : (
                      <>
                        <Copy className="w-4 h-4" />
                        Copy bảng {activeTab === "all" ? "Tất cả" : activeTab === "uncategorized" ? "Chưa phân loại" : appliedCategories.find(c => c.id === activeTab)?.name || ""}
                      </>
                    )}
                  </button>
                </div>
              )}
            </div>

          </section>
        </div>
      </main>

      {/* STATS MODAL */}
      {showStats && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className={`bg-white rounded-2xl w-full ${selectedLog ? "max-w-6xl" : "max-w-lg"} shadow-xl overflow-hidden flex flex-col max-h-[90vh] transition-all`}>
            <div className="p-4 border-b border-gray-100 flex justify-between items-center bg-gray-50">
              <h2 className="text-lg font-bold flex items-center gap-2">
                {selectedLog ? (
                  <button
                    onClick={() => setSelectedLog(null)}
                    className="p-1 -ml-1 rounded-md text-gray-500 hover:text-gray-800 hover:bg-gray-200 transition cursor-pointer"
                    aria-label="Quay lại lịch sử"
                  >
                    <ArrowLeft className="w-5 h-5" />
                  </button>
                ) : (
                  <BarChart2 className="w-5 h-5 text-indigo-600" />
                )}
                {selectedLog ? "Chi tiết lần phân loại" : "Thống kê sử dụng"}
              </h2>
              <button onClick={() => { setShowStats(false); setSelectedLog(null); }} className="text-gray-400 hover:text-gray-600 p-1 flex-shrink-0 cursor-pointer transition">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className={`${selectedLog ? "p-0" : "p-6"} flex-1 overflow-y-auto`}>
              {isLoadingStats ? (
                <div className="flex justify-center py-10"><Loader2 className="w-8 h-8 animate-spin text-indigo-600" /></div>
              ) : selectedLog ? (
                <div>
                  <div className="px-5 py-4 bg-indigo-50 border-b border-indigo-100 flex flex-wrap items-center justify-between gap-2">
                    <div>
                      <p className="text-sm text-indigo-700">Thời gian xử lý</p>
                      <p className="font-semibold text-gray-900">{new Date(selectedLog.time).toLocaleString('vi-VN')}</p>
                    </div>
                    <span className="font-semibold text-indigo-800 bg-white px-3 py-1.5 rounded-lg border border-indigo-200">
                      {selectedLog.rowCount} dòng
                    </span>
                  </div>
                  <div className="overflow-auto">
                    <table className="w-full min-w-[950px] text-left border-collapse text-sm">
                      <thead className="bg-gray-50 sticky top-0 z-10 border-b border-gray-200">
                        <tr>
                          <th className="px-4 py-3 font-medium text-gray-500 w-14 text-center">STT</th>
                          <th className="px-4 py-3 font-medium text-gray-500 whitespace-nowrap">Nhóm phân loại</th>
                          <th className="px-4 py-3 font-medium text-gray-500 whitespace-nowrap">Ngày giờ</th>
                          <th className="px-4 py-3 font-medium text-gray-500 whitespace-nowrap">Tên chủ tài khoản</th>
                          <th className="px-4 py-3 font-medium text-gray-500 min-w-[300px]">Chi tiết giao dịch</th>
                          <th className="px-4 py-3 font-medium text-gray-500 whitespace-nowrap">Tiền ra</th>
                          <th className="px-4 py-3 font-medium text-gray-500 whitespace-nowrap">Tiền vào</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100">
                        {(selectedLog.records || []).map((record, index) => (
                          <tr key={`${record.soThamChieu}-${index}`} className="hover:bg-gray-50">
                            <td className="px-4 py-3 text-center text-gray-500">{index + 1}</td>
                            <td className="px-4 py-3">
                              <span className={`inline-flex rounded-full px-2.5 py-1 text-xs font-semibold whitespace-nowrap ${record.categoryName === "Chưa phân loại" ? "bg-gray-100 text-gray-600" : "bg-indigo-50 text-indigo-700"}`}>
                                {record.categoryName}
                              </span>
                            </td>
                            <td className="px-4 py-3 text-gray-600 whitespace-nowrap">{formatDateStr(record.ngayGioGiaoDich)}</td>
                            <td className="px-4 py-3 text-gray-700">{record.tenChuTaiKhoan}</td>
                            <td className="px-4 py-3 text-gray-700 whitespace-pre-wrap break-words">{record.chiTietGiaoDich}</td>
                            <td className="px-4 py-3 text-gray-700 whitespace-nowrap">{record.tienRa !== "" ? new Intl.NumberFormat('en-US').format(cleanNumber(record.tienRa)) : ""}</td>
                            <td className="px-4 py-3 text-gray-700 whitespace-nowrap">{record.tienVao !== "" ? new Intl.NumberFormat('en-US').format(cleanNumber(record.tienVao)) : ""}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              ) : (
                <div className="space-y-6">
                  <div className="bg-indigo-50 rounded-xl p-4 flex flex-col items-center justify-center border border-indigo-100 shadow-sm">
                     <span className="text-sm text-indigo-800 font-medium mb-1">Tổng lượt tải file lên</span>
                     <span className="text-4xl font-bold text-indigo-600">{statsData.total}</span>
                  </div>
                  <div>
                    <h3 className="text-sm font-semibold text-gray-700 mb-3 border-b flex items-center pb-2">
                       Lịch sử 100 lần gần nhất
                    </h3>
                    {statsData.logs.length === 0 ? (
                      <p className="text-sm text-gray-500 text-center py-4 bg-gray-50 rounded-lg">Chưa có dữ liệu tải lên.</p>
                    ) : (
                      <div className="space-y-2">
                        {statsData.logs.map((log, idx) => {
                          const loadDate = new Date(log.time);
                          return (
                            <button
                              key={log.id || `${log.time}-${idx}`}
                              type="button"
                              disabled={!log.records}
                              onClick={() => log.records && setSelectedLog(log)}
                              className={`w-full flex justify-between items-center text-sm p-3 bg-gray-50 transition rounded-lg border border-gray-100 text-left ${log.records ? "hover:bg-indigo-50 hover:border-indigo-200 cursor-pointer group" : "cursor-not-allowed opacity-70"}`}
                            >
                              <span className="text-gray-600 font-medium">{loadDate.toLocaleString('vi-VN')}</span>
                              <span className="flex items-center gap-2">
                                {!log.records && <span className="text-xs text-gray-400">Log cũ, không có chi tiết</span>}
                                <span className="font-semibold text-gray-800 bg-white shadow-sm px-2.5 py-1 rounded-md border border-gray-200">
                                  {log.rowCount} dòng
                                </span>
                                {log.records && <ChevronRight className="w-4 h-4 text-gray-400 group-hover:text-indigo-600" />}
                              </span>
                            </button>
                          );
                        })}
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
