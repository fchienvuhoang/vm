"use client";

import { useState, useMemo, useEffect, type CSSProperties } from "react";
import { parseExcelBankStatement } from "@/lib/excelParser";
import { Category, GoogleSheetFile, GoogleSheetTab, ParsedRecord, SheetSyncRecord, UploadLog, UploadLogRecord } from "@/lib/types";
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
  ChevronRight,
  ChevronDown,
  CircleAlert,
  ExternalLink,
  Link2,
  Send,
  Unlink
} from "lucide-react";
import { 
  checkAuthStatus, 
  checkPasswordAction, 
  getCategoriesAction, 
  saveCategoriesAction, 
  saveAppliedCategoriesAction, 
  logoutAction,
  logFileUploadAction,
  getUploadStatsAction,
  getGoogleSheetFilesAction,
  getGoogleSheetTabsAction,
  syncCategoryToSheetAction
} from "./actions";

const cleanNumber = (val: unknown) => {
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
  
  const hours = parts[3] ? parseInt(parts[3], 10) : 0;
  const minutes = parts[4] ? parseInt(parts[4], 10) : 0;
  
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

const normalizeSearchText = (value: string) => value
  .normalize("NFD")
  .replace(/[\u0300-\u036f]/g, "")
  .replace(/đ/g, "d")
  .replace(/Đ/g, "D")
  .toLowerCase();

const buildSheetSyncPayload = (records: ParsedRecord[], categoryId: string): SheetSyncRecord[] => records
  .filter((record) => record.matchedCategoryId === categoryId && !record.isFooter)
  .map((record) => ({
    soThamChieu: record.soThamChieu,
    nganHang: record.nganHang,
    soTaiKhoan: record.soTaiKhoan,
    ngayGioGiaoDich: record.ngayGioGiaoDich,
    tenChuTaiKhoan: record.tenChuTaiKhoan,
    chiTietGiaoDich: record.chiTietGiaoDich,
    tienRa: record.tienRa,
    tienVao: record.tienVao,
    ghiChu: record.ghiChu,
  }));

const createSheetSyncSignature = (category: Category, payload: SheetSyncRecord[]) => JSON.stringify([
  category.sheetLink?.spreadsheetId,
  category.sheetLink?.sheetId,
  category.name,
  payload,
]);

const getGoogleSheetUrl = (category: Category) => category.sheetLink
  ? `https://docs.google.com/spreadsheets/d/${encodeURIComponent(category.sheetLink.spreadsheetId)}/edit#gid=${category.sheetLink.sheetId}`
  : "";

type CategorySheetStatus = {
  recordCount: number;
  state: "unlinked" | "empty" | "pending" | "synced";
};

export default function Home() {
  const [categories, setCategories] = useState<Category[]>([]);
  const [appliedCategories, setAppliedCategories] = useState<Category[]>([]);
  const [records, setRecords] = useState<ParsedRecord[]>([]);
  
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
    } else {
      setLoginError("Mật khẩu không đúng");
    }
  };

  // Form states for category
  const [newCatKeywords, setNewCatKeywords] = useState("");
  const [editingId, setEditingId] = useState<string | null>(null);
  const [showCategoryModal, setShowCategoryModal] = useState(false);
  const [categoryToDelete, setCategoryToDelete] = useState<Category | null>(null);

  // Tab & UI status states
  const [activeTab, setActiveTab] = useState<string>("all");
  const [isApplying, setIsApplying] = useState(false);
  const [applyMessage, setApplyMessage] = useState("");
  const [copiedTab, setCopiedTab] = useState(false);
  const [copiedAll, setCopiedAll] = useState(false);
  const [linkingCategory, setLinkingCategory] = useState<Category | null>(null);
  const [sheetFiles, setSheetFiles] = useState<GoogleSheetFile[]>([]);
  const [sheetTabs, setSheetTabs] = useState<GoogleSheetTab[]>([]);
  const [selectedSpreadsheetId, setSelectedSpreadsheetId] = useState("");
  const [selectedSheetId, setSelectedSheetId] = useState("");
  const [sheetFileSearch, setSheetFileSearch] = useState("");
  const [isSheetFileDropdownOpen, setIsSheetFileDropdownOpen] = useState(false);
  const [isLoadingSheets, setIsLoadingSheets] = useState(false);
  const [sheetLinkError, setSheetLinkError] = useState("");
  const [syncingCategoryId, setSyncingCategoryId] = useState<string | null>(null);
  const [sheetSyncMessage, setSheetSyncMessage] = useState("");
  const [successfulSheetSyncSignatures, setSuccessfulSheetSyncSignatures] = useState<Record<string, string>>({});

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

  const loadTabs = async (spreadsheetId: string, preferredSheetId = "") => {
    setSelectedSpreadsheetId(spreadsheetId);
    setSelectedSheetId("");
    setSheetTabs([]);
    setSheetLinkError("");
    if (!spreadsheetId) return;

    setIsLoadingSheets(true);
    const result = await getGoogleSheetTabsAction(spreadsheetId);
    if (result.success) {
      setSheetTabs(result.tabs);
      const preferredExists = result.tabs.some((tab) => String(tab.sheetId) === preferredSheetId);
      setSelectedSheetId(preferredExists ? preferredSheetId : String(result.tabs[0]?.sheetId ?? ""));
    } else {
      setSheetLinkError(result.error);
    }
    setIsLoadingSheets(false);
  };

  const handleOpenSheetLink = async (category: Category) => {
    setLinkingCategory(category);
    setSheetFiles([]);
    setSheetTabs([]);
    setSheetLinkError("");
    setSelectedSpreadsheetId(category.sheetLink?.spreadsheetId || "");
    setSelectedSheetId(category.sheetLink ? String(category.sheetLink.sheetId) : "");
    setSheetFileSearch(category.sheetLink?.spreadsheetName || "");
    setIsSheetFileDropdownOpen(false);
    setIsLoadingSheets(true);

    const result = await getGoogleSheetFilesAction();
    if (!result.success) {
      setSheetLinkError(result.error);
      setIsLoadingSheets(false);
      return;
    }
    setSheetFiles(result.files);
    const initialFileId = category.sheetLink?.spreadsheetId || result.files[0]?.id || "";
    const initialFile = result.files.find((file) => file.id === initialFileId);
    setSheetFileSearch(initialFile?.name || "");
    setIsLoadingSheets(false);
    await loadTabs(initialFileId, category.sheetLink ? String(category.sheetLink.sheetId) : "");
  };

  const handleSaveSheetLink = async () => {
    if (!linkingCategory) return;
    const file = sheetFiles.find((item) => item.id === selectedSpreadsheetId);
    const tab = sheetTabs.find((item) => String(item.sheetId) === selectedSheetId);
    if (!file || !tab) {
      setSheetLinkError("Hãy chọn đầy đủ file và tab cần liên kết.");
      return;
    }
    const sheetLink = {
      spreadsheetId: file.id,
      spreadsheetName: file.name,
      sheetId: tab.sheetId,
      sheetName: tab.title,
    };
    const updatedCategories = categories.map((category) => category.id === linkingCategory.id ? {
      ...category,
      sheetLink,
    } : category);
    const updatedAppliedCategories = appliedCategories.map((category) => category.id === linkingCategory.id
      ? { ...category, sheetLink }
      : category);
    setCategories(updatedCategories);
    setAppliedCategories(updatedAppliedCategories);
    await Promise.all([
      saveCategoriesAction(updatedCategories),
      saveAppliedCategoriesAction(updatedAppliedCategories),
    ]);
    setSuccessfulSheetSyncSignatures((current) => {
      const next = { ...current };
      delete next[linkingCategory.id];
      return next;
    });
    setLinkingCategory(null);
  };

  const handleRemoveSheetLink = async (categoryId: string) => {
    const updatedCategories = categories.map((category) => category.id === categoryId
      ? { ...category, sheetLink: undefined }
      : category);
    const updatedAppliedCategories = appliedCategories.map((category) => category.id === categoryId
      ? { ...category, sheetLink: undefined }
      : category);
    setCategories(updatedCategories);
    setAppliedCategories(updatedAppliedCategories);
    await Promise.all([
      saveCategoriesAction(updatedCategories),
      saveAppliedCategoriesAction(updatedAppliedCategories),
    ]);
    setSuccessfulSheetSyncSignatures((current) => {
      const next = { ...current };
      delete next[categoryId];
      return next;
    });
  };

  const handleSyncCategory = async (category: Category) => {
    const payload = buildSheetSyncPayload(computedRecords, category.id);
    if (!category.sheetLink || payload.length === 0) return;
    const syncSignature = createSheetSyncSignature(category, payload);
    setSyncingCategoryId(category.id);
    setSheetSyncMessage("");
    const result = await syncCategoryToSheetAction(category.id, payload);
    if (result.success) {
      setSuccessfulSheetSyncSignatures((current) => ({ ...current, [category.id]: syncSignature }));
      setSheetSyncMessage(`Đã gửi ${result.added} dòng sang ${category.sheetLink.spreadsheetName} / ${category.sheetLink.sheetName}${result.skipped > 0 ? `; bỏ qua ${result.skipped} dòng trùng nội dung và mã tham chiếu` : ""}.`);
    } else {
      setSheetSyncMessage(`Không thể gửi: ${result.error}`);
    }
    setSyncingCategoryId(null);
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
    setShowCategoryModal(false);
  };

  const handleEditCategory = (cat: Category) => {
    setNewCatKeywords(cat.keywords.join(", "));
    setEditingId(cat.id);
    setShowCategoryModal(true);
  };

  const handleOpenAddCategory = () => {
    setEditingId(null);
    setNewCatKeywords("");
    setShowCategoryModal(true);
  };

  const handleCloseCategoryModal = () => {
    setEditingId(null);
    setNewCatKeywords("");
    setShowCategoryModal(false);
  };

  const handleRemoveCategory = (id: string) => {
    const updated = categories.filter((c) => c.id !== id);
    setCategories(updated);
    saveCategoriesAction(updated);

    if (editingId === id) {
      handleCloseCategoryModal();
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

    const isCategoryTab = activeTab !== "all" && activeTab !== "uncategorized";
    navigator.clipboard.writeText(isCategoryTab ? rows : `${headers}\n${rows}`);
    setCopiedTab(true);
    setTimeout(() => setCopiedTab(false), 2000);
  };

  const handleCopyAllData = () => {
    if (computedRecords.length === 0) return;

    const headers = ["STT", ...TABLE_COLUMNS.map(c => c.label)].join("\t");
    
    const rows = computedRecords.filter(r => !r.isFooter).map((record, index) => {
      const rowStr = TABLE_COLUMNS.map(col => {
        let val = record[col.key];
        if (col.key === "ngayGioGiaoDich") val = formatDateStr(String(val || ""));
        else if (col.key === "tienRa" || col.key === "tienVao") {
           const num = cleanNumber(val);
           if (!isNaN(num)) val = num;
        }
        return String(val || "").replace(/\n/g, " ");
      }).join("\t");
      
      return `${index + 1}\t${rowStr}`;
    }).join("\n");

    navigator.clipboard.writeText(`${headers}\n${rows}`);
    setCopiedAll(true);
    setTimeout(() => setCopiedAll(false), 2000);
  };

  // Compute matches
  const computedRecords = classifyRecords(records, appliedCategories);

  const categorySheetStatuses = new Map<string, CategorySheetStatus>(appliedCategories.map((category) => {
    const payload = buildSheetSyncPayload(computedRecords, category.id);
    let state: CategorySheetStatus["state"] = "unlinked";
    if (category.sheetLink && payload.length === 0) {
      state = "empty";
    } else if (category.sheetLink) {
      const signature = createSheetSyncSignature(category, payload);
      state = successfulSheetSyncSignatures[category.id] === signature ? "synced" : "pending";
    }
    return [category.id, { recordCount: payload.length, state }];
  }));

  // Filter records based on active tab
  const filteredRecords = (() => {
    if (activeTab === "all") return computedRecords;
    if (activeTab === "uncategorized")
      return computedRecords.filter((r) => !r.matchedCategoryId);

    return computedRecords.filter((r) => r.matchedCategoryId === activeTab);
  })();

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
      .filter(([, count]) => count > 1) // appear more than once
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

  const filteredSheetFiles = useMemo(() => {
    const query = normalizeSearchText(sheetFileSearch.trim());
    if (!query) return sheetFiles;
    return sheetFiles.filter((file) => normalizeSearchText(file.name).includes(query));
  }, [sheetFiles, sheetFileSearch]);

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

      <main className="mx-auto grid max-w-6xl grid-cols-1 items-start gap-6 lg:grid-cols-10">

        {/* TOP SECTION */}
        <div className="contents">
          {/* File Upload Panel */}
          <section className="order-1 flex flex-col items-center justify-between gap-4 rounded-2xl border border-gray-100 bg-white p-4 shadow-sm sm:flex-row lg:col-span-10">
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
          <section className="order-3 relative flex flex-col overflow-hidden rounded-2xl border border-gray-100 bg-white p-5 shadow-sm lg:sticky lg:top-6 lg:col-span-3 lg:col-start-8 lg:row-start-2">
            <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
              <h2 className="text-xl font-semibold flex items-center gap-2">
                <Filter className="w-5 h-5 text-indigo-600" />
                Cài Đặt Thiện Pháp
              </h2>
              <button
                onClick={handleOpenAddCategory}
                className="flex cursor-pointer items-center gap-2 rounded-lg bg-indigo-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-indigo-700"
              >
                <Plus className="h-4 w-4" />
                Thêm thiện pháp
              </button>
            </div>

            {/* List of draft categories */}
            <div className="mb-4 max-h-[calc(100vh-330px)] min-h-[220px] flex-1 space-y-2 overflow-y-auto pr-1">
              {categories.map((cat) => (
                <div key={cat.id} className="rounded-lg border border-gray-200 bg-white p-2.5 transition hover:border-indigo-200 hover:shadow-sm">
                  <div className="flex items-center justify-between gap-2">
                    <h3 className="min-w-0 truncate text-sm font-semibold text-gray-900" title={cat.name}>{cat.name}</h3>
                    <div className="flex shrink-0 gap-0.5">
                      <button
                        onClick={() => handleEditCategory(cat)}
                        className="cursor-pointer rounded p-1 text-gray-400 transition hover:bg-blue-50 hover:text-blue-600"
                        title="Sửa"
                      >
                        <Pencil className="w-4 h-4" />
                      </button>
                      <button
                        onClick={() => setCategoryToDelete(cat)}
                        className="cursor-pointer rounded p-1 text-gray-400 transition hover:bg-red-50 hover:text-red-600"
                        title="Xóa"
                      >
                        <Trash2 className="w-4 h-4" />
                      </button>
                    </div>
                  </div>
                  <div className="mt-1 flex min-w-0 items-center gap-1 text-[11px] text-gray-500" title={cat.keywords.join(", ")}>
                    <Tag className="h-3 w-3 shrink-0 text-indigo-400" />
                    <span className="truncate">{cat.keywords.join(", ")}</span>
                  </div>
                  <div className="mt-2 flex items-center justify-between gap-2 border-t border-gray-100 pt-2">
                    <div className="min-w-0 text-[11px] text-gray-500">
                      {cat.sheetLink ? (
                        <span className="flex min-w-0 items-center gap-1 text-emerald-700" title={`${cat.sheetLink.spreadsheetName} / ${cat.sheetLink.sheetName}`}>
                          <Link2 className="h-3 w-3 flex-shrink-0" />
                          <span className="truncate">{cat.sheetLink.spreadsheetName} / {cat.sheetLink.sheetName}</span>
                        </span>
                      ) : <span className="text-gray-400">Chưa liên kết</span>}
                    </div>
                    <div className="flex flex-shrink-0 gap-0.5">
                      {cat.sheetLink && (
                        <button
                          onClick={() => handleRemoveSheetLink(cat.id)}
                          className="cursor-pointer rounded p-1 text-gray-400 transition hover:bg-red-50 hover:text-red-600"
                          title="Gỡ liên kết"
                        >
                          <Unlink className="w-3.5 h-3.5" />
                        </button>
                      )}
                      <button
                        onClick={() => handleOpenSheetLink(cat)}
                        className="cursor-pointer rounded-md bg-indigo-50 px-2 py-1 text-[11px] font-medium text-indigo-700 transition hover:bg-indigo-100"
                      >
                        {cat.sheetLink ? "Đổi liên kết" : "Liên kết file"}
                      </button>
                    </div>
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
        <div className="order-2 min-w-0 lg:col-span-7 lg:col-start-1 lg:row-start-2">
          <section className="relative flex min-h-[500px] flex-col overflow-hidden rounded-2xl border border-gray-100 bg-white shadow-sm">

            {/* Result filter */}
            <div className="w-full flex-shrink-0 border-b border-gray-200 bg-gray-50">
              <div className="p-3">
                <p className="mb-2 text-xs font-semibold uppercase tracking-wide text-gray-500">Tình trạng gửi theo thiện pháp</p>
                <div className="grid grid-cols-2 gap-2 md:grid-cols-3">
                  <button
                    onClick={() => setActiveTab("all")}
                    aria-current={activeTab === "all" ? "page" : undefined}
                    className={`min-w-0 cursor-pointer rounded-lg border px-3 py-2 text-left transition ${activeTab === "all" ? "border-indigo-600 bg-indigo-600 text-white shadow-sm ring-2 ring-indigo-200" : "border-gray-200 bg-white text-gray-700 hover:border-indigo-300"}`}
                  >
                    <span className="block truncate text-sm font-semibold">Tất cả</span>
                    <span className={`text-[11px] ${activeTab === "all" ? "text-indigo-100" : "text-gray-500"}`}>
                      {computedRecords.filter((record) => !record.isFooter).length} giao dịch
                    </span>
                  </button>

                  {appliedCategories.map((category) => {
                    const status = categorySheetStatuses.get(category.id) || { recordCount: 0, state: "unlinked" as const };
                    const isActive = activeTab === category.id;
                    const statusColor = status.state === "synced"
                      ? "border-emerald-300 bg-emerald-100 text-emerald-800"
                      : status.state === "pending"
                        ? "border-red-300 bg-red-100 text-red-800"
                        : status.state === "empty"
                          ? "border-emerald-200 bg-white text-gray-700"
                          : "border-gray-200 bg-white text-gray-600";

                    return (
                      <button
                        key={category.id}
                        onClick={() => setActiveTab(category.id)}
                        aria-current={isActive ? "page" : undefined}
                        className={`min-w-0 cursor-pointer rounded-lg border px-3 py-2 text-left transition hover:shadow-sm ${statusColor} ${isActive ? "ring-2 ring-indigo-500 ring-offset-1" : ""}`}
                      >
                        <span className="flex items-center justify-between gap-2">
                          <span className="truncate text-sm font-semibold" title={category.name}>{category.name}</span>
                          <span className="shrink-0 text-xs font-bold">{status.recordCount}</span>
                        </span>
                        {status.state === "synced" && (
                          <span className="mt-0.5 flex items-center gap-1 text-[11px] font-semibold text-emerald-700">
                            <CheckCircle className="h-3 w-3" /> Đã gửi
                          </span>
                        )}
                        {status.state === "pending" && (
                          <span className="mt-0.5 flex items-center gap-1 text-[11px] font-semibold text-red-700">
                            <CircleAlert className="h-3 w-3" /> Chưa gửi
                          </span>
                        )}
                        {status.state === "empty" && (
                          <span className="mt-0.5 flex items-center gap-1 text-[11px] text-emerald-700">
                            <Link2 className="h-3 w-3" /> Đã liên kết · Chưa có GD
                          </span>
                        )}
                        {status.state === "unlinked" && (
                          <span className="mt-0.5 flex items-center gap-1 text-[11px] text-gray-400">
                            <Unlink className="h-3 w-3" /> Chưa liên kết
                          </span>
                        )}
                      </button>
                    );
                  })}

                  <button
                    onClick={() => setActiveTab("uncategorized")}
                    aria-current={activeTab === "uncategorized" ? "page" : undefined}
                    className={`min-w-0 cursor-pointer rounded-lg border px-3 py-2 text-left transition ${activeTab === "uncategorized" ? "border-indigo-600 bg-indigo-600 text-white shadow-sm ring-2 ring-indigo-200" : "border-gray-200 bg-white text-gray-700 hover:border-indigo-300"}`}
                  >
                    <span className="block truncate text-sm font-semibold">Chưa phân loại</span>
                    <span className={`text-[11px] ${activeTab === "uncategorized" ? "text-indigo-100" : "text-gray-500"}`}>
                      {computedRecords.filter((record) => !record.matchedCategoryId && !record.isFooter).length} giao dịch
                    </span>
                  </button>
                </div>
              </div>
            </div>

            {/* Table Area */}
            <div className="flex-1 flex flex-col min-w-0 bg-white">
              {/* Header Copy Button */}
              {(filteredRecords.length > 0 || (activeTab !== "all" && activeTab !== "uncategorized")) && (
                <div className="p-2 border-b border-gray-100 flex flex-wrap items-center justify-end gap-2 bg-white z-10 sticky top-0">
                  {activeTab !== "all" && activeTab !== "uncategorized" && (() => {
                    const category = appliedCategories.find((item) => item.id === activeTab);
                    if (!category) return null;
                    const categorySheetStatus = categorySheetStatuses.get(category.id);
                    const isSynced = categorySheetStatus?.state === "synced";
                    const transactionCount = categorySheetStatus?.recordCount || 0;
                    return (
                      <div className="mr-auto flex flex-wrap items-center gap-2">
                        <button
                          onClick={() => handleOpenSheetLink(category)}
                          className="flex items-center gap-1.5 rounded-lg border border-indigo-200 bg-indigo-50 px-3 py-1.5 text-sm font-medium text-indigo-700 transition hover:bg-indigo-100 cursor-pointer"
                        >
                          <Link2 className="h-4 w-4" />
                          {category.sheetLink ? "Đổi liên kết" : "Liên kết Google Sheet"}
                        </button>
                        <button
                          onClick={() => handleSyncCategory(category)}
                          disabled={!category.sheetLink || syncingCategoryId === category.id}
                          title={category.sheetLink ? `Gửi sang ${category.sheetLink.spreadsheetName} / ${category.sheetLink.sheetName}` : "Hãy liên kết Google Sheet trước khi gửi"}
                          className="flex justify-center items-center cursor-pointer gap-2 px-3 py-1.5 text-sm font-medium rounded-lg bg-emerald-600 text-white hover:bg-emerald-700 disabled:bg-gray-200 disabled:text-gray-500 disabled:cursor-not-allowed transition shadow-sm"
                        >
                          {syncingCategoryId === category.id
                            ? <Loader2 className="w-4 h-4 animate-spin" />
                            : isSynced
                              ? <CheckCircle className="w-4 h-4" />
                              : <Send className="w-4 h-4" />}
                          {syncingCategoryId === category.id
                            ? `Đang gửi ${transactionCount} giao dịch...`
                            : isSynced
                              ? `Đã gửi ${transactionCount} giao dịch · Gửi lại`
                              : `Gửi sang Google Sheet (${transactionCount} giao dịch)`}
                        </button>
                        {category.sheetLink && (
                          <a
                            href={getGoogleSheetUrl(category)}
                            target="_blank"
                            rel="noopener noreferrer"
                            title={`Mở ${category.sheetLink.spreadsheetName} / ${category.sheetLink.sheetName}`}
                            className="flex items-center gap-1.5 rounded-lg border border-emerald-200 bg-emerald-50 px-3 py-1.5 text-sm font-medium text-emerald-700 transition hover:bg-emerald-100"
                          >
                            <ExternalLink className="h-4 w-4" />
                            Mở Google Sheet
                          </a>
                        )}
                      </div>
                    );
                  })()}
                  {filteredRecords.length > 0 && (
                    <>
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
                    </>
                  )}
                </div>
              )}

              {sheetSyncMessage && (
                <div className={`mx-3 mt-3 px-3 py-2 rounded-lg border text-sm ${sheetSyncMessage.startsWith("Đã gửi") ? "bg-emerald-50 border-emerald-200 text-emerald-700" : "bg-red-50 border-red-200 text-red-700"}`}>
                  {sheetSyncMessage}
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

                            const styleObj: CSSProperties = {
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

      {/* CATEGORY ADD / EDIT MODAL */}
      {showCategoryModal && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4"
          onMouseDown={(event) => {
            if (event.target === event.currentTarget) handleCloseCategoryModal();
          }}
        >
          <div
            role="dialog"
            aria-modal="true"
            aria-labelledby="category-modal-title"
            className="flex max-h-[90vh] w-full max-w-lg flex-col overflow-hidden rounded-2xl bg-white shadow-xl"
          >
            <div className="flex items-center justify-between border-b border-gray-100 bg-gray-50 p-4">
              <div>
                <h2 id="category-modal-title" className="text-lg font-bold text-gray-900">
                  {editingId ? "Chỉnh sửa thiện pháp" : "Thêm thiện pháp"}
                </h2>
                <p className="mt-1 text-xs text-gray-500">
                  Tên thiện pháp được lấy từ từ khóa đầu tiên.
                </p>
              </div>
              <button
                type="button"
                onClick={handleCloseCategoryModal}
                aria-label="Đóng"
                className="cursor-pointer rounded-lg p-1.5 text-gray-400 transition hover:bg-gray-200 hover:text-gray-600"
              >
                <X className="h-5 w-5" />
              </button>
            </div>

            <form
              onSubmit={(event) => {
                event.preventDefault();
                handleAddOrEditCategory();
              }}
              className="flex min-h-0 flex-1 flex-col"
            >
              <div className="flex-1 overflow-y-auto p-5">
                <label htmlFor="category-keywords" className="mb-1.5 block text-sm font-medium text-gray-700">
                  Danh sách từ khóa
                </label>
                <input
                  id="category-keywords"
                  type="text"
                  value={newCatKeywords}
                  onChange={(event) => setNewCatKeywords(event.target.value)}
                  placeholder="tp44, hoa sen, ..."
                  autoFocus
                  className="w-full rounded-lg border border-gray-300 px-3 py-2.5 text-sm outline-none focus:border-indigo-500 focus:ring-2 focus:ring-indigo-500"
                />
                <p className="mt-1.5 text-xs text-gray-500">Các từ khóa cách nhau bằng dấu phẩy.</p>

                {suggestedKeywords.length > 0 && (
                  <div className="mt-4 rounded-lg border border-amber-100 bg-amber-50/60 p-3">
                    <span className="mb-2 flex items-center gap-1 text-xs font-medium text-amber-700">
                      <Sparkles className="h-3.5 w-3.5" />
                      Gợi ý từ khóa thông minh — bấm để thêm
                    </span>
                    <div className="flex flex-wrap gap-1.5">
                      {suggestedKeywords.map((keyword, index) => (
                        <button
                          key={index}
                          type="button"
                          onClick={() => {
                            const current = newCatKeywords.trim();
                            setNewCatKeywords(current ? `${current}, ${keyword}` : keyword);
                          }}
                          className="cursor-pointer rounded-md bg-white px-2 py-1 text-[11px] font-medium text-amber-700 shadow-sm transition hover:bg-amber-100"
                        >
                          {keyword}
                        </button>
                      ))}
                    </div>
                  </div>
                )}
              </div>

              <div className="flex justify-end gap-2 border-t border-gray-100 bg-gray-50 p-4">
                <button
                  type="button"
                  onClick={handleCloseCategoryModal}
                  className="cursor-pointer rounded-lg px-4 py-2 text-sm font-medium text-gray-600 transition hover:bg-gray-200"
                >
                  Hủy
                </button>
                <button
                  type="submit"
                  disabled={!newCatKeywords.trim()}
                  className="flex cursor-pointer items-center gap-2 rounded-lg bg-indigo-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-indigo-700 disabled:cursor-not-allowed disabled:bg-gray-300"
                >
                  {editingId ? <Save className="h-4 w-4" /> : <Plus className="h-4 w-4" />}
                  {editingId ? "Lưu chỉnh sửa" : "Thêm thiện pháp"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* CATEGORY DELETE CONFIRMATION MODAL */}
      {categoryToDelete && (
        <div
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/50 p-4"
          onMouseDown={(event) => {
            if (event.target === event.currentTarget) setCategoryToDelete(null);
          }}
        >
          <div
            role="alertdialog"
            aria-modal="true"
            aria-labelledby="delete-category-title"
            aria-describedby="delete-category-description"
            className="w-full max-w-md overflow-hidden rounded-2xl bg-white shadow-xl"
          >
            <div className="flex items-start justify-between border-b border-gray-100 p-5">
              <div>
                <h2 id="delete-category-title" className="text-lg font-bold text-gray-900">
                  Xóa thiện pháp?
                </h2>
                <p id="delete-category-description" className="mt-2 text-sm leading-relaxed text-gray-600">
                  Bạn có chắc muốn xóa thiện pháp <strong className="text-gray-900">{categoryToDelete.name}</strong>? Thao tác này sẽ xóa cấu hình từ khóa và liên kết Google Sheet của thiện pháp.
                </p>
              </div>
              <button
                type="button"
                onClick={() => setCategoryToDelete(null)}
                aria-label="Đóng"
                className="ml-4 shrink-0 cursor-pointer rounded-lg p-1.5 text-gray-400 transition hover:bg-gray-100 hover:text-gray-600"
              >
                <X className="h-5 w-5" />
              </button>
            </div>
            <div className="flex justify-end gap-2 bg-gray-50 p-4">
              <button
                type="button"
                onClick={() => setCategoryToDelete(null)}
                className="cursor-pointer rounded-lg px-4 py-2 text-sm font-medium text-gray-600 transition hover:bg-gray-200"
              >
                Hủy
              </button>
              <button
                type="button"
                onClick={() => {
                  handleRemoveCategory(categoryToDelete.id);
                  setCategoryToDelete(null);
                }}
                autoFocus
                className="flex cursor-pointer items-center gap-2 rounded-lg bg-red-600 px-4 py-2 text-sm font-semibold text-white transition hover:bg-red-700"
              >
                <Trash2 className="h-4 w-4" />
                Xóa thiện pháp
              </button>
            </div>
          </div>
        </div>
      )}

      {/* GOOGLE SHEET LINK MODAL */}
      {linkingCategory && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl w-full max-w-lg shadow-xl overflow-hidden">
            <div className="p-4 border-b border-gray-100 flex justify-between items-center bg-gray-50">
              <div>
                <h2 className="text-lg font-bold flex items-center gap-2">
                  <Link2 className="w-5 h-5 text-indigo-600" />
                  Liên kết Google Sheet
                </h2>
                <p className="text-xs text-gray-500 mt-1">Thiện pháp: {linkingCategory.name}</p>
              </div>
              <button onClick={() => setLinkingCategory(null)} className="text-gray-400 hover:text-gray-600 p-1 cursor-pointer">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-5 space-y-4">
              {sheetLinkError && (
                <div className="bg-red-50 border border-red-200 text-red-700 text-sm rounded-lg p-3">{sheetLinkError}</div>
              )}
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1.5">File Google Sheet trong folder được chia sẻ</label>
                <div
                  className="relative"
                  onBlur={(event) => {
                    if (!event.currentTarget.contains(event.relatedTarget as Node | null)) {
                      setIsSheetFileDropdownOpen(false);
                    }
                  }}
                >
                  <Search className="pointer-events-none absolute left-3 top-1/2 z-10 h-4 w-4 -translate-y-1/2 text-gray-400" />
                  <input
                    id="google-sheet-file-search"
                    role="combobox"
                    aria-expanded={isSheetFileDropdownOpen}
                    aria-controls="google-sheet-file-options"
                    aria-autocomplete="list"
                    value={sheetFileSearch}
                    onFocus={() => setIsSheetFileDropdownOpen(true)}
                    onChange={(event) => {
                      const nextSearch = event.target.value;
                      setSheetFileSearch(nextSearch);
                      setIsSheetFileDropdownOpen(true);

                      const selectedFile = sheetFiles.find((file) => file.id === selectedSpreadsheetId);
                      if (!selectedFile || nextSearch !== selectedFile.name) {
                        setSelectedSpreadsheetId("");
                        setSelectedSheetId("");
                        setSheetTabs([]);
                      }
                    }}
                    disabled={isLoadingSheets && sheetFiles.length === 0}
                    placeholder={isLoadingSheets && sheetFiles.length === 0 ? "Đang tải danh sách file..." : "Gõ để tìm file..."}
                    autoComplete="off"
                    className="w-full rounded-lg border border-gray-300 bg-white py-2.5 pl-9 pr-10 text-sm outline-none focus:ring-2 focus:ring-indigo-500 disabled:bg-gray-100"
                  />
                  <button
                    type="button"
                    aria-label="Mở danh sách file Google Sheet"
                    onClick={() => setIsSheetFileDropdownOpen((isOpen) => !isOpen)}
                    disabled={isLoadingSheets && sheetFiles.length === 0}
                    className="absolute right-0 top-0 flex h-full w-10 cursor-pointer items-center justify-center text-gray-400 disabled:cursor-not-allowed"
                  >
                    <ChevronDown className={`h-4 w-4 transition-transform ${isSheetFileDropdownOpen ? "rotate-180" : ""}`} />
                  </button>

                  {isSheetFileDropdownOpen && !(isLoadingSheets && sheetFiles.length === 0) && (
                    <div
                      id="google-sheet-file-options"
                      role="listbox"
                      className="absolute z-40 mt-1 max-h-60 w-full overflow-y-auto rounded-lg border border-gray-200 bg-white p-1 shadow-lg"
                    >
                      {filteredSheetFiles.length > 0 ? filteredSheetFiles.map((file) => (
                        <button
                          key={file.id}
                          type="button"
                          role="option"
                          aria-selected={file.id === selectedSpreadsheetId}
                          onClick={() => {
                            setSheetFileSearch(file.name);
                            setIsSheetFileDropdownOpen(false);
                            void loadTabs(file.id);
                          }}
                          className={`flex w-full cursor-pointer items-center justify-between rounded-md px-3 py-2 text-left text-sm hover:bg-indigo-50 ${file.id === selectedSpreadsheetId ? "bg-indigo-50 font-medium text-indigo-700" : "text-gray-700"}`}
                        >
                          <span className="truncate">{file.name}</span>
                          {file.id === selectedSpreadsheetId && <Check className="ml-2 h-4 w-4 shrink-0" />}
                        </button>
                      )) : (
                        <p className="px-3 py-3 text-center text-sm text-gray-500">Không tìm thấy file phù hợp</p>
                      )}
                    </div>
                  )}
                </div>
              </div>
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-1.5">Tab nhận giao dịch</label>
                <select
                  value={selectedSheetId}
                  onChange={(event) => setSelectedSheetId(event.target.value)}
                  disabled={!selectedSpreadsheetId || isLoadingSheets}
                  className="w-full border border-gray-300 rounded-lg px-3 py-2.5 text-sm bg-white focus:ring-2 focus:ring-indigo-500 outline-none disabled:bg-gray-100"
                >
                  <option value="">{isLoadingSheets ? "Đang tải các tab..." : "Chọn tab"}</option>
                  {sheetTabs.filter((tab) => !tab.hidden).map((tab) => <option key={tab.sheetId} value={tab.sheetId}>{tab.title}</option>)}
                </select>
              </div>
              <div className="bg-amber-50 border border-amber-200 text-amber-800 text-xs rounded-lg p-3 leading-relaxed">
                Dữ liệu luôn được nối tiếp ở cuối sheet theo thứ tự cột: STT · Thời gian · Chủ tài khoản · Nội dung · Tiền ra · Tiền vào · Ghi chú · Quỹ (để trống). Các dòng đã có sẽ không bị ghi đè.
              </div>
            </div>
            <div className="p-4 border-t border-gray-100 flex justify-end gap-2 bg-gray-50">
              <button onClick={() => setLinkingCategory(null)} className="px-4 py-2 text-sm font-medium text-gray-600 hover:bg-gray-200 rounded-lg cursor-pointer">Hủy</button>
              <button
                onClick={handleSaveSheetLink}
                disabled={!selectedSpreadsheetId || !selectedSheetId || isLoadingSheets}
                className="px-4 py-2 text-sm font-semibold text-white bg-indigo-600 hover:bg-indigo-700 disabled:bg-gray-300 rounded-lg cursor-pointer disabled:cursor-not-allowed"
              >
                Lưu liên kết
              </button>
            </div>
          </div>
        </div>
      )}

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
