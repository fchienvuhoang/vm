"use server";

import { Redis } from "@upstash/redis";
import { cookies } from "next/headers";
import { Category, GoogleSheetFile, GoogleSheetTab, SheetSyncRecord, UploadLogRecord } from "@/lib/types";
import { getGoogleClients, syncRecordsToSheet } from "@/lib/googleSheets";

// Khởi tạo Redis client
const redis = new Redis({
  url: process.env.KV_REST_API_URL || "",
  token: process.env.KV_REST_API_TOKEN || "",
});

const AUTH_COOKIE = "vimutti_auth";

// -- XÁC THỰC MẬT KHẨU --
export async function checkPasswordAction(password: string) {
  const correctPassword = process.env.APP_PASSWORD || "123";
  if (password === correctPassword) {
    // Lưu session qua cookie HTTP-Only
    (await cookies()).set(AUTH_COOKIE, "true", {
      httpOnly: true,
      secure: process.env.NODE_ENV === "production",
      maxAge: 30 * 24 * 60 * 60, // 30 ngày
      path: "/",
    });
    return true;
  }
  return false;
}

export async function checkAuthStatus() {
  const cookieStore = await cookies();
  return cookieStore.get(AUTH_COOKIE)?.value === "true";
}

export async function logoutAction() {
  (await cookies()).delete(AUTH_COOKIE);
}

// -- QUẢN LÝ TỪ KHÓA --
export async function getCategoriesAction() {
  try {
    const categories = await redis.get<Category[]>("vimutti_categories");
    const appliedCategories = await redis.get<Category[]>(
      "vimutti_appliedCategories"
    );
    return {
      categories: categories || [],
      appliedCategories: appliedCategories || [],
    };
  } catch (error) {
    console.error("Lỗi khi tải từ khóa từ Redis:", error);
    return { categories: [], appliedCategories: [] };
  }
}

export async function saveCategoriesAction(categories: Category[]) {
  // Kiểm tra quyền
  if (!(await checkAuthStatus())) return { success: false, error: "Unauthorized" };

  try {
    await redis.set("vimutti_categories", categories);
    return { success: true };
  } catch (error) {
    console.error("Lỗi khi lưu categories vào Redis:", error);
    return { success: false, error: "Không thể lưu vào CSDL" };
  }
}

export async function saveAppliedCategoriesAction(categories: Category[]) {
  if (!(await checkAuthStatus())) return { success: false, error: "Unauthorized" };

  try {
    await redis.set("vimutti_appliedCategories", categories);
    return { success: true };
  } catch (error) {
    console.error("Lỗi khi lưu applied categories vào Redis:", error);
    return { success: false, error: "Không thể lưu vào CSDL" };
  }
}

// -- KẾT NỐI GOOGLE DRIVE / SHEETS --
export async function getGoogleSheetFilesAction() {
  if (!(await checkAuthStatus())) return { success: false as const, error: "Unauthorized", files: [] as GoogleSheetFile[] };

  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
  if (!folderId) {
    return { success: false as const, error: "Chưa cấu hình GOOGLE_DRIVE_FOLDER_ID.", files: [] as GoogleSheetFile[] };
  }

  try {
    const { drive } = getGoogleClients();
    const response = await drive.files.list({
      q: `'${folderId.replace(/'/g, "\\'")}' in parents and trashed = false and mimeType = 'application/vnd.google-apps.spreadsheet'`,
      fields: "files(id,name,modifiedTime)",
      orderBy: "name",
      pageSize: 1000,
      supportsAllDrives: true,
      includeItemsFromAllDrives: true,
    });
    const files = (response.data.files || [])
      .filter((file): file is typeof file & { id: string; name: string } => Boolean(file.id && file.name))
      .map((file) => ({ id: file.id, name: file.name, modifiedTime: file.modifiedTime || undefined }));
    return { success: true as const, files };
  } catch (error) {
    console.error("Lỗi khi tải danh sách Google Sheets:", error);
    return { success: false as const, error: error instanceof Error ? error.message : "Không thể đọc folder Google Drive.", files: [] as GoogleSheetFile[] };
  }
}

export async function getGoogleSheetTabsAction(spreadsheetId: string) {
  if (!(await checkAuthStatus())) return { success: false as const, error: "Unauthorized", tabs: [] as GoogleSheetTab[] };

  try {
    const { sheets } = getGoogleClients();
    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      fields: "sheets(properties(sheetId,title,hidden))",
    });
    const tabs = (response.data.sheets || [])
      .map((sheet) => sheet.properties)
      .filter((properties): properties is NonNullable<typeof properties> & { sheetId: number; title: string } => (
        properties?.sheetId !== undefined && Boolean(properties.title)
      ))
      .filter((properties) => properties.title !== "_VIMUTTI_SYNC")
      .map((properties) => ({ sheetId: properties.sheetId, title: properties.title, hidden: Boolean(properties.hidden) }));
    return { success: true as const, tabs };
  } catch (error) {
    console.error("Lỗi khi tải danh sách tab:", error);
    return { success: false as const, error: error instanceof Error ? error.message : "Không thể đọc các tab trong file.", tabs: [] as GoogleSheetTab[] };
  }
}

export async function syncCategoryToSheetAction(categoryId: string, records: SheetSyncRecord[]) {
  if (!(await checkAuthStatus())) return { success: false as const, error: "Unauthorized" };
  if (!Array.isArray(records) || records.length > 5000) return { success: false as const, error: "Dữ liệu gửi lên không hợp lệ hoặc vượt quá 5.000 dòng." };

  try {
    const appliedCategories = await redis.get<Category[]>("vimutti_appliedCategories") || [];
    const category = appliedCategories.find((item) => item.id === categoryId);
    if (!category?.sheetLink) return { success: false as const, error: "Thiện pháp này chưa được liên kết với Google Sheet." };
    return await syncRecordsToSheet(category.sheetLink, category.name, records);
  } catch (error) {
    console.error("Lỗi khi đồng bộ Google Sheet:", error);
    return { success: false as const, error: error instanceof Error ? error.message : "Không thể đồng bộ dữ liệu sang Google Sheet." };
  }
}

// -- THEO DÕI SỬ DỤNG LƯỢNG LƯU TRỮ --
export async function logFileUploadAction(rowCount: number, records: UploadLogRecord[]) {
  // Không cần xác thực cũng có thể ghi log ẩn danh, hoặc cần xác thực
  if (!(await checkAuthStatus())) return { success: false };

  try {
    const time = new Date().toISOString();
    const safeRecords = records.slice(0, Math.max(0, rowCount)).map((record) => ({
      soThamChieu: String(record.soThamChieu || ""),
      ngayGioGiaoDich: String(record.ngayGioGiaoDich || ""),
      tenChuTaiKhoan: String(record.tenChuTaiKhoan || ""),
      chiTietGiaoDich: String(record.chiTietGiaoDich || ""),
      tienRa: typeof record.tienRa === "number" ? record.tienRa : String(record.tienRa || ""),
      tienVao: typeof record.tienVao === "number" ? record.tienVao : String(record.tienVao || ""),
      categoryName: String(record.categoryName || "Chưa phân loại"),
    }));

    // Lưu ảnh chụp kết quả phân loại để có thể xem lại từ lịch sử.
    await redis.lpush("vimutti_upload_logs", {
      id: crypto.randomUUID(),
      time,
      rowCount,
      records: safeRecords,
    });
    await redis.ltrim("vimutti_upload_logs", 0, 99);
    
    // Tăng tổng số file đã xử lý
    await redis.incr("vimutti_total_uploads");
    
    return { success: true };
  } catch (error) {
    console.error("Lỗi khi lưu log upload:", error);
    return { success: false };
  }
}

export async function getUploadStatsAction() {
  if (!(await checkAuthStatus())) return { total: 0, logs: [] };

  try {
    const total = await redis.get<number>("vimutti_total_uploads") || 0;
    const logs = await redis.lrange<{ time: string; rowCount: number }>("vimutti_upload_logs", 0, -1) || [];
    return { total, logs };
  } catch (error) {
    console.error("Lỗi khi lấy log upload:", error);
    return { total: 0, logs: [] };
  }
}
