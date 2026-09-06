"use server";

import { Redis } from "@upstash/redis";
import { cookies } from "next/headers";
import { Category, UploadLogRecord } from "@/lib/types";

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
