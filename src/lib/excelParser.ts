import * as xlsx from 'xlsx';
import { ParsedRecord } from './types';

function parseVietnameseDate(dateStr: string): number {
  if (!dateStr) return Number.MAX_SAFE_INTEGER;
  
  const parts = dateStr.match(/\d+/g);
  if (!parts || parts.length < 3) {
    const fallback = new Date(dateStr).getTime();
    return isNaN(fallback) ? Number.MAX_SAFE_INTEGER : fallback;
  }
  
  let day = parseInt(parts[0], 10);
  let month = parseInt(parts[1], 10) - 1;
  let year = parseInt(parts[2], 10);
  
  if (parts[0].length === 4) {
     year = parseInt(parts[0], 10);
     month = parseInt(parts[1], 10) - 1;
     day = parseInt(parts[2], 10);
  } else if (year < 100) {
     year += 2000;
  }
  
  const hours = parts[3] ? parseInt(parts[3], 10) : 0;
  const minutes = parts[4] ? parseInt(parts[4], 10) : 0;
  const seconds = parts[5] ? parseInt(parts[5], 10) : 0;
  
  return new Date(year, month, day, hours, minutes, seconds).getTime();
}

export async function parseExcelBankStatement(file: File): Promise<ParsedRecord[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = xlsx.read(data, { type: 'array' });
        
        // Grab the first sheet
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Convert to JSON using array format to easily map standard columns
        const rawJsonData = xlsx.utils.sheet_to_json<unknown[]>(worksheet, { header: 1, defval: "", raw: false, dateNF: 'dd-mm-yyyy hh:mm' });
        
        // Skip first 10 rows (header/metadata)
        const dataRows = rawJsonData.slice(10);
        
        const records: ParsedRecord[] = dataRows
          .map((row) => {
            const chiTietGiaoDich = String(row[6] || "").trim();
            const soThamChieu = String(row[0] || "").trim();
            const ngayGioGiaoDich = String(row[1] || "").trim();
            
            const firstColsText = (soThamChieu + " " + ngayGioGiaoDich + " " + chiTietGiaoDich).toLowerCase();
            const isFooter = firstColsText.includes("tổng") || firstColsText.includes("total") || (!ngayGioGiaoDich || !/\d/.test(ngayGioGiaoDich));

            return {
              soThamChieu,
              ngayGioGiaoDich,
              loaiGiaoDich: String(row[2] || "").trim(),
              nganHang: String(row[3] || "").trim(),
              soTaiKhoan: String(row[4] || "").trim(),
              tenChuTaiKhoan: String(row[5] || "").trim(),
              chiTietGiaoDich,
              tienRa: typeof row[7] === "number" ? row[7] : String(row[7] || ""),
              tienVao: typeof row[8] === "number" ? row[8] : String(row[8] || ""),
              ghiChu: String(row[9] || "").trim(),
              searchText: chiTietGiaoDich,
              isFooter
            };
          })
          .filter((r) => {
            if (!r.ngayGioGiaoDich && !r.chiTietGiaoDich && !r.tienRa && !r.tienVao) return false;
            // Cho phép hiển thị footer mang thông tin tiền
            if (r.isFooter && !r.tienRa && !r.tienVao && !r.chiTietGiaoDich) return false;
            return true;
          })
          .sort((a, b) => parseVietnameseDate(a.ngayGioGiaoDich) - parseVietnameseDate(b.ngayGioGiaoDich));

        resolve(records);
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = (error) => {
      reject(error);
    };

    reader.readAsArrayBuffer(file);
  });
}
