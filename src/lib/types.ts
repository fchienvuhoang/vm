export interface Category {
  id: string;
  name: string;
  keywords: string[];
  sheetLink?: SheetLink;
}

export interface SheetLink {
  spreadsheetId: string;
  spreadsheetName: string;
  sheetId: number;
  sheetName: string;
}

export interface GoogleSheetFile {
  id: string;
  name: string;
  modifiedTime?: string;
}

export interface GoogleSheetTab {
  sheetId: number;
  title: string;
  hidden: boolean;
}

export interface SheetSyncRecord {
  soThamChieu: string;
  nganHang: string;
  soTaiKhoan: string;
  ngayGioGiaoDich: string;
  tenChuTaiKhoan: string;
  chiTietGiaoDich: string;
  tienRa: string | number;
  tienVao: string | number;
  ghiChu: string;
}

export interface ParsedRecord {
  soThamChieu: string;
  ngayGioGiaoDich: string;
  loaiGiaoDich: string;
  nganHang: string;
  soTaiKhoan: string;
  tenChuTaiKhoan: string;
  chiTietGiaoDich: string;
  tienRa: string | number;
  tienVao: string | number;
  ghiChu: string;

  searchText: string;
  matchedCategoryId?: string;
  isFooter?: boolean;
}

export interface UploadLogRecord {
  soThamChieu: string;
  ngayGioGiaoDich: string;
  tenChuTaiKhoan: string;
  chiTietGiaoDich: string;
  tienRa: string | number;
  tienVao: string | number;
  categoryName: string;
}

export interface UploadLog {
  id?: string;
  time: string;
  rowCount: number;
  records?: UploadLogRecord[];
}
