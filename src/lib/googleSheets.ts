import "server-only";

import { google } from "googleapis";
import { SheetLink, SheetSyncRecord } from "@/lib/types";

function getGoogleAuth() {
  const email = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL;
  const privateKey = process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n");
  if (!email || !privateKey) {
    throw new Error("Chưa cấu hình GOOGLE_SERVICE_ACCOUNT_EMAIL và GOOGLE_PRIVATE_KEY.");
  }

  return new google.auth.JWT({
    email,
    key: privateKey,
    scopes: [
      "https://www.googleapis.com/auth/drive.readonly",
      "https://www.googleapis.com/auth/spreadsheets",
    ],
  });
}

export function getGoogleClients() {
  const auth = getGoogleAuth();
  return {
    drive: google.drive({ version: "v3", auth }),
    sheets: google.sheets({ version: "v4", auth }),
  };
}

export function buildTransactionContent(record: SheetSyncRecord) {
  const content = String(record.chiTietGiaoDich || "").trim();
  const reference = String(record.soThamChieu || "").trim();
  if (!reference) return content;
  return content ? `${content}\nSố tham chiếu: ${reference}` : `Số tham chiếu: ${reference}`;
}

function normalizeTransactionContent(value: unknown) {
  return String(value ?? "")
    .trim()
    .replace(/\r\n?/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n+/g, "\n")
    .toLowerCase();
}

function quoteSheetName(name: string) {
  return `'${name.replace(/'/g, "''")}'`;
}

export async function syncRecordsToSheet(link: SheetLink, categoryName: string, records: SheetSyncRecord[]) {
  const { sheets } = getGoogleClients();
  const target = quoteSheetName(link.sheetName);

  if (records.length === 0) {
    return { success: true as const, added: 0, skipped: 0 };
  }

  // Read all destination columns so the new rows always start below every
  // existing value, even when column A contains gaps.
  const targetResponse = await sheets.spreadsheets.values.get({
    spreadsheetId: link.spreadsheetId,
    range: `${target}!A:H`,
  });
  const currentValues = targetResponse.data.values || [];
  const nextRowNumber = currentValues.length + 1;
  const existingContents = new Set(
    currentValues
      .map((row) => normalizeTransactionContent(row[3]))
      .filter(Boolean),
  );
  const uniqueRecords = records.filter((record) => {
    const contentKey = normalizeTransactionContent(buildTransactionContent(record));
    if (!contentKey || existingContents.has(contentKey)) return false;
    existingContents.add(contentKey);
    return true;
  });

  if (uniqueRecords.length === 0) {
    return { success: true as const, added: 0, skipped: records.length };
  }

  const lastSequenceNumber = currentValues.reduce((largest, row) => {
    const candidate = Number(row[0]);
    return Number.isInteger(candidate) && candidate > largest ? candidate : largest;
  }, 0);
  const values = uniqueRecords.map((record, index) => [
    lastSequenceNumber + index + 1,
    record.ngayGioGiaoDich,
    record.tenChuTaiKhoan,
    buildTransactionContent(record),
    record.tienRa,
    record.tienVao,
    record.ghiChu,
    categoryName,
  ]);

  // INSERT_ROWS appends new physical rows and never overwrites existing cell
  // values. Starting below the last used row also avoids Google selecting an
  // earlier logical table when the sheet contains blank rows.
  const appendResponse = await sheets.spreadsheets.values.append({
    spreadsheetId: link.spreadsheetId,
    range: `${target}!A${nextRowNumber}:H`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values },
  });
  const appendedRange = appendResponse.data.updates?.updatedRange || "";
  const appendedRows = appendedRange.match(/!A(\d+):H(\d+)$/);
  if (appendedRows) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: link.spreadsheetId,
      requestBody: {
        requests: [{
          repeatCell: {
            range: {
              sheetId: link.sheetId,
              startRowIndex: Number(appendedRows[1]) - 1,
              endRowIndex: Number(appendedRows[2]),
              startColumnIndex: 3,
              endColumnIndex: 4,
            },
            cell: { userEnteredFormat: { wrapStrategy: "WRAP" } },
            fields: "userEnteredFormat.wrapStrategy",
          },
        }],
      },
    });
  }
  return { success: true as const, added: uniqueRecords.length, skipped: records.length - uniqueRecords.length };
}
