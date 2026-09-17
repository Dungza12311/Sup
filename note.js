const SPREADSHEET_ID = "1hJfzW1HEbRMFBiPAeGK5CTnsjQEoFNud8OU_zWWVLeY";
const SHEET_ID = 673545916;


/* ================================
   KIỂM TRA WEB APP
================================ */

function doGet() {
  return ContentService
    .createTextOutput("VISNAM SUPPORT SYSTEM - WEB APP OK")
    .setMimeType(ContentService.MimeType.TEXT);
}


/* ================================
   XÓA TOÀN BỘ MÀU NỀN TRONG SHEET
================================ */

function xoaToanBoMau() {

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const sheet = ss.getSheets().find(
    function(s) {
      return s.getSheetId() === SHEET_ID;
    }
  );

  if (!sheet) {
    throw new Error(
      "Không tìm thấy Sheet có GID = " + SHEET_ID
    );
  }

  // Lấy toàn bộ vùng đang sử dụng
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const lastColumn = Math.max(sheet.getLastColumn(), 9);

  // Xóa toàn bộ màu nền
  sheet
    .getRange(1, 1, lastRow, lastColumn)
    .setBackground("#ffffff");

  return "Đã chuyển toàn bộ Sheet sang màu trắng.";
}


/* ================================
   NHẬN DỮ LIỆU TỪ HTML
================================ */

function doPost(e) {

  try {

    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Không nhận được dữ liệu từ HTML.");
    }


    // Đọc JSON
    let data;

    try {
      data = JSON.parse(e.postData.contents);
    } catch (error) {
      throw new Error(
        "Dữ liệu gửi lên không phải JSON hợp lệ."
      );
    }


    // ==============================
    // LẤY DỮ LIỆU
    // ==============================

    const mst = String(data.mst || "").trim();

    const sdt = String(data.sdt || "").trim();

    const tencty = String(data.tencty || "").trim();

    const loai = String(data.loai || "").trim();

    const vanDe = String(data.vanDe || "").trim();

    const cskh = String(data.cskh || "").trim();

    const nguoiHoTro = String(
      data.nguoiHoTro || "Dũng"
    ).trim();


    // ==============================
    // CHUẨN HÓA
    // ==============================

    const dichVu = loai.toUpperCase();

    let kyThuat = "";

    if (nguoiHoTro !== "") {

      kyThuat =
        nguoiHoTro.charAt(0).toUpperCase() +
        nguoiHoTro.slice(1).toLowerCase();

    }


    // ==============================
    // MỞ SHEET
    // ==============================

    const ss = SpreadsheetApp.openById(
      SPREADSHEET_ID
    );


    const sheet = ss.getSheets().find(
      function(s) {
        return s.getSheetId() === SHEET_ID;
      }
    );


    if (!sheet) {
      throw new Error(
        "Không tìm thấy Sheet có GID = " + SHEET_ID
      );
    }


    // ==============================
    // DÒNG MỚI
    // ==============================

    const nextRow = sheet.getLastRow() + 1;


    // ==============================
    // GHI 9 CỘT
    // ==============================

    sheet
      .getRange(nextRow, 1, 1, 9)
      .setValues([[
        "",          // A - Trống
        mst,         // B - MST
        sdt,         // C - SĐT
        tencty,      // D - Tên công ty
        dichVu,      // E - Dịch vụ
        kyThuat,     // F - Kỹ thuật
        vanDe,       // G - Vấn đề hỗ trợ
        "",          // H - Trống
        cskh         // I - CSKH
      ]]);


    // ==============================
    // DÒNG MỚI LUÔN TRẮNG
    // ==============================

    sheet
      .getRange(nextRow, 1, 1, 9)
      .setBackground("#ffffff");


    // ==============================
    // ĐẢM BẢO A VÀ H TRỐNG
    // ==============================

    sheet.getRange(nextRow, 1).clearContent();

    sheet.getRange(nextRow, 8).clearContent();


    // ==============================
    // TRẢ KẾT QUẢ
    // ==============================

    return ContentService
      .createTextOutput(
        JSON.stringify({
          success: true,
          message: "Đã lưu thành công.",
          row: nextRow
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );


  } catch (error) {

    console.error(error);

    return ContentService
      .createTextOutput(
        JSON.stringify({
          success: false,
          message: error.message
        })
      )
      .setMimeType(
        ContentService.MimeType.JSON
      );
  }
}