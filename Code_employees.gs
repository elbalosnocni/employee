function doGet(e) {
  return doPost(e);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Khong co du lieu gui den"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    var data = JSON.parse(e.postData.contents);
    var spreadsheetId = data.spreadsheet_id; 
    var sheetName = data.sheet_name || "DSCNV"; 
    
    if (!spreadsheetId) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Thieu thong tin spreadsheet_id"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = SpreadsheetApp.openById(spreadsheetId);
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Khong tim thay sheet co ten: " + sheetName}))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    var allRows = data.values; 
    
    if (Array.isArray(allRows) && allRows.length > 0) {
      var lastRow = sheet.getLastRow();
      var numCols = allRows[0].length; // SỬA CHUẨN: Lấy chính xác số lượng cột thực tế gửi sang (16 cột)
      
      // 1. Tải toàn bộ dữ liệu cũ trên Google Sheet vào bộ nhớ tạm để xử lý nhanh
      var sheetData = [];
      var cccdMap = {}; // Tra cứu dòng cũ dựa trên số CCCD
      
      if (lastRow > 1) {
        sheetData = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
        for (var r = 0; r < sheetData.length; r++) {
          var cccdKey = String(sheetData[r][5]); // CHUẨN XÁC: Số CCCD nằm ở Cột F (Index số 5 trong mảng Google Sheet)
          if (cccdKey) {
            cccdKey = cccdKey.replace(/'/g, "").trim();
            cccdMap[cccdKey] = r; // Lưu lại vị trí hàng trong mảng tạm sheetData
          }
        }
      }
      
      var countInsert = 0;
      var countUpdate = 0;
      
      // 2. Duyệt dữ liệu mới gửi từ VBA sang
      for (var i = 0; i < allRows.length; i++) {
        var rowData = allRows[i];
        
        // Chuẩn hóa CCCD lấy từ vị trí index 5 của dòng dữ liệu gửi sang
        var rawCCCD = String(rowData[5]).replace(/'/g, "").trim();
        
        // CHUẨN XÁC VỊ TRÍ NGÀY SINH: Index 3 là cột F (Ngày sinh Nam), Index 4 là cột G (Ngày sinh Nữ)
        rowData[3] = parseDateToStandard(rowData[3]); // Ngày sinh Nam (Lên Google sheet là cột D)
        rowData[4] = parseDateToStandard(rowData[4]); // Ngày sinh Nữ (Lên Google sheet là cột E)
        
        if (cccdMap.hasOwnProperty(rawCCCD)) {
          // TRÙNG CĂN CƯỚC -> CẬP NHẬT ĐÈ TRONG MẢNG TẠM
          var indexInSheetData = cccdMap[rawCCCD];
          sheetData[indexInSheetData] = rowData;
          countUpdate++;
        } else {
          // KHÔNG TRÙNG -> THÊM MỚI VÀO CUỐI MẢNG TẠM
          sheetData.push(rowData);
          cccdMap[rawCCCD] = sheetData.length - 1; 
          countInsert++;
        }
      }
      
      // 3. GHI HÀNG LOẠT XUỐNG SHEET (Chỉ gọi lệnh ghi đúng 1 lần duy nhất để chống Timeout)
      if (sheetData.length > 0) {
        // Xóa sạch vùng dữ liệu cũ dưới dòng tiêu đề để tránh bị lem hàng thừa cũ khi ghi đè
        if (lastRow > 1) {
          sheet.getRange(2, 1, sheet.getLastRow(), numCols).clearContent();
        }
        // Đổ mảng dữ liệu đã tối ưu xuống Google Sheet
        sheet.getRange(2, 1, sheetData.length, numCols).setValues(sheetData);
        
        // ĐỊNH DẠNG NGÀY CHUẨN XÁC: Định dạng Cột D (Cột số 4) và Cột E (Cột số 5) thành Ngày tháng định dạng VN
        sheet.getRange(2, 4, sheetData.length, 2).setNumberFormat("dd/mm/yyyy");
        
        // Ép kiểu hiển thị Cột F (Cột số 6 - CCCD) thành Plain Text để bảo vệ số 0 đầu số căn cước
        sheet.getRange(2, 6, sheetData.length, 1).setNumberFormat("@");
      }
      
      return ContentService.createTextOutput(JSON.stringify({
        "status": "success", 
        "message": "Hoan thanh! Them moi: " + countInsert + " dong, Cap nhat trung: " + countUpdate + " dong."
      })).setMimeType(ContentService.MimeType.JSON);
      
    } else {
      return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": "Du lieu 'values' trong hoac sai dinh dang"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// Hàm phân tích và định dạng ngày tháng sang chuẩn ISO yyyy-mm-dd để Google tự nhận biết kiểu Date
function parseDateToStandard(dateStr) {
  if (!dateStr || String(dateStr).trim() === "") return "";
  var cleanStr = String(dateStr).trim();
  
  var parts = cleanStr.split("/");
  if (parts.length === 3) {
    var day = parts[0];
    var month = parts[1];
    var year = parts[2];
    
    if (day.length === 1) day = "0" + day;
    if (month.length === 1) month = "0" + month;
    
    if (!isNaN(day) && !isNaN(month) && !isNaN(year)) {
      return year + "-" + month + "-" + day;
    }
  }
  return cleanStr; 
}
