/**
 * Google Apps Script để ghi nhận thông tin đăng ký từ landing page TrainghiemAI.html
 * Hỗ trợ ghi nhận URL đăng ký (bao gồm toàn bộ tham số UTM để tracking nguồn khách hàng)
 * 
 * Cách cài đặt:
 * 1. Mở Google Sheet lưu trữ data.
 * 2. Vào Tiện ích mở rộng (Extensions) -> Apps Script.
 * 3. Tạo một tệp script mới (ví dụ: CaptureLeads.gs), xóa code cũ và dán toàn bộ đoạn code dưới đây vào.
 * 4. Nhấn lưu (Ctrl+S / Cmd+S).
 * 5. Nhấn nút "Triển khai" (Deploy) -> "Triển khai mới" (New deployment).
 * 6. Chọn loại triển khai là "Ứng dụng web" (Web App).
 * 7. Thiết lập cấu hình:
 *    - Thực thi dưới danh nghĩa: "Tôi" (Execute as: Me)
 *    - Ai có quyền truy cập: "Mọi người" (Who has access: Anyone)
 * 8. Nhấn "Triển khai" (Deploy), cấp quyền cho ứng dụng và copy URL Web App được cấp.
 * 9. Dán URL Web App này thay thế cho GOOGLE_APP_SCRIPT_URL trong file TrainghiemAI.html ở dòng 1718.
 * 
 * CÁCH HOẠT ĐỘNG:
 * - Nếu Sheet của bạn KHÔNG có hàng tiêu đề đầu tiên (dòng 1 là data luôn):
 *   Script sẽ ghi theo đúng thứ tự cột cố định:
 *   A: Thời gian | B: Họ tên | C: Số điện thoại | D: Email | E: Khu vực | F: Vai trò | G: Link đăng ký | H: utm_source | I: utm_medium | J: utm_campaign | K: utm_content | L: utm_term
 * 
 * - Nếu Sheet của bạn CÓ hàng tiêu đề ở dòng 1:
 *   Bạn chỉ cần đặt tên các cột bất kỳ chứa các chữ: "Thời gian", "Họ tên", "Số điện thoại", "Email", "Khu vực", "Vai trò", "Link đăng ký", "utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term".
 *   Script sẽ tự tìm cột tương ứng để ghi (không quan trọng thứ tự cột).
 */

function doPost(e) {
  // Tránh lỗi CORS bằng cách trả về header phù hợp
  const headersResponse = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type'
  };
  
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const parameter = e.parameter || {};
    const timestamp = new Date();
    
    // Kiểm tra xem dòng 1 có phải là Dữ liệu hay Tiêu đề
    const firstCell = sheet.getRange(1, 1).getValue();
    let hasHeader = false;
    
    if (firstCell && !(firstCell instanceof Date) && isNaN(Date.parse(firstCell))) {
      // Nếu ô đầu tiên là chữ (không phải ngày tháng), kiểm tra tiếp xem có chứa các từ khóa tiêu đề không
      const lastCol = Math.min(sheet.getLastColumn(), 5);
      const firstRowValues = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      const headerKeywords = ['họ tên', 'tên', 'name', 'sđt', 'phone', 'email', 'khu vực', 'area'];
      hasHeader = firstRowValues.some(val => 
        headerKeywords.some(keyword => val.toString().toLowerCase().includes(keyword))
      );
    }
    
    let rowData = [];
    
    if (hasHeader) {
      // 1. Nếu có hàng tiêu đề: Map dữ liệu động theo tên cột (không quan trọng thứ tự)
      const lastCol = Math.max(sheet.getLastColumn(), 1);
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      
      for (let i = 0; i < headers.length; i++) {
        const colHeader = headers[i].toString().trim().toLowerCase();
        
        if (!colHeader) {
          rowData.push('');
          continue;
        }
        
        // Map từng cột
        if (['timestamp', 'thời gian', 'ngày đăng ký', 'ngày'].includes(colHeader)) {
          rowData.push(timestamp);
        } else if (['name', 'họ tên', 'họ và tên', 'fullname'].includes(colHeader)) {
          rowData.push(parameter.name || '');
        } else if (['phone', 'số điện thoại', 'sđt', 'telephone'].includes(colHeader)) {
          const phone = parameter.phone || '';
          rowData.push(phone ? "'" + phone : '');
        } else if (['email'].includes(colHeader)) {
          rowData.push(parameter.email || '');
        } else if (['area', 'khu vực', 'địa chỉ', 'tỉnh'].includes(colHeader)) {
          rowData.push(parameter.area || '');
        } else if (['role', 'vai trò', 'đối tượng', 'nhóm'].includes(colHeader)) {
          rowData.push(parameter.role || '');
        } else if (['link đăng ký', 'register link', 'url', 'link', 'nguồn đăng ký', 'register_link'].includes(colHeader)) {
          rowData.push(parameter.register_link || '');
        } else if (['utm_source', 'utm source', 'nguồn'].includes(colHeader)) {
          rowData.push(parameter.utm_source || '');
        } else if (['utm_medium', 'utm medium', 'kênh'].includes(colHeader)) {
          rowData.push(parameter.utm_medium || '');
        } else if (['utm_campaign', 'utm campaign', 'chiến dịch'].includes(colHeader)) {
          rowData.push(parameter.utm_campaign || '');
        } else if (['utm_content', 'utm content', 'nội dung quảng cáo'].includes(colHeader)) {
          rowData.push(parameter.utm_content || '');
        } else if (['utm_term', 'utm term', 'từ khóa'].includes(colHeader)) {
          rowData.push(parameter.utm_term || '');
        } else {
          // Thử map trực tiếp với tên parameter truyền lên
          rowData.push(parameter[headers[i]] || '');
        }
      }
    } else {
      // 2. Nếu KHÔNG có hàng tiêu đề (Dòng 1 là data luôn): Ghi theo thứ tự mặc định cố định
      const phone = parameter.phone || '';
      rowData = [
        timestamp,                              // A
        parameter.name || '',                   // B
        phone ? "'" + phone : '',               // C
        parameter.email || '',                  // D
        parameter.area || '',                   // E
        parameter.role || '',                   // F
        parameter.register_link || '',          // G
        parameter.utm_source || '',             // H
        parameter.utm_medium || '',             // I
        parameter.utm_campaign || '',           // J
        parameter.utm_content || '',            // K
        parameter.utm_term || ''                // L
      ];
    }
    
    // Ghi hàng dữ liệu vào sheet
    sheet.appendRow(rowData);
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success', 
      message: 'Ghi dữ liệu thành công',
      timestamp: timestamp
    }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders(headersResponse);
    
  } catch (err) {
    Logger.log("Lỗi: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    }))
    .setMimeType(ContentService.MimeType.JSON)
    .setHeaders(headersResponse);
  }
}

// Xử lý cả request OPTIONS (Preflight request trong CORS)
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .setHeaders({
      'Access-Control-Allow-Origin': '*',
      'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
      'Access-Control-Allow-Headers': 'Content-Type'
    });
}
