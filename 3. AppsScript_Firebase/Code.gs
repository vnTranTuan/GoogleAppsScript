/**
 * Google Apps Script (Code.gs)
 * Chứa logic chính để điều khiển Sidebar và tương tác với Firestore.
 * Cần các hằng số cấu hình từ file Utils.gs.
 */

// ----------------------------------------------------------------------
// 1. HÀM CHẠY KHI MỞ FILE VÀ HIỂN THỊ SIDEBAR
// ----------------------------------------------------------------------

/**
 * Hàm onOpen() (Simple Trigger)
 * Chạy tự động khi Google Sheet được mở để tạo Menu Tùy chỉnh.
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🔥 Firebase Tools')
        .addItem('Mở Input Sidebar', 'showFirebaseSidebar')
        .addToUi();
  } catch (e) {
    Logger.log('Không thể tạo Menu Tùy chỉnh: ' + e.toString());
  }
}

/**
 * Hiển thị Sidebar (Thanh bên) sử dụng file FirebaseSidebar.html
 */
function showFirebaseSidebar() {
  try {
    // Tải nội dung HTML từ file FirebaseSidebar.html
    const html = HtmlService.createTemplateFromFile('FirebaseSidebar');
    const sidebar = html.evaluate().setTitle('Ghi Dữ Liệu vào Firestore');
    
    // Hiển thị Sidebar cho người dùng
    SpreadsheetApp.getUi().showSidebar(sidebar);
  } catch (e) {
    // Hiển thị lỗi nếu không thể mở Sidebar
    SpreadsheetApp.getUi().alert('LỖI: Không thể mở Sidebar. Chi tiết: ' + e.message);
  }
}


// ----------------------------------------------------------------------
// 2. HÀM XỬ LÝ LƯU DỮ LIỆU (SERVER-SIDE FUNCTION)
// ----------------------------------------------------------------------

/**
 * Hàm đọc dữ liệu từ Google Sheet (vùng A2:B5) và gửi mỗi hàng thành một document Firestore.
 * @return {string} Trả về thông báo thành công hoặc thất bại.
 */
function saveDataToFirestore() {
  
  try {
    // 1. Đọc dữ liệu từ Sheet (đọc toàn bộ vùng INPUT_RANGE, ví dụ A2:B5)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const range = sheet.getRange(INPUT_RANGE);
    const data = range.getValues(); // Lấy tất cả các hàng trong range
    let savedCount = 0;
    
    // Chuẩn bị URL Firestore (không đổi)
    const firestoreUrl = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/${FIRESTORE_COLLECTION}?key=${WEB_API_KEY}`;
    
    // Lặp qua từng HÀNG (record) trong dữ liệu
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const nameValue = row[0]; // Giá trị cột A
      const numericValue = row[1]; // Giá trị cột B
      const rowIndex = i + range.getRow(); // Dòng thực tế trong Sheet (ví dụ: 2, 3, 4, 5)

      // Bỏ qua hàng nếu cả hai cột đều trống
      if (nameValue === "" && numericValue === "") {
        Logger.log(`Bỏ qua dòng ${rowIndex} vì trống.`);
        continue;
      }
      
      // Kiểm tra giá trị bắt buộc/tính hợp lệ (có thể tùy chỉnh)
      if (nameValue === "" || numericValue === "") {
        throw new Error(`Dữ liệu không hợp lệ ở dòng ${rowIndex}. Vui lòng kiểm tra lại cột Tên (A) hoặc Giá trị (B).`);
      }
      
      // Kiểm tra nếu giá trị thứ hai KHÔNG phải là số
      if (isNaN(numericValue) || numericValue === null || typeof numericValue === 'string') {
         // UrlFetchApp.fetch yêu cầu giá trị số phải được bọc trong Number()
         // Hoặc đảm bảo rằng dữ liệu trong Sheet là định dạng số.
         // Tuy nhiên, đối với data read từ Sheets, nếu là số thì nó sẽ là Number, 
         // nếu không thì nó là String (cần kiểm tra isNaN)
         if (typeof numericValue !== 'number' && isNaN(Number(numericValue))) {
            throw new Error(`Dữ liệu không hợp lệ ở dòng ${rowIndex}. Giá trị cột B phải là một số.`);
         }
      }

      // 2. Chuẩn bị Payload cho Firestore cho record HIỆN TẠI
      const payload = {
        fields: {
          timestamp: { timestampValue: new Date().toISOString() },
          name: { stringValue: nameValue.toString() },
          value: { doubleValue: Number(numericValue) }, // Dùng doubleValue cho số
          sheetSource: { stringValue: ss.getName() }
        }
      };
      
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      // 3. Gửi Yêu cầu tới Firestore REST API
      const response = UrlFetchApp.fetch(firestoreUrl, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 200) {
        savedCount++;
        const docName = JSON.parse(responseText).name.split('/').pop();
        Logger.log(`Đã lưu thành công dòng ${rowIndex} (ID: ${docName}).`);
      } else {
        // Nếu có lỗi ở bất kỳ record nào, ném lỗi và dừng toàn bộ quá trình
        const errorDetail = JSON.parse(responseText).error.message;
        Logger.log(`LỖI FIRESTORE API ở dòng ${rowIndex} (${responseCode}): ${errorDetail}`);
        throw new Error(`LỖI FIRESTORE API ở dòng ${rowIndex}: ${errorDetail}`);
      }
    } // Hết vòng lặp FOR

    if (savedCount === 0) {
      return `Hoàn tất. Không có dòng dữ liệu hợp lệ nào được tìm thấy trong vùng ${INPUT_RANGE}.`;
    }

    // Thông báo thành công cuối cùng
    return `Thành công! Đã lưu ${savedCount} bản ghi từ vùng ${INPUT_RANGE} vào Firestore.`;
    
  } catch (e) {
    Logger.log('LỖI HỆ THỐNG: ' + e.toString());
    // Trả về chuỗi lỗi để Sidebar hiển thị
    return 'LỖI HỆ THỐNG: ' + e.message;
  }
}







