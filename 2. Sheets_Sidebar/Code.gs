/**
 * Hàm onOpen() chạy tự động khi Google Sheet được mở.
 * Nó tạo ra một menu tùy chỉnh để người dùng dễ dàng chọn chức năng.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu của tôi') // Tên menu
      .addItem('Thông tin người dùng', 'getUserInfo') // Tên mục menu và hàm gọi
      .addItem('Sidebar của tôi', 'showMySidebar') // Tên mục menu và hàm gọi
      .addToUi();
}


function showMySidebar() {
  //Tải template Tải nội dung HTML từ file có tên là Sidebar.html. Tên file phải khớp chính xác.
  const html = HtmlService.createTemplateFromFile('Sidebar');

   //	Tạo đối tượng Sidebar	Chuyển đổi chuỗi HTML thành đối tượng HtmlOutput mà Google có thể hiển thị.
   //	Đặt tiêu đề	Đặt tiêu đề cho cửa sổ Sidebar (tiêu đề xuất hiện ở thanh trên cùng).
  const sidebar = html.evaluate()
    .setTitle('Sidebar tự định nghĩa');   

  SpreadsheetApp.getUi().showSidebar(sidebar);
}


/**
 * Hàm getUserInfo: Lấy và hiển thị thông tin của người dùng hiện tại.
 * * LƯU Ý QUAN TRỌNG:
 * - Session.getActiveUser().getEmail() chỉ trả về email nếu tài khoản là Google Workspace 
 * HOẶC nếu script được chạy dưới dạng Web App.
 * - Đối với tài khoản cá nhân, nó có thể trả về một chuỗi rỗng trừ khi script được 
 * chạy với quyền cao hơn (như khi người dùng bấm vào menu).
 */
function getUserInfo() {
  const ui = SpreadsheetApp.getUi();

  try {
    // Lấy thông tin người dùng đang thực thi script
    const user = Session.getActiveUser();
    const userEmail = user.getEmail() || 'Không có sẵn (Có thể là tài khoản cá nhân)';
    const userName = user.getUsername() || 'Không có sẵn';
    
    // Lấy múi giờ của script
    const scriptTimezone = Session.getScriptTimeZone();

    // Lấy tên người dùng đang sử dụng 
    // (Nếu không có email, có thể dùng cách này để thử lấy tên chung)
    const effectiveUser = Session.getEffectiveUser();
    const effectiveEmail = effectiveUser.getEmail() || 'Không có sẵn (Quyền thực thi)';
    
    // Tạo nội dung hộp thoại
    const infoMessage = `
      Người dùng Kích hoạt (Active User):
      - Email: ${userEmail}
      - Tên người dùng: ${userName}

      Người dùng Thực thi (Effective User):
      - Email: ${effectiveEmail} 
      
      Thông tin Môi trường:
      - Múi giờ Script: ${scriptTimezone}
    `;
    ui.alert('Thông tin Người dùng', infoMessage, ui.ButtonSet.OK);

  } catch (e) {
    // Bắt lỗi nếu script không có quyền truy cập thông tin email
    Logger.log('Lỗi khi lấy thông tin người dùng: ' + e.toString());
    ui.alert('Lỗi', 'Không thể lấy thông tin người dùng. Vui lòng đảm bảo script đã được cấp quyền truy cập email.', ui.ButtonSet.OK);
  }
}


function alertAnotherFunction() {
  const ui = SpreadsheetApp.getUi();
    ui.alert('Chức năng khác', 'Bạn có thể bổ sung thêm nội dung chức năng khác ở đây', ui.ButtonSet.OK);

}

