// ============================================================================
// CHECKLIST - DATA MODEL CHỨA KẾT QUẢ PHÂN TÍCH EMAIL
// ============================================================================
// File: CheckList.cs
// Mô tả: Định nghĩa các class dữ liệu cho cửa sổ xác nhận trước khi gửi email
// ============================================================================

using System.Collections.Generic;

namespace OutlookOkan.Types
{
    // =========================================================================
    // LỚP CHECKLIST - KẾT QUẢ PHÂN TÍCH EMAIL
    // =========================================================================
    // Class này chứa TẤT CẢ thông tin được hiển thị trong cửa sổ xác nhận:
    // - Danh sách cảnh báo (Alerts)
    // - Danh sách người nhận (To/Cc/Bcc)
    // - Danh sách file đính kèm (Attachments)
    // - Thông tin email (Subject, Body, Sender)
    // - Cờ kiểm soát (IsCanNotSendMail, DeferredMinutes)
    // =========================================================================
    public sealed class CheckList
    {
        /// <summary>Danh sách các cảnh báo hiển thị ở đầu cửa sổ xác nhận</summary>
        public List<Alert> Alerts { get; set; } = new List<Alert>();
        
        /// <summary>Danh sách người nhận TO - hiển thị màu đỏ nếu external</summary>
        public List<Address> ToAddresses { get; set; } = new List<Address>();
        
        /// <summary>Danh sách người nhận CC - hiển thị màu đỏ nếu external</summary>
        public List<Address> CcAddresses { get; set; } = new List<Address>();
        
        /// <summary>Danh sách người nhận BCC - hiển thị màu đỏ nếu external</summary>
        public List<Address> BccAddresses { get; set; } = new List<Address>();
        
        /// <summary>Danh sách file đính kèm với metadata (size, type, dangerous flags)</summary>
        public List<Attachment> Attachments { get; set; } = new List<Attachment>();
        
        /// <summary>Địa chỉ email người gửi (ví dụ: user@company.com)</summary>
        public string Sender { get; set; }
        
        /// <summary>Domain của người gửi (ví dụ: @company.com) - dùng để xác định internal/external</summary>
        public string SenderDomain { get; set; }
        
        /// <summary>Số lượng domain bên ngoài trong danh sách recipients</summary>
        public int RecipientExternalDomainNumAll { get; set; }
        
        /// <summary>Tiêu đề email</summary>
        public string Subject { get; set; }
        
        /// <summary>Loại email: HTML/Plain/Rich Text/Meeting Request/Task Request</summary>
        public string MailType { get; set; }
        
        /// <summary>Nội dung email dạng plain text - hiển thị trong preview</summary>
        public string MailBody { get; set; }
        
        /// <summary>Nội dung email dạng HTML - dùng để kiểm tra embedded images</summary>
        public string MailHtmlBody { get; set; }
        
        /// <summary>CỜ QUAN TRỌNG: Nếu true → CHẶN hoàn toàn, không cho gửi email</summary>
        public bool IsCanNotSendMail { get; set; }
        
        /// <summary>Lý do không cho gửi email - hiển thị trong dialog lỗi</summary>
        public string CanNotSendMailMessage { get; set; }
        
        /// <summary>Số phút trì hoãn gửi email (0 = gửi ngay)</summary>
        public int DeferredMinutes { get; set; }
        
        /// <summary>Đường dẫn file tạm - dùng cho việc kiểm tra attachment</summary>
        public string TempFilePath { get; set; }
    }

    // =========================================================================
    // LỚP ALERT - THÔNG BÁO CẢNH BÁO
    // =========================================================================
    // Hiển thị ở đầu cửa sổ xác nhận, có thể có checkbox để user xác nhận đã đọc
    // =========================================================================
    public sealed class Alert
    {
        /// <summary>Nội dung cảnh báo hiển thị cho user</summary>
        public string AlertMessage { get; set; }
        
        /// <summary>Nếu true → hiển thị nổi bật (màu đỏ/icon cảnh báo)</summary>
        public bool IsImportant { get; set; }
        
        /// <summary>Nếu true → đây là thông báo thông tin, không phải cảnh báo</summary>
        public bool IsWhite { get; set; }
        
        /// <summary>Trạng thái checkbox - user phải check để có thể gửi email</summary>
        public bool IsChecked { get; set; }
    }

    // =========================================================================
    // LỚP ATTACHMENT - THÔNG TIN FILE ĐÍNH KÈM
    // =========================================================================
    // Chứa metadata của file và các cờ cảnh báo (quá lớn, nguy hiểm, mã hóa)
    // =========================================================================
    public sealed class Attachment
    {
        /// <summary>Tên file (ví dụ: "Report.pdf")</summary>
        public string FileName { get; set; }
        
        /// <summary>Loại file (File/Cloud Link/Embedded)</summary>
        public string FileType { get; set; }
        
        /// <summary>Dung lượng file đã format (ví dụ: "2.5 MB")</summary>
        public string FileSize { get; set; }
        
        /// <summary>Đường dẫn tạm của file - dùng để mở preview</summary>
        public string FilePath { get; set; }
        
        /// <summary>Text hiển thị cho nút Open (đa ngôn ngữ)</summary>
        public string Open { get; set; }
        
        /// <summary>Cờ: file quá lớn (vượt ngưỡng cài đặt)</summary>
        public bool IsTooBig { get; set; }
        
        /// <summary>Cờ: file nguy hiểm (có macro, extension đáng ngờ)</summary>
        public bool IsDangerous { get; set; }
        
        /// <summary>Cờ: file ZIP có mã hóa (password protected)</summary>
        public bool IsEncrypted { get; set; }
        
        /// <summary>Cờ: có thể mở preview file này không</summary>
        public bool IsCanOpen { get; set; }
        
        /// <summary>Cờ: không bắt buộc phải mở xem trước khi gửi</summary>
        public bool IsNotMustOpenBeforeCheck { get; set; }
        
        /// <summary>Trạng thái checkbox - user phải check để xác nhận</summary>
        public bool IsChecked { get; set; }
    }

    // =========================================================================
    // LỚP ADDRESS - THÔNG TIN ĐỊA CHỈ NGƯỜI NHẬN
    // =========================================================================
    // Chứa email và các cờ kiểm soát hiển thị trong cửa sổ xác nhận
    // =========================================================================
    public sealed class Address
    {
        /// <summary>Địa chỉ email (ví dụ: "Nguyễn Văn A (a@company.com)")</summary>
        public string MailAddress { get; set; }
        
        /// <summary>Cờ: địa chỉ bên ngoài công ty → hiển thị màu đỏ</summary>
        public bool IsExternal { get; set; }
        
        /// <summary>Cờ: địa chỉ trong whitelist → tự động check</summary>
        public bool IsWhite { get; set; }
        
        /// <summary>Cờ: bỏ qua xác nhận cho địa chỉ này</summary>
        public bool IsSkip { get; set; }
        
        /// <summary>Trạng thái checkbox - user phải check tất cả để gửi</summary>
        public bool IsChecked { get; set; }
    }
}