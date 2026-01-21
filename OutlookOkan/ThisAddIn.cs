// ============================================================================
// OUTLOOKOKAN - ĐIỂM VÀO CHÍNH (ENTRY POINT)
// ============================================================================
// File: ThisAddIn.cs
// Mô tả: Đây là file chính của VSTO Add-in, xử lý tất cả sự kiện từ Outlook
// Tác giả: t-miyake | Dịch comment: AI Assistant
// ============================================================================

// --- CÁC THƯ VIỆN SỬ DỤNG ---
using OutlookOkan.Handlers;   // Xử lý file: CSV, Mail Header, Office, PDF, ZIP
using OutlookOkan.Helpers;    // Các hàm hỗ trợ native Windows
using OutlookOkan.Models;     // Business logic chính (GenerateCheckList)
using OutlookOkan.Services;   // Dịch vụ hỗ trợ (đa ngôn ngữ)
using OutlookOkan.Types;      // Các data model (CheckList, Alert, Address, v.v.)
using OutlookOkan.Views;      // Các cửa sổ WPF (Confirmation, Settings, About)
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;  // Để làm việc với COM objects (Outlook)
using System.Windows;                   // WPF MessageBox
using System.Windows.Interop;           // Để set Owner cho WPF windows
using Outlook = Microsoft.Office.Interop.Outlook;  // Thư viện COM của Outlook
using Word = Microsoft.Office.Interop.Word;        // Thư viện COM của Word (dùng cho WordEditor)

namespace OutlookOkan
{
    /// <summary>
    /// LỚP CHÍNH CỦA ADD-IN
    /// ======================
    /// Đây là lớp "partial" - phần còn lại được sinh tự động bởi Visual Studio
    /// trong file ThisAddIn.Designer.cs
    /// 
    /// LUỒNG HOẠT ĐỘNG:
    /// 1. Outlook khởi động → Gọi ThisAddIn_Startup()
    /// 2. User chọn email   → Gọi CurrentExplorer_SelectionChange() 
    /// 3. User gửi email    → Gọi Application_ItemSend() ← QUAN TRỌNG NHẤT
    /// 4. User mở attachment → Gọi BeforeAttachmentRead()
    /// </summary>
    public partial class ThisAddIn
    {
        // =====================================================================
        // CÁC BIẾN LƯU TRỮ CÀI ĐẶT (SETTINGS)
        // =====================================================================

        /// <summary>
        /// Cài đặt chung của add-in (ngôn ngữ, bật/tắt tính năng, v.v.)
        /// Được load từ file: %APPDATA%\Noraneko\OutlookOkan\GeneralSetting.csv
        /// </summary>
        private readonly GeneralSetting _generalSetting = new GeneralSetting();

        /// <summary>
        /// Cài đặt bảo mật cho email nhận (kiểm tra SPF, DKIM, DMARC, v.v.)
        /// Được load từ file: SecurityForReceivedMail.csv
        /// </summary>
        private readonly SecurityForReceivedMail _securityForReceivedMail = new SecurityForReceivedMail();

        /// <summary>
        /// Danh sách từ khóa cảnh báo trong tiêu đề email nhận
        /// Ví dụ: "[긴급]" → hiện cảnh báo khi mở email có tiêu đề chứa từ này
        /// </summary>
        private readonly List<AlertKeywordOfSubjectWhenOpeningMail> _alertKeywordOfSubjectWhenOpeningMail = new List<AlertKeywordOfSubjectWhenOpeningMail>();

        // =====================================================================
        // CÁC BIẾN COM OBJECTS CỦA OUTLOOK
        // =====================================================================

        /// <summary>
        /// Quản lý tất cả các cửa sổ Inspector (cửa sổ soạn/đọc email riêng)
        /// Dùng để bắt sự kiện khi user mở email từ Outbox (đang chờ gửi)
        /// </summary>
        private Outlook.Inspectors _inspectors;

        /// <summary>
        /// Cửa sổ Explorer hiện tại (cửa sổ chính của Outlook)
        /// Dùng để bắt sự kiện khi user chọn email khác
        /// </summary>
        private Outlook.Explorer _currentExplorer;

        /// <summary>
        /// Email đang được chọn hiện tại
        /// Cần lưu lại để có thể gỡ event handler khi chọn email khác
        /// </summary>
        private Outlook.MailItem _currentMailItem;

        /// <summary>
        /// Namespace MAPI - điểm truy cập vào dữ liệu Outlook
        /// Dùng để lấy thông tin các folder mặc định
        /// </summary>
        private Outlook.NameSpace _mapiNamespace;

        /// <summary>
        /// Danh sách tên các folder KHÔNG cần kiểm tra bảo mật
        /// (Calendar, Contacts, Drafts, Sent Items, v.v.)
        /// Dùng HashSet để tìm kiếm nhanh O(1)
        /// </summary>
        private HashSet<string> _excludedFolderNames;

        // =====================================================================
        // TẠO RIBBON (NÚT BẤM TRÊN THANH CÔNG CỤ OUTLOOK)
        // =====================================================================

        /// <summary>
        /// TẠO RIBBON EXTENSION
        /// ====================
        /// Phương thức này được gọi bởi VSTO framework để tạo các nút bấm
        /// trên thanh Ribbon của Outlook (Settings, About, Help, Verify Header)
        /// </summary>
        /// <returns>Đối tượng Ribbon chứa các nút bấm</returns>
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();  // Xem chi tiết trong file Ribbon.cs
        }

        // =====================================================================
        // SỰ KIỆN KHỞI ĐỘNG ADD-IN
        // =====================================================================

        /// <summary>
        /// KHỞI ĐỘNG ADD-IN (KHI OUTLOOK MỞ)
        /// ==================================
        /// Đây là điểm bắt đầu của add-in, được gọi khi Outlook khởi động.
        /// 
        /// TRÌNH TỰ THỰC HIỆN:
        /// 1. Load cài đặt ngôn ngữ → đổi giao diện nếu cần
        /// 2. Load cài đặt bảo mật cho email nhận
        /// 3. Nếu bật bảo mật: đăng ký event khi chọn email
        /// 4. Đăng ký event khi gửi email (QUAN TRỌNG NHẤT)
        /// </summary>
        /// <param name="sender">Đối tượng gửi sự kiện (Outlook)</param>
        /// <param name="e">Thông tin sự kiện</param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            // -----------------------------------------------------------------
            // BƯỚC 1: LOAD CÀI ĐẶT CHUNG VÀ NGÔN NGỮ
            // -----------------------------------------------------------------
            // Phải load setting sớm để đổi ngôn ngữ Ribbon trước khi hiển thị
            // Nếu load muộn, Ribbon sẽ hiện bằng ngôn ngữ mặc định (Nhật)
            LoadGeneralSetting(isLaunch: true);

            // Nếu user đã chọn ngôn ngữ → đổi ngôn ngữ giao diện
            if (!(_generalSetting.LanguageCode is null))
            {
                ResourceService.Instance.ChangeCulture(_generalSetting.LanguageCode);
            }

            // -----------------------------------------------------------------
            // BƯỚC 2: LOAD CÀI ĐẶT BẢO MẬT EMAIL NHẬN
            // -----------------------------------------------------------------
            LoadSecurityForReceivedMail();

            // Chỉ thiết lập nếu user bật tính năng bảo mật email nhận
            if (_securityForReceivedMail.IsEnableSecurityForReceivedMail)
            {
                try
                {
                    // Lấy cửa sổ Explorer hiện tại (cửa sổ chính Outlook)
                    _currentExplorer = Application.ActiveExplorer();

                    if (_currentExplorer is null)
                    {
                        // Trường hợp không có Explorer (hiếm khi xảy ra)
                        System.Diagnostics.Debug.WriteLine($"ThisAddIn_Startup: ActiveExplorer is null.");
                    }
                    else
                    {
                        // Lấy namespace MAPI để truy cập dữ liệu Outlook
                        _mapiNamespace = Application.GetNamespace("MAPI");

                        if (_mapiNamespace is null)
                        {
                            // Không thể kết nối MAPI (có thể do không có mạng)
                            MessageBox.Show(
                                Properties.Resources.IsNoInternetCantUseSecurityForReceivedMail,
                                Properties.Resources.AppName,
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                        }
                        else
                        {
                            // -------------------------------------------------
                            // TẠO DANH SÁCH FOLDER LOẠI TRỪ
                            // -------------------------------------------------
                            // Các folder này không phải là hộp thư đến,
                            // không cần kiểm tra bảo mật khi user chọn item
                            _excludedFolderNames = new HashSet<string>{
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderCalendar).Name,      // Lịch
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts).Name,      // Danh bạ
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderDrafts).Name,        // Bản nháp
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderJournal).Name,       // Nhật ký
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderNotes).Name,         // Ghi chú
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderOutbox).Name,        // Hộp thư đi
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderRssFeeds).Name,      // RSS
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail).Name,      // Đã gửi
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderServerFailures).Name,// Lỗi server
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderLocalFailures).Name, // Lỗi local
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSyncIssues).Name,    // Lỗi đồng bộ
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderTasks).Name,         // Công việc
                                _mapiNamespace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderToDo).Name           // Việc cần làm
                            };

                            // Load danh sách từ khóa cảnh báo trong tiêu đề email
                            LoadAlertKeywordOfSubjectWhenOpeningMailsData();

                            // Đăng ký event: khi user chọn email khác → kiểm tra bảo mật
                            _currentExplorer.SelectionChange += CurrentExplorer_SelectionChange;
                        }
                    }
                }
                catch (Exception exception)
                {
                    // ---------------------------------------------------------
                    // XỬ LÝ LỖI KHI THIẾT LẬP BẢO MẬT
                    // ---------------------------------------------------------
                    // Kiểm tra mã lỗi để hiện thông báo phù hợp
                    MessageBox.Show(
                        exception.HResult == ComErrorCodes.MkEUnavailable
                            ? Properties.Resources.IsNoInternetCantUseSecurityForReceivedMail  // Không có mạng
                            : Properties.Resources.CantUseSecurityForReceivedMail,             // Lỗi khác
                        Properties.Resources.AppName,
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
            }

            // -----------------------------------------------------------------
            // BƯỚC 3: ĐĂNG KÝ CÁC EVENT QUAN TRỌNG
            // -----------------------------------------------------------------

            // Lấy danh sách tất cả Inspector (cửa sổ soạn/đọc email)
            _inspectors = Application.Inspectors;

            // Đăng ký event: khi mở cửa sổ Inspector mới
            // → dùng để cảnh báo khi mở email đang chờ gửi từ Outbox
            _inspectors.NewInspector += OpenOutboxItemInspector;

            // ★★★ QUAN TRỌNG NHẤT ★★★
            // Đăng ký event: khi user click nút Send
            // → đây là nơi add-in can thiệp để hiện cửa sổ xác nhận
            Application.ItemSend += Application_ItemSend;
        }

        // =====================================================================
        // SỰ KIỆN KHI USER CHỌN EMAIL KHÁC (BẢO MẬT EMAIL NHẬN)
        // =====================================================================

        /// <summary>
        /// Lưu EntryID của email đang chọn để tránh xử lý lại cùng một email
        /// EntryID là mã định danh duy nhất của mỗi email trong Outlook
        /// </summary>
        private string _currentMailItemEntryId = "";

        /// <summary>
        /// XỬ LÝ KHI USER CHỌN EMAIL KHÁC
        /// ==============================
        /// Được gọi mỗi khi user click vào email khác trong danh sách.
        /// Thực hiện các kiểm tra bảo mật cho email nhận.
        /// 
        /// CÁC BƯỚC XỬ LÝ:
        /// 1. Bỏ qua nếu đang ở folder không cần kiểm tra (Calendar, Contacts, v.v.)
        /// 2. Bỏ qua nếu chọn nhiều email hoặc không phải MailItem
        /// 3. Bỏ qua nếu là cùng email đã chọn trước đó
        /// 4. Kiểm tra từ khóa cảnh báo trong tiêu đề
        /// 5. Phân tích header (SPF, DKIM, DMARC) để phát hiện email giả mạo
        /// 6. Đăng ký event kiểm tra attachment nếu email có file đính kèm
        /// </summary>
        private void CurrentExplorer_SelectionChange()
        {
            // -----------------------------------------------------------------
            // BƯỚC 1: KIỂM TRA FOLDER HIỆN TẠI
            // -----------------------------------------------------------------
            var currentExplorer = Application.ActiveExplorer();
            var currentFolderName = currentExplorer.CurrentFolder.Name;

            // Bỏ qua nếu đang ở folder không cần kiểm tra
            // (Calendar, Contacts, Drafts, Sent Items, v.v.)
            if (_excludedFolderNames.Contains(currentFolderName)) return;

            // -----------------------------------------------------------------
            // BƯỚC 2: KIỂM TRA SELECTION HỢP LỆ
            // -----------------------------------------------------------------
            var selection = currentExplorer.Selection;

            // Bỏ qua nếu không có selection hoặc chọn nhiều email
            if (selection is null || selection.Count != 1) return;

            // Bỏ qua nếu item được chọn không phải là MailItem
            // (có thể là Meeting, Task, Contact, v.v.)
            if (!(selection[1] is Outlook.MailItem selectedMail)) return;

            // Bỏ qua nếu user click lại vào cùng email đang chọn
            // (để không hiện cảnh báo nhiều lần)
            if (_currentMailItemEntryId == selectedMail.EntryID) return;

            // -----------------------------------------------------------------
            // BƯỚC 3: DỌN DẸP EVENT HANDLER CŨ
            // -----------------------------------------------------------------
            // Gỡ event handler từ email cũ trước khi gán email mới
            // Điều này tránh memory leak và event handler chồng chéo
            if (_currentMailItem != null)
            {
                _currentMailItem.BeforeAttachmentRead -= BeforeAttachmentRead;
            }

            // Cập nhật email hiện tại
            _currentMailItem = selectedMail;
            _currentMailItemEntryId = _currentMailItem.EntryID;

            // -----------------------------------------------------------------
            // BƯỚC 4: KIỂM TRA TỪ KHÓA CẢNH BÁO TRONG TIÊU ĐỀ
            // -----------------------------------------------------------------
            // Ví dụ: tiêu đề chứa "[긴급]" (khẩn cấp) → hiện cảnh báo
            if (_securityForReceivedMail.IsEnableAlertKeywordOfSubjectWhenOpeningMailsData)
            {
                var subject = selectedMail.Subject;

                // Tìm từ khóa cảnh báo trong tiêu đề email
                // Dùng FirstOrDefault để tìm từ khóa đầu tiên match
                var settings = _alertKeywordOfSubjectWhenOpeningMail
                    .FirstOrDefault(x => subject.Contains(x.AlertKeyword));

                if (!(settings is null))
                {
                    // Tạo thông báo cảnh báo
                    var message = Properties.Resources.AlertOfReceivedMailSubject
                        + Environment.NewLine
                        + "[" + settings.AlertKeyword + "]";

                    // Nếu có custom message → dùng custom message
                    if (!string.IsNullOrEmpty(settings.Message))
                    {
                        message = settings.Message;
                    }

                    // Hiện cảnh báo cho user
                    MessageBox.Show(
                        message,
                        Properties.Resources.Warning,
                        MessageBoxButton.OK,
                        MessageBoxImage.Exclamation);
                }
            }

            // -----------------------------------------------------------------
            // BƯỚC 5: PHÂN TÍCH EMAIL HEADER (SPF, DKIM, DMARC)
            // -----------------------------------------------------------------
            // Đây là kiểm tra quan trọng để phát hiện email giả mạo (spoofing)
            if (_securityForReceivedMail.IsEnableMailHeaderAnalysis)
            {
                // Lấy email header từ MAPI property
                // 0x007D001E = PR_TRANSPORT_MESSAGE_HEADERS (full email headers)
                var header = selectedMail.PropertyAccessor.GetProperty(Constants.PR_TRANSPORT_MESSAGE_HEADERS);

                // Phân tích header để lấy kết quả SPF, DKIM, DMARC
                var analysisResults = MailHeaderHandler.ValidateEmailHeader(header.ToString());

                if (!(analysisResults is null))
                {
                    // Kiểm tra xem có phải mail nội bộ không
                    // Mail nội bộ thường không có SPF/DKIM/DMARC
                    var isInternalMail = analysisResults["SPF"] == "NONE"
                        && analysisResults["DKIM"] == "NONE"
                        && analysisResults["DMARC"] == "NONE"
                        && analysisResults["Internal"] == "TRUE";

                    // Chỉ cảnh báo cho email từ bên ngoài
                    // Email nội bộ (internal) không cần cảnh báo
                    if (!isInternalMail)
                    {
                        // Tạo thông báo hiển thị kết quả phân tích
                        var message = "";
                        foreach (KeyValuePair<string, string> entry in analysisResults)
                        {
                            message += ($"{entry.Key}: {entry.Value}") + Environment.NewLine;
                        }

                        // Kiểm tra nguy cơ giả mạo (spoofing risk)
                        if (_securityForReceivedMail.IsShowWarningWhenSpoofingRisk)
                        {
                            if (_securityForReceivedMail.IsShowWarningWhenDmarcNotImplemented)
                            {
                                //DMARCがPASSでない場合、常に警告
                                if (analysisResults["DMARC"] != "PASS")
                                {
                                    _ = MessageBox.Show(Properties.Resources.SpoofingRiskWaring + Environment.NewLine + Properties.Resources.SpfDkimWaring2 + Environment.NewLine + Environment.NewLine + message, Properties.Resources.Warning, MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                            }
                            else
                            {
                                var selfGeneratedDmarcResult = MailHeaderHandler.DetermineDmarcResult(analysisResults["SPF"], analysisResults["SPF Alignment"], analysisResults["DKIM"], analysisResults["DKIM Alignment"]);

                                if (analysisResults["DMARC"] != "PASS" && analysisResults["DMARC"] != "BESTGUESSPASS" && selfGeneratedDmarcResult == "FAIL")
                                {
                                    _ = MessageBox.Show(Properties.Resources.SpoofingRiskWaring + Environment.NewLine + Properties.Resources.SpfDkimWaring2 + Environment.NewLine + Environment.NewLine + message, Properties.Resources.Warning, MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                            }
                        }
                        else
                        {
                            //「なりすまし(送信元偽装)の危険性がある場合に警告する。」機能が有効な場合、SPFやDKIM単独の確認は行わない。

                            //SPFレコードの検証に失敗した場合に警告を表示する。
                            if (_securityForReceivedMail.IsShowWarningWhenSpfFails)
                            {
                                if (analysisResults["SPF"] == "FAIL" || analysisResults["SPF"] == "NONE")
                                {

                                    _ = MessageBox.Show(Properties.Resources.SpfWarning1 + Environment.NewLine + Properties.Resources.SpfDkimWaring2 + Environment.NewLine + Environment.NewLine + message, Properties.Resources.Warning, MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                            }

                            //DKIMレコードの検証に失敗した場合に警告を表示する。
                            if (_securityForReceivedMail.IsShowWarningWhenDkimFails)
                            {
                                if (analysisResults["DKIM"] == "FAIL")
                                {
                                    _ = MessageBox.Show(Properties.Resources.DkimWarning1 + Environment.NewLine + Properties.Resources.SpfDkimWaring2 + Environment.NewLine + Environment.NewLine + message, Properties.Resources.Warning, MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                            }
                        }
                    }
                }
            }

            //添付ファイルを開くときの警告機能
            if (_securityForReceivedMail.IsEnableWarningFeatureWhenOpeningAttachments && selectedMail.Attachments.Count != 0)
            {
                _currentMailItem.BeforeAttachmentRead -= BeforeAttachmentRead;
                _currentMailItem.BeforeAttachmentRead += BeforeAttachmentRead;
            }
        }

        /// <summary>
        /// 添付ファイルを開く時の分析と警告
        /// </summary>
        /// <param name="attachment"></param>
        /// <param name="cancel"></param>
        private void BeforeAttachmentRead(Outlook.Attachment attachment, ref bool cancel)
        {
            //添付ファイルを開く前の警告機能
            if (_securityForReceivedMail.IsWarnBeforeOpeningAttachments)
            {
                var dialogResult = MessageBox.Show(Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Properties.Resources.OpenAttachmentWarning2 + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (dialogResult == MessageBoxResult.Yes)
                {
                    //Open file.
                }
                else
                {
                    cancel = true;
                    return;
                }
            }

            if (_securityForReceivedMail.IsWarnBeforeOpeningEncryptedZip || _securityForReceivedMail.IsWarnLinkFileInTheZip || _securityForReceivedMail.IsWarnOneFileInTheZip || _securityForReceivedMail.IsWarnOfficeFileWithMacroInTheZip || _securityForReceivedMail.IsWarnBeforeOpeningAttachmentsThatContainMacros)
            {
                var tempDirectoryPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
                _ = Directory.CreateDirectory(tempDirectoryPath);
                var tempFilePath = Path.Combine(tempDirectoryPath, Guid.NewGuid().ToString("N"));
                attachment.SaveAsFile(tempFilePath);

                if (_securityForReceivedMail.IsWarnBeforeOpeningEncryptedZip || _securityForReceivedMail.IsWarnLinkFileInTheZip || _securityForReceivedMail.IsWarnOneFileInTheZip || _securityForReceivedMail.IsWarnOfficeFileWithMacroInTheZip)
                {
                    var zipTools = new ZipFileHandler();
                    var izEncryptedZip = zipTools.CheckZipIsEncryptedAndGetIncludeExtensions(tempFilePath);

                    //暗号化ZIPファイルの場合の警告
                    if (_securityForReceivedMail.IsWarnBeforeOpeningEncryptedZip && izEncryptedZip)
                    {
                        var dialogResult = MessageBox.Show(Properties.Resources.AttatchmentIsEncryptedZip + Environment.NewLine + Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                        if (dialogResult == MessageBoxResult.Yes)
                        {
                            //Open file.
                        }
                        else
                        {
                            cancel = true;
                            try
                            {
                                File.Delete(tempFilePath);
                            }
                            catch (Exception ex)
                            {
                                // Log temp file cleanup error
                                System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to delete temp file (EncryptedZip): {ex.Message}");
                            }
                            return;
                        }
                    }

                    //Zip内にlinkファイルがある場合の警告
                    if (_securityForReceivedMail.IsWarnLinkFileInTheZip)
                    {
                        if (zipTools.IncludeExtensions.Contains(".lnk") || zipTools.IsContainsShortcut)
                        {
                            var dialogResult = MessageBox.Show(Properties.Resources.SuspiciousAttachmentZip_link + Environment.NewLine + Environment.NewLine + Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Error);
                            if (dialogResult == MessageBoxResult.Yes)
                            {
                                //Open file.
                            }
                            else
                            {
                                cancel = true;
                                try
                                {
                                    File.Delete(tempFilePath);
                                }
                                catch (Exception ex)
                                {
                                    // Log temp file cleanup error
                                    System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to delete temp file (LinkInZip): {ex.Message}");
                                }
                                return;
                            }
                        }
                    }

                    //Zip内にOneNoteファイルがある場合の警告
                    if (_securityForReceivedMail.IsWarnOneFileInTheZip)
                    {
                        if (zipTools.IncludeExtensions.Contains(".one"))
                        {
                            var dialogResult = MessageBox.Show(Properties.Resources.SuspiciousAttachmentZip_one + Environment.NewLine + Environment.NewLine + Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Error);
                            if (dialogResult == MessageBoxResult.Yes)
                            {
                                //Open file.
                            }
                            else
                            {
                                cancel = true;
                                try
                                {
                                    File.Delete(tempFilePath);
                                }
                                catch (Exception ex)
                                {
                                    // Log temp file cleanup error
                                    System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to delete temp file (OneFileInZip): {ex.Message}");
                                }
                                return;
                            }
                        }
                    }

                    //Zip内にマクロ付きOfficeファイルがある場合の警告
                    if (_securityForReceivedMail.IsWarnOfficeFileWithMacroInTheZip)
                    {
                        if (zipTools.IncludeExtensions.Contains(".docm") | zipTools.IncludeExtensions.Contains(".xlsm") | zipTools.IncludeExtensions.Contains(".pptm"))
                        {
                            var dialogResult = MessageBox.Show(Properties.Resources.SuspiciousAttachmentZip_macro + Environment.NewLine + Environment.NewLine + Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Error);
                            if (dialogResult == MessageBoxResult.Yes)
                            {
                                //Open file.
                            }
                            else
                            {
                                cancel = true;
                                try
                                {
                                    File.Delete(tempFilePath);
                                }
                                catch (Exception ex)
                                {
                                    // Log temp file cleanup error
                                    System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to delete temp file (MacroInZip): {ex.Message}");
                                }
                                return;
                            }
                        }
                    }
                }

                //Officeファイル内にマクロが含まれている場合の警告
                if (_securityForReceivedMail.IsWarnBeforeOpeningAttachmentsThatContainMacros)
                {
                    if (OfficeFileHandler.CheckOfficeFileHasVbProject(tempFilePath, Path.GetExtension(attachment.FileName).ToLower()))
                    {
                        var dialogResult = MessageBox.Show(Properties.Resources.SuspiciousAttachment_macro + Environment.NewLine + Properties.Resources.OpenAttachmentWarning1 + Environment.NewLine + Environment.NewLine + attachment.FileName, Properties.Resources.OpenAttachmentWarning1, MessageBoxButton.YesNo, MessageBoxImage.Exclamation);
                        if (dialogResult == MessageBoxResult.Yes)
                        {
                            //Open file.
                        }
                        else
                        {
                            cancel = true;
                            try
                            {
                                File.Delete(tempFilePath);
                            }
                            catch (Exception ex)
                            {
                                // Log temp file cleanup error
                                System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to delete temp file (MacroFile): {ex.Message}");
                            }
                            return;
                        }
                    }
                }

                try
                {
                    File.Delete(tempFilePath);
                }
                catch (Exception ex)
                {
                    // Log temp file cleanup error
                    System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to delete temp file (Cleanup): {ex.Message}");
                }
            }
        }

        /// <summary>
        /// 送信トレイのメールアイテムを開く際の警告。
        /// </summary>
        /// <param name="inspector">Inspector</param>
        private void OpenOutboxItemInspector(Outlook.Inspector inspector)
        {
            if (!(inspector.CurrentItem is Outlook.MailItem currentItem)) return;

            //送信保留中のメールのみ警告対象とする。
            if (currentItem.Submitted)
            {
                _ = MessageBox.Show(Properties.Resources.CanceledSendingMailMessage, Properties.Resources.CanceledSendingMail, MessageBoxButton.OK, MessageBoxImage.Warning);

                //再編集のため、配信指定日時をクリアする。
                currentItem.DeferredDeliveryTime = new DateTime(4501, 1, 1, 0, 0, 0);
                currentItem.Save();
            }

            ((Outlook.InspectorEvents_Event)inspector).Close += () =>
            {
                if (currentItem != null)
                {
                    _ = Marshal.ReleaseComObject(currentItem);
                    currentItem = null;
                }

                if (inspector != null)
                {
                    _ = Marshal.ReleaseComObject(inspector);
                    inspector = null;
                }
            };
        }

        // =====================================================================
        // ★★★ PHƯƠNG THỨC QUAN TRỌNG NHẤT: XỬ LÝ KHI GỬI EMAIL ★★★
        // =====================================================================

        /// <summary>
        /// XỬ LÝ KHI USER GỬI EMAIL (CLICK NÚT SEND)
        /// ==========================================
        /// Đây là logic cốt lõi của OutlookOkan. Khi user click Send:
        /// 1. Kiểm tra loại item (Mail, Meeting, Task)
        /// 2. Tạo CheckList chứa thông tin cần xác nhận
        /// 3. Hiện cửa sổ xác nhận nếu cần
        /// 4. Cho phép hoặc hủy gửi email
        /// 
        /// LƯU Ý: Nếu có lỗi, add-in sẽ hỏi user có muốn gửi không,
        /// để tránh trường hợp add-in lỗi làm user không gửi được email.
        /// </summary>
        /// <param name="item">Item đang gửi (MailItem, MeetingItem, TaskRequestItem)</param>
        /// <param name="cancel">Đặt = true để hủy gửi</param>
        private void Application_ItemSend(object item, ref bool cancel)
        {
            // -----------------------------------------------------------------
            // BƯỚC 0: KIỂM TRA CỬA SỔ OUTLOOK CÓ KHẢ DỤNG KHÔNG
            // -----------------------------------------------------------------
            try
            {
                // Lấy handle cửa sổ để set làm Owner cho dialog
                var activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
                _ = new NativeMethods(activeWindow).Handle;
            }
            catch (Exception)
            {
                // Không lấy được cửa sổ → cho phép gửi email bình thường
                // (tránh block email khi add-in có lỗi)
                return;
            }

            // -----------------------------------------------------------------
            // BƯỚC 1: BỎ QUA CÁC LOẠI MESSAGE ĐẶC BIỆT
            // -----------------------------------------------------------------

            // Bỏ qua Moderation Reply (Approve/Reject)
            // Nếu chặn những loại này, user không thể duyệt email được
            if (((dynamic)item).MessageClass == "IPM.Note.Microsoft.Approval.Reply.Approve"
                || ((dynamic)item).MessageClass == "IPM.Note.Microsoft.Approval.Reply.Reject")
                return;

            // Bỏ qua Meeting Response (Accept/Tentative/Decline)
            // Đây là phản hồi tự động khi user accept/decline meeting
            if (((dynamic)item).MessageClass == "IPM.Schedule.Meeting.Resp.Pos"   // Accept
                || ((dynamic)item).MessageClass == "IPM.Schedule.Meeting.Resp.Tent"  // Tentative
                || ((dynamic)item).MessageClass == "IPM.Schedule.Meeting.Resp.Neg")  // Decline
                return;

            // -----------------------------------------------------------------
            // BƯỚC 2: LOAD CÀI ĐẶT MỚI NHẤT
            // -----------------------------------------------------------------
            // User có thể thay đổi settings sau khi Outlook khởi động
            // nên phải load lại mỗi lần gửi email
            LoadGeneralSetting(isLaunch: false);
            if (!(_generalSetting.LanguageCode is null))
            {
                ResourceService.Instance.ChangeCulture(_generalSetting.LanguageCode);
            }

            // -----------------------------------------------------------------
            // BƯỚC 3: LOAD CÁC CÀI ĐẶT TỰ ĐỘNG
            // -----------------------------------------------------------------

            // Instantiate SettingsService to load all settings centrally
            var settingsService = new SettingsService();

            // Cài đặt tự động thêm text vào body email
            var autoAddMessageSetting = settingsService.AutoAddMessageSetting;

            // Danh sách recipients cần tự động xóa (ví dụ: bcc mặc định)
            var autoDeleteRecipients = settingsService.AutoDeleteRecipients ?? new List<AutoDeleteRecipient>();

            // -----------------------------------------------------------------
            // BƯỚC 4: XỬ LÝ CHÍNH - TẠO CHECKLIST VÀ HIỆN XÁC NHẬN
            // -----------------------------------------------------------------
            // Dùng try-catch lồng để đảm bảo:
            // - Nếu có lỗi nghiêm trọng: hỏi user có muốn gửi không
            // - Tránh trường hợp add-in lỗi làm user không gửi được email
            var type = typeof(Outlook.MailItem);
            try
            {
                // ---------------------------------------------------------
                // WORKAROUND: FIX LỖI OUTLOOK KHÔNG CẬP NHẬT BODY
                // ---------------------------------------------------------
                // Khi attach file dạng link, body không tự cập nhật
                // Trick: chèn space rồi xóa để trigger update
                try
                {
                    var mailItemWordEditor = (Word.Document)((dynamic)item).GetInspector.WordEditor;
                    var range = mailItemWordEditor.Range(0, 0);
                    range.InsertAfter(" ");
                    range = mailItemWordEditor.Range(0, 0);
                    _ = range.Delete();
                }
                catch (Exception)
                {
                    // Bỏ qua nếu không có WordEditor
                }

                // ---------------------------------------------------------
                // LẤY DANH BẠ NẾU CẦN
                // ---------------------------------------------------------
                // Chỉ lấy danh bạ nếu có bật tính năng liên quan
                // (để tránh truy cập không cần thiết)
                Outlook.MAPIFolder contacts = null;
                if (_generalSetting.IsAutoCheckRegisteredInContacts
                    || _generalSetting.IsWarningIfRecipientsIsNotRegistered
                    || _generalSetting.IsProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    contacts = Application.ActiveExplorer().Session
                        .GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                }

                // ---------------------------------------------------------
                // TẠO CHECKLIST TÙY THEO LOẠI ITEM
                // ---------------------------------------------------------
                var generateCheckList = new GenerateCheckList();
                CheckList checklist;

                switch (item)
                {
                    // ===== MAIL THÔNG THƯỜNG =====
                    case Outlook.MailItem mailItem:
                        // Xóa recipients theo cài đặt (nếu có)
                        var isRemovedOfMailItem = DeleteRecipients(
                            mailItem.Recipients, autoDeleteRecipients);

                        // Kiểm tra còn recipients không sau khi xóa
                        if (mailItem.Recipients.Count == 0)
                        {
                            _ = MessageBox.Show(
                                Properties.Resources.ErrorByAutoDeleteReRecipients,
                                Properties.Resources.AppName,
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                            cancel = true;
                            return;
                        }

                        type = typeof(Outlook.MailItem);

                        // ★ GỌI GENERATECHECKLIST - LOGIC CỐTLÕI ★
                        // Phân tích email và tạo danh sách cần kiểm tra
                        checklist = generateCheckList.GenerateCheckListFromMail(
                            mailItem, _generalSetting, contacts, autoAddMessageSetting, settingsService);

                        // Thêm cảnh báo nếu có recipient bị xóa tự động
                        if (isRemovedOfMailItem)
                        {
                            checklist.Alerts.Add(new Alert
                            {
                                AlertMessage = Properties.Resources.RemovedRecipietnsMessage,
                                IsImportant = true,
                                IsWhite = true,
                                IsChecked = true  // Đã check sẵn vì là thông báo
                            });
                        }
                        break;

                    // ===== MEETING REQUEST (NẾU CÓ BẬT XÁC NHẬN) =====
                    case Outlook.MeetingItem meetingItem
                        when _generalSetting.IsShowConfirmationAtSendMeetingRequest:

                        var isRemovedOfMeetingItem = DeleteRecipients(
                            meetingItem.Recipients, autoDeleteRecipients);

                        if (meetingItem.Recipients.Count == 0)
                        {
                            _ = MessageBox.Show(
                                Properties.Resources.ErrorByAutoDeleteReRecipients,
                                Properties.Resources.AppName,
                                MessageBoxButton.OK,
                                MessageBoxImage.Warning);
                            cancel = true;
                            return;
                        }

                        type = typeof(Outlook.MeetingItem);
                        checklist = generateCheckList.GenerateCheckListFromMail(
                            meetingItem, _generalSetting, contacts, autoAddMessageSetting, settingsService);

                        if (isRemovedOfMeetingItem)
                        {
                            checklist.Alerts.Add(new Alert
                            {
                                AlertMessage = Properties.Resources.RemovedRecipietnsMessage,
                                IsImportant = true,
                                IsWhite = true,
                                IsChecked = true
                            });
                        }
                        break;

                    // ===== MEETING REQUEST (KHÔNG BẬT XÁC NHẬN) =====
                    case Outlook.MeetingItem _:
                        return;  // Cho phép gửi ngay, không xác nhận

                    // ===== TASK REQUEST (NẾU CÓ BẬT XÁC NHẬN) =====
                    case Outlook.TaskRequestItem taskRequestItem
                        when _generalSetting.IsShowConfirmationAtSendTaskRequest:

                        type = typeof(Outlook.TaskRequestItem);
                        checklist = generateCheckList.GenerateCheckListFromMail(
                            taskRequestItem, _generalSetting, contacts, autoAddMessageSetting, settingsService);
                        break;

                    // ===== TASK REQUEST (KHÔNG BẬT XÁC NHẬN) =====
                    case Outlook.TaskRequestItem _:
                        return;  // Cho phép gửi ngay

                    // ===== LOẠI KHÁC (Contact, Note, v.v.) =====
                    default:
                        return;  // Cho phép gửi ngay
                }

                // ---------------------------------------------------------
                // TỰ ĐỘNG CHECK CÁC ĐỊA CHỈ CÙNG DOMAIN
                // ---------------------------------------------------------
                // Nếu bật: tự động check các địa chỉ nội bộ (cùng domain)
                if (_generalSetting.IsAutoCheckIfAllRecipientsAreSameDomain)
                {
                    // Check sẵn các địa chỉ To không phải external
                    foreach (var to in checklist.ToAddresses.Where(to => !to.IsExternal))
                    {
                        to.IsChecked = true;
                    }

                    // Check sẵn các địa chỉ Cc không phải external
                    foreach (var cc in checklist.CcAddresses.Where(cc => !cc.IsExternal))
                    {
                        cc.IsChecked = true;
                    }

                    // Check sẵn các địa chỉ Bcc không phải external
                    foreach (var bcc in checklist.BccAddresses.Where(bcc => !bcc.IsExternal))
                    {
                        bcc.IsChecked = true;
                    }
                }

                if (_generalSetting.IsEnableRecipientsAreSortedByDomain)
                {
                    checklist.ToAddresses = checklist.ToAddresses.OrderBy(x => x.MailAddress.Substring((int)Math.Sqrt(Math.Pow(x.MailAddress.IndexOf("@", StringComparison.Ordinal), 2)))).ToList();
                    checklist.CcAddresses = checklist.CcAddresses.OrderBy(x => x.MailAddress.Substring((int)Math.Sqrt(Math.Pow(x.MailAddress.IndexOf("@", StringComparison.Ordinal), 2)))).ToList();
                    checklist.BccAddresses = checklist.BccAddresses.OrderBy(x => x.MailAddress.Substring((int)Math.Sqrt(Math.Pow(x.MailAddress.IndexOf("@", StringComparison.Ordinal), 2)))).ToList();
                }

                if (checklist.IsCanNotSendMail)
                {
                    //送信禁止条件に該当するため、確認画面を表示するのではなく、送信禁止画面を表示する。
                    //このタイミングで落ちると、メールが送信されてしまうので、念のためのTry Catch。
                    try
                    {
                        _ = MessageBox.Show(checklist.CanNotSendMailMessage, Properties.Resources.SendingForbid, MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                    catch (Exception ex)
                    {
                        //Do nothing.
                        System.Diagnostics.Debug.WriteLine($"Error showing sending forbid message: {ex}");
                    }
                    finally
                    {
                        cancel = true;
                    }

                    cancel = true;
                }
                else if (IsShowConfirmationWindow(checklist))
                {
                    //OutlookのWindowを親として確認画面をモーダル表示。
                    var confirmationWindow = new ConfirmationWindow(checklist, item);
                    var activeWindow = Globals.ThisAddIn.Application.ActiveWindow();
                    var outlookHandle = new NativeMethods(activeWindow).Handle;
                    _ = new WindowInteropHelper(confirmationWindow) { Owner = outlookHandle };

                    var dialogResult = confirmationWindow.ShowDialog() ?? false;

                    if (dialogResult)
                    {
                        //メール本文への文言の自動追加はメール送信時に実行する。
                        AutoAddMessageToBody(autoAddMessageSetting, item, type == typeof(Outlook.MailItem));

                        //Send Mail.
                    }
                    else
                    {
                        cancel = true;
                    }
                }
                else
                {
                    //メール本文への文言の自動追加はメール送信時に実行する。
                    AutoAddMessageToBody(autoAddMessageSetting, item, type == typeof(Outlook.MailItem));

                    //Send Mail.
                }
            }
            catch (Exception e)
            {
                var dialogResult = MessageBox.Show(Properties.Resources.IsCanNotShowConfirmation + Environment.NewLine + Properties.Resources.SendMailConfirmation + Environment.NewLine + Environment.NewLine + e.Message, Properties.Resources.IsCanNotShowConfirmation, MessageBoxButton.YesNo, MessageBoxImage.Warning);
                if (dialogResult == MessageBoxResult.Yes)
                {
                    //メール本文への文言の自動追加はメール送信時に実行する。
                    AutoAddMessageToBody(autoAddMessageSetting, item, type == typeof(Outlook.MailItem));

                    //Send Mail.
                }
                else
                {
                    cancel = true;
                }
            }
        }

        /// <summary>
        /// 受信メールに関するセキュリティ機能の設定を読み込む
        /// </summary>
        private void LoadSecurityForReceivedMail()
        {
            var securityForReceivedMail = CsvFileHandler.ReadCsv<SecurityForReceivedMail>(typeof(SecurityForReceivedMailMap), "SecurityForReceivedMail.csv").ToList();
            if (securityForReceivedMail.Count == 0) return;

            _securityForReceivedMail.IsEnableSecurityForReceivedMail = securityForReceivedMail[0].IsEnableSecurityForReceivedMail;
            _securityForReceivedMail.IsEnableAlertKeywordOfSubjectWhenOpeningMailsData = securityForReceivedMail[0].IsEnableAlertKeywordOfSubjectWhenOpeningMailsData;
            _securityForReceivedMail.IsEnableMailHeaderAnalysis = securityForReceivedMail[0].IsEnableMailHeaderAnalysis;
            _securityForReceivedMail.IsShowWarningWhenSpfFails = securityForReceivedMail[0].IsShowWarningWhenSpfFails;
            _securityForReceivedMail.IsShowWarningWhenDkimFails = securityForReceivedMail[0].IsShowWarningWhenDkimFails;
            _securityForReceivedMail.IsEnableWarningFeatureWhenOpeningAttachments = securityForReceivedMail[0].IsEnableWarningFeatureWhenOpeningAttachments;
            _securityForReceivedMail.IsWarnBeforeOpeningAttachments = securityForReceivedMail[0].IsWarnBeforeOpeningAttachments;
            _securityForReceivedMail.IsWarnBeforeOpeningEncryptedZip = securityForReceivedMail[0].IsWarnBeforeOpeningEncryptedZip;
            _securityForReceivedMail.IsWarnLinkFileInTheZip = securityForReceivedMail[0].IsWarnLinkFileInTheZip;
            _securityForReceivedMail.IsWarnOneFileInTheZip = securityForReceivedMail[0].IsWarnOneFileInTheZip;
            _securityForReceivedMail.IsWarnOfficeFileWithMacroInTheZip = securityForReceivedMail[0].IsWarnOfficeFileWithMacroInTheZip;
            _securityForReceivedMail.IsWarnBeforeOpeningAttachmentsThatContainMacros = securityForReceivedMail[0].IsWarnBeforeOpeningAttachmentsThatContainMacros;
            _securityForReceivedMail.IsShowWarningWhenSpoofingRisk = securityForReceivedMail[0].IsShowWarningWhenSpoofingRisk;
            _securityForReceivedMail.IsShowWarningWhenDmarcNotImplemented = securityForReceivedMail[0].IsShowWarningWhenDmarcNotImplemented;
        }

        /// <summary>
        /// 受信したメールの件名の警告対象となる設定を読み込む。
        /// </summary>
        private void LoadAlertKeywordOfSubjectWhenOpeningMailsData()
        {
            var alertKeywordOfSubjectWhenOpeningMails = CsvFileHandler.ReadCsv<AlertKeywordOfSubjectWhenOpeningMail>(typeof(AlertKeywordOfSubjectWhenOpeningMailMap), "AlertKeywordOfSubjectWhenOpeningMailList.csv").Where(x => !string.IsNullOrEmpty(x.AlertKeyword));
            _alertKeywordOfSubjectWhenOpeningMail.AddRange(alertKeywordOfSubjectWhenOpeningMails);
        }

        /// <summary>
        /// 一般設定を設定ファイルから読み込む。
        /// </summary>
        /// <param name="isLaunch">Outlookの起動時か否か</param>
        private void LoadGeneralSetting(bool isLaunch)
        {
            var generalSetting = CsvFileHandler.ReadCsv<GeneralSetting>(typeof(GeneralSettingMap), "GeneralSetting.csv").ToList();
            if (generalSetting.Count == 0) return;

            _generalSetting.LanguageCode = generalSetting[0].LanguageCode;

            if (isLaunch) return;

            _generalSetting.EnableForgottenToAttachAlert = generalSetting[0].EnableForgottenToAttachAlert;
            _generalSetting.IsDoNotConfirmationIfAllRecipientsAreSameDomain = generalSetting[0].IsDoNotConfirmationIfAllRecipientsAreSameDomain;
            _generalSetting.IsDoDoNotConfirmationIfAllWhite = generalSetting[0].IsDoDoNotConfirmationIfAllWhite;
            _generalSetting.IsAutoCheckIfAllRecipientsAreSameDomain = generalSetting[0].IsAutoCheckIfAllRecipientsAreSameDomain;
            _generalSetting.IsShowConfirmationToMultipleDomain = generalSetting[0].IsShowConfirmationToMultipleDomain;
            _generalSetting.EnableGetContactGroupMembers = generalSetting[0].EnableGetContactGroupMembers;
            _generalSetting.EnableGetExchangeDistributionListMembers = generalSetting[0].EnableGetExchangeDistributionListMembers;
            _generalSetting.ContactGroupMembersAreWhite = generalSetting[0].ContactGroupMembersAreWhite;
            _generalSetting.ExchangeDistributionListMembersAreWhite = generalSetting[0].ExchangeDistributionListMembersAreWhite;
            _generalSetting.IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles = generalSetting[0].IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles;
            _generalSetting.IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain = generalSetting[0].IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain;
            _generalSetting.IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain = generalSetting[0].IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain;
            _generalSetting.IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain = generalSetting[0].IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain;
            _generalSetting.IsEnableRecipientsAreSortedByDomain = generalSetting[0].IsEnableRecipientsAreSortedByDomain;
            _generalSetting.IsAutoAddSenderToBcc = generalSetting[0].IsAutoAddSenderToBcc;
            _generalSetting.IsAutoCheckRegisteredInContacts = generalSetting[0].IsAutoCheckRegisteredInContacts;
            _generalSetting.IsAutoCheckRegisteredInContactsAndMemberOfContactLists = generalSetting[0].IsAutoCheckRegisteredInContactsAndMemberOfContactLists;
            _generalSetting.IsCheckNameAndDomainsFromRecipients = generalSetting[0].IsCheckNameAndDomainsFromRecipients;
            _generalSetting.IsWarningIfRecipientsIsNotRegistered = generalSetting[0].IsWarningIfRecipientsIsNotRegistered;
            _generalSetting.IsProhibitsSendingMailIfRecipientsIsNotRegistered = generalSetting[0].IsProhibitsSendingMailIfRecipientsIsNotRegistered;
            _generalSetting.IsShowConfirmationAtSendMeetingRequest = generalSetting[0].IsShowConfirmationAtSendMeetingRequest;
            _generalSetting.IsAutoAddSenderToCc = generalSetting[0].IsAutoAddSenderToCc;
            _generalSetting.IsCheckNameAndDomainsIncludeSubject = generalSetting[0].IsCheckNameAndDomainsIncludeSubject;
            _generalSetting.IsCheckNameAndDomainsFromSubject = generalSetting[0].IsCheckNameAndDomainsFromSubject;
            _generalSetting.IsShowConfirmationAtSendTaskRequest = generalSetting[0].IsShowConfirmationAtSendTaskRequest;
            _generalSetting.IsAutoCheckAttachments = generalSetting[0].IsAutoCheckAttachments;
            _generalSetting.IsCheckKeywordAndRecipientsIncludeSubject = generalSetting[0].IsCheckKeywordAndRecipientsIncludeSubject;
        }

        /// <summary>
        /// 全てのチェック対象がチェックされているか否かの判定。(ホワイトリスト登録の宛先など、事前にチェックされている場合がある)
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <returns>全てのチェック対象がチェックされているか否か</returns>
        private bool IsAllChecked(CheckList checkList)
        {
            var isToAddressesCompleteChecked = checkList.ToAddresses.Count(x => x.IsChecked) == checkList.ToAddresses.Count;
            var isCcAddressesCompleteChecked = checkList.CcAddresses.Count(x => x.IsChecked) == checkList.CcAddresses.Count;
            var isBccAddressesCompleteChecked = checkList.BccAddresses.Count(x => x.IsChecked) == checkList.BccAddresses.Count;
            var isAlertsCompleteChecked = checkList.Alerts.Count(x => x.IsChecked) == checkList.Alerts.Count;
            var isAttachmentsCompleteChecked = checkList.Attachments.Count(x => x.IsChecked) == checkList.Attachments.Count;

            return isToAddressesCompleteChecked && isCcAddressesCompleteChecked && isBccAddressesCompleteChecked && isAlertsCompleteChecked && isAttachmentsCompleteChecked;
        }

        /// <summary>
        /// 全ての宛先が内部(社内)ドメインであるか否かの判定。
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <returns>全ての宛先が内部(社内)ドメインであるか否か</returns>
        private bool IsAllRecipientsAreSameDomain(CheckList checkList)
        {
            var isAllToRecipientsAreSameDomain = checkList.ToAddresses.Count(x => !x.IsExternal) == checkList.ToAddresses.Count;
            var isAllCcRecipientsAreSameDomain = checkList.CcAddresses.Count(x => !x.IsExternal) == checkList.CcAddresses.Count;
            var isAllBccRecipientsAreSameDomain = checkList.BccAddresses.Count(x => !x.IsExternal) == checkList.BccAddresses.Count;

            return isAllToRecipientsAreSameDomain && isAllCcRecipientsAreSameDomain && isAllBccRecipientsAreSameDomain;
        }

        /// <summary>
        /// 送信前の確認画面の表示有無を判定。
        /// </summary>
        /// <param name="checklist">CheckList</param>
        /// <returns>送信前の確認画面の表示有無</returns>
        private bool IsShowConfirmationWindow(CheckList checklist)
        {
            if (checklist.RecipientExternalDomainNumAll >= 2 && _generalSetting.IsShowConfirmationToMultipleDomain)
            {
                //全ての宛先が確認対象だが、複数のドメインが宛先に含まれる場合は確認画面を表示するオプションが有効かつその状態のため、スキップしない。
                //他の判定より優先されるため、常に先に確認して、先にreturnする。
                return true;
            }

            if (_generalSetting.IsDoNotConfirmationIfAllRecipientsAreSameDomain && IsAllRecipientsAreSameDomain(checklist))
            {
                //全ての受信者が送信者と同一ドメインの場合に確認画面を表示しないオプションが有効かつその状態のためスキップ。
                return false;
            }

            if (checklist.ToAddresses.Count(x => x.IsSkip) == checklist.ToAddresses.Count && checklist.CcAddresses.Count(x => x.IsSkip) == checklist.CcAddresses.Count && checklist.BccAddresses.Count(x => x.IsSkip) == checklist.BccAddresses.Count)
            {
                //全ての宛先が確認画面スキップ対象のためスキップ。
                return false;
            }

            if (_generalSetting.IsDoDoNotConfirmationIfAllWhite && IsAllChecked(checklist))
            {
                //全てにチェックが入った状態の場合に確認画面を表示しないオプションが有効かつその状態のためスキップ。
                return false;
            }

            //どのようなオプション条件にも該当しないため、通常通り確認画面を表示する。
            return true;
        }

        /// <summary>
        /// メール本文へ文言を自動追加する。
        /// </summary>
        /// <param name="autoAddMessageSetting">自動追加する文言の設定</param>
        /// <param name="item">mailItem</param>
        /// <param name="isMailItem">mailItemか否か</param>
        private void AutoAddMessageToBody(AutoAddMessage autoAddMessageSetting, object item, bool isMailItem)
        {
            //一旦、通常のメールのみ対象とする。
            if (!isMailItem) return;

            if (autoAddMessageSetting.IsAddToStart)
            {
                var mailItemWordEditor = (Word.Document)((dynamic)item).GetInspector.WordEditor;
                var range = mailItemWordEditor.Range(0, 0);
                range.InsertBefore(autoAddMessageSetting.MessageOfAddToStart + Environment.NewLine + Environment.NewLine);
            }

            if (autoAddMessageSetting.IsAddToEnd)
            {
                var mailItemWordEditor = (Word.Document)((dynamic)item).GetInspector.WordEditor;
                var range = mailItemWordEditor.Range();
                range.InsertAfter(Environment.NewLine + Environment.NewLine + autoAddMessageSetting.MessageOfAddToEnd);
            }
        }

        /// <summary>
        /// 受信メールの宛先を削除する。
        /// </summary>
        /// <param name="recipients">mailItem.Recipients</param>
        /// <param name="autoDeleteRecipients">対象のドメインやメールアドレス</param>
        private bool DeleteRecipients(Outlook.Recipients recipients, List<AutoDeleteRecipient> autoDeleteRecipients)
        {
            var isRemoved = false;
            if (recipients is null || autoDeleteRecipients is null || !autoDeleteRecipients.Any())
            {
                return false;
            }

            for (var i = recipients.Count; i >= 1; i--)
            {
                var recipient = recipients[i];
                var address = recipient.Address.ToLower();

                foreach (var recipientToDelete in autoDeleteRecipients.Select(settings => settings.Recipient.ToLower()))
                {
                    if (recipientToDelete.StartsWith("@") && address.EndsWith(recipientToDelete))
                    {
                        recipients.Remove(i);
                        isRemoved = true;
                        break;
                    }

                    if (address.Equals(recipientToDelete))
                    {
                        recipients.Remove(i);
                        isRemoved = true;
                        break;
                    }
                }
            }

            if (!isRemoved) return false;

            recipients.ResolveAll();
            return true;
        }

        #region VSTO generated code

        private void InternalStartup() => Startup += ThisAddIn_Startup;

        #endregion
    }
}