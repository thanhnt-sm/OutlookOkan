// ============================================================================
// GENERATECHECKLIST - LOGIC CỐT LÕI CỦA OUTLOOKOKAN
// ============================================================================
// File: GenerateCheckList.cs
// Mô tả: Phân tích email trước khi gửi và tạo danh sách các mục cần kiểm tra
// Kích thước: 2383 dòng (cần refactor trong tương lai)
// ============================================================================

// --- CÁC THƯ VIỆN SỬ DỤNG ---
using OutlookOkan.Properties;
using OutlookOkan.Types;
using OutlookOkan.Services;    // [NEW]
using OutlookOkan.Helpers;     // [NEW]
using OutlookOkan.Handlers;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan.Models
{
    // =========================================================================
    // LỚP GENERATECHECKLIST - TẠO DANH SÁCH KIỂM TRA TRƯỚC KHI GỬI EMAIL
    // =========================================================================
    // 
    // ĐÂY LÀ LỚP QUAN TRỌNG NHẤT CỦA OUTLOOKOKAN!
    // 
    // CHỨC NĂNG CHÍNH:
    // 1. Phân tích email (recipients, attachments, body, subject)
    // 2. Kiểm tra các quy tắc (whitelist, keywords, domains)
    // 3. Tự động thêm CC/BCC nếu cần
    // 4. Tạo CheckList chứa các cảnh báo và mục cần xác nhận
    //
    // LƯU Ý VỀ CODE QUALITY:
    // - File này có 2383 dòng → nên được refactor thành nhiều class nhỏ hơn
    // - Có nhiều Thread.Sleep() để xử lý lỗi COM → không tối ưu
    // - Có nhiều try-catch rỗng → cần thêm logging
    // =========================================================================
    public sealed class GenerateCheckList
    {
        // =====================================================================
        // CÁC BIẾN INSTANCE
        // =====================================================================

        /// <summary>
        /// CheckList kết quả - chứa tất cả thông tin cần hiển thị trong cửa sổ xác nhận
        /// </summary>
        private CheckList _checkList = new CheckList();

        /// <summary>
        /// Danh sách địa chỉ được phép (whitelist)
        /// Các địa chỉ trong whitelist sẽ được tự động check trong cửa sổ xác nhận
        /// </summary>
        private List<Whitelist> _whitelist;

        /// <summary>
        /// Bộ đếm để tạo ID duy nhất cho các địa chỉ không lấy được thông tin
        /// Ví dụ: "FailedToGetInformation_1", "FailedToGetInformation_2", v.v.
        /// </summary>
        private int _failedToGetInformationOfRecipientsMailAddressCounter;

        /// <summary>
        /// Cờ đánh dấu item đang xử lý là Meeting Request
        /// Meeting Request có cấu trúc khác với Mail thông thường
        /// </summary>
        public bool IsMeetingItem;

        /// <summary>
        /// Cờ đánh dấu item đang xử lý là Task Request
        /// Task Request cần lấy thông tin từ AssociatedTask
        /// </summary>
        public bool IsTaskRequestItem;

        // =====================================================================
        // PHƯƠNG THỨC CHÍNH: TẠO CHECKLIST TỪ EMAIL
        // =====================================================================

        /// <summary>
        /// TẠO DANH SÁCH KIỂM TRA TỪ EMAIL (PHƯƠNG THỨC QUAN TRỌNG NHẤT)
        /// ==============================================================
        /// Đây là entry point chính của class, được gọi từ ThisAddIn.Application_ItemSend()
        /// 
        /// TRÌNH TỰ XỬ LÝ:
        /// 1. Load tất cả settings từ CSV files
        /// 2. Lấy thông tin email (subject, body, sender)
        /// 3. Lấy thông tin attachments
        /// 4. Kiểm tra quên đính kèm file
        /// 5. Kiểm tra từ khóa cảnh báo
        /// 6. Phân tích danh sách recipients (To, Cc, Bcc)
        /// 7. Tự động thêm CC/BCC nếu cần
        /// 8. Kiểm tra các quy tắc (domains, contacts, v.v.)
        /// 9. Trả về CheckList hoàn chỉnh
        /// </summary>
        /// <typeparam name="T">Loại item (MailItem, MeetingItem, TaskRequestItem)</typeparam>
        /// <param name="item">Item email đang gửi</param>
        /// <param name="generalSetting">Cài đặt chung của add-in</param>
        /// <param name="contacts">Folder danh bạ (có thể null nếu không cần)</param>
        /// <param name="autoAddMessageSetting">Cài đặt tự động thêm text vào body</param>
        /// <returns>CheckList chứa tất cả thông tin và cảnh báo</returns>
        internal CheckList GenerateCheckListFromMail<T>(T item, GeneralSetting generalSetting, Outlook.MAPIFolder contacts, AutoAddMessage autoAddMessageSetting, SettingsService settingsService)
        {
            // Initialize local whitelist from settings (copy to allow modification during processing)
            _whitelist = new List<Whitelist>(settingsService.Whitelist);

            switch (item)
            {
                case Outlook.MailItem mailItem:
                    _checkList.MailType = GetMailBodyFormat(mailItem.BodyFormat) ?? Resources.FailedToGetInformation;
                    _checkList.MailBody = GetMailBody(mailItem.BodyFormat, mailItem.Body ?? Resources.FailedToGetInformation);
                    _checkList.MailBody = AddMessageToBodyPreview(_checkList.MailBody, autoAddMessageSetting);

                    _checkList.MailHtmlBody = mailItem.HTMLBody ?? Resources.FailedToGetInformation;
                    break;
                case Outlook.MeetingItem meetingItem:
                    IsMeetingItem = true;
                    _checkList.MailType = Resources.MeetingRequest;
                    _checkList.MailBody = string.IsNullOrEmpty(meetingItem.Body) ? Resources.FailedToGetInformation : meetingItem.Body.Replace("\r\n\r\n", "\r\n");

                    if (meetingItem.RTFBody is byte[] byteArray)
                    {
                        var encoding = new System.Text.ASCIIEncoding();
                        _checkList.MailHtmlBody = encoding.GetString(byteArray);
                    }
                    else
                    {
                        _checkList.MailHtmlBody = _checkList.MailBody;
                    }
                    break;
                case Outlook.TaskRequestItem taskRequestItem:
                    IsTaskRequestItem = true;
                    _checkList.MailType = Resources.TaskRequest;

                    var associatedTask = taskRequestItem.GetAssociatedTask(false);
                    Thread.Sleep(10);
                    _checkList.MailBody = string.IsNullOrEmpty(associatedTask.Body) ? Resources.FailedToGetInformation : associatedTask.Body.Replace("\r\n\r\n", "\r\n");

                    if (associatedTask.RTFBody is byte[] bodyByteArray)
                    {
                        var encoding = new System.Text.ASCIIEncoding();
                        _checkList.MailHtmlBody = encoding.GetString(bodyByteArray);
                    }
                    else
                    {
                        _checkList.MailHtmlBody = _checkList.MailBody;
                    }
                    break;
            }

            _checkList.Subject = ((dynamic)item).Subject ?? Resources.FailedToGetInformation;

            _checkList = GetSenderAndSenderDomain(in item, _checkList);
            settingsService.InternalDomainList.Add(new InternalDomain { Domain = _checkList.SenderDomain });

            _checkList = GetAttachmentsInformation(in item, _checkList, generalSetting.IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles, settingsService.AttachmentsSetting, _checkList.MailHtmlBody, generalSetting.IsAutoCheckAttachments);
            _checkList = CheckForgotAttach(_checkList, generalSetting);
            _checkList = CheckKeyword(_checkList, settingsService.AlertKeywordAndMessageList);
            _checkList = CheckKeywordForSubject(_checkList, settingsService.AlertKeywordAndMessageForSubjectList);

            var displayNameAndRecipient = IsTaskRequestItem ? MakeDisplayNameAndRecipient(((Outlook.TaskRequestItem)item).GetAssociatedTask(false).Recipients, new DisplayNameAndRecipient(), generalSetting, false) : (DisplayNameAndRecipient)MakeDisplayNameAndRecipient(((dynamic)item).Recipients, new DisplayNameAndRecipient(), generalSetting, IsMeetingItem);

            var autoAddRecipients = AutoAddCcAndBcc(item, generalSetting, displayNameAndRecipient, settingsService.AutoCcBccKeywordList, settingsService.AutoCcBccAttachedFilesList, settingsService.AutoCcBccRecipientList, CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, settingsService.InternalDomainList, false), _checkList.Sender, generalSetting.IsAutoAddSenderToBcc, generalSetting.IsAutoAddSenderToCc);
            if (autoAddRecipients?.Count > 0)
            {
                displayNameAndRecipient = MakeDisplayNameAndRecipient(autoAddRecipients, displayNameAndRecipient, generalSetting, IsMeetingItem);
                _ = ((dynamic)item).Recipients.ResolveAll();
            }

            displayNameAndRecipient = ExternalDomainsChangeToBccIfNeeded(item, displayNameAndRecipient, settingsService.ExternalDomainsWarningAndAutoChangeToBccSetting, settingsService.InternalDomainList, CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, settingsService.InternalDomainList, true), _checkList.SenderDomain, _checkList.Sender, settingsService.ForceAutoChangeRecipientsToBccSetting);

            _checkList = GetRecipient(_checkList, displayNameAndRecipient, settingsService.AlertAddressList, settingsService.InternalDomainList);
            _checkList = CheckRecipientsAndAttachments(_checkList, settingsService.AttachmentsSetting.IsAttachmentsProhibited, settingsService.AttachmentsSetting.IsWarningWhenAttachedRealFile, settingsService.AttachmentProhibitedRecipientsList, settingsService.RecipientsAndAttachmentsNameList, settingsService.AttachmentAlertRecipientsList);
            _checkList = CheckMailBodyAndRecipient(_checkList, displayNameAndRecipient, settingsService.NameAndDomainsList, generalSetting.IsCheckNameAndDomainsFromRecipients, generalSetting.IsCheckNameAndDomainsIncludeSubject, generalSetting.IsCheckNameAndDomainsFromSubject);
            _checkList = CheckKeywordAndRecipient(_checkList, displayNameAndRecipient, settingsService.KeywordAndRecipientsList, generalSetting.IsCheckKeywordAndRecipientsIncludeSubject);
            _checkList.RecipientExternalDomainNumAll = CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, settingsService.InternalDomainList, false);
            _checkList = ExternalDomainsWarningIfNeeded(_checkList, settingsService.ExternalDomainsWarningAndAutoChangeToBccSetting, CountRecipientExternalDomains(displayNameAndRecipient, _checkList.SenderDomain, settingsService.InternalDomainList, true), settingsService.ForceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc);
            _checkList.DeferredMinutes = CalcDeferredMinutes(displayNameAndRecipient, settingsService.DeferredDeliveryMinutesList, generalSetting.IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain, _checkList.RecipientExternalDomainNumAll);

            if (!(contacts is null))
            {
                var contactsList = MakeContactsList(contacts);
                _checkList = AutoCheckRegisteredItemsInContacts(_checkList, displayNameAndRecipient, contactsList, generalSetting.IsAutoCheckRegisteredInContacts);
                _checkList = AddAlertOrProhibitsSendingMailIfIfRecipientsIsNotRegistered(_checkList, displayNameAndRecipient, contactsList, settingsService.InternalDomainList, generalSetting.IsWarningIfRecipientsIsNotRegistered, generalSetting.IsProhibitsSendingMailIfRecipientsIsNotRegistered);
            }

            if (settingsService.AttachmentsSetting.IsIgnoreMustOpenBeforeCheckTheAttachedFilesIfInternalDomain && _checkList.Attachments.Any() && _checkList.RecipientExternalDomainNumAll == 0)
            {
                foreach (var attachment in _checkList.Attachments)
                {
                    attachment.IsNotMustOpenBeforeCheck = true;
                }
            }

            return _checkList;
        }

        /// <summary>
        /// Lấy địa chỉ người gửi và tên miền người gửi.
        /// </summary>
        /// <param name="item">Item</param>
        /// <param name="checkList">CheckList</param>
        /// <returns>CheckList</returns>
        private CheckList GetSenderAndSenderDomain<T>(in T item, CheckList checkList)
        {
            try
            {
                if (typeof(T) == typeof(Outlook.MailItem) && !string.IsNullOrEmpty(((Outlook.MailItem)item).SentOnBehalfOfName))
                {
                    // Trường hợp gửi thay mặt.
                    checkList.Sender = ((Outlook.MailItem)item).Sender?.Address ?? Resources.FailedToGetInformation;

                    if (IsValidEmailAddress(checkList.Sender))
                    {
                        // Nếu có thể lấy được địa chỉ email thì sử dụng như bình thường.
                        checkList.SenderDomain = checkList.Sender.Substring(checkList.Sender.IndexOf("@", StringComparison.Ordinal));
                        checkList.Sender = $@"{checkList.Sender} ([{((Outlook.MailItem)item).SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                    }
                    else
                    {
                        // Trường hợp gửi thay mặt và là CN của Exchange.
                        checkList.Sender = $@"[{((Outlook.MailItem)item).SentOnBehalfOfName}] {Resources.SentOnBehalf}";
                        checkList.SenderDomain = @"------------------";

                        Outlook.ExchangeDistributionList exchangeDistributionList = null;
                        Outlook.ExchangeUser exchangeUser = null;

                        var sender = ((Outlook.MailItem)item).Sender;

                        ComRetryHelper.Execute(() =>
                        {
                            exchangeDistributionList = sender?.GetExchangeDistributionList();
                            exchangeUser = sender?.GetExchangeUser();
                        });

                        if (!(exchangeUser is null))
                        {
                            // Người dùng gửi thay mặt.
                            checkList.Sender = $@"{exchangeUser.PrimarySmtpAddress} ([{((Outlook.MailItem)item).SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                            checkList.SenderDomain = exchangeUser.PrimarySmtpAddress.Substring(exchangeUser.PrimarySmtpAddress.IndexOf("@", StringComparison.Ordinal));
                        }

                        if (!(exchangeDistributionList is null))
                        {
                            // Danh sách phân phối gửi thay mặt.
                            checkList.Sender = $@"{exchangeDistributionList.PrimarySmtpAddress} ([{((Outlook.MailItem)item).SentOnBehalfOfName}] {Resources.SentOnBehalf})";
                            checkList.SenderDomain = exchangeDistributionList.PrimarySmtpAddress.Substring(exchangeDistributionList.PrimarySmtpAddress.IndexOf("@", StringComparison.Ordinal));
                        }
                    }
                }
                else
                {
                    checkList.Sender = ((dynamic)item).SendUsingAccount?.SmtpAddress ?? Resources.FailedToGetInformation;

                    if (((dynamic)item).SenderEmailType == "EX" && !IsValidEmailAddress(checkList.Sender))
                    {
                        var tempOutlookApp = new Outlook.Application();
                        var tempRecipient = tempOutlookApp.Session.CreateRecipient(((dynamic)item).SenderEmailAddress);

                        _ = tempRecipient.Resolve();
                        Thread.Sleep(10);
                        var addressEntry = tempRecipient.AddressEntry;

                        ComRetryHelper.Execute(() =>
                        {
                            var exchangeUser = addressEntry?.GetExchangeUser();
                            checkList.Sender = exchangeUser?.PrimarySmtpAddress ?? Resources.FailedToGetInformation;
                        });
                    }
                    else
                    {
                        if (!IsValidEmailAddress(checkList.Sender))
                        {
                            checkList.Sender = ((dynamic)item).SenderEmailAddress ?? Resources.FailedToGetInformation;
                        }
                    }

                    if (!IsValidEmailAddress(checkList.Sender))
                    {
                        checkList.Sender = Resources.FailedToGetInformation;
                    }

                    checkList.SenderDomain = checkList.Sender == Resources.FailedToGetInformation ? "------------------" : checkList.Sender.Substring(checkList.Sender.IndexOf("@", StringComparison.Ordinal));
                }
            }
            catch (Exception)
            {
                try
                {
                    if (IsTaskRequestItem)
                    {
                        if (item is Outlook.TaskRequestItem taskRequest)
                        {
                            var associatedTask = taskRequest.GetAssociatedTask(false);

                            if (associatedTask != null)
                            {
                                var senderAddress = associatedTask.SendUsingAccount?.SmtpAddress;
                                checkList.Sender = senderAddress ?? Resources.FailedToGetInformation;

                                Marshal.ReleaseComObject(associatedTask);
                                checkList.SenderDomain = checkList.Sender == Resources.FailedToGetInformation ? "------------------" : checkList.Sender.Substring(checkList.Sender.IndexOf("@", StringComparison.Ordinal));
                            }
                        }
                    }
                    else
                    {
                        checkList.Sender = Resources.FailedToGetInformation;
                        checkList.SenderDomain = @"------------------";
                    }
                }
                catch (Exception)
                {
                    checkList.Sender = Resources.FailedToGetInformation;
                    checkList.SenderDomain = @"------------------";
                }
            }

            return checkList;
        }

        /// <summary>
        /// Lấy nội dung email dưới dạng văn bản (text).
        /// </summary>
        /// <param name="mailBodyFormat">Định dạng nội dung email</param>
        /// <param name="mailBody">Nội dung email</param>
        /// <returns>Nội dung email (dạng văn bản)</returns>
        private string GetMailBody(Outlook.OlBodyFormat mailBodyFormat, string mailBody)
        {
            // Để tránh vấn đề xuống dòng thành 2 dòng, chỉ thay thế 2 dòng xuống dòng liên tiếp thành 1 dòng trong trường hợp định dạng HTML.
            return mailBodyFormat == Outlook.OlBodyFormat.olFormatHTML ? mailBody.Replace("\r\n\r\n", "\r\n") : mailBody;
        }

        /// <summary>
        /// Đếm số lượng tên miền bên ngoài (trừ tên miền nội bộ).
        /// </summary>
        /// <param name="displayNameAndRecipient">Địa chỉ và tên người nhận</param>
        /// <param name="senderDomain">Tên miền người gửi</param>
        /// <param name="internalDomain">Cài đặt tên miền nội bộ</param>
        /// <param name="isToAndCcOnly">Chỉ áp dụng cho To và Cc hay không</param>
        /// <returns>Số lượng tên miền bên ngoài</returns>
        private int CountRecipientExternalDomains(DisplayNameAndRecipient displayNameAndRecipient, string senderDomain, IEnumerable<InternalDomain> internalDomain, bool isToAndCcOnly)
        {
            var domainList = new HashSet<string>();

            if (isToAndCcOnly)
            {
                foreach (var recipient in displayNameAndRecipient.To.Select(mail => mail.Key).Where(recipient => recipient != Resources.FailedToGetInformation && recipient.Contains("@")))
                {
                    _ = domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }

                foreach (var recipient in displayNameAndRecipient.Cc.Select(mail => mail.Key).Where(recipient => recipient != Resources.FailedToGetInformation && recipient.Contains("@")))
                {
                    _ = domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }
            }
            else
            {
                foreach (var recipient in displayNameAndRecipient.All.Select(mail => mail.Key).Where(recipient => recipient != Resources.FailedToGetInformation && recipient.Contains("@")))
                {
                    domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }
            }

            var externalDomainsCount = domainList.Count;

            foreach (var _ in internalDomain.Where(internalDomainSetting => domainList.Any(domain => domain.EndsWith(internalDomainSetting.Domain)) && !senderDomain.EndsWith(internalDomainSetting.Domain)))
            {
                externalDomainsCount--;
            }

            // Vì là đếm số lượng tên miền bên ngoài, nên nếu bao gồm tên miền người gửi thì trừ đi.
            if (domainList.Contains(senderDomain))
            {
                return externalDomainsCount - 1;
            }

            return externalDomainsCount;
        }

        /// <summary>
        /// Lấy địa chỉ email và tên hiển thị của người nhận.
        /// </summary>
        /// <param name="recipient">Người nhận email</param>
        /// <returns>Địa chỉ email và tên hiển thị</returns>
        private IEnumerable<NameAndRecipient> GetNameAndRecipient(Outlook.Recipient recipient)
        {
            _failedToGetInformationOfRecipientsMailAddressCounter++;

            var mailAddress = Resources.FailedToGetInformation + "_" + _failedToGetInformationOfRecipientsMailAddressCounter;
            if (IsValidEmailAddress(recipient.Name))
            {
                mailAddress = recipient.Name;
            }
            else
            {
                if (IsValidEmailAddress(recipient.Address)) mailAddress = recipient.Address;
            }

            if (!IsValidEmailAddress(mailAddress))
            {
                try
                {
                    var propertyAccessor = recipient.PropertyAccessor;
                    Thread.Sleep(20);

                    // COM Retry Pattern using Helper
                    mailAddress = ComRetryHelper.Execute(() =>
                        propertyAccessor.GetProperty(Constants.PR_SMTP_ADDRESS).ToString())
                        ?? mailAddress;
                }
                catch (Exception ex)
                {
                    // Log error for debugging purposes
                    System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to get recipient info (1): {ex.Message}");
                }
            }

            if (!IsValidEmailAddress(mailAddress))
            {
                var tempOutlookApp = new Outlook.Application();
                var tempRecipient = tempOutlookApp.Session.CreateRecipient(recipient.Address);

                try
                {
                    _ = recipient.Resolve();
                    var propertyAccessor = tempRecipient.AddressEntry.PropertyAccessor;
                    Thread.Sleep(20);

                    mailAddress = ComRetryHelper.Execute(() =>
                        propertyAccessor.GetProperty(Constants.PR_SMTP_ADDRESS).ToString())
                        ?? mailAddress;
                }
                catch (Exception ex)
                {
                    // Log error for debugging purposes
                    System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to get recipient info (2): {ex.Message}");
                }
            }

            string nameAndMailAddress;
            if (string.IsNullOrEmpty(recipient.Name))
            {
                nameAndMailAddress = mailAddress ?? Resources.FailedToGetInformation;
            }
            else
            {
                nameAndMailAddress = recipient.Name.Contains($@" ({mailAddress})") ? recipient.Name : recipient.Name + $@" ({mailAddress})";
            }

            if (!IsValidEmailAddress(mailAddress)) mailAddress = nameAndMailAddress;

            return new List<NameAndRecipient> { new NameAndRecipient { MailAddress = mailAddress, NameAndMailAddress = nameAndMailAddress } };
        }

        /// <summary>
        /// Mở rộng danh sách phân phối Exchange để lấy địa chỉ email và tên hiển thị. (Không mở rộng lồng nhau)
        /// </summary>
        /// <param name="recipient">Người nhận email</param>
        /// <param name="enableGetExchangeDistributionListMembers">Cài đặt bật/tắt mở rộng danh sách phân phối</param>
        /// <param name="exchangeDistributionListMembersAreWhite">Cài đặt xem các địa chỉ được mở rộng từ danh sách phân phối có được coi là whitelist hay không</param>
        /// <returns>Địa chỉ email và tên hiển thị</returns>
        private IEnumerable<NameAndRecipient> GetExchangeDistributionListMembers(Outlook.Recipient recipient, bool enableGetExchangeDistributionListMembers, bool exchangeDistributionListMembersAreWhite)
        {
            _failedToGetInformationOfRecipientsMailAddressCounter++;
            Outlook.OlAddressEntryUserType recipientAddressEntryUserType;
            try
            {
                recipientAddressEntryUserType = recipient.AddressEntry.AddressEntryUserType;
            }
            catch (Exception)
            {
                return null;
            }

            if (recipientAddressEntryUserType != Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry) return null;

            Outlook.ExchangeDistributionList distributionList = null;
            Outlook.AddressEntries addressEntries = null;

            try
            {
                var addressEntry = recipient.AddressEntry;

                ComRetryHelper.Execute(() =>
                {
                    distributionList = addressEntry?.GetExchangeDistributionList();

                    if (enableGetExchangeDistributionListMembers)
                    {
                        addressEntries = distributionList?.GetExchangeDistributionListMembers();
                    }
                });

                if (distributionList is null) return null;

                var exchangeDistributionListMembers = new List<NameAndRecipient>();

                if (addressEntries is null || addressEntries.Count == 0)
                {
                    exchangeDistributionListMembers.Add(new NameAndRecipient { MailAddress = distributionList.PrimarySmtpAddress ?? Resources.FailedToGetInformation + "_" + _failedToGetInformationOfRecipientsMailAddressCounter, NameAndMailAddress = (distributionList.Name ?? Resources.FailedToGetInformation) + $@" ({distributionList.PrimarySmtpAddress ?? Resources.DistributionList})" });

                    return exchangeDistributionListMembers;
                }

                var externalRecipientCounter = 1;
                var tempOutlookApp = new Outlook.Application();
                foreach (Outlook.AddressEntry member in addressEntries)
                {
                    var tempRecipient = tempOutlookApp.Session.CreateRecipient(member.Address);
                    var mailAddress = Resources.FailedToGetInformation + "_" + _failedToGetInformationOfRecipientsMailAddressCounter;

                    try
                    {
                        _ = tempRecipient.Resolve();
                        var propertyAccessor = tempRecipient.AddressEntry.PropertyAccessor;
                        Thread.Sleep(20);

                        mailAddress = ComRetryHelper.Execute(() =>
                            propertyAccessor.GetProperty(Constants.PR_SMTP_ADDRESS).ToString())
                            ?? mailAddress;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error resolving recipient: {ex}");
                    }

                    // Danh sách phân phối lồng nhau gây tải lớn cho máy chủ Exchange và tốn thời gian lấy nên không mở rộng.
                    exchangeDistributionListMembers.Add(new NameAndRecipient { MailAddress = mailAddress, NameAndMailAddress = (member.Name ?? Resources.FailedToGetInformation) + $@" ({mailAddress})", IncludedGroupAndList = $@" [{distributionList.Name}]" });

                    if (exchangeDistributionListMembersAreWhite)
                    {
                        _whitelist.Add(new Whitelist { WhiteName = mailAddress });
                    }
                }

                return exchangeDistributionListMembers;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in GetExchangeDistributionListMembers: {ex}");
                return null;
            }
        }

        /// <summary>
        /// Mở rộng nhóm liên hệ để lấy địa chỉ email và tên hiển thị. (Tự động mở rộng lồng nhau.)
        /// </summary>
        /// <param name="recipient">Người nhận email</param>
        /// <param name="contactGroupId">GroupID đã kiểm tra</param>
        /// <param name="enableGetContactGroupMembers">Cài đặt bật/tắt mở rộng nhóm liên hệ</param>
        /// <param name="contactGroupMembersAreWhite">Cài đặt xem các địa chỉ được mở rộng từ nhóm liên hệ có được coi là whitelist hay không</param>
        /// <returns>Địa chỉ email và tên hiển thị</returns>
        private IEnumerable<NameAndRecipient> GetContactGroupMembers(Outlook.Recipient recipient, string contactGroupId, bool enableGetContactGroupMembers, bool contactGroupMembersAreWhite)
        {
            var contactGroupMembers = new List<NameAndRecipient>();
            if (!enableGetContactGroupMembers)
            {
                contactGroupMembers.Add(new NameAndRecipient { MailAddress = recipient.Name, NameAndMailAddress = recipient.Name + $@" [{Resources.ContactGroup}]" });
                return contactGroupMembers;
            }

            string entryId;
            if (contactGroupId is null)
            {
                var entryIdLength = Convert.ToInt32(recipient.AddressEntry.ID.Substring(66, 2) + recipient.AddressEntry.ID.Substring(64, 2), 16) * 2;
                entryId = recipient.AddressEntry.ID.Substring(72, entryIdLength);
            }
            else
            {
                // ID trong trường hợp lồng nhau.
                entryId = recipient.AddressEntry.ID.Substring(42);
            }

            if (contactGroupId?.Contains(entryId) == true) return null;

            contactGroupId = contactGroupId + entryId + ",";

            var tempOutlookApp = new Outlook.Application().GetNamespace("MAPI");
            var distList = (Outlook.DistListItem)tempOutlookApp.GetItemFromID(entryId);

            for (var i = 1; i < distList.MemberCount + 1; i++)
            {
                var member = distList.GetMember(i);
                contactGroupMembers.AddRange(member.Address == "Unknown"
                    ? GetContactGroupMembers(member, contactGroupId, true, contactGroupMembersAreWhite)
                    : GetNameAndRecipient(member));
            }

            foreach (var nameAndRecipient in contactGroupMembers)
            {
                nameAndRecipient.IncludedGroupAndList += $@" [{distList.DLName}]";

                if (contactGroupMembersAreWhite)
                {
                    _whitelist.Add(new Whitelist { WhiteName = nameAndRecipient.MailAddress });
                }
            }

            return contactGroupMembers;
        }

        /// <summary>
        /// Khớp tên hiển thị của người nhận với tên hiển thị và địa chỉ email. (Theo đặc tả của Outlook, tên hiển thị có thể không chứa địa chỉ email.)
        /// </summary>
        /// <param name="recipients">Người nhận email</param>
        /// <param name="displayNameAndRecipient">Địa chỉ và tên người nhận</param>
        /// <param name="generalSetting">Cài đặt chung</param>
        /// <param name="isMeetingItem">Có phải là item cuộc họp không</param>
        /// <returns>Địa chỉ và tên người nhận</returns>
        private DisplayNameAndRecipient MakeDisplayNameAndRecipient(IEnumerable recipients, DisplayNameAndRecipient displayNameAndRecipient, GeneralSetting generalSetting, bool isMeetingItem)
        {
            foreach (Outlook.Recipient recipient in recipients)
            {
                var recipientAddressEntryUserType = Outlook.OlAddressEntryUserType.olOtherAddressEntry;
                try
                {
                    if (isMeetingItem)
                    {
                        if (!recipient.Sendable) continue;
                    }

                    recipientAddressEntryUserType = recipient.AddressEntry.AddressEntryUserType;
                }
                catch (Exception ex)
                {
                    // Log error for debugging purposes
                    System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to get AddressEntryUserType: {ex.Message}");
                }

                var nameAndRecipients = new List<NameAndRecipient>();

                switch (recipientAddressEntryUserType)
                {
                    case Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry:
                        var exchangeMembers = GetExchangeDistributionListMembers(recipient, generalSetting.EnableGetExchangeDistributionListMembers, generalSetting.ExchangeDistributionListMembersAreWhite);
                        if (exchangeMembers is null)
                        {
                            nameAndRecipients.AddRange(GetNameAndRecipient(recipient));
                            break;
                        }
                        nameAndRecipients.AddRange(exchangeMembers);
                        break;
                    case Outlook.OlAddressEntryUserType.olOutlookDistributionListAddressEntry:
                        var addressEntryMembers = GetContactGroupMembers(recipient, null, generalSetting.EnableGetContactGroupMembers, generalSetting.ContactGroupMembersAreWhite);
                        if (addressEntryMembers is null)
                        {
                            nameAndRecipients.AddRange(GetNameAndRecipient(recipient));
                            break;
                        }
                        nameAndRecipients.AddRange(addressEntryMembers);
                        break;
                    default:
                        nameAndRecipients.AddRange(GetNameAndRecipient(recipient));
                        break;
                }

                foreach (var nameAndRecipient in nameAndRecipients)
                {
                    if (displayNameAndRecipient.All.ContainsKey(nameAndRecipient.MailAddress))
                    {
                        displayNameAndRecipient.All[nameAndRecipient.MailAddress] += nameAndRecipient.IncludedGroupAndList;
                    }
                    else
                    {
                        displayNameAndRecipient.All[nameAndRecipient.MailAddress] = nameAndRecipient.NameAndMailAddress + nameAndRecipient.IncludedGroupAndList;
                    }

                    displayNameAndRecipient.MailRecipientsIndex.Add(new MailItemsRecipientAndMailAddress
                    {
                        MailAddress = nameAndRecipient.MailAddress,
                        MailItemsRecipient = recipient.Address,
                        Type = recipient.Type
                    });

                    switch (recipient.Type)
                    {
                        case 1:
                            if (displayNameAndRecipient.To.ContainsKey(nameAndRecipient.MailAddress))
                            {
                                displayNameAndRecipient.To[nameAndRecipient.MailAddress] += nameAndRecipient.IncludedGroupAndList;
                            }
                            else
                            {
                                displayNameAndRecipient.To[nameAndRecipient.MailAddress] = nameAndRecipient.NameAndMailAddress + nameAndRecipient.IncludedGroupAndList;
                            }
                            continue;
                        case 2:
                            if (displayNameAndRecipient.Cc.ContainsKey(nameAndRecipient.MailAddress))
                            {
                                displayNameAndRecipient.Cc[nameAndRecipient.MailAddress] += nameAndRecipient.IncludedGroupAndList;
                            }
                            else
                            {
                                displayNameAndRecipient.Cc[nameAndRecipient.MailAddress] = nameAndRecipient.NameAndMailAddress + nameAndRecipient.IncludedGroupAndList;
                            }
                            continue;
                        case 3:
                            if (displayNameAndRecipient.Bcc.ContainsKey(nameAndRecipient.MailAddress))
                            {
                                displayNameAndRecipient.Bcc[nameAndRecipient.MailAddress] += nameAndRecipient.IncludedGroupAndList;
                            }
                            else
                            {
                                displayNameAndRecipient.Bcc[nameAndRecipient.MailAddress] = nameAndRecipient.NameAndMailAddress + nameAndRecipient.IncludedGroupAndList;
                            }
                            continue;
                        default:
                            continue;
                    }
                }
            }

            return displayNameAndRecipient;
        }

        /// <summary>
        /// Kiểm tra việc quên đính kèm tệp.
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="generalSetting">Cài đặt chung</param>
        /// <returns>CheckList</returns>
        private CheckList CheckForgotAttach(CheckList checkList, GeneralSetting generalSetting)
        {
            if (checkList.Attachments.Count >= 1) return checkList;

            if (!generalSetting.EnableForgottenToAttachAlert) return checkList;

            if (checkList.MailBody.ToLower().Contains(Resources.AttachmentsKeyword))
            {
                checkList.Alerts.Add(new Alert { AlertMessage = Resources.ForgottenToAttachAlert, IsImportant = true, IsWhite = false, IsChecked = false });
            }

            return checkList;
        }

        /// <summary>
        /// Lấy định dạng email và trả về chuỗi hiển thị.
        /// </summary>
        /// <param name="bodyFormat">Định dạng email</param>
        /// <returns>Định dạng email</returns>
        private string GetMailBodyFormat(Outlook.OlBodyFormat bodyFormat)
        {
            switch (bodyFormat)
            {
                case Outlook.OlBodyFormat.olFormatUnspecified:
                    return Resources.Unknown;
                case Outlook.OlBodyFormat.olFormatPlain:
                    return Resources.Text;
                case Outlook.OlBodyFormat.olFormatHTML:
                    return Resources.HTML;
                case Outlook.OlBodyFormat.olFormatRichText:
                    return Resources.RichText;
                default:
                    return Resources.Unknown;
            }
        }

        /// <summary>
        /// Nếu có từ khóa đã đăng ký trong nội dung, hiển thị thông báo cảnh báo đã đăng ký.
        /// </summary>
        /// <param name="checkList">CheckList</param>>
        /// <param name="alertKeywordAndMessageList">Cài đặt từ khóa cảnh báo</param>>
        /// <returns>CheckList</returns>
        private CheckList CheckKeyword(CheckList checkList, IReadOnlyCollection<AlertKeywordAndMessage> alertKeywordAndMessageList)
        {
            if (alertKeywordAndMessageList.Count == 0) return checkList;

            foreach (var alertKeywordAndMessage in alertKeywordAndMessageList)
            {
                if (!checkList.MailBody.Contains(alertKeywordAndMessage.AlertKeyword) && alertKeywordAndMessage.AlertKeyword != "*") continue;

                var alertMessage = Resources.DefaultAlertMessage + $"[{alertKeywordAndMessage.AlertKeyword}]";
                if (!string.IsNullOrEmpty(alertKeywordAndMessage.Message)) alertMessage = alertKeywordAndMessage.Message;

                checkList.Alerts.Add(new Alert { AlertMessage = alertMessage, IsImportant = true, IsWhite = false, IsChecked = false });

                if (!alertKeywordAndMessage.IsCanNotSend) continue;

                checkList.IsCanNotSendMail = true;
                checkList.CanNotSendMailMessage = alertMessage;
            }

            return checkList;
        }

        /// <summary>
        /// Nếu có từ khóa đã đăng ký trong tiêu đề, hiển thị thông báo cảnh báo đã đăng ký.
        /// </summary>
        /// <param name="checkList">CheckList</param>>
        /// <param name="alertKeywordAndMessageForSubjectList">Cài đặt từ khóa cảnh báo</param>>
        /// <returns>CheckList</returns>
        private CheckList CheckKeywordForSubject(CheckList checkList, IReadOnlyCollection<AlertKeywordAndMessageForSubject> alertKeywordAndMessageForSubjectList)
        {
            if (alertKeywordAndMessageForSubjectList.Count == 0) return checkList;

            foreach (var alertKeywordAndMessage in alertKeywordAndMessageForSubjectList)
            {
                if (!checkList.Subject.Contains(alertKeywordAndMessage.AlertKeyword) && alertKeywordAndMessage.AlertKeyword != "*") continue;

                var alertMessage = Resources.DefaultAlertMessage + $"[{alertKeywordAndMessage.AlertKeyword}]";
                if (!string.IsNullOrEmpty(alertKeywordAndMessage.Message)) alertMessage = alertKeywordAndMessage.Message;

                checkList.Alerts.Add(new Alert { AlertMessage = alertMessage, IsImportant = true, IsWhite = false, IsChecked = false });

                if (!alertKeywordAndMessage.IsCanNotSend) continue;

                checkList.IsCanNotSendMail = true;
                checkList.CanNotSendMailMessage = alertMessage;
            }

            return checkList;
        }

        /// <summary>
        /// Thêm người nhận vào Cc hoặc Bcc nếu thỏa mãn điều kiện.
        /// </summary>
        /// <param name="item">Item</param>
        /// <param name="generalSetting">Cài đặt chung</param>
        /// <param name="displayNameAndRecipient">Cài đặt tên hiển thị và địa chỉ</param>
        /// <param name="autoCcBccKeywordList">Cài đặt tự động thêm Cc/Bcc theo từ khóa</param>
        /// <param name="autoCcBccAttachedFilesList">Cài đặt tự động thêm Cc/Bcc theo tệp đính kèm</param>
        /// <param name="autoCcBccRecipientList">Cài đặt tự động thêm Cc/Bcc theo người nhận</param>
        /// <param name="externalDomainCount">Số lượng tên miền bên ngoài</param>
        /// <param name="sender">Sender info from CheckList</param>
        /// <param name="isAutoAddSenderToBcc">Có tự động thêm địa chỉ người gửi vào Bcc hay không</param>
        /// <param name="isAutoAddSenderToCc">Có tự động thêm địa chỉ người gửi vào Cc hay không</param>
        /// <returns>Danh sách địa chỉ được tự động thêm vào Cc hoặc Bcc</returns>
        private List<Outlook.Recipient> AutoAddCcAndBcc<T>(T item, GeneralSetting generalSetting, DisplayNameAndRecipient displayNameAndRecipient, IReadOnlyCollection<AutoCcBccKeyword> autoCcBccKeywordList, IReadOnlyCollection<AutoCcBccAttachedFile> autoCcBccAttachedFilesList, IReadOnlyCollection<AutoCcBccRecipient> autoCcBccRecipientList, int externalDomainCount, string sender, bool isAutoAddSenderToBcc, bool isAutoAddSenderToCc)
        {
            var autoAddedCcAddressList = new List<string>();
            var autoAddedBccAddressList = new List<string>();
            var autoAddRecipients = new List<Outlook.Recipient>();

            if (autoCcBccKeywordList.Count != 0 && !(externalDomainCount == 0 && generalSetting.IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain))
            {
                foreach (var autoCcBccKeyword in autoCcBccKeywordList)
                {
                    if (!_checkList.MailBody.Contains(autoCcBccKeyword.Keyword) || !autoCcBccKeyword.AutoAddAddress.Contains("@")) continue;

                    if (autoCcBccKeyword.CcOrBcc == CcOrBcc.Cc)
                    {
                        if (!autoAddedCcAddressList.Contains(autoCcBccKeyword.AutoAddAddress) && !displayNameAndRecipient.Cc.ContainsKey(autoCcBccKeyword.AutoAddAddress))
                        {
                            var recipient = ((dynamic)item).Recipients.Add(autoCcBccKeyword.AutoAddAddress);
                            recipient.Type = (int)Outlook.OlMailRecipientType.olCC;

                            autoAddRecipients.Add(recipient);
                            autoAddedCcAddressList.Add(autoCcBccKeyword.AutoAddAddress);
                        }
                    }
                    else if (!autoAddedBccAddressList.Contains(autoCcBccKeyword.AutoAddAddress) && !displayNameAndRecipient.Bcc.ContainsKey(autoCcBccKeyword.AutoAddAddress))
                    {
                        var recipient = ((dynamic)item).Recipients.Add(autoCcBccKeyword.AutoAddAddress);
                        recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;

                        autoAddRecipients.Add(recipient);
                        autoAddedBccAddressList.Add(autoCcBccKeyword.AutoAddAddress);
                    }

                    _checkList.Alerts.Add(new Alert { AlertMessage = Resources.AutoAddDestination + $@"[{autoCcBccKeyword.CcOrBcc}] [{autoCcBccKeyword.AutoAddAddress}] (" + Resources.ApplicableKeywords + $" 「{autoCcBccKeyword.Keyword}」)", IsImportant = false, IsWhite = true, IsChecked = true });

                    _whitelist.Add(new Whitelist { WhiteName = autoCcBccKeyword.AutoAddAddress });
                }
            }

            // Chỉ thực hiện thêm Cc hoặc Bcc nếu số lượng tệp đính kèm thuộc diện cảnh báo khác 0.
            if (_checkList.Attachments.Count != 0 && !(externalDomainCount == 0 && generalSetting.IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain))
            {
                if (autoCcBccAttachedFilesList.Count != 0)
                {
                    foreach (var autoCcBccAttachedFile in autoCcBccAttachedFilesList)
                    {
                        if (autoCcBccAttachedFile.CcOrBcc == CcOrBcc.Cc)
                        {
                            if (!autoAddedCcAddressList.Contains(autoCcBccAttachedFile.AutoAddAddress) && !displayNameAndRecipient.Cc.ContainsKey(autoCcBccAttachedFile.AutoAddAddress))
                            {
                                var recipient = ((dynamic)item).Recipients.Add(autoCcBccAttachedFile.AutoAddAddress);
                                recipient.Type = (int)Outlook.OlMailRecipientType.olCC;

                                autoAddRecipients.Add(recipient);
                                autoAddedCcAddressList.Add(autoCcBccAttachedFile.AutoAddAddress);
                            }
                        }
                        else if (!autoAddedBccAddressList.Contains(autoCcBccAttachedFile.AutoAddAddress) && !displayNameAndRecipient.Bcc.ContainsKey(autoCcBccAttachedFile.AutoAddAddress))
                        {
                            var recipient = ((dynamic)item).Recipients.Add(autoCcBccAttachedFile.AutoAddAddress);
                            recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;

                            autoAddRecipients.Add(recipient);
                            autoAddedBccAddressList.Add(autoCcBccAttachedFile.AutoAddAddress);
                        }

                        _checkList.Alerts.Add(new Alert { AlertMessage = Resources.AutoAddDestination + $@"[{autoCcBccAttachedFile.CcOrBcc}] [{autoCcBccAttachedFile.AutoAddAddress}] (" + Resources.Attachments + ")", IsImportant = false, IsWhite = true, IsChecked = true });

                        _whitelist.Add(new Whitelist { WhiteName = autoCcBccAttachedFile.AutoAddAddress });
                    }
                }
            }

            if (autoCcBccRecipientList.Count != 0)
            {
                foreach (var autoCcBccRecipient in autoCcBccRecipientList)
                {
                    if (!displayNameAndRecipient.All.Any(recipient => recipient.Key.Contains(autoCcBccRecipient.TargetRecipient)) || !autoCcBccRecipient.AutoAddAddress.Contains("@")) continue;

                    if (autoCcBccRecipient.CcOrBcc == CcOrBcc.Cc)
                    {
                        if (!autoAddedCcAddressList.Contains(autoCcBccRecipient.AutoAddAddress) && !displayNameAndRecipient.Cc.ContainsKey(autoCcBccRecipient.AutoAddAddress))
                        {
                            var recipient = ((dynamic)item).Recipients.Add(autoCcBccRecipient.AutoAddAddress);
                            recipient.Type = (int)Outlook.OlMailRecipientType.olCC;

                            autoAddRecipients.Add(recipient);
                            autoAddedCcAddressList.Add(autoCcBccRecipient.AutoAddAddress);
                        }
                    }
                    else if (!autoAddedBccAddressList.Contains(autoCcBccRecipient.AutoAddAddress) && !displayNameAndRecipient.Bcc.ContainsKey(autoCcBccRecipient.AutoAddAddress))
                    {
                        var recipient = ((dynamic)item).Recipients.Add(autoCcBccRecipient.AutoAddAddress);
                        recipient.Type = (int)Outlook.OlMailRecipientType.olBCC;

                        autoAddRecipients.Add(recipient);
                        autoAddedBccAddressList.Add(autoCcBccRecipient.AutoAddAddress);
                    }

                    _checkList.Alerts.Add(new Alert { AlertMessage = Resources.AutoAddDestination + $@"[{autoCcBccRecipient.CcOrBcc}] [{autoCcBccRecipient.AutoAddAddress}] (" + Resources.ApplicableDestination + $" 「{autoCcBccRecipient.TargetRecipient}」)", IsImportant = false, IsWhite = true, IsChecked = true });

                    _whitelist.Add(new Whitelist { WhiteName = autoCcBccRecipient.AutoAddAddress });
                }
            }

            // Nếu tùy chọn luôn thêm bản thân vào Cc hoặc Bcc được bật, hãy thực hiện việc đó.
            if (isAutoAddSenderToCc || isAutoAddSenderToBcc)
            {
                var addSenderToCc = isAutoAddSenderToCc;
                var addSenderToBcc = isAutoAddSenderToBcc;

                var mailItemSender = ((dynamic)item).SenderEmailAddress;

                if (typeof(T) == typeof(Outlook.MailItem))
                {
                    if (!string.IsNullOrEmpty(((Outlook.MailItem)item).SentOnBehalfOfName) && !string.IsNullOrEmpty(((Outlook.MailItem)item).Sender.Address))
                    {
                        mailItemSender = ((Outlook.MailItem)item).Sender.Address;
                    }
                }

                var counter = 0;
                while (counter <= 5)
                {
                    counter++;
                    try
                    {
                        foreach (Outlook.Recipient recipient in ((dynamic)item).Recipients)
                        {
                            switch (recipient.Type)
                            {
                                case (int)Outlook.OlMailRecipientType.olBCC when recipient.Address.Equals(mailItemSender):
                                    addSenderToBcc = false;
                                    break;
                                case (int)Outlook.OlMailRecipientType.olCC when recipient.Address.Equals(mailItemSender):
                                    addSenderToCc = false;
                                    break;
                            }
                        }
                        counter = 6;
                        break;
                    }
                    catch (Exception)
                    {
                        Thread.Sleep(10);
                    }
                }

                if (addSenderToCc || addSenderToCc)
                {
                    if (IsX500Address(mailItemSender))
                    {
                        var exchangePrimarySmtpAddress = GetExchangePrimarySmtpAddress(mailItemSender);
                        if (exchangePrimarySmtpAddress != null)
                        {
                            mailItemSender = exchangePrimarySmtpAddress;
                        }
                    }
                }

                if (addSenderToCc)
                {
                    counter = 0;
                    while (counter <= 3)
                    {
                        counter++;
                        try
                        {
                            var senderAsRecipient = ((dynamic)item).Recipients.Add(mailItemSender);
                            Thread.Sleep(150);

                            _ = senderAsRecipient.Resolve();
                            Thread.Sleep(150);

                            senderAsRecipient.Type = (int)Outlook.OlMailRecipientType.olCC;
                            autoAddRecipients.Add(senderAsRecipient);
                            mailItemSender = senderAsRecipient.Address;
                            counter = 4;
                        }
                        catch (Exception)
                        {
                            Thread.Sleep(10);
                        }
                    }
                }

                if (addSenderToBcc)
                {
                    counter = 0;
                    while (counter < 3)
                    {
                        counter++;
                        try
                        {
                            var senderAsRecipient = ((dynamic)item).Recipients.Add(mailItemSender);
                            Thread.Sleep(150);

                            _ = senderAsRecipient.Resolve();
                            Thread.Sleep(150);

                            senderAsRecipient.Type = (int)Outlook.OlMailRecipientType.olBCC;
                            autoAddRecipients.Add(senderAsRecipient);
                            mailItemSender = senderAsRecipient.Address;
                            counter = 4;
                        }
                        catch (Exception)
                        {
                            Thread.Sleep(10);
                        }
                    }
                }

                _whitelist.Add(new Whitelist { WhiteName = sender, IsSkipConfirmation = false });
            }

            return autoAddRecipients;
        }

        /// <summary>
        /// Lấy danh sách tên tệp đính kèm được nhúng trong HTML.
        /// </summary>
        /// <param name="item">Item</param>
        /// <param name="mailHtmlBody">Nội dung thư (định dạng HTML)</param>
        /// <returns>Danh sách tên tệp đính kèm nhúng</returns>
        private List<string> MakeEmbeddedAttachmentsList<T>(T item, string mailHtmlBody)
        {
            if (typeof(T) == typeof(Outlook.MailItem))
            {
                // Chỉ xử lý nếu định dạng là HTML.
                if (((Outlook.MailItem)item).BodyFormat != Outlook.OlBodyFormat.olFormatHTML) return null;
            }

            var matches = Regex.Matches(mailHtmlBody, @"cid:.*?@");

            if (matches.Count == 0) return null;

            var embeddedAttachmentsName = new List<string>();
            foreach (var data in matches)
            {
                embeddedAttachmentsName.Add(data.ToString().Replace(@"cid:", "").Replace(@"@", ""));
            }

            return embeddedAttachmentsName;
        }

        /// <summary>
        /// Lấy thông tin tệp đính kèm và kích thước tệp, sau đó thêm vào CheckList.
        /// </summary>
        /// <param name="item">Item</param>
        /// <param name="checkList">CheckList</param>
        /// <param name="isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles">Cài đặt bỏ qua tệp đính kèm nhúng trong HTML</param>
        /// <param name="attachmentsSetting">Cài đặt liên quan đến tệp đính kèm</param>
        /// <param name="mailHtmlBody">Nội dung thư (định dạng HTML)</param>
        /// <param name="isAutoCheckAttachments">Có tự động kiểm tra hay không</param>
        /// <returns>CheckList</returns>
        private CheckList GetAttachmentsInformation<T>(in T item, CheckList checkList, bool isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles, AttachmentsSetting attachmentsSetting, string mailHtmlBody, bool isAutoCheckAttachments)
        {
            if (((dynamic)item).Attachments.Count == 0) return checkList;

            var embeddedAttachmentsName = new List<string>();
            if (isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles)
            {
                embeddedAttachmentsName = MakeEmbeddedAttachmentsList(item, mailHtmlBody);
            }

            var tempDirectoryPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            _ = Directory.CreateDirectory(tempDirectoryPath);
            checkList.TempFilePath = tempDirectoryPath;

            for (var i = 0; i < ((dynamic)item).Attachments.Count; i++)
            {
                var fileSize = "?KB";
                if (((dynamic)item).Attachments[i + 1].Size != 0)
                {
                    var sizeInKb = (double)((dynamic)item).Attachments[i + 1].Size / 1024;
                    string formattedSize;
                    if (sizeInKb >= 1 || sizeInKb == 0)
                    {
                        formattedSize = Math.Round(sizeInKb).ToString("##,###");
                    }
                    else
                    {
                        formattedSize = sizeInKb.ToString("0.###");
                    }

                    fileSize = formattedSize + "KB";
                }

                if (((dynamic)item).Attachments[i + 1].Size >= 10485760)
                {
                    checkList.Alerts.Add(new Alert { AlertMessage = Resources.IsBigAttachedFile + $"[{((dynamic)item).Attachments[i + 1].FileName}]", IsChecked = false, IsImportant = true, IsWhite = false });
                }

                // Tránh lỗi không lấy được loại tệp trong một số trường hợp.
                string fileType;
                try
                {
                    fileType = ((dynamic)item).Attachments[i + 1].FileName.Substring(((dynamic)item).Attachments[i + 1].FileName.LastIndexOf(".", StringComparison.Ordinal));
                }
                catch (Exception)
                {
                    fileType = Resources.Unknown;
                }

                var isDangerous = false;
                if (fileType == ".exe")
                {
                    checkList.Alerts.Add(new Alert { AlertMessage = Resources.IsAttachedExe + $"[{((dynamic)item).Attachments[i + 1].FileName}]", IsChecked = false, IsImportant = true, IsWhite = false });
                    isDangerous = true;
                }

                string fileName;
                try
                {
                    fileName = ((dynamic)item).Attachments[i + 1].FileName;
                }
                catch (Exception)
                {
                    fileName = Resources.Unknown;
                }

                // Bỏ qua các tệp đính kèm mà việc lấy thông tin hoàn toàn thất bại (ví dụ: hình ảnh nhúng định dạng Rich Text).
                if (fileName == Resources.Unknown && fileSize == "?KB" && fileType == Resources.Unknown) continue;

                // Chủ đích bỏ qua chứng chỉ của email có chữ ký số.
                if (fileType == ".p7s" || fileType == "p7s") continue;

                var isEncrypted = false;
                var tempFilePath = "";
                var isCanOpen = false;
                try
                {

                    if ((attachmentsSetting.IsWarningWhenEncryptedZipIsAttached || attachmentsSetting.IsProhibitedWhenEncryptedZipIsAttached) && fileName != Resources.Unknown)
                    {
                        if (attachmentsSetting.IsEnableAllAttachedFilesAreDetectEncryptedZip || fileType == ".zip" || fileType == "zip")
                        {
                            tempFilePath = Path.Combine(tempDirectoryPath, fileName);
                            ((dynamic)item).Attachments[i + 1].SaveAsFile(tempFilePath);

                            var zipTools = new ZipFileHandler();
                            if (zipTools.CheckZipIsEncryptedAndGetIncludeExtensions(tempFilePath))
                            {
                                File.Delete(tempFilePath);

                                isEncrypted = true;
                                AddAlerts(Resources.AttachedIsAnEncryptedZipFile + $" [{fileName}]", true, false, false);

                                if (attachmentsSetting.IsProhibitedWhenEncryptedZipIsAttached)
                                {
                                    checkList.IsCanNotSendMail = true;
                                    checkList.CanNotSendMailMessage = Resources.AttachedIsAnEncryptedZipFile + $"{Environment.NewLine}[{fileName}]";
                                }
                            }

                            File.Delete(tempFilePath);
                        }
                    }

                    if (attachmentsSetting.IsEnableOpenAttachedFiles && attachmentsSetting.TargetAttachmentFileExtensionOfOpen.ToLower().Contains(fileType.ToLower()))
                    {
                        // Sao chép vào thư mục tạm thời để mở tệp và ghi lại đường dẫn đó.
                        tempFilePath = Path.Combine(tempDirectoryPath, fileName);
                        ((dynamic)item).Attachments[i + 1].SaveAsFile(tempFilePath);
                        isCanOpen = true;
                    }
                }
                catch (Exception ex)
                {
                    // Log error for debugging purposes
                    System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to save attachment for opening: {ex.Message}");
                }

                var isChecked = false;
                if (!attachmentsSetting.IsMustOpenBeforeCheckTheAttachedFiles)
                {
                    isChecked = isAutoCheckAttachments;
                }

                if (embeddedAttachmentsName is null)
                {
                    checkList.Attachments.Add(new Attachment
                    {
                        FileName = fileName,
                        FileSize = fileSize,
                        FileType = fileType,
                        IsTooBig = ((dynamic)item).Attachments[i + 1].Size >= 10485760,
                        IsEncrypted = isEncrypted,
                        IsChecked = isChecked,
                        IsDangerous = isDangerous,
                        IsCanOpen = isCanOpen,
                        IsNotMustOpenBeforeCheck = !(attachmentsSetting.IsEnableOpenAttachedFiles && attachmentsSetting.IsMustOpenBeforeCheckTheAttachedFiles && isCanOpen),
                        Open = isCanOpen ? Resources.Open : "---",
                        FilePath = tempFilePath
                    });

                    continue;
                }

                // Bỏ qua tệp nhúng HTML.
                if (!embeddedAttachmentsName.Contains(fileName))
                {
                    checkList.Attachments.Add(new Attachment
                    {
                        FileName = fileName,
                        FileSize = fileSize,
                        FileType = fileType,
                        IsTooBig = ((dynamic)item).Attachments[i + 1].Size >= 10485760,
                        IsEncrypted = isEncrypted,
                        IsChecked = isChecked,
                        IsDangerous = isDangerous,
                        IsCanOpen = isCanOpen,
                        IsNotMustOpenBeforeCheck = !(attachmentsSetting.IsEnableOpenAttachedFiles && attachmentsSetting.IsMustOpenBeforeCheckTheAttachedFiles && isCanOpen),
                        Open = isCanOpen ? Resources.Open : "---",
                        FilePath = tempFilePath
                    });
                }
            }

            if (IsTaskRequestItem)
            {
                var targetAttachment = checkList.Attachments.FirstOrDefault(x => x.FileType == ".msg");
                if (targetAttachment != null)
                {
                    checkList.Attachments.Remove(targetAttachment);
                }
            }

            return checkList;
        }

        /// <summary>
        /// Hiển thị cảnh báo nếu có địa chỉ không phải là địa chỉ đề xuất (dựa trên tên và tên miền đã đăng ký) nằm trong danh sách người nhận.
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="displayNameAndRecipient">Cài đặt tên hiển thị và địa chỉ</param>
        /// <param name="nameAndDomainsList">Danh sách tên và tên miền</param>
        /// <param name="isCheckNameAndDomainsFromRecipients">Có hiển thị cảnh báo ngay cả khi không có tên người nhận trong nội dung hay không</param>
        /// <param name="isCheckNameAndDomainsIncludeSubject">Có bao gồm tiêu đề vào đối tượng kiểm tra hay không</param>
        /// <param name="isCheckNameAndDomainsFromSubject">Có hiển thị cảnh báo ngay cả khi không có tên người nhận trong tiêu đề hay không</param>
        /// <returns>CheckList</returns>
        private CheckList CheckMailBodyAndRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IEnumerable<NameAndDomains> nameAndDomainsList, bool isCheckNameAndDomainsFromRecipients, bool isCheckNameAndDomainsIncludeSubject, bool isCheckNameAndDomainsFromSubject)
        {
            if (displayNameAndRecipient is null) return checkList;

            // Nếu có giá trị cài đặt trống sẽ gây ra cảnh báo sai, nên loại bỏ các giá trị trống.
            var cleanedNameAndDomains = nameAndDomainsList.Where(nameAndDomain => !string.IsNullOrEmpty(nameAndDomain.Domain) && !string.IsNullOrEmpty(nameAndDomain.Name)).ToList();
            if (!cleanedNameAndDomains.Any()) return checkList;

            if (isCheckNameAndDomainsFromRecipients || (isCheckNameAndDomainsIncludeSubject && isCheckNameAndDomainsFromSubject))
            {
                var domainCandidateRecipients = new List<string[]>();
                foreach (var nameAndDomain in cleanedNameAndDomains)
                {
                    foreach (var recipient in displayNameAndRecipient.All)
                    {
                        if (recipient.Key.EndsWith(nameAndDomain.Domain) || recipient.Key == nameAndDomain.Domain)
                        {
                            domainCandidateRecipients.Add(new[] { recipient.Value, nameAndDomain.Name });
                        }
                    }
                }

                if (isCheckNameAndDomainsFromRecipients)
                {
                    var domainCandidateRecipientsHitCount = domainCandidateRecipients.Where(domainAndName => !checkList.MailBody.Contains(domainAndName[1])).Count(domainAndName => !domainAndName[0].Contains(checkList.SenderDomain));
                    if (domainCandidateRecipientsHitCount >= domainCandidateRecipients.Count)
                    {
                        foreach (var domainAndName in domainCandidateRecipients.Where(domainAndName => !checkList.MailBody.Contains(domainAndName[1])).Where(domainAndName => !domainAndName[0].Contains(checkList.SenderDomain)))
                        {
                            checkList.Alerts.Add(new Alert
                            {
                                AlertMessage = $"{domainAndName[0]} : {Resources.CanNotFindTheLinkedName} ({domainAndName[1]})",
                                IsImportant = true,
                                IsWhite = false,
                                IsChecked = false
                            });
                        }
                    }
                }

                if (isCheckNameAndDomainsIncludeSubject && isCheckNameAndDomainsFromSubject)
                {
                    var domainCandidateRecipientsHitCount = domainCandidateRecipients.Where(domainAndName => !checkList.Subject.Contains(domainAndName[1])).Count(domainAndName => !domainAndName[0].Contains(checkList.SenderDomain));
                    if (domainCandidateRecipientsHitCount >= domainCandidateRecipients.Count)
                    {
                        foreach (var domainAndName in domainCandidateRecipients.Where(domainAndName => !checkList.Subject.Contains(domainAndName[1])).Where(domainAndName => !domainAndName[0].Contains(checkList.SenderDomain)))
                        {
                            checkList.Alerts.Add(new Alert
                            {
                                AlertMessage = $"{domainAndName[0]} : {Resources.CanNotFindTheLinkedNameInSubject} ({domainAndName[1]})",
                                IsImportant = true,
                                IsWhite = false,
                                IsChecked = false
                            });
                        }
                    }
                }
            }

            var targetText = checkList.MailBody;
            if (isCheckNameAndDomainsIncludeSubject) { targetText += checkList.Subject; }

            var recipientCandidateDomains = (from nameAndDomain in cleanedNameAndDomains where targetText.Contains(nameAndDomain.Name) select nameAndDomain.Domain).ToList();
            // Nếu không tìm thấy ứng viên người nhận nào, không làm gì thêm. (Vì trường hợp không tìm thấy phổ biến hơn, nên nếu cảnh báo sẽ gây phiền nhiễu.)
            if (recipientCandidateDomains.Count == 0) return checkList;

            foreach (var recipient in displayNameAndRecipient.All)
            {
                if (recipientCandidateDomains.Any(domains => recipient.Key.EndsWith(domains) || domains.Equals(recipient.Key))) continue;

                // Tên miền người gửi không thuộc diện cảnh báo.
                if (!recipient.Key.Contains(checkList.SenderDomain))
                {
                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = recipient.Value + " : " + Resources.IsAlertAddressMaybeIrrelevant,
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
                }
            }

            return checkList;
        }

        /// <summary>
        /// Cảnh báo nếu không có người nhận tương ứng với từ khóa đã đăng ký.
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="displayNameAndRecipient">Cài đặt tên hiển thị và địa chỉ</param>
        /// <param name="keywordAndRecipients">Danh sách từ khóa và người nhận</param>
        /// <param name="isCheckKeywordAndRecipientsIncludeSubject">Có bao gồm tiêu đề vào đối tượng kiểm tra hay không</param>
        /// <returns>CheckList</returns>
        private CheckList CheckKeywordAndRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IEnumerable<KeywordAndRecipients> keywordAndRecipients, bool isCheckKeywordAndRecipientsIncludeSubject)
        {
            if (displayNameAndRecipient is null) return checkList;

            // Nếu có giá trị cài đặt trống sẽ gây ra cảnh báo sai, nên loại bỏ các giá trị trống.
            var cleanedKeywordAndRecipients = keywordAndRecipients.Where(keywordAndRecipient => !string.IsNullOrEmpty(keywordAndRecipient.Keyword) && !string.IsNullOrEmpty(keywordAndRecipient.Recipient)).ToList();
            if (!cleanedKeywordAndRecipients.Any()) return checkList;

            var targetText = checkList.MailBody;
            if (isCheckKeywordAndRecipientsIncludeSubject) { targetText += checkList.Subject; }

            var candidateRecipients = cleanedKeywordAndRecipients.Where(cleanedKeywordAndRecipient => targetText.Contains(cleanedKeywordAndRecipient.Keyword)).ToList();

            foreach (var candidateRecipient in candidateRecipients)
            {
                var isNeedAlert = true;
                foreach (var recipient in displayNameAndRecipient.All)
                {
                    if (recipient.Key.EndsWith(candidateRecipient.Recipient) || recipient.Key.Equals(candidateRecipient.Recipient))
                    {
                        isNeedAlert = false;
                        break;
                    }
                }

                if (isNeedAlert)
                {
                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = Resources.KeywordAndRecipientsAlert1 + candidateRecipient.Keyword + Resources.KeywordAndRecipientsAlert2 + candidateRecipient.Recipient + Resources.KeywordAndRecipientsAlert3,
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
                }
            }

            return checkList;
        }

        /// <summary>
        /// Lấy địa chỉ email người nhận và thêm vào CheckList.
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="displayNameAndRecipient">Cài đặt tên hiển thị và địa chỉ</param>
        /// <param name="alertAddressList">Cài đặt địa chỉ cảnh báo</param>
        /// <param name="internalDomainList">Cài đặt tên miền nội bộ</param>
        /// <returns>CheckList</returns>
        private CheckList GetRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IReadOnlyCollection<AlertAddress> alertAddressList, IReadOnlyCollection<InternalDomain> internalDomainList)
        {
            if (displayNameAndRecipient is null) return checkList;

            foreach (var to in displayNameAndRecipient.To)
            {
                var isExternal = true;
                foreach (var _ in internalDomainList.Where(internalDomainSetting => to.Key.EndsWith(internalDomainSetting.Domain)))
                {
                    isExternal = false;
                }

                if (to.Value.Contains(Resources.DistributionList) && to.Key.Contains(Resources.FailedToGetInformation))
                {
                    isExternal = false;
                }

                var isWhite = _whitelist.Count != 0 && _whitelist.Any(x => to.Key.EndsWith(x.WhiteName) || to.Key == x.WhiteName);
                var isSkip = false;

                if (isWhite)
                {
                    foreach (var whitelist in _whitelist.Where(whitelist => to.Key.Contains(whitelist.WhiteName)))
                    {
                        isSkip = whitelist.IsSkipConfirmation;
                    }
                }

                checkList.ToAddresses.Add(new Address { MailAddress = to.Value, IsExternal = isExternal, IsWhite = isWhite, IsChecked = isWhite, IsSkip = isSkip });

                if (alertAddressList.Count == 0) continue;

                foreach (var alertAddress in alertAddressList)
                {
                    if (!to.Key.EndsWith(alertAddress.TargetAddress)) continue;

                    if (alertAddress.IsCanNotSend)
                    {
                        checkList.IsCanNotSendMail = true;
                        checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{to.Value}]";
                        continue;
                    }

                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = string.IsNullOrEmpty(alertAddress.Message) ? Resources.IsAlertAddressToAlert + $"[{to.Value}]" : alertAddress.Message + $"[{to.Value}]",
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
                }
            }

            foreach (var cc in displayNameAndRecipient.Cc)
            {
                var isExternal = true;
                foreach (var _ in internalDomainList.Where(internalDomainSetting => cc.Key.EndsWith(internalDomainSetting.Domain)))
                {
                    isExternal = false;
                }

                if (cc.Value.Contains(Resources.DistributionList) && cc.Key.Contains(Resources.FailedToGetInformation))
                {
                    isExternal = false;
                }

                var isWhite = _whitelist.Count != 0 && _whitelist.Any(x => cc.Key.EndsWith(x.WhiteName) || cc.Key == x.WhiteName);
                var isSkip = false;

                if (isWhite)
                {
                    foreach (var whitelist in _whitelist.Where(whitelist => cc.Key.Contains(whitelist.WhiteName)))
                    {
                        isSkip = whitelist.IsSkipConfirmation;
                    }
                }

                checkList.CcAddresses.Add(new Address { MailAddress = cc.Value, IsExternal = isExternal, IsWhite = isWhite, IsChecked = isWhite, IsSkip = isSkip });

                if (alertAddressList.Count == 0) continue;

                foreach (var alertAddress in alertAddressList)
                {
                    if (!cc.Key.EndsWith(alertAddress.TargetAddress)) continue;

                    if (alertAddress.IsCanNotSend)
                    {
                        checkList.IsCanNotSendMail = true;
                        checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{cc.Value}]";
                        continue;
                    }

                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = string.IsNullOrEmpty(alertAddress.Message) ? Resources.IsAlertAddressToAlert + $"[{cc.Value}]" : alertAddress.Message + $"[{cc.Value}]",
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
                }
            }

            foreach (var bcc in displayNameAndRecipient.Bcc)
            {
                var isExternal = true;
                foreach (var _ in internalDomainList.Where(internalDomainSetting => bcc.Key.EndsWith(internalDomainSetting.Domain)))
                {
                    isExternal = false;
                }

                if (bcc.Value.Contains(Resources.DistributionList) && bcc.Key.Contains(Resources.FailedToGetInformation))
                {
                    isExternal = false;
                }

                var isWhite = _whitelist.Count != 0 && _whitelist.Any(x => bcc.Key.EndsWith(x.WhiteName) || bcc.Key == x.WhiteName);
                var isSkip = false;

                if (isWhite)
                {
                    foreach (var whitelist in _whitelist.Where(whitelist => bcc.Key.Contains(whitelist.WhiteName)))
                    {
                        isSkip = whitelist.IsSkipConfirmation;
                    }
                }

                checkList.BccAddresses.Add(new Address { MailAddress = bcc.Value, IsExternal = isExternal, IsWhite = isWhite, IsChecked = isWhite, IsSkip = isSkip });

                if (alertAddressList.Count == 0) continue;

                foreach (var alertAddress in alertAddressList)
                {
                    if (!bcc.Key.EndsWith(alertAddress.TargetAddress)) continue;

                    if (alertAddress.IsCanNotSend)
                    {
                        checkList.IsCanNotSendMail = true;
                        checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{bcc.Value}]";
                        continue;
                    }

                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = string.IsNullOrEmpty(alertAddress.Message) ? Resources.IsAlertAddressToAlert + $"[{bcc.Value}]" : alertAddress.Message + $"[{bcc.Value}]",
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
                }
            }

            return checkList;
        }

        /// <summary>
        /// Tính toán thời gian trì hoãn gửi. (Trả về thời gian trì hoãn dài nhất phù hợp với điều kiện.)
        /// </summary>
        /// <param name="displayNameAndRecipient">Cài đặt tên hiển thị và địa chỉ</param>
        /// <param name="deferredDeliveryMinutes">Cài đặt trì hoãn gửi</param>
        /// <param name="isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain">Có sử dụng tính năng này khi số tên miền bên ngoài bằng 0 hay không</param>
        /// <param name="externalDomainCount">Số lượng tên miền bên ngoài</param>
        /// <returns>Thời gian trì hoãn gửi (phút)</returns>
        private int CalcDeferredMinutes(DisplayNameAndRecipient displayNameAndRecipient, IReadOnlyCollection<DeferredDeliveryMinutes> deferredDeliveryMinutes, bool isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain, int externalDomainCount)
        {
            if (deferredDeliveryMinutes.Count == 0) return 0;
            if (externalDomainCount == 0 && isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain) return 0;

            var deferredMinutes = 0;

            // Nếu đã đăng ký chỉ với "@", hãy coi đó là thời gian trì hoãn gửi mặc định.
            foreach (var config in deferredDeliveryMinutes.Where(config => config.TargetAddress == "@"))
            {
                deferredMinutes = config.DeferredMinutes;
            }

            if (displayNameAndRecipient.To.Count != 0)
            {
                foreach (var toRecipients in displayNameAndRecipient.To)
                {
                    foreach (var config in deferredDeliveryMinutes.Where(config => toRecipients.Key.Contains(config.TargetAddress) && deferredMinutes < config.DeferredMinutes))
                    {
                        deferredMinutes = config.DeferredMinutes;
                    }
                }
            }

            if (displayNameAndRecipient.Cc.Count != 0)
            {
                foreach (var ccRecipients in displayNameAndRecipient.Cc)
                {
                    foreach (var config in deferredDeliveryMinutes.Where(config => ccRecipients.Key.Contains(config.TargetAddress) && deferredMinutes < config.DeferredMinutes))
                    {
                        deferredMinutes = config.DeferredMinutes;
                    }
                }
            }

            if (displayNameAndRecipient.Bcc.Count != 0)
            {
                foreach (var bccRecipients in displayNameAndRecipient.Bcc)
                {
                    foreach (var config in deferredDeliveryMinutes.Where(config => bccRecipients.Key.Contains(config.TargetAddress) && deferredMinutes < config.DeferredMinutes))
                    {
                        deferredMinutes = config.DeferredMinutes;
                    }
                }
            }

            return deferredMinutes;
        }

        /// <summary>
        /// Hiển thị cảnh báo hoặc cấm gửi email nếu số lượng tên miền bên ngoài trong To và Cc vượt quá giá trị quy định.
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="settings">Cài đặt cảnh báo số lượng tên miền bên ngoài và tự động chuyển sang Bcc</param>
        /// <param name="externalDomainNumToAndCc">Số lượng tên miền bên ngoài trong To và Cc</param>
        /// <param name="isForceAutoChangeRecipientsToBcc">Có bắt buộc chuyển tất cả người nhận sang Bcc hay không</param>
        /// <returns>CheckList</returns>
        private CheckList ExternalDomainsWarningIfNeeded(CheckList checkList, ExternalDomainsWarningAndAutoChangeToBcc settings, int externalDomainNumToAndCc, bool isForceAutoChangeRecipientsToBcc)
        {
            // Nếu tính năng bắt buộc chuyển đổi Bcc được bật, hãy bỏ qua tính năng này.
            if (isForceAutoChangeRecipientsToBcc) return checkList;

            if (settings.TargetToAndCcExternalDomainsNum > externalDomainNumToAndCc) return checkList;

            if (settings.IsProhibitedWhenLargeNumberOfExternalDomains)
            {
                checkList.IsCanNotSendMail = true;
                checkList.CanNotSendMailMessage = Resources.ProhibitedWhenLargeNumberOfExternalDomainsAlert + $"[{settings.TargetToAndCcExternalDomainsNum}]";

                return checkList;
            }

            if (settings.IsWarningWhenLargeNumberOfExternalDomains && !settings.IsAutoChangeToBccWhenLargeNumberOfExternalDomains)
            {
                checkList.Alerts.Add(new Alert
                {
                    AlertMessage = Resources.LargeNumberOfExternalDomainAlert + $"[{settings.TargetToAndCcExternalDomainsNum}]",
                    IsImportant = true,
                    IsWhite = false,
                    IsChecked = false
                });

                return checkList;
            }

            return checkList;
        }

        /// <summary>
        /// Xóa địa chỉ người nhận được chỉ định khỏi To và Cc, và thêm vào Bcc.
        /// </summary>
        /// <param name="item">Item</param>
        /// <param name="mailItemsRecipientAndMailAddress">Địa chỉ email và Recipient của MailItem</param>
        /// <param name="senderMailAddress">Địa chỉ email người gửi</param>
        /// <param name="isNeedsAddToSender">Có thêm địa chỉ người gửi vào To hay không</param>
        private void ChangeToBcc<T>(T item, IReadOnlyCollection<MailItemsRecipientAndMailAddress> mailItemsRecipientAndMailAddress, string senderMailAddress, bool isNeedsAddToSender)
        {
            if ((dynamic)item is null) return;

            var targetMailAddressAndRecipient = new Dictionary<string, string>();

            foreach (Outlook.Recipient recipient in ((dynamic)item).Recipients)
            {
                foreach (var target in mailItemsRecipientAndMailAddress)
                {
                    if (recipient.Address == target.MailItemsRecipient) targetMailAddressAndRecipient[target.MailAddress] = target.MailItemsRecipient;
                }
            }

            // Nếu sử dụng Index để xóa, Index sẽ bị lệch và không thể xóa chính xác nhiều mục, vì vậy hãy tìm đối tượng cần xóa và xóa nó.
            var targetCount = targetMailAddressAndRecipient.Count;
            while (targetCount > 0)
            {
                foreach (var target in targetMailAddressAndRecipient)
                {
                    foreach (Outlook.Recipient recipient in ((dynamic)item).Recipients)
                    {
                        if (recipient.Address != target.Value) continue;
                        ((dynamic)item).Recipients.Remove(recipient.Index);
                        targetCount--;
                    }
                }
            }

            foreach (var addTarget in targetMailAddressAndRecipient.Select(mailAddress => ((dynamic)item).Recipients.Add(mailAddress.Key)))
            {
                addTarget.Type = (int)Outlook.OlMailRecipientType.olBCC;
            }

            if (isNeedsAddToSender)
            {
                var senderRecipient = ((dynamic)item).Recipients.Add(senderMailAddress);
                senderRecipient.Type = (int)Outlook.OlMailRecipientType.olTo;
            }

            _ = ((dynamic)item).Recipients.ResolveAll();
        }

        /// <summary>
        /// Nếu điều kiện được đáp ứng, chuyển đổi các địa chỉ bên ngoài trong To và Cc sang Bcc.
        /// </summary>
        /// <param name="item">Item</param>
        /// <param name="displayNameAndRecipient">Cài đặt tên hiển thị và địa chỉ</param>
        /// <param name="settings">Cài đặt cảnh báo số lượng tên miền bên ngoài và tự động chuyển sang Bcc</param>
        /// <param name="internalDomains">Tên miền nội bộ</param>
        /// <param name="externalDomainNumToAndCc">Số lượng tên miền bên ngoài trong To và Cc</param>
        /// <param name="senderDomain">Tên miền người gửi</param>
        /// <param name="senderMailAddress">Địa chỉ người gửi</param>
        /// <param name="forceAutoChangeRecipientsToBccSetting">Cài đặt bắt buộc chuyển đổi người nhận sang Bcc</param>
        /// <returns>DisplayNameAndRecipient</returns>
        private DisplayNameAndRecipient ExternalDomainsChangeToBccIfNeeded<T>(T item, DisplayNameAndRecipient displayNameAndRecipient, ExternalDomainsWarningAndAutoChangeToBcc settings, ICollection<InternalDomain> internalDomains, int externalDomainNumToAndCc, string senderDomain, string senderMailAddress, ForceAutoChangeRecipientsToBcc forceAutoChangeRecipientsToBccSetting)
        {
            if ((!settings.IsAutoChangeToBccWhenLargeNumberOfExternalDomains || settings.IsProhibitedWhenLargeNumberOfExternalDomains || settings.TargetToAndCcExternalDomainsNum > externalDomainNumToAndCc) && !forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc) return displayNameAndRecipient;

            if (forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc && forceAutoChangeRecipientsToBccSetting.IsIncludeInternalDomain)
            {
                internalDomains.Clear();
            }
            else
            {
                internalDomains.Add(new InternalDomain { Domain = senderDomain });
            }

            var removeTarget = new List<string>();

            foreach (var to in displayNameAndRecipient.To)
            {
                var isInternal = false;
                foreach (var _ in internalDomains.Where(internalDomain => to.Key.EndsWith(internalDomain.Domain)))
                {
                    isInternal = true;
                }
                if (isInternal) continue;

                displayNameAndRecipient.Bcc[to.Key] = to.Value;
                removeTarget.Add(to.Key);
            }
            foreach (var target in removeTarget)
            {
                _ = displayNameAndRecipient.To.Remove(target);
            }

            removeTarget.Clear();

            foreach (var cc in displayNameAndRecipient.Cc)
            {
                var isInternal = false;
                foreach (var _ in internalDomains.Where(internalDomain => cc.Key.EndsWith(internalDomain.Domain)))
                {
                    isInternal = true;
                }
                if (isInternal) continue;

                displayNameAndRecipient.Bcc[cc.Key] = cc.Value;
                removeTarget.Add(cc.Key);
            }
            foreach (var target in removeTarget)
            {
                _ = displayNameAndRecipient.Cc.Remove(target);
            }

            if (forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc)
            {
                AddAlerts(Resources.ForceAutoChangeRecipientsToBccAlert + $"[{settings.TargetToAndCcExternalDomainsNum}]", false, false, true);
            }
            else
            {
                AddAlerts(Resources.ExternalDomainsChangeToBccAlert + $"[{settings.TargetToAndCcExternalDomainsNum}]", true, false, false);
            }
            // Nếu To không còn tồn tại, thêm người gửi vào To.
            var isNeedsAddToSender = false;
            var thisSenderMailAddress = forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc && !string.IsNullOrEmpty(forceAutoChangeRecipientsToBccSetting.ToRecipient) ? forceAutoChangeRecipientsToBccSetting.ToRecipient : senderMailAddress;
            if (displayNameAndRecipient.To.Count == 0)
            {
                displayNameAndRecipient.To[thisSenderMailAddress] = thisSenderMailAddress;
                isNeedsAddToSender = true;

                AddAlerts(thisSenderMailAddress == senderMailAddress
                        ? Resources.AutoAddSendersAddressToAlert
                        : Resources.AutoAddToRecipientByForceAutoChangeRecipientsToBccAddressToAlert, true, false, false);
            }

            var targetMailRecipientsIndex = new List<MailItemsRecipientAndMailAddress>();
            // Không làm gì với những người nhận vốn đã ở trong Bcc.
            foreach (var recipient in displayNameAndRecipient.MailRecipientsIndex.Where(recipient => recipient.Type != (int)Outlook.OlMailRecipientType.olBCC))
            {
                var isExternal = true;
                foreach (var _ in internalDomains.Where(internalDomain => recipient.MailAddress.EndsWith(internalDomain.Domain)))
                {
                    isExternal = false;
                }

                if (isExternal) targetMailRecipientsIndex.Add(recipient);
            }

            ChangeToBcc(item, targetMailRecipientsIndex, thisSenderMailAddress, isNeedsAddToSender);

            return displayNameAndRecipient;
        }

        /// <summary>
        /// Kiểm tra mối liên kết giữa người nhận và tên tệp đính kèm.
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="isAttachmentsProhibited">Có cấm gửi email có tệp đính kèm hay không</param>
        /// <param name="isWarningWhenAttachedRealFile">Nếu tệp thực được đính kèm, có hiển thị cảnh báo khuyến nghị đính kèm dưới dạng liên kết hay không</param>
        /// <param name="attachmentProhibitedRecipientsList">Cài đặt người nhận bị cấm gửi tệp đính kèm</param>
        /// <param name="recipientsAndAttachmentsNameList">Cài đặt liên kết giữa người nhận và tên tệp đính kèm</param>
        /// <param name="attachmentAlertRecipientsList">Cài đặt người nhận cảnh báo tệp đính kèm và nội dung cảnh báo</param>
        /// <returns>CheckList</returns>
        private CheckList CheckRecipientsAndAttachments(CheckList checkList, bool isAttachmentsProhibited, bool isWarningWhenAttachedRealFile, IReadOnlyCollection<AttachmentProhibitedRecipients> attachmentProhibitedRecipientsList, IReadOnlyCollection<RecipientsAndAttachmentsName> recipientsAndAttachmentsNameList, IReadOnlyCollection<AttachmentAlertRecipients> attachmentAlertRecipientsList)
        {
            if (checkList.Attachments.Count <= 0) return checkList;

            if (isAttachmentsProhibited)
            {
                checkList.IsCanNotSendMail = true;
                checkList.CanNotSendMailMessage = Resources.AttachmentsProhibitedMessage;

                // Vì việc gửi email có tệp đính kèm bị cấm, không làm gì thêm.
                return checkList;
            }

            if (attachmentProhibitedRecipientsList.Count > 0)
            {
                var prohibitedRecipients = "";
                var isProhibited = false;

                foreach (var prohibitedRecipient in attachmentProhibitedRecipientsList)
                {
                    foreach (var to in checkList.ToAddresses.Where(to => to.MailAddress.Contains(prohibitedRecipient.Recipient)))
                    {
                        checkList.IsCanNotSendMail = true;
                        isProhibited = true;
                        prohibitedRecipients += " " + to.MailAddress;
                    }

                    foreach (var cc in checkList.CcAddresses.Where(cc => cc.MailAddress.Contains(prohibitedRecipient.Recipient)))
                    {
                        checkList.IsCanNotSendMail = true;
                        isProhibited = true;
                        prohibitedRecipients += " " + cc.MailAddress;
                    }

                    foreach (var bcc in checkList.BccAddresses.Where(bcc => bcc.MailAddress.Contains(prohibitedRecipient.Recipient)))
                    {
                        checkList.IsCanNotSendMail = true;
                        isProhibited = true;
                        prohibitedRecipients += " " + bcc.MailAddress;
                    }
                }

                if (isProhibited)
                {
                    checkList.CanNotSendMailMessage = Resources.AttachmentProhibitedRecipientsMessage + "：" + prohibitedRecipients;

                    // Vì đây là người nhận bị cấm gửi tệp đính kèm, không làm gì thêm.
                    return checkList;
                }
            }

            if (attachmentAlertRecipientsList.Count > 0)
            {
                foreach (var attachmentAlertRecipient in attachmentAlertRecipientsList)
                {
                    foreach (var to in checkList.ToAddresses.Where(to => to.MailAddress.Contains(attachmentAlertRecipient.Recipient)))
                    {
                        AddAlerts(string.IsNullOrEmpty(attachmentAlertRecipient.Message) ? Resources.AttachmentAlertRecipientsMessage + $"[{to.MailAddress}]" : attachmentAlertRecipient.Message + $"[{to.MailAddress}]", true, false, false);
                    }
                    foreach (var cc in checkList.CcAddresses.Where(cc => cc.MailAddress.Contains(attachmentAlertRecipient.Recipient)))
                    {
                        AddAlerts(string.IsNullOrEmpty(attachmentAlertRecipient.Message) ? Resources.AttachmentAlertRecipientsMessage + $"[{cc.MailAddress}]" : attachmentAlertRecipient.Message + $"[{cc.MailAddress}]", true, false, false);
                    }
                    foreach (var bcc in checkList.BccAddresses.Where(bcc => bcc.MailAddress.Contains(attachmentAlertRecipient.Recipient)))
                    {
                        AddAlerts(string.IsNullOrEmpty(attachmentAlertRecipient.Message) ? Resources.AttachmentAlertRecipientsMessage + $"[{bcc.MailAddress}]" : attachmentAlertRecipient.Message + $"[{bcc.MailAddress}]", true, false, false);
                    }
                }
            }

            if (recipientsAndAttachmentsNameList.Count > 0)
            {
                foreach (var recipientsAndAttachmentsName in recipientsAndAttachmentsNameList)
                {
                    foreach (var attachment in checkList.Attachments.Where(attachment => attachment.FileName.Contains(recipientsAndAttachmentsName.AttachmentsName)))
                    {
                        foreach (var to in checkList.ToAddresses.Where(to => to.IsExternal))
                        {
                            if (!to.MailAddress.Contains(recipientsAndAttachmentsName.Recipient))
                            {
                                AddAlerts(Resources.RecipientsAndAttachmentsNameMessage + "：" + to.MailAddress + " / " + attachment.FileName, true, true, false);
                            }
                        }

                        foreach (var cc in checkList.CcAddresses.Where(cc => cc.IsExternal))
                        {
                            if (!cc.MailAddress.Contains(recipientsAndAttachmentsName.Recipient))
                            {
                                AddAlerts(Resources.RecipientsAndAttachmentsNameMessage + "：" + cc.MailAddress + " / " + attachment.FileName, true, true, false);
                            }
                        }

                        foreach (var bcc in checkList.BccAddresses.Where(bcc => bcc.IsExternal))
                        {
                            if (!bcc.MailAddress.Contains(recipientsAndAttachmentsName.Recipient))
                            {
                                AddAlerts(Resources.RecipientsAndAttachmentsNameMessage + "：" + bcc.MailAddress + " / " + attachment.FileName, true, true, false);
                            }
                        }
                    }
                }
            }

            if (isWarningWhenAttachedRealFile)
            {
                AddAlerts(Resources.RecommendationOfAttachFileAsLink, false, true, false);
            }

            return checkList;
        }

        /// <summary>
        /// Tự động đánh dấu kiểm tra trước cho các người nhận đã đăng ký trong Danh bạ (Sổ địa chỉ).
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="displayNameAndRecipient">Cài đặt tên hiển thị và địa chỉ</param>
        /// <param name="contactsList">Danh bạ (Sổ địa chỉ)</param>
        /// <param name="isAutoCheckRegisteredInContacts">Có tự động đánh dấu kiểm tra cho người nhận đã đăng ký trong danh bạ hay không</param>
        /// <returns>CheckList</returns>
        private CheckList AutoCheckRegisteredItemsInContacts(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IEnumerable<string> contactsList, bool isAutoCheckRegisteredInContacts)
        {
            if (!isAutoCheckRegisteredInContacts) return checkList;

            foreach (var mailItemsRecipient in contactsList.SelectMany(contact => displayNameAndRecipient.MailRecipientsIndex.Where(mailItemsRecipient => contact == mailItemsRecipient.MailAddress || contact == mailItemsRecipient.MailItemsRecipient)))
            {
                foreach (var toAddress in checkList.ToAddresses.Where(toAddress => toAddress.MailAddress.Contains(mailItemsRecipient.MailAddress)))
                {
                    toAddress.IsChecked = true;
                }

                foreach (var ccAddress in checkList.CcAddresses.Where(ccAddress => ccAddress.MailAddress.Contains(mailItemsRecipient.MailAddress)))
                {
                    ccAddress.IsChecked = true;
                }

                foreach (var bccAddress in checkList.BccAddresses.Where(bccAddress => bccAddress.MailAddress.Contains(mailItemsRecipient.MailAddress)))
                {
                    bccAddress.IsChecked = true;
                }
            }

            return checkList;
        }

        /// <summary>
        /// Hiển thị cảnh báo hoặc chặn gửi đối với người nhận chưa đăng ký trong Danh bạ (Sổ địa chỉ).
        /// </summary>
        /// <param name="checkList">CheckList</param>
        /// <param name="displayNameAndRecipient">Cài đặt tên hiển thị và địa chỉ</param>
        /// <param name="contactsList">Danh bạ (Sổ địa chỉ)</param>
        /// <param name="internalDomainList">Tên miền nội bộ</param>
        /// <param name="isWarningIfRecipientsIsNotRegistered">Có hiển thị cảnh báo cho người nhận chưa đăng ký hay không</param>
        /// <param name="isProhibitsSendingMailIfRecipientsIsNotRegistered">Có cấm gửi mail nếu tồn tại người nhận chưa đăng ký hay không</param>
        /// <returns>CheckList</returns>
        private CheckList AddAlertOrProhibitsSendingMailIfIfRecipientsIsNotRegistered(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, IEnumerable<string> contactsList, IReadOnlyCollection<InternalDomain> internalDomainList, bool isWarningIfRecipientsIsNotRegistered, bool isProhibitsSendingMailIfRecipientsIsNotRegistered)
        {
            if (!(isWarningIfRecipientsIsNotRegistered || isProhibitsSendingMailIfRecipientsIsNotRegistered)) return checkList;

            var selectedContactsList = contactsList.SelectMany(contact => displayNameAndRecipient.MailRecipientsIndex.Where(mailItemsRecipient => contact == mailItemsRecipient.MailAddress || contact == mailItemsRecipient.MailItemsRecipient)).Select(x => x.MailAddress).ToList();

            foreach (var to in displayNameAndRecipient.To.Where(to => !selectedContactsList.Contains(to.Key)))
            {
                // Không áp dụng cho tên miền nội bộ
                if (internalDomainList.Any(internalDomain => to.Key.EndsWith(internalDomain.Domain))) continue;

                if (isProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.ProhibitsSendingMailIfRecipientsIsNotRegisteredMessage + $" [{to.Value}]";
                    return checkList;
                }

                AddAlerts(Resources.WarningIfRecipientsIsNotRegisteredMessage + $" [{to.Value}]", true, false, false);
            }

            foreach (var cc in displayNameAndRecipient.Cc.Where(cc => !selectedContactsList.Contains(cc.Key)))
            {
                // Không áp dụng cho tên miền nội bộ
                if (internalDomainList.Any(internalDomain => cc.Key.EndsWith(internalDomain.Domain))) continue;

                if (isProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.ProhibitsSendingMailIfRecipientsIsNotRegisteredMessage + $" [{cc.Value}]";
                    return checkList;
                }

                AddAlerts(Resources.WarningIfRecipientsIsNotRegisteredMessage + $" [{cc.Value}]", true, false, false);
            }

            foreach (var bcc in displayNameAndRecipient.Bcc.Where(bcc => !selectedContactsList.Contains(bcc.Key)))
            {
                // Không áp dụng cho tên miền nội bộ
                if (internalDomainList.Any(internalDomain => bcc.Key.EndsWith(internalDomain.Domain))) continue;

                if (isProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.ProhibitsSendingMailIfRecipientsIsNotRegisteredMessage + $" [{bcc.Value}]";
                    return checkList;
                }

                AddAlerts(Resources.WarningIfRecipientsIsNotRegisteredMessage + $" [{bcc.Value}]", true, false, false);
            }

            return checkList;
        }

        /// <summary>
        /// Thêm văn bản vào bản xem trước nội dung. (Tại thời điểm này, chưa thêm vào nội dung thư thực tế. *Vì có khả năng việc gửi sẽ bị hủy)
        /// </summary>
        /// <param name="mailBody">Nội dung thư (định dạng văn bản)</param>
        /// <param name="autoAddMessageSetting">autoAddMessageSetting</param>
        /// <returns>Nội dung thư (định dạng văn bản)</returns>
        private string AddMessageToBodyPreview(string mailBody, AutoAddMessage autoAddMessageSetting)
        {
            if (mailBody == Resources.FailedToGetInformation) return mailBody;

            if (autoAddMessageSetting.IsAddToStart)
            {
                mailBody = autoAddMessageSetting.MessageOfAddToStart + Environment.NewLine + Environment.NewLine + mailBody;
                AddAlerts(Resources.AddedTextAtTheBeginning, false, false, true);
            }

            if (autoAddMessageSetting.IsAddToEnd)
            {
                mailBody = mailBody + Environment.NewLine + autoAddMessageSetting.MessageOfAddToEnd;
                AddAlerts(Resources.AddedTextAtTheEnd, false, false, true);
            }

            return mailBody;
        }

        /// <summary>
        /// Thêm cảnh báo.
        /// </summary>
        /// <param name="alertMessage">Nội dung cảnh báo</param>
        /// <param name="isImportant">Có quan trọng hay không</param>
        /// <param name="isWhite">Tạm thời chưa sử dụng</param>
        /// <param name="isChecked">Có tự động kiểm tra hay không</param>
        private void AddAlerts(string alertMessage, bool isImportant, bool isWhite, bool isChecked)
        {
            _checkList.Alerts.Add(new Alert
            {
                AlertMessage = alertMessage,
                IsImportant = isImportant,
                IsWhite = isWhite,
                IsChecked = isChecked
            });
        }

        #region Tools

        /// <summary>
        /// Xác định xem có phải là địa chỉ email hay không.
        /// </summary>
        /// <param name="emailAddress">Chuỗi ký tự cần xác định</param>
        /// <returns>Có phải là địa chỉ email hay không</returns>
        private bool IsValidEmailAddress(string emailAddress)
        {
            if (string.IsNullOrWhiteSpace(emailAddress)) return false;

            try
            {
                emailAddress = Regex.Replace(emailAddress, @"(@)(.+)$", DomainMapper, RegexOptions.None, TimeSpan.FromMilliseconds(500));
                string DomainMapper(Match match)
                {
                    var idnMapping = new IdnMapping();
                    var domainName = idnMapping.GetAscii(match.Groups[2].Value);
                    return match.Groups[1].Value + domainName;
                }
            }
            catch (Exception)
            {
                return false;
            }

            try
            {
                return Regex.IsMatch(emailAddress, @"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" + @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-0-9a-z]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$",
                    RegexOptions.IgnoreCase, TimeSpan.FromMilliseconds(500));
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Lấy tất cả địa chỉ đã đăng ký trong danh bạ.
        /// </summary>
        /// <param name="contacts">Thư mục danh bạ</param>
        /// <returns>Danh sách địa chỉ đã đăng ký trong danh bạ</returns>
        private List<string> MakeContactsList(Outlook.MAPIFolder contacts)
        {
            if (contacts is null) return null;

            var contactsList = new List<string>();
            foreach (dynamic contact in contacts.Items)
            {
                if (!(contact is Outlook.ContactItem foundContact))
                {
                    try
                    {
                        var entryId = contact.EntryID;

                        var tempOutlookApp = new Outlook.Application().GetNamespace("MAPI");
                        var distList = (Outlook.DistListItem)tempOutlookApp.GetItemFromID(entryId);

                        for (var i = 1; i < distList.MemberCount + 1; i++)
                        {
                            if (!(distList.GetMember(i).Address is null) && distList.GetMember(i).Address != "Unknown")
                            {
                                contactsList.Add(distList.GetMember(i).Address);
                            }

                        }
                    }
                    catch (Exception ex)
                    {
                        // Log error for debugging purposes
                        System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Failed to get dist list members: {ex.Message}");
                    }
                }
                else
                {
                    if (!(foundContact.Email1Address is null))
                    {
                        contactsList.Add(foundContact.Email1Address);
                        if (IsValidEmailAddress(foundContact.Email1Address)) continue;
                        // Nếu địa chỉ đã đăng ký không phải là địa chỉ email, có khả năng đó là CN(X.500) của Exchange, vì vậy hãy đăng ký cả địa chỉ đã chuyển đổi sang email thông thường.
                        var exchangePrimarySmtpAddress = GetExchangePrimarySmtpAddress(foundContact.Email1Address);
                        if (!(exchangePrimarySmtpAddress is null))
                        {
                            contactsList.Add(exchangePrimarySmtpAddress);
                        }
                    }
                    else if (!(foundContact.Email2Address is null))
                    {
                        contactsList.Add(foundContact.Email2Address);
                        if (IsValidEmailAddress(foundContact.Email2Address)) continue;
                        // Nếu địa chỉ đã đăng ký không phải là địa chỉ email, có khả năng đó là CN(X.500) của Exchange, vì vậy hãy đăng ký cả địa chỉ đã chuyển đổi sang email thông thường.
                        var exchangePrimarySmtpAddress = GetExchangePrimarySmtpAddress(foundContact.Email2Address);
                        if (!(exchangePrimarySmtpAddress is null))
                        {
                            contactsList.Add(exchangePrimarySmtpAddress);
                        }
                    }
                    else if (!(foundContact.Email3Address is null))
                    {
                        contactsList.Add(foundContact.Email3Address);
                        if (IsValidEmailAddress(foundContact.Email3Address)) continue;
                        // Nếu địa chỉ đã đăng ký không phải là địa chỉ email, có khả năng đó là CN(X.500) của Exchange, vì vậy hãy đăng ký cả địa chỉ đã chuyển đổi sang email thông thường.
                        var exchangePrimarySmtpAddress = GetExchangePrimarySmtpAddress(foundContact.Email3Address);
                        if (!(exchangePrimarySmtpAddress is null))
                        {
                            contactsList.Add(exchangePrimarySmtpAddress);
                        }
                    }
                }

            }

            return contactsList;
        }

        /// <summary>
        /// Chuyển đổi địa chỉ định dạng X500 sang địa chỉ email thông thường.
        /// </summary>
        /// <param name="x500">Địa chỉ định dạng X500</param>
        /// <returns>Địa chỉ email thông thường</returns>
        private string GetExchangePrimarySmtpAddress(string x500)
        {
            if (string.IsNullOrEmpty(x500)) return null;
            if (!IsX500Address(x500)) return null;

            try
            {
                var tempOutlookApp = new Outlook.Application();

                return ComRetryHelper.Execute(() =>
                {
                    var tempRecipient = tempOutlookApp.Session.CreateRecipient(x500);
                    _ = tempRecipient.Resolve();

                    var addressEntry = tempRecipient.AddressEntry;
                    if (addressEntry is null) return null;

                    switch (addressEntry.AddressEntryUserType)
                    {
                        case Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry:
                            {
                                var exchangeUser = addressEntry.GetExchangeUser();
                                return exchangeUser?.PrimarySmtpAddress;
                            }
                        case Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry:
                            {
                                var distributionList = addressEntry.GetExchangeDistributionList();
                                return distributionList?.PrimarySmtpAddress;
                            }
                        case Outlook.OlAddressEntryUserType.olExchangeRemoteUserAddressEntry:
                            {
                                var exchangeUser = addressEntry.GetExchangeUser();
                                return exchangeUser?.PrimarySmtpAddress;
                            }
                        case Outlook.OlAddressEntryUserType.olSmtpAddressEntry:
                            return addressEntry.Address;
                    }
                    return null;
                });
            }
            catch (Exception)
            {
                // Force commit: Ensure CI catches this fix
                return null;
            }
        }

        /// <summary>
        /// Xác định xem địa chỉ email có phải là định dạng X500 hay không
        /// </summary>
        /// <param name="emailAddress">Địa chỉ email</param>
        /// <returns>Có phải định dạng X500 hay không</returns>
        private bool IsX500Address(string emailAddress)
        {
            const string x500Pattern = @"^/o=.*/ou=.*/cn=.*";

            return Regex.IsMatch(emailAddress, x500Pattern, RegexOptions.IgnoreCase);
        }

        #endregion
    }
}