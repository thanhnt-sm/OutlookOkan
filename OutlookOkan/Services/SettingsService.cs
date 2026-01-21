using OutlookOkan.Handlers;
using OutlookOkan.Models;
using OutlookOkan.Types;
using System.Collections.Generic;
using System.Linq;

namespace OutlookOkan.Services
{
    public class SettingsService
    {
        public List<Whitelist> Whitelist { get; private set; }
        public List<AlertKeywordAndMessage> AlertKeywordAndMessageList { get; private set; }
        public List<AlertKeywordAndMessageForSubject> AlertKeywordAndMessageForSubjectList { get; private set; }
        public List<AutoCcBccKeyword> AutoCcBccKeywordList { get; private set; }
        public List<AutoCcBccAttachedFile> AutoCcBccAttachedFilesList { get; private set; }
        public List<AutoCcBccRecipient> AutoCcBccRecipientList { get; private set; }
        public List<AlertAddress> AlertAddressList { get; private set; }
        public List<NameAndDomains> NameAndDomainsList { get; private set; }
        public List<KeywordAndRecipients> KeywordAndRecipientsList { get; private set; }
        public List<DeferredDeliveryMinutes> DeferredDeliveryMinutesList { get; private set; }
        public List<InternalDomain> InternalDomainList { get; private set; }
        public ExternalDomainsWarningAndAutoChangeToBcc ExternalDomainsWarningAndAutoChangeToBccSetting { get; private set; }
        public AttachmentsSetting AttachmentsSetting { get; private set; }
        public List<RecipientsAndAttachmentsName> RecipientsAndAttachmentsNameList { get; private set; }
        public List<AttachmentProhibitedRecipients> AttachmentProhibitedRecipientsList { get; private set; }
        public List<AttachmentAlertRecipients> AttachmentAlertRecipientsList { get; private set; }
        public ForceAutoChangeRecipientsToBcc ForceAutoChangeRecipientsToBccSetting { get; private set; }
        public AutoAddMessage AutoAddMessageSetting { get; private set; }
        public List<AutoDeleteRecipient> AutoDeleteRecipients { get; private set; }

        public SettingsService()
        {
            LoadSettings();
        }

        public void LoadSettings()
        {
            Whitelist = CsvFileHandler.ReadCsv<Whitelist>(typeof(WhitelistMap), "Whitelist.csv")
                .Where(x => !string.IsNullOrEmpty(x.WhiteName)).ToList();

            AlertKeywordAndMessageList = CsvFileHandler.ReadCsv<AlertKeywordAndMessage>(typeof(AlertKeywordAndMessageMap), "AlertKeywordAndMessageList.csv")
                .Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).ToList();

            AlertKeywordAndMessageForSubjectList = CsvFileHandler.ReadCsv<AlertKeywordAndMessageForSubject>(typeof(AlertKeywordAndMessageForSubjectMap), "AlertKeywordAndMessageListForSubject.csv")
                .Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).ToList();

            AutoCcBccKeywordList = CsvFileHandler.ReadCsv<AutoCcBccKeyword>(typeof(AutoCcBccKeywordMap), "AutoCcBccKeywordList.csv")
                .Where(x => !string.IsNullOrEmpty(x.AutoAddAddress) && !string.IsNullOrEmpty(x.Keyword)).ToList();

            AutoCcBccAttachedFilesList = CsvFileHandler.ReadCsv<AutoCcBccAttachedFile>(typeof(AutoCcBccAttachedFileMap), "AutoCcBccAttachedFileList.csv")
                .Where(x => !string.IsNullOrEmpty(x.AutoAddAddress)).ToList();

            AutoCcBccRecipientList = CsvFileHandler.ReadCsv<AutoCcBccRecipient>(typeof(AutoCcBccRecipientMap), "AutoCcBccRecipientList.csv")
                .Where(x => !string.IsNullOrEmpty(x.AutoAddAddress) && !string.IsNullOrEmpty(x.TargetRecipient)).ToList();

            AlertAddressList = CsvFileHandler.ReadCsv<AlertAddress>(typeof(AlertAddressMap), "AlertAddressList.csv")
                .Where(x => !string.IsNullOrEmpty(x.TargetAddress)).ToList();

            NameAndDomainsList = CsvFileHandler.ReadCsv<NameAndDomains>(typeof(NameAndDomainsMap), "NameAndDomains.csv")
                .Where(x => !string.IsNullOrEmpty(x.Domain) && !string.IsNullOrEmpty(x.Name)).ToList();

            KeywordAndRecipientsList = CsvFileHandler.ReadCsv<KeywordAndRecipients>(typeof(KeywordAndRecipientsMap), "keywordAndRecipientsList.csv")
                .Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.Recipient)).ToList();

            DeferredDeliveryMinutesList = CsvFileHandler.ReadCsv<DeferredDeliveryMinutes>(typeof(DeferredDeliveryMinutesMap), "DeferredDeliveryMinutes.csv")
                .Where(x => !string.IsNullOrEmpty(x.TargetAddress)).ToList();

            InternalDomainList = CsvFileHandler.ReadCsv<InternalDomain>(typeof(InternalDomainMap), "InternalDomainList.csv")
                .Where(x => !string.IsNullOrEmpty(x.Domain)).ToList();

            // Cảnh báo tên miền bên ngoài
            ExternalDomainsWarningAndAutoChangeToBccSetting = new ExternalDomainsWarningAndAutoChangeToBcc();
            var extList = CsvFileHandler.ReadCsv<ExternalDomainsWarningAndAutoChangeToBcc>(typeof(ExternalDomainsWarningAndAutoChangeToBccMap), "ExternalDomainsWarningAndAutoChangeToBccSetting.csv");
            if (extList.Count > 0) ExternalDomainsWarningAndAutoChangeToBccSetting = extList[0];

            // Cài đặt tệp đính kèm
            AttachmentsSetting = new AttachmentsSetting();
            var attList = CsvFileHandler.ReadCsv<AttachmentsSetting>(typeof(AttachmentsSettingMap), "AttachmentsSetting.csv");
            if (attList.Count > 0) AttachmentsSetting = attList[0];

            if (string.IsNullOrEmpty(AttachmentsSetting.TargetAttachmentFileExtensionOfOpen))
                AttachmentsSetting.TargetAttachmentFileExtensionOfOpen = ".pdf,.txt,.csv,.rtf,.htm,.html,.doc,.docx,.xls,.xlm,.xlsm,.xlsx,.ppt,.pptx,.bmp,.gif,.jpg,.jpeg,.png,.tif,.pub,.vsd,.vsdx";

            RecipientsAndAttachmentsNameList = CsvFileHandler.ReadCsv<RecipientsAndAttachmentsName>(typeof(RecipientsAndAttachmentsNameMap), "RecipientsAndAttachmentsName.csv")
                .Where(x => !string.IsNullOrEmpty(x.Recipient) && !string.IsNullOrEmpty(x.AttachmentsName)).ToList();

            AttachmentProhibitedRecipientsList = CsvFileHandler.ReadCsv<AttachmentProhibitedRecipients>(typeof(AttachmentProhibitedRecipientsMap), "AttachmentProhibitedRecipients.csv")
                .Where(x => !string.IsNullOrEmpty(x.Recipient)).ToList();

            AttachmentAlertRecipientsList = CsvFileHandler.ReadCsv<AttachmentAlertRecipients>(typeof(AttachmentAlertRecipientsMap), "AttachmentAlertRecipients.csv")
                .Where(x => !string.IsNullOrEmpty(x.Recipient)).ToList();

            // Bắt buộc tự động chuyển sang BCC
            ForceAutoChangeRecipientsToBccSetting = new ForceAutoChangeRecipientsToBcc();
            var forceList = CsvFileHandler.ReadCsv<ForceAutoChangeRecipientsToBcc>(typeof(ForceAutoChangeRecipientsToBccMap), "ForceAutoChangeRecipientsToBcc.csv");
            if (forceList.Count > 0) ForceAutoChangeRecipientsToBccSetting = forceList[0];

            // Tự động thêm tin nhắn
            AutoAddMessageSetting = new AutoAddMessage();
            var autoMsgList = CsvFileHandler.ReadCsv<AutoAddMessage>(typeof(AutoAddMessageMap), "AutoAddMessage.csv");
            if (autoMsgList.Count > 0) AutoAddMessageSetting = autoMsgList[0];

            // Tự động xóa người nhận
            AutoDeleteRecipients = CsvFileHandler.ReadCsv<AutoDeleteRecipient>(typeof(AutoDeleteRecipientMap), "AutoDeleteRecipientList.csv")
                .Where(x => !string.IsNullOrEmpty(x.Recipient))
                .ToList();
        }
    }
}
