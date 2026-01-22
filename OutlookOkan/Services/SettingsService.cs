using OutlookOkan.Handlers;
using OutlookOkan.Models;
using OutlookOkan.Types;
using System.Collections.Generic;
using System.Linq;
using System; // Added for StringComparer and DateTime

namespace OutlookOkan.Services
{
    public class SettingsService
    {
        public Dictionary<string, bool> Whitelist { get; private set; } = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase); // Changed to Dictionary<string, bool> to store IsSkipConfirmation

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

        private readonly Dictionary<string, System.DateTime> _fileTimestamps = new Dictionary<string, System.DateTime>();
        private readonly string _basePath = CsvFileHandler.DirectoryPath;

        public void LoadSettings()
        {
            // Load Whitelist optimized as Dictionary<string, bool>
            LoadIfChanged("Whitelist.csv", () =>
            {
                var rawList = CsvFileHandler.ReadCsv<Whitelist>(typeof(WhitelistMap), "Whitelist.csv");
                Whitelist = rawList
                    .Where(x => !string.IsNullOrEmpty(x.WhiteName))
                    .GroupBy(x => x.WhiteName, StringComparer.OrdinalIgnoreCase) // Deduplicate
                    .ToDictionary(g => g.Key, g => g.First().IsSkipConfirmation, StringComparer.OrdinalIgnoreCase);
            });

            LoadIfChanged("AlertKeywordAndMessageList.csv", () => AlertKeywordAndMessageList = CsvFileHandler.ReadCsv<AlertKeywordAndMessage>(typeof(AlertKeywordAndMessageMap), "AlertKeywordAndMessageList.csv").Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).ToList());
            LoadIfChanged("AlertKeywordAndMessageListForSubject.csv", () => AlertKeywordAndMessageForSubjectList = CsvFileHandler.ReadCsv<AlertKeywordAndMessageForSubject>(typeof(AlertKeywordAndMessageForSubjectMap), "AlertKeywordAndMessageListForSubject.csv").Where(x => !string.IsNullOrEmpty(x.AlertKeyword)).ToList());
            LoadIfChanged("AutoCcBccKeywordList.csv", () => AutoCcBccKeywordList = CsvFileHandler.ReadCsv<AutoCcBccKeyword>(typeof(AutoCcBccKeywordMap), "AutoCcBccKeywordList.csv").Where(x => !string.IsNullOrEmpty(x.AutoAddAddress) && !string.IsNullOrEmpty(x.Keyword)).ToList());
            LoadIfChanged("AutoCcBccAttachedFileList.csv", () => AutoCcBccAttachedFilesList = CsvFileHandler.ReadCsv<AutoCcBccAttachedFile>(typeof(AutoCcBccAttachedFileMap), "AutoCcBccAttachedFileList.csv").Where(x => !string.IsNullOrEmpty(x.AutoAddAddress)).ToList());
            LoadIfChanged("AutoCcBccRecipientList.csv", () => AutoCcBccRecipientList = CsvFileHandler.ReadCsv<AutoCcBccRecipient>(typeof(AutoCcBccRecipientMap), "AutoCcBccRecipientList.csv").Where(x => !string.IsNullOrEmpty(x.AutoAddAddress) && !string.IsNullOrEmpty(x.TargetRecipient)).ToList());
            LoadIfChanged("AlertAddressList.csv", () => AlertAddressList = CsvFileHandler.ReadCsv<AlertAddress>(typeof(AlertAddressMap), "AlertAddressList.csv").Where(x => !string.IsNullOrEmpty(x.TargetAddress)).ToList());
            LoadIfChanged("NameAndDomains.csv", () => NameAndDomainsList = CsvFileHandler.ReadCsv<NameAndDomains>(typeof(NameAndDomainsMap), "NameAndDomains.csv").Where(x => !string.IsNullOrEmpty(x.Domain) && !string.IsNullOrEmpty(x.Name)).ToList());
            LoadIfChanged("keywordAndRecipientsList.csv", () => KeywordAndRecipientsList = CsvFileHandler.ReadCsv<KeywordAndRecipients>(typeof(KeywordAndRecipientsMap), "keywordAndRecipientsList.csv").Where(x => !string.IsNullOrEmpty(x.Keyword) && !string.IsNullOrEmpty(x.Recipient)).ToList());
            LoadIfChanged("DeferredDeliveryMinutes.csv", () => DeferredDeliveryMinutesList = CsvFileHandler.ReadCsv<DeferredDeliveryMinutes>(typeof(DeferredDeliveryMinutesMap), "DeferredDeliveryMinutes.csv").Where(x => !string.IsNullOrEmpty(x.TargetAddress)).ToList());
            LoadIfChanged("InternalDomainList.csv", () => InternalDomainList = CsvFileHandler.ReadCsv<InternalDomain>(typeof(InternalDomainMap), "InternalDomainList.csv").Where(x => !string.IsNullOrEmpty(x.Domain)).ToList());

            LoadIfChanged("ExternalDomainsWarningAndAutoChangeToBccSetting.csv", () =>
            {
                ExternalDomainsWarningAndAutoChangeToBccSetting = new ExternalDomainsWarningAndAutoChangeToBcc();
                var extList = CsvFileHandler.ReadCsv<ExternalDomainsWarningAndAutoChangeToBcc>(typeof(ExternalDomainsWarningAndAutoChangeToBccMap), "ExternalDomainsWarningAndAutoChangeToBccSetting.csv");
                if (extList.Count > 0) ExternalDomainsWarningAndAutoChangeToBccSetting = extList[0];
            });

            LoadIfChanged("AttachmentsSetting.csv", () =>
            {
                AttachmentsSetting = new AttachmentsSetting();
                var attList = CsvFileHandler.ReadCsv<AttachmentsSetting>(typeof(AttachmentsSettingMap), "AttachmentsSetting.csv");
                if (attList.Count > 0) AttachmentsSetting = attList[0];
                if (string.IsNullOrEmpty(AttachmentsSetting?.TargetAttachmentFileExtensionOfOpen))
                {
                    AttachmentsSetting = AttachmentsSetting ?? new AttachmentsSetting();
                    AttachmentsSetting.TargetAttachmentFileExtensionOfOpen = ".pdf,.txt,.csv,.rtf,.htm,.html,.doc,.docx,.xls,.xlm,.xlsm,.xlsx,.ppt,.pptx,.bmp,.gif,.jpg,.jpeg,.png,.tif,.pub,.vsd,.vsdx";
                }
            });

            LoadIfChanged("RecipientsAndAttachmentsName.csv", () => RecipientsAndAttachmentsNameList = CsvFileHandler.ReadCsv<RecipientsAndAttachmentsName>(typeof(RecipientsAndAttachmentsNameMap), "RecipientsAndAttachmentsName.csv").Where(x => !string.IsNullOrEmpty(x.Recipient) && !string.IsNullOrEmpty(x.AttachmentsName)).ToList());
            LoadIfChanged("AttachmentProhibitedRecipients.csv", () => AttachmentProhibitedRecipientsList = CsvFileHandler.ReadCsv<AttachmentProhibitedRecipients>(typeof(AttachmentProhibitedRecipientsMap), "AttachmentProhibitedRecipients.csv").Where(x => !string.IsNullOrEmpty(x.Recipient)).ToList());
            LoadIfChanged("AttachmentAlertRecipients.csv", () => AttachmentAlertRecipientsList = CsvFileHandler.ReadCsv<AttachmentAlertRecipients>(typeof(AttachmentAlertRecipientsMap), "AttachmentAlertRecipients.csv").Where(x => !string.IsNullOrEmpty(x.Recipient)).ToList());

            LoadIfChanged("ForceAutoChangeRecipientsToBcc.csv", () =>
            {
                ForceAutoChangeRecipientsToBccSetting = new ForceAutoChangeRecipientsToBcc();
                var forceList = CsvFileHandler.ReadCsv<ForceAutoChangeRecipientsToBcc>(typeof(ForceAutoChangeRecipientsToBccMap), "ForceAutoChangeRecipientsToBcc.csv");
                if (forceList.Count > 0) ForceAutoChangeRecipientsToBccSetting = forceList[0];
            });

            LoadIfChanged("AutoAddMessage.csv", () =>
            {
                AutoAddMessageSetting = new AutoAddMessage();
                var autoMsgList = CsvFileHandler.ReadCsv<AutoAddMessage>(typeof(AutoAddMessageMap), "AutoAddMessage.csv");
                if (autoMsgList.Count > 0) AutoAddMessageSetting = autoMsgList[0];
            });

            LoadIfChanged("AutoDeleteRecipientList.csv", () => AutoDeleteRecipients = CsvFileHandler.ReadCsv<AutoDeleteRecipient>(typeof(AutoDeleteRecipientMap), "AutoDeleteRecipientList.csv").Where(x => !string.IsNullOrEmpty(x.Recipient)).ToList());
        }

        private void LoadIfChanged(string fileName, System.Action loadAction)
        {
            var fullPath = System.IO.Path.Combine(_basePath, fileName);
            if (!System.IO.File.Exists(fullPath))
            {
                // File doesn't exist, execute load (which should handle empty/default) or skip?
                // CsvHandler.ReadCsv handles missing files by returning empty list.
                // But we should track that we "checked" it.
                // If it doesn't exist, maybe we don't cache timestamp (or cache DateTime.MinValue).
                // Let's run loadAction to ensure properties are initialized to empty.
                loadAction();
                return;
            }

            var currentWriteTime = System.IO.File.GetLastWriteTimeUtc(fullPath);
            if (_fileTimestamps.TryGetValue(fileName, out var cachedTime) && cachedTime == currentWriteTime)
            {
                return; // Not changed
            }

            loadAction();
            _fileTimestamps[fileName] = currentWriteTime;
        }
    }
}
