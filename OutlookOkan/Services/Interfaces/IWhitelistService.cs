using System.Collections.Generic;
using OutlookOkan.Types;

namespace OutlookOkan.Services.Interfaces
{
    public interface IWhitelistService
    {
        int CountRecipientExternalDomains(DisplayNameAndRecipient displayNameAndRecipient, string senderDomain, IEnumerable<InternalDomain> internalDomain, bool isToAndCcOnly);

        DisplayNameAndRecipient ExternalDomainsChangeToBccIfNeeded(CheckList checkList, object item, DisplayNameAndRecipient displayNameAndRecipient, ExternalDomainsWarningAndAutoChangeToBcc externalDomainsWarningAndAutoChangeToBccSetting, IEnumerable<InternalDomain> internalDomainList, int externalDomainsCount, string senderDomain, string sender, ForceAutoChangeRecipientsToBcc forceAutoChangeRecipientsToBccSetting);

        CheckList ExternalDomainsWarningIfNeeded(CheckList checkList, ExternalDomainsWarningAndAutoChangeToBcc externalDomainsWarningAndAutoChangeToBccSetting, int externalDomainsCount, bool isForceAutoChangeRecipientsToBcc);

        CheckList CheckMailBodyAndRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, List<NameAndDomains> nameAndDomainsList, bool isCheckNameAndDomainsFromRecipients, bool isCheckNameAndDomainsIncludeSubject, bool isCheckNameAndDomainsFromSubject);

        CheckList CheckKeywordAndRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, List<KeywordAndRecipients> keywordAndRecipientsList, bool isCheckKeywordAndRecipientsIncludeSubject);

        CheckList GetRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, List<AlertAddress> alertAddressList, List<InternalDomain> internalDomainList);

        CheckList CheckRecipientsAndAttachments(CheckList checkList, bool isAttachmentsProhibited, bool isWarningWhenAttachedRealFile, List<AttachmentProhibitedRecipients> attachmentProhibitedRecipientsList, List<RecipientsAndAttachmentsName> recipientsAndAttachmentsNameList, List<AttachmentAlertRecipients> attachmentAlertRecipientsList);

        CheckList AddAlertOrProhibitsSendingMailIfIfRecipientsIsNotRegistered(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, List<string> contactsList, List<InternalDomain> internalDomainList, bool isWarningIfRecipientsIsNotRegistered, bool isProhibitsSendingMailIfRecipientsIsNotRegistered);

        CheckList AutoCheckRegisteredItemsInContacts(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, List<string> contactsList, bool isAutoCheckRegisteredInContacts);

        bool IsWhitelisted(string address);
        void AddToWhitelist(string address, bool isSkipConfirmation);
    }
}
