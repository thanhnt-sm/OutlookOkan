using System.Collections.Generic;
using OutlookOkan.Types;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan.Services.Interfaces
{
    public interface IEmailAnalyzer
    {
        string GetMailBody(Outlook.OlBodyFormat mailBodyFormat, string mailBody);
        
        string GetMailBodyFormat(Outlook.OlBodyFormat bodyFormat);

        string AddMessageToBodyPreview(string mailBody, AutoAddMessage autoAddMessage);

        CheckList GetAttachmentsInformation<T>(T item, CheckList checkList, bool isNotTreatedAsAttachmentsAtHtmlEmbeddedFiles, AttachmentsSetting attachmentsSetting, string mailHtmlBody, bool isAutoCheckAttachments);
        
        CheckList CheckForgotAttach(CheckList checkList, GeneralSetting generalSetting);
        
        CheckList CheckKeyword(CheckList checkList, List<AlertKeywordAndMessage> alertKeywordAndMessageList);
        
        CheckList CheckKeywordForSubject(CheckList checkList, List<AlertKeywordAndMessageForSubject> alertKeywordAndMessageForSubjectList);
        
        int CalcDeferredMinutes(DisplayNameAndRecipient displayNameAndRecipient, List<DeferredDeliveryMinutes> deferredDeliveryMinutesList, bool isDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain, int recipientExternalDomainNumAll);
    }
}
