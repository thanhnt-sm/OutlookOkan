using System.Collections;
using System.Collections.Generic;
using OutlookOkan.Types;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan.Services.Interfaces
{
    public interface IRecipientResolver
    {
        CheckList GetSenderAndSenderDomain<T>(T item, CheckList checkList);

        DisplayNameAndRecipient MakeDisplayNameAndRecipient(IEnumerable recipients, DisplayNameAndRecipient displayNameAndRecipient, GeneralSetting generalSetting, bool isMeetingItem);



        List<string> MakeContactsList(Outlook.MAPIFolder contacts);

        List<Outlook.Recipient> AutoAddCcAndBcc<T>(T item, GeneralSetting generalSetting, DisplayNameAndRecipient displayNameAndRecipient, List<AutoCcBccKeyword> autoCcBccKeywordList, List<AutoCcBccAttachedFile> autoCcBccAttachedFilesList, List<AutoCcBccRecipient> autoCcBccRecipientList, int externalDomainCount, string sender, bool isAutoAddSenderToBcc, bool isAutoAddSenderToCc);
    }
}
