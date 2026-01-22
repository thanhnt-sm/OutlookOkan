// ==============================================================================
// Outlook Stubs for Verification Build (FINAL v3)
// ==============================================================================
#if VERIFY_BUILD

#nullable disable
using System;
using System.Runtime.InteropServices;

[assembly: System.Resources.NeutralResourcesLanguage("en")]

namespace Microsoft.Office.Interop.Outlook
{
    [ComImport, Guid("00063001-0000-0000-C000-000000000046"), CoClass(typeof(ApplicationClass))]
    public interface Application
    {
        NameSpace Session { get; set; }
        NameSpace GetNamespace(string type);
        Inspectors Inspectors { get; set; }
        Explorer ActiveExplorer();
        object ActiveWindow();
        event ApplicationEvents_11_ItemSendEventHandler ItemSend;
    }

    public class ApplicationClass : Application
    {
        public NameSpace Session { get; set; }
        public NameSpace GetNamespace(string type) => Session;
        public Inspectors Inspectors { get; set; }
        public Explorer ActiveExplorer() => null;
        public object ActiveWindow() => null;
        public event ApplicationEvents_11_ItemSendEventHandler ItemSend;
    }

    public delegate void ApplicationEvents_11_ItemSendEventHandler(object Item, ref bool Cancel);
    public delegate void InspectorsEvents_NewInspectorEventHandler(Inspector Inspector);
    public delegate void ExplorerEvents_10_SelectionChangeEventHandler();
    public delegate void InspectorEvents_CloseEventHandler();
    public delegate void ItemEvents_10_BeforeAttachmentReadEventHandler(Attachment Attachment, ref bool Cancel);

    public interface NameSpace
    {
        MAPIFolder GetDefaultFolder(OlDefaultFolders folderType);
        Recipient CreateRecipient(string name);
        object GetItemFromID(string entryId);
        string CurrentUser { get; }
    }

    public interface Explorer
    {
        Selection Selection { get; }
        NameSpace Session { get; }
        MAPIFolder CurrentFolder { get; }
        event ExplorerEvents_10_SelectionChangeEventHandler SelectionChange;
    }

    public interface Inspector
    {
        object WordEditor { get; }
        object CurrentItem { get; }
        event InspectorEvents_CloseEventHandler Close;
    }

    public interface InspectorEvents_Event
    {
        event InspectorEvents_CloseEventHandler Close;
    }

    public interface Inspectors : System.Collections.IEnumerable
    {
        int Count { get; }
        event InspectorsEvents_NewInspectorEventHandler NewInspector;
    }

    public interface Selection : System.Collections.IEnumerable
    {
        int Count { get; }
        object this[int index] { get; }
    }

    public interface MailItem
    {
        string Subject { get; set; }
        string Body { get; set; }
        string HTMLBody { get; set; }
        object RTFBody { get; }
        string SenderEmailAddress { get; }
        string SenderEmailType { get; }
        string SenderName { get; }
        string SentOnBehalfOfName { get; }
        string EntryID { get; }
        AddressEntry Sender { get; }
        OlBodyFormat BodyFormat { get; }
        Recipients Recipients { get; }
        Attachments Attachments { get; }
        Account SendUsingAccount { get; set; }
        int InternetCodepage { get; set; }
        bool Submitted { get; set; }
        System.DateTime DeferredDeliveryTime { get; set; }
        void Send();
        void Close(OlInspectorClose saveMode);
        void Save();
        PropertyAccessor PropertyAccessor { get; }
        event ItemEvents_10_BeforeAttachmentReadEventHandler BeforeAttachmentRead;
    }

    public interface MeetingItem
    {
        string Subject { get; }
        string Body { get; }
        object RTFBody { get; }
        Recipients Recipients { get; }
        Attachments Attachments { get; }
    }

    public interface TaskItem
    {
        string Subject { get; }
        string Body { get; }
        object RTFBody { get; }
        Recipients Recipients { get; }
        Account SendUsingAccount { get; }
    }

    public interface TaskRequestItem
    {
        TaskItem GetAssociatedTask(bool addToTaskList);
    }

    public interface ContactItem
    {
        string Email1Address { get; }
        string Email2Address { get; }
        string Email3Address { get; }
        string FullName { get; }
    }

    public interface DistListItem
    {
        string DLName { get; }
        int MemberCount { get; }
        Recipient GetMember(int index);
    }
    
    public interface AppointmentItem
    {
        string Subject { get; }
        string Body { get; }
        Recipients Recipients { get; }
        int InternetCodepage { get; set; }
    }

    public interface Attachments : System.Collections.IEnumerable
    {
        int Count { get; }
        Attachment this[int index] { get; }
    }

    public interface Attachment
    {
        string FileName { get; }
        string DisplayName { get; }
        OlAttachmentType Type { get; }
        int Size { get; }
        string PathName { get; }
        void SaveAsFile(string path);
        PropertyAccessor PropertyAccessor { get; }
    }

    public interface Recipients : System.Collections.IEnumerable
    {
        int Count { get; }
        Recipient this[int index] { get; }
        Recipient Add(string name);
        bool ResolveAll();
        void Remove(int index);
    }

    public interface Recipient
    {
        string Name { get; }
        string Address { get; }
        int Type { get; set; }
        bool Sendable { get; }
        bool Resolve();
        AddressEntry AddressEntry { get; }
        PropertyAccessor PropertyAccessor { get; }
        int Index { get; }
        void Delete();
    }

    public interface AddressEntries : System.Collections.IEnumerable
    {
        int Count { get; }
        AddressEntry this[int index] { get; }
    }

    public interface AddressEntry
    {
        string ID { get; }
        string Address { get; }
        string Name { get; }
        OlAddressEntryUserType AddressEntryUserType { get; }
        ExchangeUser GetExchangeUser();
        ExchangeDistributionList GetExchangeDistributionList();
        PropertyAccessor PropertyAccessor { get; }
    }

    public interface ExchangeUser
    {
        string PrimarySmtpAddress { get; }
        string Name { get; }
        string Alias { get; }
    }

    public interface ExchangeDistributionList
    {
        string PrimarySmtpAddress { get; }
        string Name { get; }
        AddressEntries GetExchangeDistributionListMembers();
    }

    public interface MAPIFolder
    {
        Items Items { get; }
        string Name { get; }
        MAPIFolder Parent { get; }
        Folders Folders { get; }
    }

    public interface Folders : System.Collections.IEnumerable
    {
        int Count { get; }
        MAPIFolder this[int index] { get; }
    }

    public interface Items : System.Collections.IEnumerable
    {
        int Count { get; }
        object this[int index] { get; }
    }

    public interface PropertyAccessor
    {
        object GetProperty(string schemaName);
        void SetProperty(string schemaName, object value);
    }

    public interface Account
    {
        string SmtpAddress { get; }
        string DisplayName { get; }
    }

    // ENUMS
    // ==========================================================================

    public enum OlBodyFormat
    {
        olFormatUnspecified = 0,
        olFormatPlain = 1,
        olFormatHTML = 2,
        olFormatRichText = 3
    }

    public enum OlAddressEntryUserType
    {
        olExchangeUserAddressEntry = 0,
        olExchangeDistributionListAddressEntry = 1,
        olExchangePublicFolderAddressEntry = 2,
        olExchangeAgentAddressEntry = 3,
        olExchangeOrganizationAddressEntry = 4,
        olOutlookContactAddressEntry = 10,
        olOutlookDistributionListAddressEntry = 11,
        olLdapAddressEntry = 20,
        olSmtpAddressEntry = 30,
        olOtherAddressEntry = 40,
        olExchangeRemoteUserAddressEntry = 5
    }

    public enum OlAttachmentType
    {
        olByValue = 1,
        olByReference = 4,
        olEmbeddeditem = 5,
        olOLE = 6
    }
    
    public enum OlMailRecipientType
    {
        olOriginator = 0,
        olTo = 1,
        olCC = 2,
        olBCC = 3
    }

    public enum OlDefaultFolders
    {
        olFolderDeletedItems = 3,
        olFolderOutbox = 4,
        olFolderSentMail = 5,
        olFolderInbox = 6,
        olFolderContacts = 10,
        olFolderDrafts = 16,
        olFolderJournal = 11,
        olFolderNotes = 12,
        olFolderRssFeeds = 25,
        olFolderServerFailures = 22,
        olFolderLocalFailures = 21,
        olFolderSyncIssues = 20,
        olFolderTasks = 13,
        olFolderToDo = 28,
        olFolderCalendar = 9
    }

    public enum OlInspectorClose
    {
        olSave = 0,
        olDiscard = 1,
        olPromptForSave = 2
    }
}

namespace Microsoft.Office.Core
{
    public enum MsoTriState
    {
        msoTrue = -1,
        msoFalse = 0,
        msoCTrue = 1,
        msoTriStateToggle = -3,
        msoTriStateMixed = -2
    }

    public enum MsoAutomationSecurity
    {
        msoAutomationSecurityLow = 1,
        msoAutomationSecurityByUI = 2,
        msoAutomationSecurityForceDisable = 3
    }
    
    public interface IRibbonExtensibility
    {
        string GetCustomUI(string RibbonID);
    }
    
    public interface IRibbonUI
    {
        void Invalidate();
        void InvalidateControl(string ControlID);
    }
    
    public interface IRibbonControl
    {
        string Id { get; }
        object Context { get; }
        string Tag { get; }
    }
}

namespace Microsoft.Office.Interop.Word
{
    [ComImport, Guid("00020970-0000-0000-C000-000000000046"), CoClass(typeof(ApplicationClass))]
    public interface Application
    {
        Application Application { get; }
        bool Visible { get; set; }
        WordDocuments Documents { get; set; }
        Microsoft.Office.Core.MsoAutomationSecurity AutomationSecurity { get; set; }
    }
    
    public class ApplicationClass : Application
    {
        public Application Application => this;
        public bool Visible { get; set; }
        public WordDocuments Documents { get; set; }
        public Microsoft.Office.Core.MsoAutomationSecurity AutomationSecurity { get; set; }
    }

    public interface WordDocuments : System.Collections.IEnumerable
    {
        Document Open(string fileName, object confirmConversions = null, object readOnly = null,
            object addToRecentFiles = null, object passwordDocument = null, object passwordTemplate = null,
            object revert = null, object writePasswordDocument = null, object writePasswordTemplate = null,
            object format = null, object encoding = null, object visible = null, object openAndRepair = null,
            object documentDirection = null, object noEncodingDialog = null, object xmlTransform = null,
            object PasswordDocument = null, object Visible = null);
    }

    public interface Document
    {
        bool HasVBProject { get; }
        void Close(object saveChanges = null);
        Range Range(object start = null, object end = null);
    }
    
    public interface Range
    {
        string Text { get; }
        int Delete(object unit = null, object count = null);
        void InsertBefore(string text);
        void InsertAfter(string text);
        Range InsertParagraphAfter();
    }
}

namespace Microsoft.Office.Interop.Excel
{
    [ComImport, Guid("000208D5-0000-0000-C000-000000000046"), CoClass(typeof(ApplicationClass))]
    public interface Application
    {
        Application Application { get; }
        bool Visible { get; set; }
        bool EnableEvents { get; set; }
        Workbooks Workbooks { get; set; }
        Microsoft.Office.Core.MsoAutomationSecurity AutomationSecurity { get; set; }
    }
    
    public class ApplicationClass : Application
    {
        public Application Application => this;
        public bool Visible { get; set; }
        public bool EnableEvents { get; set; }
        public Workbooks Workbooks { get; set; }
        public Microsoft.Office.Core.MsoAutomationSecurity AutomationSecurity { get; set; }
    }

    public interface Workbooks
    {
        Workbook Open(string fileName, object updateLinks = null, object readOnly = null,
            object format = null, object password = null, object writeResPassword = null,
            object ignoreReadOnlyRecommended = null, object origin = null, object delimiter = null,
            object editable = null, object notify = null, object converter = null, object addToMru = null,
            object local = null, object corruptLoad = null, object Password = null);
    }

    public interface Workbook
    {
        bool HasVBProject { get; }
        void Close(object saveChanges = null);
    }
}

namespace Microsoft.Office.Interop.PowerPoint
{
    public class Application
    {
        public Presentations Presentations { get; set; }
    }

    public interface Presentations
    {
        Presentation Open(string fileName, Microsoft.Office.Core.MsoTriState readOnly,
            Microsoft.Office.Core.MsoTriState untitled, Microsoft.Office.Core.MsoTriState withWindow);
    }

    public interface Presentation
    {
        bool HasVBProject { get; }
        void Close();
    }
}

namespace Microsoft.VisualStudio.Tools.Applications.Runtime
{
    // Stub for VSTO runtime
}

namespace Microsoft.Office.Tools
{
    // Stub for Office tools
}

namespace Microsoft.Office.Tools.Outlook
{
    public class AddInBase
    {
        public Microsoft.Office.Interop.Outlook.Application Application { get; set; }
        
        protected virtual Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return null;
        }

        public event System.EventHandler Startup;
        public event System.EventHandler Shutdown;
    }
}

#endif
