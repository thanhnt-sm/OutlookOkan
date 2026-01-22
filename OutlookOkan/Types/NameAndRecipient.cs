namespace OutlookOkan.Types
{
    public sealed class NameAndRecipient
    {
        public string MailAddress { get; set; }
        public string NameAndMailAddress { get; set; }
        public string IncludedGroupAndList { get; set; }
        
        /// <summary>
        /// [OPTIMIZATION] Flag for truncation warning when DL has too many members
        /// </summary>
        public bool IsWarning { get; set; } = false;
    }
}