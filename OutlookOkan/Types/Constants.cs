namespace OutlookOkan.Types
{
    public static class Constants
    {
        // MAPI Property Tags
        public const string PR_TRANSPORT_MESSAGE_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";
        public const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

        // Retry Logic
        public const int MAX_RETRY_COUNT = 100;
        public const int RETRY_DELAY_MS = 10;
        public const int RETRY_LONG_DELAY_MS = 20;

        // Error Codes
        public const int RPC_E_CALL_REJECTED = -2147418111; // 0x80010001 (Sometimes defined as standard COM error)
        public const int RPC_E_SERVERCALL_RETRYLATER = -2147417846; // 0x8001010A
        // Note: 0x80004004 (E_ABORT) was used in original code
        public const int E_ABORT = -2147467260; 
    }
}
