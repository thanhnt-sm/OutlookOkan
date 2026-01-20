namespace OutlookOkan.Types
{
    public static class ComErrorCodes
    {
        /// <summary>
        /// RPC_E_CALL_REJECTED (0x80010001)
        /// Occurs when the COM server (Outlook) is busy.
        /// </summary>
        public const int RpcECallRejected = -2147418111;

        /// <summary>
        /// MK_E_UNAVAILABLE (0x800401E3)
        /// Operation unavailable.
        /// </summary>
        public const int MkEUnavailable = -2147221021;

        /// <summary>
        /// E_ABORT (0x80004004)
        /// Operation aborted.
        /// </summary>
        public const int EAbort = -2147467260;
        
        /// <summary>
        /// E_FAIL (0x80004005)
        /// Unspecified failure.
        /// </summary>
        public const int EFail = -2147467259;
    }
}
