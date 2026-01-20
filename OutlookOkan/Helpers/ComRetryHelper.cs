using System;
using System.Runtime.InteropServices;
using System.Threading;
using OutlookOkan.Types;

namespace OutlookOkan.Helpers
{
    /// <summary>
    /// Provides standardized retry logic for COM operations that might fail
    /// due to Outlook being busy (RPC_E_CALL_REJECTED, etc).
    /// </summary>
    public static class ComRetryHelper
    {
        /// <summary>
        /// Executes a function that returns a value, with retry logic for COM exceptions.
        /// </summary>
        public static T Execute<T>(Func<T> action, T defaultValue = default)
        {
            var errorCount = 0;
            while (errorCount < Constants.MAX_RETRY_COUNT)
            {
                try
                {
                    return action();
                }
                catch (COMException e)
                {
                    if (IsRetryable(e))
                    {
                        Thread.Sleep(Constants.RETRY_DELAY_MS);
                        errorCount++;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[OutlookOkan] COM fatal error: {e}");
                        break;
                    }
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine($"[OutlookOkan] General error in retry helper: {e}");
                    break;
                }
            }
            return defaultValue;
        }

        /// <summary>
        /// Executes an action (void), with retry logic for COM exceptions.
        /// </summary>
        public static void Execute(Action action)
        {
            var errorCount = 0;
            while (errorCount < Constants.MAX_RETRY_COUNT)
            {
                try
                {
                    action();
                    return;
                }
                catch (COMException e)
                {
                    if (IsRetryable(e))
                    {
                        Thread.Sleep(Constants.RETRY_DELAY_MS);
                        errorCount++;
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[OutlookOkan] COM fatal error: {e}");
                        return;
                    }
                }
                catch (Exception e)
                {
                    System.Diagnostics.Debug.WriteLine($"[OutlookOkan] General error in retry helper: {e}");
                    return;
                }
            }
        }

        private static bool IsRetryable(COMException e)
        {
            // E_ABORT (0x80004004) or RPC errors often need retry
            // RPC_E_CALL_REJECTED (0x80010001)
            // RPC_E_SERVERCALL_RETRYLATER (0x8001010A)
            return e.ErrorCode == Constants.E_ABORT || 
                   e.ErrorCode == -2147418111 || // 0x80010001
                   e.ErrorCode == -2147417846;   // 0x8001010A
        }
    }
}
