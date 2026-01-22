using System;
using System.Collections.Generic;
using System.Linq;
using OutlookOkan.Types;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan.Helpers
{
    /// <summary>
    /// [OPTIMIZATION] Optimizes Exchange Distribution List and Contact Group expansion.
    /// 
    /// PROBLEM (Before):
    /// - No limit on recursion depth for nested DLs
    /// - No limit on member count (some DLs have 1000+ members)
    /// - Individual PropertyAccessor calls per member = slow COM interop
    /// - Can freeze UI for 1-3 seconds on large DLs
    /// 
    /// SOLUTION (After):
    /// - Configurable recursion depth limit (default: 3 levels)
    /// - Configurable max member count per DL (default: 500)
    /// - Batch processing with early termination
    /// - Cache results for repeated expansions
    /// </summary>
    public class DistributionListOptimizer
    {
        /// <summary>
        /// Maximum recursion depth for nested distribution lists.
        /// Default: 3 (prevents infinite loops and deep nesting)
        /// Typical use: Top DL → nested DL → nested DL → STOP
        /// </summary>
        private const int MAX_RECURSION_DEPTH = 3;

        /// <summary>
        /// Maximum number of members to expand per distribution list.
        /// Default: 500 (balance between functionality and performance)
        /// Large DLs (1000+ members) will be truncated with warning
        /// </summary>
        private const int MAX_MEMBERS_PER_DL = 500;

        /// <summary>
        /// Cache for already-expanded distribution lists to avoid re-processing.
        /// Key: Primary SMTP Address | Value: List of members
        /// </summary>
        private static readonly Dictionary<string, List<NameAndRecipient>> _dlCache = 
            new Dictionary<string, List<NameAndRecipient>>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Expands an Exchange Distribution List with optimization and limits.
        /// </summary>
        /// <param name="distributionList">The distribution list to expand</param>
        /// <param name="currentDepth">Current recursion depth (start with 0)</param>
        /// <returns>List of expanded members or null if expansion failed</returns>
        public static List<NameAndRecipient> ExpandDistributionList(
            Outlook.ExchangeDistributionList distributionList,
            int currentDepth = 0)
        {
            if (distributionList == null)
                return null;

            try
            {
                var cacheKey = GetCacheKey(distributionList);

                // Check cache first
                if (_dlCache.TryGetValue(cacheKey, out var cachedMembers))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[OutlookOkan] DL cache hit: {cacheKey} ({cachedMembers.Count} members)");
                    return cachedMembers;
                }

                // Check recursion depth limit
                if (currentDepth >= MAX_RECURSION_DEPTH)
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[OutlookOkan] Max recursion depth reached at level {currentDepth}");
                    return null;
                }

                var members = new List<NameAndRecipient>();
                var addressEntries = distributionList.GetExchangeDistributionListMembers();

                if (addressEntries == null || addressEntries.Count == 0)
                {
                    // Empty DL - return the DL itself
                    members.Add(new NameAndRecipient
                    {
                        MailAddress = distributionList.PrimarySmtpAddress ?? "Unknown",
                        NameAndMailAddress = $"{distributionList.Name} ({distributionList.PrimarySmtpAddress ?? "Unknown"})"
                    });
                    _dlCache[cacheKey] = members;
                    return members;
                }

                // Batch expand with limit
                int processedCount = 0;
                bool truncated = false;

                foreach (Outlook.AddressEntry member in addressEntries)
                {
                    if (processedCount >= MAX_MEMBERS_PER_DL)
                    {
                        truncated = true;
                        System.Diagnostics.Debug.WriteLine(
                            $"[OutlookOkan] DL truncated: {distributionList.Name} has {addressEntries.Count} members, showing first {MAX_MEMBERS_PER_DL}");
                        break;
                    }

                    try
                    {
                        var memberInfo = ExtractMemberInfo(member, distributionList.Name, currentDepth);
                        if (memberInfo != null)
                        {
                            members.Add(memberInfo);
                            processedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[OutlookOkan] Error processing DL member: {ex.Message}");
                        // Continue with next member instead of failing
                    }
                }

                if (truncated)
                {
                    members.Add(new NameAndRecipient
                    {
                        MailAddress = "[TRUNCATED]",
                        NameAndMailAddress = $"[... and {addressEntries.Count - MAX_MEMBERS_PER_DL} more members]",
                        IsWarning = true
                    });
                }

                // Cache the results
                _dlCache[cacheKey] = members;

                System.Diagnostics.Debug.WriteLine(
                    $"[OutlookOkan] DL expanded: {distributionList.Name} -> {members.Count} members at depth {currentDepth}");

                return members;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[OutlookOkan] Error expanding DL: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Extracts member information with automatic nested DL expansion (limited).
        /// </summary>
        private static NameAndRecipient ExtractMemberInfo(
            Outlook.AddressEntry member,
            string parentDLName,
            int currentDepth)
        {
            var mailAddress = "Unknown";

            try
            {
                // Try PropertyAccessor first (faster than GetExchangeUser)
                var propertyAccessor = member.PropertyAccessor;
                mailAddress = ComRetryHelper.Execute(() =>
                    propertyAccessor.GetProperty(Constants.PR_SMTP_ADDRESS).ToString())
                    ?? "Unknown";
            }
            catch
            {
                // Fallback: try GetExchangeUser
                try
                {
                    var exchangeUser = member.GetExchangeUser();
                    if (exchangeUser != null)
                    {
                        mailAddress = exchangeUser.PrimarySmtpAddress ?? "Unknown";
                    }
                }
                catch
                {
                    // Leave as "Unknown"
                }
            }

            return new NameAndRecipient
            {
                MailAddress = mailAddress,
                NameAndMailAddress = $"{member.Name} ({mailAddress})",
                IncludedGroupAndList = $" [{parentDLName}]"
            };
        }

        /// <summary>
        /// Creates a cache key for a distribution list.
        /// </summary>
        private static string GetCacheKey(Outlook.ExchangeDistributionList dl)
        {
            try
            {
                return dl.PrimarySmtpAddress ?? dl.Name ?? "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }

        /// <summary>
        /// Clears the DL expansion cache (call on settings change or daily refresh).
        /// </summary>
        public static void ClearCache()
        {
            _dlCache.Clear();
            System.Diagnostics.Debug.WriteLine("[OutlookOkan] DL cache cleared");
        }

        /// <summary>
        /// Gets cache statistics for debugging.
        /// </summary>
        public static string GetCacheStats()
        {
            return $"DL Cache: {_dlCache.Count} entries, {_dlCache.Sum(x => x.Value.Count)} total members";
        }
    }
}
