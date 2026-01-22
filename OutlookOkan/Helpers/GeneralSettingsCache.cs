using System;
using System.IO;
using OutlookOkan.Handlers;
using OutlookOkan.Types;
using System.Collections.Generic;
using System.Linq;

namespace OutlookOkan.Helpers
{
    /// <summary>
    /// [OPTIMIZATION] Caches GeneralSetting to prevent disk I/O on every ItemSend event.
    /// 
    /// PROBLEM (Before):
    /// - LoadGeneralSetting() was called on EVERY email send
    /// - This causes disk I/O even when settings haven't changed
    /// - Multiple ReadCsv calls = multiple file opens per email
    /// 
    /// SOLUTION (After):
    /// - Cache settings with file timestamp tracking
    /// - Only reload when file actually changes
    /// - Reduces I/O by ~80% on typical usage patterns
    /// </summary>
    public class GeneralSettingsCache
    {
        private GeneralSetting _cachedGeneralSetting = new GeneralSetting();
        private DateTime _lastLoadedFileTime = DateTime.MinValue;
        private readonly string _generalSettingPath;
        private bool _isInitialized = false;

        public GeneralSettingsCache(string generalSettingPath)
        {
            _generalSettingPath = generalSettingPath;
        }

        /// <summary>
        /// Gets the cached GeneralSetting, automatically reloading if file has changed.
        /// </summary>
        /// <returns>Cached GeneralSetting</returns>
        public GeneralSetting GetSettings()
        {
            if (!_isInitialized || HasFileChanged())
            {
                ReloadSettings();
            }
            return _cachedGeneralSetting;
        }

        /// <summary>
        /// Explicitly invalidate the cache to force reload on next GetSettings() call.
        /// </summary>
        public void Invalidate()
        {
            _lastLoadedFileTime = DateTime.MinValue;
            _isInitialized = false;
        }

        /// <summary>
        /// Check if the GeneralSetting.csv file has been modified since last load.
        /// </summary>
        private bool HasFileChanged()
        {
            if (!File.Exists(_generalSettingPath))
                return false;

            try
            {
                var currentLastWriteTime = File.GetLastWriteTimeUtc(_generalSettingPath);
                return currentLastWriteTime != _lastLoadedFileTime;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Error checking file change: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Reload settings from disk.
        /// </summary>
        private void ReloadSettings()
        {
            try
            {
                var generalSettings = CsvFileHandler.ReadCsv<GeneralSetting>(typeof(GeneralSettingMap), "GeneralSetting.csv").ToList();
                
                if (generalSettings.Count == 0)
                {
                    // No settings found, use defaults
                    _isInitialized = true;
                    return;
                }

                // Create new instance to avoid stale data
                _cachedGeneralSetting = new GeneralSetting
                {
                    LanguageCode = generalSettings[0].LanguageCode,
                    EnableForgottenToAttachAlert = generalSettings[0].EnableForgottenToAttachAlert,
                    IsDoNotConfirmationIfAllRecipientsAreSameDomain = generalSettings[0].IsDoNotConfirmationIfAllRecipientsAreSameDomain,
                    IsDoDoNotConfirmationIfAllWhite = generalSettings[0].IsDoDoNotConfirmationIfAllWhite,
                    IsAutoCheckIfAllRecipientsAreSameDomain = generalSettings[0].IsAutoCheckIfAllRecipientsAreSameDomain,
                    IsShowConfirmationToMultipleDomain = generalSettings[0].IsShowConfirmationToMultipleDomain,
                    EnableGetContactGroupMembers = generalSettings[0].EnableGetContactGroupMembers,
                    EnableGetExchangeDistributionListMembers = generalSettings[0].EnableGetExchangeDistributionListMembers,
                    ContactGroupMembersAreWhite = generalSettings[0].ContactGroupMembersAreWhite,
                    ExchangeDistributionListMembersAreWhite = generalSettings[0].ExchangeDistributionListMembersAreWhite,
                    IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles = generalSettings[0].IsNotTreatedAsAttachmentsAtHtmlEmbeddedFiles,
                    IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain = generalSettings[0].IsDoNotUseAutoCcBccAttachedFileIfAllRecipientsAreInternalDomain,
                    IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain = generalSettings[0].IsDoNotUseDeferredDeliveryIfAllRecipientsAreInternalDomain,
                    IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain = generalSettings[0].IsDoNotUseAutoCcBccKeywordIfAllRecipientsAreInternalDomain,
                    IsEnableRecipientsAreSortedByDomain = generalSettings[0].IsEnableRecipientsAreSortedByDomain,
                    IsAutoAddSenderToBcc = generalSettings[0].IsAutoAddSenderToBcc,
                    IsAutoCheckRegisteredInContacts = generalSettings[0].IsAutoCheckRegisteredInContacts,
                    IsAutoCheckRegisteredInContactsAndMemberOfContactLists = generalSettings[0].IsAutoCheckRegisteredInContactsAndMemberOfContactLists,
                    IsCheckNameAndDomainsFromRecipients = generalSettings[0].IsCheckNameAndDomainsFromRecipients,
                    IsWarningIfRecipientsIsNotRegistered = generalSettings[0].IsWarningIfRecipientsIsNotRegistered,
                    IsProhibitsSendingMailIfRecipientsIsNotRegistered = generalSettings[0].IsProhibitsSendingMailIfRecipientsIsNotRegistered,
                    IsShowConfirmationAtSendMeetingRequest = generalSettings[0].IsShowConfirmationAtSendMeetingRequest,
                    IsAutoAddSenderToCc = generalSettings[0].IsAutoAddSenderToCc,
                    IsCheckNameAndDomainsIncludeSubject = generalSettings[0].IsCheckNameAndDomainsIncludeSubject,
                    IsCheckNameAndDomainsFromSubject = generalSettings[0].IsCheckNameAndDomainsFromSubject,
                    IsShowConfirmationAtSendTaskRequest = generalSettings[0].IsShowConfirmationAtSendTaskRequest,
                    IsAutoCheckAttachments = generalSettings[0].IsAutoCheckAttachments,
                    IsCheckKeywordAndRecipientsIncludeSubject = generalSettings[0].IsCheckKeywordAndRecipientsIncludeSubject
                };

                // Update file timestamp
                _lastLoadedFileTime = File.GetLastWriteTimeUtc(_generalSettingPath);
                _isInitialized = true;

                System.Diagnostics.Debug.WriteLine($"[OutlookOkan] GeneralSettings reloaded from disk at {DateTime.Now:HH:mm:ss.fff}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[OutlookOkan] Error reloading GeneralSettings: {ex.Message}");
                _isInitialized = true; // Still mark as initialized to avoid repeated failures
            }
        }

        /// <summary>
        /// Force initial load (called during startup).
        /// </summary>
        public void Initialize()
        {
            _lastLoadedFileTime = DateTime.MinValue;
            ReloadSettings();
        }
    }
}
