using System;
using System.Collections.Generic;
using System.Linq;
using OutlookOkan.Properties;
using OutlookOkan.Services.Interfaces;
using OutlookOkan.Types;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookOkan.Services
{
    public class WhitelistService : IWhitelistService
    {
        private readonly Dictionary<string, bool> _whitelist;

        public WhitelistService()
        {
            _whitelist = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        }

        public WhitelistService(Dictionary<string, bool> initialWhitelist)
        {
            _whitelist = new Dictionary<string, bool>(initialWhitelist, StringComparer.OrdinalIgnoreCase);
        }

        public bool IsWhitelisted(string address)
        {
            return _whitelist.ContainsKey(address);
        }

        public void AddToWhitelist(string address, bool isSkipConfirmation)
        {
            _whitelist[address] = isSkipConfirmation;
        }

        public int CountRecipientExternalDomains(DisplayNameAndRecipient displayNameAndRecipient, string senderDomain, IEnumerable<InternalDomain> internalDomain, bool isToAndCcOnly)
        {
            var domainList = new HashSet<string>();

            if (isToAndCcOnly)
            {
                foreach (var recipient in displayNameAndRecipient.To.Select(mail => mail.Key).Where(recipient => recipient != Resources.FailedToGetInformation && recipient.Contains("@")))
                {
                    _ = domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }

                foreach (var recipient in displayNameAndRecipient.Cc.Select(mail => mail.Key).Where(recipient => recipient != Resources.FailedToGetInformation && recipient.Contains("@")))
                {
                    _ = domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }
            }
            else
            {
                foreach (var recipient in displayNameAndRecipient.All.Select(mail => mail.Key).Where(recipient => recipient != Resources.FailedToGetInformation && recipient.Contains("@")))
                {
                    domainList.Add(recipient.Substring(recipient.IndexOf("@", StringComparison.Ordinal)));
                }
            }

            var externalDomainsCount = domainList.Count;

            foreach (var _ in internalDomain.Where(internalDomainSetting => domainList.Any(domain => domain.EndsWith(internalDomainSetting.Domain)) && !senderDomain.EndsWith(internalDomainSetting.Domain)))
            {
                externalDomainsCount--;
            }

            if (domainList.Contains(senderDomain))
            {
                return externalDomainsCount - 1;
            }

            return externalDomainsCount;
        }

        public DisplayNameAndRecipient ExternalDomainsChangeToBccIfNeeded(CheckList checkList, object item, DisplayNameAndRecipient displayNameAndRecipient, ExternalDomainsWarningAndAutoChangeToBcc externalDomainsWarningAndAutoChangeToBccSetting, IEnumerable<InternalDomain> internalDomainList, int externalDomainsCount, string senderDomain, string sender, ForceAutoChangeRecipientsToBcc forceAutoChangeRecipientsToBccSetting)
        {
            var internalDomains = internalDomainList.ToList();

            if ((!externalDomainsWarningAndAutoChangeToBccSetting.IsAutoChangeToBccWhenLargeNumberOfExternalDomains || externalDomainsWarningAndAutoChangeToBccSetting.IsProhibitedWhenLargeNumberOfExternalDomains || externalDomainsWarningAndAutoChangeToBccSetting.TargetToAndCcExternalDomainsNum > externalDomainsCount) && !forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc) return displayNameAndRecipient;

            if (forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc && forceAutoChangeRecipientsToBccSetting.IsIncludeInternalDomain)
            {
                internalDomains.Clear();
            }
            else
            {
                internalDomains.Add(new InternalDomain { Domain = senderDomain });
            }

            var removeTarget = new List<string>();

            foreach (var to in displayNameAndRecipient.To)
            {
                var isInternal = false;
                foreach (var _ in internalDomains.Where(internalDomain => to.Key.EndsWith(internalDomain.Domain)))
                {
                    isInternal = true;
                }
                if (isInternal) continue;

                displayNameAndRecipient.Bcc[to.Key] = to.Value;
                removeTarget.Add(to.Key);
            }
            foreach (var target in removeTarget)
            {
                _ = displayNameAndRecipient.To.Remove(target);
            }

            removeTarget.Clear();

            foreach (var cc in displayNameAndRecipient.Cc)
            {
                var isInternal = false;
                foreach (var _ in internalDomains.Where(internalDomain => cc.Key.EndsWith(internalDomain.Domain)))
                {
                    isInternal = true;
                }
                if (isInternal) continue;

                displayNameAndRecipient.Bcc[cc.Key] = cc.Value;
                removeTarget.Add(cc.Key);
            }
            foreach (var target in removeTarget)
            {
                _ = displayNameAndRecipient.Cc.Remove(target);
            }

            if (forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc)
            {
                AddAlerts(checkList, Resources.ForceAutoChangeRecipientsToBccAlert + $"[{externalDomainsWarningAndAutoChangeToBccSetting.TargetToAndCcExternalDomainsNum}]", false, false, true);
            }
            else
            {
                AddAlerts(checkList, Resources.ExternalDomainsChangeToBccAlert + $"[{externalDomainsWarningAndAutoChangeToBccSetting.TargetToAndCcExternalDomainsNum}]", true, false, false);
            }

            var isNeedsAddToSender = false;
            var senderMailAddress = sender; // Use sender parameter
            var thisSenderMailAddress = forceAutoChangeRecipientsToBccSetting.IsForceAutoChangeRecipientsToBcc && !string.IsNullOrEmpty(forceAutoChangeRecipientsToBccSetting.ToRecipient) ? forceAutoChangeRecipientsToBccSetting.ToRecipient : senderMailAddress;
            if (displayNameAndRecipient.To.Count == 0)
            {
                displayNameAndRecipient.To[thisSenderMailAddress] = thisSenderMailAddress;
                isNeedsAddToSender = true;

                AddAlerts(checkList, thisSenderMailAddress == senderMailAddress
                        ? Resources.AutoAddSendersAddressToAlert
                        : Resources.AutoAddToRecipientByForceAutoChangeRecipientsToBccAddressToAlert, true, false, false);
            }

            var targetMailRecipientsIndex = new List<MailItemsRecipientAndMailAddress>();
            foreach (var recipient in displayNameAndRecipient.MailRecipientsIndex.Where(recipient => recipient.Type != (int)Outlook.OlMailRecipientType.olBCC))
            {
                var isExternal = true;
                foreach (var _ in internalDomains.Where(internalDomain => recipient.MailAddress.EndsWith(internalDomain.Domain)))
                {
                    isExternal = false;
                }

                if (isExternal) targetMailRecipientsIndex.Add(recipient);
            }

            ChangeToBcc(item, targetMailRecipientsIndex, thisSenderMailAddress, isNeedsAddToSender);

            return displayNameAndRecipient;
        }

        public CheckList ExternalDomainsWarningIfNeeded(CheckList checkList, ExternalDomainsWarningAndAutoChangeToBcc externalDomainsWarningAndAutoChangeToBccSetting, int externalDomainsCount, bool isForceAutoChangeRecipientsToBcc)
        {
            if (isForceAutoChangeRecipientsToBcc) return checkList;

            if (externalDomainsCount < externalDomainsWarningAndAutoChangeToBccSetting.TargetCount || !externalDomainsWarningAndAutoChangeToBccSetting.IsWarning) return checkList;

            checkList.Alerts.Add(new Alert { AlertMessage = Resources.WarningExternalDomainsNumber + $" ({externalDomainsCount})", IsImportant = true, IsWhite = false, IsChecked = false });

            if (externalDomainsWarningAndAutoChangeToBccSetting.IsProhibited)
            {
                checkList.IsCanNotSendMail = true;
                checkList.CanNotSendMailMessage = Resources.WarningExternalDomainsNumber + $" ({externalDomainsCount})";
            }

            return checkList;
        }

        public CheckList CheckMailBodyAndRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, List<NameAndDomains> nameAndDomainsList, bool isCheckNameAndDomainsFromRecipients, bool isCheckNameAndDomainsIncludeSubject, bool isCheckNameAndDomainsFromSubject)
        {
            if (displayNameAndRecipient is null) return checkList;

            var cleanedNameAndDomains = nameAndDomainsList.Where(nameAndDomain => !string.IsNullOrEmpty(nameAndDomain.Domain) && !string.IsNullOrEmpty(nameAndDomain.Name)).ToList();
            if (!cleanedNameAndDomains.Any()) return checkList;

            if (isCheckNameAndDomainsFromRecipients || (isCheckNameAndDomainsIncludeSubject && isCheckNameAndDomainsFromSubject))
            {
                var domainCandidateRecipients = new List<string[]>();
                foreach (var nameAndDomain in cleanedNameAndDomains)
                {
                    foreach (var recipient in displayNameAndRecipient.All)
                    {
                        if (recipient.Key.EndsWith(nameAndDomain.Domain) || recipient.Key == nameAndDomain.Domain)
                        {
                            domainCandidateRecipients.Add(new[] { recipient.Value, nameAndDomain.Name });
                        }
                    }
                }

                if (isCheckNameAndDomainsFromRecipients)
                {
                    var domainCandidateRecipientsHitCount = domainCandidateRecipients.Where(domainAndName => !checkList.MailBody.Contains(domainAndName[1])).Count(domainAndName => !domainAndName[0].Contains(checkList.SenderDomain));
                    if (domainCandidateRecipientsHitCount >= domainCandidateRecipients.Count)
                    {
                        foreach (var domainAndName in domainCandidateRecipients.Where(domainAndName => !checkList.MailBody.Contains(domainAndName[1])).Where(domainAndName => !domainAndName[0].Contains(checkList.SenderDomain)))
                        {
                            checkList.Alerts.Add(new Alert
                            {
                                AlertMessage = $"{domainAndName[0]} : {Resources.CanNotFindTheLinkedName} ({domainAndName[1]})",
                                IsImportant = true,
                                IsWhite = false,
                                IsChecked = false
                            });
                        }
                    }
                }

                if (isCheckNameAndDomainsIncludeSubject && isCheckNameAndDomainsFromSubject)
                {
                    var domainCandidateRecipientsHitCount = domainCandidateRecipients.Where(domainAndName => !checkList.Subject.Contains(domainAndName[1])).Count(domainAndName => !domainAndName[0].Contains(checkList.SenderDomain));
                    if (domainCandidateRecipientsHitCount >= domainCandidateRecipients.Count)
                    {
                        foreach (var domainAndName in domainCandidateRecipients.Where(domainAndName => !checkList.Subject.Contains(domainAndName[1])).Where(domainAndName => !domainAndName[0].Contains(checkList.SenderDomain)))
                        {
                            checkList.Alerts.Add(new Alert
                            {
                                AlertMessage = $"{domainAndName[0]} : {Resources.CanNotFindTheLinkedNameInSubject} ({domainAndName[1]})",
                                IsImportant = true,
                                IsWhite = false,
                                IsChecked = false
                            });
                        }
                    }
                }
            }

            var targetText = checkList.MailBody;
            if (isCheckNameAndDomainsIncludeSubject) { targetText += checkList.Subject; }

            var recipientCandidateDomains = (from nameAndDomain in cleanedNameAndDomains where targetText.Contains(nameAndDomain.Name) select nameAndDomain.Domain).ToList();
            if (recipientCandidateDomains.Count == 0) return checkList;

            foreach (var recipient in displayNameAndRecipient.All)
            {
                if (recipientCandidateDomains.Any(domains => recipient.Key.EndsWith(domains) || domains.Equals(recipient.Key))) continue;

                if (!recipient.Key.Contains(checkList.SenderDomain))
                {
                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = recipient.Value + " : " + Resources.IsAlertAddressMaybeIrrelevant,
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
                }
            }

            return checkList;
        }

        public CheckList CheckKeywordAndRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, List<KeywordAndRecipients> keywordAndRecipientsList, bool isCheckKeywordAndRecipientsIncludeSubject)
        {
            if (displayNameAndRecipient is null) return checkList;

            var cleanedKeywordAndRecipients = keywordAndRecipientsList.Where(keywordAndRecipient => !string.IsNullOrEmpty(keywordAndRecipient.Keyword) && !string.IsNullOrEmpty(keywordAndRecipient.Recipient)).ToList();
            if (!cleanedKeywordAndRecipients.Any()) return checkList;

            var targetText = checkList.MailBody;
            if (isCheckKeywordAndRecipientsIncludeSubject) { targetText += checkList.Subject; }

            var candidateRecipients = cleanedKeywordAndRecipients.Where(cleanedKeywordAndRecipient => targetText.Contains(cleanedKeywordAndRecipient.Keyword)).ToList();

            foreach (var candidateRecipient in candidateRecipients)
            {
                var isNeedAlert = true;
                foreach (var recipient in displayNameAndRecipient.All)
                {
                    if (recipient.Key.Contains(candidateRecipient.Recipient))
                    {
                        isNeedAlert = false;
                        break;
                    }
                }

                if (isNeedAlert)
                {
                    checkList.Alerts.Add(new Alert
                    {
                        AlertMessage = Resources.KeywordAndRecipientsAlert1 + candidateRecipient.Keyword + Resources.KeywordAndRecipientsAlert2 + candidateRecipient.Recipient + Resources.KeywordAndRecipientsAlert3,
                        IsImportant = true,
                        IsWhite = false,
                        IsChecked = false
                    });
                }
            }

            return checkList;
        }

        public CheckList GetRecipient(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, List<AlertAddress> alertAddressList, List<InternalDomain> internalDomainList)
        {
            if (displayNameAndRecipient is null) return checkList;

            // Helper to process address lists
            void ProcessAddressList(Dictionary<string, string> addresses, List<Address> outputList)
            {
                foreach (var addr in addresses)
                {
                    var isExternal = true;
                    foreach (var _ in internalDomainList.Where(internalDomainSetting => addr.Key.EndsWith(internalDomainSetting.Domain)))
                    {
                        isExternal = false;
                    }

                    if (addr.Value.Contains(Resources.DistributionList) && addr.Key.Contains(Resources.FailedToGetInformation))
                    {
                        isExternal = false;
                    }

                    var isWhite = false;
                    var isSkip = false;
                    if (_whitelist.Count > 0)
                    {
                        if (_whitelist.TryGetValue(addr.Key, out var skip)) { isWhite = true; isSkip = skip; }
                        else
                        {
                            var idx = addr.Key.IndexOf("@", StringComparison.Ordinal);
                            if (idx >= 0)
                            {
                                var domain = addr.Key.Substring(idx);
                                if (_whitelist.TryGetValue(domain, out var skipDom)) { isWhite = true; isSkip = skipDom; }
                                else if (_whitelist.TryGetValue(domain.Substring(1), out var skipDomNoAt)) { isWhite = true; isSkip = skipDomNoAt; }
                            }
                        }
                    }

                    outputList.Add(new Address { MailAddress = addr.Value, IsExternal = isExternal, IsWhite = isWhite, IsChecked = isWhite, IsSkip = isSkip });

                    if (alertAddressList.Count == 0) continue;

                    foreach (var alertAddress in alertAddressList)
                    {
                        if (!addr.Key.EndsWith(alertAddress.TargetAddress)) continue;

                        if (alertAddress.IsCanNotSend)
                        {
                            checkList.IsCanNotSendMail = true;
                            checkList.CanNotSendMailMessage = Resources.SendingForbidAddress + $"[{addr.Value}]";
                            continue;
                        }

                        checkList.Alerts.Add(new Alert
                        {
                            AlertMessage = string.IsNullOrEmpty(alertAddress.Message) ? Resources.IsAlertAddressToAlert + $"[{addr.Value}]" : alertAddress.Message + $"[{addr.Value}]",
                            IsImportant = true,
                            IsWhite = false,
                            IsChecked = false
                        });
                    }
                }
            }

            ProcessAddressList(displayNameAndRecipient.To, checkList.ToAddresses);
            ProcessAddressList(displayNameAndRecipient.Cc, checkList.CcAddresses);
            ProcessAddressList(displayNameAndRecipient.Bcc, checkList.BccAddresses);

            return checkList;
        }

        public CheckList CheckRecipientsAndAttachments(CheckList checkList, bool isAttachmentsProhibited, bool isWarningWhenAttachedRealFile, List<AttachmentProhibitedRecipients> attachmentProhibitedRecipientsList, List<RecipientsAndAttachmentsName> recipientsAndAttachmentsNameList, List<AttachmentAlertRecipients> attachmentAlertRecipientsList)
        {
            if (!checkList.Attachments.Any()) return checkList;

            if (isAttachmentsProhibited)
            {
                checkList.IsCanNotSendMail = true;
                checkList.CanNotSendMailMessage = Resources.AttachmentsProhibitedMessage;
                return checkList;
            }

            if (attachmentProhibitedRecipientsList.Count > 0)
            {
                var prohibitedRecipients = "";
                var isProhibited = false;

                foreach (var prohibitedRecipient in attachmentProhibitedRecipientsList)
                {
                    foreach (var to in checkList.ToAddresses.Where(to => to.MailAddress.Contains(prohibitedRecipient.Recipient)))
                    {
                        checkList.IsCanNotSendMail = true;
                        isProhibited = true;
                        prohibitedRecipients += " " + to.MailAddress;
                    }

                    foreach (var cc in checkList.CcAddresses.Where(cc => cc.MailAddress.Contains(prohibitedRecipient.Recipient)))
                    {
                        checkList.IsCanNotSendMail = true;
                        isProhibited = true;
                        prohibitedRecipients += " " + cc.MailAddress;
                    }

                    foreach (var bcc in checkList.BccAddresses.Where(bcc => bcc.MailAddress.Contains(prohibitedRecipient.Recipient)))
                    {
                        checkList.IsCanNotSendMail = true;
                        isProhibited = true;
                        prohibitedRecipients += " " + bcc.MailAddress;
                    }
                }

                if (isProhibited)
                {
                    checkList.CanNotSendMailMessage = Resources.AttachmentProhibitedRecipientsMessage + "：" + prohibitedRecipients;
                    return checkList;
                }
            }

            if (attachmentAlertRecipientsList.Count > 0)
            {
                foreach (var attachmentAlertRecipient in attachmentAlertRecipientsList)
                {
                    foreach (var to in checkList.ToAddresses.Where(to => to.MailAddress.Contains(attachmentAlertRecipient.Recipient)))
                    {
                        AddAlerts(checkList, string.IsNullOrEmpty(attachmentAlertRecipient.Message) ? Resources.AttachmentAlertRecipientsMessage + $"[{to.MailAddress}]" : attachmentAlertRecipient.Message + $"[{to.MailAddress}]", true, false, false);
                    }
                    foreach (var cc in checkList.CcAddresses.Where(cc => cc.MailAddress.Contains(attachmentAlertRecipient.Recipient)))
                    {
                        AddAlerts(checkList, string.IsNullOrEmpty(attachmentAlertRecipient.Message) ? Resources.AttachmentAlertRecipientsMessage + $"[{cc.MailAddress}]" : attachmentAlertRecipient.Message + $"[{cc.MailAddress}]", true, false, false);
                    }
                    foreach (var bcc in checkList.BccAddresses.Where(bcc => bcc.MailAddress.Contains(attachmentAlertRecipient.Recipient)))
                    {
                        AddAlerts(checkList, string.IsNullOrEmpty(attachmentAlertRecipient.Message) ? Resources.AttachmentAlertRecipientsMessage + $"[{bcc.MailAddress}]" : attachmentAlertRecipient.Message + $"[{bcc.MailAddress}]", true, false, false);
                    }
                }
            }

            if (recipientsAndAttachmentsNameList.Count > 0)
            {
                foreach (var recipientsAndAttachmentsName in recipientsAndAttachmentsNameList)
                {
                    foreach (var attachment in checkList.Attachments.Where(attachment => attachment.FileName.Contains(recipientsAndAttachmentsName.AttachmentsName)))
                    {
                        foreach (var to in checkList.ToAddresses.Where(to => to.IsExternal))
                        {
                            if (!to.MailAddress.Contains(recipientsAndAttachmentsName.Recipient))
                            {
                                AddAlerts(checkList, Resources.RecipientsAndAttachmentsNameMessage + "：" + to.MailAddress + " / " + attachment.FileName, true, true, false);
                            }
                        }

                        foreach (var cc in checkList.CcAddresses.Where(cc => cc.IsExternal))
                        {
                            if (!cc.MailAddress.Contains(recipientsAndAttachmentsName.Recipient))
                            {
                                AddAlerts(checkList, Resources.RecipientsAndAttachmentsNameMessage + "：" + cc.MailAddress + " / " + attachment.FileName, true, true, false);
                            }
                        }

                        foreach (var bcc in checkList.BccAddresses.Where(bcc => bcc.IsExternal))
                        {
                            if (!bcc.MailAddress.Contains(recipientsAndAttachmentsName.Recipient))
                            {
                                AddAlerts(checkList, Resources.RecipientsAndAttachmentsNameMessage + "：" + bcc.MailAddress + " / " + attachment.FileName, true, true, false);
                            }
                        }
                    }
                }
            }

            if (isWarningWhenAttachedRealFile)
            {
                AddAlerts(checkList, Resources.RecommendationOfAttachFileAsLink, false, true, false);
            }

            return checkList;
        }

        public CheckList AddAlertOrProhibitsSendingMailIfIfRecipientsIsNotRegistered(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, List<string> contactsList, List<InternalDomain> internalDomainList, bool isWarningIfRecipientsIsNotRegistered, bool isProhibitsSendingMailIfRecipientsIsNotRegistered)
        {
            if (!(isWarningIfRecipientsIsNotRegistered || isProhibitsSendingMailIfRecipientsIsNotRegistered)) return checkList;

            var selectedContactsList = contactsList.SelectMany(contact => displayNameAndRecipient.MailRecipientsIndex.Where(mailItemsRecipient => contact == mailItemsRecipient.MailAddress || contact == mailItemsRecipient.MailItemsRecipient)).Select(x => x.MailAddress).ToList();

            void CheckContacts(Dictionary<string, string> recipients)
            {
                foreach (var recipient in recipients.Where(to => !selectedContactsList.Contains(to.Key)))
                {
                    if (internalDomainList.Any(internalDomain => recipient.Key.EndsWith(internalDomain.Domain, StringComparison.OrdinalIgnoreCase))) continue;

                    if (isProhibitsSendingMailIfRecipientsIsNotRegistered)
                    {
                        checkList.IsCanNotSendMail = true;
                        checkList.CanNotSendMailMessage = Resources.ProhibitsSendingMailIfRecipientsIsNotRegisteredMessage + $" [{recipient.Value}]";
                        return; // Note: In original code, it returns explicitly. Here we are in a lambda/local func. We need to handle this.
                        // Since we can't easily return from parent, we loop and check checkList state outside.
                    }

                    AddAlerts(checkList, Resources.WarningIfRecipientsIsNotRegisteredMessage + $" [{recipient.Value}]", true, false, false);
                }
            }

            // Re-implementing loops to allow early return
            foreach (var to in displayNameAndRecipient.To.Where(to => !selectedContactsList.Contains(to.Key)))
            {
                if (internalDomainList.Any(internalDomain => to.Key.EndsWith(internalDomain.Domain, StringComparison.OrdinalIgnoreCase))) continue;

                if (isProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.ProhibitsSendingMailIfRecipientsIsNotRegisteredMessage + $" [{to.Value}]";
                    return checkList;
                }
                AddAlerts(checkList, Resources.WarningIfRecipientsIsNotRegisteredMessage + $" [{to.Value}]", true, false, false);
            }

            foreach (var cc in displayNameAndRecipient.Cc.Where(cc => !selectedContactsList.Contains(cc.Key)))
            {
                if (internalDomainList.Any(internalDomain => cc.Key.EndsWith(internalDomain.Domain, StringComparison.OrdinalIgnoreCase))) continue;

                if (isProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.ProhibitsSendingMailIfRecipientsIsNotRegisteredMessage + $" [{cc.Value}]";
                    return checkList;
                }
                AddAlerts(checkList, Resources.WarningIfRecipientsIsNotRegisteredMessage + $" [{cc.Value}]", true, false, false);
            }

            foreach (var bcc in displayNameAndRecipient.Bcc.Where(bcc => !selectedContactsList.Contains(bcc.Key)))
            {
                if (internalDomainList.Any(internalDomain => bcc.Key.EndsWith(internalDomain.Domain, StringComparison.OrdinalIgnoreCase))) continue;

                if (isProhibitsSendingMailIfRecipientsIsNotRegistered)
                {
                    checkList.IsCanNotSendMail = true;
                    checkList.CanNotSendMailMessage = Resources.ProhibitsSendingMailIfRecipientsIsNotRegisteredMessage + $" [{bcc.Value}]";
                    return checkList;
                }
                AddAlerts(checkList, Resources.WarningIfRecipientsIsNotRegisteredMessage + $" [{bcc.Value}]", true, false, false);
            }

            return checkList;
        }

        public CheckList AutoCheckRegisteredItemsInContacts(CheckList checkList, DisplayNameAndRecipient displayNameAndRecipient, List<string> contactsList, bool isAutoCheckRegisteredInContacts)
        {
            if (!isAutoCheckRegisteredInContacts) return checkList;

            foreach (var mailItemsRecipient in contactsList.SelectMany(contact => displayNameAndRecipient.MailRecipientsIndex.Where(mailItemsRecipient => contact == mailItemsRecipient.MailAddress || contact == mailItemsRecipient.MailItemsRecipient)))
            {
                foreach (var toAddress in checkList.ToAddresses.Where(toAddress => toAddress.MailAddress.Contains(mailItemsRecipient.MailAddress)))
                {
                    toAddress.IsChecked = true;
                }

                foreach (var ccAddress in checkList.CcAddresses.Where(ccAddress => ccAddress.MailAddress.Contains(mailItemsRecipient.MailAddress)))
                {
                    ccAddress.IsChecked = true;
                }

                foreach (var bccAddress in checkList.BccAddresses.Where(bccAddress => bccAddress.MailAddress.Contains(mailItemsRecipient.MailAddress)))
                {
                    bccAddress.IsChecked = true;
                }
            }

            return checkList;
        }

        // Helper method to add alerts
        private void AddAlerts(CheckList checkList, string alertMessage, bool isImportant, bool isWhite, bool isChecked)
        {
            checkList.Alerts.Add(new Alert
            {
                AlertMessage = alertMessage,
                IsImportant = isImportant,
                IsWhite = isWhite,
                IsChecked = isChecked
            });
        }

        private static void ChangeToBcc<T>(T item, IReadOnlyCollection<MailItemsRecipientAndMailAddress> mailItemsRecipientAndMailAddress, string senderMailAddress, bool isNeedsAddToSender)
        {
            if ((dynamic)item is null) return;

            var targetMailAddressAndRecipient = new Dictionary<string, string>();

            foreach (Outlook.Recipient recipient in ((dynamic)item).Recipients)
            {
                foreach (var target in mailItemsRecipientAndMailAddress)
                {
                    if (recipient.Address == target.MailItemsRecipient) targetMailAddressAndRecipient[target.MailAddress] = target.MailItemsRecipient;
                }
            }

            var targetCount = targetMailAddressAndRecipient.Count;
            while (targetCount > 0)
            {
                foreach (var target in targetMailAddressAndRecipient)
                {
                    foreach (Outlook.Recipient recipient in ((dynamic)item).Recipients)
                    {
                        if (recipient.Address != target.Value) continue;
                        ((dynamic)item).Recipients.Remove(recipient.Index);
                        targetCount--;
                    }
                }
            }

            foreach (var addTarget in targetMailAddressAndRecipient.Select(mailAddress => ((dynamic)item).Recipients.Add(mailAddress.Key)))
            {
                addTarget.Type = (int)Outlook.OlMailRecipientType.olBCC;
            }

            if (isNeedsAddToSender)
            {
                var senderRecipient = ((dynamic)item).Recipients.Add(senderMailAddress);
                senderRecipient.Type = (int)Outlook.OlMailRecipientType.olTo;
            }

            _ = ((dynamic)item).Recipients.ResolveAll();
        }
    }
}
