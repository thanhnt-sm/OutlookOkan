using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OutlookOkan.Handlers
{
    /// <summary>
    /// Thực hiện phân tích tiêu đề email
    /// </summary>
    internal static class MailHeaderHandler
    {
        /// <summary>
        /// Phân tích tiêu đề email và trả về kết quả xác minh như SPF, DKIM, DMARC
        /// </summary>
        /// <param name="emailHeader">Tiêu đề email</param>
        /// <returns>Kết quả phân tích</returns>
        internal static Dictionary<string, string> ValidateEmailHeader(string emailHeader)
        {
            var results = new Dictionary<string, string>
            {
                ["From Domain"] = "NONE",
                ["ReturnPath Domain"] = "NONE",
                ["SPF"] = "NONE",
                ["SPF IP"] = "NONE",
                ["SPF Alignment"] = "NONE",
                ["DKIM"] = "NONE",
                ["DKIM Domain"] = "NONE",
                ["DKIM Alignment"] = "NONE",
                ["DMARC"] = "NONE",
                ["Internal"] = "FALSE"
            };

            if (string.IsNullOrEmpty(emailHeader))
            {
                return null;
            }

            if (IsInternalMail(emailHeader))
            {
                results["Internal"] = "TRUE";
            }

            var fromDomain = string.Empty;
            var fromRegex = new Regex(@"^From:\s*.*(?:\r?\n\s+.*)*", RegexOptions.IgnoreCase | RegexOptions.Multiline);
            var fromMatch = fromRegex.Match(emailHeader);
            if (fromMatch.Success)
            {
                var fromHeader = fromMatch.Value;
                var domainRegex = new Regex(@"<.*?@(?<domain>[^\s>]+)>", RegexOptions.IgnoreCase);
                var domainMatch = domainRegex.Match(fromHeader);

                if (!domainMatch.Success)
                {
                    var alternativeDomainRegex = new Regex(@"[^<\s]+@(?<domain>[^\s>]+)", RegexOptions.IgnoreCase);
                    domainMatch = alternativeDomainRegex.Match(fromHeader);
                }

                fromDomain = domainMatch.Success ? domainMatch.Groups["domain"].Value : string.Empty;
                results["From Domain"] = fromDomain;
            }

            // Xác minh SPF
            var spfRegex = new Regex(@"Received-SPF:\s*(?<result>pass|fail|softfail|neutral|temperror|permerror|none).*\b(does\s+not\s+)?designate[s]?\s+(?<ip>[^ ]+)\s+as", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var spfMatch = spfRegex.Match(emailHeader);
            if (spfMatch.Success)
            {
                results["SPF"] = spfMatch.Groups["result"].Value.ToUpper();
                results["SPF IP"] = spfMatch.Groups["ip"].Value;
            }

            // Xác minh SPF Alignment
            var returnPathRegex = new Regex(@"Return-Path:\s*.*@(?<domain>[^\s>]+)");
            var returnPathMatch = returnPathRegex.Match(emailHeader);
            if (returnPathMatch.Success && fromDomain != string.Empty)
            {
                var returnPathDomain = returnPathMatch.Groups["domain"].Value;
                results["ReturnPath Domain"] = returnPathDomain;
                results["SPF Alignment"] = returnPathDomain.Equals(fromDomain, StringComparison.OrdinalIgnoreCase) || returnPathDomain.ToLower().Contains(fromDomain.ToLower()) || fromDomain.ToLower().Contains(returnPathDomain.ToLower()) ? "PASS" : "FAIL";
            }

            // Xác minh DKIM
            var dkimRegex = new Regex(@"Authentication-Results:.*?dkim=(?<result>pass|policy|fail|softfail|hardfail|neutral|temperror|permerror|none).*?header.d=(?<domain>[^(;| )]+)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var dkimMatch = dkimRegex.Match(emailHeader);
            if (dkimMatch.Success)
            {
                results["DKIM"] = dkimMatch.Groups["result"].Value.ToUpper();
            }

            // Xác minh DKIM Alignment
            var dkimSignatureRegex = new Regex(@"DKIM-Signature:.*?d=(?<domain>[^(;| )]+)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var dkimMatches = dkimSignatureRegex.Matches(emailHeader);
            var dkimAlignmentPass = false;
            var dkimDomains = new List<string>();

            foreach (Match match in dkimMatches)
            {
                var dkimDomain = match.Groups["domain"].Value;
                if (string.IsNullOrEmpty(dkimDomain)) continue;

                dkimDomains.Add(dkimDomain);
                if (dkimDomain.Equals(fromDomain, StringComparison.OrdinalIgnoreCase) ||
                    dkimDomain.ToLower().Contains(fromDomain.ToLower()) ||
                    fromDomain.ToLower().Contains(dkimDomain.ToLower()))
                {
                    dkimAlignmentPass = true;
                }
            }
            results["DKIM Domain"] = string.Join(", ", dkimDomains);
            results["DKIM Alignment"] = dkimAlignmentPass ? "PASS" : "FAIL";

            // Xác minh DMARC
            var dmarcRegex = new Regex(@"Authentication-Results:.*?dmarc=(?<result>pass|bestguesspass|softfail|fail|none)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
            var dmarcMatch = dmarcRegex.Match(emailHeader);
            if (dmarcMatch.Success)
            {
                results["DMARC"] = dmarcMatch.Groups["result"].Value.ToUpper();
            }

            return results;
        }

        /// <summary>
        /// Tự xác định kết quả xác minh DMARC
        /// </summary>
        /// <param name="spfResult"></param>
        /// <param name="spfAlignmentResult"></param>
        /// <param name="dkimResult"></param>
        /// <param name="dkimAlignmentResult"></param>
        /// <returns>Kết quả xác minh DMARC</returns>
        public static string DetermineDmarcResult(string spfResult, string spfAlignmentResult, string dkimResult, string dkimAlignmentResult)
        {
            if (string.IsNullOrEmpty(spfResult) || string.IsNullOrEmpty(spfAlignmentResult) || string.IsNullOrEmpty(dkimResult) || string.IsNullOrEmpty(dkimAlignmentResult))
            {
                return "FAIL";
            }

            // Coi NONE là FAIL
            spfResult = spfResult.ToUpper() == "NONE" ? "FAIL" : spfResult.ToUpper();
            spfAlignmentResult = spfAlignmentResult.ToUpper() == "NONE" ? "FAIL" : spfAlignmentResult.ToUpper();
            dkimResult = dkimResult.ToUpper() == "NONE" ? "FAIL" : dkimResult.ToUpper();
            dkimAlignmentResult = dkimAlignmentResult.ToUpper() == "NONE" ? "FAIL" : dkimAlignmentResult.ToUpper();

            var key = $"{spfResult}_{spfAlignmentResult}_{dkimResult}_{dkimAlignmentResult}";

            //Xác thực SPF_SPF Alignment_Xác thực DKIM_DKIM Alignment
            var dmarcResults = new Dictionary<string, string>
            {
                { "PASS_PASS_PASS_PASS", "PASS" }, // Cả xác thực và alignment đều thành công
                { "PASS_PASS_PASS_FAIL", "PASS" }, // Xác thực SPF và alignment thành công, xác thực DKIM thành công
                { "PASS_PASS_FAIL_PASS", "PASS" }, // Xác thực SPF và alignment thành công, alignment DKIM thành công
                { "PASS_PASS_FAIL_FAIL", "PASS" }, // Xác thực SPF và alignment thành công
                { "PASS_FAIL_PASS_PASS", "PASS" }, // Xác thực SPF thành công, xác thực DKIM và alignment thành công
                { "FAIL_PASS_PASS_PASS", "PASS" }, // Alignment SPF thành công, xác thực DKIM và alignment thành công
                { "FAIL_FAIL_PASS_PASS", "PASS" }, // Xác thực DKIM và alignment thành công
                
                { "PASS_FAIL_PASS_FAIL", "FAIL" }, // Xác thực SPF thành công, xác thực DKIM thành công
                { "PASS_FAIL_FAIL_PASS", "FAIL" }, // Xác thực SPF thành công, alignment DKIM thành công
                { "PASS_FAIL_FAIL_FAIL", "FAIL" }, // Xác thực SPF thành công
                { "FAIL_PASS_PASS_FAIL", "FAIL" }, // Alignment SPF thành công, xác thực DKIM thành công
                { "FAIL_PASS_FAIL_PASS", "FAIL" }, // Alignment SPF thành công, alignment DKIM thành công
                { "FAIL_PASS_FAIL_FAIL", "FAIL" }, // Alignment SPF thành công
                { "FAIL_FAIL_PASS_FAIL", "FAIL" }, // Xác thực DKIM thành công
                { "FAIL_FAIL_FAIL_PASS", "FAIL" }, // Alignment DKIM thành công
                { "FAIL_FAIL_FAIL_FAIL", "FAIL" }  // Tất cả đều thất bại
            };
            return dmarcResults.TryGetValue(key, out var result) ? result : "FAIL";
        }

        /// <summary>
        /// Xác định xem có phải là email nội bộ hay không
        /// </summary>
        /// <param name="emailHeader">Tiêu đề email</param>
        /// <returns>Kết quả xác định</returns>
        internal static bool IsInternalMail(string emailHeader)
        {
            // Lấy tất cả các tiêu đề Received
            var receivedRegex = new Regex(@"^Received:.*", RegexOptions.Multiline);
            var matches = receivedRegex.Matches(emailHeader);

            var receivedHeaders = (from Match match in matches select match.Value).ToList();

            // Nếu số lượng tiêu đề nhận được nhiều, xác định là email bên ngoài
            if (receivedHeaders.Count > 3)
            {
                return false;
            }

            // Nếu có nhiều tiêu đề nhận, kiểm tra xem tên miền liên tiếp có khớp nhau không
            var domainRegex = new Regex(@"from\s([^\s]+)", RegexOptions.IgnoreCase);
            string previousDomain = null;

            foreach (var currentDomain in from header in receivedHeaders select domainRegex.Match(header) into domainMatch where domainMatch.Success select ExtractMainDomain(domainMatch.Groups[1].Value))
            {
                if (previousDomain != null && previousDomain.Equals(currentDomain, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                previousDomain = currentDomain;
            }

            return false;
        }

        private static string ExtractMainDomain(string domain)
        {
            var parts = domain.Split('.');
            var length = parts.Length;

            return length > 2 ? string.Join(".", parts.Skip(length - 3)) : domain;
        }
    }
}