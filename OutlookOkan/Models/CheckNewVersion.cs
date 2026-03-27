using System.Reflection;

namespace OutlookOkan.Models
{
    /// <summary>
    /// Kiểm tra phiên bản mới.
    /// [SECURITY] Đã vô hiệu hóa kết nối internet - add-in hoạt động hoàn toàn offline.
    /// </summary>
    internal sealed class CheckNewVersion
    {
        /// <summary>
        /// Luôn trả về false - tính năng kiểm tra phiên bản mới đã bị vô hiệu hóa
        /// để đảm bảo add-in không thực hiện bất kỳ kết nối internet nào.
        /// </summary>
        /// <returns>Luôn trả về false</returns>
        internal bool IsCanDownloadNewVersion()
        {
            return false;
        }

        /// <summary>
        /// Lấy phiên bản hiện tại của add-in.
        /// </summary>
        /// <returns>Phiên bản hiện tại</returns>
        internal int GetCurrentVersion()
        {
            var assemblyName = Assembly.GetExecutingAssembly().GetName();
            return int.Parse(assemblyName.Version.Major.ToString() + assemblyName.Version.Minor.ToString() + assemblyName.Version.Build.ToString() + assemblyName.Version.Revision.ToString());
        }
    }
}