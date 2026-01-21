using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;

namespace OutlookOkan.Handlers
{
    public sealed class ZipFileHandler
    {
        internal readonly List<string> IncludeExtensions = new List<string>();
        internal bool IsContainsShortcut;

        /// <summary>
        /// Xác định xem có phải là tệp ZIP được mã hóa (ZIP có mật khẩu) hay không.
        /// </summary>
        /// <param name="filePath">Đường dẫn đầy đủ của tệp cần kiểm tra</param>
        /// <returns>Có phải là ZIP được mã hóa hay không</returns>
        internal bool CheckZipIsEncryptedAndGetIncludeExtensions(string filePath)
        {
            // Nếu đính kèm dưới dạng liên kết, tệp thực tế có thể không tồn tại.
            if (!File.Exists(filePath)) return false;

            var isEncrypted = false;

            try
            {
                using (var zipFile = new ZipFile(filePath))
                {
                    try
                    {
                        foreach (ZipEntry entry in zipFile)
                        {
                            if (!entry.IsFile) continue;

                            if (entry.IsCrypted)
                            {
                                isEncrypted = true;
                            }

                            var extension = Path.GetExtension(entry.Name);
                            IncludeExtensions.Add(extension.ToLower());

                            if (!isEncrypted)
                            {
                                if (IsShortcutFile(zipFile, entry))
                                {
                                    IsContainsShortcut = true;
                                }
                            }
                        }
                    }
                    catch (NotSupportedException)
                    {
                        isEncrypted = true;
                    }
                    catch (Exception)
                    {
                        //Do nothing
                    }
                }
            }
            catch (Exception)
            {
                isEncrypted = false;
            }

            return isEncrypted;
        }

        private bool IsShortcutFile(ZipFile zipFile, ZipEntry entry)
        {
            try
            {
                using (var stream = zipFile.GetInputStream(entry))
                {
                    var buffer = new byte[20];
                    var bytesRead = stream.Read(buffer, 0, buffer.Length);

                    if (bytesRead >= 20)
                    {
                        return buffer[0] == 0x4C && buffer[1] == 0x00 && buffer[2] == 0x00 && buffer[3] == 0x00 &&
                               buffer[4] == 0x01 && buffer[5] == 0x14 && buffer[6] == 0x02 && buffer[7] == 0x00 &&
                               buffer[8] == 0x00 && buffer[9] == 0x00 && buffer[10] == 0x00 && buffer[11] == 0x00 &&
                               buffer[12] == 0xC0 && buffer[13] == 0x00 && buffer[14] == 0x00 && buffer[15] == 0x00 &&
                               buffer[16] == 0x00 && buffer[17] == 0x00 && buffer[18] == 0x00 && buffer[19] == 0x46;
                    }

                    return false;
                }
            }
            catch (Exception)
            {
                return false;
            }
        }
    }
}