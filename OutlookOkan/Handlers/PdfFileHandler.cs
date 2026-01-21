using PdfSharp.Pdf.IO;
using System;
using System.IO;

namespace OutlookOkan.Handlers
{
    internal static class PdfFileHandler
    {
        internal static bool CheckPdfIsEncrypted(string filePath)
        {
            // Nếu đính kèm dưới dạng liên kết, tệp thực tế có thể không tồn tại.
            if (!File.Exists(filePath)) return false;

            try
            {
                PdfReader.Open(filePath, PdfDocumentOpenMode.ReadOnly).Dispose();
            }
            catch (PdfReaderException)
            {
                return true;
            }
            catch (Exception)
            {
                return false;
            }

            return false;
        }
    }
}