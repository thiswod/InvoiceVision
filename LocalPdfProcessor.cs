using System;
using System.IO;
using System.Text;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace InvoiceVision
{
    /// <summary>
    /// 本地PDF处理器，用于从PDF文件中提取文本内容
    /// </summary>
    public class LocalPdfProcessor
    {
        /// <summary>
        /// 从PDF文件中提取所有文本内容
        /// </summary>
        /// <param name="pdfPath">PDF文件路径</param>
        /// <returns>提取的文本内容</returns>
        public string ExtractText(string pdfPath)
        {
            if (!File.Exists(pdfPath))
            {
                throw new FileNotFoundException($"PDF文件不存在: {pdfPath}");
            }

            StringBuilder textBuilder = new StringBuilder();

            try
            {
                using (PdfDocument document = PdfDocument.Open(pdfPath))
                {
                    foreach (Page page in document.GetPages())
                    {
                        textBuilder.AppendLine(page.Text);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"提取PDF文本时出错: {ex.Message}", ex);
            }

            return textBuilder.ToString();
        }

        /// <summary>
        /// 从PDF文件中提取第一页的文本内容（用于发票通常只有一页）
        /// </summary>
        /// <param name="pdfPath">PDF文件路径</param>
        /// <returns>第一页的文本内容</returns>
        public string ExtractFirstPageText(string pdfPath)
        {
            if (!File.Exists(pdfPath))
            {
                throw new FileNotFoundException($"PDF文件不存在: {pdfPath}");
            }

            try
            {
                using (PdfDocument document = PdfDocument.Open(pdfPath))
                {
                    if (document.NumberOfPages > 0)
                    {
                        Page firstPage = document.GetPage(1);
                        return firstPage.Text;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"提取PDF文本时出错: {ex.Message}", ex);
            }

            return string.Empty;
        }
    }
}

