using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using WodToolkit.Json;

namespace InvoiceVision
{
    public partial class Form1 : Form
    {
        private readonly LocalPdfInvoiceExtractor localPdfInvoiceExtractor = new();
        private readonly List<string> extractedTempDirectories = new();
        private readonly List<InvoiceData> invoiceResults = new();
        private BaiDu? baiDu;

        public Form1()
        {
            InitializeComponent();
            LoadConfiguration();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void LoadConfiguration()
        {
            try
            {
                var builder = new ConfigurationBuilder()
                    .SetBasePath(AppContext.BaseDirectory)
                    .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);

                IConfiguration configuration = builder.Build();
                string apiKey = configuration["BaiduOCR:ApiKey"] ?? string.Empty;
                string secretKey = configuration["BaiduOCR:SecretKey"] ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(apiKey) && !string.IsNullOrWhiteSpace(secretKey))
                {
                    baiDu = new BaiDu(apiKey, secretKey);
                }
            }
            catch
            {
                baiDu = null;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            CleanupExtractedTempDirectories();
            base.OnFormClosed(e);
        }

        private void BtnSelectImages_Click(object sender, EventArgs e)
        {
            using OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "发票文件和压缩包|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.pdf;*.zip|图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif|PDF文件|*.pdf|压缩包|*.zip|所有文件|*.*";
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "选择发票图片、PDF或压缩包";

            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            CleanupExtractedTempDirectories();
            listBoxImages.Items.Clear();

            var collectedFiles = new List<string>();
            foreach (string selectedPath in openFileDialog.FileNames)
            {
                collectedFiles.AddRange(ExpandSelectedPath(selectedPath));
            }

            foreach (string filePath in collectedFiles.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                listBoxImages.Items.Add(filePath);
            }

            btnStart.Enabled = listBoxImages.Items.Count > 0;
            labelStatus.Text = $"已载入 {listBoxImages.Items.Count} 个文件";

            if (listBoxImages.Items.Count == 0)
            {
                MessageBox.Show("没有找到可识别的 PDF 或图片文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (listBoxImages.Items.Count == 0)
            {
                MessageBox.Show("请先选择图片、PDF或压缩包。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool containsImageFiles = listBoxImages.Items
                .Cast<string>()
                .Any(path => !IsPdfFile(path));

            if (containsImageFiles && baiDu == null)
            {
                MessageBox.Show(
                    "当前选择中包含图片文件，但未找到可用的百度 OCR 配置。\n\nPDF 已改为本地解析；如果还要识别图片，请在 appsettings.json 中配置百度 OCR。",
                    "缺少图片识别配置",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
                return;
            }

            btnStart.Enabled = false;
            btnSelectImages.Enabled = false;
            btnExport.Enabled = false;
            progressBar.Visible = true;
            progressBar.Maximum = listBoxImages.Items.Count;
            progressBar.Value = 0;
            superListView.Items.Clear();
            invoiceResults.Clear();

            try
            {
                ProcessFiles();
                labelStatus.Text = $"处理完成，共识别 {invoiceResults.Count} 条记录";
                btnExport.Enabled = invoiceResults.Count > 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理过程中出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelStatus.Text = "处理失败";
            }
            finally
            {
                btnStart.Enabled = true;
                btnSelectImages.Enabled = true;
                progressBar.Visible = false;
            }
        }

        private void ListBoxImages_DoubleClick(object sender, EventArgs e)
        {
            if (listBoxImages.SelectedItem is not string filePath || string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            if (!File.Exists(filePath))
            {
                MessageBox.Show($"文件不存在：{filePath}", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开文件失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ProcessFiles()
        {
            const int minDelayMs = 500;
            int processedCount = 0;

            foreach (string filePath in listBoxImages.Items.Cast<string>())
            {
                try
                {
                    if (!IsPdfFile(filePath) && processedCount > 0)
                    {
                        Thread.Sleep(minDelayMs);
                    }

                    ProcessSingleFile(filePath);
                    processedCount++;
                    progressBar.Value = processedCount;
                    labelStatus.Text = $"正在处理... ({processedCount}/{listBoxImages.Items.Count})";
                    Application.DoEvents();
                }
                catch (Exception ex)
                {
                    processedCount++;
                    progressBar.Value = processedCount;
                    labelStatus.Text = $"处理 {Path.GetFileName(filePath)} 出错：{ex.Message}";
                    Application.DoEvents();
                }
            }
        }

        private void ProcessSingleFile(string filePath)
        {
            InvoiceData invoiceData = IsPdfFile(filePath)
                ? localPdfInvoiceExtractor.Extract(filePath)
                : ProcessImageWithOcr(filePath);

            invoiceResults.Add(invoiceData);
            AddInvoiceToListView(invoiceData);
        }

        private InvoiceData ProcessImageWithOcr(string imagePath)
        {
            if (baiDu == null)
            {
                throw new InvalidOperationException("未配置百度 OCR，无法处理图片文件。");
            }

            byte[] fileBytes = File.ReadAllBytes(imagePath);
            string base64Data = Convert.ToBase64String(fileBytes);
            string fileType = GetFileType(imagePath);
            string resultJson = baiDu.vat_invoice(base64Data, fileType);
            dynamic result = EasyJson.ParseJsonToDynamic(resultJson);

            if (result.words_result == null)
            {
                string errorMsg = "识别结果为空";

                try
                {
                    if (result.error_code != null)
                    {
                        errorMsg = $"API 返回错误：{result.error_msg ?? "未知错误"} (错误码: {result.error_code})";
                    }
                }
                catch
                {
                }

                throw new Exception(errorMsg);
            }

            return ParseInvoiceData(result.words_result, imagePath);
        }

        private InvoiceData ParseInvoiceData(dynamic wordsResult, string imagePath)
        {
            string invoiceNum = GetStringValue(wordsResult.InvoiceNum);
            string invoiceCode = GetStringValue(wordsResult.InvoiceCode);

            if (string.IsNullOrEmpty(invoiceCode) && !string.IsNullOrEmpty(invoiceNum))
            {
                invoiceCode = invoiceNum;
            }

            return new InvoiceData
            {
                ImagePath = imagePath,
                InvoiceNum = invoiceNum,
                InvoiceCode = invoiceCode,
                InvoiceDate = GetStringValue(wordsResult.InvoiceDate),
                PurchaserName = GetStringValue(wordsResult.PurchaserName),
                PurchaserRegisterNum = GetStringValue(wordsResult.PurchaserRegisterNum),
                SellerName = GetStringValue(wordsResult.SellerName),
                SellerRegisterNum = GetStringValue(wordsResult.SellerRegisterNum),
                TotalAmount = GetStringValue(wordsResult.TotalAmount),
                TotalTax = GetStringValue(wordsResult.TotalTax),
                AmountInFiguers = GetStringValue(wordsResult.AmountInFiguers),
                InvoiceType = GetStringValue(wordsResult.InvoiceType),
                RawData = wordsResult
            };
        }

        private static string GetFileType(string imagePath)
        {
            return Path.GetExtension(imagePath).ToLowerInvariant() switch
            {
                ".jpg" => "jpeg",
                ".jpeg" => "jpeg",
                ".png" => "png",
                ".bmp" => "bmp",
                ".gif" => "gif",
                ".pdf" => "pdf",
                _ => "png"
            };
        }

        private static bool IsPdfFile(string path)
        {
            return string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsZipFile(string path)
        {
            return string.Equals(Path.GetExtension(path), ".zip", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsSupportedInvoiceFile(string path)
        {
            string extension = Path.GetExtension(path);
            return extension.Equals(".jpg", StringComparison.OrdinalIgnoreCase)
                || extension.Equals(".jpeg", StringComparison.OrdinalIgnoreCase)
                || extension.Equals(".png", StringComparison.OrdinalIgnoreCase)
                || extension.Equals(".bmp", StringComparison.OrdinalIgnoreCase)
                || extension.Equals(".gif", StringComparison.OrdinalIgnoreCase)
                || extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase);
        }

        private IEnumerable<string> ExpandSelectedPath(string path)
        {
            if (IsZipFile(path))
            {
                return ExtractSupportedFilesFromZip(path);
            }

            return IsSupportedInvoiceFile(path) ? new[] { path } : Array.Empty<string>();
        }

        private IEnumerable<string> ExtractSupportedFilesFromZip(string zipPath)
        {
            if (!File.Exists(zipPath))
            {
                return Array.Empty<string>();
            }

            string extractRoot = Path.Combine(
                Path.GetTempPath(),
                "InvoiceVision",
                $"{Path.GetFileNameWithoutExtension(zipPath)}_{Guid.NewGuid():N}");

            Directory.CreateDirectory(extractRoot);
            extractedTempDirectories.Add(extractRoot);

            var extractedFiles = new List<string>();

            using ZipArchive archive = ZipFile.OpenRead(zipPath);
            foreach (ZipArchiveEntry entry in archive.Entries)
            {
                if (string.IsNullOrEmpty(entry.Name))
                {
                    continue;
                }

                if (!IsSupportedInvoiceFile(entry.FullName))
                {
                    continue;
                }

                string destinationPath = Path.GetFullPath(Path.Combine(extractRoot, entry.FullName));
                if (!destinationPath.StartsWith(extractRoot, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string? directory = Path.GetDirectoryName(destinationPath);
                if (!string.IsNullOrWhiteSpace(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                entry.ExtractToFile(destinationPath, overwrite: true);
                extractedFiles.Add(destinationPath);
            }

            return extractedFiles;
        }

        private void CleanupExtractedTempDirectories()
        {
            foreach (string directory in extractedTempDirectories)
            {
                try
                {
                    if (Directory.Exists(directory))
                    {
                        Directory.Delete(directory, recursive: true);
                    }
                }
                catch
                {
                }
            }

            extractedTempDirectories.Clear();
        }

        private static string GetStringValue(dynamic value)
        {
            return value == null ? string.Empty : value.ToString();
        }

        private void AddInvoiceToListView(InvoiceData invoice)
        {
            ListViewItem item = new ListViewItem(invoice.InvoiceNum);
            item.SubItems.Add(invoice.InvoiceCode);
            item.SubItems.Add(invoice.InvoiceDate);
            item.SubItems.Add(invoice.PurchaserName);
            item.SubItems.Add(invoice.SellerName);
            item.SubItems.Add(invoice.TotalAmount);
            item.SubItems.Add(invoice.TotalTax);
            item.SubItems.Add(invoice.AmountInFiguers);
            item.SubItems.Add(invoice.ImagePath);
            item.Tag = invoice;
            superListView.Items.Add(item);
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (invoiceResults.Count == 0)
            {
                MessageBox.Show("没有可导出的数据。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel文件|*.xlsx|所有文件|*.*";
            saveFileDialog.FileName = $"发票识别结果_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            saveFileDialog.Title = "保存Excel文件";

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            try
            {
                ExportToExcel(saveFileDialog.FileName);
                MessageBox.Show("导出成功。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportToExcel(string fileName)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("发票识别结果");

            worksheet.Cells[1, 1].Value = "发票号码";
            worksheet.Cells[1, 2].Value = "发票代码";
            worksheet.Cells[1, 3].Value = "开票日期";
            worksheet.Cells[1, 4].Value = "购买方名称";
            worksheet.Cells[1, 5].Value = "购买方税号";
            worksheet.Cells[1, 6].Value = "销售方名称";
            worksheet.Cells[1, 7].Value = "销售方税号";
            worksheet.Cells[1, 8].Value = "金额合计";
            worksheet.Cells[1, 9].Value = "税额";
            worksheet.Cells[1, 10].Value = "价税合计";
            worksheet.Cells[1, 11].Value = "发票类型";
            worksheet.Cells[1, 12].Value = "文件路径";

            using (var range = worksheet.Cells[1, 1, 1, 12])
            {
                range.Style.Font.Bold = true;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(Color.LightGray);
            }

            for (int i = 0; i < invoiceResults.Count; i++)
            {
                InvoiceData invoice = invoiceResults[i];
                int row = i + 2;
                worksheet.Cells[row, 1].Value = invoice.InvoiceNum;
                worksheet.Cells[row, 2].Value = invoice.InvoiceCode;
                worksheet.Cells[row, 3].Value = invoice.InvoiceDate;
                worksheet.Cells[row, 4].Value = invoice.PurchaserName;
                worksheet.Cells[row, 5].Value = invoice.PurchaserRegisterNum;
                worksheet.Cells[row, 6].Value = invoice.SellerName;
                worksheet.Cells[row, 7].Value = invoice.SellerRegisterNum;
                worksheet.Cells[row, 8].Value = invoice.TotalAmount;
                worksheet.Cells[row, 9].Value = invoice.TotalTax;
                worksheet.Cells[row, 10].Value = invoice.AmountInFiguers;
                worksheet.Cells[row, 11].Value = invoice.InvoiceType;
                worksheet.Cells[row, 12].Value = invoice.ImagePath;
            }

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
            package.SaveAs(new FileInfo(fileName));
        }
    }
}
