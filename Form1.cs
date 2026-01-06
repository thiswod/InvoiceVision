using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using WodToolkit.Json;

namespace InvoiceVision
{
    public partial class Form1 : Form
    {
        private BaiDu? baiDu;
        private List<InvoiceData> invoiceResults = new List<InvoiceData>();
        private string? apiKey;
        private string? secretKey;

        public Form1()
        {
            InitializeComponent();
            LoadConfiguration();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void LoadConfiguration()
        {
            try
            {
                var builder = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                var configuration = builder.Build();
                apiKey = configuration["BaiduOCR:ApiKey"] ?? "";
                secretKey = configuration["BaiduOCR:SecretKey"] ?? "";

                if (string.IsNullOrWhiteSpace(apiKey) || string.IsNullOrWhiteSpace(secretKey))
                {
                    MessageBox.Show(
                        "请在 appsettings.json 文件中配置百度OCR API密钥！\n\n" +
                        "请参考 appsettings.example.json 文件格式进行配置。",
                        "配置错误",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                baiDu = new BaiDu(apiKey, secretKey);
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show(
                    "未找到 appsettings.json 配置文件！\n\n" +
                    "请复制 appsettings.example.json 为 appsettings.json 并配置您的API密钥。",
                    "配置文件缺失",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"加载配置文件时出错：{ex.Message}\n\n" +
                    "请检查 appsettings.json 文件格式是否正确。",
                    "配置错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void BtnSelectImages_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "图片和PDF文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.pdf|图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif|PDF文件|*.pdf|所有文件|*.*";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "选择发票图片或PDF文件";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    listBoxImages.Items.Clear();
                    foreach (string fileName in openFileDialog.FileNames)
                    {
                        listBoxImages.Items.Add(fileName);
                    }
                    btnStart.Enabled = listBoxImages.Items.Count > 0;
                }
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
            if (baiDu == null)
            {
                MessageBox.Show(
                    "API密钥未配置！\n\n" +
                    "请配置 appsettings.json 文件中的百度OCR API密钥。",
                    "配置错误",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (listBoxImages.Items.Count == 0)
            {
                MessageBox.Show("请先选择图片或PDF文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                ProcessImages();
                labelStatus.Text = $"识别完成，共识别 {invoiceResults.Count} 张发票";
                btnExport.Enabled = invoiceResults.Count > 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"识别过程中出错：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelStatus.Text = "识别失败";
            }
            finally
            {
                btnStart.Enabled = true;
                btnSelectImages.Enabled = true;
                progressBar.Visible = false;
            }
        }

        private void ProcessImages()
        {
            int minDelayMs = 500; // 最小间隔500ms，确保不超过2 QPS
            int processedCount = 0;

            foreach (string imagePath in listBoxImages.Items.Cast<string>())
            {
                try
                {
                    // 控制QPS：每次请求之间至少间隔500ms
                    if (processedCount > 0)
                    {
                        System.Threading.Thread.Sleep(minDelayMs);
                    }

                    ProcessSingleImage(imagePath);
                    processedCount++;
                    progressBar.Value = processedCount;
                    labelStatus.Text = $"正在识别... ({processedCount}/{listBoxImages.Items.Count})";
                    Application.DoEvents(); // 更新UI
                }
                catch (Exception ex)
                {
                    labelStatus.Text = $"处理 {Path.GetFileName(imagePath)} 时出错: {ex.Message}";
                    processedCount++;
                    progressBar.Value = processedCount;
                    Application.DoEvents(); // 更新UI
                }
            }
        }

        private void ProcessSingleImage(string imagePath)
        {
            try
            {
                // 读取文件并转换为base64
                byte[] fileBytes = File.ReadAllBytes(imagePath);
                string base64Data = Convert.ToBase64String(fileBytes);

                // 获取文件类型（根据文件扩展名）
                string fileType = "png"; // 默认
                string extension = Path.GetExtension(imagePath).ToLower();
                if (extension == ".jpg" || extension == ".jpeg")
                    fileType = "jpeg";
                else if (extension == ".png")
                    fileType = "png";
                else if (extension == ".bmp")
                    fileType = "bmp";
                else if (extension == ".gif")
                    fileType = "gif";
                else if (extension == ".pdf")
                    fileType = "pdf";

                // 调用API识别
                if (baiDu == null)
                {
                    throw new InvalidOperationException("API密钥未配置，无法进行识别。");
                }
                string resultJson = baiDu.vat_invoice(base64Data, fileType);

                // 解析JSON结果
                dynamic result = EasyJson.ParseJsonToDynamic(resultJson);
                
                // 检查是否有words_result字段，有则表示识别成功
                if (result.words_result != null)
                {
                    var invoiceData = ParseInvoiceData(result.words_result, imagePath);
                    invoiceResults.Add(invoiceData);
                    AddInvoiceToListView(invoiceData);
                }
                else
                {
                    // 如果没有words_result，可能是出错了，尝试获取错误信息
                    string errorMsg = "识别结果为空";
                    try
                    {
                        if (result.error_code != null)
                        {
                            errorMsg = $"API返回错误: {result.error_msg ?? "未知错误"} (错误码: {result.error_code})";
                        }
                    }
                    catch
                    {
                        // 如果无法获取错误信息，使用默认消息
                    }
                    throw new Exception(errorMsg);
                }
            }
            catch (Exception ex)
            {
                string fileType = Path.GetExtension(imagePath).ToLower() == ".pdf" ? "PDF文件" : "图片";
                throw new Exception($"处理{fileType} {Path.GetFileName(imagePath)} 时出错: {ex.Message}", ex);
            }
        }

        private InvoiceData ParseInvoiceData(dynamic wordsResult, string imagePath)
        {
            string invoiceNum = GetStringValue(wordsResult.InvoiceNum);
            string invoiceCode = GetStringValue(wordsResult.InvoiceCode);
            
            // 如果发票代码为空，使用InvoiceNum作为发票代码
            // 根据用户反馈，InvoiceNum实际上就是发票代码
            if (string.IsNullOrEmpty(invoiceCode) && !string.IsNullOrEmpty(invoiceNum))
            {
                invoiceCode = invoiceNum;
            }

            var invoice = new InvoiceData
            {
                ImagePath = imagePath,
                InvoiceNum = invoiceNum,
                InvoiceCode = invoiceCode,
                InvoiceDate = GetStringValue(wordsResult.InvoiceDate),
                PurchaserName = GetStringValue(wordsResult.PurchaserName),
                SellerName = GetStringValue(wordsResult.SellerName),
                TotalAmount = GetStringValue(wordsResult.TotalAmount),
                TotalTax = GetStringValue(wordsResult.TotalTax),
                AmountInFiguers = GetStringValue(wordsResult.AmountInFiguers),
                InvoiceType = GetStringValue(wordsResult.InvoiceType),
                RawData = wordsResult
            };
            return invoice;
        }

        private string GetStringValue(dynamic value)
        {
            if (value == null) return "";
            return value.ToString();
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
                MessageBox.Show("没有可导出的数据！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel文件|*.xlsx|所有文件|*.*";
                saveFileDialog.FileName = $"发票识别结果_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                saveFileDialog.Title = "保存Excel文件";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExportToExcel(saveFileDialog.FileName);
                        MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"导出失败：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ExportToExcel(string fileName)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("发票识别结果");

                // 设置表头
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

                // 设置表头样式
                using (var range = worksheet.Cells[1, 1, 1, 12])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // 填充数据
                for (int i = 0; i < invoiceResults.Count; i++)
                {
                    var invoice = invoiceResults[i];
                    int row = i + 2;
                    worksheet.Cells[row, 1].Value = invoice.InvoiceNum;
                    worksheet.Cells[row, 2].Value = invoice.InvoiceCode;
                    worksheet.Cells[row, 3].Value = invoice.InvoiceDate;
                    worksheet.Cells[row, 4].Value = invoice.PurchaserName;
                    worksheet.Cells[row, 5].Value = GetStringValue(invoice.RawData?.PurchaserRegisterNum);
                    worksheet.Cells[row, 6].Value = invoice.SellerName;
                    worksheet.Cells[row, 7].Value = GetStringValue(invoice.RawData?.SellerRegisterNum);
                    worksheet.Cells[row, 8].Value = invoice.TotalAmount;
                    worksheet.Cells[row, 9].Value = invoice.TotalTax;
                    worksheet.Cells[row, 10].Value = invoice.AmountInFiguers;
                    worksheet.Cells[row, 11].Value = invoice.InvoiceType;
                    worksheet.Cells[row, 12].Value = invoice.ImagePath;
                }

                // 自动调整列宽
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // 保存文件
                package.SaveAs(new FileInfo(fileName));
            }
        }

        private class InvoiceData
        {
            public string InvoiceNum { get; set; } = "";
            public string InvoiceCode { get; set; } = "";
            public string InvoiceDate { get; set; } = "";
            public string PurchaserName { get; set; } = "";
            public string SellerName { get; set; } = "";
            public string TotalAmount { get; set; } = "";
            public string TotalTax { get; set; } = "";
            public string AmountInFiguers { get; set; } = "";
            public string InvoiceType { get; set; } = "";
            public string ImagePath { get; set; } = "";
            public dynamic? RawData { get; set; }
        }
    }
}
