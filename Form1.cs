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
                
                // 记录调试信息到文件
                System.Text.StringBuilder logBuilder = new System.Text.StringBuilder();
                logBuilder.AppendLine($"[{DateTime.Now}] 开始调用API...");
                
                string resultJson = baiDu.vat_invoice(base64Data, fileType);
                logBuilder.AppendLine($"[{DateTime.Now}] API调用完成");
                
                // 输出API返回结果的前500个字符，以便了解其结构
                logBuilder.AppendLine($"[{DateTime.Now}] API返回结果前500个字符: {resultJson.Substring(0, Math.Min(500, resultJson.Length))}");

                // 保存API返回的结果到文件，以便查看其结构
                try
                {
                    string fileName = Path.GetFileNameWithoutExtension(imagePath);
                    string outputPath = $"api_result_{fileName}.json";
                    System.IO.File.WriteAllText(outputPath, resultJson, System.Text.Encoding.UTF8);
                    logBuilder.AppendLine($"[{DateTime.Now}] API返回结果已保存到 {outputPath} 文件");
                }
                catch (Exception ex)
                {
                    logBuilder.AppendLine($"[{DateTime.Now}] 保存API结果时出错: {ex.Message}");
                    logBuilder.AppendLine($"[{DateTime.Now}] 错误堆栈: {ex.StackTrace}");
                }

                // 解析JSON结果
                logBuilder.AppendLine($"[{DateTime.Now}] 开始解析JSON结果...");
                dynamic result = EasyJson.ParseJsonToDynamic(resultJson);
                logBuilder.AppendLine($"[{DateTime.Now}] JSON结果解析完成");
                
                // 保存调试日志到文件
                try
                {
                    string logPath = "debug_log.txt";
                    System.IO.File.AppendAllText(logPath, logBuilder.ToString(), System.Text.Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    // 忽略保存日志时的错误
                }
                
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
                PurchaserRegisterNum = GetStringValue(wordsResult.PurchaserRegisterNum),
                SellerName = GetStringValue(wordsResult.SellerName),
                SellerRegisterNum = GetStringValue(wordsResult.SellerRegisterNum),
                TotalAmount = GetStringValue(wordsResult.TotalAmount),
                TotalTax = GetStringValue(wordsResult.TotalTax),
                AmountInFiguers = GetStringValue(wordsResult.AmountInFiguers),
                InvoiceType = GetStringValue(wordsResult.InvoiceType),
                RawData = wordsResult
            };

            // 提取商品明细信息
            try
            {
                // 记录调试信息
                System.Text.StringBuilder logBuilder = new System.Text.StringBuilder();
                logBuilder.AppendLine($"[{DateTime.Now}] 开始提取商品明细信息");

                // 根据API返回的实际结构提取商品明细信息
                // 商品信息分散在不同的字段中，都是数组形式
                var commodityNames = GetArrayValue(wordsResult.CommodityName);
                var commodityUnits = GetArrayValue(wordsResult.CommodityUnit);
                var commodityNums = GetArrayValue(wordsResult.CommodityNum);
                var commodityPrices = GetArrayValue(wordsResult.CommodityPrice);
                var commodityAmounts = GetArrayValue(wordsResult.CommodityAmount);
                var commodityTaxRates = GetArrayValue(wordsResult.CommodityTaxRate);
                var commodityTaxes = GetArrayValue(wordsResult.CommodityTax);

                // 计算商品数量，取所有数组中长度最大的那个
                int itemCount = Math.Max(
                    Math.Max(Math.Max(commodityNames.Length, commodityUnits.Length), 
                    Math.Max(commodityNums.Length, commodityPrices.Length)),
                    Math.Max(Math.Max(commodityAmounts.Length, commodityTaxRates.Length), 
                    commodityTaxes.Length)
                );

                logBuilder.AppendLine($"[{DateTime.Now}] 商品数量: {itemCount}");
                logBuilder.AppendLine($"[{DateTime.Now}] 商品名称数量: {commodityNames.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] 商品单位数量: {commodityUnits.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] 商品数量数量: {commodityNums.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] 商品单价数量: {commodityPrices.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] 商品金额数量: {commodityAmounts.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] 商品税率数量: {commodityTaxRates.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] 商品税额数量: {commodityTaxes.Length}");

                // 提取商品明细信息
                for (int i = 0; i < itemCount; i++)
                {
                    var commodityItem = new CommodityItem
                    {
                        Name = i < commodityNames.Length ? commodityNames[i] : "",
                        Unit = i < commodityUnits.Length ? commodityUnits[i] : "",
                        Quantity = i < commodityNums.Length ? commodityNums[i] : "",
                        Price = i < commodityPrices.Length ? commodityPrices[i] : "",
                        Amount = i < commodityAmounts.Length ? commodityAmounts[i] : "",
                        TaxRate = i < commodityTaxRates.Length ? commodityTaxRates[i] : "",
                        Tax = i < commodityTaxes.Length ? commodityTaxes[i] : ""
                    };
                    
                    // 尝试从商品名称中提取规格型号
                    if (string.IsNullOrEmpty(commodityItem.Specification) && !string.IsNullOrEmpty(commodityItem.Name))
                    {
                        // 简单的规则：如果商品名称包含空格，尝试将最后一部分作为规格型号
                        string[] parts = commodityItem.Name.Split(' ');
                        if (parts.Length > 1)
                        {
                            commodityItem.Specification = parts[parts.Length - 1];
                        }
                    }
                    
                    invoice.CommodityItems.Add(commodityItem);
                    logBuilder.AppendLine($"[{DateTime.Now}] 添加商品: {commodityItem.Name}");
                    logBuilder.AppendLine($"[{DateTime.Now}] 商品单位: {commodityItem.Unit}");
                    logBuilder.AppendLine($"[{DateTime.Now}] 商品数量: {commodityItem.Quantity}");
                    logBuilder.AppendLine($"[{DateTime.Now}] 商品单价: {commodityItem.Price}");
                    logBuilder.AppendLine($"[{DateTime.Now}] 商品金额: {commodityItem.Amount}");
                    logBuilder.AppendLine($"[{DateTime.Now}] 商品税率: {commodityItem.TaxRate}");
                    logBuilder.AppendLine($"[{DateTime.Now}] 商品税额: {commodityItem.Tax}");
                }

                logBuilder.AppendLine($"[{DateTime.Now}] 商品明细提取完成，共 {invoice.CommodityItems.Count} 个商品");
                
                // 保存调试日志
                try
                {
                    System.IO.File.AppendAllText("parse_log.txt", logBuilder.ToString(), System.Text.Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    // 忽略保存日志时的错误
                }
            }
            catch (Exception ex)
            {
                // 商品明细解析失败，记录错误但不影响整体解析
                try
                {
                    System.Text.StringBuilder logBuilder = new System.Text.StringBuilder();
                    logBuilder.AppendLine($"[{DateTime.Now}] 解析商品明细时出错: {ex.Message}");
                    logBuilder.AppendLine($"[{DateTime.Now}] 错误堆栈: {ex.StackTrace}");
                    System.IO.File.AppendAllText("parse_error_log.txt", logBuilder.ToString(), System.Text.Encoding.UTF8);
                }
                catch
                {
                    // 忽略保存错误日志时的错误
                }
            }

            return invoice;
        }

        // 获取数组类型的值，返回字符串数组
        private string[] GetArrayValue(dynamic value)
        {
            try
            {
                if (value == null)
                    return new string[0];
                
                // 检查是否为数组
                var enumerable = value as System.Collections.IEnumerable;
                if (enumerable != null)
                {
                    List<string> result = new List<string>();
                    foreach (var item in enumerable)
                    {
                        try
                        {
                            // 对于动态对象，尝试直接访问word属性
                            if (item != null)
                            {
                                dynamic dynamicItem = item;
                                if (dynamicItem.word != null)
                                {
                                    result.Add(dynamicItem.word.ToString());
                                }
                                else
                                {
                                    // 尝试直接转换为字符串
                                    result.Add(item.ToString());
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // 记录错误并尝试其他方式
                            try
                            {
                                System.IO.File.AppendAllText(
                                    "array_value_error_log.txt", 
                                    $"[{DateTime.Now}] 提取数组值时出错: {ex.Message}\n", 
                                    System.Text.Encoding.UTF8
                                );
                            }
                            catch
                            {
                                // 忽略错误
                            }
                            // 尝试直接转换为字符串
                            if (item != null)
                            {
                                result.Add(item.ToString());
                            }
                        }
                    }
                    return result.ToArray();
                }
                else
                {
                    // 尝试直接转换为字符串
                    return new string[] { value.ToString() };
                }
            }
            catch (Exception ex)
            {
                // 记录错误
                try
                {
                    System.IO.File.AppendAllText(
                        "array_value_error_log.txt", 
                        $"[{DateTime.Now}] 提取数组值时出错: {ex.Message}\n", 
                        System.Text.Encoding.UTF8
                    );
                }
                catch
                {
                    // 忽略错误
                }
                return new string[0];
            }
        }

        private string GetStringValue(dynamic value)
        {
            if (value == null) return "";
            return value.ToString();
        }

        private void AddInvoiceToListView(InvoiceData invoice)
        {
            if (invoice.CommodityItems.Count > 0)
            {
                // 如果有商品明细，为每个商品明细创建一行数据
                for (int i = 0; i < invoice.CommodityItems.Count; i++)
                {
                    var item = invoice.CommodityItems[i];
                    ListViewItem listItem = new ListViewItem(i == 0 ? invoice.InvoiceNum : "");
                    listItem.SubItems.Add(i == 0 ? invoice.InvoiceCode : "");
                    listItem.SubItems.Add(i == 0 ? invoice.InvoiceDate : "");
                    listItem.SubItems.Add(i == 0 ? invoice.PurchaserName : "");
                    listItem.SubItems.Add(i == 0 ? invoice.SellerName : "");
                    listItem.SubItems.Add(i == 0 ? invoice.PurchaserRegisterNum : "");
                    listItem.SubItems.Add(i == 0 ? invoice.SellerRegisterNum : "");
                    listItem.SubItems.Add(i == 0 ? invoice.TotalAmount : "");
                    listItem.SubItems.Add(i == 0 ? invoice.TotalTax : "");
                    listItem.SubItems.Add(i == 0 ? invoice.AmountInFiguers : "");
                    listItem.SubItems.Add(item.Name);
                    listItem.SubItems.Add(item.Specification);
                    listItem.SubItems.Add(item.Unit);
                    listItem.SubItems.Add(item.Quantity);
                    listItem.SubItems.Add(item.Price);
                    listItem.SubItems.Add(item.Amount);
                    listItem.SubItems.Add(item.TaxRate);
                    listItem.SubItems.Add(item.Tax);
                    listItem.SubItems.Add(i == 0 ? invoice.ImagePath : "");
                    listItem.Tag = invoice;
                    superListView.Items.Add(listItem);
                }
            }
            else
            {
                // 如果没有商品明细，创建一行基本信息
                ListViewItem item = new ListViewItem(invoice.InvoiceNum);
                item.SubItems.Add(invoice.InvoiceCode);
                item.SubItems.Add(invoice.InvoiceDate);
                item.SubItems.Add(invoice.PurchaserName);
                item.SubItems.Add(invoice.SellerName);
                item.SubItems.Add(invoice.PurchaserRegisterNum);
                item.SubItems.Add(invoice.SellerRegisterNum);
                item.SubItems.Add(invoice.TotalAmount);
                item.SubItems.Add(invoice.TotalTax);
                item.SubItems.Add(invoice.AmountInFiguers);
                item.SubItems.Add("");
                item.SubItems.Add("");
                item.SubItems.Add("");
                item.SubItems.Add("");
                item.SubItems.Add("");
                item.SubItems.Add("");
                item.SubItems.Add("");
                item.SubItems.Add("");
                item.SubItems.Add(invoice.ImagePath);
                item.Tag = invoice;
                superListView.Items.Add(item);
            }
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
                worksheet.Cells[1, 8].Value = "商品名称";
                worksheet.Cells[1, 9].Value = "规格型号";
                worksheet.Cells[1, 10].Value = "单位";
                worksheet.Cells[1, 11].Value = "数量";
                worksheet.Cells[1, 12].Value = "单价";
                worksheet.Cells[1, 13].Value = "金额";
                worksheet.Cells[1, 14].Value = "税率";
                worksheet.Cells[1, 15].Value = "税额";
                worksheet.Cells[1, 16].Value = "金额合计";
                worksheet.Cells[1, 17].Value = "税额合计";
                worksheet.Cells[1, 18].Value = "价税合计";
                worksheet.Cells[1, 19].Value = "发票类型";
                worksheet.Cells[1, 20].Value = "文件路径";

                // 设置表头样式
                using (var range = worksheet.Cells[1, 1, 1, 20])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // 填充数据
                int row = 2;
                foreach (var invoice in invoiceResults)
                {
                    // 为每个发票创建一行，所有商品明细都在这一行中
                    worksheet.Cells[row, 1].Value = invoice.InvoiceNum;
                    worksheet.Cells[row, 2].Value = invoice.InvoiceCode;
                    worksheet.Cells[row, 3].Value = invoice.InvoiceDate;
                    worksheet.Cells[row, 4].Value = invoice.PurchaserName;
                    worksheet.Cells[row, 5].Value = invoice.PurchaserRegisterNum;
                    worksheet.Cells[row, 6].Value = invoice.SellerName;
                    worksheet.Cells[row, 7].Value = invoice.SellerRegisterNum;
                    
                    if (invoice.CommodityItems.Count > 0)
                    {
                        // 收集所有商品明细信息，用分号连接
                        var names = new List<string>();
                        var specifications = new List<string>();
                        var units = new List<string>();
                        var quantities = new List<string>();
                        var prices = new List<string>();
                        var amounts = new List<string>();
                        var taxRates = new List<string>();
                        var taxes = new List<string>();
                        
                        foreach (var item in invoice.CommodityItems)
                        {
                            names.Add(item.Name);
                            specifications.Add(item.Specification);
                            units.Add(item.Unit);
                            quantities.Add(item.Quantity);
                            prices.Add(item.Price);
                            amounts.Add(item.Amount);
                            taxRates.Add(item.TaxRate);
                            taxes.Add(item.Tax);
                        }
                        
                        // 将收集的信息用分号连接并填充到单元格
                        worksheet.Cells[row, 8].Value = string.Join("; ", names);
                        worksheet.Cells[row, 9].Value = string.Join("; ", specifications);
                        worksheet.Cells[row, 10].Value = string.Join("; ", units);
                        worksheet.Cells[row, 11].Value = string.Join("; ", quantities);
                        worksheet.Cells[row, 12].Value = string.Join("; ", prices);
                        worksheet.Cells[row, 13].Value = string.Join("; ", amounts);
                        worksheet.Cells[row, 14].Value = string.Join("; ", taxRates);
                        worksheet.Cells[row, 15].Value = string.Join("; ", taxes);
                    }
                    else
                    {
                        // 如果没有商品明细，留空
                        worksheet.Cells[row, 8].Value = "";
                        worksheet.Cells[row, 9].Value = "";
                        worksheet.Cells[row, 10].Value = "";
                        worksheet.Cells[row, 11].Value = "";
                        worksheet.Cells[row, 12].Value = "";
                        worksheet.Cells[row, 13].Value = "";
                        worksheet.Cells[row, 14].Value = "";
                        worksheet.Cells[row, 15].Value = "";
                    }
                    
                    // 填充发票的其他信息
                    worksheet.Cells[row, 16].Value = invoice.TotalAmount;
                    worksheet.Cells[row, 17].Value = invoice.TotalTax;
                    worksheet.Cells[row, 18].Value = invoice.AmountInFiguers;
                    worksheet.Cells[row, 19].Value = invoice.InvoiceType;
                    worksheet.Cells[row, 20].Value = invoice.ImagePath;
                    
                    row++;
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
            public string PurchaserRegisterNum { get; set; } = "";
            public string SellerName { get; set; } = "";
            public string SellerRegisterNum { get; set; } = "";
            public string TotalAmount { get; set; } = "";
            public string TotalTax { get; set; } = "";
            public string AmountInFiguers { get; set; } = "";
            public string InvoiceType { get; set; } = "";
            public string ImagePath { get; set; } = "";
            public List<CommodityItem> CommodityItems { get; set; } = new List<CommodityItem>();
            public dynamic? RawData { get; set; }
        }

        private class CommodityItem
        {
            public string Name { get; set; } = "";
            public string Specification { get; set; } = "";
            public string Unit { get; set; } = "";
            public string Quantity { get; set; } = "";
            public string Price { get; set; } = "";
            public string Amount { get; set; } = "";
            public string TaxRate { get; set; } = "";
            public string Tax { get; set; } = "";
        }
    }
}
