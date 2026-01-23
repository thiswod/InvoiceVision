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
                        "璇峰湪 appsettings.json 鏂囦欢涓厤缃櫨搴CR API瀵嗛挜锛乗n\n" +
                        "璇峰弬鑰?appsettings.example.json 鏂囦欢鏍煎紡杩涜閰嶇疆銆?,
                        "閰嶇疆閿欒",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                baiDu = new BaiDu(apiKey, secretKey);
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show(
                    "鏈壘鍒?appsettings.json 閰嶇疆鏂囦欢锛乗n\n" +
                    "璇峰鍒?appsettings.example.json 涓?appsettings.json 骞堕厤缃偍鐨凙PI瀵嗛挜銆?,
                    "閰嶇疆鏂囦欢缂哄け",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"鍔犺浇閰嶇疆鏂囦欢鏃跺嚭閿欙細{ex.Message}\n\n" +
                    "璇锋鏌?appsettings.json 鏂囦欢鏍煎紡鏄惁姝ｇ‘銆?,
                    "閰嶇疆閿欒",
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
                openFileDialog.Filter = "鍥剧墖鍜孭DF鏂囦欢|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.pdf|鍥剧墖鏂囦欢|*.jpg;*.jpeg;*.png;*.bmp;*.gif|PDF鏂囦欢|*.pdf|鎵�鏈夋枃浠秥*.*";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "閫夋嫨鍙戠エ鍥剧墖鎴朠DF鏂囦欢";

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
                    "API瀵嗛挜鏈厤缃紒\n\n" +
                    "璇烽厤缃?appsettings.json 鏂囦欢涓殑鐧惧害OCR API瀵嗛挜銆?,
                    "閰嶇疆閿欒",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (listBoxImages.Items.Count == 0)
            {
                MessageBox.Show("璇峰厛閫夋嫨鍥剧墖鎴朠DF鏂囦欢锛?, "鎻愮ず", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                labelStatus.Text = $"璇嗗埆瀹屾垚锛屽叡璇嗗埆 {invoiceResults.Count} 寮犲彂绁?;
                btnExport.Enabled = invoiceResults.Count > 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"璇嗗埆杩囩▼涓嚭閿欙細{ex.Message}", "閿欒", MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelStatus.Text = "璇嗗埆澶辫触";
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
            int minDelayMs = 500; // 鏈�灏忛棿闅?00ms锛岀‘淇濅笉瓒呰繃2 QPS
            int processedCount = 0;

            foreach (string imagePath in listBoxImages.Items.Cast<string>())
            {
                try
                {
                    // 鎺у埗QPS锛氭瘡娆¤姹備箣闂磋嚦灏戦棿闅?00ms
                    if (processedCount > 0)
                    {
                        System.Threading.Thread.Sleep(minDelayMs);
                    }

                    ProcessSingleImage(imagePath);
                    processedCount++;
                    progressBar.Value = processedCount;
                    labelStatus.Text = $"姝ｅ湪璇嗗埆... ({processedCount}/{listBoxImages.Items.Count})";
                    Application.DoEvents(); // 鏇存柊UI
                }
                catch (Exception ex)
                {
                    labelStatus.Text = $"澶勭悊 {Path.GetFileName(imagePath)} 鏃跺嚭閿? {ex.Message}";
                    processedCount++;
                    progressBar.Value = processedCount;
                    Application.DoEvents(); // 鏇存柊UI
                }
            }
        }

        private void ProcessSingleImage(string imagePath)
        {
            try
            {
                // 璇诲彇鏂囦欢骞惰浆鎹负base64
                byte[] fileBytes = File.ReadAllBytes(imagePath);
                string base64Data = Convert.ToBase64String(fileBytes);

                // 鑾峰彇鏂囦欢绫诲瀷锛堟牴鎹枃浠舵墿灞曞悕锛?
                string fileType = "png"; // 榛樿
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

                // 璋冪敤API璇嗗埆
                if (baiDu == null)
                {
                    throw new InvalidOperationException("API瀵嗛挜鏈厤缃紝鏃犳硶杩涜璇嗗埆銆?);
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
                MessageBox.Show("娌℃湁鍙鍑虹殑鏁版嵁锛?, "鎻愮ず", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel鏂囦欢|*.xlsx|鎵�鏈夋枃浠秥*.*";
                saveFileDialog.FileName = $"鍙戠エ璇嗗埆缁撴灉_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                saveFileDialog.Title = "淇濆瓨Excel鏂囦欢";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExportToExcel(saveFileDialog.FileName);
                        MessageBox.Show("瀵煎嚭鎴愬姛锛?, "鎻愮ず", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"瀵煎嚭澶辫触锛歿ex.Message}", "閿欒", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ExportToExcel(string fileName)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("鍙戠エ璇嗗埆缁撴灉");

                // 璁剧疆琛ㄥご
                worksheet.Cells[1, 1].Value = "鍙戠エ鍙风爜";
                worksheet.Cells[1, 2].Value = "鍙戠エ浠ｇ爜";
                worksheet.Cells[1, 3].Value = "寮�绁ㄦ棩鏈?;
                worksheet.Cells[1, 4].Value = "璐拱鏂瑰悕绉?;
                worksheet.Cells[1, 5].Value = "璐拱鏂圭◣鍙?;
                worksheet.Cells[1, 6].Value = "閿�鍞柟鍚嶇О";
                worksheet.Cells[1, 7].Value = "閿�鍞柟绋庡彿";
                worksheet.Cells[1, 8].Value = "鍟嗗搧鍚嶇О";
                worksheet.Cells[1, 9].Value = "瑙勬牸鍨嬪彿";
                worksheet.Cells[1, 10].Value = "鍗曚綅";
                worksheet.Cells[1, 11].Value = "鏁伴噺";
                worksheet.Cells[1, 12].Value = "鍗曚环";
                worksheet.Cells[1, 13].Value = "閲戦";
                worksheet.Cells[1, 14].Value = "绋庣巼";
                worksheet.Cells[1, 15].Value = "绋庨";
                worksheet.Cells[1, 16].Value = "閲戦鍚堣";
                worksheet.Cells[1, 17].Value = "绋庨鍚堣";
                worksheet.Cells[1, 18].Value = "浠风◣鍚堣";
                worksheet.Cells[1, 19].Value = "鍙戠エ绫诲瀷";
                worksheet.Cells[1, 20].Value = "鏂囦欢璺緞";

                // 璁剧疆琛ㄥご鏍峰紡
                using (var range = worksheet.Cells[1, 1, 1, 20])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // 濉厖鏁版嵁
                int row = 2;
                foreach (var invoice in invoiceResults)
                {
                    // 涓烘瘡涓彂绁ㄥ垱寤轰竴琛岋紝鎵�鏈夊晢鍝佹槑缁嗛兘鍦ㄨ繖涓�琛屼腑
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

                // 鑷姩璋冩暣鍒楀
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // 淇濆瓨鏂囦欢
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
