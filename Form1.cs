using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace InvoiceVision
{
    public partial class Form1 : Form
    {
        private BaiDu? baiDu;
        private List<InvoiceData> invoiceResults = new List<InvoiceData>();

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
                // UmiOCR 不需要 API 密钥，直接初始化
                baiDu = new BaiDu();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"初始化 OCR 时出错：{ex.Message}",
                    "初始化错误",
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
                    "OCR未初始化！\n\n" +
                    "请重新启动程序。",
                    "初始化错误",
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
            // UmiOCR 是本地 OCR，不需要 QPS 控制
            int processedCount = 0;

            foreach (string imagePath in listBoxImages.Items.Cast<string>())
            {
                try
                {
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

                // 调用UmiOCR识别
                if (baiDu == null)
                {
                    throw new InvalidOperationException("OCR未初始化，无法进行识别。");
                }
                string resultText = baiDu.vat_invoice(imagePath, fileType);

                // UmiOCR 返回的是纯文本，需要解析
                if (string.IsNullOrWhiteSpace(resultText))
                {
                    throw new Exception("识别结果为空");
                }

                // 解析文本结果，提取发票信息
                var invoiceData = ParseInvoiceDataFromText(resultText, imagePath);
                invoiceResults.Add(invoiceData);
                AddInvoiceToListView(invoiceData);
            }
            catch (Exception ex)
            {
                string fileType = Path.GetExtension(imagePath).ToLower() == ".pdf" ? "PDF文件" : "图片";
                throw new Exception($"处理{fileType} {Path.GetFileName(imagePath)} 时出错: {ex.Message}", ex);
            }
        }

        private InvoiceData ParseInvoiceDataFromText(string text, string imagePath)
        {
            // 从文本中提取发票信息（使用正则表达式和关键字匹配）
            var invoice = new InvoiceData
            {
                ImagePath = imagePath,
                InvoiceNum = ExtractInvoiceNum(text),
                InvoiceCode = ExtractInvoiceCode(text),
                InvoiceDate = ExtractInvoiceDate(text),
                PurchaserName = ExtractPurchaserName(text),
                SellerName = ExtractSellerName(text),
                TotalAmount = ExtractTotalAmount(text),
                TotalTax = ExtractTotalTax(text),
                AmountInFiguers = ExtractAmountInFigures(text),
                InvoiceType = ExtractInvoiceType(text),
                RawData = text
            };

            // 如果发票代码为空，使用InvoiceNum作为发票代码
            if (string.IsNullOrEmpty(invoice.InvoiceCode) && !string.IsNullOrEmpty(invoice.InvoiceNum))
            {
                invoice.InvoiceCode = invoice.InvoiceNum;
            }

            return invoice;
        }

        private string ExtractInvoiceNum(string text)
        {
            // 匹配：发票号码: 25447000001498458680
            var match = Regex.Match(text, @"发票号码[：:]\s*(\d{10,})");
            if (match.Success)
                return match.Groups[1].Value;
            
            return ExtractField(text, new[] { "发票号码", "发票号" });
        }

        private string ExtractInvoiceCode(string text)
        {
            // 发票代码通常在发票号码行的下一行，格式类似：914419000585344943
            // 先尝试匹配"发票代码"关键字
            var match = Regex.Match(text, @"发票代码[：:]\s*(\d{10,})");
            if (match.Success)
                return match.Groups[1].Value;
            
            // 如果没有"发票代码"关键字，尝试在发票号码行附近查找
            // 格式：发票号码: 25447000001498458680 开票日期: 2025年11月14日
            //       914419000585344943  （下一行通常是发票代码）
            var invoiceNumMatch = Regex.Match(text, @"发票号码[：:]\s*\d{10,}.*?开票日期");
            if (invoiceNumMatch.Success)
            {
                int startPos = invoiceNumMatch.Index + invoiceNumMatch.Length;
                string afterDate = text.Substring(startPos);
                // 查找下一行的数字（通常是10-20位）
                var codeMatch = Regex.Match(afterDate, @"^\s*(\d{10,20})", RegexOptions.Multiline);
                if (codeMatch.Success)
                    return codeMatch.Groups[1].Value;
            }
            
            return ExtractField(text, new[] { "发票代码", "代码" });
        }

        private string ExtractInvoiceDate(string text)
        {
            // 匹配：开票日期: 2025年11月14日 或 2025年11月25日
            var match = Regex.Match(text, @"开票日期[：:]\s*(\d{4}年\d{1,2}月\d{1,2}日)");
            if (match.Success)
                return match.Groups[1].Value;
            
            return ExtractField(text, new[] { "开票日期", "日期" });
        }

        private string ExtractPurchaserName(string text)
        {
            // 匹配：购买方信息后的名称
            // 格式：购买方信息\n税\n名称: 郑州琳之星通讯有限公司统一社会信用代码/纳税人识别号:
            var match = Regex.Match(text, @"购买方信息[\s\S]*?名称[：:]\s*([^\n统一社会信用代码纳税人识别号/]{2,50})");
            if (match.Success)
            {
                string name = match.Groups[1].Value.Trim();
                // 移除可能的后缀
                name = Regex.Replace(name, @"统一.*$", "").Trim();
                if (!string.IsNullOrWhiteSpace(name))
                    return name;
            }
            
            // 备用方案：直接查找"名称: "后面的公司名
            var nameMatch = Regex.Match(text, @"购买方[\s\S]*?名称[：:]\s*([^\n统一社会信用代码纳税人识别号/：:]{2,50})");
            if (nameMatch.Success)
            {
                string name = nameMatch.Groups[1].Value.Trim();
                name = Regex.Replace(name, @"统一.*$", "").Trim();
                if (!string.IsNullOrWhiteSpace(name))
                    return name;
            }
            
            return ExtractField(text, new[] { "购买方", "买方" });
        }

        private string ExtractSellerName(string text)
        {
            // 匹配：销售方信息后的名称
            // 格式：销售\n方信息\n\n名称: 华为终端有限公司
            var match = Regex.Match(text, @"销售[\s\S]*?方信息[\s\S]*?名称[：:]\s*([^\n统一社会信用代码纳税人识别号/：:]{2,50})");
            if (match.Success)
            {
                string name = match.Groups[1].Value.Trim();
                name = Regex.Replace(name, @"统一.*$", "").Trim();
                if (!string.IsNullOrWhiteSpace(name))
                    return name;
            }
            
            // 备用方案：查找"销售"关键字后的"名称: "
            var nameMatch = Regex.Match(text, @"销售[\s\S]*?名称[：:]\s*([^\n统一社会信用代码纳税人识别号/：:]{2,50})");
            if (nameMatch.Success)
            {
                string name = nameMatch.Groups[1].Value.Trim();
                name = Regex.Replace(name, @"统一.*$", "").Trim();
                if (!string.IsNullOrWhiteSpace(name))
                    return name;
            }
            
            return ExtractField(text, new[] { "销售方", "卖方" });
        }

        private string ExtractTotalAmount(string text)
        {
            // 匹配：金额合计 或 合计 后的数字
            var match = Regex.Match(text, @"(?:金额合计|合计)[：:]*\s*(?:¥|￥)?\s*([\d,]+\.?\d*)");
            if (match.Success)
                return match.Groups[1].Value;
            
            // 备用方案：查找"金额"关键字后的数字（格式：金额 2211.50）
            var amountMatch = Regex.Match(text, @"金额\s+(?:¥|￥)?\s*([\d,]+\.?\d*)");
            if (amountMatch.Success)
                return amountMatch.Groups[1].Value;
            
            return ExtractField(text, new[] { "金额合计", "合计", "金额" });
        }

        private string ExtractTotalTax(string text)
        {
            // 匹配：税额后的数字
            var match = Regex.Match(text, @"税额\s*(?:¥|￥)?\s*([\d,]+\.?\d*)");
            if (match.Success)
                return match.Groups[1].Value;
            
            return ExtractField(text, new[] { "税额", "税" });
        }

        private string ExtractAmountInFigures(string text)
        {
            // 匹配：价税合计（小写）¥1983.32 或类似格式
            var match = Regex.Match(text, @"价税合计[（(（]*[小]*[写]*[）)）]*[：:]*\s*(?:¥|￥)?\s*([\d,]+\.?\d*)");
            if (match.Success)
                return match.Groups[1].Value;
            
            // 备用方案：查找"（小写）¥"后的数字（格式：（小写）¥1983.32）
            var amountMatch = Regex.Match(text, @"（小写）[¥￥]\s*([\d,]+\.?\d*)");
            if (amountMatch.Success)
                return amountMatch.Groups[1].Value;
            
            // 再尝试：查找"小写"后的数字
            var amountMatch2 = Regex.Match(text, @"小写[）)]*\s*[¥￥]?\s*([\d,]+\.?\d*)");
            if (amountMatch2.Success)
                return amountMatch2.Groups[1].Value;
            
            return ExtractField(text, new[] { "价税合计", "总计" });
        }

        private string ExtractInvoiceType(string text)
        {
            // 匹配发票类型
            var match = Regex.Match(text, @"(?:电子发票|增值税专用发票|增值税普通发票|普通发票)");
            if (match.Success)
                return match.Value;
            
            return ExtractField(text, new[] { "发票类型", "类型" });
        }

        private string ExtractField(string text, string[] keywords)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            foreach (var keyword in keywords)
            {
                int index = text.IndexOf(keyword, StringComparison.OrdinalIgnoreCase);
                if (index >= 0)
                {
                    // 尝试提取关键字后的内容
                    int startIndex = index + keyword.Length;
                    string remaining = text.Substring(startIndex).Trim();
                    
                    // 跳过冒号（中文或英文）
                    if (remaining.StartsWith(":") || remaining.StartsWith("："))
                        remaining = remaining.Substring(1).Trim();
                    
                    // 跳过空格
                    remaining = remaining.TrimStart();
                    
                    // 提取到换行符、制表符或下一个关键字之前的内容
                    int endIndex = -1;
                    
                    // 查找换行符
                    int newlineIndex = remaining.IndexOfAny(new[] { '\r', '\n' });
                    if (newlineIndex > 0)
                        endIndex = newlineIndex;
                    
                    // 查找制表符
                    int tabIndex = remaining.IndexOf('\t');
                    if (tabIndex > 0 && (endIndex < 0 || tabIndex < endIndex))
                        endIndex = tabIndex;
                    
                    // 对于某些字段，查找下一个可能的关键字
                    if (keyword.Contains("号码") || keyword.Contains("日期"))
                    {
                        // 对于发票号码和日期，可能在空格分隔的同一行，查找下一个关键字
                        int spaceIndex = remaining.IndexOf(' ');
                        if (spaceIndex > 0 && spaceIndex < 50) // 限制在合理范围内
                        {
                            string potentialValue = remaining.Substring(0, spaceIndex).Trim();
                            // 如果看起来像是一个完整的值（如发票号码通常是数字），使用它
                            if (potentialValue.Length > 5 && !potentialValue.Contains("信息") && !potentialValue.Contains("名称"))
                            {
                                return potentialValue;
                            }
                        }
                    }
                    
                    if (endIndex > 0)
                    {
                        string value = remaining.Substring(0, endIndex).Trim();
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            // 清理值，移除常见的后缀关键字
                            value = value.Split(new[] { "统一", "社会", "信用", "代码", "纳税人", "识别号" }, StringSplitOptions.None)[0].Trim();
                            if (!string.IsNullOrWhiteSpace(value))
                                return value;
                        }
                    }
                    else if (remaining.Length > 0)
                    {
                        // 如果没有明确的结束符，尝试提取前100个字符
                        string value = remaining.Length > 100 ? remaining.Substring(0, 100).Trim() : remaining.Trim();
                        // 尝试在空格或常见分隔符处截断
                        int cutIndex = value.IndexOfAny(new[] { ' ', '统', '社', '信', '代', '纳', '税', '人', '识', '别', '号' });
                        if (cutIndex > 0 && cutIndex < 80)
                            value = value.Substring(0, cutIndex).Trim();
                        if (!string.IsNullOrWhiteSpace(value))
                            return value;
                    }
                }
            }
            return "";
        }

        private string GetStringValue(dynamic value)
        {
            if (value == null) return "";
            return value.ToString();
        }

        /// <summary>
        /// 安全地检查动态对象是否包含指定属性
        /// </summary>
        private bool HasProperty(dynamic obj, string propertyName)
        {
            if (obj == null) return false;
            
            try
            {
                // ExpandoObject实现了IDictionary<string, object>接口
                if (obj is System.Collections.Generic.IDictionary<string, object> dict)
                {
                    return dict.ContainsKey(propertyName);
                }
                
                // 如果转换失败，尝试使用反射
                var type = ((object)obj).GetType();
                return type.GetProperty(propertyName) != null || type.GetField(propertyName) != null;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// 安全地获取动态对象的属性值
        /// </summary>
        private dynamic GetPropertyValue(dynamic obj, string propertyName)
        {
            if (obj == null) return null;
            
            try
            {
                // ExpandoObject实现了IDictionary<string, object>接口
                if (obj is System.Collections.Generic.IDictionary<string, object> dict)
                {
                    return dict.TryGetValue(propertyName, out var value) ? value : null;
                }
                
                // 如果转换失败，尝试使用反射
                var type = ((object)obj).GetType();
                var property = type.GetProperty(propertyName);
                if (property != null)
                {
                    return property.GetValue(obj);
                }
                
                var field = type.GetField(propertyName);
                if (field != null)
                {
                    return field.GetValue(obj);
                }
                
                return null;
            }
            catch
            {
                return null;
            }
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
                    worksheet.Cells[row, 5].Value = ExtractField(invoice.RawData?.ToString() ?? "", new[] { "购买方税号", "购买方纳税人识别号" });
                    worksheet.Cells[row, 6].Value = invoice.SellerName;
                    worksheet.Cells[row, 7].Value = ExtractField(invoice.RawData?.ToString() ?? "", new[] { "销售方税号", "销售方纳税人识别号" });
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
