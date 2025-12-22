using System;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using WodToolkit.Json;

namespace InvoiceVision
{
    /// <summary>
    /// 发票信息解析器，从文本中提取发票关键信息
    /// </summary>
    public class InvoiceParser
    {
        /// <summary>
        /// 从文本中解析发票信息
        /// </summary>
        /// <param name="text">PDF提取的文本内容</param>
        /// <param name="pdfPath">PDF文件路径</param>
        /// <returns>解析后的发票数据JSON字符串（格式与百度API返回格式兼容）</returns>
        public string ParseInvoice(string text, string pdfPath)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                throw new Exception("PDF文本内容为空，无法解析发票信息");
            }

            // 调试：保存提取的文本到文件（可选，用于调试）
            try
            {
                string debugPath = Path.Combine(Path.GetDirectoryName(pdfPath) ?? "", 
                    Path.GetFileNameWithoutExtension(pdfPath) + "_extracted_text.txt");
                File.WriteAllText(debugPath, text, System.Text.Encoding.UTF8);
            }
            catch
            {
                // 忽略调试文件写入错误
            }

            // 创建结果对象（模拟百度API返回格式）
            var result = new
            {
                words_result = new
                {
                    InvoiceNum = ExtractInvoiceNum(text),
                    InvoiceCode = ExtractInvoiceCode(text),
                    InvoiceDate = ExtractInvoiceDate(text),
                    PurchaserName = ExtractPurchaserName(text),
                    PurchaserRegisterNum = ExtractPurchaserRegisterNum(text),
                    SellerName = ExtractSellerName(text),
                    SellerRegisterNum = ExtractSellerRegisterNum(text),
                    TotalAmount = ExtractTotalAmount(text),
                    TotalTax = ExtractTotalTax(text),
                    AmountInFiguers = ExtractAmountInFiguers(text),
                    InvoiceType = ExtractInvoiceType(text)
                }
            };

            // 使用System.Text.Json序列化为JSON字符串
            // 注意：保持PascalCase命名以匹配百度API返回格式
            return JsonSerializer.Serialize(result, new JsonSerializerOptions 
            { 
                PropertyNamingPolicy = null, // 保持原始属性名（PascalCase）
                WriteIndented = false
            });
        }

        /// <summary>
        /// 提取发票号码
        /// </summary>
        private string ExtractInvoiceNum(string text)
        {
            // 根据实际格式：发票号码:25447000001684857775（可能是发票号码和代码连在一起）
            // 发票号码通常是8位或12位，在"发票号码:"后面
            var patterns = new[]
            {
                @"发票号码[：:]([0-9]{8,12})", // 匹配8-12位数字
                @"发票号码[：:]([0-9]+)" // 如果上面匹配不到，匹配所有数字
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                if (match.Success && match.Groups.Count > 1)
                {
                    string value = match.Groups[1].Value.Trim();
                    // 如果数字很长（可能是号码和代码连在一起），取前8-12位作为发票号码
                    if (value.Length > 12)
                    {
                        // 通常发票号码是8位或12位，如果超过12位，可能是号码+代码
                        // 尝试取最后8-12位作为发票号码
                        if (value.Length >= 20)
                        {
                            // 如果总长度>=20，可能是12位号码+10位代码，取最后12位
                            value = value.Substring(value.Length - 12);
                        }
                        else
                        {
                            // 否则取最后8位
                            value = value.Substring(value.Length - 8);
                        }
                    }
                    if (!string.IsNullOrEmpty(value))
                        return value;
                }
            }

            return "";
        }

        /// <summary>
        /// 提取发票代码
        /// </summary>
        private string ExtractInvoiceCode(string text)
        {
            // 根据实际格式，发票代码可能在"发票号码:"后面的长数字串中
            // 先尝试直接匹配"发票代码:"
            var patterns = new[]
            {
                @"发票代码[：:](\d{10,12})",
                @"发票代码[：:](\d+)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                if (match.Success && match.Groups.Count > 1)
                {
                    string value = match.Groups[1].Value.Trim();
                    if (!string.IsNullOrEmpty(value) && value.Length >= 10)
                        return value;
                }
            }

            // 如果没有单独的发票代码，尝试从"发票号码:"后面的长数字中提取
            // 格式可能是：发票号码:25447000001684857775（12位号码+10位代码）
            var invoiceNumMatch = Regex.Match(text, @"发票号码[：:]([0-9]+)", RegexOptions.IgnoreCase);
            if (invoiceNumMatch.Success && invoiceNumMatch.Groups.Count > 1)
            {
                string fullNumber = invoiceNumMatch.Groups[1].Value.Trim();
                // 如果总长度是20位或22位，可能是12位号码+10位代码
                if (fullNumber.Length == 20 || fullNumber.Length == 22)
                {
                    // 取前10-12位作为发票代码
                    return fullNumber.Substring(0, Math.Min(12, fullNumber.Length - 8));
                }
                // 如果总长度是18位，可能是8位号码+10位代码
                else if (fullNumber.Length == 18)
                {
                    return fullNumber.Substring(0, 10);
                }
            }

            return "";
        }

        /// <summary>
        /// 提取开票日期
        /// </summary>
        private string ExtractInvoiceDate(string text)
        {
            // 匹配日期格式：2025-08-06 或 2025年08月06日 或 2025/08/06
            var patterns = new[]
            {
                @"开票日期[：:\s]+(\d{4}[-年/]\d{1,2}[-月/]\d{1,2}[日]?)",
                @"开票日期[：:]\s*(\d{4}[-年/]\d{1,2}[-月/]\d{1,2}[日]?)",
                @"日期[：:\s]+(\d{4}[-年/]\d{1,2}[-月/]\d{1,2}[日]?)",
                @"(\d{4}年\d{1,2}月\d{1,2}日)",
                @"(\d{4}-\d{1,2}-\d{1,2})",
                @"(\d{4}/\d{1,2}/\d{1,2})"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                if (match.Success && match.Groups.Count > 1)
                {
                    string dateStr = match.Groups[1].Value.Trim();
                    // 统一格式化为 yyyy-MM-dd
                    dateStr = dateStr.Replace("年", "-").Replace("月", "-").Replace("日", "").Replace("/", "-");
                    // 确保日期格式正确（补零）
                    var dateParts = dateStr.Split('-');
                    if (dateParts.Length == 3)
                    {
                        if (dateParts[1].Length == 1) dateParts[1] = "0" + dateParts[1];
                        if (dateParts[2].Length == 1) dateParts[2] = "0" + dateParts[2];
                        return string.Join("-", dateParts);
                    }
                    return dateStr;
                }
            }

            return "";
        }

        /// <summary>
        /// 提取购买方名称
        /// </summary>
        private string ExtractPurchaserName(string text)
        {
            // 根据实际格式：购买方信息  91410104MAD3TRFC98名称:郑州琳之星通讯有限公司统一社会信用代码/纳税人识别号:销售方信息
            // 先找到"购买方信息"的位置，然后在这个范围内查找"名称:"
            int purchaserInfoIndex = text.IndexOf("购买方信息", StringComparison.OrdinalIgnoreCase);
            if (purchaserInfoIndex >= 0)
            {
                // 找到"销售方信息"的位置，作为结束边界
                int sellerInfoIndex = text.IndexOf("销售方信息", StringComparison.OrdinalIgnoreCase);
                int endIndex = sellerInfoIndex > purchaserInfoIndex ? sellerInfoIndex : text.Length;
                
                // 在购买方信息范围内查找
                string purchaserSection = text.Substring(purchaserInfoIndex, endIndex - purchaserInfoIndex);
                
                // 查找"名称:"后面的内容
                int nameIndex = purchaserSection.IndexOf("名称:", StringComparison.OrdinalIgnoreCase);
                if (nameIndex >= 0)
                {
                    int nameStart = nameIndex + "名称:".Length;
                    string namePart = purchaserSection.Substring(nameStart);
                    
                    // 找到"统一社会信用代码"或"销售方"的位置作为结束
                    int endNameIndex = namePart.IndexOf("统一社会信用代码", StringComparison.OrdinalIgnoreCase);
                    if (endNameIndex < 0)
                        endNameIndex = namePart.IndexOf("销售方", StringComparison.OrdinalIgnoreCase);
                    if (endNameIndex < 0)
                        endNameIndex = namePart.Length;
                    
                    string name = namePart.Substring(0, endNameIndex).Trim();
                    if (!string.IsNullOrEmpty(name) && name.Length > 1)
                    {
                        return name;
                    }
                }
            }

            return "";
        }

        /// <summary>
        /// 提取购买方纳税人识别号
        /// </summary>
        private string ExtractPurchaserRegisterNum(string text)
        {
            var patterns = new[]
            {
                @"购买方[：:]\s*纳税人识别号[：:]\s*([A-Z0-9]{15,20})",
                @"购买方[：:]\s*识别号[：:]\s*([A-Z0-9]{15,20})",
                @"纳税人识别号[：:]\s*([A-Z0-9]{15,20})"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success && match.Groups.Count > 1)
                {
                    return match.Groups[1].Value.Trim();
                }
            }

            return "";
        }

        /// <summary>
        /// 提取销售方名称
        /// </summary>
        private string ExtractSellerName(string text)
        {
            // 根据实际格式：销售方信息  914419000585344943名称:华为终端有限公司统一社会信用代码/纳税人识别号:项目名称
            // 先找到"销售方信息"的位置，然后在这个范围内查找"名称:"
            int sellerInfoIndex = text.IndexOf("销售方信息", StringComparison.OrdinalIgnoreCase);
            if (sellerInfoIndex >= 0)
            {
                // 找到"项目名称"的位置，作为结束边界
                int projectNameIndex = text.IndexOf("项目名称", sellerInfoIndex, StringComparison.OrdinalIgnoreCase);
                int endIndex = projectNameIndex > sellerInfoIndex ? projectNameIndex : text.Length;
                
                // 在销售方信息范围内查找
                string sellerSection = text.Substring(sellerInfoIndex, endIndex - sellerInfoIndex);
                
                // 查找"名称:"后面的内容
                int nameIndex = sellerSection.IndexOf("名称:", StringComparison.OrdinalIgnoreCase);
                if (nameIndex >= 0)
                {
                    int nameStart = nameIndex + "名称:".Length;
                    string namePart = sellerSection.Substring(nameStart);
                    
                    // 找到"统一社会信用代码"或"项目名称"的位置作为结束
                    int endNameIndex = namePart.IndexOf("统一社会信用代码", StringComparison.OrdinalIgnoreCase);
                    if (endNameIndex < 0)
                        endNameIndex = namePart.IndexOf("项目名称", StringComparison.OrdinalIgnoreCase);
                    if (endNameIndex < 0)
                        endNameIndex = namePart.Length;
                    
                    string name = namePart.Substring(0, endNameIndex).Trim();
                    if (!string.IsNullOrEmpty(name) && name.Length > 1)
                    {
                        return name;
                    }
                }
            }

            return "";
        }

        /// <summary>
        /// 提取销售方纳税人识别号
        /// </summary>
        private string ExtractSellerRegisterNum(string text)
        {
            var patterns = new[]
            {
                @"销售方[：:]\s*纳税人识别号[：:]\s*([A-Z0-9]{15,20})",
                @"销售方[：:]\s*识别号[：:]\s*([A-Z0-9]{15,20})"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase);
                if (match.Success && match.Groups.Count > 1)
                {
                    return match.Groups[1].Value.Trim();
                }
            }

            return "";
        }

        /// <summary>
        /// 提取金额合计（不含税）
        /// </summary>
        private string ExtractTotalAmount(string text)
        {
            // 根据实际格式：合        计¥167.26
            var patterns = new[]
            {
                @"合\s+计[¥￥]?([\d,]+\.?\d*)", // 匹配"合        计¥167.26"
                @"合计[¥￥]?([\d,]+\.?\d*)", // 匹配"合计¥167.26"
                @"金额合计[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)",
                @"不含税金额[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                if (match.Success && match.Groups.Count > 1)
                {
                    string value = match.Groups[1].Value.Trim();
                    if (!string.IsNullOrEmpty(value))
                        return value;
                }
            }

            return "";
        }

        /// <summary>
        /// 提取税额
        /// </summary>
        private string ExtractTotalTax(string text)
        {
            // 根据实际格式：合        计¥167.26¥21.74（税额在合计金额后面）
            // 先尝试匹配"合计"后面的第二个金额
            var patterns = new[]
            {
                @"合\s+计[¥￥]?[\d,]+\.?\d*[¥￥]?([\d,]+\.?\d*)", // 匹配"合        计¥167.26¥21.74"中的21.74
                @"合计[¥￥]?[\d,]+\.?\d*[¥￥]?([\d,]+\.?\d*)", // 匹配"合计¥167.26¥21.74"中的21.74
                @"税额[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)",
                @"合计税额[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)",
                @"增值税[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                if (match.Success && match.Groups.Count > 1)
                {
                    string value = match.Groups[1].Value.Trim();
                    if (!string.IsNullOrEmpty(value))
                        return value;
                }
            }

            return "";
        }

        /// <summary>
        /// 提取价税合计
        /// </summary>
        private string ExtractAmountInFiguers(string text)
        {
            // 根据实际格式：价税合计（大写）壹佰捌拾玖圆整（小写）¥189.00
            var patterns = new[]
            {
                @"价税合计[^¥]*[¥￥]?([\d,]+\.?\d*)", // 匹配"价税合计...¥189.00"
                @"价税合计[（(]大写[）)][^¥]*[（(]小写[）)][¥￥]?([\d,]+\.?\d*)", // 匹配完整格式
                @"小写[）)][¥￥]?([\d,]+\.?\d*)", // 匹配"小写）¥189.00"
                @"总计[：:\s]*[¥￥]?\s*([\d,]+\.?\d*)"
            };

            foreach (var pattern in patterns)
            {
                var match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
                if (match.Success && match.Groups.Count > 1)
                {
                    string value = match.Groups[1].Value.Trim();
                    if (!string.IsNullOrEmpty(value))
                        return value;
                }
            }

            return "";
        }

        /// <summary>
        /// 提取发票类型
        /// </summary>
        private string ExtractInvoiceType(string text)
        {
            if (text.Contains("增值税专用发票") || text.Contains("专用发票"))
            {
                return "增值税专用发票";
            }
            else if (text.Contains("增值税普通发票") || text.Contains("普通发票"))
            {
                return "增值税普通发票";
            }
            else if (text.Contains("电子发票"))
            {
                return "电子发票";
            }

            return "增值税发票";
        }
    }
}

