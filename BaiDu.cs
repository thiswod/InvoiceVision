using System;
using System.IO;
using System.Text.Json;
using WodToolKit.src.UmiOCR.Ocr;
using WodToolKit.src.UmiOCR.Doc;

namespace InvoiceVision
{
    public class BaiDu
    {
        private OCR ocr;
        private Doc doc;

        public BaiDu(string ApiKey = "", string SecretKey = "")
        {
            // UmiOCR 不需要 API 密钥，可以直接使用
            ocr = new OCR();
            doc = new Doc();
        }

        public string vat_invoice(string imagePath, string imageType = "png")
        {
            try
            {
                // 判断是 PDF 还是图片
                if (imageType.ToLower() == "pdf")
                {
                    // 使用 Doc 类的 Recognize 方法处理 PDF 文件
                    // Recognize 方法返回的结果包含 TextContent（识别文本内容）
                    var result = doc.Recognize(imagePath);
                    return result.TextContent ?? "";
                }
                else
                {
                    // 使用 OCR 类的 Ocr 方法处理图片文件
                    // Ocr 方法返回保存的 JSON 文件路径
                    string resultPath = ocr.Ocr(imagePath);
                    
                    // 读取 JSON 文件并提取文本内容
                    if (!string.IsNullOrEmpty(resultPath) && File.Exists(resultPath))
                    {
                        string jsonContent = File.ReadAllText(resultPath);
                        
                        // 解析 JSON 提取文本
                        try
                        {
                            using (JsonDocument doc = JsonDocument.Parse(jsonContent))
                            {
                                var root = doc.RootElement;
                                
                                // UmiOCR JSON 格式：{"code": 100, "data": "文本内容", ...}
                                if (root.TryGetProperty("data", out var data) && data.ValueKind == JsonValueKind.String)
                                {
                                    // data 字段是字符串，包含识别出的文本（可能有 Unicode 转义）
                                    string text = data.GetString() ?? "";
                                    // JSON 解析器会自动处理 Unicode 转义，所以直接返回即可
                                    return text;
                                }
                                
                                // 如果格式不匹配，返回原始 JSON（用于调试）
                                return jsonContent;
                            }
                        }
                        catch (JsonException)
                        {
                            // 如果 JSON 解析失败，直接返回文件内容
                            return jsonContent;
                        }
                    }
                    
                    return "";
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"OCR 识别失败: {ex.Message}", ex);
            }
        }
    }
}
