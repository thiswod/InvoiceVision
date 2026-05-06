# InvoiceVision

发票识别工具，基于 WinForms 开发。

当前版本支持三类输入：
- 图片文件：`jpg` `jpeg` `png` `bmp` `gif`
- PDF 文件：`pdf`
- 压缩包：`zip`

当前识别方式：
- PDF 发票走本地解析，不依赖百度 API
- 图片发票仍通过百度 OCR 识别
- ZIP 会先解压，再自动读取其中的 PDF 和图片文件

## 功能

- 批量导入图片、PDF、ZIP
- 本地提取 PDF 发票关键信息
- 百度 OCR 识别图片发票
- 列表展示识别结果
- 导出 Excel
- 双击左侧文件列表直接打开文件

## 运行环境

- Windows 10 或更高版本
- .NET 8

## 项目依赖

- `EPPlus 7.5.2`
- `Microsoft.Extensions.Configuration`
- `Microsoft.Extensions.Configuration.Json`
- `UglyToad.PdfPig`
- `WodToolKit`

## 构建

```powershell
dotnet build
```

发布示例：

```powershell
dotnet publish -c Release -r win-x64 --self-contained true
```

## 配置说明

程序会读取根目录下的 `appsettings.json`。

如果你只处理 PDF，不需要配置百度 OCR。

如果你要处理图片，需要配置百度 OCR：

```json
{
  "BaiduOCR": {
    "ApiKey": "your_api_key",
    "SecretKey": "your_secret_key"
  }
}
```

项目中已提供示例文件：

- `appsettings.example.json`

## 使用方式

1. 点击“选择文件”
2. 选择图片、PDF，或者 ZIP 压缩包
3. 点击“开始识别”
4. 查看右侧识别结果
5. 点击“导出Excel”导出结果

补充说明：
- 导入 ZIP 时，程序会先解压到临时目录
- 程序关闭或重新导入时，会自动清理临时目录
- 双击左侧文件列表可用系统默认程序打开文件

## 当前 PDF 提取字段

- 发票号码
- 发票代码
- 开票日期
- 购买方名称
- 购买方税号
- 销售方名称
- 销售方税号
- 金额合计
- 税额
- 价税合计

## 项目结构

```text
InvoiceVision/
├─ Form1.cs
├─ Form1.Designer.cs
├─ LocalPdfInvoiceExtractor.cs
├─ BaiDu.cs
├─ InvoiceData.cs
├─ Program.cs
├─ SuperListView.cs
├─ InvoiceVision.csproj
├─ README.md
└─ 发票提取/
```

## 注意事项

- 当前图片识别仍依赖网络和百度 OCR
- 当前 PDF 本地提取规则是按现有发票样本校准的
- 如果遇到新的票样格式，可能还需要继续补规则

## 已实现的改动

- PDF 改为本地提取
- 支持 ZIP 导入
- 支持双击打开文件
- 优化了 PDF 金额与购销方信息提取
