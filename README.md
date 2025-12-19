# InvoiceVision - 发票识别工具

一个基于百度OCR API的发票识别Windows桌面应用程序，支持批量识别发票图片和PDF文件，并将识别结果导出为Excel格式。

## 📋 功能特性

- ✅ **批量识别**：支持一次性选择多张发票图片或PDF文件进行识别
- ✅ **多格式支持**：支持 JPG、JPEG、PNG、BMP、GIF 图片格式以及 PDF 文件
- ✅ **智能识别**：使用百度OCR API识别发票关键信息
- ✅ **数据展示**：实时显示识别结果，包括发票号码、代码、日期、购买方、销售方、金额等信息
- ✅ **Excel导出**：一键导出所有识别结果到Excel文件，方便后续处理
- ✅ **进度显示**：实时显示识别进度，支持QPS控制（每秒最多2次请求）

## 🚀 快速开始

### 系统要求

- Windows 10 或更高版本
- .NET 8.0 Runtime（如果未安装，程序会自动提示）

### 安装方式

1. **下载发布版本**
   - 从 [Releases](https://github.com/your-repo/InvoiceVision/releases) 下载最新版本的 `InvoiceVision.exe`
   - 直接运行即可，无需安装

2. **从源码编译**
   ```bash
   # 克隆仓库
   git clone https://github.com/your-repo/InvoiceVision.git
   cd InvoiceVision
   
   # 使用 Visual Studio 或 .NET CLI 编译
   dotnet build --configuration Release
   
   # 发布单文件
   dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
   ```

### 配置百度OCR API

**首次使用前必须配置API密钥：**

1. **复制配置文件模板**
   - 将项目根目录下的 `appsettings.example.json` 复制为 `appsettings.json`

2. **编辑配置文件**
   - 使用文本编辑器打开 `appsettings.json`
   - 填入您的百度OCR API密钥：
   ```json
   {
     "BaiduOCR": {
       "ApiKey": "your_api_key_here",
       "SecretKey": "your_secret_key_here"
     }
   }
   ```

3. **保存配置文件**
   - 保存 `appsettings.json` 文件
   - **重要**：`appsettings.json` 已添加到 `.gitignore`，不会被提交到Git仓库

> **注意**：
> - 获取百度OCR API密钥请访问 [百度智能云](https://cloud.baidu.com/)
> - 如果未配置API密钥，程序启动时会提示配置错误
> - 配置文件格式错误时，程序会显示相应的错误提示

## 📖 使用说明

### 基本流程

1. **选择文件**
   - 点击"选择图片/PDF"按钮
   - 在文件选择对话框中选择一张或多张发票图片或PDF文件
   - 支持多选（按住 Ctrl 或 Shift 键）

2. **开始识别**
   - 点击"开始识别"按钮
   - 程序会自动处理所有选中的文件
   - 识别过程中会显示进度条和当前处理状态

3. **查看结果**
   - 识别完成后，结果会显示在列表中
   - 可以查看每张发票的详细信息：
     - 发票号码
     - 发票代码
     - 开票日期
     - 购买方名称
     - 销售方名称
     - 金额合计
     - 税额
     - 价税合计
     - 发票类型

4. **导出Excel**
   - 点击"导出Excel"按钮
   - 选择保存位置和文件名
   - 导出的Excel文件包含所有识别结果，方便后续数据分析

### 识别字段说明

| 字段名称 | 说明 |
|---------|------|
| 发票号码 | 发票的唯一编号 |
| 发票代码 | 发票的代码标识 |
| 开票日期 | 发票开具的日期 |
| 购买方名称 | 购买方的企业或个人名称 |
| 购买方税号 | 购买方的纳税人识别号 |
| 销售方名称 | 销售方的企业或个人名称 |
| 销售方税号 | 销售方的纳税人识别号 |
| 金额合计 | 不含税的金额合计 |
| 税额 | 增值税税额 |
| 价税合计 | 含税总金额 |
| 发票类型 | 发票的类型（如增值税专用发票、普通发票等） |

## 🛠️ 技术栈

- **开发语言**：C# (.NET 8.0)
- **UI框架**：Windows Forms
- **OCR服务**：百度智能云 OCR API
- **Excel处理**：EPPlus 7.5.2
- **HTTP请求**：WodToolkit 1.0.1.4

## 📦 项目结构

```
InvoiceVision/
├── Form1.cs              # 主窗体逻辑
├── Form1.Designer.cs     # 窗体设计器代码
├── BaiDu.cs              # 百度OCR API封装
├── SuperListView.cs      # 自定义列表视图控件
├── Program.cs            # 程序入口
└── InvoiceVision.csproj  # 项目配置文件
```

## ⚙️ 配置说明

### API请求限制

程序已内置QPS控制机制，每次API请求之间至少间隔500毫秒，确保不超过百度API的2 QPS限制。

### 许可证

EPPlus使用非商业许可证（NonCommercial License）。

## 🐛 常见问题

**Q: 识别失败怎么办？**  
A: 请检查：
- 图片是否清晰
- 网络连接是否正常
- API密钥是否有效
- 文件格式是否支持

**Q: 支持哪些发票类型？**  
A: 支持增值税发票（专用发票、普通发票）等百度OCR API支持的发票类型。

**Q: 可以离线使用吗？**  
A: 不可以，本程序需要联网调用百度OCR API进行识别。

**Q: 识别速度慢怎么办？**  
A: 识别速度受网络状况和API响应时间影响。程序已优化请求频率，避免超过API限制。

## 📝 更新日志

### v1.0.0
- ✅ 初始版本发布
- ✅ 支持批量识别发票图片和PDF
- ✅ 支持Excel导出功能
- ✅ 实现QPS控制机制

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

本项目采用 MIT 许可证。

## 👤 作者

InvoiceVision 开发团队

---

**注意**：使用本软件需要有效的百度OCR API密钥。请确保遵守百度智能云的服务条款和使用限制。

