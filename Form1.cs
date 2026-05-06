using System;
using System.Collections.Generic;
<<<<<<< HEAD
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
=======
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
>>>>>>> main
using System.Windows.Forms;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using WodToolkit.Json;

namespace InvoiceVision
{
    public partial class Form1 : Form
    {
<<<<<<< HEAD
        private BaiDu? baiDu;
        private List<InvoiceData> invoiceResults = new List<InvoiceData>();
        private string? apiKey;
        private string? secretKey;
=======
        private readonly LocalPdfInvoiceExtractor localPdfInvoiceExtractor = new();
        private readonly List<string> extractedTempDirectories = new();
        private readonly List<InvoiceData> invoiceResults = new();
        private BaiDu? baiDu;
>>>>>>> main

        public Form1()
        {
            InitializeComponent();
            LoadConfiguration();
<<<<<<< HEAD
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
=======
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
>>>>>>> main
        }

        private void LoadConfiguration()
        {
            try
            {
                var builder = new ConfigurationBuilder()
<<<<<<< HEAD
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

                var configuration = builder.Build();
                apiKey = configuration["BaiduOCR:ApiKey"] ?? "";
                secretKey = configuration["BaiduOCR:SecretKey"] ?? "";

                if (string.IsNullOrWhiteSpace(apiKey) || string.IsNullOrWhiteSpace(secretKey))
                {
                    MessageBox.Show(
                        "ËØ∑Âú® appsettings.json Êñá‰ª∂‰∏≠ÈÖçÁΩÆÁôæÂ∫¶OCR APIÂØÜÈí•ÔºÅ\n\n" +
                        "ËØ∑ÂèÇËÄ?appsettings.example.json Êñá‰ª∂Ê†ºÂºèËøõË°åÈÖçÁΩÆ„Ä?,
                        "ÈÖçÁΩÆÈîôËØØ",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                baiDu = new BaiDu(apiKey, secretKey);
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show(
                    "Êú™ÊâæÂà?appsettings.json ÈÖçÁΩÆÊñá‰ª∂ÔºÅ\n\n" +
                    "ËØ∑Â§çÂà?appsettings.example.json ‰∏?appsettings.json Âπ∂ÈÖçÁΩÆÊÇ®ÁöÑAPIÂØÜÈí•„Ä?,
                    "ÈÖçÁΩÆÊñá‰ª∂Áº∫Â§±",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Âä†ËΩΩÈÖçÁΩÆÊñá‰ª∂Êó∂Âá∫ÈîôÔºö{ex.Message}\n\n" +
                    "ËØ∑Ê£ÄÊü?appsettings.json Êñá‰ª∂Ê†ºÂºèÊòØÂê¶Ê≠£Á°Æ„Ä?,
                    "ÈÖçÁΩÆÈîôËØØ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
=======
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
>>>>>>> main
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

<<<<<<< HEAD
        private void BtnSelectImages_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "ÂõæÁâáÂíåPDFÊñá‰ª∂|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.pdf|ÂõæÁâáÊñá‰ª∂|*.jpg;*.jpeg;*.png;*.bmp;*.gif|PDFÊñá‰ª∂|*.pdf|ÊâÄÊúâÊñá‰ª∂|*.*";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "ÈÄâÊã©ÂèëÁ•®ÂõæÁâáÊàñPDFÊñá‰ª∂";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    listBoxImages.Items.Clear();
                    foreach (string fileName in openFileDialog.FileNames)
                    {
                        listBoxImages.Items.Add(fileName);
                    }
                    btnStart.Enabled = listBoxImages.Items.Count > 0;
                }
=======
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            CleanupExtractedTempDirectories();
            base.OnFormClosed(e);
        }

        private void BtnSelectImages_Click(object sender, EventArgs e)
        {
            using OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "ÂèëÁ•®Êñá‰ª∂ÂíåÂéãÁº©ÂåÖ|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.pdf;*.zip|ÂõæÁâáÊñá‰ª∂|*.jpg;*.jpeg;*.png;*.bmp;*.gif|PDFÊñá‰ª∂|*.pdf|ÂéãÁº©ÂåÖ|*.zip|ÊâÄÊúâÊñá‰ª∂|*.*";
            openFileDialog.Multiselect = true;
            openFileDialog.Title = "ÈÄâÊã©ÂèëÁ•®ÂõæÁâá„ÄÅPDFÊàñÂéãÁº©ÂåÖ";

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
            labelStatus.Text = $"Â∑≤ËΩΩÂÖ• {listBoxImages.Items.Count} ‰∏™Êñá‰ª∂";

            if (listBoxImages.Items.Count == 0)
            {
                MessageBox.Show("Ê≤°ÊúâÊâæÂà∞ÂèØËØÜÂà´ÁöÑ PDF ÊàñÂõæÁâáÊñá‰ª∂„ÄÇ", "ÊèêÁ§∫", MessageBoxButtons.OK, MessageBoxIcon.Warning);
>>>>>>> main
            }
        }

        private void BtnStart_Click(object sender, EventArgs e)
        {
<<<<<<< HEAD
            if (baiDu == null)
            {
                MessageBox.Show(
                    "APIÂØÜÈí•Êú™ÈÖçÁΩÆÔºÅ\n\n" +
                    "ËØ∑ÈÖçÁΩ?appsettings.json Êñá‰ª∂‰∏≠ÁöÑÁôæÂ∫¶OCR APIÂØÜÈí•„Ä?,
                    "ÈÖçÁΩÆÈîôËØØ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }

            if (listBoxImages.Items.Count == 0)
            {
                MessageBox.Show("ËØ∑ÂÖàÈÄâÊã©ÂõæÁâáÊàñPDFÊñá‰ª∂Ôº?, "ÊèêÁ§∫", MessageBoxButtons.OK, MessageBoxIcon.Warning);
=======
            if (listBoxImages.Items.Count == 0)
            {
                MessageBox.Show("ËØ∑ÂÖàÈÄâÊã©ÂõæÁâá„ÄÅPDFÊàñÂéãÁº©ÂåÖ„ÄÇ", "ÊèêÁ§∫", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            bool containsImageFiles = listBoxImages.Items
                .Cast<string>()
                .Any(path => !IsPdfFile(path));

            if (containsImageFiles && baiDu == null)
            {
                MessageBox.Show(
                    "ÂΩìÂâçÈÄâÊã©‰∏≠ÂåÖÂê´ÂõæÁâáÊñá‰ª∂Ôºå‰ΩÜÊú™ÊâæÂà∞ÂèØÁî®ÁöÑÁôæÂ∫¶ OCR ÈÖçÁΩÆ„ÄÇ\n\nPDF Â∑≤Êîπ‰∏∫Êú¨Âú∞Ëß£ÊûêÔºõÂ¶ÇÊûúËøòË¶ÅËØÜÂà´ÂõæÁâáÔºåËØ∑Âú® appsettings.json ‰∏≠ÈÖçÁΩÆÁôæÂ∫¶ OCR„ÄÇ",
                    "Áº∫Â∞ëÂõæÁâáËØÜÂà´ÈÖçÁΩÆ",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
>>>>>>> main
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
<<<<<<< HEAD
                ProcessImages();
                labelStatus.Text = $"ËØÜÂà´ÂÆåÊàêÔºåÂÖ±ËØÜÂà´ {invoiceResults.Count} Âº†ÂèëÁ•?;
=======
                ProcessFiles();
                labelStatus.Text = $"Â§ÑÁêÜÂÆåÊàêÔºåÂÖ±ËØÜÂà´ {invoiceResults.Count} Êù°ËÆ∞ÂΩï";
>>>>>>> main
                btnExport.Enabled = invoiceResults.Count > 0;
            }
            catch (Exception ex)
            {
<<<<<<< HEAD
                MessageBox.Show($"ËØÜÂà´ËøáÁ®ã‰∏≠Âá∫ÈîôÔºö{ex.Message}", "ÈîôËØØ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelStatus.Text = "ËØÜÂà´Â§±Ë¥•";
=======
                MessageBox.Show($"Â§ÑÁêÜËøáÁ®ã‰∏≠Âá∫ÈîôÔºö{ex.Message}", "ÈîôËØØ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelStatus.Text = "Â§ÑÁêÜÂ§±Ë¥•";
>>>>>>> main
            }
            finally
            {
                btnStart.Enabled = true;
                btnSelectImages.Enabled = true;
                progressBar.Visible = false;
            }
        }

<<<<<<< HEAD
        private void ProcessImages()
        {
            int minDelayMs = 500; // ÊúÄÂ∞èÈó¥Èö?00msÔºåÁ°Æ‰øù‰∏çË∂ÖËøá2 QPS
            int processedCount = 0;

            foreach (string imagePath in listBoxImages.Items.Cast<string>())
            {
                try
                {
                    // ÊéßÂà∂QPSÔºöÊØèÊ¨°ËØ∑Ê±Ç‰πãÈó¥Ëá≥Â∞ëÈó¥Èö?00ms
                    if (processedCount > 0)
                    {
                        System.Threading.Thread.Sleep(minDelayMs);
                    }

                    ProcessSingleImage(imagePath);
                    processedCount++;
                    progressBar.Value = processedCount;
                    labelStatus.Text = $"Ê≠£Âú®ËØÜÂà´... ({processedCount}/{listBoxImages.Items.Count})";
                    Application.DoEvents(); // Êõ¥Êñ∞UI
                }
                catch (Exception ex)
                {
                    labelStatus.Text = $"Â§ÑÁêÜ {Path.GetFileName(imagePath)} Êó∂Âá∫Èî? {ex.Message}";
                    processedCount++;
                    progressBar.Value = processedCount;
                    Application.DoEvents(); // Êõ¥Êñ∞UI
=======
        private void ListBoxImages_DoubleClick(object sender, EventArgs e)
        {
            if (listBoxImages.SelectedItem is not string filePath || string.IsNullOrWhiteSpace(filePath))
            {
                return;
            }

            if (!File.Exists(filePath))
            {
                MessageBox.Show($"Êñá‰ª∂‰∏çÂ≠òÂú®Ôºö{filePath}", "ÊèêÁ§∫", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
                MessageBox.Show($"ÊâìÂºÄÊñá‰ª∂Â§±Ë¥•Ôºö{ex.Message}", "ÈîôËØØ", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                    labelStatus.Text = $"Ê≠£Âú®Â§ÑÁêÜ... ({processedCount}/{listBoxImages.Items.Count})";
                    Application.DoEvents();
                }
                catch (Exception ex)
                {
                    processedCount++;
                    progressBar.Value = processedCount;
                    labelStatus.Text = $"Â§ÑÁêÜ {Path.GetFileName(filePath)} Âá∫ÈîôÔºö{ex.Message}";
                    Application.DoEvents();
>>>>>>> main
                }
            }
        }

<<<<<<< HEAD
        private void ProcessSingleImage(string imagePath)
        {
            try
            {
                // ËØªÂèñÊñá‰ª∂Âπ∂ËΩ¨Êç¢‰∏∫base64
                byte[] fileBytes = File.ReadAllBytes(imagePath);
                string base64Data = Convert.ToBase64String(fileBytes);

                // Ëé∑ÂèñÊñá‰ª∂Á±ªÂûãÔºàÊ†πÊçÆÊñá‰ª∂Êâ©Â±ïÂêçÔº?
                string fileType = "png"; // ÈªòËÆ§
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

                // Ë∞ÉÁî®APIËØÜÂà´
                if (baiDu == null)
                {
                    throw new InvalidOperationException("APIÂØÜÈí•Êú™ÈÖçÁΩÆÔºåÊó†Ê≥ïËøõË°åËØÜÂà´„Ä?);
                }
                
                // º«¬ºµ˜ ‘–≈œ¢µΩŒƒº˛
                System.Text.StringBuilder logBuilder = new System.Text.StringBuilder();
                logBuilder.AppendLine($"[{DateTime.Now}] ø™ ºµ˜”√API...");
                
                string resultJson = baiDu.vat_invoice(base64Data, fileType);
                logBuilder.AppendLine($"[{DateTime.Now}] APIµ˜”√ÕÍ≥…");
                
                //  ‰≥ˆAPI∑µªÿΩ·π˚µƒ«∞500∏ˆ◊÷∑˚£¨“‘±„¡ÀΩ‚∆‰Ω·ππ
                logBuilder.AppendLine($"[{DateTime.Now}] API∑µªÿΩ·π˚«∞500∏ˆ◊÷∑˚: {resultJson.Substring(0, Math.Min(500, resultJson.Length))}");

                // ±£¥ÊAPI∑µªÿµƒΩ·π˚µΩŒƒº˛£¨“‘±„≤Èø¥∆‰Ω·ππ
                try
                {
                    string fileName = Path.GetFileNameWithoutExtension(imagePath);
                    string outputPath = $"api_result_{fileName}.json";
                    System.IO.File.WriteAllText(outputPath, resultJson, System.Text.Encoding.UTF8);
                    logBuilder.AppendLine($"[{DateTime.Now}] API∑µªÿΩ·π˚“—±£¥ÊµΩ {outputPath} Œƒº˛");
                }
                catch (Exception ex)
                {
                    logBuilder.AppendLine($"[{DateTime.Now}] ±£¥ÊAPIΩ·π˚ ±≥ˆ¥Ì: {ex.Message}");
                    logBuilder.AppendLine($"[{DateTime.Now}] ¥ÌŒÛ∂—’ª: {ex.StackTrace}");
                }

                // Ω‚ŒˆJSONΩ·π˚
                logBuilder.AppendLine($"[{DateTime.Now}] ø™ ºΩ‚ŒˆJSONΩ·π˚...");
                dynamic result = EasyJson.ParseJsonToDynamic(resultJson);
                logBuilder.AppendLine($"[{DateTime.Now}] JSONΩ·π˚Ω‚ŒˆÕÍ≥…");
                
                // ±£¥Êµ˜ ‘»’÷æµΩŒƒº˛
                try
                {
                    string logPath = "debug_log.txt";
                    System.IO.File.AppendAllText(logPath, logBuilder.ToString(), System.Text.Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    // ∫ˆ¬‘±£¥Ê»’÷æ ±µƒ¥ÌŒÛ
                }
                
                // ºÏ≤È «∑Ò”–words_result◊÷∂Œ£¨”–‘Ú±Ì æ ∂±≥…π¶
                if (result.words_result != null)
                {
                    var invoiceData = ParseInvoiceData(result.words_result, imagePath);
                    invoiceResults.Add(invoiceData);
                    AddInvoiceToListView(invoiceData);
                }
                else
                {
                    // »Áπ˚√ª”–words_result£¨ø…ƒ‹ «≥ˆ¥Ì¡À£¨≥¢ ‘ªÒ»°¥ÌŒÛ–≈œ¢
                    string errorMsg = " ∂±Ω·π˚Œ™ø’";
                    try
                    {
                        if (result.error_code != null)
                        {
                            errorMsg = $"API∑µªÿ¥ÌŒÛ: {result.error_msg ?? "Œ¥÷™¥ÌŒÛ"} (¥ÌŒÛ¬Î: {result.error_code})";
                        }
                    }
                    catch
                    {
                        // »Áπ˚Œﬁ∑®ªÒ»°¥ÌŒÛ–≈œ¢£¨ π”√ƒ¨»œœ˚œ¢
                    }
                    throw new Exception(errorMsg);
                }
            }
            catch (Exception ex)
            {
                string fileType = Path.GetExtension(imagePath).ToLower() == ".pdf" ? "PDFŒƒº˛" : "Õº∆¨";
                throw new Exception($"¥¶¿Ì{fileType} {Path.GetFileName(imagePath)}  ±≥ˆ¥Ì: {ex.Message}", ex);
            }
=======
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
                throw new InvalidOperationException("Êú™ÈÖçÁΩÆÁôæÂ∫¶ OCRÔºåÊó†Ê≥ïÂ§ÑÁêÜÂõæÁâáÊñá‰ª∂„ÄÇ");
            }

            byte[] fileBytes = File.ReadAllBytes(imagePath);
            string base64Data = Convert.ToBase64String(fileBytes);
            string fileType = GetFileType(imagePath);
            string resultJson = baiDu.vat_invoice(base64Data, fileType);
            dynamic result = EasyJson.ParseJsonToDynamic(resultJson);

            if (result.words_result == null)
            {
                string errorMsg = "ËØÜÂà´ÁªìÊûú‰∏∫Á©∫";

                try
                {
                    if (result.error_code != null)
                    {
                        errorMsg = $"API ËøîÂõûÈîôËØØÔºö{result.error_msg ?? "Êú™Áü•ÈîôËØØ"} (ÈîôËØØÁ†Å: {result.error_code})";
                    }
                }
                catch
                {
                }

                throw new Exception(errorMsg);
            }

            return ParseInvoiceData(result.words_result, imagePath);
>>>>>>> main
        }

        private InvoiceData ParseInvoiceData(dynamic wordsResult, string imagePath)
        {
            string invoiceNum = GetStringValue(wordsResult.InvoiceNum);
            string invoiceCode = GetStringValue(wordsResult.InvoiceCode);
<<<<<<< HEAD
            
            // »Áπ˚∑¢∆±¥˙¬ÎŒ™ø’£¨ π”√InvoiceNum◊˜Œ™∑¢∆±¥˙¬Î
            // ∏˘æ›”√ªß∑¥¿°£¨InvoiceNum µº …œæÕ «∑¢∆±¥˙¬Î
=======

>>>>>>> main
            if (string.IsNullOrEmpty(invoiceCode) && !string.IsNullOrEmpty(invoiceNum))
            {
                invoiceCode = invoiceNum;
            }

<<<<<<< HEAD
            var invoice = new InvoiceData
=======
            return new InvoiceData
>>>>>>> main
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
<<<<<<< HEAD

            // Ã·»°…Ã∆∑√˜œ∏–≈œ¢
            try
            {
                // º«¬ºµ˜ ‘–≈œ¢
                System.Text.StringBuilder logBuilder = new System.Text.StringBuilder();
                logBuilder.AppendLine($"[{DateTime.Now}] ø™ ºÃ·»°…Ã∆∑√˜œ∏–≈œ¢");

                // ∏˘æ›API∑µªÿµƒ µº Ω·ππÃ·»°…Ã∆∑√˜œ∏–≈œ¢
                // …Ã∆∑–≈œ¢∑÷…¢‘⁄≤ªÕ¨µƒ◊÷∂Œ÷–£¨∂º « ˝◊È–Œ Ω
                var commodityNames = GetArrayValue(wordsResult.CommodityName);
                var commodityUnits = GetArrayValue(wordsResult.CommodityUnit);
                var commodityNums = GetArrayValue(wordsResult.CommodityNum);
                var commodityPrices = GetArrayValue(wordsResult.CommodityPrice);
                var commodityAmounts = GetArrayValue(wordsResult.CommodityAmount);
                var commodityTaxRates = GetArrayValue(wordsResult.CommodityTaxRate);
                var commodityTaxes = GetArrayValue(wordsResult.CommodityTax);

                // º∆À„…Ã∆∑ ˝¡ø£¨»°À˘”– ˝◊È÷–≥§∂»◊Ó¥Ûµƒƒ«∏ˆ
                int itemCount = Math.Max(
                    Math.Max(Math.Max(commodityNames.Length, commodityUnits.Length), 
                    Math.Max(commodityNums.Length, commodityPrices.Length)),
                    Math.Max(Math.Max(commodityAmounts.Length, commodityTaxRates.Length), 
                    commodityTaxes.Length)
                );

                logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑ ˝¡ø: {itemCount}");
                logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑√˚≥∆ ˝¡ø: {commodityNames.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑µ•Œª ˝¡ø: {commodityUnits.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑ ˝¡ø ˝¡ø: {commodityNums.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑µ•º€ ˝¡ø: {commodityPrices.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑Ω∂Ó ˝¡ø: {commodityAmounts.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑À∞¬  ˝¡ø: {commodityTaxRates.Length}");
                logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑À∞∂Ó ˝¡ø: {commodityTaxes.Length}");

                // Ã·»°…Ã∆∑√˜œ∏–≈œ¢
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
                    
                    // ≥¢ ‘¥”…Ã∆∑√˚≥∆÷–Ã·»°πÊ∏Ò–Õ∫≈
                    if (string.IsNullOrEmpty(commodityItem.Specification) && !string.IsNullOrEmpty(commodityItem.Name))
                    {
                        // ºÚµ•µƒπÊ‘Ú£∫»Áπ˚…Ã∆∑√˚≥∆∞¸∫¨ø’∏Ò£¨≥¢ ‘Ω´◊Ó∫Û“ª≤ø∑÷◊˜Œ™πÊ∏Ò–Õ∫≈
                        string[] parts = commodityItem.Name.Split(' ');
                        if (parts.Length > 1)
                        {
                            commodityItem.Specification = parts[parts.Length - 1];
                        }
                    }
                    
                    invoice.CommodityItems.Add(commodityItem);
                    logBuilder.AppendLine($"[{DateTime.Now}] ÃÌº”…Ã∆∑: {commodityItem.Name}");
                    logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑µ•Œª: {commodityItem.Unit}");
                    logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑ ˝¡ø: {commodityItem.Quantity}");
                    logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑µ•º€: {commodityItem.Price}");
                    logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑Ω∂Ó: {commodityItem.Amount}");
                    logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑À∞¬ : {commodityItem.TaxRate}");
                    logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑À∞∂Ó: {commodityItem.Tax}");
                }

                logBuilder.AppendLine($"[{DateTime.Now}] …Ã∆∑√˜œ∏Ã·»°ÕÍ≥…£¨π≤ {invoice.CommodityItems.Count} ∏ˆ…Ã∆∑");
                
                // ±£¥Êµ˜ ‘»’÷æ
                try
                {
                    System.IO.File.AppendAllText("parse_log.txt", logBuilder.ToString(), System.Text.Encoding.UTF8);
                }
                catch (Exception ex)
                {
                    // ∫ˆ¬‘±£¥Ê»’÷æ ±µƒ¥ÌŒÛ
                }
            }
            catch (Exception ex)
            {
                // …Ã∆∑√˜œ∏Ω‚Œˆ ß∞‹£¨º«¬º¥ÌŒÛµ´≤ª”∞œÏ’˚ÃÂΩ‚Œˆ
                try
                {
                    System.Text.StringBuilder logBuilder = new System.Text.StringBuilder();
                    logBuilder.AppendLine($"[{DateTime.Now}] Ω‚Œˆ…Ã∆∑√˜œ∏ ±≥ˆ¥Ì: {ex.Message}");
                    logBuilder.AppendLine($"[{DateTime.Now}] ¥ÌŒÛ∂—’ª: {ex.StackTrace}");
                    System.IO.File.AppendAllText("parse_error_log.txt", logBuilder.ToString(), System.Text.Encoding.UTF8);
                }
                catch
                {
                    // ∫ˆ¬‘±£¥Ê¥ÌŒÛ»’÷æ ±µƒ¥ÌŒÛ
                }
            }

            return invoice;
        }

        // ªÒ»° ˝◊È¿‡–Õµƒ÷µ£¨∑µªÿ◊÷∑˚¥Æ ˝◊È
        private string[] GetArrayValue(dynamic value)
        {
            try
            {
                if (value == null)
                    return new string[0];
                
                // ºÏ≤È «∑ÒŒ™ ˝◊È
                var enumerable = value as System.Collections.IEnumerable;
                if (enumerable != null)
                {
                    List<string> result = new List<string>();
                    foreach (var item in enumerable)
                    {
                        try
                        {
                            // ∂‘”⁄∂ØÃ¨∂‘œÛ£¨≥¢ ‘÷±Ω”∑√Œ word Ù–‘
                            if (item != null)
                            {
                                dynamic dynamicItem = item;
                                if (dynamicItem.word != null)
                                {
                                    result.Add(dynamicItem.word.ToString());
                                }
                                else
                                {
                                    // ≥¢ ‘÷±Ω”◊™ªªŒ™◊÷∑˚¥Æ
                                    result.Add(item.ToString());
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // º«¬º¥ÌŒÛ≤¢≥¢ ‘∆‰À˚∑Ω Ω
                            try
                            {
                                System.IO.File.AppendAllText(
                                    "array_value_error_log.txt", 
                                    $"[{DateTime.Now}] Ã·»° ˝◊È÷µ ±≥ˆ¥Ì: {ex.Message}\n", 
                                    System.Text.Encoding.UTF8
                                );
                            }
                            catch
                            {
                                // ∫ˆ¬‘¥ÌŒÛ
                            }
                            // ≥¢ ‘÷±Ω”◊™ªªŒ™◊÷∑˚¥Æ
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
                    // ≥¢ ‘÷±Ω”◊™ªªŒ™◊÷∑˚¥Æ
                    return new string[] { value.ToString() };
                }
            }
            catch (Exception ex)
            {
                // º«¬º¥ÌŒÛ
                try
                {
                    System.IO.File.AppendAllText(
                        "array_value_error_log.txt", 
                        $"[{DateTime.Now}] Ã·»° ˝◊È÷µ ±≥ˆ¥Ì: {ex.Message}\n", 
                        System.Text.Encoding.UTF8
                    );
                }
                catch
                {
                    // ∫ˆ¬‘¥ÌŒÛ
                }
                return new string[0];
            }
        }

        private string GetStringValue(dynamic value)
        {
            if (value == null) return "";
            return value.ToString();
=======
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
>>>>>>> main
        }

        private void AddInvoiceToListView(InvoiceData invoice)
        {
<<<<<<< HEAD
            if (invoice.CommodityItems.Count > 0)
            {
                // »Áπ˚”–…Ã∆∑√˜œ∏£¨Œ™√ø∏ˆ…Ã∆∑√˜œ∏¥¥Ω®“ª–– ˝æ›
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
                // »Áπ˚√ª”–…Ã∆∑√˜œ∏£¨¥¥Ω®“ª––ª˘±æ–≈œ¢
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
=======
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
>>>>>>> main
        }

        private void BtnExport_Click(object sender, EventArgs e)
        {
            if (invoiceResults.Count == 0)
            {
<<<<<<< HEAD
                MessageBox.Show("Ê≤°ÊúâÂèØÂØºÂá∫ÁöÑÊï∞ÊçÆÔº?, "ÊèêÁ§∫", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "ExcelÊñá‰ª∂|*.xlsx|ÊâÄÊúâÊñá‰ª∂|*.*";
                saveFileDialog.FileName = $"ÂèëÁ•®ËØÜÂà´ÁªìÊûú_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                saveFileDialog.Title = "‰øùÂ≠òExcelÊñá‰ª∂";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        ExportToExcel(saveFileDialog.FileName);
                        MessageBox.Show("ÂØºÂá∫ÊàêÂäüÔº?, "ÊèêÁ§∫", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"ÂØºÂá∫Â§±Ë¥•Ôºö{ex.Message}", "ÈîôËØØ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
=======
                MessageBox.Show("Ê≤°ÊúâÂèØÂØºÂá∫ÁöÑÊï∞ÊçÆ„ÄÇ", "ÊèêÁ§∫", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "ExcelÊñá‰ª∂|*.xlsx|ÊâÄÊúâÊñá‰ª∂|*.*";
            saveFileDialog.FileName = $"ÂèëÁ•®ËØÜÂà´ÁªìÊûú_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            saveFileDialog.Title = "‰øùÂ≠òExcelÊñá‰ª∂";

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            try
            {
                ExportToExcel(saveFileDialog.FileName);
                MessageBox.Show("ÂØºÂá∫ÊàêÂäü„ÄÇ", "ÊèêÁ§∫", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ÂØºÂá∫Â§±Ë¥•Ôºö{ex.Message}", "ÈîôËØØ", MessageBoxButtons.OK, MessageBoxIcon.Error);
>>>>>>> main
            }
        }

        private void ExportToExcel(string fileName)
        {
<<<<<<< HEAD
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("ÂèëÁ•®ËØÜÂà´ÁªìÊûú");

                // ËÆæÁΩÆË°®Â§¥
                worksheet.Cells[1, 1].Value = "ÂèëÁ•®Âè∑Á†Å";
                worksheet.Cells[1, 2].Value = "ÂèëÁ•®‰ª£Á†Å";
                worksheet.Cells[1, 3].Value = "ÂºÄÁ•®Êó•Êú?;
                worksheet.Cells[1, 4].Value = "Ë¥≠‰π∞ÊñπÂêçÁß?;
                worksheet.Cells[1, 5].Value = "Ë¥≠‰π∞ÊñπÁ®éÂè?;
                worksheet.Cells[1, 6].Value = "ÈîÄÂîÆÊñπÂêçÁß∞";
                worksheet.Cells[1, 7].Value = "ÈîÄÂîÆÊñπÁ®éÂè∑";
                worksheet.Cells[1, 8].Value = "ÂïÜÂìÅÂêçÁß∞";
                worksheet.Cells[1, 9].Value = "ËßÑÊ†ºÂûãÂè∑";
                worksheet.Cells[1, 10].Value = "Âçï‰Ωç";
                worksheet.Cells[1, 11].Value = "Êï∞Èáè";
                worksheet.Cells[1, 12].Value = "Âçï‰ª∑";
                worksheet.Cells[1, 13].Value = "ÈáëÈ¢ù";
                worksheet.Cells[1, 14].Value = "Á®éÁéá";
                worksheet.Cells[1, 15].Value = "Á®éÈ¢ù";
                worksheet.Cells[1, 16].Value = "ÈáëÈ¢ùÂêàËÆ°";
                worksheet.Cells[1, 17].Value = "Á®éÈ¢ùÂêàËÆ°";
                worksheet.Cells[1, 18].Value = "‰ª∑Á®éÂêàËÆ°";
                worksheet.Cells[1, 19].Value = "ÂèëÁ•®Á±ªÂûã";
                worksheet.Cells[1, 20].Value = "Êñá‰ª∂Ë∑ØÂæÑ";

                // ËÆæÁΩÆË°®Â§¥Ê†∑Âºè
                using (var range = worksheet.Cells[1, 1, 1, 20])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // Â°´ÂÖÖÊï∞ÊçÆ
                int row = 2;
                foreach (var invoice in invoiceResults)
                {
                    // ‰∏∫ÊØè‰∏™ÂèëÁ•®ÂàõÂª∫‰∏ÄË°åÔºåÊâÄÊúâÂïÜÂìÅÊòéÁªÜÈÉΩÂú®Ëøô‰∏ÄË°å‰∏≠
                    worksheet.Cells[row, 1].Value = invoice.InvoiceNum;
                    worksheet.Cells[row, 2].Value = invoice.InvoiceCode;
                    worksheet.Cells[row, 3].Value = invoice.InvoiceDate;
                    worksheet.Cells[row, 4].Value = invoice.PurchaserName;
                    worksheet.Cells[row, 5].Value = invoice.PurchaserRegisterNum;
                    worksheet.Cells[row, 6].Value = invoice.SellerName;
                    worksheet.Cells[row, 7].Value = invoice.SellerRegisterNum;
                    
                    if (invoice.CommodityItems.Count > 0)
                    {
                        //  ’ºØÀ˘”–…Ã∆∑√˜œ∏–≈œ¢£¨”√∑÷∫≈¡¨Ω”
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
                        
                        // Ω´ ’ºØµƒ–≈œ¢”√∑÷∫≈¡¨Ω”≤¢ÃÓ≥‰µΩµ•‘™∏Ò
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
                        // »Áπ˚√ª”–…Ã∆∑√˜œ∏£¨¡Ùø’
                        worksheet.Cells[row, 8].Value = "";
                        worksheet.Cells[row, 9].Value = "";
                        worksheet.Cells[row, 10].Value = "";
                        worksheet.Cells[row, 11].Value = "";
                        worksheet.Cells[row, 12].Value = "";
                        worksheet.Cells[row, 13].Value = "";
                        worksheet.Cells[row, 14].Value = "";
                        worksheet.Cells[row, 15].Value = "";
                    }
                    
                    // ÃÓ≥‰∑¢∆±µƒ∆‰À˚–≈œ¢
                    worksheet.Cells[row, 16].Value = invoice.TotalAmount;
                    worksheet.Cells[row, 17].Value = invoice.TotalTax;
                    worksheet.Cells[row, 18].Value = invoice.AmountInFiguers;
                    worksheet.Cells[row, 19].Value = invoice.InvoiceType;
                    worksheet.Cells[row, 20].Value = invoice.ImagePath;
                    
                    row++;
                }

                // Ëá™Âä®Ë∞ÉÊï¥ÂàóÂÆΩ
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // ‰øùÂ≠òÊñá‰ª∂
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
=======
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("ÂèëÁ•®ËØÜÂà´ÁªìÊûú");

            worksheet.Cells[1, 1].Value = "ÂèëÁ•®Âè∑Á†Å";
            worksheet.Cells[1, 2].Value = "ÂèëÁ•®‰ª£Á†Å";
            worksheet.Cells[1, 3].Value = "ÂºÄÁ•®Êó•Êúü";
            worksheet.Cells[1, 4].Value = "Ë¥≠‰π∞ÊñπÂêçÁß∞";
            worksheet.Cells[1, 5].Value = "Ë¥≠‰π∞ÊñπÁ®éÂè∑";
            worksheet.Cells[1, 6].Value = "ÈîÄÂîÆÊñπÂêçÁß∞";
            worksheet.Cells[1, 7].Value = "ÈîÄÂîÆÊñπÁ®éÂè∑";
            worksheet.Cells[1, 8].Value = "ÈáëÈ¢ùÂêàËÆ°";
            worksheet.Cells[1, 9].Value = "Á®éÈ¢ù";
            worksheet.Cells[1, 10].Value = "‰ª∑Á®éÂêàËÆ°";
            worksheet.Cells[1, 11].Value = "ÂèëÁ•®Á±ªÂûã";
            worksheet.Cells[1, 12].Value = "Êñá‰ª∂Ë∑ØÂæÑ";

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
>>>>>>> main
        }
    }
}
