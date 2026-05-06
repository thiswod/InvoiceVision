using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace InvoiceVision;

internal sealed class LocalPdfInvoiceExtractor
{
    public InvoiceData Extract(string pdfPath)
    {
        if (string.IsNullOrWhiteSpace(pdfPath))
        {
            throw new ArgumentException("PDF path is required.", nameof(pdfPath));
        }

        string fullText = ReadPdfText(pdfPath);
        string normalizedText = NormalizeText(fullText);
        FileNameInvoiceInfo fileNameInfo = ParseFileName(Path.GetFileName(pdfPath));

        var invoice = new InvoiceData
        {
            ImagePath = pdfPath,
            InvoiceType = "数电发票（本地PDF提取）",
            RawText = normalizedText
        };

        invoice.InvoiceNum = FirstNonEmpty(
            ExtractFirstGroup(normalizedText, @"发票号码[:：]?\s*([0-9]{8,20})"),
            ExtractFirstGroup(normalizedText, @"电子发票号码[:：]?\s*([0-9]{8,20})"),
            ExtractFirstGroup(normalizedText, @"([0-9]{20})"));

        invoice.InvoiceCode = FirstNonEmpty(
            ExtractFirstGroup(normalizedText, @"发票代码[:：]?\s*([0-9]{8,20})"),
            invoice.InvoiceNum);

        invoice.InvoiceDate = FirstNonEmpty(
            NormalizeDate(ExtractFirstGroup(normalizedText, @"开票日期[:：]?\s*([0-9]{4}[年/-][0-9]{1,2}[月/-][0-9]{1,2}日?)")),
            NormalizeDate(ExtractFirstGroup(normalizedText, @"([0-9]{4}[年/-][0-9]{1,2}[月/-][0-9]{1,2}日?)")),
            fileNameInfo.InvoiceDate);

        var purchaserParty = ExtractPartyInfo(normalizedText, "购买方信息", "销售方信息");
        var sellerParty = ExtractPartyInfo(normalizedText, "销售方信息", "项目名称");

        invoice.PurchaserRegisterNum = FirstNonEmpty(
            purchaserParty.taxNo,
            ExtractFirstGroup(normalizedText, @"购买方(?:纳税人)?识别号[:：]?\s*([0-9A-Z]{15,20})"));

        invoice.SellerRegisterNum = FirstNonEmpty(
            sellerParty.taxNo,
            ExtractFirstGroup(normalizedText, @"销售方(?:纳税人)?识别号[:：]?\s*([0-9A-Z]{15,20})"));

        invoice.PurchaserName = FirstNonEmpty(
            purchaserParty.name,
            ExtractCompanyNameByLabel(normalizedText, "购买方名称"),
            ExtractCompanyNameByTwoNameBlocks(normalizedText, takeSecond: false),
            fileNameInfo.PurchaserName);

        invoice.SellerName = FirstNonEmpty(
            sellerParty.name,
            ExtractCompanyNameByLabel(normalizedText, "销售方名称"),
            ExtractCompanyNameByTwoNameBlocks(normalizedText, takeSecond: true));

        (string amount, string tax, string total) = ExtractAmounts(normalizedText);
        invoice.TotalAmount = amount;
        invoice.TotalTax = tax;
        invoice.AmountInFiguers = FirstNonEmpty(total, fileNameInfo.AmountInFiguers);

        if (string.IsNullOrWhiteSpace(invoice.AmountInFiguers) &&
            decimal.TryParse(invoice.TotalAmount, NumberStyles.Any, CultureInfo.InvariantCulture, out var amountValue) &&
            decimal.TryParse(invoice.TotalTax, NumberStyles.Any, CultureInfo.InvariantCulture, out var taxValue))
        {
            invoice.AmountInFiguers = (amountValue + taxValue).ToString("0.##", CultureInfo.InvariantCulture);
        }

        return invoice;
    }

    private static string ReadPdfText(string pdfPath)
    {
        using PdfDocument document = PdfDocument.Open(pdfPath);
        return string.Join(Environment.NewLine, document.GetPages().Select(page => page.Text ?? string.Empty));
    }

    private static string NormalizeText(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return string.Empty;
        }

        string normalized = text
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace("\u00A0", " ")
            .Replace("（", "(")
            .Replace("）", ")")
            .Replace("：", ":");

        normalized = Regex.Replace(normalized, @"[ \t]+", " ");
        normalized = Regex.Replace(normalized, @"\n{2,}", "\n");
        return normalized.Trim();
    }

    private static FileNameInvoiceInfo ParseFileName(string fileName)
    {
        Match match = Regex.Match(
            fileName,
            @"^(?<name>.+)-(?<amount>-?\d+(?:\.\d+)?)(?:退(?<tax>\d+(?:\.\d+)?))?-(?<date>\d{4}-\d{2}-\d{2})--\d+\.pdf$",
            RegexOptions.IgnoreCase);

        if (!match.Success)
        {
            Match fallback = Regex.Match(
                fileName,
                @"^(?<date>\d{8})_(?<name>[^_]+)_.*\.pdf$",
                RegexOptions.IgnoreCase);

            if (fallback.Success)
            {
                string rawDate = fallback.Groups["date"].Value;
                return new FileNameInvoiceInfo(
                    fallback.Groups["name"].Value.Replace(" ", string.Empty).Trim(),
                    string.Empty,
                    $"{rawDate[..4]}-{rawDate.Substring(4, 2)}-{rawDate.Substring(6, 2)}");
            }

            return new FileNameInvoiceInfo(string.Empty, string.Empty, string.Empty);
        }

        return new FileNameInvoiceInfo(
            match.Groups["name"].Value.Trim(),
            match.Groups["amount"].Value.Trim(),
            match.Groups["date"].Value.Trim());
    }

    private static string ExtractFirstGroup(string text, string pattern)
    {
        Match match = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
        return match.Success && match.Groups.Count > 1 ? match.Groups[1].Value.Trim() : string.Empty;
    }

    private static string ExtractCompanyNameByLabel(string text, string label)
    {
        Match match = Regex.Match(
            text,
            $"{Regex.Escape(label)}\\s*:?\\s*([^\\n]+)",
            RegexOptions.IgnoreCase | RegexOptions.Multiline);

        if (!match.Success)
        {
            return string.Empty;
        }

        return CleanCompanyName(match.Groups[1].Value);
    }

    private static string ExtractCompanyNameByTwoNameBlocks(string text, bool takeSecond)
    {
        MatchCollection matches = Regex.Matches(text, @"名称\s*:?\s*([^\n]+)", RegexOptions.IgnoreCase | RegexOptions.Multiline);
        var names = new List<string>();

        foreach (Match match in matches)
        {
            string candidate = CleanCompanyName(match.Groups[1].Value);
            if (LooksLikeCompanyName(candidate))
            {
                names.Add(candidate);
            }
        }

        if (names.Count == 0)
        {
            return string.Empty;
        }

        if (takeSecond)
        {
            return names.Count > 1 ? names[1] : string.Empty;
        }

        return names[0];
    }

    private static (string taxNo, string name) ExtractPartyInfo(string text, string startLabel, string endLabel)
    {
        Match match = Regex.Match(
            text,
            $"{Regex.Escape(startLabel)}\\s*(?<tax>[0-9A-Z]{{15,20}})?\\s*名称\\s*:?\\s*(?<name>.+?)(?=统一社会信用代码/纳税人识别号:|{Regex.Escape(endLabel)}|项目名称|备注|$)",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        if (!match.Success)
        {
            return (string.Empty, string.Empty);
        }

        return (
            match.Groups["tax"].Value.Trim(),
            CleanCompanyName(match.Groups["name"].Value));
    }

    private static (string amount, string tax, string total) ExtractAmounts(string text)
    {
        string amount = string.Empty;
        string tax = string.Empty;
        string total = string.Empty;

        Match summaryMatch = Regex.Match(
            text,
            @"合\s*计\s*¥?\s*(?<amount>-?\d+(?:\.\d+)?)\s*¥?\s*(?<tax>-?\d+(?:\.\d+)?)\s*价税合计",
            RegexOptions.IgnoreCase | RegexOptions.Singleline);

        if (summaryMatch.Success)
        {
            amount = NormalizeMoney(summaryMatch.Groups["amount"].Value);
            tax = NormalizeMoney(summaryMatch.Groups["tax"].Value);
        }

        total = NormalizeMoney(
            FirstNonEmpty(
                ExtractFirstGroup(text, @"价税合计(?:\(大写\))?.{0,30}?\(小写\)\s*¥?\s*(-?\d+(?:\.\d+)?)"),
                ExtractAmountByLabel(text, "价税合计(小写)"),
                ExtractAmountByLabel(text, "价税合计")));

        if (string.IsNullOrWhiteSpace(total))
        {
            Match totalMatch = Regex.Match(
                text,
                @"价税合计(?:\(大写\))?.{0,30}?(?:\(小写\))?\s*(-?\d+(?:\.\d+)?)",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            if (totalMatch.Success)
            {
                total = NormalizeMoney(totalMatch.Groups[1].Value);
            }
        }

        if (string.IsNullOrWhiteSpace(amount))
        {
            amount = NormalizeMoney(ExtractAmountByLabel(text, "金额合计"));
        }

        if (string.IsNullOrWhiteSpace(tax))
        {
            tax = NormalizeMoney(ExtractFirstGroup(text, @"合\s*计\s*¥?\s*-?\d+(?:\.\d+)?\s*¥?\s*(-?\d+(?:\.\d+)?)"));
        }

        return (amount, tax, total);
    }

    private static string ExtractAmountByLabel(string text, string label)
    {
        string pattern = $"{Regex.Escape(label)}\\s*:?\\s*(-?\\d+(?:,\\d{{3}})*(?:\\.\\d+)?)";
        Match direct = Regex.Match(text, pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);
        if (direct.Success)
        {
            return direct.Groups[1].Value;
        }

        int index = text.IndexOf(label, StringComparison.OrdinalIgnoreCase);
        if (index < 0)
        {
            return string.Empty;
        }

        int length = Math.Min(80, text.Length - index);
        string window = text.Substring(index, length);
        Match nearby = Regex.Match(window, @"-?\d+(?:,\d{3})*(?:\.\d+)?");
        return nearby.Success ? nearby.Value : string.Empty;
    }

    private static string NormalizeDate(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        string normalized = value
            .Replace("年", "-")
            .Replace("月", "-")
            .Replace("日", string.Empty)
            .Replace("/", "-")
            .Trim();

        return DateTime.TryParse(normalized, out DateTime parsed)
            ? parsed.ToString("yyyy-MM-dd")
            : normalized;
    }

    private static string NormalizeMoney(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        string normalized = value.Replace(",", string.Empty).Trim();
        return decimal.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal parsed)
            ? parsed.ToString("0.00", CultureInfo.InvariantCulture)
            : normalized;
    }

    private static string CleanCompanyName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        string result = value.Trim();
        result = Regex.Split(result, @"\s+(统一社会信用代码|纳税人识别号|识别号|地址|电话|开户行|账号)\b").FirstOrDefault()?.Trim() ?? result;
        result = Regex.Split(result, @"销售方信息|购买方信息|项目名称|备注").FirstOrDefault()?.Trim() ?? result;
        result = Regex.Replace(result, @"\s*[0-9A-Z]{15,20}\s*$", string.Empty).Trim();

        if (result.Contains("名称:"))
        {
            result = result.Split("名称:", StringSplitOptions.RemoveEmptyEntries)[0].Trim();
        }

        return result;
    }

    private static bool LooksLikeCompanyName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        if (Regex.IsMatch(value, @"^[0-9A-Z]{10,}$"))
        {
            return false;
        }

        return value.Any(ch => ch >= 0x4E00 && ch <= 0x9FFF);
    }

    private static string FirstNonEmpty(params string[] values)
    {
        return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value)) ?? string.Empty;
    }

    private sealed record FileNameInvoiceInfo(string PurchaserName, string AmountInFiguers, string InvoiceDate);
}
