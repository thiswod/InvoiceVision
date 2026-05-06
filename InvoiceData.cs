namespace InvoiceVision;

internal sealed class InvoiceData
{
    public string InvoiceNum { get; set; } = string.Empty;
    public string InvoiceCode { get; set; } = string.Empty;
    public string InvoiceDate { get; set; } = string.Empty;
    public string PurchaserName { get; set; } = string.Empty;
    public string PurchaserRegisterNum { get; set; } = string.Empty;
    public string SellerName { get; set; } = string.Empty;
    public string SellerRegisterNum { get; set; } = string.Empty;
    public string TotalAmount { get; set; } = string.Empty;
    public string TotalTax { get; set; } = string.Empty;
    public string AmountInFiguers { get; set; } = string.Empty;
    public string InvoiceType { get; set; } = string.Empty;
    public string ImagePath { get; set; } = string.Empty;
    public string RawText { get; set; } = string.Empty;
    public dynamic? RawData { get; set; }
}
