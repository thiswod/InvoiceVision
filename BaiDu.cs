using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using WodToolkit.Http;
using WodToolkit.Json;

namespace InvoiceVision
{
    public class BaiDu
    {
        private HttpRequestClass http = new HttpRequestClass();
        private string Ak = "";
        private string Sk = "";
        public BaiDu(string ApiKey = "",string SecretKey = "")
        {
            Ak = ApiKey;
            Sk = SecretKey;
        }
        public string vat_invoice(string image, string imageType = "png")
        {
            string url = $"https://aip.baidubce.com/rest/2.0/ocr/v1/vat_invoice?access_token={GetAccessToken()}";
            http.Open(url, HttpMethod.Post);
            
            string requestBody;
            if (imageType.ToLower() == "pdf")
            {
                // PDF文件使用pdf_file参数，直接使用base64字符串，不需要data URI前缀
                requestBody = $"pdf_file={WebUtility.UrlEncode(image)}&seal_tag=false";
            }
            else
            {
                // 图片文件使用image参数，需要完整的数据URI格式：data:image/{type};base64,{base64字符串}
                string base64Data = image;
                if (!base64Data.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
                {
                    base64Data = $"data:image/{imageType};base64,{base64Data}";
                }
                requestBody = $"image={WebUtility.UrlEncode(base64Data)}&seal_tag=false";
            }
            
            http.Send(requestBody);
            return http.GetResponse().Body;
        }
        string GetAccessToken()
        {
            string url = "https://aip.baidubce.com/oauth/2.0/token";
            string queryString = $"grant_type=client_credentials&client_id={Ak}&client_secret={Sk}";
            http.Open(url,HttpMethod.Post);
            http.Send(queryString);
            dynamic dynamic = EasyJson.ParseJsonToDynamic(http.GetResponse().Body);
            return dynamic.access_token;
        }
    }
}
