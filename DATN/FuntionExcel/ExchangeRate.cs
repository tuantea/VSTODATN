using Newtonsoft.Json;
using System;
using RestSharp;
using System.Windows.Forms;

namespace DATN.FuntionExcel
{
    internal class ExchangeRate
    {
        public static string GetExchangeRate()
        {
            string baseCurrency = "USD";  // Đơn vị tiền tệ cơ sở
            string targetCurrency = "VND";  // Đơn vị tiền tệ mục tiêu

            string apiUrl = $"https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency={baseCurrency}&to_currency={targetCurrency}&apikey=QDV4XQKV0OC06XOX";

            RestClient client = new RestClient(apiUrl);
            RestRequest request = new RestRequest("",Method.Get);
            RestResponse response = client.Execute(request);
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                dynamic jsonResponse = JsonConvert.DeserializeObject(response.Content);
                dynamic exchangeRate = jsonResponse["Realtime Currency Exchange Rate"]["5. Exchange Rate"];
                return exchangeRate;
            }
            else
            {
                return "Failed to get exchange rate.";
            }
        }
        public static string GetExchangeRate(string baseCurrency, string targetCurrency)
        {

            string apiUrl = $"https://www.alphavantage.co/query?function=CURRENCY_EXCHANGE_RATE&from_currency={baseCurrency}&to_currency={targetCurrency}&apikey=QDV4XQKV0OC06XOX";

            RestClient client = new RestClient(apiUrl);
            RestRequest request = new RestRequest("", Method.Get);
            RestResponse response = client.Execute(request);
            if (response.StatusCode == System.Net.HttpStatusCode.OK)
            {
                dynamic jsonResponse = JsonConvert.DeserializeObject(response.Content);
                dynamic exchangeRate = jsonResponse["Realtime Currency Exchange Rate"]["5. Exchange Rate"];
                return exchangeRate;
            }
            else
            {
                return "Failed to get exchange rate.";
            }
        }
        [ExcelDna.Integration.ExcelFunction(Description = "Viết số bằng chữ tiếng Việt có dấu")]
        public static string GetExchange(
            [ExcelDna.Integration.ExcelArgument(Description = "Đơn vị cần chuyển đổi")] string baseCurrency,
            [ExcelDna.Integration.ExcelArgument(Description = "Đơn vị chuyện đổi")] string targetCurrency
            )
        {
            return GetExchangeRate(baseCurrency,targetCurrency);
        }
    }
}
