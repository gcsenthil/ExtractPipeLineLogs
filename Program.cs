using Azure.Identity;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Threading.Tasks;

class Program
{
    static async Task Main(string[] args)
    {
        string tenantId = Environment.GetEnvironmentVariable("TENANT_ID");
        string clientId = Environment.GetEnvironmentVariable("CLIENT_ID");
        string clientSecret = Environment.GetEnvironmentVariable("CLIENT_SECRET");
        string subscriptionId = Environment.GetEnvironmentVariable("SUBSCRIPTION_ID");
        string resourceGroupName = Environment.GetEnvironmentVariable("RESOURCE_GROUP_NAME");
        string factoryName = Environment.GetEnvironmentVariable("FACTORY_NAME");
        string url = "https://management.azure.com/subscriptions/<subscriptionId>/resourceGroups/<resourcegroup>/providers/Microsoft.DataFactory/factories/<datafactoryname>/queryPipelineRuns?api-version=2018-06-01";
        if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(clientSecret) ||
            string.IsNullOrEmpty(subscriptionId) || string.IsNullOrEmpty(resourceGroupName) || string.IsNullOrEmpty(factoryName))
        {
            Console.WriteLine("Please set all required environment variables.");
            return;
        }

        try
        {
        
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var lstRunInfo = await CollectRunDetails<Value>(url, tenantId, clientId, clientSecret);
        ExportDataToExcel(lstRunInfo, "PipelineLogs.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    private static async Task<List<T>> CollectRunDetails<T>(string url, string tenantId, string clientId, string clientSecret)
    {
        var cred = new ClientSecretCredential(tenantId, clientId, clientSecret);
        var accessToken = cred.GetToken(new Azure.Core.TokenRequestContext(new[] { "https://management.azure.com/.default" })).Token;

        using (HttpClient httpClient = new HttpClient())
        {
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

            List<T> lstRunInfo = new List<T>();
            string continuationToken = null;

            do
            {
                var requestBody = new
                {
                    continuationToken = continuationToken,
                    lastUpdatedAfter = "2024-03-12",
                    lastUpdatedBefore = "2024-03-13"
                };
                string requestBodyJson = System.Text.Json.JsonSerializer.Serialize(requestBody);

                var request = new HttpRequestMessage(HttpMethod.Post, url)
                {
                    Content = new StringContent(requestBodyJson, System.Text.Encoding.UTF8, "application/json")
                };

                var response = await httpClient.SendAsync(request);
                response.EnsureSuccessStatusCode();

                string responseBody = await response.Content.ReadAsStringAsync();
                var runInfoResponse = System.Text.Json.JsonSerializer.Deserialize<Root<T>>(responseBody);

                if (runInfoResponse != null && runInfoResponse.value != null)
                {
                    lstRunInfo.AddRange(runInfoResponse.value);
                }

                continuationToken = runInfoResponse.continuationToken;
            } while (!string.IsNullOrWhiteSpace(continuationToken));
            return lstRunInfo;
        }
    }

    private static void ExportDataToExcel<T>(List<T> lstRunInfo, string filePath)
    {
        FileInfo file = new FileInfo(filePath);

        using (ExcelPackage package = new ExcelPackage(file))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(ws => ws.Name == "PipelineRuns");

            if (worksheet == null)
            {
                worksheet = package.Workbook.Worksheets.Add("PipelineRuns");
            }
            else
            {
                string newWorksheetName = GetUniqueWorksheetName(package.Workbook, "PipelineRuns");
                worksheet = package.Workbook.Worksheets.Add(newWorksheetName);
            }

            // Adding headers
            int rowIndex = 1;
            foreach (var prop in typeof(T).GetProperties())
            {
                worksheet.Cells[1, rowIndex].Value = prop.Name;
                rowIndex++;
            }

            // Adding data
            int rowCount = 2;
            foreach (var item in lstRunInfo)
            {
                int colIndex = 1;
                foreach (var prop in typeof(T).GetProperties())
                {
                    var value = prop.GetValue(item);
                    worksheet.Cells[rowCount, colIndex].Value = value != null ? GetValueAsString(value) : "";
                    colIndex++;
                }
                rowCount++;
            }

            package.Save();
        }

        Console.WriteLine("Data exported to Excel successfully.");
    }

    private static string GetValueAsString(object value)
    {
        if (value is string stringValue)
        {
            return stringValue;
        }
        else if (value is List<object> listValue)
        {
            try
            {
                return JsonConvert.SerializeObject(listValue);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to serialize value: {ex.Message}");
                return "";
            }
        }
        else
        {
            try
            {
                return JsonConvert.SerializeObject(value);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to serialize value: {ex.Message}");
                return "";
            }
        }
    }

    private static string GetUniqueWorksheetName(ExcelWorkbook workbook, string baseName)
    {
        int suffix = 1;
        string newName = baseName;
        while (workbook.Worksheets.Any(ws => ws.Name == newName))
        {
            suffix++;
            newName = $"{baseName}_{suffix}";
        }
        return newName;
    }
}

public class Value
{
    public string id { get; set; }
    public string runId { get; set; }
    public object debugRunId { get; set; }
    public string runGroupId { get; set; }
    public string pipelineName { get; set; }
    public Parameters parameters { get; set; }
    public InvokedBy invokedBy { get; set; }
    public DateTime runStart { get; set; }
    public DateTime runEnd { get; set; }
    public int durationInMs { get; set; }
    public string status { get; set; }
    public string message { get; set; }
    public PipelineReturnValue pipelineReturnValue { get; set; }
    public DateTime lastUpdated { get; set; }
    public List<object> annotations { get; set; }
    public RunDimension runDimension { get; set; }
    public bool isLatest { get; set; }
}

public class Parameters
{
    public string AhcCorrelationId { get; set; }
    public string IsManualRun { get; set; }
    public string ShouldCheckPredecessor { get; set; }
    public string MinNumWorker { get; set; }
    public string MaxNumWorker { get; set; }
    public string FixedNumWorker { get; set; }
    public string ForceOverwrite { get; set; }
    public string initScriptFile { get; set; }
    public string emailReceiver { get; set; }
    public string source { get; set; }
    public string dataset { get; set; }
    public string steps { get; set; }
    public string subjectArea { get; set; }
    public string mainFolder { get; set; }
    public string rejectedEmailReceiver { get; set; }
    public string timeFrame { get; set; }
    public string additionalConfigurations { get; set; }
    public string processingFolder { get; set; }
    public string IsBeta { get; set; }
    public string emailMonitorOnlyReceiver { get; set; }
    public string reprocessFlag { get; set; }
    public string LastRunDate { get; set; }
    public string ReportExportDLSPathFax { get; set; }
    public string ReportAppNameFax { get; set; }
    public string ReportExportFolder { get; set; }
    public string PowerbiReportIdFax { get; set; }
    public string emailreceiver { get; set; }
    public string Directory { get; set; }
    public string CompanName { get; set; }
    public string ControlFileTemplateFolder { get; set; }
    public string ControlFileName { get; set; }
    public string PPOFaxFolder { get; set; }
    public string CommunicationTo { get; set; }
    public string ResultsDirectory { get; set; }
    public string CutoffDate { get; set; }
    public string FaxProcessEmail { get; set; }
    public string AHCDirectory { get; set; }
    public string ArchivePath { get; set; }
    public string FailedFolder { get; set; }
    public string varAppName { get; set; }
    public string varCompanyID { get; set; }
    public string varEmailTo { get; set; }
    public string varJobName { get; set; }
    public string Source { get; set; }
    public string varAppNameADS { get; set; }
    public string FolderPath_SourceStore { get; set; }
    public string FolderPath_DestinationStore { get; set; }
    public string BackupFolder { get; set; }
    public string FileSystem { get; set; }
    public string FileName { get; set; }
    public string CompanyID { get; set; }
    public string AppName { get; set; }
    public string ClientId { get; set; }
    public string ClientSecret { get; set; }
    public string TenantId { get; set; }
    public string KeyVaultURL { get; set; }
    public string MemberWebAPI { get; set; }
    public string AccessTokenURL { get; set; }
    public string AccessTokenScope { get; set; }
    public string Filesystem { get; set; }
    public string ClientSecretKeyVaultURL { get; set; }
    public string SubscriptionId { get; set; }
    public string ResourceGroupName { get; set; }
    public string FunctionAppName { get; set; }
    public string message { get; set; }
    public string failedMessage { get; set; }
    public string PipelineName { get; set; }
    public string batchcount { get; set; }
    public string SourceTable { get; set; }
    public string DOS { get; set; }
    public string ALTDOS { get; set; }
    public string Company { get; set; }
    public string StartDate { get; set; }
    public string EndDate { get; set; }
    public string ArchiveFolder { get; set; }
    public string processTarget { get; set; }
    public string reprocessPath { get; set; }
    public string offset { get; set; }
    public string fetch { get; set; }
    public string LoadLogKey { get; set; }
}

public class PipelineReturnValue
{
}

public class InvokedBy
{
    public string id { get; set; }
    public string name { get; set; }
    public string invokedByType { get; set; }
    public string pipelineName { get; set; }
    public string pipelineRunId { get; set; }
}

public class RunDimension
{
}

public class Root<T>
{
    public List<T> value { get; set; }
    public string continuationToken { get; set; }
}
