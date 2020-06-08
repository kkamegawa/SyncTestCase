using System;
using System.Text;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System.Net.Http;
using System.Net.Http.Headers;
using SyncTestCase.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SyncTestCase
{
    public class SharePointJson
    {
        [JsonProperty(PropertyName = "path")]
        public string FilePath;
        [JsonProperty(PropertyName = "filename")]
        public string FileName;
        [JsonProperty(PropertyName = "fullpath")]
        public string FullPath;
    }

    public static class SyncTestCase
    {
        static Workbook TestWorkBook;
        public static string HttpEndpoint;
        public static string Token;

        static TestPlan TestPlan { get; set; }

        private static HttpClient httpClient = new HttpClient();

        static ILogger Log;

        /// <summary>
        /// ExcelファイルのURLを受け取って、Azure DevOpsへREST APIを使ってテストケースを登録する
        /// </summary>
        /// <param name="req"></param>
        /// <param name="log"></param>
        /// <returns></returns>
        [FunctionName("SyncTestCase")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req, ILogger log)
        {
            Initialize();
            Uri excelUri = null;

            Log = log;

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            var data = JsonConvert.DeserializeObject<SharePointJson>(requestBody);
            if(!string.IsNullOrEmpty(data.FullPath))
            {
                excelUri = new Uri(data.FullPath);
            }
            else
            {
                excelUri = new Uri(string.Format($"{data.FilePath}/{data.FileName}"));
            }
            log.LogInformation(string.Format($"File post {excelUri.AbsoluteUri }"));

            var result = await ParseExcelFile(excelUri);

            if (result)
            {
                // 登録
                result = await RegisterTest2AzureDevOps();
                if (!result)
                {
                    log.LogError(string.Format($"Failed : register Excel File to Azure DevOps"));
                }
            }

            return result
                ? (ActionResult)new OkObjectResult($"{excelUri} file create Test Plan.")
                : new BadRequestObjectResult($"{excelUri} is invalid file. Please check your test sheet.");
        }

        static void Initialize()
        {
            if (string.IsNullOrEmpty(Token))
            {
                string personalAccessToken = Environment.GetEnvironmentVariable("AZUREDEVOPS_PAT");
                Token = Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(string.Format("{0}:{1}", "", personalAccessToken)));
            }

            if (HttpEndpoint == null)
            {
               HttpEndpoint = Environment.GetEnvironmentVariable("AZUREDEVOPS_ENDPOINT");
            }

            if (TestWorkBook == null)
            {
                TestWorkBook = new Workbook();
            }
        }

        public static async Task<string> InvokeRestAPIDelete(string pathFormat,
            string contentType = "application/json")
        {
            string responseBody = string.Empty;

            HttpResponseMessage response;
            var uri = new Uri(new Uri(HttpEndpoint), relativeUri: pathFormat);
            httpClient.DefaultRequestHeaders.Accept.Clear();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Token);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            try
            {
                using (response = await httpClient.DeleteAsync(uri.AbsoluteUri))
                {
                    var statusCode = response.EnsureSuccessStatusCode();
                    if (statusCode.StatusCode != System.Net.HttpStatusCode.OK)
                    {
                        Log.LogError(string.Format($"NG:{statusCode.StatusCode}"));
                    }
                    responseBody = await response.Content.ReadAsStringAsync();
                }
            }
            catch (Exception ex)
            {
                Log.LogError(string.Format($"Exception{ex.Message}, {responseBody}"));
                throw;
            }

            return responseBody;
        }

        public static async Task<string> InvokeRestAPIPost(string jsonValue, string pathFormat, 
            string contentType = "application/json")
        {
            string responseBody = string.Empty;

            HttpResponseMessage response;
            var uri = new Uri(new Uri(HttpEndpoint), relativeUri: pathFormat);
            httpClient.DefaultRequestHeaders.Accept.Clear();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Token);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            try
            {
                using (response = await httpClient.PostAsync(uri.AbsoluteUri, new StringContent(jsonValue, Encoding.UTF8, contentType)))
                {
                    var statusCode = response.EnsureSuccessStatusCode();
                    if (statusCode.StatusCode != System.Net.HttpStatusCode.OK)
                    {
                        Log.LogError(string.Format($"NG:{statusCode.StatusCode}"));
                    }
                    responseBody = await response.Content.ReadAsStringAsync();
                }
            }
            catch (Exception ex)
            {
                Log.LogError(string.Format($"Exception{ex.Message}, {responseBody}"));
                throw;
            }

            return responseBody;
        }

        public static async Task<string> InvokeRestAPIPatch(string jsonValue, string pathFormat)
        {
            string responseBody = string.Empty;

            HttpResponseMessage response;
            var uri = new Uri(new Uri(HttpEndpoint), relativeUri: pathFormat);
            httpClient.DefaultRequestHeaders.Accept.Clear();
            httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", Token);
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            try
            {
                using (response = await httpClient.PatchAsync(uri.AbsoluteUri, new StringContent(jsonValue, Encoding.UTF8, "application/json-patch+json")))
                {
                    var statusCode = response.EnsureSuccessStatusCode();
                    if (statusCode.StatusCode != System.Net.HttpStatusCode.OK)
                    {
                        Log.LogError(string.Format($"NG:{statusCode.StatusCode}"));
                    }
                    responseBody = await response.Content.ReadAsStringAsync();
                }
            }
            catch (Exception ex)
            {
                Log.LogError(string.Format($"Exception{ex.Message}, {responseBody}"));
                throw;
            }
            return responseBody;
        }

        /// <summary>
        /// Create Test Plans
        /// </summary>
        /// <param name="planName"></param>
        /// <returns></returns>
        static async Task<TestPlanRootobject> CreateTestPlan(TestPlan testPlan)
        {
            TestPlanRootobject testPlanRootobject = new TestPlanRootobject();


            testPlan.StartDate = DateTime.Now;
            testPlan.EndDate = testPlan.StartDate.AddDays(14);//Just Sample fix 2weeks.
            var json = JsonConvert.SerializeObject(testPlan);

            Log.LogInformation(string.Format($"Test Plan {testPlan.Name}"));
            // https://docs.microsoft.com/en-us/rest/api/azure/devops/testplan/test%20%20suites/create?view=azure-devops-rest-6.0
            var result = await InvokeRestAPIPost(json, "testplan/plans?api-version=6.0-preview.1");
            if(string.IsNullOrEmpty(result))
            {
                Log.LogError($"CreateTestPlan {testPlan.Name} create failed.");
            }
            else
            {
                // Test Plan from Azure DevOps
                testPlanRootobject = JsonConvert.DeserializeObject<TestPlanRootobject>(result);
            }
            Log.LogInformation(string.Format($"Test Plan {testPlan.Name} End"));
            return testPlanRootobject;
        }

        /// <summary>
        /// Create and encode html for Work Items in Test repro steps
        /// </summary>
        /// <param name="steps">step of test</param>
        /// <returns>encoded html string</returns>
        public static string EncodeHtmlString(TestStep steps)
        {
            var sb = new StringBuilder();
            sb.Append(string.Format($"<steps id=\"0\">"));

            string repoStep = default(string);
            string expectStep = default(string);
            for (int i = 0; i < steps.StepRepro.Count; i++)
            {
                repoStep = System.Web.HttpUtility.HtmlEncode(steps.StepRepro[i]);
                expectStep = System.Web.HttpUtility.HtmlEncode(steps.StepExpected[i]);
                sb.Append(string.Format($"<step id=\"{i + 2}\" type=\"ValidateStep\"><parameterizedString isformatted=\"true\">&lt;DIV&gt;&lt;P&gt;{repoStep}&amp;nbsp;&lt;/P&gt;&lt;/DIV&gt;</parameterizedString>"));
                sb.Append(string.Format($"<parameterizedString isformatted=\"true\">&lt;P&gt;{expectStep}&amp;nbsp;&lt;/P&gt;</parameterizedString><description/></step>"));
            }
            sb.Append("</steps>");

            return sb.ToString();
        }

        /// <summary>
        /// Create Work Item for Test and Test repro steps.
        /// </summary>
        /// <param name="testPlanID"></param>
        /// <param name="testSuiteID"></param>
        /// <param name="testCase"></param>
        /// <returns></returns>
        static async Task<bool>CreateTestCase(int testPlanID, int testSuiteID, TestCase[] testCase)
        {
            bool functionResult = true;

            Log.LogInformation(string.Format($"PlanID:{testPlanID}, SuiteID:{testSuiteID} Test Case Start"));

            foreach (var item in testCase)
            {
                var steps = EncodeHtmlString(item.TestStep);
                var registeringTestSteps = new TestCase[]
                {
                    new TestCase{Value = item.Value, Operation = "add", Path = "/fields/System.Title"},
                    new TestCase{Value = steps, Operation = "add", Path = "/fields/Microsoft.VSTS.TCM.Steps"},
                };

                TestCaseRootobject resultObject = null;
                var json = JsonConvert.SerializeObject(registeringTestSteps);
                Log.LogInformation(string.Format($"Test Work Item:{item.Value}'s Test Case") + json);
                // https://docs.microsoft.com/en-us/rest/api/azure/devops/wit/work%20items/create?view=azure-devops-rest-6.0
                var result = await InvokeRestAPIPost(json,
                    "wit/workitems/$Test%20Case?api-version=6.0-preview.3", "application/json-patch+json");
                if (string.IsNullOrEmpty(result))
                {
                    Log.LogError($"Test Work Item:{item.Value}'s TestCase  create failed.");
                    return false;
                }
                else
                {
                    resultObject = JsonConvert.DeserializeObject<TestCaseRootobject>(result);
                }

                // create association to Test Plan
                result = await InvokeRestAPIPost(string.Empty,
                    string.Format($"test/Plans/{testPlanID}/suites/{testSuiteID}/testcases/{resultObject.id}?api-version=6.0-preview.3"));
                if (string.IsNullOrEmpty(result))
                {
                    Log.LogError($"TestCase {testSuiteID} create relation failed.");
                    return false;
                }

            }
            Log.LogInformation(string.Format($"Suite ID:{testSuiteID}'s Test Case End"));

            return functionResult;
        }

        /// <summary>
        /// テストスィートを作る
        /// </summary>
        /// <param name="testPlanID"></param>
        /// <param name="testSuite"></param>
        /// <returns></returns>
        static async Task<TestSuiteResultRootobject> CreateTestSuite(int testPlanID, TestSuite testSuite)
        {
            TestSuiteResultRootobject testSuiteformAzureDevOps = new TestSuiteResultRootobject();
            var json = JsonConvert.SerializeObject(testSuite);
            Log.LogInformation(string.Format($"Plan id :{testPlanID}, Root Suite:{testSuite.Parent.ParentID} Test Suite {testSuite.Name} Start"));
            Log.LogInformation(json);

            // https://docs.microsoft.com/en-us/rest/api/azure/devops/testplan/test%20%20plans/create?view=azure-devops-rest-6.0
            var result = await InvokeRestAPIPost(json,
                string.Format($"testplan/Plans/{testPlanID}/suites/?api-version=6.0-preview.1"));
            if (string.IsNullOrEmpty(result))
            {
                Log.LogError($"TestSuite {testSuite.Name} create failed.");
            }
            else
            {
                testSuiteformAzureDevOps = JsonConvert.DeserializeObject<TestSuiteResultRootobject>(result);
            }
            Log.LogInformation(string.Format($"Test Case {testSuite.Name} end"));

            return testSuiteformAzureDevOps;
        }

        /// <summary>
        /// Invoke-Rest API to Azure DevOps
        /// 1. create Test Suite
        /// 2. create Test Plan
        /// 3. create Work Item(Type:Test) and add test repro step.
        /// </summary>
        /// <returns></returns>
        static async Task<bool>RegisterTest2AzureDevOps()
        {
            bool functionResult = true;
            TestSuiteResultRootobject testSuiteResultRootfromAzureDevOps = new TestSuiteResultRootobject();

            var testPlanfromAzureDevOps = await CreateTestPlan(TestPlan);
            if (testPlanfromAzureDevOps == null)
            {
                Log.LogError($"TestPlan {TestPlan.Name} create failed.");
                return false;
            }

            for(int i = 0;i < TestPlan.TestSuites.Count;i++)
            {
                TestPlan.TestSuites[i].Parent.ParentID = testPlanfromAzureDevOps.rootSuite.id;
                testSuiteResultRootfromAzureDevOps = await CreateTestSuite(testPlanfromAzureDevOps.id, 
                    TestPlan.TestSuites[i]);
                if (testSuiteResultRootfromAzureDevOps == null)
                {
                    Log.LogError($"TestSuite {TestPlan.TestSuites[i].Name} create failed.");
                    return false;
                }
                else
                {
                    for(int j = 0; j < TestPlan.TestSuites[i].TestCases.Count; j++)
                    {
                        functionResult = await CreateTestCase(
                            testPlanfromAzureDevOps.id,
                            testSuiteResultRootfromAzureDevOps.id,
                            TestPlan.TestSuites[i].TestCases.ToArray());
                        if (!functionResult)
                        {
                            Log.LogError($"TestCase No.{TestPlan.TestSuites[i].ID} is {TestPlan.TestSuites[i].Name} test case {TestPlan.TestSuites[i].TestCases[j].Value} create failed.");
                            return functionResult;
                        }
                    }
                }
            }

            return functionResult;
        }

        /// <summary>
        /// Excelファイルをパースして、登録
        /// </summary>
        /// <param name="excelURL"></param>
        /// <returns></returns>
        public static async Task<bool> ParseExcelFile(Uri excelURL)
        {
            bool functionResult = true;

            var excelByteArray = await httpClient.GetByteArrayAsync(excelURL);
            if (excelByteArray == null)
            {
                Log.LogInformation("excelByteArray is null");
                return false;
            }
            var excelStream = new MemoryStream(excelByteArray);
            if (excelStream == null)
            {
                Log.LogInformation("excelStream is null");
                return false;
            }

            using (var document = SpreadsheetDocument.Open(excelStream, false))
            {
                var wbPart = document.WorkbookPart;
                var stringTable = wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                var sheet = wbPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if(sheet == null)
                {
                    Log.LogError("Request workbook does not exist sheet.");
                }

                var wsheetPart = wbPart.GetPartById(sheet.Id) as WorksheetPart;
                // Tab name is name of Test Plan
                TestPlan = new TestPlan();
                TestPlan.Name = sheet.Name;

                var ws = wsheetPart.Worksheet;
                List<string[]> testAssets = new List<string[]>();
                int rownum = 0;
                char[] columnID = { 'A', 'B', 'C', 'D', };

                int indexStart = 0;
                //var suiteTuple = new List<Tuple<int, int>>();
                foreach (var (row, index) in ws.Descendants<Row>().Skip(1).Indexed())
                {
                    var cells = row.Elements<Cell>().ToArray();
                    var cellStrings = new string[4];

                    // stopping loop, when all column is empty.
                    if (cells.All(x => x.DataType.HasValue == false))
                    {
                        CreateTestPlanfromString(testAssets);
                        Log.LogInformation("All Data proceeded.");
                        break;
                    }
                    // found Test Suite Row number(all column is fill).
                    if (cells.All(x => x.DataType.HasValue == true) && cells.Length == 4)
                    {
                        if (testAssets.Count() >= 1)
                        {
                            CreateTestPlanfromString(testAssets);
                            testAssets.Clear();
                        }
                        //suiteTuple.Add(new Tuple<int, int>(indexStart, index - 1));
                        indexStart = index;
                    }
                    for (int i = 0; i < cells.Length; i++)
                    {
                        var currentColunm = cells[i].CellReference.Value.Substring(0, 1).ToCharArray();
                        rownum = Array.FindIndex(columnID, x =>
                        {
                            if (x == currentColunm[0])
                            {
                                return true;
                            }
                            return false;
                        });

                        // Test Plan or Test Step's row
                        switch (cells[i].DataType.Value)
                        {
                            case CellValues.SharedString:
                                if (stringTable != null)
                                {
                                    var value = GetSharedString(wbPart, int.Parse(cells[i].InnerText));
                                    cellStrings[rownum] = value;
                                }
                                else
                                {
                                    Log.LogError($"Invalid Type in row {index}, cell {i}");
                                }
                                break;
                            case CellValues.String:
                                cellStrings[rownum] = cells[i].InnerText;
                                break;
                            default:
                                Log.LogError($"Invalid Type in {cells[i].CellReference.Value}");
                                break;
                        }
                    }
                    testAssets.Add(cellStrings);
                }
                CreateTestPlanfromString(testAssets);
            }

            return functionResult;
        }

        /// <summary>
        /// Excelデータの文字配列からテストケースの塊を取り出す。
        /// </summary>
        /// <param name="testAssets"></param>
        private static void CreateTestPlanfromString(IList<string[]> testAssets)
        {
            // Because Row number 0 must be Test Suite, start Row number 1.
            // Get Test Case's indexes.
            var indexes = testAssets.Select((p, i) => new { Content = p, Index = i })
                .Where(x => string.IsNullOrEmpty(x.Content[1]) == false)
                .Select(x => x.Index).ToList();

            TestPlan.TestSuites.Add(
                new TestSuite
                {
                    Name = testAssets[0][0]
                });

            //processing per test suite
            int suiteNumber = TestPlan.TestSuites.Count() - 1;
            int endIndex = 0;
            for(int i = 0; i < indexes.Count(); i++)
            {
                TestPlan.TestSuites[suiteNumber].TestCases.Add(new TestCase
                {
                    Value = testAssets[indexes[i]][1]
                });
                if(i == indexes.Count() - 1)
                {
                    endIndex = testAssets.Count();
                }
                else {
                    endIndex = indexes[i + 1];
                }
                var stepLists = testAssets.Skip(indexes[i]).Take(endIndex - indexes[i]).Select(x => x[2]);
                TestPlan.TestSuites[suiteNumber].TestCases[i].TestStep.StepRepro.AddRange(stepLists);

                var stepExpectLists = testAssets.Skip(indexes[i]).Take(testAssets.Count() - indexes[i]).Select(x => x[3]);
                TestPlan.TestSuites[suiteNumber].TestCases[i].TestStep.StepExpected.AddRange(stepExpectLists);
            }
        }


        private static string GetSharedString(WorkbookPart wb_part, int index)
        {
            if (wb_part == null && wb_part.GetPartsOfType<SharedStringTablePart>().Count() == 0)
            {
                return string.Empty;
            }
            SharedStringTablePart shareStringPart = wb_part.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringItem sh_item = shareStringPart.SharedStringTable.Elements<SharedStringItem>().ElementAt(index);
            if (sh_item != null && sh_item.Text != null)
            {
                return sh_item.Text.InnerText;
            }
            else
            {
                return string.Empty;
            }
        }
    }
    public static partial class TupleEnumerable
    {
        public static IEnumerable<(T item, int index)> Indexed<T>(this IEnumerable<T> source)
        {
            if (source == null) throw new ArgumentNullException(nameof(source));

            IEnumerable<(T item, int index)> impl()
            {
                var i = 0;
                foreach (var item in source)
                {
                    yield return (item, i);
                    ++i;
                }
            }

            return impl();
        }
    }
}
