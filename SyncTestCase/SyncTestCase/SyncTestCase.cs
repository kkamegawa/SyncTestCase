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
        public static TestPlanRootobject TestPlanResultObject { get; private set; }
        public static TestSuiteRootObject TestSuiteResultObject { get; private set; }
        public static TestCaseRootobject TestCaseResultObject { get; private set; }
        public static TestCaseRelationRootobject TestRelationResultObject { get; private set; }
        public static UpdateTestCaseRootobject TestCaseUpdateResultObject { get; private set; }

        static TestPlan TestPlan { get; set; }

        private static HttpClient httpClient = new HttpClient();

        static ILogger Log;

        /// <summary>
        /// SwwharePointにあるExcelファイルのURLを受け取って、Azure DevOpsへREST APIを使ってテストケースを登録する
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
            }

            return result
                ? (ActionResult)new OkObjectResult($"{excelUri} file create Test Plan.")
                : new BadRequestObjectResult($"{excelUri} is invalid file. Please check your test sheet.");
        }

        private static void Initialize()
        {
            if (string.IsNullOrEmpty(Token))
            {
                string personalAccessToken = ""; Environment.GetEnvironmentVariable("AZUREDEVOPS_PAT");
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
            httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json-patch+json"));
            try
            {
                using (response = await httpClient.PatchAsync(uri.AbsoluteUri, new StringContent(jsonValue, Encoding.UTF8, "application/json")))
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
        /// テスト計画を作る
        /// </summary>
        /// <param name="planName"></param>
        /// <returns></returns>
        static async Task<bool>CreateTestPlan(TestPlan testPlan)
        {
            bool functionResult = true;
            
            testPlan.StartDate = DateTime.Now;
            testPlan.EndDate = testPlan.StartDate.AddDays(14);//デモなので二週間固定で
            var json = JsonConvert.SerializeObject(testPlan);

            Log.LogInformation(string.Format($"Test Plan {testPlan.Name}"));

            var result = await InvokeRestAPIPost(json, "test/plans?api-version=5.0-preview.2");
            if(string.IsNullOrEmpty(result))
            {
                functionResult = false;
                Log.LogError($"CreateTestPlan {testPlan.Name} create failed.");
            }
            else
            {
                var resultObject = JsonConvert.DeserializeObject<TestPlanRootobject>(result);
                testPlan.ID = resultObject.id;
                int suiteID = 0;
                
                if(int.TryParse(resultObject.rootSuite.id, out suiteID) == true)
                {
                    testPlan.SuiteRootID = suiteID;
                }
            }
            Log.LogInformation(string.Format($"Test Plan {testPlan.Name} End"));
            return functionResult;
        }

        public static string EncodeHtmlString(TestStep steps)
        {
            var sb = new StringBuilder();
            sb.Append(string.Format($"<steps id=\\\"0\\\" last=\\\"{steps.StepRepro.Count + 1}\\\">"));

            for (int i = 0; i < steps.StepRepro.Count; i++)
            {
                sb.Append(string.Format($"<step id=\\\"{i + 2}\\\" type=\\\"ValidateStep\\\"><parameterizedString isformatted=\\\"true\\\">&lt;DIV&gt;&lt;P&gt;{steps.StepRepro[i]}&amp;nbsp;&lt;/P&gt;&lt;/DIV&gt;</parameterizedString>"));
                sb.Append(string.Format($"<parameterizedString isformatted=\\\"true\\\">&lt;P&gt;{steps.StepExpected[i]}&amp;nbsp;&lt;/P&gt;</parameterizedString><description/></step>"));
            }
            sb.Append("</steps>");

            return sb.ToString();
        }

        /// <summary>
        /// テストケースの作成
        /// </summary>
        /// <param name="testPlanID"></param>
        /// <param name="testSuiteID"></param>
        /// <param name="testCase"></param>
        /// <returns></returns>
        static async Task<bool>CreateTestCase(int testPlanID, int testSuiteID, TestCase[] testCase)
        {
            bool functionResult = true;

            Log.LogInformation(string.Format($"PlanID:{testPlanID}, SuiteID:{testSuiteID} Test Case Start"));

            TestCaseRootobject resultObject = null;
            var json = JsonConvert.SerializeObject(testCase);
            Log.LogInformation(string.Format($"TestSuiteID:{testSuiteID}'s Test Case") + json);
            var result = await InvokeRestAPIPost(json,
                "wit/workitems/$Test%20Case?api-version=5.0-preview.3",
                "application/json-patch+json");
            if (string.IsNullOrEmpty(result))
            {
                Log.LogError($"TestSuite {testSuiteID}'s TestCase  create failed.");
                return false;
            }
            else
            {
                resultObject = JsonConvert.DeserializeObject<TestCaseRootobject>(result);
                for(int i = 0;i< testCase.Length; i++)
                {
                    testCase[i].ID = resultObject.id;
                }
            }

            //関連付け
            result = await InvokeRestAPIPost(string.Empty,
                string.Format($"test/Plans/{testPlanID}/suites/{testSuiteID}/testcases/{resultObject.id}?api-version=5.0-preview.3"));
            if (string.IsNullOrEmpty(result))
            {
                Log.LogError($"TestCase {testSuiteID} create relation failed.");
                return false;
            }
            else
            {
                /* in testing
                foreach(var item in testCase)
                {
                    Log.LogInformation(string.Format($"Test Case {item.CaseName} Start"));
                    if (item.TestStep.StepRepro.Count > 0)
                    {
                        // テストステップはhtmlエンコードしないといけない
                        item.TestStep.Value = EncodeHtmlString(item.TestStep);
                        json = JsonConvert.SerializeObject(item.TestStep);

                        result = await InvokeRestAPIPatch(json,
                            $"wit/workitems/{item.ID}?api-version=5.0-preview.3");
                        if (string.IsNullOrEmpty(result))
                        {
                            Log.LogError($"TestStep in {item.CaseName} register failed.");
                            return false;
                        }

                        var updateResult = JsonConvert.DeserializeObject<UpdateTestCaseRootobject>(result);
                    }
                }
                */
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
        static async Task<bool>CreateTestSuite(int testPlanID, TestSuite testSuite)
        {
            bool functionResult = true;
            var json = JsonConvert.SerializeObject(testSuite);
            Log.LogInformation(string.Format($"Plan id :{testPlanID}, Root Suite:{testSuite.Parent.ParentID} Test Suite {testSuite.Name} Start"));
            Log.LogInformation(json);
            var result = await InvokeRestAPIPost(json,
                string.Format($"testplan/Plans/{testPlanID}/suites/?api-version=5.0-preview.1"));
            if (string.IsNullOrEmpty(result))
            {
                functionResult = false;
                Log.LogError($"TestSuite {testSuite.Name} create failed.");
            }
            else
            {
                var resultObject = JsonConvert.DeserializeObject<TestSuiteResultRootobject>(result);
                testSuite.ID = resultObject.id;
            }
            Log.LogInformation(string.Format($"Test Case {testSuite.Name} end"));

            return functionResult;
        }

        /// <summary>
        /// Azure DevOpsへのREST API実行
        /// </summary>
        /// <returns></returns>
        static async Task<bool>RegisterTest2AzureDevOps()
        {
            bool functionResult = true;

            functionResult = await CreateTestPlan(TestPlan);
            if (!functionResult)
            {
                Log.LogError($"TestPlan {TestPlan.Name} create failed.");
                return functionResult;
            }

            //書き換える必要があるので、foreachではNG
            for(int i = 0;i < TestPlan.TestSuites.Count;i++)
            {
                TestPlan.TestSuites[i].Parent.ParentID = TestPlan.SuiteRootID;
                functionResult = await CreateTestSuite(TestPlan.ID, TestPlan.TestSuites[i]);
                if (!functionResult)
                {
                    Log.LogError($"TestSuite {TestPlan.TestSuites[i].Name} create failed.");
                    return functionResult;
                }
                else
                {
                    for(int j = 0; j < TestPlan.TestSuites[i].TestCases.Count; j++)
                    {
                        functionResult = await CreateTestCase(
                            TestPlan.ID,
                            TestPlan.TestSuites[i].ID,
                            TestPlan.TestSuites[i].TestCases.ToArray());
                        if (!functionResult)
                        {
                            Log.LogError($"TestCase No.{TestPlan.TestSuites[i].ID} is {TestPlan.TestSuites[i].Name} test case {TestPlan.TestSuites[i].TestCases[j].CaseName} create failed.");
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
                // Test Planはタブの名前
                TestPlan = new TestPlan();
                TestPlan.Name = sheet.Name;

                var ws = wsheetPart.Worksheet;
                List<string[]> testAssets = new List<string[]>();
                int count = 0;
                int rownum = 0;
                char[] columnID = { 'A', 'B', 'C', 'D', };
                foreach (var row in ws.Descendants<Row>().Skip(1))
                {
                    var cells = row.Elements<Cell>().ToArray();
                    var cellStrings = new string[4];

                    // セルに値が入っていないレコードが出れば処理終了(Excelの空行対応)。
                    if (cells.All(x => x.DataType.HasValue == false))
                    {
                        Log.LogInformation("All Data proceeded.");
                        break;
                    }

                    //各行をとりあえず文字列化して、配列＋リストに突っ込む。
                    for (int i = 0; i < cells.Length; i++)
                    {
                        var currentColunm = cells[i].CellReference.Value.Substring(0, 1).ToCharArray();
                        rownum = Array.FindIndex(columnID, x => 
                        {
                            if(x == currentColunm[0])
                            {
                                return true;
                            }
                            return false;
                        });
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
                                    Log.LogError($"Invalid Type in row {count}, cell {i}");
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
                    count++;
                    testAssets.Add(cellStrings);
                }

                //終わったらPlan配下を作る
                CreateTestPlanfromString(testAssets);
            }

            return functionResult;
        }

        private static int GetFillStatus(string[] testAssets)
        {
            int rowStatus = 2; // 0:Test Suite, 1:Test Case, 2:Test Step
            bool[] columnExist = new bool[4];
            for(int i =0;i < testAssets.Length; i++)
            {
                columnExist[i] = string.IsNullOrEmpty(testAssets[i]);
            }

            if (columnExist[0] == false && columnExist[1] == false &&
                columnExist[2] == false && columnExist[3] == false)
            {
                return 0;
            }
            if(columnExist[0] == true && columnExist[1] == false && 
                columnExist[2] == false && columnExist[3] == false)
            {
                return 1;
            }
            return rowStatus;
        }

        /// <summary>
        /// Excelデータの文字配列からテストケースの塊を取り出す。
        /// </summary>
        /// <param name="testAssets"></param>
        private static void CreateTestPlanfromString(IList<string[]> testAssets)
        {
            var caseTuple = new List<Tuple<int, int>>();
            var suiteTuple = new List<Tuple<int, int>>();
            int rowStatus = 0; // 0:Test Suite, 1:Test Case, 2:Test Step

            int suiteStartPosition = 0;
            int caseStartPosition = 0;
            // 0行目は必ずスィートなので,1行目から
            for (int i = 1; i < testAssets.Count; i++)
            {
                rowStatus = GetFillStatus(testAssets[i]);
                switch (rowStatus)
                {
                    case 0:
                        suiteTuple.Add(
                            new Tuple<int, int>(suiteStartPosition, i - 1));
                        caseTuple.Add(
                            new Tuple<int, int>(caseStartPosition, i - 1));
                        suiteStartPosition = i;
                        break;
                    case 1:
                        caseTuple.Add(
                                new Tuple<int, int>(caseStartPosition, i - 1));
                        caseStartPosition = i;
                        break;
                }
            }
            //終了するときクローズする
            suiteTuple.Add(
                new Tuple<int, int>(suiteStartPosition, testAssets.Count - 1));
            caseTuple.Add(
                new Tuple<int, int>(caseTuple[caseTuple.Count - 1].Item2 + 1, testAssets.Count - 1));

            //テストスィート単位で処理
            for(int j = 0;j < suiteTuple.Count; j++)
            {
                TestPlan.TestSuites.Add(
                    new TestSuite
                    {
                        Name = testAssets[suiteTuple[j].Item1][0]
                    });
                for (int i = 0;i < caseTuple.Count; i++)
                {
                    TestPlan.TestSuites[j].TestCases.Add(new TestCase
                    {
                        CaseName = testAssets[caseTuple[i].Item1][1]
                    });
                    if (caseTuple[i].Item1 >= suiteTuple[j].Item1 && 
                        caseTuple[i].Item2 <= suiteTuple[j].Item2)
                    {
                        IEnumerable<string[]> stepLists;
                        if(i == 0)
                        {
                            stepLists = testAssets.Take(caseTuple[i].Item2 - caseTuple[i].Item1 + 1).Select(x => x);
                        }
                        else
                        {
                            stepLists = testAssets.Skip(caseTuple[i].Item1)
                                .Take(caseTuple[i].Item2 - caseTuple[i].Item1 + 1)
                                .Select(x => x);
                        }
                        // 手順
                        TestPlan.TestSuites[j].TestCases[TestPlan.TestSuites[j].TestCases.Count - 1].TestStep.StepRepro = 
                            ConvertHolirontalToVertical(stepLists, 2);
                        // 結果
                        TestPlan.TestSuites[j].TestCases[TestPlan.TestSuites[j].TestCases.Count - 1].TestStep.StepExpected =
                            ConvertHolirontalToVertical(stepLists, 3);
                    }
                }
            }
        }

        //Excelの縦持ちになっているデータを横持ちにする
        static List<string>ConvertHolirontalToVertical(IEnumerable<string[]> source, int index)
        {
            var resultData = new List<string>();
            foreach (var item in source)
            {
                resultData.Add(item[index]);
            }

            return resultData;
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
}
