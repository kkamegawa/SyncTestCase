using Microsoft.VisualStudio.TestTools.UnitTesting;
using SyncTestCase;
using SyncTestCase.Models;
using System.Text;
using System;
using System.IO;
using Newtonsoft.Json;
using System.Net.Http;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace UnitTestSyncTestCase
{
    [TestClass]
    public class UnitTestSyncTestCase
    {
        Uri TestExcelFile;
        string PAT;
        TestPlanRootobject TestPlan;
        TestCaseRootobject TestCase;
        string ExcelUrl = string.Empty;
        string AzureDevOpsEndpoing = string.Empty;

        [TestInitialize]
        public void TestInit()
        {
            if(string.IsNullOrEmpty(Environment.GetEnvironmentVariable("EXCEL_TEST_URL")) == true)
            {
                Assert.Fail("Environment variable \"EXCEL_TEST_URL\" is null");
            }
            if(string.IsNullOrEmpty(Environment.GetEnvironmentVariable("TF_BUILD")) == true)
            {
                PAT = File.ReadAllText(@"apikeys.txt");
            }
            else
            {
                TestExcelFile = new Uri(ExcelUrl);
                PAT = Environment.GetEnvironmentVariable("AZUREDEVOPS_PAT");
            }

            if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("AZUREDEVOPS_ENDPOING_URL")) == true)
            {
                Assert.Fail("Environment variable \"AZUREDEVOPS_ENDPOING_URL\" is null");
            }
            SyncTestCase.SyncTestCase.HttpEndpoint = AzureDevOpsEndpoing;
            PAT = Convert.ToBase64String(ASCIIEncoding.ASCII.GetBytes(string.Format("{0}:{1}", "", PAT)));
            SyncTestCase.SyncTestCase.Token = PAT;
        }

        [TestMethod]
        [Priority(0)]
        public async Task ParseExcelTest()
        {
            var result = await SyncTestCase.SyncTestCase.ParseExcelFile(TestExcelFile);
            Assert.IsTrue(result == true, "読み込み失敗");
        }

        [Priority(0)]
        [TestMethod]
        public void EncodeHtmlStringTest()
        {
            TestStep testStep = new TestStep();
            var expected = "<steps id=\\\"0\\\" last=\\\"4\\\"><step id=\\\"2\\\" type=\\\"ValidateStep\\\"><parameterizedString isformatted=\\\"true\\\">&lt;DIV&gt;&lt;P&gt;1st step&amp;nbsp;&lt;/P&gt;&lt;/DIV&gt;</parameterizedString><parameterizedString isformatted=\\\"true\\\">&lt;P&gt;expect 1st&amp;nbsp;&lt;/P&gt;</parameterizedString><description/></step><step id=\\\"3\\\" type=\\\"ValidateStep\\\"><parameterizedString isformatted=\\\"true\\\">&lt;DIV&gt;&lt;P&gt;2nd step&amp;nbsp;&lt;/P&gt;&lt;/DIV&gt;</parameterizedString><parameterizedString isformatted=\\\"true\\\">&lt;P&gt;expect 2nd&amp;nbsp;&lt;/P&gt;</parameterizedString><description/></step><step id=\\\"4\\\" type=\\\"ValidateStep\\\"><parameterizedString isformatted=\\\"true\\\">&lt;DIV&gt;&lt;P&gt;3rd step&amp;nbsp;&lt;/P&gt;&lt;/DIV&gt;</parameterizedString><parameterizedString isformatted=\\\"true\\\">&lt;P&gt;expect 3rd&amp;nbsp;&lt;/P&gt;</parameterizedString><description/></step></steps>";
            testStep.StepRepro = new List<string> { "1st step", "2nd step", "3rd step" };
            testStep.StepExpected = new List<string> {"expect 1st", "expect 2nd", "expect 3rd" };
            var result = SyncTestCase.SyncTestCase.EncodeHtmlString(testStep);

            Assert.IsTrue(string.Compare(result, expected) == 0, "テストステップの期待文字列と一致しませんでした");

        }

        [TestMethod]
        [Priority(2)]
        public async Task TestCreateTestPlan()
        {
            var testPlan = new SyncTestCase.Models.TestPlan();
            testPlan.Name = "API Test Plan";
            testPlan.StartDate = new DateTime(2018, 11, 1);
            testPlan.EndDate = new DateTime(2018, 11, 30);
            var json = JsonConvert.SerializeObject(testPlan);

            var result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json, "test/plans?api-version=5.0-preview.2");
            TestPlan = JsonConvert.DeserializeObject<TestPlanRootobject>(result);
        }

        [TestMethod]
        [Priority(2)]
        public async Task TestCreateTestSuite()
        {
            var testSuite = new SyncTestCase.Models.TestSuite();
            testSuite.Name = "API Test Suite";
            testSuite.ID = TestPlan.id;
            var json = JsonConvert.SerializeObject(testSuite);

            var result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json, 
                string.Format($"testplan/Plans/{TestPlan.id}/suites/?api-version=5.0-preview.1"));
            var resultObject = JsonConvert.DeserializeObject<TestSuiteRootObject>(result);
        }

        [TestMethod]
        [Priority(2)]
        public async Task TestCreateTestCase()
        {
            var testPlan = new SyncTestCase.Models.TestPlan();
            testPlan.Name = "API Test Plan Root";
            testPlan.StartDate = new DateTime(2018, 11, 1);
            testPlan.EndDate = new DateTime(2018, 11, 30);
            var json = JsonConvert.SerializeObject(testPlan);

            var result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json, "test/plans?api-version=5.0-preview.2");
            var testPlanResult = JsonConvert.DeserializeObject<TestPlanRootobject>(result);

            var testSuite = new SyncTestCase.Models.TestSuite();
            testSuite.Name = "API Test Suite";
            testSuite.ID = TestPlan.id;
            json = JsonConvert.SerializeObject(testSuite);

            result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json,
                string.Format($"test/Plans/{TestPlan.id}/suites/?api-version=5.0-preview.3"));
            var testSuiteResult = JsonConvert.DeserializeObject<TestSuiteRootObject>(result);

            //テストケース作成
            var testCase = new SyncTestCase.Models.TestCase[1] 
            {
                new TestCase{CaseName = "Case 1", Operation = "add"},
            };
            
            json = JsonConvert.SerializeObject(testCase);

            result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json, 
                "wit/workitems/$Test%20Case?api-version=5.0-preview.3");
            TestCase = JsonConvert.DeserializeObject<TestCaseRootobject>(result);

            //関連付け
            result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json,
                string.Format($"test/Plans/{testPlanResult.id}/suites/{testSuiteResult.value[0].id}/testcases/{TestCase.id}?api-version=5.0-preview.3"));

            // ステップ登録
            var testStep = new SyncTestCase.Models.TestStep();
            testStep.StepRepro = new List<string> { "ブラウザ起動", "ログイン", "About表示", "終了" };
            // HTMLにする

            json = JsonConvert.SerializeObject(testStep);

            result = await SyncTestCase.SyncTestCase.InvokeRestAPIPatch(json,
                $"wit/workitems/{TestCase.id}?api-version=5.0-preview.3");
            var updateResult = JsonConvert.DeserializeObject<UpdateTestCaseRootobject>(result);

        }

    }
}
