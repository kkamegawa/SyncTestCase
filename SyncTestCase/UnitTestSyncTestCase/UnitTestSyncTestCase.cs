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
        List<int> RemoveTestCaseID;
        TestPlanRootobject TestPlan;
        TestCaseRootobject TestCase;
        string ExcelUrl = string.Empty;
        string AzureDevOpsEndpoing = string.Empty;

        [TestInitialize]
        public void TestInit()
        {
            RemoveTestCaseID = new List<int>();
            if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable("EXCEL_TEST_URL")) == true)
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

        [TestCleanup]
        public void TestCleanup()
        {
            if (TestPlan != null)
            {
                foreach (var testID in RemoveTestCaseID)
                {
                    var result = SyncTestCase.SyncTestCase.InvokeRestAPIDelete(
                        string.Format($"testplan/plans/{testID}?api-version=6.0-preview.1"));
                }
            }
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
            var expected = "<steps id=\"0\"><step id=\"2\" type=\"ValidateStep\"><parameterizedString isformatted=\"true\">&lt;DIV&gt;&lt;P&gt;1st step&amp;nbsp;&lt;/P&gt;&lt;/DIV&gt;</parameterizedString><parameterizedString isformatted=\"true\">&lt;P&gt;expect 1st&amp;nbsp;&lt;/P&gt;</parameterizedString><description/></step><step id=\"3\" type=\"ValidateStep\"><parameterizedString isformatted=\"true\">&lt;DIV&gt;&lt;P&gt;2nd step&amp;nbsp;&lt;/P&gt;&lt;/DIV&gt;</parameterizedString><parameterizedString isformatted=\"true\">&lt;P&gt;expect 2nd&amp;nbsp;&lt;/P&gt;</parameterizedString><description/></step><step id=\"4\" type=\"ValidateStep\"><parameterizedString isformatted=\"true\">&lt;DIV&gt;&lt;P&gt;3rd step&amp;nbsp;&lt;/P&gt;&lt;/DIV&gt;</parameterizedString><parameterizedString isformatted=\"true\">&lt;P&gt;expect 3rd&amp;nbsp;&lt;/P&gt;</parameterizedString><description/></step></steps>";
            testStep.StepRepro = new List<string> { "1st step", "2nd step", "3rd step" };
            testStep.StepExpected = new List<string> { "expect 1st", "expect 2nd", "expect 3rd" };
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

            var result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json, "testplan/plans?api-version=6.0-preview.1");
            Assert.IsTrue(string.IsNullOrEmpty(result), "response is null");
            TestPlan = JsonConvert.DeserializeObject<TestPlanRootobject>(result);
            Assert.IsTrue(TestPlan.id == 0, "REST API failed");
            RemoveTestCaseID.Add(TestPlan.id);
        }

        [TestMethod]
        [Priority(2)]
        public async Task TestCreateTestSuite()
        {
            var testPlan = new SyncTestCase.Models.TestPlan();
            testPlan.Name = "API Test Plan_TestSuite";
            testPlan.StartDate = new DateTime(2018, 11, 1);
            testPlan.EndDate = new DateTime(2018, 11, 30);
            var json = JsonConvert.SerializeObject(testPlan);

            var result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json, "testplan/plans?api-version=6.0-preview.1");
            TestPlan = JsonConvert.DeserializeObject<TestPlanRootobject>(result);

            RemoveTestCaseID.Add(TestPlan.id);

            var testSuite = new SyncTestCase.Models.TestSuite();
            testSuite.Name = "API Test Suite";
            testSuite.Parent.ParentID = TestPlan.rootSuite.id;
            json = JsonConvert.SerializeObject(testSuite);

            result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json,
                string.Format($"testplan/Plans/{TestPlan.id}/suites/?api-version=6.0-preview.1"));
            var testResult = JsonConvert.DeserializeObject<TestSuiteResultRootobject>(result);
            if (testResult.id == 0)
            {
                Assert.Fail("Response ID is zero");
            }
        }
        [TestMethod]
        [Priority(2)]
        public async Task TestWorkItemTest()
        {
            TestStep testStepObject = new TestStep();
            testStepObject.StepRepro = new List<string> { "1st step", "2nd step", "3rd step" };
            testStepObject.StepExpected = new List<string> { "expect 1st", "expect 2nd", "expect 3rd" };
            var steps = SyncTestCase.SyncTestCase.EncodeHtmlString(testStepObject);
            var testCase = new SyncTestCase.Models.TestCase[]
            {
                new TestCase{Value = "Test Case 1", Operation = "add", Path = "/fields/System.Title"},
                new TestCase{Value = steps, Operation = "add", Path = "/fields/Microsoft.VSTS.TCM.Steps"},
            };

            var json = JsonConvert.SerializeObject(testCase);

            var result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json,
                "wit/workitems/$Test%20Case?api-version=6.0-preview.3", "application/json-patch+json");
            TestCase = JsonConvert.DeserializeObject<TestCaseRootobject>(result);


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

            var result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json, "testplan/plans?api-version=6.0-preview.1");
            TestPlan = JsonConvert.DeserializeObject<TestPlanRootobject>(result);
            RemoveTestCaseID.Add(TestPlan.id);

            var testSuite = new SyncTestCase.Models.TestSuite();
            testSuite.Name = "API Test Suite";
            testSuite.Parent.ParentID = TestPlan.rootSuite.id;
            json = JsonConvert.SerializeObject(testSuite);

            result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json,
                string.Format($"testplan/Plans/{TestPlan.id}/suites/?api-version=6.0-preview.1"));
            var testSuiteResult = JsonConvert.DeserializeObject<TestSuiteResultRootobject>(result);

            //テストケース作成
            var testCase = new SyncTestCase.Models.TestCase[1]
            {
                new TestCase{Value = "Case 1", Operation = "add"},
            };

            json = JsonConvert.SerializeObject(testCase);

            result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json,
                "wit/workitems/$Test%20Case?api-version=6.0-preview.3", "application/json-patch+json");
            TestCase = JsonConvert.DeserializeObject<TestCaseRootobject>(result);

            //関連付け
            result = await SyncTestCase.SyncTestCase.InvokeRestAPIPost(json,
                string.Format($"test/Plans/{TestPlan.id}/suites/{testSuiteResult.id}/testcases/{TestCase.id}?api-version=6.0-preview.3"));

            // ステップ登録
            var testStep = new SyncTestCase.Models.TestStep();
            testStep.StepRepro = new List<string> { "ブラウザ起動", "ログイン", "About表示", "終了" };
            // HTMLにする

            json = JsonConvert.SerializeObject(testStep);

            result = await SyncTestCase.SyncTestCase.InvokeRestAPIPatch(json,
                $"wit/workitems/{TestCase.id}?api-version=6.0-preview.3");
            var updateResult = JsonConvert.DeserializeObject<UpdateTestCaseRootobject>(result);

        }


    }
}
