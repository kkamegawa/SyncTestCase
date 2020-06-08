using System;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Text;
using Newtonsoft.Json.Converters;

namespace SyncTestCase.Models
{


    public class TestPlanRootobject
    {
        public int id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
        public TestPlanRootProject project { get; set; }
        public Area area { get; set; }
        public DateTime startDate { get; set; }
        public DateTime endDate { get; set; }
        public string iteration { get; set; }
        public DateTime updatedDate { get; set; }
        public Updatedby updatedBy { get; set; }
        public Owner owner { get; set; }
        public int revision { get; set; }
        public string state { get; set; }
        public Rootsuite rootSuite { get; set; }
        public string clientUrl { get; set; }
    }

    public class TestPlanRootProject
    {
        public string id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
    }

    public class Area
    {
        public string id { get; set; }
        public string name { get; set; }
    }

    public class Updatedby
    {
        public string id { get; set; }
        public string displayName { get; set; }
        public string uniqueName { get; set; }
        public string url { get; set; }
        public string imageUrl { get; set; }
    }

    public class Owner
    {
        public string id { get; set; }
        public string displayName { get; set; }
        public string uniqueName { get; set; }
        public string url { get; set; }
        public string imageUrl { get; set; }
    }

    public class Rootsuite
    {
        public int id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
    }

    public class DateFormatConverter : IsoDateTimeConverter
    {
        public DateFormatConverter(string format)
        {
            DateTimeFormat = format;
        }
    }

    // Model for Test Plan
    public class TestPlan
    {
        [JsonProperty(PropertyName = "name")]
        public string Name { get; set; }

        [JsonConverter(typeof(DateFormatConverter), "yyyy-MM-dd")]
        [JsonProperty(PropertyName = "startDate")]
        public DateTime StartDate { get; set; }

        [JsonConverter(typeof(DateFormatConverter), "yyyy-MM-dd")]
        [JsonProperty(PropertyName = "endDate")]
        public DateTime EndDate { get; set; }

        [JsonProperty(PropertyName = "areaPath")]
        public string AreaPath { get; set; }

        [JsonProperty(PropertyName = "iteration")]
        public string Iteration { get; set; }

        [JsonProperty(PropertyName = "state")]
        public string State { get; set; }

        [JsonIgnore]
        public int ID { get; set; }

        [JsonIgnore]
        public int SuiteRootID { get; set; }

        [JsonIgnore]
        public List<TestSuite> TestSuites { get; set; }

        public TestPlan()
        {
            StartDate = DateTime.Now;
            EndDate = DateTime.Now.AddDays(14);
            TestSuites = new List<TestSuite>();
        }
    }

    public struct ParentSuite
    {
        [JsonProperty(PropertyName = "id")]
        public int ParentID { get; set; }
    }
    /// <summary>
    /// Model for Test Suite
    /// </summary>
    public class TestSuite
    {
        [JsonProperty(PropertyName = "name")]
        public string Name { get; set; }
        [JsonProperty(PropertyName = "suiteType")]
        public const string SuiteType = "StaticTestSuite"; // This is prototype, Static base only.
        [JsonProperty(PropertyName = "parentSuite")]
        public ParentSuite Parent;
        
        [JsonIgnore]
        public int ID { get; set; } // Test Suite ID
        [JsonIgnore]
        public List<TestCase> TestCases { get; set; }
        public TestSuite()
        {
            TestCases = new List<TestCase>();
            Parent = new ParentSuite();
        }
    }

    /// <summary>
    /// Model for Test Case
    /// </summary>
    public class TestCase
    {
        [JsonProperty(PropertyName = "op")]
        public string Operation { get; set; }

        [JsonProperty(PropertyName = "path")]
        public string Path = "/fields/System.Title";

        [JsonProperty(PropertyName = "value")]
        public string Value { get; set; }
        [JsonIgnore]
        public int ID { get; set; }  // Test Case ID
        [JsonIgnore]
        public TestStep TestStep { get; set; }
        public TestCase()
        {
            TestStep = new TestStep();
            Operation = "add";
        }
    }

    public class TestStep
    {
        [JsonProperty(PropertyName = "op")]
        public string Operation = "add"; // operation

        [JsonProperty(PropertyName = "path")]
        public string Path = "/fields/Microsoft.VSTS.TCM.Steps"; // Test Case Path

        [JsonProperty(PropertyName = "value")]
        public string Value;

        [JsonIgnore]
        public List<string> StepRepro { get; set; } // Test Step
        [JsonIgnore]
        public List<string> StepExpected { get; set; } // Test to expected result
        public TestStep()
        {
            StepRepro = new List<string>();
            StepExpected = new List<string>();
        }
    }
}
