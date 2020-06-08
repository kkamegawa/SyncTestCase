using System;
using System.Collections.Generic;
using System.Text;

namespace SyncTestCase.Models
{


    public class TestSuiteResultRootobject
    {
        public int id { get; set; }
        public int revision { get; set; }
        public Project project { get; set; }
        public Lastupdatedby lastUpdatedBy { get; set; }
        public DateTime lastUpdatedDate { get; set; }
        public Plan plan { get; set; }
        public _Links1 _links { get; set; }
        public string suiteType { get; set; }
        public string name { get; set; }
        public Parentsuite parentSuite { get; set; }
        public bool inheritDefaultConfigurations { get; set; }
        public Defaultconfiguration[] defaultConfigurations { get; set; }
    }

    public class Project
    {
        public string id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
        public string state { get; set; }
        public string visibility { get; set; }
    }

    public class Lastupdatedby
    {
        public string displayName { get; set; }
        public string url { get; set; }
        public _Links _links { get; set; }
        public string id { get; set; }
        public string uniqueName { get; set; }
        public string imageUrl { get; set; }
        public string descriptor { get; set; }
    }

    public class _Links
    {
        public Avatar avatar { get; set; }
    }

    public class Avatar
    {
        public string href { get; set; }
    }

    public class Plan
    {
        public int id { get; set; }
        public string name { get; set; }
    }

    public class _Links1
    {
        public _Self _self { get; set; }
        public Testcases testCases { get; set; }
        public Testpoints testPoints { get; set; }
    }

    public class _Self
    {
        public string href { get; set; }
    }

    public class Testcases
    {
        public string href { get; set; }
    }

    public class Testpoints
    {
        public string href { get; set; }
    }

    public class Parentsuite
    {
        public int id { get; set; }
        public string name { get; set; }
    }

    public class Defaultconfiguration
    {
        public int id { get; set; }
        public string name { get; set; }
    }



    public class TestSuiteRootObject
    {
        public Value[] value { get; set; }
        public int count { get; set; }
    }

    public class Value
    {
        public int id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
        public Project project { get; set; }
        public TestSuiteRootPlan plan { get; set; }
        public Parent parent { get; set; }
        public int revision { get; set; }
        public int testCaseCount { get; set; }
        public string suiteType { get; set; }
        public string testCasesUrl { get; set; }
        public bool inheritDefaultConfigurations { get; set; }
        public string state { get; set; }
        public TestSuiteRootLastupdatedby lastUpdatedBy { get; set; }
        public DateTime lastUpdatedDate { get; set; }
    }

    public class TestSuiteRootPlan
    {
        public int id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
    }

    public class Parent
    {
        public int id { get; set; }
        public string name { get; set; }
        public string url { get; set; }
    }

    public class TestSuiteRootLastupdatedby
    {
        public string id { get; set; }
        public string displayName { get; set; }
        public string uniqueName { get; set; }
        public string url { get; set; }
        public string imageUrl { get; set; }
    }
}
