using System;
using System.Collections.Generic;
using System.Text;

namespace SyncTestCase.Models
{

    public class TestCaseRelationRootobject
    {
        public TestCaseValue[] value { get; set; }
        public int count { get; set; }
    }

    public class TestCaseValue
    {
        public Testcase testCase { get; set; }
        public Pointassignment[] pointAssignments { get; set; }
    }

    public class Testcase
    {
        public int id { get; set; }
        public string url { get; set; }
        public string webUrl { get; set; }
    }

    public class Pointassignment
    {
        public Configuration configuration { get; set; }
        public Tester tester { get; set; }
    }

    public class Configuration
    {
        public int id { get; set; }
        public string name { get; set; }
    }

    public class Tester
    {
        public int id { get; set; }
        public string displayName { get; set; }
        public string uniqueName { get; set; }
        public string url { get; set; }
        public string imageUrl { get; set; }
    }
}
