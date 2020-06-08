using System;
using System.Collections.Generic;
using System.Text;

namespace SyncTestCase.Models
{

    public class UpdateTestCaseRootobject
    {
        public TestOperation[] TestCaseProperty { get; set; }
    }

    public class TestOperation
    {
        public string op { get; set; }
        public string path { get; set; }
        public object value { get; set; }

        public TestOperation()
        {
            path = "/fields/Microsoft.VSTS.TCM.Steps";
        }
    }
}
