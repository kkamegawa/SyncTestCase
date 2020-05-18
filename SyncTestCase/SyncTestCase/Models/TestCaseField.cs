using System;
using System.Collections.Generic;
using System.Text;

namespace SyncTestCase.Models
{

    public class UpdateTestCaseRootobject
    {
        public Class1[] Property1 { get; set; }
    }

    public class Class1
    {
        public string op { get; set; }
        public string path { get; set; }
        public object value { get; set; }
    }
}
