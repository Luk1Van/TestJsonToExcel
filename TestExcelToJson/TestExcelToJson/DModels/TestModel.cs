using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExcelToJson.DModels
{
    public class TestModel
    {
        public class Policy
        {
            public Driver driver { get; set; }
        }

        public class Driver
        {
            public string driverAge {get; set;}
            public string name { get; set; }
        }
    }
}
