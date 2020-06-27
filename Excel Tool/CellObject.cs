using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace Excel_Tool
{
    
    class CellObject
    {
        public string Param1 { get; set; }

        public string Param2 { get; set; }

        public CellObject(string param1, string param2)
        {
            Param1 = param1;
            Param2 = param2;
        }

        public CellObject()
        {
        }
    }
}
