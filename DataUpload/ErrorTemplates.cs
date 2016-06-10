using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DataUpload
{
    public class ErrorTemplates
    {
        public string ErrorType { get; set; }
        public string Field_1 { get; set; }
        public string Field_2 { get; set; }
        public int Row { get; set; }
        public string ErrorComments { get; set; }
        public List<string> IncorrectHeaderList { get; set; }

    }
}