using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DataUpload
{
    public class ErrorTemplates
    {
        public string MappingErrorType { get; set; }
        public string Field_1 { get; set; }
        public string Field_2 { get; set; }
        public int Row { get; set; }
        public string IncorrectHeaders { get; set; }
        public string EmptyFinalBeatName { get; set; }
        public string EmptyESM { get; set; }
        public string HierarchyBeak { get; set; }
        public string PhoneError { get; set; }
    }
}