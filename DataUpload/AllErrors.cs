using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DataUpload
{
    public class AllErrors
    {
        public List<ErrorTemplates> Error { get;set; }
        public List<WarningTemplates>Warning { get; set; }
    }
}