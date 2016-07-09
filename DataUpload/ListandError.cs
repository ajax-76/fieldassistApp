using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DataUpload
{
    public class ListandError
    {
        public List<ErrorTemplates> error { get; set; }
        public List<FileHeadersBeatHierarchy> listBeatHierarchy { get; set; }
        public List<FileHeadersBeatPlanAddition> listBeatPlan { get; set; }
        public List<FileHeadersLocationAddtion> listLocation { get; set; }
        public List<FileHeadersProductAddition> listProductAddition { get; set; }
    }
}