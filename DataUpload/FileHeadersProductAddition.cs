using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DataUpload
{
    public class FileHeadersProductAddition
    {
        public string PrimaryCategory { get; set; }
        public string SecondaryCategory { get; set; }
        public string Product { get; set; }
        public string Variant { get; set; }
        public string Price { get; set; }
        public string Unit { get; set; }
        public string DisplayCategory { get; set; }
        public string Image { get; set; }
        public string Description { get; set; }
        public string StandardUnitConversionFactor { get; set; }
        public string StandardUnit { get; set; }
        public string ProductCode { get; set; }
        public string VariantCode { get; set; }
        public string ProductCategory { get; set; }
        public int Row { get; set; }
       
    }
}