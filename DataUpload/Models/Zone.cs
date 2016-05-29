using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DataUpload.Models
{
    public class Zone
    {
        public int ID { get; set; }
        public string ZSMName { get; set; }
        public string ZSMEmailId { get; set; }
        public string SecondaryEmailId { get; set; }
        public string ZoneName { get; set; }
    }
}