using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using CsvHelper;
using CsvHelper.Configuration;


namespace DataUpload
{
    public class FileHeaders
    {
       

            public string NSM { get; set; }
            public string NSMZone { get; set; }
            public string NSMEmailId { get; set; }
            public string NSMSecondaryEmailId { get; set; }
            public string ZSM { get; set; }
            public string ZSMEmailId { get; set; }
            public string ZSMZone { get; set; }
            public string ZSMSecondaryEmailId { get; set; }
            public string RSM { get; set; }
            public string RSMZone { get; set; }
            public string RSMEmailId { get; set; }
            public string RSMSecondaryEmailId { get; set; }
            public string ASM { get; set; }
            public string ASMZone { get; set; }
            public string ASMEmailId { get; set; }
            public string ASMSecondaryEmailId { get; set; }
            public string ESM { get; set; }
            public string ESMZone { get; set; }
            public string ESMEmailId { get; set; }
            public string ESMSecondaryEmailId { get; set; }
            public string ESMContactNumber { get; set; }
            public string ESMHQ { get; set; }
            public string FinalBeatName { get; set; }
            public string ESMErpId { get; set; }
            public string BeatErpId { get; set; }
            public string BeatDistrict { get; set; }
            public string BeatState { get; set; }
            public string BeatZone { get; set; }
            public string DistributorName { get; set; }
            public string DistributorLocation { get; set; }
            public string DistributorErpId { get; set; }
            public string DistributorEmailId { get; set; }

        
    }
    public sealed class MyClass : CsvClassMap<FileHeaders>
    {
        public MyClass()
        {
            Map(x => x.NSM).Name("NSM");
            Map(x => x.NSMZone).Name("NSMZone");
            Map(x => x.NSMEmailId).Name("NSMEmailId");
            Map(x => x.NSMSecondaryEmailId).Name("NSMSecondaryEmailId");
            Map(x => x.ZSM).Name("ZSM");
            Map(x => x.ZSMZone).Name("ZSMZone");
            Map(x => x.ZSMEmailId).Name("ZSMEmailId");
            Map(x => x.ZSMSecondaryEmailId).Name("ZSMSecondaryEmailId");
            Map(x => x.RSM).Name("RSM");
            Map(x => x.RSMZone).Name("RSMZone");
            Map(x => x.RSMEmailId).Name("RSMEmailId");
            Map(x => x.RSMSecondaryEmailId).Name("RSMSecondaryEmailId");
            Map(x => x.ASM).Name("ASM");
            Map(x => x.ASMZone).Name("ASMZone");
            Map(x => x.ASMEmailId).Name("ASMEmailId");
            Map(x => x.ASMSecondaryEmailId).Name("ASMSecondaryEmailId");
            Map(x => x.ESM).Name("ESM");
            Map(x => x.ESMZone).Name("ESMZone");
            Map(x => x.ESMEmailId).Name("ESMEmailId");
            Map(x => x.ESMSecondaryEmailId).Name("ESMSecondaryEmailId");
            Map(x => x.ESMContactNumber).Name("ESMContactNumber");
            Map(x => x.ESMHQ).Name("ESMHQ");
            Map(x => x.ESMErpId).Name("ESMErpId");
            Map(x => x.FinalBeatName).Name("FinalaBeatName");
            Map(x => x.BeatZone).Name("BeatZone");
            Map(x => x.BeatDistrict).Name("BeatDistrict");
            Map(x=>x.BeatState).Name("BeatState");
            Map(x => x.BeatErpId).Name("BeatErpId");
            Map(x => x.DistributorName).Name("DistriutorName");
            Map(x => x.DistributorEmailId).Name("DistributorEmailId");
            Map(x => x.DistributorLocation).Name("DistributorLocation");
            Map(x=>x.DistributorErpId).Name("DistribitorErpId");

        }
    }
}