using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OfficeOpenXml;
using DataUpload.Controllers;
namespace DataUpload
{
    public class ValidationChecker
    {
        public List<ErrorTemplates> Checker(ExcelWorksheet sheet, List<FileHeadersBeatHierarchy> list1,List<FileHeadersLocationAddtion> list2,List<FileHeadersProductAddition>list3,List<FileHeadersBeatPlanAddition> list4)
        {
            List<ErrorTemplates> newError = new List<ErrorTemplates>();
            columnIndex indexer = new columnIndex();
            MappingValidations checks = new MappingValidations();
            var file = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.Start.Row, sheet.Dimension.End.Column];

            for (int i = sheet.Dimension.Start.Column - 1; i <= sheet.Dimension.End.Column - 1; i++)                 //To find Empty cells.(algo will be updated)
            {
                var value = ((object[,])file.Value)[0, i];
                if (value != null)
                {
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("NSM").ToLower())
                    {
                        indexer.NSM = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("NSMZone").ToLower())
                    {
                        indexer.NSMZone = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("NSMEmailId").ToLower())
                    {
                        indexer.NSMEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("NSMSecondaryEmailId").ToLower())
                    {
                        indexer.NSMSecondaryEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("ZSM").ToLower())
                    {
                        indexer.ZSM = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("ZSMZone").ToLower())
                    {
                        indexer.ZSMZone = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("ZSMEmailId").ToLower())
                    {
                        indexer.ZSMEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("ZSMSecondaryEmailId").ToLower())
                    {
                        indexer.ZSMSecondaryEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("RSM").ToLower())
                    {
                        indexer.RSM = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "RSMZone".ToLower())
                    {
                        indexer.RSMZone = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "RSMEmailId".ToLower())
                    {
                        indexer.RSMEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "RSMSecondaryEmailId".ToLower())
                    {
                        indexer.RSMSecondaryEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ASM".ToLower())
                    {
                        indexer.ASM = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ASMZone".ToLower())
                    {
                        indexer.ASMZone = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ASMEmailId".ToLower())
                    {
                        indexer.ASMEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ASMSecondaryEmailId".ToLower())
                    {
                        indexer.ASMSecondaryEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESM".ToLower())
                    {
                        indexer.ESM = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMZone".ToLower())
                    {
                        indexer.ESMZone = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMEmailId".ToLower())
                    {
                        indexer.ESMEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMSecondaryEmailId".ToLower())
                    {
                        indexer.ESMSecondaryEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMContactNumber".ToLower())
                    {
                        indexer.ESMContactNumber = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMHQ".ToLower())
                    {
                        indexer.ESMHQ = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMErpId".ToLower())
                    {
                        indexer.ESMErpId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "FinalBeatName".ToLower())
                    {
                        indexer.FinalBeatName = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "BeatErpId".ToLower())
                    {
                        indexer.BeatErpId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "BeatDistrict".ToLower())
                    {
                        indexer.BeatDistrict = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "BeatState".ToLower())
                    {
                        indexer.BeatState = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "BeatZone".ToLower())
                    {
                        indexer.BeatZone = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "DistributorName".ToLower())
                    {
                        indexer.DistributorName = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "DistributorLocation".ToLower())
                    {
                        indexer.DistributorLocation = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "DistributorErpId".ToLower())
                    {
                        indexer.DistributorErpId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "DistributorEmailId".ToLower())
                    {
                        indexer.DistributorEmailId = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ShopName".ToLower())
                    {
                        indexer.ShopName = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ShopErpId".ToLower())
                    {
                        indexer.ShopName = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "PrimaryCategory".ToLower())
                    {
                        indexer.PrimaryCategory = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "SecondaryCategory".ToLower())
                    {
                        indexer.SecondaryCategory = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "Product".ToLower())
                    {
                        indexer.Product = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "Variant".ToLower())
                    {
                        indexer.Variant = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "Price".ToLower())
                    {
                        indexer.Price = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "Unit".ToLower())
                    {
                        indexer.Unit = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "DisplayCategory".ToLower())
                    {
                        indexer.DisplayCategory = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "Image".ToLower())
                    {
                        indexer.Image = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "Description".ToLower())
                    {
                        indexer.Description = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "StandardUnitConversionFactor".ToLower())
                    {
                        indexer.StandardUnitConversionFactor = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "StandardUnit".ToLower())
                    {
                        indexer.StandardUnit = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ProductCode".ToLower())
                    {
                        indexer.ProductCode = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "VariantCode".ToLower())
                    {
                        indexer.VariantCode = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ProductCategory".ToLower())
                    {
                        indexer.ProductCategory = i + 1;
                    }
                    if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "BeatDay".ToLower())
                    {
                        indexer.BeatDay = i + 1;
                    }
                }
            }
            //hierarchy break
            if (list1 != null)
            {
                foreach (var item in list1)
                {
                    if (item.NSM == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "NSM";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if (item.NSMEmailId == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "NSMEmaialId";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if (item.ZSM == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "ZSM";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if (item.ZSMEmailId == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "ZSMEmailId";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if (item.RSM == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "RSM";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if (item.RSMEmailId == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "RSMEmaiilId";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if (item.ASM == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "ASM";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if (item.ASMEmailId == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "ASMEmaiId";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if (item.ESM == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "ESM";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if (item.DistributorName == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "DistributorName";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if (item.ESMContactNumber == "")
                    {

                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "ESMContactNumber";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                }
                newError.AddRange(checks.HierarchyError(sheet, indexer.ESM, indexer.ASM, indexer.RSM, indexer.ZSM, indexer.NSM));
                //Phone Digits checking
                newError.AddRange(checks.CheckPhoneDigit(sheet, indexer.ESMContactNumber));
                //Relationship Check
                var query1 = list1.GroupBy(x => x.NSM, x => new Mapping { c1 = x.NSMZone, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query1, indexer.NSM, indexer.NSMZone, "NSM", "NSMZone"));
                //Nsm to Nsm EmailID
                var query2 = list1.GroupBy(x => x.NSM, x => new Mapping { c1 = x.NSMEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query2, indexer.NSM, indexer.NSMEmailId, "NSM", "NSMEmailId"));
                //Nsm to Secondary EmailID
                var query3 = list1.GroupBy(x => x.NSM, x => new Mapping { c1 = x.NSMSecondaryEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query3, indexer.NSM, indexer.NSMSecondaryEmailId, "NSM", "NSMSecondaryEmailId"));
                //Nsm to Zsm
                var query4 = list1.GroupBy(x => x.NSM, x => new Mapping { c1 = x.ZSM, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query4, indexer.NSM, indexer.ZSM, "NSM", "ZSM"));
                //Zsm to Zsm Zone
                var query5 = list1.GroupBy(x => x.ZSM, x => new Mapping { c1 = x.ZSMZone, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query5, indexer.ZSM, indexer.ZSMZone, "ZSM", "ZSMZone"));
                //Zsm to Zsm EmailId
                var query6 = list1.GroupBy(x => x.ZSM, x => new Mapping { c1 = x.ZSMEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query6, indexer.ZSM, indexer.ZSMEmailId, "ZSM", "ZSMEmailId"));
                //Zsm to Zsm Secondary Email ID
                var query7 = list1.GroupBy(x => x.ZSM, x => new Mapping { c1 = x.ZSMSecondaryEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query7, indexer.ZSM, indexer.ZSMSecondaryEmailId, "ZSM", "ZSMSecondaryEmailId"));
                //Zsm to Rsm
                var query8 = list1.GroupBy(x => x.ZSM, x => new Mapping { c1 = x.RSM, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query8, indexer.ZSM, indexer.RSM, "ZSM", "RSM"));
                //Rsm to Rsm Zone
                var query9 = list1.GroupBy(x => x.RSM, x => new Mapping { c1 = x.RSMZone, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query9, indexer.RSM, indexer.RSMZone, "RSM", "RSMZone"));
                //Rsm to Rsm EmailId
                var query10 = list1.GroupBy(x => x.RSM, x => new Mapping { c1 = x.RSMEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query10, indexer.RSM, indexer.RSMEmailId, "RSM", "RSMEmailId"));
                //Rsm to Rsm SecondaryEmailId
                var query11 = list1.GroupBy(x => x.RSM, x => new Mapping { c1 = x.RSMSecondaryEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query11, indexer.RSM, indexer.RSMSecondaryEmailId, "RSM", "RSMSecondaryEmailId"));
                //Rsm to Asm
                var query12 = list1.GroupBy(x => x.RSM, x => new Mapping { c1 = x.ASM, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query12, indexer.RSM, indexer.ASM, "RSM", "ASM"));
                //Asm to Asm Zone
                var query13 = list1.GroupBy(x => x.ASM, x => new Mapping { c1 = x.ASMZone, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query13, indexer.ASM, indexer.ASMZone, "ASM", "ASMZone"));
                //Asm to Asm EmailId
                var query14 = list1.GroupBy(x => x.ASM, x => new Mapping { c1 = x.ASMEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query14, indexer.ASM, indexer.ASMEmailId, "ASM", "ASMEmailId"));
                //Asm to Secondary EmailId
                var query15 = list1.GroupBy(x => x.ASM, x => new Mapping { c1 = x.ASMSecondaryEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query15, indexer.ASM, indexer.ASMSecondaryEmailId, "ASM", "ASMSecondaryEmailId"));
                //Asm to Esm
                var query16 = list1.GroupBy(x => x.ASM, x => new Mapping { c1 = x.ESM, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query16, indexer.ASM, indexer.ESM, "ASM", "ESM"));
                //Esm to Esm Zone
                var query17 = list1.GroupBy(x => x.ESM, x => new Mapping { c1 = x.ESMZone, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query17, indexer.ESM, indexer.ESMZone, "ESM", "ESMZone"));
                //Esm to Esm EmailId
                var query18 = list1.GroupBy(x => x.ESM, x => new Mapping { c1 = x.ESMEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query18, indexer.ESM, indexer.ESMEmailId, "ESM", "ESMEmailId"));
                //Esm to Esm Secondary EmailId
                var query19 = list1.GroupBy(x => x.ESM, x => new Mapping { c1 = x.ESMSecondaryEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query19, indexer.ESM, indexer.ESMSecondaryEmailId, "ESM", "ESMSecondaryEmailId"));
                //Esm to Esm HQ
                var query20 = list1.GroupBy(x => x.ESM, x => new Mapping { c1 = x.ESMHQ, row = x.Row }).ToList();
                newError.AddRange(checks.AttributeMapping(query20, indexer.ESM, indexer.ESMHQ, "ESM", "ESMHQ"));
                //ESM to ESMErpId
                var query21 = list1.GroupBy(x => x.ESM, x => new Mapping { c1 = x.ESMErpId, row = x.Row }).ToList();
                newError.AddRange(checks.One2OneValidationCheck(query21, indexer.ESM, indexer.ESMErpId, "ESM", "ESMErpId"));
                //ESM to ESMContact Number
                var query22 = list1.GroupBy(x => x.ESM, x => new Mapping { c1 = x.ESMContactNumber, row = x.Row }).ToList();
                newError.AddRange(checks.One2OneValidationCheck(query22, indexer.ESM, indexer.ESMContactNumber, "ESM", "ESMContactNumber"));
                //BeatDistrict to FinalBeatName
                var query23 = list1.GroupBy(x => x.BeatDistrict, x => new Mapping { c1 = x.FinalBeatName, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query23, indexer.BeatDistrict, indexer.FinalBeatName, "BeatDistrict", "FinalBeatName"));
                //BeatState FinalBeatName
                var query24 = list1.GroupBy(x => x.BeatState, x => new Mapping { c1 = x.FinalBeatName, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query24, indexer.BeatState, indexer.FinalBeatName, "BeatState", "FinalBeatName"));
                //BeatZone to FinalBeatName
                var query25 = list1.GroupBy(x => x.BeatZone, x => new Mapping { c1 = x.FinalBeatName, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query25, indexer.BeatZone, indexer.FinalBeatName, "BeatZone", "FinalBeatName"));
                //BeatName to BeatErpId
                var query_25 = list1.GroupBy(x => x.FinalBeatName, x => new Mapping { c1 = x.BeatErpId, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query_25, indexer.FinalBeatName, indexer.BeatErpId, "FinalBeatName", "BeatErpId"));
                //FinalBeatName to DistributorName
                var query26 = list1.GroupBy(x => x.DistributorName, x => new Mapping { c1 = x.FinalBeatName, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query26, indexer.DistributorName, indexer.FinalBeatName, "DistributorName", "FinalBeatName"));
                //FinalBeatName to DistributorLocation
                var query27 = list1.GroupBy(x => x.FinalBeatName, x => new Mapping { c1 = x.DistributorLocation, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query27, indexer.DistributorLocation, indexer.FinalBeatName, "DistributorLocation", "FinalBeatName"));
                //FinalBeatName to DistributorLocation
                var query28 = list1.GroupBy(x => x.FinalBeatName, x => new Mapping { c1 = x.DistributorEmailId, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query28, indexer.DistributorEmailId, indexer.FinalBeatName, "DistributorEmailId", "FinalBeatName"));
                //DistributorName to DistributorErpId
                var query29 = list1.GroupBy(x => x.DistributorName, x => new Mapping { c1 = x.DistributorErpId, row = x.Row }).ToList();
                newError.AddRange(checks.One2OneValidationCheck(query29, indexer.DistributorName, indexer.DistributorErpId, "DistributorName", "DistributorErpId"));
                //Emails Check
                newError.AddRange(checks.EmailCheck(sheet, indexer.NSMEmailId, "NSMEmailId"));
                newError.AddRange(checks.EmailCheck(sheet, indexer.NSMSecondaryEmailId, "NSMSecondaryEmailId"));
                newError.AddRange(checks.EmailCheck(sheet, indexer.ZSMEmailId, "ZSMEmailId"));
                newError.AddRange(checks.EmailCheck(sheet, indexer.ZSMSecondaryEmailId, "ZSMSecondaryEmailId"));
                newError.AddRange(checks.EmailCheck(sheet, indexer.RSMEmailId, "RSMEmailId"));
                newError.AddRange(checks.EmailCheck(sheet, indexer.ZSMSecondaryEmailId, "RSMSecondaryEmailId"));
                newError.AddRange(checks.EmailCheck(sheet, indexer.ASMEmailId, "ASMEmailId"));
                newError.AddRange(checks.EmailCheck(sheet, indexer.ASMSecondaryEmailId, "ASMSecondaryEmailId"));
                newError.AddRange(checks.EmailCheck(sheet, indexer.ESMEmailId, "ESMEmailId"));
                newError.AddRange(checks.EmailCheck(sheet, indexer.ESMSecondaryEmailId, "ESMSecondaryEmailId"));
                newError.AddRange(checks.CheckStateAndDistrict(sheet, indexer.BeatState, indexer.BeatDistrict));
                //Unique ErpId
                Random rnd = new Random();
                var query30 = list1.Select(x => new Mapping { c1 = x.BeatErpId ?? rnd.Next(0,100000).ToString(), row = x.Row }).GroupBy(y => y.c1).Where(y => y.Count() > 1).ToList();
                newError.AddRange(checks.Unique(query30, indexer.BeatErpId, "BeatErpId"));
                var query31 = list1.Select(x => new Mapping { c1 = x.ESMErpId ?? rnd.Next(0, 100000).ToString(), row = x.Row }).GroupBy(y => y.c1).Where(y => y.Count() > 1).ToList();
                newError.AddRange(checks.Unique(query31, indexer.ESMErpId, "ESMErpId"));
            }

            //Checks For Location Addition
            if (list2 != null)
            {
                foreach(var item in list2)
                {
                    if(item.FinalBeatName=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "FinalBeatNamae";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.ShopName=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "ShopName";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.MarketName=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "MarketName";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.City=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "City";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.State=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "State";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.Country=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "Country";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    

                }
                Random rnd = new Random();
                //     var query= list1.GroupBy(x => x.DistributorName, x => new Mapping { c1 = x.DistributorErpId, row = x.Row }).ToList();
                var query31 = list2.Select(x => new Mapping { c1 = x.ShopCode ?? rnd.Next(0, 100000).ToString(), row = x.Row }).GroupBy(y => y.c1).Where(y => y.Count() > 1).ToList();
                //  newError.AddRange(checks.One2OneValidationCheck(query, indexer.ShopName, indexer.FinalBeatName, "ShopName", "ShopErpId"));
                newError.AddRange(checks.Unique(query31, indexer.ShopErpId, "ShopErpId"));
                List<Mapping> mapList = new List<Mapping>();
                foreach(var item in list2)
                {
                    Mapping map = new Mapping();
                    map.c1 = item.ShopName + item.Address + item.MarketName + item.City;
                    map.row = item.Row;
                    mapList.Add(map);
                }
                var query32 = mapList.Select(x => new Mapping { c1 = x.c1 ?? rnd.Next(0, 100000).ToString(), row = x.row }).GroupBy(y => y.c1).Where(y=>y.Count()>1).ToList();
                newError.AddRange(checks.Unique(query32, 1, "ShopName,Address,MarketName,City"));
            }
            // Beat State and District
            
            //Product Addtion 
            if (list3 != null)
            {
                foreach(var item in list3)
                {
                    if(item.PrimaryCategory=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "PrimaryCategory";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.SecondaryCategory=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "SecondaryCategory";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.Product=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Character Length Error";
                        temp.Field_1 = "Product";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.Price=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "Price";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.Unit=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "Unit";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    
                }
                var query1 = list3.GroupBy(x => x.PrimaryCategory, x => new Mapping { c1 = x.SecondaryCategory, row = x.Row }).ToList();
                var query2 = list3.GroupBy(x => x.SecondaryCategory, x => new Mapping { c1 = x.DisplayCategory, row = x.Row }).ToList();
                var query3 = list3.GroupBy(x => x.DisplayCategory, x => new Mapping { c1 = x.Product, row = x.Row }).ToList();
                
                var query5 = list3.GroupBy(x => x.ProductCode, x => new Mapping { c1 = x.VariantCode, row = x.Row }).ToList();
                var query6 = list3.GroupBy(x => x.Product, x => new Mapping { c1 = x.Variant, row = x.Row }).ToList();
                
                newError.AddRange(checks.One2ManyValidationCheck(query1, indexer.PrimaryCategory, indexer.SecondaryCategory, "PrimaryCategory", "SecondaryCategory"));
                newError.AddRange(checks.One2ManyValidationCheck(query2, indexer.SecondaryCategory, indexer.DisplayCategory, "SecondaryCategory", "DisplayCategory"));
                newError.AddRange(checks.One2ManyValidationCheck(query3, indexer.DisplayCategory, indexer.Product, "DisplayCategory", "Product"));
                
                newError.AddRange(checks.One2OneValidationCheck(query5, indexer.ProductCode, indexer.VariantCode, "ProductCode", "VariantCode"));
                newError.AddRange(checks.UniqueProductVariant(query6, indexer.Product, indexer.Variant));
                
            }
            if(list4!=null)
            {
                foreach(var item in list4)
                {
                    if(item.ESM=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "ESM";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.FinalBeatName=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "FinalBeatName";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.BeatPlanStartDate=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "BeatPlanStartDate";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.BeatPeriod=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "BeatPeriod";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    if(item.BeatDay=="")
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = "Null Entry";
                        temp.Field_1 = "BeatDay";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    try
                    {
                        if(int.Parse(item.BeatPeriod)<1)
                        {
                            ErrorTemplates temp = new ErrorTemplates();
                            temp.ErrorType = "Beat period cannot be less than 1";
                            temp.Field_1 = "BeatPeriod";
                            temp.Row = item.Row;
                            newError.Add(temp);
                        }

                    }
                    catch
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = " BeatDay must be integer";
                        temp.Field_1 = "BeatDay";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                    try
                    {
                        if (int.Parse(item.BeatDay) >int.Parse(item.BeatPeriod))
                        {
                            ErrorTemplates temp = new ErrorTemplates();
                            temp.ErrorType = "Beat period cannot be less than 1";
                            temp.Field_1 = "BeatPeriod";
                            temp.Row = item.Row;
                            newError.Add(temp);
                        }

                    }
                    catch
                    {
                        ErrorTemplates temp = new ErrorTemplates();
                        temp.ErrorType = " BeatDay must be integer";
                        temp.Field_1 = "BeatDay";
                        temp.Row = item.Row;
                        newError.Add(temp);
                    }
                }
                var query = list4.GroupBy(x => x.BeatDay, x => new Mapping { c1 = x.FinalBeatName, row = x.Row }).ToList();
                newError.AddRange(checks.One2ManyValidationCheck(query, indexer.BeatDay, indexer.FinalBeatName, "BeatDay", "FinalBeatName"));
            }
            return newError;
        }
        public List<WarningTemplates> WarningChecks(ExcelWorksheet sheet, List<WarningTemplates> error)
        {
            List<WarningTemplates> warningError = new List<WarningTemplates>();
            columnIndex indexer = new columnIndex();
            MappingValidations checks = new MappingValidations();
            var file = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.Start.Row, sheet.Dimension.End.Column];
            for (int i = sheet.Dimension.Start.Column - 1; i <= sheet.Dimension.End.Column - 1; i++)                 //To find Empty cells.(algo will be updated)
            {
                var value = ((object[,])file.Value)[0, i]; 
                if (value != null)
                {
                    if (value.ToString().Trim().Replace(" ", "").ToLower() == "NSM".ToLower())
                    {
                        indexer.NSM = i + 1;
                    }

                    if (value.ToString().Trim().Replace(" ", "").ToLower() == "ZSM".ToLower())
                    {
                        indexer.ZSM = i + 1;
                    }

                    if (value.ToString().Trim().Replace(" ", "").ToLower() == "RSM".ToLower())
                    {
                        indexer.RSM = i + 1;
                    }

                    if (value.ToString().Trim().Replace(" ", "").ToLower() == "ASM".ToLower())
                    {
                        indexer.ASM = i + 1;
                    }

                    if (value.ToString().Trim().Replace(" ", "").ToLower() == "ESM".ToLower())
                    {
                        indexer.ESM = i + 1;
                    }


                }

               
               
            }
            warningError.AddRange(checks.HierarchyWarning(sheet, indexer.ESM, indexer.ASM, indexer.RSM, indexer.ZSM, indexer.NSM));
            return warningError;
        }
    }
}