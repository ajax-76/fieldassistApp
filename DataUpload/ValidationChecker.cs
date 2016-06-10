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
        public List<ErrorTemplates> Checker(ExcelWorksheet sheet,List<ErrorTemplates>error)
        {
            List<ErrorTemplates> newError = new List<ErrorTemplates>();
            columnIndex indexer = new columnIndex();
            MappingValidations checks = new MappingValidations();
            var file = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.Start.Row, sheet.Dimension.End.Column];

            for (int i = sheet.Dimension.Start.Column - 1; i <= sheet.Dimension.End.Column - 1; i++)                 //To find Empty cells.(algo will be updated)
            {
                if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("NSM").ToLower())
                {
                    indexer.NSM = i + 1;
                }
                if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("NSMZone").ToLower())
                {
                    indexer.NSMZone = i + 1;
                }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("NSMEmailId").ToLower())
                 {
                     indexer.NSMEmailId = i+1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("NSMSecondaryEmailId").ToLower())
                 {
                     indexer.NSMSecondaryEmailId = i+1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("ZSM").ToLower())
                 {
                     indexer.ZSM = i+1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("ZSMZone").ToLower())
                 {
                     indexer.ZSMZone = i+1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("ZSMEmailId").ToLower())
                 {
                     indexer.ZSMEmailId = i+1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("ZSMSecondaryEmailId").ToLower())
                 {
                     indexer.ZSMSecondaryEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == ("RSM").ToLower())
                 {
                     indexer.RSM = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "RSMZone".ToLower())
                 {
                     indexer.RSMZone = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "RSMEmailId".ToLower())
                 {
                     indexer.RSMEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "RSMSecondaryEmailId".ToLower())
                 {
                     indexer.RSMSecondaryEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ASM".ToLower())
                 {
                     indexer.ASM = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ASMZone".ToLower())
                 {
                     indexer.ASMZone = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ASMEmailId".ToLower())
                 {
                     indexer.ASMEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ASMSecondaryEmailId".ToLower())
                 {
                     indexer.ASMSecondaryEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESM".ToLower())
                 {
                     indexer.ESM = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMZone".ToLower())
                 {
                     indexer.ESMZone = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMEmailId".ToLower())
                 {
                     indexer.ESMEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMSecondaryEmailId".ToLower())
                 {
                     indexer.ESMSecondaryEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMContactNumber".ToLower())
                 {
                     indexer.ESMContactNumber = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMHQ".ToLower())
                 {
                     indexer.ESMHQ = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESMErpId".ToLower())
                 {
                     indexer.ESMErpId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "FinalBeatName".ToLower())
                 {
                     indexer.FinalBeatName = i + 1;
                 }
                if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "BeatErpId".ToLower())
                 {
                     indexer.BeatErpId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "BeatDistrict".ToLower())
                 {
                     indexer.BeatDistrict = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "BeatState".ToLower())
                 {
                     indexer.BeatZone = i + 1;
                 }
                 if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "BeatZone".ToLower())
                 {
                     indexer.BeatState = i + 1;
                 }
                 if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "DistributorName".ToLower())
                 {
                     indexer.DistributorName = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "DistributorLocation".ToLower())
                 {
                     indexer.DistributorLocation = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "DistributorErpId".ToLower())
                 {
                     indexer.DistributorErpId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "DistributorEmailId".ToLower())
                 {
                     indexer.DistributorEmailId = i + 1;
                 }
            }
            //hierarchy break
            newError.AddRange(checks.HierarchyError(sheet, indexer.ESM,indexer.ASM, indexer.RSM, indexer.ZSM, indexer.NSM, error));
            //Phone Digits checking
            newError.AddRange(checks.CheckPhoneDigit(sheet, indexer.ESMContactNumber,error));
            //Relationship Check
            newError.AddRange(checks.AttributeMapping(sheet, indexer.NSM, indexer.NSMZone, "NSM", "NSMZone", error));
            //Nsm to Nsm EmailID
            newError.AddRange(checks.AttributeMapping(sheet, indexer.NSM,indexer.NSMEmailId, "NSM", "NSMEmailId", error));
            //Nsm to Secondary EmailID
            newError.AddRange(checks.AttributeMapping(sheet, indexer.NSM, indexer.NSMSecondaryEmailId, "NSM", "NSMSecondaryEmailId", error));
            //Nsm to Zsm
            newError.AddRange(checks.One2ManyValidationCheck(sheet, indexer.NSM, indexer.ZSM, "NSM", "ZSM", error));
            //Zsm to Zsm Zone
            newError.AddRange(checks.AttributeMapping(sheet, indexer.ZSM, indexer.ZSMZone, "ZSM", "ZSMZone", error));
            //Zsm to Zsm EmailId
            newError.AddRange(checks.AttributeMapping(sheet, indexer.ZSM, indexer.ZSMEmailId, "ZSM", "ZSMEmailId", error));
            //Zsm to Zsm Secondary Email ID
            newError.AddRange(checks.AttributeMapping(sheet, indexer.ZSM, indexer.ZSMSecondaryEmailId, "ZSM", "ZSMSecondaryEmailId", error));
            //Zsm to Rsm
            newError.AddRange(checks.One2ManyValidationCheck(sheet, indexer.ZSM, indexer.RSM, "ZSM", "RSM", error));
            //Rsm to Rsm Zone
            newError.AddRange(checks.AttributeMapping(sheet, indexer.RSM, indexer.RSMZone, "RSM", "RSMZone", error));
            //Rsm to Rsm EmailId
            newError.AddRange(checks.AttributeMapping(sheet, indexer.RSM, indexer.RSMEmailId, "RSM", "RSMEmailId", error));
            //Rsm to Rsm SecondaryEmailId
            newError.AddRange(checks.AttributeMapping(sheet, indexer.RSM, indexer.RSMSecondaryEmailId, "RSM", "RSMSecondaryEmailId", error));
            //Rsm to Asm
            newError.AddRange(checks.One2ManyValidationCheck(sheet, indexer.RSM, indexer.ASM,"RSM","ASM",error));
            //Asm to Asm Zone
            newError.AddRange(checks.AttributeMapping(sheet,indexer.ASM, indexer.ASMZone, "ASM", "ASMZone", error));
            //Asm to Asm EmailId
            newError.AddRange(checks.AttributeMapping(sheet, indexer.ASM, indexer.ASMEmailId, "ASM", "ASMEmailId", error));
            //Asm to Secondary EmailId
            newError.AddRange(checks.AttributeMapping(sheet, indexer.ASM, indexer.ASMSecondaryEmailId, "ASM", "ASMSecondaryEmailId", error));
            //Asm to Esm
            newError.AddRange(checks.One2ManyValidationCheck(sheet, indexer.ASM, indexer.ESM,"ASM","ESM",error));
            //Esm to Esm Zone
            newError.AddRange(checks.AttributeMapping(sheet, indexer.ESM, indexer.ESMZone, "ESM", "ESMZone", error));
            //Esm to Esm EmailId
            newError.AddRange(checks.AttributeMapping(sheet, indexer.ESM, indexer.ESMEmailId, "ESM", "ESMEmailId", error));
            //Esm to Esm Secondary EmailId
            newError.AddRange(checks.AttributeMapping(sheet, indexer.ESM, indexer.ESMSecondaryEmailId, "ESM", "ESMSecondaryEmailId", error));
            //Esm to Esm HQ
            newError.AddRange(checks.AttributeMapping(sheet, indexer.ESM,indexer.ESMHQ, "ESM", "ESMHQ", error));
            //ESM to ESMErpId
            newError.AddRange(checks.One2OneValidationCheck(sheet, indexer.ESM, indexer.ESMErpId, "ESM", "ESMErpId", error));
            //ESM to ESMContact Number
            newError.AddRange(checks.One2OneValidationCheck(sheet, indexer.ESM, indexer.ESMContactNumber, "ESM", "ESMContactNumber", error));
            //BeatDistrict to FinalBeatName
            newError.AddRange(checks.One2ManyValidationCheck(sheet, indexer.BeatDistrict, indexer.FinalBeatName,"BeatDistrict","FinalBeatName",error));
            //BeatState FinalBeatName
            newError.AddRange(checks.One2ManyValidationCheck(sheet, indexer.BeatState, indexer.FinalBeatName,"BeatState","FinalBeatName",error));
            //BeatZone to FinalBeatName
            newError.AddRange(checks.One2ManyValidationCheck(sheet, indexer.BeatZone, indexer.FinalBeatName,"BeatZone","FinalBeatName",error));
            //FinalBeatName to DistributorName
            newError.AddRange(checks.One2ManyValidationCheck(sheet, indexer.FinalBeatName, indexer.DistributorName,"FinalBeatName","DistributorName",error));
            //FinalBeatName to DistributorLocation
            newError.AddRange(checks.One2ManyValidationCheck(sheet, indexer.FinalBeatName, indexer.DistributorLocation, "FinalBeatName", "DistributorLocation", error));
            //FinalBeatName to DistributorLocation
            newError.AddRange(checks.One2ManyValidationCheck(sheet, indexer.FinalBeatName, indexer.DistributorEmailId, "FinalBeatName", "DistributorEmailId", error));
            //DistributorName to DistributorErpId
            newError.AddRange(checks.One2OneValidationCheck(sheet, indexer.DistributorName, indexer.DistributorErpId,"DistributorName","DistributorErpId",error));
            //Emails Check
            newError.AddRange(checks.EmailCheck(sheet, indexer.NSMEmailId, "NSMEmailId", error));
            newError.AddRange(checks.EmailCheck(sheet, indexer.NSMSecondaryEmailId, "NSMSecondaryEmailId", error));
            newError.AddRange(checks.EmailCheck(sheet, indexer.ZSMEmailId, "ZSMEmailId", error));
            newError.AddRange(checks.EmailCheck(sheet, indexer.ZSMSecondaryEmailId, "ZSMSecondaryEmailId", error));
            newError.AddRange(checks.EmailCheck(sheet, indexer.RSMEmailId, "RSMEmailId", error));
            newError.AddRange(checks.EmailCheck(sheet, indexer.ZSMSecondaryEmailId, "RSMSecondaryEmailId", error));
            newError.AddRange(checks.EmailCheck(sheet, indexer.ASMEmailId, "ASMEmailId", error));
            newError.AddRange(checks.EmailCheck(sheet, indexer.ASMSecondaryEmailId, "ASMSecondaryEmailId", error));
            newError.AddRange(checks.EmailCheck(sheet, indexer.ESMEmailId, "ESMEmailId", error));
            newError.AddRange(checks.EmailCheck(sheet, indexer.ESMSecondaryEmailId, "ESMSecondaryEmailId", error));
            newError.AddRange(checks.EmailCheck(sheet, indexer.DistributorEmailId, "DistributorEmailId", error));
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
                if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "NSM".ToLower())
                {
                    indexer.NSM = i + 1;
                }
                
                if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ZSM".ToLower())
                {
                    indexer.ZSM = i + 1;
                }
                
                if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "RSM".ToLower())
                {
                    indexer.RSM = i + 1;
                }
                
                if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ASM".ToLower())
                {
                    indexer.ASM = i + 1;
                }
                
                if (((object[,])file.Value)[0, i].ToString().Trim().Replace(" ", "").ToLower() == "ESM".ToLower())
                {
                    indexer.ESM = i + 1;
                }
                
                
            }

            error.AddRange(checks.HierarchyWarning(sheet,indexer.ESM, indexer.ASM, indexer.RSM, indexer.ZSM, indexer.NSM, error));
            return error;
        }
    }
}