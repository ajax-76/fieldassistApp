using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace ConsoleApplication1
{
    class ValidationChecker
    {
        

        public void Checker(ExcelWorksheet sheet)
        {
            columnIndex indexer = new columnIndex();
            MappingValidations checks = new MappingValidations();
            var file = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.Start.Row, sheet.Dimension.End.Column];

            for (int i = sheet.Dimension.Start.Column - 1; i <= sheet.Dimension.End.Column - 1; i++)                 //To find Empty cells.(algo will be updated)
            {
               if (((object[,])file.Value)[0, i].ToString().Trim() == "NSM")
                  {
                    indexer.NSM = i+1;
                  }
                if (((object[,])file.Value)[0, i].ToString().Trim() == "NSMZone") 
                {
                    indexer.NSMZone = i+1;
                }
                if (((object[,])file.Value)[0, i].ToString().Trim() == "ZSM")
                {
                    indexer.ZSM = i + 1;
                }
                if (((object[,])file.Value)[0, i].ToString().Trim() == "FinalBeatName")
                {
                    indexer.FinalBeatName = i + 1;
                }
                if (((object[,])file.Value)[0, i].ToString().Trim() == "BeatState")
                {
                    indexer.BeatZone = i + 1;
                }
                if (((object[,])file.Value)[0, i].ToString().Trim() == "BeatZone")
                {
                    indexer.BeatState = i + 1;
                }
                if (((object[,])file.Value)[0, i].ToString().Trim() == "DistributorName")
                {
                    indexer.DistributorName = i + 1;
                }
                if (((object[,])file.Value)[0, i].ToString().Trim() == "DistributorErpId")
                {
                    indexer.DistributorErpId = i + 1;
                }
                /* if(((object[,])file.Value)[0, i].ToString().Trim() =="NSMEmailId")
                 {
                     indexer.NSMEmailId = i+1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="NSMSecondaryEmailId")
                 {
                     indexer.NSMSecondaryEmailId = i+1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ZSM")
                 {
                     indexer.ZSM = i+1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ZSMZone")
                 {
                     indexer.ZSMZone = i+1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ZSMEmailId")
                 {
                     indexer.ZSMEmailId = i+1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ZSMSecondaryEmailId")
                 {
                     indexer.ZSMSecondaryEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="RSM")
                 {
                     indexer.RSM = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="RSMZone")
                 {
                     indexer.RSMZone = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="RSMEmailId")
                 {
                     indexer.RSMEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="RSMSecondaryEmailId")
                 {
                     indexer.RSMSecondaryEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ASM")
                 {
                     indexer.ASM = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ASMZone")
                 {
                     indexer.ASMZone = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ASMEmailId")
                 {
                     indexer.ASMEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ASMSecondaryEmailId")
                 {
                     indexer.ASMSecondaryEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ESM")
                 {
                     indexer.ESM = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ESMZone")
                 {
                     indexer.ESMZone = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ESMEmailId")
                 {
                     indexer.ESMEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim()=="ESMSecondaryEmailId")
                 {
                     indexer.ESMSecondaryEmailId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="ESMContactNumber")
                 {
                     indexer.ESMContactNumber = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="ESMHQ")
                 {
                     indexer.ESMErpId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="ESMErpId")
                 {
                     indexer.ESMErpId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="FinalBeatName")
                 {
                     indexer.FinalBeatName = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="BeatErpId")
                 {
                     indexer.BeatErpId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="BeatDistrict")
                 {
                     indexer.BeatDistrict = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="BeatState")
                 {
                     indexer.BeatZone = i + 1;
                 }
                 if (((object[,])file.Value)[0, i].ToString().Trim() == "BeatZone")
                 {
                     indexer.BeatState = i + 1;
                 }
                 if (((object[,])file.Value)[0, i].ToString().Trim() =="DistributorName")
                 {
                     indexer.DistributorName = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="DistributorLocation")
                 {
                     indexer.DistributorLocation = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="DistributorErpId")
                 {
                     indexer.DistributorErpId = i + 1;
                 }
                 if(((object[,])file.Value)[0, i].ToString().Trim() =="DistributorEmailId")
                 {
                     indexer.DistributorEmailId = i + 1;
                 }*/
            }
            //Relationship Check
            checks.One2ManyValidationCheck(sheet, indexer.NSM, indexer.NSMZone,"NSM","NSMZone");
          /*  checks.One2ManyValidationCheck(sheet, indexer.NSM, indexer.ZSM);
            checks.One2ManyValidationCheck(sheet, indexer.BeatState, indexer.FinalBeatName);
            checks.One2ManyValidationCheck(sheet, indexer.BeatZone, indexer.FinalBeatName);
            checks.One2ManyValidationCheck(sheet, indexer.FinalBeatName, indexer.DistributorName);
            checks.One2OneValidationCheck(sheet, indexer.DistributorName, indexer.DistributorErpId);
            //Nsm to Nsm EmailID
            /*  checks.One2ManyValidationCheck(sheet, indexer.NSM, indexer.NSMEmailId);
              //Nsm to Secondary EmailID
              checks.One2ManyValidationCheck(sheet, indexer.NSM, indexer.NSMSecondaryEmailId);
              //Nsm to Zsm
              checks.One2ManyValidationCheck(sheet, indexer.NSM, indexer.ZSM);
              //Zsm to Zsm Zone
              checks.One2ManyValidationCheck(sheet, indexer.ZSM, indexer.ZSMZone);
              //Zsm to Zsm EmailId
              checks.One2ManyValidationCheck(sheet, indexer.ZSM, indexer.ZSMEmailId);
              //Zsm to Zsm Secondary Email ID
              checks.One2ManyValidationCheck(sheet, indexer.ZSM, indexer.ZSMSecondaryEmailId);
              //Zsm to Rsm
              checks.One2ManyValidationCheck(sheet, indexer.ZSM, indexer.RSM);
              //Rsm to Rsm Zone
              checks.One2ManyValidationCheck(sheet, indexer.RSM, indexer.RSMZone);
              //Rsm to Rsm EmailId
              checks.One2ManyValidationCheck(sheet, indexer.RSM, indexer.RSMEmailId);
              //Rsm to Rsm SecondaryEmailId
              checks.One2ManyValidationCheck(sheet, indexer.RSM, indexer.RSMSecondaryEmailId);
              //Rsm to Asm
              checks.One2ManyValidationCheck(sheet, indexer.RSM, indexer.ASM);
              //Asm to Asm Zone
              checks.One2ManyValidationCheck(sheet,indexer.ASM, indexer.ASMZone);
              //Asm to Asm EmailId
              checks.One2ManyValidationCheck(sheet, indexer.ASM, indexer.ASMEmailId);
              //Asm to Secondary EmailId
              checks.One2ManyValidationCheck(sheet, indexer.ASM, indexer.ASMSecondaryEmailId);
              //Asm to Esm
              checks.One2ManyValidationCheck(sheet, indexer.ASM, indexer.ESM);
              //Esm to Esm Zone
              checks.One2ManyValidationCheck(sheet, indexer.ESM, indexer.ESMZone);
              //Esm to Esm EmailId
              checks.One2ManyValidationCheck(sheet, indexer.ESM, indexer.ESMEmailId);
              //Esm to Esm Secondary EmailId
              checks.One2ManyValidationCheck(sheet, indexer.ESM, indexer.ESMSecondaryEmailId);
              //Esm to Esm HQ
              checks.One2ManyValidationCheck(sheet, indexer.ESM,indexer.ESMHQ);
              //ESM to ESMErpId
              checks.One2OneValidationCheck(sheet, indexer.ESM, indexer.ESMErpId);
              //ESM to ESMContact Number
              checks.One2OneValidationCheck(sheet, indexer.ESM, indexer.ESMContactNumber);
              //
              checks.One2ManyValidationCheck(sheet, indexer.BeatDistrict, indexer.FinalBeatName);
              checks.One2ManyValidationCheck(sheet, indexer.BeatState, indexer.FinalBeatName);
              checks.One2ManyValidationCheck(sheet, indexer.BeatZone, indexer.FinalBeatName);
              checks.One2ManyValidationCheck(sheet, indexer.FinalBeatName, indexer.DistributorName);
              checks.One2ManyValidationCheck(sheet, indexer.FinalBeatName, indexer.DistributorLocation);
              checks.One2ManyValidationCheck(sheet, indexer.FinalBeatName, indexer.DistributorEmailId);
              checks.One2OneValidationCheck(sheet, indexer.DistributorName, indexer.DistributorErpId);*/
        }
    }
}

