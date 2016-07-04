using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;                                                                                   
using System.Web.Mvc;                                                                                   
using OfficeOpenXml;
using CsvHelper;
using CsvHelper.Configuration;
using System.Text.RegularExpressions;

namespace DataUpload.Controllers
{
    
   
    public class UploadController : Controller
    {

        // GET: Upload
        public ActionResult Upload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            string path = null;
            int count = 0;
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            List<WarningTemplates> warningTemp = new List<WarningTemplates>();
            AllErrors allErrors = new AllErrors();
            allErrors.Error = null;
            allErrors.ShowHeader = null;
            allErrors.WarnHeaders = null;
            allErrors.Warning = null;
            try
            {

                if (file.ContentLength > 0)
                {

                    string filename = Path.GetFileName(file.FileName);

                    path = Server.MapPath("~/Seed Data/" + filename);
                    file.SaveAs(path);
                    ValidationChecker mapChecker = new ValidationChecker();
                    List<string> WarningfileHeaders = new List<string>();
                    List<string> ErrorfileHeaders = new List<string>();
                    WarningfileHeaders.Add("NSM".ToLower());                //---------------------cell1
                    WarningfileHeaders.Add("NSMZone".ToLower());
                    WarningfileHeaders.Add("NSMEmailId".ToLower());         //---------------------cell3
                    WarningfileHeaders.Add("NSMSecondaryEmailId".ToLower());//---------------------cell4
                    WarningfileHeaders.Add("ZSM".ToLower());                //---------------------cell5
                    WarningfileHeaders.Add("ZSMEmailId".ToLower());         //---------------------cell6
                    WarningfileHeaders.Add("ZSMZone".ToLower());            //---------------------cell7
                    WarningfileHeaders.Add("ZSMSecondaryEmailId".ToLower());//---------------------cell8
                    WarningfileHeaders.Add("RSM".ToLower());                //---------------------cell9
                    WarningfileHeaders.Add("RSMEmailId".ToLower());         //---------------------cell10
                    WarningfileHeaders.Add("RSMSecondaryEmailId".ToLower());//---------------------cell11
                    WarningfileHeaders.Add("ASM".ToLower());                //---------------------cell12
                    WarningfileHeaders.Add("ASMEmailId".ToLower());         //---------------------cell13
                    WarningfileHeaders.Add("ASMZone".ToLower());            //---------------------cell14
                    WarningfileHeaders.Add("ASMSecondaryEmailId".ToLower());//---------------------cell15
                    WarningfileHeaders.Add("ESM".ToLower());                //---------------------cell16
                    WarningfileHeaders.Add("ESMEmailId".ToLower());         //---------------------cell17
                    WarningfileHeaders.Add("ESMZone".ToLower());            //---------------------cell18
                    WarningfileHeaders.Add("ESMSecondaryEmailId".ToLower());//---------------------cell19
                    WarningfileHeaders.Add("ESMContactNumber".ToLower());   //---------------------cell20
                    WarningfileHeaders.Add("ESMHQ".ToLower());
                    WarningfileHeaders.Add("ESMErpId".ToLower());//---------------------cell21
                    WarningfileHeaders.Add("FinalBeatName".ToLower());      //---------------------cell22
                    WarningfileHeaders.Add("BeatErpId".ToLower());          //---------------------cell23
                    WarningfileHeaders.Add("BeatDistrict".ToLower());       //---------------------cell24
                    WarningfileHeaders.Add("BeatState".ToLower());          //---------------------cell25
                    WarningfileHeaders.Add("BeatZone".ToLower());           //---------------------cell26
                    WarningfileHeaders.Add("DistributorName".ToLower());    //---------------------cell27
                    WarningfileHeaders.Add("DistributorLocation".ToLower());//---------------------cell28
                    WarningfileHeaders.Add("DistributorErpId".ToLower());   //---------------------cell29
                    WarningfileHeaders.Add("DistributorEmailId".ToLower()); //---------------------cell20
                    ErrorfileHeaders.Add("FinalBeatName".ToLower());
                    ErrorfileHeaders.Add("BeatState".ToLower());
                    ErrorfileHeaders.Add("BeatDistrict".ToLower());
                    try
                    {
                        var xfile = new FileInfo(path);
                        ExcelPackage package = new ExcelPackage(xfile);
                        ExcelWorksheet sheet = package.Workbook.Worksheets[1];

                        var header = sheet.Cells[1, 1, 1, sheet.Dimension.End.Column];
                        var fileField = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
                        List<string> headerError = new List<string>();
                        for (int i = sheet.Dimension.Start.Column; i <= sheet.Dimension.End.Column; i++)
                        {
                            var value = sheet.Cells[1, i].Value;
                            if (value != null)
                            {
                                string newValue = value.ToString().Trim().Replace(" ", "").ToLower();
                                if (newValue == "FinalBeatName".ToLower())
                                {
                                    headerError.Add(newValue);
                                }
                                if (newValue == "BeatState".ToLower())
                                {
                                    headerError.Add(newValue);
                                }
                                if (newValue == "BeatDistrict".ToLower())
                                {
                                    headerError.Add(newValue);
                                }

                            }
                        }
                        var difference1 = ErrorfileHeaders.Except(headerError);
                        if (difference1.Any())
                        {
                            ErrorTemplates error = new ErrorTemplates();
                            error.ErrorType = "Missing Header file";
                            error.ErrorComments = "Cannot go further Ensure these headers must be present" + string.Join(",", difference1);
                            errorTemp.Add(error);
                            allErrors.Error = errorTemp;
                            return View(allErrors);

                        }
                        List<string> headerCheck = new List<string>();
                        for (int i = 0; i < sheet.Dimension.End.Column; i++)
                        {
                            if (((object[,])fileField.Value)[0, i] != null)
                            {
                                headerCheck.Add(((object[,])fileField.Value)[0, i].ToString().Replace(" ", "").ToLower());
                            }

                        }

                        var difference = WarningfileHeaders.Except(headerCheck);
                        if (difference.Any())
                        {
                            WarningTemplates warningTemplates = new WarningTemplates();
                            warningTemplates.Comments = "These Headers are missing do you wish to continue";
                            warningTemplates.Field = string.Join(",", difference);
                            warningTemp.Add(warningTemplates);
                            
                        }

                        

                        warningTemp.AddRange(mapChecker.WarningChecks(sheet, warningTemp));//Warnings
                       

                        

                        //Excel to object  

                        List<FileHeadersBeatHierarchy> list = new List<FileHeadersBeatHierarchy>();
                        List<FileHeadersProductAddition> list2 = new List<FileHeadersProductAddition>();
                        
                        List<FileHeadersLocationAddtion> list3 = new List<FileHeadersLocationAddtion>();
                        List<FileHeadersBeatPlanAddition> list4 = new List<FileHeadersBeatPlanAddition>();
                        for (int i = sheet.Dimension.Start.Row; i < sheet.Dimension.End.Row; i++)
                        {
                            FileHeadersBeatHierarchy records = new FileHeadersBeatHierarchy();
                            for (int j = sheet.Dimension.Start.Column - 1; j < sheet.Dimension.End.Column; j++)
                            {
                                var value = ((object[,])fileField.Value)[0, j];
                                if (value != null)
                                {
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "NSM".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.NSM = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.NSM = "";
                                            records.Row = i;
                                        }
                                        if(records.NSM.Replace(" ", "").Count()>=50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "NSM";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                            
                                        }

                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "NSMZone".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.NSMZone = ((object[,])fileField.Value)[i, j].ToString();
                                        }
                                        else
                                        {
                                            records.NSMZone = "";
                                        }
                                        if(records.NSMZone.Replace(" ", "").Count()>=100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "NSMZone";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "NSMEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.NSMEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.NSMEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.NSMEmailId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "NSMEmailId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "NSMSecondaryEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.NSMSecondaryEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.NSMSecondaryEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.NSMSecondaryEmailId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "NSMSecondaryEmaiId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ZSM".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ZSM = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ZSM = "";
                                            records.Row = i;
                                        }
                                        if (records.ZSM.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ZSM";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ZSMZone".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ZSMZone = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ZSMZone = "";
                                            records.Row = i;
                                        }
                                        if (records.ZSMZone.Replace(" ", "").Count() >= 100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ZSMZone";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ZSMEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ZSMEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ZSMEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.ZSMEmailId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ZSMEmailId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ZSMSecondaryEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ZSMSecondaryEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ZSMSecondaryEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.ZSMSecondaryEmailId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ZSMSecondaryEmailId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "RSM".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.RSM = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.RSM = "";
                                            records.Row = i;
                                        }
                                        if (records.RSM.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "RSM";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "RSMZone".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.RSMZone = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.RSMZone = "";
                                            records.Row = i;
                                        }
                                        if (records.RSMZone.Replace(" ", "").Count() >=100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "RSMZone";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "RSMEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.RSMEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.RSMEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.RSMEmailId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "RSMEmailId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "RSMSecondaryEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.RSMSecondaryEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.RSMSecondaryEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.RSMSecondaryEmailId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "RSMSecondaryEmailId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ASM".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ASM = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ASM = "";
                                            records.Row = i;
                                        }
                                        if (records.ASM.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ASM";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ASMZone".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ASMZone = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ASMZone = "";
                                            records.Row = i;
                                        }
                                        if (records.ASMZone.Replace(" ", "").Count() >= 100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ASMZone";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ASMEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ASMEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ASMEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.ASMEmailId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ASMEmailId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ASMSecondaryEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ASMSecondaryEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ASMSecondaryEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.ASMSecondaryEmailId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ASMSecondaryEmailId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ESM".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ESM = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ESM = "";
                                            records.Row = i;
                                        }
                                        if (records.ESM.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ESM";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ESMEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ESMEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ESMEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.ESMEmailId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ESMEmailId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ESMSecondaryEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ESMSecondaryEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ESMSecondaryEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.ESMSecondaryEmailId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ESMSecondaryEmailId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ESMZone".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ESMZone = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ESMZone = "";
                                            records.Row = i;
                                        }
                                        if (records.ESMZone.Replace(" ", "").Count() >= 100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ESMZone";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ESMContactNumber".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ESMContactNumber = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ESMContactNumber = "";
                                            records.Row = i;
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ESMHQ".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ESMHQ = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ESMHQ = "";
                                            records.Row = i;
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ESMErpId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ESMErpId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ESMErpId = "";
                                            records.Row = i;
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "FinalBeatName".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.FinalBeatName = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.FinalBeatName = "";
                                            records.Row = i;
                                        }
                                        if (records.FinalBeatName.Replace(" ", "").Count() >=100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "FinalBeatName";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "BeatState".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.BeatState = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.BeatState = "";
                                            records.Row = i;
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "BeatDistrict".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.BeatDistrict = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.BeatDistrict = "";
                                            records.Row = i;
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "BeatZone".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.BeatZone = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.BeatZone = "";
                                            records.Row = i;
                                        }
                                        if (records.BeatZone.Replace(" ", "").Count() >= 100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "BeatZone";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "BeatErpId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.BeatErpId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.BeatErpId = "";
                                            records.Row = i;
                                        }
                                        if (records.BeatErpId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "BeatErpId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "DistributorName".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.DistributorName = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.DistributorName = "";
                                            records.Row = i;
                                        }
                                        if (records.DistributorName.Replace(" ", "").Count() >= 100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "DistributorName";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "DistributorLocation".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.DistributorLocation = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.DistributorLocation = "";
                                            records.Row = i;
                                        }
                                        if (records.DistributorLocation.Replace(" ", "").Count() >= 100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "DistributorLocation";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "DistributorEmailId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.DistributorEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.DistributorEmailId = "";
                                            records.Row = i;
                                        }
                                        if (records.DistributorEmailId.Replace(" ", "").Count() >= 100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "DistributorEmailId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "DistributorErpId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.DistributorErpId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.DistributorErpId = "";
                                            records.Row = i;
                                        }
                                        if (records.DistributorErpId.Replace(" ", "").Count() >= 100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "DistributorErpId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);
                                        }
                                    }
                                }
                            }
                            list.Add(records);
                        }
                        errorTemp.AddRange(mapChecker.Checker(sheet, list, list3, list2,list4));
                        
                        if (errorTemp != null)
                        {
                            allErrors.Error = errorTemp.GroupBy(a => new { a.Row, a.Field_1, a.Field_2, a.ErrorType, a.ErrorComments, a.IncorrectHeaderList }, (key, g) => g.FirstOrDefault()).ToList();
                            allErrors.Warning = warningTemp.OrderBy(x => x.Row).GroupBy(a => new { a.Row, a.Field, a.Comments }, (key, g) => g.FirstOrDefault()).ToList();
                            return View(allErrors);
                        }
                        //write csv


                        MemoryStream memory = new MemoryStream();
                        StreamWriter streamwiter = new StreamWriter(memory);
                        var newCsv = new CsvWriter(streamwiter);
                        newCsv.WriteHeader<FileHeadersBeatHierarchy>();

                        foreach (var item in list)
                        {
                            FileHeadersBeatHierarchy temp = new FileHeadersBeatHierarchy();
                            temp.NSM = item.NSM;
                            temp.NSMEmailId = item.NSMEmailId;
                            temp.NSMSecondaryEmailId = item.NSMSecondaryEmailId;
                            temp.NSMZone = item.NSMZone;
                            temp.ZSM = item.ZSM;
                            temp.ZSMEmailId = item.ZSMEmailId;
                            temp.ZSMSecondaryEmailId = item.ZSMSecondaryEmailId;
                            temp.ZSMZone = item.ZSMZone;
                            temp.RSM = item.RSM;
                            temp.RSMEmailId = item.RSMEmailId;
                            temp.RSMSecondaryEmailId = item.RSMSecondaryEmailId;
                            temp.RSMZone = item.RSMZone;
                            temp.ASM = item.ASM;
                            temp.ASMEmailId = item.ASMEmailId;
                            temp.ASMSecondaryEmailId = item.ASMSecondaryEmailId;
                            temp.ASMZone = item.ASMZone;
                            temp.ESM = item.ESM;
                            temp.ESMEmailId = item.ESMEmailId;
                            temp.ESMSecondaryEmailId = item.ESMSecondaryEmailId;
                            temp.ESMContactNumber = item.ESMContactNumber;
                            temp.ESMErpId = item.ESMErpId;
                            temp.ESMHQ = item.ESMHQ;
                            temp.ESMZone = item.ESMZone;
                            temp.FinalBeatName = item.FinalBeatName;
                            temp.BeatDistrict = item.BeatDistrict;
                            temp.BeatErpId = item.BeatErpId;
                            temp.BeatState = item.BeatState;
                            temp.BeatZone = item.BeatZone;
                            temp.DistributorName = item.DistributorName;
                            temp.DistributorEmailId = item.DistributorEmailId;
                            temp.DistributorLocation = item.DistributorLocation;
                            temp.DistributorErpId = item.DistributorErpId;
                            newCsv.WriteRecord<FileHeadersBeatHierarchy>(temp);
                        }
                        streamwiter.Flush();
                        return File(memory.ToArray(), "text/csv", "data.csv");
                    }
                    catch (Exception ex)
                    {
                        var c = count;
                        var error = ex;
                        return View("Error");
                    }
                }
            }
            catch (Exception ex)
            {
                var error = ex;
                return View("Error");
            }


            return null;

        }
        //Program for Location Addition

        public ActionResult Upload_LocationAddition()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Upload_LocationAddition(HttpPostedFileBase file)
        {
            string path = null;

            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            List<WarningTemplates> warningTemp = new List<WarningTemplates>();
            AllErrors allErrors = new AllErrors();
            allErrors.Error = null;
            allErrors.ShowHeader = null;
            allErrors.WarnHeaders = null;
            allErrors.Warning = null;
            try
            {

                if (file.ContentLength > 0)
                {

                    string filename = Path.GetFileName(file.FileName);

                    path = Server.MapPath("~/Seed Data/" + filename);
                    file.SaveAs(path);
                    ValidationChecker mapChecker = new ValidationChecker();
                    List<string> WarningfileHeaders = new List<string>();
                    List<string> ErrorfileHeaders = new List<string>();
                    WarningfileHeaders.Add("ShopName".ToLower());
                    WarningfileHeaders.Add("Address".ToLower());
                    WarningfileHeaders.Add("Email".ToLower());
                    WarningfileHeaders.Add("Tin".ToLower());
                    WarningfileHeaders.Add("Pin".ToLower());
                    WarningfileHeaders.Add("MarketName".ToLower());
                    WarningfileHeaders.Add("City".ToLower());
                    WarningfileHeaders.Add("State".ToLower());
                    WarningfileHeaders.Add("Country".ToLower());
                    WarningfileHeaders.Add("ShopCode".ToLower());
                    WarningfileHeaders.Add("ShopType".ToLower());
                    WarningfileHeaders.Add("Segmentation".ToLower());
                    WarningfileHeaders.Add("OwnersName".ToLower());
                    WarningfileHeaders.Add("OwnersContactNumber".ToLower());
                    WarningfileHeaders.Add("FinalBeatName".ToLower());
                    WarningfileHeaders.Add("BeatErpId".ToLower());
                    ErrorfileHeaders.Add("FinalBeatName".ToLower());
                    try
                    {
                        var xfile = new FileInfo(path);
                        ExcelPackage package = new ExcelPackage(xfile);
                        ExcelWorksheet sheet = package.Workbook.Worksheets[1];

                        var header = sheet.Cells[1, 1, 1, sheet.Dimension.End.Column];
                        var fileField = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
                        List<string> headerCheck = new List<string>();
               
                        for (int i = 0; i < sheet.Dimension.End.Column; i++)
                        {
                            if (((object[,])fileField.Value)[0, i] != null)
                            {
                                headerCheck.Add(((object[,])fileField.Value)[0, i].ToString().Replace(" ", "").ToLower());
                            }

                        }

                        var difference = WarningfileHeaders.Except(headerCheck);
                        if (difference.Any())
                        {
                            WarningTemplates warningTemplates = new WarningTemplates();
                            warningTemplates.Comments = "These Headers are missing do you wish to continue";
                            warningTemplates.Field = string.Join(",",difference);
                            warningTemp.Add(warningTemplates);
                            
                        }
                        List<string> headerError = new List<string>();
                        for (int i = sheet.Dimension.Start.Column-1; i < sheet.Dimension.End.Column; i++)
                        {
                            var value = ((object[,])fileField.Value)[0, i];
                            if (value != null)
                            {
                                string newValue = value.ToString().Trim().Replace(" ", "").ToLower();
                                if (newValue == "FinalBeatName".ToLower())
                                {
                                    headerError.Add(newValue);
                                }
                                
                            }
                        }
                        var difference1 = ErrorfileHeaders.Except(headerError);
                        if (difference1.Any())
                        {
                            ErrorTemplates error = new ErrorTemplates();
                            error.ErrorType = "Missing Header file";
                            error.ErrorComments = "Cannot go further Ensure these headers must be present" + string.Join(",", difference1);
                            errorTemp.Add(error);
                            allErrors.Error = errorTemp;
                            return View(allErrors);
                        }
                      
                        List<FileHeadersLocationAddtion> list = new List<FileHeadersLocationAddtion>();
                        List<FileHeadersProductAddition> list3 = new List<FileHeadersProductAddition>();
                        List<FileHeadersBeatHierarchy> list2 = new List<FileHeadersBeatHierarchy>();
                        List<FileHeadersBeatPlanAddition> list4 = new List<FileHeadersBeatPlanAddition>();


                        for (int i = sheet.Dimension.Start.Row; i < sheet.Dimension.End.Row; i++)
                        {
                            FileHeadersLocationAddtion records = new FileHeadersLocationAddtion();
                            for (int j = sheet.Dimension.Start.Column - 1; j < sheet.Dimension.End.Column; j++)
                            {
                                var value = ((object[,])fileField.Value)[0, j];
                                if (value != null)
                                {
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ShopName".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ShopName = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ShopName = "";
                                            records.Row = i;
                                        }
                                        if (records.ShopName.Replace(" ", "").Count() >= 100)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ShopName";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "Address".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.Address = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.Address = "";
                                            records.Row = i;
                                        }
                                        if (records.Address.Replace(" ", "").Count() >= 500)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "Address";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "Email".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.Email = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.Email = "";
                                            records.Row = i;
                                        }
                                        if (records.Email.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "Email";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "Tin".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.Tin = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.Tin = "";
                                            records.Row = i;
                                        }
                                        if (records.Tin.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "Tin";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "Pin".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.Pin = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.Pin  = "";
                                            records.Row = i;
                                        }
                                        if (records.Pin.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "Pin";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "MarketName".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.MarketName = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.MarketName = "";
                                            records.Row = i;
                                        }
                                        if (records.MarketName.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "MarketName";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "City".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.City = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.City = "";
                                            records.Row = i;
                                        }
                                        if (records.City.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "City";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "State".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.State = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.State = "";
                                            records.Row = i;
                                        }
                                        if (records.State.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "State";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "Country".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.Country = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.Country = "";
                                            records.Row = i;
                                        }
                                        if (records.Country.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "Country";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ShopCode".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ShopCode = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ShopCode = "";
                                            records.Row = i;
                                        }
                                        if (records.ShopCode.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ShopCode";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "ShopType".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.ShopType = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.ShopType = "";
                                            records.Row = i;
                                        }
                                        if (records.ShopType.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "ShopType";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "Segmentation".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.Segmentation = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.Segmentation = "";
                                            records.Row = i;
                                        }
                                        if (records.Segmentation.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "Segmentation";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "OwnersName".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.OwnersName = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.OwnersName = "";
                                            records.Row = i;
                                        }
                                        if (records.OwnersName.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "OwnersName";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "OwnersContactNumber".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.OwnersContactNumber = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.OwnersContactNumber = "";
                                            records.Row = i;
                                        }
                                        if (records.OwnersContactNumber.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "OwnersContactNumber";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "FinalBeatName".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.FinalBeatName = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.FinalBeatName = "";
                                            records.Row = i;
                                        }
                                        if (records.FinalBeatName.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "FinalBeatName";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                    if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "").ToLower() == "BeatErpId".ToLower())
                                    {
                                        if (((object[,])fileField.Value)[i, j] != null)
                                        {
                                            records.BeatErpId = ((object[,])fileField.Value)[i, j].ToString();
                                            records.Row = i;
                                        }
                                        else
                                        {
                                            records.BeatErpId = "";
                                            records.Row = i;
                                        }
                                        if (records.BeatErpId.Replace(" ", "").Count() >= 50)
                                        {
                                            ErrorTemplates temp = new ErrorTemplates();
                                            temp.ErrorType = "Character Length Error";
                                            temp.Field_1 = "BeatErpId";
                                            temp.Row = i;
                                            errorTemp.Add(temp);

                                        }
                                    }
                                }
                            }
                            list.Add(records);
                        }

                                        //    warningTemp = mapChecker.WarningChecks(sheet, warningTemp);//Warnings
                        errorTemp.AddRange(mapChecker.Checker(sheet,list2,list,list3,list4));//Mapping checking

                        if (errorTemp != null)
                        {
                            allErrors.Error = errorTemp.GroupBy(a => new { a.Row, a.Field_1, a.Field_2, a.ErrorType, a.ErrorComments, a.IncorrectHeaderList }, (key, g) => g.FirstOrDefault()).ToList();
                            allErrors.Warning = warningTemp.OrderBy(x => x.Row).GroupBy(a => new { a.Row, a.Field, a.Comments }, (key, g) => g.FirstOrDefault()).ToList();
                           // allErrors.ShowHeader = null;
                            return View(allErrors);
                        }
                        MemoryStream memory = new MemoryStream();
                        StreamWriter streamwiter = new StreamWriter(memory);
                        var newCsv = new CsvWriter(streamwiter);
                        newCsv.WriteHeader<FileHeadersLocationAddtion>();

                        foreach (var item in list)
                        {
                            FileHeadersLocationAddtion temp = new FileHeadersLocationAddtion();
                            temp.ShopName = item.ShopName;
                            temp.Address = item.Address;
                            temp.BeatErpId = item.BeatErpId;
                            temp.City = item.City;
                            temp.Country = item.Country;
                            temp.Email = item.Email;
                            temp.FinalBeatName = item.FinalBeatName;
                            temp.MarketName = item.MarketName;
                            temp.OwnersContactNumber = item.OwnersContactNumber;
                            temp.OwnersName = item.OwnersName;
                            temp.Pin = item.Pin;
                            temp.Segmentation = item.Segmentation;
                            temp.ShopCode = item.ShopCode;
                            temp.ShopType = item.ShopType;
                            temp.State = item.State;
                            temp.Tin = item.Tin;
                            newCsv.WriteRecord<FileHeadersLocationAddtion>(temp);
                        }
                        streamwiter.Flush();
                        return File(memory.ToArray(), "text/csv", "data.csv");

                    }
                    catch
                    {
                        return View("Error");
                    }
                    

                }
            }
            catch
            {
                return View("Error");
            }
            return null;
        }
        public ActionResult ProductAddition()
        {
            return View();
        }
        [HttpPost]
        public ActionResult ProductAddition(HttpPostedFileBase file)
        {   

            string path = null;
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            List<WarningTemplates> warningTemp = new List<WarningTemplates>();
            AllErrors allErrors = new AllErrors();
            allErrors.Error = null;
            allErrors.ShowHeader = null;
            allErrors.WarnHeaders = null;
            allErrors.Warning = null;
            int c1 = 0;
            int c2 = 0;
            try
            {
                if (file.ContentLength > 0)
                {

                    string filename = Path.GetFileName(file.FileName);

                    path = Server.MapPath("~/Seed Data/" + filename);
                    file.SaveAs(path);
                    ValidationChecker mapChecker = new ValidationChecker();
                    List<string> fileHeadersWarning = new List<string>();
                    List<string> fileHeadersError = new List<string>();
                    fileHeadersWarning.Add("PrimaryCategory".ToLower());
                    fileHeadersWarning.Add("SecondaryCategory".ToLower());
                    fileHeadersWarning.Add("Product".ToLower());
                    fileHeadersWarning.Add("Variant".ToLower());
                    fileHeadersWarning.Add("Price".ToLower());
                    fileHeadersWarning.Add("Unit".ToLower());
                    fileHeadersWarning.Add("DisplayCategory".ToLower());
                    fileHeadersWarning.Add("Image".ToLower());
                    fileHeadersWarning.Add("Description".ToLower());
                    fileHeadersWarning.Add("StandardUnitConversionFactor".ToLower());
                    fileHeadersWarning.Add("StandardUnit".ToLower());
                    fileHeadersWarning.Add("ProductCode".ToLower());
                    fileHeadersWarning.Add("VariantCode".ToLower());
                    fileHeadersError.Add("PrimaryCategory".ToLower());
                    fileHeadersError.Add("SecondaryCategory".ToLower());
                    fileHeadersError.Add("Product".ToLower());
                    fileHeadersError.Add("Unit".ToLower());

                    var xfile = new FileInfo(path);
                    ExcelPackage package = new ExcelPackage(xfile);
                    ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                    var fileField = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
                    List<string> headerError = new List<string>();
                    for (int i = sheet.Dimension.Start.Column; i <= sheet.Dimension.End.Column; i++)
                    {
                        var value = sheet.Cells[1, i].Value;
                        if (value != null)
                        {
                            string newValue = value.ToString().Trim().Replace(" ", "").ToLower();
                            if (newValue == "PrimaryCategory".ToLower())
                            {
                                headerError.Add(newValue);
                            }
                            if (newValue == "SecondaryCategory".ToLower())
                            {
                                headerError.Add(newValue);
                            }
                            if (newValue == "Product".ToLower())
                            {
                                headerError.Add(newValue);
                            }
                            if (newValue == "Price".ToLower())
                            {
                                headerError.Add(newValue);
                            }
                            if (newValue == "Unit".ToLower())
                            {
                                headerError.Add(newValue);
                            }
                        }
                    }
                    var difference = fileHeadersError.Except(headerError);
                    if (difference.Any())
                    {
                        ErrorTemplates error = new ErrorTemplates();
                        error.ErrorType = "Missing Header file";
                        error.ErrorComments = "Cannot go further Ensure these headers must be present" + string.Join(",", difference);
                        errorTemp.Add(error);
                        allErrors.Error = errorTemp;

                        return View(allErrors);


                    }
                    List<string> headerCheck = new List<string>();
                    for (int i = 1; i <= sheet.Dimension.End.Column; i++)
                    {
                        if (sheet.Cells[1, i].Value != null)
                        {
                            headerCheck.Add(sheet.Cells[1, i].Value.ToString().Replace(" ", "").ToLower());
                        }

                    }
                    var difference2 = fileHeadersWarning.Except(headerCheck);
                    if (difference2.Any())
                    {
                        WarningTemplates warningTemplates = new WarningTemplates();
                        warningTemplates.Comments = "These Headers are missing do you wish to continue";
                        warningTemplates.Field = string.Join(",", difference2);
                        warningTemp.Add(warningTemplates);

                    }
              
                    List<FileHeadersProductAddition> list = new List<FileHeadersProductAddition>();
                    List<FileHeadersBeatHierarchy> list2 = new List<FileHeadersBeatHierarchy>();
                    List<FileHeadersLocationAddtion> list3 = new List<FileHeadersLocationAddtion>();
                    List<FileHeadersBeatPlanAddition> list4 = new List<FileHeadersBeatPlanAddition>();

                    for (int i = sheet.Dimension.Start.Row+1; i <= sheet.Dimension.End.Row; i++)
                    {
                        FileHeadersProductAddition records = new FileHeadersProductAddition();
                        for (int j = sheet.Dimension.Start.Column; j <= sheet.Dimension.End.Column; j++)
                        {
                            var value = sheet.Cells[1, j].Value;
                            if (value != null)
                            {
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "PrimaryCategory".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.PrimaryCategory = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.PrimaryCategory = "";
                                        records.Row = i;
                                    }
                                    if (records.PrimaryCategory.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "PrimaryCategory";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "SecondaryCategory".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.SecondaryCategory = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.SecondaryCategory = "";
                                        records.Row = i;
                                    }
                                    if (records.SecondaryCategory.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "SecondaryCategory";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "Product".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.Product = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.Product = "";
                                        records.Row = i;
                                    }
                                    if (records.Product.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "Product";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "Variant".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.Variant = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.Variant = "";
                                        records.Row = i;
                                    }
                                    if (records.Variant.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "Variant";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "Price".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.Price = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.Price = "";
                                        records.Row = i;
                                    }
                                    if (records.Price.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "Price";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "Unit".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.Unit = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.Unit = "";
                                        records.Row = i;
                                    }
                                    if (records.Unit.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "Unit";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "DisplayCategory".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.DisplayCategory = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.DisplayCategory = "";
                                        records.Row = i;
                                    }
                                    if (records.DisplayCategory.Replace(" ", "").Count() >= 30)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "DisplayCategory";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "Image".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.Image = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.Image = "";
                                        records.Row = i;
                                    }
                                    if (records.Image.Replace(" ", "").Count() >= 40)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "Image";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "Description".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.Description = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.Description = "";
                                        records.Row = i;
                                    }
                                    if (records.Description.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "Description";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "StandardUnitConversionFactor".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.StandardUnitConversionFactor = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.StandardUnitConversionFactor = "";
                                        records.Row = i;
                                    }
                                    if (records.StandardUnitConversionFactor.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "StandardUnitConversionFactor";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "StandardUnit".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.StandardUnit = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.StandardUnit = "";
                                        records.Row = i;
                                    }
                                    if (records.StandardUnit.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "StandardUnit";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "ProductCode".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.ProductCode = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.ProductCode = "";
                                        records.Row = i;
                                    }
                                    if (records.ProductCode.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ProductCode";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "VariantCode".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.VariantCode = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.VariantCode = "";
                                        records.Row = i;
                                    }
                                    if (records.VariantCode.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "VariantCode";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "ProductCategory".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.ProductCategory = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.ProductCategory = "";
                                        records.Row = i;    
                                    }
                                    if (records.ProductCategory.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ProductCategory";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                            }
                        }
                        list.Add(records);
                    }
                    errorTemp = errorTemp.GroupBy(a => new { a.Row, a.Field_1, a.Field_2, a.ErrorType, a.ErrorComments, a.IncorrectHeaderList }, (key, g) => g.FirstOrDefault()).ToList();
                    
                    errorTemp.AddRange(mapChecker.Checker(sheet,list2,list3,list,list4));
                   





                    if (errorTemp.Count() != 0)
                    {
                        errorTemp = errorTemp.Take(50).ToList();
                        allErrors.Error = errorTemp.GroupBy(a => new { a.Row, a.Field_1, a.Field_2, a.ErrorType, a.ErrorComments, a.IncorrectHeaderList }, (key, g) => g.FirstOrDefault()).ToList();
                        allErrors.Warning = warningTemp.OrderBy(x => x.Row).GroupBy(a => new { a.Row, a.Field, a.Comments }, (key, g) => g.FirstOrDefault()).ToList();
                    //    allErrors.WarnHeaders = null;
                        return View(allErrors);
                    }
                    MemoryStream memory = new MemoryStream();
                    StreamWriter streamwiter = new StreamWriter(memory);
                    var newCsv = new CsvWriter(streamwiter);
                    newCsv.WriteHeader<FileHeadersProductAddition>();

                    foreach (var item in list)
                    {
                        FileHeadersProductAddition temp = new FileHeadersProductAddition();
                        temp.Description = item.Description;
                        temp.DisplayCategory = item.DisplayCategory;
                        temp.Image = item.Image;
                        temp.Price = item.Price;
                        temp.PrimaryCategory = item.PrimaryCategory;
                        temp.Product = item.Product;
                        temp.ProductCategory = item.ProductCategory;
                        temp.ProductCode = item.ProductCode;
                        temp.SecondaryCategory = item.SecondaryCategory;
                        temp.StandardUnit = item.StandardUnit;
                        temp.StandardUnitConversionFactor = item.StandardUnitConversionFactor;
                        temp.Unit = item.Unit;
                        temp.Variant = item.Variant;
                        temp.VariantCode = item.VariantCode;
                        newCsv.WriteRecord<FileHeadersProductAddition>(temp);
                    }
                    streamwiter.Flush();
                    return File(memory.ToArray(), "text/csv", "data.csv");
                }
            }
            catch (Exception ex)
            {
                var some = c1;
                var some2 = c2;
                var excep = ex;
                return View("Error");
            }
            return null;
        }
        public ActionResult BeatPlanAddition()
        {
            return View();
        }
        [HttpPost]
        public ActionResult BeatPlanAddtion (HttpPostedFileBase file)
        {
            string path = null;
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            List<WarningTemplates> warningTemp = new List<WarningTemplates>();
            AllErrors allErrors = new AllErrors();
            allErrors.Error = null;
            allErrors.ShowHeader = null;
            allErrors.WarnHeaders = null;
            allErrors.Warning = null;
            int c1 = 0;
            int c2 = 0;
            try
            {
                if (file.ContentLength > 0)
                {

                    string filename = Path.GetFileName(file.FileName);

                    path = Server.MapPath("~/Seed Data/" + filename);
                    file.SaveAs(path);
                    ValidationChecker mapChecker = new ValidationChecker();
                    List<string> fileHeadersWarning = new List<string>();
                    List<string> fileHeadersError = new List<string>();
                    fileHeadersWarning.Add("ESM".ToLower());
                    fileHeadersWarning.Add("FinalBeatName".ToLower());
                    fileHeadersWarning.Add("BeatDay".ToLower());
                    fileHeadersWarning.Add("BeatPlanStartDate".ToLower());
                    fileHeadersError.Add("ESM".ToLower());
                    fileHeadersError.Add("FinalBeatName".ToLower());
                    

                    var xfile = new FileInfo(path);
                    ExcelPackage package = new ExcelPackage(xfile);
                    ExcelWorksheet sheet = package.Workbook.Worksheets[1];
                    var fileField = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
                    List<string> headerError = new List<string>();
                    for (int i = sheet.Dimension.Start.Column; i <= sheet.Dimension.End.Column; i++)
                    {
                        var value = sheet.Cells[1, i].Value;
                        if (value != null)
                        {
                            string newValue = value.ToString().Trim().Replace(" ", "").ToLower();
                            if (newValue == "ESM".ToLower())
                            {
                                headerError.Add(newValue);
                            }
                            if (newValue == "FinalBeatName".ToLower())
                            {
                                headerError.Add(newValue);
                            }
                           
                        }
                    }
                    var difference = fileHeadersError.Except(headerError);
                    if (difference.Any())
                    {
                        ErrorTemplates error = new ErrorTemplates();
                        error.ErrorType = "Missing Header file";
                        error.ErrorComments = "Cannot go further Ensure these headers must be present" + string.Join(",", difference);
                        errorTemp.Add(error);
                        allErrors.Error = errorTemp;

                        return View(allErrors);


                    }
                    List<string> headerCheck = new List<string>();
                    for (int i = 1; i <= sheet.Dimension.End.Column; i++)
                    {
                        if (sheet.Cells[1, i].Value != null)
                        {
                            headerCheck.Add(sheet.Cells[1, i].Value.ToString().Replace(" ", "").ToLower());
                        }

                    }
                    var difference2 = fileHeadersWarning.Except(headerCheck);
                    if (difference2.Any())
                    {
                        WarningTemplates warningTemplates = new WarningTemplates();
                        warningTemplates.Comments = "These Headers are missing do you wish to continue";
                        warningTemplates.Field = string.Join(",", difference2);
                        warningTemp.Add(warningTemplates);

                    }

                    List<FileHeadersProductAddition> list = new List<FileHeadersProductAddition>();
                    List<FileHeadersBeatHierarchy> list2 = new List<FileHeadersBeatHierarchy>();
                    List<FileHeadersLocationAddtion> list3 = new List<FileHeadersLocationAddtion>();
                    List<FileHeadersBeatPlanAddition> list4 = new List<FileHeadersBeatPlanAddition>();
                    for (int i = sheet.Dimension.Start.Row + 1; i <= sheet.Dimension.End.Row; i++)
                    {
                        FileHeadersBeatPlanAddition records = new FileHeadersBeatPlanAddition();
                        for (int j = sheet.Dimension.Start.Column; j <= sheet.Dimension.End.Column; j++)
                        {
                            var value = sheet.Cells[1, j].Value;
                            if (value != null)
                            {
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "ESM".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.ESM = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.ESM = "";
                                        records.Row = i;
                                    }
                                    if (records.ESM.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "PrimaryCategory";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "FinalBeatName".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.FinalBeatName = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.FinalBeatName = "";
                                        records.Row = i;
                                    }
                                    if (records.FinalBeatName.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "SecondaryCategory";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatPlanStartDate".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.BeatPlanStartDate = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.BeatPlanStartDate = "";
                                        records.Row = i;
                                    }
                                    if (records.BeatPlanStartDate.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "BeatPlanStartDate";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatPeriod".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.BeatPeriod = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.BeatPeriod = "";
                                        records.Row = i;
                                    }
                                    if (records.BeatPeriod.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "BeatPeriod";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatDay".ToLower())
                                {
                                    if (sheet.Cells[i, j].Value != null)
                                    {
                                        records.BeatDay = sheet.Cells[i, j].Value.ToString();
                                        records.Row = i;
                                    }
                                    else
                                    {
                                        records.BeatDay = "";
                                        records.Row = i;
                                    }
                                    if (records.BeatDay.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "Price";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                
                            }
                            list4.Add(records);
                        }
                        
                    }
                    errorTemp = errorTemp.GroupBy(a => new { a.Row, a.Field_1, a.Field_2, a.ErrorType, a.ErrorComments, a.IncorrectHeaderList }, (key, g) => g.FirstOrDefault()).ToList();

                    errorTemp.AddRange(mapChecker.Checker(sheet, list2, list3, list,list4));






                    if (errorTemp.Count() != 0)
                    {
                        errorTemp = errorTemp.Take(50).ToList();
                        allErrors.Error = errorTemp.GroupBy(a => new { a.Row, a.Field_1, a.Field_2, a.ErrorType, a.ErrorComments, a.IncorrectHeaderList }, (key, g) => g.FirstOrDefault()).ToList();
                        allErrors.Warning = warningTemp.OrderBy(x => x.Row).GroupBy(a => new { a.Row, a.Field, a.Comments }, (key, g) => g.FirstOrDefault()).ToList();
                        //    allErrors.WarnHeaders = null;
                        return View(allErrors);
                    }
                    MemoryStream memory = new MemoryStream();
                    StreamWriter streamwiter = new StreamWriter(memory);
                    var newCsv = new CsvWriter(streamwiter);
                    newCsv.WriteHeader<FileHeadersBeatPlanAddition>();

                    foreach (var item in list4)
                    {
                        FileHeadersBeatPlanAddition temp = new FileHeadersBeatPlanAddition();
                        temp.BeatDay = item.BeatDay;
                        temp.BeatPeriod = item.BeatPeriod;
                        temp.BeatPlanStartDate = item.BeatPlanStartDate;
                        temp.ESM = item.ESM;
                        temp.FinalBeatName = item.FinalBeatName;
                        newCsv.WriteRecord<FileHeadersBeatPlanAddition>(temp);
                    }
                    streamwiter.Flush();
                    return File(memory.ToArray(), "text/csv", "data.csv");
                }
            }
            catch (Exception ex)
            {
                var some = c1;
                var some2 = c2;
                var excep = ex;
                return View("Error");
            }
            return null;
        }
    }
    
}
