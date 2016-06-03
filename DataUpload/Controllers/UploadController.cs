using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;

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
            
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            List<WarningTemplates> warningTemp = new List<WarningTemplates>();
            AllErrors allErrors = new AllErrors();            
            try
            {

                if (file.ContentLength > 0)
                {

                    string filename = Path.GetFileName(file.FileName);

                    path = Path.Combine(Server.MapPath("~/Seed Data/"), filename);
                    file.SaveAs(path);
                    ValidationChecker mapChecker = new ValidationChecker();
                    List<string> fileHeaders = new List<string>();
                       fileHeaders.Add("NSM");                //---------------------cell1
                       fileHeaders.Add("NSMZone");
                       fileHeaders.Add("NSMEmailId");         //---------------------cell3
                       fileHeaders.Add("NSMSecondaryEmailId");//---------------------cell4
                       fileHeaders.Add("ZSM");                //---------------------cell5
                       fileHeaders.Add("ZSMEmailId");         //---------------------cell6
                       fileHeaders.Add("ZSMZone");            //---------------------cell7
                       fileHeaders.Add("ZSMSecondaryEmailId");//---------------------cell8
                       fileHeaders.Add("RSM");                //---------------------cell9
                       fileHeaders.Add("RSMEmailId");         //---------------------cell10
                       fileHeaders.Add("RSMSecondaryEmailId");
                       fileHeaders.Add("RSMZone");//---------------------cell11
                       fileHeaders.Add("ASM");                //---------------------cell12
                       fileHeaders.Add("ASMEmailId");         //---------------------cell13
                       fileHeaders.Add("ASMZone");            //---------------------cell14
                       fileHeaders.Add("ASMSecondaryEmailId");//---------------------cell15
                       fileHeaders.Add("ESM");                //---------------------cell16
                       fileHeaders.Add("ESMEmailId");         //---------------------cell17
                       fileHeaders.Add("ESMZone");            //---------------------cell18
                       fileHeaders.Add("ESMSecondaryEmailId");//---------------------cell19
                       fileHeaders.Add("ESMContactNumber");   //---------------------cell20
                       fileHeaders.Add("ESMHQ");
                       fileHeaders.Add("ESMErpId");//---------------------cell21
                       fileHeaders.Add("FinalBeatName");      //---------------------cell22
                       fileHeaders.Add("BeatErpId");          //---------------------cell23
                       fileHeaders.Add("BeatDistrict");       //---------------------cell24
                       fileHeaders.Add("BeatState");          //---------------------cell25
                       fileHeaders.Add("BeatZone");           //---------------------cell26
                       fileHeaders.Add("DistributorName");    //---------------------cell27
                       fileHeaders.Add("DistributorLocation");//---------------------cell28
                       fileHeaders.Add("DistributorErpId");   //---------------------cell29
                       fileHeaders.Add("DistributorEmailId"); //---------------------cell20*/
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

                            headerCheck.Add(((object[,])fileField.Value)[0, i].ToString().Replace(" ",""));


                        }
                        var difference = fileHeaders.Except(headerCheck);
                        if (difference.Any())
                        {
                            ErrorTemplates errorTemplates = new ErrorTemplates();
                            AllErrors tempErrors = new AllErrors();
                            errorTemplates.IncorrectHeaders = "Set the Headers correctly";
                            errorTemp.Add(errorTemplates);
                            tempErrors.Error = errorTemp;
                            tempErrors.Warning = warningTemp;
                            return View(tempErrors);
                        }

                        for (int i = sheet.Dimension.Start.Column - 1; i <= sheet.Dimension.End.Column - 1; i++) //To find Empty cells.(algo will be updated)
                        {
                            if (((object[,])fileField.Value)[0, i].ToString().Replace(" ","") == "FinalBeatName")
                            {
                                for (int j = sheet.Dimension.Start.Row - 1; j <= sheet.Dimension.End.Row - 1; j++)
                                {
                                    if (((object[,])fileField.Value)[j, i] == null)
                                    {
                                        var index = ((object[,])fileField.Value)[j, i];
                                        ErrorTemplates errorTemplates = new ErrorTemplates();
                                        AllErrors tempErrors = new AllErrors();
                                        errorTemplates.EmptyFinalBeatName = "There is an empty field";
                                        errorTemplates.Row = j + 1;
                                        errorTemp.Add(errorTemplates);
                                        tempErrors.Error = errorTemp;
                                        tempErrors.Warning = warningTemp;
                                        return View(tempErrors);
                                    }
                                }
                            }

                        }
                        for (int i = sheet.Dimension.Start.Column - 1; i <= sheet.Dimension.End.Column - 1; i++) //To find Empty cells.(algo will be updated)
                        {
                            if (((object[,])fileField.Value)[0, i].ToString().Replace(" ", "") == "ESM")
                            {
                                for (int j = sheet.Dimension.Start.Row - 1; j <= sheet.Dimension.End.Row - 1; j++)
                                {
                                    if (((object[,])fileField.Value)[j, i] == null)
                                    {
                                        var index = ((object[,])fileField.Value)[j, i];
                                        ErrorTemplates errorTemplates = new ErrorTemplates();
                                        AllErrors tempErrors = new AllErrors();
                                        errorTemplates.EmptyESM = "There is an empty field";
                                        errorTemp.Add(errorTemplates);
                                        tempErrors.Error = errorTemp;
                                        tempErrors.Warning = warningTemp;
                                        return View(tempErrors);
                                    }
                                }
                            }

                        }
                        warningTemp = mapChecker.WarningChecks(sheet, warningTemp);//Warnings
                        errorTemp = mapChecker.Checker(sheet,errorTemp);//Mapping checking
                        if(errorTemp!=null)
                        {
                            allErrors.Error = errorTemp.OrderBy(x=>x.Row).GroupBy(a => new { a.Row, a.Field_1, a.Field_2,a.EmptyESM,a.EmptyFinalBeatName,a.HierarchyBeak,a.IncorrectHeaders,a.MappingErrorType,a.PhoneError }, (key, g) => g.FirstOrDefault()).ToList();
                            allErrors.Warning = warningTemp.OrderBy(x=>x.Row).GroupBy(a=>new { a.Row,a.Field,a.Comments},(key,g)=>g.FirstOrDefault()).ToList();
                            return View(allErrors);
                        }
                        var fileinfo = new FileInfo(@"C:\Docs\Visual Studio 2015\Projects\DataUpload\DataUpload\Seed Data\file.csv");
                        package.SaveAs(fileinfo);
                        UploadController upload = new UploadController();
                        MemoryStream memory = new MemoryStream();                                        
                    }
                    catch (Exception ex)
                    {
                        string err = ex.Message;
                        View("Error");
                    }
                }
            }
            catch
            {
                View("Error");
            }


            return null;

        }
    }
    }
