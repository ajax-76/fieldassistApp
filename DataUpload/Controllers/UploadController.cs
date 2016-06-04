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
                              allErrors.Error = errorTemp.GroupBy(a => new { a.Row, a.Field_1, a.Field_2,a.EmptyESM,a.EmptyFinalBeatName,a.HierarchyBeak,a.IncorrectHeaders,a.MappingErrorType,a.PhoneError,a.LinkRow }, (key, g) => g.FirstOrDefault()).ToList();
                              allErrors.Warning = warningTemp.OrderBy(x=>x.Row).GroupBy(a=>new { a.Row,a.Field,a.Comments},(key,g)=>g.FirstOrDefault()).ToList();
                              return View(allErrors);
                          }

                        //Excel to object  

                        List<FileHeaders> list = new List<FileHeaders>();
                        
                        for (int i=sheet.Dimension.Start.Row;i<sheet.Dimension.End.Row;i++)
                        {
                            FileHeaders records = new FileHeaders();
                            for (int j=sheet.Dimension.Start.Column-1;j<sheet.Dimension.End.Column;j++)
                            {
                                if(((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "NSM")
                                {
                                    if(((object[,])fileField.Value)[i, j]!=null)
                                    {
                                        records.NSM = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.NSM = "";
                                    }
                                    
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "NSMZone")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.NSMZone = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.NSMZone = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "NSMEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.NSMEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.NSMEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "NSMSecondaryEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.NSMSecondaryEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.NSMSecondaryEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ZSM")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ZSM = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ZSM = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ZSMZone")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ZSMZone = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ZSMZone = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ZSMEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ZSMEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ZSMEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ZSMSecondaryEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ZSMSecondaryEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ZSMSecondaryEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "RSM")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.RSM = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.RSM = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "RSMZone")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.RSMZone = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.RSMZone = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "RSMEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.RSMEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.RSMEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "RSMSecondaryEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.RSMSecondaryEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.RSMSecondaryEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ASM")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ASM = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ASM = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ASMZone")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ASMZone = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ASMZone = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ASMEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ASMEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ASMEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ASMSecondaryEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ASMSecondaryEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ASMSecondaryEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ESM")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ESM = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ESM = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ESMEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ESMEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ESMEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ESMSecondaryEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ESMSecondaryEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ESMSecondaryEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ESMZone")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ESMZone = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ESMZone = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ESMContactNumber")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ESMContactNumber = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ESMContactNumber = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ESMHQ")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ESMHQ = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ESMHQ = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "ESMErpId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.ESMErpId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.ESMErpId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "FinalBeatName")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.FinalBeatName = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.FinalBeatName = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "BeatState")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.BeatState = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.BeatState = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "BeatDistrict")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.BeatDistrict = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.BeatDistrict = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "BeatZone")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.BeatZone = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.BeatZone = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "BeatErpId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.BeatErpId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.BeatErpId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "DistributorName")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.DistributorName = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.DistributorName = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "DistributorLocation")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.DistributorLocation = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.DistributorLocation = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "DistributorEmailId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.DistributorEmailId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.DistributorEmailId = "";
                                    }
                                }
                                if (((object[,])fileField.Value)[0, j].ToString().Trim().Replace(" ", "") == "DistributorErpId")
                                {
                                    if (((object[,])fileField.Value)[i, j] != null)
                                    {
                                        records.DistributorErpId = ((object[,])fileField.Value)[i, j].ToString();
                                    }
                                    else
                                    {
                                        records.DistributorErpId = "";
                                    }
                                }
                            }
                            list.Add(records);
                        }

                        //write csv

                        

                        MemoryStream memory = new MemoryStream();
                        StreamWriter streamwiter = new StreamWriter(memory);
                        var newCsv = new CsvWriter(streamwiter);
                        newCsv.WriteHeader<FileHeaders>();
                        
                        foreach (var item in list)
                        {
                            FileHeaders temp = new FileHeaders();
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
                            newCsv.WriteRecord<FileHeaders>(temp);
                        }
                        streamwiter.Flush();
                        return File(memory.ToArray(), "text/csv", "data.csv");                                        
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
