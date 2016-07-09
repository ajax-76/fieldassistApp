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
using FieldAssist.DataAccessLayer.Models.EFModels;
//using FieldAssist.DataAccessLayer.Repositories;
using DataUpload.DataUploadRepository;
using DataUpload.ExcelToObject;


namespace DataUpload.Controllers
{
    
   
    public class UploadController : Controller
    {

        // GET: Upload
        public ActionResult Upload()
        {
            return View();
        }
        //Passive Checks Beat Hierarchy
        #region
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
                        ExcelToObject.ExcelToObject excelToObject = new ExcelToObject.ExcelToObject();

                        ListandError listerror = new ListandError();

                        listerror=excelToObject.listWithErrorBeatHiearachy(sheet);
                        list = listerror.listBeatHierarchy;
                        errorTemp.AddRange(listerror.error);
                        errorTemp.AddRange(mapChecker.Checker(sheet, list, list3, list2,list4));
                        
                       if (errorTemp != null)
                        {
                            allErrors.Error = errorTemp.GroupBy(a => new { a.Row, a.Field_1, a.Field_2, a.ErrorType, a.ErrorComments, a.IncorrectHeaderList }, (key, g) => g.FirstOrDefault()).ToList();
                            allErrors.Warning = warningTemp.OrderBy(x => x.Row).GroupBy(a => new { a.Row, a.Field, a.Comments }, (key, g) => g.FirstOrDefault()).ToList();
                            return View(allErrors);
                        }
                        //write csv

                        Session["BeatHierarchy"]=list;
                        return RedirectToAction("CompanySelect");
                       
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
        #endregion  
        CompanyDataContext db = new CompanyDataContext();
        public ActionResult CompanySelect()
        {
            List<Company> list = db.Companies.ToList();
            
            return View(list);
        }
        //BeatHierarchy Addition--Active Checks
        #region
        [HttpPost]
        public ActionResult BeatHierarchyAdditionCheck(string ID )
        {
            int? Id = int.Parse(ID);
            Session["CompanyID"] = Id;
            List<FileHeadersBeatHierarchy> beatHierarchyList = (List<FileHeadersBeatHierarchy>)Session["BeatHierarchy"];           
            NSMRepository NSMRepo = new NSMRepository();
            List<NationalSalesManager> NSMList = NSMRepo.GetNSM(Id);
            ZSMRepository ZSMRepo = new ZSMRepository();
            List<ZonalSalesManager> ZSMList = ZSMRepo.GetZSM(Id);
            RSMRepository RSMRepo = new RSMRepository();
            List<RegionalSalesManager> RSMList = RSMRepo.GetRSM(Id);
            ASMRepository ASMRepo = new ASMRepository();
            List<AreaSalesManager> ASMList = ASMRepo.GetASM(Id);
            ESMRepository ESMRepo = new ESMRepository();
            List<ClientEmployee> ESMList = ESMRepo.GetESM(Id);
            BEATRepository BEATRepo = new BEATRepository();
            List<LocationBeat> BEATList = BEATRepo.GetBEAT(Id);
            DistributorRepository DistributorRepo = new DistributorRepository();
            List<Distributor> DIstributorList = DistributorRepo.GetDistributor(Id);
            List<FACompanyZone> CompanyZone = db.FACompanyZones.Where(x => x.CompanyId == Id).ToList();
            List<ErrorTemplates> errorList = new List<ErrorTemplates>();
            foreach(var item in beatHierarchyList)
            {
               
                if(NSMList.Exists(x=>x.Name==item.NSM))
                {
                    ErrorTemplates er = new ErrorTemplates();
                    er.ErrorComments ="NSM : "+ item.NSM + " already exists cannot be added";
                    errorList.Add(er);
                }
                
                if(ZSMList.Exists(x=>x.Name==item.ZSM))
                {
                    ErrorTemplates er = new ErrorTemplates();
                    er.ErrorComments = "ZSM : " + item.ZSM + " already exists cannot be added";
                    errorList.Add(er);
                }
                
                if(RSMList.Exists(x=>x.Name==item.RSM))
                {
                    ErrorTemplates er = new ErrorTemplates();
                    er.ErrorComments = "RSM : " + item.RSM + " already exists cannot be added";
                    errorList.Add(er);
                }
                
                if(ASMList.Exists(x=>x.Name==item.ASM))
                {
                    ErrorTemplates er = new ErrorTemplates();
                    er.ErrorComments = "ASM : " + item.ASM + " already exists cannot be added";
                    errorList.Add(er);
                }
                
                if(ESMList.Exists(x=>x.Name==item.ESM))
                {
                    ErrorTemplates er = new ErrorTemplates();
                    er.ErrorComments = "ESM : " + item.ESM + " already exists cannot be added";
                    errorList.Add(er);
                }
               
                if(BEATList.Exists(x=>x.Name==item.FinalBeatName))
                {
                    ErrorTemplates er = new ErrorTemplates();
                    er.ErrorComments = "FinalBeatName : " + item.FinalBeatName + " already exists cannot be added";
                    errorList.Add(er);
                }
                if (BEATList.Exists(x => x.ErpId == item.BeatErpId) && item.BeatErpId!=null)
                {
                    ErrorTemplates er = new ErrorTemplates();
                    er.ErrorComments = "BeatErpId : " + item.BeatErpId +" cannot be assigned as it already existing ";
                    errorList.Add(er);
                }

                if (DIstributorList.Exists(x=>x.Name==item.DistributorName))
                {
                    ErrorTemplates er = new ErrorTemplates();
                    er.ErrorComments = "Distrbutor : " + item.DistributorName + " already exists cannot be added";
                    errorList.Add(er);
                }
                if(CompanyZone.Exists(x=>x.Name==item.ESMZone))
                {
                    ErrorTemplates er = new ErrorTemplates();
                    er.ErrorComments = "ESMZone : " + item.ESMZone + " already exists cannot be added";
                    errorList.Add(er);
                }
            }

            return View(errorList);
        }
        public ActionResult BeatHierarchyAddition()
        {
            List<FileHeadersBeatHierarchy> list = (List<FileHeadersBeatHierarchy>)Session["BeatHierarchy"];
            int? id = (int?)Session["CompanyID"];
            int newId = id ?? default(int);
            var list1 = list.GroupBy(x => x.NSM);


            foreach (var keyitem in list1)
            {
                foreach (var item in keyitem)
                {
                    NationalSalesManager NSM = new NationalSalesManager
                    {
                        Name = keyitem.Key,
                        EmailId = item.NSMEmailId,
                        SecondaryEmailId = item.NSMSecondaryEmailId,
                        Zone = item.NSMZone,
                        CompanyId = newId
                    };
                    NSMRepository NSMRepo = new NSMRepository();
                    List<NationalSalesManager> NSMList = NSMRepo.GetNSM(newId);

                    if (!NSMList.Exists(x=>x.Name==NSM.Name))
                    {
                        NSMRepo.AddNSM(NSM);
                    }
                    //Insert Regions
                    long NSMPk = db.NationalSalesManagers.Where(x => x.Name == item.NSM).Select(x => x.Id).FirstOrDefault();

                    ZonalSalesManager ZSM = new ZonalSalesManager();
                    ZSM.Name = item.ZSM;
                    ZSM.EmailId = item.ZSMEmailId;
                    ZSM.SecondaryEmailId = item.ZSMSecondaryEmailId;
                    ZSM.Zone = item.ZSMZone;
                    ZSM.NationalSalesManagerId = NSMPk;
                    ZSM.CompanyId = newId;


                    ZSMRepository ZSMRepo = new ZSMRepository();
                    List<ZonalSalesManager> ZSMList = ZSMRepo.GetZSM(newId);
                    if (!ZSMList.Exists(x=>x.Name==ZSM.Name))
                    {
                        ZSMRepo.AddZSM(ZSM);
                    }

                    long ZSMPk = db.ZonalSalesManagers.Where(x => x.Name == item.ZSM).Select(x => x.Id).FirstOrDefault();

                    RegionalSalesManager RSM = new RegionalSalesManager();
                    RSM.Name = item.RSM;
                    RSM.EmailId = item.RSMEmailId;
                    RSM.SecondaryEmailId = item.RSMSecondaryEmailId;
                    RSM.Zone = item.RSMZone;
                    RSM.ZonalSalesManagerId = ZSMPk;
                    RSM.CompanyId = newId;

                    RSMRepository RSMRepo = new RSMRepository();
                    List<RegionalSalesManager> RSMList = RSMRepo.GetRSM(newId);
                    if(!RSMList.Exists(x=>x.Name==RSM.Name))
                    {
                        RSMRepo.AddRSM(RSM);
                    }

                    long RSMPk = db.RegionalSalesManagers.Where(x => x.Name == item.RSM).Select(x => x.Id).FirstOrDefault();

                    AreaSalesManager ASM = new AreaSalesManager();
                    ASM.Name = item.ASM;
                    ASM.EmailId = item.ASMEmailId;
                    ASM.SecondaryEmailId = item.ASMSecondaryEmailId;
                    ASM.Zone = item.ASMZone;
                    ASM.RegionalSalesManagerId = RSMPk;
                    ASM.CompanyId = newId;

                    ASMRepository ASMRepo = new ASMRepository();
                    List<AreaSalesManager> ASMList = ASMRepo.GetASM(newId);
                    if(!ASMList.Exists(x=>x.Name==ASM.Name))
                    {
                        ASMRepo.AddASM(ASM);
                    }

                    long ASMPk = db.ClientEmployees.Where(x => x.Name == item.ASM).Select(x => x.Id).FirstOrDefault();

                    ClientEmployee ESM = new ClientEmployee();
                    ESM.Name = item.ESM;
                    ESM.EmailId = item.ESMEmailId;
                    ESM.SecondaryEmailId = item.ESMSecondaryEmailId;
                 //   ESM.Zone = item.ESMZone;
                    ESM.ClientSideId = item.ESMErpId;
                    ESM.AreaSalesManagerId = ASMPk;
                    ESM.Company = newId;

                    ESMRepository ESMRepo = new ESMRepository();
                    List<ClientEmployee> ESMList = ESMRepo.GetESM(newId);
                    if(!ESMList.Exists(x=>x.Name==ESM.Name))
                    {
                        ESMRepo.AddESM(ESM);
                    }

                    LocationBeat BEAT = new LocationBeat();
                    BEAT.Name = item.FinalBeatName;
                    BEAT.ErpId = item.BeatErpId;
                    BEAT.City = item.BeatDistrict;
                    BEAT.State = item.BeatState;
                    BEAT.DivisionZone = item.BeatZone;

                    BEATRepository BEATRepo = new BEATRepository();
                    List<LocationBeat> BEATList = BEATRepo.GetBEAT(newId);

                    if(!BEATList.Exists(x=>x.Name==BEAT.Name))
                    {
                        BEATRepo.AddBEAT(BEAT);
                    }
                    Distributor Distributor = new Distributor();
                    Distributor.Name = item.DistributorName;
                    Distributor.EmailId = item.DistributorEmailId;
                    Distributor.ClientSideId = item.DistributorErpId;
                    Distributor.Place = item.DistributorLocation;

                    DistributorRepository DistributorRepo = new DistributorRepository();
                    List<Distributor> DIstributorList = DistributorRepo.GetDistributor(newId);

                    if(!DIstributorList.Exists(x=>x.Name==Distributor.Name))
                    {
                        DistributorRepo.AddDistributor(Distributor);
                    }

                    DistributorBeatMappingRepository DBMRepo = new DistributorBeatMappingRepository();
                    List<DistributorBeatMapping> DBMList = DBMRepo.GetDistributorBeatMap(newId);

                    long BeatId = db.LocationBeats.Where(x => x.Name == item.FinalBeatName).Select(x => x.Id).FirstOrDefault();
                    long DistributorId = db.Distributors.Where(x => x.Name == item.DistributorName).Select(x => x.Id).FirstOrDefault();

                    DistributorBeatMapping DBM = new DistributorBeatMapping();
                    DBM.BeatId = BeatId;
                    DBM.DistributorId = DistributorId;
                    DBM.CompanyId = newId;

                    if (!DBMList.Contains(DBM))
                    {
                        DBMRepo.AddDistributorBeatMap(DBM);
                    }


                }
            }
            return View();
        }
        #endregion


        //Program for Location Addition Passive Checks..

        public ActionResult Upload_LocationAddition()
        {
            return View();
        }
        #region
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
                    WarningfileHeaders.Add("ShopType".ToLower());
                    WarningfileHeaders.Add("Segmentation".ToLower());
                    WarningfileHeaders.Add("OwnersName".ToLower());
                    WarningfileHeaders.Add("OwnersContactNumber".ToLower());
                    WarningfileHeaders.Add("FinalBeatName".ToLower());
                    WarningfileHeaders.Add("ShopErpId".ToLower());
                    WarningfileHeaders.Add("ISFocused".ToLower());
                    WarningfileHeaders.Add("ConsumerType".ToLower());
                    WarningfileHeaders.Add("OutletPotential".ToLower());
                    WarningfileHeaders.Add("CodeId".ToLower());
                    ErrorfileHeaders.Add("FinalBeatName".ToLower());
                    ErrorfileHeaders.Add("ShopName".ToLower());
                    ErrorfileHeaders.Add("Market".ToLower());
                    ErrorfileHeaders.Add("City".ToLower());
                    ErrorfileHeaders.Add("State".ToLower());
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
                                
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "FinalBeatName".ToLower())
                                {
                                    headerError.Add(value.ToString().Trim().Replace(" ", "").ToLower());
                                }                              
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "ShopName".ToLower())
                                {
                                    headerError.Add(value.ToString().Trim().Replace(" ", "").ToLower());
                                }
                                if (value.ToString().Trim().Replace(" ", "").ToLower() == "Market".ToLower()) 
                                {
                                    headerError.Add(value.ToString().Trim().Replace(" ", "").ToLower());
                                }
                                if(value.ToString().Trim().Replace(" ", "").ToLower()=="City")
                                {
                                    headerError.Add(value.ToString().Trim().Replace(" ", "").ToLower());
                                }
                                if(value.ToString().Trim().Replace(" ", "").ToLower()=="State")
                                {
                                    headerError.Add(value.ToString().Trim().Replace(" ", "").ToLower());
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

                        ExcelToObject.ExcelToObject excelToObject = new ExcelToObject.ExcelToObject();

                        ListandError listerror = new ListandError();
                        listerror = excelToObject.listErrorLocationAddition(sheet);

                        errorTemp.AddRange(listerror.error);
                        list = listerror.listLocation;


                        //    warningTemp = mapChecker.WarningChecks(sheet, warningTemp);//Warnings
                        errorTemp.AddRange(mapChecker.Checker(sheet,list2,list,list3,list4));//Mapping checking

                        if (errorTemp != null)
                        {
                            allErrors.Error = errorTemp.GroupBy(a => new { a.Row, a.Field_1, a.Field_2, a.ErrorType, a.ErrorComments, a.IncorrectHeaderList }, (key, g) => g.FirstOrDefault()).ToList();
                            allErrors.Warning = warningTemp.OrderBy(x => x.Row).GroupBy(a => new { a.Row, a.Field, a.Comments }, (key, g) => g.FirstOrDefault()).ToList();
                           // allErrors.ShowHeader = null;
                            return View(allErrors);
                        }
                        Session["LocationAdditionObjectList"] = list;

                        return RedirectToAction("CompanySelect");

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
        #endregion
            [HttpPost]
          public ActionResult LocationAddition_AdditionCheck(string ID)
          {
            int id = int.Parse(ID);
            Session["companyId"] = id;
            List<FileHeadersLocationAddtion> list = (List<FileHeadersLocationAddtion>)Session["LocationAdditionObjectList"];

            LocationAdditionRepository LOCRepo = new LocationAdditionRepository();
            List<Location> LOCList = LOCRepo.GetLocation(id);
            List<ErrorTemplates> errorList = new List<ErrorTemplates>();

            foreach(var item in list)
            {

            }
            return null;


        }

        //Passive Checks ProductAddition
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

                    ExcelToObject.ExcelToObject excelToObject = new ExcelToObject.ExcelToObject();

                        ListandError listerror = new ListandError();

                        listerror=excelToObject.listErrrorProductAddition(sheet);
                        list = listerror.listProductAddition;
                        errorTemp.AddRange(listerror.error);

                  
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

        //Passive Checks BeatPlanAddition

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
                    ExcelToObject.ExcelToObject excelToObject = new ExcelToObject.ExcelToObject();

                    ListandError listerror = new ListandError();

                    listerror = excelToObject.listErrrorProductAddition(sheet);
                    list = listerror.listProductAddition;
                    errorTemp.AddRange(listerror.error);

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
