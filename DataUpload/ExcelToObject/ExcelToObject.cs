using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OfficeOpenXml;

namespace DataUpload.ExcelToObject
{
    public class ExcelToObject
    {
        public ListandError listWithErrorBeatHiearachy(ExcelWorksheet sheet)
        {
            var fileField = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
            
            ListandError listerror = new ListandError();
            List<FileHeadersBeatHierarchy> list = new List<FileHeadersBeatHierarchy>();
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            int startrow = sheet.Dimension.Start.Row;
            int endrow = sheet.Dimension.End.Row;
            int startcolomn = sheet.Dimension.Start.Column;
            int endcolomn = sheet.Dimension.End.Column;
            try
            {
                for (int i = startrow + 1; i <= endrow; i++)
                {
                    FileHeadersBeatHierarchy records = new FileHeadersBeatHierarchy();
                    for (int j = startcolomn; j <= endcolomn; j++)
                    {
                        var value = sheet.Cells[1, j].Value;
                        if (value != null)
                        {
                            var newValue = sheet.Cells[i, j].Value;
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "NSM".ToLower())
                            {

                                if (newValue != null)
                                {
                                    records.NSM = newValue.ToString();
                                    records.Row = i;
                                    if (records.NSM.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "NSM";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }

                                }
                                else
                                {
                                    records.NSM = null;
                                    records.Row = i;
                                }


                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "NSMZone".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.NSMZone = newValue.ToString();
                                    records.Row = i;
                                    if (records.NSMZone.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "NSMZone";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }
                                }
                                else
                                {
                                    records.NSMZone = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "NSMEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.NSMEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.NSMEmailId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "NSMEmailId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.NSMEmailId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "NSMSecondaryEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.NSMSecondaryEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.NSMSecondaryEmailId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "NSMSecondaryEmaiId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.NSMSecondaryEmailId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ZSM".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ZSM = newValue.ToString();
                                    records.Row = i;
                                    if (records.ZSM.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ZSM";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ZSM = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ZSMZone".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ZSMZone = newValue.ToString();
                                    records.Row = i;
                                    if (records.ZSMZone.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ZSMZone";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ZSMZone = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ZSMEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ZSMEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.ZSMEmailId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ZSMEmailId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ZSMEmailId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ZSMSecondaryEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ZSMSecondaryEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.ZSMSecondaryEmailId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ZSMSecondaryEmailId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ZSMSecondaryEmailId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "RSM".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.RSM = newValue.ToString();
                                    records.Row = i;
                                    if (records.RSM.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "RSM";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.RSM = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "RSMZone".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.RSMZone = newValue.ToString();
                                    records.Row = i;
                                    if (records.RSMZone.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "RSMZone";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.RSMZone = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "RSMEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.RSMEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.RSMEmailId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "RSMEmailId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.RSMEmailId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "RSMSecondaryEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.RSMSecondaryEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.RSMSecondaryEmailId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "RSMSecondaryEmailId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.RSMSecondaryEmailId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ASM".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ASM = newValue.ToString();
                                    records.Row = i;
                                    if (records.ASM.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ASM";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ASM = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ASMZone".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ASMZone = newValue.ToString();
                                    records.Row = i;
                                    if (records.ASMZone.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ASMZone";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ASMZone = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ASMEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ASMEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.ASMEmailId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ASMEmailId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ASMEmailId = "";
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ASMSecondaryEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ASMSecondaryEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.ASMSecondaryEmailId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ASMSecondaryEmailId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ASMSecondaryEmailId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ESM".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ESM = newValue.ToString();
                                    records.Row = i;
                                    if (records.ESM.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ESM";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ESM = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ESMEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ESMEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.ESMEmailId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ESMEmailId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ESMEmailId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ESMSecondaryEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ESMSecondaryEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.ESMSecondaryEmailId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ESMSecondaryEmailId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ESMSecondaryEmailId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ESMZone".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ESMZone = newValue.ToString();
                                    records.Row = i;
                                    if (records.ESMZone.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ESMZone";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.ESMZone = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ESMContactNumber".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ESMContactNumber = newValue.ToString();
                                    records.Row = i;
                                }
                                else
                                {
                                    records.ESMContactNumber = null;
                                    records.Row = i;
                                }
                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ESMHQ".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ESMHQ = newValue.ToString();
                                    records.Row = i;
                                }
                                else
                                {
                                    records.ESMHQ = null;
                                    records.Row = i;
                                }
                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ESMErpId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ESMErpId = newValue.ToString();
                                    records.Row = i;
                                }
                                else
                                {
                                    records.ESMErpId = null;
                                    records.Row = i;
                                }
                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "FinalBeatName".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.FinalBeatName = newValue.ToString();
                                    records.Row = i;
                                    if (records.FinalBeatName.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "FinalBeatName";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.FinalBeatName = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatState".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.BeatState = newValue.ToString();
                                    records.Row = i;
                                }
                                else
                                {
                                    records.BeatState = null;
                                    records.Row = i;
                                }
                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatDistrict".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.BeatDistrict = newValue.ToString();
                                    records.Row = i;
                                }
                                else
                                {
                                    records.BeatDistrict = null;
                                    records.Row = i;
                                }
                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatZone".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.BeatZone = newValue.ToString();
                                    records.Row = i;
                                    if (records.BeatZone.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "BeatZone";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.BeatZone = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatErpId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.BeatErpId = newValue.ToString();
                                    records.Row = i;
                                    if (records.BeatErpId.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "BeatErpId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }


                                }
                                else
                                {
                                    records.BeatErpId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "DistributorName".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.DistributorName = newValue.ToString();
                                    records.Row = i;
                                    if (records.DistributorName.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "DistributorName";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.DistributorName = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "DistributorLocation".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.DistributorLocation = newValue.ToString();
                                    records.Row = i;
                                    if (records.DistributorLocation.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "DistributorLocation";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.DistributorLocation = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "DistributorEmailId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.DistributorEmailId = newValue.ToString();
                                    records.Row = i;
                                    if (records.DistributorEmailId.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "DistributorEmailId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.DistributorEmailId = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "DistributorErpId".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.DistributorErpId = newValue.ToString();
                                    records.Row = i;
                                    if (records.DistributorErpId.Replace(" ", "").Count() >= 100)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "DistributorErpId";
                                        temp.Row = i;
                                        errorTemp.Add(temp);
                                    }

                                }
                                else
                                {
                                    records.DistributorErpId = null;
                                    records.Row = i;
                                }

                            }
                        }
                    }
                    list.Add(records);
                }
            }
            catch(Exception ex)
            {
                var x = ex.Message;
            }
            listerror.error = errorTemp;
            listerror.listBeatHierarchy = list;
            return listerror;
        }

        public ListandError listErrorLocationAddition(ExcelWorksheet sheet)
        {
            var fileField = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
            ListandError listerror = new ListandError();
            List<FileHeadersLocationAddtion> list = new List<FileHeadersLocationAddtion>();
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();

            int startrow = sheet.Dimension.Start.Row;
            int endrow = sheet.Dimension.End.Row;
            int startcolomn = sheet.Dimension.Start.Column;
            int endcolomn = sheet.Dimension.End.Column;

            for (int i = startrow +1; i <= endrow; i++)
            {
                FileHeadersLocationAddtion records = new FileHeadersLocationAddtion();
                for (int j = startrow; j <=endcolomn; j++)
                {
                    var value = sheet.Cells[1, j].Value;
                    if (value != null)
                    {
                        var newValue = sheet.Cells[i, j].Value;
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "ShopName".ToLower())
                        {
                            if (newValue!= null)
                            {
                                records.ShopName = newValue.ToString();
                                records.Row = i;
                                if (records.ShopName.Replace(" ", "").Count() >= 100)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "ShopName";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.ShopName = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Address".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Address = newValue.ToString();
                                records.Row = i;
                                if (records.Address.Replace(" ", "").Count() >= 500)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Address";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Address = null;
                                records.Row = i;
                            }
                           
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Email".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Email = newValue.ToString();
                                records.Row = i;
                                if (records.Email.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Email";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Email = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Tin".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Tin = newValue.ToString();
                                records.Row = i;
                                if (records.Tin.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Tin";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Tin = null;
                                records.Row = i;
                            }
                           
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Pin".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Pin = newValue.ToString();
                                records.Row = i;
                                if (records.Pin.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Pin";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Pin = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "MarketName".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.MarketName = newValue.ToString();
                                records.Row = i;
                                if (records.MarketName.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "MarketName";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.MarketName = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "City".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.City = newValue.ToString();
                                records.Row = i;
                                if (records.City.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "City";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.City = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "State".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.State = newValue.ToString();
                                records.Row = i;
                                if (records.State.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "State";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.State = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Country".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Country = newValue.ToString();
                                records.Row = i;
                                if (records.Country.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Country";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Country = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "ShopCode".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.ShopCode = newValue.ToString();
                                records.Row = i;
                                if (records.ShopCode.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "ShopCode";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.ShopCode = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "ShopType".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.ShopType = newValue.ToString();
                                records.Row = i;
                                if (records.ShopType.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "ShopType";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.ShopType = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Segmentation".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Segmentation = newValue.ToString();
                                records.Row = i;
                                if (records.Segmentation.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Segmentation";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Segmentation = null;
                                records.Row = i;
                            }
                           
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "OwnersName".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.OwnersName = newValue.ToString();
                                records.Row = i;
                                if (records.OwnersName.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "OwnersName";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.OwnersName = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "OwnersContactNumber".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.OwnersContactNumber = newValue.ToString();
                                records.Row = i;
                                if (records.OwnersContactNumber.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "OwnersContactNumber";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.OwnersContactNumber = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "FinalBeatName".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.FinalBeatName = newValue.ToString();
                                records.Row = i;
                                if (records.FinalBeatName.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "FinalBeatName";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.FinalBeatName = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatErpId".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.BeatErpId = newValue.ToString();
                                records.Row = i;
                                if (records.BeatErpId.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "BeatErpId";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.BeatErpId = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "ISFocused".ToLower())
                        {
                            if (newValue != null)
                            {
                                if (newValue.ToString().ToLower() == "true" || newValue.ToString().ToLower() == "false")
                                {
                                    records.ISFocused = ((object[,])fileField.Value)[i, j].ToString();
                                    records.Row = i;
                                }
                                else
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "ISFocused is only true false type";
                                    temp.Field_1 = "ISFocused";
                                    temp.Row = i;
                                    errorTemp.Add(temp);
                                }
                            }
                            else
                            {
                                records.ISFocused = null;
                                records.Row = i;
                            }


                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "ConsumerType".ToLower())
                        {
                            if (newValue != null)
                            {
                                if (newValue.ToString().ToLower() == "shop" || newValue.ToString().ToLower() == "persosn")
                                {
                                    records.ConsumerType = ((object[,])fileField.Value)[i, j].ToString();
                                    records.Row = i;
                                }
                                else
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "ConsumerType is only shop or person";
                                    temp.Field_1 = "ConsumerType";
                                    temp.Row = i;
                                    errorTemp.Add(temp);
                                }
                            }
                            else
                            {
                                records.ConsumerType = null;
                                records.Row = i;
                            }


                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "OutletPotential".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.OutletPotential = newValue.ToString();
                                records.Row = i;
                                if (records.OutletPotential.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "State";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.OutletPotential = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "CodeId".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.CodeId = newValue.ToString();
                                records.Row = i;
                                if (records.State.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "State";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.CodeId = null;
                                records.Row = i;
                            }
                           
                        }


                    }
                    
                }
                list.Add(records);
            }
            listerror.error = errorTemp;
            listerror.listLocation = list;
            return listerror;
        }
            
        public ListandError listErrrorProductAddition(ExcelWorksheet sheet)
        {
            var fileField = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
            ListandError listerror = new ListandError();
            List<FileHeadersProductAddition> list = new List<FileHeadersProductAddition>();
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            int startrow = sheet.Dimension.Start.Row;
            int endrow = sheet.Dimension.End.Row;
            int startcolomn = sheet.Dimension.Start.Column;
            int endcolomn = sheet.Dimension.End.Column;
            for (int i = startrow + 1; i <= endrow; i++)
            {
                FileHeadersProductAddition records = new FileHeadersProductAddition();
                for (int j = startcolomn; j <= endcolomn; j++)
                {
                    var value = sheet.Cells[1, j].Value;
                    if (value != null)
                    {
                        var newValue = sheet.Cells[i, j].Value;
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "PrimaryCategory".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.PrimaryCategory = newValue.ToString();
                                records.Row = i;
                                if (records.PrimaryCategory.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "PrimaryCategory";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.PrimaryCategory = null;
                                records.Row = i;
                            }

                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "SecondaryCategory".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.SecondaryCategory = newValue.ToString();
                                records.Row = i;
                                if (records.SecondaryCategory.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "SecondaryCategory";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.SecondaryCategory = null;
                                records.Row = i;
                            }

                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Product".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Product = newValue.ToString();
                                records.Row = i;
                                if (records.Product.Replace(" ", "").Count() >= 100)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Product";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Product = null;
                                records.Row = i;
                            }

                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Variant".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Variant = newValue.ToString();
                                records.Row = i;
                                if (records.Variant.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Variant";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Variant = null;
                                records.Row = i;
                            }

                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Price".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Price = newValue.ToString();
                                records.Row = i;
                                if (records.Price.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Price";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Price = null;
                                records.Row = i;
                            }

                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Unit".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Unit = newValue.ToString();
                                records.Row = i;
                                if (records.Unit.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Unit";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Unit = null;
                                records.Row = i;
                            }

                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "DisplayCategory".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.DisplayCategory = newValue.ToString();
                                records.Row = i;
                                if (records.DisplayCategory.Replace(" ", "").Count() >= 30)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "DisplayCategory";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.DisplayCategory = null;
                                records.Row = i;
                            }

                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Image".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Image = newValue.ToString();
                                records.Row = i;
                                if (records.Image.Replace(" ", "").Count() >= 40)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Image";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.Image = null;
                                records.Row = i;
                            }

                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "Description".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.Description = newValue.ToString();
                                records.Row = i;
                                if (records.Description.Replace(" ", "").Count() >= 100)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Description";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }

                                else
                                {
                                    records.Description = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "StandardUnitConversionFactor".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.StandardUnitConversionFactor = newValue.ToString();
                                    records.Row = i;
                                    if (records.StandardUnitConversionFactor.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "StandardUnitConversionFactor";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                else
                                {
                                    records.StandardUnitConversionFactor = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "StandardUnit".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.StandardUnit = newValue.ToString();
                                    records.Row = i;
                                    if (records.StandardUnit.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "StandardUnit";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                else
                                {
                                    records.StandardUnit = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ProductCode".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ProductCode = newValue.ToString();
                                    records.Row = i;
                                    if (records.ProductCode.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ProductCode";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                else
                                {
                                    records.ProductCode = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "VariantCode".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.VariantCode = newValue.ToString();
                                    records.Row = i;
                                    if (records.VariantCode.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "VariantCode";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                else
                                {
                                    records.VariantCode = null;
                                    records.Row = i;
                                }

                            }
                            if (value.ToString().Trim().Replace(" ", "").ToLower() == "ProductCategory".ToLower())
                            {
                                if (newValue != null)
                                {
                                    records.ProductCategory = newValue.ToString();
                                    records.Row = i;
                                    if (records.ProductCategory.Replace(" ", "").Count() >= 50)
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "Character Length Error";
                                        temp.Field_1 = "ProductCategory";
                                        temp.Row = i;
                                        errorTemp.Add(temp);

                                    }
                                }
                                else
                                {
                                    records.ProductCategory = null;
                                    records.Row = i;
                                }

                            }
                        }
                    }
                    
                }
                list.Add(records);
            }
            listerror.error = errorTemp;
            listerror.listProductAddition = list;
            return listerror;
        }
        public ListandError ListErrorBeatPlanAddition (ExcelWorksheet sheet)
        {
            var fileField = sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, sheet.Dimension.End.Row, sheet.Dimension.End.Column];
            ListandError listerror = new ListandError();
            List<FileHeadersBeatPlanAddition> list4 = new List<FileHeadersBeatPlanAddition>();
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            int startrow = sheet.Dimension.Start.Row;
            int endrow = sheet.Dimension.End.Row;
            int startcolomn = sheet.Dimension.Start.Column;
            int endcolomn = sheet.Dimension.End.Column;

            for (int i = startrow + 1; i <= endrow; i++)
            {
                FileHeadersBeatPlanAddition records = new FileHeadersBeatPlanAddition();
                for (int j = startcolomn; j <= endcolomn; j++)
                {
                    var value = sheet.Cells[1, j].Value;
                    if (value != null)
                    {
                        var newValue = sheet.Cells[i, j].Value;
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "ESM".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.ESM = newValue.ToString();
                                records.Row = i;
                                if (records.ESM.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "PrimaryCategory";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.ESM = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "FinalBeatName".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.FinalBeatName = newValue.ToString();
                                records.Row = i;
                                if (records.FinalBeatName.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "SecondaryCategory";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.FinalBeatName = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatPlanStartDate".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.BeatPlanStartDate = newValue.ToString();
                                records.Row = i;
                                if (records.BeatPlanStartDate.Replace(" ", "").Count() >= 100)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "BeatPlanStartDate";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.BeatPlanStartDate = null;
                                records.Row = i;
                            }
                           
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatPeriod".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.BeatPeriod = newValue.ToString();
                                records.Row = i;
                                if (records.BeatPeriod.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "BeatPeriod";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.BeatPeriod = null;
                                records.Row = i;
                            }
                            
                        }
                        if (value.ToString().Trim().Replace(" ", "").ToLower() == "BeatDay".ToLower())
                        {
                            if (newValue != null)
                            {
                                records.BeatDay = newValue.ToString();
                                records.Row = i;
                                if (records.BeatDay.Replace(" ", "").Count() >= 50)
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "Character Length Error";
                                    temp.Field_1 = "Price";
                                    temp.Row = i;
                                    errorTemp.Add(temp);

                                }
                            }
                            else
                            {
                                records.BeatDay = null;
                                records.Row = i;
                            }
                            
                        }

                    }
                    list4.Add(records);
                }

                

            }
            listerror.error = errorTemp;
            listerror.listBeatPlan = list4;
            return listerror;
        }
    }
}