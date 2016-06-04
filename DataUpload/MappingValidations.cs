using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using OfficeOpenXml;
using DataUpload.Controllers;

namespace DataUpload
{
    public class MappingValidations
    {
        public List<ErrorTemplates> One2ManyValidationCheck(ExcelWorksheet file, int flag_coloumn, int map_coloumn, string flagString, string mapString,List<ErrorTemplates>errorTemp)
        {
            try
            {
                // var flagCell = file.Cells[start_row, start_coloumn];
                for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
                {
                    var flag = file.Cells[i, flag_coloumn].Value;
                    var map = file.Cells[i, map_coloumn].Value;
                    //  int count = 0;
                    for (int j = 2; j <= file.Dimension.End.Row; j++)
                    {
                        if (j != i)
                        {
                            var x = file.Cells[j, map_coloumn].Value;      
                            if (x == map)
                            {
                                var y = file.Cells[j, flag_coloumn];
                                if (y.Value != flag)
                                {
                                    ErrorTemplates error = new ErrorTemplates();
                                    error.MappingErrorType = "One to Many";
                                    error.Field_1 = flagString;
                                    error.Field_2 = mapString;
                                    error.Row = j;
                                    error.LinkRow = i;
                                    errorTemp.Add(error);
                                }

                            }
                        }
                    }
                }
                return errorTemp;

            }
            catch(Exception ex)
            {
                var x = ex.Message;
                
            }
            return null;
        }
        public List<ErrorTemplates>  One2OneValidationCheck(ExcelWorksheet file, int flag_coloumn, int map_coloumn, string flagString, string mapString, List<ErrorTemplates> errorTemp)
        {
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {
                var flag = file.Cells[i, flag_coloumn].Value;
                var map = file.Cells[i, map_coloumn].Value;
                // int count = 0;
                for (int j = 2; j <= file.Dimension.End.Row; j++)
                {
                    if (j != i)
                    {
                        if (file.Cells[j, flag_coloumn].Value == flag)
                        {
                            if (file.Cells[j, map_coloumn].Value != map)
                            {
                                ErrorTemplates error = new ErrorTemplates();
                                error.MappingErrorType = "One to One";
                                error.Field_1 = flagString;
                                error.Field_2 = mapString;
                                error.Row = j;
                                error.LinkRow = i;
                                errorTemp.Add(error);
                            }
                        }
                        else
                        {
                            if(file.Cells[j,flag_coloumn].Value==map)
                            {
                                ErrorTemplates error = new ErrorTemplates();
                                error.MappingErrorType = "One to One";
                                error.Field_1 = flagString;
                                error.Field_2 = mapString;
                                error.Row = j;
                                error.LinkRow = i;
                                errorTemp.Add(error);
                            }
                        }

                    }
                }
            }
            return errorTemp;
        }
       public List<ErrorTemplates>AttributeMapping(ExcelWorksheet file, int flag_coloumn, int map_coloumn, string flagString, string mapString, List<ErrorTemplates> errorTemp)
        {
            try
            {
                for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
                {
                    var map = file.Cells[i, map_coloumn].Value;
                    var flag = file.Cells[i, flag_coloumn].Value;
                    for (int j = file.Dimension.Start.Row + 1; j <= file.Dimension.End.Row; j++)
                    {
                        if (j != i)
                        {
                            if (file.Cells[j, flag_coloumn].Value == flag)
                            {
                                if (file.Cells[j, map_coloumn].Value != map)
                                {
                                    ErrorTemplates error = new ErrorTemplates();
                                    error.MappingErrorType = "Attribute Mapping";
                                    error.Field_1 = flagString;
                                    error.Field_2 = mapString;
                                    error.Row = j;
                                    error.LinkRow = i;
                                  
                                    errorTemp.Add(error);
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                var x = ex.Message;
            }
            return errorTemp;
        }
        public List<WarningTemplates> HierarchyWarning(ExcelWorksheet file,int ASM_index,int RSM_index,int ZSM_index,int NSM_index,List<WarningTemplates> errorTemp)
        {
            for(int i=file.Dimension.Start.Row+1;i<=file.Dimension.End.Row;i++)
            {
                var flag_ASM = file.Cells[i, ASM_index].Value;
                if(flag_ASM==null)
                {
                    var flag_RSM = file.Cells[i, RSM_index].Value;
                    if(flag_RSM==null)
                    {
                        var flag_ZSM = file.Cells[i, ZSM_index].Value;
                        if(flag_ZSM==null)
                        {
                            var flag_NSM = file.Cells[i, NSM_index].Value;
                            if(flag_NSM==null)
                            {
                                WarningTemplates warning = new WarningTemplates();
                                warning.Comments = "Warning Hierarchy chain is missing";
                                warning.Field = "ASM";              
                                warning.Row = i;
                                errorTemp.Add(warning);
                            }
                        }
                    }
                }

            }
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {

                var flag_RSM = file.Cells[i, RSM_index].Value;
                if (flag_RSM == null)
                {
                    var flag_ZSM = file.Cells[i, ZSM_index].Value;
                    if (flag_ZSM == null)
                    {
                        var flag_NSM = file.Cells[i, NSM_index].Value;
                        if (flag_NSM == null)
                        {
                            WarningTemplates warning = new WarningTemplates();
                            warning.Comments = "Warning Hierarachy Chain is missing";
                            warning.Field = "RSM";
                            warning.Row = i;
                            errorTemp.Add(warning);
                        }
                    }
                }
            }
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {

               
                    var flag_ZSM = file.Cells[i, ZSM_index].Value;
                    if (flag_ZSM == null)
                    {
                        var flag_NSM = file.Cells[i, NSM_index].Value;
                        if (flag_NSM == null)
                        {
                        WarningTemplates warning = new WarningTemplates();
                        warning.Comments = "Warning hierarchy chain is missing";
                        warning.Field = "ZSM";
                        warning.Row = i;
                        errorTemp.Add(warning);
                    }
                    }
                
            }
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {


                
                    var flag_NSM = file.Cells[i, NSM_index].Value;
                    if (flag_NSM == null)
                    {
                    WarningTemplates warning = new WarningTemplates();
                    warning.Comments = "Warning Hierarchy chain is missing";
                    warning.Field = "NSM";
                    warning.Row = i;
                    errorTemp.Add(warning);
                }
                

            }
            return errorTemp;
        }
        public List<ErrorTemplates>HierarchyError(ExcelWorksheet file, int ASM_index, int RSM_index, int ZSM_index, int NSM_index, List<ErrorTemplates> errorTemp)
        {
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {
                var flag_ASM = file.Cells[i, ASM_index].Value;
                if (flag_ASM == null)
                {
                    var flag_RSM = file.Cells[i, RSM_index].Value;
                    if (flag_RSM != null)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.HierarchyBeak = "Error Hierachy is broken";
                        errors.Field_1 = "RSM";
                        errors.Row = i;
                        errors.LinkRow = i;
                        errorTemp.Add(errors);
                    }
                }

            }
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {
                var flag_RSM = file.Cells[i, RSM_index].Value;
                if (flag_RSM == null)
                {
                    var flag_ZSM = file.Cells[i, ZSM_index].Value;
                    if (flag_ZSM != null)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.HierarchyBeak = "Error Hierachy is broken";
                        errors.Field_1 = "ZSM";
                        errors.Row = i;
                        errors.LinkRow = i;
                        errorTemp.Add(errors);
                    }
                }

            }
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {
                var flag_ZSM = file.Cells[i, ZSM_index].Value;
                if (flag_ZSM == null)
                {
                    var flag_ASM = file.Cells[i, ASM_index].Value;
                    if (flag_ASM != null)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.HierarchyBeak = "Error Hierachy is broken";
                        errors.Field_1 = "NSM";
                        errors.Row = i;
                        errors.LinkRow = i;
                        errorTemp.Add(errors);
                    }
                }

            }
            return errorTemp;
        }
        public List<ErrorTemplates> CheckPhoneDigit(ExcelWorksheet file,int ESM_flag,List<ErrorTemplates>errorTemp)
        {
            int phone = 0;
            for(int i=file.Dimension.Start.Row+1;i<=file.Dimension.End.Row;i++)
            {
                phone = i;
                try
                {
                    if(file.Cells[i,ESM_flag].Value!=null)
                    {
                        string number = file.Cells[i, ESM_flag].Value.ToString().Trim();
                        int count = number.Length;
                        if(count!=10)
                        {
                            ErrorTemplates errors = new ErrorTemplates();
                            errors.PhoneError = "Phone Number Should be of 10 Digit";
                            errors.Field_1 = "ESM Contact Number";
                            errors.Row = i;
                            errors.LinkRow = i;
                            errorTemp.Add(errors);
                        }
                    }
                }
                catch
                {
                    ErrorTemplates errors = new ErrorTemplates();
                    errors.PhoneError = "Phone Number is incorrect";
                    errors.Field_1 = "ESM Contact Number";
                    errors.Row = phone;
                    errors.LinkRow = i;
                    errorTemp.Add(errors);
                }
                
            }
            return errorTemp;

        }
    }
}