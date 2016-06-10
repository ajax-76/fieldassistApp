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
                if (flag_coloumn != 0 && map_coloumn != 0)
                {
                    for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
                    {

                        var map = file.Cells[i, map_coloumn].Value;
                        if (map != null)
                        {
                            map = map.ToString().ToLower();
                        }
                        else
                        {
                            map = "";
                        }

                        var flag = file.Cells[i, flag_coloumn].Value;
                        if (flag != null)
                        {
                            flag = flag.ToString().ToLower();
                        }
                        else
                        {
                            flag = "";
                        }
                        //  int count = 0;
                        for (int j = 2; j <= file.Dimension.End.Row; j++)
                        {
                            if (j != i)
                            {
                                var x = file.Cells[j, map_coloumn].Value;
                                if (x != null)
                                {
                                    x = x.ToString().ToLower();
                                }
                                else
                                {
                                    x = "";
                                }
                                if (x.Equals(map))
                                {
                                    var y = file.Cells[j, flag_coloumn].Value;
                                    if (y != null)
                                    {
                                        y = y.ToString().ToLower();
                                    }
                                    else
                                    {
                                        y = "";
                                    }
                                    if (!y.Equals(flag))
                                    {
                                        ErrorTemplates error = new ErrorTemplates();
                                        error.ErrorType = "One to Many Mapping";
                                        error.Field_1 = flagString;
                                        error.Field_2 = mapString;
                                        error.Row = j;
                                        error.ErrorComments = "More than one " + mapString + " is mapped to " + flagString;
                                        // error.LinkRow = i;
                                        errorTemp.Add(error);
                                    }

                                }
                            }
                        }
                    }
                }
                return errorTemp.OrderBy(x=>x.Row).ToList();

            }
            catch(Exception ex)
            {
                var x = ex.Message;
                
            }
            return null;
        }
        public List<ErrorTemplates>  One2OneValidationCheck(ExcelWorksheet file, int flag_coloumn, int map_coloumn, string flagString, string mapString, List<ErrorTemplates> errorTemp)
        {
            if (flag_coloumn != 0 && map_coloumn != 0)
            {
                for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
                {
                    var map = file.Cells[i, map_coloumn].Value;
                    if (map != null)
                    {
                        map = map.ToString().ToLower();
                    }
                    else
                    {
                        map = "";
                    }
                    var flag = file.Cells[i, flag_coloumn].Value;
                    if (flag != null)
                    {
                        flag = flag.ToString().ToLower();
                    }
                    else
                    {
                        flag = "";
                    }
                    // int count = 0;
                    for (int j = 2; j <= file.Dimension.End.Row; j++)
                    {
                        if (j != i)
                        {
                            var x = file.Cells[j, flag_coloumn].Value;
                            if (x != null)
                            {
                                x = x.ToString().ToLower();
                            }
                            else
                            {
                                x = "";
                            }
                            if (x.Equals(flag))
                            {
                                var y = file.Cells[j, map_coloumn].Value;
                                if (y != null)
                                {
                                    y = y.ToString().ToLower();
                                }
                                else
                                {
                                    y = "";
                                }
                                if (!y.Equals(map))
                                {
                                    ErrorTemplates error = new ErrorTemplates();
                                    error.ErrorType = "One to One Mapping";
                                    error.Field_1 = flagString;
                                    error.Field_2 = mapString;
                                    error.Row = j;
                                    error.ErrorComments = "";
                                    //   error.LinkRow = i;
                                    errorTemp.Add(error);
                                }
                            }
                            else
                            {
                                var z = file.Cells[j, flag_coloumn].Value;
                                if (z != null)
                                {
                                    z = z.ToString().ToLower();
                                }
                                else
                                {
                                    z = "";
                                }
                                if (z.Equals(map))
                                {
                                    ErrorTemplates error = new ErrorTemplates();
                                    error.ErrorType = "One to One";
                                    error.Field_1 = flagString;
                                    error.Field_2 = mapString;
                                    error.Row = j;
                                    //    error.LinkRow = i;
                                    errorTemp.Add(error);
                                }
                            }

                        }
                    }
                }
            }
            return errorTemp.OrderBy(x => x.Row).ToList();
        }
       public List<ErrorTemplates>AttributeMapping(ExcelWorksheet file, int flag_coloumn, int map_coloumn, string flagString, string mapString, List<ErrorTemplates> errorTemp)
        {
            try
            {
                if (flag_coloumn != 0 && map_coloumn != 0)
                {
                    for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
                    {
                        var map = file.Cells[i, map_coloumn].Value;
                        if (map != null)
                        {
                            map = map.ToString().ToLower();
                        }

                        else
                        {
                            map = "";
                        }
                        var flag = file.Cells[i, flag_coloumn].Value;
                        if (flag != null)
                        {
                            flag = flag.ToString().ToLower();
                        }
                        else
                        {
                            flag = "";
                        }
                        for (int j = file.Dimension.Start.Row + 1; j <= file.Dimension.End.Row; j++)
                        {
                            if (j != i)
                            {
                                var x = file.Cells[j, flag_coloumn].Value;
                                if (x != null)
                                {
                                    x = x.ToString().ToLower();
                                }
                                else
                                {
                                    x = "";
                                }

                                if (x.Equals(flag))
                                {
                                    var y = file.Cells[j, map_coloumn].Value;
                                    if (y != null)
                                    {
                                        y = y.ToString().ToLower();
                                    }
                                    else
                                    {
                                        y = "";
                                    }
                                    if (!y.Equals(map))
                                    {
                                        ErrorTemplates error = new ErrorTemplates();
                                        error.ErrorType = "Attribute Mapping";
                                        error.Field_1 = flagString;
                                        error.Field_2 = mapString;
                                        error.Row = j;
                                        error.ErrorComments = "One " + flagString + " contains more than one " + mapString;
                                        // error.LinkRow = i;                                 
                                        errorTemp.Add(error);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                var x = ex.Message;
                var temp = flag_coloumn;
                var temp1 = map_coloumn;
            }
            return errorTemp.OrderBy(x => x.Row).ToList();
        }
        public List<WarningTemplates> HierarchyWarning(ExcelWorksheet file,int ESM_index,int ASM_index,int RSM_index,int ZSM_index,int NSM_index,List<WarningTemplates> errorTemp)
        {
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {
                if(ESM_index==0)
                {
                    break;
                }
                var flag_ESM = file.Cells[i, ESM_index].Value;
                if (flag_ESM == null)
                {
                    var flag_ASM = file.Cells[i, ASM_index].Value;
                    if(ASM_index==0)
                    {
                        break;
                    }
                    if (flag_ASM == null)
                    {
                        if(RSM_index==0)
                        {
                            break;
                        }
                        var flag_RSM = file.Cells[i, RSM_index].Value;
                        if (flag_RSM == null)
                        {
                            if(ZSM_index==0)
                            {
                                break;
                            }
                            var flag_ZSM = file.Cells[i, ZSM_index].Value;
                            if (flag_ZSM == null)
                            {
                                if(NSM_index==0)
                                {
                                    break;
                                }
                                var flag_NSM = file.Cells[i, NSM_index].Value;
                                if (flag_NSM == null)
                                {
                                    WarningTemplates warning = new WarningTemplates();
                                    warning.Comments = "Warning Hierarchy chain is missing";
                                    warning.Field = "ESM";
                                    warning.Row = i;
                                    errorTemp.Add(warning);
                                }
                            }
                        }
                    }
                }

            }
            for (int i=file.Dimension.Start.Row+1;i<=file.Dimension.End.Row;i++)
            {
                if(ASM_index==0)
                {
                    break;
                }
                var flag_ASM = file.Cells[i, ASM_index].Value;
                if(flag_ASM==null)
                {
                    if(RSM_index==0)
                    {
                        break;
                    }
                    var flag_RSM = file.Cells[i, RSM_index].Value;
                    if(flag_RSM==null)
                    {
                        if(ZSM_index==0)
                        {
                            break;
                        }
                        var flag_ZSM = file.Cells[i, ZSM_index].Value;
                        if(flag_ZSM==null)
                        {
                            if(NSM_index==0)
                            {
                                break;
                            }
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
                if(RSM_index==0)
                {
                    break;
                }
                var flag_RSM = file.Cells[i, RSM_index].Value;
                if (flag_RSM == null)
                {
                    if(ZSM_index==0)
                    {
                        break;
                    }
                    var flag_ZSM = file.Cells[i, ZSM_index].Value;
                    if (flag_ZSM == null)
                    {
                        if(NSM_index==0)
                        {
                            break;
                        }
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
                if(ZSM_index==0)
                {
                    break;
                }
               
                    var flag_ZSM = file.Cells[i, ZSM_index].Value;
                    if (flag_ZSM == null)
                    {
                    if(NSM_index==0)
                    {
                        break;
                    }
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

                if(NSM_index==0)
                {
                    break;
                }
                
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
            return errorTemp.OrderByDescending(x => x.Row).ToList();
        }
        public List<ErrorTemplates>HierarchyError(ExcelWorksheet file,int ESM_index, int ASM_index, int RSM_index, int ZSM_index, int NSM_index, List<ErrorTemplates> errorTemp)
        {
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {
                if(ESM_index==0)
                {
                    break;
                }
                var flag_ESM = file.Cells[i, ESM_index].Value;
                if (flag_ESM != null)
                {
                    if(ASM_index==0)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "ASM";
                        errors.Row = i;
                        errors.ErrorComments = "Field ASM is not present";
                        errorTemp.Add(errors);
                    }
                    var flag_ASM = file.Cells[i, ASM_index].Value;
                    if (flag_ASM == null)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "ASM";
                        errors.Row = i;
                        errors.ErrorComments = "hierarchy is broken for field ASM";
                        errorTemp.Add(errors);
                    }
                }

            }
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {
                if(ASM_index==0)
                {
                    break;
                }
                var flag_ASM = file.Cells[i, ASM_index].Value;
                if (flag_ASM != null)
                {
                    if(RSM_index==0)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "RSM";
                        errors.Row = i;
                        errors.ErrorComments = "field RSM is not present";
                        errorTemp.Add(errors);
                    }
                    var flag_RSM = file.Cells[i, RSM_index].Value;
                    if (flag_RSM == null)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "RSM";
                        errors.Row = i;
                        errors.ErrorComments="hierarchy is broken for field RSM";
                        errorTemp.Add(errors);
                    }
                }

            }
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {
                if(RSM_index==0)
                {
                    break;
                }
                var flag_RSM = file.Cells[i, RSM_index].Value;
                if (flag_RSM != null)
                {
                    if(ZSM_index==0)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "ZSM";
                        errors.Row = i;
                        errors.ErrorComments = "field ZSM is not present";
                        //  errors.LinkRow = i;
                        errorTemp.Add(errors);
                    }
                    var flag_ZSM = file.Cells[i, ZSM_index].Value;
                    if (flag_ZSM == null)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "ZSM";
                        errors.Row = i;
                        errors.ErrorComments = "hierarchy is broken for field ZSM";
                      //  errors.LinkRow = i;
                        errorTemp.Add(errors);
                    }
                }

            }
            for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {
                if(ZSM_index==0)
                {
                    break;
                }
                var flag_ZSM = file.Cells[i, ZSM_index].Value;
                if (flag_ZSM != null)
                {
                    if(NSM_index==0)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "NSM";
                        errors.Row = i;
                        errors.ErrorComments = "field is nog present for NSM";
                        //  errors.LinkRow = i;
                        errorTemp.Add(errors);
                    }
                    var flag_NSM = file.Cells[i, NSM_index].Value;
                    if (flag_NSM == null)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "NSM";
                        errors.Row = i;
                        errors.ErrorComments = "hierarchy is broken for field NSM";
                      //  errors.LinkRow = i;
                        errorTemp.Add(errors);
                    }
                }

            }
            return errorTemp.OrderBy(x => x.Row).ToList();
        }
        public List<ErrorTemplates> CheckPhoneDigit(ExcelWorksheet file,int ESM_flag,List<ErrorTemplates>errorTemp)
        {
            int phone = 0;
            if (ESM_flag != 0)
            {
                for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
                {
                    phone = i;
                    try
                    {
                        if (file.Cells[i, ESM_flag].Value != null)
                        {
                            string number = file.Cells[i, ESM_flag].Value.ToString().Trim();
                            int count = number.Length;
                            if (count != 10)
                            {
                                ErrorTemplates errors = new ErrorTemplates();
                                errors.ErrorType = "Wrong Phone Number ";
                                errors.Field_1 = "ESM Contact Number";
                                errors.Row = i;
                                errors.ErrorComments = "phone number should of 10 digit";
                                //     errors.LinkRow = i;
                                errorTemp.Add(errors);
                            }
                        }
                    }
                    catch
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Phone Number is incorrect";
                        errors.Field_1 = "ESM Contact Number";
                        errors.Row = phone;
                        errors.ErrorComments = "phone number should of 10 digit";
                        // errors.LinkRow = i;
                        errorTemp.Add(errors);
                    }

                }
            }
            return errorTemp.OrderBy(x => x.Row).ToList();

        }
        public List<ErrorTemplates> EmailCheck(ExcelWorksheet file ,int columnIndex,string field,List<ErrorTemplates> errorTemp)
        {
            if (columnIndex != 0)
            {
                for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
                {
                    var value = file.Cells[i, columnIndex].Value;
                    string newValue = "";
                    if (value != null)
                    {
                        newValue = value.ToString();
                    }

                    RegexUtilities util = new RegexUtilities();
                    if (value != null)
                    {
                        if (!util.IsValidEmail(newValue))
                        {
                            ErrorTemplates error = new ErrorTemplates();
                            error.ErrorType = "Email Format";
                            error.ErrorComments = "Email Format is incorrect";
                            error.Field_1 = field;
                            error.Row = i;
                            errorTemp.Add(error);
                        }
                    }

                }
            }
            return errorTemp.OrderBy(x=>x.Row).ToList();
        }
    }
}