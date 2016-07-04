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
        public List<ErrorTemplates> One2ManyValidationCheck(List<IGrouping<string,Mapping>> query1, int flag_coloumn, int map_coloumn, string flagString, string mapString)
        {
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            int count1 = 0;
            int count2 = 0;
            try
            {
                // var flagCell = file.Cells[start_row, start_coloumn];
                if (flag_coloumn != 0 && map_coloumn != 0)
                {

                    List<List<Mapping>> mappin = new List<List<Mapping>>();
                    foreach (var key in query1)
                    {
                        List<Mapping> tempMap = new List<Mapping>();
                        foreach (var item in key)
                        {
                            Mapping map = new Mapping();
                            map.c1 = item.c1;
                            map.row = item.row;
                            tempMap.Add(map);
                        }
                        mappin.Add(tempMap);
                    }
                    foreach (var item in mappin)
                    {
                        var comparer = item;
                        var list1 = item.Select(x => x.c1).ToList();

                        foreach (var item2 in mappin)
                        {

                            var list2 = item2.Select(x => x.c1).ToList();
                            if (item != item2)
                            {
                                var differenceMap = list1.Intersect(list2).ToList();
                                if (differenceMap.Count() != 0)
                                {
                                    List<int> q = new List<int>();
                                    List<int> r = new List<int>();
                                    foreach (var key in differenceMap)
                                    {
                                        q = item.Where(x => x.c1 == key).Select(x => x.row).ToList();
                                        r = item2.Where(x => x.c1 == key).Select(x => x.row).ToList();
                                    }


                                    foreach (var num in q.Distinct())
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "One to Many mapping";
                                        temp.Field_1 = flagString;
                                        temp.Field_2 = mapString;
                                        temp.Row = num;
                                        errorTemp.Add(temp);
                                        
                                    }

                                    foreach (var num in r.Distinct())
                                    {
                                        ErrorTemplates temp = new ErrorTemplates();
                                        temp.ErrorType = "One to Many mapping";
                                        temp.Field_1 = flagString;
                                        temp.Field_2 = mapString;
                                        temp.Row = num;
                                        errorTemp.Add(temp);
                                        
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
                var c = count1;
                var c2 = count2;
            }
            return null;
        }
        public List<ErrorTemplates>  One2OneValidationCheck(List<IGrouping<string,Mapping>> query1, int flag_coloumn, int map_coloumn, string flagString, string mapString)
        {
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            if (flag_coloumn != 0 && map_coloumn != 0)
            {

                List<List<Mapping>> mappin = new List<List<Mapping>>();
                foreach (var key in query1)
                {
                    List<Mapping> tempMap = new List<Mapping>();
                    foreach (var item in key)
                    {
                        Mapping map = new Mapping();
                        map.c1 = item.c1;
                        map.row = item.row;
                        tempMap.Add(map);
                    }
                    mappin.Add(tempMap);
                }
                foreach (var item in mappin)
                {
                    var comparer = item;
                    var list1 = item.Select(x => x.c1).ToList();
                    var key1 = list1.Distinct().Count();
                    if(key1!=1)
                    {
                        foreach (var num in item)
                        {
                            ErrorTemplates temp = new ErrorTemplates();
                            temp.ErrorType = "One to One mapping";
                            temp.Field_1 = flagString;
                            temp.Field_2 = mapString;
                            temp.Row = num.row;
                            errorTemp.Add(temp);

                        }
                    }
                    foreach (var item2 in mappin)
                    {

                        var list2 = item2.Select(x => x.c1).ToList();
                        if (item != item2)
                        {
                            var differenceMap = list1.Intersect(list2).ToList();
                            if (differenceMap.Count() != 0)
                            {
                                List<int> q = new List<int>();
                                List<int> r = new List<int>();
                                foreach (var key in differenceMap)
                                {
                                    q = item.Where(x => x.c1 == key).Select(x => x.row).ToList();
                                    r = item2.Where(x => x.c1 == key).Select(x => x.row).ToList();
                                }


                                foreach (var num in q.Distinct())
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "One to One mapping";
                                    temp.Field_1 = flagString;
                                    temp.Field_2 = mapString;
                                    temp.Row = num;
                                    errorTemp.Add(temp);

                                }

                                foreach (var num in r.Distinct())
                                {
                                    ErrorTemplates temp = new ErrorTemplates();
                                    temp.ErrorType = "One to One mapping";
                                    temp.Field_1 = flagString;
                                    temp.Field_2 = mapString;
                                    temp.Row = num;
                                    errorTemp.Add(temp);

                                }
                            }
                        }
                    }
                }
            }
            return errorTemp.OrderBy(x => x.Row).ToList();
        }
       public List<ErrorTemplates>AttributeMapping(List<IGrouping<string, Mapping>> query1, int flag_coloumn, int map_coloumn, string flagString, string mapString)
        {
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            try
            {

                if (flag_coloumn != 0 && map_coloumn != 0)
                {

                    List<List<Mapping>> mappin = new List<List<Mapping>>();
                    foreach (var key in query1)
                    {
                        List<Mapping> tempMap = new List<Mapping>();
                        foreach (var item in key)
                        {
                            Mapping map = new Mapping();
                            map.c1 = item.c1;
                            map.row = item.row;
                            tempMap.Add(map);
                        }
                        mappin.Add(tempMap);
                    }
                    foreach (var item in mappin)
                    {
                        var comparer = item;
                        var list1 = item.Select(x => x.c1).ToList();
                        var key1 = list1.Distinct().Count();
                        if (key1 != 1)
                        {
                            foreach (var num in item)
                            {
                                ErrorTemplates temp = new ErrorTemplates();
                                temp.ErrorType = "Attribute Mapping";
                                temp.Field_1 = flagString;
                                temp.Field_2 = mapString;
                                temp.Row = num.row;
                                errorTemp.Add(temp);

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
        public List<WarningTemplates> HierarchyWarning(ExcelWorksheet file,int ESM_index,int ASM_index,int RSM_index,int ZSM_index,int NSM_index)
        {
            List<WarningTemplates> errorTemp = new List<WarningTemplates>();
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
        public List<ErrorTemplates>HierarchyError(ExcelWorksheet file,int ESM_index, int ASM_index, int RSM_index, int ZSM_index, int NSM_index)
        {
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
                for (int i = file.Dimension.Start.Row + 1; i <= file.Dimension.End.Row; i++)
            {
                if(ESM_index==0)
                {
                    break;
                }
                var flag_ESM = file.Cells[i, ESM_index].Value;
                if (flag_ESM != null)
                {
                    if (ASM_index == 0)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "ASM";
                        errors.ErrorComments = "Header ASM is missing";
                        errorTemp.Add(errors);
                        break;
                    }
                    else
                    {
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
                    if (RSM_index == 0)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "RSM";
                        errors.ErrorComments = "Header RSM is missing";
                        errorTemp.Add(errors);
                        break;
                    }
                    else
                    {
                        var flag_RSM = file.Cells[i, RSM_index].Value;
                        if (flag_RSM == null)
                        {
                            ErrorTemplates errors = new ErrorTemplates();
                            errors.ErrorType = "Error Hierachy is broken";
                            errors.Field_1 = "RSM";
                            errors.Row = i;
                            errors.ErrorComments = "hierarchy is broken for field RSM";
                            errorTemp.Add(errors);
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
                if (flag_RSM != null)
                {
                    if (ZSM_index == 0)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "ZSM";
                        errors.ErrorComments = "Header ZSM is missing";
                        //  errors.LinkRow = i;
                        errorTemp.Add(errors);
                        break;
                    }
                    else
                    {
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
                    if (NSM_index == 0)
                    {
                        ErrorTemplates errors = new ErrorTemplates();
                        errors.ErrorType = "Error Hierachy is broken";
                        errors.Field_1 = "NSM";
                        errors.ErrorComments = "Header NSM is missing";
                        errorTemp.Add(errors);
                        break;
                    }
                    else
                    {
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

            }
            return errorTemp.OrderBy(x => x.Row).ToList();
        }
        public List<ErrorTemplates> CheckPhoneDigit(ExcelWorksheet file,int ESM_flag)
        {
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();

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
                            if (number.Any(x=>char.IsLetter(x)))
                            {
                                ErrorTemplates errors = new ErrorTemplates();
                                errors.ErrorType = "Phone Number is incorrect";
                                errors.Field_1 = "ESM Contact Number";
                                errors.Row = phone;
                                errors.ErrorComments = "phone number should only contain Numeric values";
                                // errors.LinkRow = i;
                                errorTemp.Add(errors);
                            }
                            
                            else if (count != 10)
                            {
                                ErrorTemplates errors = new ErrorTemplates();
                                errors.ErrorType = "Phone Number is incorrect ";
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
        public List<ErrorTemplates> EmailCheck(ExcelWorksheet file ,int columnIndex,string field)
        {
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
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
        public List<ErrorTemplates> CheckStateAndDistrict(ExcelWorksheet file,int flag_State,int flag_District)
        {
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();


            StateDistrict beatDistrict = new StateDistrict();
            
            for(int i=file.Dimension.Start.Row+1;i<=file.Dimension.End.Row;i++)
            {
                if (flag_State != 0 && flag_District!=0)
                {
                    var statevalue = file.Cells[i, flag_State].Value;
                    var districtValue = file.Cells[i, flag_District].Value;
                    if (statevalue != null)
                    {
                        var newValue = statevalue.ToString().Replace(" ", "").ToLower();
                        var listState = beatDistrict.GetAllStates();
                        var boole = listState.Contains(newValue);
                        if (boole == false)
                        {
                            ErrorTemplates error = new ErrorTemplates();
                            error.ErrorType = "State";
                            error.ErrorComments = "State field is out of the list";
                            error.Field_1 = newValue;
                            error.Row = i;
                            errorTemp.Add(error);
                        }
                        else
                        {
                            if(districtValue!=null)
                            {
                                var newValue2 = districtValue.ToString().Replace(" ", "").ToLower();
                                var listDistrict = beatDistrict.GetDistrictsOfState(newValue);
                                var boole2 = listDistrict.Contains(newValue2);
                                if(boole2==false)
                                {
                                    ErrorTemplates error = new ErrorTemplates();
                                    error.ErrorType = "District";
                                    error.ErrorComments = "District field is out of the list for State  "+newValue;
                                    error.Field_1 = newValue2;
                                    error.Row = i;
                                    errorTemp.Add(error);
                                }
                            }
                        }
                    }
                }               
            } 
            return errorTemp.OrderBy(x => x.Row).ToList();
        }
        public List<ErrorTemplates> UniqueProductVariant(List<IGrouping<string, Mapping>> query, int product_index,int variant_index)
        {
            List<ErrorTemplates> errorTemp = new List<ErrorTemplates>();
            List<string> List = new List<string>();
            List<Mapping> maper = new List<Mapping>();
            if (product_index != 0 && variant_index != 0)
            {
                List<List<Mapping>> mappin = new List<List<Mapping>>();
                foreach (var key in query)
                {
                    List<Mapping> tempMap = new List<Mapping>();
                    foreach (var item in key)
                    {
                        Mapping map = new Mapping();
                        map.c1 = item.c1;
                        map.row = item.row;
                        tempMap.Add(map);
                    }
                    if(tempMap.Count>1)
                    {
                        foreach(var item2 in tempMap)
                        {
                            ErrorTemplates error = new ErrorTemplates();
                            error.ErrorType = "UniqueProductVariantError";
                            error.Field_1 = item2.row.ToString();
                            errorTemp.Add(error);
                        }
                    }
                }
                
                }

            
            return errorTemp.OrderBy(x => x.Row).ToList();
        }
        
        
    }
}