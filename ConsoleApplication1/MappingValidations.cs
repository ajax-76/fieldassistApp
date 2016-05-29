using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
    class MappingValidations
    {
        public void One2ManyValidationCheck(ExcelWorksheet file, int flag_coloumn, int map_coloumn)
        {
            // var flagCell = file.Cells[start_row, start_coloumn];
            for (int i = file.Dimension.Start.Row; i <= file.Dimension.End.Row; i++)
            {
                var flag = file.Cells[i, flag_coloumn];
                var map = file.Cells[i, map_coloumn];
                //  int count = 0;
                for (int j = 1; j <= file.Dimension.End.Row; j++)
                {
                    if (j != i)
                    {
                        if (file.Cells[j, map_coloumn] == map)
                        {
                            if (file.Cells[j, flag_coloumn] != flag)
                            {
                                Console.WriteLine("many to many map is between at row: {0} coloumn:{1} and row: {2} coloumn: {3}", j, flag_coloumn, j, map_coloumn);
                            }
                        }
                    }
                }
            }
        }
        public void One2OneValidationCheck(ExcelWorksheet file, int flag_coloumn, int map_coloumn)
        {
            for (int i = file.Dimension.Start.Row; i <= file.Dimension.End.Row; i++)
            {
                var flag = file.Cells[i, flag_coloumn];
                var map = file.Cells[i, map_coloumn];
                // int count = 0;
                for (int j = 1; j <= file.Dimension.End.Row; j++)
                {
                    if (j != i)
                    {
                        if (file.Cells[j, flag_coloumn] == flag)
                        {
                            if (file.Cells[j, map_coloumn] != map)
                            {
                                Console.WriteLine("one to one mapping is incorrect between row: {0} coloumn: {1} and row:{2} coloumn: {3}", j, flag_coloumn, j, map_coloumn);
                            }
                        }
                    }
                }
            }
        }
    }
}
