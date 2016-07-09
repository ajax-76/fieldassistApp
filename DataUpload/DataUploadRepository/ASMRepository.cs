using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using FieldAssist.DataAccessLayer.Models.EFModels;
namespace DataUpload.DataUploadRepository
{
    public class ASMRepository
    {
        CompanyDataContext db = new CompanyDataContext();
        public List<AreaSalesManager> GetASM(int? id)
        {
            return db.AreaSalesManagers.Where(x => x.CompanyId == id).ToList();
        }
        public void AddASM(AreaSalesManager ASM)
        {
            db.AreaSalesManagers.Add(ASM);
            db.SaveChanges();
        }
    }
}