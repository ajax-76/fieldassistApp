using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using FieldAssist.DataAccessLayer.Models.EFModels;
namespace DataUpload.DataUploadRepository
{
    public class RSMRepository
    {
        CompanyDataContext db = new CompanyDataContext();
        public List<RegionalSalesManager> GetRSM(int? id)
        {
            return db.RegionalSalesManagers.Where(x => x.CompanyId == id).ToList();
        }
        public void AddRSM(RegionalSalesManager RSM)
        {
            db.RegionalSalesManagers.Add(RSM);
            db.SaveChanges();
        }
    }
}