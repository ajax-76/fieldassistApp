using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using FieldAssist.DataAccessLayer.Models.EFModels;

namespace DataUpload.DataUploadRepository
{
    
    public class NSMRepository
    {
        CompanyDataContext db = new CompanyDataContext();
        public List<NationalSalesManager> GetNSM(int? id)
        {
            return db.NationalSalesManagers.Where(x => x.CompanyId == id).ToList();
        }
        public void  AddNSM(NationalSalesManager NSM)
        {
            db.NationalSalesManagers.Add(NSM);
            db.SaveChanges();
        }
    }
}