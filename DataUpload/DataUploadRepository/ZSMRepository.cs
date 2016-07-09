using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using FieldAssist.DataAccessLayer.Models.EFModels;


namespace DataUpload.DataUploadRepository
{
    public class ZSMRepository
    {
        CompanyDataContext db = new CompanyDataContext();
        public List<ZonalSalesManager> GetZSM(int? id)
        {
            return db.ZonalSalesManagers.Where(x => x.CompanyId == id).ToList();
        }
        public void AddZSM(ZonalSalesManager ZSM)
        {
            db.ZonalSalesManagers.Add(ZSM);
        }
    }
}