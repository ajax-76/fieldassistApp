using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using FieldAssist.DataAccessLayer.Models.EFModels;
namespace DataUpload.DataUploadRepository
{
    public class ESMRepository
    {
        CompanyDataContext db = new CompanyDataContext();
        public List<ClientEmployee> GetESM(int? id)
        {
            return db.ClientEmployees.Where(x => x.Company == id).ToList();
        }
        public void AddESM(ClientEmployee ESM)
        {
            db.ClientEmployees.Add(ESM);
            db.SaveChanges();
        }
    }
}