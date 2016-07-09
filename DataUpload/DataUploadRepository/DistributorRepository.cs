using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using FieldAssist.DataAccessLayer.Models.EFModels;

namespace DataUpload.DataUploadRepository
{
    public class DistributorRepository
    {
        CompanyDataContext db = new CompanyDataContext();
        public List<Distributor> GetDistributor(int? id)
        {
            return db.Distributors.Where(x => x.CompanyId == id).ToList();
        }
        public void AddDistributor(Distributor Distributor)
        {
            db.Distributors.Add(Distributor);
            db.SaveChanges();
        }
    }
}