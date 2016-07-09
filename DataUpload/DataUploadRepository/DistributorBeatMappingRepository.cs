using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using FieldAssist.DataAccessLayer.Models.EFModels;

namespace DataUpload.DataUploadRepository
{
    public class DistributorBeatMappingRepository
    {
        CompanyDataContext db = new CompanyDataContext();
        public List<DistributorBeatMapping> GetDistributorBeatMap(int ? id)
        {
            return db.DistributorBeatMappings.Where(x => x.CompanyId == id).ToList();
        }
        public void AddDistributorBeatMap(DistributorBeatMapping DBM)
        {
            db.DistributorBeatMappings.Add(DBM);
            db.SaveChanges();
        }
    }
}