using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using FieldAssist.DataAccessLayer.Models.EFModels;
namespace DataUpload.DataUploadRepository
{
    
    public class LocationAdditionRepository
    {
        CompanyDataContext db = new CompanyDataContext();

        public List<Location> GetLocation(int id)
        {
            return db.Locations.Where(x => x.CompanyId == id).ToList();
        }


    }
}