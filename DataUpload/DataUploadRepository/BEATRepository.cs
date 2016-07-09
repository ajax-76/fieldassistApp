using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using FieldAssist.DataAccessLayer.Models.EFModels;
namespace DataUpload.DataUploadRepository
{
    public class BEATRepository
    {
        CompanyDataContext db = new CompanyDataContext();
        public List<LocationBeat> GetBEAT(int? id)
        {
            return db.LocationBeats.Where(x => x.Company == id).ToList();
        }
        public void AddBEAT(LocationBeat BEAT)
        {
            db.LocationBeats.Add(BEAT);
            db.SaveChanges();
        }
    }
}