using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace DataUpload.Controllers
{
    public class UploadController : Controller
    {
        // GET: Upload
        public ActionResult Upload()
        {
            return View();
        }
        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            string path = null;


            try
            {

                if (file.ContentLength > 0)
                {

                    string filename = Path.GetFileName(file.FileName);

                    path = Path.Combine(Server.MapPath("~/Seed Data/"), filename);
                    file.SaveAs(path);


                    string Params = "\"" + path + "\"";
                    
                    Process newProcess = new Process();
                    newProcess.StartInfo.UseShellExecute = false;
                    newProcess = Process.Start(@"C:\Users\Lenovo\Documents\Visual Studio 2015\Projects\DataUpload\ConsoleApplication1\bin\Debug\ConsoleApplication1.exe",Params);
                    newProcess.Close();



                }
            }
            catch
            {
                View("Error");
            }


            return null;

        }
    }
    }
