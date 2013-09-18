using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExportDemo1;
using ExporterObjects;

namespace ExportDemo2.Controllers
{
    public partial class HomeController : Controller
    {
        public virtual ActionResult Index()
        {
            ViewBag.Message = "Export to various formats";

            return View();
        }

        public virtual ActionResult About()
        {
            return View();
        }
        public virtual ActionResult ExportTo(int id)
        {
            List<Electronics> list = Electronics.GetData();
            ExportList<Electronics> exp = new ExportList<Electronics>();
            exp.PathTemplateFolder = Server.MapPath("~/ExportTemplates/electronics");

            string filePathExport = Server.MapPath("~/exports/a" + ExportBase.GetFileExtension((ExportToFormat)id));
            exp.ExportTo(list, (ExportToFormat)id, filePathExport);
            return this.File(filePathExport, "application/octet-stream", System.IO.Path.GetFileName(filePathExport));

        }
    }
}
