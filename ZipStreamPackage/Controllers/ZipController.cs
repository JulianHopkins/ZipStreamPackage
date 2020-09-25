using ICSharpCode.SharpZipLib.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using ZipStreamPackage.cs;

namespace ZipStreamPackage.Controllers
{
    public class ZipController : Controller
    {
        // GET: Zip
        public ActionResult Index()
        {

            SolidZipModel[] chain = new SolidZipModel[] {
                new SolidZipModel() { FileName = "Photo1.jpg", Url = "http://gdb.rferl.org/C439247B-B9B6-4D99-A512-D9D1951FE4BE_cx0_cy10_cw0_w987_r1_s_r1.jpg"},
                new SolidZipModel() {FileName = "Photo2.jpg", Url = "http://hdwallpapersrocks.com/wp-content/uploads/2013/09/Good-morning-green-land-best-wishes.jpg"},
                new SolidZipModel() {FileName = "Photo3.jpg", Url = "http://hdwallpapersrocks.com/wp-content/uploads/2013/09/Good-morning-flowery-tea-for-you.jpg"}
            };

            return new SolidZipResult(chain, "photoarchive");  

        }
    }
}