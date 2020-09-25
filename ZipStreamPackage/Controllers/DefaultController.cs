using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace ZipStreamPackage.Controllers
{
    public class DefaultController : Controller
    {
        // GET: Default
        public void Index()
        {

            string[] Photos = new[] { "https://www.google.ru/url?sa=i&rct=j&q=&esrc=s&source=images&cd=&cad=rja&uact=8&ved=0ahUKEwjhp5roxYXPAhXhB5oKHcK7BckQjRwIBw&url=http%3A%2F%2Fpozdravka.com%2Fphoto%2Fpozdravlenija%2Fs_dnem_rozhdenija%2F24&psig=AFQjCNEwBWFRfqob7zZUSjWhh53EGUgbFQ&ust=1473622515400060", "https://www.google.ru/url?sa=i&rct=j&q=&esrc=s&source=images&cd=&cad=rja&uact=8&ved=0ahUKEwia0sWBxoXPAhWEa5oKHS0PCcUQjRwIBw&url=http%3A%2F%2Fpozdravik.com%2Fotkritki-5&psig=AFQjCNEwBWFRfqob7zZUSjWhh53EGUgbFQ&ust=1473622515400060", "https://www.google.ru/url?sa=i&rct=j&q=&esrc=s&source=images&cd=&cad=rja&uact=8&ved=0ahUKEwi0idyJxoXPAhWlZpoKHUVgDcUQjRwIBw&url=http%3A%2F%2Fphotoshablon.ru%2Fphoto%2Fkartinki_pozdravlenija_s_dnem_rozhdenija%2F3-2&psig=AFQjCNEwBWFRfqob7zZUSjWhh53EGUgbFQ&ust=1473622515400060" };
            //Stream originalStream = wc.OpenRead(si)
            WebClient wc = new WebClient();

                HttpContext.Response.Clear();
                HttpContext.Response.ContentType = "application/octet-stream";
                HttpContext.Response.AppendHeader("Connection", "keep-alive");
                HttpContext.Response.AppendHeader("Content-Disposition", "attachment;filename=" + "test.zip");
            int i = 0;
                using (ZipArchive archive = new ZipArchive(Response.OutputStream,  ZipArchiveMode.Create))
                {

                foreach (var si in Photos)
                {
                    ZipArchiveEntry readmeEntry = archive.CreateEntry(i.ToString());
                    i++;
                    CopyStream(wc.OpenRead(si), readmeEntry.Open());

                }

                }
                HttpContext.Response.Flush();
                HttpContext.Response.End();
               // return File(compressedStream, "application/octet-stream", string.Format("Archive_{0}", ans.ShootingId));
              }


        private static void CopyStream(Stream source, Stream target)
        {
            const int bufSize = 0x1000;
            byte[] buf = new byte[bufSize];
            int bytesRead = 0;
            while ((bytesRead = source.Read(buf, 0, bufSize)) > 0)
                target.Write(buf, 0, bytesRead);
        }
    }

}