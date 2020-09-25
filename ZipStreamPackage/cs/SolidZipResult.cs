using ICSharpCode.SharpZipLib.Zip;
using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace ZipStreamPackage.cs
{
    public class SolidZipResult : ActionResult
    {

        public SolidZipResult(SolidZipModel[] chain, string filename)
        {
            this.chain = chain;
            this.filename = filename;
        }

        private SolidZipModel[] chain;
        private string filename { get; set; }
 

        public override void ExecuteResult(ControllerContext context)
        {
            int level = 0;
            context.HttpContext.Response.Clear();
            context.HttpContext.Response.ContentType = "application/octet-stream";
            context.HttpContext.Response.AppendHeader("Connection", "keep-alive");
            context.HttpContext.Response.AppendHeader("Content-Disposition", string.Format("attachment;filename={0}.zip", filename));
            using (ZipOutputStream s = new ZipOutputStream(context.HttpContext.Response.OutputStream))
            {
                s.UseZip64 = UseZip64.Off;
                if (level != -1)
                    s.SetLevel(level);

                byte[] buffer = new byte[4096];

                foreach(var link in chain)
                {
                    ZipEntry entry = new ZipEntry(link.FileName);
                    s.PutNextEntry(entry);
                    using (Stream fs = new WebClient().OpenRead(link.Url))
                    {
                        int sourceBytes;
                        do
                        {
                            sourceBytes = fs.Read(buffer, 0, buffer.Length);
                            s.Write(buffer, 0, sourceBytes);
                        }
                        while (sourceBytes > 0);
                    }
                }

                s.Finish();
                s.Close();
            }
            context.HttpContext.Response.Flush();
            context.HttpContext.Response.End();

        }
    }
}