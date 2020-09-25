using NPOI.HSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ZipStreamPackage.cs
{
    public class ExcelResult : ActionResult
    {

        public ExcelResult(HSSFWorkbook hssfworkbook, string filename)
        {
            this.hssfworkbook = hssfworkbook;
            this.filename = filename;
        }

        private HSSFWorkbook hssfworkbook;
        private string filename;

        private MemoryStream WriteToStream()
        {

            MemoryStream file = new MemoryStream();
            hssfworkbook.Write(file);
            return file;
        }

        public override void ExecuteResult(ControllerContext context)
        {
            var Response = context.HttpContext.Response;
            Response.ContentType = "application/vnd.ms-excel";
            Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", filename));
            Response.Clear();
            Response.BinaryWrite(WriteToStream().GetBuffer());
            Response.End();

        }
    }
}