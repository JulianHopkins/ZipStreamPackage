using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;

namespace ZipStreamPackage.cs
{


    public class ExcellBuilder
    {
        public const string sheetname = "Лист1";
        public HSSFWorkbook GetWorkbook(string Company, string Subject)
        {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = Company;
            hssfworkbook.DocumentSummaryInformation = dsi;

            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = Subject;
            hssfworkbook.SummaryInformation = si;
            return hssfworkbook;
        }


        public void BuildCells(HSSFWorkbook hssfworkbook, List<MergedExcel> store, BorderStyle bs = BorderStyle.Medium)
        {
            ISheet sheet1 = null;
            ICellStyle style1 = null;
            ICellStyle style2 = null;
            ICellStyle style3 = null;
            List<ICellStyle> styles = null;
            if (hssfworkbook.NumberOfSheets == 0)
            {
                sheet1 = hssfworkbook.CreateSheet(sheetname);
                sheet1.SetMargin(MarginType.RightMargin, (double)0.5);
                sheet1.SetMargin(MarginType.TopMargin, (double)0.6);
                sheet1.SetMargin(MarginType.LeftMargin, (double)0.4);
                sheet1.SetMargin(MarginType.BottomMargin, (double)0.3);

                sheet1.HorizontallyCenter = true;

                style1 = hssfworkbook.CreateCellStyle();
                style2 = hssfworkbook.CreateCellStyle();
                style3 = hssfworkbook.CreateCellStyle();

                style1.Alignment = HorizontalAlignment.Left;
                style1.VerticalAlignment = VerticalAlignment.Top;
                style2.Alignment = HorizontalAlignment.Center;
                style2.VerticalAlignment = VerticalAlignment.Top;
                style3.Alignment = HorizontalAlignment.Right;
                style3.VerticalAlignment = VerticalAlignment.Top;

                style1.BorderBottom = BorderStyle.Thin;
                style1.BorderLeft = BorderStyle.Thin;
                style1.BorderRight = BorderStyle.Thin;
                style1.BorderTop = BorderStyle.Thin;

                style2.BorderBottom = BorderStyle.Thin;
                style2.BorderLeft = BorderStyle.Thin;
                style2.BorderRight = BorderStyle.Thin;
                style2.BorderTop = BorderStyle.Thin;

                style3.BorderBottom = BorderStyle.Thin;
                style3.BorderLeft = BorderStyle.Thin;
                style3.BorderRight = BorderStyle.Thin;
                style3.BorderTop = BorderStyle.Thin;

                styles = new List<ICellStyle>() { style1, style2, style3 };

                sheet1.PrintSetup.NoColor = true;
                sheet1.PrintSetup.Landscape = true;
                sheet1.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet1.FitToPage = true;
                sheet1.IsPrintGridlines = false;
                sheet1.SetActiveCell(-1, -1);
            }
            else sheet1 = hssfworkbook.GetSheetAt(0);

            int oldrow = 0;

            foreach (var rowgrp in store.GroupBy(k => k.row).OrderBy(g => g.First().row))
            {
                int nowrow = rowgrp.First().row;
                int row = sheet1.LastRowNum + (nowrow - oldrow);
                IRow row1 = sheet1.CreateRow(row);

                int rows = rowgrp.Max(f => f.rows);
                oldrow = nowrow;
                int begincell = 0;
                int endcell = 0;
                ICell cell = null;
                foreach (MergedExcel mcell in rowgrp.OrderBy(h => h.rowcount))
                {
                    begincell = row1.Cells.Count();
                    endcell = begincell + mcell.cols - 1;
                    cell = row1.CreateCell(begincell);


                    cell.SetCellValue(mcell.value);
                    cell.CellStyle = styles.FirstOrDefault(f => f.Alignment == mcell.horal);
                    CellRangeAddress region = new CellRangeAddress(row, row + rows - 1, begincell, endcell);
                    sheet1.AddMergedRegion(region);

                    if (endcell > begincell)
                    {
                        for (int icell = begincell + 1; icell <= endcell; icell++)
                        {
                            ICell tmp = row1.CreateCell(icell);
                            tmp.CellStyle = cell.CellStyle;
                        }
                    }


                }
                if (rows > 1)
                {
                    for (int rs = row + 1; rs < rows + row; rs++)
                    {
                        IRow rowtmp = null;
                        if (rs > sheet1.LastRowNum) rowtmp = sheet1.CreateRow(sheet1.LastRowNum + 1);
                        else rowtmp = sheet1.GetRow(rs);
                        for (int icell = begincell; icell <= endcell; icell++)
                        {
                            ICell tmp = rowtmp.CreateCell(icell);
                            tmp.CellStyle = cell.CellStyle;
                        }

                    }
                }

            }

            for (int i = 0; i < store.Max(f => f.rowcount) + 1; i++)
            {

                sheet1.AutoSizeColumn(i, true);

                int wth = sheet1.GetColumnWidth(i);
                sheet1.SetColumnWidth(i, wth + 500);
            }


        }

        public class MergedExcel
        {
            ///<summary>
            ///Счетчик колонки в строке
            ///</summary>
            public int rowcount { get; set; }
            ///<summary>
            ///Номер строки
            ///</summary>
            public int row { get; set; }
            ///<summary>
            ///Число строк
            ///</summary>
            public int rows { get; set; }
            ///<summary>
            ///Число объединяемых клеток по горизонтали
            ///</summary>
            public int cols { get; set; }
            ///<summary>
            ///Толщина верхнего бордюра
            ///</summary>
            public bool top { get; set; }
            ///<summary>
            ///Толщина правого бордюра
            ///</summary>
            public bool right { get; set; }
            ///<summary>
            ///Толщина нижнего бордюра
            ///</summary>
            public bool bottom { get; set; }
            ///<summary>
            ///Толщина левого бордюра
            ///</summary>
            public bool left { get; set; }
            ///<summary>
            ///Значение 
            ///</summary>
            public string value { get; set; }
            ///<summary>
            ///Horizontal Alignment 
            ///</summary>
            public HorizontalAlignment horal { get; set; }

        }

        public class CellsHolder
        {
            public List<MergedExcel> store;
            public DateTime? datereport;
            private int width;
            public CellsHolder(DateTime? datereport, int width)
            {
                store = new List<MergedExcel>();
                this.datereport = datereport;
                this.width = width;
            }
            public void AddCell(int row, int cols, bool top, bool right, bool bottom, bool left, string value, HorizontalAlignment horal)
            {
                int rowcount = 0;
                if (store.Count > 0) rowcount = store.Where(k => k.row == row).Count();
                int rows = 0;
                string v = SplittingStr(value, width * cols, "", out rows);
                store.Add(new MergedExcel() { rowcount = rowcount, row = row, cols = cols, top = top, right = right, bottom = bottom, left = left, horal = horal, value = v, rows = rows });

            }
            private string SplittingStr(string value, int width, string first, out int rows)
            {
                rows = 0;
                if (String.IsNullOrWhiteSpace(value)) return value;
                Func<int, string, int> spacesearch = (beg, str) =>
                {
                    for (int y = beg; y < str.Length; y++)
                    {
                        if ((Char.IsWhiteSpace(value[y])) || (Char.IsPunctuation(value[y]))) return y + 1;

                    }
                    return str.Length;
                };
                int begin = 0;
                int end = width - first.Length;
                StringBuilder sb = new StringBuilder();
                string predict = first;
                while (begin < value.Length)
                {
                    if (end > value.Length) end = value.Length;
                    else end = spacesearch(end, value);
                    sb.AppendLine(String.Format("{1}{0}", value.Substring(begin, end - begin), predict));
                    predict = "";
                    begin = end;
                    end = end + width;
                    ++rows;

                }
                return sb.ToString();


            }
        }
    }
}