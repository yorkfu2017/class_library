﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebApplication2
{
    public class style
    {
        //=============================================================================================================================

        /*  
          var workbook = new XSSFWorkbook();
           var sheet = workbook.CreateSheet("Commission");
           var row = sheet.CreateRow(0);

           var bStylehead = workbook.CreateCellStyle();
           bStylehead.BorderBottom = BorderStyle.Thin;
   bStylehead.BorderLeft = BorderStyle.Thin;
   bStylehead.BorderRight = BorderStyle.Thin;
   bStylehead.BorderTop = BorderStyle.Thin;
   bStylehead.Alignment = HorizontalAlignment.Center;
   bStylehead.VerticalAlignment = VerticalAlignment.Center;         
   bStylehead.FillBackgroundColor = HSSFColor.Green.Index;

   row.CreateCell(0);
   row.CreateCell(1);

   var r2 = sheet.CreateRow(1);
           r2.CreateCell(0, CellType.String).SetCellValue("Name");
           r2.CreateCell(1, CellType.String).SetCellValue("Address");
           r2.CreateCell(2, CellType.String).SetCellValue("city");
           r2.CreateCell(3, CellType.String).SetCellValue("state");

           var cra = new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 1);
           var cra1 = new NPOI.SS.Util.CellRangeAddress(0, 0, 2, 3);
           sheet.AddMergedRegion(cra);                         
   sheet.AddMergedRegion(cra1);         

   ICell cell = sheet.GetRow(0).GetCell(0);
           cell.SetCellType(CellType.String);
   cell.SetCellValue("Supplier Provided Data");
   cell.CellStyle = bStylehead;

   ICell cell1 = sheet.GetRow(0).GetCell(1);
           cell1.SetCellType(CellType.String);
   cell1.SetCellValue("Deal Provided Data");
   cell1.CellStyle = bStylehead; 
   
   */



        //=============================================================================================================================


        /*
                 var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Commission");
        var row = sheet.CreateRow(0);

        var cellStyleBorder = workbook.CreateCellStyle();
        cellStyleBorder.BorderBottom = BorderStyle.Thin;
        cellStyleBorder.BorderLeft = BorderStyle.Thin;
        cellStyleBorder.BorderRight = BorderStyle.Thin;
        cellStyleBorder.BorderTop = BorderStyle.Thin;
        cellStyleBorder.Alignment = HorizontalAlignment.Center;
        cellStyleBorder.VerticalAlignment = VerticalAlignment.Center;

        var cellStyleBorderAndColorGreen = workbook.CreateCellStyle();
        cellStyleBorderAndColorGreen.CloneStyleFrom(cellStyleBorder);
        cellStyleBorderAndColorGreen.FillPattern = FillPattern.SolidForeground;
        ((XSSFCellStyle)cellStyleBorderAndColorGreen).SetFillForegroundColor(new XSSFColor(new byte[] { 198, 239, 206 }));

        var cellStyleBorderAndColorYellow = workbook.CreateCellStyle();
        cellStyleBorderAndColorYellow.CloneStyleFrom(cellStyleBorder);
        cellStyleBorderAndColorYellow.FillPattern = FillPattern.SolidForeground;
        ((XSSFCellStyle)cellStyleBorderAndColorYellow).SetFillForegroundColor(new XSSFColor(new byte[] { 255, 235, 156 }));

        row.CreateCell(0);
        row.CreateCell(1);
        row.CreateCell(2);
        row.CreateCell(3);

        var r2 = sheet.CreateRow(1);
        r2.CreateCell(0, CellType.String).SetCellValue("Name");
        r2.Cells[0].CellStyle = cellStyleBorderAndColorGreen;
        r2.CreateCell(1, CellType.String).SetCellValue("Address");
        r2.Cells[1].CellStyle = cellStyleBorderAndColorGreen;
        r2.CreateCell(2, CellType.String).SetCellValue("city");
        r2.Cells[2].CellStyle = cellStyleBorderAndColorYellow;
        r2.CreateCell(3, CellType.String).SetCellValue("state");
        r2.Cells[3].CellStyle = cellStyleBorderAndColorYellow;
        var cra = new NPOI.SS.Util.CellRangeAddress(0, 0, 0, 1);
        var cra1 = new NPOI.SS.Util.CellRangeAddress(0, 0, 2, 3);
        sheet.AddMergedRegion(cra);
        sheet.AddMergedRegion(cra1);

        ICell cell = sheet.GetRow(0).GetCell(0);
        cell.SetCellType(CellType.String);
        cell.SetCellValue("Supplier Provided Data");
        cell.CellStyle = cellStyleBorderAndColorGreen;
        sheet.GetRow(0).GetCell(1).CellStyle = cellStyleBorderAndColorGreen;

        ICell cell1 = sheet.GetRow(0).GetCell(2);
        cell1.SetCellType(CellType.String);
        cell1.SetCellValue("Deal Provided Data");
        cell1.CellStyle = cellStyleBorderAndColorYellow;
        sheet.GetRow(0).GetCell(3).CellStyle = cellStyleBorderAndColorYellow;

        using (FileStream fs = new FileStream(@"c:\temp\excel\test.xlsx", FileMode.Create, FileAccess.Write))
        {
            workbook.Write(fs);
        }
                 */
        //=============================================================================================================================
    }
}