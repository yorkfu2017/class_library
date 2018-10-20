using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
namespace WebApplication2
{
    public class ExcelHelp_York
    {
        /*
        HSSFWorkbook hssfworkbook;
        IFont font;
        NPOI.HSSF.Util.HSSFColor fillForegroundColor;
        FillPattern fillPattern;
        HSSFColor fillBackgroundColor;
        HorizontalAlignment ha;
        VerticalAlignment va;
        */


        /*
          foreach (PropertyDescriptor prop in props)
            {
                TableHeaderCell hcell = new TableHeaderCell();
                hcell.Text = prop.Name;
             //   hcell.BackColor = System.Drawing.Color.Yellow;
                row.Cells.Add(hcell);
            }
             */


        public static void WriteExcelWithNPOI(String extension, DataSet dt)
        {
            //excel文件格式判断
            //要区分.xlsx和.xls文件格式，所用的sheet对象区别
            /*xlsx是从Office2007开始使用的，是用新的基于XML的压缩文件格式取代了其目前专有的默认文件格式，
             * 在传统的文件名扩展名后面添加了字母x（即：docx取代doc、.xlsx取代xls等等），使其占用空间更小。
           
            */
            IWorkbook workbook;

            if (extension == "xlsx")
            {
                workbook = new XSSFWorkbook();
                // XSSFWorkbook workbook_XSS = new XSSFWorkbook();
                // 右击文件 属性信息
                {
                    //新版本的信息我也不知道在哪里设置。。。
                }

            }
            else if (extension == "xls")
            {
                //workbook = new HSSFWorkbook();
                HSSFWorkbook workbook_HSS = new HSSFWorkbook();
                // 右击文件 属性信息
                {
                    DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                    dsi.Company = "http://www.yongfa365.com/";
                    workbook_HSS.DocumentSummaryInformation = dsi;

                    SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                    si.Author = "柳永法"; //填加xls文件作者信息
                    si.ApplicationName = "NPOI测试程序"; //填加xls文件创建程序信息
                    si.LastAuthor = "柳永法2"; //填加xls文件最后保存者信息
                    si.Comments = "说明信息"; //填加xls文件作者信息
                    si.Title = "NPOI测试"; //填加xls文件标题信息
                    si.Subject = "NPOI测试Demo";//填加文件主题信息
                    si.CreateDateTime = DateTime.Now;
                    workbook_HSS.SummaryInformation = si;
                }
                workbook = workbook_HSS;
            }
            else
            {
                //throw new Exception("This format is not supported");
                workbook = new XSSFWorkbook();
            }

            // dll refered NPOI.dll and NPOI.OOXML  
            for (int k = 0; k < dt.Tables.Count; k++)
            {
                DataTable dtSource = dt.Tables[k];


                ICellStyle dateStyle = workbook.CreateCellStyle();
                IDataFormat format = workbook.CreateDataFormat();
                dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

                //取得列宽      相当于遍历啊，相当耗费资源啊
                int[] arrColWidth = new int[dtSource.Columns.Count];
                foreach (DataColumn item in dtSource.Columns)
                {
                    arrColWidth[item.Ordinal] = System.Text.Encoding.GetEncoding(936).GetBytes(item.ColumnName.ToString()).Length;
                }
                for (int i = 0; i < dtSource.Rows.Count; i++)
                {
                    for (int j = 0; j < dtSource.Columns.Count; j++)
                    {
                        int intTemp = System.Text.Encoding.GetEncoding(936).GetBytes(dtSource.Rows[i][j].ToString()).Length;
                        if (intTemp > arrColWidth[j])
                        {
                            arrColWidth[j] = intTemp;
                        }
                    }
                }

                //int rowIndex = 0;

                //ISheet sheet1 = workbook.CreateSheet(dtSource.TableName.ToString());

                ////make a header row  
                //IRow row1 = sheet1.CreateRow(0);

                //for (int j = 0; j < dtSource.Columns.Count; j++)
                //{

                //    ICell cell = row1.CreateCell(j);

                //    String columnName = dtSource.Columns[j].ToString();
                //    cell.SetCellValue(columnName);
                //}

                ////loops through data  
                //for (int i = 0; i < dtSource.Rows.Count; i++)
                //{
                //    IRow row = sheet1.CreateRow(i + 1);
                //    for (int j = 0; j < dtSource.Columns.Count; j++)
                //    {

                //        ICell cell = row.CreateCell(j);
                //        String columnName = dtSource.Columns[j].ToString();
                //        cell.SetCellValue(dtSource.Rows[i][columnName].ToString());
                //    }
                //}

                int rowIndex = 0;
                ISheet sheet1 = workbook.CreateSheet(dtSource.TableName.ToString());
                foreach (DataRow row in dtSource.Rows)
                {
                    var cra = new NPOI.SS.Util.CellRangeAddress(100, 100, 100, 100);
                    var cra1 = new NPOI.SS.Util.CellRangeAddress(10, 10, 10, 10);
                    sheet1.AddMergedRegion(cra);
                    sheet1.AddMergedRegion(cra1);
                    // 新建表，填充表头，填充列头，样式
                    if (rowIndex == 65535 || rowIndex == 0)
                    {


                       // 表头及样式
                        {
                            IRow headerRow = sheet1.CreateRow(0);
                            headerRow.HeightInPoints = 100;
                            headerRow.CreateCell(0).SetCellValue("The Great Wall is a Great Building");
                            headerRow.CreateCell(1).SetCellValue("The Great Wall is a Great Building1");

                            ICellStyle headStyle = workbook.CreateCellStyle();
                            //headStyle.Alignment = CellHorizontalAlignment.CENTER;
                            IFont font = workbook.CreateFont();
                            font.FontHeightInPoints = 100;
                            font.Boldweight = 1000;
                            headStyle.SetFont(font);

                            //单元格样式
                            var cellStyleBorder = workbook.CreateCellStyle();
                            cellStyleBorder.BorderBottom = BorderStyle.Thin;
                            cellStyleBorder.BorderLeft = BorderStyle.Thin;
                            cellStyleBorder.BorderRight = BorderStyle.Thin;
                            cellStyleBorder.BorderTop = BorderStyle.Thin;
                            cellStyleBorder.Alignment = HorizontalAlignment.Center;
                            cellStyleBorder.VerticalAlignment = VerticalAlignment.Center;

                            var cellStyleBorderAndColorGreen = workbook.CreateCellStyle();
                            cellStyleBorderAndColorGreen.CloneStyleFrom(cellStyleBorder);
                            cellStyleBorderAndColorGreen.CloneStyleFrom(headStyle);
                            //cellStyleBorderAndColorGreen.FillPattern = FillPattern.SolidForeground;
                            cellStyleBorderAndColorGreen.FillPattern = FillPattern.ThickBackwardDiagonals;
                            //cellStyleBorderAndColorGreen.FillBackgroundColor = //new SolidColorBrush(Color.FromArgb(255, 255, 0, 0, 0));


                            headerRow.GetCell(0).CellStyle = cellStyleBorderAndColorGreen;
                            headerRow.GetCell(1).CellStyle = headStyle;

                            //sheet.AddMergedRegion(new Region(0, 0, 0, dtSource.Columns.Count - 1));
                            //headerRow.Dispose();
                        }
                       


                      // 列头及样式
                        {
                            //行样式
                            IRow headerRow = sheet1.CreateRow(1);
                            ICellStyle headStyle = workbook.CreateCellStyle();
                            headStyle.BorderBottom = BorderStyle.Medium;
                            headStyle.BorderLeft = BorderStyle.Thin;
                            headStyle.BorderRight = BorderStyle.Thin;
                            headStyle.BorderTop = BorderStyle.Thin;
                            headStyle.Alignment = HorizontalAlignment.Center;
                            headStyle.VerticalAlignment = VerticalAlignment.Center;
                            headStyle.FillBackgroundColor = HSSFColor.Green.Index;

                            //headStyle.Alignment = CellHorizontalAlignment.CENTER;
                            IFont font = workbook.CreateFont();
                            font.FontHeightInPoints = 10;
                            font.Boldweight = 700;
                            headStyle.SetFont(font);

                            foreach (DataColumn column in dtSource.Columns)
                            {
                                headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);
                                headerRow.GetCell(column.Ordinal).CellStyle = headStyle;

                                //设置列宽
                                sheet1.SetColumnWidth(column.Ordinal, (arrColWidth[column.Ordinal] + 1) * 256);

                            }
                            // headerRow.Dispose();
                        }
                       

                        rowIndex = 2;
                    }
                   


                   // 填充内容
                    IRow dataRow = sheet1.CreateRow(rowIndex);
                    foreach (DataColumn column in dtSource.Columns)
                    {
                        ICell newCell = dataRow.CreateCell(column.Ordinal);

                        string drValue = row[column].ToString();

                        switch (column.DataType.ToString())
                        {
                            case "System.String"://字符串类型
                                newCell.SetCellValue(drValue);
                                break;
                            case "System.DateTime"://日期类型
                                DateTime dateV;
                                DateTime.TryParse(drValue, out dateV);
                                newCell.SetCellValue(dateV);

                                newCell.CellStyle = dateStyle;//格式化显示
                                break;
                            case "System.Boolean"://布尔型
                                bool boolV = false;
                                bool.TryParse(drValue, out boolV);
                                newCell.SetCellValue(boolV);
                                break;
                            case "System.Int16"://整型
                            case "System.Int32":
                            case "System.Int64":
                            case "System.Byte":
                                int intV = 0;
                                int.TryParse(drValue, out intV);
                                newCell.SetCellValue(intV);
                                break;
                            case "System.Decimal"://浮点型
                            case "System.Double":
                                double doubV = 0;
                                double.TryParse(drValue, out doubV);
                                newCell.SetCellValue(doubV);
                                break;
                            case "System.DBNull"://空值处理
                                newCell.SetCellValue("");
                                break;
                            default:
                                newCell.SetCellValue("");
                                break;
                        }

                    }
                    

                    rowIndex++;
                }
            }

            using (var exportData = new MemoryStream())
            {
                exportData.Flush();
                exportData.Position = 0;
                //HttpResponse Response;
                //System.Web.HttpContext httpContext
                //HttpResponse.Clear();
                workbook.Write(exportData);
                if (extension == "xlsx") //xlsx file format  
                {/*
                    HttpResponse response = httpContext.Response;             
                    response.Clear();
                    response.BufferOutput = true;
                    response.StatusCode = 200; // HttpStatusCode.OK;
                    response.Write("Hello");
                    response.ContentType = "text/xml";
                    response.End();
                                     */
                 
                    HttpContext curContext = HttpContext.Current;
                    // curContext.Response.ContentType = "application/vnd.ms-excel";
                    curContext.Response.ContentEncoding = System.Text.Encoding.UTF8;
                    curContext.Response.Charset = "";
                    curContext.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    curContext.Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", "tpms_Dict"+DateTime.Now.ToString()+".xlsx"));
                    curContext.Response.BinaryWrite(exportData.ToArray());
                    curContext.Response.End();
                }
                else if (extension == "xls")  //xls file format  
                {
                    HttpContext curContext = HttpContext.Current;
                    curContext.Response.ContentEncoding = System.Text.Encoding.UTF8;
                    curContext.Response.Charset = "";
                    curContext.Response.ContentType = "application/vnd.ms-excel";
                    curContext.Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", "tpms_Dict" + DateTime.Now.ToString() + ".xls"));
                    curContext.Response.BinaryWrite(exportData.GetBuffer());
                    curContext.Response.End();
                }
                workbook.Dispose();

            }
        }
    }
}


    