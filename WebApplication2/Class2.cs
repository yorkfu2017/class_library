using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace WebApplication2
{
    public class Class2
    {

        public static void datatoexcel(string excelFilePath)
        {
            IWorkbook workbook;
            using (FileStream stream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(stream);
            }

            ISheet sheet = workbook.GetSheetAt(0); // zero-based index of your target sheet
            DataTable dt = new DataTable(sheet.SheetName);

            // write header row
            IRow headerRow = sheet.GetRow(0);
            foreach (ICell headerCell in headerRow)
            {
                dt.Columns.Add(headerCell.ToString());
            }

            // write the rest
            int rowIndex = 0;
            foreach (IRow row in sheet)
            {
                // skip header row
                if (rowIndex++ == 0) continue;
                DataRow dataRow = dt.NewRow();
                dataRow.ItemArray = row.Cells.Select(c => c.ToString()).ToArray();
                dt.Rows.Add(dataRow);
            }
        }



        public  DataTable datatoexcel_1(String Path)
        {
            XSSFWorkbook wb;
            XSSFSheet sh;
            String Sheet_name;

            using (var fs = new FileStream(Path, FileMode.Open, FileAccess.Read))
            {
                wb = new XSSFWorkbook(fs);

                Sheet_name = wb.GetSheetAt(0).SheetName;  //get first sheet name
            }
            DataTable DT = new DataTable();
            DT.Rows.Clear();
            DT.Columns.Clear();

            // get sheet
            sh = (XSSFSheet)wb.GetSheet(Sheet_name);

            int i = 0;
            while (sh.GetRow(i) != null)
            {
                // add neccessary columns
                if (DT.Columns.Count < sh.GetRow(i).Cells.Count)
                {
                    for (int j = 0; j < sh.GetRow(i).Cells.Count; j++)
                    {
                        DT.Columns.Add("", typeof(string));
                    }
                }

                // add row
                DT.Rows.Add();

                // write row value
                for (int j = 0; j < sh.GetRow(i).Cells.Count; j++)
                {
                    var cell = sh.GetRow(i).GetCell(j);

                    if (cell != null)
                    {
                        // TODO: you can add more cell types capatibility, e. g. formula
                        switch (cell.CellType)
                        {
                            case NPOI.SS.UserModel.CellType.Numeric:
                                DT.Rows[i][j] = sh.GetRow(i).GetCell(j).NumericCellValue;
                                //dataGridView1[j, i].Value = sh.GetRow(i).GetCell(j).NumericCellValue;

                                break;
                            case NPOI.SS.UserModel.CellType.String:
                                DT.Rows[i][j] = sh.GetRow(i).GetCell(j).StringCellValue;

                                break;
                        }
                    }
                }

                i++;
            }

            return DT;
        }

        public DataTable datatoexcel_2(Stream str)
        {
            XSSFWorkbook hssfworkbook = new XSSFWorkbook(str);
            ISheet sheet = hssfworkbook.GetSheetAt(0);
            str.Close();

            DataTable dt = new DataTable();
            IRow headerRow = sheet.GetRow(0);
            System.Collections.IEnumerator rows = sheet.GetRowEnumerator();

            int colCount = headerRow.LastCellNum;
            int rowCount = sheet.LastRowNum;

            for (int c = 0; c < colCount; c++)
                dt.Columns.Add(headerRow.GetCell(c).ToString());

            while (rows.MoveNext())
            {
                IRow row = (XSSFRow)rows.Current;
                DataRow dr = dt.NewRow();

                for (int i = 0; i < colCount; i++)
                {
                    ICell cell = row.GetCell(i);

                    if (cell != null)
                        dr[i] = cell.ToString();
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        #region excel to datatable 

        private static ISheet GetFileStream(string fullFilePath)
        {
            var fileExtension = Path.GetExtension(fullFilePath);
            string sheetName;
            ISheet sheet = null;
            //要区分.xlsx和.xls文件格式，所用的sheet对象区别
            /*xlsx是从Office2007开始使用的，是用新的基于XML的压缩文件格式取代了其目前专有的默认文件格式，
             * 在传统的文件名扩展名后面添加了字母x（即：docx取代doc、.xlsx取代xls等等），使其占用空间更小。
             * */
            switch (fileExtension)
            {
                case ".xlsx":
                    using (var fs = new FileStream(fullFilePath, FileMode.Open, FileAccess.Read))
                    {
                        var wb = new XSSFWorkbook(fs);
                        sheetName = wb.GetSheetAt(0).SheetName;
                        sheet = (XSSFSheet)wb.GetSheet(sheetName);
                    }
                    break;
                case ".xls":
                    using (var fs = new FileStream(fullFilePath, FileMode.Open, FileAccess.Read))
                    {
                        var wb = new HSSFWorkbook(fs);
                        sheetName = wb.GetSheetAt(0).SheetName;
                        sheet = (HSSFSheet)wb.GetSheet(sheetName);
                    }
                    break;
            }
            return sheet;
        }

        private static DataTable GetRequestsDataFromExcel(string fullFilePath)
        {
            try
            {
                var sh = GetFileStream(fullFilePath);
                var dtExcelTable = new DataTable();
                dtExcelTable.Rows.Clear();
                dtExcelTable.Columns.Clear();
                var headerRow = sh.GetRow(0);
                int colCount = headerRow.LastCellNum;
                for (var c = 0; c < colCount; c++)
                    dtExcelTable.Columns.Add(headerRow.GetCell(c).ToString());
                var i = 1;
                var currentRow = sh.GetRow(i);
                while (currentRow != null)
                {
                    var dr = dtExcelTable.NewRow();
                    for (var j = 0; j < currentRow.Cells.Count; j++)
                    {
                        var cell = currentRow.GetCell(j);
                        //获得的cell 数据类型 处理
                        if (cell != null)
                            switch (cell.CellType)
                            {
                                case CellType.Numeric:
                                    dr[j] = DateUtil.IsCellDateFormatted(cell)
                                        ? cell.DateCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture)
                                        : cell.NumericCellValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                                    break;
                                case CellType.String:
                                    dr[j] = cell.StringCellValue;
                                    break;
                                case CellType.Blank:
                                    dr[j] = string.Empty;
                                    break;
                            }
                    }
                    dtExcelTable.Rows.Add(dr);
                    i++;
                    currentRow = sh.GetRow(i);
                }
                return dtExcelTable;
            }
            catch (Exception e)
            {
                throw;
            }
        }
        #endregion

    }
}