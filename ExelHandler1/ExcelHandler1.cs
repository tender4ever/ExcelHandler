using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelDriver
{

    /// <summary>
    /// 需要參考 Microsoft.Office.Interop.Excel.dll
    /// Excel 檔案讀取器
    /// </summary>
    class ExcelHandler1 : IDisposable
    {

        #region "Property"

        /// <summary>
        /// File 檔案路徑
        /// </summary>
        private string FilePath { get; set; }

        /// <summary>
        /// Sheet
        /// </summary>
        private int Sheet { get; set; }


        private int aRowCount;

        /// <summary>
        /// Row Count
        /// </summary>
        public int RowCount
        {
            get { return aRowCount; }
        }


        private int aColumnCount;

        /// <summary>
        /// Column Count
        /// </summary>
        public int ColumnCount
        {
            get { return aColumnCount; }
        }

        #endregion



        #region "Excel Property"

        Excel.Application xlApp;

        Excel.Workbook xlWorkbook;

        Excel._Worksheet xlWorksheet;

        Excel.Range xlRange;

        #endregion



        #region "建構子'

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="aFilePath"></param>
        public ExcelHandler1(string aFilePath, int sSheet)
        {
            FilePath = aFilePath;

            Sheet = sSheet;

            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(FilePath);
            xlWorksheet = xlWorkbook.Sheets[Sheet];
            xlRange = xlWorksheet.UsedRange;

            aRowCount = xlRange.Rows.Count;
            aColumnCount = xlRange.Columns.Count;
        }


        /// <summary>
        /// 釋放記憶體
        /// </summary>
        public void Dispose()
        {
            this.Dispose();
        }

        #endregion



        #region "Method"

        /// <summary>
        /// 讀取檔案資料
        /// </summary>
        /// <param name="aRow"> 行 </param>
        /// <param name="aColumn"> 列 </param>
        /// <returns></returns>
        public string Read(string aRow, string aColumn)
        {
            int Row = int.Parse(aRow);
            int Column = int.Parse(aColumn);

            if (xlRange.Cells[Row, Column].Value2 == null)
            {
                return null;
            }
            else
            {
                return xlRange.Cells[Row, Column].Value2.ToString();
            }
         
        }


        /// <summary>
        /// 寫入檔案資料
        /// </summary>
        /// <param name="aRow"> 行 </param>
        /// <param name="aColumn"> 列 </param>
        /// <param name="aValue"> 設定值 </param>
        public void Wirte(string aRow, string aColumn, string aValue)
        {
            int Row = int.Parse(aRow);
            int Column = int.Parse(aColumn);

            xlRange.Cells[Row, Column].Value2 = aValue;

        }


        /// <summary>
        /// 關閉 Excel Reader
        /// </summary>
        public void Close()
        {
            xlWorkbook.Save();

            xlWorkbook.Close(Type.Missing, Type.Missing, Type.Missing);
            
            xlWorkbook = null;

            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
        }

        #endregion

    }
}
