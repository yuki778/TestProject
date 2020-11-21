using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication5
{

    class ExcelManager : IDisposable
    {

        #region Excel操作用オブジェクト

        public Application xlApp = null;
        public Workbooks xlBooks = null;
        public Workbook xlBook = null;
        public Sheets xlSheets = null;
        public Worksheet xlSheet = null;
        public Range xlRange = null;
        public Range xlCells = null;
        public bool isDispose = false;

        #endregion

        #region コンストラクタ・デストラクタ

        //コンストラクタ
        public ExcelManager()
        {
            xlApp = new Application();
        }

        //デストラクタ
        ~ExcelManager()
        {
            if (isDispose == false)
            {
                Dispose();
            }
        }

        #endregion

        #region リソース開放

        public void Dispose()
        {
            ReleaseExcelComObject(EnumReleaseMode.App);
            isDispose = true;
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Excelリソース解放
        /// </summary>
        /// <param name="ReleaseMode">リリース対象Enum</param>
        private void ReleaseExcelComObject(EnumReleaseMode ReleaseMode)
        {
            try
            {
                // xlSheet解放
                if (xlSheet != null)
                {
                    Marshal.ReleaseComObject(xlSheet);
                    xlSheet = null;
                }
                if (ReleaseMode == EnumReleaseMode.Sheet)
                    return;

                // xlSheets解放
                if (xlSheets != null)
                {
                    Marshal.ReleaseComObject(xlSheets);
                    xlSheets = null;
                }
                if (ReleaseMode == EnumReleaseMode.Sheets)
                    return;

                // xlBook解放
                if (xlBook != null)
                {
                    Marshal.ReleaseComObject(xlBook);
                    xlBook = null;
                }
                if (ReleaseMode == EnumReleaseMode.Book)
                    return;

                // xlBooks解放
                if (xlBooks != null)
                {
                    Marshal.ReleaseComObject(xlBooks);
                    xlBooks = null;
                }
                if (ReleaseMode == EnumReleaseMode.Books)
                    return;

                // xlApp解放
                if (xlApp != null)
                {
                    Marshal.ReleaseComObject(xlApp);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// リリース対象
        /// </summary>
        private enum EnumReleaseMode
        {
            Sheet,
            Sheets,
            Book,
            Books,
            App
        }

        #endregion

        public void test(object[,] setValue)
        {
            
            try
            {

                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlBooks = xlApp.Workbooks;
                xlBook = xlBooks.Add();
                xlSheets = xlBook.Worksheets;
                xlSheet = xlSheets[1];
                xlApp.Visible = true;

                int iColCnt = 0;
                int iRowCnt = 0;

                Range xlCellsFrom = null;
                Range xlRangeFrom = null;
                Range xlCellsTo = null;
                Range xlRangeTo = null;   
                Range xlTarget = null;

                //配列の要素数取得
                iRowCnt = setValue.GetLength(0) - 1;
                iColCnt = setValue.GetLength(1) - 1;

                // 貼り付け位置
                int col = 4;
                int row = 5;
                xlCellsFrom = xlSheet.Cells;
                xlRangeFrom = xlCellsFrom[row, col];
                xlCellsTo = xlSheet.Cells;
                xlRangeTo = xlCellsTo[row + iRowCnt, col + iColCnt];
                xlTarget = xlSheet.Range[xlRangeFrom, xlRangeTo];
                xlTarget.Value = setValue;

                // 画像貼付
                string imagefile = @"TestPath";
                double Left = xlRangeFrom.Left;
                double Top = xlRangeFrom.Top;
                float Width = 0;
                float Height = 0;

                Microsoft.Office.Interop.Excel.Shape shape = xlSheet.Shapes.AddPicture(imagefile, MsoTriState.msoTrue, MsoTriState.msoTrue, xlRangeTo.Left, xlRangeTo.Top, Width, Height);
                shape.ScaleHeight(0.5F, Microsoft.Office.Core.MsoTriState.msoCTrue);
                shape.ScaleWidth(0.5F, Microsoft.Office.Core.MsoTriState.msoCTrue);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                this.ReleaseExcelComObject(EnumReleaseMode.App);
            }
            
        }

    }

}
