using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace WindowsFormsApplication5
{

    class ExcelController
    {

        public void OutPutTest()
        {

            object[,] setValue = new object[9, 4];

            for (int i = 0; i < 9; i++)
            {
                for (int j = 0; j < 4; j++)
                {
                    setValue[i, j] = string.Format("{0}行{1}列目", i + 1, j + 1);
                }
            }

            ExcelManager em = new ExcelManager();
            em.test(setValue);

        }

    }

}
