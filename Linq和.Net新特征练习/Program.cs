using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Linq和.Net新特征练习
{
    class Program
    {
        public static string connString = "Server=.;DataBase=Test;Uid=sa;Pwd=abcd1234...";
        static void Main(string[] args)
        {
            //SearchFromArr();


            SqlConnection conn = new SqlConnection(connString);
            SqlCommand cmd = new SqlCommand("Select *from People", conn);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataSet ds = new DataSet();
            da.Fill(ds);
            if (ds.Tables[0]!=null)
            {
                DataTable ddt = ds.Tables[0];
                doExport(ddt, @"D:\1111.xlsx", "people");
            }


        }

        public static void SearchFromArr()
        {
            var arr = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9 };
            //使用linq和lambda表达式从数组中获取所有偶数，并求三次方
            //var list = arr.Where(item => item % 2 == 0)
            //    .Select(item => Math.Pow(item,3))
            //    .OrderBy(item => item);

            var list = from a in arr
                       where a % 2 == 0
                       orderby a descending
                       select Math.Pow(a, 3);

            foreach (var item in list)
            {
                Console.WriteLine(item);
            }


        }

        private static void doExport(DataTable dt, string toFileName, string strSheetName)
        {
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application(); //Execl的操作类
                                                                                                                 //读取保存目标的对象bai
            Microsoft.Office.Interop.Excel.Workbook bookDest = excel.Workbooks._Open(toFileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value
            , Missing.Value, Missing.Value, Missing.Value, Missing.Value
            , Missing.Value, Missing.Value, Missing.Value, Missing.Value);//打开要导出到的Execl文件的du工作薄。--ps:关于Missing类在这里的作用，我也不知道...囧
            Microsoft.Office.Interop.Excel.Worksheet sheetDest = bookDest.Worksheets.Add(Missing.Value, Missing.Value, Missing.Value, Missing.Value) as Microsoft.Office.Interop.Excel.Worksheet;//给工作薄添加一个Sheet
            sheetDest.Name = strSheetName;//自己定义名字O(∩_∩)O哈哈~
            int rowIndex = 1;
            int colIndex = 0;
            excel.Application.Workbooks.Add(true);//这句不写不知道会不会报错
            foreach (DataColumn col in dt.Columns)
            {
                colIndex++;
                sheetDest.Cells[1, colIndex] = col.ColumnName;//Execl中的第一列,把DataTable的列名先导进去
            }
            //导入数据行
            foreach (DataRow row in dt.Rows)
            {
                rowIndex++;
                colIndex = 0;
                foreach (DataColumn col in dt.Columns)
                {
                    colIndex++;
                    sheetDest.Cells[rowIndex, colIndex] = row[col.ColumnName].ToString();
                }
            }
            bookDest.Saved = true;
            bookDest.Save();
            excel.Quit();
            excel = null;
            GC.Collect();//垃圾回收
        }
    }
}
