using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using CsvHelper;
using System.Text.RegularExpressions;

namespace ExcelTool
{
    public class ExcelUtilities
    {
        public struct SinhVien
        {
            public string Stt;
            public string Ho;
            public string Ten;
            public string HoTen;
            public string NgaySinh;
        }
        Application xlApp = new Application();
        Workbook xlWorkBook;
        Worksheet xlWorkSheet;
        Range range;
        private Range khoangDuLieu;
        List<SinhVien> listSV = new List<SinhVien>();
        public ExcelUtilities()
        {

        }
        public ExcelUtilities(string path, int sheet)
        {
            xlWorkBook = xlApp.Workbooks.Open(path);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(sheet);

        }

        public void ReadExcel()
        {
            try
            {

                khoangDuLieu = xlWorkSheet.UsedRange;
                var soDong = khoangDuLieu.Rows.Count;
                var soCot = khoangDuLieu.Columns.Count;
                for (int dong = 1; dong <= soDong; dong++)
                {
                    SinhVien sv = new SinhVien();// goi structure
                    sv.Stt = LayDuLieu(khoangDuLieu.Cells[dong, 1] as Range);
                    sv.Ho = LayDuLieu(khoangDuLieu.Cells[dong, 2] as Range);
                    sv.Ten = LayDuLieu(khoangDuLieu.Cells[dong, 3] as Range);
                    sv.HoTen = LayDuLieu(khoangDuLieu.Cells[dong, 4] as Range);
                    sv.NgaySinh = LayDuLieu(khoangDuLieu.Cells[dong, 5] as Range);
                    listSV.Add(sv);
                }

                foreach (var item in listSV)
                {
                    Console.WriteLine("{0}, {1}, {2}, {3}, {4}", item.Stt, item.Ho, item.Ten, item.HoTen, item.NgaySinh);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                Marshal.ReleaseComObject(khoangDuLieu);
                Marshal.ReleaseComObject(xlWorkSheet);
                xlWorkBook.Close();
                Marshal.ReleaseComObject(xlWorkBook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }

        private string LayDuLieu(Range cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }
            return Convert.ToString(cell.Value2);
        }

        public void writeCSV(string tenFile)
        {
            string phanChia = ",";
            StringBuilder sb = new StringBuilder();
            foreach (var item in listSV)
            {
                sb.AppendLine(string.Join(phanChia, item.Stt, item.Ho,
                    item.Ten,
                    item.HoTen,
                    item.NgaySinh));
            }
            File.WriteAllText(tenFile, sb.ToString());
            Console.WriteLine("Success Chou Chou");
        }
    }
}
