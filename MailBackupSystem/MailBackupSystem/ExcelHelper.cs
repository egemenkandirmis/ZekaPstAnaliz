using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;

namespace MailBackupSystem
{
    class ExcelHelper
    {
        ISheet sheet;
        XSSFWorkbook workbook;
        IRow CurrentRow;
        int counter = 0;
        public ExcelHelper()
        {
            workbook = new XSSFWorkbook();

            sheet = workbook.CreateSheet("Report");
            IRow HeaderRow = sheet.CreateRow(0);

            //Create The Actual Cells
            CreateCell(HeaderRow, 0, "Tarih");
            CreateCell(HeaderRow, 1, "Gönderen");
            CreateCell(HeaderRow, 2, "Alıcı");
            CreateCell(HeaderRow, 3, "Cc");
            CreateCell(HeaderRow, 4, "Bcc");
            CreateCell(HeaderRow, 5, "Konu");
            CreateCell(HeaderRow, 6, "İçerik");
            CreateCell(HeaderRow, 7, "Dizin");
        }

        public void CreateData(List<string> icerik ,int rowIndex)
        {

            CurrentRow = sheet.CreateRow(rowIndex);
            counter = 0;
            foreach (var item in icerik)
            {
                CreateCell(CurrentRow, counter, item.ToString());
                counter++;
            }
            counter = 0;
            AutoSize();

        }

        public void AutoSize()
        {
            // This will be used to calculate the merge area
            int lastColumNum = sheet.GetRow(0).LastCellNum;
            for (int i = 0; i < 2; i++)
            {
                sheet.AutoSizeColumn(i);
                GC.Collect();
            }
        }

        public void WriteExcel(string excelPath)
        {
            if (!Directory.Exists(excelPath))
                Directory.CreateDirectory(excelPath);

            using (var fileData = new FileStream(excelPath + @"\MailReport.xlsx", FileMode.Create))
            {
                workbook.Write(fileData);
            }
        }


        private static void CreateCell(IRow CurrentRow, int CellIndex, string Value)
        {
            ICell Cell = CurrentRow.CreateCell(CellIndex);
            if(Value.Length <= 32700)
                Cell.SetCellValue(Value);
            if(Value.Length > 32700)
            {
                var temp = Value.Substring(0, 32000);
                Cell.SetCellValue(temp);
            }
        }

    }
}
