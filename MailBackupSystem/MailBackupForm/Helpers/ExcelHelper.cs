using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using MailBackupForm.Properties;

namespace MailBackupForm.Helpers
{
    class ExcelHelper
    {
        ISheet sheet;
        XSSFWorkbook workbook;
        IRow CurrentRow;
        static int rowCounter = 1;
        Form1 form = new Form1();
        public ExcelHelper()
        {
            Console.WriteLine("EXCEL HELPER");
            workbook = new XSSFWorkbook();

            sheet = workbook.CreateSheet("Report");
            IRow HeaderRow = sheet.CreateRow(0);

            //Create The Actual Cells
            CreateCell(HeaderRow, 0, "Tarih");
            CreateCell(HeaderRow, 1, "Gönderen");
            CreateCell(HeaderRow, 2, "Konu");
            CreateCell(HeaderRow, 3, "Alıcı");
            CreateCell(HeaderRow, 4, "Bcc");
            CreateCell(HeaderRow, 5, "Cc");
            CreateCell(HeaderRow, 6, "İçerik");
            CreateCell(HeaderRow, 7, "Dizin");
            //WriteExcel(@"C:\Users\kandi\Desktop\SamplePST\ExcelDocs\");
        }

        public void ReadFromTxt(string txtPath, string excelPath)
        {
            if (!Directory.Exists(excelPath))
                Directory.CreateDirectory(excelPath);


            CurrentRow = sheet.CreateRow(rowCounter);

            foreach (var item in File.ReadAllLines(txtPath + @"\output.txt"))
            {

                var mail = JsonConvert.DeserializeObject<Mail>(item);

                CreateCell(CurrentRow, 0, mail.DeliveryTime.ToString());
                CreateCell(CurrentRow, 1, mail.SenderEmailAddress);
                CreateCell(CurrentRow, 2, mail.Subject);
                CreateCell(CurrentRow, 3, mail.DisplayTo);
                CreateCell(CurrentRow, 4, mail.DisplayBcc);
                CreateCell(CurrentRow, 5, mail.DisplayCc);
                CreateCell(CurrentRow, 6, mail.Body);
                CreateCell(CurrentRow, 7, mail.FilePath);
                rowCounter++;
                CurrentRow = sheet.CreateRow(rowCounter);
            }
            Console.WriteLine("Toplam mail sayısı: " + rowCounter);

            using (var fileData = new FileStream(excelPath + @"\MailReport.xlsx", FileMode.Create))
            {
                workbook.Write(fileData);
            }
        }




        private static void CreateCell(IRow CurrentRow, int CellIndex, string Value)
        {
            ICell Cell = CurrentRow.CreateCell(CellIndex);
            if (Value == null)
                Value = "";
            if (Value.Length <= 32700)
                Cell.SetCellValue(Value);
            if (Value.Length > 32700)
            {
                var temp = Value.Substring(0, 32000);
                Cell.SetCellValue(temp);
            }
        }

    }
}
