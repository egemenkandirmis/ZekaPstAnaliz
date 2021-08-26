using Spire.Email.Outlook;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using Spire.Email;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace MailBackupSystem
{
    class Program
    {

        
        //private static string path = @"C:\Users\kandi\Desktop\SamplePST\Mail\$$$$$\";
        private static string path = @"";

        //private static string pstPath = @"C:\Users\kandi\Desktop\SamplePST\archive.pst";
        private static string pstPath = @"";
        //private static string txtPath = @"C:\Users\kandi\Desktop\SamplePST\TxtDocs\";
        private static string txtPath = @"";
        //private static string excelPath = @"C:\Users\kandi\Desktop\SamplePST\ExcelDocs\";
        private static string excelPath = @"";


        static void Main(string[] args)
        {
            DateTime start = DateTime.Now;
            Console.WriteLine("Başlangıç: "+ start.ToString());
            Clear();
            MailHelper mailHelper = new MailHelper(pstPath, path);
            ExcelHelper excelHelper = new ExcelHelper();
            excelHelper.ReadFromTxt(txtPath, excelPath);
            DateTime end = DateTime.Now;
            Console.WriteLine("Bitiş: " + end.ToString());
            var fark = end.Subtract(start).TotalMinutes;
            Console.WriteLine("Geçen süre: " + fark);
            Console.ReadKey();
        }

        public static void Clear()
        {
            if (Directory.Exists(excelPath))
                Directory.Delete(excelPath,true);


            if (Directory.Exists(path.Split(@"$")[0]))
                Directory.Delete(path.Split(@"$")[0],true);


            if (Directory.Exists(txtPath))
                Directory.Delete(txtPath,true); 
        }

        public static void SetPstFileName(string pstName)
        {
            pstPath = pstName;
        }
        public static void SetRootFolderName(string folderName)
        {
            var fullFolderName = folderName + @"\Mail\$$$$$";
            var fullTxtFolderName = folderName + @"\Mail\TxtDocs";
            var fullExcelFolderName = folderName + @"\Mail\ExcelDocs";
            path = fullFolderName;
            excelPath = fullExcelFolderName;
            txtPath = fullTxtFolderName;
        }


    }
}

