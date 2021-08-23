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

        private static string path = @"C:\Users\kandi\Desktop\SamplePST\Mail\$$$$$\";
        //private static string path = @"C:\Users\kandi\Desktop\SamplePST\Export\";
        private static string pstPath = @"C:\Users\kandi\Desktop\SamplePST\yahya.kisioglu.pst";
        private static string excelPath = @"C:\Users\kandi\Desktop\SamplePST\ExcelDocs\";
        //private static static string pstPath = @"C:\Users\kandi\Desktop\SamplePST\archive.pst";
        //private static string pstPath = @"C:\Users\kandi\Desktop\SamplePST\ege.pst";
        static void Main(string[] args)
        {
            //MailHelper mailHelper = new MailHelper(pstPath,path);
            ExcelHelper excelHelper = new ExcelHelper();
            excelHelper.ReadFromTxt(@"C:\Users\kandi\Desktop\SamplePST\TxtDocs\", excelPath);
        }

    }
}

