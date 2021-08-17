using Spire.Email.Outlook;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using Spire.Email;

namespace MailBackupSystem
{
    class Program
    {

        private static string path = @"C:\Users\kandi\Desktop\SamplePST\$$$$$\";
        //private static string path = @"C:\Users\kandi\Desktop\SamplePST\Export\";
        private static string pstPath = @"C:\Users\kandi\Desktop\SamplePST\yahya.kisioglu.pst";
        //private static static string pstPath = @"C:\Users\kandi\Desktop\SamplePST\archive.pst";
        //private static string pstPath = @"C:\Users\kandi\Desktop\SamplePST\ege.pst";
        static void Main(string[] args)
        {
            MailHelper mailHelper = new MailHelper(pstPath, path);
        }
    }
}

