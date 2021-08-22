using Spire.Email.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;


namespace MailBackupSystem
{
    public class MailHelper
    {
        private OutlookFile outlookFile;
        private static string excelPath = @"C:\Users\kandi\Desktop\SamplePST\ExcelDocs\";
        ExcelHelper excelHelper;
        private static int rowIndex = 1;
        public List<string> icerik = new List<string>();


        public MailHelper(string pstPath, string recordPath)
        {
            excelHelper = new ExcelHelper();
            outlookFile = new OutlookFile(pstPath);
            var firstSubFolder = outlookFile.GetSubFolders().Where(c => c.ContainerClass == "").ToList();
            foreach (var sf in firstSubFolder)
            {
                //if (sf.Name.ToLower().Contains("gelen"))
                //  continue;
                Islem(sf, Path.Combine(recordPath, sf.Name));
            }
        }


        public void Islem(OutlookFolder folder, string recordFolderName)
        {
            Console.WriteLine(folder.Name);

            var klasorList = KlasorAl(folder);
            MailAl(folder, recordFolderName);
            if (klasorList != null)
            {
                foreach (var item in klasorList)
                {
                    Islem(item, Path.Combine(recordFolderName, folder.Name));
                }
            }
        }

        public List<OutlookFolder> KlasorAl(OutlookFolder folder)
        {
            return folder.GetSubFolders().Select(c => c).ToList();

        }

        public void MailAl(OutlookFolder folder, string recordFolderName)
        {

            foreach (var mail in folder.EnumerateOutlookItem())
            {
                var path = Path.Combine(recordFolderName, mail.GetHashCode().ToString());

                if (mail.DisplayTo == null)
                    continue;
                SeperateMail(mail, recordFolderName, mail.DisplayTo, "yahya.kisioglu@zafer.gov", "gov");
                SeperateMail(mail, recordFolderName, mail.DisplayTo, "yahya.kisioglu@zafer.org", "org");


                if (mail.DisplayCc == null)
                    continue;
                SeperateMail(mail, recordFolderName, mail.DisplayCc, "yahya.kisioglu@zafer.gov", "govCc");
                SeperateMail(mail, recordFolderName, mail.DisplayCc, "yahya.kisioglu@zafer.org", "orgCc");


                if (mail.DisplayBcc == null)
                    continue;
                SeperateMail(mail, recordFolderName, mail.DisplayBcc, "yahya.kisioglu@zafer.gov", "govBcc");
                SeperateMail(mail, recordFolderName, mail.DisplayBcc, "yahya.kisioglu@zafer.org", "orgBcc");


                if (mail.SenderEmailAddress == null)
                    continue;
                SeperateMail(mail, recordFolderName, mail.SenderEmailAddress, "yahya.kisioglu@zafer.gov", "govGonderilen");
                SeperateMail(mail, recordFolderName, mail.SenderEmailAddress, "yahya.kisioglu@zafer.org", "orgGonderilen");


            }
        }

        public bool SeperateMail(OutlookItem mail, string recordFolderName, string contains1, string contains2, string replaceWith)
        {

            if (contains1.ToLower().Contains(contains2))
            {
                var pathGov = recordFolderName.Replace("$$$$$", replaceWith);
                ValidFolder(pathGov);
                // MailYazdir(mail, path);
                MailKaydet(mail, pathGov);
                //  EklentiKaydet(mail, path);
                return true;
            }
            if (!contains1.ToLower().Contains(contains2))
            {
                var pathGov = recordFolderName.Replace("$$$$$", "diger");
                ValidFolder(pathGov);
                // MailYazdir(mail, path);
                try
                {
                    MailKaydet(mail, pathGov);

                }
                catch (Exception e)
                {

                    Console.WriteLine(e.Message); ;
                }
                //  EklentiKaydet(mail, path);
                return true;
            }

            return false;

        }

        public void SaveOtherMails(OutlookItem mail, string recordFolderName, string replaceWith)
        {
            var path = recordFolderName.Replace("$$$$$", replaceWith);
            ValidFolder(path);
            MailKaydet(mail, path);
        }


        public void MailKaydet(OutlookItem outlookItem, string recordFolderName)
        {
            if (outlookItem.DeliveryTime != null)
                icerik.Add(outlookItem.DeliveryTime.ToString());
            else
            {
                icerik.Add("N/A");
            }

            if (outlookItem.SenderEmailAddress != null)
                icerik.Add(outlookItem.SenderEmailAddress.ToString());
            else
            {
                icerik.Add("N/A");
            }

            if (outlookItem.DisplayTo != null)
                icerik.Add(outlookItem.DisplayTo.ToString());
            else
            {
                icerik.Add("N/A");
            }

            if (outlookItem.DisplayCc != null)
                icerik.Add(outlookItem.DisplayCc.ToString());
            else
            {
                icerik.Add("N/A");
            }

            if (outlookItem.DisplayBcc != null)
                icerik.Add(outlookItem.DisplayBcc.ToString());
            else
            {
                icerik.Add("N/A");
            }

            if (outlookItem.Subject != null)
                icerik.Add(outlookItem.Subject.ToString());
            else
            {
                icerik.Add("N/A");
            }

            if (outlookItem.Body != null)
                icerik.Add(outlookItem.Body.ToString());
            else
            {
                icerik.Add("N/A");
            }

            if (outlookItem != null)
                icerik.Add(recordFolderName + @"\" + outlookItem.GetHashCode());
            else
            {
                icerik.Add("N/A");
            }

            try
            {
                outlookItem.SaveToStream(recordFolderName + @"\" + outlookItem.GetHashCode() + ".msg");
            }

            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }
            excelHelper.CreateData(icerik, rowIndex);
            excelHelper.WriteExcel(excelPath);
            rowIndex++;
            icerik.Clear();

        }

        public void EklentiKaydet(OutlookItem outlookItem, string recordFolderName)
        {
            if (outlookItem.Attachments.Count > 0)
            {
                foreach (var i in outlookItem.Attachments)
                {
                    if (i.Extension == null || i.FileName == null)
                        continue;
                    i.SaveToFile(recordFolderName + @"\" + i.FileName + i.Extension);
                }
            }
        }

        public void ValidFolder(string folderName)
        {
            if (!Directory.Exists(folderName))
                Directory.CreateDirectory(folderName);

        }

    }
}
