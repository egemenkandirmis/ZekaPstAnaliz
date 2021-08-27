using Spire.Email.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;
using MailBackupForm.Properties;

namespace MailBackupForm.Helpers
{
    public class MailHelper
    {
        private OutlookFile outlookFile;
        public List<string> icerik = new List<string>();
        string txtPath = @"";


        public MailHelper(string pstPath, string recordPath, string txtPath)
        {
            Directory.CreateDirectory(txtPath);
            this.txtPath = txtPath;
            outlookFile = new OutlookFile(pstPath);
            var firstSubFolder = outlookFile.GetSubFolders().Where(c => c.ContainerClass == "" || c.ContainerClass == "IPF.Note").ToList();
            var x = outlookFile.Folders.ToList();
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
                ValidFolder(txtPath);
                MailYazdir(mail, txtPath, pathGov);
                MailKaydet(mail, pathGov);
                return true;
            }
            if (!contains1.ToLower().Contains(contains2))
            {
                var pathGov = recordFolderName.Replace("$$$$$", "diger");
                ValidFolder(txtPath);
                ValidFolder(pathGov);
                MailYazdir(mail, txtPath, pathGov);
                try
                {
                    MailKaydet(mail, pathGov);

                }
                catch (Exception e)
                {

                    Console.WriteLine(e.Message); ;
                }

                return true;
            }

            return false;

        }

        public void SaveOtherMails(OutlookItem mail, string recordFolderName, string replaceWith)
        {
            var path = recordFolderName.Replace("$$$$$", replaceWith);
            ValidFolder(path);
            ValidFolder(txtPath);
            MailKaydet(mail, path);
        }

        public void MailYazdir(OutlookItem outlookItem, string txtPath, string recordFolderName)
        {
            Mail mail = new Mail
            {
                Body = outlookItem.Body,
                DeliveryTime = outlookItem.DeliveryTime,
                DisplayBcc = outlookItem.DisplayBcc,
                DisplayCc = outlookItem.DisplayCc,
                DisplayTo = outlookItem.DisplayTo,
                SenderEmailAddress = outlookItem.SenderEmailAddress,
                Subject = outlookItem.Subject,
                FilePath = @"\\"+recordFolderName + @"\" + outlookItem.GetHashCode()+".msg"
            };

            var str = Newtonsoft.Json.JsonConvert.SerializeObject(mail);

            //File.Create(txtPath + @"output.txt");


            //File.AppendAllText(txtPath + @"\output.txt", "[");
         

            // serialize JSON to a string and then write string to a file
            //File.AppendAllText(txtPath + @"\output.txt", JsonConvert.SerializeObject(mail));

            // serialize JSON directly to a file
            using (StreamWriter file = File.AppendText(txtPath + @"\output.txt"))
            {
                file.WriteLine(str);
                //File.AppendAllText(txtPath + @"\output.txt", str + ",");
            }
        }

        public void MailKaydet(OutlookItem outlookItem, string recordFolderName)
        {

            try
            {

                outlookItem.SaveToStream(recordFolderName + @"\" + outlookItem.GetHashCode() + ".msg");
                
            }

            catch (Exception e)
            {

                Console.WriteLine(e.Message);
            }


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
