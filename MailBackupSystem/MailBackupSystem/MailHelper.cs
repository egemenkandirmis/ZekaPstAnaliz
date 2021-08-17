﻿using Spire.Email.Outlook;
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


        public MailHelper(string pstPath, string recordPath)
        {
            outlookFile = new OutlookFile(pstPath);
            var firstSubFolder = outlookFile.GetSubFolders().Where(c => c.ContainerClass == "").ToList();
            foreach (var sf in firstSubFolder)
            {
                Islem(sf, Path.Combine(recordPath, sf.Name));
            }
        }

        public void Islem(OutlookFolder folder, string recordFolderName)
        {
            
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
           
                ValidFolder(recordFolderName);
               // MailYazdir(mail, path);
                MailKaydet(mail, recordFolderName);
              //  EklentiKaydet(mail, path);
            }
        }

        public void MailYazdir(OutlookItem outlookItem, string recordFolderName)
        {
            var icerik = $@"{outlookItem.DeliveryTime}ß{outlookItem.SenderEmailAddress}ß{outlookItem.Subject}ß{outlookItem.DisplayName}ß{outlookItem.DisplayTo}ß{outlookItem.DisplayBcc}ß{outlookItem.DisplayCc}ß{outlookItem.Body}";
            
            using (StreamWriter sw = File.AppendText(recordFolderName + @"\content.txt"))
            {
                sw.WriteLine(icerik);
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