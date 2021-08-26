using MailBackupForm.Helpers;
using MailBackupForm.Properties;
using Spire.Email.Outlook;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MailBackupForm
{
    public partial class Form1 : Form
    {
        //private static string path = @"C:\Users\kandi\Desktop\SamplePST\Mail\$$$$$\";
        private static string path = @"";

        //private static string pstPath = @"C:\Users\kandi\Desktop\SamplePST\archive.pst";
        private static string pstPath = @"";
        //private static string txtPath = @"C:\Users\kandi\Desktop\SamplePST\TxtDocs\";
        private static string txtPath = @"";
        //private static string excelPath = @"C:\Users\kandi\Desktop\SamplePST\ExcelDocs\";
        private static string excelPath = @"";

        private static int indexCounter = 0;
        private int mailCount;


        DateTime start;
        DateTime end;
        public Form1()
        {
            InitializeComponent();
            //btnExcel.Enabled = false;
           
            if (path == "" || pstPath == "")
            {
                btnAyikla.Enabled = false;
            }
        }

        private void btnOpenPst_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            pstPath = openFileDialog1.FileName;
            textBox1.Text = pstPath;
            //btnExcel.Enabled = false;
        }

        private void btnSelectFolder_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            string folderName = folderBrowserDialog1.SelectedPath;
            textBox2.Text = folderName;
            DialogResult choice = MessageBox.Show("Seçtiğiniz klasör içerisine 'Gov','Org' ve 'Diğer' maillerinizin kaydedileceği 'Mail' klasörü,  'Mail' klasörü içerisine '.msg' dosyalarının detaylarının yer alacağı 'TxtDocs' klasörü ve excel raporunun yer alacağı 'ExcelDocs' klasörü oluşturulacaktır. Onaylıyor musunuz?", "Bilgilendirme Penceresi", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Information);
            if (choice == DialogResult.Yes)
            {
                var fullFolderName = folderName + @"\Mail\$$$$$";
                var fullTxtFolderName = folderName + @"\Mail\TxtDocs";
                var fullExcelFolderName = folderName + @"\Mail\ExcelDocs";
                path = fullFolderName;
                excelPath = fullExcelFolderName;
                txtPath = fullTxtFolderName;
                if (path != "" && pstPath != "")
                {
                    btnAyikla.Enabled = true;
                }
            }
            else if (choice == DialogResult.No)
            {
                //Hayır seçeneğine tıklandığında çalıştırılacak kodlar

                MessageBox.Show("İşlem iptal ediliyor.", "Bilgilendirme Penceresi", MessageBoxButtons.OK, MessageBoxIcon.Error);
                folderBrowserDialog1.SelectedPath = "";
                folderName = "";
                path = "";
                excelPath = "";
                txtPath = "";
            }

            textBox1.Text = pstPath;
            textBox2.Text = folderName;
        }

        private void btnAyikla_Click(object sender, EventArgs e)
        {
            backgroundWorker1.RunWorkerAsync();
            btnOpenPst.Enabled = false;
            btnSelectFolder.Enabled = false;
            btnAyikla.Enabled = false;

        }


        public void Clear()
        {
            if (Directory.Exists(excelPath))
                Directory.Delete(excelPath, true);


            if (Directory.Exists(path.Split('$')[0]))
                Directory.Delete(path.Split('$')[0], true);


            if (Directory.Exists(txtPath))
                Directory.Delete(txtPath, true);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            start = DateTime.Now;
            Clear();
            Thread.Sleep(100);
            backgroundWorker1.ReportProgress(0);

            Console.WriteLine("Başlangıç: " + start.ToString());


            MailHelper mailHelper = new MailHelper(pstPath, path, txtPath);
            Thread.Sleep(100);
            backgroundWorker1.ReportProgress(0);

            ExcelHelper excelHelper = new ExcelHelper();
            excelHelper.ReadFromTxt(txtPath, excelPath);
            Thread.Sleep(100);
            backgroundWorker1.ReportProgress(0);

            
            end = DateTime.Now;
            Thread.Sleep(100);
            backgroundWorker1.ReportProgress(0);
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            UpdateListBox(indexCounter);
            indexCounter++;
            
        }

        private void UpdateListBox(int index)
        {
            switch (index)
            {
                case 0:
                    listBox1.Items.Add("Başlama saati: " + start.ToString());
                    listBox1.Items.Add("Mailler ayıklanıyor..");
                    break;
                case 1:
                    listBox1.Items.Add("Mailler ayıklandı..");
                    listBox1.Items.Add("Excel raporu oluşturuluyor..");
                    break;
                case 2:
                    listBox1.Items.Add("Excel raporu oluşturuldu..");
                    listBox1.Items.Add("Toplam mail sayısı: " + mailCount);
                    listBox1.Items.Add("Bitiş saati: " + start.ToString());
                    var fark = end.Subtract(start).TotalMinutes;
                    listBox1.Items.Add("Geçen süre: " + fark + " dakika");
                    break;
                default:
                    break;
            }
        }

        public void GetTotalMailCount(int count)
        {
            mailCount = count;
        }


        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnExcel.Enabled = true;
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(excelPath + @"\MailReport.xlsx");
        }

        private void btnMsg_Click(object sender, EventArgs e)
        {
            OutlookItem outlookItem = new OutlookItem();
            openFileDialog1.ShowDialog();
            var a = openFileDialog1.FileName;
            outlookItem.LoadFromFile(a);
            tbGonderen.Text = outlookItem.SenderEmailAddress;
            tbAlici.Text = outlookItem.DisplayTo;
            tbTarih.Text = outlookItem.DeliveryTime.ToString();
            tbBcc.Text = outlookItem.DisplayBcc;
            tbCc.Text = outlookItem.DisplayCc;
            tbKonu.Text = outlookItem.Subject;
            richTextBoxİcerik.Text = outlookItem.Body;
        }
    }
}
