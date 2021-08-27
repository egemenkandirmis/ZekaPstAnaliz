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
        private string rootPath = "";
        private string rootFolder = "";
        //private static string pstPath = @"C:\Users\kandi\Desktop\SamplePST\archive.pst";
        private static string pstPath = @"";
        //private static string txtPath = @"C:\Users\kandi\Desktop\SamplePST\TxtDocs\";
        private static string txtPath = @"";
        //private static string excelPath = @"C:\Users\kandi\Desktop\SamplePST\ExcelDocs\";
        private static string excelPath = @"";

        private static int indexCounter = 0;
        private static int mailCount;

        WebBrowser webView;


        DateTime start;
        DateTime end;
        public Form1()
        {
            InitializeComponent();
            btnExcel.Enabled = false;

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
            btnExcel.Enabled = false;
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
                rootPath = folderName + @"\Mail";
                rootFolder = folderName;
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
            webView = new WebBrowser();
            tableLayoutPanel2.Controls.Add(webView, 1, 6);
            webView.Dock = DockStyle.Fill;
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            start = DateTime.Now;
            backgroundWorker1.ReportProgress(0);
            Clear();
            Thread.Sleep(100);
            backgroundWorker1.ReportProgress(0);

            Console.WriteLine("Başlangıç: " + start.ToShortTimeString());


            MailHelper mailHelper = new MailHelper(pstPath, path, txtPath);
            Thread.Sleep(1000);
            backgroundWorker1.ReportProgress(0);

            ExcelHelper excelHelper = new ExcelHelper();
            excelHelper.ReadFromTxt(txtPath, excelPath);
            Thread.Sleep(1500);
            backgroundWorker1.ReportProgress(0);


            end = DateTime.Now;
            Thread.Sleep(1000);
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
                    listBox1.Items.Add("Klasörler hazırlanıyor...");
                    break;
                case 1:
                    listBox1.Items.Add("Başlama saati: " + start.ToShortTimeString());
                    listBox1.Items.Add("Mailler ayıklanıyor...");
                    break;
                case 2:
                    listBox1.Items.Add("Mailler ayıklandı.");
                    listBox1.Items.Add("Excel raporu oluşturuluyor...");
                    break;
                case 3:
                    Console.WriteLine("helllooooo");
                    Console.WriteLine(mailCount);
                    listBox1.Items.Add("Excel raporu oluşturuldu.");
                    listBox1.Items.Add("Toplam mail sayısı: " + mailCount.ToString());
                    listBox1.Items.Add("Bitiş saati: " + start.ToShortTimeString());
                    var fark = end.Subtract(start).TotalMinutes;
                    fark = Math.Round(fark, 2);
                    listBox1.Items.Add("Geçen süre: " + fark + " dakika");
                    break;
                default:
                    break;
            }
        }

        public void SetTotalMailCount(int count)
        {
            Console.WriteLine("TOTAL MAİL: " + count);
            mailCount = count;
            Console.WriteLine("TOTAL MAİL: " + mailCount);
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

            ListDirectory(treeView1, rootPath);




            //OutlookItem outlookItem = new OutlookItem();
            //openFileDialog1.ShowDialog();
            //var a = openFileDialog1.FileName;
            //outlookItem.LoadFromFile(a);
            //tbGonderen.Text = outlookItem.SenderEmailAddress;
            //tbAlici.Text = outlookItem.DisplayTo;
            //tbTarih.Text = outlookItem.DeliveryTime.ToString();
            //tbBcc.Text = outlookItem.DisplayBcc;
            //tbCc.Text = outlookItem.DisplayCc;
            //tbKonu.Text = outlookItem.Subject;
            //richTextBoxİcerik.Text = outlookItem.Body;
        }



        private void ListDirectory(TreeView treeView, string path)
        {
            if (path == "")
            {
                MessageBox.Show("Lütfen Kayıt Yolunu Seçiniz.");
                return;
            }

            treeView.Nodes.Clear();
            var rootDirectoryInfo = new DirectoryInfo(path);
            treeView.Nodes.Add(CreateDirectoryNode(rootDirectoryInfo));
        }

        private static TreeNode CreateDirectoryNode(DirectoryInfo directoryInfo)
        {
            var directoryNode = new TreeNode(directoryInfo.Name);
            foreach (var directory in directoryInfo.GetDirectories())
                directoryNode.Nodes.Add(CreateDirectoryNode(directory));
            foreach (var file in directoryInfo.GetFiles())
                directoryNode.Nodes.Add(new TreeNode(file.Name));
            return directoryNode;
        }

        private void treeView1_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            
            if (treeView1.SelectedNode == null)
            {
                return;
            }
            var txt = treeView1.SelectedNode.Text;
            if (txt.Length < 4)
            {
                return;
            }


            if (txt.Substring(txt.Length - 4, 4) != ".msg")
            {
                return;
            }


            OutlookItem outlookItem = new OutlookItem();

            var a = rootFolder + @"\" + treeView1.SelectedNode.FullPath;
            outlookItem.LoadFromFile(a);
            tbGonderen.Text = outlookItem.SenderEmailAddress;
            tbAlici.Text = outlookItem.DisplayTo;
            tbTarih.Text = outlookItem.DeliveryTime.ToString();
            tbBcc.Text = outlookItem.DisplayBcc;
            tbCc.Text = outlookItem.DisplayCc;
            tbKonu.Text = outlookItem.Subject;
           
            if (outlookItem.BodyHtml != null)
                webView.DocumentText = outlookItem.BodyHtml;
            if (outlookItem.BodyHtml == null)
                webView.DocumentText = "Mail İçeriği Görüntülenemiyor";



        }

        private void tableLayoutPanel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
