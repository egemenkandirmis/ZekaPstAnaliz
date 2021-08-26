
namespace MailBackupForm
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.btnOpenPst = new System.Windows.Forms.Button();
            this.btnSelectFolder = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.btnAyikla = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.btnExcel = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btnMsg = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tbTarih = new System.Windows.Forms.TextBox();
            this.tbKonu = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tbAlici = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tbCc = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tbBcc = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.tbGonderen = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.richTextBoxİcerik = new System.Windows.Forms.RichTextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(168, 61);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(242, 20);
            this.textBox1.TabIndex = 1;
            // 
            // btnOpenPst
            // 
            this.btnOpenPst.Location = new System.Drawing.Point(67, 61);
            this.btnOpenPst.Name = "btnOpenPst";
            this.btnOpenPst.Size = new System.Drawing.Size(75, 23);
            this.btnOpenPst.TabIndex = 0;
            this.btnOpenPst.Text = "Gözat";
            this.btnOpenPst.UseVisualStyleBackColor = true;
            this.btnOpenPst.Click += new System.EventHandler(this.btnOpenPst_Click);
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Location = new System.Drawing.Point(67, 108);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(75, 23);
            this.btnSelectFolder.TabIndex = 2;
            this.btnSelectFolder.Text = "Klasör Seç";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            this.btnSelectFolder.Click += new System.EventHandler(this.btnSelectFolder_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(168, 110);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(242, 20);
            this.textBox2.TabIndex = 3;
            // 
            // btnAyikla
            // 
            this.btnAyikla.Location = new System.Drawing.Point(67, 158);
            this.btnAyikla.Name = "btnAyikla";
            this.btnAyikla.Size = new System.Drawing.Size(71, 23);
            this.btnAyikla.TabIndex = 4;
            this.btnAyikla.Text = "AYIKLA";
            this.btnAyikla.UseVisualStyleBackColor = true;
            this.btnAyikla.Click += new System.EventHandler(this.btnAyikla_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(64, 45);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = ".pst Dosyası Seç";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(64, 92);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(90, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Kayıt Klasörü Seç";
            // 
            // listBox1
            // 
            this.listBox1.AllowDrop = true;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(168, 158);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(242, 290);
            this.listBox1.TabIndex = 8;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // btnExcel
            // 
            this.btnExcel.Location = new System.Drawing.Point(168, 454);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(98, 37);
            this.btnExcel.TabIndex = 10;
            this.btnExcel.Text = "Excel Rapor Göster";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(477, 45);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(91, 13);
            this.label3.TabIndex = 12;
            this.label3.Text = ".msg Dosyası Seç";
            // 
            // btnMsg
            // 
            this.btnMsg.Location = new System.Drawing.Point(480, 61);
            this.btnMsg.Name = "btnMsg";
            this.btnMsg.Size = new System.Drawing.Size(75, 23);
            this.btnMsg.TabIndex = 11;
            this.btnMsg.Text = "Gözat";
            this.btnMsg.UseVisualStyleBackColor = true;
            this.btnMsg.Click += new System.EventHandler(this.btnMsg_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.richTextBoxİcerik);
            this.groupBox1.Controls.Add(this.tbGonderen);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.tbBcc);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.tbCc);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.tbAlici);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.tbKonu);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.tbTarih);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Location = new System.Drawing.Point(583, 61);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(408, 387);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = ".msg Dosyası";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 53);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(34, 13);
            this.label4.TabIndex = 14;
            this.label4.Text = "Tarih:";
            // 
            // tbTarih
            // 
            this.tbTarih.Location = new System.Drawing.Point(100, 50);
            this.tbTarih.Name = "tbTarih";
            this.tbTarih.ReadOnly = true;
            this.tbTarih.Size = new System.Drawing.Size(242, 20);
            this.tbTarih.TabIndex = 14;
            // 
            // tbKonu
            // 
            this.tbKonu.Location = new System.Drawing.Point(100, 76);
            this.tbKonu.Name = "tbKonu";
            this.tbKonu.ReadOnly = true;
            this.tbKonu.Size = new System.Drawing.Size(242, 20);
            this.tbKonu.TabIndex = 15;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 79);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(35, 13);
            this.label5.TabIndex = 16;
            this.label5.Text = "Konu:";
            // 
            // tbAlici
            // 
            this.tbAlici.Location = new System.Drawing.Point(100, 102);
            this.tbAlici.Name = "tbAlici";
            this.tbAlici.ReadOnly = true;
            this.tbAlici.Size = new System.Drawing.Size(242, 20);
            this.tbAlici.TabIndex = 17;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(8, 105);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(29, 13);
            this.label6.TabIndex = 18;
            this.label6.Text = "Alıcı:";
            // 
            // tbCc
            // 
            this.tbCc.Location = new System.Drawing.Point(100, 128);
            this.tbCc.Name = "tbCc";
            this.tbCc.ReadOnly = true;
            this.tbCc.Size = new System.Drawing.Size(242, 20);
            this.tbCc.TabIndex = 19;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(8, 131);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(23, 13);
            this.label7.TabIndex = 20;
            this.label7.Text = "Cc:";
            // 
            // tbBcc
            // 
            this.tbBcc.Location = new System.Drawing.Point(100, 154);
            this.tbBcc.Name = "tbBcc";
            this.tbBcc.ReadOnly = true;
            this.tbBcc.Size = new System.Drawing.Size(242, 20);
            this.tbBcc.TabIndex = 21;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(8, 157);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(29, 13);
            this.label8.TabIndex = 22;
            this.label8.Text = "Bcc:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(8, 183);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(36, 13);
            this.label9.TabIndex = 24;
            this.label9.Text = "İçerik:";
            // 
            // tbGonderen
            // 
            this.tbGonderen.Location = new System.Drawing.Point(100, 24);
            this.tbGonderen.Name = "tbGonderen";
            this.tbGonderen.ReadOnly = true;
            this.tbGonderen.Size = new System.Drawing.Size(242, 20);
            this.tbGonderen.TabIndex = 27;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(8, 27);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(57, 13);
            this.label10.TabIndex = 28;
            this.label10.Text = "Gönderen:";
            // 
            // richTextBoxİcerik
            // 
            this.richTextBoxİcerik.Location = new System.Drawing.Point(100, 183);
            this.richTextBoxİcerik.Name = "richTextBoxİcerik";
            this.richTextBoxİcerik.ReadOnly = true;
            this.richTextBoxİcerik.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.richTextBoxİcerik.Size = new System.Drawing.Size(242, 188);
            this.richTextBoxİcerik.TabIndex = 29;
            this.richTextBoxİcerik.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1021, 524);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnMsg);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnAyikla);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.btnSelectFolder);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.btnOpenPst);
            this.Name = "Form1";
            this.Text = "MailBackup";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.Button btnOpenPst;
        private System.Windows.Forms.Button btnSelectFolder;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Button btnAyikla;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox listBox1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnMsg;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RichTextBox richTextBoxİcerik;
        private System.Windows.Forms.TextBox tbGonderen;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox tbBcc;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox tbCc;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox tbAlici;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbKonu;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbTarih;
        private System.Windows.Forms.Label label4;
    }
}

