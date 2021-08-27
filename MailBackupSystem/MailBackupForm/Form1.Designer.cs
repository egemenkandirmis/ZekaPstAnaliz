
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
            this.btnMsg = new System.Windows.Forms.Button();
            this.tbGonderen = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.tbBcc = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.tbCc = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tbAlici = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tbKonu = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tbTarih = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel3 = new System.Windows.Forms.TableLayoutPanel();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.tableLayoutPanel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.White;
            this.textBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.textBox1.Location = new System.Drawing.Point(209, 3);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(157, 23);
            this.textBox1.TabIndex = 1;
            // 
            // btnOpenPst
            // 
            this.btnOpenPst.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnOpenPst.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnOpenPst.Location = new System.Drawing.Point(128, 3);
            this.btnOpenPst.Name = "btnOpenPst";
            this.btnOpenPst.Size = new System.Drawing.Size(75, 23);
            this.btnOpenPst.TabIndex = 0;
            this.btnOpenPst.Text = "Gözat";
            this.btnOpenPst.UseVisualStyleBackColor = true;
            this.btnOpenPst.Click += new System.EventHandler(this.btnOpenPst_Click);
            // 
            // btnSelectFolder
            // 
            this.btnSelectFolder.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnSelectFolder.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnSelectFolder.Location = new System.Drawing.Point(128, 32);
            this.btnSelectFolder.Name = "btnSelectFolder";
            this.btnSelectFolder.Size = new System.Drawing.Size(75, 23);
            this.btnSelectFolder.TabIndex = 2;
            this.btnSelectFolder.Text = "Klasör Seç";
            this.btnSelectFolder.UseVisualStyleBackColor = true;
            this.btnSelectFolder.Click += new System.EventHandler(this.btnSelectFolder_Click);
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.Color.White;
            this.textBox2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.textBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.textBox2.Location = new System.Drawing.Point(209, 32);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(157, 23);
            this.textBox2.TabIndex = 3;
            // 
            // btnAyikla
            // 
            this.btnAyikla.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnAyikla.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnAyikla.Location = new System.Drawing.Point(128, 61);
            this.btnAyikla.Name = "btnAyikla";
            this.btnAyikla.Size = new System.Drawing.Size(75, 23);
            this.btnAyikla.TabIndex = 4;
            this.btnAyikla.Text = "Ayıkla";
            this.btnAyikla.UseVisualStyleBackColor = true;
            this.btnAyikla.Click += new System.EventHandler(this.btnAyikla_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(119, 17);
            this.label1.TabIndex = 5;
            this.label1.Text = "*.Pst Dosyası Seç";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label2.Location = new System.Drawing.Point(3, 29);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(119, 17);
            this.label2.TabIndex = 6;
            this.label2.Text = "Kayıt Klasörü Seç";
            // 
            // listBox1
            // 
            this.listBox1.AllowDrop = true;
            this.listBox1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.HorizontalScrollbar = true;
            this.listBox1.ItemHeight = 16;
            this.listBox1.Location = new System.Drawing.Point(209, 90);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(157, 592);
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
            this.btnExcel.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnExcel.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnExcel.Location = new System.Drawing.Point(128, 90);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(75, 23);
            this.btnExcel.TabIndex = 10;
            this.btnExcel.Text = "Excel Rapor Göster";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // btnMsg
            // 
            this.btnMsg.Dock = System.Windows.Forms.DockStyle.Left;
            this.btnMsg.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.btnMsg.Location = new System.Drawing.Point(3, 3);
            this.btnMsg.Name = "btnMsg";
            this.btnMsg.Size = new System.Drawing.Size(75, 23);
            this.btnMsg.TabIndex = 11;
            this.btnMsg.Text = "Yenile";
            this.btnMsg.UseVisualStyleBackColor = true;
            this.btnMsg.Click += new System.EventHandler(this.btnMsg_Click);
            // 
            // tbGonderen
            // 
            this.tbGonderen.BackColor = System.Drawing.Color.White;
            this.tbGonderen.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbGonderen.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.tbGonderen.Location = new System.Drawing.Point(80, 3);
            this.tbGonderen.Name = "tbGonderen";
            this.tbGonderen.ReadOnly = true;
            this.tbGonderen.Size = new System.Drawing.Size(884, 23);
            this.tbGonderen.TabIndex = 27;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label10.Location = new System.Drawing.Point(3, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(71, 16);
            this.label10.TabIndex = 28;
            this.label10.Text = "Gönderen:";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label9.Location = new System.Drawing.Point(3, 120);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(43, 16);
            this.label9.TabIndex = 24;
            this.label9.Text = "İçerik:";
            // 
            // tbBcc
            // 
            this.tbBcc.BackColor = System.Drawing.Color.White;
            this.tbBcc.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbBcc.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.tbBcc.Location = new System.Drawing.Point(80, 103);
            this.tbBcc.Name = "tbBcc";
            this.tbBcc.ReadOnly = true;
            this.tbBcc.Size = new System.Drawing.Size(884, 23);
            this.tbBcc.TabIndex = 21;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label8.Location = new System.Drawing.Point(3, 100);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(34, 16);
            this.label8.TabIndex = 22;
            this.label8.Text = "Bcc:";
            // 
            // tbCc
            // 
            this.tbCc.BackColor = System.Drawing.Color.White;
            this.tbCc.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbCc.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.tbCc.Location = new System.Drawing.Point(80, 83);
            this.tbCc.Name = "tbCc";
            this.tbCc.ReadOnly = true;
            this.tbCc.Size = new System.Drawing.Size(884, 23);
            this.tbCc.TabIndex = 19;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label7.Location = new System.Drawing.Point(3, 80);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(27, 16);
            this.label7.TabIndex = 20;
            this.label7.Text = "Cc:";
            // 
            // tbAlici
            // 
            this.tbAlici.BackColor = System.Drawing.Color.White;
            this.tbAlici.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbAlici.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.tbAlici.Location = new System.Drawing.Point(80, 63);
            this.tbAlici.Name = "tbAlici";
            this.tbAlici.ReadOnly = true;
            this.tbAlici.Size = new System.Drawing.Size(884, 23);
            this.tbAlici.TabIndex = 17;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label6.Location = new System.Drawing.Point(3, 60);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(36, 16);
            this.label6.TabIndex = 18;
            this.label6.Text = "Alıcı:";
            // 
            // tbKonu
            // 
            this.tbKonu.BackColor = System.Drawing.Color.White;
            this.tbKonu.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbKonu.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.tbKonu.Location = new System.Drawing.Point(80, 43);
            this.tbKonu.Name = "tbKonu";
            this.tbKonu.ReadOnly = true;
            this.tbKonu.Size = new System.Drawing.Size(884, 23);
            this.tbKonu.TabIndex = 15;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label5.Location = new System.Drawing.Point(3, 40);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(41, 16);
            this.label5.TabIndex = 16;
            this.label5.Text = "Konu:";
            // 
            // tbTarih
            // 
            this.tbTarih.BackColor = System.Drawing.Color.White;
            this.tbTarih.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbTarih.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.tbTarih.Location = new System.Drawing.Point(80, 23);
            this.tbTarih.Name = "tbTarih";
            this.tbTarih.ReadOnly = true;
            this.tbTarih.Size = new System.Drawing.Size(884, 23);
            this.tbTarih.TabIndex = 14;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.label4.Location = new System.Drawing.Point(3, 20);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(42, 16);
            this.label4.TabIndex = 14;
            this.label4.Text = "Tarih:";
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.btnOpenPst, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.textBox1, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.btnSelectFolder, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.textBox2, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.btnAyikla, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.btnExcel, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.listBox1, 2, 3);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Left;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.Size = new System.Drawing.Size(369, 685);
            this.tableLayoutPanel1.TabIndex = 15;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 2;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel2.Controls.Add(this.label10, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.tbGonderen, 1, 0);
            this.tableLayoutPanel2.Controls.Add(this.label4, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.tbBcc, 1, 5);
            this.tableLayoutPanel2.Controls.Add(this.label5, 0, 2);
            this.tableLayoutPanel2.Controls.Add(this.tbCc, 1, 4);
            this.tableLayoutPanel2.Controls.Add(this.label6, 0, 3);
            this.tableLayoutPanel2.Controls.Add(this.tbAlici, 1, 3);
            this.tableLayoutPanel2.Controls.Add(this.label9, 0, 6);
            this.tableLayoutPanel2.Controls.Add(this.tbKonu, 1, 2);
            this.tableLayoutPanel2.Controls.Add(this.tbTarih, 1, 1);
            this.tableLayoutPanel2.Controls.Add(this.label7, 0, 4);
            this.tableLayoutPanel2.Controls.Add(this.label8, 0, 5);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(369, 221);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 7;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(967, 464);
            this.tableLayoutPanel2.TabIndex = 16;
            this.tableLayoutPanel2.Paint += new System.Windows.Forms.PaintEventHandler(this.tableLayoutPanel2_Paint);
            // 
            // tableLayoutPanel3
            // 
            this.tableLayoutPanel3.ColumnCount = 1;
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel3.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel3.Controls.Add(this.treeView1, 0, 1);
            this.tableLayoutPanel3.Controls.Add(this.btnMsg, 0, 0);
            this.tableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Top;
            this.tableLayoutPanel3.Location = new System.Drawing.Point(369, 0);
            this.tableLayoutPanel3.Name = "tableLayoutPanel3";
            this.tableLayoutPanel3.RowCount = 2;
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel3.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel3.Size = new System.Drawing.Size(967, 221);
            this.tableLayoutPanel3.TabIndex = 17;
            // 
            // treeView1
            // 
            this.treeView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeView1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(162)));
            this.treeView1.Location = new System.Drawing.Point(3, 32);
            this.treeView1.Name = "treeView1";
            this.treeView1.Size = new System.Drawing.Size(961, 186);
            this.treeView1.TabIndex = 13;
            this.treeView1.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeView1_NodeMouseDoubleClick);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1336, 685);
            this.Controls.Add(this.tableLayoutPanel2);
            this.Controls.Add(this.tableLayoutPanel3);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "Form1";
            this.Text = "MailBackup";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.tableLayoutPanel3.ResumeLayout(false);
            this.ResumeLayout(false);

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
        private System.Windows.Forms.Button btnMsg;
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
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel3;
        private System.Windows.Forms.TreeView treeView1;
    }
}

