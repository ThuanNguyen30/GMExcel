namespace GMExcel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.dapanfile = new System.Windows.Forms.OpenFileDialog();
            this.txtdapan = new System.Windows.Forms.TextBox();
            this.Browse = new System.Windows.Forms.Button();
            this.btnbrowsefo = new System.Windows.Forms.Button();
            this.txtthumuccham = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.foldercham = new System.Windows.Forms.FolderBrowserDialog();
            this.btnthc = new System.Windows.Forms.Button();
            this.btnthoat = new System.Windows.Forms.Button();
            this.txttieude = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btbbrowsekq = new System.Windows.Forms.Button();
            this.txtthumucketqua = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtdiemsapxep = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txttenlop = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txtdapan
            // 
            this.txtdapan.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtdapan.Location = new System.Drawing.Point(43, 25);
            this.txtdapan.Name = "txtdapan";
            this.txtdapan.Size = new System.Drawing.Size(544, 22);
            this.txtdapan.TabIndex = 0;
            this.txtdapan.Text = "C:\\Users\\QUANGTHUAN\\Desktop\\GMExcel\\DAPAN\\DAPAN.XLSX";
            // 
            // Browse
            // 
            this.Browse.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Browse.Location = new System.Drawing.Point(593, 24);
            this.Browse.Name = "Browse";
            this.Browse.Size = new System.Drawing.Size(75, 23);
            this.Browse.TabIndex = 1;
            this.Browse.Text = "Browse";
            this.Browse.UseVisualStyleBackColor = true;
            this.Browse.Click += new System.EventHandler(this.Browse_Click);
            // 
            // btnbrowsefo
            // 
            this.btnbrowsefo.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnbrowsefo.Location = new System.Drawing.Point(593, 80);
            this.btnbrowsefo.Name = "btnbrowsefo";
            this.btnbrowsefo.Size = new System.Drawing.Size(75, 23);
            this.btnbrowsefo.TabIndex = 3;
            this.btnbrowsefo.Text = "Browse";
            this.btnbrowsefo.UseVisualStyleBackColor = true;
            this.btnbrowsefo.Click += new System.EventHandler(this.btnbrowsefo_Click);
            // 
            // txtthumuccham
            // 
            this.txtthumuccham.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtthumuccham.Location = new System.Drawing.Point(43, 80);
            this.txtthumuccham.Name = "txtthumuccham";
            this.txtthumuccham.Size = new System.Drawing.Size(544, 22);
            this.txtthumuccham.TabIndex = 2;
            this.txtthumuccham.Text = "C:\\Users\\QUANGTHUAN\\Desktop\\GMExcel\\CHAM";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(43, 3);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 16);
            this.label1.TabIndex = 4;
            this.label1.Text = "File excel đáp án";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(39, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(149, 16);
            this.label2.TabIndex = 5;
            this.label2.Text = "Thư mục chứa bài chấm";
            // 
            // btnthc
            // 
            this.btnthc.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnthc.Location = new System.Drawing.Point(174, 235);
            this.btnthc.Name = "btnthc";
            this.btnthc.Size = new System.Drawing.Size(152, 42);
            this.btnthc.TabIndex = 6;
            this.btnthc.Text = "THỰC HIỆN CHẤM";
            this.btnthc.UseVisualStyleBackColor = true;
            this.btnthc.Click += new System.EventHandler(this.btnthc_Click);
            // 
            // btnthoat
            // 
            this.btnthoat.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnthoat.Location = new System.Drawing.Point(359, 235);
            this.btnthoat.Name = "btnthoat";
            this.btnthoat.Size = new System.Drawing.Size(152, 42);
            this.btnthoat.TabIndex = 7;
            this.btnthoat.Text = "THOÁT";
            this.btnthoat.UseVisualStyleBackColor = true;
            this.btnthoat.Click += new System.EventHandler(this.btnthoat_Click);
            // 
            // txttieude
            // 
            this.txttieude.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txttieude.Location = new System.Drawing.Point(45, 187);
            this.txttieude.Name = "txttieude";
            this.txttieude.Size = new System.Drawing.Size(143, 22);
            this.txttieude.TabIndex = 9;
            this.txttieude.Text = "STT";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(39, 168);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(156, 16);
            this.label3.TabIndex = 10;
            this.label3.Text = "Giá trị đầu tiên cột tiêu đề";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(42, 112);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(158, 16);
            this.label6.TabIndex = 14;
            this.label6.Text = "Thư mục chứa file kết quả";
            // 
            // btbbrowsekq
            // 
            this.btbbrowsekq.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btbbrowsekq.Location = new System.Drawing.Point(593, 132);
            this.btbbrowsekq.Name = "btbbrowsekq";
            this.btbbrowsekq.Size = new System.Drawing.Size(75, 23);
            this.btbbrowsekq.TabIndex = 13;
            this.btbbrowsekq.Text = "Browse";
            this.btbbrowsekq.UseVisualStyleBackColor = true;
            this.btbbrowsekq.Click += new System.EventHandler(this.btbbrowsekq_Click);
            // 
            // txtthumucketqua
            // 
            this.txtthumucketqua.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtthumucketqua.Location = new System.Drawing.Point(43, 132);
            this.txtthumucketqua.Name = "txtthumucketqua";
            this.txtthumucketqua.Size = new System.Drawing.Size(544, 22);
            this.txtthumucketqua.TabIndex = 12;
            this.txtthumucketqua.Text = "C:\\Users\\QUANGTHUAN\\Desktop\\GMExcel\\KETQUA";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(249, 168);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(132, 16);
            this.label4.TabIndex = 16;
            this.label4.Text = "Điểm sắp xếp dữ liệu";
            // 
            // txtdiemsapxep
            // 
            this.txtdiemsapxep.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtdiemsapxep.Location = new System.Drawing.Point(252, 187);
            this.txtdiemsapxep.Name = "txtdiemsapxep";
            this.txtdiemsapxep.Size = new System.Drawing.Size(128, 22);
            this.txtdiemsapxep.TabIndex = 9;
            this.txtdiemsapxep.Text = "2";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(422, 168);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(99, 16);
            this.label5.TabIndex = 18;
            this.label5.Text = "Tên file kết quả";
            // 
            // txttenlop
            // 
            this.txttenlop.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txttenlop.Location = new System.Drawing.Point(425, 187);
            this.txttenlop.Name = "txttenlop";
            this.txttenlop.Size = new System.Drawing.Size(162, 22);
            this.txttenlop.TabIndex = 17;
            this.txttenlop.Text = "TEN_LOP.CSV";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(680, 298);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txttenlop);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.btbbrowsekq);
            this.Controls.Add(this.txtthumucketqua);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtdiemsapxep);
            this.Controls.Add(this.txttieude);
            this.Controls.Add(this.btnthoat);
            this.Controls.Add(this.btnthc);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnbrowsefo);
            this.Controls.Add(this.txtthumuccham);
            this.Controls.Add(this.Browse);
            this.Controls.Add(this.txtdapan);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "GMExcel Tool v1.0";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog dapanfile;
        private System.Windows.Forms.TextBox txtdapan;
        private System.Windows.Forms.Button Browse;
        private System.Windows.Forms.Button btnbrowsefo;
        private System.Windows.Forms.TextBox txtthumuccham;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.FolderBrowserDialog foldercham;
        private System.Windows.Forms.Button btnthc;
        private System.Windows.Forms.Button btnthoat;
		private System.Windows.Forms.TextBox txttieude;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Button btbbrowsekq;
		private System.Windows.Forms.TextBox txtthumucketqua;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtdiemsapxep;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txttenlop;
    }
}

