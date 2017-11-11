namespace OptimaToPP
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
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.button1 = new System.Windows.Forms.Button();
            this.XLSpath = new System.Windows.Forms.TextBox();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.btnConvert = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(12, 541);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(105, 17);
            this.linkLabel1.TabIndex = 6;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Damian Suwała";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnConvert);
            this.groupBox1.Controls.Add(this.btnOpenFile);
            this.groupBox1.Controls.Add(this.XLSpath);
            this.groupBox1.Location = new System.Drawing.Point(13, 26);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(448, 127);
            this.groupBox1.TabIndex = 7;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Wskaż plik XLS";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.button1);
            this.groupBox2.Location = new System.Drawing.Point(18, 159);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(448, 80);
            this.groupBox2.TabIndex = 8;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Zapisz jako XML dla Poczty Polskiej";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(6, 21);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(423, 41);
            this.button1.TabIndex = 0;
            this.button1.Text = "Eksportuj plik XLM";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // XLSpath
            // 
            this.XLSpath.Location = new System.Drawing.Point(149, 30);
            this.XLSpath.Name = "XLSpath";
            this.XLSpath.Size = new System.Drawing.Size(280, 22);
            this.XLSpath.TabIndex = 0;
            // 
            // btnOpenFile
            // 
            this.btnOpenFile.Location = new System.Drawing.Point(6, 21);
            this.btnOpenFile.Name = "btnOpenFile";
            this.btnOpenFile.Size = new System.Drawing.Size(120, 41);
            this.btnOpenFile.TabIndex = 1;
            this.btnOpenFile.Text = "Otwórz plik";
            this.btnOpenFile.UseVisualStyleBackColor = true;
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(6, 68);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(423, 41);
            this.btnConvert.TabIndex = 2;
            this.btnConvert.Text = "Konwertuj";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(478, 252);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.linkLabel1);
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.LinkLabel linkLabel1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.TextBox XLSpath;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button button1;
    }
}

