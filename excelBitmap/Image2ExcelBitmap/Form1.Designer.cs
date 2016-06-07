namespace Image2ExcelBitmap
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.ViewPicureBox = new System.Windows.Forms.PictureBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ImageFilePathTextBox = new System.Windows.Forms.TextBox();
            this.ExcelFilePathTextBox = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.StartButton = new System.Windows.Forms.Button();
            this.ProcessProgressBar = new System.Windows.Forms.ProgressBar();
            this.ProcessLable = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ImageInfoRichTextBox = new System.Windows.Forms.RichTextBox();
            ((System.ComponentModel.ISupportInitialize)(this.ViewPicureBox)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(276, 29);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(101, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "Image File Path:";
            // 
            // ViewPicureBox
            // 
            this.ViewPicureBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ViewPicureBox.Location = new System.Drawing.Point(3, 17);
            this.ViewPicureBox.Name = "ViewPicureBox";
            this.ViewPicureBox.Size = new System.Drawing.Size(252, 222);
            this.ViewPicureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.ViewPicureBox.TabIndex = 1;
            this.ViewPicureBox.TabStop = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.ViewPicureBox);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(258, 242);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "View";
            // 
            // ImageFilePathTextBox
            // 
            this.ImageFilePathTextBox.Location = new System.Drawing.Point(383, 26);
            this.ImageFilePathTextBox.Name = "ImageFilePathTextBox";
            this.ImageFilePathTextBox.Size = new System.Drawing.Size(151, 21);
            this.ImageFilePathTextBox.TabIndex = 3;
            // 
            // ExcelFilePathTextBox
            // 
            this.ExcelFilePathTextBox.Location = new System.Drawing.Point(383, 53);
            this.ExcelFilePathTextBox.Name = "ExcelFilePathTextBox";
            this.ExcelFilePathTextBox.Size = new System.Drawing.Size(151, 21);
            this.ExcelFilePathTextBox.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(276, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 12);
            this.label2.TabIndex = 4;
            this.label2.Text = "Excel File Path:";
            // 
            // StartButton
            // 
            this.StartButton.Location = new System.Drawing.Point(383, 93);
            this.StartButton.Name = "StartButton";
            this.StartButton.Size = new System.Drawing.Size(151, 161);
            this.StartButton.TabIndex = 6;
            this.StartButton.Text = "Open Image File";
            this.StartButton.UseVisualStyleBackColor = true;
            this.StartButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // ProcessProgressBar
            // 
            this.ProcessProgressBar.Location = new System.Drawing.Point(73, 260);
            this.ProcessProgressBar.Name = "ProcessProgressBar";
            this.ProcessProgressBar.Size = new System.Drawing.Size(461, 23);
            this.ProcessProgressBar.TabIndex = 7;
            // 
            // ProcessLable
            // 
            this.ProcessLable.AutoSize = true;
            this.ProcessLable.Location = new System.Drawing.Point(10, 264);
            this.ProcessLable.Name = "ProcessLable";
            this.ProcessLable.Size = new System.Drawing.Size(29, 12);
            this.ProcessLable.TabIndex = 8;
            this.ProcessLable.Text = "100%";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ImageInfoRichTextBox);
            this.groupBox2.Location = new System.Drawing.Point(278, 93);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(99, 158);
            this.groupBox2.TabIndex = 9;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Image Info";
            this.groupBox2.Enter += new System.EventHandler(this.groupBox2_Enter);
            // 
            // ImageInfoRichTextBox
            // 
            this.ImageInfoRichTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.ImageInfoRichTextBox.Location = new System.Drawing.Point(3, 17);
            this.ImageInfoRichTextBox.Name = "ImageInfoRichTextBox";
            this.ImageInfoRichTextBox.ReadOnly = true;
            this.ImageInfoRichTextBox.Size = new System.Drawing.Size(93, 138);
            this.ImageInfoRichTextBox.TabIndex = 0;
            this.ImageInfoRichTextBox.Text = "";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(543, 295);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.ProcessLable);
            this.Controls.Add(this.ProcessProgressBar);
            this.Controls.Add(this.StartButton);
            this.Controls.Add(this.ExcelFilePathTextBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.ImageFilePathTextBox);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Image To Excel Bitmap";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ViewPicureBox)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox ViewPicureBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox ImageFilePathTextBox;
        private System.Windows.Forms.TextBox ExcelFilePathTextBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button StartButton;
        private System.Windows.Forms.ProgressBar ProcessProgressBar;
        private System.Windows.Forms.Label ProcessLable;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RichTextBox ImageInfoRichTextBox;
    }
}

