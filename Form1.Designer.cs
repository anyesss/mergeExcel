namespace excelMerge
{
    partial class MainForm
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
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.file1 = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.selectFile3 = new System.Windows.Forms.Button();
            this.file3 = new System.Windows.Forms.TextBox();
            this.selectFile2 = new System.Windows.Forms.Button();
            this.file2 = new System.Windows.Forms.TextBox();
            this.selectFile1 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // file1
            // 
            this.file1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.file1.Location = new System.Drawing.Point(24, 34);
            this.file1.Name = "file1";
            this.file1.Size = new System.Drawing.Size(347, 21);
            this.file1.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.selectFile3);
            this.groupBox1.Controls.Add(this.file3);
            this.groupBox1.Controls.Add(this.selectFile2);
            this.groupBox1.Controls.Add(this.file2);
            this.groupBox1.Controls.Add(this.selectFile1);
            this.groupBox1.Controls.Add(this.file1);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(502, 160);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "选择Excel文件";
            // 
            // selectFile3
            // 
            this.selectFile3.Location = new System.Drawing.Point(390, 114);
            this.selectFile3.Name = "selectFile3";
            this.selectFile3.Size = new System.Drawing.Size(93, 23);
            this.selectFile3.TabIndex = 12;
            this.selectFile3.Text = "选择文件3";
            this.selectFile3.UseVisualStyleBackColor = true;
            this.selectFile3.Click += new System.EventHandler(this.SelectFile3_Click);
            // 
            // file3
            // 
            this.file3.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.file3.Location = new System.Drawing.Point(24, 115);
            this.file3.Name = "file3";
            this.file3.Size = new System.Drawing.Size(347, 21);
            this.file3.TabIndex = 11;
            // 
            // selectFile2
            // 
            this.selectFile2.Location = new System.Drawing.Point(390, 73);
            this.selectFile2.Name = "selectFile2";
            this.selectFile2.Size = new System.Drawing.Size(93, 23);
            this.selectFile2.TabIndex = 10;
            this.selectFile2.Text = "选择文件2";
            this.selectFile2.UseVisualStyleBackColor = true;
            this.selectFile2.Click += new System.EventHandler(this.SelectFile2_Click);
            // 
            // file2
            // 
            this.file2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.file2.Location = new System.Drawing.Point(24, 74);
            this.file2.Name = "file2";
            this.file2.Size = new System.Drawing.Size(347, 21);
            this.file2.TabIndex = 9;
            // 
            // selectFile1
            // 
            this.selectFile1.Location = new System.Drawing.Point(390, 33);
            this.selectFile1.Name = "selectFile1";
            this.selectFile1.Size = new System.Drawing.Size(93, 23);
            this.selectFile1.TabIndex = 8;
            this.selectFile1.Text = "选择文件1";
            this.selectFile1.UseVisualStyleBackColor = true;
            this.selectFile1.Click += new System.EventHandler(this.SelectFile1_Click);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(205, 192);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(123, 45);
            this.button1.TabIndex = 2;
            this.button1.Text = "开始合并";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(526, 253);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "自动合并Excel";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox file1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button selectFile1;
        private System.Windows.Forms.Button selectFile3;
        private System.Windows.Forms.TextBox file3;
        private System.Windows.Forms.Button selectFile2;
        private System.Windows.Forms.TextBox file2;
    }
}

