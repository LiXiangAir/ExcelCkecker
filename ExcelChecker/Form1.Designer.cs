namespace ExcelChecker
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.importExcel = new System.Windows.Forms.Button();
            this.browse = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(178, 26);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(344, 26);
            this.textBox1.TabIndex = 6;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 12F);
            this.label1.Location = new System.Drawing.Point(11, 31);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(176, 16);
            this.label1.TabIndex = 5;
            this.label1.Text = "选择Excel(.xls)文件：";
            // 
            // importExcel
            // 
            this.importExcel.Enabled = false;
            this.importExcel.Font = new System.Drawing.Font("宋体", 12F);
            this.importExcel.Location = new System.Drawing.Point(470, 65);
            this.importExcel.Name = "importExcel";
            this.importExcel.Size = new System.Drawing.Size(124, 29);
            this.importExcel.TabIndex = 9;
            this.importExcel.Text = "检测重复词汇";
            this.importExcel.UseVisualStyleBackColor = true;
            this.importExcel.Click += new System.EventHandler(this.importExcel_Click);
            // 
            // browse
            // 
            this.browse.Font = new System.Drawing.Font("宋体", 12F);
            this.browse.Location = new System.Drawing.Point(526, 26);
            this.browse.Name = "browse";
            this.browse.Size = new System.Drawing.Size(68, 26);
            this.browse.TabIndex = 7;
            this.browse.Text = "导入";
            this.browse.UseVisualStyleBackColor = true;
            this.browse.Click += new System.EventHandler(this.browse_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(611, 115);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.importExcel);
            this.Controls.Add(this.browse);
            this.Name = "Form1";
            this.Text = "Excel检测器";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button importExcel;
        private System.Windows.Forms.Button browse;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
    }
}

