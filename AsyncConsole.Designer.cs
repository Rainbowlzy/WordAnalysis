namespace WordAnalysis
{
    partial class AsyncConsole
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
            this.openFileDialogDefault = new System.Windows.Forms.OpenFileDialog();
            this.textBoxFolder = new System.Windows.Forms.TextBox();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.folderBrowserDialogDefault = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // openFileDialogDefault
            // 
            this.openFileDialogDefault.FileName = "openFileDialog1";
            // 
            // textBoxFolder
            // 
            this.textBoxFolder.Dock = System.Windows.Forms.DockStyle.Top;
            this.textBoxFolder.Location = new System.Drawing.Point(0, 0);
            this.textBoxFolder.Name = "textBoxFolder";
            this.textBoxFolder.Size = new System.Drawing.Size(852, 21);
            this.textBoxFolder.TabIndex = 1;
            this.textBoxFolder.TextChanged += new System.EventHandler(this.textBoxFolder_TextChanged);
            this.textBoxFolder.DoubleClick += new System.EventHandler(this.textBoxFolder_DoubleClick);
            // 
            // txtOutput
            // 
            this.txtOutput.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtOutput.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtOutput.Location = new System.Drawing.Point(0, 21);
            this.txtOutput.Multiline = true;
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.Size = new System.Drawing.Size(852, 527);
            this.txtOutput.TabIndex = 2;
            // 
            // AsyncConsole
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(852, 548);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.textBoxFolder);
            this.Name = "AsyncConsole";
            this.Text = "AsyncConsole";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.OpenFileDialog openFileDialogDefault;
        private System.Windows.Forms.TextBox textBoxFolder;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialogDefault;
    }
}

