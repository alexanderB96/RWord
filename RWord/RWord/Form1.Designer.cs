namespace RWord
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.filepyt = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.Provodnik = new System.Windows.Forms.TreeView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // filepyt
            // 
            this.filepyt.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.filepyt.Location = new System.Drawing.Point(279, 12);
            this.filepyt.Name = "filepyt";
            this.filepyt.Size = new System.Drawing.Size(347, 20);
            this.filepyt.TabIndex = 0;
            this.filepyt.Text = "кликните для выбора файла. . .";
            this.filepyt.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.filepyt.MouseClick += new System.Windows.Forms.MouseEventHandler(this.filepyt_MouseClick);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // Provodnik
            // 
            this.Provodnik.ImageIndex = 0;
            this.Provodnik.ImageList = this.imageList1;
            this.Provodnik.Location = new System.Drawing.Point(2, 1);
            this.Provodnik.Name = "Provodnik";
            this.Provodnik.SelectedImageIndex = 0;
            this.Provodnik.Size = new System.Drawing.Size(197, 508);
            this.Provodnik.TabIndex = 1;
            this.Provodnik.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.Provodnik_BeforeExpand);
            this.Provodnik.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.Provodnik_AfterSelect);
            this.Provodnik.Click += new System.EventHandler(this.Provodnik_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "drive-icon.png");
            this.imageList1.Images.SetKeyName(1, "dvd-case-icon.png");
            this.imageList1.Images.SetKeyName(2, "folder-documents-icon.png");
            this.imageList1.Images.SetKeyName(3, "document-word-icon.png");
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.DarkRed;
            this.label1.Location = new System.Drawing.Point(202, 496);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "label1";
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(205, 226);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(120, 95);
            this.listBox1.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(231, 90);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 4;
            this.button1.Text = "button1";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 511);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Provodnik);
            this.Controls.Add(this.filepyt);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "RWord";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox filepyt;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        public System.Windows.Forms.TreeView Provodnik;
        public System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Button button1;
    }
}

