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
            this.filepyt = new System.Windows.Forms.TextBox();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.Provodnik = new System.Windows.Forms.TreeView();
            this.label1 = new System.Windows.Forms.Label();
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
            this.Provodnik.Location = new System.Drawing.Point(2, 1);
            this.Provodnik.Name = "Provodnik";
            this.Provodnik.Size = new System.Drawing.Size(161, 508);
            this.Provodnik.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.Color.DarkRed;
            this.label1.Location = new System.Drawing.Point(169, 496);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(35, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "label1";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 511);
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
    }
}

