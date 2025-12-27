namespace ExcelWork20
{
    partial class Work20
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
            this.buttonOpen = new System.Windows.Forms.Button();
            this.Data_entry = new System.Windows.Forms.Button();
            this.exit = new System.Windows.Forms.Button();
            this.info = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonOpen
            // 
            this.buttonOpen.Location = new System.Drawing.Point(12, 12);
            this.buttonOpen.Name = "buttonOpen";
            this.buttonOpen.Size = new System.Drawing.Size(145, 47);
            this.buttonOpen.TabIndex = 0;
            this.buttonOpen.Text = "Create File";
            this.buttonOpen.UseVisualStyleBackColor = true;
            this.buttonOpen.Click += new System.EventHandler(this.buttonOpen_Click);
            // 
            // Data_entry
            // 
            this.Data_entry.Location = new System.Drawing.Point(187, 12);
            this.Data_entry.Name = "Data_entry";
            this.Data_entry.Size = new System.Drawing.Size(145, 47);
            this.Data_entry.TabIndex = 1;
            this.Data_entry.Text = "Data_entry";
            this.Data_entry.UseVisualStyleBackColor = true;
            this.Data_entry.Click += new System.EventHandler(this.Data_entry_Click);
            // 
            // exit
            // 
            this.exit.Location = new System.Drawing.Point(12, 86);
            this.exit.Name = "exit";
            this.exit.Size = new System.Drawing.Size(145, 47);
            this.exit.TabIndex = 3;
            this.exit.Text = "exit";
            this.exit.UseVisualStyleBackColor = true;
            this.exit.Click += new System.EventHandler(this.exit_Click);
            // 
            // info
            // 
            this.info.Location = new System.Drawing.Point(187, 86);
            this.info.Name = "info";
            this.info.Size = new System.Drawing.Size(145, 47);
            this.info.TabIndex = 4;
            this.info.Text = "info";
            this.info.UseVisualStyleBackColor = true;
            this.info.Click += new System.EventHandler(this.info_Click);
            // 
            // Work20
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(360, 150);
            this.Controls.Add(this.info);
            this.Controls.Add(this.exit);
            this.Controls.Add(this.Data_entry);
            this.Controls.Add(this.buttonOpen);
            this.Name = "Work20";
            this.Text = "Work20";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonOpen;
        private System.Windows.Forms.Button Data_entry;
        private System.Windows.Forms.Button exit;
        private System.Windows.Forms.Button info;
    }
}

