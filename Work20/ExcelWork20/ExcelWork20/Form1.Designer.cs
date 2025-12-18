namespace ExcelWork20
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
            this.buttonOpen = new System.Windows.Forms.Button();
            this.Data_entry = new System.Windows.Forms.Button();
            this.exit = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // buttonOpen
            // 
            this.buttonOpen.Location = new System.Drawing.Point(98, 160);
            this.buttonOpen.Name = "buttonOpen";
            this.buttonOpen.Size = new System.Drawing.Size(145, 47);
            this.buttonOpen.TabIndex = 0;
            this.buttonOpen.Text = "Create File";
            this.buttonOpen.UseVisualStyleBackColor = true;
            this.buttonOpen.Click += new System.EventHandler(this.buttonOpen_Click);
            // 
            // Data_entry
            // 
            this.Data_entry.Location = new System.Drawing.Point(331, 160);
            this.Data_entry.Name = "Data_entry";
            this.Data_entry.Size = new System.Drawing.Size(142, 47);
            this.Data_entry.TabIndex = 1;
            this.Data_entry.Text = "Data_entry";
            this.Data_entry.UseVisualStyleBackColor = true;
            this.Data_entry.Click += new System.EventHandler(this.Data_entry_Click);
            // 
            // exit
            // 
            this.exit.Location = new System.Drawing.Point(548, 160);
            this.exit.Name = "exit";
            this.exit.Size = new System.Drawing.Size(127, 47);
            this.exit.TabIndex = 3;
            this.exit.Text = "exit";
            this.exit.UseVisualStyleBackColor = true;
            this.exit.Click += new System.EventHandler(this.exit_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.exit);
            this.Controls.Add(this.Data_entry);
            this.Controls.Add(this.buttonOpen);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button buttonOpen;
        private System.Windows.Forms.Button Data_entry;
        private System.Windows.Forms.Button exit;
    }
}

