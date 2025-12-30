namespace WordWork19
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
            this.CreateWord = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.FormFile = new System.Windows.Forms.Button();
            this.checkColor1 = new System.Windows.Forms.ComboBox();
            this.input1 = new System.Windows.Forms.Button();
            this.Input = new System.Windows.Forms.TextBox();
            this.Back = new System.Windows.Forms.Button();
            this.Preview = new System.Windows.Forms.Button();
            this.CreateReceipt = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CreateWord
            // 
            this.CreateWord.Location = new System.Drawing.Point(12, 12);
            this.CreateWord.Name = "CreateWord";
            this.CreateWord.Size = new System.Drawing.Size(116, 23);
            this.CreateWord.TabIndex = 0;
            this.CreateWord.Text = "Создать";
            this.CreateWord.UseVisualStyleBackColor = true;
            this.CreateWord.Click += new System.EventHandler(this.CreateWord_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(12, 41);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(776, 282);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            this.richTextBox1.TextChanged += new System.EventHandler(this.richTextBox1_TextChanged);
            // 
            // FormFile
            // 
            this.FormFile.Location = new System.Drawing.Point(172, 12);
            this.FormFile.Name = "FormFile";
            this.FormFile.Size = new System.Drawing.Size(157, 23);
            this.FormFile.TabIndex = 2;
            this.FormFile.Text = "Заполнить";
            this.FormFile.UseVisualStyleBackColor = true;
            this.FormFile.Click += new System.EventHandler(this.FormFile_Click);
            // 
            // checkColor1
            // 
            this.checkColor1.FormattingEnabled = true;
            this.checkColor1.Location = new System.Drawing.Point(667, 329);
            this.checkColor1.Name = "checkColor1";
            this.checkColor1.Size = new System.Drawing.Size(121, 24);
            this.checkColor1.TabIndex = 4;
            this.checkColor1.SelectedIndexChanged += new System.EventHandler(this.checkColor1_SelectedIndexChanged);
            // 
            // input1
            // 
            this.input1.Location = new System.Drawing.Point(297, 373);
            this.input1.Name = "input1";
            this.input1.Size = new System.Drawing.Size(170, 65);
            this.input1.TabIndex = 5;
            this.input1.Text = "Ввести";
            this.input1.UseVisualStyleBackColor = true;
            this.input1.Click += new System.EventHandler(this.input1_Click);
            // 
            // Input
            // 
            this.Input.Location = new System.Drawing.Point(12, 331);
            this.Input.Name = "Input";
            this.Input.Size = new System.Drawing.Size(658, 22);
            this.Input.TabIndex = 3;
            this.Input.TextChanged += new System.EventHandler(this.Input_TextChanged);
            // 
            // Back
            // 
            this.Back.Location = new System.Drawing.Point(110, 373);
            this.Back.Name = "Back";
            this.Back.Size = new System.Drawing.Size(157, 65);
            this.Back.TabIndex = 6;
            this.Back.Text = "Назад";
            this.Back.UseVisualStyleBackColor = true;
            this.Back.Click += new System.EventHandler(this.Back_Click);
            // 
            // Preview
            // 
            this.Preview.Location = new System.Drawing.Point(500, 373);
            this.Preview.Name = "Preview";
            this.Preview.Size = new System.Drawing.Size(170, 65);
            this.Preview.TabIndex = 7;
            this.Preview.Text = "Предпросмотр";
            this.Preview.UseVisualStyleBackColor = true;
            this.Preview.Click += new System.EventHandler(this.Preview_Click);
            // 
            // CreateReceipt
            // 
            this.CreateReceipt.Location = new System.Drawing.Point(375, 12);
            this.CreateReceipt.Name = "CreateReceipt";
            this.CreateReceipt.Size = new System.Drawing.Size(157, 23);
            this.CreateReceipt.TabIndex = 8;
            this.CreateReceipt.Text = "Создать квитанцию";
            this.CreateReceipt.UseVisualStyleBackColor = true;
            this.CreateReceipt.Click += new System.EventHandler(this.CreateReceipt_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.CreateReceipt);
            this.Controls.Add(this.Preview);
            this.Controls.Add(this.Back);
            this.Controls.Add(this.input1);
            this.Controls.Add(this.checkColor1);
            this.Controls.Add(this.Input);
            this.Controls.Add(this.FormFile);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.CreateWord);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button CreateWord;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button FormFile;
        private System.Windows.Forms.ComboBox checkColor1;
        private System.Windows.Forms.Button input1;
        private System.Windows.Forms.TextBox Input;
        private System.Windows.Forms.Button Back;
        private System.Windows.Forms.Button Preview;
        private System.Windows.Forms.Button CreateReceipt;
    }
}

