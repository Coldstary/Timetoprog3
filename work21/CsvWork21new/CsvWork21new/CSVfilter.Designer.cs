namespace CsvWork21new
{
    partial class CSVfilter
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
            this.CmbFilterType = new System.Windows.Forms.ComboBox();
            this.monthCalendar = new System.Windows.Forms.MonthCalendar();
            this.lstResults = new System.Windows.Forms.ListBox();
            this.TxtSearchParam1 = new System.Windows.Forms.TextBox();
            this.TxtSearchParam2 = new System.Windows.Forms.TextBox();
            this.TxtSearchParam3 = new System.Windows.Forms.TextBox();
            this.LblParam1 = new System.Windows.Forms.Label();
            this.LblParam2 = new System.Windows.Forms.Label();
            this.LblParam3 = new System.Windows.Forms.Label();
            this.BtnExecute = new System.Windows.Forms.Button();
            this.BtnReserve = new System.Windows.Forms.Button();
            this.BtnShowAll = new System.Windows.Forms.Button();
            this.checkAdmin = new System.Windows.Forms.CheckBox();
            this.Booking = new System.Windows.Forms.Button();
            this.DiseaseInfo = new System.Windows.Forms.Button();
            this.Exit = new System.Windows.Forms.Button();
            this.Reset = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CmbFilterType
            // 
            this.CmbFilterType.FormattingEnabled = true;
            this.CmbFilterType.Location = new System.Drawing.Point(12, 45);
            this.CmbFilterType.Name = "CmbFilterType";
            this.CmbFilterType.Size = new System.Drawing.Size(295, 24);
            this.CmbFilterType.TabIndex = 0;
            this.CmbFilterType.SelectedIndexChanged += new System.EventHandler(this.CmbFilterType_SelectedIndexChanged);
            // 
            // monthCalendar
            // 
            this.monthCalendar.Location = new System.Drawing.Point(33, 105);
            this.monthCalendar.Name = "monthCalendar";
            this.monthCalendar.TabIndex = 1;
            this.monthCalendar.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar_DateChanged);
            // 
            // lstResults
            // 
            this.lstResults.FormattingEnabled = true;
            this.lstResults.ItemHeight = 16;
            this.lstResults.Location = new System.Drawing.Point(326, 45);
            this.lstResults.Name = "lstResults";
            this.lstResults.Size = new System.Drawing.Size(574, 340);
            this.lstResults.TabIndex = 2;
            this.lstResults.SelectedIndexChanged += new System.EventHandler(this.lstResults_SelectedIndexChanged);
            // 
            // TxtSearchParam1
            // 
            this.TxtSearchParam1.Location = new System.Drawing.Point(111, 415);
            this.TxtSearchParam1.Name = "TxtSearchParam1";
            this.TxtSearchParam1.Size = new System.Drawing.Size(157, 22);
            this.TxtSearchParam1.TabIndex = 4;
            this.TxtSearchParam1.TextChanged += new System.EventHandler(this.TxtSearchParam1_TextChanged);
            // 
            // TxtSearchParam2
            // 
            this.TxtSearchParam2.Location = new System.Drawing.Point(386, 415);
            this.TxtSearchParam2.Name = "TxtSearchParam2";
            this.TxtSearchParam2.Size = new System.Drawing.Size(157, 22);
            this.TxtSearchParam2.TabIndex = 5;
            this.TxtSearchParam2.TextChanged += new System.EventHandler(this.TxtSearchParam2_TextChanged);
            // 
            // TxtSearchParam3
            // 
            this.TxtSearchParam3.Location = new System.Drawing.Point(643, 415);
            this.TxtSearchParam3.Name = "TxtSearchParam3";
            this.TxtSearchParam3.Size = new System.Drawing.Size(157, 22);
            this.TxtSearchParam3.TabIndex = 6;
            this.TxtSearchParam3.TextChanged += new System.EventHandler(this.TxtSearchParam3_TextChanged);
            // 
            // LblParam1
            // 
            this.LblParam1.AutoSize = true;
            this.LblParam1.Location = new System.Drawing.Point(108, 396);
            this.LblParam1.Name = "LblParam1";
            this.LblParam1.Size = new System.Drawing.Size(83, 16);
            this.LblParam1.TabIndex = 7;
            this.LblParam1.Text = "expectation1";
            this.LblParam1.Click += new System.EventHandler(this.LblParam1_Click);
            // 
            // LblParam2
            // 
            this.LblParam2.AutoSize = true;
            this.LblParam2.Location = new System.Drawing.Point(383, 396);
            this.LblParam2.Name = "LblParam2";
            this.LblParam2.Size = new System.Drawing.Size(83, 16);
            this.LblParam2.TabIndex = 8;
            this.LblParam2.Text = "expectation2";
            this.LblParam2.Click += new System.EventHandler(this.LblParam2_Click);
            // 
            // LblParam3
            // 
            this.LblParam3.AutoSize = true;
            this.LblParam3.Location = new System.Drawing.Point(640, 396);
            this.LblParam3.Name = "LblParam3";
            this.LblParam3.Size = new System.Drawing.Size(83, 16);
            this.LblParam3.TabIndex = 9;
            this.LblParam3.Text = "expectation3";
            this.LblParam3.Click += new System.EventHandler(this.LblParam3_Click);
            // 
            // BtnExecute
            // 
            this.BtnExecute.Location = new System.Drawing.Point(906, 45);
            this.BtnExecute.Name = "BtnExecute";
            this.BtnExecute.Size = new System.Drawing.Size(157, 55);
            this.BtnExecute.TabIndex = 10;
            this.BtnExecute.Text = "btnExecute";
            this.BtnExecute.UseVisualStyleBackColor = true;
            this.BtnExecute.Click += new System.EventHandler(this.BtnExecute_Click);
            // 
            // BtnReserve
            // 
            this.BtnReserve.Location = new System.Drawing.Point(906, 198);
            this.BtnReserve.Name = "BtnReserve";
            this.BtnReserve.Size = new System.Drawing.Size(157, 55);
            this.BtnReserve.TabIndex = 11;
            this.BtnReserve.Text = "BtnReserve";
            this.BtnReserve.UseVisualStyleBackColor = true;
            this.BtnReserve.Click += new System.EventHandler(this.BtnReserve_Click);
            // 
            // BtnShowAll
            // 
            this.BtnShowAll.Location = new System.Drawing.Point(906, 117);
            this.BtnShowAll.Name = "BtnShowAll";
            this.BtnShowAll.Size = new System.Drawing.Size(157, 55);
            this.BtnShowAll.TabIndex = 12;
            this.BtnShowAll.Text = "btnShowAll";
            this.BtnShowAll.UseVisualStyleBackColor = true;
            this.BtnShowAll.Click += new System.EventHandler(this.BtnShowAll_Click);
            // 
            // checkAdmin
            // 
            this.checkAdmin.AutoSize = true;
            this.checkAdmin.Location = new System.Drawing.Point(848, 536);
            this.checkAdmin.Name = "checkAdmin";
            this.checkAdmin.Size = new System.Drawing.Size(235, 20);
            this.checkAdmin.TabIndex = 14;
            this.checkAdmin.Text = "Вход от имени администратора";
            this.checkAdmin.UseVisualStyleBackColor = true;
            this.checkAdmin.CheckedChanged += new System.EventHandler(this.checkAdmin_CheckedChanged);
            // 
            // Booking
            // 
            this.Booking.Location = new System.Drawing.Point(906, 476);
            this.Booking.Name = "Booking";
            this.Booking.Size = new System.Drawing.Size(157, 54);
            this.Booking.TabIndex = 15;
            this.Booking.Text = "Проверка Брони";
            this.Booking.UseVisualStyleBackColor = true;
            this.Booking.Visible = false;
            this.Booking.Click += new System.EventHandler(this.Booking_Click);
            // 
            // DiseaseInfo
            // 
            this.DiseaseInfo.Location = new System.Drawing.Point(111, 456);
            this.DiseaseInfo.Name = "DiseaseInfo";
            this.DiseaseInfo.Size = new System.Drawing.Size(157, 54);
            this.DiseaseInfo.TabIndex = 16;
            this.DiseaseInfo.Text = "Информация о болезнях";
            this.DiseaseInfo.UseVisualStyleBackColor = true;
            this.DiseaseInfo.Visible = false;
            this.DiseaseInfo.Click += new System.EventHandler(this.DiseaseInfo_Click);
            // 
            // Exit
            // 
            this.Exit.Location = new System.Drawing.Point(481, 517);
            this.Exit.Name = "Exit";
            this.Exit.Size = new System.Drawing.Size(153, 39);
            this.Exit.TabIndex = 17;
            this.Exit.Text = "Exit";
            this.Exit.UseVisualStyleBackColor = true;
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // Reset
            // 
            this.Reset.Location = new System.Drawing.Point(906, 425);
            this.Reset.Name = "Reset";
            this.Reset.Size = new System.Drawing.Size(157, 35);
            this.Reset.TabIndex = 18;
            this.Reset.Text = "Сброс";
            this.Reset.UseVisualStyleBackColor = true;
            this.Reset.Visible = false;
            this.Reset.Click += new System.EventHandler(this.Reset_Click);
            // 
            // CSVfilter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1095, 568);
            this.Controls.Add(this.Reset);
            this.Controls.Add(this.Exit);
            this.Controls.Add(this.DiseaseInfo);
            this.Controls.Add(this.Booking);
            this.Controls.Add(this.checkAdmin);
            this.Controls.Add(this.BtnShowAll);
            this.Controls.Add(this.BtnReserve);
            this.Controls.Add(this.BtnExecute);
            this.Controls.Add(this.LblParam3);
            this.Controls.Add(this.LblParam2);
            this.Controls.Add(this.LblParam1);
            this.Controls.Add(this.TxtSearchParam3);
            this.Controls.Add(this.TxtSearchParam2);
            this.Controls.Add(this.TxtSearchParam1);
            this.Controls.Add(this.lstResults);
            this.Controls.Add(this.monthCalendar);
            this.Controls.Add(this.CmbFilterType);
            this.Name = "CSVfilter";
            this.Text = "CSVfilter";
            this.Load += new System.EventHandler(this.CSVfilter_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox CmbFilterType;
        private System.Windows.Forms.MonthCalendar monthCalendar;
        private System.Windows.Forms.ListBox lstResults;
        private System.Windows.Forms.TextBox TxtSearchParam1;
        private System.Windows.Forms.TextBox  TxtSearchParam2;
        private System.Windows.Forms.TextBox TxtSearchParam3;
        private System.Windows.Forms.Label LblParam1;
        private System.Windows.Forms.Label LblParam2;
        private System.Windows.Forms.Label LblParam3;
        private System.Windows.Forms.Button BtnExecute;
        private System.Windows.Forms.Button BtnReserve;
        private System.Windows.Forms.Button BtnShowAll;
        private System.Windows.Forms.CheckBox checkAdmin;
        private System.Windows.Forms.Button Booking;
        private System.Windows.Forms.Button DiseaseInfo;
        private System.Windows.Forms.Button Exit;
        private System.Windows.Forms.Button Reset;
    }
}

