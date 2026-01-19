namespace Сoursework1
{
    partial class Tablecreator
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
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.btnOpenFiles = new System.Windows.Forms.Button();
            this.btnExportCSV = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.lblStatus = new System.Windows.Forms.Label();
            this.btnClear = new System.Windows.Forms.Button();
            this.ButtonExportWord = new System.Windows.Forms.Button();
            this.ButtonExportExcel = new System.Windows.Forms.Button();
            this.Exit = new System.Windows.Forms.Button();
            this.BtnOpenCSV = new System.Windows.Forms.Button();
            this.BtnOpenExcel = new System.Windows.Forms.Button();
            this.BtnOpenWORD = new System.Windows.Forms.Button();
            this.CreateChart = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.CreateChartBtn = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.CreateChart)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOpenFiles
            // 
            this.btnOpenFiles.Location = new System.Drawing.Point(278, 12);
            this.btnOpenFiles.Name = "btnOpenFiles";
            this.btnOpenFiles.Size = new System.Drawing.Size(143, 60);
            this.btnOpenFiles.TabIndex = 13;
            this.btnOpenFiles.Text = "btnOpenFiles";
            this.btnOpenFiles.Click += new System.EventHandler(this.btnOpenFiles_Click);
            // 
            // btnExportCSV
            // 
            this.btnExportCSV.Location = new System.Drawing.Point(440, 12);
            this.btnExportCSV.Name = "btnExportCSV";
            this.btnExportCSV.Size = new System.Drawing.Size(145, 23);
            this.btnExportCSV.TabIndex = 1;
            this.btnExportCSV.Text = "btnExportCSV";
            this.btnExportCSV.UseVisualStyleBackColor = true;
            this.btnExportCSV.Click += new System.EventHandler(this.btnExportCSV_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 114);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersWidth = 51;
            this.dataGridView1.RowTemplate.Height = 24;
            this.dataGridView1.Size = new System.Drawing.Size(1387, 216);
            this.dataGridView1.TabIndex = 2;
            this.dataGridView1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellContentClick);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Location = new System.Drawing.Point(707, 95);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(58, 16);
            this.lblStatus.TabIndex = 3;
            this.lblStatus.Text = "lblStatus";
            // 
            // btnClear
            // 
            this.btnClear.Location = new System.Drawing.Point(985, 52);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(157, 23);
            this.btnClear.TabIndex = 4;
            this.btnClear.Text = "btnClear";
            this.btnClear.UseVisualStyleBackColor = true;
            // 
            // ButtonExportWord
            // 
            this.ButtonExportWord.Location = new System.Drawing.Point(620, 12);
            this.ButtonExportWord.Name = "ButtonExportWord";
            this.ButtonExportWord.Size = new System.Drawing.Size(145, 23);
            this.ButtonExportWord.TabIndex = 5;
            this.ButtonExportWord.Text = "buttonExportWord";
            this.ButtonExportWord.UseVisualStyleBackColor = true;
            this.ButtonExportWord.Click += new System.EventHandler(this.ButtonExportWord_Click);
            // 
            // ButtonExportExcel
            // 
            this.ButtonExportExcel.Location = new System.Drawing.Point(800, 13);
            this.ButtonExportExcel.Name = "ButtonExportExcel";
            this.ButtonExportExcel.Size = new System.Drawing.Size(145, 23);
            this.ButtonExportExcel.TabIndex = 6;
            this.ButtonExportExcel.Text = "buttonExportExcel";
            this.ButtonExportExcel.UseVisualStyleBackColor = true;
            this.ButtonExportExcel.Click += new System.EventHandler(this.ButtonExportExcel_Click);
            // 
            // Exit
            // 
            this.Exit.Location = new System.Drawing.Point(985, 12);
            this.Exit.Name = "Exit";
            this.Exit.Size = new System.Drawing.Size(157, 23);
            this.Exit.TabIndex = 7;
            this.Exit.Text = "Exit";
            this.Exit.UseVisualStyleBackColor = true;
            this.Exit.Click += new System.EventHandler(this.Exit_Click);
            // 
            // BtnOpenCSV
            // 
            this.BtnOpenCSV.Location = new System.Drawing.Point(440, 52);
            this.BtnOpenCSV.Name = "BtnOpenCSV";
            this.BtnOpenCSV.Size = new System.Drawing.Size(145, 23);
            this.BtnOpenCSV.TabIndex = 8;
            this.BtnOpenCSV.Text = "btnOpenCSV";
            this.BtnOpenCSV.UseVisualStyleBackColor = true;
            this.BtnOpenCSV.Click += new System.EventHandler(this.BtnOpenCSV_Click);
            // 
            // BtnOpenExcel
            // 
            this.BtnOpenExcel.Location = new System.Drawing.Point(800, 52);
            this.BtnOpenExcel.Name = "BtnOpenExcel";
            this.BtnOpenExcel.Size = new System.Drawing.Size(145, 23);
            this.BtnOpenExcel.TabIndex = 9;
            this.BtnOpenExcel.Text = "btnOpenExcel";
            this.BtnOpenExcel.UseVisualStyleBackColor = true;
            this.BtnOpenExcel.Click += new System.EventHandler(this.BtnOpenExcel_Click);
            // 
            // BtnOpenWORD
            // 
            this.BtnOpenWORD.Location = new System.Drawing.Point(620, 51);
            this.BtnOpenWORD.Name = "BtnOpenWORD";
            this.BtnOpenWORD.Size = new System.Drawing.Size(145, 23);
            this.BtnOpenWORD.TabIndex = 10;
            this.BtnOpenWORD.Text = "btnOpenWord";
            this.BtnOpenWORD.UseVisualStyleBackColor = true;
            this.BtnOpenWORD.Click += new System.EventHandler(this.BtnOpenWORD_Click);
            // 
            // CreateChart
            // 
            chartArea1.Name = "ChartArea1";
            this.CreateChart.ChartAreas.Add(chartArea1);
            legend1.Name = "Legend1";
            this.CreateChart.Legends.Add(legend1);
            this.CreateChart.Location = new System.Drawing.Point(12, 350);
            this.CreateChart.Name = "CreateChart";
            series1.ChartArea = "ChartArea1";
            series1.Legend = "Legend1";
            series1.Name = "Series1";
            this.CreateChart.Series.Add(series1);
            this.CreateChart.Size = new System.Drawing.Size(1376, 300);
            this.CreateChart.TabIndex = 11;
            this.CreateChart.Text = "CreateChart";
            this.CreateChart.Click += new System.EventHandler(this.CreateChart_Click);
            // 
            // CreateChartBtn
            // 
            this.CreateChartBtn.Location = new System.Drawing.Point(1172, 12);
            this.CreateChartBtn.Name = "CreateChartBtn";
            this.CreateChartBtn.Size = new System.Drawing.Size(143, 61);
            this.CreateChartBtn.TabIndex = 12;
            this.CreateChartBtn.Text = "CreateChartBtn";
            this.CreateChartBtn.UseVisualStyleBackColor = true;
            this.CreateChartBtn.Click += new System.EventHandler(this.CreateChartBtn_Click);
            // 
            // Tablecreator
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1424, 700);
            this.Controls.Add(this.CreateChartBtn);
            this.Controls.Add(this.CreateChart);
            this.Controls.Add(this.BtnOpenWORD);
            this.Controls.Add(this.BtnOpenExcel);
            this.Controls.Add(this.BtnOpenCSV);
            this.Controls.Add(this.Exit);
            this.Controls.Add(this.ButtonExportExcel);
            this.Controls.Add(this.ButtonExportWord);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnExportCSV);
            this.Controls.Add(this.btnOpenFiles);
            this.Name = "Tablecreator";
            this.Text = "table creator";
            this.Load += new System.EventHandler(this.Tablecreator_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.CreateChart)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpenFiles;
        private System.Windows.Forms.Button btnExportCSV;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.Button ButtonExportWord;
        private System.Windows.Forms.Button ButtonExportExcel;
        private System.Windows.Forms.Button Exit;
        private System.Windows.Forms.Button BtnOpenCSV;
        private System.Windows.Forms.Button BtnOpenExcel;
        private System.Windows.Forms.Button BtnOpenWORD;
        private System.Windows.Forms.DataVisualization.Charting.Chart CreateChart;
        private System.Windows.Forms.Button CreateChartBtn;
    }
}

