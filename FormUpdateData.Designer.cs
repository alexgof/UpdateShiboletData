namespace UpdateShiboletData
{
    partial class FormUpdateData
    {

        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            groupBox1 = new GroupBox();
            btnSaveDataToSQL = new Button();
            textBoxFilePath = new TextBox();
            btnSelectFile = new Button();
            openFileDialog1 = new OpenFileDialog();
            btnSaveDataToExcel = new Button();
            groupBox1.SuspendLayout();
            SuspendLayout();
            // 
            // groupBox1
            // 
            groupBox1.BackColor = SystemColors.Control;
            groupBox1.Controls.Add(btnSaveDataToExcel);
            groupBox1.Controls.Add(btnSaveDataToSQL);
            groupBox1.Controls.Add(textBoxFilePath);
            groupBox1.Controls.Add(btnSelectFile);
            groupBox1.Dock = DockStyle.Fill;
            groupBox1.Font = new Font("Calibri", 18F, FontStyle.Bold, GraphicsUnit.Point);
            groupBox1.Location = new Point(0, 0);
            groupBox1.Name = "groupBox1";
            groupBox1.Padding = new Padding(6);
            groupBox1.Size = new Size(800, 450);
            groupBox1.TabIndex = 0;
            groupBox1.TabStop = false;
            groupBox1.Text = "קליטה נתוני שיבולת ";
            // 
            // btnSaveDataToSQL
            // 
            btnSaveDataToSQL.Font = new Font("Calibri", 15.75F, FontStyle.Bold, GraphicsUnit.Point);
            btnSaveDataToSQL.Location = new Point(320, 111);
            btnSaveDataToSQL.Name = "btnSaveDataToSQL";
            btnSaveDataToSQL.Size = new Size(263, 33);
            btnSaveDataToSQL.TabIndex = 3;
            btnSaveDataToSQL.Text = "שמירה נתונים למאגר";
            btnSaveDataToSQL.UseVisualStyleBackColor = true;
            btnSaveDataToSQL.Click += btnSaveData_Click;
            // 
            // textBoxFilePath
            // 
            textBoxFilePath.Location = new Point(50, 48);
            textBoxFilePath.Name = "textBoxFilePath";
            textBoxFilePath.Size = new Size(533, 37);
            textBoxFilePath.TabIndex = 2;
            // 
            // btnSelectFile
            // 
            btnSelectFile.Font = new Font("Calibri", 15.75F, FontStyle.Bold, GraphicsUnit.Point);
            btnSelectFile.Location = new Point(610, 47);
            btnSelectFile.Name = "btnSelectFile";
            btnSelectFile.Size = new Size(162, 37);
            btnSelectFile.TabIndex = 0;
            btnSelectFile.Text = "בחירת קובץ";
            btnSelectFile.UseVisualStyleBackColor = true;
            btnSelectFile.Click += buttonSelectFile_Click;
            // 
            // openFileDialog1
            // 
            openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnSaveDataToExcel
            // 
            btnSaveDataToExcel.Font = new Font("Calibri", 15.75F, FontStyle.Bold, GraphicsUnit.Point);
            btnSaveDataToExcel.Location = new Point(320, 160);
            btnSaveDataToExcel.Name = "btnSaveDataToExcel";
            btnSaveDataToExcel.Size = new Size(263, 33);
            btnSaveDataToExcel.TabIndex = 3;
            btnSaveDataToExcel.Text = "שמירה נתונים לקובץ";
            btnSaveDataToExcel.UseVisualStyleBackColor = true;
            btnSaveDataToExcel.Click += btnSaveDataToExcel_Click;
            // 
            // FormUpdateData
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoScroll = true;
            BackColor = SystemColors.GradientInactiveCaption;
            ClientSize = new Size(800, 450);
            Controls.Add(groupBox1);
            Name = "FormUpdateData";
            Text = "קליטת נתונים V1.4";
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private GroupBox groupBox1;
        private Button btnSelectFile;
        private OpenFileDialog openFileDialog1;
        private TextBox textBoxFilePath;
        private Button btnSaveDataToSQL;
        private Button btnSaveDataToExcel;
    }
}
