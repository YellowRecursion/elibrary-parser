namespace eLIBRARYparsing
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.inputPathTextBox = new System.Windows.Forms.TextBox();
            this.selectInputPathButton = new System.Windows.Forms.Button();
            this.startButton = new System.Windows.Forms.Button();
            this.openExcelFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.logs = new System.Windows.Forms.RichTextBox();
            this.loginField = new System.Windows.Forms.TextBox();
            this.passwordField = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // inputPathTextBox
            // 
            this.inputPathTextBox.Location = new System.Drawing.Point(11, 156);
            this.inputPathTextBox.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.inputPathTextBox.Name = "inputPathTextBox";
            this.inputPathTextBox.ReadOnly = true;
            this.inputPathTextBox.Size = new System.Drawing.Size(426, 26);
            this.inputPathTextBox.TabIndex = 0;
            // 
            // selectInputPathButton
            // 
            this.selectInputPathButton.Location = new System.Drawing.Point(11, 191);
            this.selectInputPathButton.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.selectInputPathButton.Name = "selectInputPathButton";
            this.selectInputPathButton.Size = new System.Drawing.Size(426, 51);
            this.selectInputPathButton.TabIndex = 2;
            this.selectInputPathButton.Text = "Select file with names";
            this.selectInputPathButton.UseVisualStyleBackColor = true;
            this.selectInputPathButton.Click += new System.EventHandler(this.OpenExcelFileDialog);
            // 
            // startButton
            // 
            this.startButton.Location = new System.Drawing.Point(11, 636);
            this.startButton.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.startButton.Name = "startButton";
            this.startButton.Size = new System.Drawing.Size(425, 55);
            this.startButton.TabIndex = 4;
            this.startButton.Text = "Start";
            this.startButton.UseVisualStyleBackColor = true;
            this.startButton.Click += new System.EventHandler(this.StartButton_Click);
            // 
            // openExcelFileDialog
            // 
            this.openExcelFileDialog.FileName = "Excel файл";
            this.openExcelFileDialog.Filter = "Файлы Excel|*.xlsx*";
            this.openExcelFileDialog.FileOk += new System.ComponentModel.CancelEventHandler(this.ExcelFileDialog_FileOk);
            // 
            // logs
            // 
            this.logs.AccessibleDescription = "";
            this.logs.Font = new System.Drawing.Font("Microsoft Sans Serif", 11F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.logs.Location = new System.Drawing.Point(456, 11);
            this.logs.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.logs.Name = "logs";
            this.logs.ReadOnly = true;
            this.logs.Size = new System.Drawing.Size(654, 680);
            this.logs.TabIndex = 5;
            this.logs.Text = "";
            // 
            // loginField
            // 
            this.loginField.Location = new System.Drawing.Point(226, 404);
            this.loginField.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.loginField.Name = "loginField";
            this.loginField.Size = new System.Drawing.Size(210, 26);
            this.loginField.TabIndex = 8;
            // 
            // passwordField
            // 
            this.passwordField.Location = new System.Drawing.Point(226, 441);
            this.passwordField.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.passwordField.Name = "passwordField";
            this.passwordField.Size = new System.Drawing.Size(210, 26);
            this.passwordField.TabIndex = 9;
            this.passwordField.UseSystemPasswordChar = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
            this.pictureBox1.Location = new System.Drawing.Point(11, 11);
            this.pictureBox1.Margin = new System.Windows.Forms.Padding(2);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(425, 144);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 10;
            this.pictureBox1.TabStop = false;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 408);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 20);
            this.label1.TabIndex = 11;
            this.label1.Text = "Elibrary login";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 445);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(133, 20);
            this.label2.TabIndex = 12;
            this.label2.Text = "Elibrary password";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1116, 699);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.passwordField);
            this.Controls.Add(this.loginField);
            this.Controls.Add(this.logs);
            this.Controls.Add(this.startButton);
            this.Controls.Add(this.selectInputPathButton);
            this.Controls.Add(this.inputPathTextBox);
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(1138, 755);
            this.MinimumSize = new System.Drawing.Size(1138, 755);
            this.Name = "MainForm";
            this.Text = "eLibraryParser";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox inputPathTextBox;
        private System.Windows.Forms.Button selectInputPathButton;
        private System.Windows.Forms.Button startButton;
        private System.Windows.Forms.OpenFileDialog openExcelFileDialog;
        private System.Windows.Forms.RichTextBox logs;
        private System.Windows.Forms.TextBox loginField;
        private System.Windows.Forms.TextBox passwordField;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

