namespace MergeExcel
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.SourceFileLbl = new System.Windows.Forms.Label();
            this.SourceFileBtn = new System.Windows.Forms.Button();
            this.SourceFileCheckedListBox = new System.Windows.Forms.CheckedListBox();
            this.label2 = new System.Windows.Forms.Label();
            this.TargFileTextbox = new System.Windows.Forms.TextBox();
            this.TargetFileBtn = new System.Windows.Forms.Button();
            this.OKBtn = new System.Windows.Forms.Button();
            this.CancelBtn = new System.Windows.Forms.Button();
            this.radioButton_Rows = new System.Windows.Forms.RadioButton();
            this.radioButton_Columns = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBox_KeyColumn1 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox_KeyColumn2 = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // SourceFileLbl
            // 
            this.SourceFileLbl.AutoSize = true;
            this.SourceFileLbl.Location = new System.Drawing.Point(48, 111);
            this.SourceFileLbl.Name = "SourceFileLbl";
            this.SourceFileLbl.Size = new System.Drawing.Size(144, 17);
            this.SourceFileLbl.TabIndex = 0;
            this.SourceFileLbl.Text = "Source MS Excel File:";
            this.SourceFileLbl.Click += new System.EventHandler(this.label1_Click);
            // 
            // SourceFileBtn
            // 
            this.SourceFileBtn.Location = new System.Drawing.Point(51, 40);
            this.SourceFileBtn.Name = "SourceFileBtn";
            this.SourceFileBtn.Size = new System.Drawing.Size(212, 49);
            this.SourceFileBtn.TabIndex = 1;
            this.SourceFileBtn.Text = "Open File";
            this.SourceFileBtn.UseVisualStyleBackColor = true;
            this.SourceFileBtn.Click += new System.EventHandler(this.SourceFileBtn_Click);
            // 
            // SourceFileCheckedListBox
            // 
            this.SourceFileCheckedListBox.FormattingEnabled = true;
            this.SourceFileCheckedListBox.Location = new System.Drawing.Point(51, 160);
            this.SourceFileCheckedListBox.Name = "SourceFileCheckedListBox";
            this.SourceFileCheckedListBox.Size = new System.Drawing.Size(656, 157);
            this.SourceFileCheckedListBox.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(48, 422);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(137, 17);
            this.label2.TabIndex = 3;
            this.label2.Text = "Target MS Excel File";
            // 
            // TargFileTextbox
            // 
            this.TargFileTextbox.Location = new System.Drawing.Point(51, 469);
            this.TargFileTextbox.Name = "TargFileTextbox";
            this.TargFileTextbox.Size = new System.Drawing.Size(656, 22);
            this.TargFileTextbox.TabIndex = 4;
            this.TargFileTextbox.TextChanged += new System.EventHandler(this.TargFileTextbox_TextChanged);
            // 
            // TargetFileBtn
            // 
            this.TargetFileBtn.Location = new System.Drawing.Point(51, 356);
            this.TargetFileBtn.Name = "TargetFileBtn";
            this.TargetFileBtn.Size = new System.Drawing.Size(212, 49);
            this.TargetFileBtn.TabIndex = 5;
            this.TargetFileBtn.Text = "Target File";
            this.TargetFileBtn.UseVisualStyleBackColor = true;
            this.TargetFileBtn.Click += new System.EventHandler(this.TargetFileBtn_Click);
            // 
            // OKBtn
            // 
            this.OKBtn.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.OKBtn.Location = new System.Drawing.Point(723, 542);
            this.OKBtn.Name = "OKBtn";
            this.OKBtn.Size = new System.Drawing.Size(145, 41);
            this.OKBtn.TabIndex = 6;
            this.OKBtn.Text = "OK";
            this.OKBtn.UseVisualStyleBackColor = true;
            this.OKBtn.Click += new System.EventHandler(this.OKBtn_Click);
            // 
            // CancelBtn
            // 
            this.CancelBtn.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CancelBtn.Location = new System.Drawing.Point(991, 542);
            this.CancelBtn.Name = "CancelBtn";
            this.CancelBtn.Size = new System.Drawing.Size(132, 41);
            this.CancelBtn.TabIndex = 7;
            this.CancelBtn.Text = "Cancel";
            this.CancelBtn.UseVisualStyleBackColor = true;
            this.CancelBtn.Click += new System.EventHandler(this.CancelBtn_Click);
            // 
            // radioButton_Rows
            // 
            this.radioButton_Rows.AutoSize = true;
            this.radioButton_Rows.Location = new System.Drawing.Point(6, 96);
            this.radioButton_Rows.Name = "radioButton_Rows";
            this.radioButton_Rows.Size = new System.Drawing.Size(107, 21);
            this.radioButton_Rows.TabIndex = 8;
            this.radioButton_Rows.Text = "Merge Rows";
            this.radioButton_Rows.UseVisualStyleBackColor = true;
            this.radioButton_Rows.CheckedChanged += new System.EventHandler(this.radioButton_Rows_CheckedChanged);
            // 
            // radioButton_Columns
            // 
            this.radioButton_Columns.AutoSize = true;
            this.radioButton_Columns.Checked = true;
            this.radioButton_Columns.Location = new System.Drawing.Point(6, 45);
            this.radioButton_Columns.Name = "radioButton_Columns";
            this.radioButton_Columns.Size = new System.Drawing.Size(127, 21);
            this.radioButton_Columns.TabIndex = 8;
            this.radioButton_Columns.TabStop = true;
            this.radioButton_Columns.Text = "Merge Columns";
            this.radioButton_Columns.UseVisualStyleBackColor = true;
            this.radioButton_Columns.CheckedChanged += new System.EventHandler(this.radioButton_Columns_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButton_Rows);
            this.groupBox1.Controls.Add(this.radioButton_Columns);
            this.groupBox1.Location = new System.Drawing.Point(713, 373);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(327, 133);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Merge Rows or Columns";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(713, 160);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(178, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "Key Column : Source File 1";
            // 
            // textBox_KeyColumn1
            // 
            this.textBox_KeyColumn1.Location = new System.Drawing.Point(713, 189);
            this.textBox_KeyColumn1.Name = "textBox_KeyColumn1";
            this.textBox_KeyColumn1.Size = new System.Drawing.Size(437, 22);
            this.textBox_KeyColumn1.TabIndex = 4;
            this.textBox_KeyColumn1.Text = "NationalID";
            this.textBox_KeyColumn1.TextChanged += new System.EventHandler(this.KeyColumnTextbox_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(713, 238);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(182, 17);
            this.label3.TabIndex = 3;
            this.label3.Text = "Key Column : Source File  2";
            // 
            // textBox_KeyColumn2
            // 
            this.textBox_KeyColumn2.Location = new System.Drawing.Point(713, 280);
            this.textBox_KeyColumn2.Name = "textBox_KeyColumn2";
            this.textBox_KeyColumn2.Size = new System.Drawing.Size(437, 22);
            this.textBox_KeyColumn2.TabIndex = 4;
            this.textBox_KeyColumn2.Text = "NationalID";
            this.textBox_KeyColumn2.TextChanged += new System.EventHandler(this.KeyColumnTextbox_TextChanged);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1162, 615);
            this.Controls.Add(this.CancelBtn);
            this.Controls.Add(this.OKBtn);
            this.Controls.Add(this.TargetFileBtn);
            this.Controls.Add(this.textBox_KeyColumn2);
            this.Controls.Add(this.textBox_KeyColumn1);
            this.Controls.Add(this.TargFileTextbox);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.SourceFileCheckedListBox);
            this.Controls.Add(this.SourceFileBtn);
            this.Controls.Add(this.SourceFileLbl);
            this.Controls.Add(this.groupBox1);
            this.Name = "MainForm";
            this.Text = "Excel Merger";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label SourceFileLbl;
        private System.Windows.Forms.Button SourceFileBtn;
        private System.Windows.Forms.CheckedListBox SourceFileCheckedListBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TargFileTextbox;
        private System.Windows.Forms.Button TargetFileBtn;
        private System.Windows.Forms.Button OKBtn;
        private System.Windows.Forms.Button CancelBtn;
        private System.Windows.Forms.RadioButton radioButton_Rows;
        private System.Windows.Forms.RadioButton radioButton_Columns;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBox_KeyColumn1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox_KeyColumn2;
    }
}

