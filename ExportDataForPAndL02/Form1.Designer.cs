namespace ExportDataForPAndL02
{
    partial class Form1
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
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.btProcesss = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.button2 = new System.Windows.Forms.Button();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.label3 = new System.Windows.Forms.Label();
            this.ccbListMonth = new System.Windows.Forms.ComboBox();
            this.button4 = new System.Windows.Forms.Button();
            this.openFileDialog2 = new System.Windows.Forms.OpenFileDialog();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(143, 14);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(339, 20);
            this.textBox1.TabIndex = 12;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(12, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(109, 23);
            this.button1.TabIndex = 11;
            this.button1.Text = "From File";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // btProcesss
            // 
            this.btProcesss.Location = new System.Drawing.Point(12, 106);
            this.btProcesss.Name = "btProcesss";
            this.btProcesss.Size = new System.Drawing.Size(111, 23);
            this.btProcesss.TabIndex = 15;
            this.btProcesss.Text = "Processing";
            this.btProcesss.UseVisualStyleBackColor = true;
            this.btProcesss.Click += new System.EventHandler(this.button3_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(10, 55);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(111, 23);
            this.button2.TabIndex = 16;
            this.button2.Text = "To File";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(143, 57);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(339, 20);
            this.textBox2.TabIndex = 17;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 163);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(220, 13);
            this.label3.TabIndex = 20;
            this.label3.Text = "Select From Month Column in Destination File";
            // 
            // ccbListMonth
            // 
            this.ccbListMonth.FormattingEnabled = true;
            this.ccbListMonth.Location = new System.Drawing.Point(143, 194);
            this.ccbListMonth.Name = "ccbListMonth";
            this.ccbListMonth.Size = new System.Drawing.Size(182, 21);
            this.ccbListMonth.TabIndex = 22;
            this.ccbListMonth.SelectedIndexChanged += new System.EventHandler(this.ccbListMonth_SelectedIndexChanged);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(12, 238);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(124, 23);
            this.button4.TabIndex = 24;
            this.button4.Text = "Get Data";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // openFileDialog2
            // 
            this.openFileDialog2.FileName = "openFileDialog2";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(497, 273);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.ccbListMonth);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btProcesss);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Export Data For Consolidated P&L 02";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button btProcesss;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox textBox2;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox ccbListMonth;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.OpenFileDialog openFileDialog2;
    }
}

