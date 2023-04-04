
namespace ProgramPrognos
{
    partial class ExcelForm
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
            this.Excelbutton = new System.Windows.Forms.Button();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.TBlastsem = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.RB_minorprog = new System.Windows.Forms.RadioButton();
            this.RB_homeinst = new System.Windows.Forms.RadioButton();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Excelbutton
            // 
            this.Excelbutton.Location = new System.Drawing.Point(980, 300);
            this.Excelbutton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Excelbutton.Name = "Excelbutton";
            this.Excelbutton.Size = new System.Drawing.Size(174, 74);
            this.Excelbutton.TabIndex = 0;
            this.Excelbutton.Text = "Make Excel files";
            this.Excelbutton.UseVisualStyleBackColor = true;
            this.Excelbutton.Click += new System.EventHandler(this.Excelbutton_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(52, 46);
            this.richTextBox1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(584, 426);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            // 
            // TBlastsem
            // 
            this.TBlastsem.Location = new System.Drawing.Point(1070, 172);
            this.TBlastsem.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.TBlastsem.Name = "TBlastsem";
            this.TBlastsem.Size = new System.Drawing.Size(82, 26);
            this.TBlastsem.TabIndex = 2;
            this.TBlastsem.Text = "VT23";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(879, 177);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(182, 20);
            this.label1.TabIndex = 3;
            this.label1.Text = "Last semester with data:";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.RB_minorprog);
            this.panel1.Controls.Add(this.RB_homeinst);
            this.panel1.Location = new System.Drawing.Point(894, 420);
            this.panel1.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(300, 83);
            this.panel1.TabIndex = 4;
            // 
            // RB_minorprog
            // 
            this.RB_minorprog.AutoSize = true;
            this.RB_minorprog.Location = new System.Drawing.Point(4, 40);
            this.RB_minorprog.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.RB_minorprog.Name = "RB_minorprog";
            this.RB_minorprog.Size = new System.Drawing.Size(243, 24);
            this.RB_minorprog.TabIndex = 1;
            this.RB_minorprog.Text = "Also programs with minor part";
            this.RB_minorprog.UseVisualStyleBackColor = true;
            // 
            // RB_homeinst
            // 
            this.RB_homeinst.AutoSize = true;
            this.RB_homeinst.Checked = true;
            this.RB_homeinst.Location = new System.Drawing.Point(4, 5);
            this.RB_homeinst.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.RB_homeinst.Name = "RB_homeinst";
            this.RB_homeinst.Size = new System.Drawing.Size(209, 24);
            this.RB_homeinst.TabIndex = 0;
            this.RB_homeinst.TabStop = true;
            this.RB_homeinst.Text = "Home inst programs only";
            this.RB_homeinst.UseVisualStyleBackColor = true;
            // 
            // ExcelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1200, 811);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TBlastsem);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.Excelbutton);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "ExcelForm";
            this.Text = "ExcelForm";
            this.Load += new System.EventHandler(this.ExcelForm_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button Excelbutton;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.TextBox TBlastsem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.RadioButton RB_minorprog;
        private System.Windows.Forms.RadioButton RB_homeinst;
    }
}