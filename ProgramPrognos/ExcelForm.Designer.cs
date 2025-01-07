
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
            this.convertfolderbutton = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.coursecheckbutton = new System.Windows.Forms.Button();
            this.LBinst = new System.Windows.Forms.ListBox();
            this.Prognosbutton = new System.Windows.Forms.Button();
            this.CBsortprog = new System.Windows.Forms.CheckBox();
            this.CBspecialinput = new System.Windows.Forms.CheckBox();
            this.omstallningbutton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.TBstartyear = new System.Windows.Forms.TextBox();
            this.TBendyear = new System.Windows.Forms.TextBox();
            this.CBtriangel = new System.Windows.Forms.CheckBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // Excelbutton
            // 
            this.Excelbutton.Location = new System.Drawing.Point(653, 145);
            this.Excelbutton.Name = "Excelbutton";
            this.Excelbutton.Size = new System.Drawing.Size(116, 48);
            this.Excelbutton.TabIndex = 0;
            this.Excelbutton.Text = "Skapa Excel planeringsfil per institution";
            this.Excelbutton.UseVisualStyleBackColor = true;
            this.Excelbutton.Click += new System.EventHandler(this.Excelbutton_Click);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(35, 30);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(391, 278);
            this.richTextBox1.TabIndex = 1;
            this.richTextBox1.Text = "";
            // 
            // TBlastsem
            // 
            this.TBlastsem.Location = new System.Drawing.Point(713, 49);
            this.TBlastsem.Name = "TBlastsem";
            this.TBlastsem.Size = new System.Drawing.Size(56, 20);
            this.TBlastsem.TabIndex = 2;
            this.TBlastsem.Text = "VT24";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(586, 52);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Last semester with data:";
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.RB_minorprog);
            this.panel1.Controls.Add(this.RB_homeinst);
            this.panel1.Location = new System.Drawing.Point(596, 332);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(200, 54);
            this.panel1.TabIndex = 4;
            // 
            // RB_minorprog
            // 
            this.RB_minorprog.AutoSize = true;
            this.RB_minorprog.Checked = true;
            this.RB_minorprog.Location = new System.Drawing.Point(3, 26);
            this.RB_minorprog.Name = "RB_minorprog";
            this.RB_minorprog.Size = new System.Drawing.Size(162, 17);
            this.RB_minorprog.TabIndex = 1;
            this.RB_minorprog.TabStop = true;
            this.RB_minorprog.Text = "Also programs with minor part";
            this.RB_minorprog.UseVisualStyleBackColor = true;
            // 
            // RB_homeinst
            // 
            this.RB_homeinst.AutoSize = true;
            this.RB_homeinst.Location = new System.Drawing.Point(3, 3);
            this.RB_homeinst.Name = "RB_homeinst";
            this.RB_homeinst.Size = new System.Drawing.Size(140, 17);
            this.RB_homeinst.TabIndex = 0;
            this.RB_homeinst.Text = "Home inst programs only";
            this.RB_homeinst.UseVisualStyleBackColor = true;
            // 
            // convertfolderbutton
            // 
            this.convertfolderbutton.Location = new System.Drawing.Point(79, 448);
            this.convertfolderbutton.Name = "convertfolderbutton";
            this.convertfolderbutton.Size = new System.Drawing.Size(147, 41);
            this.convertfolderbutton.TabIndex = 5;
            this.convertfolderbutton.Text = "Convert excel files in folder";
            this.convertfolderbutton.UseVisualStyleBackColor = true;
            this.convertfolderbutton.Click += new System.EventHandler(this.convertfolderbutton_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // coursecheckbutton
            // 
            this.coursecheckbutton.Location = new System.Drawing.Point(79, 366);
            this.coursecheckbutton.Name = "coursecheckbutton";
            this.coursecheckbutton.Size = new System.Drawing.Size(147, 55);
            this.coursecheckbutton.TabIndex = 6;
            this.coursecheckbutton.Text = "Check courses against planning files";
            this.coursecheckbutton.UseVisualStyleBackColor = true;
            this.coursecheckbutton.Click += new System.EventHandler(this.coursecheckbutton_Click);
            // 
            // LBinst
            // 
            this.LBinst.FormattingEnabled = true;
            this.LBinst.Location = new System.Drawing.Point(547, 406);
            this.LBinst.Name = "LBinst";
            this.LBinst.SelectionMode = System.Windows.Forms.SelectionMode.MultiSimple;
            this.LBinst.Size = new System.Drawing.Size(249, 160);
            this.LBinst.TabIndex = 7;
            // 
            // Prognosbutton
            // 
            this.Prognosbutton.Location = new System.Drawing.Point(653, 84);
            this.Prognosbutton.Name = "Prognosbutton";
            this.Prognosbutton.Size = new System.Drawing.Size(116, 41);
            this.Prognosbutton.TabIndex = 8;
            this.Prognosbutton.Text = "Skapa gemensam prognosfil";
            this.Prognosbutton.UseVisualStyleBackColor = true;
            this.Prognosbutton.Click += new System.EventHandler(this.Prognosbutton_Click);
            // 
            // CBsortprog
            // 
            this.CBsortprog.AutoSize = true;
            this.CBsortprog.Checked = true;
            this.CBsortprog.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CBsortprog.Location = new System.Drawing.Point(612, 249);
            this.CBsortprog.Name = "CBsortprog";
            this.CBsortprog.Size = new System.Drawing.Size(101, 17);
            this.CBsortprog.TabIndex = 9;
            this.CBsortprog.Text = "Sort program list";
            this.CBsortprog.UseVisualStyleBackColor = true;
            // 
            // CBspecialinput
            // 
            this.CBspecialinput.AutoSize = true;
            this.CBspecialinput.Location = new System.Drawing.Point(612, 226);
            this.CBspecialinput.Name = "CBspecialinput";
            this.CBspecialinput.Size = new System.Drawing.Size(128, 17);
            this.CBspecialinput.TabIndex = 10;
            this.CBspecialinput.Text = "Read special data file";
            this.CBspecialinput.UseVisualStyleBackColor = true;
            // 
            // omstallningbutton
            // 
            this.omstallningbutton.Location = new System.Drawing.Point(232, 448);
            this.omstallningbutton.Name = "omstallningbutton";
            this.omstallningbutton.Size = new System.Drawing.Size(125, 41);
            this.omstallningbutton.TabIndex = 11;
            this.omstallningbutton.Text = "Läs omställningsstudiestöd";
            this.omstallningbutton.UseVisualStyleBackColor = true;
            this.omstallningbutton.Click += new System.EventHandler(this.omstallningbutton_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(496, 156);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(31, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "Från:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(499, 179);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(23, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "Till:";
            // 
            // TBstartyear
            // 
            this.TBstartyear.Location = new System.Drawing.Point(547, 148);
            this.TBstartyear.Name = "TBstartyear";
            this.TBstartyear.Size = new System.Drawing.Size(67, 20);
            this.TBstartyear.TabIndex = 14;
            this.TBstartyear.Text = "2023";
            // 
            // TBendyear
            // 
            this.TBendyear.Location = new System.Drawing.Point(547, 179);
            this.TBendyear.Name = "TBendyear";
            this.TBendyear.Size = new System.Drawing.Size(67, 20);
            this.TBendyear.TabIndex = 15;
            this.TBendyear.Text = "2030";
            // 
            // CBtriangel
            // 
            this.CBtriangel.AutoSize = true;
            this.CBtriangel.Checked = true;
            this.CBtriangel.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CBtriangel.Location = new System.Drawing.Point(612, 272);
            this.CBtriangel.Name = "CBtriangel";
            this.CBtriangel.Size = new System.Drawing.Size(105, 17);
            this.CBtriangel.TabIndex = 16;
            this.CBtriangel.Text = "Flik med trianglar";
            this.CBtriangel.UseVisualStyleBackColor = true;
            // 
            // ExcelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 578);
            this.Controls.Add(this.CBtriangel);
            this.Controls.Add(this.TBendyear);
            this.Controls.Add(this.TBstartyear);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.omstallningbutton);
            this.Controls.Add(this.CBspecialinput);
            this.Controls.Add(this.CBsortprog);
            this.Controls.Add(this.Prognosbutton);
            this.Controls.Add(this.LBinst);
            this.Controls.Add(this.coursecheckbutton);
            this.Controls.Add(this.convertfolderbutton);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TBlastsem);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.Excelbutton);
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
        private System.Windows.Forms.Button convertfolderbutton;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button coursecheckbutton;
        private System.Windows.Forms.ListBox LBinst;
        private System.Windows.Forms.Button Prognosbutton;
        private System.Windows.Forms.CheckBox CBsortprog;
        private System.Windows.Forms.CheckBox CBspecialinput;
        private System.Windows.Forms.Button omstallningbutton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox TBstartyear;
        private System.Windows.Forms.TextBox TBendyear;
        private System.Windows.Forms.CheckBox CBtriangel;
    }
}