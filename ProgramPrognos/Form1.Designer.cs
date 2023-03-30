
namespace ProgramPrognos
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
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.Quitbutton = new System.Windows.Forms.Button();
            this.readdatabutton = new System.Windows.Forms.Button();
            this.businessbutton = new System.Windows.Forms.Button();
            this.proddisplaybutton = new System.Windows.Forms.Button();
            this.loadscenariobutton = new System.Windows.Forms.Button();
            this.savescenariobutton = new System.Windows.Forms.Button();
            this.datafolderlabel = new System.Windows.Forms.Label();
            this.docfolderlabel = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.examtestbutton = new System.Windows.Forms.Button();
            this.buttonExamforecast = new System.Windows.Forms.Button();
            this.TBxrounds = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.RBnursingexam = new System.Windows.Forms.RadioButton();
            this.RBteacherexam = new System.Windows.Forms.RadioButton();
            this.RBallexam = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.CBFK = new System.Windows.Forms.CheckBox();
            this.CBshortprogram = new System.Windows.Forms.CheckBox();
            this.CBlongprogram = new System.Windows.Forms.CheckBox();
            this.CBfutureadm = new System.Windows.Forms.CheckBox();
            this.TB_moneylimit = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.TB_endyear = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.AppRegButton = new System.Windows.Forms.Button();
            this.FKbutton = new System.Windows.Forms.Button();
            this.HSTbutton = new System.Windows.Forms.Button();
            this.Excelplanningbutton = new System.Windows.Forms.Button();
            this.coursedatabutton = new System.Windows.Forms.Button();
            this.testbutton = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(18, 25);
            this.richTextBox1.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(505, 612);
            this.richTextBox1.TabIndex = 0;
            this.richTextBox1.Text = "";
            // 
            // Quitbutton
            // 
            this.Quitbutton.Location = new System.Drawing.Point(924, 782);
            this.Quitbutton.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.Quitbutton.Name = "Quitbutton";
            this.Quitbutton.Size = new System.Drawing.Size(153, 80);
            this.Quitbutton.TabIndex = 1;
            this.Quitbutton.Text = "Quit";
            this.Quitbutton.UseVisualStyleBackColor = true;
            this.Quitbutton.Click += new System.EventHandler(this.Quitbutton_Click);
            // 
            // readdatabutton
            // 
            this.readdatabutton.Location = new System.Drawing.Point(924, 702);
            this.readdatabutton.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.readdatabutton.Name = "readdatabutton";
            this.readdatabutton.Size = new System.Drawing.Size(153, 74);
            this.readdatabutton.TabIndex = 2;
            this.readdatabutton.Text = "Read data";
            this.readdatabutton.UseVisualStyleBackColor = true;
            this.readdatabutton.Click += new System.EventHandler(this.readdatabutton_Click);
            // 
            // businessbutton
            // 
            this.businessbutton.Enabled = false;
            this.businessbutton.Location = new System.Drawing.Point(924, 571);
            this.businessbutton.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.businessbutton.Name = "businessbutton";
            this.businessbutton.Size = new System.Drawing.Size(153, 60);
            this.businessbutton.TabIndex = 3;
            this.businessbutton.Text = "Extrapolate business as usual";
            this.businessbutton.UseVisualStyleBackColor = true;
            this.businessbutton.Click += new System.EventHandler(this.businessbutton_Click);
            // 
            // proddisplaybutton
            // 
            this.proddisplaybutton.Enabled = false;
            this.proddisplaybutton.Location = new System.Drawing.Point(924, 402);
            this.proddisplaybutton.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.proddisplaybutton.Name = "proddisplaybutton";
            this.proddisplaybutton.Size = new System.Drawing.Size(153, 71);
            this.proddisplaybutton.TabIndex = 4;
            this.proddisplaybutton.Text = "Production per institution";
            this.proddisplaybutton.UseVisualStyleBackColor = true;
            this.proddisplaybutton.Click += new System.EventHandler(this.button1_Click);
            // 
            // loadscenariobutton
            // 
            this.loadscenariobutton.Location = new System.Drawing.Point(924, 638);
            this.loadscenariobutton.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.loadscenariobutton.Name = "loadscenariobutton";
            this.loadscenariobutton.Size = new System.Drawing.Size(153, 55);
            this.loadscenariobutton.TabIndex = 5;
            this.loadscenariobutton.Text = "Load scenario";
            this.loadscenariobutton.UseVisualStyleBackColor = true;
            this.loadscenariobutton.Click += new System.EventHandler(this.loadscenariobutton_Click);
            // 
            // savescenariobutton
            // 
            this.savescenariobutton.Enabled = false;
            this.savescenariobutton.Location = new System.Drawing.Point(924, 505);
            this.savescenariobutton.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.savescenariobutton.Name = "savescenariobutton";
            this.savescenariobutton.Size = new System.Drawing.Size(153, 60);
            this.savescenariobutton.TabIndex = 6;
            this.savescenariobutton.Text = "Save scenario";
            this.savescenariobutton.UseVisualStyleBackColor = true;
            this.savescenariobutton.Click += new System.EventHandler(this.savescenariobutton_Click);
            // 
            // datafolderlabel
            // 
            this.datafolderlabel.AutoSize = true;
            this.datafolderlabel.Location = new System.Drawing.Point(555, 37);
            this.datafolderlabel.Name = "datafolderlabel";
            this.datafolderlabel.Size = new System.Drawing.Size(51, 20);
            this.datafolderlabel.TabIndex = 8;
            this.datafolderlabel.Text = "label2";
            // 
            // docfolderlabel
            // 
            this.docfolderlabel.AutoSize = true;
            this.docfolderlabel.Location = new System.Drawing.Point(555, 69);
            this.docfolderlabel.Name = "docfolderlabel";
            this.docfolderlabel.Size = new System.Drawing.Size(51, 20);
            this.docfolderlabel.TabIndex = 9;
            this.docfolderlabel.Text = "label3";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // examtestbutton
            // 
            this.examtestbutton.Enabled = false;
            this.examtestbutton.Location = new System.Drawing.Point(924, 318);
            this.examtestbutton.Name = "examtestbutton";
            this.examtestbutton.Size = new System.Drawing.Size(153, 60);
            this.examtestbutton.TabIndex = 10;
            this.examtestbutton.Text = "Test exam forecast";
            this.examtestbutton.UseVisualStyleBackColor = true;
            this.examtestbutton.Click += new System.EventHandler(this.examtestbutton_Click);
            // 
            // buttonExamforecast
            // 
            this.buttonExamforecast.Enabled = false;
            this.buttonExamforecast.Location = new System.Drawing.Point(924, 231);
            this.buttonExamforecast.Name = "buttonExamforecast";
            this.buttonExamforecast.Size = new System.Drawing.Size(153, 62);
            this.buttonExamforecast.TabIndex = 11;
            this.buttonExamforecast.Text = "Exam forecast per program";
            this.buttonExamforecast.UseVisualStyleBackColor = true;
            this.buttonExamforecast.Click += new System.EventHandler(this.buttonExamforecast_Click);
            // 
            // TBxrounds
            // 
            this.TBxrounds.Location = new System.Drawing.Point(780, 582);
            this.TBxrounds.Name = "TBxrounds";
            this.TBxrounds.Size = new System.Drawing.Size(100, 26);
            this.TBxrounds.TabIndex = 12;
            this.TBxrounds.Text = "20";
            this.TBxrounds.TextChanged += new System.EventHandler(this.TBxrounds_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(748, 558);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(159, 20);
            this.label1.TabIndex = 13;
            this.label1.Text = "Extrapolation rounds:";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.RBnursingexam);
            this.groupBox1.Controls.Add(this.RBteacherexam);
            this.groupBox1.Controls.Add(this.RBallexam);
            this.groupBox1.Location = new System.Drawing.Point(718, 214);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 122);
            this.groupBox1.TabIndex = 14;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Exam area";
            // 
            // RBnursingexam
            // 
            this.RBnursingexam.AutoSize = true;
            this.RBnursingexam.Location = new System.Drawing.Point(6, 85);
            this.RBnursingexam.Name = "RBnursingexam";
            this.RBnursingexam.Size = new System.Drawing.Size(138, 24);
            this.RBnursingexam.TabIndex = 2;
            this.RBnursingexam.TabStop = true;
            this.RBnursingexam.Text = "Nursing exams";
            this.RBnursingexam.UseVisualStyleBackColor = true;
            // 
            // RBteacherexam
            // 
            this.RBteacherexam.AutoSize = true;
            this.RBteacherexam.Location = new System.Drawing.Point(6, 55);
            this.RBteacherexam.Name = "RBteacherexam";
            this.RBteacherexam.Size = new System.Drawing.Size(142, 24);
            this.RBteacherexam.TabIndex = 1;
            this.RBteacherexam.TabStop = true;
            this.RBteacherexam.Text = "Teacher exams";
            this.RBteacherexam.UseVisualStyleBackColor = true;
            // 
            // RBallexam
            // 
            this.RBallexam.AutoSize = true;
            this.RBallexam.Checked = true;
            this.RBallexam.Location = new System.Drawing.Point(6, 25);
            this.RBallexam.Name = "RBallexam";
            this.RBallexam.Size = new System.Drawing.Size(51, 24);
            this.RBallexam.TabIndex = 0;
            this.RBallexam.TabStop = true;
            this.RBallexam.Text = "All";
            this.RBallexam.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.CBFK);
            this.groupBox2.Controls.Add(this.CBshortprogram);
            this.groupBox2.Controls.Add(this.CBlongprogram);
            this.groupBox2.Location = new System.Drawing.Point(718, 342);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(200, 112);
            this.groupBox2.TabIndex = 15;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Type of study";
            // 
            // CBFK
            // 
            this.CBFK.AutoSize = true;
            this.CBFK.Checked = true;
            this.CBFK.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CBFK.Location = new System.Drawing.Point(6, 85);
            this.CBFK.Name = "CBFK";
            this.CBFK.Size = new System.Drawing.Size(55, 24);
            this.CBFK.TabIndex = 2;
            this.CBFK.Text = "FK";
            this.CBFK.UseVisualStyleBackColor = true;
            // 
            // CBshortprogram
            // 
            this.CBshortprogram.AutoSize = true;
            this.CBshortprogram.Checked = true;
            this.CBshortprogram.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CBshortprogram.Location = new System.Drawing.Point(6, 55);
            this.CBshortprogram.Name = "CBshortprogram";
            this.CBshortprogram.Size = new System.Drawing.Size(145, 24);
            this.CBshortprogram.TabIndex = 1;
            this.CBshortprogram.Text = "Program (short)";
            this.CBshortprogram.UseVisualStyleBackColor = true;
            // 
            // CBlongprogram
            // 
            this.CBlongprogram.AutoSize = true;
            this.CBlongprogram.Checked = true;
            this.CBlongprogram.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CBlongprogram.Location = new System.Drawing.Point(6, 25);
            this.CBlongprogram.Name = "CBlongprogram";
            this.CBlongprogram.Size = new System.Drawing.Size(139, 24);
            this.CBlongprogram.TabIndex = 0;
            this.CBlongprogram.Text = "Program (long)";
            this.CBlongprogram.UseVisualStyleBackColor = true;
            // 
            // CBfutureadm
            // 
            this.CBfutureadm.AutoSize = true;
            this.CBfutureadm.Checked = true;
            this.CBfutureadm.CheckState = System.Windows.Forms.CheckState.Checked;
            this.CBfutureadm.Location = new System.Drawing.Point(702, 612);
            this.CBfutureadm.Name = "CBfutureadm";
            this.CBfutureadm.Size = new System.Drawing.Size(216, 24);
            this.CBfutureadm.TabIndex = 16;
            this.CBfutureadm.Text = "Include future admissions";
            this.CBfutureadm.UseVisualStyleBackColor = true;
            // 
            // TB_moneylimit
            // 
            this.TB_moneylimit.Location = new System.Drawing.Point(752, 458);
            this.TB_moneylimit.Name = "TB_moneylimit";
            this.TB_moneylimit.Size = new System.Drawing.Size(145, 26);
            this.TB_moneylimit.TabIndex = 17;
            this.TB_moneylimit.Text = "1000000";
            this.TB_moneylimit.TextChanged += new System.EventHandler(this.textBox1_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(555, 462);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(191, 20);
            this.label2.TabIndex = 18;
            this.label2.Text = "Minibelopp särredovisning";
            // 
            // TB_endyear
            // 
            this.TB_endyear.Location = new System.Drawing.Point(774, 682);
            this.TB_endyear.Name = "TB_endyear";
            this.TB_endyear.Size = new System.Drawing.Size(100, 26);
            this.TB_endyear.TabIndex = 19;
            this.TB_endyear.Text = "2028";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(698, 685);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 20);
            this.label3.TabIndex = 20;
            this.label3.Text = "End year";
            // 
            // AppRegButton
            // 
            this.AppRegButton.Location = new System.Drawing.Point(924, 162);
            this.AppRegButton.Name = "AppRegButton";
            this.AppRegButton.Size = new System.Drawing.Size(154, 52);
            this.AppRegButton.TabIndex = 21;
            this.AppRegButton.Text = "Applicant to registered";
            this.AppRegButton.UseVisualStyleBackColor = true;
            this.AppRegButton.Click += new System.EventHandler(this.AppRegButton_Click);
            // 
            // FKbutton
            // 
            this.FKbutton.Location = new System.Drawing.Point(18, 682);
            this.FKbutton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.FKbutton.Name = "FKbutton";
            this.FKbutton.Size = new System.Drawing.Size(153, 60);
            this.FKbutton.TabIndex = 22;
            this.FKbutton.Text = "Analyze course data";
            this.FKbutton.UseVisualStyleBackColor = true;
            this.FKbutton.Click += new System.EventHandler(this.FKbutton_Click);
            // 
            // HSTbutton
            // 
            this.HSTbutton.Location = new System.Drawing.Point(20, 752);
            this.HSTbutton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.HSTbutton.Name = "HSTbutton";
            this.HSTbutton.Size = new System.Drawing.Size(178, 57);
            this.HSTbutton.TabIndex = 23;
            this.HSTbutton.Text = "Read hst_hpr_utfall_budget";
            this.HSTbutton.UseVisualStyleBackColor = true;
            this.HSTbutton.Click += new System.EventHandler(this.HSTbutton_Click);
            // 
            // Excelplanningbutton
            // 
            this.Excelplanningbutton.Location = new System.Drawing.Point(924, 25);
            this.Excelplanningbutton.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Excelplanningbutton.Name = "Excelplanningbutton";
            this.Excelplanningbutton.Size = new System.Drawing.Size(153, 52);
            this.Excelplanningbutton.TabIndex = 24;
            this.Excelplanningbutton.Text = "Make Excel planning file";
            this.Excelplanningbutton.UseVisualStyleBackColor = true;
            this.Excelplanningbutton.Click += new System.EventHandler(this.Excelplanningbutton_Click);
            // 
            // coursedatabutton
            // 
            this.coursedatabutton.Location = new System.Drawing.Point(924, 86);
            this.coursedatabutton.Name = "coursedatabutton";
            this.coursedatabutton.Size = new System.Drawing.Size(153, 52);
            this.coursedatabutton.TabIndex = 25;
            this.coursedatabutton.Text = "Read course data";
            this.coursedatabutton.UseVisualStyleBackColor = true;
            this.coursedatabutton.Click += new System.EventHandler(this.coursedatabutton_Click);
            // 
            // testbutton
            // 
            this.testbutton.Location = new System.Drawing.Point(226, 685);
            this.testbutton.Name = "testbutton";
            this.testbutton.Size = new System.Drawing.Size(124, 57);
            this.testbutton.TabIndex = 26;
            this.testbutton.Text = "Test stuff";
            this.testbutton.UseVisualStyleBackColor = true;
            this.testbutton.Click += new System.EventHandler(this.testbutton_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1341, 878);
            this.Controls.Add(this.testbutton);
            this.Controls.Add(this.coursedatabutton);
            this.Controls.Add(this.Excelplanningbutton);
            this.Controls.Add(this.HSTbutton);
            this.Controls.Add(this.FKbutton);
            this.Controls.Add(this.AppRegButton);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.TB_endyear);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.TB_moneylimit);
            this.Controls.Add(this.CBfutureadm);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.TBxrounds);
            this.Controls.Add(this.buttonExamforecast);
            this.Controls.Add(this.examtestbutton);
            this.Controls.Add(this.docfolderlabel);
            this.Controls.Add(this.datafolderlabel);
            this.Controls.Add(this.savescenariobutton);
            this.Controls.Add(this.loadscenariobutton);
            this.Controls.Add(this.proddisplaybutton);
            this.Controls.Add(this.businessbutton);
            this.Controls.Add(this.readdatabutton);
            this.Controls.Add(this.Quitbutton);
            this.Controls.Add(this.richTextBox1);
            this.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.Name = "Form1";
            this.Text = "Form1";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button Quitbutton;
        private System.Windows.Forms.Button readdatabutton;
        private System.Windows.Forms.Button businessbutton;
        private System.Windows.Forms.Button proddisplaybutton;
        private System.Windows.Forms.Button loadscenariobutton;
        private System.Windows.Forms.Button savescenariobutton;
        private System.Windows.Forms.Label datafolderlabel;
        private System.Windows.Forms.Label docfolderlabel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button examtestbutton;
        private System.Windows.Forms.Button buttonExamforecast;
        private System.Windows.Forms.TextBox TBxrounds;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton RBnursingexam;
        private System.Windows.Forms.RadioButton RBteacherexam;
        private System.Windows.Forms.RadioButton RBallexam;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox CBFK;
        private System.Windows.Forms.CheckBox CBshortprogram;
        private System.Windows.Forms.CheckBox CBlongprogram;
        private System.Windows.Forms.CheckBox CBfutureadm;
        private System.Windows.Forms.TextBox TB_moneylimit;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox TB_endyear;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button AppRegButton;
        private System.Windows.Forms.Button FKbutton;
        private System.Windows.Forms.Button HSTbutton;
        private System.Windows.Forms.Button Excelplanningbutton;
        private System.Windows.Forms.Button coursedatabutton;
        private System.Windows.Forms.Button testbutton;
    }
}

