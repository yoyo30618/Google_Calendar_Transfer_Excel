namespace Google_Calendar_Transfer_Excel
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            btnExport = new Button();
            dateTimePickerStart = new DateTimePicker();
            dateTimePickerEnd = new DateTimePicker();
            Btn_Login = new Button();
            comboBoxCalendars = new ComboBox();
            Chk_DefaultRules = new CheckedListBox();
            Rbt_Default = new RadioButton();
            Rbt_Spec = new RadioButton();
            label1 = new Label();
            Tbx_SpecialRule = new RichTextBox();
            label2 = new Label();
            label3 = new Label();
            label4 = new Label();
            button1 = new Button();
            SuspendLayout();
            // 
            // btnExport
            // 
            btnExport.Enabled = false;
            btnExport.Font = new Font("標楷體", 12F);
            btnExport.Location = new Point(383, 493);
            btnExport.Name = "btnExport";
            btnExport.Size = new Size(189, 45);
            btnExport.TabIndex = 0;
            btnExport.Text = "匯出日曆內容到EXCEL";
            btnExport.UseVisualStyleBackColor = true;
            btnExport.Click += btnExport_Click;
            // 
            // dateTimePickerStart
            // 
            dateTimePickerStart.Enabled = false;
            dateTimePickerStart.Font = new Font("標楷體", 12F);
            dateTimePickerStart.Location = new Point(31, 100);
            dateTimePickerStart.Name = "dateTimePickerStart";
            dateTimePickerStart.Size = new Size(239, 27);
            dateTimePickerStart.TabIndex = 1;
            // 
            // dateTimePickerEnd
            // 
            dateTimePickerEnd.Enabled = false;
            dateTimePickerEnd.Font = new Font("標楷體", 12F);
            dateTimePickerEnd.Location = new Point(309, 100);
            dateTimePickerEnd.Name = "dateTimePickerEnd";
            dateTimePickerEnd.Size = new Size(239, 27);
            dateTimePickerEnd.TabIndex = 2;
            // 
            // Btn_Login
            // 
            Btn_Login.Font = new Font("標楷體", 12F);
            Btn_Login.Location = new Point(31, 12);
            Btn_Login.Name = "Btn_Login";
            Btn_Login.Size = new Size(121, 50);
            Btn_Login.TabIndex = 3;
            Btn_Login.Text = "登入帳戶";
            Btn_Login.UseVisualStyleBackColor = true;
            Btn_Login.Click += Btn_Login_Click;
            // 
            // comboBoxCalendars
            // 
            comboBoxCalendars.Enabled = false;
            comboBoxCalendars.Font = new Font("標楷體", 12F);
            comboBoxCalendars.FormattingEnabled = true;
            comboBoxCalendars.Location = new Point(200, 38);
            comboBoxCalendars.Name = "comboBoxCalendars";
            comboBoxCalendars.Size = new Size(236, 24);
            comboBoxCalendars.TabIndex = 4;
            // 
            // Chk_DefaultRules
            // 
            Chk_DefaultRules.Enabled = false;
            Chk_DefaultRules.Font = new Font("標楷體", 12F);
            Chk_DefaultRules.FormattingEnabled = true;
            Chk_DefaultRules.Location = new Point(31, 163);
            Chk_DefaultRules.Name = "Chk_DefaultRules";
            Chk_DefaultRules.Size = new Size(240, 180);
            Chk_DefaultRules.TabIndex = 5;
            // 
            // Rbt_Default
            // 
            Rbt_Default.AutoSize = true;
            Rbt_Default.Checked = true;
            Rbt_Default.Enabled = false;
            Rbt_Default.Font = new Font("標楷體", 12F);
            Rbt_Default.Location = new Point(31, 133);
            Rbt_Default.Name = "Rbt_Default";
            Rbt_Default.Size = new Size(121, 20);
            Rbt_Default.TabIndex = 7;
            Rbt_Default.TabStop = true;
            Rbt_Default.Text = "月曆固定欄位";
            Rbt_Default.UseVisualStyleBackColor = true;
            Rbt_Default.CheckedChanged += radioButton1_CheckedChanged;
            // 
            // Rbt_Spec
            // 
            Rbt_Spec.AutoSize = true;
            Rbt_Spec.Enabled = false;
            Rbt_Spec.Font = new Font("標楷體", 12F);
            Rbt_Spec.Location = new Point(299, 135);
            Rbt_Spec.Name = "Rbt_Spec";
            Rbt_Spec.Size = new Size(137, 20);
            Rbt_Spec.TabIndex = 8;
            Rbt_Spec.Text = "自定義匯出格式";
            Rbt_Spec.UseVisualStyleBackColor = true;
            Rbt_Spec.CheckedChanged += radioButton2_CheckedChanged;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("標楷體", 12F);
            label1.Location = new Point(12, 362);
            label1.Name = "label1";
            label1.Size = new Size(487, 176);
            label1.TabIndex = 6;
            label1.Text = resources.GetString("label1.Text");
            // 
            // Tbx_SpecialRule
            // 
            Tbx_SpecialRule.Enabled = false;
            Tbx_SpecialRule.Font = new Font("標楷體", 12F);
            Tbx_SpecialRule.Location = new Point(309, 161);
            Tbx_SpecialRule.Name = "Tbx_SpecialRule";
            Tbx_SpecialRule.Size = new Size(239, 184);
            Tbx_SpecialRule.TabIndex = 10;
            Tbx_SpecialRule.Text = "";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("標楷體", 12F);
            label2.Location = new Point(282, 107);
            label2.Name = "label2";
            label2.Size = new Size(15, 16);
            label2.TabIndex = 11;
            label2.Text = "~";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("標楷體", 12F);
            label3.Location = new Point(31, 81);
            label3.Name = "label3";
            label3.Size = new Size(103, 16);
            label3.TabIndex = 12;
            label3.Text = "匯出日期區間";
            // 
            // label4
            // 
            label4.AutoSize = true;
            label4.Font = new Font("標楷體", 12F);
            label4.Location = new Point(200, 12);
            label4.Name = "label4";
            label4.Size = new Size(71, 16);
            label4.TabIndex = 13;
            label4.Text = "選擇月曆";
            // 
            // button1
            // 
            button1.Enabled = false;
            button1.Font = new Font("標楷體", 12F);
            button1.Location = new Point(522, 12);
            button1.Name = "button1";
            button1.Size = new Size(50, 30);
            button1.TabIndex = 14;
            button1.Text = "登出";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(584, 567);
            Controls.Add(button1);
            Controls.Add(label4);
            Controls.Add(label3);
            Controls.Add(label2);
            Controls.Add(Tbx_SpecialRule);
            Controls.Add(Rbt_Spec);
            Controls.Add(Rbt_Default);
            Controls.Add(Chk_DefaultRules);
            Controls.Add(comboBoxCalendars);
            Controls.Add(Btn_Login);
            Controls.Add(dateTimePickerEnd);
            Controls.Add(dateTimePickerStart);
            Controls.Add(btnExport);
            Controls.Add(label1);
            Name = "Form1";
            Text = "Form1";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnExport;
        private DateTimePicker dateTimePickerStart;
        private DateTimePicker dateTimePickerEnd;
        private Button Btn_Login;
        private ComboBox comboBoxCalendars;
        private CheckedListBox Chk_DefaultRules;
        private RadioButton Rbt_Default;
        private RadioButton Rbt_Spec;
        private Label label1;
        private RichTextBox Tbx_SpecialRule;
        private Label label2;
        private Label label3;
        private Label label4;
        private Button button1;
    }
}
