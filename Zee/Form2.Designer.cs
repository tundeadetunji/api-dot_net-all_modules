namespace Zee
{
    partial class Form2
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
            this.label1 = new System.Windows.Forms.Label();
            this.titleDropDown = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.nameTextBox = new System.Windows.Forms.TextBox();
            this.genderDropDown = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.professionDropDown = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.yearsOfExperienceTextBox = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.resetButton = new System.Windows.Forms.Button();
            this.submitButton = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.summaryListBox = new System.Windows.Forms.ListBox();
            this.summaryTextBox = new System.Windows.Forms.TextBox();
            this.removeButton = new System.Windows.Forms.Button();
            this.upButton = new System.Windows.Forms.Button();
            this.clearButton = new System.Windows.Forms.Button();
            this.copyToListBoxButton = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 19);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 19);
            this.label1.TabIndex = 0;
            this.label1.Text = "Title";
            // 
            // titleDropDown
            // 
            this.titleDropDown.FormattingEnabled = true;
            this.titleDropDown.Location = new System.Drawing.Point(12, 41);
            this.titleDropDown.Name = "titleDropDown";
            this.titleDropDown.Size = new System.Drawing.Size(121, 27);
            this.titleDropDown.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 71);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(45, 19);
            this.label2.TabIndex = 2;
            this.label2.Text = "Name";
            // 
            // nameTextBox
            // 
            this.nameTextBox.Location = new System.Drawing.Point(12, 93);
            this.nameTextBox.Name = "nameTextBox";
            this.nameTextBox.Size = new System.Drawing.Size(214, 26);
            this.nameTextBox.TabIndex = 3;
            // 
            // genderDropDown
            // 
            this.genderDropDown.FormattingEnabled = true;
            this.genderDropDown.Location = new System.Drawing.Point(12, 144);
            this.genderDropDown.Name = "genderDropDown";
            this.genderDropDown.Size = new System.Drawing.Size(121, 27);
            this.genderDropDown.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 122);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 19);
            this.label3.TabIndex = 4;
            this.label3.Text = "Gender";
            // 
            // professionDropDown
            // 
            this.professionDropDown.FormattingEnabled = true;
            this.professionDropDown.Location = new System.Drawing.Point(12, 196);
            this.professionDropDown.Name = "professionDropDown";
            this.professionDropDown.Size = new System.Drawing.Size(121, 27);
            this.professionDropDown.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(12, 174);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(99, 19);
            this.label4.TabIndex = 6;
            this.label4.Text = "Profession";
            // 
            // yearsOfExperienceTextBox
            // 
            this.yearsOfExperienceTextBox.Location = new System.Drawing.Point(12, 248);
            this.yearsOfExperienceTextBox.Name = "yearsOfExperienceTextBox";
            this.yearsOfExperienceTextBox.Size = new System.Drawing.Size(214, 26);
            this.yearsOfExperienceTextBox.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 226);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(180, 19);
            this.label5.TabIndex = 8;
            this.label5.Text = "Years of Experience";
            // 
            // resetButton
            // 
            this.resetButton.Location = new System.Drawing.Point(16, 291);
            this.resetButton.Name = "resetButton";
            this.resetButton.Size = new System.Drawing.Size(75, 23);
            this.resetButton.TabIndex = 10;
            this.resetButton.Text = "RESET";
            this.resetButton.UseVisualStyleBackColor = true;
            this.resetButton.Click += new System.EventHandler(this.resetButton_Click);
            // 
            // submitButton
            // 
            this.submitButton.Location = new System.Drawing.Point(16, 320);
            this.submitButton.Name = "submitButton";
            this.submitButton.Size = new System.Drawing.Size(75, 23);
            this.submitButton.TabIndex = 11;
            this.submitButton.Text = "SUBMIT";
            this.submitButton.UseVisualStyleBackColor = true;
            this.submitButton.Click += new System.EventHandler(this.submitButton_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(334, 19);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 19);
            this.label6.TabIndex = 12;
            this.label6.Text = "Summary";
            // 
            // summaryListBox
            // 
            this.summaryListBox.FormattingEnabled = true;
            this.summaryListBox.ItemHeight = 19;
            this.summaryListBox.Location = new System.Drawing.Point(556, 41);
            this.summaryListBox.Name = "summaryListBox";
            this.summaryListBox.Size = new System.Drawing.Size(252, 289);
            this.summaryListBox.TabIndex = 13;
            // 
            // summaryTextBox
            // 
            this.summaryTextBox.Location = new System.Drawing.Point(336, 41);
            this.summaryTextBox.Multiline = true;
            this.summaryTextBox.Name = "summaryTextBox";
            this.summaryTextBox.Size = new System.Drawing.Size(214, 289);
            this.summaryTextBox.TabIndex = 14;
            // 
            // removeButton
            // 
            this.removeButton.Location = new System.Drawing.Point(814, 70);
            this.removeButton.Name = "removeButton";
            this.removeButton.Size = new System.Drawing.Size(75, 23);
            this.removeButton.TabIndex = 16;
            this.removeButton.Text = "REMOVE";
            this.removeButton.UseVisualStyleBackColor = true;
            this.removeButton.Click += new System.EventHandler(this.removeButton_Click);
            // 
            // upButton
            // 
            this.upButton.Location = new System.Drawing.Point(814, 41);
            this.upButton.Name = "upButton";
            this.upButton.Size = new System.Drawing.Size(75, 23);
            this.upButton.TabIndex = 15;
            this.upButton.Text = "UP";
            this.upButton.UseVisualStyleBackColor = true;
            this.upButton.Click += new System.EventHandler(this.upButton_Click);
            // 
            // clearButton
            // 
            this.clearButton.Location = new System.Drawing.Point(814, 99);
            this.clearButton.Name = "clearButton";
            this.clearButton.Size = new System.Drawing.Size(75, 23);
            this.clearButton.TabIndex = 17;
            this.clearButton.Text = "CLEAR";
            this.clearButton.UseVisualStyleBackColor = true;
            this.clearButton.Click += new System.EventHandler(this.clearButton_Click);
            // 
            // copyToListBoxButton
            // 
            this.copyToListBoxButton.Location = new System.Drawing.Point(336, 336);
            this.copyToListBoxButton.Name = "copyToListBoxButton";
            this.copyToListBoxButton.Size = new System.Drawing.Size(75, 23);
            this.copyToListBoxButton.TabIndex = 18;
            this.copyToListBoxButton.Text = "COPY";
            this.copyToListBoxButton.UseVisualStyleBackColor = true;
            this.copyToListBoxButton.Click += new System.EventHandler(this.copyToListBoxButton_Click);
            // 
            // Form2
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 19F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1025, 392);
            this.Controls.Add(this.copyToListBoxButton);
            this.Controls.Add(this.clearButton);
            this.Controls.Add(this.removeButton);
            this.Controls.Add(this.upButton);
            this.Controls.Add(this.summaryTextBox);
            this.Controls.Add(this.summaryListBox);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.submitButton);
            this.Controls.Add(this.resetButton);
            this.Controls.Add(this.yearsOfExperienceTextBox);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.professionDropDown);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.genderDropDown);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.nameTextBox);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.titleDropDown);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("JetBrains Mono", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Form2";
            this.Text = "Form2";
            this.Load += new System.EventHandler(this.Form2_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox titleDropDown;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox nameTextBox;
        private System.Windows.Forms.ComboBox genderDropDown;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox professionDropDown;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox yearsOfExperienceTextBox;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button resetButton;
        private System.Windows.Forms.Button submitButton;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ListBox summaryListBox;
        private System.Windows.Forms.TextBox summaryTextBox;
        private System.Windows.Forms.Button removeButton;
        private System.Windows.Forms.Button upButton;
        private System.Windows.Forms.Button clearButton;
        private System.Windows.Forms.Button copyToListBoxButton;
    }
}