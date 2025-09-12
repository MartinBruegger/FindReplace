namespace FindReplace
{
    partial class FormModFavorite
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
            this.materialLabel1 = new MaterialSkin.Controls.MaterialLabel();
            this.textBox_Description = new MaterialSkin.Controls.MaterialTextBox();
            this.textBox_Files = new MaterialSkin.Controls.MaterialTextBox();
            this.textBox_FindText = new MaterialSkin.Controls.MaterialTextBox();
            this.textBox_ReplaceText = new MaterialSkin.Controls.MaterialTextBox();
            this.textBox_Directory = new MaterialSkin.Controls.MaterialTextBox();
            this.checkbox_SubDirectory = new MaterialSkin.Controls.MaterialCheckbox();
            this.checkbox_ByDate = new MaterialSkin.Controls.MaterialCheckbox();
            this.numericUpDown_Days = new System.Windows.Forms.NumericUpDown();
            this.materialLabel2 = new MaterialSkin.Controls.MaterialLabel();
            this.checkbox_MatchCase = new MaterialSkin.Controls.MaterialCheckbox();
            this.checkbox_MatchWord = new MaterialSkin.Controls.MaterialCheckbox();
            this.checkbox_RegEx = new MaterialSkin.Controls.MaterialCheckbox();
            this.materialLabel3 = new MaterialSkin.Controls.MaterialLabel();
            this.materialButton_OK = new MaterialSkin.Controls.MaterialButton();
            this.materialButton_Cancel = new MaterialSkin.Controls.MaterialButton();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_Days)).BeginInit();
            this.SuspendLayout();
            // 
            // materialLabel1
            // 
            this.materialLabel1.AutoSize = true;
            this.materialLabel1.Depth = 0;
            this.materialLabel1.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.materialLabel1.Location = new System.Drawing.Point(27, 6);
            this.materialLabel1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialLabel1.Name = "materialLabel1";
            this.materialLabel1.Size = new System.Drawing.Size(175, 19);
            this.materialLabel1.TabIndex = 0;
            this.materialLabel1.Text = "Add or Update a Favorite";
            // 
            // textBox_Description
            // 
            this.textBox_Description.AnimateReadOnly = false;
            this.textBox_Description.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox_Description.Depth = 0;
            this.textBox_Description.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.textBox_Description.Hint = "Description";
            this.textBox_Description.LeadingIcon = null;
            this.textBox_Description.Location = new System.Drawing.Point(31, 58);
            this.textBox_Description.MaxLength = 500;
            this.textBox_Description.MouseState = MaterialSkin.MouseState.OUT;
            this.textBox_Description.Multiline = false;
            this.textBox_Description.Name = "textBox_Description";
            this.textBox_Description.Size = new System.Drawing.Size(460, 50);
            this.textBox_Description.TabIndex = 1;
            this.textBox_Description.Text = "";
            this.textBox_Description.TrailingIcon = null;
            // 
            // textBox_Files
            // 
            this.textBox_Files.AnimateReadOnly = false;
            this.textBox_Files.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox_Files.Depth = 0;
            this.textBox_Files.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.textBox_Files.Hint = "Files";
            this.textBox_Files.LeadingIcon = null;
            this.textBox_Files.Location = new System.Drawing.Point(31, 170);
            this.textBox_Files.MaxLength = 500;
            this.textBox_Files.MouseState = MaterialSkin.MouseState.OUT;
            this.textBox_Files.Multiline = false;
            this.textBox_Files.Name = "textBox_Files";
            this.textBox_Files.Size = new System.Drawing.Size(460, 50);
            this.textBox_Files.TabIndex = 4;
            this.textBox_Files.Text = "";
            this.textBox_Files.TrailingIcon = null;
            // 
            // textBox_FindText
            // 
            this.textBox_FindText.AnimateReadOnly = false;
            this.textBox_FindText.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox_FindText.Depth = 0;
            this.textBox_FindText.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.textBox_FindText.Hint = "Find Text";
            this.textBox_FindText.LeadingIcon = null;
            this.textBox_FindText.Location = new System.Drawing.Point(31, 226);
            this.textBox_FindText.MaxLength = 2000;
            this.textBox_FindText.MouseState = MaterialSkin.MouseState.OUT;
            this.textBox_FindText.Multiline = false;
            this.textBox_FindText.Name = "textBox_FindText";
            this.textBox_FindText.Size = new System.Drawing.Size(460, 50);
            this.textBox_FindText.TabIndex = 7;
            this.textBox_FindText.Text = "";
            this.textBox_FindText.TrailingIcon = null;
            // 
            // textBox_ReplaceText
            // 
            this.textBox_ReplaceText.AnimateReadOnly = false;
            this.textBox_ReplaceText.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox_ReplaceText.Depth = 0;
            this.textBox_ReplaceText.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.textBox_ReplaceText.Hint = "Replace Text";
            this.textBox_ReplaceText.LeadingIcon = null;
            this.textBox_ReplaceText.Location = new System.Drawing.Point(31, 335);
            this.textBox_ReplaceText.MaxLength = 2000;
            this.textBox_ReplaceText.MouseState = MaterialSkin.MouseState.OUT;
            this.textBox_ReplaceText.Multiline = false;
            this.textBox_ReplaceText.Name = "textBox_ReplaceText";
            this.textBox_ReplaceText.Size = new System.Drawing.Size(460, 50);
            this.textBox_ReplaceText.TabIndex = 11;
            this.textBox_ReplaceText.Text = "";
            this.textBox_ReplaceText.TrailingIcon = null;
            // 
            // textBox_Directory
            // 
            this.textBox_Directory.AnimateReadOnly = false;
            this.textBox_Directory.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox_Directory.Depth = 0;
            this.textBox_Directory.Font = new System.Drawing.Font("Roboto", 16F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.textBox_Directory.Hint = "Directory";
            this.textBox_Directory.LeadingIcon = null;
            this.textBox_Directory.Location = new System.Drawing.Point(31, 114);
            this.textBox_Directory.MaxLength = 2000;
            this.textBox_Directory.MouseState = MaterialSkin.MouseState.OUT;
            this.textBox_Directory.Multiline = false;
            this.textBox_Directory.Name = "textBox_Directory";
            this.textBox_Directory.Size = new System.Drawing.Size(460, 50);
            this.textBox_Directory.TabIndex = 2;
            this.textBox_Directory.Text = "";
            this.textBox_Directory.TrailingIcon = null;
            // 
            // checkbox_SubDirectory
            // 
            this.checkbox_SubDirectory.AutoSize = true;
            this.checkbox_SubDirectory.Depth = 0;
            this.checkbox_SubDirectory.Location = new System.Drawing.Point(494, 127);
            this.checkbox_SubDirectory.Margin = new System.Windows.Forms.Padding(0);
            this.checkbox_SubDirectory.MouseLocation = new System.Drawing.Point(-1, -1);
            this.checkbox_SubDirectory.MouseState = MaterialSkin.MouseState.HOVER;
            this.checkbox_SubDirectory.Name = "checkbox_SubDirectory";
            this.checkbox_SubDirectory.ReadOnly = false;
            this.checkbox_SubDirectory.Ripple = true;
            this.checkbox_SubDirectory.Size = new System.Drawing.Size(136, 37);
            this.checkbox_SubDirectory.TabIndex = 3;
            this.checkbox_SubDirectory.Text = "Subdirectories";
            this.checkbox_SubDirectory.UseVisualStyleBackColor = true;
            // 
            // checkbox_ByDate
            // 
            this.checkbox_ByDate.AutoSize = true;
            this.checkbox_ByDate.Depth = 0;
            this.checkbox_ByDate.Location = new System.Drawing.Point(494, 177);
            this.checkbox_ByDate.Margin = new System.Windows.Forms.Padding(0);
            this.checkbox_ByDate.MouseLocation = new System.Drawing.Point(-1, -1);
            this.checkbox_ByDate.MouseState = MaterialSkin.MouseState.HOVER;
            this.checkbox_ByDate.Name = "checkbox_ByDate";
            this.checkbox_ByDate.ReadOnly = false;
            this.checkbox_ByDate.Ripple = true;
            this.checkbox_ByDate.Size = new System.Drawing.Size(90, 37);
            this.checkbox_ByDate.TabIndex = 5;
            this.checkbox_ByDate.Text = "By Date";
            this.checkbox_ByDate.UseVisualStyleBackColor = true;
            // 
            // numericUpDown_Days
            // 
            this.numericUpDown_Days.Location = new System.Drawing.Point(692, 186);
            this.numericUpDown_Days.Maximum = new decimal(new int[] {
            999,
            0,
            0,
            0});
            this.numericUpDown_Days.Name = "numericUpDown_Days";
            this.numericUpDown_Days.Size = new System.Drawing.Size(43, 20);
            this.numericUpDown_Days.TabIndex = 6;
            // 
            // materialLabel2
            // 
            this.materialLabel2.AutoSize = true;
            this.materialLabel2.Depth = 0;
            this.materialLabel2.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.materialLabel2.Location = new System.Drawing.Point(596, 186);
            this.materialLabel2.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialLabel2.Name = "materialLabel2";
            this.materialLabel2.Size = new System.Drawing.Size(72, 19);
            this.materialLabel2.TabIndex = 99;
            this.materialLabel2.Text = "Max Days";
            // 
            // checkbox_MatchCase
            // 
            this.checkbox_MatchCase.AutoSize = true;
            this.checkbox_MatchCase.Depth = 0;
            this.checkbox_MatchCase.Location = new System.Drawing.Point(174, 277);
            this.checkbox_MatchCase.Margin = new System.Windows.Forms.Padding(0);
            this.checkbox_MatchCase.MouseLocation = new System.Drawing.Point(-1, -1);
            this.checkbox_MatchCase.MouseState = MaterialSkin.MouseState.HOVER;
            this.checkbox_MatchCase.Name = "checkbox_MatchCase";
            this.checkbox_MatchCase.ReadOnly = false;
            this.checkbox_MatchCase.Ripple = true;
            this.checkbox_MatchCase.Size = new System.Drawing.Size(119, 37);
            this.checkbox_MatchCase.TabIndex = 8;
            this.checkbox_MatchCase.Text = "Match Case";
            this.checkbox_MatchCase.UseVisualStyleBackColor = true;
            // 
            // checkbox_MatchWord
            // 
            this.checkbox_MatchWord.AutoSize = true;
            this.checkbox_MatchWord.Depth = 0;
            this.checkbox_MatchWord.Location = new System.Drawing.Point(333, 277);
            this.checkbox_MatchWord.Margin = new System.Windows.Forms.Padding(0);
            this.checkbox_MatchWord.MouseLocation = new System.Drawing.Point(-1, -1);
            this.checkbox_MatchWord.MouseState = MaterialSkin.MouseState.HOVER;
            this.checkbox_MatchWord.Name = "checkbox_MatchWord";
            this.checkbox_MatchWord.ReadOnly = false;
            this.checkbox_MatchWord.Ripple = true;
            this.checkbox_MatchWord.Size = new System.Drawing.Size(121, 37);
            this.checkbox_MatchWord.TabIndex = 9;
            this.checkbox_MatchWord.Text = "Match Word";
            this.checkbox_MatchWord.UseVisualStyleBackColor = true;
            // 
            // checkbox_RegEx
            // 
            this.checkbox_RegEx.AutoSize = true;
            this.checkbox_RegEx.Depth = 0;
            this.checkbox_RegEx.Location = new System.Drawing.Point(498, 277);
            this.checkbox_RegEx.Margin = new System.Windows.Forms.Padding(0);
            this.checkbox_RegEx.MouseLocation = new System.Drawing.Point(-1, -1);
            this.checkbox_RegEx.MouseState = MaterialSkin.MouseState.HOVER;
            this.checkbox_RegEx.Name = "checkbox_RegEx";
            this.checkbox_RegEx.ReadOnly = false;
            this.checkbox_RegEx.Ripple = true;
            this.checkbox_RegEx.Size = new System.Drawing.Size(170, 37);
            this.checkbox_RegEx.TabIndex = 10;
            this.checkbox_RegEx.Text = "Regular Expression";
            this.checkbox_RegEx.UseVisualStyleBackColor = true;
            // 
            // materialLabel3
            // 
            this.materialLabel3.AutoSize = true;
            this.materialLabel3.Depth = 0;
            this.materialLabel3.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.materialLabel3.Location = new System.Drawing.Point(38, 286);
            this.materialLabel3.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialLabel3.Name = "materialLabel3";
            this.materialLabel3.Size = new System.Drawing.Size(56, 19);
            this.materialLabel3.TabIndex = 14;
            this.materialLabel3.Text = "Options";
            // 
            // materialButton_OK
            // 
            this.materialButton_OK.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.materialButton_OK.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.materialButton_OK.Depth = 0;
            this.materialButton_OK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.materialButton_OK.HighEmphasis = true;
            this.materialButton_OK.Icon = null;
            this.materialButton_OK.Location = new System.Drawing.Point(41, 417);
            this.materialButton_OK.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.materialButton_OK.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialButton_OK.Name = "materialButton_OK";
            this.materialButton_OK.NoAccentTextColor = System.Drawing.Color.Empty;
            this.materialButton_OK.Size = new System.Drawing.Size(64, 36);
            this.materialButton_OK.TabIndex = 12;
            this.materialButton_OK.Text = "OK";
            this.materialButton_OK.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.materialButton_OK.UseAccentColor = false;
            this.materialButton_OK.UseVisualStyleBackColor = true;
            // 
            // materialButton_Cancel
            // 
            this.materialButton_Cancel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.materialButton_Cancel.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.materialButton_Cancel.Depth = 0;
            this.materialButton_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.materialButton_Cancel.HighEmphasis = true;
            this.materialButton_Cancel.Icon = null;
            this.materialButton_Cancel.Location = new System.Drawing.Point(159, 417);
            this.materialButton_Cancel.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.materialButton_Cancel.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialButton_Cancel.Name = "materialButton_Cancel";
            this.materialButton_Cancel.NoAccentTextColor = System.Drawing.Color.Empty;
            this.materialButton_Cancel.Size = new System.Drawing.Size(77, 36);
            this.materialButton_Cancel.TabIndex = 13;
            this.materialButton_Cancel.Text = "Cancel";
            this.materialButton_Cancel.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.materialButton_Cancel.UseAccentColor = false;
            this.materialButton_Cancel.UseVisualStyleBackColor = true;
            // 
            // FormModFavorite
            // 
            this.AcceptButton = this.materialButton_OK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.materialButton_Cancel;
            this.ClientSize = new System.Drawing.Size(776, 496);
            this.ControlBox = false;
            this.Controls.Add(this.materialButton_Cancel);
            this.Controls.Add(this.materialButton_OK);
            this.Controls.Add(this.materialLabel3);
            this.Controls.Add(this.checkbox_RegEx);
            this.Controls.Add(this.checkbox_MatchWord);
            this.Controls.Add(this.checkbox_MatchCase);
            this.Controls.Add(this.materialLabel2);
            this.Controls.Add(this.numericUpDown_Days);
            this.Controls.Add(this.checkbox_ByDate);
            this.Controls.Add(this.checkbox_SubDirectory);
            this.Controls.Add(this.textBox_Directory);
            this.Controls.Add(this.textBox_ReplaceText);
            this.Controls.Add(this.textBox_FindText);
            this.Controls.Add(this.textBox_Files);
            this.Controls.Add(this.textBox_Description);
            this.Controls.Add(this.materialLabel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormModFavorite";
            this.Padding = new System.Windows.Forms.Padding(2, 52, 2, 2);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "FormModFavorite";
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown_Days)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private MaterialSkin.Controls.MaterialLabel materialLabel1;
        private MaterialSkin.Controls.MaterialTextBox textBox_Description;
        private MaterialSkin.Controls.MaterialTextBox textBox_Files;
        private MaterialSkin.Controls.MaterialTextBox textBox_FindText;
        private MaterialSkin.Controls.MaterialTextBox textBox_ReplaceText;
        private MaterialSkin.Controls.MaterialTextBox textBox_Directory;
        private MaterialSkin.Controls.MaterialCheckbox checkbox_SubDirectory;
        private MaterialSkin.Controls.MaterialCheckbox checkbox_ByDate;
        private System.Windows.Forms.NumericUpDown numericUpDown_Days;
        private MaterialSkin.Controls.MaterialLabel materialLabel2;
        private MaterialSkin.Controls.MaterialCheckbox checkbox_MatchCase;
        private MaterialSkin.Controls.MaterialCheckbox checkbox_MatchWord;
        private MaterialSkin.Controls.MaterialCheckbox checkbox_RegEx;
        private MaterialSkin.Controls.MaterialLabel materialLabel3;
        private MaterialSkin.Controls.MaterialButton materialButton_OK;
        private MaterialSkin.Controls.MaterialButton materialButton_Cancel;
    }
}