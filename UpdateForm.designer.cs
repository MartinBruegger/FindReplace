namespace FindReplace
{
    partial class UpdateForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UpdateForm));
            this.linkInfo = new System.Windows.Forms.LinkLabel();
            this.materialButton_Update = new MaterialSkin.Controls.MaterialButton();
            this.materialButton_Cancel = new MaterialSkin.Controls.MaterialButton();
            this.materialLabel1 = new MaterialSkin.Controls.MaterialLabel();
            this.lblInfo = new MaterialSkin.Controls.MaterialLabel();
            this.SuspendLayout();
            // 
            // linkInfo
            // 
            this.linkInfo.AutoSize = true;
            this.linkInfo.LinkColor = System.Drawing.Color.Teal;
            this.linkInfo.Location = new System.Drawing.Point(29, 89);
            this.linkInfo.Name = "linkInfo";
            this.linkInfo.Size = new System.Drawing.Size(52, 13);
            this.linkInfo.TabIndex = 2;
            this.linkInfo.TabStop = true;
            this.linkInfo.Text = "More Info";
            this.linkInfo.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkInfo_LinkClicked);
            // 
            // materialButton_Update
            // 
            this.materialButton_Update.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.materialButton_Update.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.materialButton_Update.Depth = 0;
            this.materialButton_Update.HighEmphasis = true;
            this.materialButton_Update.Icon = null;
            this.materialButton_Update.Location = new System.Drawing.Point(134, 89);
            this.materialButton_Update.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.materialButton_Update.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialButton_Update.Name = "materialButton_Update";
            this.materialButton_Update.NoAccentTextColor = System.Drawing.Color.Empty;
            this.materialButton_Update.Size = new System.Drawing.Size(77, 36);
            this.materialButton_Update.TabIndex = 5;
            this.materialButton_Update.Text = "Update";
            this.materialButton_Update.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.materialButton_Update.UseAccentColor = false;
            this.materialButton_Update.UseVisualStyleBackColor = true;
            // 
            // materialButton_Cancel
            // 
            this.materialButton_Cancel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.materialButton_Cancel.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.materialButton_Cancel.Depth = 0;
            this.materialButton_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.materialButton_Cancel.HighEmphasis = true;
            this.materialButton_Cancel.Icon = null;
            this.materialButton_Cancel.Location = new System.Drawing.Point(281, 89);
            this.materialButton_Cancel.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.materialButton_Cancel.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialButton_Cancel.Name = "materialButton_Cancel";
            this.materialButton_Cancel.NoAccentTextColor = System.Drawing.Color.Empty;
            this.materialButton_Cancel.Size = new System.Drawing.Size(77, 36);
            this.materialButton_Cancel.TabIndex = 6;
            this.materialButton_Cancel.Text = "Cancel";
            this.materialButton_Cancel.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.materialButton_Cancel.UseAccentColor = false;
            this.materialButton_Cancel.UseVisualStyleBackColor = true;
            // 
            // materialLabel1
            // 
            this.materialLabel1.AutoSize = true;
            this.materialLabel1.Depth = 0;
            this.materialLabel1.Font = new System.Drawing.Font("Roboto", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.materialLabel1.Location = new System.Drawing.Point(28, 18);
            this.materialLabel1.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialLabel1.Name = "materialLabel1";
            this.materialLabel1.Size = new System.Drawing.Size(120, 19);
            this.materialLabel1.TabIndex = 7;
            this.materialLabel1.Text = "Update Available";
            // 
            // lblInfo
            // 
            this.lblInfo.AutoSize = true;
            this.lblInfo.Depth = 0;
            this.lblInfo.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.lblInfo.FontType = MaterialSkin.MaterialSkinManager.fontType.Caption;
            this.lblInfo.Location = new System.Drawing.Point(28, 52);
            this.lblInfo.MouseState = MaterialSkin.MouseState.HOVER;
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(243, 14);
            this.lblInfo.TabIndex = 8;
            this.lblInfo.Text = "A new version {0} was made available on {1}.";
            // 
            // UpdateForm
            // 
            this.AcceptButton = this.materialButton_Update;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightGray;
            this.CancelButton = this.materialButton_Cancel;
            this.ClientSize = new System.Drawing.Size(435, 140);
            this.ControlBox = false;
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.materialLabel1);
            this.Controls.Add(this.materialButton_Cancel);
            this.Controls.Add(this.materialButton_Update);
            this.Controls.Add(this.linkInfo);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UpdateForm";
            this.Padding = new System.Windows.Forms.Padding(2, 52, 2, 2);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "{0} {1}";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.LinkLabel linkInfo;
        private MaterialSkin.Controls.MaterialButton materialButton_Update;
        private MaterialSkin.Controls.MaterialButton materialButton_Cancel;
        private MaterialSkin.Controls.MaterialLabel materialLabel1;
        private MaterialSkin.Controls.MaterialLabel lblInfo;
    }
}