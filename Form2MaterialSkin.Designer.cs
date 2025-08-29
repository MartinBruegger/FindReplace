namespace FindReplace
{
    partial class Form2MaterialSkin
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2MaterialSkin));
            this.listView1 = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.materialLabel2 = new MaterialSkin.Controls.MaterialLabel();
            this.materialLabel_ZipFile = new MaterialSkin.Controls.MaterialLabel();
            this.materialButton_Restore = new MaterialSkin.Controls.MaterialButton();
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.SuspendLayout();
            // 
            // listView1
            // 
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2,
            this.columnHeader3});
            this.listView1.FullRowSelect = true;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(10, 76);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(660, 151);
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "File";
            this.columnHeader1.Width = 445;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "Date";
            this.columnHeader2.Width = 145;
            // 
            // materialLabel2
            // 
            this.materialLabel2.AutoSize = true;
            this.materialLabel2.Depth = 0;
            this.materialLabel2.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.materialLabel2.FontType = MaterialSkin.MaterialSkinManager.fontType.Caption;
            this.materialLabel2.Location = new System.Drawing.Point(10, 240);
            this.materialLabel2.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialLabel2.Name = "materialLabel2";
            this.materialLabel2.Size = new System.Drawing.Size(138, 14);
            this.materialLabel2.TabIndex = 3;
            this.materialLabel2.Text = "Restore from .zip Archive:";
            // 
            // materialLabel_ZipFile
            // 
            this.materialLabel_ZipFile.AutoSize = true;
            this.materialLabel_ZipFile.Depth = 0;
            this.materialLabel_ZipFile.Font = new System.Drawing.Font("Roboto", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Pixel);
            this.materialLabel_ZipFile.FontType = MaterialSkin.MaterialSkinManager.fontType.Caption;
            this.materialLabel_ZipFile.Location = new System.Drawing.Point(150, 240);
            this.materialLabel_ZipFile.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialLabel_ZipFile.Name = "materialLabel_ZipFile";
            this.materialLabel_ZipFile.Size = new System.Drawing.Size(116, 14);
            this.materialLabel_ZipFile.TabIndex = 4;
            this.materialLabel_ZipFile.Text = "materialLabel_ZipFile";
            // 
            // materialButton_Restore
            // 
            this.materialButton_Restore.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.materialButton_Restore.Density = MaterialSkin.Controls.MaterialButton.MaterialButtonDensity.Default;
            this.materialButton_Restore.Depth = 0;
            this.materialButton_Restore.HighEmphasis = true;
            this.materialButton_Restore.Icon = null;
            this.materialButton_Restore.Location = new System.Drawing.Point(10, 269);
            this.materialButton_Restore.Margin = new System.Windows.Forms.Padding(4, 6, 4, 6);
            this.materialButton_Restore.MouseState = MaterialSkin.MouseState.HOVER;
            this.materialButton_Restore.Name = "materialButton_Restore";
            this.materialButton_Restore.NoAccentTextColor = System.Drawing.Color.Empty;
            this.materialButton_Restore.Size = new System.Drawing.Size(84, 36);
            this.materialButton_Restore.TabIndex = 2;
            this.materialButton_Restore.Text = "Restore";
            this.materialButton_Restore.Type = MaterialSkin.Controls.MaterialButton.MaterialButtonType.Contained;
            this.materialButton_Restore.UseAccentColor = false;
            this.materialButton_Restore.UseVisualStyleBackColor = true;
            this.materialButton_Restore.Click += new System.EventHandler(this.MaterialButton_Restore_Click);
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "Size";
            this.columnHeader3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.columnHeader3.Width = 65;
            // 
            // Form2MaterialSkin
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(677, 320);
            this.Controls.Add(this.materialButton_Restore);
            this.Controls.Add(this.materialLabel_ZipFile);
            this.Controls.Add(this.materialLabel2);
            this.Controls.Add(this.listView1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Form2MaterialSkin";
            this.Text = "Select version to restore";
            this.Load += new System.EventHandler(this.Form2MaterialSkin_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private MaterialSkin.Controls.MaterialLabel materialLabel2;
        private MaterialSkin.Controls.MaterialLabel materialLabel_ZipFile;
        private MaterialSkin.Controls.MaterialButton materialButton_Restore;
        private System.Windows.Forms.ColumnHeader columnHeader3;
    }
}