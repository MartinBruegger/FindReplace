using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Windows.Forms;

namespace FindReplace
{
    public partial class FormModFavorite : MaterialForm
    {
        public String TextBox_DescriptionValue => this.textBox_Description.Text;
        public String TextBox_DirectoryValue => this.textBox_Directory.Text;
        public bool Checkbox_ByDateValue => this.checkbox_ByDate.Checked;
        public bool Checkbox_SubDirectoryValue => this.checkbox_SubDirectory.Checked;
        public String TextBox_FilesValue => this.textBox_Files.Text;
        public decimal NumericUpDown_DaysValue => this.numericUpDown_Days.Value;
        public String TextBox_FindTextValue => this.textBox_FindText.Text;
        public bool Checkbox_MatchCaseValue => this.checkbox_MatchCase.Checked;
        public bool Checkbox_MatchWordValue => this.checkbox_MatchWord.Checked;
        public bool Checkbox_RegExValue => this.checkbox_RegEx.Checked;
        public String TextBox_ReplaceTextValue => this.textBox_ReplaceText.Text;
        public FormModFavorite(string desc, string dir, string subDirs, string files, string byDate, string fileAge, string findText,
            string matchCase, string matchWord, string regEx, string replaceText )
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            textBox_Description.Text   = desc;
            textBox_Directory.Text     = dir;
            if (subDirs == "True")     checkbox_SubDirectory.Checked = true;
            textBox_Files.Text         = files;
            if (byDate == "True")      checkbox_ByDate.Checked = true;
            numericUpDown_Days.Value   = Convert.ToDecimal(fileAge);
            textBox_FindText.Text      = findText;
            if (matchCase == "True")   checkbox_MatchCase.Checked = true;  
            if (matchWord == "True")   checkbox_MatchWord.Checked = true;
            if (regEx == "True")       checkbox_RegEx.Checked = true;
            textBox_ReplaceText.Text   = replaceText;
        }

        private void MaterialButton_OK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
