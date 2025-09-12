using MaterialSkin;
using MaterialSkin.Controls;
using System.Windows.Forms;

namespace FindReplace
{
    // In UpdateForm are 2 Buttons
    // - Update:    materialButton_Update           in materialButton_Update DialogResult= OK
    // - Cancel:    materialButton_Cancel           in materialButton_Cancel DialogResult= Cancel (Default)
    // Forms AcceptButton: materialButton_Update
    //       CancelButton: materialButton_Cancel
    public partial class UpdateForm : MaterialForm
    {
        public UpdateForm()
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
        }

        public string Info { get { return lblInfo.Text; } set { lblInfo.Text = value; } }
        public string MoreInfoLink { get; set; }

        private void LinkInfo_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            linkInfo.LinkVisited = true;
            System.Diagnostics.Process.Start(MoreInfoLink);
        }
    }
}
