using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;

namespace FindReplace
{
    public partial class FormRestore : MaterialForm
    {
        bool argBackupLocal;
        public FormRestore(string backupDir, bool backupLocal, string sourceFileName)  
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            argBackupLocal = backupLocal;
            string entryName = sourceFileName;
            if (backupLocal)
                entryName = Path.GetFileName(sourceFileName);
            string zipFile = backupDir + @"\" + FileTools.zipArchiveName;
            materialLabel_ZipFile.Text = zipFile;
            int fileCounter = 0;
            using (ZipArchive zipArchive = ZipFile.OpenRead(zipFile))
            {
                foreach (ZipArchiveEntry entry in zipArchive.Entries)
                {
                    if (entry.FullName == entryName)
                    {
                        fileCounter++;
                        ListViewItem item = new ListViewItem(new string[]
                        {
                            sourceFileName,
                            entry.LastWriteTime.DateTime.ToString("yyyy/MM/dd HH:mm:ss"),
                            FileTools.GetFileSizeHuman(sourceFileName, entry.Length)
                        })
                        {
                            Tag = fileCounter
                        };
                        listView1.Items.Add(item);
                    }
                }
            }
        }

        private void MaterialButton_Restore_Click(object sender, EventArgs e)
        {
            ListView.SelectedListViewItemCollection fileselection = listView1.SelectedItems;
            foreach (ListViewItem item in fileselection)
            {
                int ID = Convert.ToInt16(item.Tag);
                if (ID == 0)
                    MessageBox.Show("Please select a file version to restore.");
                else
                {
                    FileTools.RestoreFile(Path.GetDirectoryName(materialLabel_ZipFile.Text), argBackupLocal, item.SubItems[0].Text, ID);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }            
        }
    }
}
