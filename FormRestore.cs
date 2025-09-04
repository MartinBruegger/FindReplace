using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;

namespace FindReplace
{
    public partial class FormRestore : MaterialForm
    {
        public FormRestore(string backup_dir, string file)  //public Form2(string backup_dir, string file)
        {
            InitializeComponent();
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);

            //materialLabel_File.Text = file;
            string zipFile = backup_dir + @"\" + FileTools.zipArchiveName;
            materialLabel_ZipFile.Text = zipFile;

            //string[] row1 = { "" };
            int fileCounter = 0;

            using (ZipArchive zipArchive = ZipFile.OpenRead(zipFile))
            {
                foreach (ZipArchiveEntry entry in zipArchive.Entries)
                {
                    if (entry.FullName == file)
                    {
                        fileCounter++;
                        ListViewItem item = new ListViewItem(new string[]
                        {
                            file,
                            entry.LastWriteTime.DateTime.ToString("yyyy/MM/dd HH:mm:ss"),
                            FileTools.GetFileSizeHuman(file, entry.Length)
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
            //int ID = 0;           // TAG contains ID of list_files
            //string file;                      
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;

            foreach (ListViewItem item in fileselection)
            {
                int ID = Convert.ToInt16(item.Tag);
                //file = item.SubItems[0].Text;
                if (ID == 0)
                    MessageBox.Show("Please select a file version to restore.");
                else
                {
                    FileTools.RestoreFile(Path.GetDirectoryName(materialLabel_ZipFile.Text), item.SubItems[0].Text, ID);
                    this.DialogResult = DialogResult.OK;
                    this.Close();
                }
            }            
        }
    }
}
