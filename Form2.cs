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
using System.Diagnostics;

namespace FindReplace
{
    public partial class Form2 : Form
    {
        public Form2(string backup_dir, string file)
        {
            InitializeComponent();

            //Point point1 = this.Location;
            //MessageBox.Show("Location:" + point1);
            //this.StartPosition = FormStartPosition.CenterParent(+30, 30);

            label_File.Text = file;
            string zipFile = backup_dir + @"\" + FileTools.zipArchiveName;
            label_ZipFile.Text = zipFile;

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
                            entry.LastWriteTime.DateTime.ToString(),
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

        private void Button_Restore_Click(object sender, EventArgs e)
        {
            int ID = 0;           // TAG contains ID of list_files            
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;

            foreach (ListViewItem item in fileselection)
            {
                ID = Convert.ToInt16(item.Tag);
            }
            if (ID == 0)
                MessageBox.Show("Please select a file version to restore.");
            else
            {
                FileTools.RestoreFile(Path.GetDirectoryName(label_ZipFile.Text), label_File.Text, ID);
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void Button_View_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Code View follows ...");
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            // CenterParent plus Offset 150m50
            // ... shift a bit to the right, don't cover parents preview and listivew 
            if (Owner != null)
                Location = new Point(Owner.Location.X + Owner.Width / 2 - Width / 2 + 150,
                    Owner.Location.Y + Owner.Height / 2 - Height / 2 + 50);
        }
    }
}
