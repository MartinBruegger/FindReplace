using FindReplace.Properties;
using MaterialSkin;
using MaterialSkin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolBar;

namespace FindReplace
{
    public partial class Form1MaterialSkin : MaterialForm
    {
        private readonly List<string> listFiles = new List<string>();
        private readonly List<string> listDirectories = new List<string>();


        private class MatchInfo
        {
            public int ID { get; set; }
            public int Matches { get; set; }
            public DateTime Date { get; set; }
            public string Size { get; set; }
            public MatchInfo(int _ID, int _Matches, DateTime _Date, string _Size)
            {
                ID = _ID;
                Matches = _Matches;
                Date = _Date;
                Size = _Size;
            }
        }
        private readonly List<MatchInfo> listMatches = new List<MatchInfo>();
        private class PreviewMatchClass
        {
            public int Match { get; set; }
            public string Value { get; set; }
            public int Position { get; set; }
            public PreviewMatchClass(int _Match, string _Value, int _Position)
            {
                Match = _Match;
                Value = _Value;
                Position = _Position;
            }
        }
        private class PreviewFindClass
        {
            public int Match { get; set; }
            public string Value { get; set; }
            public int Position { get; set; }
            public PreviewFindClass(int _Match, string _Value, int _Position)
            {
                Match = _Match;
                Value = _Value;
                Position = _Position;
            }
        }

        private readonly List<PreviewMatchClass> listPreviewMatches = new List<PreviewMatchClass>();
        private readonly MaterialSkinManager TManager = MaterialSkinManager.Instance;

        private string bg_worker1_msg;                          // temporary location for backgroundworker1 messages
        private bool NetworkDriveMapped = false;
        private int preview_match_x = 0;
        private string selectedPath = string.Empty;
        private string selectedFileNames = string.Empty;
        private string selectedTextString = string.Empty;
        private string selectedRegXText = string.Empty;
        private bool selectedSubDirectory = false;
        private bool selectedByDate = false;
        private double selectedFileAge = 0;
        private bool matchCase = false;
        private bool matchWord = false;
        private bool regEx = false;
        private string selectedReplaceString = string.Empty;
        private string selectedBackupDir = string.Empty;
        private readonly string argumentPath = string.Empty;
        private readonly string historyFile = "Favorites.xml";
        private string colorScheme;
        public Form1MaterialSkin(string[] argumentFile)
        {
            if (argumentFile.Length > 0)
            {
                if (Directory.Exists(argumentFile[0]))
                    argumentPath = argumentFile[0];
                else
                    argumentPath = Path.GetDirectoryName(argumentFile[0]);
            }
            InitializeComponent();
            BuildTreeLevel0();
            richTextBoxAbout.LoadFile("FindReplace.rtf");
            materialLabel_AppVersion.Text = Assembly.GetEntryAssembly().GetName().Version.ToString();
            materialLabel_AppDate.Text = Directory.GetLastWriteTime(AppDomain.CurrentDomain.BaseDirectory + "FindReplace.exe").ToString("yyyy'/'MM'/'dd HH:mm");
            var materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);
        }
        private void BuildTreeLevel0()
        {
            treeV_Directories.Nodes.Clear();
            //get a list of the drives
            DriveInfo[] drives = DriveInfo.GetDrives();
            //foreach (string drive in drives)
            foreach (DriveInfo drive in drives)
            {
                //DriveInfo di = new DriveInfo(drive);
                int driveImage;

                switch (drive.DriveType)    //set the drive's icon
                {
                    case DriveType.CDRom:
                        driveImage = 2;
                        break;
                    case DriveType.Network:
                        driveImage = 3;
                        break;
                    case DriveType.NoRootDirectory:
                        driveImage = 1;
                        break;
                    case DriveType.Unknown:
                        driveImage = 1;
                        break;
                    default:
                        driveImage = 0;
                        break;
                }
                TreeNode node = new TreeNode(drive.Name.Substring(0, 2), driveImage, driveImage)
                {
                    Tag = drive.Name
                };

                if (drive.IsReady == true)
                {
                    node.Nodes.Add("...");
                }
                treeV_Directories.Nodes.Add(node);
            }
        }
        private void TreeView1_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            {
                if (e.Node.Nodes.Count > 0)
                {
                    if (e.Node.Nodes[0].Text == "..." && e.Node.Nodes[0].Tag == null)
                    {
                        e.Node.Nodes.Clear();
                        //get the list of sub direcotires
                        string[] dirs = Directory.GetDirectories(e.Node.Tag.ToString());
                        foreach (string dir in dirs)
                        {
                            DirectoryInfo di = new DirectoryInfo(dir);
                            TreeNode node = new TreeNode(di.Name, 0, 1);
                            try
                            {
                                //keep the directory's full path in the tag for use later
                                node.Tag = dir;
                                node.ImageIndex = 4;
                                //if the directory has sub directories add the place holder
                                DirectoryInfo[] diArr = di.GetDirectories();
                                if (diArr.Length > 0)
                                {
                                    node.Nodes.Add(null, "...", 0, 0);
                                }
                            }
                            catch (System.UnauthorizedAccessException)
                            {
                                //display a locked folder icon
                                node.ImageIndex = 5;
                                node.SelectedImageIndex = 5;
                            }
                            catch
                            {
                                MessageBox.Show("DirectoryLister");
                            }
                            finally
                            {
                                e.Node.Nodes.Add(node);
                            }
                        }
                    }
                }
            }
        }
        private void TreeView1_BeforeSelect(object sender, TreeViewCancelEventArgs e)
        {
            TreeNode newSelected = e.Node;
            if (newSelected.ImageIndex == 5)
            {
                MessageBox.Show("No authorization for this directory.");
                treeV_Directories.SelectedNode = null;
                selectedPath = "";
                return;
            }
            selectedPath = newSelected.FullPath;
            newSelected.SelectedImageIndex = 6;
        }
        private void ToolStripMenu_Tree_Explorer_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", treeV_Directories.SelectedNode.FullPath);
        }

        private void MaterialSwitch_Theme_CheckedChanged(object sender, EventArgs e)
        {
            if (materialSwitch_Theme.Checked)
                TManager.Theme = MaterialSkinManager.Themes.LIGHT;
            else
            {
                TManager.Theme = MaterialSkinManager.Themes.DARK;
                SetToolStripMenuBackground();
            }                
        }
        private void SetToolStripMenuBackground()
        {
            ToolStripMenuBackground(contextMenuStrip_Tree);
            ToolStripMenuBackground(contextMenuStrip_Result);
            ToolStripMenuBackground(contextMenuStrip_Favorites);

        }
        private void ToolStripMenuBackground(ContextMenuStrip contextMenuStrip)
        {
            foreach (ToolStripMenuItem tsmi in contextMenuStrip.Items)          // When Theme "dark": set ForeColor SlateGray. Contrast on Black Background/SkyBlue Selection 
            {                                   
                tsmi.ForeColor = Color.SlateGray;
            }
        }
        private void MaterialRadioButton_CSDefault_CheckedChanged(object sender, EventArgs e)
        {
            if (materialRadioButton_CSDefault.Checked)
                SetColorScheme(string.Empty);
        }

        private void MaterialRadioButton_CSOrange_CheckedChanged(object sender, EventArgs e)
        {
            if (materialRadioButton_CSOrange.Checked)
                SetColorScheme("Orange");
        }

        private void MaterialRadioButton_CSGreen_CheckedChanged(object sender, EventArgs e)
        {
            if (materialRadioButton_CSGreen.Checked)
                SetColorScheme("Green");
        }

        private void MaterialRadioButton_CSBlue_CheckedChanged(object sender, EventArgs e)
        {
            if (materialRadioButton_CSBlue.Checked)
                SetColorScheme("Blue");
        }
        private void SetColorScheme(string color)
        {
            switch (color)
            {
                case "Orange":
                    TManager.ColorScheme = new ColorScheme(Primary.Orange800, Primary.Orange900, Primary.Orange500, Accent.Orange200, TextShade.WHITE);
                    break;
                case "Green":
                    TManager.ColorScheme = new ColorScheme(Primary.Green800, Primary.Green900, Primary.Green500, Accent.Green200, TextShade.WHITE);
                    break;
                case "Blue":
                    TManager.ColorScheme = new ColorScheme(Primary.Blue800, Primary.Blue900, Primary.Blue500, Accent.Blue200, TextShade.WHITE);
                    break;
                default:
                    TManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);
                    break;
            }
            colorScheme = color;
            if (!materialSwitch_Theme.Checked) 
                SetToolStripMenuBackground();
        }

        private void ShowUpdateDialog(Version appVersion, Version newVersion, XDocument doc)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<Version, Version, XDocument>(ShowUpdateDialog), appVersion, newVersion, doc);
                return;
            }

            using (UpdateForm f = new UpdateForm())
            {
                f.Text = string.Format(f.Text, Application.ProductName, appVersion);
                f.MoreInfoLink = (string)doc.Root.Element("info");
                f.Info = string.Format(f.Info, newVersion, (DateTime)doc.Root.Element("date"));
                if (f.ShowDialog(this) == DialogResult.OK)
                {
                    Updater.LaunchUpdater(doc);
                    this.Close();
                }
            }
            SetToolStripMenuBackground();
        }

        private void MaterialButton_Update_Click(object sender, EventArgs e)
        {
            UpdateStatus status = Updater.CheckForUpdate(ShowUpdateDialog);
            if (status == UpdateStatus.UpdateFailed)
                MessageBox.Show(this, "Failed to check for update.  Please ty again later.", "Warning");
            else if (status == UpdateStatus.NoUpdate)
                MessageBox.Show(this, "There are no updates available at this time.", "Update Check");
        }

        private void RichTextBoxAbout_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.LinkText);
        }

        private void MaterialButton_ArchiveDirectory_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog
            {
                ShowNewFolderButton = false,
                SelectedPath = materialTextBox_ArchiveDirectory.Text,
                RootFolder = Environment.SpecialFolder.MyComputer
            };
            // Show the FolderBrowserDialog.
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                materialTextBox_ArchiveDirectory.Text = folderDlg.SelectedPath;
            }
        }

        private void Form1MaterialSkin_Load(object sender, EventArgs e)
        {
            // Upgrade?
            string configPath = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath;
            if (!File.Exists(configPath))
            {
                //Existing user config does not exist, so load settings from previous assembly
                Settings.Default.Upgrade();
                Settings.Default.Reload();
                Settings.Default.Save();
            }
            //if (Properties.Settings.Default.F1Size.Width == 0) Properties.Settings.Default.Upgrade();

            if (Properties.Settings.Default.F1Size.Width == 0 || Properties.Settings.Default.F1Size.Height == 0)
            {
                // first start
                // optional: add default values
            }
            else
            {
                this.WindowState = Properties.Settings.Default.F1State;

                // we don't want a minimized window at startup
                if (this.WindowState == FormWindowState.Minimized) this.WindowState = FormWindowState.Normal;

                this.Location = Properties.Settings.Default.F1Location;
                this.Size = Properties.Settings.Default.F1Size;
            }
            materialRadioButtonBRLocal.Checked = Properties.Settings.Default.F1BRLocal;
            materialRadioButtonBRCentral.Checked = Properties.Settings.Default.F1BRCentral;
            materialRadioButtonBRNone.Checked = Properties.Settings.Default.F1BRNone;
            materialTextBox_ArchiveDirectory.Text = Properties.Settings.Default.F1BRDir;
            if (Properties.Settings.Default.F1Theme == "light")
            {
                TManager.Theme = MaterialSkinManager.Themes.LIGHT;
                materialSwitch_Theme.Checked = true;
            }                
            else
            {
                TManager.Theme = MaterialSkinManager.Themes.DARK;
                materialSwitch_Theme.Checked = false;
            }
            string color;
            switch (Properties.Settings.Default.F1ColorScheme)
            {
                case "Orange":
                    color = "Orange";
                    materialRadioButton_CSOrange.Checked = true;
                    break;
                case "Green":
                    color = "Green";
                    materialRadioButton_CSGreen.Checked = true;
                    break;
                case "Blue":
                    color = "Blue";
                    materialRadioButton_CSBlue.Checked = true;
                    break;
                default:
                    color = string.Empty;
                    materialRadioButton_CSDefault.Checked = true;
                    break;
            }
            SetColorScheme(color);

            if (Properties.Settings.Default.F1SubDirectory)
                materialCheckbox_SubDirectory.Checked = true;
            if (!string.IsNullOrEmpty(Properties.Settings.Default.F1Filenames))
                materialTextBox_Filenames.Text = Properties.Settings.Default.F1Filenames;
            if (Properties.Settings.Default.F1SelectByDate)
                numericUpDown_FileAge.Value = Properties.Settings.Default.F1FileAge;
            if (!string.IsNullOrEmpty(Properties.Settings.Default.F1FindString))
                materialTextBox_FindString.Text = Properties.Settings.Default.F1FindString;
            matchCase = Properties.Settings.Default.F1MatchCase;
            PictureBox_Checked(pictureBox_MatchCase, matchCase);
            matchWord = Properties.Settings.Default.F1MatchWord;
            PictureBox_Checked(pictureBox_MatchWord, matchWord);
            regEx = Properties.Settings.Default.F1RegEx;
            PictureBox_Checked(pictureBox_RegEx, regEx);
            if (!string.IsNullOrEmpty(Properties.Settings.Default.F1ReplaceString))
                materialTextBox_ReplaceString.Text = Properties.Settings.Default.F1ReplaceString;
            selectedPath = Properties.Settings.Default.F1SelectedPath;
            if (!string.IsNullOrEmpty(selectedPath))
                if (SelectPathInTreeView(selectedPath) == false)
                {
                    SetMessage("Arguments from last session loaded, but directory \"" + selectedPath + "\" does not exist.", true);
                    selectedPath = string.Empty;
                }

            if (File.Exists(historyFile))
            {
                try
                {
                    XDocument doc = XDocument.Load(historyFile);
                    foreach (var dm in doc.Descendants("Favorite"))
                    {
                        ListViewItem item = new ListViewItem(new string[]
                        {
                    dm.Element("Title")?.Value,
                    dm.Element("Path")?.Value,
                    dm.Element("SubDirectories")?.Value,
                    dm.Element("FileNames")?.Value,
                    dm.Element("ByDate")?.Value,
                    dm.Element("Days")?.Value,
                    dm.Element("FindText")?.Value,
                    dm.Element("MatchCase")?.Value,
                    dm.Element("MatchWord")?.Value,
                    dm.Element("RegEx")?.Value,
                    dm.Element("ReplaceText")?.Value
                        });
                        listV_Favorites.Items.Add(item);
                    }
                    materialTabControl1.SelectedIndex = 2;
                }
                
                catch (Exception ex)
                {
                    MessageBox.Show("Error: "+ ex.Message);
                }
            }
            ListV_Favorites_SetColumnSizes();
        }
        private void ListV_Favorites_SetColumnSizes()
        {
            listV_Favorites.BeginUpdate();

            //Auto size using header
            listV_Favorites.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);

            //Auto size using data
            listV_Favorites.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listV_Favorites.AutoResizeColumn(2, ColumnHeaderAutoResizeStyle.HeaderSize);        // Column "Subdirectories"
            listV_Favorites.AutoResizeColumn(4, ColumnHeaderAutoResizeStyle.HeaderSize);        // Column "By Date"
            listV_Favorites.AutoResizeColumn(5, ColumnHeaderAutoResizeStyle.HeaderSize);        // Column "Days"
            listV_Favorites.AutoResizeColumn(7, ColumnHeaderAutoResizeStyle.HeaderSize);        // Column "Match Case"
            listV_Favorites.AutoResizeColumn(8, ColumnHeaderAutoResizeStyle.HeaderSize);        // Column "Match Word"
            listV_Favorites.AutoResizeColumn(9, ColumnHeaderAutoResizeStyle.HeaderSize);        // Column "RegExp":
            listV_Favorites.AutoResizeColumn(10, ColumnHeaderAutoResizeStyle.HeaderSize);       // Column "Replace Text"
            
            listV_Favorites.EndUpdate();
        }

        private void Form1MaterialSkin_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (listV_Favorites.Items.Count > 0)
                using (XmlWriter writer = XmlWriter.Create(historyFile))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("Favorites");
                    foreach (ListViewItem item in listV_Favorites.Items)
                    {
                        writer.WriteStartElement("Favorite");

                        writer.WriteElementString("Title", item.SubItems[0].Text);
                        writer.WriteElementString("Path", item.SubItems[1].Text);
                        writer.WriteElementString("SubDirectories", item.SubItems[2].Text);
                        writer.WriteElementString("FileNames", item.SubItems[3].Text);
                        writer.WriteElementString("ByDate", item.SubItems[4].Text);
                        writer.WriteElementString("Days", item.SubItems[5].Text);
                        writer.WriteElementString("FindText", item.SubItems[6].Text);
                        writer.WriteElementString("MatchCase", item.SubItems[7].Text);
                        writer.WriteElementString("MatchWord", item.SubItems[8].Text);
                        writer.WriteElementString("RegEx", item.SubItems[9].Text);
                        writer.WriteElementString("ReplaceText", item.SubItems[10].Text);

                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }
            if (NetworkDriveMapped)
            {
                Utility.NetworkDrive.DisconnectNetworkDrive(materialComboBox_NetworkDrive.Text, true);
            }
        }

        private void Form1MaterialSkin_FormClosing(object sender, FormClosingEventArgs e)
        {
            Properties.Settings.Default.F1State = this.WindowState;
            if (this.WindowState == FormWindowState.Normal)
            {
                // save location and size if the state is normal
                Properties.Settings.Default.F1Location = this.Location;
                Properties.Settings.Default.F1Size = this.Size;
            }
            else
            {
                // save the RestoreBounds if the form is minimized or maximized!
                Properties.Settings.Default.F1Location = this.RestoreBounds.Location;
                Properties.Settings.Default.F1Size = this.RestoreBounds.Size;
            }

            Properties.Settings.Default.F1BRLocal = materialRadioButtonBRLocal.Checked;
            Properties.Settings.Default.F1BRCentral = materialRadioButtonBRCentral.Checked;
            Properties.Settings.Default.F1BRNone = materialRadioButtonBRNone.Checked;
            Properties.Settings.Default.F1BRDir = materialTextBox_ArchiveDirectory.Text;

            if (materialSwitch_Theme.Checked)
                Properties.Settings.Default.F1Theme = "light";
            else
                Properties.Settings.Default.F1Theme = "dark";
            if (materialRadioButton_CSDefault.Checked)
                Properties.Settings.Default.F1ColorScheme = string.Empty;
            if (materialRadioButton_CSOrange.Checked)
                Properties.Settings.Default.F1ColorScheme = "Orange";
            if (materialRadioButton_CSGreen.Checked)
                Properties.Settings.Default.F1ColorScheme = "Green";
            if (materialRadioButton_CSBlue.Checked)
                Properties.Settings.Default.F1ColorScheme = "Blue";
            
            Properties.Settings.Default.F1SelectedPath = selectedPath;
            Properties.Settings.Default.F1FindString = materialTextBox_FindString.Text;
            Properties.Settings.Default.F1ReplaceString = materialTextBox_ReplaceString.Text;
            Properties.Settings.Default.F1Filenames = materialTextBox_Filenames.Text;
            Properties.Settings.Default.F1FileAge = (int)numericUpDown_FileAge.Value;
            Properties.Settings.Default.F1SubDirectory = materialCheckbox_SubDirectory.Checked;
            Properties.Settings.Default.F1SelectByDate = materialCheckbox_SelectByDate.Checked;
            Properties.Settings.Default.F1MatchCase   = matchCase;
            Properties.Settings.Default.F1MatchWord = matchWord;
            Properties.Settings.Default.F1RegEx = regEx;

            // don't forget to save the settings
            Properties.Settings.Default.Save();
        }

        private bool SelectPathInTreeView(string selectedPath)
        {
            {
                bool root_node = true;
                bool expand_node = true;

                string[] words = selectedPath.Split('\\');
                int word_count = words.Count();
                int words_processed = 0;
                if (Directory.Exists(selectedPath))
                {
                    foreach (string word in words)
                    {
                        words_processed++;
                        if (words_processed == word_count)
                        {
                            expand_node = false;
                        }

                        if (words_processed > 1)
                        {
                            root_node = false;
                        }
                        if (!PreExpand(word, root_node, expand_node))
                        {
                            return true;
                        }
                    }
                    return true;
                }
                else
                {
                    SetMessage("Directory: \"" + selectedPath + "\" does not exist.", true);
                    return false;
                }
            }
        }
        private bool PreExpand(string text, bool root_node, bool expand_node)
        {
            if (root_node)
            {
                foreach (TreeNode t in treeV_Directories.Nodes)
                {
                    if (String.Equals(t.Text, text, StringComparison.OrdinalIgnoreCase))
                    {
                        treeV_Directories.SelectedNode = t;
                        if (expand_node) { treeV_Directories.SelectedNode.Expand(); }
                        else { treeV_Directories.Select(); }
                        return true;
                    }
                }
                SetMessage("Directory: \"" + selectedPath + "\" does not exist in current Tree View.", true);
                return false;
            }
            else
            {
                foreach (TreeNode t in treeV_Directories.SelectedNode.Nodes)
                {
                    if (String.Equals(t.Text, text, StringComparison.OrdinalIgnoreCase))
                    {
                        treeV_Directories.SelectedNode = t;
                        if (expand_node) { treeV_Directories.SelectedNode.Expand(); }
                        else { treeV_Directories.Select(); }
                        return true;
                    }
                }
            }
            return false;
        }

        private void ListV_Favorites_KeyDown(object sender, KeyEventArgs e)
        {
            if (Keys.Delete == e.KeyCode)
            {
                ToolStripMenu_Favorites_Delete_Click(sender, e);
                ListV_Favorites_SetColumnSizes();
            }
        }

        private void ListV_Favorites_SelectedIndexChanged(object sender, EventArgs e)
        {
            System.Windows.Forms.ListView.SelectedListViewItemCollection FavoritesSelection = this.listV_Favorites.SelectedItems;
            foreach (ListViewItem item in FavoritesSelection)
            {
                item.Selected = true;
            }
        }
        private void ToolStripMenu_Favorites_Load_Click(object sender, EventArgs e)
        {
            listV_Result.Items.Clear();
            treeV_Directories.CollapseAll();
            foreach (ListViewItem item in listV_Favorites.SelectedItems)
            {
                selectedPath = item.SubItems[1].Text;
                if (item.SubItems[2].Text == "True")
                    materialCheckbox_SubDirectory.Checked = true;
                else
                    materialCheckbox_SubDirectory.Checked = false;
                materialTextBox_Filenames.Text = item.SubItems[3].Text;
                if (item.SubItems[4].Text == "True")
                    materialCheckbox_SelectByDate.Checked = true;
                else
                    materialCheckbox_SelectByDate.Checked = false;
                numericUpDown_FileAge.Value = Convert.ToDecimal(item.SubItems[5].Text);
                materialTextBox_FindString.Text = item.SubItems[6].Text;
                if (item.SubItems[7].Text == "True")
                    matchCase = true;
                else
                    matchCase = false;
                PictureBox_Checked(pictureBox_MatchCase, matchCase);
                if (item.SubItems[8].Text == "True")
                    matchWord = true;
                else
                    matchWord = false;
                PictureBox_Checked(pictureBox_MatchWord, matchWord);
                if (item.SubItems[9].Text == "True")
                    regEx = true;
                else
                    regEx = false;
                PictureBox_Checked(pictureBox_RegEx, regEx);
                materialTextBox_ReplaceString.Text = item.SubItems[10].Text;
                SetMessage("Favorites: \"" + item.SubItems[0].Text + "\" loaded.");
                if (!string.IsNullOrEmpty(selectedPath))
                    if (SelectPathInTreeView(selectedPath) == false)
                    {
                        SetMessage(string.Format("Favorites: \"{0}\" loaded, but directory \"{1}\" does not exist.", item.SubItems[0].Text, selectedPath), true);
                        selectedPath = string.Empty;
                    }
            }
        }

        private void ListV_Favorites_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            ToolStripMenu_Favorites_LoadAndFind_Click(sender, e);
        }
        private void ToolStripMenu_Favorites_LoadAndFind_Click(object sender, EventArgs e)
        {
            ToolStripMenu_Favorites_Load_Click(sender, e);
            if (!string.IsNullOrEmpty(selectedPath))
            {
                MaterialButton_Find_Click(sender, e);
            }
        }
        private void ToolStripMenu_Favorites_Delete_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listV_Favorites.SelectedItems)
            {
                SetMessage(string.Format("Favorites: \"{0}\" deleted.", item.SubItems[0].Text));
                listV_Favorites.Items.Remove(item);
            }
        }

        private void MaterialButton_Find_Click(object sender, EventArgs e)
        {
            FindOrReplace("Find");
        }

        private void MaterialButton_Replace_Click(object sender, EventArgs e)
        {
            FindOrReplace("Replace");
        }
        private void FindOrReplace(string action)
        {
            listV_Result.Items.Clear();
            SetMessage("Processing - generating a list of files ...");
            if (string.IsNullOrEmpty(materialTextBox_FindString.Text))
            {
                MessageBox.Show("Please enter a \"Find Text\".");
                return;
            }
            if (action == "Replace")
                if (string.IsNullOrEmpty(materialTextBox_FindString.Text))
                {
                    MessageBox.Show("Please enter a \"Replace Text\".");
                    return;
                }
            if (string.IsNullOrEmpty(selectedPath))
            {
                MessageBox.Show("Please select a starting directory.");
                return;
            }
            materialButton_Find.Visible = false;
            materialButton_Replace.Visible = false;
            materialButton_Cancel.Visible = true;
            if (string.IsNullOrEmpty(materialTextBox_Filenames.Text)) materialTextBox_Filenames.Text = "*";
            selectedFileNames = materialTextBox_Filenames.Text;
            selectedTextString = materialTextBox_FindString.Text;
            selectedSubDirectory = materialCheckbox_SubDirectory.Checked;
            selectedByDate = materialCheckbox_SelectByDate.Checked;
            selectedFileAge = Convert.ToDouble(numericUpDown_FileAge.Value);
            selectedReplaceString = materialTextBox_ReplaceString.Text;
            selectedBackupDir = materialTextBox_ArchiveDirectory.Text;
            materialProgressBar1.Visible = true;
            materialProgressBar1.Maximum = 100;
            materialProgressBar1.Step = 1;
            materialProgressBar1.Value = 40;
            ResetPreview();
            //listView_Result.Items.Clear();  
            backgroundWorker1.RunWorkerAsync(action);
            //materialTabControl1.SelectedIndex = 0;
        }

        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
           
            string action = (string)e.Argument;
            int total_files_skipped = GetFiles(e);
            int total_directories = listDirectories.Count();
            int total_files = listFiles.Count();
            int total_matches = 0;
            int total_file_matches = 0;
            int total_file_replaced = 0;
            string backupDir = string.Empty;
            if (materialRadioButtonBRCentral.Checked)
                backupDir = selectedBackupDir;

            bg_worker1_msg = "0 Files processed.";
            for (int i = 0; i < total_files; i++)
            {
                if (backgroundWorker1.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                backgroundWorker1.ReportProgress((int)((double)i / total_files * 60)+40, listFiles[i]);
                int matches = FindStringInFile(listFiles[i], i, listMatches);
                if (matches > 0)
                {
                    total_matches += matches;
                    total_file_matches++;
                    if (action == "Replace")
                    {
                        if (materialRadioButtonBRLocal.Checked)
                            backupDir = Path.GetDirectoryName(listFiles[i]);
                        if (ReplaceString(backupDir, materialRadioButtonBRLocal.Checked, listFiles[i]))
                            total_file_replaced++;
                    }
                }
            }
            if (action == "Replace")                
                        bg_worker1_msg = string.Format("{0} directories, {1} files skipped, {2} files processed: total {3} matches in {4} files, {5} files changed",
                             total_directories.ToString("N0", CultureInfo.InvariantCulture),
                             total_files_skipped.ToString("N0", CultureInfo.InvariantCulture),
                             total_files.ToString("N0", CultureInfo.InvariantCulture),
                             total_matches.ToString("N0", CultureInfo.InvariantCulture),
                             total_file_matches.ToString("N0", CultureInfo.InvariantCulture),
                             total_file_replaced.ToString("N0", CultureInfo.InvariantCulture));
            else
                bg_worker1_msg = string.Format("{0} directories, {1} files skipped, {2} files processed: total {3} matches in {4} files.",
                             total_directories.ToString("N0", CultureInfo.InvariantCulture),
                             total_files_skipped.ToString("N0", CultureInfo.InvariantCulture),
                             total_files.ToString("N0", CultureInfo.InvariantCulture),
                             total_matches.ToString("N0", CultureInfo.InvariantCulture),
                             total_file_matches.ToString("N0", CultureInfo.InvariantCulture));
        }
        private int GetFiles(DoWorkEventArgs e)
        {
            List<string> blockFileExtensions = new List<string>()
        {
            ".7z",
            ".avi",
            ".bin",
            ".bmp",
            ".cab",
            ".cache",
            ".catalogs",
            ".chm",
            ".com",
            ".cpl",
            ".cur",
            ".dat",
            ".DAT",
            ".db",
            ".dbf",
            ".dll",
            ".dmp",
            ".doc",
            ".docx",
            ".edb",
            ".exe",
            ".gif",
            ".hdmp",
            ".ico",
            ".ide-shm",
            ".ide-wal",
            ".ide",
            ".ifs",
            ".iso",
            ".jar",
            ".jpeg",
            ".jpg",
            ".ldb",
            ".lnk",
            ".lock",
            ".LOG1",
            ".LOG2",
            ".mdf",
            ".mkv",
            ".mov",
            ".mov",
            ".mp3",
            ".mp4",
            ".mpeg",
            ".mpg",
            ".msi",
            ".nupkg",
            ".ods",
            ".ova",
            ".pdb",
            ".pdf",
            ".pgn",
            ".pkg",
            ".png",
            ".pps",
            ".ppt",
            ".pptx",
            ".rar",
            ".rmskin",
            ".rtf",
            ".sqlite",
            ".svn",
            ".svn_base",
            ".sys",
            ".tar",
            ".tar.gz",
            ".tif",
            ".tiff",
            ".ttf",
            ".wav",
            ".wks",
            ".wmv",
            ".wps",
            ".xlr",
            ".xls",
            ".xlsx",
            ".zip"
        };
            string path = selectedPath;
            int total_files_skipped = 0;
            SearchOption searchOption = SearchOption.TopDirectoryOnly;
            if (selectedSubDirectory)
                searchOption = SearchOption.AllDirectories;
            if (path.Length < 3)
                path += "\\";                     // C: returns 0 files - use C:\ instead ...
            listFiles.Clear();                    // multiple Find/Replace pressed; reset listFiles
            listDirectories.Clear();
            listMatches.Clear();
            listPreviewMatches.Clear();

            DateTime fileCompareDate = DateTime.Now.AddDays(-Convert.ToDouble(selectedFileAge));

            char[] delimiters = new char[] { ',', ';', ' ' };
            int x_searchPatterns = 0;
            string[] searchPatterns = selectedFileNames.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            string fileExtension;
            
            foreach (var searchPattern in searchPatterns)
            {
                x_searchPatterns++;
                var files = GetDirectoryFiles(path, searchPattern, searchOption, listDirectories, x_searchPatterns);
                foreach (string file in files)
                {
                    fileExtension = Path.GetExtension(file);
                    if (!string.IsNullOrEmpty(fileExtension)) // skip files without Extensions
                    {
                        if (blockFileExtensions.Contains(fileExtension, StringComparer.OrdinalIgnoreCase)) // skip files with Extension in above List  
                            total_files_skipped++;
                        else
                        {
                            if (selectedByDate)
                            {
                                if (File.GetLastWriteTime(file).Date >= fileCompareDate.Date)
                                    listFiles.Add(file);
                            }
                            else
                                listFiles.Add(file);
                        }
                    }
                    if (backgroundWorker1.CancellationPending)
                    {
                        e.Cancel = true;
                        return 0;
                    }
                }
            }
            listFiles.Sort();
            return total_files_skipped;
        }
        /// <summary>
        /// A safe way to get all the files in a directory and sub directory without crashing on UnauthorizedException or PathTooLongException
        /// </summary>
        /// <param name="rootPath">Starting directory</param>
        /// <param name="patternMatch">Filename pattern match</param>
        /// <param name="searchOption">Search subdirectories or only top level directory for files</param>
        /// <param name="listDirectories">List containing directory names</param>
        /// <param name="x_loop">Flag to control Add to list; 1=Add</param>
        /// <returns>List of files</returns>
        public static IEnumerable<string> GetDirectoryFiles(string rootPath, string patternMatch,
            SearchOption searchOption, List<string> listDirectories, int x_loop)
        {
            var foundFiles = Enumerable.Empty<string>();
            if (x_loop == 1)
            {
                listDirectories.Add(rootPath);
            }

            if (searchOption == SearchOption.AllDirectories)
            {
                try
                {
                    IEnumerable<string> subDirs = Directory.EnumerateDirectories(rootPath);
                    foreach (string dir in subDirs)
                    {
                        foundFiles = foundFiles.Concat(GetDirectoryFiles(dir, patternMatch, searchOption, listDirectories, x_loop)); // Add files in subdirectories recursively to the list
                    }
                }
                catch (UnauthorizedAccessException) { }
                catch (PathTooLongException) { }
            }
            try
            {
                foundFiles = foundFiles.Concat(Directory.EnumerateFiles(rootPath, patternMatch)); // Add files from the current directory               
            }
            catch (UnauthorizedAccessException) { }
            return foundFiles;
        }
        private int FindStringInFile(string file, int id = -1, List<MatchInfo> listMatches = null)
        {
            string text = FileTools.ReadFileString(file);
            int matches = 0;
            selectedRegXText = selectedTextString;
            if (!regEx)                                                                         // When RegEx is not selected
                selectedRegXText = ProtectRegExpCharacterEscape(selectedRegXText);              // go protect all Escape Characters with a leading '\'
            if (matchWord)                                                                      // When "Match Word" is selected
                selectedRegXText = "\\b" + selectedRegXText + "\\b";                            // enclose findText with \b

            RegexOptions options = RegexOptions.Multiline;
            if (matchCase == false) 
                options = RegexOptions.Multiline | RegexOptions.IgnoreCase;
            try
            {
                foreach (Match m in Regex.Matches(text, selectedRegXText, options))  
                { matches++; }
                if ((matches > 0) && (id > -1))
                    listMatches.Add(new MatchInfo(id, matches,
                        File.GetLastWriteTime(file), FileTools.GetFileSizeHuman(file)));
            }
            catch (Exception e)
            {
                backgroundWorker1.CancelAsync();
                MessageBox.Show("Error: "
                    + e.Message
                    + Environment.NewLine
                    + Environment.NewLine
                    + "Possible cause: Input Error in FIND TEXT, Single Backslash not escaped with Backslash ?");
            }
            return matches;
        }
        private bool ReplaceString(string backupDir, bool backupLocal, string file)
        {
            string fileOld = FileTools.ReadFileString(file);
            RegexOptions options = RegexOptions.Multiline;
            if (matchCase == false)
                options = RegexOptions.Multiline | RegexOptions.IgnoreCase;
            Regex rgx = new Regex(selectedRegXText, options);
            string fileNew = rgx.Replace(fileOld, selectedReplaceString);
            if (!fileOld.Equals(fileNew))
            {
                FileTools.BackupFile(backupDir, backupLocal, file);
                try
                {
                    File.WriteAllText(file, fileNew, Encoding.Default);
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error: " + e.Message);
                    return false;
                }
                return true;
            }
            return false;
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            materialProgressBar1.Value = e.ProgressPercentage;
            SetMessage("Processing: " + e.UserState as string);
        }

        private void BackgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // First, handle the case where an exception was thrown.
            if (e.Error != null)
            {
                _ = MessageBox.Show(e.Error.Message);
            }
            if (e.Cancelled)
                SetMessage("Process aborted (CANCEL pressed).", true);
            else
                SetMessage("Process completed. (" + bg_worker1_msg + ")");
            materialProgressBar1.Value = 0;
            materialProgressBar1.Visible = false;
            materialButton_Find.Visible = true;
            materialButton_Replace.Visible = true;
            materialButton_Cancel.Visible = false;
            //listView_Result.Items.Clear();
            richTextBox_Preview.Text = string.Empty;
            if (listMatches.Count() == 0)
            {
                return;
            }
            string[] directoriesOld;
            string[] directoriesNew;
            string directoryNameOld = selectedPath;
            string directoryNameNew;
            string file;
            int imageIndex;
            bool fileInResult = false;
            int firstFile = 0;

            // ListViewItem (string item text, int image index)
            // 1st Item: Text and Tag = starting directory
            // ListViewItem item0 = new ListViewItem(directoryNameOld, 0)
            imageIndex = 0;                                                                     // imageIndex 0 = Directory
            ListViewItem item0 = new ListViewItem(directoryNameOld, imageIndex)
            {
                IndentCount = 0,
                Tag = directoryNameOld
            };
            listV_Result.Items.Add(item0);                                                      // top directory
            directoriesOld = directoryNameOld.Split('\\');
            int indentOffset = directoriesOld.Count() - 1;
            foreach (MatchInfo foundItem in listMatches)
            {
                file = listFiles[foundItem.ID];
                directoryNameNew = Path.GetDirectoryName(file);
                directoriesNew = directoryNameNew.Split('\\');
                if (string.Compare(directoryNameOld, directoryNameNew) == 0)                
                {
                    // ListViewItem (string array subitems and text, int image index)
                    // new item in existing directory 
                    imageIndex = 1;                                                             // imageIndex 1 = File
                    ListViewItem item1 = new ListViewItem(new string[]
                        {
                            Path.GetFileName(file),
                            foundItem.Matches.ToString(),
                            foundItem.Date.ToString("yyyy/MM/dd HH:mm:ss"),
                            foundItem.Size
                        }, imageIndex)
                    {
                        IndentCount = directoriesNew.Count() - indentOffset,
                        Tag = foundItem.ID
                    };
                    listV_Result.Items.Add(item1);                                              // file in  top directory                       
                    directoriesOld = directoriesNew;
                    directoryNameOld = directoryNameNew;
                    if (!fileInResult)
                    {
                        fileInResult = true;
                        firstFile = listV_Result.Items.Count;
                    }
                }
                else
                { // add directory tree
                    int oldElements = directoriesOld.Count();
                    int startLoop = indentOffset + 1;
                    imageIndex = 0;                                                             // imageIndex 0 = Directory
                    directoryNameOld = selectedPath;
                    for (int i = indentOffset + 1; i < directoriesNew.Count(); i++)
                    {
                        directoryNameOld = directoryNameOld + @"\" + directoriesNew[i];
                        if (i > oldElements - 1)
                        {
                            ListViewItem item1 = new ListViewItem(directoriesNew[i], imageIndex)
                            {
                                IndentCount = i - indentOffset,
                                Tag = directoryNameOld
                            };
                            listV_Result.Items.Add(item1);                                      // sub-directory
                        }
                        else
                        {
                            if (string.Compare(directoriesNew[i], directoriesOld[i]) > 0)
                            {
                                ListViewItem item1 = new ListViewItem(directoriesNew[i], imageIndex)
                                {
                                    IndentCount = i - indentOffset,
                                    Tag = directoryNameOld
                                };
                                listV_Result.Items.Add(item1);                                  // sub-directory
                                oldElements = 0;
                            }
                        }
                    }
                    // ListViewItem (string array subitems and text, int image index)
                    // new file in new directory 
                    imageIndex = 1;                                                             // imageIndex 1 = File
                    ListViewItem item2 = new ListViewItem(new string[]
                        {
                            Path.GetFileName(file),
                            foundItem.Matches.ToString(),
                            //foundItem.Date.ToString(),
                            foundItem.Date.ToString("yyyy/MM/dd HH:mm:ss"),
                            foundItem.Size
                        }, imageIndex)
                    {
                        IndentCount = directoriesNew.Count() - indentOffset,
                        Tag = foundItem.ID
                    };
                    listV_Result.Items.Add(item2);                                              // file in sub-directory
                    directoriesOld = directoriesNew;
                    directoryNameOld = directoryNameNew;
                    if (!fileInResult)
                    {
                        fileInResult = true;
                        firstFile = listV_Result.Items.Count;
                    }
                }
            }
            if (listV_Result.Items.Count > 0)
            {
                listV_Result.BeginUpdate();

                //Auto size using header
                listV_Result.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
                //Auto size using data
                listV_Result.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);      // Resize all Columns with ColumnContent 
                listV_Result.AutoResizeColumn(1, ColumnHeaderAutoResizeStyle.HeaderSize);       // Exception "Matches":    HeaderSize
                int listV_ResultWidth = 0;                                                      // calculate position for setting the SplitterDistance
                foreach (ColumnHeader colHeader in listV_Result.Columns)                        // .. get Width of all Columns, add Locations and Margins
                {
                    listV_ResultWidth += colHeader.Width;
                }
                splitContainer1.SplitterDistance = listV_ResultWidth + 54;                      // .. MaterialCard: Location 3, Margin 14; ListView: Location 10 =27 *2=54

                listV_Result.EndUpdate();
                listV_Result.Items[firstFile - 1].Selected = true;
                listV_Result.Items[firstFile - 1].Focused = true;
                materialTabControl1.SelectedIndex = 0;
            }
            
        }
        private void ResetPreview()
        {
            materialTextBox_PrFindString.Text = string.Empty;
            PictureBox_Checked(pictureBox_PRMatchCase, matchCase);
            PictureBox_Checked(pictureBox_PRMatchWord, matchWord);
            PictureBox_Checked(pictureBox_PRRegEx, regEx);
            materialLabel_PrMatches.Text = string.Empty;
            richTextBox_Preview.Text = string.Empty;
            richTextBox_Preview.Font = new Font("Consolas", 8F);  // Lucida Console
        }

        private void MaterialButton_Cancel_Click(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
        }
        private void MaterialButton_FavoritesAdd_Click(object sender, EventArgs e)
        {
            using (FormModFavorite f = new FormModFavorite("",
                    selectedPath, 
                    (materialCheckbox_SubDirectory.Checked) ? "True" : "False",
                    materialTextBox_Filenames.Text,
                    (materialCheckbox_SelectByDate.Checked) ? "True" : "False",
                    numericUpDown_FileAge.Value.ToString(),
                    materialTextBox_FindString.Text,
                    (matchCase) ? "True" : "False",
                    (matchWord) ? "True" : "False",
                    (regEx) ? "True" : "False",
                    materialTextBox_ReplaceString.Text))
            {
                if (f.ShowDialog(this) == DialogResult.OK)
                {
                    ListViewItem itemNew = new ListViewItem(new string[]
                   {
                            f.TextBox_DescriptionValue,
                            f.TextBox_DirectoryValue,
                            (f.Checkbox_SubDirectoryValue) ? "True" : "False",
                            f.TextBox_FilesValue,
                            (f.Checkbox_ByDateValue) ? "True" : "False",
                            f.NumericUpDown_DaysValue.ToString(),
                            f.TextBox_FindTextValue,
                            (f.Checkbox_MatchCaseValue) ? "True" : "False",
                            (f.Checkbox_MatchWordValue) ? "True" : "False",
                            (f.Checkbox_RegExValue) ? "True" : "False",
                            f.TextBox_ReplaceTextValue,
                    });
                    listV_Favorites.Items.Add(itemNew);
                    SetMessage("Favorites: \"" + f.TextBox_DescriptionValue + "\" added.");
                }
                f.Dispose();
                ListV_Favorites_SetColumnSizes();
            }
            SetToolStripMenuBackground();
        }


        private void ToolStripMenu_Favorites_Update_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listV_Favorites.SelectedItems)
            {
                using (FormModFavorite f = new FormModFavorite(item.SubItems[0].Text, item.SubItems[1].Text, item.SubItems[2].Text, item.SubItems[3].Text,
                    item.SubItems[4].Text, item.SubItems[5].Text, item.SubItems[6].Text, item.SubItems[7].Text, item.SubItems[8].Text, item.SubItems[9].Text,
                    item.SubItems[10].Text))
                {
                    if (f.ShowDialog(this) == DialogResult.OK)
                    {
                        int itemIndex = listV_Favorites.SelectedIndices[0];
                        listV_Favorites.Items.Remove(item);
                        ListViewItem itemNew = new ListViewItem(new string[]
                       {
                            f.TextBox_DescriptionValue,
                            f.TextBox_DirectoryValue,
                            (f.Checkbox_SubDirectoryValue) ? "True" : "False",
                            f.TextBox_FilesValue,
                            (f.Checkbox_ByDateValue) ? "True" : "False",
                            f.NumericUpDown_DaysValue.ToString(),
                            f.TextBox_FindTextValue,
                            (f.Checkbox_MatchCaseValue) ? "True" : "False",
                            (f.Checkbox_MatchWordValue) ? "True" : "False",
                            (f.Checkbox_RegExValue) ? "True" : "False",
                            f.TextBox_ReplaceTextValue,
                        });
                        listV_Favorites.Items.Add(itemNew);
                        SetMessage("Favorites: \"" + f.TextBox_DescriptionValue + "\" updated.");
                    }
                    f.Dispose();
                }
            }
            SetToolStripMenuBackground();
        }
        private string ProtectRegExpCharacterEscape (string text)       // When RegEx is not selected: protect all Escape Characters with a leading '\'
        {
            char[] escapeCharacter = { '\\','.', '$', '^', '{', '}', '[', ']', '(', ')',  '|', '*', '+', '?' };  // important: start with \ 
            string newChar;                                                                                      
            for (int i = 0; i < escapeCharacter.Length; i ++)
            {
                newChar = "\\" + escapeCharacter[i];
                text = text.Replace(escapeCharacter[i].ToString(),newChar);
            }
            return text;                                                    // Example: C:\temp\file.txt ==> C:\\temp\\file\.txt 
        }
        private void FindInPreview()
        {
            int matches = 0;
            string findText = materialTextBox_PrFindString.Text;
            listPreviewMatches.Clear();
            if (pictureBox_PRRegEx.BorderStyle == BorderStyle.None )        // When RegEx is not selected
                findText = ProtectRegExpCharacterEscape( findText );        // go protect all Escape Characters with a leading '\'
            if (pictureBox_PRMatchWord.BorderStyle != BorderStyle.None)     // When "Match Word" is selected
                findText = "\\b" + findText + "\\b";                        // enclose findText with \b
            
            RegexOptions options = RegexOptions.Multiline;
            if (pictureBox_PRMatchCase.BorderStyle == BorderStyle.None)
                options = RegexOptions.Multiline | RegexOptions.IgnoreCase;
            try
            {
                foreach (Match m in Regex.Matches(richTextBox_Preview.Text, findText, options))
                {
                    matches++;
                    ColorKeyword(m.Value, m.Index);
                    listPreviewMatches.Add(new PreviewMatchClass(matches,m.Value,m.Index));
                    if (matches == 1000)                    // limit highlighting of found strings to 1000
                        break;                              // ... otherwise it takes too long and is not responding
                }
                preview_match_x = 0;
                PreviewFindMatchToggler("none");
            }
            catch (Exception e)
            {
                MessageBox.Show("Error: "
                    + e.Message
                    + Environment.NewLine
                    + Environment.NewLine
                    + "Possible cause: Input Error in FIND TEXT, Single Backslash not escaped with Backslash ?");
            }
        }
        private void ColorKeyword(string word, int startIndex)                  // Highlite all found strings
        {
            richTextBox_Preview.Select(startIndex, word.Length);
            richTextBox_Preview.SelectionBackColor = Color.Orchid;
        }
        private void ColorKeywordBackground(string word, int startIndex, bool onOff)    // Selected string LimeGreen, reset old selection to Orchid
        {
            richTextBox_Preview.Select(startIndex, word.Length);
            if (onOff)
            {
                richTextBox_Preview.SelectionBackColor = Color.LimeGreen;           
                richTextBox_Preview.Select(startIndex, word.Length);
                richTextBox_Preview.ScrollToCaret();
            }
            else
                richTextBox_Preview.SelectionBackColor = Color.Orchid;
        }
        private void PreviewFindMatchToggler(string direction)
        {                                                                           // Find DOWN or UP: Select next match and scroll to it
            if (listPreviewMatches.Count > 0)
            {
                if ("up down".Contains(direction))    
                {
                    ColorKeywordBackground(listPreviewMatches[preview_match_x].Value, listPreviewMatches[preview_match_x].Position, false);
                    if (direction == "down")
                    {
                        if (preview_match_x + 1 == listPreviewMatches.Count)
                            preview_match_x = 0;
                        else
                            preview_match_x++;
                    }
                    else
                    {
                        if (preview_match_x == 0)
                            preview_match_x = listPreviewMatches.Count - 1;
                        else
                            preview_match_x--;
                    }
                }
                ColorKeywordBackground(listPreviewMatches[preview_match_x].Value, listPreviewMatches[preview_match_x].Position, true);
            }
            if (listPreviewMatches.Count == 0)
                materialLabel_PrMatches.Text = string.Empty;
            else
                materialLabel_PrMatches.Text = preview_match_x + 1 + "/" + listPreviewMatches.Count;
        }

        private void ListView_Result_SelectedIndexChanged(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of listFiles
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches            
            System.Windows.Forms.ListView.SelectedListViewItemCollection fileselection = this.listV_Result.SelectedItems;
            ResetPreview();
            foreach (ListViewItem item in fileselection)
            {
                if (item.ImageIndex == 1)  // File selected
                {
                    ID = Convert.ToInt16(item.Tag);
                    file = listFiles[ID];
                    richTextBox_Preview.Text = FileTools.ReadFileString(file);
                    materialTextBox_PrFindString.Text = materialTextBox_FindString.Text;
                    FindInPreview();
                }
            }
        }
        
        private void PictureBox_HoverMsg(PictureBox pictureBox, Label label, string message)
        {   
            if (string.IsNullOrEmpty(message))
                label.Visible = false;                      // reset Hover Message
            else
            {
                label.Visible = true;                       // set Hover Message
                Point outPoint = new Point                  // above the Picture Box
                {
                    X = pictureBox.Left,
                    Y = pictureBox.Top - 16
                };
                label.Text = message;
                label.Location = outPoint;
            }
        }
        private void PictureBox_Checked(PictureBox pictureBox, bool pb_checked)  // PictureBox when selected: paint BottmLine in AccentColor
        {
            Bitmap bmp = new Bitmap(pictureBox.Image);
            int x, y, z = bmp.Height - 5;
            for (x = 0; x < bmp.Width; x++)
            {
                for (y = 0; y < bmp.Height; y++)
                {
                    Color oldPixelColor = bmp.GetPixel(x, y);
                    if (y > z)
                    {
                        if (pb_checked)
                            oldPixelColor = TManager.ColorScheme.AccentColor;
                        else
                            oldPixelColor = Color.FromKnownColor(KnownColor.Transparent); // Alpha 0=fully transparent
                        bmp.SetPixel(x, y, oldPixelColor);
                    }
                    else
                        bmp.SetPixel(x, y, oldPixelColor);
                }
            }
            pictureBox.Image = bmp;
            if (pb_checked)
                pictureBox.BorderStyle = BorderStyle.FixedSingle;
            else
                pictureBox.BorderStyle = BorderStyle.None;
        }
        // Picture Boxes on the Left Side 
        private void PictureBox_MatchCase_MouseHover(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_MatchCase, label_Panel2PB, "Match Case");
        }

        private void PictureBox_MatchCase_MouseLeave(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_MatchCase, label_Panel2PB, string.Empty);
        }
        private void PictureBox_MatchCase_Click(object sender, EventArgs e)
        {
            if (matchCase)
                matchCase = false;
            else
                matchCase = true;
            PictureBox_Checked(pictureBox_MatchCase, matchCase);
        }
        private void PictureBox_MatchWord_MouseHover(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_MatchWord, label_Panel2PB, "Match Word");
        }
        private void PictureBox_MatchWord_MouseLeave(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_MatchWord, label_Panel2PB, string.Empty);
        }
        private void PictureBox_MatchWord_Click(object sender, EventArgs e)
        {
            if (matchWord)
                matchWord = false;
            else
                matchWord = true;
            PictureBox_Checked(pictureBox_MatchWord, matchWord);
        }
        private void PictureBox_RegEx_MouseHover(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_RegEx, label_Panel2PB, "Regular Expression");
        }
        private void PictureBox_RegEx_MouseLeave(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_RegEx, label_Panel2PB, string.Empty);
        }
        private void PictureBox_RegEx_Click(object sender, EventArgs e)
        {
            if (regEx)
                regEx = false;
            else
                regEx = true;
            PictureBox_Checked(pictureBox_RegEx, regEx);
        }
        // Picture Boxes on the Right; Preview Section (names start with PictureBox_PRF)
        private void PictureBox_PRFindPage_MouseHover(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_PRFindPage, label_SCPanel2PB, "Find in Page");
        }

        private void PictureBox_PRFindPage_MouseLeave(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_PRFindPage, label_SCPanel2PB, string.Empty);
        }

        private void PictureBox_PRFindPage_Click(object sender, EventArgs e)        // Need to reload the file because of Highlighting strings
        {
            materialLabel_PrMatches.Text = string.Empty;
            int ID;             // TAG contains ID of listFiles
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches            
            System.Windows.Forms.ListView.SelectedListViewItemCollection fileselection = this.listV_Result.SelectedItems;
            foreach (ListViewItem item in fileselection)
            {
                if (item.ImageIndex == 1)  // File selected
                {
                    ID = Convert.ToInt16(item.Tag);
                    file = listFiles[ID];
                    richTextBox_Preview.Text = FileTools.ReadFileString(file);
                    FindInPreview();
                }
            }
        }
        private void PictureBox_PRMatchCase_Click(object sender, EventArgs e)
        {
            if (pictureBox_PRMatchCase.BorderStyle == BorderStyle.FixedSingle)
                PictureBox_Checked(pictureBox_PRMatchCase, false);
            else
                PictureBox_Checked(pictureBox_PRMatchCase, true);
        }

        private void PictureBox_PRMatchCase_MouseHover(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_PRMatchCase, label_SCPanel2PB, "Match Case");
        }

        private void PictureBox_PRMatchCase_MouseLeave(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_PRMatchCase, label_SCPanel2PB, string.Empty);
        }

        private void PictureBox_PRBackward_Click(object sender, EventArgs e)
        {
            PreviewFindMatchToggler("up");
        }
        private void PictureBox_PRForward_Click(object sender, EventArgs e)
        {
            PreviewFindMatchToggler("down");
        }

        private void PictureBox_PRMatchWord_MouseHover(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_PRMatchWord, label_SCPanel2PB, "Match Word");
        }

        private void PictureBox_PRMatchWord_MouseLeave(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_PRMatchWord, label_SCPanel2PB, string.Empty);
        }

        private void PictureBox_PRMatchWord_Click(object sender, EventArgs e)
        {
            if (pictureBox_PRMatchWord.BorderStyle == BorderStyle.FixedSingle)
                PictureBox_Checked(pictureBox_PRMatchWord, false);
            else
                PictureBox_Checked(pictureBox_PRMatchWord, true);
        }

        private void PictureBox_PRRegEx_MouseHover(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_PRRegEx, label_SCPanel2PB, "Regular Expression");
        }

        private void PictureBox_PRRegEx_MouseLeave(object sender, EventArgs e)
        {
            PictureBox_HoverMsg(pictureBox_PRRegEx, label_SCPanel2PB, string.Empty);
        }

        private void PictureBox_PRRegEx_Click(object sender, EventArgs e)
        {
            if (pictureBox_PRRegEx.BorderStyle == BorderStyle.FixedSingle)
                PictureBox_Checked(pictureBox_PRRegEx, false);
            else
                PictureBox_Checked(pictureBox_PRRegEx, true);
        }

        private void ToolStripMenu_Result_Edit_Click(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of listFiles
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches            
            System.Windows.Forms.ListView.SelectedListViewItemCollection fileselection = this.listV_Result.SelectedItems;
            foreach (ListViewItem item in fileselection)
            {
                if (item.ImageIndex == 1)  // File selected
                {
                    ID = Convert.ToInt16(item.Tag);
                    file = listFiles[ID];
                    Process p = new Process();
                    p.StartInfo.UseShellExecute = true;
                    p.StartInfo.FileName = file;
                    try { p.Start(); }
                    catch (Exception EX)
                    {
                        MessageBox.Show(string.Format(EX.Message));
                    }
                }
            }
        }

        private void ToolStripMenu_Result_Explorer_Click(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of listFiles
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches            
            System.Windows.Forms.ListView.SelectedListViewItemCollection fileselection = this.listV_Result.SelectedItems;
            ProcessStartInfo startInfo = new ProcessStartInfo();
            foreach (ListViewItem item in fileselection)
            {
                if (item.ImageIndex == 1)                                   // File selected. Get Filename from listFiles[]
                {
                    ID = Convert.ToInt16(item.Tag);
                    file = listFiles[ID];                    
                    startInfo.Arguments = @"/n, /select, " + file;
                } else                                                      // Directory selected. Directoryname in item.Tag
                    startInfo.Arguments = item.Tag.ToString();
                startInfo.FileName = "explorer.exe";
                Process.Start(startInfo);
            }
        }

        private void ToolStripMenu_Result_Replace_Text_Click(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of listFiles
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches
            string backupDir;   // Backup central, local or Empty
            selectedReplaceString = materialTextBox_ReplaceString.Text;
            System.Windows.Forms.ListView.SelectedListViewItemCollection fileselection = this.listV_Result.SelectedItems;
            foreach (ListViewItem item in fileselection)
            {
                if (item.ImageIndex == 1)                                   // File selected. Get Filename from listFiles[]
                {
                    ID = Convert.ToInt16(item.Tag);
                    file = listFiles[ID];
                    if (materialRadioButtonBRLocal.Checked)
                        backupDir = Path.GetDirectoryName(file);
                    else
                    {
                    if (materialRadioButtonBRCentral.Checked)
                        backupDir = materialTextBox_ArchiveDirectory.Text;
                    else
                        backupDir = string.Empty;
                    }
                    if (ReplaceString(backupDir, materialRadioButtonBRLocal.Checked, file))
                    {
                        item.SubItems[1].Text = FindStringInFile(file).ToString();
                        item.SubItems[2].Text = File.GetLastWriteTime(file).ToString("yyyy/MM/dd HH:mm:ss");
                        item.SubItems[3].Text = FileTools.GetFileSizeHuman(file);
                        ListView_Result_SelectedIndexChanged(listV_Result, EventArgs.Empty);  // Refresh File Preview after restore
                        SetMessage("Replace: file contents changed.");
                    }
                }
            }
        }

        private void ToolStripMenu_Result_Restore_Click(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of listFiles
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches  
            string backupDir;   // Backup local or central
            System.Windows.Forms.ListView.SelectedListViewItemCollection fileselection = this.listV_Result.SelectedItems;
            if (materialRadioButtonBRNone.Checked)
            {
                SetMessage("Nothing to restore - Option \"Do not archive\" is checked.", true);
                return;
            }

            foreach (ListViewItem item in fileselection)
            {
                ID = Convert.ToInt16(item.Tag);
                file = listFiles[ID];
                if (materialRadioButtonBRLocal.Checked)
                    backupDir = Path.GetDirectoryName(file);
                else
                    backupDir = selectedBackupDir;

                switch (FileTools.CountBackupVersions(backupDir, materialRadioButtonBRLocal.Checked, file))
                {
                    case 0:
                        SetMessage("Restore: no archived version of file \"" + file + "\" found", true);
                        break;
                    case 1:
                        if (FileTools.RestoreFile(backupDir, materialRadioButtonBRLocal.Checked, file, 1))
                        {
                            item.SubItems[1].Text = FindStringInFile(file).ToString();
                            item.SubItems[2].Text = File.GetLastWriteTime(file).ToString("yyyy/MM/dd HH:mm:ss");
                            item.SubItems[3].Text = FileTools.GetFileSizeHuman(file);
                            ListView_Result_SelectedIndexChanged(listV_Result, EventArgs.Empty);  // Refresh File Preview after restore
                            SetMessage("Restore: file \"" + file + "\" successful restored.");
                        }
                        else
                        {
                            SetMessage("Restore: restore of file \"" + file + "\" failed.", true);
                        }
                        break;
                    default:
                        FormRestore f = new FormRestore(backupDir, materialRadioButtonBRLocal.Checked, file);
                        if (f.ShowDialog(this) == DialogResult.OK)
                        {
                            item.SubItems[1].Text = FindStringInFile(file).ToString();
                            item.SubItems[2].Text = File.GetLastWriteTime(file).ToString("yyyy/MM/dd HH:mm:ss");
                            item.SubItems[3].Text = FileTools.GetFileSizeHuman(file);
                            ListView_Result_SelectedIndexChanged(listV_Result, EventArgs.Empty);  // Refresh File Preview after restore
                            SetMessage("Restore: file \"" + file + "\" successful restored.");
                        }
                        f.Dispose();
                        SetToolStripMenuBackground();
                        break;
                }
            }
        }

        private void RichTextBox_Preview_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F3)                   // find next in page when F3 is pressed
                PreviewFindMatchToggler("down");
        }
        private void SetMessage(string message, bool warning = false)
        {
            label_Message.Text = message;
            if (warning)
            {
                if (!pictureBox_Warning.Visible)
                    pictureBox_Warning.Visible = true;
            }
            else
            {
                if (pictureBox_Warning.Visible)
                    pictureBox_Warning.Visible = false;
            }
        }
        private enum MoveDirection { Up = -1, Down = 1 };
        private void ToolStripMenu_Favorites_MoveUp_Click(object sender, EventArgs e)
        {
            MoveItems(listV_Favorites, MoveDirection.Up);
        }

        private void ToolStripMenu_Favorites_MoveDown_Click(object sender, EventArgs e)
        {
            MoveItems(listV_Favorites, MoveDirection.Down);
        }
        
        private void MoveItems(System.Windows.Forms.ListView  sender, MoveDirection direction)
        {
            bool valid = sender.SelectedItems.Count > 0 &&
                        ((direction == MoveDirection.Down && (sender.SelectedItems[sender.SelectedItems.Count - 1].Index < sender.Items.Count - 1))
                        || (direction == MoveDirection.Up && (sender.SelectedItems[0].Index > 0)));
            if (valid)
            {
                bool start = true;
                int first_idx = 0;
                List<ListViewItem> items = new List<ListViewItem>();
                foreach (ListViewItem i in sender.SelectedItems)
                {
                    if (start)
                    {
                        first_idx = i.Index;
                        start = false;
                    }
                    items.Add(i);
                }
                sender.BeginUpdate();
                foreach (ListViewItem i in sender.SelectedItems) i.Remove();
                if (direction == MoveDirection.Up)
                {
                    int insert_to = first_idx - 1;
                    foreach (ListViewItem i in items)
                    {
                        sender.Items.Insert(insert_to, i);
                        insert_to++;
                    }
                }
                else
                {
                    int insert_to = first_idx + 1;
                    foreach (ListViewItem i in items)
                    {
                        sender.Items.Insert(insert_to, i);
                        insert_to++;
                    }
                }
                sender.EndUpdate();
            }
        }

        private void treeV_Directories_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void treeV_Directories_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                if (file.StartsWith("\\\\"))
                {
                    string UNCdirectory = file;
                    //string mapDriveMsg;
                    if (!Directory.Exists(file))
                        UNCdirectory = Path.GetDirectoryName(file);
                    string mapDriveMsg = Utility.NetworkDrive.MapNetworkDrive(materialComboBox_NetworkDrive.Text, UNCdirectory);
                    if (string.IsNullOrEmpty(mapDriveMsg))
                    {
                        NetworkDriveMapped = true;
                        BuildTreeLevel0();
                        selectedPath = materialComboBox_NetworkDrive.Text + ":";
                        SetMessage("Drag and Drop - directory: \"" + file + "\" mapped as network drive \"" + materialComboBox_NetworkDrive.Text + "\"");
                    }
                    else
                    {
                        BuildTreeLevel0();
                        selectedPath = null;
                        SetMessage("Drag and Drop: " + mapDriveMsg, true);
                    }
                }
                else
                {
                    if (Directory.Exists(file))
                        selectedPath = file;
                    else
                        selectedPath = Path.GetDirectoryName(file);
                    SetMessage("Drag and Drop - selected directory is \"" + selectedPath + "\"");
                }
            }
            treeV_Directories.CollapseAll();
            if (!string.IsNullOrEmpty(selectedPath))
                _ = SelectPathInTreeView(selectedPath);
        }
    }
}
