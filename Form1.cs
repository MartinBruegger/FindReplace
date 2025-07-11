// Form1.cs
//
// Copyright 2025 Martin Bruegger
using System;
using System.Globalization;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Reflection;

namespace FindReplace
{
    public partial class Form1 : Form
    {
        List<string> list_files = new List<string>();
        List<string> list_directories = new List<string>();        
        string bg_worker1_msg;
        bool NetworkDriveMapped = false;

        class FoundInfo
        {
            public int ID { get; set; }
            public int Matches { get; set; }
            public DateTime Date { get; set; }
            public string Size { get; set; }
             public FoundInfo(int _ID, int _Matches, DateTime _Date, string _Size)
            {
                ID = _ID;
                Matches = _Matches;
                Date = _Date;
                Size = _Size;
            }
        }

        List<FoundInfo> list_founds = new List<FoundInfo>();

        class PreviewFindClass
        {
            public int Match { get; set; }
            public string Value { get; set; }
            public int Position { get; set; }
            public PreviewFindClass(int _Match, string _Value,  int _Position)
            {
                Match = _Match;
                Value = _Value;
                Position = _Position; 
            }
        }

        private List<PreviewFindClass> list_PreviewFounds = new List<PreviewFindClass>();
        private ListViewColumnSorter lvwColumnSorter;
        
        int preview_match_x = 0;
        string selectedPath = string.Empty;
        string argumentPath = string.Empty;
        string historyFile = "FindReplace_History.xml";


        public Form1(string[] argumentFile)
        {
            if (argumentFile.Length > 0)
            {    
                if (Directory.Exists(argumentFile[0]))
                    argumentPath = argumentFile[0];
                else
                    argumentPath = Path.GetDirectoryName(argumentFile[0]);
            }
            InitializeComponent();
            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);
            lvwColumnSorter = new ListViewColumnSorter();
            this.listView2.ListViewItemSorter = lvwColumnSorter;
            tabControl1.SelectedIndex = 0;            
            richTextBoxAbout.LoadFile("FindReplace.rtf");            
            label_AppVersion.Text = Assembly.GetEntryAssembly().GetName().Version.ToString();         
            label_BuildDate.Text = Directory.GetLastWriteTime(AppDomain.CurrentDomain.BaseDirectory + "FindReplace.exe").ToString("yyyy'/'MM'/'dd HH:mm");
            BuildTreeLevel0();
        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);            
            foreach (string file in files)
            {
                if (file.StartsWith("\\\\"))
                {
                    string UNCdirectory = file;
                    //string mapDriveMsg;
                    if (! Directory.Exists(file))
                         UNCdirectory = Path.GetDirectoryName(file);
                    string mapDriveMsg = Utility.NetworkDrive.MapNetworkDrive(textBox_DriveLetter.Text, UNCdirectory);
                    if ( string.IsNullOrEmpty (mapDriveMsg))
                    {
                        NetworkDriveMapped = true;
                        BuildTreeLevel0();
                        selectedPath = textBox_DriveLetter.Text + ":";
                        textBox_Msg.Text = "Drag and Drop - directory: " + file + " mapped as network drive \"" + textBox_DriveLetter.Text + "\"";
                        textBox_Msg.BackColor = Color.LimeGreen;
                    }
                    else
                    {
                        BuildTreeLevel0();
                        selectedPath = null;
                        textBox_Msg.Text = "Drag and Drop Error - " + mapDriveMsg;
                        textBox_Msg.BackColor = Color.Orange; ;
                    }
                }
                else
                {
                    if (Directory.Exists(file))
                        selectedPath = file;
                    else
                        selectedPath = Path.GetDirectoryName(file);
                    textBox_Msg.Text = "Drag and Drop - selected directory is \"" + selectedPath + "\"";
                    textBox_Msg.BackColor = Color.LimeGreen;
                }
            }
            treeView1.CollapseAll();
            if (! string.IsNullOrEmpty(selectedPath)) 
                SelectPathInTreeView(selectedPath);
        }

        private void BuildTreeLevel0()
        {
            treeView1.Nodes.Clear();
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
                treeView1.Nodes.Add(node);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Upgrade?
            if (Properties.Settings.Default.F1Size.Width == 0) Properties.Settings.Default.Upgrade();

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

            treeView1.HideSelection = false;
            progressBar1.Visible = false;
            button_Abort.Visible = false;
            var app_settings = ConfigurationManager.AppSettings;            
            if (argumentPath.Length > 0)
                selectedPath = argumentPath;
            else
            {
                selectedPath = null;
                if (!string.IsNullOrEmpty(app_settings["PATH"]))
                {
                    if (Directory.Exists(app_settings["PATH"]))
                        selectedPath = app_settings["PATH"];
                }                   
            }                
            textBox_fileNames.Text = app_settings["FILENAMES"];
            textBox_Find.Text = app_settings["FIND"];
            textBox_Replace.Text = app_settings["REPLACE"];            
            if (!string.IsNullOrEmpty(app_settings["IGNORE_CASE"]))            
                checkBox_IgnoreCase.Checked = true;
            if (!string.IsNullOrEmpty(app_settings["INCL_SUBDIRS"]))            
                checkBox_Subdir.Checked = true;            
            if (!string.IsNullOrEmpty(app_settings["LASTMODIFICATIONDATE"]))
            {
                checkBox_LastModificationDate.Checked = true;
                tabControl1.SelectedIndex = 1;
            }                       
            numericUpDown_FileAge.Value = Convert.ToDecimal( app_settings["FILEAGE"]);

            if (!string.IsNullOrEmpty(app_settings["BACKUP_DIR"]))
            {
                if (Directory.Exists(app_settings["BACKUP_DIR"]))
                    textBox_BackupCentral.Text = app_settings["BACKUP_DIR"];                
            } 
            if (!string.IsNullOrEmpty(app_settings["BACKUP_LOCAL"]))
                checkBox_BackupLocal.Checked = true;
            if (!string.IsNullOrEmpty(app_settings["BACKUP_CENTRAL"]))
                checkBox_BackupCentral.Checked = true;            
            if (!string.IsNullOrEmpty(app_settings["NOBACKUP"]))
                checkBox_NoBackup.Checked = true;
            if ((checkBox_BackupLocal.Checked == false) && (checkBox_BackupCentral.Checked == false))
                checkBox_NoBackup.Checked = true;
            if (!string.IsNullOrEmpty(selectedPath))
                SelectPathInTreeView(selectedPath);
            if (!string.IsNullOrEmpty(app_settings["DRIVE_LETTER"]))
                textBox_DriveLetter.Text = app_settings["DRIVE_LETTER"];
            else
                textBox_DriveLetter.Text = "B";

            if (File.Exists(historyFile))
            {
                XDocument doc = XDocument.Load(historyFile);
                foreach (var dm in doc.Descendants("History"))
                {
                    ListViewItem item = new ListViewItem(new string[]
                    {
                    dm.Element("Text").Value,
                    dm.Element("SelectedPath").Value,
                    dm.Element("IncludeSubdirs").Value,
                    dm.Element("FileNames").Value,
                    dm.Element("LastModificationDate").Value,
                    dm.Element("FileAge").Value,
                    dm.Element("Find").Value,
                    dm.Element("IgnoreCase").Value,
                    dm.Element("Replace").Value
                    });
                    listView2.Items.Add(item);
                }
            }
            listView2.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listView2.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }

        private bool SelectPathInTreeView ( string selectedPath)
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
                    textBox_Msg.Text = "Directory: " + selectedPath + " does not exist.";
                    textBox_Msg.BackColor = Color.Orange;
                    return false;
                }                    
            }
        }

        private bool PreExpand(string text, bool root_node, bool expand_node)
        {
            if (root_node)
            {
                foreach (TreeNode t in treeView1.Nodes)
                {                    
                    if (String.Equals(t.Text, text, StringComparison.OrdinalIgnoreCase))                    
                    {
                        treeView1.SelectedNode = t;
                        if (expand_node) { treeView1.SelectedNode.Expand(); }
                        else { treeView1.Select(); }
                        return true;
                    }
                }                
                textBox_Msg.Text = "Directory: " + selectedPath + " does not exist in current Tree View.";
                textBox_Msg.BackColor = Color.Orange;
                return false;
            }
            else
            {
                foreach (TreeNode t in treeView1.SelectedNode.Nodes)
                {                    
                    if (String.Equals(t.Text, text, StringComparison.OrdinalIgnoreCase))
                    {
                        treeView1.SelectedNode = t;
                        if (expand_node) { treeView1.SelectedNode.Expand(); }
                        else { treeView1.Select(); }
                        return true;
                    }
                }
            }
            return false;
        }

        private void TreeView1_BeforeSelect(object sender, TreeViewCancelEventArgs e)
        { 
            TreeNode newSelected = e.Node;
            if (newSelected.ImageIndex == 5)
            {
                MessageBox.Show("No authorization for this directory.");
                treeView1.SelectedNode = null;
                selectedPath = "";
                return;
            }
            selectedPath = newSelected.FullPath;
            newSelected.SelectedImageIndex = 6;
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

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (selectedPath != string.Empty) AddUpdateAppSettings("PATH", selectedPath);
            else DeleteAppSettings("PATH");
            if (textBox_fileNames.Text != string.Empty) AddUpdateAppSettings("FILENAMES", textBox_fileNames.Text);
            else DeleteAppSettings("FILENAMES");
            if (textBox_Find.Text != string.Empty) AddUpdateAppSettings("FIND", textBox_Find.Text);
            else DeleteAppSettings("FIND");
            if (textBox_Replace.Text != string.Empty) AddUpdateAppSettings("REPLACE", textBox_Replace.Text);
            else DeleteAppSettings("REPLACE");
            if (checkBox_IgnoreCase.Checked) AddUpdateAppSettings("IGNORE_CASE", "x");
            else DeleteAppSettings("IGNORE_CASE");
            if (checkBox_Subdir.Checked) AddUpdateAppSettings("INCL_SUBDIRS", "x");
            else DeleteAppSettings("INCL_SUBDIRS");
            if (checkBox_LastModificationDate.Checked) AddUpdateAppSettings("LASTMODIFICATIONDATE", "x");
            else DeleteAppSettings("LASTMODIFICATIONDATE");
            if (numericUpDown_FileAge.Value > 0) AddUpdateAppSettings("FILEAGE", Convert.ToString(numericUpDown_FileAge.Value));
            else DeleteAppSettings("FILEAGE");
            if (checkBox_BackupLocal.Checked) AddUpdateAppSettings("BACKUP_LOCAL", "x");
            else DeleteAppSettings("BACKUP_LOCAL");
            if (checkBox_BackupCentral.Checked) AddUpdateAppSettings("BACKUP_CENTRAL", "x");
            else DeleteAppSettings("BACKUP_CENTRAL");
            if (textBox_BackupCentral.Text != string.Empty) AddUpdateAppSettings("BACKUP_DIR", textBox_BackupCentral.Text);
            else DeleteAppSettings("BACKUP_DIR");
            if (checkBox_NoBackup.Checked) AddUpdateAppSettings("NOBACKUP", "x");
            else DeleteAppSettings("NOBACKUP");
            if (textBox_DriveLetter.Text != string.Empty) AddUpdateAppSettings("DRIVE_LETTER", textBox_DriveLetter.Text);
            else DeleteAppSettings("DRIVE_LETTER");

            if (listView2.Items.Count > 0)
                using (XmlWriter writer = XmlWriter.Create(historyFile))
                {
                    writer.WriteStartDocument();
                    writer.WriteStartElement("Histories");
                    foreach (ListViewItem item in listView2.Items)
                    {
                        writer.WriteStartElement("History");

                        writer.WriteElementString("Text", item.SubItems[0].Text);
                        writer.WriteElementString("SelectedPath", item.SubItems[1].Text);
                        writer.WriteElementString("IncludeSubdirs", item.SubItems[2].Text);
                        writer.WriteElementString("FileNames", item.SubItems[3].Text);
                        writer.WriteElementString("LastModificationDate", item.SubItems[4].Text);
                        writer.WriteElementString("FileAge", item.SubItems[5].Text);
                        writer.WriteElementString("Find", item.SubItems[6].Text);
                        writer.WriteElementString("IgnoreCase", item.SubItems[7].Text);
                        writer.WriteElementString("Replace", item.SubItems[8].Text);                        

                        writer.WriteEndElement();
                    }
                    writer.WriteEndElement();
                    writer.WriteEndDocument();
                }    
            if (NetworkDriveMapped)
            {
                Utility.NetworkDrive.DisconnectNetworkDrive(textBox_DriveLetter.Text, true);
            }
                

        }

        static void AddUpdateAppSettings(string key, string value)
        {
            try
            {
                var config_file = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = config_file.AppSettings.Settings;
                if (settings[key] == null)
                {
                    settings.Add(key, value);
                }
                else
                {
                    settings[key].Value = value;
                }
                config_file.AppSettings.SectionInformation.ForceSave = true;
                config_file.Save(ConfigurationSaveMode.Modified);
                ConfigurationManager.RefreshSection(config_file.AppSettings.SectionInformation.Name);
            }
            catch (ConfigurationErrorsException)
            {
                MessageBox.Show("Error writing app settings (AddUpdateAppSettings)");
            }
        }

        static void DeleteAppSettings(string key)
        {
            try
            {
                var config_file = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                var settings = config_file.AppSettings.Settings;
                if (settings[key] != null)
                {
                    settings.Remove(key);
                }
                config_file.AppSettings.SectionInformation.ForceSave = true;
                config_file.Save(ConfigurationSaveMode.Modified);
            }
            catch (ConfigurationErrorsException)
            {
                MessageBox.Show("Error writing app settings (DeleteAppSettings)");
            }
        }

        private void Do_Find(object sender, EventArgs e)
        {
            FindOrReplace("Find");
        }

        private void Do_Replace(object sender, EventArgs e)
        {
            FindOrReplace("Replace");
        }

        private void FindOrReplace (string action)
        {
            textBox_Msg.BackColor = SystemColors.Control;
            if (string.IsNullOrEmpty(textBox_Find.Text))
            {
                MessageBox.Show("Please enter a \"Find Text\".");
                return;
            }
            if (action == "Replace")
                if (string.IsNullOrEmpty(textBox_Replace.Text))
                {
                    MessageBox.Show("Please enter a \"Replace Text\".");
                    return;
                }

            if (string.IsNullOrEmpty(selectedPath))
            {
                MessageBox.Show("Please select a starting directory.");
                return;
            }
            progressBar1.Visible = true;
            button_Find.Visible = false;
            button_Replace.Visible = false;
            button_Abort.Visible = true;
            label_Progress.Text = action;
            if (string.IsNullOrEmpty(textBox_fileNames.Text)) textBox_fileNames.Text = "*";
            progressBar1.Maximum = 100;
            progressBar1.Step = 1;
            progressBar1.Value = 0;
            ResetPreview();
            listView1.Items.Clear();

            backgroundWorker1.RunWorkerAsync(action);

            tabControl1.SelectedIndex = 0;
        }

        private void ResetPreview()
        {
            textBox_PreviewFind.Text = string.Empty;
            checkBox_PreviewCase.Checked = false;            
            textBox_FilePreview.Text = string.Empty;
            label_PreviewMatches.Text = string.Empty;
            richTextBox_Preview.Text = string.Empty;
        }

        private int GetFiles()
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
            ".png",
            ".pps",
            ".ppt",
            ".pptx",
            ".rar",
            ".rmskin",
            ".rtf",
            ".sqlite",
            ".sys",
            ".tar",
            ".tar.gz",
            ".tif",
            ".tiff",
            ".wav",
            ".wks",
            ".wmv",
            ".wps",
            ".xlr",
            ".xls",
            ".xlsx",
            ".zip"
        };
            int total_files_skipped = 0;
            SearchOption searchOption = SearchOption.TopDirectoryOnly;
            if (checkBox_Subdir.Checked)
            {
                searchOption = SearchOption.AllDirectories;
            }

            string path = selectedPath;
            if (path.Length < 3)
            {
                path = path + "\\";                     // C: returns 0 files - use C:\ instead ...
            }

            list_files.Clear();                          // multiple Find/Replace pressed; reset list_files
            list_directories.Clear();
            list_founds.Clear();
            list_PreviewFounds.Clear();

            DateTime fileCompareDate = DateTime.Now.AddDays(- Convert.ToDouble( numericUpDown_FileAge.Value));            

            char[] delimiters = new char[] { ',', ';', ' ' };
            int x_searchPatterns = 0;            
            string[] searchPatterns = textBox_fileNames.Text.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            string fileExtension;
            foreach (var searchPattern in searchPatterns)
            {
                x_searchPatterns++;
                var files = GetDirectoryFiles(path, searchPattern, searchOption, list_directories, x_searchPatterns);
                foreach (string file in files)
                {
                    fileExtension = Path.GetExtension(file);
                    if (!string.IsNullOrEmpty(fileExtension)) // skip files without Extensions
                    {
                        if (blockFileExtensions.Contains(fileExtension, StringComparer.OrdinalIgnoreCase)) // skip files with Extension in above List  
                            total_files_skipped++;
                        else                        
                        {
                            if (checkBox_LastModificationDate.Checked)
                            {
                                //fileDate = File.GetLastWriteTime(file);
                                if (File.GetLastWriteTime(file).Date >= fileCompareDate.Date) list_files.Add(file);
                            }
                            else
                            {
                                list_files.Add(file);
                            }                                
                        }
                    }                      
                }
            }
            list_files.Sort();
            return total_files_skipped;
        }

        /// <summary>
        /// A safe way to get all the files in a directory and sub directory without crashing on UnauthorizedException or PathTooLongException
        /// </summary>
        /// <param name="rootPath">Starting directory</param>
        /// <param name="patternMatch">Filename pattern match</param>
        /// <param name="searchOption">Search subdirectories or only top level directory for files</param>
        /// <param name="list_directories">List containing directory names</param>
        /// <param name="x_loop">Flag to control Add to list; 1=Add</param>
        /// <returns>List of files</returns>
        public static IEnumerable<string> GetDirectoryFiles(string rootPath, string patternMatch, 
            SearchOption searchOption, List<string> list_directories, int x_loop)
        {
            var foundFiles = Enumerable.Empty<string>();
            if (x_loop == 1)
            {
                list_directories.Add(rootPath);
            }

            if (searchOption == SearchOption.AllDirectories)
            {
                try
                {
                    IEnumerable<string> subDirs = Directory.EnumerateDirectories(rootPath);
                    foreach (string dir in subDirs)
                    {
                        foundFiles = foundFiles.Concat(GetDirectoryFiles(dir, patternMatch, searchOption, list_directories, x_loop)); // Add files in subdirectories recursively to the list
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

        
            
        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string action = (string) e.Argument;
            
            GetFiles();
            int total_directories = list_directories.Count();
            int total_files = list_files.Count();
            int matches = 0;
            int total_matches = 0;
            int total_file_matches = 0;
            int total_files_skipped = GetFiles();
            string backupDir = string.Empty;
            if (checkBox_BackupCentral.Checked)
                backupDir = textBox_BackupCentral.Text;            

            bg_worker1_msg = "0 Files processed.";
            for (int i = 0; i < total_files; i++)
            {
                if (backgroundWorker1.CancellationPending)
                {
                    e.Cancel = true;
                    bg_worker1_msg = bg_worker1_msg + " - Aborted";
                    return;
                }
                backgroundWorker1.ReportProgress((int)((double)i / total_files * 100), list_files[i]);
                matches = FindStringInFile(list_files[i], i, list_founds);
                if (matches > 0)
                {
                    total_matches = total_matches + matches;
                    total_file_matches++;
                    if ( action == "Replace")
                    {                        
                        if (checkBox_BackupLocal.Checked)
                            backupDir = Path.GetDirectoryName(list_files[i]);
                        ReplaceString(backupDir, list_files[i]);
                    }
                }                
            }
            bg_worker1_msg = string.Format("{0} directories, {1} files skipped, {2} files processed: total {3} matches in {4} files.",
                     total_directories.ToString("N0", CultureInfo.InvariantCulture),                     
                     total_files_skipped.ToString("N0", CultureInfo.InvariantCulture),
                     total_files.ToString("N0", CultureInfo.InvariantCulture),
                     total_matches.ToString("N0", CultureInfo.InvariantCulture),
                     total_file_matches.ToString("N0", CultureInfo.InvariantCulture));
        }

        private int FindStringInFile(string file, int id = -1, List<FoundInfo> list_founds = null)
        {
            string text = FileTools.ReadFileString(file);
            int matches = 0;
            RegexOptions options = RegexOptions.Multiline;
            if (checkBox_IgnoreCase.Checked == true) 
                options = RegexOptions.Multiline | RegexOptions.IgnoreCase;
            try
            {
                foreach (Match m in Regex.Matches(text, textBox_Find.Text, options))
                { matches++; }
                if ((matches > 0) && (id > -1) )          
                    list_founds.Add(new FoundInfo(id, matches, 
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


        private bool ReplaceString(string backupDir, string file)
        {
            string fileOld = FileTools.ReadFileString(file);
            RegexOptions options = RegexOptions.Multiline;
            if (checkBox_IgnoreCase.Checked == true)
                options = RegexOptions.Multiline | RegexOptions.IgnoreCase;
            Regex rgx = new Regex(textBox_Find.Text, options);
            string fileNew = rgx.Replace(fileOld, textBox_Replace.Text);
            if (!fileOld.Equals(fileNew))
            {
                FileTools.BackupFile(backupDir, file);
                try
                {
                    File.WriteAllText(file, fileNew, Encoding.Default);
                    textBox_Msg.Text = "Replace: file contents changed.";
                    textBox_Msg.BackColor = Color.LimeGreen;
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error: " + e.Message );
                    textBox_Msg.Text = "Replace: Write failed.";
                    textBox_Msg.BackColor = Color.Orange;
                    return false;
                }
                return true;
            }
            textBox_Msg.Text = "Replace: nothing found to replace in file..";
            textBox_Msg.BackColor = Color.Gold;
            return false;
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar1.Value = e.ProgressPercentage;
            textBox_Msg.Text = "Processing: " + e.UserState as string;
        }

        private void BackgroundWorker1__RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            textBox_Msg.Text = label_Progress.Text + " completed. (" + bg_worker1_msg + ")";
            progressBar1.Value = 0;
            progressBar1.Visible = false;
            button_Find.Visible = true;
            button_Replace.Visible = true;
            button_Abort.Visible = false;
            label_Progress.Text = "Find and Replace";
            listView1.Items.Clear();
            richTextBox_Preview.Text = string.Empty;
            textBox_FilePreview.Text = string.Empty;
            if (list_founds.Count() == 0)
            {
                textBox_Msg.BackColor = Color.Gold;
                return;
            }
            textBox_Msg.BackColor = Color.LimeGreen;
            string[] directoriesOld;
            string[] directoriesNew;
            string directoryNameOld = selectedPath;
            string directoryNameNew;
            string file;
            int indentOffset = 0;

            // ListViewItem (string item text, int image index)
            // 1st Item: Text and Tag = starting directory
            ListViewItem item0 = new ListViewItem(directoryNameOld, 0)
            {
                IndentCount = 0,
                Tag = directoryNameOld
            };
            listView1.Items.Add(item0);
            directoriesOld = directoryNameOld.Split('\\');
            indentOffset = directoriesOld.Count() - 1;
            foreach (FoundInfo foundItem in list_founds)
            {
                file = list_files[foundItem.ID];
                directoryNameNew = Path.GetDirectoryName(file);
                directoriesNew = directoryNameNew.Split('\\');
                if (string.Compare(directoryNameOld, directoryNameNew) == 0)
                {
                    // ListViewItem (string array subitems and text, int image index)
                    // new item in existing directory 
                    ListViewItem item1 = new ListViewItem(new string[]
                        {
                            Path.GetFileName(file),
                            foundItem.Matches.ToString(),
                            foundItem.Date.ToString(),
                            foundItem.Size
                        }, 1)
                    {
                        IndentCount = directoriesNew.Count() - indentOffset,
                        Tag = foundItem.ID
                    };
                    listView1.Items.Add(item1);
                    directoriesOld = directoriesNew;
                    directoryNameOld = directoryNameNew;
                }
                else
                { // add directory tree
                    int oldElements = directoriesOld.Count();
                    int startLoop = indentOffset + 1;
                    directoryNameOld = selectedPath;
                    for (int i = indentOffset + 1; i < directoriesNew.Count(); i++)
                    {
                        directoryNameOld = directoryNameOld + @"\" + directoriesNew[i];                        
                        if (i > oldElements - 1)
                        {
                            ListViewItem item1 = new ListViewItem(directoriesNew[i], 0)
                            {
                                IndentCount = i - indentOffset,
                                Tag = directoryNameOld
                            };
                            listView1.Items.Add(item1);         
                        }
                        else
                        {                           
                            if (string.Compare(directoriesNew[i], directoriesOld[i]) > 0)
                            {
                                ListViewItem item1 = new ListViewItem(directoriesNew[i], 0)
                                {
                                    IndentCount = i - indentOffset,
                                    Tag = directoryNameOld
                                };
                                listView1.Items.Add(item1);
                                oldElements = 0; 
                            }
                        }
                    }
                    // ListViewItem (string array subitems and text, int image index)
                    // new item in new directory 
                    ListViewItem item2 = new ListViewItem(new string[]
                        {
                            Path.GetFileName(file),
                            foundItem.Matches.ToString(),
                            foundItem.Date.ToString(),
                            foundItem.Size
                        }, 1)
                    {
                        IndentCount = directoriesNew.Count() - indentOffset,
                        Tag = foundItem.ID
                    };
                    listView1.Items.Add(item2);
                    directoriesOld = directoriesNew;
                    directoryNameOld = directoryNameNew;
                }
            }
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
        }



        private void Button_Abort(object sender, EventArgs e)
        {
            backgroundWorker1.CancelAsync();
        }

        private void ListView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of list_files
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches            
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;
            ResetPreview();
            foreach (ListViewItem item in fileselection)
            {
                if (item.ImageIndex == 1)  // File selected
                {
                    ID = Convert.ToInt16(item.Tag);
                    file = list_files[ID];
                    textBox_FilePreview.Text = file;
                    richTextBox_Preview.Text = FileTools.ReadFileString(file);                    
                    textBox_PreviewFind.Text = textBox_Find.Text;
                    if (checkBox_IgnoreCase.Checked)
                        checkBox_PreviewCase.Checked = true;
                    else
                        checkBox_PreviewCase.Checked = false;
                    FindInPreview();
                }                
            }
        }

        private void FindInPreview()
        {
            int matches = 0;
            int index = 0;
            list_PreviewFounds.Clear();
            RegexOptions options = RegexOptions.Multiline;
            if (checkBox_PreviewCase.Checked == true)            
                options = RegexOptions.Multiline | RegexOptions.IgnoreCase;            
            try
            {
                foreach (Match m in Regex.Matches(richTextBox_Preview.Text, textBox_PreviewFind.Text, options))
                {
                    matches++;
                    if (matches == 1) { index = m.Index; }
                    ColorKeyword(m.Value, Color.OrangeRed, m.Index);
                    list_PreviewFounds.Add(new PreviewFindClass(matches, m.Value, m.Index));
                    if (matches == 1000)                    // limit highlighting of found strings to 1000
                        break;                              // ... otherwise it takes too long and is not responding
                }
                preview_match_x = 0;
                PreviewFindMatchToggler("down");
                richTextBox_Preview.SelectionStart = index;  // Scroll to 1st match in richTextBox
                richTextBox_Preview.ScrollToCaret();             
              
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

        private void PreviewFindMatchToggler(string direction)
        {
            richTextBox_Preview.SelectAll();            
            richTextBox_Preview.SelectionBackColor = SystemColors.ControlLight;
            if (list_PreviewFounds.Count > 0)
            {
                if (direction == "down")
                {
                    if (preview_match_x == list_PreviewFounds.Count)
                        preview_match_x = 1;
                    else
                        preview_match_x++;
                }
                else
                {
                    if (preview_match_x == 1)
                        preview_match_x = list_PreviewFounds.Count;
                    else
                        preview_match_x--;
                }                
                PreviewFindClass List = list_PreviewFounds[preview_match_x - 1];                
                richTextBox_Preview.Select(List.Position, List.Value.Length);                
                richTextBox_Preview.SelectionBackColor = Color.Gold;
                richTextBox_Preview.ScrollToCaret();
            }
            label_PreviewMatches.Text = preview_match_x + "/" + list_PreviewFounds.Count;
        }
                

        private void ColorKeyword(string word, Color color, int startIndex)
        {
            richTextBox_Preview.Select(startIndex, word.Length);
            richTextBox_Preview.SelectionColor = color;
        }

        private void Do_PreviewFind(object sender, EventArgs e)
        {            
            if (!string.IsNullOrEmpty(textBox_PreviewFind.Text))
            {
                richTextBox_Preview.SelectAll();
                richTextBox_Preview.SelectionColor = SystemColors.WindowText;
                FindInPreview();
            }
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
        }


        private void RichTextBoxAbout_LinkClicked(object sender, LinkClickedEventArgs e)
        {           
            System.Diagnostics.Process.Start(e.LinkText);
        }

        private void PreviewMatchUp(object sender, EventArgs e)
        {
            PreviewFindMatchToggler("up");
        }

        private void PreviewMatchDown(object sender, EventArgs e)
        {
            PreviewFindMatchToggler("down");
        }      
        

        private void RichTextBox_Preview__PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
        {
            if (e.KeyCode == Keys.F3)
            {
                PreviewFindMatchToggler("down");
            }
        }

        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {            
            Process.Start("explorer.exe", treeView1.SelectedNode.FullPath);            
        }

        private void ExplorerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of list_files
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches            
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;
            ProcessStartInfo startInfo = new ProcessStartInfo();
            foreach (ListViewItem item in fileselection)
            {
                ID = Convert.ToInt16(item.Tag);
                file = list_files[ID];
                startInfo.FileName = "explorer.exe";
                startInfo.Arguments = @"/n, /select, " + file;
                Process.Start(startInfo);
            }
        }

        private void NotepadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of list_files
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches            
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;
            ProcessStartInfo startInfo = new ProcessStartInfo();

            foreach (ListViewItem item in fileselection)
            {
                ID = Convert.ToInt16(item.Tag);
                file = list_files[ID];
                startInfo.FileName = "notepad.exe";
                startInfo.Arguments = file;
                var process = Process.Start(startInfo);     // start notepad 
                process.WaitForExit();                      // ... wait for its exit - so we can properly refresh listview
                item.SubItems[1].Text = FindStringInFile(file).ToString();
                item.SubItems[2].Text = File.GetLastWriteTime(file).ToString();
                item.SubItems[3].Text = FileTools.GetFileSizeHuman(file);
                ListView1_SelectedIndexChanged(listView1, EventArgs.Empty);  // Refresh File Preview after restore
            }
        }

        private void Button_BackupCentral_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDlg = new FolderBrowserDialog
            {
                ShowNewFolderButton = false,
                SelectedPath = textBox_BackupCentral.Text,
                RootFolder = Environment.SpecialFolder.MyComputer
            };
            // Show the FolderBrowserDialog.
            DialogResult result = folderDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox_BackupCentral.Text = folderDlg.SelectedPath;
            }
        }

        private void CheckBox_BackupLocal_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_BackupLocal.Checked)
            {
                checkBox_BackupCentral.Checked = false;
                checkBox_NoBackup.Checked = false;
            }
        }

        private void CheckBox_BackupCentral_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_BackupCentral.Checked)
            {
                checkBox_BackupLocal.Checked = false;
                checkBox_NoBackup.Checked = false;
                if (string.IsNullOrEmpty(textBox_BackupCentral.Text))
                {
                    tabControl1.SelectedIndex = 1;
                    MessageBox.Show("Please define a valid Archive directory.");
                }
                                 
                else                
                    if (!Directory.Exists(textBox_BackupCentral.Text))
                {
                    tabControl1.SelectedIndex = 1;
                    MessageBox.Show(string.Format("Archive directory {0} not found, please define a valid directory",
                            textBox_BackupCentral.Text));
                }
                                    
            }  
        }

        private void CheckBox_NoBackup_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_NoBackup.Checked)
            {
                checkBox_BackupLocal.Checked = false;
                checkBox_BackupCentral.Checked = false;
            }
        }

        private void TextBox_BackupCentral_TextChanged(object sender, EventArgs e)
        {
            if (!Directory.Exists(textBox_BackupCentral.Text))
                MessageBox.Show(string.Format("Directory {0} not found, please define a valid directory",
                            textBox_BackupCentral.Text));
        }

        private void BackupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of list_files
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches   
            string backupDir;   // Backup local or central
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;
            if (checkBox_NoBackup.Checked)
            {
                MessageBox.Show("Please select a Backup Location in tab \"More Options\".");
                return;
            }
            foreach (ListViewItem item in fileselection)
            {
                ID = Convert.ToInt16(item.Tag);
                file = list_files[ID];
                if (checkBox_BackupLocal.Checked)
                    backupDir = Path.GetDirectoryName(file);
                else
                    backupDir = textBox_BackupCentral.Text;
                if (FileTools.BackupFile(backupDir, file))
                    textBox_Msg.Text = "Backup: file archived in " + backupDir + @"\" + FileTools.zipArchiveName;
                    textBox_Msg.BackColor = Color.LimeGreen;
            }
        }


        private void RestoreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of list_files
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches  
            string backupDir;   // Backup local or central
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;
            if (checkBox_NoBackup.Checked)
            {
                MessageBox.Show("Please select a Backup Location in tab \"More Options\".");
                return;
            }

            foreach (ListViewItem item in fileselection)
            {
                ID = Convert.ToInt16(item.Tag);
                file = list_files[ID];
                if (checkBox_BackupLocal.Checked)
                    backupDir = Path.GetDirectoryName(file);
                else
                    backupDir = textBox_BackupCentral.Text;

                switch (FileTools.CountBackupVersions(backupDir, file))
                {
                    case 0:
                        textBox_Msg.Text = "Restore: no archived version of that file found";
                        textBox_Msg.BackColor = Color.Gold;
                        break;
                    case 1:
                        if (FileTools.RestoreFile(backupDir, file, 1))
                        {
                            item.SubItems[1].Text = FindStringInFile(file).ToString();
                            item.SubItems[2].Text = File.GetLastWriteTime(file).ToString();
                            item.SubItems[3].Text = FileTools.GetFileSizeHuman(file);
                            ListView1_SelectedIndexChanged(listView1, EventArgs.Empty);  // Refresh File Preview after restore
                            textBox_Msg.Text = "Restore: file restored.";
                            textBox_Msg.BackColor = Color.LimeGreen;
                        }
                        else
                        {
                            textBox_Msg.Text = "Restore: Restore failed.";
                            textBox_Msg.BackColor = Color.Orange;
                        }                        
                        break;
                    default:
                        Form2 f = new Form2(backupDir, file);
                        if (f.ShowDialog(this) == DialogResult.OK)
                        {
                            item.SubItems[1].Text = FindStringInFile(file).ToString();
                            item.SubItems[2].Text = File.GetLastWriteTime(file).ToString();
                            item.SubItems[3].Text = FileTools.GetFileSizeHuman(file);
                            ListView1_SelectedIndexChanged(listView1, EventArgs.Empty);  // Refresh File Preview after restore
                            textBox_Msg.Text = "Restore: file restored.";
                            textBox_Msg.BackColor = Color.LimeGreen;
                        }
                        f.Dispose();
                        break;
                }
            }
        }

        private void ReplaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int ID;             // TAG contains ID of list_files
            string file;        // ID -> list_founds -> match-values from REGEXP.Matches   
            string backupDir;   // Backup central, local or Empty
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;

            foreach (ListViewItem item in fileselection)
            {                
                ID = Convert.ToInt16(item.Tag);
                file = list_files[ID];
                if (checkBox_BackupLocal.Checked)
                    backupDir = Path.GetDirectoryName(file);
                else
                {
                    if (checkBox_BackupCentral.Checked)
                        backupDir = textBox_BackupCentral.Text;
                    else
                        backupDir = string.Empty;
                }
                if (ReplaceString(backupDir, file))
                {
                    item.SubItems[1].Text = FindStringInFile(file).ToString();
                    item.SubItems[2].Text = File.GetLastWriteTime(file).ToString();
                    item.SubItems[3].Text = FileTools.GetFileSizeHuman(file);
                    ListView1_SelectedIndexChanged(listView1, EventArgs.Empty);  // Refresh File Preview after restore                    
                }                
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
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
            // don't forget to save the settings
            Properties.Settings.Default.Save();
        }

        private void ExplorerToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;
            ProcessStartInfo startInfo = new ProcessStartInfo();
            foreach (ListViewItem item in fileselection)
            {
                Process.Start("explorer.exe", Convert.ToString(item.Tag));
            }
        }

        private void ListView1_MouseClick(object sender, MouseEventArgs e)
        {
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;            
            if (e.Button == MouseButtons.Right)
            {
                foreach (ListViewItem item in fileselection)
                {
                    if (item.ImageIndex == 1)   // File clicked - ContextMenu 2                   
                        this.contextMenuStrip2.Show(this.listView1, e.Location);
                    else                        // Folder clicked - ContextMenu 3
                        this.contextMenuStrip3.Show(this.listView1, e.Location);                    
                }                
            }
        }

        private void ListView1_DoubleClick(object sender, MouseEventArgs e)
        {            
            ListView.SelectedListViewItemCollection fileselection = this.listView1.SelectedItems;
            foreach (ListViewItem item in fileselection)
            {
                if (item.ImageIndex == 1)   // File double clicked - ContextMenu 2 
                    Process.Start(list_files[Convert.ToInt16(item.Tag)]);
                else                        // Folder double clicked - ContextMenu 3
                    Process.Start("explorer.exe", Convert.ToString(item.Tag));                
            }            
        }

        private void ListView2_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)                
                    lvwColumnSorter.Order = SortOrder.Descending;                
                else                
                    lvwColumnSorter.Order = SortOrder.Ascending;                
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.listView2.Sort();
        }

        private void LoadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            treeView1.CollapseAll();
            foreach (ListViewItem item in listView2.SelectedItems)
            {
                selectedPath = item.SubItems[1].Text;
                if (item.SubItems[2].Text == "True")
                    checkBox_Subdir.Checked = true;
                else
                    checkBox_Subdir.Checked = false;
                textBox_fileNames.Text = item.SubItems[3].Text;
                if (item.SubItems[4].Text == "True")
                    checkBox_LastModificationDate.Checked = true;
                else
                    checkBox_LastModificationDate.Checked = false;
                numericUpDown_FileAge.Value = Convert.ToDecimal(item.SubItems[5].Text);
                textBox_Find.Text = item.SubItems[6].Text;
                if (item.SubItems[7].Text == "True")
                    checkBox_IgnoreCase.Checked = true;
                else
                    checkBox_IgnoreCase.Checked = false;
                textBox_Replace.Text = item.SubItems[8].Text;
                textBox_Msg.Text = "Favorites: entry loaded.";
                textBox_Msg.BackColor = Color.LimeGreen;
                if (!string.IsNullOrEmpty(selectedPath))
                    if (SelectPathInTreeView(selectedPath) == false)
                    {
                        textBox_Msg.Text = "Error - favorites: entry loaded, but directory: " + selectedPath + " does not exist.";
                        textBox_Msg.BackColor = Color.Orange;
                        selectedPath = string.Empty;
                    }
                        
            }
        }

        private void DeleteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView2.SelectedItems)
            {
                listView2.Items.Remove(item);
                textBox_Msg.Text = "Favorites: entry deleted.";
                textBox_Msg.BackColor = Color.LimeGreen;
            }
        }
            

        private void EditToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in listView2.SelectedItems)
            {
                string input = Microsoft.VisualBasic.Interaction.InputBox("Modify description of selected item", "Edit Favorites", item.SubItems[0].Text, -1, -1);                
                if (!string.IsNullOrEmpty(input))
                {
                    item.SubItems[0].Text = input;                    
                    textBox_Msg.Text = "Favorites: description edited.";
                    textBox_Msg.BackColor = Color.LimeGreen;
                }
            }
        }

        private void Button_Add_Click(object sender, EventArgs e)
        {
            string input = Microsoft.VisualBasic.Interaction.InputBox("Add description for the actual search", "Add Favorites", "", -1, -1);
            if (!string.IsNullOrEmpty(input))
            {                
                ListViewItem item = new ListViewItem(new string[]
                    {
                    input,
                    selectedPath,
                    (checkBox_Subdir.Checked ) ? "True" : "False",
                    textBox_fileNames.Text,
                    (checkBox_LastModificationDate.Checked ) ? "True" : "False",
                    numericUpDown_FileAge.Value.ToString(),
                    textBox_Find.Text,
                    (checkBox_IgnoreCase.Checked ) ? "True" : "False",
                    textBox_Replace.Text
                    });
                listView2.Items.Add(item);
                textBox_Msg.Text = "Favorites: new entry added.";
                textBox_Msg.BackColor = Color.LimeGreen;
            }
        }

        private void LoadAndFindToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadToolStripMenuItem_Click(sender, e);
            if (!string.IsNullOrEmpty(selectedPath))
            {
                Do_Find(sender, e);
            }
                
            
        }

        private void Button_Check4Update_Click(object sender, EventArgs e)
        {
            UpdateStatus status = Updater.CheckForUpdate(ShowUpdateDialog);
            if (status == UpdateStatus.UpdateFailed)
                MessageBox.Show(this, "Failed to check for update.  Please ty again later.", "Warning");
            else if (status == UpdateStatus.NoUpdate)
                MessageBox.Show(this, "There are no updates available at this time.", "Update Check");
        }

        private void ListView2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            LoadAndFindToolStripMenuItem_Click(sender, e);
        }

        private void ListView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (Keys.Delete == e.KeyCode)
            {
                foreach (ListViewItem item in listView2.SelectedItems)
                {
                    listView2.Items.Remove(item);
                    textBox_Msg.Text = "Favorites: entry deleted.";
                    textBox_Msg.BackColor = Color.LimeGreen;
                }
            }
        }

        private void ListView2_SelectedIndexChanged(object sender, EventArgs e)
        {                   
            ListView.SelectedListViewItemCollection FavoritesSelection = this.listView2.SelectedItems;            
            foreach (ListViewItem item in FavoritesSelection)
            {
                item.Selected = true;                
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex ==2)
            {
                listView2.Focus();
            }
        }

        private void Leave_textBox_DriveLetter(object sender, EventArgs e)
        {
            if (!Regex.IsMatch(textBox_DriveLetter.Text,@"[a-zA-Z]"))
            {
                textBox_DriveLetter.Select();
                MessageBox.Show("Only letters are valid.");
            }
                
        }

        private void button_CheckForUpdate_Click(object sender, EventArgs e)
        {
            UpdateStatus status = Updater.CheckForUpdate(ShowUpdateDialog);
            if (status == UpdateStatus.UpdateFailed) 
                MessageBox.Show(this, "Check for update failed. Please try later", "Warning");
            else if (status == UpdateStatus.NoUpdate)
                MessageBox.Show(this, "There are no updates available.", "Update Check");
        }
    }
}
