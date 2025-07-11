using System;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.IO.Compression;

namespace FindReplace
{
    class FileTools
    {
        public const string zipArchiveName = @"FindReplace_Backup.zip";

        public static bool BackupFile(string backup_dir, string file)
        {
            if (string.IsNullOrEmpty(backup_dir))
                return true;
            string zipFile = backup_dir + @"\" + zipArchiveName;
            try
            {
                if (!File.Exists(zipFile))        // create empty .zp file if not exists    
                {
                    using (ZipArchive newFile = ZipFile.Open(zipFile, ZipArchiveMode.Create))
                    {
                    }
                }
                if (!FileExistsInArchive(zipFile, file))
                {
                    using (ZipArchive zipArchive = ZipFile.Open(zipFile, ZipArchiveMode.Update))
                    {
                        zipArchive.CreateEntryFromFile(file, file, CompressionLevel.Optimal);
                    }
                }
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show("Backup: Error" + Environment.NewLine + e.Message);
                return false;
            }            
        }

        public static bool RestoreFile(string backup_dir, string file, int fileNumber)
        {
            string zipFile = backup_dir + @"\" + zipArchiveName;
            int fileCounter = 0;
            try
            {
                if (File.Exists(zipFile))
                {
                    using (ZipArchive zipArchive = ZipFile.OpenRead(zipFile))
                    {
                        foreach (ZipArchiveEntry entry in zipArchive.Entries)
                        {
                            if (entry.FullName == file)
                            {
                                fileCounter++;
                                if (fileNumber == fileCounter)
                                {
                                    //MessageBox.Show(string.Format("Restore File: {0} \n LastWriteTime: {1}", 
                                        //entry.FullName, entry.LastWriteTime.DateTime));     
                                    entry.ExtractToFile(file, true);                  
                                }
                            }
                                
                        }
                    }
                }               
            }
            catch (Exception e)
            {
                MessageBox.Show("Restore: Error" + Environment.NewLine + e.Message);
                return false;
            }
            return true;
        }

        private static bool FileExistsInArchive(string zipFile, string file)
        {
            if (!File.Exists(zipFile))
                return false;
            DateTime fileLastWriteTime = File.GetLastWriteTime(file);
            using (ZipArchive zipArchive = ZipFile.OpenRead(zipFile))
            {
                foreach (ZipArchiveEntry entry in zipArchive.Entries)
                {
                    if (entry.FullName == file)
                    {
                        TimeSpan timeSpan = fileLastWriteTime - entry.LastWriteTime.DateTime;
                        //MessageBox.Show(String.Format("File:   {0}\nArch: {1}\nTimeSpan: {2}", 
                        //    fileLastWriteTime, entry.LastWriteTime.DateTime, timeSpan.Seconds ));                        
                        // Microsoft: ..  DateTimeOffset value is converted to a timestamp format that is specific to zip archives. 
                        // This format supports a resolution of two seconds.
                        if (timeSpan.Seconds < 3)
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        public static int CountBackupVersions(string backup_dir, string file)
        {
            int fileCounter = 0;
            string zipFile = backup_dir + @"\" + zipArchiveName;
            if (!File.Exists(zipFile))
                return 0;           
            using (ZipArchive zipArchive = ZipFile.OpenRead(zipFile))
            {
                foreach (ZipArchiveEntry entry in zipArchive.Entries)
                {
                    if (entry.FullName == file)
                        fileCounter++;                    
                }
            }
            return fileCounter;
        }

        public static string GetFileSizeHuman(string file, double len = 0)
        {
            string[] sizes = { "B", "KB", "MB", "GB", "TB" };
            int order = 0;
            if (len == 0)
                len = new FileInfo(file).Length;            
            while (len >= 1024 && order < sizes.Length - 1)
            {
                order++;
                len = len / 1024;
            }
            return String.Format("{0} {1}", Math.Ceiling(len), sizes[order]);
        }

        public static string ReadFileString(string path)
        {
            try
            {
                // Use StreamReader to consume the entire text file.
                using (StreamReader reader = new StreamReader(path, Encoding.Default, true))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception)
            {
                // Fixing Errorsituation "The process cannot access the file 'filename' 
                //                        because it is being used by another process"
                // Example: file dsmerror.log, process TSM Server
                return ReadFileString2(path);
            }
        }

        public static string ReadFileString2(string path)
        {
            try
            {
                using (FileStream fs = File.Open(path, System.IO.FileMode.Open,
                                                        System.IO.FileAccess.Read,
                                                        System.IO.FileShare.ReadWrite))
                using (BufferedStream bs = new BufferedStream(fs))               
                using (StreamReader sr = new StreamReader(bs, Encoding.Default, true))
                {
                    string s;
                    StringBuilder builder = new StringBuilder();
                    while ((s = sr.ReadLine()) != null)
                    {
                        builder.Append(s);
                        builder.Append(Environment.NewLine);
                    }
                    return builder.ToString();
                }
            }
            catch (Exception)
            {
                //MessageBox.Show("ReadFileString2 Error: " + e.Message);
                //Debug.WriteLine("ReadFileString2 Error: " + e.Message);
                return "";
            }
        }
    }
}
