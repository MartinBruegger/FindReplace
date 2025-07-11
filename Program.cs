// Program.cs
//
// Copyright 2025 Martin Bruegger

using System;
using System.Windows.Forms;

namespace FindReplace
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Updater.UpdateUpdater();    // Update the updater when file not in use.
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1(args));
        }
    }
}
