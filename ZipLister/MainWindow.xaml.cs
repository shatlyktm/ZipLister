/*
The MIT License (MIT)

Copyright (c) 2014 Charles Justin Henck

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.IO;
using Ionic.Zip;
using Kent.Boogaart.KBCsv;

namespace ZipLister
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }


        private void tbSelectOut_Click(object sender, RoutedEventArgs e)
        {
            // Configure open file dialog box
            string path = tbZipFile.Text.ToLower();
            string filename = tbZipFile.Text.Split('\\').Last();
            path = path.Remove(path.Length - filename.Length);
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.InitialDirectory = path;
            dlg.FileName = filename.ToLower().Replace(".zip", "_contents.csv");
            dlg.DefaultExt = ".csv"; // Default file extension
            dlg.Filter = "CSV File|*.csv"; // Filter files by extension 

            // Show open file dialog box
            bool? result = dlg.ShowDialog();

            // Process open file dialog box results 
            if (result == true)
            {
                // Open document 
                tbOutputCSV.Text = dlg.FileName;
            }
        }

        private void tbSelectZip_Click(object sender, RoutedEventArgs e)
        {
            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "ToList"; // Default file name
            dlg.DefaultExt = ".zip"; // Default file extension
            dlg.Filter = "Zip Files|*.zip"; // Filter files by extension 

            // Show open file dialog box
            bool? result = dlg.ShowDialog();

            // Process open file dialog box results 
            if (result == true)
            {
                // Open document 
                tbZipFile.Text = dlg.FileName;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CsvWriter outCSV = new CsvWriter(tbOutputCSV.Text);
                outCSV.WriteRecord("TYPE","PATH", "FILENAME");
                System.IO.Stream s = new FileStream(tbZipFile.Text, FileMode.Open);
                String zipName = tbZipFile.Text.Split('\\').Last();
                RecurseZip("\\" + zipName, s, outCSV);
                outCSV.Close();
                if (MessageBox.Show("Complete! Would you like to open the file?", "Complete!", MessageBoxButton.YesNo) == MessageBoxResult.Yes) 
                    System.Diagnostics.Process.Start(tbOutputCSV.Text);
            }
            catch (Exception excp)
            {
                MessageBox.Show("Error: " + excp.Message);
            }
        }

        private void RecurseZip(string path, Stream toReadStream, CsvWriter outCSV)
        {
            string existingZipName = path.Split('/').Last();
            ZipFile existingZip = null;
            existingZip = ZipFile.Read((Stream)toReadStream);
            foreach (ZipEntry ze in existingZip.Entries)
            {
                if (ze.IsDirectory)
                    continue;
                List<string> parts = new List<string>(ze.FileName.Split('/'));
                string filename = parts[parts.Count - 1];
                if (filename.StartsWith("."))
                    continue; 
                parts.RemoveAt(parts.Count - 1);
                string pathToEntry = path + "\\" + String.Join("\\",parts);
                outCSV.WriteRecord("IN ZIP", pathToEntry, filename);
                if (filename.EndsWith(".zip", StringComparison.CurrentCultureIgnoreCase))
                {
                    MemoryStream nestedZip = new MemoryStream((int)ze.UncompressedSize);
                    ze.Extract(nestedZip);
                    nestedZip.Position = 0;
                    this.RecurseZip(pathToEntry + filename, nestedZip, outCSV);
                    nestedZip.Close();
                }
            }
            toReadStream.Close();
        }

        private void RecurseFolder(string relativeTo, string folderPath, CsvWriter outCSV)
        {
            string path = relativeTo + folderPath;
            foreach (string dirFilename in Directory.EnumerateFiles(path))
            {
                List<string> parts = new List<string>(dirFilename.Split('\\'));
                string filename = parts[parts.Count - 1];
                outCSV.WriteRecord("IN FOLDER", folderPath , filename);
                if (filename.EndsWith(".zip", StringComparison.CurrentCultureIgnoreCase))
                {
                    Stream s = new FileStream(dirFilename,FileMode.Open);
                    RecurseZip(folderPath + filename,s,outCSV);

                }
            }
            foreach (string dirname in Directory.EnumerateDirectories(path))
            {
                RecurseFolder(relativeTo, dirname.Replace(relativeTo,""), outCSV);
            }
        }

        private void tbSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            // Configure open file dialog box

            System.Windows.Forms.FolderBrowserDialog fbd = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult dr = fbd.ShowDialog();      

            // Show open file dialog box
      

            // Process open file dialog box results 
            if (dr.HasFlag(System.Windows.Forms.DialogResult.OK))
            {
                // Open document 
                tbFolder.Text = fbd.SelectedPath;
            }
        }

        private void tbSelectOutFolder_Click(object sender, RoutedEventArgs e)
        {
            // Configure open file dialog box
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.InitialDirectory = tbFolder.Text;
            dlg.FileName = tbFolder.Text.Split(new char[]{'\\'},StringSplitOptions.RemoveEmptyEntries).Last() + "_contents.csv"; // Default file name
            dlg.DefaultExt = ".csv"; // Default file extension
            dlg.Filter = "CSV File|*.csv"; // Filter files by extension 

            // Show open file dialog box
            bool? result = dlg.ShowDialog();

            // Process open file dialog box results 
            if (result == true)
            {
                // Open document 
                tbOutputCSVFolder.Text = dlg.FileName;
            }
        }

        private void ButtonFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CsvWriter outCSV = new CsvWriter(tbOutputCSVFolder.Text);
                outCSV.WriteRecord("TYPE", "PATH", "FILENAME");
                RecurseFolder(tbFolder.Text, "\\", outCSV);
                outCSV.Close();
                if (MessageBox.Show("Complete! Would you like to open the file?", "Complete!", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                    System.Diagnostics.Process.Start(tbOutputCSVFolder.Text);
            }
            catch (Exception excp)
            {
                MessageBox.Show("Error: " + excp.Message);
            }
        }
    }

}
