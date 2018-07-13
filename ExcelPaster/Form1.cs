using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelPaster
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            if (Properties.Settings.Default.RecentFiles == null)
            {
                Properties.Settings.Default.RecentFiles = new System.Collections.Specialized.StringCollection();
            }
            else
            {
                foreach(string filename in Properties.Settings.Default.RecentFiles)
                {
                    comboBox_FileLocation.Items.Add(filename);
                }
            }
        }
        public enum ButtonState
        {
            READY = 0,
            COPYING = 1
        }
        private void EnableButtons(ButtonState mode)
        {
            switch (mode)
            {
                case ButtonState.COPYING:
                    btn_Cancel1.Enabled = true;
                    btn_SelectFile.Enabled = false;
                    btn_StartCopyDirect.Enabled = false;
                    btn_StartCopyFile.Enabled = false;
                    break;
                case ButtonState.READY:
                    btn_Cancel1.Enabled = false;
                    btn_SelectFile.Enabled = true;
                    btn_StartCopyDirect.Enabled = true;
                    btn_StartCopyFile.Enabled = true;
                    break;
                default:
                    break;
            }
        }

        private void SetFileMostRecent(string file)
        {
            Properties.Settings.Default.RecentFiles.Remove(file);
            Properties.Settings.Default.RecentFiles.Insert(0,file);
            Properties.Settings.Default.Save();
            foreach (string filename in Properties.Settings.Default.RecentFiles)
            {
                comboBox_FileLocation.Items.Add(filename);
            }

        }
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void btn_SelectFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = openFileDialog1.FileName;
                if ( !string.IsNullOrWhiteSpace(result))
                {
                    comboBox_FileLocation.Text = result;
                    if (!Properties.Settings.Default.RecentFiles.Contains(result))
                    {
                        Properties.Settings.Default.RecentFiles.Add(result);
                        Properties.Settings.Default.Save();
                    }
                    else
                    {
                        SetFileMostRecent(result);
                    }
                }
            }
            
 
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void btn_StartCopyFile_Click(object sender, EventArgs e)
        {
            string CSVFile = comboBox_FileLocation.Text;
            if (CSVFile.Count() > 0)
            {
                EnableButtons(ButtonState.COPYING);

                if (!BgWorker.IsBusy)
                {
                    BgWorker.RunWorkerAsync(CSVFile);
                }
            }
            else
            {
                label_Status.Text = "No File Selected";
            }
        }
        private void BgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bg = sender as BackgroundWorker;
            string fileLoc = (string)e.Argument;
            try
            {
                FileInfo fInfo = new FileInfo(fileLoc);
                if (!fInfo.Exists)
                {
                   
                    e.Cancel = true;
                  
                }
                if (fInfo.Extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                {
                    CSVReader reader = new CSVReader();
                    reader.ParseCSV(fInfo.FullName);
                    Typer typer = new Typer();
                    Thread.Sleep(5000);
                    typer.TypeCSVArray(reader.GetArrayStorage());
                }
            }
         
            finally
            {
                EnableButtons(ButtonState.READY);
            }
        }
    }
}
