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
           
            textBox_StartCopyDelayDirect.Text = Properties.Settings.Default.DelayTime.ToString();
            textBox_StartCopyDelayFile.Text = Properties.Settings.Default.DelayTime.ToString();
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
                    textBox_StartCopyDelayDirect.Enabled = false;
                    textBox_StartCopyDelayFile.Enabled = false;
                    break;
                case ButtonState.READY:
                    btn_Cancel1.Enabled = false;
                    btn_SelectFile.Enabled = true;
                    btn_StartCopyDirect.Enabled = true;
                    btn_StartCopyFile.Enabled = true;
                    textBox_StartCopyDelayDirect.Enabled = true;
                    textBox_StartCopyDelayFile.Enabled = true;
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
                label_Status.Text = "Loading File...";
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
                    typer.ih.LoadDriver();
                    if (!bg.CancellationPending)
                    {
                        float dTime = Properties.Settings.Default.DelayTime;
                        bg.ReportProgress((int)dTime);
                        while (dTime >= 1)
                        {
                            if (!bg.CancellationPending)
                            {
                                dTime--;
                                Thread.Sleep(1000);
                                bg.ReportProgress((int)dTime);
                            }
                            else
                            {
                                e.Cancel = true;
                                break;
                            }
                        }
                        typer.TypeCSVArray(reader.GetArrayStorage(), bg);
                        if (bg.CancellationPending)
                        {
                            e.Cancel = true;
                        }
                    }
                    else
                    {
                        e.Cancel = true;
                    }
                    
                }
            }
         
            finally
            {
               
            }
        }
        private void BgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage > 0)
            {
                label_Status.Text = "Press Any Key at least Once \n Starting in " + e.ProgressPercentage;
            }
            else
            {
                label_Status.Text = "Copying...";
            }
            
        }
        private void BgWorker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            EnableButtons(ButtonState.READY);
            if (e.Cancelled)
            {
                label_Status.Text = "Canceled";
            }
            else
            {
                label_Status.Text = "Finished";
            }
        }

        private void btn_Cancel1_Click(object sender, EventArgs e)
        {
            CancelBGWorker();
        }
        private void CancelBGWorker()
        {
            BgWorker.CancelAsync();
        }

        private void textBox_StartCopyDelayFile_TextChanged(object sender, EventArgs e)
        {
            float newValue = 0;
            bool isNumber = float.TryParse(textBox_StartCopyDelayFile.Text, out newValue);
            if (isNumber)
            {
                if (newValue < 60)
                {
                    Properties.Settings.Default.DelayTime = newValue;

                }

                textBox_StartCopyDelayDirect.Text = Properties.Settings.Default.DelayTime.ToString();
                textBox_StartCopyDelayFile.Text = Properties.Settings.Default.DelayTime.ToString();
            }
           

        }
    }
}
