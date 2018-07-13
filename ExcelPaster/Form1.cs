using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
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
        private void EnableButtons(string mode)
        {
            switch (mode)
            {
                case "COPYING":
                    btn_Cancel1.Enabled = true;
                    btn_SelectFile.Enabled = false;
                    btn_StartCopyDirect.Enabled = false;
                    btn_StartCopyFile.Enabled = false;
                    break;
                case "READY":
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
                //System.IO.StreamReader sr = new
                //   System.IO.StreamReader(openFileDialog1.FileName);
                //MessageBox.Show(sr.ReadToEnd());
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
                //sr.Close();
            }
            
            //using (var fbd = openFileDialog1)
            //{
            //    DialogResult result = fbd.ShowDialog();

            //    if (result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            //    {
            //        comboBox_FileLocation.Text = fbd.SelectedPath;
            //        if (!Properties.Settings.Default.RecentFiles.Contains(fbd.SelectedPath))
            //        {
            //            Properties.Settings.Default.RecentFiles.Add(fbd.SelectedPath);
            //        }
            //        else
            //        {
            //            SetFileMostRecent(fbd.SelectedPath);
            //        }
            //    }
            //}
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

    
    }
}
