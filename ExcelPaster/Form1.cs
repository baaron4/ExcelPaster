using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ModbusTCP;
using Syncfusion.XPS;
using System.Windows.Automation;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using System.Runtime.InteropServices.ComTypes;
using System.Net.Http.Headers;
using System.Linq.Expressions;
using System.Net.Mail;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ExcelPaster
{
    public partial class MainForm : Form
    {
        IPAddress addressIP;
        IPAddress submask;
        IPAddress gateway;
        Dictionary<int,NetworkInterface> adapterList= new Dictionary<int, NetworkInterface>();
        NetworkInterface selectedAdapter;
        List<PadInfo> PadInfo = new List<PadInfo>();
        List<String> Companys;
        List<String> Pads;
        List<String> Devices;

        private ModbusTCP.Master MBmaster;


        public MainForm()
        {
            InitializeComponent();
            if (Properties.Settings.Default.RecentFiles == null)
            {
                Properties.Settings.Default.RecentFiles = new System.Collections.Specialized.StringCollection();
            }
            else
            {
                foreach (string filename in Properties.Settings.Default.RecentFiles)
                {
                    comboBox_FileLocation.Items.Add(filename);
                }
            }

            label_Version.Text = "V " + Application.ProductVersion;
            //textBox_StartCopyDelayDirect.Text = Properties.Settings.Default.DelayTime.ToString();
            textBox_StartCopyDelayFile.Text = Properties.Settings.Default.DelayTime.ToString();
            comboBox_TargetProgramCSV.SelectedIndex = Properties.Settings.Default.TargetProgram;

            //selectedAdapter = NetworkInterface.GetAllNetworkInterfaces().Where(n => n.NetworkInterfaceType != NetworkInterfaceType.Loopback).First(n => n.OperationalStatus == OperationalStatus.Up);
            selectedAdapter = NetworkInterface.GetAllNetworkInterfaces().FirstOrDefault();
            LoadAdapters();

            SetPadDB();

            //Reports
            comboBox_ReportType.SelectedIndex = 0;
            comboBox_HexaneCalc.SelectedIndex = 0;
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
                    //btn_StartCopyDirect.Enabled = false;
                    btn_StartCopyFile.Enabled = false;
                   // textBox_StartCopyDelayDirect.Enabled = false;
                    textBox_StartCopyDelayFile.Enabled = false;
                    break;
                case ButtonState.READY:
                    btn_Cancel1.Enabled = false;
                    btn_SelectFile.Enabled = true;
                   // btn_StartCopyDirect.Enabled = true;
                    btn_StartCopyFile.Enabled = true;
                    //textBox_StartCopyDelayDirect.Enabled = true;
                    textBox_StartCopyDelayFile.Enabled = true;
                    break;
                default:
                    break;
            }
        }

        private void SetFileMostRecent(string file)
        {
            Properties.Settings.Default.RecentFiles.Remove(file);
            Properties.Settings.Default.RecentFiles.Insert(0, file);
            Properties.Settings.Default.Save();
            foreach (string filename in Properties.Settings.Default.RecentFiles)
            {
                comboBox_FileLocation.Items.Add(filename);
            }

        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //Set Default Drop Downs
            comboBox_ReqMBFunc.SelectedIndex = 2;
            comboBox_ReqDataType.SelectedIndex = 0;
            comboBox_ReqFormat.SelectedIndex = 1;

            //Set Grid View to default
            int number = 5;
            int startValue = 7000;
            if (number > 0)
            {
                for (int count = 0; count < number; count++)
                {
                    string[] rowValues = new string[] {(startValue + count).ToString(), ""};
                    dataGridView_ReqData.Rows.Add(rowValues );
                }
            }
            
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
                if (!string.IsNullOrWhiteSpace(result))
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
        private enum TargetProgram
        {
            TxT = 0,
            Excel = 1,
            PCCU = 2,
            Realflo = 3,
            NewAGA = 4,
            OldAGA = 5,
            NewModWorx = 6,
            OldModWorx = 7
        }

        private void btn_StartCopyFile_Click(object sender, EventArgs e)
        {
            List<string> BGWorkStorage = new List<string>();
            string sourceLoc = comboBox_FileLocation.Text;
            if (sourceLoc.Count() > 0)
            {
                label_Status.Text = "Loading File...";
                EnableButtons(ButtonState.COPYING);

                if (!BgWorker.IsBusy)
                {
                    BGWorkStorage.Add(sourceLoc);
                    BGWorkStorage.Add(textBox_KeypressDelay.Text);
                    BGWorkStorage.Add(textBox_KeyStateChange.Text);
                    BgWorker.RunWorkerAsync(BGWorkStorage);
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
            List<string> BGWorkerList = (List<string>)e.Argument;
            string sourceLoc = BGWorkerList[0]; //set file source
            int KeypressDelay;
            int KeypressStateDelay;
            //parse out delay values
            if (!Int32.TryParse(BGWorkerList[1], out KeypressDelay)) MessageBox.Show("KeypressDelay not valid. Must be an integer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (!Int32.TryParse(BGWorkerList[2], out KeypressStateDelay)) MessageBox.Show("KeypressStatDelay not valid. Musst be an integer.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

            if (File.Exists(sourceLoc))
            {
                bool remove = false; //used to determine whether to delete csv after copying (realflow and xmv)
                TargetProgram target = ((TargetProgram)Properties.Settings.Default.TargetProgram);
                //Generate Realflow CSV from run report
                if(target == TargetProgram.Realflo)
                {
                    int dirIndex = sourceLoc.LastIndexOf("\\") + 1;
                    string outputLoc = sourceLoc.Remove(dirIndex);
                    ReportGenerator rG = new ReportGenerator();
                    sourceLoc = rG.GenerateRealfloCSV(sourceLoc, outputLoc); //set source to new csv
                    if (sourceLoc == null)
                    {
                        MessageBox.Show("Failed to generate CSV.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        e.Cancel = true;
                    }
                    else remove = true;
                }
                //Generate XMV CSV from run report
                else if (target == TargetProgram.NewAGA)
                {
                    int dirIndex = sourceLoc.LastIndexOf("\\") + 1;
                    string outputLoc = sourceLoc.Remove(dirIndex);
                    ReportGenerator rG = new ReportGenerator();
                    sourceLoc = rG.GenerateNewAGA3CSV(sourceLoc, outputLoc); //set source to new csv
                    if (sourceLoc == null)
                    {
                        MessageBox.Show("Failed to generate CSV.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        e.Cancel = true;
                    }
                    else remove = true;
                }
                else if (target == TargetProgram.OldAGA)
                {
                    int dirIndex = sourceLoc.LastIndexOf("\\") + 1;
                    string outputLoc = sourceLoc.Remove(dirIndex);
                    ReportGenerator rG = new ReportGenerator();
                    sourceLoc = rG.GenerateOldAGA3CSV(sourceLoc, outputLoc); //set source to new csv
                    if (sourceLoc == null)
                    {
                        MessageBox.Show("Failed to generate CSV.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        e.Cancel = true;
                    }
                    else remove = true;
                }
                //Generate ModWorx CSV from run report
                else if (target == TargetProgram.NewModWorx || target == TargetProgram.OldModWorx)
                {
                    int dirIndex = sourceLoc.LastIndexOf("\\") + 1;
                    string outputLoc = sourceLoc.Remove(dirIndex);
                    ReportGenerator rG = new ReportGenerator();
                    sourceLoc = rG.GenerateModWorxCSV(sourceLoc, outputLoc); //set source to new CSV
                    if (sourceLoc == null)
                    {
                        MessageBox.Show("Failed to generate CSV.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        e.Cancel = true;
                    }
                    else remove = true;
                }

                //convert excel to csv wasn't working. Skipping for now

                if (sourceLoc.Split('.')[sourceLoc.Split('.').Length - 1].ToLower() == "csv")//check if csv file
                {
                    CSVReader reader = new CSVReader();
                    reader.ParseCSV(sourceLoc, "");
                    Typer typer = new Typer();
                    typer.strokeDelay = KeypressDelay;
                    typer.ih.kscdelay = KeypressStateDelay;

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
                        if (target == TargetProgram.TxT)
                        {
                            typer.TypeCSVtoText(reader.GetArrayStorage(), bg);
                        }
                        else if (target == TargetProgram.Excel)
                        {
                            typer.TypeCSVtoExcel(reader.GetArrayStorage(), bg);
                        }
                        else if (target == TargetProgram.PCCU)
                        {
                            typer.TypeCSVtoPCCU(reader.GetArrayStorage(), bg);
                        }
                        else if (target == TargetProgram.NewAGA || target == TargetProgram.OldAGA)
                        {
                            typer.TypeCSVtoAGA(reader.GetArrayStorage(), bg);
                        }
                        else if (target == TargetProgram.Realflo)
                        {
                            typer.TypeCSVtoRealflo(reader.GetArrayStorage(), bg);
                        }
                        else if (target == TargetProgram.NewModWorx)
                        {
                            typer.TyperCSVtoNewModWorx(reader.GetArrayStorage(), bg);
                        }
                        else if (target == TargetProgram.OldModWorx)
                        {
                            typer.TypeCSVtoOldModWorx(reader.GetArrayStorage(), bg);
                        }
                        if (bg.CancellationPending)
                        {
                            e.Cancel = true;
                        }
                    }
                    //delete temporary csv
                    if (remove) File.Delete(sourceLoc);
                }
                else
                {
                    MessageBox.Show("CSV file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    e.Cancel = true;
                }

            }
            else
            {
                MessageBox.Show("Could not find file", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
                //return;
            }
        }


        private void BgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage > 0)
            {
                label_Status.Text = "Press Any Key at least Once. Starting in " + e.ProgressPercentage;
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
                    Properties.Settings.Default.Save();
                }

                //textBox_StartCopyDelayDirect.Text = Properties.Settings.Default.DelayTime.ToString();
                textBox_StartCopyDelayFile.Text = Properties.Settings.Default.DelayTime.ToString();
            }


        }

        private void comboBox_TargetProgramCSV_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.TargetProgram = comboBox_TargetProgramCSV.SelectedIndex;
            Properties.Settings.Default.Save();
            int index = comboBox_TargetProgramCSV.SelectedIndex;
            switch (index)
            {
                case (int)TargetProgram.TxT:
                    openFileDialog1.Filter = "excel files | *.csv; *.xlsx";
                    pictureBox2.Image = ExcelPaster.Properties.Resources.TXT;
                    break;
                case (int)TargetProgram.Excel:
                    openFileDialog1.Filter = "excel files | *.csv; *.xlsx";
                    pictureBox2.Image = ExcelPaster.Properties.Resources.No_Photo;
                    break;
                case (int)TargetProgram.PCCU:
                    openFileDialog1.Filter = "excel files | *.csv; *.xlsx";
                    pictureBox2.Image = ExcelPaster.Properties.Resources.ExcelToPCCU;
                    break;
                case (int)TargetProgram.Realflo:
                    openFileDialog1.Filter = "Run Reports | *.3.txt";
                    pictureBox2.Image = ExcelPaster.Properties.Resources.Realflo;
                    break;
                case (int)TargetProgram.NewAGA:
                    openFileDialog1.Filter = "Run Reports | *.3.txt";
                    pictureBox2.Image = ExcelPaster.Properties.Resources.NewAGA;
                    break;
                case (int)TargetProgram.OldAGA:
                    openFileDialog1.Filter = "Run Reports | *.3.txt";
                    pictureBox2.Image = ExcelPaster.Properties.Resources.No_Photo;
                    break;
                case (int)TargetProgram.NewModWorx:
                    openFileDialog1.Filter = "Run Reports | *.3.txt";
                    pictureBox2.Image = ExcelPaster.Properties.Resources.No_Photo;
                    break;
                case (int)TargetProgram.OldModWorx:
                    openFileDialog1.Filter = "Run Reports | *.3.txt";
                    pictureBox2.Image = ExcelPaster.Properties.Resources.No_Photo;
                    break;
                default:
                    break;
            }
        }

        private void textBox_IPAdress_TextChanged(object sender, EventArgs e)
        {
            //check to see if input iks valid/invalid or matching
            bool isIP = false;
            string text = textBox_IPAdress.Text;
            IPAddress address;
            if (IPAddress.TryParse(text, out address))
            {
                switch (address.AddressFamily)
                {
                    case System.Net.Sockets.AddressFamily.InterNetwork:
                        // we have IPv4
                        isIP = true;
                        break;
                    case System.Net.Sockets.AddressFamily.InterNetworkV6:
                        // we have IPv6
                        isIP = false;
                        break;
                    default:
                        // umm... yeah... I'm going to need to take your red packet and...
                        isIP = false;
                        break;
                }
            }

            if (isIP)
            {
                string curIP = addressIP.ToString();
                if (curIP == address.ToString())
                {
                    //Make green
                    textBox_IPAdress.BackColor = Color.LightGreen;
                }
                else
                {
                    //Make yellow 
                    textBox_IPAdress.BackColor = Color.LightYellow;
                }
            }
            else
            {
                //make red
                textBox_IPAdress.BackColor = Color.LightPink;
            }
        }
        public static IPAddress GetLocalIPAddress(NetworkInterface adapter)
        {
            int trys = 0;
            while (trys < 10)
            {
                foreach (UnicastIPAddressInformation unicastIPAddressInformation in adapter.GetIPProperties().UnicastAddresses)
                {
                    if (unicastIPAddressInformation.Address.AddressFamily == AddressFamily.InterNetwork)
                    {

                        return unicastIPAddressInformation.Address;

                    }
                }
                trys++;
                System.Threading.Thread.Sleep(1000);

            }
            
            throw new ArgumentException(string.Format("Can't find address"));
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            //check to see if input iks valid/invalid or matching
            bool isIP = false;
            string text = textBox2.Text;
            IPAddress address;
            if (IPAddress.TryParse(text, out address))
            {
                switch (address.AddressFamily)
                {
                    case System.Net.Sockets.AddressFamily.InterNetwork:
                        // we have IPv4
                        isIP = true;
                        break;
                    case System.Net.Sockets.AddressFamily.InterNetworkV6:
                        // we have IPv6
                        isIP = false;
                        break;
                    default:
                        // umm... yeah... I'm going to need to take your red packet and...
                        isIP = false;
                        break;
                }
            }

            if (isIP)
            {
                string curIP = submask.ToString();
                if (curIP == address.ToString())
                {
                    //Make green
                    textBox2.BackColor = Color.LightGreen;
                }
                else
                {
                    //Make yellow 
                    textBox2.BackColor = Color.LightYellow;
                }
            }
            else
            {
                //make red
                textBox2.BackColor = Color.LightPink;
            }
        }
        public static IPAddress GetSubnetMask(IPAddress address,NetworkInterface adapter)
        {
           
                foreach (UnicastIPAddressInformation unicastIPAddressInformation in adapter.GetIPProperties().UnicastAddresses)
                {
                    if (unicastIPAddressInformation.Address.AddressFamily == AddressFamily.InterNetwork)
                    {
                        if (address.Equals(unicastIPAddressInformation.Address))
                        {
                            return unicastIPAddressInformation.IPv4Mask;
                        }
                    }
                }
            
            throw new ArgumentException(string.Format("Can't find subnetmask for IP address '{0}'", address));
        }
        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            //check to see if input iks valid/invalid or matching
            bool isIP = false;
            string text = textBox3.Text;
            IPAddress address;
            if (IPAddress.TryParse(text, out address))
            {
                switch (address.AddressFamily)
                {
                    case System.Net.Sockets.AddressFamily.InterNetwork:
                        // we have IPv4
                        isIP = true;
                        break;
                    case System.Net.Sockets.AddressFamily.InterNetworkV6:
                        // we have IPv6
                        isIP = false;
                        break;
                    default:
                        // umm... yeah... I'm going to need to take your red packet and...
                        isIP = false;
                        break;
                }
            }

            if (isIP)
            {
                string curIP =gateway.ToString();
                if (curIP == address.ToString())
                {
                    //Make green
                    textBox3.BackColor = Color.LightGreen;
                }
                else
                {
                    //Make yellow 
                    textBox3.BackColor = Color.LightYellow;
                }
            }
            else
            {
                //make red
                textBox3.BackColor = Color.LightPink;
            }
        }
        public static IPAddress GetDefaultGateway(NetworkInterface adapter)
        {
            if (adapter.GetIPProperties().GatewayAddresses.Count() > 0)
            {
                string what = adapter.GetIPProperties().GatewayAddresses[0].Address.ToString();

                IPAddress addre = IPAddress.Parse(what);

                return addre;

            }
            else
            {
                IPAddress blankaddress = IPAddress.Parse("0.0.0.0");
                return blankaddress;
            }
            
        //    return NetworkInterface
        //.GetAllNetworkInterfaces()
        //.Where(n => n.OperationalStatus == OperationalStatus.Up)
        //.Where(n => n.NetworkInterfaceType != NetworkInterfaceType.Loopback)
        //.SelectMany(n => n.GetIPProperties()?.GatewayAddresses)
        //.Select(g => g?.Address)
        //.Where(a => a != null)
        //// .Where(a => a.AddressFamily == AddressFamily.InterNetwork)
        //// .Where(a => Array.FindIndex(a.GetAddressBytes(), b => b != 0) >= 0)
        //.FirstOrDefault();
        }

        private void comboBox_NetworkAdapter_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedAdapter = adapterList[comboBox_NetworkAdapter.SelectedIndex];

            addressIP = GetLocalIPAddress(selectedAdapter);
            IPAdress_Status.Text = addressIP.ToString();
            textBox_IPAdress.Text = addressIP.ToString();

            submask = GetSubnetMask(addressIP, selectedAdapter);
            SubMask_Status.Text = submask.ToString();
            textBox2.Text = submask.ToString();

            gateway = GetDefaultGateway(selectedAdapter);
            DefGate_Status.Text = gateway.ToString();
            textBox3.Text = gateway.ToString();           
            
            //label36.Text = selectedAdapter.Description.ToString();
        }

        private void LoadAdapters()
        {
           //TODO: If IP is dynamic dont load the IP into text box
            int indexer = 0;
            foreach (NetworkInterface adapter in NetworkInterface.GetAllNetworkInterfaces().Where(n => n.NetworkInterfaceType != NetworkInterfaceType.Loopback))
            {
                //this will match the order of known adapters with the combo box in case they change
                if (!adapter.Name.Contains("Local Area Connection"))
                {
                    adapterList.Add(indexer, adapter);
                    comboBox_NetworkAdapter.Items.Insert(indexer, adapter.Name);
                    indexer++;
                }
            }
            
            comboBox_NetworkAdapter.SelectedItem = selectedAdapter.Name;
            
          
            addressIP = GetLocalIPAddress(selectedAdapter);
            IPAdress_Status.Text = addressIP.ToString();
            textBox_IPAdress.Text = addressIP.ToString();

            submask = GetSubnetMask(addressIP, selectedAdapter);
            SubMask_Status.Text = submask.ToString();
            textBox2.Text = submask.ToString();

            gateway = GetDefaultGateway(selectedAdapter);
            DefGate_Status.Text = gateway.ToString();
            textBox3.Text = gateway.ToString();


            string startAddress = addressIP.ToString().Remove(addressIP.ToString().LastIndexOf(".") + 1) +"0";
            string endAddress = addressIP.ToString().Remove(addressIP.ToString().LastIndexOf(".") + 1) + "254";
            textBox_IPScanStart.Text = startAddress;
            textBox_IPScanStop.Text = endAddress;


        }
        private void button_RefreshAdapter_Click(object sender, EventArgs e)
        {
            adapterList.Clear();
            comboBox_NetworkAdapter.Items.Clear();
            LoadAdapters();
        }

        private void button_ApplyIPChanges_Click(object sender, EventArgs e)
        {

            string newAddressString = "";
            string newSubMaskString = "";
            string newGatewayString = "";

            IPAddress newAddress;
            if (IPAddress.TryParse(textBox_IPAdress.Text, out newAddress))
            {
                newAddressString = newAddress.ToString();
                //Check that IP does not match another adapter
                foreach (var netInt in adapterList)
                {
                    if (selectedAdapter == netInt.Value)
                    {
                        continue;
                    }
                    IPAddress ipTest = GetLocalIPAddress(netInt.Value);
                    if (newAddress.ToString() == ipTest.ToString())
                    {
                        var confirmResult = MessageBox.Show("IP Address is the same as " + netInt.Value.Name,
                                     "Try a different IP",
                                     MessageBoxButtons.OK);
                        return;
                    }
                }
                
            }
            else
            {
                newAddressString = "0.0.0.0";
            }
            
            IPAddress newSubMask;
            if (IPAddress.TryParse(textBox2.Text, out newSubMask))
            {
                newSubMaskString = newSubMask.ToString();
            }
            else
            {
                newSubMaskString = "0.0.0.0";
            }

            IPAddress newGateway;
            if (IPAddress.TryParse(textBox3.Text, out newGateway))
            {
                newGatewayString = newGateway.ToString();
            }
            else
            {
                newGatewayString = "0.0.0.0";
            }

            Process p = new Process();
            ProcessStartInfo psi = new ProcessStartInfo("netsh", "interface ip set address \""+selectedAdapter.Name +"\" static "+newAddressString+" "+newSubMaskString +" "+newGatewayString+" 1");
            p.StartInfo = psi;
            p.StartInfo.Verb = "runas"; 
            p.Start();

            p.WaitForExit();
            adapterList.Clear();
            comboBox_NetworkAdapter.Items.Clear();
            LoadAdapters();
            textBox_IPAdress.BackColor = Color.LightGreen;
            textBox2.BackColor = Color.LightGreen;
            textBox3.BackColor = Color.LightGreen;
        }

        private void button_ChangeDBFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = openFileDialog2.FileName;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_DBFile.Text = result;
                    Properties.Settings.Default.DatabaseFileLoc = result;
                    Properties.Settings.Default.Save();
                    SetPadDB();
                }
            }
        }
        private void SetPadDB()
        {
            if (Properties.Settings.Default.DatabaseFileLoc != "")
            {
                comboBox_DBFile.Text = Properties.Settings.Default.DatabaseFileLoc;
                //Load Pad DB
                CSVReader padReader = new CSVReader();
                padReader.ParseCSV(Properties.Settings.Default.DatabaseFileLoc,"");
                PadInfo.Clear();
                foreach (List<string> ListS in padReader.GetArrayStorage())
                {
                    if (ListS.Count >= 5)
                    {
                        PadInfo pInfo = new PadInfo(ListS[0], ListS[1], ListS[2], ListS[3], ListS[4], ListS[5]);
                       
                        PadInfo.Add(pInfo);
                    }
                    
                }
                Companys = PadInfo.Select(s => s.Company).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                Pads = PadInfo.Select(s => s.PadName).Distinct().ToList();
                Devices = PadInfo.Select(s => s.DeviceName).Distinct().ToList();

                //comboBox_NewCompany.Items.Clear();
                //comboBox_AddDBCompany.Items.Clear();

                //comboBox_NewPad.Items.Clear();
                //comboBox_AddDBPad.Items.Clear();

                //comboBox_NewDevice.Items.Clear();
                //comboBox_AddDBDevice.Items.Clear();

                comboBox_NewCompany.Items.AddRange(Companys.ToArray());
                comboBox_AddDBCompany.Items.AddRange(Companys.ToArray());

                comboBox_NewPad.Items.AddRange(Pads.ToArray());
                comboBox_AddDBPad.Items.AddRange(Pads.ToArray());

                comboBox_NewDevice.Items.AddRange(Devices.ToArray());
                comboBox_AddDBDevice.Items.AddRange(Devices.ToArray());
            }
            
        }

        private void comboBox_NewCompany_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox_NewPad.Items.Clear();
            comboBox_NewPad.Items.AddRange(PadInfo.Where(c=>c.Company == comboBox_NewCompany.SelectedItem.ToString()).Select(s => s.PadName).Distinct().ToArray());
            comboBox_NewPad.SelectedIndex = 0;
            comboBox_NewDevice.SelectedIndex = 0;
            textBox_DBAddress.Text = "";
            textBox_DBSubMask.Text = "";
            textBox_DBGateway.Text = "";
            comboBox_NewPad.Enabled = true;
        }

        private void comboBox_NewPad_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox_NewDevice.Items.Clear();

            comboBox_NewDevice.Items.AddRange(PadInfo.Where(c => c.Company == comboBox_NewCompany.SelectedItem.ToString())
                .Where(c => c.PadName == comboBox_NewPad.SelectedItem.ToString()).Select(s => s.DeviceName).Distinct().ToArray());
            comboBox_NewDevice.SelectedIndex = 0;
            textBox_DBAddress.Text = "";
            textBox_DBSubMask.Text = "";
            textBox_DBGateway.Text = "";
            comboBox_NewDevice.Enabled = true;

            //Add all IPs to List
            listView_DBAddresses.Items.Clear();
            System.Version number= Version.Parse("0.0.0.0");
            List<PadInfo> padInfoList = PadInfo.Where(c => c.Company == comboBox_NewCompany.SelectedItem.ToString())
                .Where(c => c.PadName == comboBox_NewPad.SelectedItem.ToString()).OrderByDescending(c => Version.TryParse(c.IPAddress,out number)).Reverse().ToList();
            foreach (PadInfo pad in padInfoList)
            {
                ListViewItem lvi = new ListViewItem(pad.IPAddress.ToString());
                //lvi.SubItems.Add();
                lvi.SubItems.Add(pad.DeviceName.ToString());

                listView_DBAddresses.Items.Add(lvi);
            }
            
        }

        private void comboBox_NewDevice_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_NewCompany.SelectedItem.ToString() != "" && comboBox_NewPad.SelectedItem.ToString() != "" && comboBox_NewDevice.SelectedItem.ToString() != "")
            {
                //search for IP
                PadInfo pi = PadInfo.Where(x => x.Company == comboBox_NewCompany.SelectedItem.ToString()).Where(x => x.PadName == comboBox_NewPad.SelectedItem.ToString()).First(x => x.DeviceName == comboBox_NewDevice.SelectedItem.ToString());
                textBox_DBAddress.Text = pi.IPAddress;
                textBox_DBSubMask.Text = pi.SubnetMask;
                textBox_DBGateway.Text = pi.Gateway;

            }
            else
            {
                textBox_DBAddress.Text = "";
                textBox_DBSubMask.Text = "";
                textBox_DBGateway.Text = "";
            }
        }
        private void comboBox_AddDBPad_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button_AddIPInfo_Click(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.DatabaseFileLoc != "")
            {
                //Search PadInfo for existing entry
                string addcompany = comboBox_AddDBCompany.Text;
                string addpad = comboBox_AddDBPad.Text;
                string adddevice = comboBox_AddDBDevice.Text;

                PadInfo possibleDevice = PadInfo.Where(c => c.Company == addcompany).Where(p => p.PadName == addpad).FirstOrDefault(d => d.DeviceName == adddevice);
                if (possibleDevice == null)
                {
                    PadInfo.Add(new PadInfo(addcompany,addpad,adddevice,textBox_AddDBAddress.Text,textBox_AddDB_SubMask.Text,textBox_AddDBGateway.Text));

                    if (File.Exists(Properties.Settings.Default.DatabaseFileLoc))
                    {

                        string appendInfo = addcompany +","+ addpad + "," + adddevice + "," + textBox_AddDBAddress.Text + "," + textBox_AddDB_SubMask.Text + "," + textBox_AddDBGateway.Text;
                        
                        File.AppendAllText(Properties.Settings.Default.DatabaseFileLoc, Environment.NewLine + appendInfo);
                        MessageBox.Show(addcompany + "," + addpad + "," + adddevice + "," + textBox_AddDBAddress.Text + "," + textBox_AddDB_SubMask.Text + "," + textBox_AddDBGateway.Text,"Added Entry" );
                        //PadInfo = new List<PadInfo>();
                        //Companys = new List<string>();
                        //Pads = new List<string>();
                        //Devices = new List<string>();
                        SetPadDB();
                    }

                   
                }
            }
        }

        private void textBox_KeypressDelay_TextChanged(object sender, EventArgs e)
        {

        }

        private void button_OpenFile_Click(object sender, EventArgs e)
        {
            if ( !string.IsNullOrEmpty(Properties.Settings.Default.DatabaseFileLoc))
            {
                Process.Start(Properties.Settings.Default.DatabaseFileLoc);
            }
            
        }

        //private BackgroundWorker pingWorker = new BackgroundWorker();
        public class PingObject
        {

            public PingObject(IPAddress ip,int trys)
            {
                this.IP = ip;
                this.Trys = trys;

            }
            public IPAddress IP;
            public int Trys;
            

        }

        private void button_Ping_Click(object sender, EventArgs e)
        {
            //TODO: Change to Ping.exe and run pings into a list with graphics
            IPAddress newAddress;
            IPAddress.TryParse(textBox_DBAddress.Text, out newAddress);
            int pingTrys = Int32.Parse(textBox1_PingTrys.Text.ToString());
            PingObject po = new PingObject(newAddress,pingTrys);
            //-----------------------------------------------------------------

          
            if (newAddress != null)
            {

                try
                {
                    pingWorker.RunWorkerAsync(argument: po);
                    label_PingResults.Text += "\nPinging " + textBox_DBAddress.Text + "....";
                }
                catch
                {
                }
                

            }
            else
            {
                label_PingResults.Text = "'"+ textBox_DBAddress.Text + "' was not a valid IP Address";
            }
            //-------------------------------------------------------------------------
        }
       
        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void pingWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            PingObject po = (PingObject)e.Argument;
            IPAddress newAddress = po.IP;
            int pingCount = 0;
            int pingTrys = po.Trys;
            int successCount = 0;

            /*
            Process p = new Process();
            // No need to use the CMD processor - just call ping directly.
            p.StartInfo.FileName = "ping.exe";
            p.StartInfo.Arguments = "-a " + newAddress.ToString();
            p.StartInfo.RedirectStandardOutput = true;
            p.StartInfo.UseShellExecute = false;
            p.StartInfo.CreateNoWindow = true;
            p.Start();
            p.WaitForExit();

            var output = p.StandardOutput.ReadToEnd();
            pingWorker.ReportProgress(100,output);
            */
           
            Ping ping = new Ping();
            PingReply pingReply = null ;
            try
            {
                pingReply = ping.Send(newAddress.ToString());
                // check when the ping is not success
                while (pingTrys > 0)
                {
                    while (!(pingReply.Status.ToString() == "Success") & pingTrys > 0)
                    {
                        pingTrys--;
                        pingCount++;
                        //Console.WriteLine(pingReply.Status.ToString());
                        // check after the ping is n success


                        var output = "\n    " + pingReply.Status.ToString();
                        pingWorker.ReportProgress(100, output);
                        pingReply = ping.Send(newAddress.ToString());


                        if (pingWorker.CancellationPending)
                        {
                            break;
                        }
                    }

                    if (pingReply.Status.ToString().Equals("Success"))
                    {
                        pingTrys--;
                        pingCount++;
                        // Console.WriteLine(pingReply.Status.ToString());
                        var output = "\n    " + pingReply.Status.ToString() + " " + pingReply.RoundtripTime.ToString() + "ms";
                        pingWorker.ReportProgress(100, output);
                        pingReply = ping.Send(newAddress.ToString());
                        successCount++;
                    }

                    if (pingWorker.CancellationPending)
                    {
                        pingTrys = -1;
                        var output = "\n    Ping Request Canceled";
                        pingWorker.ReportProgress(100, output);

                    }
                    
                }
                if (pingTrys <= 0)
                {
                    var output = "\nPinged " + pingCount + " Times. \nSuccessRate of " + Math.Round((float)successCount / (float)pingCount, 2)*100f + "%";
                    pingWorker.ReportProgress(100, output);
                }

            }
            catch
            {
                var output = "Not a valid IP Address";
                pingWorker.ReportProgress(100, output);
            }
           

        
        }


        private void pingWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //label_PingResults.Text = "";
            label_PingResults.Text += (string)e.UserState;// output;
            panel1.VerticalScroll.Value = panel1.VerticalScroll.Maximum;
        }

        private void button_SetDynamic_Click(object sender, EventArgs e)
        {
            Process p = new Process();
            ProcessStartInfo psi = new ProcessStartInfo("netsh", "interface ip set address \"" + selectedAdapter.Name + "\" dhcp");
            p.StartInfo = psi;
            p.StartInfo.Verb = "runas";
            p.Start();

            p.WaitForExit();
            System.Threading.Thread.Sleep(2000);
            
            //adapterList.Clear();
            //comboBox_NetworkAdapter.Items.Clear();
            //LoadAdapters();
            //textBox_IPAdress.BackColor = Color.LightGreen;
            //textBox2.BackColor = Color.LightGreen;
            //textBox3.BackColor = Color.LightGreen; 
        }

       


       

        private void button_OpenTODOFile_Click_1(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.TODOFileLoc.Count() > 0) Process.Start(Properties.Settings.Default.TODOFileLoc);
            else MessageBox.Show("No file Selected.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }


        private void button_CancelPing_Click(object sender, EventArgs e)
        {
            pingWorker.CancelAsync();
        }

        private void button_ClearPings_Click(object sender, EventArgs e)
        {
            label_PingResults.Text = "";
        }

        private void button_AdapterOptionsControlPanel_Click(object sender, EventArgs e)
        {
            //var cplPath = System.IO.Path.Combine(Environment.SystemDirectory, "control.exe");
            //System.Diagnostics.Process.Start(cplPath, "/name Microsoft.NetworkConnections");
            System.Diagnostics.Process.Start("NCPA.cpl");
        }

        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }
        public static uint ParseIP(string ip)
        {
            byte[] b = ip.Split('.').Select(s => Byte.Parse(s)).ToArray();
            if (BitConverter.IsLittleEndian) Array.Reverse(b);
            return BitConverter.ToUInt32(b, 0);
        }

        public static string FormatIP(uint ip)
        {
            byte[] b = BitConverter.GetBytes(ip);
            if (BitConverter.IsLittleEndian) Array.Reverse(b);
            return String.Join(".", b.Select(n => n.ToString()));
        }


        int threadCount = 0;
        public void PingThread(int itemID,string itemAddress)
        {
            
            //listView_ScannedPadIPs.Items[itemID].Name = "";

            IPAddress pingingIP = IPAddress.Parse(itemAddress);
            Ping ping = new Ping();
            PingReply pingReply = null;
            try
            {
                pingReply = ping.Send(pingingIP.ToString());
                int count = 5000;
                while (!(pingReply.Status.ToString() == "Success"))
                {
                    System.Threading.Thread.Sleep(500);
                    count = count - 500;
                    if (count <= 0)
                    {
                        break;
                    }
                }
                if (pingReply.Status.ToString() == "Success")
                {
                    listView_ScannedPadIPs.Items[itemID].SubItems[0].Text = "1";
                }
                else
                {
                    listView_ScannedPadIPs.Items[itemID].SubItems[0].Text = "0";
                }
                
            }
            catch
            {
            }
            threadCount--;
        }
        
        private void button_ScanNetwork_Click(object sender, EventArgs e)
        {

            //Create List to Scan
            string StartIP = textBox_IPScanStart.Text.ToString();
            string StopIP = textBox_IPScanStop.Text.ToString();

            uint startInt = ParseIP(StartIP);
            uint stopInt = ParseIP(StopIP);
            uint IPCount = stopInt - startInt;

            string[] range = new string[IPCount];
            for (uint i = 0; i < IPCount; i++)
            {
                string curIP = FormatIP(startInt + i);
                ListViewItem lvi = new ListViewItem("n/s");
                lvi.SubItems.Add(curIP);
                lvi.SubItems.Add("n/s");
                listView_ScannedPadIPs.Items.Add(lvi);
                range[i] = curIP;
            }

            //Scan over List
          
            int threadMaxCount = 100;
            foreach (ListViewItem item in listView_ScannedPadIPs.Items)
            {
                while (threadCount > threadMaxCount)
                {
                    System.Threading.Thread.Sleep(500);
                }
                
                //Create new thread
                threadCount++;
                int itemId = item.Index;
                string itemAddress = item.SubItems[1].Text.ToString();
                Thread t = new Thread(() => PingThread(itemId,itemAddress));
                t.Start();
               
               
            }
        }

        private void textBox_IPScanStart_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
        private byte[] MBRequest = {0x01, 0x01, 0x1b, 0x58, 0x00 ,0x05 };
        private byte[] MBData;
        private void btnConnect_Click(object sender, EventArgs e)
        {
            try
            {
                // Create new modbus master and add event functions
                MBmaster = new Master(textbox_ModbusIP.Text,ushort.Parse( textBox_Port.Text), true);
                MBmaster.OnResponseData += new ModbusTCP.Master.ResponseData(MBmaster_OnResponseData);
                MBmaster.OnException += new ModbusTCP.Master.ExceptionData(MBmaster_OnException);
                MBmaster.OnSocketData += new ModbusTCP.Master.SocketData(MBmaster_OnSocketData);

            }
            catch (SystemException error)
            {
                MessageBox.Show(error.Message);
            }
        }
        // ------------------------------------------------------------------------
        // Event for socket data
        // ------------------------------------------------------------------------
        private void MBmaster_OnSocketData(byte[] values,bool IsRecieve)
        {
            // ------------------------------------------------------------------
            // Seperate calling threads
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Master.SocketData(MBmaster_OnSocketData), new object[] {  values,IsRecieve });
                return;
            }
            if (IsRecieve)
            {
                textBox_ReqData.Text += ">>> " + BitConverter.ToString(values) + System.Environment.NewLine;
            }
            else
            {
                textBox_ReqData.Text += "<<< " + BitConverter.ToString(values) + System.Environment.NewLine;
            }
            
        }
        // ------------------------------------------------------------------------
        // Event for response data
        // ------------------------------------------------------------------------
        private void MBmaster_OnResponseData(ushort ID, byte unit, byte function, byte[] values)
        {
            // ------------------------------------------------------------------
            // Seperate calling threads
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Master.ResponseData(MBmaster_OnResponseData), new object[] { ID, unit, function, values });
                return;
            }

            // ------------------------------------------------------------------------
            // Identify requested data
            switch (ID)
            {
                case 1:
                    group_Data.Text = "Read coils";
                    MBData = values;
                    //Change MB Data ListView
                    ReadValuesDataGrid();
                    break;
                case 2:
                    group_Data.Text = "Read discrete inputs";
                    MBData = values;
                    //Change MB Data ListView
                    ReadValuesDataGrid();
                    break;
                case 3:
                    group_Data.Text = "Read holding register";
                    MBData = values;
                    //Change MB Data ListView
                    ReadValuesDataGrid();
                    break;
                case 4:
                    group_Data.Text = "Read input register";
                    MBData = values;
                    //Change MB Data ListView
                    ReadValuesDataGrid();
                    break;
                case 5:
                    group_Data.Text = "Write single coil";
                    break;
                case 6:
                    group_Data.Text = "Write multiple coils";
                    break;
                case 7:
                    group_Data.Text = "Write single register";
                    break;
                case 8:
                    group_Data.Text = "Write multiple register";
                    break;
            }
            
        }
        private void ReadValuesDataGrid()
        {
            if (MBData != null)
            {
                float divisor = 2;
                switch(comboBox_ReqDataType.SelectedIndex)
                {
                    case 0:
                        divisor = 2;
                        break;
                    case 1:
                        divisor = 4;
                        break;
                    case 2:
                        divisor = 4;
                        break;
                    case 3:
                        divisor = 0.125f;
                        break;
                    default:
                        break;
                }
                int multiplier = 1;
                if (comboBox_ReqFormat.SelectedIndex > 0)// && (comboBox_ReqDataType.SelectedIndex == 1 || comboBox_ReqDataType.SelectedIndex == 2))
                {
                    multiplier = 2;
                }
                if (MBData.Count()*multiplier/divisor >= dataGridView_ReqData.Rows.Count)
                {
                    for (int counter = 0; counter < dataGridView_ReqData.Rows.Count; counter++)
                    {
                        byte[] byteArray;
                        if (comboBox_ReqDataType.SelectedIndex == 0 && comboBox_ReqFormat.SelectedIndex == 0)
                        {
                            //Note only works for 16 bit ints
                            byteArray = new byte[] { MBData[counter * 2], MBData[(counter * 2) + 1] };
                            if (BitConverter.IsLittleEndian)
                                Array.Reverse(byteArray);
                            int dataValue = BitConverter.ToInt16(byteArray, 0);
                            dataGridView_ReqData.Rows[counter].Cells[1].Value = (dataValue).ToString();
                        }
                        else if (comboBox_ReqDataType.SelectedIndex == 1 && comboBox_ReqFormat.SelectedIndex == 0)
                        {

                            //Note only works for 32 bit ints
                            byteArray = new byte[] { MBData[counter * 4], MBData[(counter * 4) + 1], MBData[(counter * 4) + 2], MBData[(counter * 4) + 3] };
                            if (BitConverter.IsLittleEndian)
                                Array.Reverse(byteArray);
                            int dataValue = BitConverter.ToInt32(byteArray, 0);
                            dataGridView_ReqData.Rows[counter].Cells[1].Value = (dataValue).ToString();

                        }
                        else if (comboBox_ReqDataType.SelectedIndex == 2 && comboBox_ReqFormat.SelectedIndex == 0)
                        {

                            //Note only works for floats
                            byteArray = new byte[] { MBData[counter * 4], MBData[(counter * 4) + 1], MBData[(counter * 4) + 2], MBData[(counter * 4) + 3] };
                            if (BitConverter.IsLittleEndian)
                                Array.Reverse(byteArray);
                            float dataValue = BitConverter.ToSingle(byteArray, 0);
                            dataGridView_ReqData.Rows[counter].Cells[1].Value = (dataValue).ToString();


                        }
                        else if (comboBox_ReqDataType.SelectedIndex == 3)
                        {

                            //Note only works for bools
                            byte[] oneByte = new byte[] { MBData[((int)(counter / 8))] } ;
                            if (BitConverter.IsLittleEndian)
                                Array.Reverse(oneByte);


                            bool dataValue = (oneByte[0] & (1 << counter % 8)) == 0 ? false : true;
                            // bool dataValue = BitConverter.ToBoolean( byteArray,counter);//BitConverter.ToInt16(byteArray, 0);
                             dataGridView_ReqData.Rows[counter].Cells[1].Value = (dataValue).ToString();
                        }

                        //16 Bit Formats selected
                        if (comboBox_ReqFormat.SelectedIndex == 1 || comboBox_ReqFormat.SelectedIndex == 2)
                        {
                            int size = Int16.Parse(textBox_ReqSize.Text);
                            if (size % 2 != 0)
                            {
                                textBox_ReqSize.Text = (size - 1).ToString();
                            }
                            //Data
                            
                            //Doing 16 Bit even tho 32 is selected. Will put 32 bit in Data Pair
                            byteArray = new byte[] { MBData[counter * 2], MBData[(counter * 2) + 1] };
                            if (BitConverter.IsLittleEndian)
                                Array.Reverse(byteArray);
                            int dataValue = BitConverter.ToUInt16(byteArray, 0);
                            dataGridView_ReqData.Rows[counter].Cells[1].Value = (dataValue).ToString();
                            
                            //Data Pair
                            if ((counter + 1) % 2 == 0)
                            {
                                if (comboBox_ReqDataType.SelectedIndex == 1)
                                {
                                    //Int 32 Type
                                    Int32 valueInt32 = 0;
                                    if (comboBox_ReqFormat.SelectedIndex == 1)
                                    {
                                        //16 Bit Modicon
                                        //int combined = (highBits << 16) | lowBits;
                                        valueInt32 = (UInt16.Parse(dataGridView_ReqData.Rows[counter].Cells[1].Value.ToString()) << 16) | UInt16.Parse(dataGridView_ReqData.Rows[counter-1].Cells[1].Value.ToString());


                                    }
                                    else if (comboBox_ReqFormat.SelectedIndex == 2)
                                    {
                                        //16 Bit Word Swapped
                                        valueInt32 = (UInt16.Parse(dataGridView_ReqData.Rows[counter-1].Cells[1].Value.ToString()) << 16) | UInt16.Parse(dataGridView_ReqData.Rows[counter].Cells[1].Value.ToString());
                                    }


                                    dataGridView_ReqData.Rows[counter].Cells[2].Value = valueInt32;
                                }
                                else if (comboBox_ReqDataType.SelectedIndex == 2)
                                {
                                    //Float Type
                                    float valueFloat = 0;
                                    if (comboBox_ReqFormat.SelectedIndex == 1)
                                    {
                                        //16 Bit Modicon
                                        //int combined = (highBits << 16) | lowBits;
                                        valueFloat = BitConverter.ToSingle( BitConverter.GetBytes(( UInt16.Parse(dataGridView_ReqData.Rows[counter].Cells[1].Value.ToString()) << 16) 
                                            | UInt16.Parse(dataGridView_ReqData.Rows[counter - 1].Cells[1].Value.ToString())),0);


                                    }
                                    else if (comboBox_ReqFormat.SelectedIndex == 2)
                                    {
                                        //16 Bit Word Swapped
                                        valueFloat = BitConverter.ToSingle(BitConverter.GetBytes((UInt16.Parse(dataGridView_ReqData.Rows[counter - 1].Cells[1].Value.ToString()) << 16) 
                                            | UInt16.Parse(dataGridView_ReqData.Rows[counter - 1].Cells[1].Value.ToString())),0);
                                    }


                                    dataGridView_ReqData.Rows[counter].Cells[2].Value = valueFloat;
                                }

                            }
                        }

                    }
                }
                else
                {
                    MessageBox.Show("Too few bytes to fit target data type","Data Size Error");
                }
                
            }
        }
        // ------------------------------------------------------------------------
        // Modbus TCP slave exception
        // ------------------------------------------------------------------------
        private void MBmaster_OnException(ushort id, byte unit, byte function, byte exception)
        {
            string exc = "Modbus says error: ";
            switch (exception)
            {
                case Master.excIllegalFunction: exc += "Illegal function!"; break;
                case Master.excIllegalDataAdr: exc += "Illegal data address!"; break;
                case Master.excIllegalDataVal: exc += "Illegal data value!"; break;
                case Master.excSlaveDeviceFailure: exc += "Slave device failure!"; break;
                case Master.excAck: exc += "Acknowledge!"; break;
                case Master.excGatePathUnavailable: exc += "Gateway path unavailabale!"; break;
                case Master.excExceptionTimeout: exc += "Slave timed out!"; break;
                case Master.excExceptionConnectionLost: exc += "Connection is lost!"; break;
                case Master.excExceptionNotConnected: exc += "Not connected!"; break;
            }

            MessageBox.Show(exc, "Modbus slave exception");
        }

        private void UpdateMBReq(byte[] value, int code)
        {
            switch (code)
            {
                case 1:
                    //ID
                    MBRequest[0] = value[0];
                    break;
                case 2:
                    //Func
                    MBRequest[1] = value[0];
                    break;
                case 3:
                    //Start address
                    MBRequest[2] = value[0];
                    MBRequest[3] = value[1];
                    break;
                case 4:
                    //Size
                    MBRequest[4] = value[0];
                    MBRequest[5] = value[1];
                    break;
                   

            }
            textBox_MBRequestStructure.Text = BitConverter.ToString(MBRequest);
        }
        private void textBox_ReqID_TextChanged(object sender, EventArgs e)
        {
            try
            {
                Byte[] b = { Byte.Parse(textBox_ReqID.Text) };
                UpdateMBReq(b, 1);
                textBox_ReqID.BackColor = Color.White;
            }
            catch
            {
                //Not valid
                textBox_ReqID.BackColor = Color.Red;
            }
        }

        private void comboBox_ReqMBFunc_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Byte[] b = { 0x00 };
                switch (comboBox_ReqMBFunc.SelectedIndex)
                {
                    case 0:
                        b[0] = 0x01;
                        UpdateMBReq(b, 2);
                        break;
                    case 1:
                        b[0] = 0x02 ;
                        UpdateMBReq(b, 2);
                        break;
                    case 2:
                        b[0] = 0x03;
                        UpdateMBReq(b, 2);
                        break;
                    case 3:
                        b[0] = 0x04;
                        UpdateMBReq(b, 2);
                        break;
                    case 4:
                        b[0] = 0x05;
                        UpdateMBReq(b, 2);
                        textBox_ReqSize.Text = "1";
                        break;
                    case 5:
                        b[0] = 0x06;
                        UpdateMBReq(b, 2);
                        textBox_ReqSize.Text = "1";
                        break;
                    case 6:
                        b[0] = 0x0F;
                        UpdateMBReq(b, 2);
                        break;
                    case 7:
                        b[0] = 0x10;
                        UpdateMBReq(b, 2);
                        break;



                }
                textBox_ReqID.BackColor = Color.White;
            }
            catch
            {
                comboBox_ReqMBFunc.BackColor = Color.Red;
            }
        }

        private void textBox_ReqStrAddress_TextChanged(object sender, EventArgs e)
        {
            try
            {
               //Change MB Request Data
                int number = Int16.Parse(textBox_ReqStrAddress.Text);

                byte[] splitb = BitConverter.GetBytes(number);
                byte[] b = { splitb[1], splitb[0] };
                UpdateMBReq(b, 3);

                //Change MB Data ListView
                for(int counter = 0; counter < Int16.Parse(textBox_ReqSize.Text);counter++)
                {
                    //listView_ReqData.Items[counter].Text = (number + counter).ToString();
                    dataGridView_ReqData.Rows[counter].Cells[0].Value = (number + counter).ToString();
                }
                

                //Show if Valid
                textBox_ReqStrAddress.BackColor = Color.White;


            }
            catch
            {
                //Not valid
                textBox_ReqStrAddress.BackColor = Color.Red;
            }
        }

        private void textBox_ReqSize_TextChanged(object sender, EventArgs e)
        {
            try
            {
                //Change MB Request Data
                int number = Int16.Parse(textBox_ReqSize.Text);

                byte[] splitb = BitConverter.GetBytes(number);
                byte[] b = { splitb[1], splitb[0] };
                UpdateMBReq(b, 4);

                //Change MB Data ListView
               // int lviCount = listView_ReqData.Items.Count;
                int lviCount = dataGridView_ReqData.Rows.Count;
                //int lviLastValue = Int16.Parse(listView_ReqData.Items[lviCount - 1].Text);
                int lviLastValue = Int32.Parse(dataGridView_ReqData.Rows[lviCount - 1].Cells[0].Value.ToString());
                if (number > lviCount)
                {
                    for (int count = 0; count < number - lviCount; count++)
                    {
                        //listView_ReqData.Items.Add(new ListViewItem( (lviLastValue+count +1).ToString() ));
                        string[] rowValues = new string[] { (lviLastValue + count + 1).ToString(), "" };
                        dataGridView_ReqData.Rows.Add(rowValues);
                        //dataGridView_ReqData.Rows.Add((lviLastValue + count + 1));
                    }
                }else 
                if (number < lviCount)
                {
                    
                    for (int count =0 ; count < lviCount - number; count++)
                    {
                        //listView_ReqData.Items.RemoveAt(lviCount - count - 1);
                        dataGridView_ReqData.Rows.RemoveAt(lviCount - count - 1);
                    }
                }

                //Show if Valid
                textBox_ReqSize.BackColor = Color.White;
            }
            catch(Exception exc)
            {
                //Not valid
                textBox_ReqSize.BackColor = Color.Red;
            }
        }

        private void button_ReqSend_Click(object sender, EventArgs e)
        {
            if (MBmaster == null)
            {
                return;
            }
            
            ushort ID = 1;
            byte unit = Convert.ToByte(textBox_ReqID.Text);
            ushort StartAddress = ushort.Parse(textBox_ReqStrAddress.Text);
            UInt16 Length = Convert.ToUInt16(textBox_ReqSize.Text);

            switch (comboBox_ReqMBFunc.SelectedIndex)
            {
                case 0:
                    ID = 1;
                    MBmaster.ReadCoils(ID, unit, StartAddress, Length);
                    break;
                case 1:
                    ID = 2;
                    MBmaster.ReadDiscreteInputs(ID, unit, StartAddress, Length);
                    break;
                case 2:
                    ID = 3;
                    MBmaster.ReadHoldingRegister(ID, unit, StartAddress, Length);
                    break;
                case 3:
                    ID = 4;
                    MBmaster.ReadInputRegister(ID, unit, StartAddress, Length);
                    break;
                case 4:
                    ID = 5;
                    
                    MBmaster.WriteSingleCoils(ID, unit, StartAddress, Convert.ToBoolean(GetData()[0]));
                    break;
                case 5:
                    ID = 6;
                    MBmaster.WriteSingleRegister(ID, unit, StartAddress, GetData());
                    break;
                case 6:
                    ID = 15;
                    MBmaster.WriteMultipleCoils(ID, unit, StartAddress, Length, GetData());
                    break;
                case 7:
                    ID = 16;
                    MBmaster.WriteMultipleRegister(ID, unit, StartAddress, GetData());
                    break;



            }

            
        }
        private byte[] GetData()
        {
            int count = dataGridView_ReqData.Rows.Count;
            byte[] data;

            if (comboBox_ReqDataType.SelectedIndex == 0)
            {
                //16 bit Int
                data = new byte[count * 2];
                for (int n = 0; n < count; n++)
                {
                    byte[] bit16byte = BitConverter.GetBytes(Int16.Parse(dataGridView_ReqData.Rows[n].Cells[1].Value.ToString()));
                    data[2 * n] = bit16byte[1];
                    data[(2 * n) + 1] = bit16byte[0];
                }
                return data;
            }
            else if (comboBox_ReqDataType.SelectedIndex == 1)
            {
                //32 bit Int
                data = new byte[count * 4];
                for (int n = 0; n < count; n++)
                {
                    byte[] bit32byte = BitConverter.GetBytes(Int32.Parse(dataGridView_ReqData.Rows[n].Cells[1].Value.ToString()));
                    data[4 * n] = bit32byte[1];
                    data[(4 * n) + 1] = bit32byte[0];
                    data[(4 * n) + 2] = bit32byte[3];
                    data[(4 * n) + 3] = bit32byte[2];
                }
                return data;
            }
            else if (comboBox_ReqDataType.SelectedIndex == 2)
            {
                //32 bit Float
                data = new byte[count * 4];
                for (int n = 0; n < count; n++)
                {
                    byte[] bit32byte = BitConverter.GetBytes(float.Parse(dataGridView_ReqData.Rows[n].Cells[1].Value.ToString()));
                    data[4 * n] = bit32byte[1];
                    data[(4 * n) + 1] = bit32byte[0];
                    data[(4 * n) + 2] = bit32byte[3];
                    data[(4 * n) + 3] = bit32byte[2];
                }
                return data;
            }
            else if (comboBox_ReqDataType.SelectedIndex == 3)
            {
                //Bool
                data = new byte[count/8 +1];
                bool[] boolArray = new bool[8];
                for (int n = 0; n < count; n++)
                {
                   boolArray[n%8] = bool.Parse(dataGridView_ReqData.Rows[n].Cells[1].Value.ToString());
                    if (n%8 == 7 || n ==count-1)
                    {
                        int index = 7;
                        // Loop through the array
                        foreach (bool b in boolArray)
                        {
                            // if the element is 'true' set the bit at that position
                            if (b)
                                data[n / 8] |= (byte)(1 << (7 - index));

                            index--;
                        }
                        
                    }
                    
                }
                return data;
            }

            return null;
            

        }
        private void comboBox_ReqDataType_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReadValuesDataGrid();
        }

        private void comboBox_ReqFormat_SelectedIndexChanged(object sender, EventArgs e)
        {
            ReadValuesDataGrid();
        }

        static void ConvertExcelToCsv(string excelFilePath, string csvOutputFile, int worksheetNumber = 1)
        {
            if (!File.Exists(excelFilePath)) throw new FileNotFoundException(excelFilePath);
            if (File.Exists(csvOutputFile)) throw new ArgumentException("File exists: " + csvOutputFile);

            // connection string
            var cnnStr = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=\"Excel 8.0;IMEX=1;HDR=NO\"", excelFilePath);
            var cnn = new OleDbConnection(cnnStr);

            // get schema, then data
            var dt = new DataTable();
            try
            {
                cnn.Open();
                var schemaTable = cnn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                if (schemaTable.Rows.Count < worksheetNumber) throw new ArgumentException("The worksheet number provided cannot be found in the spreadsheet");
                string worksheet = schemaTable.Rows[worksheetNumber - 1]["table_name"].ToString().Replace("'", "");
                string sql = String.Format("select * from [{0}]", worksheet);
                var da = new OleDbDataAdapter(sql, cnn);
                da.Fill(dt);
            }
            catch (Exception e)
            {
                // ???
                throw e;
            }
            finally
            {
                // free resources
                cnn.Close();
            }

            // write out CSV data
            using (var wtr = new StreamWriter(csvOutputFile))
            {
                foreach (DataRow row in dt.Rows)
                {
                    bool firstLine = true;
                    foreach (DataColumn col in dt.Columns)
                    {
                        if (!firstLine) { wtr.Write(","); } else { firstLine = false; }
                        var data = row[col.ColumnName].ToString().Replace("\"", "\"\"");
                        wtr.Write(String.Format("\"{0}\"", data));
                    }
                    wtr.WriteLine();
                }
            }
        }

        private void button_ReportChooseFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog4.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = openFileDialog4.FileName;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_ReportSource.Text = result;
                }
            }
        }

        private void button_ReportChooseLoc_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = folderBrowserDialog1.SelectedPath;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_ReportOutput.Text = result;
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



        private void button_ReportGenerate_Click(object sender, EventArgs e)
        {
            string reportType = "";
            bool success = false;
            //TODO: Save Recents to Properties and set dropDownLists
            if (comboBox_SourceFolder.Text == "")
            {
                if (File.Exists(comboBox_ReportSource.Text))
                {

                    if (Directory.Exists(comboBox_ReportOutput.Text))
                    {
                        ReportGenerator rG = new ReportGenerator();

                        //Report Type
                        if (comboBox_ReportType.SelectedIndex == 0)
                        {
                            success = rG.GenerateExcelCalReport(comboBox_ReportSource.Text, comboBox_ReportOutput.Text,checkBox_showReport.Checked);
                            reportType = "Limerock Excel Report";
                        }
                        else if (comboBox_ReportType.SelectedIndex == 1)
                        {
                            success = rG.GenerateLimerockReport(comboBox_ReportSource.Text, comboBox_HexaneCalc.SelectedIndex, comboBox_ReportOutput.Text,checkBox_showReport.Checked);
                            reportType = "RimRock PDF Report";
                        }
                        else if (comboBox_ReportType.SelectedIndex == 2)
                        {
                            success = rG.GenerateSpreadsheet1(comboBox_ReportSource.Text, comboBox_ReportOutput.Text,checkBox_showReport.Checked);
                            reportType = "Option 1 Spreadsheet";
                        }
                        else if (comboBox_ReportType.SelectedIndex == 3)
                        {
                            success = rG.GenerateRunReportRename(comboBox_ReportSource.Text, comboBox_ReportOutput.Text, textBox_meterID.Text, textBox_meterDesc.Text, checkBox_doAll.Checked,checkBox_showReport.Checked);
                            if (checkBox_doAll.Checked) reportType = "Renamed Reports";
                            else reportType = "Renamed Report";
                        }
                        else if (comboBox_ReportType.SelectedIndex == 4)
                        {
                            success = rG.excelToPDF(comboBox_ReportSource.Text, comboBox_ReportOutput.Text, checkBox_showReport.Checked);
                            reportType = "Excel to PDF";
                        }
                        else if (comboBox_ReportType.SelectedIndex == 6)
                        {
                            success = rG.DriversLicencePDFFillout(comboBox_ReportSource.Text, comboBox_ReportOutput.Text, checkBox_showReport.Checked);
                            reportType = "Drivers Licence PDF Fill Out";
                        }
                        else
                        {
                            MessageBox.Show("Single file generation not supported.", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        if (success)
                        {
                            if(!checkBox_showReport.Checked)MessageBox.Show("Successfully generated " + reportType, "Complete", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        else MessageBox.Show("Failed to generated " + reportType, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    }


                    else
                    {
                        MessageBox.Show("Select a valid Output Folder", "Invalid Output Folder Location",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                }
                else
                {
                    MessageBox.Show("Select a valid Source File", "Invalid Source File Location",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else {
                if (Directory.Exists(comboBox_SourceFolder.Text))
                {
                    string[] files = Directory.GetFiles(comboBox_SourceFolder.Text, openFileDialog4.FileName, SearchOption.AllDirectories);
                    foreach (string file in files)
                    {
                        if (File.Exists(file))
                        {
                            if (Directory.Exists(comboBox_ReportOutput.Text))
                            {
                                ReportGenerator rG = new ReportGenerator();

                                //Limerock Report
                                if (comboBox_ReportType.SelectedIndex == 0)
                                {
                                    success = rG.GenerateExcelCalReport(file, comboBox_ReportOutput.Text, checkBox_showReport.Checked);
                                    reportType = "Limerock Excel Reports";
                                }
                                else if (comboBox_ReportType.SelectedIndex == 1)
                                {
                                    success = rG.GenerateLimerockReport(file, comboBox_HexaneCalc.SelectedIndex, comboBox_ReportOutput.Text, checkBox_showReport.Checked);
                                    reportType = "Limerock PDF Reports";
                                }
                                else if (comboBox_ReportType.SelectedIndex == 2)
                                {
                                    success = rG.GenerateSpreadsheet1(file, comboBox_ReportOutput.Text, checkBox_showReport.Checked);
                                    reportType = "Option 1 Spreadsheets";
                                }
                                else if (comboBox_ReportType.SelectedIndex == 4)
                                {
                                    string outputLoc = comboBox_ReportOutput.Text + "\\" + textBox_ovintivDirectory.Text; ;
                                    string[] directories = {outputLoc + @" Run Reports\Run 3", outputLoc + @" Run Reports", outputLoc + @" Spreadsheets", outputLoc + @" All" };
                                    if (textBox_ovintivDirectory.Text.Any())
                                    {
                                        if (rG.breaksFileNameRules(textBox_ovintivDirectory.Text))
                                        {
                                            MessageBox.Show("Folder name breaks Window's naming rules.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                            break;
                                        }
                                        
                                        if (file == files.First())
                                        {
                                            for(int i = 0; i < 4; i++)
                                            {
                                                
                                                if(i != 1) Directory.CreateDirectory(directories[i]);
                                            }
                                        }
                                        bool show;
                                        if (file == files.Last()) show = checkBox_showReport.Checked;
                                        else show = false;
                                        success = rG.OvintivSendOut(file, outputLoc, show);
                                        reportType = "Ovintiv send-out"; 
                                    }
                                    else
                                    {
                                        MessageBox.Show("Please enter a folder name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        break;
                                    }
                                    if (success)
                                    {
                                        if(file == files.Last())
                                        {
                                            for(int i = 1; i < 4; i++) //Compress folders (make zips)
                                            {
                                                System.IO.Compression.ZipFile.CreateFromDirectory(directories[i], directories[i] + ".zip");
                                            }
                                            for(int i = 0; i < 4; i++)//Remove folders
                                            {
                                                string[] filesToRemove = Directory.GetFiles(directories[i], openFileDialog4.FileName, SearchOption.AllDirectories);
                                                foreach (string fileToRemove in filesToRemove)
                                                {
                                                    File.Delete(fileToRemove);
                                                }
                                                Directory.Delete(directories[i]);
                                            }

                                            if (!checkBox_showReport.Checked) MessageBox.Show("Successfully generated " + reportType, "Complete", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                                        }
                                    }
                                    else MessageBox.Show("Failed to generated " + reportType, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                                else if (comboBox_ReportType.SelectedIndex == 5)
                                {
                                    success = rG.excelToPDF(file, comboBox_ReportOutput.Text, checkBox_showReport.Checked);
                                    reportType = "Excel to PDF";
                                }
                               
                                else
                                {
                                    MessageBox.Show("Directory generation not supported.", "Sorry", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    return;
                                }
                                if (!success)
                                {
                                    MessageBox.Show("Failed to generate with " + file, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);

                                }
                            }
                            else
                            {
                                MessageBox.Show("Select a valid Output Folder", "Invalid Output Folder Location",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Select a valid Source Folder", "Invalid Source File Location",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else MessageBox.Show("Directory not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
        }
    

        private void comboBox_ReportType_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Update Report Type Image and hide/show options
            int index = comboBox_ReportType.SelectedIndex;
            switch (index) {
                case 0:
                    //Limerock Excel
                    pictureBox1.Image = ExcelPaster.Properties.Resources.Excel_Report;
                    openFileDialog4.Filter = "Excel Files | *.xlsx";
                    openFileDialog4.FileName = "*.xlsx";
                    panel_renameOptions.Visible = false;
                    label_renameFileInfo.Visible = false;
                    panel_sourceFolder.Visible = true;
                    panel_hexCalc.Visible = true;
                    panel_ovintivDirectory.Visible = false;
                    label_namingScheme.Visible = false;
                    break;
                    
                case 1:
                    //Limerock pdf
                    pictureBox1.Image = ExcelPaster.Properties.Resources.LimerockPDF;
                    openFileDialog4.Filter = "Notepad Files | *.txt";
                    openFileDialog4.FileName = "*.txt";
                    panel_renameOptions.Visible = false;
                    label_renameFileInfo.Visible = false;
                    panel_sourceFolder.Visible = true;
                    panel_sourceFile.Visible = true;
                    panel_hexCalc.Visible = true;
                    panel_ovintivDirectory.Visible = false;
                    label_namingScheme.Visible = false;
                    break;
                case 2:
                    //PCCU spreadsheet option 1
                    pictureBox1.Image = ExcelPaster.Properties.Resources.spread1;
                    openFileDialog4.Filter = "Notepad Files | *.txt";
                    openFileDialog4.FileName = "*.txt";
                    panel_renameOptions.Visible = false;
                    label_renameFileInfo.Visible = false;
                    panel_sourceFolder.Visible = true;
                    panel_sourceFile.Visible = true;
                    panel_hexCalc.Visible = false;
                    panel_ovintivDirectory.Visible = false;
                    label_namingScheme.Visible = false;
                    break;
                case 3:
                    //Run Report Rename
                    pictureBox1.Image = ExcelPaster.Properties.Resources.Run_Report;
                    openFileDialog4.Filter = "Notepad Files | *.txt";
                    openFileDialog4.FileName = "*.txt";
                    panel_renameOptions.Visible = true;
                    panel_sourceFolder.Visible = false;
                    panel_sourceFile.Visible = true;
                    panel_hexCalc.Visible = false;
                    panel_ovintivDirectory.Visible = false;
                    label_namingScheme.Visible = true;
                    if (checkBox_doAll.Checked) label_renameFileInfo.Visible = true;
                    else label_renameFileInfo.Visible = false;
                    break;
                case 4:
                    //Ovintiv send-out
                    pictureBox1.Image = ExcelPaster.Properties.Resources.send_out;
                    openFileDialog4.Filter = "Notepad Files | *.txt";
                    openFileDialog4.FileName = "*.txt";
                    panel_renameOptions.Visible = false;
                    panel_sourceFolder.Visible = true;
                    panel_sourceFile.Visible = false;
                    label_renameFileInfo.Visible = false;
                    panel_hexCalc.Visible = false;
                    panel_ovintivDirectory.Visible = true;
                    label_namingScheme.Visible = true;
                    break;
                case 6:
                    //Drivers Licences fill out
                    pictureBox1.Image = ExcelPaster.Properties.Resources.DriversLicPic;
                    openFileDialog4.Filter = "Excel Files | *.xlsx";
                    openFileDialog4.FileName = "*.xlsx";
                    panel_renameOptions.Visible = false;
                    panel_sourceFolder.Visible = false;
                    panel_sourceFile.Visible = true;
                    label_renameFileInfo.Visible = false;
                    panel_hexCalc.Visible = false;
                    panel_ovintivDirectory.Visible = false;
                    label_namingScheme.Visible = false;
                    break;
                default:
                    pictureBox1.Image = ExcelPaster.Properties.Resources.No_Report;
                    panel_renameOptions.Visible = false;
                    label_renameFileInfo.Visible = false;
                    break;
            } 

        }

        private void comboBox_ReportSource_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Link;
            else
                e.Effect = DragDropEffects.None;
        }

        private void comboBox_ReportSource_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = e.Data.GetData(DataFormats.FileDrop) as string[]; // get all files droppeds  
            if (files != null && files.Any())
                comboBox_ReportSource.Text = files.First(); //select the first one  
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = folderBrowserDialog1.SelectedPath;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_SourceFolder.Text = result;
                   
                }
            }
        }

        private void doAll_checkbox_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox_doAll.Checked)
            {
                label_renameFileInfo.Visible = true;
            }
            else
            {
                label_renameFileInfo.Visible = false;
            }
        }

        private void comboBox_ReportSource_TextChanged(object sender, EventArgs e)
        {
            if (comboBox_ReportSource.Text == "")
            {
                comboBox_ReportOutput.Text = "";
                return;
            }
            else
            {
                comboBox_SourceFolder.Text = "";
                int i = comboBox_ReportSource.Text.LastIndexOf('\\') + 1;
                if (i > 0 && i < comboBox_ReportSource.Text.Length)
                {
                    comboBox_ReportOutput.Text = comboBox_ReportSource.Text.Remove(i);
                } 
            }
        }

        private void comboBox_SourceFolder_TextChanged(object sender, EventArgs e)
        {
            if (comboBox_SourceFolder.Text == "") return;
            else
            {
                comboBox_ReportSource.Text = "";
                comboBox_ReportOutput.Text = comboBox_SourceFolder.Text; 
            }
        }

        private void comboBox_DTFSourceFile_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button_DTFChangeSource_Click(object sender, EventArgs e)
        {
            if (openFileDialog5.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = openFileDialog5.FileName;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_DTFSourceFile.Text = result;
    
                }
            }
        }

        private void button_DTFChangeOutput_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = folderBrowserDialog1.SelectedPath;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_DTFOutputLocation.Text = result;
                    
                }
            }
        }

        private void button_DTFExtract_Click(object sender, EventArgs e)
        {
            if (comboBox_DTFSourceFile.Text != null)
            {
                if (File.Exists(comboBox_DTFSourceFile.Text))
                {
                    if (comboBox_DTFOutputLocation.Text != null)
                    {
                        if (Directory.Exists(comboBox_DTFOutputLocation.Text))
                        {
                            DTFReader reader = new DTFReader();
                            reader.ExtractRegisters( comboBox_DTFSourceFile.Text);
                            reader.SaveRegisters( comboBox_DTFOutputLocation.Text, checkBox_DTFShowOutput.Checked);
                        }
                        else
                        {
                            MessageBox.Show("Please select a output folder that exists!");
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please select a output folder location!");
                        return;
                    }
                    
                }
                else 
                {
                    MessageBox.Show("Please select a DTF file that exists!");
                    return;
                }
            }
            else 
            {
                MessageBox.Show("Please select a DTF file source!");
                return;
            }
        }

        private void button_DTFReplace_Click(object sender, EventArgs e)
        {

        }


        private void button_DTFChangeTransFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog5.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = openFileDialog5.FileName;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_DTFTransFile.Text = result;

                }
            }
        }

        private void label32_Click(object sender, EventArgs e)
        {

        }

        private void btn_ChangeSourceMRBs_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = folderBrowserDialog1.SelectedPath;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_MRBSourceFolder.Text = result;

                }
            }
        }

        private void btn_ChangeMRBOutput_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = folderBrowserDialog1.SelectedPath;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_MRBOutput.Text = result;

                }
            }
        }

        private void btn_MRBFindAndReplace_Click(object sender, EventArgs e)
        {
            if (comboBox_MRBSourceFolder.Text == null)
            {
                MessageBox.Show("Please select a MRB folder source!");
                return;

            }

            if (comboBox_MRBOutput.Text == null)
            {
                MessageBox.Show("Please select a output folder location!");
                return;
            }

            if (!Directory.Exists(comboBox_MRBOutput.Text))
            {
                MessageBox.Show("Please select a output folder that exists!");
                return;
            }
                        
            FileDataReplacor editor = new FileDataReplacor();
            if (radioButton_ReplaceInt16.Checked)
            {
                if (comboBox_MRBFind.Text == null)
                {
                    MessageBox.Show("Please Enter a Int16 to Find!");
                    return;
                }

                if (comboBox_MRBReplace.Text == null)
                {
                    MessageBox.Show("Please Enter a Int16 to Reaplace!");
                    return;
                }
                editor.ReplaceInt16(comboBox_MRBSourceFolder.Text, comboBox_MRBOutput.Text, byte.Parse(comboBox_MRBFind.Text), byte.Parse(comboBox_MRBReplace.Text));
            }

            if (radioButton_replaceReg.Checked)
            {
                if (comboBox_MRBFindReg.Text == null)
                {
                    MessageBox.Show("Please Enter a Register to Find!");
                    return;
                }

                if (comboBox_MRBReplaceReg.Text == null)
                {
                    MessageBox.Show("Please Enter a Register to Reaplace!");
                    return;
                }
                editor.ReplaceRegister(comboBox_MRBSourceFolder.Text, comboBox_MRBOutput.Text, comboBox_MRBFindReg.Text,comboBox_MRBReplaceReg.Text);
            }

            if (radioButton_ReplaceApp.Checked)
            {
                if (comboBox_MRBFindApp.Text == null)
                {
                    MessageBox.Show("Please Enter a App to Find!");
                    return;
                }

                if (checkBox_MRBReplaceAllApps.Checked == false)
                {
                    if (comboBox_MRBReplaceApp.Text == null)
                    {
                        MessageBox.Show("Please Enter a App to Reaplace!");
                        return;
                    }
                    editor.ReplaceApp(comboBox_MRBSourceFolder.Text, comboBox_MRBOutput.Text, comboBox_MRBFindApp.Text, comboBox_MRBReplaceApp.Text);
                }
                else 
                {
                    editor.GenerateAllApps(comboBox_MRBSourceFolder.Text, comboBox_MRBOutput.Text, comboBox_MRBFindApp.Text);
                }
                
                
            }

            if (radioButton_ReplaceMultReg.Checked)
            {
                if (comboBox_FindReplaceTemplate.Text == null)
                {
                    MessageBox.Show("Please select a .CSV file!");
                    return;
                }

                editor.ReplaceMultipleRegister(comboBox_MRBSourceFolder.Text, comboBox_MRBOutput.Text, comboBox_FindReplaceTemplate.Text);
            }

            if (radioButton_ReplaceMultStrings.Checked)
            {
                if (comboBox_FindReplaceStringsLocations.Text == null)
                {
                    MessageBox.Show("Please select a .CSV file!");
                    return;
                }

                editor.ReplaceMultipleRegister(comboBox_MRBSourceFolder.Text, comboBox_MRBOutput.Text, comboBox_FindReplaceTemplate.Text);
            }


        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void label43_Click(object sender, EventArgs e)
        {

        }

        private void label33_Click(object sender, EventArgs e)
        {

        }

        private void comboBox_MRBOutput_SelectedIndexChanged(object sender, EventArgs e)
        {

        }


        private void button_ChooseFindReplaceTemplate_Click(object sender, EventArgs e)
        {
            if (openFileDialog5.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = openFileDialog5.FileName;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_FindReplaceTemplate.Text = result;

                }
            }
        }

        private void button_ChooseFindReplaceStrings_Click(object sender, EventArgs e)
        {
            if (openFileDialog5.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = openFileDialog5.FileName;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    comboBox_FindReplaceStringsLocations.Text = result;

                }
            }
        }

        public void AppendLog(string value)
        {
            if (InvokeRequired)
            {
                this.Invoke(new Action<string>(AppendLog), new object[] { value });
                return;
            }
            textBox_AutoPCCULog.Text += value;
        }
        public void PostToLogs(string line)
        {
            try
            {
                Debug.WriteLine(line);

                AppendLog(DateTime.Now.ToString() + "> " + line + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
        }
        private void button_PCCUCollect_Click(object sender, EventArgs e)
        {
            //Run Auto Collect Task
            Task.Run(() => RunAutoCollect(textBox_PCCUInstallLocation.Text, checkBox_CloseOnComplete.Checked));

           

        }

        private AutomationElement GetUIWindow(Process proc,int retrys)
        {
            AutomationElement element = null;
            try 
            {
                element = AutomationElement.FromHandle(proc.MainWindowHandle);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
           
            if (element != null) return element;
            else 
            {
                // Retry X Times
                int retrycount = retrys;
                while (element == null)
                {
                    if (retrycount <= 0) return null;
                    System.Threading.Thread.Sleep(1000);
                    Debug.WriteLine("Retrying to find window: "+ proc.MainWindowHandle + " Attempt#: " + retrycount);
                    try 
                    {
                        element = AutomationElement.FromHandle(proc.MainWindowHandle);
                    } 
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                    
                    retrycount--;
                }
                return element;
            }
        }
        private AutomationElement GetUIElement(AutomationElement parent, TreeScope scope, Condition cond, int retrys)
        {
            AutomationElement element = null;
            try
            {
                element = parent.FindFirst(scope, cond);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
            }
           
            if (element != null) return element;
            else
            {
                // Retry X Times
                int retrycount = retrys;
                while (element == null)
                {
                    if (retrycount <= 0) return null;
                    System.Threading.Thread.Sleep(1000);
                    Debug.WriteLine("Retrying to find element: " + cond + " Attempt#: " + retrycount);
                    try 
                    {
                        element = parent.FindFirst(scope, cond);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.Message);
                    }
                    retrycount--;
                }
                return element;
            }
        }
        private void WaitForStatusBar(AutomationElement parent, string waitUntilText)
        {
            AutomationElement statusbar = GetUIElement(parent, TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.StatusBar), 1);
            int maxseconds = 20;
            while (statusbar.Current.Name != waitUntilText)
            {
                if (maxseconds <= 0) break;
                System.Threading.Thread.Sleep(1000);
                statusbar = GetUIElement(parent, TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.StatusBar), 1);
                maxseconds--;
            }
        }
        private AutomationElement WaitForWindowName(Process proc, string waitUntilText, int maxSeconds)
        {
            AutomationElement collectionWindow = null ;
           
            while (true)
            {
                if (maxSeconds <= 0) return null ;
                System.Threading.Thread.Sleep(1000);
                collectionWindow = GetUIWindow(proc, 10);
                if (collectionWindow != null) 
                {
                    if (collectionWindow.Current.Name == waitUntilText)
                    {
                        return collectionWindow;
                    }
                } 
                maxSeconds--;
            }
        }
        private async Task<int> RunAutoCollect(string exeLocation, bool closeOnComplete)
        {
            //Check for current open PCCUs
            PostToLogs("Looking for PCCU...");
            Process[] pccusOpen = Process.GetProcessesByName("PCCU32");
            Process pccu = null;

            if (pccusOpen.Count() == 0)
            {
                if (exeLocation == "")
                {
                    PostToLogs("No PCCU Location Specified! Ending Collect.");
                    return 0;
                }
                //None Open. Open a new PCCU
                PostToLogs("No Open PCCUs! Opening new PCCU...");

                var process = new Process
                {
                    StartInfo = { FileName = exeLocation + "\\pccu32.exe", 
                        UseShellExecute = true, 
                        Verb = "runas", 
                        WorkingDirectory = exeLocation,
                        WindowStyle = ProcessWindowStyle.Normal
                    },
                    EnableRaisingEvents = true
                };

                process.Start();
                pccu = process;
            }
            else
            {
                PostToLogs("Open PCCU Found!");
                pccu = pccusOpen.First();
            }

            //Start hammering for window details
            AutomationElement mainWindow = GetUIWindow(pccu,10);
            //Find Collect button on toolbar and press it
            PostToLogs("Opening Collect...");
            AutomationElement toolbar = GetUIElement(mainWindow,TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ToolBar),10);
            AutomationElement entercollectionbutton = GetUIElement(toolbar,TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Collect"),3);
            InvokePattern patentercollectionbutton = entercollectionbutton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
            patentercollectionbutton.Invoke();

            //Check for 20 seconds if connection was made
            AutomationElement collectionWindow = WaitForWindowName(pccu,"PCCU32 - [Collect]", Int32.Parse(textBox_MaxConnectionTimeCollect.Text));
            if (collectionWindow != null)
            {
                PostToLogs("Starting Collection...");
                //Start collection
                AutomationElement collectwindow = GetUIElement(collectionWindow, TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Collect"), 10);
                AutomationElement collectpane = GetUIElement(collectwindow, TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane), 3);
                AutomationElement collectbutton = GetUIElement(collectpane, TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Collect"), 3);
                InvokePattern collectpatt = collectbutton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                collectpatt.Invoke();

                //Wait for collection to complete
                PostToLogs("Collecting...");
                WaitForStatusBar(collectionWindow, "Collection Complete");

                //close out
                AutomationElement closebutton = GetUIElement(collectpane, TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, "Close"), 20);
                InvokePattern closepatt = closebutton.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                closepatt.Invoke();
                PostToLogs("Closing Collect...");
            }
            else 
            {
                PostToLogs("Failed to Connect to Totalflow! Ending Collect.");
                return 0;
            }
           

            if (closeOnComplete)
            {
                PostToLogs("Closing PCCU...");
                pccu.Kill();
                pccu.WaitForExit();
                pccu.Dispose();
            }

            PostToLogs("Looking for new files...");
            //Find new files
            List<FileInfo> Files = new DirectoryInfo(textBox_CollectSource.Text)
                .GetFiles("*.csv")
                .Where(x => x.LastWriteTime.AddMinutes(5) > DateTime.Now).ToList();
            PostToLogs("Found " + Files.Count() + " new file(s)");

            
            //Send Emails
            if (checkBox_SendEmailEnable.Checked && Files.Count() > 0)
            {
                PostToLogs("Sending to "+ textBox_SendCollectToEmail.Text + " Email with Attachment(s)...");
                await Task.Run(() => EmailResults(textBox_SendCollectToEmail.Text,
                    "AutoCollect Results: " + DateTime.Now,
                    "Data Auto Collected from PCCU is attached.",
                    Files));
            }
            return 1;
        }

        private async Task<int> EmailResults(string email, string subject, string body, List<FileInfo> attachments)
        {
            try
            {
                MailMessage mail = new MailMessage("from@email.com", email);
                mail.Subject = subject;
                mail.Body = body;
                foreach (FileInfo attachment in attachments)
                {
                    mail.Attachments.Add(new Attachment(attachment.FullName));
                }
                

                SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                SmtpServer.Port = 587;
                SmtpServer.UseDefaultCredentials = false;
                SmtpServer.Credentials = new System.Net.NetworkCredential("user@name.com", "password");
                //SmtpServer.EnableSsl = true;
                SmtpServer.Send(mail);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message);
                return 0;
            }
            return 1;
        }



        private void button_ChangePCCUInstallLocation_Click(object sender, EventArgs e)
        {
           
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = folderBrowserDialog1.SelectedPath;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    textBox_PCCUInstallLocation.Text = result;
                }
            }
        }

        private void button_ChangeCollectFolder_Click(object sender, EventArgs e)
        {
           
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                string result = folderBrowserDialog1.SelectedPath;
                if (!string.IsNullOrWhiteSpace(result))
                {
                    textBox_CollectSource.Text = result;
                }
            }
        }

        private void timer_AutoCollectTimer_Tick(object sender, EventArgs e)
        {
            //Every Minute Check
            if (checkBox_ScheduleEnable.Checked)
            {
                DateTime startTime = new DateTime(dateTimePicker_ScheduleStartDate.Value.Year,
                    dateTimePicker_ScheduleStartDate.Value.Month,
                    dateTimePicker_ScheduleStartDate.Value.Day,
                    dateTimePicker_ScheduleStartTime.Value.Hour,
                    dateTimePicker_ScheduleStartTime.Value.Minute,
                    0);
                float countdown = DateTime.Now.Subtract(startTime).Hours % float.Parse(textBox_ScheduleInterval.Text);
                label_CollectCountdown.Text = "Next Collect in: " + countdown + " hours";
                if (countdown == 0)
                {
                    Task.Run(() => RunAutoCollect(textBox_PCCUInstallLocation.Text, checkBox_CloseOnComplete.Checked));
                }
            }
        }
    }
}
