using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            textBox_StartCopyDelayDirect.Text = Properties.Settings.Default.DelayTime.ToString();
            textBox_StartCopyDelayFile.Text = Properties.Settings.Default.DelayTime.ToString();
            comboBox_TargetProgramCSV.SelectedIndex = Properties.Settings.Default.TargetProgram;

            selectedAdapter = NetworkInterface.GetAllNetworkInterfaces().Where(n => n.NetworkInterfaceType != NetworkInterfaceType.Loopback).First(n => n.OperationalStatus == OperationalStatus.Up);
            LoadAdapters();

            SetPadDB();
            
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
            Properties.Settings.Default.RecentFiles.Insert(0, file);
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

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private enum TargetProgram
        {
            TxT = 0,
            Excel = 1,
            PCCU = 2
        }

        private void btn_StartCopyFile_Click(object sender, EventArgs e)
        {
            List<string> BGWorkStorage = new List<string>();
            string CSVFile = comboBox_FileLocation.Text;
            if (CSVFile.Count() > 0)
            {
                label_Status.Text = "Loading File...";
                EnableButtons(ButtonState.COPYING);

                if (!BgWorker.IsBusy)
                {
                    BGWorkStorage.Add(CSVFile);
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
            string fileLoc = BGWorkerList[0];
            int KeypressDelay = Int32.Parse(BGWorkerList[1]);
            int KeypressStateDelay = Int32.Parse(BGWorkerList[2]);

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
                        TargetProgram tgt = ((TargetProgram)Properties.Settings.Default.TargetProgram);
                        if (tgt == TargetProgram.TxT)
                        {
                            typer.TypeCSVtoText(reader.GetArrayStorage(), bg);
                        }
                        else if (tgt == TargetProgram.Excel)
                        {
                            typer.TypeCSVtoExcel(reader.GetArrayStorage(), bg);
                        }
                        else if (tgt == TargetProgram.PCCU)
                        {
                            typer.TypeCSVtoPCCU(reader.GetArrayStorage(), bg);
                        }

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
                    Properties.Settings.Default.Save();
                }

                textBox_StartCopyDelayDirect.Text = Properties.Settings.Default.DelayTime.ToString();
                textBox_StartCopyDelayFile.Text = Properties.Settings.Default.DelayTime.ToString();
            }


        }

        private void comboBox_TargetProgramCSV_SelectedIndexChanged(object sender, EventArgs e)
        {
            Properties.Settings.Default.TargetProgram = comboBox_TargetProgramCSV.SelectedIndex;
            Properties.Settings.Default.Save();
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
            
                foreach (UnicastIPAddressInformation unicastIPAddressInformation in adapter.GetIPProperties().UnicastAddresses)
                {
                    if (unicastIPAddressInformation.Address.AddressFamily == AddressFamily.InterNetwork)
                    {

                        return unicastIPAddressInformation.Address;

                    }
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
        }

        private void LoadAdapters()
        {
           
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
        
        }
        private void button_RefreshAdapter_Click(object sender, EventArgs e)
        {
            adapterList.Clear();
            comboBox_NetworkAdapter.Items.Clear();
            LoadAdapters();
        }

        private void button_ApplyIPChanges_Click(object sender, EventArgs e)
        {
            IPAddress newAddress;
            IPAddress.TryParse( textBox_IPAdress.Text, out newAddress);

            IPAddress newSubMask;
            IPAddress.TryParse(textBox2.Text, out newSubMask);

            IPAddress newGateway;
            IPAddress.TryParse(textBox3.Text, out newGateway);

            Process p = new Process();
            ProcessStartInfo psi = new ProcessStartInfo("netsh", "interface ip set address \""+selectedAdapter.Name +"\" static "+newAddress.ToString()+" "+newSubMask.ToString() +" "+newGateway.ToString()+" 1");
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
                padReader.ParseCSV(Properties.Settings.Default.DatabaseFileLoc);
                foreach (List<string> ListS in padReader.GetArrayStorage())
                {
                    if (ListS.Count >= 5)
                    {
                        PadInfo pInfo = new PadInfo(ListS[0], ListS[1], ListS[2], ListS[3], ListS[4], ListS[5]);

                        PadInfo.Add(pInfo);
                    }
                    
                }
                Companys = PadInfo.Select(s => s.Company).Distinct().ToList();
                Pads = PadInfo.Select(s => s.PadName).Distinct().ToList();
                Devices = PadInfo.Select(s => s.DeviceName).Distinct().ToList();

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
                        //PadInfo = new List<PadInfo>();
                        //Companys = new List<string>();
                        //Pads = new List<string>();
                        //Devices = new List<string>();
                        //SetPadDB();
                    }

                   
                }
            }
        }

        private void textBox_KeypressDelay_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
