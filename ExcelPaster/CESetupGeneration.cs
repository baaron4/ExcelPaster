using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelPaster
{
    public class CESetupGeneration
    {
        dbDevicesDataSet.tblModelDataTable tblModelDataTable = new dbDevicesDataSet.tblModelDataTable();

        FlowLayoutPanel[] flps = new FlowLayoutPanel[10];

        dbDevicesDataSet db = new dbDevicesDataSet();

        string[] labelArray = new string[] { "Number of Wells:", "Number of Treaters:","Number of Oil Tanks:", "Number of Salt Water Tanks:",
        "Number of Fresh Water Tanks:", "Number of SW Disposal Systems:", "Number of FW Systems:","Number of flare Systems:","Number of Recycle Pumps:",
        "Number of Glycole Heaters:"};

        int[] defaultSetupValues = new int[] { 8, 8, 12, 12, 1, 1, 1, 2, 1, 1 };

        string[] wellLabelArray = new string[] { "Well Number", "Well Name", "Drive Type","MOV Model","Tubing PSI","Casing PSI" };
        public enum ProjectType : int
        { 
            KODA_MultiWell = 0,
            CPE_MultiWell = 1
        }
        public TabControl GenerateSetupInterface(ProjectType pType, TabControl tabControl)
        {
            LoadDatabase();

            TabControl newTabControl = tabControl;
            newTabControl.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            if (pType == ProjectType.KODA_MultiWell)
            {
                //Systems Tab
                newTabControl.TabPages[0].Text = "Systems";
                newTabControl.TabPages[0].AutoScroll = true;

                flps[0] = new FlowLayoutPanel();
                flps[0].AutoSize = true;
                flps[0].FlowDirection = FlowDirection.LeftToRight;

                Label label_sitename = new Label();
                label_sitename.Text = "Site Name:";
                label_sitename.Width = 200;
                flps[0].Controls.Add(label_sitename);
                TextBox textBox_SiteName = new TextBox();
                textBox_SiteName.Text = "Name of Site";
                flps[0].Controls.Add(textBox_SiteName);
                flps[0].SetFlowBreak(textBox_SiteName, true);

                for (int i = 0; i < labelArray.Length;  i++) {

                    Label nLabel = new Label();
                    nLabel.Text = labelArray[i];
                    nLabel.Width = 200;
                    flps[0].Controls.Add(nLabel);
                    TextBox ntextBox = new TextBox();
                    ntextBox.Text = Convert.ToString(defaultSetupValues[i]);
                    flps[0].Controls.Add(ntextBox);
                    flps[0].SetFlowBreak(ntextBox, true);
                }
                newTabControl.TabPages[0].Controls.Add(flps[0]);

                //Well Systems Tab
                newTabControl.TabPages.Add( "Well Systems");
                newTabControl.TabPages[1].AutoScroll = true;

                flps[1] = new FlowLayoutPanel();
                flps[1].AutoSize = true;
                flps[1].FlowDirection = FlowDirection.LeftToRight;

                for (int i = 0; i < wellLabelArray.Length; i++)
                {
                    Label nLabel = new Label();
                    nLabel.Text = wellLabelArray[i];
                    nLabel.Width = 100;
                    flps[1].Controls.Add(nLabel);
                }
                flps[1].SetFlowBreak(flps[1].Controls[wellLabelArray.Length-1], true);
                //Content
                int valveid = db.tblMVFeeder.FirstOrDefault(x => x.Name == "Valve").ID;
                for (int i = 0; i < defaultSetupValues[0]; i++)
                {
                    Label nLabel = new Label();
                    nLabel.Text = "Well " + (i+1);
                    nLabel.Width = 100;
                    flps[1].Controls.Add(nLabel);

                    TextBox ntextBox = new TextBox();
                    ntextBox.Text = Convert.ToString("Well " + (i+1));
                    flps[1].Controls.Add(ntextBox);

                    ComboBox ndriveComboBox = new ComboBox();
                    ndriveComboBox.Items.AddRange(new string[] { "Freeflowing", "ESP","PumpJack", "Rotoflex","Jetpump","GasLift" });
                    flps[1].Controls.Add(ndriveComboBox);

                    ComboBox nmovComboBox = new ComboBox();
                    nmovComboBox.Items.AddRange(db.tblModel.Where(y => y.modelMV == valveid).Select(z=> z.modelName).ToArray());
                    flps[1].Controls.Add(nmovComboBox);
                }
                newTabControl.TabPages[1].Controls.Add(flps[1]);
            }
            return newTabControl;
        }

        private void LoadDatabase()
        {
            
            dbDevicesDataSetTableAdapters.tblBrandFeederTableAdapter brandFeederTableAdapter = new dbDevicesDataSetTableAdapters.tblBrandFeederTableAdapter();
            brandFeederTableAdapter.Fill(db.tblBrandFeeder);

            dbDevicesDataSetTableAdapters.tblModelTableAdapter modelTableAdapter = new dbDevicesDataSetTableAdapters.tblModelTableAdapter();
            modelTableAdapter.Fill(db.tblModel);

            dbDevicesDataSetTableAdapters.tblMVFeederTableAdapter MVTableAdapter = new dbDevicesDataSetTableAdapters.tblMVFeederTableAdapter();
            MVTableAdapter.Fill(db.tblMVFeeder);

            dbDevicesDataSetTableAdapters.tblSpecificationsTableAdapter specificationsTableAdapter = new dbDevicesDataSetTableAdapters.tblSpecificationsTableAdapter();
            specificationsTableAdapter.Fill(db.tblSpecifications);

            dbDevicesDataSetTableAdapters.tblTypeFeederTableAdapter typeTableAdapter = new dbDevicesDataSetTableAdapters.tblTypeFeederTableAdapter();
            typeTableAdapter.Fill(db.tblTypeFeeder);
        }
    }
}
