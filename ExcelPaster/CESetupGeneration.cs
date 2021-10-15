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
       
        FlowLayoutPanel[] flps = new FlowLayoutPanel[10];

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
            TabControl newTabControl = tabControl;
            if (pType == ProjectType.KODA_MultiWell)
            {
                //Systems Tab
                newTabControl.TabPages[0].Text = "Systems";

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
                newTabControl.TabPages[1].Controls.Add(flps[1]);
            }
            return newTabControl;
        }
    }
}
