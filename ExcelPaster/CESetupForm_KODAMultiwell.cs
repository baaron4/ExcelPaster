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
    public partial class CESetupForm_KODAMultiwell : Form
    {
        public CESetupForm_KODAMultiwell()
        {
            InitializeComponent();
        }

        private void CESetupForm_KODAMultiwell_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'dbDevicesDataSet.tblModel' table. You can move, or remove it, as needed.
            this.tblModelTableAdapter.Fill(this.dbDevicesDataSet.tblModel);

        }

        private void label88_Click(object sender, EventArgs e)
        {

        }

        private void textBox38_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox39_TextChanged(object sender, EventArgs e)
        {

        }

        private void label87_Click(object sender, EventArgs e)
        {

        }

        private void textBox37_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
