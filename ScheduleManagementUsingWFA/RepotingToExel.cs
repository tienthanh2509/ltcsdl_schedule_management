using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ScheduleManagementUsingWFA
{
    public partial class RepotingToExel : Form
    {
        public RepotingToExel()
        {
            InitializeComponent();
        }

        private void RepotingToExel_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'ScheduleManagementDataSet.XepLich' table. You can move, or remove it, as needed.
            

            this.reportViewer1.RefreshReport();
        }
    }
}
