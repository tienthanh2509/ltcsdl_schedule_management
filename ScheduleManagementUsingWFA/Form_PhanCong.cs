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
    public partial class Form_PhanCong : Form
    {
        public Form_PhanCong()
        {
            InitializeComponent();
        }

        private void Form_PhanCong_Load(object sender, EventArgs e)
        {

            this.reportViewer1.RefreshReport();
        }
    }
}
