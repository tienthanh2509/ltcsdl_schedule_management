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
    public partial class frm_Chinh : Form
    {
        public frm_Chinh()
        {
            InitializeComponent();
        }

        private void frm_Chinh_Load(object sender, EventArgs e)
        {
            frm_Chinh frm = new frm_Chinh();
            frm.Width = 1024;
            frm.Height = 768;
        }
    }
}
