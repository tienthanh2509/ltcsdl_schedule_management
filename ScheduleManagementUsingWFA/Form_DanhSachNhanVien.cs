using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
namespace ScheduleManagementUsingWFA
{
    public partial class Form_DanhSachNhanVien : Form
    {
        public Form_DanhSachNhanVien()
        {
            InitializeComponent();
            // disable control

            //disableButtonChonLichRanh(false);



        }
        ScheduleManagementDataContext db = new ScheduleManagementDataContext();
        List<ChiTietNhanVien> chiTietNhanViens = new List<ChiTietNhanVien>();
        List<NhanVien> nhanViens = new List<NhanVien>();
        List<ThoiGianLamViec> thoiGianLamViecs = new List<ThoiGianLamViec>();
        //List<CaLamViec> caLamViecs = new List<CaLamViec>();
        //List<ViTriLamViec> viTriLamViec = new List<ViTriLamViec>();
        //List<PhanCongLamViec> phanCongLamViec = new List<PhanCongLamViec>();



        // enable/disable
        private void disableCommonNhanVien(bool b)
        {

            txtTenNhanVien.Enabled = b;
            cmbThoiGianLV.Enabled = b;
        }
        private void disableButtonNhanVien(bool b1, bool b2)
        {
            btnCapNhatNV.Enabled = b1;
            btnXoaNV.Enabled = b1;
            btnThemNV.Enabled = b1;
            btnLuu.Enabled = b2;
            btnHuy.Enabled = b2;
        }
        private void disableCommonChonLichRanh(bool b)
        {
            cmbChonTenNV.Enabled = b;
            //cmbMauNV.Enabled = b;
        }
        private void disableButtonChonLichRanh(bool b)
        {
            btnCapNhatLichRanh.Enabled = !b;
            btnLuuLichRanh.Enabled = b;
            btnHuyLichRanh.Enabled = b;
        }
        private void disableDgvChonLichRanh(bool b)
        {
            chkBoxChonTatCaCaChieu.Enabled = b;
            chkBoxChonTatCaCaSang.Enabled = b;
            chkBoxChonTatCaCaToi.Enabled = b;
            dgvCaChieu.Enabled = b;
            dgvCaSang.Enabled = b;
            dgvCaToi.Enabled = b;
        }

        private void loadDataTabChonLichRanh()
        {


            // datasource -> combobox
            cmbChonTenNV.DataSource = nhanViens;
            cmbChonTenNV.DisplayMember = "HoTen";
            cmbChonTenNV.ValueMember = "Id";
            chiTietNhanViens = db.ChiTietNhanViens.ToList();
            // binding data -> combobox
            //int temp;
            //bool b=Int32.TryParse(cmb2.SelectedValue.ToString(),out temp);

            try
            {
                var mau = ((from ctnv in chiTietNhanViens
                            where ctnv.Id == (int)cmbChonTenNV.SelectedValue
                            select ctnv.Mau).ToList()).FirstOrDefault();

                cmbMauNV.BackColor = ColorTranslator.FromHtml(mau.ToString());

                // databinding to datagridview
                //DataSet dsLichRanh = new DataSet();
                // BindingSource bs = new BindingSource();
                //  bs.DataSource = typeof(ChiTietNhanVien);

                // chon ra lich ranh ca sang cua nhan vien dang chon
                var dsLichRanhCaSang = (from ctnv in chiTietNhanViens
                                        where (ctnv.Id == Convert.ToInt32(cmbChonTenNV.SelectedValue) && ctnv.CaLV == 1)
                                        select ctnv.ThuLV).ToList();

                // chon ra lich ranh ca chieu cua nhan vien dang chon
                var dsLichRanhCaChieu = (from ctnv in chiTietNhanViens
                                         where (ctnv.Id == Convert.ToInt32(cmbChonTenNV.SelectedValue) && ctnv.CaLV == 2)
                                         select ctnv.ThuLV).ToList();
                // chon ra lich ranh ca toi cua nhan vien dang chon
                var dsLichRanhCaToi = (from ctnv in chiTietNhanViens
                                       where (ctnv.Id == Convert.ToInt32(cmbChonTenNV.SelectedValue) && ctnv.CaLV == 3)
                                       select ctnv.ThuLV).ToList();
                foreach (int j in dsLichRanhCaSang)
                {
                    dgvCaSang.Rows[0].Cells[j - 2].Value = true;
                }
                foreach (int j in dsLichRanhCaChieu)
                {
                    dgvCaChieu.Rows[0].Cells[j - 2].Value = true;
                }
                foreach (int j in dsLichRanhCaToi)
                {
                    dgvCaToi.Rows[0].Cells[j - 2].Value = true;
                }
            }
            catch (Exception)
            {

                clearOrSelectAllColumnsChecked(false);
            }

        }
        private void Form_DanhSachNhanVien_Load(object sender, EventArgs e)
        {
            tabNhanVien.AttachedControl.Select();
            disableCommonNhanVien(false);
            disableButtonNhanVien(true, false);
            nhanViens = db.NhanViens.ToList();
            // datasource -> datagridview
            dgvDSNhanVien.DataSource = nhanViens;

            thoiGianLamViecs = db.ThoiGianLamViecs.ToList();
            // datasource -> combobox
            cmbThoiGianLV.DataSource = thoiGianLamViecs;
            cmbThoiGianLV.DisplayMember = "ThoiGian";
            cmbThoiGianLV.ValueMember = "Id";


            //// datasource -> combobox
            //cmbChonTenNV.DataSource = nhanViens;
            //cmbChonTenNV.DisplayMember = "HoTen";
            //cmbChonTenNV.ValueMember = "Id";
            //// binding data -> combobox
            //var mau = (from ctnv in db.ChiTietNhanViens
            //           where ctnv.Id == (int)cmbChonTenNV.SelectedValue
            //           select ctnv.Mau).FirstOrDefault();
            //cmbMauNV.BackColor = ColorTranslator.FromHtml(mau.ToString());

            //// databinding to datagridview
            ////DataSet dsLichRanh = new DataSet();
            //// BindingSource bs = new BindingSource();
            ////  bs.DataSource = typeof(ChiTietNhanVien);

            //// chon ra lich ranh ca sang cua nhan vien dang chon
            //var dsLichRanhCaSang = (from ctnv in db.ChiTietNhanViens
            //                        where (ctnv.Id == Convert.ToInt32(cmbChonTenNV.SelectedValue) && ctnv.CaLV == 1)
            //                        select ctnv.ThuLV).ToList();

            //// chon ra lich ranh ca chieu cua nhan vien dang chon
            //var dsLichRanhCaChieu = (from ctnv in db.ChiTietNhanViens
            //                         where (ctnv.Id == Convert.ToInt32(cmbChonTenNV.SelectedValue) && ctnv.CaLV == 2)
            //                         select ctnv.ThuLV).ToList();
            //// chon ra lich ranh ca toi cua nhan vien dang chon
            //var dsLichRanhCaToi = (from ctnv in db.ChiTietNhanViens
            //                       where (ctnv.Id == Convert.ToInt32(cmbChonTenNV.SelectedValue) && ctnv.CaLV == 3)
            //                       select ctnv.ThuLV).ToList();
            //foreach (int j in dsLichRanhCaSang)
            //{
            //    dgvCaSang.Rows[0].Cells[j - 2].Value = true;
            //}
            //foreach (int j in dsLichRanhCaChieu)
            //{
            //    dgvCaChieu.Rows[0].Cells[j - 2].Value = true;
            //}
            //foreach (int j in dsLichRanhCaToi)
            //{
            //    dgvCaToi.Rows[0].Cells[j - 2].Value = true;
            //}
            //loadDataTabChonLichRanh();
        }

        private void expandableSplitter1_ExpandedChanged(object sender, DevComponents.DotNetBar.ExpandedChangeEventArgs e)
        {

        }

        private void tabControl1_Click(object sender, EventArgs e)
        {

        }

        private void labelX1_Click(object sender, EventArgs e)
        {

        }

        private void labelX2_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void tabControlPanel3_Click(object sender, EventArgs e)
        {

        }

        private void dgvCaSang_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView10_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {

        }

        private void tabControlPanel1_Click(object sender, EventArgs e)
        {

        }

        private void cmbMauNV_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbMauNV_Click(object sender, EventArgs e)
        {
            DialogResult dr = colorDialog1.ShowDialog();
            if (dr == DialogResult.OK)
                cmbMauNV.BackColor = colorDialog1.Color;
        }

        private void btnLuuLichRanh_Click(object sender, EventArgs e)
        {
            // variables
            int i;
            List<ChiTietNhanVien> ctnv = new List<ChiTietNhanVien>();
            ChiTietNhanVien ct;


            // dgvCaSang
            for (i = 0; i < 7; i++)
            {
                if (dgvCaSang.Rows[0].Cells[i].Value != null && (bool)dgvCaSang.Rows[0].Cells[i].Value)
                {
                    // Checked
                    ct = new ChiTietNhanVien();
                    ct.Mau = HexConverter(cmbMauNV.BackColor);
                    ct.ThuLV = (i + 2);
                    ct.CaLV = 1;
                    ct.Id = Convert.ToInt32(cmbChonTenNV.SelectedValue);
                    ctnv.Add(ct);
                }
                else if (dgvCaSang.Rows[0].Cells[i].Value == null)
                {
                    // Unchecked
                }
            }
            // dgvCaChieu
            for (i = 0; i < 7; i++)
            {
                if (dgvCaChieu.Rows[0].Cells[i].Value != null && (bool)dgvCaChieu.Rows[0].Cells[i].Value)
                {
                    // Checked
                    ct = new ChiTietNhanVien();
                    ct.Mau = HexConverter(cmbMauNV.BackColor);
                    ct.ThuLV = (i + 2);
                    ct.CaLV = 2;
                    ct.Id = Convert.ToInt32(cmbChonTenNV.SelectedValue);
                    ctnv.Add(ct);

                }
                else if (dgvCaChieu.Rows[0].Cells[i].Value == null)
                {
                    // Unchecked
                }
            }
            // dgvCaToi
            for (i = 0; i < 7; i++)
            {
                if (dgvCaToi.Rows[0].Cells[i].Value != null && (bool)dgvCaToi.Rows[0].Cells[i].Value)
                {
                    // Checked
                    ct = new ChiTietNhanVien();
                    ct.Mau = HexConverter(cmbMauNV.BackColor);
                    ct.ThuLV = (i + 2);
                    ct.CaLV = 3;
                    ct.Id = Convert.ToInt32(cmbChonTenNV.SelectedValue);
                    ctnv.Add(ct);

                }
                else if (dgvCaToi.Rows[0].Cells[i].Value == null)
                {
                    // Unchecked
                }
            }
            db.ChiTietNhanViens.InsertAllOnSubmit(ctnv);
            db.SubmitChanges();
            //db.Refresh(System.Data.Linq.RefreshMode.KeepChanges);
            MessageBox.Show("Đã lưu nhân viên thành công!");
            clearOrSelectAllColumnsChecked(false);
        }




        private void btnHuyLichRanh_Click(object sender, EventArgs e)
        {
            tabLichRanh_Click(sender, e);
            //clearOrSelectAllColumnsChecked(false);
            //cmbChonTenNV.Text = "";
            //cmbMauNV.BackColor = Color.White;

        }
        // convert color to hex=================================================================
        private static String HexConverter(System.Drawing.Color c)
        {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }
        // convert color to RGB=================================================================
        private static String RGBConverter(System.Drawing.Color c)
        {
            return "RGB(" + c.R.ToString() + "," + c.G.ToString() + "," + c.B.ToString() + ")";
        }
        private void clearOrSelectAllColumnsChecked(bool b)
        {
            for (int i = 0; i < 7; i++)
            {
                if (b == false)
                {
                    dgvCaSang.Rows[0].Cells[i].Value = b;
                    dgvCaChieu.Rows[0].Cells[i].Value = b;
                    dgvCaToi.Rows[0].Cells[i].Value = b;
                }
                else
                {
                    dgvCaSang.Rows[0].Cells[i].Value = !b;
                    dgvCaChieu.Rows[0].Cells[i].Value = !b;
                    dgvCaToi.Rows[0].Cells[i].Value = !b;
                }
            }
        }
        private void clearAllColumnsCheckedCaSang(bool b)
        {
            for (int i = 0; i < 7; i++)
            {
                if (b == true)
                    dgvCaSang.Rows[0].Cells[i].Value = false;
                else
                    dgvCaSang.Rows[0].Cells[i].Value = true;

            }
        }
        private void clearAllColumnsCheckedCaChieu(bool b)
        {
            for (int i = 0; i < 7; i++)
            {

                if (b == true)
                    dgvCaChieu.Rows[0].Cells[i].Value = false;
                else
                    dgvCaChieu.Rows[0].Cells[i].Value = true;

            }
        }
        private void clearAllColumnsCheckedCaToi(bool b)
        {
            for (int i = 0; i < 7; i++)
            {

                if (b == true)
                    dgvCaToi.Rows[0].Cells[i].Value = false;
                else
                    dgvCaToi.Rows[0].Cells[i].Value = true;
            }
        }
        private void cmbChonTenNV_SelectedIndexChanged(object sender, EventArgs e)
        {
            clearOrSelectAllColumnsChecked(false);
            loadDataTabChonLichRanh();
            //// datasource -> combobox
            //cmbChonTenNV.DataSource = nhanViens;
            //cmbChonTenNV.DisplayMember = "HoTen";
            //cmbChonTenNV.ValueMember = "Id";
            //// binding data -> combobox
            //var mau = (from ctnv in db.ChiTietNhanViens
            //           where ctnv.Id == (int)cmbChonTenNV.SelectedValue
            //           select ctnv.Mau).FirstOrDefault();
            //cmbMauNV.BackColor = ColorTranslator.FromHtml(mau.ToString());

            //// databinding to datagridview
            ////DataSet dsLichRanh = new DataSet();
            //// BindingSource bs = new BindingSource();
            ////  bs.DataSource = typeof(ChiTietNhanVien);

            //// chon ra lich ranh ca sang cua nhan vien dang chon
            //var dsLichRanhCaSang = (from ctnv in db.ChiTietNhanViens
            //                        where (ctnv.Id == Convert.ToInt32(cmbChonTenNV.SelectedValue) && ctnv.CaLV == 1)
            //                        select ctnv.ThuLV).ToList();

            //// chon ra lich ranh ca chieu cua nhan vien dang chon
            //var dsLichRanhCaChieu = (from ctnv in db.ChiTietNhanViens
            //                         where (ctnv.Id == Convert.ToInt32(cmbChonTenNV.SelectedValue) && ctnv.CaLV == 2)
            //                         select ctnv.ThuLV).ToList();
            //// chon ra lich ranh ca toi cua nhan vien dang chon
            //var dsLichRanhCaToi = (from ctnv in db.ChiTietNhanViens
            //                       where (ctnv.Id == Convert.ToInt32(cmbChonTenNV.SelectedValue) && ctnv.CaLV == 3)
            //                       select ctnv.ThuLV).ToList();
            //foreach (int j in dsLichRanhCaSang)
            //{
            //    dgvCaSang.Rows[0].Cells[j - 2].Value = true;
            //}
            //foreach (int j in dsLichRanhCaChieu)
            //{
            //    dgvCaChieu.Rows[0].Cells[j - 2].Value = true;
            //}
            //foreach (int j in dsLichRanhCaToi)
            //{
            //    dgvCaToi.Rows[0].Cells[j - 2].Value = true;
            //}
        }

        private void chkBoxChonTatCaCaSang_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBoxChonTatCaCaSang.Checked)
            {
                clearAllColumnsCheckedCaSang(false);
            }
            else
            {
                clearAllColumnsCheckedCaSang(true);
            }
        }

        private void chkBoxChonTatCaCaChieu_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBoxChonTatCaCaChieu.Checked)
            {
                clearAllColumnsCheckedCaChieu(false);
            }
            else
            {
                clearAllColumnsCheckedCaChieu(true);
            }
        }

        private void chkBoxChonTatCaCaToi_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBoxChonTatCaCaToi.Checked)
            {
                clearAllColumnsCheckedCaToi(false);
            }
            else
            {


                clearAllColumnsCheckedCaToi(true);
            }
        }

        private void loadTab(bool bTabNhanVien, bool bTabLichRanh, bool bTabPhanCong)
        {
            tabPhanCong.AttachedControl.Enabled = bTabPhanCong;
            tabLichRanh.AttachedControl.Enabled = bTabLichRanh;
            tabNhanVien.AttachedControl.Enabled = bTabNhanVien;
        }

        private void tabNhanVien_Click(object sender, EventArgs e)
        {
            loadTab(true, false, false);
            nhanViens = db.NhanViens.ToList();

            thoiGianLamViecs = db.ThoiGianLamViecs.ToList();
            // datasource -> combobox
            cmbThoiGianLV.DataSource = thoiGianLamViecs;
            cmbThoiGianLV.DisplayMember = "ThoiGian";
            cmbThoiGianLV.ValueMember = "Id";

            // datasource -> datagridview
            dgvDSNhanVien.DataSource = nhanViens;
        }

        private void tabLichRanh_Click(object sender, EventArgs e)
        {
            loadTab(false, true, false);
            loadDataTabChonLichRanh();
            disableButtonChonLichRanh(false);
            disableCommonChonLichRanh(true);
            disableDgvChonLichRanh(false);
        }
        bool them = false;
        private void tabPhanCong_Click(object sender, EventArgs e)
        {
            loadTab(false, false, true);
        }

        private void btnThemNV_Click(object sender, EventArgs e)
        {
            them = true;
            disableButtonNhanVien(false, true);
            disableCommonNhanVien(true);
        }
        
        private void btnLuu_Click(object sender, EventArgs e)
        {
            if (them)
            {
                NhanVien nv = new NhanVien();
                nv.HoTen = txtTenNhanVien.Text;
                nv.ThoiGianLV = Convert.ToInt32(cmbThoiGianLV.SelectedValue);
                db.NhanViens.InsertOnSubmit(nv);
                db.SubmitChanges();
                MessageBox.Show("Lưu nhân viên thành công!");
                btnHuy_Click(sender, e);
            }
            else {
                db.NhanViens.I
                db.SubmitChanges();
            }
        }


        private void btnHuy_Click(object sender, EventArgs e)
        {
            Form_DanhSachNhanVien_Load(sender, e);
            txtTenNhanVien.Text = "";
            cmbThoiGianLV.SelectedItem = "";
        }

        private void btnXoaNV_Click(object sender, EventArgs e)
        {
            int index = Convert.ToInt32(dgvDSNhanVien.CurrentCell.RowIndex);
            int getId = Convert.ToInt32(dgvDSNhanVien.Rows[index].Cells[0].Value);

            var nv = db.NhanViens.Single(x => x.Id == getId);
            db.NhanViens.DeleteOnSubmit(nv);
            db.SubmitChanges();

            MessageBox.Show("Đã xóa nhân viên thành công!");
            tabNhanVien_Click(sender, e);
        }

        private void btnCapNhatLichRanh_Click(object sender, EventArgs e)
        {
            disableButtonChonLichRanh(true);
            disableCommonChonLichRanh(false);
            disableDgvChonLichRanh(true);

        }

        private void btnCapNhatNV_Click(object sender, EventArgs e)
        {
            them = false;
            disableButtonNhanVien(false, true);
            disableCommonNhanVien(true);
            
        }

        private void dgvDSNhanVien_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int index = Convert.ToInt32(dgvDSNhanVien.CurrentCell.RowIndex);
            //int getId = Convert.ToInt32(dgvDSNhanVien.Rows[index].Cells[0].Value);
            int getIdTGLV=Convert.ToInt32(dgvDSNhanVien.Rows[index].Cells[2].Value);

            txtTenNhanVien.Text = dgvDSNhanVien.Rows[index].Cells[1].Value.ToString();

            ThoiGianLamViec tglv = db.ThoiGianLamViecs.Single(x=>x.Id==getIdTGLV);

            cmbThoiGianLV.Text = tglv.ThoiGian.ToString();
            
            
            
        }

        private void dgvDSNhanVien_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dgvDSNhanVien_CellClick(sender, e);
            
        }
    }
}
