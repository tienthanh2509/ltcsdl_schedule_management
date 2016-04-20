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
using OfficeExel = Microsoft.Office.Interop.Excel;
using Microsoft.Reporting.WinForms;
using System.Reflection;

namespace ScheduleManagementUsingWFA
{
    public partial class Form_Main : Form
    {
        public Form_Main()
        {
            InitializeComponent();
        }
        // Initialize LINQ to Classes =========================
        ScheduleManagementDataContext db = new ScheduleManagementDataContext();

        // Initialize lists ===================================
        List<ChiTietNhanVien> chiTietNhanViens;
        List<NhanVien> nhanViens;
        List<ThoiGianLamViec> thoiGianLamViecs;
        //List<CaLamViec> caLamViecs;
        //List<ViTriLamViec> viTriLamViec;
        //List<PhanCongLamViec> phanCongLamViec;

        // Initialize variables ===============================
        bool them = false; // flag phân biệt insert/update
        int getIdNV;
        int getIdTGLV;


        /*=================================================== Methods ================================================================================*/

        // ====================================== Tab Nhan Vien ===================================
        #region
        // enable/disable control =======================================
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

        // ====================================== Tab Lich Ranh ===================================
        // enable/disable control ============================
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

        private void checkOrUncheckCheckBox(bool b)
        {
            chkBoxChonTatCaCaChieu.Checked = b;
            chkBoxChonTatCaCaSang.Checked = b;
            chkBoxChonTatCaCaToi.Checked = b;
        }

        private void disableCheckBox(bool b)
        {
            chkBoxChonTatCaCaChieu.Enabled = b;
            chkBoxChonTatCaCaSang.Enabled = b;
            chkBoxChonTatCaCaToi.Enabled = b;
        }

        private void disableDgvChonLichRanh(bool b)
        {
            dgvCaChieu.Enabled = b;
            dgvCaSang.Enabled = b;
            dgvCaToi.Enabled = b;
        }

        // Check/Uncheck all CheckBoxCells =============================================
        private void clearOrSelectAllColumnsChecked(bool b)
        {
            for (int i = 0; i < 7; i++)// Duyệt hết 7 ngày trong tuần
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

        // Check/Uncheck all CheckBoxCells ca sáng ====================================
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

        // Check/Uncheck all CheckBoxCells ca chiều ====================================
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

        // Check/Uncheck all CheckBoxCells ca tối ======================================
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

        // Load data tab ============================================================
        private void loadDataTabChonLichRanh()
        {
            // datasource -> combobox
            cmbChonTenNV.DataSource = nhanViens;
            cmbChonTenNV.DisplayMember = "HoTen";
            cmbChonTenNV.ValueMember = "Id";
            chiTietNhanViens = db.ChiTietNhanViens.ToList();

            // Kiểm tra nếu đã tồn tại màu thì sẽ load vào combobox
            // Nếu chưa tồn tại thì sẽ tương ứng với nhân viên này chưa được cập nhật lịch rãnh và màu.
            try
            {
                them = false;// flag

                // Lọc ra màu của nhân viên đang chọn trong combox
                var mau = ((from ctnv in chiTietNhanViens
                            where ctnv.Id == (int)cmbChonTenNV.SelectedValue
                            select ctnv.Mau).ToList()).FirstOrDefault();

                // set background color
                cmbMauNV.BackColor = ColorTranslator.FromHtml(mau.ToString());

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

                // load lịch rãnh vào từng datagridview
                foreach (int j in dsLichRanhCaSang)// ca sáng
                {
                    dgvCaSang.Rows[0].Cells[j - 2].Value = true;
                }
                foreach (int j in dsLichRanhCaChieu)// ca chiều
                {
                    dgvCaChieu.Rows[0].Cells[j - 2].Value = true;
                }
                foreach (int j in dsLichRanhCaToi)// ca tối
                {
                    dgvCaToi.Rows[0].Cells[j - 2].Value = true;
                }
            }
            catch (Exception)
            {
                them = true; // flag

                // uncheck all CheckBoxCell 
                clearOrSelectAllColumnsChecked(false);
            }

        }
        #endregion
        // ====================================== Tab Phan Cong =========================
        #region
        // Xếp lịch ==========================
        private void xepLich()
        {
            // Drop table PhanCongLamViec
            db.PhanCongLamViecs.DeleteAllOnSubmit(from pclv in db.PhanCongLamViecs
                                                  select pclv);
            db.SubmitChanges();
            // Create list to getting data
            List<ChiTietNhanVien> lLichRanhCaSang = new List<ChiTietNhanVien>();
            List<ChiTietNhanVien> lLichRanhCaChieu = new List<ChiTietNhanVien>();
            List<ChiTietNhanVien> lLichRanhCaToi = new List<ChiTietNhanVien>();

            // Chọn ra tất cả lịch rãnh mỗi ngày, mỗi ca lưu vào list tương ứng
            var getLichRanhCaSang = (from ctnv in db.ChiTietNhanViens
                                     where ctnv.CaLV == 1
                                     select ctnv);
            var getLichRanhCaChieu = (from ctnv in db.ChiTietNhanViens
                                      where ctnv.CaLV == 2
                                      select ctnv);
            var getLichRanhCaToi = (from ctnv in db.ChiTietNhanViens
                                    where ctnv.CaLV == 3
                                    select ctnv);

            // Set data into lists
            lLichRanhCaSang = getLichRanhCaSang.ToList();
            lLichRanhCaChieu = getLichRanhCaChieu.ToList();
            lLichRanhCaToi = getLichRanhCaToi.ToList();

            // Chọn ra theo từng ngày, theo mỗi ca số lượng lịch rãnh
            var lCTNVTheoNgay = new List<List<ChiTietNhanVien>>();
            int i, j;
            for (j = 0; j < 7; j++)// 7 ngày trong tuần
            {
                List<ChiTietNhanVien> ctnv = new List<ChiTietNhanVien>();
                // Ca sáng
                for (i = 0; i < lLichRanhCaSang.Count; i++)
                {
                    if (lLichRanhCaSang[i].ThuLV == (j + 2))
                    {

                        ctnv.Add(lLichRanhCaSang[i]);
                    }
                }
                // Ca chiều
                for (i = 0; i < lLichRanhCaChieu.Count; i++)
                {
                    if (lLichRanhCaChieu[i].ThuLV == (j + 2))
                    {
                        ctnv.Add(lLichRanhCaChieu[i]);
                    }
                }
                // Ca tối
                for (i = 0; i < lLichRanhCaToi.Count; i++)
                {
                    if (lLichRanhCaToi[i].ThuLV == (j + 2))
                    {
                        ctnv.Add(lLichRanhCaToi[i]);
                    }
                }
                lCTNVTheoNgay.Add(ctnv);
            }

            /* ============= Algorithm =================
             * Kiểm tra số lượng nhân viên trong ca
             *  - Ca sáng: Min = 4 người, Max = 5 người
             *  - Ca chiều: Min = 2 người, Max = 3 người
             *  - Ca tối:  Min = Max = 2 người (sau này có thể mở rộng thêm 1 người)
             * Trong mỗi ca, nếu số lượng người rãnh > Max, thì ta phải xét đến thời gian 
             * làm việc của nhân viên:
             *  - Đối với ca sáng: phải có ít nhất 3 người có TGLV >= 2 tháng
             *  - Đối với ca chiều: 2 người (phòng trường hợp 1 người đi giao hàng thì tại quán vẫn còn lại 1 người) 
             *  - Đối với ca tối: 1 người
            **/


            var lPCLV = new List<List<PhanCongLamViec>>();// Lưu lịch phân công trong tuần

            //DataTable dt = new DataTable();

            dgvLich.Rows.Clear();

            dgvLich.ColumnCount = 7;
            dgvLich.Columns[0].Name = "Thứ 2";
            dgvLich.Columns[1].Name = "Thứ 3";
            dgvLich.Columns[2].Name = "Thứ 4";
            dgvLich.Columns[3].Name = "Thứ 5";
            dgvLich.Columns[4].Name = "Thứ 6";
            dgvLich.Columns[5].Name = "Thứ 7";
            dgvLich.Columns[6].Name = "Chủ nhật";
            string[] row = new string[] { "", "", "", "", "", "", "" };
            for (i = 0; i < 12; i++)
            {
                dgvLich.Rows.Add(row);
                for (j = 0; j < 7; j++)
                {

                    dgvLich.Rows[i].Cells[j].Value = "";
                }
            }
            for (i = 0; i < 7; i++)// Duyệt qua 7 ngày trong tuần
            {
                List<PhanCongLamViec> pclv = new List<PhanCongLamViec>();// Lưu lịch làm việc trong 1 ngày

                PhanCongLamViec pc; // Lưu thông tin PhanCongLamViec

                // Chọn ra tất cả nhân viên mỗi ca trong 1 ngày
                List<ChiTietNhanVien> caSang = lCTNVTheoNgay.ElementAt(i).FindAll(x => x.CaLV == 1);
                List<ChiTietNhanVien> caChieu = lCTNVTheoNgay.ElementAt(i).FindAll(x => x.CaLV == 2);
                List<ChiTietNhanVien> caToi = lCTNVTheoNgay.ElementAt(i).FindAll(x => x.CaLV == 3);
                string temp = "";
                string mau = "";
                // Kiểm tra số lượng nhân viên mỗi ca và tiến hành đưa dữ liệu vào lPCLV
                if (caSang.Count() <= 5)
                {
                    for (j = 0; j < caSang.Count(); j++)
                    {
                        pc = new PhanCongLamViec();
                        pc.MaNV = caSang.ElementAt(j).Id;
                        pc.MaCaLV = caSang.ElementAt(j).CaLV;
                        pc.ThuLamViec = caSang.ElementAt(j).ThuLV;
                        pclv.Add(pc);
                        mau = (from cts in db.ChiTietNhanViens
                               where cts.Id == caSang.ElementAt(j).Id
                               select cts.Mau).First();
                        temp = (from nvs in db.NhanViens
                                where nvs.Id == caSang.ElementAt(j).Id
                                select nvs.HoTen).First();
                        dgvLich.Rows[j].Cells[i].Value = temp;
                        dgvLich.Rows[j].Cells[i].Style.BackColor = ColorTranslator.FromHtml(mau);
                    }
                    if (j <= 5)
                    {
                        for (int jj = j; jj < 6; jj++)
                            dgvLich.Rows[jj].Cells[i].Value = "";
                    }
                }
                else
                {

                }
                if (caChieu.Count() <= 3)
                {
                    for (j = 0; j < caChieu.Count(); j++)
                    {
                        pc = new PhanCongLamViec();
                        pc.MaNV = caChieu.ElementAt(j).Id;
                        pc.MaCaLV = caChieu.ElementAt(j).CaLV;
                        pc.ThuLamViec = caChieu.ElementAt(j).ThuLV;
                        pclv.Add(pc);
                        mau = (from cts in db.ChiTietNhanViens
                               where cts.Id == caChieu.ElementAt(j).Id
                               select cts.Mau).First();
                        temp = (from nvs in db.NhanViens
                                where nvs.Id == caChieu.ElementAt(j).Id
                                select nvs.HoTen).First();
                        dgvLich.Rows[j + 6].Cells[i].Value = temp;
                        dgvLich.Rows[j + 6].Cells[i].Style.BackColor = ColorTranslator.FromHtml(mau);
                    }
                    if (j <= 3)
                    {
                        for (int jj = j; jj < 4; jj++)
                            dgvLich.Rows[jj + 6].Cells[i].Value = "";
                    }
                }
                else
                {

                }
                if (caToi.Count() <= 2)
                {
                    for (j = 0; j < caToi.Count(); j++)
                    {
                        pc = new PhanCongLamViec();
                        pc.MaNV = caToi.ElementAt(j).Id;
                        pc.MaCaLV = caToi.ElementAt(j).CaLV;
                        pc.ThuLamViec = caToi.ElementAt(j).ThuLV;
                        pclv.Add(pc);
                        mau = (from cts in db.ChiTietNhanViens
                               where cts.Id == caToi.ElementAt(j).Id
                               select cts.Mau).First();
                        temp = (from nvs in db.NhanViens
                                where nvs.Id == caToi.ElementAt(j).Id
                                select nvs.HoTen).First();
                        dgvLich.Rows[j + 10].Cells[i].Value = temp;
                        dgvLich.Rows[j + 10].Cells[i].Style.BackColor = ColorTranslator.FromHtml(mau);
                    }
                    if (j <= 1)
                    {
                        for (int jj = j; jj < caToi.Count(); jj++)
                            dgvLich.Rows[jj + 10].Cells[i].Value = "";
                    }
                }
                else
                {

                }
                lPCLV.Add(pclv);
                db.PhanCongLamViecs.InsertAllOnSubmit(lPCLV.ElementAt(i));
                db.SubmitChanges();
            }


            //MessageBox.Show("Đã xếp lịch thành công!");
        }

        #endregion

        // ====================================== General methods ===================================
        // Enable/Disable tab ===========================================================
        #region
        private void loadTab(bool bTabNhanVien, bool bTabLichRanh, bool bTabPhanCong)
        {
            tabPhanCong.AttachedControl.Enabled = bTabPhanCong;
            tabLichRanh.AttachedControl.Enabled = bTabLichRanh;
            tabNhanVien.AttachedControl.Enabled = bTabNhanVien;
        }

        // convert color to hex =======================================================
        private static String HexConverter(System.Drawing.Color c)
        {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

        // convert color to RGB ================================================================
        private static String RGBConverter(System.Drawing.Color c)
        {
            return "RGB(" + c.R.ToString() + "," + c.G.ToString() + "," + c.B.ToString() + ")";
        }
        #endregion
        /*=================================================Events =============================================================================*/

        // ====================================== General events ===============================
        #region
        // Xử lý sự kiện khi Form_DanhSachNhanVien load=========================
        private void Form_DanhSachNhanVien_Load(object sender, EventArgs e)
        {
            // load tab NhanVien
            loadTab(true, false, false);

            // disable controls
            disableButtonNhanVien(true, false);
            disableCommonNhanVien(false);

            // get data NhanVien
            nhanViens = db.NhanViens.ToList();
            thoiGianLamViecs = db.ThoiGianLamViecs.ToList();

            var customizeNhanViens = from n in nhanViens // outer sequence
                                     join t in thoiGianLamViecs //inner sequence 
                                     on n.ThoiGianLV equals t.Id // key selector 
                                     select new
                                     { // result selector 
                                         IDHidden = n.Id,
                                         HoTen = n.HoTen,
                                         ThoiGianLV = t.ThoiGian,
                                         ThoiGianLVHidden = n.ThoiGianLV
                                     };

            // datasource -> datagridview
            dgvDSNhanVien.DataSource = customizeNhanViens.ToList();
            dgvDSNhanVien.Columns["IDHidden"].Visible = false;
            dgvDSNhanVien.Columns["ThoiGianLVHidden"].Visible = false;

            // get data ThoiGianLamViec
            thoiGianLamViecs = db.ThoiGianLamViecs.ToList();

            // datasource -> combobox
            cmbThoiGianLV.DataSource = thoiGianLamViecs;
            cmbThoiGianLV.DisplayMember = "ThoiGian";
            cmbThoiGianLV.ValueMember = "Id";

        }
        #endregion
        // ====================================== Tab Nhan Vien ===============================
        #region
        // Xử lý sự kiện khi tabNhanVien được click =====================
        private void tabNhanVien_Click(object sender, EventArgs e)
        {
            Form_DanhSachNhanVien_Load(sender, e);
        }

        // Xử lý sự kiện khi button btnThemNV được click ================
        private void btnThemNV_Click(object sender, EventArgs e)
        {
            them = true;
            disableButtonNhanVien(false, true);
            disableCommonNhanVien(true);
        }

        // Xử lý sự kiện khi button btnLuu được click ===================
        private void btnLuu_Click(object sender, EventArgs e)
        {
            NhanVien nv;

            if (them)
            {
                nv = new NhanVien();
            }
            else
            {
                nv = db.NhanViens.Single(x => x.Id == getIdNV);
            }

            // get data
            nv.HoTen = txtTenNhanVien.Text;
            nv.ThoiGianLV = Convert.ToInt32(cmbThoiGianLV.SelectedValue);

            if (them)
            {
                db.NhanViens.InsertOnSubmit(nv);
                MessageBox.Show("Đã thêm nhân viên mới thành công!");
            }
            else
            {
                MessageBox.Show("Lưu thông tin nhân viên thành công!");
            }

            db.SubmitChanges();

            // recall event
            btnHuy_Click(sender, e);
        }

        // Xử lý sự kiện khi button btnHuy được click =============================
        private void btnHuy_Click(object sender, EventArgs e)
        {
            // recall event
            Form_DanhSachNhanVien_Load(sender, e);

            // set empty
            txtTenNhanVien.Text = "";
            cmbThoiGianLV.SelectedItem = "";
        }

        // Xử lý sự kiện khi button btnXoaNV được click ==========================
        private void btnXoaNV_Click(object sender, EventArgs e)
        {
            // get index current row
            int index = Convert.ToInt32(dgvDSNhanVien.CurrentCell.RowIndex);

            // get id nhân viên đang được chọn trên datagridview
            int getId = Convert.ToInt32(dgvDSNhanVien.Rows[index].Cells[0].Value);

            // Tìm nhân viên đang chọn và tiến hành xóa
            var nv = db.NhanViens.Single(x => x.Id == getId);
            db.NhanViens.DeleteOnSubmit(nv);
            db.SubmitChanges();

            // Thông báo
            MessageBox.Show("Đã xóa nhân viên thành công!");

            // recall event
            tabNhanVien_Click(sender, e);
        }

        // Xử lý sự kiện khi button btnCapNhatNV được click ======================
        private void btnCapNhatNV_Click(object sender, EventArgs e)
        {
            them = false; // nhận dạng xem người dùng đang muốn thêm hay cập nhật

            // disable/enable control
            disableButtonNhanVien(false, true);
            disableCommonNhanVien(true);
        }
        #endregion
        // ====================================== Tab Lich Ranh ===============================
        // get color and set background combox equals color selected ==========
        #region
        private void cmbMauNV_Click(object sender, EventArgs e)
        {
            DialogResult dr = colorDialog1.ShowDialog();
            if (dr == DialogResult.OK)
                cmbMauNV.BackColor = colorDialog1.Color;
        }

        // Lưu lịch rãnh và màu của nhân viên đang chọn =======================
        private void btnLuuLichRanh_Click(object sender, EventArgs e)
        {
            List<ChiTietNhanVien> ctnv = new List<ChiTietNhanVien>();
            ChiTietNhanVien ct;
            int i;

            // Kiểm tra nếu nhân viên đang chọn đã tồn tại lịch rãnh và màu
            if (them == false)
            {
                // Lọc ra tất cả các chi tiết nhân viên có id = id của nhân viên trong comboxbox đang chọn
                var vlCTNV = db.ChiTietNhanViens.Where(x => x.Id == Convert.ToInt32(cmbChonTenNV.SelectedValue)).ToList();

                List<ChiTietNhanVien> lCTNV = new List<ChiTietNhanVien>();
                for (int k = 0; k < vlCTNV.Count; k++)
                {
                    lCTNV.Add(vlCTNV.ElementAt(k));
                }

                // Xóa tất cả các chi tiết nhân viên đã chọn
                db.ChiTietNhanViens.DeleteAllOnSubmit(lCTNV);
                db.SubmitChanges();
            }

            // Lấy dữ liệu và lưu vào list chiTietNhanViens và cập nhật vào csdl
            for (i = 0; i < 7; i++)
            {
                // dgvCaSang
                if (dgvCaSang.Rows[0].Cells[i].Value != null && (bool)dgvCaSang.Rows[0].Cells[i].Value)// CheckBoxCell checked 
                {
                    // thuộc tính của từng chi tiết nhân viên
                    ct = new ChiTietNhanVien();
                    ct.Mau = HexConverter(cmbMauNV.BackColor);
                    ct.ThuLV = (i + 2);
                    ct.CaLV = 1;
                    ct.Id = Convert.ToInt32(cmbChonTenNV.SelectedValue);
                    // thêm vào list chiTietNhanViens
                    ctnv.Add(ct);
                }

                // dgvCaChieu
                if (dgvCaChieu.Rows[0].Cells[i].Value != null && (bool)dgvCaChieu.Rows[0].Cells[i].Value)
                {
                    ct = new ChiTietNhanVien();
                    ct.Mau = HexConverter(cmbMauNV.BackColor);
                    ct.ThuLV = (i + 2);
                    ct.CaLV = 2;
                    ct.Id = Convert.ToInt32(cmbChonTenNV.SelectedValue);
                    ctnv.Add(ct);
                }

                // dgvCaToi
                if (dgvCaToi.Rows[0].Cells[i].Value != null && (bool)dgvCaToi.Rows[0].Cells[i].Value)
                {
                    ct = new ChiTietNhanVien();
                    ct.Mau = HexConverter(cmbMauNV.BackColor);
                    ct.ThuLV = (i + 2);
                    ct.CaLV = 3;
                    ct.Id = Convert.ToInt32(cmbChonTenNV.SelectedValue);
                    ctnv.Add(ct);
                }
            }
            // Thêm vào ChiTietNhanViens
            db.ChiTietNhanViens.InsertAllOnSubmit(ctnv);

            // Lưu thay đổi
            db.SubmitChanges();

            // Thông báo và uncheked tất cả CheckBoxCell
            MessageBox.Show("Đã lưu lịch rãnh của nhân viên thành công!");
            clearOrSelectAllColumnsChecked(false);

            // reload tab lịch rãnh
            tabLichRanh_Click(sender, e);
        }

        // Xử lý sự kiện khi button btnHuyLichRanh được click ==========================
        private void btnHuyLichRanh_Click(object sender, EventArgs e)
        {
            tabLichRanh_Click(sender, e);
        }

        // Xử lý sự kiện khi giá trị trong cmbChonTenNV thay đổi ======================
        private void cmbChonTenNV_SelectedIndexChanged(object sender, EventArgs e)
        {
            clearOrSelectAllColumnsChecked(false);
            loadDataTabChonLichRanh();
        }

        // Xử lý sự kiện khi chkBoxChonTatCaCaSang check/uncheck =======================
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

        // Xử lý sự kiện khi chkBoxChonTatCaCaSang check/uncheck =======================
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

        // Xử lý sự kiện khi chkBoxChonTatCaCaSang check/uncheck =====================
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


        // Xử lý sự kiện khi tabLichRanh được click ====================================
        private void tabLichRanh_Click(object sender, EventArgs e)
        {
            loadTab(false, true, false);
            loadDataTabChonLichRanh();
            disableButtonChonLichRanh(false);
            disableCommonChonLichRanh(true);
            disableDgvChonLichRanh(false);
            disableCheckBox(false);
            checkOrUncheckCheckBox(false);
        }

        // Xử lý sự kiện khi button btnCapNhatLichRanh được click ====================
        private void btnCapNhatLichRanh_Click(object sender, EventArgs e)
        {
            // disable/enable control tabChonLichRanh
            disableButtonChonLichRanh(true);
            disableCommonChonLichRanh(false);
            disableDgvChonLichRanh(true);
            disableCheckBox(true);
        }

        // Xử lý sự kiện khi cell trong dgvDSNhanVien được click =======================
        private void dgvDSNhanVien_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // get index current row
            int index = Convert.ToInt32(dgvDSNhanVien.CurrentCell.RowIndex);

            // get id nhân viên
            getIdNV = Convert.ToInt32(dgvDSNhanVien.Rows[index].Cells["IDHidden"].Value);

            // get id thời gian làm việc
            getIdTGLV = Convert.ToInt32(dgvDSNhanVien.Rows[index].Cells["ThoiGianLVHidden"].Value);

            // gán giá trị
            txtTenNhanVien.Text = dgvDSNhanVien.Rows[index].Cells[1].Value.ToString();

            // tham chiếu đến bảng ThoiGianLamViec để lấy ra thời gian làm việc
            ThoiGianLamViec tglv = db.ThoiGianLamViecs.Single(x => x.Id == getIdTGLV);

            // set text cho combobox
            cmbThoiGianLV.Text = tglv.ThoiGian.ToString();
        }

        // Xử lý sự kiện khi nội dung cell trong dgvDSNhanVien được click =======================
        private void dgvDSNhanVien_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // recall event
            dgvDSNhanVien_CellClick(sender, e);
        }
        #endregion
        // =================================== Tab Phan Cong ====================================
        // Xử lý sự kiện khi tabPhanCong được click
        #region
        private void tabPhanCong_Click(object sender, EventArgs e)
        {
            loadTab(false, false, true);
        }



        private void btnTienHanhXepLich_Click(object sender, EventArgs e)
        {
            timer1.Enabled = true;
            timer1.Start();


        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            //pgrBarPhanCong.Increment(10);
            if (pgrBarPhanCong.Value == 50)
                xepLich();
            if (pgrBarPhanCong.Value >= 100)
            {

                timer1.Enabled = false;
                timer1.Stop();
                pgrBarPhanCong.Value = 0;
                MessageBox.Show("Xếp lịch hoàn thành!");
                
            }
            else
                pgrBarPhanCong.Value += 25;
        }

        private void btnHuyXepLich_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            timer1.Stop();
            pgrBarPhanCong.Value = 0;
        }
        #endregion

        private void buttonX1_Click(object sender, EventArgs e)
        {
            OfficeExel.Application Excel = new OfficeExel.Application();

            OfficeExel.Workbook wb = Excel.Workbooks.Add(OfficeExel.XlSheetType.xlWorksheet);

            OfficeExel.Worksheet ws = (OfficeExel.Worksheet)Excel.ActiveSheet;

            

            //textBox2.Text = DateTimeExtensions.LastDayOfWeek(DateTimeExtensions.LastDayOfWeek(dtToday).AddDays(1)).ToShortDateString();
            Excel.Visible = true;

            //DateTime dtToday = new DateTime();
            
            

            int getDay = (int)DateTimeExtensions.FirstDayOfMonth(DateTime.Now).DayOfWeek;
            switch(getDay){
                case 0:{// monday
                    getDay = 0;
                    break;
                }
                case 1:{// tuesday
                    getDay=7-1;
                    break;                
                }
                case 2:{// wednessday
                    getDay=7-2;
                    break;
                }
                case 3:{// thursday
                    getDay=7-3;
                    break;
                }
                case 4:{// friday
                    getDay=7-4;
                    break;
                }
                case 5:{// saturday
                    getDay=7-5;
                    break;
                }
                case 6:{// sunday
                    getDay=7-6;
                    break;
                }
            }


            ws.Cells[1, 1] = "LỊCH LÀM VIỆC X - COFFEE TUẦN " + ((Convert.ToInt32(DateTime.Now.Day) + getDay) / 7) + " THÁNG " + DateTime.Now.Month + "." + DateTime.Now.Year;
            ws.Range[ws.Cells[1,1],ws.Cells[1,10]].Merge();
            ws.Cells[1, 1].Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Violet);
            ws.Cells[1, 1].Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White);
            ws.Range[ws.Cells[4, 1], ws.Cells[8, 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.AliceBlue);
            ws.Range[ws.Cells[10, 1], ws.Cells[12, 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.AliceBlue);
            ws.Range[ws.Cells[14, 1], ws.Cells[15, 3]].Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.AliceBlue);
            

            ws.Cells[2, 1] = "Ca";
            ws.Cells[2, 2] = "Thời gian";
            ws.Cells[2, 3] = "Vị trí";
            ws.Cells[2, 4] = "Thứ 2";
            ws.Cells[2, 5] = "Thứ 3";
            ws.Cells[2, 6] = "Thứ 4";
            ws.Cells[2, 7] = "Thứ 5";
            ws.Cells[2, 8] = "Thứ 6";
            ws.Cells[2, 9] = "Thứ 7";
            ws.Cells[2, 10] = "Chủ nhật";

            ws.Cells[4, 1] = "Sáng";
            ws.Cells[10, 1] = "Chiều";
            ws.Cells[14, 1] = "Tối";

            ws.Cells[3, 4] = DateTimeExtensions.LastDayOfWeek(DateTime.Now).AddDays(1).ToShortDateString();
            ws.Cells[3, 5] = DateTimeExtensions.LastDayOfWeek(DateTime.Now).AddDays(2).ToShortDateString();
            ws.Cells[3, 6] = DateTimeExtensions.LastDayOfWeek(DateTime.Now).AddDays(3).ToShortDateString();
            ws.Cells[3, 7] = DateTimeExtensions.LastDayOfWeek(DateTime.Now).AddDays(4).ToShortDateString();
            ws.Cells[3, 8] = DateTimeExtensions.LastDayOfWeek(DateTime.Now).AddDays(5).ToShortDateString();
            ws.Cells[3, 9] = DateTimeExtensions.LastDayOfWeek(DateTime.Now).AddDays(6).ToShortDateString();
            ws.Cells[3, 10] = DateTimeExtensions.LastDayOfWeek(DateTimeExtensions.LastDayOfWeek(DateTime.Now).AddDays(1)).ToShortDateString();

            ws.Range[ws.Cells[2, 1], ws.Cells[3, 1]].Merge();
            ws.Range[ws.Cells[2, 2], ws.Cells[3, 2]].Merge();
            ws.Range[ws.Cells[2, 3], ws.Cells[3, 3]].Merge();
            ws.Range[ws.Cells[4, 1], ws.Cells[8, 1]].Merge();
            ws.Range[ws.Cells[10, 1], ws.Cells[12, 1]].Merge();
            ws.Range[ws.Cells[14, 1], ws.Cells[15, 1]].Merge();



            for (int i = 0; i < dgvLich.Rows.Count; i++)
            {
                for (int j = 0; j < 7; j++)
                {

                    ws.Cells[i + 4, j + 4] = dgvLich.Rows[i].Cells[j].Value;
                    ws.Cells[i + 4, j + 4].Interior.Color = System.Drawing.ColorTranslator.ToOle(dgvLich.Rows[i].Cells[j].Style.BackColor);
                }
            }

            //ws.Range[ws.Cells[1, 1], ws.Cells[1, 2]].Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Blue);
            ws.Range[ws.Cells[9, 1], ws.Cells[9, 10]].Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.White);
            ws.Range[ws.Cells[13, 1], ws.Cells[13, 10]].Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.White);
            ws.Range[ws.Cells[16, 1], ws.Cells[16, 10]].Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.White);

            Excel.StandardFont = "Times New Roman";
            Excel.StandardFontSize = 14;
            Excel.Cells.Font.Bold = true;
            Excel.Cells.RowHeight = 20;
            Excel.Cells.ColumnWidth = 12.5;

            Excel.Cells.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            Excel.Cells.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
            Excel.Cells.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
        }
    }
}
