/*
===============================================================================
QLNSData.cs - CÁC CLASS DÙNG CHUNG CHO PROJECT QuanLyNhanSu
===============================================================================

MỤC ĐÍCH
- Gom toàn bộ class xử lý dữ liệu vào một file để dễ tìm và sửa.
- File này KHÔNG tự động kết nối các class với các Form.
- Mỗi Form gọi class cần dùng trong các sự kiện Load, Click, ... của Form.
- Form chịu trách nhiệm giao diện; class DAL chịu trách nhiệm đọc/ghi SQL Server.

FORM NÀO DÙNG CLASS NÀO?

1. FormMain
   - Chỉ mở các Form con; thường không cần truy vấn dữ liệu trực tiếp.

2. FormDanhMuc
   - Dùng: DanhMucDAL. Nạp dữ liệu: GetAll("PhongBan").
   - Thêm/sửa/xóa: Insert(...), Update(...), Delete(...).
   - Tên danh mục hợp lệ được liệt kê trong phần DanhMucDAL bên dưới.

3. FormNhanVien
   - Dùng: NhanVienDAL và NhanVienModel.
   - Nạp bảng: GetAll(); xem một hồ sơ: GetByMaNV(maNV).
   - Thêm/sửa: tạo NhanVienModel từ các control, rồi gọi Insert(nv)/Update(nv).
   - Xóa: Delete(maNV).

4. FormQTDT
   - Dùng: QTDTDAL và QTDTModel.
   - Nạp bảng: GetAll(); thêm/sửa/xóa: Insert(q), Update(q), Delete(maQTDT).

5. FormQTCT
   - Dùng: QTCTDAL và QTCTModel.
   - Nạp bảng: GetAll(); thêm/sửa/xóa: Insert(q), Update(q), Delete(maQTCT).

6. FormQuaTrinhNV
   - Dùng: QuaTrinhNhanVienDAL.
   - Khi chọn nhân viên, nạp hai bảng riêng:
       dgvQTDT.DataSource = QuaTrinhNhanVienDAL.GetQTDT(maNV);
       dgvQTCT.DataSource = QuaTrinhNhanVienDAL.GetQTCT(maNV);

7. FormTimKiemNV
   - Dùng: NhanVienDAL.TimKiem(tenNV, maPhong, maCV, maHV, maNN).
   - Tiêu chí không dùng thì truyền null; có thể kết hợp nhiều tiêu chí.

8. FormTimKiemQTDT
   - Dùng: QTDTDAL.TimKiem(maNV, maTruong, maHT).
   - Tiêu chí không dùng thì truyền null.

9. FormBaoCaoNhanVien
   - Dùng: BaoCaoDAL.BaoCaoNhanVienTheoNgoaiNguChuyenMon(maNN, maCM).

10. FormBaoCaoQTCT
    - Dùng: BaoCaoDAL.BaoCaoQTCT(maNV, maPhong).

11. FormBaoCaoQTDT
    - Dùng: BaoCaoDAL.BaoCaoQTDT(maNV, maTruong, maNganh).

12. FormBaoCaoQTDT_ChuaKetThuc
    - Dùng: BaoCaoDAL.BaoCaoQTDTChuaKetThuc().

VÍ DỤ GỌI CLASS TỪ FORM

A. Nạp dữ liệu khi Form mở (đổi tên DataGridView cho đúng với Form của bạn):
    private void FormNhanVien_Load(object sender, EventArgs e)
    {
        dgvNhanVien.DataSource = NhanVienDAL.GetAll();
    }

B. Tìm kiếm nhân viên (SelectedValue chỉ dùng khi ComboBox đã được nạp dữ liệu):
    dgvNhanVien.DataSource = NhanVienDAL.TimKiem(
        txtTimTen.Text,
        cboPhong.SelectedValue?.ToString(),
        cboChucVu.SelectedValue?.ToString(),
        cboHocVan.SelectedValue?.ToString(),
        cboNgoaiNgu.SelectedValue?.ToString());

C. Thêm nhân viên: tạo object chứa dữ liệu rồi gọi DAL:
    NhanVienModel nv = new NhanVienModel
    {
        MaNV = txtMaNV.Text.Trim(),
        HoTen = txtHoTen.Text.Trim(),
        MaPhong = cboPhong.SelectedValue?.ToString(),
        MaCV = cboChucVu.SelectedValue?.ToString(),
        MaHV = cboHocVan.SelectedValue?.ToString(),
        MaCM = cboChuyenMon.SelectedValue?.ToString()
    };
    NhanVienDAL.Insert(nv);
    dgvNhanVien.DataSource = NhanVienDAL.GetAll();

D. Nạp danh sách vào ComboBox:
    cboPhong.DataSource = DanhMucDAL.GetAll("PhongBan");
    cboPhong.DisplayMember = "TenPhong";
    cboPhong.ValueMember = "MaPhong";

LƯU Ý
- Sửa Db.ConnectionString nếu SQL Server trên máy bạn không phải .\SQLEXPRESS.
- Project cần có package Microsoft.Data.SqlClient.
- Ví dụ chỉ minh họa cách gọi class; tên control phải đổi theo Form thực tế.
- Insert/Update nhân viên lưu cả danh sách ngoại ngữ trong NhanVienModel.MaNgoaiNgu.
- Xóa nhân viên có QTDT/QTCT liên quan có thể bị SQL Server từ chối do khóa ngoại.
===============================================================================
*/

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Microsoft.Data.SqlClient;

namespace QuanLyNhanSu
{
    // ============================================================
    // 1. DÙNG CHUNG CHO TOÀN BỘ ỨNG DỤNG: kết nối SQL Server
    // Form không cần tạo kết nối riêng; các DAL sử dụng Db.
    // ============================================================
    public static class Db
    {
        // Chuỗi kết nối dùng chung cho toàn bộ ứng dụng. Chỉ sửa tại đây khi cần.
        public static string ConnectionString { get; set; } =
            @"Server=.\SQLEXPRESS;Database=QLNS_QTDT_QTCT;Integrated Security=True;TrustServerCertificate=True;";

        // Tạo tham số SQL an toàn. Không nối trực tiếp giá trị người dùng vào câu SQL.
        public static SqlParameter P(string name, object value)
        {
            return new SqlParameter(name, value ?? DBNull.Value);
        }

        // Đọc dữ liệu; kết quả là DataTable để gán cho DataGridView/ComboBox.
        public static DataTable Query(string sql, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
            {
                if (parameters != null && parameters.Length > 0)
                    cmd.Parameters.AddRange(parameters);

                DataTable table = new DataTable();
                adapter.Fill(table);
                return table;
            }
        }

        // Dùng cho INSERT/UPDATE/DELETE; kết quả là số dòng bị ảnh hưởng.
        public static int Execute(string sql, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                if (parameters != null && parameters.Length > 0)
                    cmd.Parameters.AddRange(parameters);

                conn.Open();
                return cmd.ExecuteNonQuery();
            }
        }

        // Dùng khi chỉ cần lấy một giá trị, ví dụ COUNT(*). Các Form hiện tại chưa bắt buộc dùng.
        public static object Scalar(string sql, params SqlParameter[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            using (SqlCommand cmd = new SqlCommand(sql, conn))
            {
                if (parameters != null && parameters.Length > 0)
                    cmd.Parameters.AddRange(parameters);

                conn.Open();
                return cmd.ExecuteScalar();
            }
        }

        // Chuyển chuỗi null/rỗng/chỉ có khoảng trắng thành null để bỏ qua tiêu chí lọc.
        public static string Clean(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value.Trim();
        }
    }

    // ============================================================
    // 2. MODEL: lớp chứa dữ liệu để truyền từ Form sang DAL
    // Không tự truy vấn SQL.
    // ============================================================
    // Dữ liệu một hồ sơ nhân viên; MaNgoaiNgu chứa các mã ngoại ngữ được chọn.
    public class NhanVienModel
    {
        public string MaNV { get; set; } = "";
        public string HoTen { get; set; } = "";
        public string MaGT { get; set; }
        public string MaCV { get; set; }
        public string MaDT { get; set; }
        public string MaTG { get; set; }
        public DateTime? NgaySinh { get; set; }
        public string QueQuan { get; set; }
        public string MaNS { get; set; }
        public string DiaChi { get; set; }
        public string DienThoai { get; set; }
        public string SoCMND { get; set; }
        public string MaPhong { get; set; }
        public string MaHV { get; set; }
        public string MaCM { get; set; }
        public string VoChong { get; set; }
        public string Con { get; set; }
        public List<string> MaNgoaiNgu { get; set; } = new List<string>();
    }

    // Dữ liệu một quá trình đào tạo; dùng cho QTDTDAL.Insert/Update.
    public class QTDTModel
    {
        public string MaQTDT { get; set; } = "";
        public string MaNV { get; set; } = "";
        public DateTime? TuNgay { get; set; }
        public DateTime? DenNgay { get; set; }
        public string MaTruong { get; set; }
        public string MaHT { get; set; }
        public string MaNganh { get; set; }
        public string MaXL { get; set; }
    }

    // Dữ liệu một quá trình công tác; dùng cho QTCTDAL.Insert/Update.
    public class QTCTModel
    {
        public string MaQTCT { get; set; } = "";
        public string MaNV { get; set; } = "";
        public DateTime? TuNgay { get; set; }
        public DateTime? DenNgay { get; set; }
        public string MaPhong { get; set; }
        public string MaCV { get; set; }
    }

    // ============================================================
    // 3. FormDanhMuc -> DanhMucDAL
    // Cũng dùng để nạp ComboBox ở các Form khác.
    // ============================================================
    public static class DanhMucDAL
    {
        private class CauHinhDanhMuc
        {
            public string Bang { get; set; }
            public string Khoa { get; set; }
            public string[] Cot { get; set; }

            public CauHinhDanhMuc(string bang, string khoa, params string[] cot)
            {
                Bang = bang;
                Khoa = khoa;
                Cot = cot;
            }
        }

        // Key là tên danh mục; tên bảng/cột được khai báo sẵn để tránh SQL injection.
        private static readonly Dictionary<string, CauHinhDanhMuc> DanhMuc =
            new Dictionary<string, CauHinhDanhMuc>(StringComparer.OrdinalIgnoreCase)
            {
                { "GioiTinh", new CauHinhDanhMuc("GioiTinh", "MaGT", "TenGT") },
                { "ChucVu", new CauHinhDanhMuc("ChucVu", "MaCV", "TenCV") },
                { "DanToc", new CauHinhDanhMuc("DanToc", "MaDT", "TenDT") },
                { "TonGiao", new CauHinhDanhMuc("TonGiao", "MaTG", "TenTG", "GhiChu") },
                { "NoiSinh", new CauHinhDanhMuc("NoiSinh", "MaNS", "TenNS") },
                { "NgoaiNgu", new CauHinhDanhMuc("NgoaiNgu", "MaNN", "TenNN") },
                { "PhongBan", new CauHinhDanhMuc("PhongBan", "MaPhong", "TenPhong", "DienThoai") },
                { "HocVan", new CauHinhDanhMuc("HocVan", "MaHV", "TenHV") },
                { "ChuyenMon", new CauHinhDanhMuc("ChuyenMon", "MaCM", "TenCM") },
                { "TruongDT", new CauHinhDanhMuc("TruongDT", "MaTruong", "TenTruong", "DiaChi", "DienThoai") },
                { "HinhThuc", new CauHinhDanhMuc("HinhThuc", "MaHT", "TenHT") },
                { "NganhDT", new CauHinhDanhMuc("NganhDT", "MaNganh", "TenNganh") },
                { "XepLoai", new CauHinhDanhMuc("XepLoai", "MaXL", "TenXL") }
            };

        private static CauHinhDanhMuc GetConfig(string loai)
        {
            if (loai == null || !DanhMuc.TryGetValue(loai, out CauHinhDanhMuc config))
                throw new ArgumentException("Tên danh mục không hợp lệ.", nameof(loai));
            return config;
        }

        // Lấy danh sách một danh mục. Ví dụ: GetAll("PhongBan").
        // Dùng nạp DataGridView trong FormDanhMuc hoặc nạp ComboBox ở các Form khác.
        public static DataTable GetAll(string loai)
        {
            CauHinhDanhMuc c = GetConfig(loai);
            return Db.Query("SELECT * FROM [" + c.Bang + "] ORDER BY [" + c.Khoa + "]");
        }

        // Thêm danh mục. values gồm khóa chính và các cột dữ liệu.
        // Ví dụ PhongBan cần các khóa: "MaPhong", "TenPhong", "DienThoai".
        public static int Insert(string loai, Dictionary<string, object> values)
        {
            CauHinhDanhMuc c = GetConfig(loai);
            List<string> cols = new List<string> { c.Khoa };
            cols.AddRange(c.Cot);
            string columnSql = string.Join(", ", cols.ConvertAll(x => "[" + x + "]"));
            string paramSql = string.Join(", ", cols.Select((x, i) => "@p" + i));
            string sql = "INSERT INTO [" + c.Bang + "] (" + columnSql + ") VALUES (" + paramSql + ")";
            List<SqlParameter> pars = new List<SqlParameter>();
            for (int i = 0; i < cols.Count; i++)
                pars.Add(Db.P("@p" + i, values != null && values.ContainsKey(cols[i]) ? values[cols[i]] : null));
            return Db.Execute(sql, pars.ToArray());
        }

        // Sửa thông tin theo khóa chính. values chỉ cần chứa các cột dữ liệu, không cần khóa.
        public static int Update(string loai, string maKhoa, Dictionary<string, object> values)
        {
            CauHinhDanhMuc c = GetConfig(loai);
            string setSql = string.Join(", ", c.Cot.Select((x, i) => "[" + x + "]=@p" + i));
            string sql = "UPDATE [" + c.Bang + "] SET " + setSql + " WHERE [" + c.Khoa + "]=@Khoa";
            List<SqlParameter> pars = new List<SqlParameter>();
            for (int i = 0; i < c.Cot.Length; i++)
                pars.Add(Db.P("@p" + i, values != null && values.ContainsKey(c.Cot[i]) ? values[c.Cot[i]] : null));
            pars.Add(Db.P("@Khoa", maKhoa));
            return Db.Execute(sql, pars.ToArray());
        }

        // Xóa một dòng danh mục. Nếu được bảng khác tham chiếu, khóa ngoại có thể chặn xóa.
        public static int Delete(string loai, string maKhoa)
        {
            CauHinhDanhMuc c = GetConfig(loai);
            return Db.Execute("DELETE FROM [" + c.Bang + "] WHERE [" + c.Khoa + "]=@Khoa",
                Db.P("@Khoa", maKhoa));
        }
    }

    // ============================================================
    // 4. FormNhanVien, FormTimKiemNV -> NhanVienDAL
    // CRUD hồ sơ, nạp danh sách và tìm kiếm nhân viên.
    // ============================================================
    public static class NhanVienDAL
    {
        private const string SelectSql = @"
SELECT nv.*,
       gt.TenGT, cv.TenCV, dt.TenDT, tg.TenTG, ns.TenNS,
       pb.TenPhong, hv.TenHV, cm.TenCM,
       STUFF((
           SELECT ', ' + nn.TenNN
           FROM NV_NgoaiNgu mapNN
           INNER JOIN NgoaiNgu nn ON nn.MaNN = mapNN.MaNN
           WHERE mapNN.MaNV = nv.MaNV
           FOR XML PATH(''), TYPE
       ).value('.', 'varchar(max)'), 1, 2, '') AS NgoaiNgu
FROM HoSoNV nv
LEFT JOIN GioiTinh gt ON gt.MaGT = nv.MaGT
LEFT JOIN ChucVu cv ON cv.MaCV = nv.MaCV
LEFT JOIN DanToc dt ON dt.MaDT = nv.MaDT
LEFT JOIN TonGiao tg ON tg.MaTG = nv.MaTG
LEFT JOIN NoiSinh ns ON ns.MaNS = nv.MaNS
LEFT JOIN PhongBan pb ON pb.MaPhong = nv.MaPhong
LEFT JOIN HocVan hv ON hv.MaHV = nv.MaHV
LEFT JOIN ChuyenMon cm ON cm.MaCM = nv.MaCM ";

        // FormNhanVien: nạp toàn bộ danh sách vào DataGridView.
        // Ví dụ: dgvNhanVien.DataSource = NhanVienDAL.GetAll();
        public static DataTable GetAll()
        {
            return Db.Query(SelectSql + " ORDER BY nv.MaNV");
        }

        // Lấy một hồ sơ theo mã nhân viên, thường dùng khi chọn dòng để xem/sửa.
        public static DataTable GetByMaNV(string maNV)
        {
            return Db.Query(SelectSql + " WHERE nv.MaNV = @MaNV", Db.P("@MaNV", maNV));
        }

        // FormTimKiemNV (hoặc FormNhanVien): tìm theo nhiều tiêu chí kết hợp.
        // Truyền null cho tiêu chí không lọc. Ngoại ngữ được tìm qua bảng NV_NgoaiNgu.
        public static DataTable TimKiem(string tenNV = null, string maPhong = null,
            string maCV = null, string maHV = null, string maNN = null)
        {
            string sql = SelectSql + @"
WHERE (@TenNV IS NULL OR nv.HoTen LIKE '%' + @TenNV + '%')
  AND (@MaPhong IS NULL OR nv.MaPhong = @MaPhong)
  AND (@MaCV IS NULL OR nv.MaCV = @MaCV)
  AND (@MaHV IS NULL OR nv.MaHV = @MaHV)
  AND (@MaNN IS NULL OR EXISTS (
      SELECT 1 FROM NV_NgoaiNgu x WHERE x.MaNV = nv.MaNV AND x.MaNN = @MaNN
  ))
ORDER BY nv.MaNV";

            return Db.Query(sql,
                Db.P("@TenNV", Db.Clean(tenNV)),
                Db.P("@MaPhong", Db.Clean(maPhong)),
                Db.P("@MaCV", Db.Clean(maCV)),
                Db.P("@MaHV", Db.Clean(maHV)),
                Db.P("@MaNN", Db.Clean(maNN)));
        }

        // Thêm hồ sơ mới. NhanVienModel.MaNgoaiNgu sẽ được lưu vào NV_NgoaiNgu.
        // Gọi sau khi đã kiểm tra các trường bắt buộc trên FormNhanVien.
        public static void Insert(NhanVienModel nv)
        {
            using (SqlConnection conn = new SqlConnection(Db.ConnectionString))
            {
                conn.Open();
                using (SqlTransaction tran = conn.BeginTransaction())
                {
                    try
                    {
                        string sql = @"INSERT INTO HoSoNV
(MaNV, HoTen, MaGT, MaCV, MaDT, MaTG, NgaySinh, QueQuan, MaNS, DiaChi,
 DienThoai, SoCMND, MaPhong, MaHV, MaCM, VoChong, Con)
VALUES
(@MaNV, @HoTen, @MaGT, @MaCV, @MaDT, @MaTG, @NgaySinh, @QueQuan, @MaNS, @DiaChi,
 @DienThoai, @SoCMND, @MaPhong, @MaHV, @MaCM, @VoChong, @Con)";

                        using (SqlCommand cmd = new SqlCommand(sql, conn, tran))
                        {
                            AddNhanVienParameters(cmd, nv);
                            cmd.ExecuteNonQuery();
                        }

                        SaveLanguages(conn, tran, nv.MaNV, nv.MaNgoaiNgu, false);
                        tran.Commit();
                    }
                    catch
                    {
                        tran.Rollback();
                        throw;
                    }
                }
            }
        }

        // Cập nhật hồ sơ và thay toàn bộ danh sách ngoại ngữ theo MaNgoaiNgu.
        public static void Update(NhanVienModel nv)
        {
            using (SqlConnection conn = new SqlConnection(Db.ConnectionString))
            {
                conn.Open();
                using (SqlTransaction tran = conn.BeginTransaction())
                {
                    try
                    {
                        string sql = @"UPDATE HoSoNV SET
HoTen=@HoTen, MaGT=@MaGT, MaCV=@MaCV, MaDT=@MaDT, MaTG=@MaTG,
NgaySinh=@NgaySinh, QueQuan=@QueQuan, MaNS=@MaNS, DiaChi=@DiaChi,
DienThoai=@DienThoai, SoCMND=@SoCMND, MaPhong=@MaPhong, MaHV=@MaHV,
MaCM=@MaCM, VoChong=@VoChong, Con=@Con
WHERE MaNV=@MaNV";

                        using (SqlCommand cmd = new SqlCommand(sql, conn, tran))
                        {
                            AddNhanVienParameters(cmd, nv);
                            cmd.ExecuteNonQuery();
                        }

                        SaveLanguages(conn, tran, nv.MaNV, nv.MaNgoaiNgu, true);
                        tran.Commit();
                    }
                    catch
                    {
                        tran.Rollback();
                        throw;
                    }
                }
            }
        }

        // Xóa hồ sơ và các dòng NV_NgoaiNgu.
        // Không tự động xóa QTDT/QTCT. Nếu còn bản ghi liên quan, khóa ngoại sẽ chặn xóa.
        public static void Delete(string maNV)
        {
            using (SqlConnection conn = new SqlConnection(Db.ConnectionString))
            {
                conn.Open();
                using (SqlTransaction tran = conn.BeginTransaction())
                {
                    try
                    {
                        using (SqlCommand cmd = new SqlCommand(
                            "DELETE FROM NV_NgoaiNgu WHERE MaNV=@MaNV", conn, tran))
                        {
                            cmd.Parameters.Add(Db.P("@MaNV", maNV));
                            cmd.ExecuteNonQuery();
                        }
                        using (SqlCommand cmd = new SqlCommand(
                            "DELETE FROM HoSoNV WHERE MaNV=@MaNV", conn, tran))
                        {
                            cmd.Parameters.Add(Db.P("@MaNV", maNV));
                            cmd.ExecuteNonQuery();
                        }
                        tran.Commit();
                    }
                    catch
                    {
                        tran.Rollback();
                        throw;
                    }
                }
            }
        }

        private static void AddNhanVienParameters(SqlCommand cmd, NhanVienModel nv)
        {
            cmd.Parameters.Add(Db.P("@MaNV", nv.MaNV));
            cmd.Parameters.Add(Db.P("@HoTen", nv.HoTen));
            cmd.Parameters.Add(Db.P("@MaGT", Db.Clean(nv.MaGT)));
            cmd.Parameters.Add(Db.P("@MaCV", Db.Clean(nv.MaCV)));
            cmd.Parameters.Add(Db.P("@MaDT", Db.Clean(nv.MaDT)));
            cmd.Parameters.Add(Db.P("@MaTG", Db.Clean(nv.MaTG)));
            cmd.Parameters.Add(Db.P("@NgaySinh", nv.NgaySinh));
            cmd.Parameters.Add(Db.P("@QueQuan", Db.Clean(nv.QueQuan)));
            cmd.Parameters.Add(Db.P("@MaNS", Db.Clean(nv.MaNS)));
            cmd.Parameters.Add(Db.P("@DiaChi", Db.Clean(nv.DiaChi)));
            cmd.Parameters.Add(Db.P("@DienThoai", Db.Clean(nv.DienThoai)));
            cmd.Parameters.Add(Db.P("@SoCMND", Db.Clean(nv.SoCMND)));
            cmd.Parameters.Add(Db.P("@MaPhong", Db.Clean(nv.MaPhong)));
            cmd.Parameters.Add(Db.P("@MaHV", Db.Clean(nv.MaHV)));
            cmd.Parameters.Add(Db.P("@MaCM", Db.Clean(nv.MaCM)));
            cmd.Parameters.Add(Db.P("@VoChong", Db.Clean(nv.VoChong)));
            cmd.Parameters.Add(Db.P("@Con", Db.Clean(nv.Con)));
        }

        private static void SaveLanguages(SqlConnection conn, SqlTransaction tran,
            string maNV, List<string> languages, bool clearOld)
        {
            if (clearOld)
            {
                using (SqlCommand cmd = new SqlCommand(
                    "DELETE FROM NV_NgoaiNgu WHERE MaNV=@MaNV", conn, tran))
                {
                    cmd.Parameters.Add(Db.P("@MaNV", maNV));
                    cmd.ExecuteNonQuery();
                }
            }

            if (languages == null) return;
            foreach (string maNN in languages.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct())
            {
                using (SqlCommand cmd = new SqlCommand(
                    "INSERT INTO NV_NgoaiNgu (MaNV, MaNN) VALUES (@MaNV, @MaNN)", conn, tran))
                {
                    cmd.Parameters.Add(Db.P("@MaNV", maNV));
                    cmd.Parameters.Add(Db.P("@MaNN", maNN));
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }

    // ============================================================
    // 5. FormQTCT -> QTCTDAL; FormQuaTrinhNV cũng dùng để xem lịch sử.
    // CRUD và tìm kiếm quá trình công tác.
    // ============================================================
    public static class QTCTDAL
    {
        private const string SelectSql = @"
SELECT q.*, nv.HoTen, pb.TenPhong, cv.TenCV
FROM QTCT q
LEFT JOIN HoSoNV nv ON nv.MaNV=q.MaNV
LEFT JOIN PhongBan pb ON pb.MaPhong=q.MaPhong
LEFT JOIN ChucVu cv ON cv.MaCV=q.MaCV ";

        // Nạp tất cả QTCT kèm họ tên, tên phòng ban và tên chức vụ.
        public static DataTable GetAll()
        {
            return Db.Query(SelectSql + " ORDER BY q.MaQTCT");
        }

        // Lọc QTCT theo nhân viên và/hoặc phòng ban; null = không lọc tiêu chí đó.
        public static DataTable TimKiem(string maNV = null, string maPhong = null)
        {
            string sql = SelectSql + @"
WHERE (@MaNV IS NULL OR q.MaNV=@MaNV)
  AND (@MaPhong IS NULL OR q.MaPhong=@MaPhong)
ORDER BY q.MaQTCT";
            return Db.Query(sql,
                Db.P("@MaNV", Db.Clean(maNV)),
                Db.P("@MaPhong", Db.Clean(maPhong)));
        }

        // Thêm một QTCT.
        public static int Insert(QTCTModel q)
        {
            return Db.Execute(@"INSERT INTO QTCT
(MaQTCT, MaNV, TuNgay, DenNgay, MaPhong, MaCV)
VALUES (@MaQTCT, @MaNV, @TuNgay, @DenNgay, @MaPhong, @MaCV)",
                Db.P("@MaQTCT", q.MaQTCT), Db.P("@MaNV", q.MaNV),
                Db.P("@TuNgay", q.TuNgay), Db.P("@DenNgay", q.DenNgay),
                Db.P("@MaPhong", Db.Clean(q.MaPhong)), Db.P("@MaCV", Db.Clean(q.MaCV)));
        }

        // Sửa QTCT theo MaQTCT.
        public static int Update(QTCTModel q)
        {
            return Db.Execute(@"UPDATE QTCT SET MaNV=@MaNV, TuNgay=@TuNgay,
DenNgay=@DenNgay, MaPhong=@MaPhong, MaCV=@MaCV WHERE MaQTCT=@MaQTCT",
                Db.P("@MaQTCT", q.MaQTCT), Db.P("@MaNV", q.MaNV),
                Db.P("@TuNgay", q.TuNgay), Db.P("@DenNgay", q.DenNgay),
                Db.P("@MaPhong", Db.Clean(q.MaPhong)), Db.P("@MaCV", Db.Clean(q.MaCV)));
        }

        // Xóa QTCT theo mã quá trình.
        public static int Delete(string maQTCT)
        {
            return Db.Execute("DELETE FROM QTCT WHERE MaQTCT=@MaQTCT", Db.P("@MaQTCT", maQTCT));
        }
    }

    // ============================================================
    // 6. FormQTDT, FormTimKiemQTDT -> QTDTDAL; FormQuaTrinhNV cũng dùng.
    // CRUD và tìm kiếm quá trình đào tạo.
    // ============================================================
    public static class QTDTDAL
    {
        private const string SelectSql = @"
SELECT q.*, nv.HoTen, tr.TenTruong, ht.TenHT, ng.TenNganh, xl.TenXL
FROM QTDT q
LEFT JOIN HoSoNV nv ON nv.MaNV=q.MaNV
LEFT JOIN TruongDT tr ON tr.MaTruong=q.MaTruong
LEFT JOIN HinhThuc ht ON ht.MaHT=q.MaHT
LEFT JOIN NganhDT ng ON ng.MaNganh=q.MaNganh
LEFT JOIN XepLoai xl ON xl.MaXL=q.MaXL ";

        // Nạp tất cả QTDT, kèm tên nhân viên/trường/hình thức/ngành/xếp loại.
        public static DataTable GetAll()
        {
            return Db.Query(SelectSql + " ORDER BY q.MaQTDT");
        }

        // Dùng cho FormTimKiemQTDT và FormQuaTrinhNV.
        // Lọc theo mã nhân viên, trường đào tạo, hình thức đào tạo; null = không lọc.
        public static DataTable TimKiem(string maNV = null, string maTruong = null, string maHT = null)
        {
            string sql = SelectSql + @"
WHERE (@MaNV IS NULL OR q.MaNV=@MaNV)
  AND (@MaTruong IS NULL OR q.MaTruong=@MaTruong)
  AND (@MaHT IS NULL OR q.MaHT=@MaHT)
ORDER BY q.MaQTDT";
            return Db.Query(sql,
                Db.P("@MaNV", Db.Clean(maNV)),
                Db.P("@MaTruong", Db.Clean(maTruong)),
                Db.P("@MaHT", Db.Clean(maHT)));
        }

        // Thêm một QTDT. Giá trị trả về là số dòng được thêm.
        public static int Insert(QTDTModel q)
        {
            return Db.Execute(@"INSERT INTO QTDT
(MaQTDT, MaNV, TuNgay, DenNgay, MaTruong, MaHT, MaNganh, MaXL)
VALUES (@MaQTDT, @MaNV, @TuNgay, @DenNgay, @MaTruong, @MaHT, @MaNganh, @MaXL)",
                Db.P("@MaQTDT", q.MaQTDT), Db.P("@MaNV", q.MaNV),
                Db.P("@TuNgay", q.TuNgay), Db.P("@DenNgay", q.DenNgay),
                Db.P("@MaTruong", Db.Clean(q.MaTruong)), Db.P("@MaHT", Db.Clean(q.MaHT)),
                Db.P("@MaNganh", Db.Clean(q.MaNganh)), Db.P("@MaXL", Db.Clean(q.MaXL)));
        }

        // Sửa QTDT theo MaQTDT. Giá trị trả về là số dòng được sửa.
        public static int Update(QTDTModel q)
        {
            return Db.Execute(@"UPDATE QTDT SET MaNV=@MaNV, TuNgay=@TuNgay, DenNgay=@DenNgay,
MaTruong=@MaTruong, MaHT=@MaHT, MaNganh=@MaNganh, MaXL=@MaXL
WHERE MaQTDT=@MaQTDT",
                Db.P("@MaQTDT", q.MaQTDT), Db.P("@MaNV", q.MaNV),
                Db.P("@TuNgay", q.TuNgay), Db.P("@DenNgay", q.DenNgay),
                Db.P("@MaTruong", Db.Clean(q.MaTruong)), Db.P("@MaHT", Db.Clean(q.MaHT)),
                Db.P("@MaNganh", Db.Clean(q.MaNganh)), Db.P("@MaXL", Db.Clean(q.MaXL)));
        }

        // Xóa QTDT theo mã quá trình.
        public static int Delete(string maQTDT)
        {
            return Db.Execute("DELETE FROM QTDT WHERE MaQTDT=@MaQTDT", Db.P("@MaQTDT", maQTDT));
        }
    }

    // ============================================================
    // 7. FormQuaTrinhNV -> lấy QTDT và QTCT theo một MaNV.
    // ============================================================
    public static class QuaTrinhNhanVienDAL
    {
        // Gán kết quả cho DataGridView QTDT: dgvQTDT.DataSource = GetQTDT(maNV).
        public static DataTable GetQTDT(string maNV)
        {
            return QTDTDAL.TimKiem(maNV, null, null);
        }

        // Gán kết quả cho DataGridView QTCT: dgvQTCT.DataSource = GetQTCT(maNV).
        public static DataTable GetQTCT(string maNV)
        {
            return QTCTDAL.TimKiem(maNV, null);
        }
    }

    // ============================================================
    // 8. Các FormBaoCao -> BaoCaoDAL
    // Các hàm trả về DataTable để gán vào DataGridView.
    // ============================================================
    public static class BaoCaoDAL
    {
        // FormBaoCaoQTDT: báo cáo theo nhân viên, trường đào tạo, ngành đào tạo.
        // Truyền null cho tiêu chí không chọn; có thể lọc một hoặc nhiều tiêu chí.
        public static DataTable BaoCaoQTDT(string maNV = null, string maTruong = null, string maNganh = null)
        {
            string sql = @"
SELECT q.MaQTDT, q.MaNV, nv.HoTen, q.TuNgay, q.DenNgay,
       q.MaTruong, tr.TenTruong, q.MaHT, ht.TenHT,
       q.MaNganh, ng.TenNganh, q.MaXL, xl.TenXL
FROM QTDT q
LEFT JOIN HoSoNV nv ON nv.MaNV=q.MaNV
LEFT JOIN TruongDT tr ON tr.MaTruong=q.MaTruong
LEFT JOIN HinhThuc ht ON ht.MaHT=q.MaHT
LEFT JOIN NganhDT ng ON ng.MaNganh=q.MaNganh
LEFT JOIN XepLoai xl ON xl.MaXL=q.MaXL
WHERE (@MaNV IS NULL OR q.MaNV=@MaNV)
  AND (@MaTruong IS NULL OR q.MaTruong=@MaTruong)
  AND (@MaNganh IS NULL OR q.MaNganh=@MaNganh)
ORDER BY q.MaQTDT";
            return Db.Query(sql,
                Db.P("@MaNV", Db.Clean(maNV)), Db.P("@MaTruong", Db.Clean(maTruong)),
                Db.P("@MaNganh", Db.Clean(maNganh)));
        }

        // FormBaoCaoQTCT: báo cáo theo nhân viên và/hoặc phòng ban; null = bỏ qua tiêu chí.
        public static DataTable BaoCaoQTCT(string maNV = null, string maPhong = null)
        {
            string sql = @"
SELECT q.MaQTCT, q.MaNV, nv.HoTen, q.TuNgay, q.DenNgay,
       q.MaPhong, pb.TenPhong, q.MaCV, cv.TenCV
FROM QTCT q
LEFT JOIN HoSoNV nv ON nv.MaNV=q.MaNV
LEFT JOIN PhongBan pb ON pb.MaPhong=q.MaPhong
LEFT JOIN ChucVu cv ON cv.MaCV=q.MaCV
WHERE (@MaNV IS NULL OR q.MaNV=@MaNV)
  AND (@MaPhong IS NULL OR q.MaPhong=@MaPhong)
ORDER BY q.MaQTCT";
            return Db.Query(sql,
                Db.P("@MaNV", Db.Clean(maNV)), Db.P("@MaPhong", Db.Clean(maPhong)));
        }

        // FormBaoCaoNhanVien: danh sách nhân viên có ĐỒNG THỜI ngoại ngữ và chuyên môn đã chọn.
        public static DataTable BaoCaoNhanVienTheoNgoaiNguChuyenMon(string maNN, string maCM)
        {
            if (string.IsNullOrWhiteSpace(maNN) || string.IsNullOrWhiteSpace(maCM))
                throw new ArgumentException("Hay chon ca ngoại ngữ va chuyen mon.");

            string sql = @"
SELECT nv.MaNV, nv.HoTen, pb.TenPhong, cv.TenCV, hv.TenHV,
       nn.TenNN, cm.TenCM
FROM HoSoNV nv
INNER JOIN NV_NgoaiNgu mapNN ON mapNN.MaNV=nv.MaNV
INNER JOIN NgoaiNgu nn ON nn.MaNN=mapNN.MaNN
LEFT JOIN ChuyenMon cm ON cm.MaCM=nv.MaCM
LEFT JOIN PhongBan pb ON pb.MaPhong=nv.MaPhong
LEFT JOIN ChucVu cv ON cv.MaCV=nv.MaCV
LEFT JOIN HocVan hv ON hv.MaHV=nv.MaHV
WHERE mapNN.MaNN=@MaNN AND nv.MaCM=@MaCM
ORDER BY nv.MaNV";
            return Db.Query(sql, Db.P("@MaNN", maNN), Db.P("@MaCM", maCM));
        }

        // FormBaoCaoQTDT_ChuaKetThuc: QTDT đã bắt đầu và chưa qua ngày DenNgay.
        // Nếu DenNgay = null, coi như chưa có ngày kết thúc. Quy ước này cần phù hợp nghiệp vụ.
        // Không cần tham số; ngày hiện tại được lấy từ SQL Server.
        public static DataTable BaoCaoQTDTChuaKetThuc()
        {
            string sql = @"
SELECT q.MaQTDT, q.MaNV, nv.HoTen, q.TuNgay, q.DenNgay,
       q.MaTruong, tr.TenTruong, q.MaHT, ht.TenHT,
       q.MaNganh, ng.TenNganh, q.MaXL, xl.TenXL
FROM QTDT q
LEFT JOIN HoSoNV nv ON nv.MaNV=q.MaNV
LEFT JOIN TruongDT tr ON tr.MaTruong=q.MaTruong
LEFT JOIN HinhThuc ht ON ht.MaHT=q.MaHT
LEFT JOIN NganhDT ng ON ng.MaNganh=q.MaNganh
LEFT JOIN XepLoai xl ON xl.MaXL=q.MaXL
WHERE q.TuNgay <= CONVERT(date, GETDATE())
  AND (q.DenNgay IS NULL OR q.DenNgay >= CONVERT(date, GETDATE()))
ORDER BY q.DenNgay, q.MaQTDT";
            return Db.Query(sql);
        }
    }
}
