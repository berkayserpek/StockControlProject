using ClosedXML.Excel;
using Entities;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StockControlProject
{
    public partial class Form1 : Form
    {
        DataTable importExcel = new DataTable();
        string connectionString = "Server=RABIYAGEZER\\BERKAY;Initial Catalog=Test;User Id=sa;Password=berkay345..";
        private int rowIndex = 0;
        public Form1()
        {
            InitializeComponent();
            InitializeDropDownItems();
        }

        private void InitializeDropDownItems()
        {
            // Dropdown'ı doldur
            FillDropDownItems();
            btnPrint.Visible = false;
            btnImportExcel.Visible = false;
            dataGridView1.Visible = false;
            btnClear.Visible = false;

        }

        private void FillDropDownItems()
        {
            try
            {
                // Veritabanı bağlantısı
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    // Prosedür çağırma
                    SqlCommand stkCommand = new SqlCommand("GetUniqueItems", connection);
                    stkCommand.CommandType = CommandType.StoredProcedure;

                    // Veritabanı bağlantısını açma
                    connection.Open();

                    // Verileri okuma
                    SqlDataReader stkReader = stkCommand.ExecuteReader();

                    List<STK> stkNameList = new List<STK>();
                    List<STK> stkCodeList = new List<STK>();


                    stkCodeList.Add(new STK { ID = 0, MalKodu = "Seçiniz", MalAdi = "" });
                    stkNameList.Add(new STK { ID = 0, MalAdi = "Seçiniz", MalKodu = "" });

                    // Dropdown'u doldurma
                    while (stkReader.Read())
                    {
                        STK stk = new STK();
                        stk.ID = Convert.ToInt32(stkReader["ID"].ToString());
                        stk.MalKodu = stkReader["MalKodu"].ToString();
                        stk.MalAdi = stkReader["MalAdi"].ToString();

                        cBoxItemName.DisplayMember = "MalAdi";
                        cBoxItemName.ValueMember = "MalKodu";

                        cBoxItemCode.DisplayMember = "MalKodu";
                        cBoxItemCode.ValueMember = "MalKodu";

                        cBoxItemCode.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                        cBoxItemCode.AutoCompleteSource = AutoCompleteSource.ListItems;


                        cBoxItemName.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                        cBoxItemName.AutoCompleteSource = AutoCompleteSource.ListItems;


                        stkNameList.Add(stk);
                        stkCodeList.Add(stk);

                    }
                    cBoxItemCode.DataSource = stkCodeList;
                    cBoxItemName.DataSource = stkNameList;

                    // Veritabanı bağlantısını kapatma
                    connection.Close();
                }
            }
            catch (Exception ex)
            {
                // Hata durumunda kullanıcıya bilgi verme
                MessageBox.Show("Bir hata oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnResult_Click(object sender, EventArgs e)
        {
            if (CheckValues())
                GetResult();
        }


        private bool CheckValues()
        {
            string itemCode;
            string itemName;
            DateTime startDate;
            DateTime endDate;
            if (cBoxItemCode.SelectedValue.ToString() != "Seçiniz")
            {
                itemCode = cBoxItemCode.SelectedValue.ToString();
            }
            if (cBoxItemCode.SelectedValue.ToString() != "Seçiniz")
            {
                itemName = cBoxItemName.SelectedValue.ToString();
            }

            if (cBoxItemCode.SelectedIndex == -1)
            {
                cBoxItemCode.SelectedIndex = 0;
            }

            if (cBoxItemName.SelectedIndex == -1)
            {
                cBoxItemName.SelectedIndex = 0;
            }

            if (cBoxItemCode.SelectedValue.ToString() == "Seçiniz" && cBoxItemName.SelectedValue.ToString() == "")
            {
                MessageBox.Show("Lütfen MalKodu veya MalAdı alanlarından bir tanesini seçiniz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cBoxItemCode.SelectedValue = "Seçiniz";
                cBoxItemName.SelectedValue = "";
                return false;

            }

            startDate = dtpStartDate.Value.Date;
            endDate = dtpEndDate.Value.Date;


            if (startDate > endDate)
            {
                MessageBox.Show("Başlangıç tarihi bitiş tarihinden küçük olamaz!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                dtpStartDate.Value = DateTime.Now;
                dtpEndDate.Value = DateTime.Now;
                return false;
            }
            return true;
        }

        private void GetResult()
        {
            try
            {
                // Veritabanı bağlantısını oluştur
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    // Prosedürü çağırma
                    SqlCommand stiCommand = new SqlCommand("GetStockTransactions", connection);
                    stiCommand.CommandType = CommandType.StoredProcedure;

                    // Parametreleri belirleme
                    var startDate = Convert.ToInt32(dtpStartDate.Value.ToOADate());
                    var endDate = Convert.ToInt32(dtpEndDate.Value.ToOADate());

                    // Prosedür için parametreleri ayarlama
                    stiCommand.Parameters.AddWithValue("@BaslangicTarihi", startDate);
                    stiCommand.Parameters.AddWithValue("@BitisTarihi", endDate);
                    if (cBoxItemCode.SelectedValue != "Seçiniz")
                    {
                        stiCommand.Parameters.AddWithValue("@MalKodu", cBoxItemCode.SelectedValue);
                    }
                    else
                    {
                        stiCommand.Parameters.AddWithValue("@MalKodu", cBoxItemName.SelectedValue);
                    }

                    // Veritabanı bağlantısını açma
                    connection.Open();


                    // Verileri okuma
                    SqlDataReader stiReader = stiCommand.ExecuteReader();
                    // Bir DataTable oluştur
                    DataTable dataTable = new DataTable();

                    dataTable.Columns.Add("SıraNo", typeof(int));
                    dataTable.Columns.Add("EvrakNo", typeof(string));
                    dataTable.Columns.Add("IslemTuru", typeof(string));
                    dataTable.Columns.Add("Tarih", typeof(DateTime));
                    dataTable.Columns.Add("GirisMiktar", typeof(decimal));
                    dataTable.Columns.Add("CikisMiktar", typeof(decimal));
                    dataTable.Columns.Add("Stok", typeof(decimal));

                    // Stok miktarını hesaplama
                    decimal stok = 0;
                    while (stiReader.Read())
                    {
                        int siraNo = Convert.ToInt32(stiReader["SiraNo"]);
                        string evrakNo = stiReader["EvrakNo"].ToString();
                        string islemTur = stiReader["IslemTur"].ToString();
                        DateTime tarih = Convert.ToDateTime(stiReader["Tarih"]);
                        decimal girisMiktar = Convert.ToDecimal(stiReader["GirisMiktar"]);
                        decimal cikisMiktar = Convert.ToDecimal(stiReader["CikisMiktar"]);

                        if (islemTur == "Giriş") // Giriş işlemi
                        {
                            stok += girisMiktar;
                            dataTable.Rows.Add(siraNo, evrakNo, islemTur, tarih, girisMiktar, 0, stok);
                        }
                        else // Çıkış işlemi
                        {
                            stok -= cikisMiktar;
                            dataTable.Rows.Add(siraNo, evrakNo, islemTur, tarih, 0, cikisMiktar, stok);
                        }
                    }

                    // SqlDataReader kapat
                    stiReader.Close();

                    if (dataTable.Rows.Count == 0)
                    {
                        MessageBox.Show("Girilen değerlere göre sonuç bulunamadı!", "Uyarı", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        dataTable.Columns.Clear();
                        dataTable.Rows.Clear();
                        dataGridView1.Visible = false;
                        btnPrint.Visible = false;
                        btnImportExcel.Visible = false;
                        btnClear.Visible = false;

                    }

                    if (dataTable.Rows.Count > 0)
                    {
                        dataGridView1.Visible = true;

                        btnImportExcel.Visible = true;
                        btnPrint.Visible = true;
                        btnClear.Visible = true;
                        importExcel = dataTable;

                        // DataGridView'e DataTable'ı atama
                        dataGridView1.DataSource = dataTable;
                    }

                    // Veritabanı bağlantısını kapatma
                    connection.Close();
                }


            }
            catch (Exception ex)
            {
                // Hata durumunda kullanıcıya bilgi verme
                MessageBox.Show("Bir hata oluştu", "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            dtpStartDate.Value = DateTime.Now;
            dtpEndDate.Value = DateTime.Now;
            cBoxItemCode.SelectedIndex = 0;
            cBoxItemName.SelectedIndex = 0;
            btnImportExcel.Visible = false;
            btnPrint.Visible = false;
            dataGridView1.Visible = false;
            btnClear.Visible = false;
        }

        private void btnImportExcel_Click(object sender, EventArgs e)
        {
            using (var workbook = new XLWorkbook())
            {
                // Yeni bir çalışma sayfası oluştur
                var worksheet = workbook.Worksheets.Add("Veriler");

                // DataTable'dan gelen verileri excel sayfasına aktar
                worksheet.Cell(1, 1).InsertTable(importExcel);

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel Dosyaları|*.xlsx|Tüm Dosyalar|*.*";
                saveFileDialog.Title = "Excel Dosyasını Kaydet";
                saveFileDialog.ShowDialog();

                if (saveFileDialog.FileName != "")
                {
                    workbook.SaveAs(saveFileDialog.FileName);
                }
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            PrintDocument printDocument = new PrintDocument();
            printDocument.PrintPage += PrintDocument_PrintPage;
            printDocument.DefaultPageSettings.Landscape = true;

            PrintPreviewDialog printPreviewDialog = new PrintPreviewDialog();
            printPreviewDialog.Document = printDocument;

            printPreviewDialog.ShowDialog();
        }

        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            Graphics g = e.Graphics;
            Font font = new Font("Arial", 12);
            Brush brush = Brushes.Black;

            float yPos = 0;
            float leftMargin = e.MarginBounds.Left;
            float topMargin = e.MarginBounds.Top;
            float columnWidth = e.MarginBounds.Width / importExcel.Columns.Count;


            //başlıkları atamak istedim fakat tam verimli olmadı.
            //foreach (DataColumn column in importExcel.Columns)
            //{
            //    g.DrawString(column.ColumnName, font, brush, leftMargin, topMargin + yPos);
            //    leftMargin += columnWidth;
            //}
            //yPos += font.GetHeight();

            while (rowIndex < importExcel.Rows.Count)
            {
                DataRow row = importExcel.Rows[rowIndex];
                for (int j = 0; j < importExcel.Columns.Count; j++)
                {
                    g.DrawString(row[j].ToString(), font, brush, leftMargin, topMargin + yPos);
                    leftMargin += columnWidth;
                }
                yPos += font.GetHeight();
                leftMargin = e.MarginBounds.Left;
                rowIndex++;

                if (yPos + font.GetHeight() > e.MarginBounds.Bottom)
                {
                    e.HasMorePages = true;
                    return;
                }
            }

            e.HasMorePages = false;
            rowIndex = 0;
        }

        private void btnResult_Click_1(object sender, EventArgs e)
        {

        }
    }
}
