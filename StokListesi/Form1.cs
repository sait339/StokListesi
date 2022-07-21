using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace StokListesi
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        DataTable crm = new DataTable();
        DataTable netsis = new DataTable();
        DataTable sonuc = new DataTable();
        ArrayList kilitsonuc = new ArrayList();
        ArrayList fiyatsonuc = new ArrayList();
        private void Form1_Load(object sender, EventArgs e)
        {
            sonuc.Columns.Add("STOK_KODU", typeof(String));
            sonuc.Columns.Add("STOK_ADI", typeof(String));
            sonuc.Columns.Add("CRM_KILIT", typeof(String));
            sonuc.Columns.Add("CRM_FIYAT1", typeof(float));
            sonuc.Columns.Add("CRM_FIYATGRUBU", typeof(String));
            sonuc.Columns.Add("NETSIS_KILIT", typeof(String));
            sonuc.Columns.Add("NETSIS_FIYAT1", typeof(float));
            sonuc.Columns.Add("NETSIS_FIYATGRUBU", typeof(String));
            sonuc.Columns.Add("KILITSONUC", typeof(String));
            sonuc.Columns.Add("FIYATSONUC", typeof(String));
            try
            {
                Connections.MySqlBaglanti.Open();
                MySqlCommand cmd = new MySqlCommand("SELECT *FROM bt_stokkarsilastirma", Connections.MySqlBaglanti);
                MySqlDataAdapter dr = new MySqlDataAdapter(cmd);
                dr.Fill(crm);
                Connections.MySqlBaglanti.Close();


                Connections.SqlBaglanti.Open();
                SqlCommand ccmd = new SqlCommand("SELECT *FROM BT_STOKKARSILASTIRMA", Connections.SqlBaglanti);
                SqlDataAdapter ddr = new SqlDataAdapter(ccmd);
                ddr.Fill(netsis);
                Connections.SqlBaglanti.Close();

                var results = from CRM in crm.AsEnumerable()
                              join NETSIS in netsis.AsEnumerable() on CRM["STOK_KODU"] equals NETSIS["STOK_KODU"]
                              select new
                              {
                                  STOK_KODU = CRM["STOK_KODU"],
                                  STOK_ADI =  CRM["STOK_ADI"],
                                  CRM_KILIT =  CRM["KILIT"],
                                  CRM_FIYAT1 =  CRM["FIYAT1"],
                                  CRM_FIYATGRUBU =  CRM["FIYATGRUBU"],
                                  NETSIS_KILIT = NETSIS["KILIT"],
                                  NETSIS_FIYAT1 =  NETSIS["FIYAT1"],
                                  NETSIS_FIYATGRUBU =  NETSIS["FIYATGRUBU"]
                              };

                foreach (var item in results)
                {
                    Convert.ToDouble(item.CRM_FIYAT1);
                    Convert.ToDouble(item.NETSIS_FIYAT1);
                    if (item.NETSIS_KILIT.Equals(item.CRM_KILIT))
                    {
                        kilitsonuc.Add("DOĞRU");
                    }
                    else
                    {
                        kilitsonuc.Add("YANLIŞ");
                    }
                    if (item.NETSIS_FIYAT1.Equals(item.CRM_FIYAT1))
                    {
                        fiyatsonuc.Add("DOĞRU");
                    }
                    else
                    {
                        fiyatsonuc.Add("YANLIŞ");
                    }
                }
                int i = 0;
                    foreach (var item in results)
                {
                    sonuc.Rows.Add(item.STOK_KODU, item.STOK_ADI, item.CRM_KILIT, Convert.ToDouble(item.CRM_FIYAT1), item.CRM_FIYATGRUBU, item.NETSIS_KILIT, Convert.ToDouble(item.NETSIS_FIYAT1), item.NETSIS_FIYATGRUBU, kilitsonuc[i], fiyatsonuc[i]);
                    i++;
                }
                

                DataView dv = sonuc.DefaultView;
                dv.RowFilter = String.Format("CRM_FIYATGRUBU LIKE '3' AND NETSIS_FIYATGRUBU LIKE '1'");
                dataGridView1.DataSource = dv;

                dataGridView1.DataSource = sonuc;

            }
            catch
            {

            }
            


        }

        public static void Excel_Disa_Aktar(DataGridView dataGridView1)
        {

            SaveFileDialog save = new SaveFileDialog();
            save.OverwritePrompt = false;
            save.Title = "Excel Dosyaları";
            save.DefaultExt = "xlsx";
            save.Filter = "xlsx Dosyaları (*.xlsx)|*.xlsx|Tüm Dosyalar(*.*)|*.*";

            if (save.ShowDialog() == DialogResult.OK)
            {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
                app.Visible = true;
                worksheet = workbook.Sheets["Sayfa1"];
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "Excel Dışa Aktarım";
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }

                }
                workbook.SaveAs(save.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            }

        }



        private void button1_Click(object sender, EventArgs e)
        {
            DataView dv = sonuc.DefaultView;
            dv.RowFilter = String.Format("CRM_FIYATGRUBU LIKE '3' AND NETSIS_FIYATGRUBU LIKE '1' AND KILITSONUC LIKE 'YANLIŞ'");
            dataGridView1.DataSource = dv;
            dataGridView1.DataSource = sonuc;
            Excel_Disa_Aktar(dataGridView1);



        }

        private void button2_Click(object sender, EventArgs e)
        {
            DataView dv = sonuc.DefaultView;
            dv.RowFilter = String.Format("CRM_FIYATGRUBU LIKE '3' AND NETSIS_FIYATGRUBU LIKE '1' AND FIYATSONUC LIKE 'YANLIŞ'");
            dataGridView1.DataSource = dv;
            dataGridView1.DataSource = sonuc;
            Excel_Disa_Aktar(dataGridView1);
        }
    }
}
