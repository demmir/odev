using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace toplu_ekle
{
    public partial class toplu_kekle : Form

    {
        SqlConnection conn = new SqlConnection("Data Source = GEFORCE\\SQLEXPRESS; Initial Catalog = toplu_ekle; Integrated Security = True");
        public toplu_kekle()

        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Dosya Seç";
            dialog.FileName = textBox1.Text;
            dialog.Filter = "Excel Dosyaları|*.xlsx;*.xlsm;*.xlsb;*.xltx;*.xltm;*.xls;*.xlt;*.xls;*.xml;*.xml;*.xlam;*.xla;*.xlw;*.xlr;";
            dialog.RestoreDirectory = true;
            dialog.CheckFileExists = false;

            if (dialog.ShowDialog() == DialogResult.OK)
                      
            {

                textBox1.Text = dialog.FileName;

                //dosya yolu seçildikten sonrası 

                OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + textBox1.Text + ";Extended Properties=Excel 12.0;");
                //"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                //"Provider=Microsoft.Jet.OleDb.4.0; Data Source = " + textBox1.Text + "; Extended Properties = \"Excel 8.0;HDR=Yes;IMEX=1\";"
                conn.Open();
                OleDbDataAdapter adapt = new OleDbDataAdapter("Select * from[Tablo1]", conn);
                DataSet sd = new DataSet();
                DataTable dt = new DataTable();
                adapt.Fill(dt);
                this.dataGridView1.DataSource = dt.DefaultView;


                

            }



        }

        void FillGrid()
        {
            conn.Open();
            SqlDataAdapter da = new SqlDataAdapter("select * from kitapEkle order by kayıt_ID", conn);
            DataTable dtab = new DataTable();
            da.Fill(dtab);
            dataGridView1.DataSource = dtab;
            conn.Close();



        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void toplu_kekle_Load(object sender, EventArgs e)
        {
            FillGrid();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            conn.Open();
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                SqlCommand cmd = new SqlCommand("Insert into kitapEkle(kitapno,kitapAdi,kitapYazari,baski,basimyili,basimyeri,kitapSayfaSayisi,kitapYayinevi,stok,toplamOkunma)vaues('"+dataGridView1.Rows[i].Cells[2].Value+"')", conn);
                cmd.ExecuteNonQuery();
            }
            conn.Close();
            FillGrid();

        }
    }
}
