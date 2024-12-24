using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics.Eventing.Reader;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;
using System.Runtime.InteropServices;
using Application = System.Windows.Forms.Application;

namespace GorevliKayitFormu
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            KisiListele();
            comboBox1.SelectedItem = null;

        }

        private void btnSaat_Click(object sender, EventArgs e)
        {
            label7.Text = DateTime.Now.ToShortTimeString(); 
        }

        private void button3_Click(object sender, EventArgs e)
        {
            label8.Text = DateTime.Now.ToShortDateString();
        }

        OleDbConnection baglanti;
        OleDbCommand komut;
        OleDbDataAdapter da;

        void KisiListele()
        {
            baglanti = new OleDbConnection("Provider=Microsoft.ACE.OleDb.12.0; Data Source=vtPersonel.accdb");
            baglanti.Open();
            da = new OleDbDataAdapter("SELECT *FROM Personel", baglanti);
            //var Datatable = new System.Data.DataTable();
            var tablo = new System.Data.DataTable();
            da.Fill(tablo);
            dataGridView1.DataSource = tablo;
            baglanti.Close();
        }
        private void btnEkle_Click(object sender, EventArgs e)
        {
            
            string sorgu = "INSERT INTO Personel (AD_SOYAD,TC,KURUM,TELEFON,AFET_GRUBU,EKİPTEKİ_KİŞİ_SAYISI,GİRİŞ_TARİHİ,GİRİŞ_SAATİ) values (@adsoyad,@tc,@kurum,@telefon,@afetgrubu,@ekiptekikisisayisi,@giristarihi,@girissaati)";
            komut = new OleDbCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@adsoyad", txtAdSoyad.Text);
            komut.Parameters.AddWithValue("@tc", txtTC.Text);
            komut.Parameters.AddWithValue("@kurum", txtKurum.Text);
            komut.Parameters.AddWithValue("@telefon", mskTextBox1.Text);
            komut.Parameters.AddWithValue("@afetgrubu", comboBox1.SelectedItem);
            komut.Parameters.AddWithValue("@ekiptekikisisayisi", txtKisiSayisi.Text);
            komut.Parameters.AddWithValue("@giristarihi", label8.Text);
            komut.Parameters.AddWithValue("@girissaati", label7.Text);
            baglanti.Open();
            komut.ExecuteNonQuery();
            baglanti.Close();
            KisiListele();                       
        }

        private void btnTemizle_Click(object sender, EventArgs e)
        {
            txtAdSoyad.Clear();
            txtKurum.Clear();
            mskTextBox1.Clear();
            txtTC.Clear();
            txtKisiSayisi.Clear();
            txtNo.Clear();
            label7.Text = default;
            label8.Text = default;
            comboBox1.SelectedIndex = 0;
        }

        

       /* private void btnSil_Click(object sender, EventArgs e)
        {
            string sorgu = "DELETE FROM Personel WHERE TC=@tc";
            komut = new OleDbCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@tc", dataGridView1.CurrentRow.Cells[1].Value);
            baglanti.Open();
            komut.ExecuteNonQuery();
            baglanti.Close() ;
            KisiListele();
        }*/

        

       /* private void btnGuncelle_Click(object sender, EventArgs e)
        {
            string sorgu = "UPDATE Personel Set AD_SOYAD=@adsoyad,TC=@tc,KURUM=@kurum,TELEFON=@telefon,AFET_GRUBU=@afetgrubu,EKİPTEKİ_KİŞİ_SAYISI=@ekiptekikisisayisi,GİRİŞ_TARİHİ=@giristarihi,GİRİŞ_SAATİ=@girissaati WHERE TC=@tc";
            komut = new OleDbCommand(sorgu, baglanti);
            komut.Parameters.AddWithValue("@adsoyad", txtAdSoyad.Text);
            komut.Parameters.AddWithValue("@tc", txtTC.Text);
            komut.Parameters.AddWithValue("@kurum", txtKurum.Text);
            komut.Parameters.AddWithValue("@telefon", mskTextBox1.Text);
            komut.Parameters.AddWithValue("@afetgrubu", comboBox1.SelectedItem);
            komut.Parameters.AddWithValue("@ekiptekikisisayisi", txtKisiSayisi.Text);
            komut.Parameters.AddWithValue("@giristarihi", label8.Text);
            komut.Parameters.AddWithValue("@girissaati", label7.Text);
            komut.Parameters.AddWithValue("@no", Convert.ToInt32(txtNo.Text));
            baglanti.Open();
            komut.ExecuteNonQuery();
            baglanti.Close();
            KisiListele();




        }*/

        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {

            txtNo.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            txtAdSoyad.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            txtTC.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            txtKurum.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
            mskTextBox1.Text = dataGridView1.CurrentRow.Cells[4].Value.ToString();
            comboBox1.SelectedItem = dataGridView1.CurrentRow.Cells[5].Value.ToString();
            txtKisiSayisi.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            label8.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            label7.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
        }

        
        private void button1_Click(object sender, EventArgs e)
        {
            Excel.Application excel = new Excel.Application();
            excel.Visible = true;
            object Missing = Type.Missing;
            Workbook workbook = excel.Workbooks.Add(Missing);
            Worksheet sheet1 = (Worksheet)workbook.Sheets[1];
            int StartCol = 1;
            int StartRow = 1;
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
            {
                Range myRange = (Range)sheet1.Cells[StartRow, StartCol + j];
                myRange.Value2 = dataGridView1.Columns[j].HeaderText;
            }
            StartRow++;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {

                    Range myRange = (Range)sheet1.Cells[StartRow + i, StartCol + j];
                    myRange.Value2 = dataGridView1[j, i].Value == null ? "" : dataGridView1[j, i].Value;
                    myRange.Select();


                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Copyright © 2024 Tüm Hakları Saklıdır.  Bu program Tokat AFAD Görevli Kaydı yapmak için kullanılmaktadır. İçerisindeki bilgiler üçüncü şahıslara verilemez.                                           Bilgi ve tasarım için: ilyas.yesil@afad.gov.tr","Yasal Uyarı!",MessageBoxButtons.OK);
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult secim = new DialogResult();
            secim = MessageBox.Show("Programdan Çıkış Yapılıyor.Emin misiniz?", "Çıkış", MessageBoxButtons.YesNo);

            switch (secim)
            {
                case DialogResult.Yes: Application.Exit();
                    break;
                case DialogResult.No: e.Cancel = true;
                    break;
            }




           /* if (secim == DialogResult.Yes)
            {
                Environment.Exit(0);
            }
            else if (secim == DialogResult.No)
            {
                e.Cancel = true;
            }*/
        }
    }
}
