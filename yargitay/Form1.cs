using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Xml;

using System.Runtime.InteropServices;
using System.Drawing.Printing;
using System.Diagnostics;
namespace yargitay
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        
   
        }
        OleDbConnection conn;//bağlatı için değişken tanımlandı
        OleDbCommand com;// sorgu cümleleri için "     "
        bool test=false;
        private void Form1_Load(object sender, EventArgs e)
        {

            try
            {
                conn = new OleDbConnection("Provider=Microsoft.jet.oledb.4.0;data source= kitapvt.mdb");
                //bağlantı cümlesi access için 

            }
            catch 
            { 
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
           {
                saveFileDialog1.InitialDirectory = Application.ExecutablePath;

                saveFileDialog1.AddExtension = true;
            /// özeti kayıt etmek için tanımlanan saveFileDialog cümleleri
            basla:///aynı isimli özet dosyası mevcutsa tekrar dönülecek yer
                Random r = new Random();
                int ad = (int)r.Next(0, 999999999);///random oluşturulacak dosya ismi için aralık 
                string dosya;
                dosya = ad + ".rtf";//dosya adlı değişkene random oluşturulan isim ve .kyhn son eki ekleniyor
                string kayit_id;
                kayit_id = ad.ToString();//random oluşturulan ad değişkeni kayit_id değişkenine aktarılıyır
                OleDbDataAdapter da;//data adaptör oluşturuluyor
                da = new OleDbDataAdapter("select ozet from Kitap Where ozet='" + dosya + "'  ", conn);//sql cümlesi ile sorgu yapılıyor
                DataSet ds = new DataSet();//dataset tanımlaması yapılıyor
                da.Fill(ds, "Kitap");//kitap isimli tablo ile data set bağlantısı kurulup sql sorgsu çalıştırılıyor
                textBox1.Clear();//temizleme yapılıyor
                textBox1.DataBindings.Add("text", ds, "Kitap.ozet");//tablonun özet alanı text1 e aktarılıyor
                textBox1.DataBindings.Clear();

                if (textBox1.Text == "")//eğer textbox1 boş ise yani yukarıdaki sql sorgusundan değer dönmediyse
                {
                    conn.Close();
                    ozet_box.SaveFile(@dosya, RichTextBoxStreamType.RichText);
                    d_adi.Text = dosya;
                    /// ozet içindeki bilgileri fiziksel olarak /debug/bin klasörü altına .kyhn son eki ile kayıt ediyor
                    OleDbCommand com = new OleDbCommand("insert into kitap(kayit_id,kitapadi,yazar,bsmtrh,yayinevi,ozet) values ('" + kayit_id + "','" + k_adi.Text + "','" + yzr.Text + "','" + bsm.Text + "','" + y_evi.Text + "','" + d_adi.Text + "')", conn);
                    MessageBox.Show("Kayıt Tamamlandı");
                    temizle();
                    com.Connection.Open();
                    com.ExecuteNonQuery();
                    conn.Close();
                    /// sql komutu ile kitap tablosuna gerekli verileri giriyor.
                }
                else/// eğer en üstteki select sorgusundan değer dönmüşse bu komutlar çalışıyor
                {
                    MessageBox.Show("Aynı İsimli Doya Mevcut. İşlem Tekrarlanıyor LütfenBekleyin...");
                    goto basla;/// yukarıdaki başla isimli yer tutucuya gidiyor.

                }
            }
           catch { }
                }
        private void temizle()
        {
            k_adi.Clear();
            yzr.Clear();
            bsm.Clear();
            y_evi.Clear();
            ozet_box.Clear();
            g_bsm.Clear();
            g_k_adi.Clear();
            g_yayin.Clear();
            g_yzr.Clear();
            s_bsm.Clear();      /// tüm textboxları temizlemek için bir fonksiyon.
            s_kadi.Clear();
            s_yayin.Clear();
            s_yzr.Clear();
            textBox4.Clear();
            textBox2.Clear();
            Kitap_ara.Clear();
            Yazar_Ara.Clear();
            yayınevi_ara.Clear();
        }

   
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (radioButton1.Checked == false && radioButton2.Checked == false && radioButton3.Checked == false)
                {
                    MessageBox.Show("Lütfen Arama Türünü Seçiniz");
                }
                /// eğer 3 radiobutonda seçili değilse mesaj veriyor ve bizden bir radiobutonu seçmemizi istiyor
                else /// eğer herhangi biri seçili ise alttaki iflere göre işlem yapıyo.
                {
                    if (radioButton1.Checked == true)// 
                    {

                        conn.Open();
                        DataSet dtst = new DataSet();
                        //string Sql = "Select * from MUSTERI where MUSTAN like '" + Arama + "%'";
                        OleDbDataAdapter adtr = new OleDbDataAdapter("select kayit_id,kitapadi,yazar,bsmtrh,yayinevi From Kitap where kitapadi like'%" + Kitap_ara.Text + "%'", conn);

                        adtr.Fill(dtst, "Kitap");

                        DG1.DataSource = dtst.Tables["Kitap"];

                        adtr.Dispose();

                        conn.Close();
                        /// radio1 seçiliyse kitap ismine göre arama yapıyor ve datagridwieve aktarıyor

                    }
                    if (radioButton2.Checked == true)
                    {

                        conn.Open();
                        DataSet dtst1 = new DataSet();
                        OleDbDataAdapter adtr1 = new OleDbDataAdapter("select kayit_id,kitapadi,yazar,bsmtrh,yayinevi From Kitap where yazar='" + Yazar_Ara.Text + "'", conn);
                        adtr1.Fill(dtst1, "Kitap");
                        DG1.DataSource = dtst1.Tables["Kitap"];
                        adtr1.Dispose();
                        conn.Close();
                        /// radio2 seçiliyse yazar ismine göre arama yapıyor ve datagridwieve aktarıyor
                    }
                    if (radioButton3.Checked == true)
                    {

                        conn.Open();
                        DataSet dtst2 = new DataSet();
                        OleDbDataAdapter adtr2 = new OleDbDataAdapter("select kayit_id,kitapadi,yazar,bsmtrh,yayinevi From Kitap where yayinevi='" + yayınevi_ara.Text + "'", conn);
                        adtr2.Fill(dtst2, "Kitap");
                        DG1.DataSource = dtst2.Tables["Kitap"];
                        adtr2.Dispose();
                        conn.Close();
                        /// radio3 seçiliyse yayın evine göre arama yapıyor ve datagridwieve aktarıyor
                    }

                }
            }
            catch { }

        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
                {
                    Kitap_ara.Enabled = true;
                    Yazar_Ara.Enabled = false;
                    yayınevi_ara.Enabled = false;
            /// radi1 seçiliyse sadece kitap_ara textboxını aktif yapıyor

                }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
                {
                    Yazar_Ara.Enabled = true;
                    Kitap_ara.Enabled = false;
                    yayınevi_ara.Enabled = false;
                    /// radi1 seçiliyse sadece yazar_ara textboxını aktif yapıyor
                }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
                {
                    Kitap_ara.Enabled = false;
                    Yazar_Ara.Enabled=false;
                    yayınevi_ara.Enabled = true;
                    /// radi1 seçiliyse sadece yayınevi_ara textboxını aktif yapıyor
                }

        private void DG1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                conn.Close();
                /// güncelleme için textleri doldurur.
                textBox4.Text = (DG1.Rows[e.RowIndex].Cells[0].Value.ToString());
                g_k_adi.Text = (DG1.Rows[e.RowIndex].Cells[1].Value.ToString());
                g_yzr.Text = (DG1.Rows[e.RowIndex].Cells[2].Value.ToString());
                g_bsm.Text = (DG1.Rows[e.RowIndex].Cells[3].Value.ToString());
                g_yayin.Text = (DG1.Rows[e.RowIndex].Cells[4].Value.ToString());
                /// silme için textleri doldurur.
                s_kadi.Text = (DG1.Rows[e.RowIndex].Cells[1].Value.ToString());
                s_yzr.Text = (DG1.Rows[e.RowIndex].Cells[2].Value.ToString());
                s_bsm.Text = (DG1.Rows[e.RowIndex].Cells[3].Value.ToString());
                s_yayin.Text = (DG1.Rows[e.RowIndex].Cells[4].Value.ToString());
                //datagridview den seçtiğimiz satırın 1. indexinde yazan veriyi text4 e aktarıyor.           
                conn.Open();
                OleDbDataAdapter da1;
                da1 = new OleDbDataAdapter("select * from Kitap Where kayit_id='" + textBox4.Text + "'  ", conn);
                DataSet ds1 = new DataSet();
                da1.Fill(ds1, "Kitap");
                textBox2.Clear();
                textBox2.DataBindings.Add("text", ds1, "Kitap.ozet");
                textBox2.DataBindings.Clear();

                ///sql sorgusu ile belirtilen kayıt_id sine sahip .kyhn uzantılı özet dosyası seçilip text2 ye aktarılıyor

                String file;
                //String tmp;
                file = textBox2.Text;// dosya yolunu text 2 den alıyor

                ozet_box.LoadFile(file);

                conn.Close();
            }
            catch { }    
        }
       
 private void arama_Click(object sender, EventArgs e)
      {
         
              if (test == false) test = true; else test = false;///test isimli bool değer false ise true true ise false yapıyor
          baslangic://işlemi tekrarlamak için kullanılır.
              if (test == true)//eğer test ture ise alttaki komutları çalıştır
              {
                  int sonuc = Occurs(Bul.Text, ozet_box.Text);// sonuç değişkenine Occurs metdoduna veri gönder.
                  if (sonuc == 0) { MessageBox.Show("Metin Bulunamadi"); }//Dönen sonuç 0 ise metin bulunamadı mesajını göster
                  Isaretle(Bul.Text);//işaretle fonksiyonuna bul.text deki veriyi gönder
              }
              else
              {
                  sifirla(textBox5.Text);//sıfırla fonksiyonuna textbox5 deki değeri gönder ve sıfırlayı çalıştır
                  goto baslangic;/// baslangıç isimli yer tutucuya geri dön
              }
       
       }
         private static int Occurs(string search, string exp)
        {
          
                int occurs = 0, current = -1;//int değişkenler tanımlanı 
                do//do while döngüsüne girer
                {
                    current = exp.IndexOf(search, current + 1);/// yukarıdan gelen veri ye metinin index değerleri belirlenir ve current e aktarılır.
                    if (current >= 0)//current 0 dan buyuk ve esitse  occus u 1 arttıracak
                    {
                        occurs++;
                    }
                } while (current >= 0);///current 0 dan büyük olduğu sürece dönmeye devam edecek
                return occurs;/// çağırıldığı yere tekrar dönecek

         
        }
         private void Isaretle(string arananMetin)
         {
             try
             {
                 int index = 0;
                 while (index >= 0)// index 0 dan büyük ve eşit olduğu sürece dönecek bir while döngüsü tanımlanır.
                 {
                     textBox5.Text = Bul.Text;//text5 e aranan kelimeyi aktarır
                     index = ozet_box.Find(arananMetin, index, RichTextBoxFinds.None);/// index e özet_box içinde bulunan kelimelerin indexini atat 
                     ozet_box.SelectionColor = Color.DarkRed;// seçme işlemini yapar bulunan kelimeyi kırmızı yapar

                     if (index >= 0)
                     {
                         index++;/// index değeri 0 dan büyük olduğu sürce indexi 1 arttıtır ve başa döner.
                     }
                 }
             }
             catch { }
         }

         private void sifirla(string arananMetin)
         {
             try
             {
                 int index = 0;
                 while (index >= 0)///işaretle işleminin tam tersini yapar bu işlem ikinci arama için işaretli yerlerin seçimi kaldırır.
                 {
                     index = ozet_box.Find(arananMetin, index, RichTextBoxFinds.None);
                     ozet_box.SelectionColor = Color.Black;///seçim kaldırma işi burada yapılır
                     if (index >= 0)
                     {
                         index++;
                     }
                 }
                 if (test == false) test = true; else test = false;///test ture ise false false ise true yapar
             }
             catch { }
         }

 private void Bul_TextChanged(object sender, EventArgs e)
 {
     int i = 0;
     int a = Bul.Text.Length;
     if (a == 0)
     {
         
     }
     else
     {
         if (i <= a)
         {
             button3.PerformClick();
         }

     }
 }

 private void guncelle_Click(object sender, EventArgs e)
 {
     try
     {
         conn.Close();

         OleDbCommand gncl = new OleDbCommand("update Kitap set kitapadi= '" + g_k_adi.Text + "' where kayit_id = '" + textBox4.Text + "'", conn);
         gncl.Connection.Open();
         gncl.ExecuteNonQuery();
         conn.Close();
         OleDbCommand com1 = new OleDbCommand("update Kitap set yazar= '" + g_yzr.Text + "' where kayit_id = '" + textBox4.Text + "'", conn);
         com1.Connection.Open();
         com1.ExecuteNonQuery();
         conn.Close();
         OleDbCommand com2 = new OleDbCommand("update Kitap set bsmtrh= '" + g_bsm.Text + "' where kayit_id = '" + textBox4.Text + "'", conn);
         com2.Connection.Open();
         com2.ExecuteNonQuery();
         conn.Close();
         OleDbCommand com3 = new OleDbCommand("update Kitap set yayinevi= '" + g_yayin.Text + "' where kayit_id = '" + textBox4.Text + "'", conn);
         com3.Connection.Open();
         com3.ExecuteNonQuery();
         conn.Close();
         string dosya = textBox2.Text;
         ozet_box.SaveFile(@dosya, RichTextBoxStreamType.RichText);

         MessageBox.Show("Güncelleme İşlemi Tamamlandı");
         temizle();
     }
     catch { }
     
 }

 

 private void button7_Click(object sender, EventArgs e)
 {
     colorDialog1.ShowDialog();
  ozet_box.SelectionColor = colorDialog1.Color;
      
 }

 private void panel4_Paint(object sender, PaintEventArgs e)
 {

 }

 private void button4_Click(object sender, EventArgs e)
 {
     try
     {
         ozet_box.SelectionFont = new Font(ozet_box.SelectionFont, FontStyle.Bold);
         ///seçili metni kalın yapar    
     }
     catch { }
 }

 private void button5_Click(object sender, EventArgs e)
 {
     try{
     ozet_box.SelectionFont = new Font(ozet_box.SelectionFont, FontStyle.Underline);
     ///seçili metnin altını çizer    
  }
     catch { }
 }

 private void button6_Click(object sender, EventArgs e)
 {
     try{
     ozet_box.SelectionFont = new Font(ozet_box.SelectionFont, FontStyle.Italic);
     /// seçili metni eğik yapar
     }
     catch { }
     }

 private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
 {try{
     int a;
     textBox3.Text = comboBox1.SelectedItem.ToString();
     a =Convert.ToInt32( textBox3.Text);
     ozet_box.SelectionFont = new Font(ozet_box.Font.FontFamily, a);
     /// cobobox 1 de seçilen sayıyı seçili metnin boyutu olarak ayarlar
 }
 catch { }
 }

 private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
 {try{
     string t;
     t = comboBox2.SelectedItem.ToString();
     int a;
     textBox3.Text = comboBox1.SelectedItem.ToString();
     a = Convert.ToInt32(textBox3.Text);
     ozet_box.SelectionFont = new Font(t, a);
     /// comobox 2 de seçilen yazı tipini seçili metnin yazı tipi yapar
 }
 catch { }
 }

           
 private void sil_btn_Click(object sender, EventArgs e)
 {
     try{
     conn.Close();
     // seçili olan kayıdı siler
     OleDbCommand sil = new OleDbCommand("delete from Kitap where kayit_id= '" + textBox4.Text + "' ", conn);
        sil.Connection.Open();
        sil.ExecuteNonQuery();
        conn.Close();
        string a = textBox2.Text;
        System.IO.File.Delete(a);
        kayit_Getir();
        MessageBox.Show("Silme İşlemi Tamamlandı");
     temizle();//temizle alt progmını çalıştırır.
     
     }
     catch { }
     }

 private void yazdırToolStripMenuItem_Click(object sender, EventArgs e)
 {
     ozet_box.Print();
 }

 private void kapatToolStripMenuItem_Click(object sender, EventArgs e)
 {
     this.Close();
 }

 private void nasılKullanılırToolStripMenuItem_Click(object sender, EventArgs e)
 {
     Form form2 = new Form2();
     form2.Show();
 }

 private void programHakkındaToolStripMenuItem_Click(object sender, EventArgs e)
 {
     MessageBox.Show("Sanal Kütüphane V.1.0 Temmuz 2012");
 }
 private void kayit_Getir()
 {
     conn.Open();
     DataSet dtst = new DataSet();
     //string Sql = "Select * from MUSTERI where MUSTAN like '" + Arama + "%'";
     OleDbDataAdapter adtr = new OleDbDataAdapter("select kayit_id,kitapadi,yazar,bsmtrh,yayinevi From Kitap", conn);

     adtr.Fill(dtst, "Kitap");

     DG1.DataSource = dtst.Tables["Kitap"];

     adtr.Dispose();

     conn.Close();
 }
 private void button8_Click(object sender, EventArgs e)
 {
     kayit_Getir();
 }





 
           
        
    }
}
