using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Threading;


namespace BilgiYarismasi
{
    public partial class Form1 : Form
    {
        string path = "C:\\Users\\berat\\Desktop\\berat\\visualstudio_prpjects\\BilgiYarismasi\\Kitap4.xlsx";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        string gelen_ad, gelen_soyad;
        public Form1(string gelen_ad , string gelen_soyad)
        {
            InitializeComponent();

            this.gelen_ad = gelen_ad;
            this.gelen_soyad = gelen_soyad;

            label3.Text = gelen_ad;
            label4.Text = gelen_soyad;
            label8.Text = string.Empty;

            

        }

        int tiklanma_sayisi = -1;
        int soru_id = 0;
        int verilen_dogru_cevap_sayisi = 0;
        int[] sorulan_sorular = new int[44];
        int kayitli_soru_sayisi = 0;
        int sorulan_soru_sayisi = 0;
        private void Form1_Load(object sender, EventArgs e)
        {
            kayitli_soru_sayisi  = sorulan_sorular.Length - 1;
            timer1.Interval = 1000;
            timer1.Start();
            
            tiklanma_sayisi++;

            richTextBox1.Enabled = false;

            Random rnd = new Random();
            soru_id = rnd.Next(1, kayitli_soru_sayisi);
            sorulan_sorular[tiklanma_sayisi] = soru_id;
            Readquestion(soru_id);

            button1.Text = ReadChoice(soru_id, 2);
            button2.Text = ReadChoice(soru_id, 3);
            button3.Text = ReadChoice(soru_id, 4);
            button4.Text = ReadChoice(soru_id, 5);

            sorulan_soru_sayisi++;
            label7.Text = sorulan_soru_sayisi.ToString();
            
        }


        public bool SoruKontrol(int x)
        {
            
            for (int i = 0; i < sorulan_sorular.Length; i++)
            {
                if(sorulan_sorular[i] == x)
                {
                    
                    return true;
                }

            }


            return false;
        }


        public void Readquestion(int t)
        {
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];
                                 //cells[satir , stun]
            _Excel.Range cell = ws.Cells[t, 1];
            string CellValue = cell.Value;
            richTextBox1.Text = CellValue;
            

        }


        public string ReadChoice(int t, int z)
        {
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];

            _Excel.Range cell = ws.Cells[t, z];
            string CellValue = cell.Value;
            return CellValue;
        }

        public string ReadAnswer(int t)
        {
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[1];

            _Excel.Range cell = ws.Cells[t, 6];
            string CellValue = cell.Value;
            return CellValue;
           
        }


        string veda = "Yanlış cevap verdiniz.Tekrar görüşmek dileğiyle...";
        private void button1_Click(object sender, EventArgs e)
        {
            string cevap = ReadAnswer(soru_id);
            if(cevap == "a")
            {
                sure_sayac = 15;
                if (sorulan_soru_sayisi - ek_surenin_kullanildigi_soru_numarasi > 3)
                {
                    button8.Visible = true;
                }
                if (sorulan_soru_sayisi - yari_yariyanin_kullanildigi_soru > 8)
                {
                    button6.Visible = true;
                }
                if (sorulan_soru_sayisi - cift_cevap_kullanildigi_soru_numarasi > 5)
                {
                    button5.Enabled = true;
                }
                if (sorulan_soru_sayisi - pasin_kullanildigi_soru > 6)
                {
                    button7.Visible = true;
                }


                if(button5.Enabled == false)
                {
                    button1.BackColor = DefaultBackColor;
                    button2.BackColor = DefaultBackColor;
                    button3.BackColor = DefaultBackColor;
                    button4.BackColor = DefaultBackColor;
                }


                if (sorulan_soru_sayisi == sorulan_sorular.Length - 1)
                {
                    MessageBox.Show("Tebrikler yarışmayı başarıyla tamamladınız !!!");
                }
                bool kontrol;
                do
                {
                    Random rnd = new Random();
                    soru_id = rnd.Next(1, kayitli_soru_sayisi);

                    kontrol = SoruKontrol(soru_id);


                }while(kontrol);


                verilen_dogru_cevap_sayisi++;
                label8.Text = verilen_dogru_cevap_sayisi.ToString();
                button1.BackColor = Color.GreenYellow;
                tiklanma_sayisi++;
                sorulan_soru_sayisi++;
                label7.Text = sorulan_soru_sayisi.ToString();

                if (button1.Visible == false)
                    button1.Visible = true;
                if (button2.Visible == false)
                    button2.Visible = true;
                if (button3.Visible == false)
                    button3.Visible = true;
                if (button4.Visible == false)
                    button4.Visible = true;


                sorulan_sorular[tiklanma_sayisi] = soru_id;
                Readquestion(soru_id);
                Thread.Sleep(500);

                button1.BackColor = DefaultBackColor;
                button2.BackColor = DefaultBackColor;
                button3.BackColor = DefaultBackColor;
                button4.BackColor = DefaultBackColor;

                button1.Text = ReadChoice(soru_id, 2);
                button2.Text = ReadChoice(soru_id, 3);
                button3.Text = ReadChoice(soru_id, 4);
                button4.Text = ReadChoice(soru_id, 5);
                
            }
            else
            {
                button1.BackColor = Color.Red;

                if (button5.Enabled == false && (button2.BackColor != Color.Red && button3.BackColor != Color.Red && button4.BackColor != Color.Red))
                {
                    return;
                }
                if (button5.Enabled == false && (button2.BackColor == Color.Red || button3.BackColor == Color.Red || button4.BackColor == Color.Red))
                    button1.BackColor = DefaultBackColor;

               
                MessageBox.Show(veda);
                Thread.Sleep(500);
                System.Windows.Forms.Application.Exit();
            }

            
        }


        
   
        private void button2_Click(object sender, EventArgs e)
        {
            string cevap = ReadAnswer(soru_id);
            if (cevap == "b")
            {
                sure_sayac = 15;
                if(sorulan_soru_sayisi - ek_surenin_kullanildigi_soru_numarasi > 3 )
                {
                    button8.Visible = true;
                }
                if (sorulan_soru_sayisi - yari_yariyanin_kullanildigi_soru > 8)
                {
                    button6.Visible = true;
                }
                if (sorulan_soru_sayisi - cift_cevap_kullanildigi_soru_numarasi > 7)
                {
                    button5.Enabled = true;
                }
                if (sorulan_soru_sayisi - pasin_kullanildigi_soru > 7)
                {
                    button7.Visible = true;
                }


                if (button5.Enabled == false)
                {
                    button1.BackColor = DefaultBackColor;
                    button2.BackColor = DefaultBackColor;
                    button3.BackColor = DefaultBackColor;
                    button4.BackColor = DefaultBackColor;
                }

                if (sorulan_soru_sayisi == sorulan_sorular.Length - 1)
                {
                    MessageBox.Show("Tebrikler yarışmayı başarıyla tamamladınız !!!");

                }
                bool kontrol;
                do
                {
                    Random rnd = new Random();
                    soru_id = rnd.Next(1, kayitli_soru_sayisi);

                    kontrol = SoruKontrol(soru_id);


                } while (kontrol);

                button2.BackColor = Color.GreenYellow;
                verilen_dogru_cevap_sayisi++;
                label8.Text = verilen_dogru_cevap_sayisi.ToString();
                tiklanma_sayisi++;
                sorulan_soru_sayisi++;
                label7.Text = sorulan_soru_sayisi.ToString();

                if (button1.Visible == false)
                    button1.Visible = true;
                if (button2.Visible == false)
                    button2.Visible = true;
                if (button3.Visible == false)
                    button3.Visible = true;
                if (button4.Visible == false)
                    button4.Visible = true;


                sorulan_sorular[tiklanma_sayisi] = soru_id;
                Readquestion(soru_id);
                Thread.Sleep(500);
                button2.BackColor = DefaultBackColor;

                button1.Text = ReadChoice(soru_id, 2);
                button2.Text = ReadChoice(soru_id, 3);  // 10
                button3.Text = ReadChoice(soru_id, 4);
                button4.Text = ReadChoice(soru_id, 5);

            }
            else
            {
                button2.BackColor = Color.Red;

                if (button5.Enabled == false && (button1.BackColor != Color.Red && button3.BackColor != Color.Red && button4.BackColor != Color.Red))
                {
                    return;
                }
                if (button5.Enabled == false && (button1.BackColor == Color.Red || button3.BackColor == Color.Red || button4.BackColor == Color.Red))
                    button2.BackColor = DefaultBackColor;

                MessageBox.Show(veda);
                Thread.Sleep(500);
                System.Windows.Forms.Application.Exit();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string cevap = ReadAnswer(soru_id);
            if (cevap == "c")
            {
                sure_sayac = 15;
                if (sorulan_soru_sayisi - ek_surenin_kullanildigi_soru_numarasi > 3)
                {
                    button8.Visible = true;
                }
                if (sorulan_soru_sayisi - yari_yariyanin_kullanildigi_soru > 8)
                {
                    button6.Visible = true;
                }
                if (sorulan_soru_sayisi - cift_cevap_kullanildigi_soru_numarasi > 5)
                {
                    button5.Enabled = true;
                }
                if (sorulan_soru_sayisi - pasin_kullanildigi_soru > 6)
                {
                    button7.Visible = true;
                }


                if (button5.Enabled == false)
                {
                    button1.BackColor = DefaultBackColor;
                    button2.BackColor = DefaultBackColor;
                    button3.BackColor = DefaultBackColor;
                    button4.BackColor = DefaultBackColor;
                }


                if (sorulan_soru_sayisi == sorulan_sorular.Length - 1)
                {
                    MessageBox.Show("Tebrikler yarışmayı başarıyla tamamladınız !!!");

                }
                bool kontrol;
                do
                {
                    Random rnd = new Random();
                    soru_id = rnd.Next(1, kayitli_soru_sayisi);

                    kontrol = SoruKontrol(soru_id);


                } while (kontrol);

                button3.BackColor = Color.GreenYellow;
                verilen_dogru_cevap_sayisi++;
                label8.Text = verilen_dogru_cevap_sayisi.ToString();
                tiklanma_sayisi++;
                sorulan_soru_sayisi++;
                label7.Text = sorulan_soru_sayisi.ToString();

                if (button1.Visible == false)
                    button1.Visible = true;
                if (button2.Visible == false)
                    button2.Visible = true;
                if (button3.Visible == false)
                    button3.Visible = true;
                if (button4.Visible == false)
                    button4.Visible = true;


                sorulan_sorular[tiklanma_sayisi] = soru_id;
                Readquestion(soru_id);
                Thread.Sleep(500);
                button3.BackColor = DefaultBackColor;

                button1.Text = ReadChoice(soru_id, 2);
                button2.Text = ReadChoice(soru_id, 3);
                button3.Text = ReadChoice(soru_id, 4);
                button4.Text = ReadChoice(soru_id, 5);
            }
            else
            {
                button3.BackColor = Color.Red;

                if (button5.Enabled == false && (button1.BackColor != Color.Red && button2.BackColor != Color.Red && button4.BackColor != Color.Red))
                {
                    return;
                }
                if (button5.Enabled == false && (button1.BackColor == Color.Red || button2.BackColor == Color.Red || button4.BackColor == Color.Red))
                    button3.BackColor = DefaultBackColor;

                MessageBox.Show(veda);
                Thread.Sleep(500);
                System.Windows.Forms.Application.Exit();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string cevap = ReadAnswer(soru_id);   
            if (cevap == "d")
            {
                sure_sayac = 15;
                if (sorulan_soru_sayisi - ek_surenin_kullanildigi_soru_numarasi > 3)
                {
                    button8.Visible = true;
                }
                if(sorulan_soru_sayisi - yari_yariyanin_kullanildigi_soru > 8)
                {
                    button6.Visible = true;
                }
                if(sorulan_soru_sayisi - cift_cevap_kullanildigi_soru_numarasi > 5)
                {
                    button5.Enabled = true;
                }
                if(sorulan_soru_sayisi - pasin_kullanildigi_soru > 6)
                {
                    button7.Visible = true;
                }


                if (button5.Enabled == false)
                {
                    button1.BackColor = DefaultBackColor;
                    button2.BackColor = DefaultBackColor;
                    button3.BackColor = DefaultBackColor;
                    button4.BackColor = DefaultBackColor;
                }


                if (sorulan_soru_sayisi == sorulan_sorular.Length - 1)
                {
                    MessageBox.Show("Tebrikler yarışmayı başarıyla tamamladınız !!!");
                    System.Windows.Forms.Application.Exit();
                }
                bool kontrol;
                do
                {
                    Random rnd = new Random();
                    soru_id = rnd.Next(1, kayitli_soru_sayisi);

                    kontrol = SoruKontrol(soru_id);


                } while (kontrol);

                button4.BackColor = Color.GreenYellow;
                verilen_dogru_cevap_sayisi++;
                label7.Text = verilen_dogru_cevap_sayisi.ToString();
                tiklanma_sayisi++;
                sorulan_soru_sayisi++;
                label8.Text = sorulan_soru_sayisi.ToString();

                if (button1.Visible == false)
                    button1.Visible = true;
                if (button2.Visible == false)
                    button2.Visible = true;
                if (button3.Visible == false)
                    button3.Visible = true;
                if (button4.Visible == false)
                    button4.Visible = true;

                Readquestion(soru_id);
                Thread.Sleep(500);
                button4.BackColor = DefaultBackColor;

                button1.Text = ReadChoice(soru_id, 2);
                button2.Text = ReadChoice(soru_id, 3);
                button3.Text = ReadChoice(soru_id, 4);
                button4.Text = ReadChoice(soru_id, 5);
            }
            else
            {
                button4.BackColor = Color.Red;

                if (button5.Enabled == false && (button1.BackColor != Color.Red && button2.BackColor != Color.Red && button3.BackColor != Color.Red))
                {
                    return;
                }
                if (button5.Enabled == false && (button1.BackColor == Color.Red || button2.BackColor == Color.Red || button3.BackColor == Color.Red))
                    button4.BackColor = DefaultBackColor;

                MessageBox.Show(veda);
                Thread.Sleep(500);
                System.Windows.Forms.Application.Exit();
            }
        }

        int yari_yariyanin_kullanildigi_soru = 0;
        private void button6_Click(object sender, EventArgs e)
        {

            yari_yariyanin_kullanildigi_soru = sorulan_soru_sayisi;
            button6.Visible = false;
            string cevap = ReadAnswer(soru_id);

            int elenecek1, elenecek2 = 0;

            int[] cevap_kumesi = new int[3]; 
             



            if(cevap == "a")
            {

                cevap_kumesi[0] = 1;
                cevap_kumesi[1] = 2;
                cevap_kumesi[2] = 3;


                Random rnd = new Random();
                int xx = rnd.Next(0, 2);

                elenecek1 = xx;

                int xy = rnd.Next(0, 2);
                if (elenecek1 == xy)
                {
                    do
                    {
                        xy = rnd.Next(0, 2);
                    } while (elenecek1 ==  xy);
                    elenecek2 =  xy;
                }


                int elenecek_button1 = cevap_kumesi[elenecek1];
                int elenecek_button2 = cevap_kumesi[elenecek2];

                if (elenecek_button1 == 1)
                    button2.Visible = false;
                else if (elenecek_button1 == 2)
                    button3.Visible = false;
                else if (elenecek_button1 == 3)
                    button4.Visible = false;
                else
                    ;


                if (elenecek_button2 == 1)
                    button2.Visible = false;
                else if (elenecek_button2 == 2)
                    button3.Visible = false;
                else if (elenecek_button2 == 3)
                    button4.Visible = false;
                else
                    ;


            }


            if (cevap == "b")
            {

                cevap_kumesi[0] = 0;
                cevap_kumesi[1] = 2;
                cevap_kumesi[2] = 3;


                Random rnd = new Random();
                int xx = rnd.Next(0, 2);

                elenecek1 = xx;

                int xy = rnd.Next(0, 2);
                if (elenecek1 == xy)
                {
                    do
                    {
                        xy = rnd.Next(0, 2);
                    } while (elenecek1 == xy);
                    elenecek2 = xy;
                }


                int elenecek_button1 = cevap_kumesi[elenecek1];
                int elenecek_button2 = cevap_kumesi[elenecek2];

                if (elenecek_button1 == 0)
                    button1.Visible = false;
                else if (elenecek_button1 == 2)
                    button3.Visible = false;
                else if (elenecek_button1 == 3)
                    button4.Visible = false;
                else
                    ;


                if (elenecek_button2 == 0)
                    button1.Visible = false;
                else if (elenecek_button2 == 2)
                    button3.Visible = false;
                else if (elenecek_button2 == 3)
                    button4.Visible = false;
                else
                    ;


            }


            if (cevap == "c")
            {

                cevap_kumesi[0] = 0;
                cevap_kumesi[1] = 1;
                cevap_kumesi[2] = 3;


                Random rnd = new Random();
                int xx = rnd.Next(0, 2);

                elenecek1 = xx;

                int xy = rnd.Next(0, 2);
                if (elenecek1 == xy)
                {
                    do
                    {
                        xy = rnd.Next(0, 2);
                    } while (elenecek1 == xy);
                    elenecek2 = xy;
                }


                int elenecek_button1 = cevap_kumesi[elenecek1];
                int elenecek_button2 = cevap_kumesi[elenecek2];

                if (elenecek_button1 == 0)
                    button1.Visible = false;
                else if (elenecek_button1 == 1)
                    button2.Visible = false;
                else if (elenecek_button1 == 3)
                    button4.Visible = false;
                else
                    ;


                if (elenecek_button2 == 0)
                    button1.Visible = false;
                else if (elenecek_button2 == 1)
                    button2.Visible = false;
                else if (elenecek_button2 == 3)
                    button4.Visible = false;
                else
                    ;


            }

            if (cevap == "d")
            {

                cevap_kumesi[0] = 0;
                cevap_kumesi[1] = 1;
                cevap_kumesi[2] = 2;


                Random rnd = new Random();
                int xx = rnd.Next(0, 2);

                elenecek1 = xx;

                int xy = rnd.Next(0, 2);
                if (elenecek1 == xy)
                {
                    do
                    {
                        xy = rnd.Next(0, 2);
                    } while (elenecek1 == xy);
                    elenecek2 = xy;
                }


                int elenecek_button1 = cevap_kumesi[elenecek1];
                int elenecek_button2 = cevap_kumesi[elenecek2];

                if (elenecek_button1 == 0)
                    button1.Visible = false;
                else if (elenecek_button1 == 1)
                    button2.Visible = false;
                else if (elenecek_button1 == 2)
                    button3.Visible = false;
                else
                    ;


                if (elenecek_button2 == 0)
                    button1.Visible = false;
                else if (elenecek_button2 == 1)
                    button2.Visible = false;
                else if (elenecek_button2 == 2)
                    button3.Visible = false;
                else
                    ;


            }


        }

        int pasin_kullanildigi_soru = 0;
        private void button7_Click(object sender, EventArgs e)
        {
            pasin_kullanildigi_soru = sorulan_soru_sayisi;
            sure_sayac = 15;
            button7.Visible = false;
            bool kontrol;
            do
            {
                Random rnd = new Random();
                soru_id = rnd.Next(1, kayitli_soru_sayisi);

                kontrol = SoruKontrol(soru_id);


            } while (kontrol);

            
            tiklanma_sayisi++;

            if (button1.Visible == false)
                button1.Visible = true;
            if (button2.Visible == false)
                button2.Visible = true;
            if (button3.Visible == false)
                button3.Visible = true;
            if (button4.Visible == false)
                button4.Visible = true;

            Readquestion(soru_id);
            
           

            button1.Text = ReadChoice(soru_id, 2);
            button2.Text = ReadChoice(soru_id, 3);
            button3.Text = ReadChoice(soru_id, 4);
            button4.Text = ReadChoice(soru_id, 5);

            sorulan_soru_sayisi++;
            label7.Text = sorulan_soru_sayisi.ToString();
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        int sure_sayac = 15;
        private void timer1_Tick(object sender, EventArgs e)
        {
            
            label10.Text = sure_sayac.ToString();
            sure_sayac--;
            if (sure_sayac == 0)
            {
                this.Visible = false;
                MessageBox.Show("Sureniz bitmiştir !!!");
                Thread.Sleep(500);
                System.Windows.Forms.Application.Exit();
            }
        }

        int ek_surenin_kullanildigi_soru_numarasi = 0;
        private void button8_Click(object sender, EventArgs e)
        {
            ek_surenin_kullanildigi_soru_numarasi = sorulan_soru_sayisi;
            button8.Visible = false;
            sure_sayac += 15;
        }

        int cift_cevap_kullanildigi_soru_numarasi = 0;
        private void button5_Click_1(object sender, EventArgs e)
        {
            cift_cevap_kullanildigi_soru_numarasi = sorulan_soru_sayisi;
            button5.Enabled = false;

        }

       
    }
}
