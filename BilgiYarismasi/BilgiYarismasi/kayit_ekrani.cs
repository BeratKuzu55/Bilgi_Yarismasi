using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


namespace BilgiYarismasi
{
    public partial class kayit_ekrani : Form
    {
        public kayit_ekrani()
        {
            InitializeComponent();
        }

        private void kayit_ekrani_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string gelen_ad = textBox1.Text;
            string gelen_soyad = textBox2.Text;

            if(gelen_ad != null && gelen_ad != "" && gelen_soyad != null && gelen_soyad != "")
            {
                Form1 yarisma = new Form1(gelen_ad , gelen_soyad);
                this.Visible = false;
                yarisma.Visible = true;
            }
        }
    }
}
