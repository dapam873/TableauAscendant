using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TableauAscendant
{
    public partial class Form2 : Form
    {

        /// <summary>
        /// 
        /// </summary>
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            VersionLb.Text= "Version " + Application.ProductVersion + "B";
        }

        private void PambrunLLb_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("http://pambrun.net");
            }
            catch
            {
                MessageBox.Show("Impossible de se connecter.");
            }
        }

        private void VersionLb_Click(object sender, EventArgs e)
        {

        }
    }
}
