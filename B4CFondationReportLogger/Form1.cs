using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace B4CFondationReportLogger
{
    public partial class AcctivityRepport : Form
    {
        public AcctivityRepport()
        {
            InitializeComponent();
        }

        private void BtnConfig_Click(object sender, EventArgs e)
        {
            new FormConfiguration().Show();
            this.Hide();
        }

        private void bunifuImageButton2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

    } 
}
