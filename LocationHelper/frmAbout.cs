using System;
using System.Configuration;
using System.Linq;
using System.Windows.Forms;
using LocationHelper.Helper;

namespace LocationHelper
{
    public partial class FrmAbout : Form
    {
        public FrmAbout()
        {
            InitializeComponent();
        }

        private void frmAbout_Load(object sender, EventArgs e)
        {
            label1.Text=string.Format(label1.Text, ConfigHelper.Name);
        }
    }
}
