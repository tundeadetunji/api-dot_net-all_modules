using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FluentFTP;
using System.IO;

namespace Zee
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string server = "ftp.radianthope.org.ng";
        string username = "lbtwieye";
        string password = "h568]ImZFq7;zG";
        string remoteFolderPathEscapedString = @"dev.radianthope.org.ng\wwwroot\images\";
        private void Form1_Load(object sender, EventArgs e)
        {
            new Form2().Show();
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;


            button1.Enabled = true;

        }
    }
}
