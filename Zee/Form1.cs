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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;

            using (FtpClient ftp = new FtpClient(server, new System.Net.NetworkCredential { UserName = username, Password = password }))
            {
                FtpListItem[] listing = ftp.GetListing();
                foreach (FtpListItem ftpItem in listing)
                {
                    if (ftpItem.Type != FtpObjectType.File)
                        continue;

                    using (MemoryStream ms = new MemoryStream())
                    {
                        ftp.DownloadBytes(out ms, ftpItem.Name);

                    }

                }

            }


            button1.Enabled = true;

        }
    }
}
