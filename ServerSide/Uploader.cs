using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using FluentFTP;

namespace iNovation.ServerSide
{
    public class Uploader
    {

        private string server { get; set; }
        private string username { get; set; }
        private string password { get; set; }

        public Uploader(string ftp_server, string username, string password)
        {
            this.server = ftp_server;
            this.username = username;
            this.password = password;

        }


        FtpClient CreateFtpClient()
        {
            return new FtpClient(this.server, new NetworkCredential { UserName = username, Password = password });
        }


        /// <summary>
        /// Uploads content of file.
        /// </summary>
        /// <param name="filename_with_extension"></param>
        /// <param name="fileBytes"></param>
        /// <param name="remoteFolderPathEscapedString"></param>
        /// <example>
        /// new Uploader(...).uploadBytes("1.jpg", bannerPic.FileBytes, directory()) //from ASP.NET
        /// new Uploader(...).uploadBytes("1.jpg", PictureFromStream(...), directory()) //from regular Desktop application
        /// </example>
        public void uploadBytes(string filename_with_extension, byte[] fileBytes, string remoteFolderPathEscapedString = "")
        {
            using (FtpClient ftp = CreateFtpClient())
            {
                ftp.UploadBytes(fileBytes, remoteFolderPathEscapedString + filename_with_extension);
            }

        }


        /// <summary>
        /// Called from C# desktop application (from OpenFileDialog thereabouts). 
        /// Prefer uploadBytes().
        /// </summary>
        /// <param name="info"></param>
        /// <param name="remoteFolderPathEscapedString"></param>
        public void uploadFile(FileInfo info, string remoteFolderPathEscapedString = "")
        {
            using (FtpClient ftp = CreateFtpClient())
            {
                using (FileStream stream = File.OpenRead(info.FullName))
                {
                    ftp.UploadStream(stream, remoteFolderPathEscapedString + info.Name);
                }

            }

        }

    }
}
