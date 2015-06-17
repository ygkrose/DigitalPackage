using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using FSCAPIATLLib;
using System.IO;
using System.Security.Cryptography.X509Certificates;

namespace retrieveCert
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            chkOEFile(((string[])e.Data.GetData(DataFormats.FileDrop, false))[0]);
        }

        private void chkOEFile(string SourcePath)
        {
            
            //FSXMLATLClass fsxml = new FSXMLATLClass();
            //GPKICryptATLClass fsgpki = new GPKICryptATLClass();
            FSCAPI X = new FSCAPIATLLib.FSCAPI();
            byte[] FileData = null;
            byte[] SignData = null;
            string SignString = "";
            BinaryReader binReader = new BinaryReader(File.Open(SourcePath, FileMode.Open));
            try
            {
                if (binReader.PeekChar() != -1)
                {
                    if (binReader.ReadString() == "S")
                    {
                        int SignDataLength = binReader.ReadInt16();
                        binReader.ReadBytes(6);
                        SignData = binReader.ReadBytes(SignDataLength);

                        int FileDataLength = binReader.ReadInt16();
                        binReader.ReadBytes(6);
                        FileData = binReader.ReadBytes(Convert.ToInt32(new FileInfo(SourcePath).Length) - 6 - 6 - SignDataLength);
                        SignString = new ASCIIEncoding().GetString(SignData);
                        //fsgpki.FSGPKI_EnumCerts();
                    }
                    else
                    {
                        //檔案沒有加簽，不驗證
                        binReader.Close();
                        label1.Text = "非OE檔案或未加簽OE檔案";
                    }
                }
                else
                {
                    label1.Text = "檔案有誤";
                }
            }
            catch
            {
            }
            finally
            {
                binReader.Close();
            }

            if (FileData != null)
            {
                using (BinaryWriter binWriter = new BinaryWriter(File.Open(SourcePath + "tmp", FileMode.Create)))
                {
                    binWriter.Write(FileData);
                }

                string strFileValue = X.FSCAPI_ReadFile(SourcePath + "tmp", 0);
                const int FS_FLAG_BASE64_ENCODE = 0x00001000;
                const int FS_FLAG_DETACHMSG = 0x00004000;

                int retValue = X.FSCAPIVerify(SignString, strFileValue, "", FS_FLAG_DETACHMSG | FS_FLAG_BASE64_ENCODE, 0, 0);
                if (retValue == 0)
                {
                    string certstr = X.GetVerifyRtnCert();
                    if (!string.IsNullOrEmpty(certstr))
                    {
                        byte[] b = Encoding.UTF8.GetBytes(certstr);
                        X509Certificate2 cert = new X509Certificate2(b);
                        X509Certificate2UI.DisplayCertificate(cert);
                    }
                    else
                        label1.Text = "取得憑證發生錯誤";
                }
                else
                    label1.Text = "檔案驗簽失敗無法取得憑證內容";
                if (File.Exists(SourcePath + "tmp"))
                    File.Delete(SourcePath + "tmp");
            }
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                // 允許拖拉動作繼續 (這時滑鼠游標應該會顯示 +)
                e.Effect = DragDropEffects.All;//
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "*.zip|*.oec|*.oes|*.oel";
            openFileDialog1.DefaultExt = "*.zip";
            openFileDialog1.FileName = "*.*";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                label1.Text = openFileDialog1.FileName;
                chkOEFile(label1.Text);
            }
            else
                label1.Text = "";
        }
    }
}
