using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace testDigitalSealed
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_DragDrop(object sender, DragEventArgs e)
        {
            string path;

            path = ((string[])e.Data.GetData(DataFormats.FileDrop, false))[0];//擷取檔案路徑.

            if (path.EndsWith(".dsicdp"))
            {
                FileInfo fi = new FileInfo(path);
                Directory.CreateDirectory(Path.Combine(fi.DirectoryName, "unzip"));
                GZip.Decompress(fi.DirectoryName, Path.Combine(fi.DirectoryName, "unzip"), fi.Name);
                MessageBox.Show("解壓至: " + Path.Combine(fi.DirectoryName, "unzip"));
            }
            else
                MessageBox.Show("檔案非封裝壓縮檔!"); 

        }

        private void Form2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop, false))
            {
                // 允許拖拉動作繼續 (這時滑鼠游標應該會顯示 +)
                e.Effect = DragDropEffects.All;//
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            //NPBDCP.DCPWSE wse = new NPBDCP.DCPWSE();
            try
            {
            //    System.Xml.XmlNode rtnd = wse.getXmlforSign("1027000001", "102K000001", "A07090000D", "資訊室", "12", "主任", "帝緯系統", "A12345", "OK", "發");
            //    wse.updateDCPData("A07090000D", "1027000002", "102K000003", "aaaaaaaaaaaaaaaaaaaaa", "cccccccccccccccccccccc", "ds_Flow_12_帝緯系統_批示", "sign_Flow_12_帝緯系統_批示", "A12345", "7");
            //    //string signstr = wse.geturl();
            }
            catch (Exception err)
            {

            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            folderBrowserDialog1.ShowDialog();
            
            GZip.Compress(folderBrowserDialog1.SelectedPath, @"c:\temp\", "test.dsicdp");  
        }
    }
}
