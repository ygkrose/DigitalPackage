using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.IO;
using DigitalSealed.Tools;
using DigitalSealed.AttachmentsDesc;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.SignDocFolder
{
    /// <summary>
    /// 檔案清單
    /// </summary>
    public class FilesList : List<ElecFileInfo>
    {
        /// <summary>
        /// 檔案數
        /// </summary>
        private string _filesCnt = "";

        /// <summary>
        /// 建構子
        /// </summary>
        public FilesList()
        {
        }

        /// <summary>
        /// 取得檔案清單標籤集內容
        /// </summary>
        /// <returns></returns>
        public XmlNode getFilesListNode()
        {
            _filesCnt = this.Count.ToString() ;
            XmlNode rtnNode = xmlTool.MakeNode("檔案清單", "檔案數", _filesCnt);
            int serial = 0;
            foreach (ElecFileInfo finfo in this)
            {
                Dictionary<string, string> attr = new Dictionary<string, string>();
                attr.Add("原始檔序號", finfo.Ofileserial);
                attr.Add("序號", serial.ToString());
                attr.Add("產生時間", finfo.GeneTime);

                XmlNode childs = xmlTool.MakeNode("電子檔案資訊", attr);
                childs.AppendChild(xmlTool.MakeNode("檔案名稱", finfo.Filename));
                childs.AppendChild(xmlTool.MakeNode("檔案大小", finfo.Fsize));
                childs.AppendChild(xmlTool.MakeNode("檔案格式", finfo.FFormat));
                rtnNode.AppendChild(childs);  
                serial++;
            }
            return rtnNode;
        }
    }

    /// <summary>
    ///  電子檔案資訊
    /// </summary>
    public class ElecFileInfo
    {
        private string _filename = "";
        /// <summary>
        /// 檔案名稱
        /// </summary>
        public string Filename
        {
            get { return _filename; }
            set { _filename = value; }
        }
        private string _fsize = "";
        /// <summary>
        /// 檔案大小
        /// </summary>
        public string Fsize
        {
            get { return _fsize; }
            set { _fsize = value; }
        }
        private string _fFormat = "";
        /// <summary>
        /// 檔案格式
        /// </summary>
        public string FFormat
        {
            get { return _fFormat; }
            set { _fFormat = value; }
        }
        private string _ofileserial = ""; //原始檔序號
        /// <summary>
        /// 原始檔序號
        /// </summary>
        public string Ofileserial
        {
            get { return _ofileserial; }
            set { _ofileserial = value; }
        }
        private string _geneTime = "";
        /// <summary>
        /// 產生時間
        /// </summary>
        public string GeneTime
        {
            get { return _geneTime; }
            set { _geneTime = value; }
        }

        /// <summary>
        /// 類別轉型將ElecAtt轉成ElecFileInfo
        /// </summary>
        /// <param name="att"></param>
        /// <returns></returns>
        public static explicit operator ElecFileInfo(ElecAtt att)
        {
            return new ElecFileInfo(att.AttName, att.Ofileserial, att.GeneTime);    
        }

        /// <summary>
        /// 電子檔案資訊建構子
        /// </summary>
        /// <param name="filename">檔案名稱</param>
        /// <param name="ofileserial">原始檔序號</param>
        /// <param name="geneTime">產生時間</param>
        public ElecFileInfo(string filename, string ofileserial,string geneTime)
        {
            _filename = filename;
            if (File.Exists(Path.Combine( ConstVar.SourceFilesPath , _filename)))
            {
                FileInfo fi = new FileInfo(Path.Combine(ConstVar.SourceFilesPath, _filename));
                _fsize = fi.Length.ToString();
                _fFormat = (fi.Extension.StartsWith(".") ? fi.Extension.Substring(1) : fi.Extension).ToUpper();
                _ofileserial = ofileserial;
                _geneTime = geneTime;
            }
            else
                throw new Exception("找不到檔案:" + Path.Combine(ConstVar.SourceFilesPath , _filename)); 
        }

       
    }

}
