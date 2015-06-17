using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.SignDocFolder
{
    /// <summary>
    /// 簽核文件夾
    /// </summary>
    public class SignDocFolder
    {
        private string _tagName = "簽核文件夾";

        public string TagName
        {
            get { return _tagName; }
            set { _tagName = value; }
        }

        private string _id = "";

        public string Id
        {
            get { return _id; }
            set { _id = value; }
        }

        private string _geneTime = "";

        public string GeneTime
        {
            get { return _geneTime; }
            set { _geneTime = value; }
        }

        private SignDocList _簽核文稿清單 = null;

        public SignDocList 簽核文稿清單
        {
            get { return _簽核文稿清單; }
            set { _簽核文稿清單 = value; }
        }

        private FilesList _檔案清單 = null;

        public FilesList 檔案清單
        {
            get { return _檔案清單; }
            set { _檔案清單 = value; }
        }

        private MergeFileList _併文清單 = null;

        public MergeFileList 併文清單
        {
            get { return _併文清單; }
            set { _併文清單 = value; }
        }

        /// <summary>
        /// 簽核文夾建構子
        /// </summary>
        /// <param name="id">ID值</param>
        /// <param name="genetime">產生時間</param>
        /// <param name="sdl">簽核文稿清單</param>
        /// <param name="fl">檔案清單</param>
        public SignDocFolder(string id,string genetime,SignDocList sdl,FilesList fl)
        {
            _id = id;
            _geneTime = genetime;
            _簽核文稿清單 = sdl;
            _檔案清單 = fl;
        }

        /// <summary>
        /// 取得簽核文件夾節點內容
        /// </summary>
        /// <returns></returns>
        public XmlNode getSignDocFolderNode()
        {
            Dictionary<string, string> attrs = new Dictionary<string, string>();
            attrs.Add("Id", _id);
            attrs.Add("產生時間", _geneTime);
            XmlNode rtnNode = xmlTool.MakeNode(_tagName , attrs);
            if (_tagName == "簽核文件夾")
            {
                if (_簽核文稿清單 != null)
                    rtnNode.AppendChild(_簽核文稿清單.getSignDocListNode());
            }
            else
                rtnNode.AppendChild(xmlTool.MakeNode("母文", _併文清單.MotherDocNo));
            rtnNode.AppendChild(_檔案清單.getFilesListNode());
            if (_併文清單.MotherDocNo  != "" && _tagName == "簽核文件夾")
                rtnNode.AppendChild(_併文清單.getMergeFileListNode());
            return rtnNode;
        }

    }
}
