using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.SignDocFolder
{
    /// <summary>
    /// 簽核文稿清單
    /// </summary>
    public class SignDocList : List<Document> 
    {
        /// <summary>
        /// 文稿數
        /// </summary>
        private string _DocCount = "";

        public string DocCount
        {
            get { return _DocCount; }
            set { _DocCount = value; }
        }

        public SignDocList()
        {
        }

        /// <summary>
        /// 取得簽核文稿清單節點內容
        /// </summary>
        /// <returns></returns>
        public XmlNode getSignDocListNode()
        {
            _DocCount = this.Count.ToString() ;
            XmlNode rtnNode = xmlTool.MakeNode("簽核文稿清單", "文稿數", _DocCount);
            int serialno = 0;
            foreach (Document mydoc in this)
            {
                Dictionary<string, string> attr = new Dictionary<string, string>();
                attr.Add("原始檔序號", mydoc.Ofileserial);
                attr.Add("序號", serialno.ToString());
                attr.Add("產生時間", mydoc.GeneTime);
                XmlNode child = xmlTool.MakeNode("文稿",attr);
                child.AppendChild(xmlTool.MakeNode("名稱",mydoc.DFilename));
                child.AppendChild(xmlTool.MakeNode("文稿類型", mydoc.DFormat));
                if (mydoc.文稿頁面檔.文稿頁面清單.Count > 0)
                    child.AppendChild(mydoc.文稿頁面檔.getDocPageFileNode());
                if (mydoc.附件清單 != null)
                    child.AppendChild(mydoc.附件清單.getAttsListNode());
                rtnNode.AppendChild(child);  
                serialno++;
            }
            
            return rtnNode;
        }

    }

    /// <summary>
    /// 文稿
    /// </summary>
    public class Document
    {
        private string _DocName = "";
        /// <summary>
        /// 文稿名稱
        /// </summary>
        public string DocName
        {
            get { return _DocName; }
            set { _DocName = value; }
        }
        private string _DFilename = "";
        /// <summary>
        /// 檔案名稱
        /// </summary>
        public string DFilename
        {
            get { return _DFilename; }
            set { _DFilename = value; }
        }
        private string _DFormat = "";
        /// <summary>
        /// 文稿類型
        /// </summary>
        public string DFormat
        {
            get { return _DFormat; }
            set { _DFormat = value; }
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

        private bool _judged = false;
        /// <summary>
        /// 文稿已決行
        /// </summary>
        public bool Judged
        {
            get { return _judged; }
            set { _judged = value; }
        }


        private DocPageFile _文稿頁面檔 = null;

        public DocPageFile 文稿頁面檔
        {
            get 
            {
                if (_文稿頁面檔 == null) _文稿頁面檔 = new DocPageFile(_geneTime);
                return _文稿頁面檔; 
            }
            set 
            {
                if (_文稿頁面檔 == null) _文稿頁面檔 = new DocPageFile(_geneTime);  
                _文稿頁面檔 = value; 
            }
        }

        private AttachmentsDesc.AttsList _附件清單 = null;

        public AttachmentsDesc.AttsList 附件清單
        {
            get { return _附件清單; }
            set { _附件清單 = value;}
        }
         

        /// <summary>
        /// 文稿建構子
        /// </summary>
        /// <param name="DName">名稱 ex:創簽/創稿</param>
        /// <param name="ofileserial">原始檔序號</param>
        /// <param name="geneTime">產生時間</param>
        /// <param name="DocTyp">文稿類型</param>
        /// <param name="DocTyp">文稿已決行</param>
        public Document(string filename,string dname, string ofileserial, string geneTime,string DocTyp,bool isjudge)
        {
            _DFilename = filename;
            _DocName = dname;
            _ofileserial = ofileserial;
            _geneTime = geneTime;
            _DFormat = DocTyp;
            _judged = isjudge;
        }

        /// <summary>
        /// 文稿建構子
        /// </summary>
        /// <param name="DName">名稱 ex:創簽/創稿</param>
        /// <param name="ofileserial">原始檔序號</param>
        /// <param name="geneTime">產生時間</param>
        /// <param name="DocTyp">文稿類型</param>
        public Document(string filename, string dname, string ofileserial, string geneTime, string DocTyp)
        {
            _DFilename = filename;
            _DocName = dname;
            _ofileserial = ofileserial;
            _geneTime = geneTime;
            _DFormat = DocTyp;
        }

    }
}
