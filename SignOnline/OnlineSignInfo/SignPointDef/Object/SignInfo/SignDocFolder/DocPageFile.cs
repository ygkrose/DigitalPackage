using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.SignDocFolder
{
    /// <summary>
    /// 文稿頁面檔
    /// </summary>
    public class DocPageFile
    {
        private string _RecordType = "單層式";
        /// <summary>
        /// 紀錄方式
        /// </summary>
        public string RecordType
        {
            get { return _RecordType; }
            set { _RecordType = value; }
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

        private DocPageList _文稿頁面清單 = null;

        public DocPageList 文稿頁面清單
        {
            get 
            {
                if (_文稿頁面清單 == null) _文稿頁面清單 = new DocPageList();
                return _文稿頁面清單; 
            }
            set 
            {
                if (_文稿頁面清單 == null) _文稿頁面清單 = new DocPageList();
                _文稿頁面清單 = value; 
            }
        } 

        /// <summary>
        /// 文稿頁面檔建構子
        /// </summary>
        /// <param name="genetime">產生時間套用致所有頁面檔產生時間</param>
        public DocPageFile(string genetime)
        {
            _geneTime = genetime;
        }

        /// <summary>
        /// 取得文稿頁面檔節點內容
        /// </summary>
        /// <returns></returns>
        public XmlNode getDocPageFileNode()
        {
            Dictionary<string, string> attrs = new Dictionary<string, string>();
            attrs.Add("記錄方式", _RecordType);
            attrs.Add("產生時間", _geneTime);
            XmlNode rtnNode = xmlTool.MakeNode("文稿頁面檔", attrs);
            if (_文稿頁面清單!=null)
                rtnNode.AppendChild(_文稿頁面清單.getDocPageListNode(_geneTime));
            return rtnNode;
        }
    }

    /// <summary>
    /// 文稿頁面清單
    /// </summary>
    public class DocPageList : List<DocPageList.Page>
    {
        private string _pagecount = "";
        /// <summary>
        /// 頁面數
        /// </summary>
        public string PageCount
        {
            get { return _pagecount; }
            set { _pagecount = value; }
        }

        /// <summary>
        /// 取得文稿頁面清單節點
        /// </summary>
        /// <returns></returns>
        public XmlNode getDocPageListNode(string _geneTime)
        {
            _pagecount = this.Count.ToString();  
            XmlNode rtnNode = xmlTool.MakeNode("文稿頁面清單", "頁面數", _pagecount);
            int serialno = 0;
            foreach (Page p in this)
            {
                Dictionary<string, string> attrs = new Dictionary<string, string>();
                attrs.Add("原始檔序號", p.Ofileserial);
                attrs.Add("序號", serialno.ToString());
                attrs.Add("產生時間", _geneTime);
                rtnNode.AppendChild(xmlTool.MakeNode("頁面", attrs));    
                serialno++; 
            }
            return rtnNode;
        }

        /// <summary>
        /// 取得各頁面節點
        /// </summary>
        /// <returns></returns>
        public XmlNodeList getAllDocPageNode(string _geneTime)
        {
            _pagecount = this.Count.ToString();
            XmlNode rtnNode = xmlTool.MakeNode("文稿頁面清單", "頁面數", _pagecount);
            int serialno = 0;
            foreach (Page p in this)
            {
                Dictionary<string, string> attrs = new Dictionary<string, string>();
                attrs.Add("原始檔序號", p.Ofileserial);
                attrs.Add("序號", serialno.ToString());
                attrs.Add("產生時間", _geneTime);
                rtnNode.AppendChild(xmlTool.MakeNode("頁面", attrs));
                serialno++;
            }
            return rtnNode.ChildNodes;
        }

        /// <summary>
        /// 頁面
        /// </summary>
        public class Page
        {
            private string _ofileserial = ""; //原始檔序號
            /// <summary>
            /// 原始檔序號
            /// </summary>
            public string Ofileserial
            {
                get { return _ofileserial; }
                set { _ofileserial = value; }
            }
            

            /// <summary>
            /// 頁面建構子
            /// </summary>
            /// <param name="oserialno">原始檔序號</param>
            public Page(string oserialno)
            {
                _ofileserial = oserialno;
            }
        }
    }
}
