using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.SignDocFolder;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.ReceivedDoc
{
    /// <summary>
    /// 來文文件夾
    /// </summary>
    public class ReceivedDoc
    {
        private string _parDocNO = "";
        private string _DocFormat = "電子"; //紙本

        public string DocFormat
        {
            get { return _DocFormat; }
            set { _DocFormat = value; }
        }
        private string _serialNo = "0";
        private bool RefDoc = false;

        //來文類型
        private string _indocType = "";

        private string _Id = "";

        public string Id
        {
            get { return _Id; }
            set { _Id = value; }
        }

        /// <summary>
        /// 產生時間
        /// </summary>
        private string _geneTime = "";

        public string GeneTime
        {
            get { return _geneTime; }
            set { _geneTime = value; }
        }

        /// <summary>
        /// 來文清單的來文數固定只有1
        /// </summary>
        private string _DocCount = "1";

        internal interface 來文介面
        {
            XmlNode getInDocNode(string genetime);
            void AddAtt(DigitalSealed.AttachmentsDesc.IAttachment att);
            void AddPage(DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.SignDocFolder.DocPageList.Page page);
            AttachmentsDesc.AttsList get附件清單();
        }

        class 電子來文 : 來文介面
        {
            private string _原始檔序號 = "0";

            private AttachmentsDesc.AttsList _附件清單 = null;
            public AttachmentsDesc.AttsList 附件清單
            {
                get { return _附件清單; }
                set { _附件清單 = value; }
            }


            #region 來文介面 成員

            XmlNode 來文介面.getInDocNode(string genetime)
            {
                XmlNode rtnNode = xmlTool.MakeNode("電子來文", "原始檔序號", _原始檔序號);

                if (_附件清單 != null)
                {
                    rtnNode.AppendChild(_附件清單.getAttsListNode());
                }
                return rtnNode;
            }

            void 來文介面.AddAtt(AttachmentsDesc.IAttachment att)
            {
                if (_附件清單 == null) _附件清單 = new AttachmentsDesc.AttsList(); 
                _附件清單.Add(att); 
            }

            void 來文介面.AddPage(SignDocFolder.DocPageList.Page page)
            {
                throw new NotImplementedException();
            }

            AttachmentsDesc.AttsList 來文介面.get附件清單()
            {
                return _附件清單; 
            }

            #endregion
        }

        class 紙本來文 : 來文介面
        {
            private string _頁面數;
            private string _原始檔序號;
            private string _序號;
            private string _產生時間;

            private SignDocFolder.DocPageList _頁面 = new DocPageList() ;
            public SignDocFolder.DocPageList 頁面
            {
                get { return _頁面; }
                set { _頁面 = value; }
            }

            private AttachmentsDesc.AttsList _附件清單 = null;
            public AttachmentsDesc.AttsList 附件清單
            {
                get { return _附件清單; }
                set { _附件清單 = value; }
            }


            #region 來文介面 成員

            XmlNode 來文介面.getInDocNode(string genetime)
            {
                XmlNode rtnNode = xmlTool.MakeNode("紙本來文", "頁面數", _頁面.Count.ToString() );
                foreach (XmlNode anode in _頁面.getAllDocPageNode(genetime))
                    rtnNode.AppendChild(anode.Clone());
                if (_附件清單 != null)
                {
                    rtnNode.AppendChild(_附件清單.getAttsListNode());
                }
                return rtnNode;
            }

            void 來文介面.AddAtt(AttachmentsDesc.IAttachment att)
            {
                if (_附件清單 == null) _附件清單 = new AttachmentsDesc.AttsList();
                _附件清單.Add(att);   
            }

            void 來文介面.AddPage(SignDocFolder.DocPageList.Page page)
            {
                _頁面.Add(page); 
            }

            AttachmentsDesc.AttsList 來文介面.get附件清單()
            {
                return _附件清單; 
            }

            #endregion
        }

        internal 來文介面 indoc;
        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="ID">xml tag id,預設給DOC_0001</param>
        /// <param name="inDocNO">來文字號</param>
        /// <param name="indocType">來文類型 例:函1</param>
        /// <param name="format">電子/紙本</param>
        public ReceivedDoc(string ID , string inDocNO, string indocType,string format)
        {
            _Id = ID == "" ? "DOC_0001" : ID;
            _parDocNO = inDocNO;
            _DocFormat = format;
            _indocType = indocType;
            if (_DocFormat == "電子")
            {
                indoc = new 電子來文();
            }
            else
            {
                indoc = new 紙本來文();
            }
        }

        /// <summary>
        /// 來文文件夾參照路徑建構子
        /// </summary>
        /// <param name="ID"></param>
        public ReceivedDoc(string RefID)
        {
            _Id = RefID;
            RefDoc = true;
        }

        /// <summary>
        /// 取得來文文件夾節點內容
        /// </summary>
        /// <returns></returns>
        public XmlNode getinDocFolder()
        {
            if (!RefDoc)
            {

                Dictionary<string, string> atts = new Dictionary<string, string>();
                atts.Clear();
                atts.Add("格式", _DocFormat);
                atts.Add("序號", _serialNo);
                atts.Add("產生時間", _geneTime);
                XmlNode rtnNode = xmlTool.MakeNode("來文", atts);
                XmlNode xnd = xmlTool.MakeNode("來文字號", "");
                xnd.AppendChild(xmlTool.MakeNode("文字", _parDocNO));
                rtnNode.AppendChild(xnd);
                rtnNode.AppendChild(xmlTool.MakeNode("來文類型", _indocType));
                rtnNode.AppendChild(indoc.getInDocNode(_geneTime));
                atts.Clear();
                atts.Add("Id", _Id);
                atts.Add("產生時間", _geneTime);
                XmlNode inXnd = xmlTool.MakeNode("來文文件夾", atts);
                XmlNode inList = xmlTool.MakeNode("來文清單", "來文數", _DocCount);
                inList.AppendChild(rtnNode);
                inXnd.AppendChild(inList);
                return inXnd;

            }
            else
                return xmlTool.MakeNode("來文文件夾參照路徑", "URI", "#" + _Id);
        }

    }
}
