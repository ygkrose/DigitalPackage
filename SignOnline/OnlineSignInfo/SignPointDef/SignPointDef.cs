using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef
{
    /// <summary>
    /// 簽核點定義
    /// </summary>
    public class SignPointDef
    {
        private string _uri = "";
        private string _Id = "";
        private SignatureTag _SignTag = null;
        private List<string> _preSignDefId = new List<string>();

        public List<string> RefSignDefId
        {
            get { return _preSignDefId; }
            set { _preSignDefId = value; }
        }

        public SignatureTag SignTag
        {
            get { return _SignTag; }
            set { _SignTag = value; }
        }
        private ObjectTag _ObjectTag = null;

        public ObjectTag ObjectTag
        {
            get { return _ObjectTag; }
            set { _ObjectTag = value; }
        }

        /// <summary>
        /// 簽核點定義的建構子
        /// </summary>
        /// <param name="flowInfoId">簽核流程Id</param>
        /// <param name="preSignDefId">上一簽核點定義Id</param>
        /// <param name="st"></param>
        /// <param name="ot"></param>
        public SignPointDef(string flowInfoId,List<string> preSignDefId,SignatureTag st,ObjectTag ot)
        {
            _uri = "#" + flowInfoId;
            _Id = "sign_" + flowInfoId;
            _SignTag = st;
            _ObjectTag = ot;
            _preSignDefId = preSignDefId;
        }

        public SignPointDef(string flowInfoId)
        {
            _uri = "#" + flowInfoId;
            _Id = "sign_" + flowInfoId;
        }

        /// <summary>
        /// 取得簽核點定義的標籤內容
        /// </summary>
        /// <returns></returns>
        public XmlNode getSignPointNode()
        {
            Dictionary<string, string> att = new Dictionary<string, string>();
            att.Add("URI", _uri);
            att.Add("Id", _Id);
            XmlNode xnd = xmlTool.MakeNode("簽核點定義", att);
            if (_SignTag != null)
                xnd.AppendChild(_SignTag.getSignatureNode(_preSignDefId));
            if (_ObjectTag != null)
                xnd.AppendChild(_ObjectTag.getObjectNode());
            return xnd;
        }


    }
}
