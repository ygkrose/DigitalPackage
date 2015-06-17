using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object
{
    /// <summary>
    /// 簽核點定義裡的Object標籤集
    /// </summary>
    public class ObjectTag
    {
        private string _Id = "";

        public string Id
        {
            get { return _Id; }
            set { _Id = value; }
        }
        private ModifyInfo.ModifyInfo _異動資訊;
        private SignInfo.SignInfo _簽核資訊;

        public SignInfo.SignInfo 簽核資訊
        {
            get { return _簽核資訊; }
            set { _簽核資訊 = value; }
        }

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="ID">ID屬性值</param>
        /// <param name="異動資訊">異動資訊</param>
        /// <param name="簽核資訊">簽核資訊</param>
        public ObjectTag(string ID,ModifyInfo.ModifyInfo 異動資訊, SignInfo.SignInfo 簽核資訊)
        {
            _Id = ID;
            _異動資訊 = 異動資訊;
            _簽核資訊 = 簽核資訊;
        }

        /// <summary>
        /// 取得Object Tag的XML內容
        /// </summary>
        /// <returns></returns>
        public XmlNode getObjectNode()
        {
            XmlNode rtnNode = xmlTool.MakeNode("Object", "Id", _Id);
            rtnNode.AppendChild(_異動資訊.getModifyInfoNode());
            rtnNode.AppendChild(_簽核資訊.getSignInfoNode());
            return rtnNode;
        }

        public FlowInfo.FlowType get異動資訊()
        {
            return _異動資訊.異動別;
        }
    }
}
