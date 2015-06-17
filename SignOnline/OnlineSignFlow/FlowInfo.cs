using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;

namespace DigitalSealed
{
    /// <summary>
    /// 簽核流程
    /// </summary>
    public class FlowInfo
    {
        static string TagName = "簽核流程";
  
        private string _Id = ""; 

        /// <summary>
        /// 簽核流程的ID值
        /// </summary>
        public string Id
        {
            get { return _Id; }
            //set { _Id = value; }
        }

        private FlowType _modeName ;

        /// <summary>
        /// 異動別:呈核,退文,核閱,決行....
        /// </summary>
        public FlowType ModeName
        {
            get { return _modeName; }
            //set { _異動別 = value; }
        }

        /// <summary>
        /// 簽核流程建構子
        /// </summary>
        /// <param name="id"></param>
        /// <param name="modename"></param>
        public FlowInfo(string id, FlowType modename)
        {
            _Id = id;
            _modeName = modename; 
        }

        public FlowInfo(DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.ModifyInfo.ModifyInfo.signer signer, FlowType modename)
        {
            _Id = "Flow_" + signer.單位代碼 + "_" +signer.姓名 + "_" +  modename.ToString();
            _modeName = modename; 
        }

        /// <summary>
        /// 產生簽核流程點的Tag
        /// </summary>
        /// <returns></returns>
        public XmlNode getFlowInfoNode()
        {
            Dictionary<string, string> attb = new Dictionary<string, string>();
            attb.Add("Id", _Id);
            attb.Add("異動別", _modeName.ToString());
            XmlNode rtnNode = xmlTool.MakeNode(TagName, attb);
            //if (_modeName == FlowType.會辦 || _modeName == FlowType.會辦陳核)
            //{
            //    XmlNode splitNd = xmlTool.MakeNode("分會點", "");
            //}

            return rtnNode;
        }

        /// <summary>
        /// 定義分會點
        /// </summary>
        /// <returns></returns>
        private XmlNode geneDepFlowInfoNode(FlowInfo[] Depfi)
        {
            return null; 
        }

        /// <summary>
        /// 異動別種類
        /// </summary>
        public enum FlowType
        {
            呈核,
            核閱,
            批示,
            決行,
            會辦,
            並會,
            後會,
            彙併辦,
            解併,
            補簽,
            清稿,
            完稿,
            並會整併,
            其他
        }
    }
}
