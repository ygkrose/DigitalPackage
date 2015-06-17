using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.ModifyInfo
{
    //異動資訊類別
    public class ModifyInfo 
    {
        static string TagName = "異動資訊";
        //簽核人員
        public class signer
        {
            private string _單位 = "";

            internal string 單位
            {
                get { return _單位; }
                set { _單位 = value; }
            }

            private string _單位代碼 = "";

            internal string 單位代碼
            {
                get { return _單位代碼; }
                set { _單位代碼 = value; }
            }

            private string _職稱 = "";

            internal string 職稱
            {
                get { return _職稱; }
                set { _職稱 = value; }
            }
            private string _姓名 = "";

            internal string 姓名
            {
                get { return _姓名; }
                set { _姓名 = value; }
            }
            private string _帳號 = "";

            internal string 帳號
            {
                get { return _帳號; }
                set { _帳號 = value; }
            }

            private string _借卡時間 = null;

            internal string 借卡時間
            {
                get { return _借卡時間; }
                set { _借卡時間 = value; }
            }

            private string _借卡原因 = null;

            internal string 借卡原因
            {
                get { return _借卡原因; }
                set { _借卡原因 = value; }
            }
            /// <summary>
            /// 建構子
            /// </summary>
            public signer(string 單位,string 單位代碼 ,string 職稱,string 姓名,string 帳號) 
            {
                _單位 = 單位;
                _單位代碼 = 單位代碼;
                _職稱 = 職稱;
                _姓名 = 姓名;
                _帳號 = 帳號;
            }

            /// <summary>
            /// 設定使用臨時卡資訊
            /// </summary>
            /// <param name="borrowCardTime">借卡時間</param>
            /// <param name="reason">借卡原因</param>
            public void setUseTempCard(string borrowCardTime, string reason)
            {
                _借卡時間 = borrowCardTime;
                _借卡原因 = reason;
            }

            /// <summary>
            /// 取得簽核人員節點資料
            /// </summary>
            /// <returns></returns>
            internal XmlNode getSignerNodeXML()
            {
                XmlNode aNode =xmlTool.MakeNode("簽核人員", "");
                aNode.AppendChild(xmlTool.MakeNode("單位", _單位));
                aNode.AppendChild(xmlTool.MakeNode("職稱", _職稱));
                aNode.AppendChild(xmlTool.MakeNode("姓名", _姓名));
                aNode.AppendChild(xmlTool.MakeNode("帳號", _帳號));
                return aNode;
            }

        }

        private FlowInfo.FlowType _異動別;

        public FlowInfo.FlowType 異動別
        {
            get { return _異動別; }
            set { _異動別 = value; }
        }
        private string _簽核意見 = "";
        private string _簽章時間 = "";
        private signer _sender;
        private signer _receiver = null;
        

        /// <summary>
        /// 建立異動資訊內容
        /// </summary>
        /// <param name="sender">傳送方</param>
        /// <param name="receiver">接收方</param>
        /// <param name="cType">異動別</param>
        /// <param name="comment">簽核意見</param>
        /// <param name="signTime">簽章時間</param>
        public ModifyInfo(signer sender, signer receiver,FlowInfo.FlowType cType,string comment,DateTime signTime)
        {
            _sender = sender;
            _receiver = receiver;
            _異動別 = cType;
            _簽核意見 = comment;
            _簽章時間 = (signTime.Year - 1911).ToString("000") + signTime.Month.ToString("00") + signTime.Day.ToString("00") + signTime.Hour.ToString("00") + signTime.Minute.ToString("00") ;          
        }

        public ModifyInfo(signer sender, signer receiver, FlowInfo.FlowType cType, string comment, string signTime)
        {
            _sender = sender;
            _receiver = receiver;
            _異動別 = cType;
            _簽核意見 = comment;
            _簽章時間 = signTime;
        }

        public ModifyInfo(signer sender, FlowInfo.FlowType cType, string comment, string signTime)
        {
            _sender = sender;
            _異動別 = cType;
            _簽核意見 = comment;
            _簽章時間 = signTime;
        }


        /// <summary>
        /// 取得異動資訊XML資料
        /// </summary>
        /// <returns></returns>
        public XmlNode getModifyInfoNode()
        {
            XmlNode aNode = xmlTool.MakeNode(TagName, "");
            aNode.AppendChild(_sender.getSignerNodeXML());
            aNode.AppendChild(xmlTool.MakeNode("異動別", _異動別.ToString()));
            aNode.AppendChild(xmlTool.MakeNode("簽核意見", _簽核意見));
            aNode.AppendChild(xmlTool.MakeNode("簽章時間", _簽章時間));
            if (_receiver != null)
            {
                XmlNode bNode = xmlTool.MakeNode("次位簽核人員", "");
                bNode.AppendChild(_receiver.getSignerNodeXML());
                aNode.AppendChild(bNode);
            }
            if (_sender.借卡時間 != null)
            {
                XmlAttribute atb =  aNode.OwnerDocument.CreateAttribute("借卡");
                atb.Value = "Y";
                aNode.Attributes.Append(atb);   
                XmlNode cNode = xmlTool.MakeNode("借卡", "");
                cNode.AppendChild(xmlTool.MakeNode("借卡日期", _sender.借卡時間));
                cNode.AppendChild(xmlTool.MakeNode("借卡原因", _sender.借卡原因));
                aNode.AppendChild(cNode);
            }
            return aNode;
        }

    }
}
