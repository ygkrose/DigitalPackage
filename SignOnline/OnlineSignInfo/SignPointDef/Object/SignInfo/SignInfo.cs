using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo
{
    /// <summary>
    /// 簽核資訊 : 簽核資訊 ((來文文件夾|來文文件夾參照路徑)? , (簽核文件夾|子文簽核文件夾))
    /// </summary>
    public class SignInfo
    {
        static string TagName = "簽核資訊";

        private ReceivedDoc.ReceivedDoc _來文文件夾 = null;
        private SignDocFolder.SignDocFolder _簽核文件夾;

        public SignDocFolder.SignDocFolder 簽核文件夾
        {
            get { return _簽核文件夾; }
            set { _簽核文件夾 = value; }
        }

        public SignInfo(SignDocFolder.SignDocFolder 簽核文件夾)
        {
            _簽核文件夾 = 簽核文件夾;
        }

        public SignInfo(SignDocFolder.SignDocFolder 簽核文件夾, ReceivedDoc.ReceivedDoc 來文文件夾)
        {
            _簽核文件夾 = 簽核文件夾;
            _來文文件夾 = 來文文件夾;
        }

        /// <summary>
        /// 取得簽核資訊節點資料
        /// </summary>
        /// <returns></returns>
        public XmlNode getSignInfoNode()
        {
            XmlNode rtnnode = xmlTool.MakeNode(TagName, "");
            if (_來文文件夾 != null)
                rtnnode.AppendChild(_來文文件夾.getinDocFolder());
            rtnnode.AppendChild(_簽核文件夾.getSignDocFolderNode());
            return rtnnode;
        }
    }
}
