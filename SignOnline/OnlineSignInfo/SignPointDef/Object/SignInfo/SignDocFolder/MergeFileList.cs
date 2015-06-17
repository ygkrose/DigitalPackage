using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Xml;
using DigitalSealed.Tools;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.SignDocFolder
{
    /// <summary>
    /// 併文清單
    /// </summary>
    public class MergeFileList
    {

        private string _MotherDocNo = "";
        /// <summary>
        /// 母文文號
        /// </summary>
        public string MotherDocNo
        {
            get { return _MotherDocNo; }
            set { _MotherDocNo = value; }
        }

        private List<string> _ChildsDocNO = new List<string>();

        /// <summary>
        /// 併文清單建構子
        /// </summary>
        /// <param name="MDocNO">母文文號</param>
        /// <param name="CDocNO">子文文號</param>
        public MergeFileList(string MDocNO,List<string> CDocNO)
        {
            _MotherDocNo = MDocNO;
            _ChildsDocNO = CDocNO;   
        }

        /// <summary>
        /// 取得併文清單節點
        /// </summary>
        /// <returns></returns>
        public XmlNode getMergeFileListNode()
        {
            XmlNode rtn = xmlTool.MakeNode("併文清單", "");
            rtn.AppendChild(xmlTool.MakeNode("母文", _MotherDocNo));
            foreach (string s in _ChildsDocNO)
                 rtn.AppendChild(xmlTool.MakeNode("子文", s));
            return rtn;
        }

    }
}
