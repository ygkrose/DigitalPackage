using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed;
using DigitalSealed.Tools;

namespace DigitalSealed.AttachmentsDesc
{
    /// <summary>
    /// 附件清單
    /// </summary>
    public class AttsList : List<IAttachment>
    {
        private string _attCount = "";

        public AttsList()
        {
            
        }

        /// <summary>
        /// 回傳附件清單標籤內容
        /// </summary>
        /// <returns></returns>
        public XmlNode getAttsListNode()
        {
            _attCount = this.Count.ToString();
            XmlNode rtnNode = xmlTool.MakeNode("附件清單", "附件數", _attCount);
            if (this[0].getAttKindNode().Name == "電子檔格式附件")
            {
                foreach (IAttachment ainfo in this)
                {
                    XmlNode childs = null;
                    Dictionary<string, string> attr = new Dictionary<string, string>();
                    attr.Add("格式", ainfo.getAttType() + "文件");
                    attr.Add("產生時間", ainfo.getGeneTime());
                    childs = xmlTool.MakeNode("附件", attr);
                    childs.AppendChild(xmlTool.MakeNode("名稱", ainfo.getAttDesc()));
                    childs.AppendChild(xmlTool.MakeNode("附件類型", (ainfo as ElecAtt).AttType));
                    childs.AppendChild(ainfo.getAttKindNode());
                    rtnNode.AppendChild(childs);
                }
            }
            else if (this[0].getAttKindNode().Name == "一般頁面檔格式附件")
            {
                XmlNode childs = null;
                Dictionary<string, string> attr = new Dictionary<string, string>();
                attr.Add("格式", "紙本掃描文件");
                attr.Add("產生時間", this[0].getGeneTime());
                childs = xmlTool.MakeNode("附件", attr);
                childs.AppendChild(xmlTool.MakeNode("名稱", "紙本附件掃描"));
                childs.AppendChild(xmlTool.MakeNode("附件類型", "TIF"));
                XmlNode 頁面檔 = xmlTool.MakeNode("一般頁面檔格式附件", "頁面數", this._attCount);
                int s = 1;
                foreach (ElecAtt ainfo in this)
                {
                    Dictionary<string, string> attrs = new Dictionary<string, string>();
                    attrs.Add("原始檔序號", ainfo.Ofileserial);
                    attrs.Add("序號", s.ToString() );
                    attrs.Add("產生時間", ainfo.GeneTime);
                    頁面檔.AppendChild(xmlTool.MakeNode("頁面", attrs));
                    s++;
                }
                childs.AppendChild(頁面檔); 
                rtnNode.AppendChild(childs);
            }
            return rtnNode;
        }
    }

  
}
