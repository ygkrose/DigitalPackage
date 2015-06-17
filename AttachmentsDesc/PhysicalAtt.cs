using System;
using System.Collections.Generic;
using System.Text;

namespace DigitalSealed.AttachmentsDesc
{
    /// <summary>
    /// 實體附件
    /// </summary>
    public class PhysicalAtt : IAttachment 
    {
        private string _attDescName = "";

        public string AttDescName
        {
            get { return _attDescName; }
            set { _attDescName = value; }
        }

        private string _geneTime = "";

        public string GeneTime
        {
            get { return _geneTime; }
            set { _geneTime = value; }
        }

        public PhysicalAtt(string attName,string genetime)
        {
            _attDescName = attName;
            _geneTime = genetime; 
        }


        #region IAttachment 成員

        string IAttachment.getAttType()
        {
            return "實體";
        }


        string IAttachment.getGeneTime()
        {
            return _geneTime;
        }

        /// <summary>
        /// 產生"一般頁面檔格式附件" or "文稿頁面檔格式附件"標籤集
        /// </summary>
        /// <returns></returns>
        System.Xml.XmlNode IAttachment.getAttKindNode()
        {
            throw new NotImplementedException();
        }


        string IAttachment.getAttDesc()
        {
            return _attDescName;
        }

        #endregion
    }
}
