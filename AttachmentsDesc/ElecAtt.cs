using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using DigitalSealed.Tools;

namespace DigitalSealed.AttachmentsDesc
{
    /// <summary>
    /// 電子附件
    /// </summary>
    public class ElecAtt : IAttachment 
    {
        private string _attName = "";
        /// <summary>
        /// 附件檔名
        /// </summary>
        public string AttName
        {
            get { return _attName; }
            set { _attName = value; }
        }
        
        private string _attType = "";
        /// <summary>
        /// 附件類型
        /// </summary>
        public string AttType
        {
            get { return _attType; }
            set { _attType = value; }
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
        private string _attKind = ""; //電子檔格式附件|一般頁面檔格式附件|文稿頁面檔格式附件

        public string AttFormatKind
        {
            get { return _attKind; }
            set { _attKind = value; }
        }

        private string _ofileserial = ""; //原始檔序號

        public string Ofileserial
        {
            get { return _ofileserial; }
            set { _ofileserial = value; }
        }


        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="attname">附件名稱</param>
        /// <param name="genetime">產生時間</param>
        /// <param name="attKind">附件種類</param>
        public ElecAtt(string attname, string genetime, AttachStyle astyle, string ofileserial)
        {
            _attName = attname;
            _geneTime = genetime;
            if (_attName.LastIndexOf(".") > -1)
                _attType = _attName.Substring(_attName.LastIndexOf(".") + 1);
            else
                _attType = "";
            _attKind = astyle.ToString() ;
            _ofileserial = ofileserial;
        }

        public ElecAtt(string attname, string genetime)
        {
            _attName = attname;
            _geneTime = genetime;
            if (_attName.LastIndexOf(".") > -1)
                _attType = _attName.Substring(_attName.LastIndexOf(".") + 1);
            else
                _attType = "";
            _attKind = AttachStyle.電子檔格式附件.ToString()  ;
        }

        #region IAttachment 成員

        string IAttachment.getAttType()
        {
            return "電子";
        }

 
        string IAttachment.getGeneTime()
        {
            return _geneTime;
        }


        /// <summary>
        /// 產生電子檔格式附件標籤集
        /// </summary>
        /// <returns></returns>
        XmlNode IAttachment.getAttKindNode()
        {
            return xmlTool.MakeNode(_attKind, "原始檔序號", _ofileserial);  
        }

        string IAttachment.getAttDesc()
        {
            return _attName;
        }

        #endregion


    }
}
