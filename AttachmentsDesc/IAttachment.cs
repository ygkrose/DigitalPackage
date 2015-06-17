using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace DigitalSealed.AttachmentsDesc
{
    public interface IAttachment
    {
        /// <summary>
        /// 取得附件型態 : 電子/實體
        /// </summary>
        /// <returns></returns>
        string getAttType();

        /// <summary>
        /// 取得產生時間
        /// </summary>
        /// <returns></returns>
        string getGeneTime();

        /// <summary>
        /// 取得附件名稱
        /// </summary>
        /// <returns></returns>
        string getAttDesc();

        /// <summary>
        /// 取得附件(電子檔格式附件|一般頁面檔格式附件|文稿頁面檔格式附件)節點
        /// </summary>
        /// <returns></returns>
        XmlNode getAttKindNode();

    }
}
