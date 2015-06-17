using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using DigitalSealed.Tools;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature;

namespace DigitalSealed.OA
{
    /// <summary>
    /// 移轉交封裝檔
    /// </summary>
    public class ETransfer
    {
        private string _移轉交封裝檔SigID = "ertransfer";

        public string 移轉交封裝檔SigID
        {
            get { return _移轉交封裝檔SigID; }
            set { _移轉交封裝檔SigID = value; }
        }
        //媒體封裝檔位置
        private string _MediaFilepath = "";

        private string _PKCS11Driver = "";

        public string PKCS11Driver
        {
            get { return _PKCS11Driver; }
            set { _PKCS11Driver = value; }
        }

        private string _HashAlgorithm = "SHA1"; //or SHA256
        /// <summary>
        /// 選用雜湊演算法目前只支援SHA1, SHA256
        /// </summary>
        public string HashAlgorithm
        {
            get { return _HashAlgorithm; }
            set { _HashAlgorithm = value; }
        }

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="媒體封裝檔">媒體封裝檔路徑檔名</param>
        public ETransfer(string 媒體封裝檔)
        {
            _MediaFilepath = 媒體封裝檔;
            if (!File.Exists(_MediaFilepath)) throw new Exception("媒體封裝檔不存在!");
        }


        public XmlDocument MakeTransferXml(string[] 電子媒體編號,string pwd)
        {
            if (電子媒體編號.Length < 1) throw new Exception("無媒體編號可移轉交!");
            XmlDocument DcpXml = new XmlDocument();
            DcpXml.XmlResolver = null;
            DcpXml.PreserveWhitespace = true;
            DcpXml.AppendChild(DcpXml.CreateXmlDeclaration("1.0", "utf-8", null));
            DcpXml.AppendChild(DcpXml.CreateDocumentType("移轉封裝檔", null, "99_ertransfer_utf8.dtd", null));
            XmlNode transfer = xmlTool.MakeNode("移轉封裝檔", "");
             
            XmlNode objtag = xmlTool.MakeNode("Object", "Id", "MediaObjId");
            objtag.AppendChild(xmlTool.MakeNode("電子媒體編號清單", ""));

            foreach (string idno in 電子媒體編號)
            {
                objtag.FirstChild.AppendChild(xmlTool.MakeNode("電子媒體編號", idno));
            }

            SignatureTag st = new SignatureTag(_移轉交封裝檔SigID);
            st.HashAlgorithm = _HashAlgorithm;
            List<string> refs = new List<string>();
            refs.Add(new FileInfo(_MediaFilepath).Name);
            refs.Add("#MediaObjId"); 
            XmlNode signaturetag = st.getSignatureNode(refs);
            signaturetag.AppendChild(objtag);
            XmlDocument tmpdoc = new XmlDocument();
            tmpdoc.AppendChild(tmpdoc.ImportNode(signaturetag,true));
            Signature sign = new Signature(tmpdoc);
            if (_PKCS11Driver != "") sign.PKCS11Driver = _PKCS11Driver;
            sign.FilePath = new FileInfo(_MediaFilepath).DirectoryName;
            XmlNode xnd = (XmlNode)sign.getSignXmlData("ertransfer", pwd).DocumentElement;
            XmlNode keyinfo = xnd.SelectSingleNode("//KeyInfo").Clone();
            xnd.RemoveChild(xnd.SelectSingleNode("//KeyInfo"));
            xnd.InsertAfter(keyinfo, xnd.SelectSingleNode("//SignatureValue"));  
            transfer.AppendChild(transfer.OwnerDocument.ImportNode(xnd,true));
            DcpXml.AppendChild(DcpXml.ImportNode(transfer,true) );
            return DcpXml;
        }

    }
}
