using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using DigitalSealed.Tools;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object;
using DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Object.SignInfo.SignDocFolder;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature
{
    /// <summary>
    /// Signature標籤
    /// </summary>
    public class SignatureTag
    {
        private string _Id = "";

        public string Id
        {
            get { return _Id; }
            set { _Id = value; }
        }

        private ObjectTag _ObjectNode = null;

        public ObjectTag ObjectNode
        {
            get { return _ObjectNode; }
            set { _ObjectNode = value; }
        }

        private string _HashAlgorithm = "SHA1"; //or SHA256 or SHA512
        /// <summary>
        /// 選用雜湊演算法目前只支援SHA1, SHA256, SHA512
        /// </summary>
        public string HashAlgorithm
        {
            get { return _HashAlgorithm; }
            set { _HashAlgorithm = value; }
        }

        private FilesList _檔案清單 = null;

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="id">Signature Id值</param>
        /// <param name="objectNode">要處理的Object節點</param>
        public SignatureTag(string id, ObjectTag objectNode)
        {
            _Id = id;
            _ObjectNode = objectNode;
            _檔案清單 = objectNode.簽核資訊.簽核文件夾.檔案清單;
        }

        public SignatureTag(string id)
        {
            _Id = id;
        }

        /// <summary>
        /// 取得尚未加簽的Signature節點內容
        /// </summary>
        /// <param name="preSignDefId">上一簽核點定義Id值</param>
        /// <returns></returns>
        public XmlNode getSignatureNode(List<string> preSignDefId)
        {
            XmlNode rtnNode = xmlTool.MakeNode("Signature", "Id", _Id);
            XmlNode _SignedInfo = xmlTool.MakeNode("SignedInfo", "");
            _SignedInfo.AppendChild(xmlTool.MakeNode("CanonicalizationMethod", "Algorithm", "http://www.w3.org/TR/2001/REC-xml-c14n-20010315"));
            if (_HashAlgorithm == "SHA1")
                _SignedInfo.AppendChild(xmlTool.MakeNode("SignatureMethod", "Algorithm", "http://www.w3.org/2000/09/xmldsig#rsa-sha1"));
            else if (_HashAlgorithm == "SHA256")
                _SignedInfo.AppendChild(xmlTool.MakeNode("SignatureMethod", "Algorithm", "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"));
            else if (_HashAlgorithm == "SHA512")
                _SignedInfo.AppendChild(xmlTool.MakeNode("SignatureMethod", "Algorithm", "http://www.w3.org/2001/04/xmldsig-more#rsa-sha512"));
            if (_檔案清單 != null)
            {
                foreach (ElecFileInfo f in _檔案清單)
                {
                    XmlNode refnode = xmlTool.MakeNode("Reference", "URI", f.Filename);
                    if (_HashAlgorithm == "SHA1")
                        refnode.AppendChild(xmlTool.MakeNode("DigestMethod", "Algorithm", "http://www.w3.org/2000/09/xmldsig#sha1"));
                    else if (_HashAlgorithm == "SHA256")
                        refnode.AppendChild(xmlTool.MakeNode("DigestMethod", "Algorithm", "http://www.w3.org/2001/04/xmlenc#sha256"));
                    else if (_HashAlgorithm == "SHA512")
                        refnode.AppendChild(xmlTool.MakeNode("DigestMethod", "Algorithm", "http://www.w3.org/2001/04/xmlenc#sha512"));
                    refnode.AppendChild(xmlTool.MakeNode("DigestValue", ""));
                    _SignedInfo.AppendChild(refnode);
                }
            }
            //加入上一個簽核點定義的tag
            foreach (string refid in preSignDefId)
            {
                XmlNode refprend = xmlTool.MakeNode("Reference", "URI", refid);
                if (_HashAlgorithm == "SHA1")
                    refprend.AppendChild(xmlTool.MakeNode("DigestMethod", "Algorithm", "http://www.w3.org/2000/09/xmldsig#sha1"));
                else if (_HashAlgorithm == "SHA256")
                    refprend.AppendChild(xmlTool.MakeNode("DigestMethod", "Algorithm", "http://www.w3.org/2001/04/xmlenc#sha256"));
                else if (_HashAlgorithm == "SHA512")
                    refprend.AppendChild(xmlTool.MakeNode("DigestMethod", "Algorithm", "http://www.w3.org/2001/04/xmlenc#sha512"));
                refprend.AppendChild(xmlTool.MakeNode("DigestValue", ""));
                _SignedInfo.AppendChild(refprend); 
            }
            if (_檔案清單 != null)
            {
                //加入目前簽核點定義的tag
                XmlNode refnd = xmlTool.MakeNode("Reference", "URI", "#" + _ObjectNode.Id);
                if (_HashAlgorithm == "SHA1")
                    refnd.AppendChild(xmlTool.MakeNode("DigestMethod", "Algorithm", "http://www.w3.org/2000/09/xmldsig#sha1"));
                else if (_HashAlgorithm == "SHA256")
                    refnd.AppendChild(xmlTool.MakeNode("DigestMethod", "Algorithm", "http://www.w3.org/2001/04/xmlenc#sha256"));
                else if (_HashAlgorithm == "SHA512")
                    refnd.AppendChild(xmlTool.MakeNode("DigestMethod", "Algorithm", "http://www.w3.org/2001/04/xmlenc#sha512"));
                refnd.AppendChild(xmlTool.MakeNode("DigestValue", ""));
                _SignedInfo.AppendChild(refnd);
            }
            rtnNode.AppendChild(_SignedInfo);  
            rtnNode.AppendChild(xmlTool.MakeNode("SignatureValue",""));
            XmlNode tmpNde = xmlTool.MakeNode("KeyInfo", "");
            XmlNode x509 = xmlTool.MakeNode("X509Data", "");
            x509.AppendChild(xmlTool.MakeNode("X509Certificate", ""));
            tmpNde.AppendChild(x509);
            rtnNode.AppendChild(tmpNde);
            return rtnNode;
        }
    }
}
