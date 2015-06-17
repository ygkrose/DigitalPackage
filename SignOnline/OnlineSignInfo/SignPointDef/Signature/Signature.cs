using System;
using System.Collections.Generic;
using System.Text;
using FSXMLCAPIATLLib;
using FSGPKICRYPTATLLib;
using System.Xml;
using System.IO;
using System.Security.Cryptography.Xml;
using System.Security.Cryptography;
using System.Collections;
using System.Security.Cryptography.X509Certificates;

namespace DigitalSealed.SignOnline.OnlineSignInfo.SignPointDef.Signature
{
    /// <summary>
    /// 加簽類別
    /// </summary>
    public class Signature 
    {
        private ArrayList arrNotPassDescs = new ArrayList();
        private ArrayList arrNotPassItems = new ArrayList();

        private string _FilePath = "";
        /// <summary>
        /// 所有檔案存放路徑
        /// </summary>
        public string FilePath
        {
            get { return _FilePath; }
            set { _FilePath = value; }
        }

        private string _XmlFileName = "";
        /// <summary>
        /// si檔檔名
        /// </summary>
        public string XmlFileName
        {
            get { return _XmlFileName; }
            set { _XmlFileName = value; }
        }

        private int _cardReaderSlot = 0;
        /// <summary>
        /// 設定要使用第幾台讀卡機,預設第一台
        /// </summary>
        public int CardReaderSlot
        {
            get { return _cardReaderSlot; }
            set { _cardReaderSlot = value; }
        }


        private string _PKCS11Driver = "";
        /// <summary>
        /// 台網元件dll檔名
        /// </summary>
        public string PKCS11Driver
        {
            set { _PKCS11Driver = value; }
        }

        private string _CertX509Str = "";
        /// <summary>
        /// 暫存x509憑證字串
        /// </summary>
        public string CertX509Str
        {
            get { return _CertX509Str; }
            set { _CertX509Str = value; }
        }

        //FSXMLATLClass fsxml = new FSXMLATLClass();
        FSGPKICRYPTATLLib.GPKICryptATLClass gpkixml = new GPKICryptATLClass();
        private XmlDocument _XmlForSign = null;
        
        /// <summary>
        /// 建構子
        /// </summary>
        public Signature()
        {         
        }

        /// <summary>
        /// 建構子
        /// </summary>
        /// <param name="filepath">需要封裝的檔案路徑</param>
        /// <param name="filename">si檔檔名</param>
        public Signature(string filepath,string filename)
        {
            _FilePath = filepath;
            _XmlFileName = filename;
        }

        /// <summary>
        /// 建構子(路徑取用ConstVar.FileWorkingPath)
        /// </summary>
        /// <param name="filename"></param>
        public Signature(string filename)
        {
            _FilePath = ConstVar.SourceFilesPath;
            _XmlFileName = filename;
        }

        /// <summary>
        ///  建構子
        /// </summary>
        /// <param name="xmlforsign">預加簽的XML資料</param>
        public Signature(XmlDocument xmlforsign)
        {
            _XmlForSign = xmlforsign;
            _XmlForSign.PreserveWhitespace = true; 
        }


        /// <summary>
        /// 取得加完簽的完整XML
        /// </summary>
        /// <param name="SignatureId">預加簽節點Id</param>
        /// <param name="pinCode">憑證密碼</param>
        /// <returns></returns>
        public XmlDocument getSignXmlData(string SignatureId, string pinCode)
        {
            string hashstr = CalcFileHash(SignatureId);
            //DigitalSealed.Tools.geneTime.getCounter("計算雜湊"); 
            if (hashstr.StartsWith("ERR:")) throw new Exception("雜湊發生例外:" + hashstr);
            string signstr = "";
            if (pinCode != "")
                signstr = SignXmlNode(hashstr, SignatureId, pinCode);
            else
                signstr = hashstr;
            //DigitalSealed.Tools.geneTime.getCounter("加簽"); 
            if (signstr.StartsWith("ERR:")) throw new Exception("加簽發生例外:" + signstr);
            XmlDocument xdoc = new XmlDocument();
            xdoc.XmlResolver = null;
            xdoc.LoadXml(signstr);
            xdoc.PreserveWhitespace = true;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //DigitalSealed.Tools.geneTime.getCounter("回傳加簽結果"); 
            return xdoc;
        }

        /// <summary>
        /// 只處理XML加簽(client端點收用)
        /// </summary>
        /// <param name="SignatureId">預加簽節點Id</param>
        /// <param name="pinCode">憑證密碼</param>
        /// <returns></returns>
        public XmlDocument SignOnlyXmlData(string SignatureId, string pinCode)
        {
            string signstr = "";
            if (pinCode != "")
                signstr = SignXmlNode(_XmlForSign.OuterXml , SignatureId, pinCode);
            else
                signstr = _XmlForSign.OuterXml;
            if (signstr.StartsWith("ERR:")) throw new Exception("Client加簽發生例外:" + signstr);
            XmlDocument xdoc = new XmlDocument();
            xdoc.XmlResolver = null;
            xdoc.LoadXml(signstr);
            xdoc.PreserveWhitespace = true;
            GC.Collect();
            GC.WaitForPendingFinalizers();

            return xdoc;
        }

        public string getHashXmlData(string SignatureId)
        {
            string hashstr = CalcFileHash(SignatureId);
            if (hashstr.StartsWith("ERR:")) throw new Exception("計算雜湊發生錯誤:" + hashstr);
            return hashstr;
        }

        public XmlDocument getReSignXmlData(string xmldata,string SignatureId, string pinCode)
        {
            string signstr = SignXmlNode(xmldata, SignatureId, pinCode);
            if (signstr.StartsWith("ERR:")) throw new Exception("補簽發生錯誤:" + signstr);
            XmlDocument xdoc = new XmlDocument();
            xdoc.XmlResolver = null;
            xdoc.LoadXml(signstr);
            xdoc.PreserveWhitespace = true;
            return xdoc;
        }

        /// <summary>
        /// 檢測簽章內容
        /// </summary>
        /// <returns>錯誤時回傳訊息陣列</returns>
        public List<string> VerifyXmlFile()
        {
            if (_XmlForSign == null)
            {
                _XmlForSign = new XmlDocument();
                _XmlForSign.XmlResolver = null;
                _XmlForSign.PreserveWhitespace = true;
                _XmlForSign.Load(Path.Combine(_FilePath, _XmlFileName));
            }
            arrNotPassDescs.Clear();
            foreach (XmlNode signd in _XmlForSign.DocumentElement.GetElementsByTagName("Signature"))
            {
                VerifySingleNode(signd);
            }
            List<string> err = new List<string>(); 
            for (int i = 0; i < arrNotPassDescs.Count; i++)
            {
                if (!err.Contains(arrNotPassItems[i] + "：" + arrNotPassDescs[i])) 
                    err.Add(arrNotPassItems[i] + "：" + arrNotPassDescs[i]);
            }
            return err;    
        }

        /// <summary>
        /// 檢測最後一點簽核點
        /// </summary>
        /// <returns></returns>
        public List<string> VerifyLastNode()
        {
            if (_XmlForSign == null)
            {
                _XmlForSign = new XmlDocument();
                _XmlForSign.XmlResolver = null;
                _XmlForSign.PreserveWhitespace = true;
                _XmlForSign.Load(Path.Combine(_FilePath, _XmlFileName));
            }
            arrNotPassDescs.Clear();
            XmlNodeList ns = _XmlForSign.DocumentElement.GetElementsByTagName("Signature");
            if (ns.Count > 0)
                VerifySingleNode(ns[ns.Count - 1]);
            List<string> err = new List<string>();
            for (int i = 0; i < arrNotPassDescs.Count; i++)
            {
                if (!err.Contains(arrNotPassItems[i] + ":" + arrNotPassDescs[i]))
                    err.Add(arrNotPassItems[i] + ":" + arrNotPassDescs[i]);
            }
            return err;  
        }

        /// <summary>
        /// 取得自然人憑證開始與結束日期
        /// </summary>
        /// <param name="KeyNum">0:開始日期 1:結束日期</param>
        /// <returns></returns>
        public string GetPGKIStartEndDate(int KeyNum)
        {
            //自然人憑證用IssuerName和SubjectName當抓憑證的條件
            //CA憑證用憑證序號當抓憑證的條件
            const int FS_FLAG_CERT_NOATTACH = 0x00000100;
            const int FS_FLAG_CERT_ATTACHALL = 0x00000200;
            const int FS_FLAG_VERIFY_CONTENT_ONLY = 0x00000000;
            const int FS_FLAG_VERIFY_CERTCHAIN = 0x00000001;
            const int FS_FLAG_VERIFY_CRL = 0x00000002;
            const int FS_FLAG_BASE64_ENCODE = 0x00001000;
            const int FS_FLAG_BASE64_DECODE = 0x00002000;
            const int FS_FLAG_DETACHMSG = 0x00004000;
            const int FS_FLAG_USE_FILE = 0x00000020;
            const int FS_FLAG_NOHASHOID = 0x00020000;
            const int FS_KU_DIGITAL_SIGNATURE = 0x00000080;
            int retValue = 0;
            string strErrMessage = "";
            string strCertDate = "";
            try
            {
                FSGPKICRYPTATLLib.GPKICryptATL Y = new FSGPKICRYPTATLLib.GPKICryptATL();
                //取得GPKI中的全部憑證
                string strCerts = null;
                //strCerts = Y.FSGPKI_EnumCertsToString(FS_KU_DIGITAL_SIGNATURE);

                strCerts = Y.FSGPKI_EnumCertsToString(FS_KU_DIGITAL_SIGNATURE);
                retValue = Y.get_lastError();
                if (retValue != 0)
                {
                    strErrMessage = ErrorMsgTable.GetErrMessage(retValue);
                    throw new Exception(strErrMessage);
                }

                //取得憑證開始與結束日期
                if (KeyNum == 0)
                    strCertDate = Y.FSCAPICertGetNotBefore(strCerts, 0);
                else
                    strCertDate = Y.FSCAPICertGetNotAfter(strCerts, 0);
                retValue = Y.get_lastError();
                if (retValue != 0)
                {
                    strErrMessage = ErrorMsgTable.GetErrMessage(retValue);
                    throw new Exception(strErrMessage);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return strCertDate;
        }

        public string UpdateSignatureValue(string bs64xdoc,string signature,string cert,string signid)
        {
            FSXMLATLClass fsxml = new FSXMLATLClass();
            try
            {
                return fsxml.FSXMLCAPI_UpdateSignatureValue(bs64xdoc, signid, signature, cert, 0);
            }
            catch (Exception err)
            {
                return "ErrA:" + err.Message;
            }
            finally
            {
                fsxml = null;
            }
        }

        public string GetSignedInfoByTemplateDigest(string bsXml,string sigId)
        {
            FSXMLATLClass fsxml = new FSXMLATLClass();
            string rtn = fsxml.FSXMLCAPI_GetSignedInfoByTemplateDigest(bsXml, sigId, 0);
            fsxml = null;
            return rtn;
        }

        private void VerifySingleNode(XmlNode signd)
        {
            //先驗hash值
            foreach (XmlNode chind in signd.SelectNodes("//SignedInfo/Reference"))
            {
                ValidRef(chind, _FilePath);
            }
            //再驗簽章值
            bool rtn = ValidSig(signd);
            if (!rtn)
            {
                if (signd.Attributes["Id"] == null)
                    arrNotPassItems.Add(signd.Name);
                else
                    arrNotPassItems.Add(signd.Attributes["Id"].Value);
                arrNotPassDescs.Add("驗簽失敗!");
            }
        }

        private string GetStrHash(string argString, string argAlgo)
        {
            if (argAlgo.ToUpper() == "HTTP://WWW.W3.ORG/2000/09/XMLDSIG#SHA1")
            {
                string str;
                SHA1CryptoServiceProvider provider = new SHA1CryptoServiceProvider();
                try
                {
                    byte[] bytes = Encoding.UTF8.GetBytes(argString);
                    str = Convert.ToBase64String(provider.ComputeHash(bytes));
                }
                catch (Exception exception1)
                {
                    str = "";
                }
                return str;
            }
            else if (argAlgo.ToLower() == "http://www.w3.org/2001/04/xmlenc#sha256" || argAlgo.ToLower() == "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256")
            {
                string str = "";
                try
                {
                    //SHA256Managed _sha256 = new SHA256Managed();
                    SHA256CryptoServiceProvider _sha256 = new SHA256CryptoServiceProvider(); 
                    byte[] bytes = Encoding.UTF8.GetBytes(argString);
                    str = Convert.ToBase64String(_sha256.ComputeHash(bytes));
                }
                catch (Exception exception1)
                {
                    str = "";
                }
                return str;
            }
            else if (argAlgo.ToLower() == "http://www.w3.org/2001/04/xmlenc#sha512")
            {
                string str = "";
                try
                {
                    //SHA512Managed _sha512 = new SHA512Managed();
                    SHA512CryptoServiceProvider _sha512 = new SHA512CryptoServiceProvider(); 
                    byte[] bytes = Encoding.UTF8.GetBytes(argString);
                    str = Convert.ToBase64String(_sha512.ComputeHash(bytes));
                }
                catch (Exception exception1)
                {
                    str = "";
                }
                return str;
            }

            return "";
        }

        private string GetFileHash(string argFilePath, string argAlgo)
        {
            if (argAlgo.ToUpper() == "HTTP://WWW.W3.ORG/2000/09/XMLDSIG#SHA1")
            {
                string str;
                try
                {
                    SHA1CryptoServiceProvider provider = new SHA1CryptoServiceProvider();
                    FileStream inputStream = new FileStream(argFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    byte[] inArray = provider.ComputeHash(inputStream);
                    inputStream.Close();
                    str = Convert.ToBase64String(inArray);
                }
                catch (Exception exception1)
                {
                    str = "";
                }
                return str;
            }
            else if (argAlgo.ToLower() == "http://www.w3.org/2001/04/xmlenc#sha256" || argAlgo.ToLower() == "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256")
            {
                string str = "";
                try
                {
                    //SHA256 provider = new SHA256Managed(); 
                    SHA256CryptoServiceProvider provider = new SHA256CryptoServiceProvider();
                    FileStream inputStream = new FileStream(argFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    byte[] inArray = provider.ComputeHash(inputStream);
                    inputStream.Close();
                    str = Convert.ToBase64String(inArray);
                }
                catch (Exception exception1)
                {
                    str = "";
                }
                return str;
            }
            else if (argAlgo.ToLower() == "http://www.w3.org/2001/04/xmlenc#sha512")
            {
                string str = "";
                try
                {
                    //SHA512Managed _sha512 = new SHA512Managed();
                    SHA512CryptoServiceProvider _sha512 = new SHA512CryptoServiceProvider(); 
                    FileStream inputStream = new FileStream(argFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    byte[] inArray = _sha512.ComputeHash(inputStream);
                    inputStream.Close();
                    str = Convert.ToBase64String(inArray);
                }
                catch (Exception exception1)
                {
                    str = "";
                }
                return str;
            }
            return "";
        }

        private void ValidRef(XmlNode argNode, string m_strEnveFilePath)
        {
            string str = argNode.Attributes["URI"].Value.ToString();
            if (!str.StartsWith("#"))
            {
                if (str.StartsWith("\\"))
                    str = str.Substring(1);
                string path = Path.Combine(m_strEnveFilePath, str);
                if (!File.Exists(path))
                {
                    str = argNode.ParentNode.ParentNode.Attributes["Id"].Value + "：" + str;
                    arrNotPassItems.Add(str);
                    arrNotPassDescs.Add("檔案已不存在!");
                }
                string argAlgo = this.GetSubNodeWithName(argNode, "DigestMethod").Attributes["Algorithm"].Value.ToString();
                string innerText = this.GetSubNodeWithName(argNode, "DigestValue").InnerText;
                if (argAlgo.ToUpper() != "HTTP://WWW.W3.ORG/2000/09/XMLDSIG#SHA1" && 
                    argAlgo.ToLower() != "http://www.w3.org/2001/04/xmlenc#sha256" &&
                    argAlgo.ToLower() != "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256" &&
                    argAlgo.ToLower() != "http://www.w3.org/2001/04/xmlenc#sha512")
                {
                    if (argNode.ParentNode.ParentNode.Attributes["Id"] != null)
                        str = argNode.ParentNode.ParentNode.Attributes["Id"].Value +"："+ str;
                    else
                        str = argNode.ParentNode.ParentNode.Name + "：" + str;
                    arrNotPassItems.Add(str);
                    arrNotPassDescs.Add("封裝檔中指定外部檔案雜湊演算法不正確!");
                }
                string fileHash = this.GetFileHash(path, argAlgo);
                if (fileHash != innerText)
                {
                    if (argNode.ParentNode.ParentNode.Attributes["Id"] != null)
                        str = argNode.ParentNode.ParentNode.Attributes["Id"].Value + "：" + str;
                    else
                        str = argNode.ParentNode.ParentNode.Name + "：" + str;
                    arrNotPassItems.Add(str);
                    arrNotPassDescs.Add("檔案雜湊值不一致,原始雜湊值[" + innerText + "],目前雜湊值[" + fileHash + "]");
                }
            }
            else
            {
                try
                {
                    str = str.Substring(1);
                    string xpath = "//*[@Id='" + str + "']";
                    XmlNodeList list = argNode.OwnerDocument.SelectNodes(xpath);
                    if (list.Count == 1)
                    {
                        string outerXml = list[0].OuterXml;
                        string str8 = this.GetSubNodeWithName(argNode, "DigestMethod").Attributes["Algorithm"].Value.ToString();
                        if (str8.ToUpper() != "HTTP://WWW.W3.ORG/2000/09/XMLDSIG#SHA1" && 
                            str8.ToLower() != "http://www.w3.org/2001/04/xmlenc#sha256" &&
                            str8.ToLower() != "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256" &&
                            str8.ToLower() != "http://www.w3.org/2001/04/xmlenc#sha512")
                        {
                            str = argNode.ParentNode.ParentNode.Attributes["Id"].Value + "：" + str;
                            arrNotPassItems.Add(str);
                            arrNotPassDescs.Add("封裝檔中指定雜湊演算法不正確");
                        }
                        outerXml = this.GetCanonicalXML(outerXml);
                        string str9 = this.GetSubNodeWithName(argNode, "DigestValue").InnerText;
                        string nodesha = "";
                        if (str8.ToLower().EndsWith("sha1"))
                        {
                            SHA1CryptoServiceProvider _sha1 = new SHA1CryptoServiceProvider();
                            nodesha = Convert.ToBase64String(_sha1.ComputeHash(C14NTrans(list[0])));
                            if (nodesha != str9)
                                nodesha = Convert.ToBase64String(_sha1.ComputeHash(C14NTransform(list[0])));
                            if (nodesha != str9)
                                nodesha = Convert.ToBase64String(_sha1.ComputeHash(C14NTransformNPW(list[0])));
                            if (nodesha != str9)
                                nodesha = Convert.ToBase64String(_sha1.ComputeHash(C14NTransformNPWENC(list[0])));
                        }
                        else if (str8.ToLower().EndsWith("sha256"))
                        {
                            //SHA256Managed _sha256 = new SHA256Managed();
                            SHA256CryptoServiceProvider _sha256 = new SHA256CryptoServiceProvider();  
                            nodesha = Convert.ToBase64String(_sha256.ComputeHash(C14NTrans(list[0])));
                            if (nodesha != str9)
                                nodesha = Convert.ToBase64String(_sha256.ComputeHash(C14NTransform(list[0])));
                            if (nodesha != str9)
                                nodesha = Convert.ToBase64String(_sha256.ComputeHash(C14NTransformNPW(list[0])));
                            if (nodesha != str9)
                                nodesha = Convert.ToBase64String(_sha256.ComputeHash(C14NTransformNPWENC(list[0])));
                        }
                        else if (str8.ToLower().EndsWith("sha512"))
                        {
                            //SHA512Managed _sha512 = new SHA512Managed();
                            SHA512CryptoServiceProvider _sha512 = new SHA512CryptoServiceProvider(); 
                            nodesha = Convert.ToBase64String(_sha512.ComputeHash(C14NTrans(list[0])));
                            if (nodesha != str9)
                                nodesha = Convert.ToBase64String(_sha512.ComputeHash(C14NTransform(list[0])));
                            if (nodesha != str9)
                                nodesha = Convert.ToBase64String(_sha512.ComputeHash(C14NTransformNPW(list[0])));
                            if (nodesha != str9)
                                nodesha = Convert.ToBase64String(_sha512.ComputeHash(C14NTransformNPWENC(list[0])));
                        }
                        string strHash = this.GetStrHash(outerXml, str8);
                        //string str9 = this.GetSubNodeWithName(argNode, "DigestValue").InnerText;
                        if (nodesha != str9)
                        {
                            string str12 = "封裝檔XML內容雜湊值不一致,原始雜湊值[" + str9 + "],目前雜湊值[" + strHash + "]";
                            str = argNode.ParentNode.ParentNode.Attributes["Id"].Value + "：" + str;
                            arrNotPassItems.Add(str);
                            arrNotPassDescs.Add(str12);
                        }
                    }
                }
                catch (Exception exception1)
                {
                    throw exception1;
                }
            }
        }

        private string GetCanonicalXML(string argXML)
        {
            string str2 = "";
            try
            {
                XmlDocument document = new XmlDocument();
                document.LoadXml(argXML);
                XmlDsigC14NTransform transform = new XmlDsigC14NTransform();
                transform.LoadInput(document);
                Stream output = (Stream)transform.GetOutput(typeof(Stream));
                str2 = new StreamReader(output).ReadToEnd();
                if (str2.IndexOf("&#xD") != -1)
                {
                    str2 = str2.Replace("&#xD;", "");
                }
            }
            catch (Exception exception1)
            {
                
            }
            return str2;
        }

        private XmlNode GetSubNodeWithName(XmlNode argNode, string argName)
        {
            int num2 = argNode.ChildNodes.Count - 1;
            for (int i = 0; i <= num2; i++)
            {
                if (argNode.ChildNodes[i].Name.ToUpper() == argName.ToUpper())
                {
                    return argNode.ChildNodes[i];
                }
            }
            return null;
        }

        private bool ValidSig(XmlNode argNode)
        {
            if (argNode.ChildNodes.Count >= 3)
            {
                try
                {
                    XmlNode subNodeWithName = this.GetSubNodeWithName(argNode, "SignedInfo");
                    XmlNode node2 = this.GetSubNodeWithName(subNodeWithName, "CanonicalizationMethod");
                    XmlNode node4 = this.GetSubNodeWithName(subNodeWithName, "SignatureMethod");
                    XmlNode node3 = this.GetSubNodeWithName(argNode, "KeyInfo");
                    XmlNode node7 = this.GetSubNodeWithName(node3, "X509Data");
                    XmlNode node6 = this.GetSubNodeWithName(node7, "X509Certificate");
                    string outerXml = subNodeWithName.OuterXml;
                    string str = node2.Attributes["Algorithm"].Value.ToString();
                    string innerText = this.GetSubNodeWithName(argNode, "SignatureValue").InnerText;
                    string str2 = node6.InnerText;
                    string str3 = node4.Attributes["Algorithm"].Value.ToString();
                    if (str.ToUpper() != "HTTP://WWW.W3.ORG/TR/2001/REC-XML-C14N-20010315")
                    {
                         
                        this.arrNotPassItems.Add(argNode.Attributes["Id"].Value + "：" + "簽核點簽章");
                        this.arrNotPassDescs.Add("驗章未通過-封裝檔指定之正規化演算法為不正確");
                    }
                    if (str3.ToUpper() != "HTTP://WWW.W3.ORG/2000/09/XMLDSIG#RSA-SHA1" && str3.ToLower() != "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256" && str3.ToLower() != "http://www.w3.org/2001/04/xmldsig-more#rsa-sha512")
                    {
                        this.arrNotPassItems.Add(argNode.Attributes["Id"].Value + "：" + "簽核點簽章");
                        this.arrNotPassDescs.Add("驗章未通過-封裝檔指定之加簽演算法不正確");
                    }
                    //byte[] sha2561 =C14NTransform(subNodeWithName);
                    //byte[] sha2562 =C14NTransformNPW(subNodeWithName);
                    //byte[] sha2563 =C14NTransformNPWENC(subNodeWithName);
                    XmlDocument document = new XmlDocument();
                    document.LoadXml(outerXml);
                    XmlDsigC14NTransform transform = new XmlDsigC14NTransform();
                    transform.LoadInput(document);
                    //byte[] sha256 = transform.GetDigestedOutput(HashAlgorithm.Create("SHA256"));  
                    Stream output = (Stream)transform.GetOutput(typeof(Stream));
                    outerXml = new StreamReader(output).ReadToEnd();
                    //HashAlgorithm managed = null; //GCB不支援
                    //byte[] rgbHash = managed.ComputeHash(bytes); //GCB不支援
                    string algstr = "";
                    byte[] rgbHash = null;
                    byte[] rawData =Convert.FromBase64String(str2)  ;//Convert.FromBase64CharArray(str2.ToCharArray(), 0, str2.Length);
                    byte[] rgbSignature = Convert.FromBase64String(innerText);
                    byte[] bytes = Encoding.UTF8.GetBytes(outerXml);

                    if (str3.ToLower().EndsWith("sha1"))
                    {
                        rgbHash = new SHA1CryptoServiceProvider().ComputeHash(bytes);
                        algstr = "SHA1";
                    }
                    else if (str3.ToLower().EndsWith("sha256"))
                    {
                        rgbHash = new SHA256CryptoServiceProvider().ComputeHash(bytes);
                        algstr = "SHA256";
                    }
                    else if (str3.ToLower().EndsWith("sha512"))
                    {
                        rgbHash = new SHA512CryptoServiceProvider().ComputeHash(bytes);
                        algstr = "SHA512";
                    }
                   
                    X509Certificate2 certificate = new X509Certificate2(rawData);
                    RSACryptoServiceProvider.UseMachineKeyStore = true;
                    //bool certchainok = certificate.Verify();//檢測憑證合法性
                    //需使用framework3.5以上，解決winxp無法使用SHA2問題
                    //bool flag3 = ((RSACryptoServiceProvider)certificate.PublicKey.Key).VerifyData(sha256, CryptoConfig.MapNameToOID(algstr), rgbSignature);
                    bool flag2 = ((RSACryptoServiceProvider)certificate.PublicKey.Key).VerifyHash(rgbHash, CryptoConfig.MapNameToOID(algstr), rgbSignature);
                    if (!flag2)
                    {
                        this.arrNotPassItems.Add(argNode.Attributes["Id"].Value + "：" + "簽核點簽章");
                        this.arrNotPassDescs.Add("驗章未通過");
                    }
                    XmlNode nextSibling = argNode.NextSibling;
                    if ((nextSibling != null) && (nextSibling.Name == "簽章時戳"))
                    {
                    }
                    return flag2;
                }
                catch (Exception exception1)
                {
                    throw exception1;
                }
            }
            return true;
        }

        private byte[] C14NTransform(XmlNode xn)
        {
            byte[] buffer3;
            byte[] bytes = Encoding.UTF8.GetBytes(xn.OuterXml);
            XmlDsigC14NTransform transform = new XmlDsigC14NTransform();
            using (MemoryStream stream = new MemoryStream(bytes))
            {
                transform.LoadInput(stream);
            }
            using (Stream stream2 = (Stream)transform.GetOutput())
            {
                buffer3 = new byte[((int)(stream2.Length - 1L)) + 1];
                stream2.Read(buffer3, 0, (int)stream2.Length);
            }
            return buffer3;
        }

        private byte[] C14NTransformNPW(XmlNode xn)
        {
            byte[] buffer3;
            XmlDocument document = new XmlDocument();
            //document.PreserveWhitespace = true;
            document.LoadXml(xn.OuterXml);
            
            byte[] bytes = Encoding.UTF8.GetBytes(document.OuterXml);
            XmlDsigC14NTransform transform = new XmlDsigC14NTransform();
            using (MemoryStream stream = new MemoryStream(bytes))
            {
                transform.LoadInput(stream);
            }
            using (Stream stream2 = (Stream)transform.GetOutput())
            {
                buffer3 = new byte[((int)(stream2.Length - 1L)) + 1];
                stream2.Read(buffer3, 0, (int)stream2.Length);
            }
            return buffer3;
        }

        private byte[] C14NTransformNPWENC(XmlNode xn)
        {
            byte[] buffer3;
            string str = "UTF-8";
            if (xn.OwnerDocument.FirstChild.NodeType == XmlNodeType.XmlDeclaration)
            {
                str = ((XmlDeclaration)xn.OwnerDocument.FirstChild).Encoding.ToUpper();
            }
            XmlDocument document = new XmlDocument();
            //document.PreserveWhitespace = true;
            document.LoadXml(xn.OuterXml);
            byte[] bytes = Encoding.UTF8.GetBytes(xn.OuterXml);
            XmlDsigC14NTransform transform = new XmlDsigC14NTransform();
            using (MemoryStream stream = new MemoryStream(bytes))
            {
                transform.LoadInput(stream);
            }
            using (Stream stream2 = (Stream)transform.GetOutput())
            {
                buffer3 = new byte[stream2.Length];
                stream2.Read(buffer3, 0, (int)stream2.Length);
            }
            
            if (str!="UNICODE" && str!="BIG5" && str!="CNS11643")
            {
                return buffer3;
            }
            System.Text.Decoder decoder = new UTF8Encoding().GetDecoder();
            char[] chars = new char[decoder.GetCharCount(buffer3, 0, buffer3.Length)];
            decoder.GetChars(buffer3, 0, buffer3.Length, chars, 0);
            string s = new string(chars);
            if (str.Equals("BIG5") || str.Equals("CNS11643"))
            {
                return Encoding.Default.GetBytes(s);
            }
            return Encoding.Unicode.GetBytes(s);
        }

        private byte[] C14NTrans(XmlNode xn)
        {
            byte[] toReturn = null;
            string _encoding = "UTF-8";//預設為UTF-8編碼
            if (xn.OwnerDocument.FirstChild.NodeType == XmlNodeType.XmlDeclaration)
            {
                //取得xml declaration中指定編碼方式
                _encoding = ((XmlDeclaration)(xn.OwnerDocument.FirstChild)).Encoding.ToUpper();
            }

            //先把字串轉為MemoryStream作為正規化方法的輸入
            byte[] encData_byte = System.Text.Encoding.UTF8.GetBytes(xn.OuterXml);
            System.Security.Cryptography.Xml.XmlDsigC14NTransform trans = new System.Security.Cryptography.Xml.XmlDsigC14NTransform();
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream(encData_byte))
            {
                trans.LoadInput(ms);
            }

            //取得輸出
            byte[] todecode_byte;
            using (System.IO.Stream stream = (System.IO.Stream)(trans.GetOutput()))
            {
                todecode_byte = new byte[stream.Length];
                stream.Read(todecode_byte, 0, (int)(stream.Length));
            }

            //依據不同的編碼方式決定傳回的位元組陣列
            switch (_encoding)
            {
                case "UTF-8":
                    toReturn = todecode_byte;
                    break;
                default:
                    string result = Encoding.UTF8.GetString(todecode_byte);

                    string code;
                    if (_encoding == "CNS11643")
                        code = "Big5";
                    else
                        code = _encoding;

                    byte[] encData_byte1 = Encoding.GetEncoding(code).GetBytes(result);
                    toReturn = encData_byte1;
                    break;
            }

            return toReturn;
        }

        /// <summary>
        /// 計算檔案hash值
        /// </summary>
        /// <param name="SignatureId"></param>
        /// <returns></returns>
        private string CalcFileHash(string SignatureId)
        {
            string rtn = "";
            if (_XmlForSign != null)
            {
                //fsxml.SetXMLReferenceDirs(_FilePath);
                //rtn = fsxml.FSXMLCAPI_CalcTemplateDigest(_XmlForSign.OuterXml, SignatureId);
                gpkixml.SetXMLReferenceDirs(_FilePath);
                rtn = gpkixml.FSGPKI_XMLCalcTemplateDigest(_XmlForSign.OuterXml, SignatureId);
            }
            else
                rtn = gpkixml.FSGPKI_XMLCalcTemplateFileDigest(Path.Combine(_FilePath, _XmlFileName), SignatureId);
                //rtn = fsxml.FSXMLCAPI_CalcTemplateFileDigest(Path.Combine(_FilePath, _XmlFileName), SignatureId);

            if (gpkixml.get_lastError() != 0)
            {
                rtn = "ERR:" + ErrorMsgTable.GetErrMessage(gpkixml.get_lastError());
                return rtn;
            }
            else
            {
                XmlDocument rtnDoc = null;
                using (MemoryStream ms = new MemoryStream(Convert.FromBase64String(rtn)))
                {
                    rtnDoc = new XmlDocument();
                    rtnDoc.XmlResolver = null;
                    rtnDoc.Load(ms);
                    rtnDoc.PreserveWhitespace = true;
                    ms.Close();
                    ms.Dispose();
                }
                if (rtnDoc != null)
                    return rtnDoc.OuterXml;
                else
                    return "ERR:XmlDoc錯誤";
            }
        }

        /// <summary>
        /// 對某一節點加簽
        /// </summary>
        /// <param name="xmlData"></param>
        /// <param name="SignatureId"></param>
        /// <param name="pinCode"></param>
        /// <returns></returns>
        private string SignXmlNode(string xmlData, string SignatureId, string pinCode)
        {
            const int FSCAPI_FLAG_SELCERT_MANUAL = 0x00000001;
            const int FSCAPI_FLAG_SELCERT_AUTO = 0x00000002;
            const int FSCAPI_FLAG_SELCERT_SELFAUTO = 0x00000003;
            const int FSCAPI_FLAG_SELCERT_AFTER = 0x00000004;
            const int FSCAPI_FLAG_SELCERT_OLDEST = 0x00000008;
            const int FSCAPI_FLAG_SELCERT_CHECKVALID = 0x00000010;
            const int FS_FLAG_USE_FILE = 0x00000020;
            const int FS_FLAG_DETACHMSG = 0x00004000;
            const int FS_FLAG_NOHASHOID = 0x00020000;

            try
            {
                string rtnNode = "";
                if (pinCode.StartsWith("serial:"))
                {
                    X509Certificate2 signCert = null;
                    X509Store store = new X509Store(StoreLocation.CurrentUser);
                    store.Open(OpenFlags.MaxAllowed);
                    X509Certificate2Enumerator enumerator = store.Certificates.GetEnumerator();
                    while (enumerator.MoveNext())
                    {
                        X509Certificate2 current = enumerator.Current;
                        //System.Windows.Forms.MessageBox.Show(current.SerialNumber + ":" + pinCode.Substring(pinCode.IndexOf(":") + 1).ToUpper()); 
                        if (current.SerialNumber == pinCode.Substring(pinCode.IndexOf(":") + 1).ToUpper())
                        {
                            signCert = current;
                            break;
                        }
                    }
                    if (signCert == null) return "ERR: 找不到相對應的憑證可加簽!";
                    string strEndDate = signCert.GetExpirationDateString();
                    if (!chkValidDate(strEndDate))
                    {
                        return "ERR:憑證已過期。";
                    }
                    //System.Windows.Forms.MessageBox.Show(signCert.Subject + " ; " + signCert.IssuerName.Name);
                    string subjectstr = "";
                    foreach (string s in signCert.Subject.Split(new char[] { ',' }))
                    {
                        string sub = s.TrimStart(new char[] { ' ' });
                        if (sub.StartsWith("DC=")) continue;
                        subjectstr += sub + "\n";
                    }

                    string issuerstr = "";
                    foreach (string s in signCert.Issuer.Split(new char[] { ',' }))
                    {
                        string sub = s.TrimStart(new char[] { ' ' });
                        if (sub.StartsWith("DC=")) continue;
                        issuerstr += sub + "\n";
                    }
                    FSXMLATLClass fsxml = new FSXMLATLClass();
                    rtnNode = fsxml.FSXMLCAPI_SignTemplateDigest(Convert.ToBase64String(UTF8Encoding.UTF8.GetBytes(xmlData)), SignatureId, subjectstr, issuerstr, 2, 0);
                    
                    //SignString = X.FSCAPIFileSign(SourcePath, SerialNum, "", "", FS_FLAG_USE_FILE | FS_FLAG_DETACHMSG | FS_FLAG_NOHASHOID, 0);
                    int retValue = fsxml.GetErrorCode();
                    fsxml = null;
                    if (retValue != 0)
                    {
                        string strErrMessage = "ERR:" + ErrorMsgTable.GetErrMessage(retValue);
                        return strErrMessage;
                    }

                }
                else
                {
                    gpkixml.FSGPKI_SetReaderSlot(0, _cardReaderSlot);

                    if (_PKCS11Driver != "")
                        gpkixml.FSGPKI_SetPKCS11Driver(_PKCS11Driver);

                    int iFlags = FSCAPI_FLAG_SELCERT_CHECKVALID;
                    if (string.IsNullOrEmpty(_CertX509Str))
                        _CertX509Str = gpkixml.FSGPKI_EnumCertsToString(0);
                    string strEndDate = gpkixml.FSCAPICertGetNotAfter(_CertX509Str, 0);
                    if (!chkValidDate(strEndDate))
                    {
                        return "ERR:憑證已過期。";
                    }
                    rtnNode = gpkixml.FSGPKI_XMLSignTemplateDigest(pinCode, Convert.ToBase64String(UTF8Encoding.UTF8.GetBytes(xmlData)), SignatureId, iFlags);

                    //rtnNode = xmlData;
                    //return rtnNode;
                    //string rtnNode = gpkixml.FSGPKI_XMLSignTemplateFile(pinCode, _FilePath + @"\afterhash.xml", SignatureId, iFlags);
                    //File.Delete(Path.Combine(_FilePath , "afterhash.xml"));
                    //rtnNode = rtnNode.Replace("\r\n\0", ""); 
                    if (gpkixml.get_lastError() != 0)
                    {
                        throw new Exception("ERR:SignXmlNode錯誤," + ErrorMsgTable.GetErrMessage(gpkixml.get_lastError()));
                    }
                    
                }
                return rtnNode;

            }
            catch (Exception err)
            {
                return "ERR:" + err.Message;
            }
            finally
            {
                //objDispose(fsxml);
                //objDispose(gpkixml);
                //fsxml = null;
                gpkixml = null;
            }
        }

        private void objDispose(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(obj);

                obj = null;

                GC.Collect();

                GC.WaitForPendingFinalizers();
            }
            catch (Exception ex)
            {
            }
        }


        /// <summary>
        /// 檢查憑證是否到期
        /// </summary>
        /// <param name="strEndDate"></param>
        /// <returns></returns>
        private bool chkValidDate(string strEndDate)
        {
            DateTime dteEndDate;
            if (!DateTime.TryParse(strEndDate, out dteEndDate))
            {
                string _Date = strEndDate.Substring(0, 4) + "/" + strEndDate.Substring(4, 2) + "/" + strEndDate.Substring(6, 2);
                if (!DateTime.TryParse(_Date, out dteEndDate))
                    return false;
            }

            if (dteEndDate < DateTime.Today)
                return false;
            else
                return true;
        }
    }
}
